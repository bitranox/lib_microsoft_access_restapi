Attribute VB_Name = "modRestPump"
Option Explicit

' modRestPump - module-level anchor for in-flight async calls plus
' the Win32 SetTimer polling pump that drives their completion.
'
' Internal-ish module: users normally go through modRest, but this
' module's Cancel functions are also exposed as modRest.Cancel / .CancelAll
' for convenience.
'
' Two responsibilities, kept together because each exists for the other:
'
'   1. Registry: Register / RemovePending / IsPending / Cancel /
'      CancelAll / PendingCount. Keeps a Dictionary keyed by Long
'      request-id; without this anchor the caller-less clsAsyncCall
'      would be released the moment SendAsync returns.
'
'   2. Polling pump: EnsurePump / TimerProc. Starts a Win32 WM_TIMER
'      whose callback iterates the registry on each tick and asks each
'      pending call to poll its own readyState. The pump self-stops
'      when the registry empties.
'
' Why polling instead of MSXML's onreadystatechange event sink: four
' COM-sink approaches were tried and all failed on this Access/MSXML
' build (see async_architecture.md). Polling sidesteps the whole class
' of issues at the cost of ~25ms completion latency.

' --- Win32 declares (MUST come before any Sub/Function in a VBA module) ---

Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long

' 25ms gives ~40Hz polling. Far below WM_TIMER coalescing thresholds,
' well below human-perceptible latency, and only one COM readyState
' read per tick per pending call -- effectively free at typical
' usage (1-5 concurrent requests).
Private Const POLL_INTERVAL_MS As Long = 25

Private g_pending As Object              ' Scripting.Dictionary, key = Long requestId
Private g_nextId  As Long
Private g_timerId As LongPtr

'==========================================================================
'  Registry
'==========================================================================

Public Function Register(ByVal asyncCall As Object) As Long
    EnsureDict
    g_nextId = g_nextId + 1
    g_pending.Add g_nextId, asyncCall
    Register = g_nextId
End Function

Public Sub RemovePending(ByVal id As Long)
    If g_pending Is Nothing Then Exit Sub
    If g_pending.Exists(id) Then g_pending.Remove id
End Sub

Public Function PendingCount() As Long
    If g_pending Is Nothing Then Exit Function
    PendingCount = g_pending.Count
End Function

Public Function IsPending(ByVal id As Long) As Boolean
    If g_pending Is Nothing Then Exit Function
    IsPending = g_pending.Exists(id)
End Function

Public Sub Cancel(ByVal id As Long)
    If g_pending Is Nothing Then Exit Sub
    If Not g_pending.Exists(id) Then Exit Sub
    Dim asyncCall As Object
    Set asyncCall = g_pending(id)
    On Error Resume Next
    asyncCall.Abort                      ' clsAsyncCall.Abort
    On Error GoTo 0
    g_pending.Remove id
End Sub

Public Sub CancelAll()
    If g_pending Is Nothing Then Exit Sub
    Dim ids As Variant
    ids = g_pending.Keys
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        Cancel CLng(ids(i))
    Next i
End Sub

Private Sub EnsureDict()
    If g_pending Is Nothing Then Set g_pending = CreateObject("Scripting.Dictionary")
End Sub

'==========================================================================
'  Polling pump
'==========================================================================

' Called by clsAsyncCall.Run right after m_http.send. Idempotent;
' starts the WM_TIMER if it isn't already running. The timer stops
' itself once g_pending is empty.
Public Sub EnsurePump()
    If g_timerId <> 0 Then Exit Sub
    g_timerId = SetTimer(0&, 0&, POLL_INTERVAL_MS, AddressOf TimerProc)
End Sub

Private Sub StopPump()
    If g_timerId = 0 Then Exit Sub
    KillTimer 0&, g_timerId
    g_timerId = 0
End Sub

' WM_TIMER callback. MUST be Public so AddressOf can resolve it. MUST NOT
' let a VBA error escape - an unhandled error inside a Win32 callback
' takes Access down with it.
Public Sub TimerProc(ByVal hWnd As LongPtr, _
                     ByVal uMsg As Long, _
                     ByVal idEvent As LongPtr, _
                     ByVal dwTime As Long)
    On Error Resume Next

    If g_pending Is Nothing Or PendingCount() = 0 Then
        StopPump
        Exit Sub
    End If

    ' Snapshot keys before iterating: PollAndMaybeDispatch removes
    ' completed entries from g_pending and would otherwise invalidate
    ' the live key sequence.
    Dim ids As Variant
    ids = g_pending.Keys
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        Dim id As Long
        id = CLng(ids(i))
        If g_pending.Exists(id) Then
            Dim asyncCall As Object
            Set asyncCall = g_pending(id)
            If Not asyncCall Is Nothing Then asyncCall.PollAndMaybeDispatch
        End If
    Next i

    If PendingCount() = 0 Then StopPump
End Sub
