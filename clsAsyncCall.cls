VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAsyncCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsAsyncCall - internal worker for one in-flight async HTTP call.
'
' Not part of the public API. Users build a clsHttpRequest and call
' .SendAsync / .SendAsyncTo; the builder constructs a clsAsyncCall,
' registers it with modRestPump, and returns the request id.
'
' Lifecycle:
'   1. Builder calls StartTyped or StartLateBound, supplying the verb,
'      url, body, headers, and a callback (either typed IHttpCallback or
'      a (target, methodName) pair). The call creates MSXML, opens,
'      applies headers, sends, and asks modRestPump.EnsurePump to begin
'      polling.
'   2. modRestPump.TimerProc calls PollAndMaybeDispatch on every tick.
'      Once MSXML reports readyState=4 the call builds a clsHttpResponse,
'      removes itself from the registry, releases MSXML, and fires the
'      callback.
'   3. Abort short-circuits everything: m_completed flips True, MSXML is
'      aborted, the http reference is released. No callback fires after
'      a successful Abort.
'
' Why polling instead of MSXML's onreadystatechange COM sink: four sink
' approaches failed on this Access/MSXML build. See async_architecture.md.

Private Const PROGID_XMLHTTP As String = "MSXML2.XMLHTTP.6.0"

Private m_http               As Object        ' MSXML2.XMLHTTP60 (late-bound)
Private m_requestId          As Long

' Exactly one of these callback shapes is populated:
Private m_typedCallback      As IHttpCallback ' typed path
Private m_lateBoundTarget    As Object        ' late-bound path
Private m_lateBoundMethod    As String        ' late-bound method name

Private m_tag                As Variant
Private m_completed          As Boolean       ' one-shot guard

'==========================================================================
'  Lifecycle
'==========================================================================

Public Sub StartTyped(ByVal requestId As Long, _
                      ByVal method As String, _
                      ByVal url As String, _
                      ByVal body As String, _
                      ByVal headers As Object, _
                      ByVal callback As IHttpCallback, _
                      ByRef tag As Variant)
    m_requestId = requestId
    Set m_typedCallback = callback
    AssignTag tag
    Run method, url, body, headers
End Sub

Public Sub StartLateBound(ByVal requestId As Long, _
                          ByVal method As String, _
                          ByVal url As String, _
                          ByVal body As String, _
                          ByVal headers As Object, _
                          ByVal target As Object, _
                          ByVal methodName As String, _
                          ByRef tag As Variant)
    m_requestId = requestId
    Set m_lateBoundTarget = target
    m_lateBoundMethod = methodName
    AssignTag tag
    Run method, url, body, headers
End Sub

Public Property Get RequestId() As Long
    RequestId = m_requestId
End Property

' Cancel the call. Idempotent. No callback fires after Abort, even if
' MSXML produces a final state transition.
Public Sub Abort()
    If m_completed Then Exit Sub
    m_completed = True
    On Error Resume Next
    If Not m_http Is Nothing Then m_http.Abort
    On Error GoTo 0
    ReleaseHttp
End Sub

' Called by modRestPump.TimerProc on every poll tick. Reads MSXML's
' current readyState; once it hits 4, loads the response, removes self
' from the registry, releases m_http, and fires the user callback.
Public Sub PollAndMaybeDispatch()
    If m_completed Then Exit Sub
    If m_http Is Nothing Then Exit Sub

    Dim rs As Long
    On Error Resume Next
    rs = m_http.readyState
    On Error GoTo 0
    If rs <> 4 Then Exit Sub

    ' Pin a strong ref on this stack frame: RemovePending drops the
    ' registry's external reference; without selfRef the class could
    ' terminate mid-method.
    Dim selfRef As clsAsyncCall
    Set selfRef = Me

    m_completed = True

    Dim resp As clsHttpResponse
    Set resp = New clsHttpResponse
    On Error Resume Next
    resp.LoadFromMSXML m_http
    If Err.Number <> 0 Then
        resp.SetTransportError "Transport error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    modRestPump.RemovePending m_requestId
    ReleaseHttp

    Dispatch resp
End Sub

'==========================================================================
'  Internals
'==========================================================================

' Setup phase. Creates the MSXML object, opens, applies headers, fires
' the send (which returns immediately in async mode), and hands
' completion detection to modRestPump. If any step raises, builds a
' transport-error response and dispatches synchronously.
Private Sub Run(ByVal method As String, _
                ByVal url As String, _
                ByVal body As String, _
                ByVal headers As Object)
    Dim stage As String
    On Error GoTo Failed

    stage = "CreateObject"
    Set m_http = CreateObject(PROGID_XMLHTTP)

    stage = "Open"
    m_http.Open method, url, True            ' True = async

    stage = "ApplyHeaders"
    ApplyHeaders headers

    ' Identity encoding: MSXML can't decode gzip / deflate. setRequestHeader
    ' on an unsupported encoding is a no-op on most builds, hence the
    ' local Resume Next.
    On Error Resume Next
    m_http.setRequestHeader "Accept-Encoding", "identity"
    On Error GoTo Failed

    stage = "send"
    m_http.send body

    modRestPump.EnsurePump
    Exit Sub

Failed:
    Dim errResp As clsHttpResponse
    Set errResp = New clsHttpResponse
    errResp.SetTransportError "Run failed at " & stage & ": #" & _
                              Err.Number & " " & Err.Description
    m_completed = True
    modRestPump.RemovePending m_requestId
    ReleaseHttp
    Dispatch errResp
End Sub

Private Sub ApplyHeaders(ByVal headers As Object)
    If headers Is Nothing Then Exit Sub
    Dim keys As Variant
    keys = headers.Keys
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        On Error Resume Next
        m_http.setRequestHeader CStr(keys(i)), CStr(headers(keys(i)))
        On Error GoTo 0
    Next i
End Sub

' Variants can hold either a primitive or an object reference; assigning
' the wrong way for the contained type raises. IsObject picks the path.
Private Sub AssignTag(ByRef tag As Variant)
    If IsObject(tag) Then
        Set m_tag = tag
    Else
        m_tag = tag
    End If
End Sub

' Fire the user callback. Errors inside the user's callback are
' swallowed (logged to Debug) - we never let user-code errors corrupt
' the pump's state machine.
Private Sub Dispatch(ByVal resp As clsHttpResponse)
    On Error Resume Next
    If Not m_typedCallback Is Nothing Then
        m_typedCallback.OnResponse resp, m_tag
    ElseIf Len(m_lateBoundMethod) > 0 And Not m_lateBoundTarget Is Nothing Then
        CallByName m_lateBoundTarget, m_lateBoundMethod, VbMethod, resp, m_tag
    End If
    If Err.Number <> 0 Then
        Debug.Print "clsAsyncCall: callback raised error " & _
                    Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub ReleaseHttp()
    On Error Resume Next
    If Not m_http Is Nothing Then Set m_http = Nothing
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    If Not m_http Is Nothing Then
        On Error Resume Next
        m_http.Abort
        On Error GoTo 0
    End If
    ReleaseHttp
End Sub
