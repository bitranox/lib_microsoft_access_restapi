VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsHttpRequest - mutable builder for one HTTP call.
'
' Created via `modRest.NewRequest` (or implicitly by the verb-convenience
' functions on modRest). Set Method / Url / Body / Headers / TimeoutMs,
' then call Send for a synchronous result or SendAsync / SendAsyncTo for
' an asynchronous one.
'
' Headers and timeout layer over modRest's module-level defaults: request-
' specific values win on conflict; if TimeoutMs stays 0 the module
' default is used. Url resolution applies modRest.BaseUrl when the url
' is relative.
'
' Usage:
'
'   ' Simple, fluent-via-With:
'   With modRest.NewRequest
'       .Method = "POST"
'       .Url = "/users"
'       .Body = "name=Alice"
'       Set r = .Send
'   End With
'
'   ' Async typed:
'   id = modRest.NewRequest.Configure("GET", "/users/42").SendAsync(Me, "user-42")
'
'   ' Async from a standard module (late-bound target):
'   id = modRest.NewRequest.Configure("GET", "/users/42") _
'                .SendAsyncTo(host, "OnUserLoaded", "user-42")

Private Const PROGID_SERVER_XMLHTTP As String = "MSXML2.ServerXMLHTTP.6.0"
Private Const JSON_CONTENT_TYPE     As String = "application/json; charset=utf-8"

Private m_method     As String
Private m_url        As String
Private m_body       As String
Private m_headers    As Object       ' Scripting.Dictionary, case-insensitive
Private m_timeoutMs  As Long         ' 0 means "use modRest.DefaultTimeoutMs"

Private Sub Class_Initialize()
    m_method = "GET"
    Set m_headers = CreateObject("Scripting.Dictionary")
    m_headers.CompareMode = 1            ' TextCompare = case-insensitive keys
End Sub

'==========================================================================
'  Public properties
'==========================================================================

Public Property Get Method() As String
    Method = m_method
End Property
Public Property Let Method(ByVal v As String)
    m_method = v
End Property

Public Property Get Url() As String
    Url = m_url
End Property
Public Property Let Url(ByVal v As String)
    m_url = v
End Property

Public Property Get Body() As String
    Body = m_body
End Property
Public Property Let Body(ByVal v As String)
    m_body = v
End Property

Public Property Get TimeoutMs() As Long
    TimeoutMs = m_timeoutMs
End Property
Public Property Let TimeoutMs(ByVal v As Long)
    m_timeoutMs = v
End Property

' Mutable Dictionary. Set request-specific headers via:
'   req.Headers.Add "X-Foo", "bar"
' or use SetHeader(key, value) which removes-then-adds.
Public Property Get Headers() As Object
    Set Headers = m_headers
End Property

'==========================================================================
'  Convenience setters
'==========================================================================

' Set method + url in one call. Returns Me so the verb shortcuts compose
' nicely: `modRest.NewRequest.Configure("GET", url).Send`.
Public Function Configure(ByVal method As String, ByVal url As String) As clsHttpRequest
    m_method = method
    m_url = url
    Set Configure = Me
End Function

' Replace any existing value for the given header key (case-insensitive).
Public Sub SetHeader(ByVal key As String, ByVal value As String)
    If m_headers.Exists(key) Then m_headers.Remove key
    m_headers.Add key, value
End Sub

' Merge another header dictionary into this request's headers. Caller-
' supplied keys override existing values. Accepts Nothing / Missing as
' a no-op for ergonomic Optional-Variant passing from the facade.
Public Sub MergeHeaders(Optional ByRef other As Variant)
    If IsMissing(other) Then Exit Sub
    If Not IsObject(other) Then Exit Sub
    If other Is Nothing Then Exit Sub
    Dim keys As Variant
    keys = other.Keys
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        SetHeader CStr(keys(i)), CStr(other(keys(i)))
    Next i
End Sub

' Serialize value as JSON, install it as Body, and set Content-Type
' (only if the caller hasn't already provided one).
Public Sub SetJsonBody(ByVal value As Variant)
    m_body = modJson.JsonStringify(value)
    If Not m_headers.Exists("Content-Type") Then
        m_headers.Add "Content-Type", JSON_CONTENT_TYPE
    End If
End Sub

'==========================================================================
'  Send
'==========================================================================

' Synchronous send. Blocks the Access STA until the response arrives or
' the timeout expires. Returns a clsHttpResponse; transport errors do not
' raise -- they populate response.ErrorMessage and leave Status at 0.
Public Function Send() As clsHttpResponse
    Dim resp As clsHttpResponse
    Set resp = New clsHttpResponse

    Dim http As Object
    On Error GoTo Failed
    Set http = CreateObject(PROGID_SERVER_XMLHTTP)

    Dim t As Long
    t = EffectiveTimeoutMs()
    http.setTimeouts t, t, t, t

    http.Open m_method, EffectiveUrl(), False        ' False = synchronous
    ApplyHeadersTo http

    On Error Resume Next
    http.setRequestHeader "Accept-Encoding", "identity"
    On Error GoTo Failed

    http.send m_body
    resp.LoadFromMSXML http
    Set Send = resp
    Exit Function

Failed:
    resp.SetTransportError "Transport error: " & Err.Description
    Set Send = resp
End Function

' Async send with a typed IHttpCallback. Returns the request id; the
' callback fires later (from modRestPump's WM_TIMER) with the
' clsHttpResponse and the supplied tag.
Public Function SendAsync(ByVal callback As IHttpCallback, _
                          Optional ByRef tag As Variant) As Long
    Dim asyncCall As clsAsyncCall
    Set asyncCall = New clsAsyncCall
    SendAsync = modRestPump.Register(asyncCall)
    asyncCall.StartTyped SendAsync, m_method, EffectiveUrl(), m_body, _
                         EffectiveHeaders(), callback, tag
End Function

' Async send with a late-bound callback target. The target's public
' method `methodName` must have signature
' `(response As clsHttpResponse, tag As Variant)`. Use this when the
' caller is a standard module that can't Implements IHttpCallback.
Public Function SendAsyncTo(ByVal target As Object, _
                            ByVal methodName As String, _
                            Optional ByRef tag As Variant) As Long
    Dim asyncCall As clsAsyncCall
    Set asyncCall = New clsAsyncCall
    SendAsyncTo = modRestPump.Register(asyncCall)
    asyncCall.StartLateBound SendAsyncTo, m_method, EffectiveUrl(), m_body, _
                             EffectiveHeaders(), target, methodName, tag
End Function

'==========================================================================
'  Effective-value accessors
'==========================================================================

' The URL this request will actually hit, with modRest.BaseUrl applied if
' the request's Url is relative. See modRest.JoinUrl for the resolution
' rules. Public so callers can debug "what URL is this gonna send to?".
Public Function EffectiveUrl() As String
    EffectiveUrl = modRest.JoinUrl(modRest.BaseUrl, m_url)
End Function

' The timeout this request will actually use: m_timeoutMs if set, else
' modRest.DefaultTimeoutMs.
Public Function EffectiveTimeoutMs() As Long
    If m_timeoutMs > 0 Then
        EffectiveTimeoutMs = m_timeoutMs
    Else
        EffectiveTimeoutMs = modRest.DefaultTimeoutMs
    End If
End Function

' A fresh Dictionary with module-default headers and request-specific
' headers merged. Request-specific keys override defaults on conflict.
' Always returns a new dict so the caller can safely retain or modify it.
' Public so callers + tests can introspect the merged set.
Public Function EffectiveHeaders() As Object
    Dim merged As Object
    Set merged = CreateObject("Scripting.Dictionary")
    merged.CompareMode = 1

    Dim defaults As Object
    Set defaults = modRest.DefaultHeaders
    If Not defaults Is Nothing Then
        Dim dk As Variant
        For Each dk In defaults.Keys
            merged.Add CStr(dk), CStr(defaults(dk))
        Next dk
    End If

    Dim rk As Variant
    For Each rk In m_headers.Keys
        If merged.Exists(CStr(rk)) Then merged.Remove CStr(rk)
        merged.Add CStr(rk), CStr(m_headers(rk))
    Next rk

    Set EffectiveHeaders = merged
End Function

' Apply effective headers to a live MSXML object (sync path).
Private Sub ApplyHeadersTo(ByVal http As Object)
    Dim merged As Object
    Set merged = EffectiveHeaders()
    Dim keys As Variant
    keys = merged.Keys
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        On Error Resume Next
        http.setRequestHeader CStr(keys(i)), CStr(merged(keys(i)))
        On Error GoTo 0
    Next i
End Sub
