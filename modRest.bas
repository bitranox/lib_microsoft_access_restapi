Attribute VB_Name = "modRest"
Option Explicit

' modRest - public facade. The only module callers normally touch.
'
' One-line verbs (sync) return a clsHttpResponse:
'
'   HttpGet, HttpDelete                              (no body)
'   HttpPost, HttpPut, HttpPatch                     (string body)
'   HttpPostJson, HttpPutJson, HttpPatchJson         (Variant -> JSON body)
'
' Same verbs with `Async` suffix return a Long request id immediately;
' completion fires the supplied IHttpCallback on the Access STA via
' modRestPump's WM_TIMER. For late-bound async (callers in standard
' modules), use the builder: `modRest.NewRequest.Configure(...).SendAsyncTo`.
'
' Configuration (shared across all calls):
'
'   modRest.BaseUrl              = "https://api.example.com"
'   modRest.DefaultTimeoutMs     = 10000
'   modRest.AddDefaultHeader        "Authorization", "Bearer " & token
'
' Cancellation:
'
'   modRest.Cancel(id), modRest.CancelAll, modRest.PendingCount
'
' Advanced: modRest.NewRequest exposes a clsHttpRequest builder for
' anything the verb facade doesn't cover (custom methods, per-call
' timeout, programmatic body construction, late-bound async).

Private g_baseUrl          As String
Private g_defaultTimeoutMs As Long
Private g_defaultHeaders   As Object         ' Scripting.Dictionary

'==========================================================================
'  Configuration
'==========================================================================

Public Property Get BaseUrl() As String
    BaseUrl = g_baseUrl
End Property
Public Property Let BaseUrl(ByVal v As String)
    g_baseUrl = v
End Property

Public Property Get DefaultTimeoutMs() As Long
    EnsureDefaults
    DefaultTimeoutMs = g_defaultTimeoutMs
End Property
Public Property Let DefaultTimeoutMs(ByVal v As Long)
    EnsureDefaults
    g_defaultTimeoutMs = v
End Property

' Read-only access to the default-headers Dictionary. To mutate, use the
' helper subs (AddDefaultHeader etc.) so we keep one consistent shape.
Public Function DefaultHeaders() As Object
    EnsureDefaults
    Set DefaultHeaders = g_defaultHeaders
End Function

Public Sub AddDefaultHeader(ByVal key As String, ByVal value As String)
    EnsureDefaults
    If g_defaultHeaders.Exists(key) Then g_defaultHeaders.Remove key
    g_defaultHeaders.Add key, value
End Sub

Public Sub RemoveDefaultHeader(ByVal key As String)
    EnsureDefaults
    If g_defaultHeaders.Exists(key) Then g_defaultHeaders.Remove key
End Sub

Public Sub ClearDefaultHeaders()
    EnsureDefaults
    g_defaultHeaders.RemoveAll
End Sub

Private Sub EnsureDefaults()
    If g_defaultHeaders Is Nothing Then
        Set g_defaultHeaders = CreateObject("Scripting.Dictionary")
        g_defaultHeaders.CompareMode = 1
        g_defaultTimeoutMs = 30000           ' 30s
    End If
End Sub

'==========================================================================
'  URL helper (exposed so callers + tests can resolve relative paths
'  against a base without having to construct a clsHttpRequest)
'==========================================================================

' Join a base URL and a (possibly relative) path. Boundary handling:
'   - Absolute path (contains "://")                  -> returned as-is
'   - Empty base                                       -> path returned as-is
'   - Base ends with "/" AND path starts with "/"     -> strip one slash
'   - Neither side has a slash and path is non-empty  -> insert one slash
'   - Otherwise                                        -> simple concatenation
Public Function JoinUrl(ByVal base As String, ByVal path As String) As String
    If Len(base) = 0 Or InStr(path, "://") > 0 Then
        JoinUrl = path
    ElseIf Right$(base, 1) = "/" And Left$(path, 1) = "/" Then
        JoinUrl = base & Mid$(path, 2)
    ElseIf Right$(base, 1) <> "/" And Left$(path, 1) <> "/" And Len(path) > 0 Then
        JoinUrl = base & "/" & path
    Else
        JoinUrl = base & path
    End If
End Function

'==========================================================================
'  Builder factory
'==========================================================================

Public Function NewRequest() As clsHttpRequest
    Set NewRequest = New clsHttpRequest
End Function

'==========================================================================
'  Sync convenience verbs
'==========================================================================

Public Function HttpGet(ByVal url As String, _
                        Optional headers As Variant, _
                        Optional ByVal timeoutMs As Long = 0) As clsHttpResponse
    Set HttpGet = Build("GET", url, "", headers, timeoutMs).Send()
End Function

Public Function HttpDelete(ByVal url As String, _
                           Optional headers As Variant, _
                           Optional ByVal timeoutMs As Long = 0) As clsHttpResponse
    Set HttpDelete = Build("DELETE", url, "", headers, timeoutMs).Send()
End Function

Public Function HttpPost(ByVal url As String, _
                         ByVal body As String, _
                         Optional headers As Variant, _
                         Optional ByVal timeoutMs As Long = 0) As clsHttpResponse
    Set HttpPost = Build("POST", url, body, headers, timeoutMs).Send()
End Function

Public Function HttpPut(ByVal url As String, _
                        ByVal body As String, _
                        Optional headers As Variant, _
                        Optional ByVal timeoutMs As Long = 0) As clsHttpResponse
    Set HttpPut = Build("PUT", url, body, headers, timeoutMs).Send()
End Function

Public Function HttpPatch(ByVal url As String, _
                          ByVal body As String, _
                          Optional headers As Variant, _
                          Optional ByVal timeoutMs As Long = 0) As clsHttpResponse
    Set HttpPatch = Build("PATCH", url, body, headers, timeoutMs).Send()
End Function

Public Function HttpPostJson(ByVal url As String, _
                             ByVal value As Variant, _
                             Optional headers As Variant, _
                             Optional ByVal timeoutMs As Long = 0) As clsHttpResponse
    Set HttpPostJson = BuildJson("POST", url, value, headers, timeoutMs).Send()
End Function

Public Function HttpPutJson(ByVal url As String, _
                            ByVal value As Variant, _
                            Optional headers As Variant, _
                            Optional ByVal timeoutMs As Long = 0) As clsHttpResponse
    Set HttpPutJson = BuildJson("PUT", url, value, headers, timeoutMs).Send()
End Function

Public Function HttpPatchJson(ByVal url As String, _
                              ByVal value As Variant, _
                              Optional headers As Variant, _
                              Optional ByVal timeoutMs As Long = 0) As clsHttpResponse
    Set HttpPatchJson = BuildJson("PATCH", url, value, headers, timeoutMs).Send()
End Function

'==========================================================================
'  Async convenience verbs (typed IHttpCallback)
'==========================================================================

Public Function HttpGetAsync(ByVal url As String, _
                             ByVal callback As IHttpCallback, _
                             Optional ByRef tag As Variant, _
                             Optional headers As Variant) As Long
    HttpGetAsync = Build("GET", url, "", headers, 0).SendAsync(callback, tag)
End Function

Public Function HttpDeleteAsync(ByVal url As String, _
                                ByVal callback As IHttpCallback, _
                                Optional ByRef tag As Variant, _
                                Optional headers As Variant) As Long
    HttpDeleteAsync = Build("DELETE", url, "", headers, 0).SendAsync(callback, tag)
End Function

Public Function HttpPostAsync(ByVal url As String, _
                              ByVal body As String, _
                              ByVal callback As IHttpCallback, _
                              Optional ByRef tag As Variant, _
                              Optional headers As Variant) As Long
    HttpPostAsync = Build("POST", url, body, headers, 0).SendAsync(callback, tag)
End Function

Public Function HttpPutAsync(ByVal url As String, _
                             ByVal body As String, _
                             ByVal callback As IHttpCallback, _
                             Optional ByRef tag As Variant, _
                             Optional headers As Variant) As Long
    HttpPutAsync = Build("PUT", url, body, headers, 0).SendAsync(callback, tag)
End Function

Public Function HttpPatchAsync(ByVal url As String, _
                               ByVal body As String, _
                               ByVal callback As IHttpCallback, _
                               Optional ByRef tag As Variant, _
                               Optional headers As Variant) As Long
    HttpPatchAsync = Build("PATCH", url, body, headers, 0).SendAsync(callback, tag)
End Function

Public Function HttpPostJsonAsync(ByVal url As String, _
                                  ByVal value As Variant, _
                                  ByVal callback As IHttpCallback, _
                                  Optional ByRef tag As Variant, _
                                  Optional headers As Variant) As Long
    HttpPostJsonAsync = BuildJson("POST", url, value, headers, 0).SendAsync(callback, tag)
End Function

Public Function HttpPutJsonAsync(ByVal url As String, _
                                 ByVal value As Variant, _
                                 ByVal callback As IHttpCallback, _
                                 Optional ByRef tag As Variant, _
                                 Optional headers As Variant) As Long
    HttpPutJsonAsync = BuildJson("PUT", url, value, headers, 0).SendAsync(callback, tag)
End Function

Public Function HttpPatchJsonAsync(ByVal url As String, _
                                   ByVal value As Variant, _
                                   ByVal callback As IHttpCallback, _
                                   Optional ByRef tag As Variant, _
                                   Optional headers As Variant) As Long
    HttpPatchJsonAsync = BuildJson("PATCH", url, value, headers, 0).SendAsync(callback, tag)
End Function

'==========================================================================
'  Cancellation
'==========================================================================

Public Sub Cancel(ByVal id As Long)
    modRestPump.Cancel id
End Sub

Public Sub CancelAll()
    modRestPump.CancelAll
End Sub

Public Function PendingCount() As Long
    PendingCount = modRestPump.PendingCount
End Function

'==========================================================================
'  Internal builder helpers
'==========================================================================

Private Function Build(ByVal method As String, _
                       ByVal url As String, _
                       ByVal body As String, _
                       ByRef headers As Variant, _
                       ByVal timeoutMs As Long) As clsHttpRequest
    Dim req As clsHttpRequest
    Set req = New clsHttpRequest
    req.Method = method
    req.Url = url
    req.Body = body
    If timeoutMs > 0 Then req.TimeoutMs = timeoutMs
    req.MergeHeaders headers
    Set Build = req
End Function

Private Function BuildJson(ByVal method As String, _
                           ByVal url As String, _
                           ByRef value As Variant, _
                           ByRef headers As Variant, _
                           ByVal timeoutMs As Long) As clsHttpRequest
    Dim req As clsHttpRequest
    Set req = New clsHttpRequest
    req.Method = method
    req.Url = url
    req.SetJsonBody value
    If timeoutMs > 0 Then req.TimeoutMs = timeoutMs
    req.MergeHeaders headers
    Set BuildJson = req
End Function
