Attribute VB_Name = "modRestTest"
Option Explicit

' modRestTest - test suite for the library.
'
' Suites:
'   Test_Json_All       - offline; JSON parser + stringifier
'   Test_Request_All    - offline; clsHttpRequest builder mechanics
'   Test_Defaults_All   - offline; modRest config (URL join, header merge)
'   Test_Sync_All       - online; sync HTTP against httpbin.org
'   Test_Async_All      - online; async HTTP + cancel against httpbin.org
'
' Run Test_All from the Immediate window. JSON / Request / Defaults
' suites need no network; Sync / Async hit https://httpbin.org.

Private Const HTTPBIN_BASE As String = "https://httpbin.org"

Private g_pass    As Long
Private g_fail    As Long
Private g_logPath As String                  ' optional log mirror

' Call this BEFORE Test_All to mirror output to a text file. Used by the
' PowerShell runner; harmless when unset (output then goes to Debug.Print
' only). Pass "" to disable.
Public Sub SetLogFile(ByVal filePath As String)
    g_logPath = filePath
    If Len(filePath) > 0 Then
        Dim fnum As Integer
        fnum = FreeFile
        Open filePath For Output As #fnum    ' truncate
        Close #fnum
    End If
End Sub

' Exposed for the runner: 0 means all green.
Public Function FailCount() As Long
    FailCount = g_fail
End Function

Public Sub Test_All()
    g_pass = 0
    g_fail = 0

    LogLine "=== modJson tests ==="
    Test_Json_All

    LogLine "=== clsHttpRequest builder tests (offline) ==="
    Test_Request_All

    LogLine "=== modRest defaults tests (offline) ==="
    Test_Defaults_All

    LogLine "=== sync HTTP tests (httpbin.org) ==="
    Test_Sync_All

    LogLine "=== async HTTP tests (httpbin.org) ==="
    Test_Async_All

    LogLine "============================"
    LogLine "RESULT: " & g_pass & " passed, " & g_fail & " failed"
End Sub

Private Sub LogLine(ByVal msg As String)
    Debug.Print msg
    If Len(g_logPath) > 0 Then
        Dim fnum As Integer
        fnum = FreeFile
        Open g_logPath For Append As #fnum
        Print #fnum, msg
        Close #fnum
    End If
End Sub

'==========================================================================
'  JSON unit tests
'==========================================================================

Public Sub Test_Json_All()
    Dim parsed As Variant

    parsed = modJson.JsonParse("null")
    AssertEq "null literal", Null, parsed

    parsed = modJson.JsonParse("true")
    AssertEq "true literal", True, parsed

    parsed = modJson.JsonParse("false")
    AssertEq "false literal", False, parsed

    parsed = modJson.JsonParse("42")
    AssertEq "integer", 42, parsed

    parsed = modJson.JsonParse("-3.14")
    AssertEq "negative float", -3.14, parsed

    parsed = modJson.JsonParse("1e3")
    AssertEq "scientific", 1000#, parsed

    parsed = modJson.JsonParse("""hello""")
    AssertEq "string", "hello", parsed

    parsed = modJson.JsonParse("""line1\nline2""")
    AssertEq "string with \n", "line1" & vbLf & "line2", parsed

    parsed = modJson.JsonParse("""\u00E4\u00F6\u00FC""")
    AssertEq "string with \uXXXX (umlauts)", ChrW$(228) & ChrW$(246) & ChrW$(252), parsed

    ' Surrogate pair: U+1F600 (grinning face) -> UTF-16 D83D DE00.
    ' Source kept ASCII-only (VBE imports .bas as cp1252, not UTF-8);
    ' we feed the parser the JSON \u escape form instead of a raw emoji.
    Dim emoji As String
    emoji = modJson.JsonParse("""\uD83D\uDE00""")
    AssertEq "string with surrogate pair length", 2, Len(emoji)
    AssertEq "string with surrogate pair cp1", AscW(ChrW$(&HD83D)), AscW(Mid$(emoji, 1, 1))

    ' Whitespace tolerance
    parsed = modJson.JsonParse("   42  ")
    AssertEq "leading/trailing whitespace", 42, parsed

    Set parsed = modJson.JsonParse("  {  ""k""  :  1  }  ")
    AssertEq "whitespace inside object", 1, parsed("k")

    ' For JSON values that decode to a VBA object (Collection for arrays,
    ' Scripting.Dictionary for objects) the caller MUST use Set. Otherwise
    ' VBA's Let-assignment invokes the contained object's default member
    ' (Collection.Item / Dictionary.Item), which requires an index and
    ' raises Runtime 450 "wrong number of arguments".
    Set parsed = modJson.JsonParse("[]")
    AssertEq "empty array Count", 0, parsed.Count

    Set parsed = modJson.JsonParse("{}")
    AssertEq "empty object Count", 0, parsed.Count

    Set parsed = modJson.JsonParse("[1,2,3]")
    AssertEq "array length", 3, parsed.Count
    AssertEq "array[1]", 1, parsed(1)
    AssertEq "array[3]", 3, parsed(3)

    Set parsed = modJson.JsonParse("{""a"":1,""b"":""x""}")
    AssertEq "object a", 1, parsed("a")
    AssertEq "object b", "x", parsed("b")

    Set parsed = modJson.JsonParse("{""dup"":1,""dup"":2}")
    AssertEq "duplicate keys -> last wins", 2, parsed("dup")

    Set parsed = modJson.JsonParse("{""nested"":{""inner"":[true,null,""ok""]}}")
    AssertEq "nested object/array", "ok", parsed("nested")("inner")(3)

    ' Large integers fall back to Double (or stay as literal string with
    ' numbersAsStrings=True).
    parsed = modJson.JsonParse("12345678901234")        ' > 2^31, fits in Double
    AssertEq "big int as Double", 12345678901234#, parsed

    ' Stringify round-trips
    AssertEq "stringify int", "42", modJson.JsonStringify(42)
    AssertEq "stringify string with quote", """he said \""hi\""""", modJson.JsonStringify("he said ""hi""")
    AssertEq "stringify true", "true", modJson.JsonStringify(True)
    AssertEq "stringify false", "false", modJson.JsonStringify(False)
    AssertEq "stringify null", "null", modJson.JsonStringify(Null)
    AssertEq "stringify empty array", "[]", modJson.JsonStringify(New Collection)

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "x", 1
    d.Add "y", "two"
    AssertEq "stringify dict", "{""x"":1,""y"":""two""}", modJson.JsonStringify(d)

    ' Stringify nested
    Dim outer As Object
    Set outer = CreateObject("Scripting.Dictionary")
    Dim inner As New Collection
    inner.Add 1
    inner.Add "two"
    outer.Add "items", inner
    AssertEq "stringify nested", "{""items"":[1,""two""]}", modJson.JsonStringify(outer)

    ' Pretty-print emits multi-line with tab indent
    Dim prettyOut As String
    prettyOut = modJson.JsonStringify(outer, True)
    AssertTrue "stringify pretty has newline", InStr(prettyOut, vbCrLf) > 0
    AssertTrue "stringify pretty has tab", InStr(prettyOut, vbTab) > 0

    ' Cycle detection: a Dictionary that contains itself must raise on
    ' stringify rather than recurse forever.
    Dim cyclic As Object
    Set cyclic = CreateObject("Scripting.Dictionary")
    cyclic.Add "self", cyclic
    On Error Resume Next
    Err.Clear
    Dim ignore As String
    ignore = modJson.JsonStringify(cyclic)
    AssertTrue "stringify cycle raises", Err.Number <> 0
    Err.Clear
    On Error GoTo 0

    ' numbersAsStrings preserves precision
    parsed = modJson.JsonParse("9007199254740993", True)
    AssertEq "numbersAsStrings keeps literal", "9007199254740993", parsed

    ' Malformed input raises
    On Error Resume Next
    Err.Clear
    parsed = modJson.JsonParse("{")
    AssertTrue "malformed raises", Err.Number <> 0
    Err.Clear
    On Error GoTo 0
End Sub

'==========================================================================
'  clsHttpRequest builder tests (offline)
'==========================================================================

Public Sub Test_Request_All()
    ResetDefaults

    ' Defaults
    Dim req As clsHttpRequest
    Set req = modRest.NewRequest
    AssertEq "default Method is GET", "GET", req.Method
    AssertEq "default Url is empty", "", req.Url
    AssertEq "default Body is empty", "", req.Body
    AssertEq "default TimeoutMs is 0 sentinel", 0, req.TimeoutMs

    ' Property setters
    req.Method = "POST"
    req.Url = "/users"
    req.Body = "name=Alice"
    req.TimeoutMs = 5000
    AssertEq "Method set", "POST", req.Method
    AssertEq "Url set", "/users", req.Url
    AssertEq "Body set", "name=Alice", req.Body
    AssertEq "TimeoutMs set", 5000, req.TimeoutMs

    ' Configure chains method + url, returns Me
    Dim req2 As clsHttpRequest
    Set req2 = modRest.NewRequest.Configure("PUT", "/x")
    AssertEq "Configure sets Method", "PUT", req2.Method
    AssertEq "Configure sets Url", "/x", req2.Url

    ' SetHeader replaces existing (case-insensitive)
    Dim req3 As clsHttpRequest
    Set req3 = modRest.NewRequest
    req3.SetHeader "X-Test", "first"
    req3.SetHeader "x-test", "second"
    AssertEq "SetHeader count is 1 (case-insensitive)", 1, req3.Headers.Count
    AssertEq "SetHeader latest value wins", "second", req3.Headers("X-Test")

    ' MergeHeaders with a dict
    Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
    h.Add "X-A", "1"
    h.Add "X-B", "2"
    req3.MergeHeaders h
    AssertEq "MergeHeaders adds entries", 3, req3.Headers.Count

    ' MergeHeaders with Nothing/Missing is a no-op
    Dim req4 As clsHttpRequest
    Set req4 = modRest.NewRequest
    req4.MergeHeaders Nothing
    AssertEq "MergeHeaders Nothing is no-op", 0, req4.Headers.Count
    Call CallMergeHeadersMissing(req4)
    AssertEq "MergeHeaders Missing is no-op", 0, req4.Headers.Count

    ' SetJsonBody serializes + adds Content-Type
    Dim body As Object: Set body = CreateObject("Scripting.Dictionary")
    body.Add "name", "Alice"
    Dim req5 As clsHttpRequest
    Set req5 = modRest.NewRequest
    req5.SetJsonBody body
    AssertEq "SetJsonBody serializes", "{""name"":""Alice""}", req5.Body
    AssertEq "SetJsonBody adds Content-Type", "application/json; charset=utf-8", req5.Headers("Content-Type")

    ' SetJsonBody does NOT override an existing Content-Type
    Dim req6 As clsHttpRequest
    Set req6 = modRest.NewRequest
    req6.SetHeader "Content-Type", "application/vnd.custom+json"
    req6.SetJsonBody body
    AssertEq "SetJsonBody preserves caller Content-Type", _
             "application/vnd.custom+json", req6.Headers("Content-Type")
End Sub

' Helper to exercise MergeHeaders' IsMissing branch (Optional-style call).
Private Sub CallMergeHeadersMissing(ByVal req As clsHttpRequest)
    req.MergeHeaders          ' no arg -> IsMissing path
End Sub

'==========================================================================
'  modRest defaults tests (offline)
'==========================================================================

Public Sub Test_Defaults_All()
    ResetDefaults

    ' URL join permutations
    AssertEq "JoinUrl empty base", "/path", modRest.JoinUrl("", "/path")
    AssertEq "JoinUrl absolute path wins", "https://x/y", _
             modRest.JoinUrl("https://api", "https://x/y")
    AssertEq "JoinUrl base/ + /path", "https://api/users", _
             modRest.JoinUrl("https://api/", "/users")
    AssertEq "JoinUrl base + path", "https://api/users", _
             modRest.JoinUrl("https://api", "users")
    AssertEq "JoinUrl base/ + path", "https://api/users", _
             modRest.JoinUrl("https://api/", "users")
    AssertEq "JoinUrl base + /path", "https://api/users", _
             modRest.JoinUrl("https://api", "/users")
    AssertEq "JoinUrl base + empty", "https://api", _
             modRest.JoinUrl("https://api", "")
    AssertEq "JoinUrl both empty", "", modRest.JoinUrl("", "")

    ' EffectiveUrl honours BaseUrl
    modRest.BaseUrl = "https://api.example.com"
    Dim req As clsHttpRequest
    Set req = modRest.NewRequest
    req.Url = "/users/42"
    AssertEq "EffectiveUrl joins BaseUrl + relative path", _
             "https://api.example.com/users/42", req.EffectiveUrl

    ' Absolute Url bypasses BaseUrl
    req.Url = "https://other.example.com/raw"
    AssertEq "EffectiveUrl preserves absolute URL", _
             "https://other.example.com/raw", req.EffectiveUrl

    ' EffectiveTimeoutMs: 0 -> module default, >0 -> per-request
    modRest.DefaultTimeoutMs = 7000
    Dim req2 As clsHttpRequest
    Set req2 = modRest.NewRequest
    AssertEq "EffectiveTimeoutMs defaults to module value", 7000, req2.EffectiveTimeoutMs
    req2.TimeoutMs = 1500
    AssertEq "EffectiveTimeoutMs per-request override", 1500, req2.EffectiveTimeoutMs

    ' Default headers + request-specific merge
    ResetDefaults
    modRest.AddDefaultHeader "X-Default", "default-val"
    modRest.AddDefaultHeader "Accept", "application/json"
    Dim req3 As clsHttpRequest
    Set req3 = modRest.NewRequest
    req3.SetHeader "X-Request", "req-val"
    Dim eff As Object
    Set eff = req3.EffectiveHeaders
    AssertEq "EffectiveHeaders has default header", "default-val", eff("X-Default")
    AssertEq "EffectiveHeaders has request header", "req-val", eff("X-Request")
    AssertEq "EffectiveHeaders has Accept default", "application/json", eff("Accept")
    AssertEq "EffectiveHeaders count = defaults + request", 3, eff.Count

    ' Per-request value overrides default with same key
    req3.SetHeader "X-Default", "overridden"
    Set eff = req3.EffectiveHeaders
    AssertEq "Request header overrides default", "overridden", eff("X-Default")

    ' RemoveDefaultHeader / ClearDefaultHeaders
    modRest.RemoveDefaultHeader "Accept"
    Set eff = modRest.NewRequest.EffectiveHeaders
    AssertTrue "RemoveDefaultHeader drops the key", Not eff.Exists("Accept")
    modRest.ClearDefaultHeaders
    Set eff = modRest.NewRequest.EffectiveHeaders
    AssertEq "ClearDefaultHeaders empties defaults", 0, eff.Count

    ' Leave state clean for the network suites.
    ResetDefaults
End Sub

'==========================================================================
'  Sync HTTP tests
'==========================================================================

Public Sub Test_Sync_All()
    ResetDefaults

    Dim r As clsHttpResponse

    Set r = modRest.HttpGet(HTTPBIN_BASE & "/get?x=1")
    AssertEq "sync GET status", 200, r.Status
    AssertTrue "sync GET IsSuccess", r.IsSuccess
    AssertEq "sync GET echoed arg", "1", r.Json("args")("x")
    AssertTrue "sync GET has Content-Type header", r.Headers.Exists("Content-Type")

    Dim body As Object
    Set body = CreateObject("Scripting.Dictionary")
    body.Add "hello", "world"
    Set r = modRest.HttpPostJson(HTTPBIN_BASE & "/post", body)
    AssertEq "sync POST status", 200, r.Status
    AssertEq "sync POST echoed body", "world", r.Json("json")("hello")

    Set r = modRest.HttpPut(HTTPBIN_BASE & "/put", "", _
                            HeadersOf("X-Test", "put"))
    AssertEq "sync PUT status", 200, r.Status

    Set r = modRest.HttpPatch(HTTPBIN_BASE & "/patch", "")
    AssertEq "sync PATCH status", 200, r.Status

    Set r = modRest.HttpDelete(HTTPBIN_BASE & "/delete")
    AssertEq "sync DELETE status", 200, r.Status

    ' UTF-8 path: /encoding/utf8 returns a UTF-8 HTML page with multi-byte chars.
    Set r = modRest.HttpGet(HTTPBIN_BASE & "/encoding/utf8")
    AssertEq "utf-8 fetch status", 200, r.Status
    AssertTrue "utf-8 body contains multibyte char", InStr(r.Body, ChrW$(&H2603)) > 0 _
        Or InStr(r.Body, ChrW$(&HE9)) > 0     ' snowman or e-acute; httpbin sample varies

    ' 4xx path: HTTP status surfaces, IsSuccess is False, no transport error.
    Set r = modRest.HttpGet(HTTPBIN_BASE & "/status/404")
    AssertEq "sync 404 status", 404, r.Status
    AssertTrue "sync 404 IsSuccess false", Not r.IsSuccess
    AssertEq "sync 404 has no ErrorMessage", "", r.ErrorMessage

    ' BaseUrl + relative path
    modRest.BaseUrl = HTTPBIN_BASE
    Set r = modRest.HttpGet("/get?x=2")
    AssertEq "sync BaseUrl + relative status", 200, r.Status
    AssertEq "sync BaseUrl + relative echoed arg", "2", r.Json("args")("x")
    ResetDefaults

    ' Default header reaches the server
    modRest.AddDefaultHeader "X-Default-Hdr", "yes"
    Set r = modRest.HttpGet(HTTPBIN_BASE & "/headers")
    AssertEq "sync default header status", 200, r.Status
    AssertEq "sync default header reaches server", "yes", _
             r.Json("headers")("X-Default-Hdr")
    ResetDefaults

    ' Transport error: short timeout against a black-hole address that
    ' won't accept SYN. Status=0 and ErrorMessage populated.
    Set r = modRest.HttpGet("http://10.255.255.1/", timeoutMs:=1500)
    AssertEq "sync transport-error status", 0, r.Status
    AssertTrue "sync transport-error IsSuccess false", Not r.IsSuccess
    AssertTrue "sync transport-error ErrorMessage set", Len(r.ErrorMessage) > 0
End Sub

'==========================================================================
'  Async HTTP tests
'==========================================================================

Public Sub Test_Async_All()
    ResetDefaults

    ' Three concurrent /delay/N requests with N=1..3 seconds. All three
    ' should complete; pumping DoEvents drives modRestPump's WM_TIMER
    ' which dispatches completions.
    Dim col As New clsResponseCollector
    Dim id1 As Long, id2 As Long, id3 As Long
    id1 = modRest.HttpGetAsync(HTTPBIN_BASE & "/delay/1", col, "a")
    id2 = modRest.HttpGetAsync(HTTPBIN_BASE & "/delay/2", col, "b")
    id3 = modRest.HttpGetAsync(HTTPBIN_BASE & "/delay/3", col, "c")

    Dim startedAt As Single
    startedAt = Timer
    Do While col.Count < 3
        DoEvents
        If Timer - startedAt > 15 Then Exit Do
    Loop
    AssertEq "async: 3 completions", 3, col.Count
    AssertEq "async: registry drained", 0, modRest.PendingCount
    If col.Results.Exists("a") Then
        AssertEq "async tag 'a' status", 200, col.Results("a").Status
    End If

    ' Cancel test
    Dim col2 As New clsResponseCollector
    Dim idC As Long
    idC = modRest.HttpGetAsync(HTTPBIN_BASE & "/delay/10", col2, "cancelMe")
    DoEvents
    modRest.Cancel idC
    AssertEq "cancel: registry empty", 0, modRest.PendingCount

    Dim deadline As Single
    deadline = Timer + 2
    Do While Timer < deadline
        DoEvents
    Loop
    AssertEq "cancel: no callback fired", 0, col2.Count

    ' Late-bound async via the builder: a plain class with a public
    ' method receives the response. This is the path standard modules
    ' use, since they cannot Implements an interface.
    Dim lateBound As New clsLateBoundReceiver
    Dim idL As Long
    With modRest.NewRequest
        idL = .Configure("GET", HTTPBIN_BASE & "/get?x=late") _
               .SendAsyncTo(lateBound, "OnHttpResponse", "late-tag")
    End With

    startedAt = Timer
    Do While lateBound.Count = 0
        DoEvents
        If Timer - startedAt > 10 Then Exit Do
    Loop
    AssertEq "late-bound async fires once", 1, lateBound.Count
    AssertEq "late-bound async passes tag", "late-tag", lateBound.LastTag
    AssertEq "late-bound async status", 200, lateBound.LastStatus
End Sub

'==========================================================================
'  Helpers
'==========================================================================

' Reset modRest state between test suites so default-state mutations
' from one suite don't leak into the next.
Private Sub ResetDefaults()
    modRest.BaseUrl = ""
    modRest.DefaultTimeoutMs = 30000
    modRest.ClearDefaultHeaders
    modRest.CancelAll
End Sub

Private Function HeadersOf(ByVal key As String, ByVal value As String) As Object
    Set HeadersOf = CreateObject("Scripting.Dictionary")
    HeadersOf.Add key, value
End Function

Private Sub AssertEq(ByVal label As String, expected As Variant, actual As Variant)
    Dim ok As Boolean
    If IsNull(expected) And IsNull(actual) Then
        ok = True
    ElseIf IsNull(expected) Or IsNull(actual) Then
        ok = False
    ElseIf IsObject(expected) Or IsObject(actual) Then
        ok = (ObjPtr(expected) = ObjPtr(actual))
    Else
        ok = (expected = actual)
    End If
    If ok Then
        g_pass = g_pass + 1
        LogLine "  PASS: " & label
    Else
        g_fail = g_fail + 1
        LogLine "  FAIL: " & label & "  expected=" & SafeToStr(expected) & _
                "  actual=" & SafeToStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal label As String, ByVal cond As Boolean)
    If cond Then
        g_pass = g_pass + 1
        LogLine "  PASS: " & label
    Else
        g_fail = g_fail + 1
        LogLine "  FAIL: " & label
    End If
End Sub

Private Function SafeToStr(v As Variant) As String
    If IsNull(v) Then SafeToStr = "Null": Exit Function
    If IsObject(v) Then SafeToStr = "<object>": Exit Function
    SafeToStr = CStr(v)
End Function
