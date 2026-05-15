# lib_access_restapi

VBA library for Microsoft Access 64-bit: synchronous and asynchronous REST
calls (GET / POST / PUT / PATCH / DELETE) with hand-rolled JSON
parse/stringify. No external dependencies — every COM component used
(`MSXML 6.0`, `Scripting.Dictionary`, `ADODB.Stream`) ships with Windows.

---

## Install

Target: **Access 64-bit**, VBA7. 32-bit Access will compile but is
untested.

### Scripted (recommended)

From the repo directory, with the target `.accdb` closed:

```powershell
powershell -ExecutionPolicy Bypass -File .\Install.ps1 -Database C:\path\to\app.accdb
```

The script imports every production module. Flags:

- `-IncludeTests` — also import `clsResponseCollector.cls` and
  `modRestTest.bas`.
- `-Force` — replace components that already exist in the target. By
  default existing modules are left untouched.

The installer always prints a clearly visible result on completion —
either an `ERROR:` block (database not found, file locked, COM /
Trust-Center problem, etc.) or a green
`Installation completed successfully.` banner with the database path
and module counts — and then waits for a keypress before closing. Safe
to launch by double-click; the window will not vanish on error. Exit
codes: `0` success, `1` unexpected failure, `2` bad arguments
(database / module file not found), `3` Access not installed,
`4` could not open database, `5` VBA project model not reachable.

Prereq: Trust Center → Macro Settings → **"Trust access to the VBA
project object model"** must be enabled (one-time, per-user Access
setting). No type-library references are required: MSXML, Scripting,
and ADODB are all created via `CreateObject` (late binding).

### Manual

1. Open your `.accdb` in Access, hit **Alt + F11** to open the VBE.
2. **File → Import File…** each of:
   - `modJson.bas`
   - `IHttpCallback.cls`
   - `clsHttpResponse.cls`
   - `clsAsyncCall.cls`
   - `clsHttpRequest.cls`
   - `modRestPump.bas`
   - `modRest.bas`
   - (optional, for tests) `clsResponseCollector.cls`, `modRestTest.bas`
3. No type-library references need to be added.

---

## Quick start — synchronous

```vba
Dim r As clsHttpResponse
Set r = modRest.HttpGet("https://httpbin.org/get?x=1")

If r.IsSuccess Then
    Debug.Print r.Status                      ' 200
    Debug.Print r.Json("args")("x")           ' "1"
Else
    Debug.Print "Failed: " & r.Status & " " & r.ErrorMessage
End If
```

POST a Dictionary as JSON:

```vba
Dim body As Object: Set body = CreateObject("Scripting.Dictionary")
body.Add "name", "Alice"
body.Add "age", 30

Dim r As clsHttpResponse
Set r = modRest.HttpPostJson("https://httpbin.org/post", body)
Debug.Print r.Json("json")("name")            ' "Alice"
```

Other verbs work the same: `HttpPut`, `HttpPutJson`, `HttpPatch`,
`HttpPatchJson`, `HttpDelete`. Every sync verb accepts an optional
`headers` (Dictionary) and `timeoutMs` (Long, milliseconds).

```vba
Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
h.Add "X-Trace-Id", "abc123"
Set r = modRest.HttpGet("https://httpbin.org/get", headers:=h, timeoutMs:=2000)
```

---

## Quick start — asynchronous (callback)

VBA has no function pointers; the "callback" is a method on an object
that implements `IHttpCallback`. The cleanest place to receive
completions is a form or class module.

```vba
' --- in your form ---
Implements IHttpCallback

Private Sub btnLoad_Click()
    modRest.HttpGetAsync "https://api.example.com/users/42", Me, "user-42"
End Sub

Private Sub IHttpCallback_OnResponse(ByVal response As clsHttpResponse, _
                                     ByVal tag As Variant)
    If Not response.IsSuccess Then
        MsgBox "HTTP " & response.Status & ": " & response.ErrorMessage
        Exit Sub
    End If
    Me.txtName = response.Json("name")
End Sub
```

`tag` is whatever you passed as the third argument — use it to tell
many in-flight requests apart when one object handles all of them.

`HttpGetAsync` returns a `Long` request ID. Hang onto it if you want to
cancel:

```vba
Dim id As Long
id = modRest.HttpGetAsync("https://httpbin.org/delay/10", Me)
' …later…
modRest.Cancel id
' or cancel everything:
modRest.CancelAll
```

---

## Async from a standard module

Standard modules can't `Implements IHttpCallback`. Use the builder's
`SendAsyncTo` method to point at a target object plus a method name:

```vba
' --- standard module ---
Public Sub LoadUser()
    Dim host As Object: Set host = New clsMyHost      ' any class
    With modRest.NewRequest
        .Configure "GET", "https://api.example.com/users/42"
        .SendAsyncTo host, "OnUserLoaded", "user-42"
    End With
End Sub
```

```vba
' --- clsMyHost ---
Public Sub OnUserLoaded(ByVal response As clsHttpResponse, ByVal tag As Variant)
    Debug.Print "tag=" & tag & " status=" & response.Status
End Sub
```

The signature must be exactly `(response As clsHttpResponse, tag As Variant)`.

---

## How async actually works (so the design makes sense)

1. You call `modRest.HttpGetAsync url, Me`. The facade:
   - Builds a `clsHttpRequest`, fills in method / url / headers.
   - Calls `req.SendAsync` which:
     - Creates a `clsAsyncCall` worker.
     - **Registers it in `modRestPump`** (a module-level `Dictionary`).
       This anchor keeps the worker alive — without it the object would
       be released the moment `SendAsync` returns and the request would
       die mid-flight.
     - Creates `MSXML2.XMLHTTP60`, opens with `async=True`, sends, and
       asks `modRestPump.EnsurePump` to start the polling timer.
   - Returns the request id. Your sub returns immediately.
2. `modRestPump` runs a Win32 `SetTimer` (~25ms cadence). Its
   `TimerProc` fires from the normal Access message pump, iterates the
   registry, and calls `PollAndMaybeDispatch` on each pending call.
3. When MSXML reaches `readyState = 4` (complete), the next poll tick:
   - Builds a `clsHttpResponse` (UTF-8 body via `ADODB.Stream`,
     parsed headers, lazy JSON).
   - Removes the call from the registry **before** dispatching, so a
     callback that pumps the message queue can't see the half-cleaned
     state.
   - Releases the MSXML object.
   - Calls your `IHttpCallback_OnResponse` (or `CallByName` fallback
     when `SendAsyncTo` was used).
4. When the registry is empty the timer stops itself.

VBA is single-threaded and cooperative — `WM_TIMER` (and therefore
completion) only fires when the message pump runs (`DoEvents`, modal
dialogs, `MsgBox`, or natural idle between procedures). A sync
`HttpGet` blocks all pending async completions until it returns.

Why polling and not `WithEvents` / `onreadystatechange`? Four COM-sink
approaches were tried and all failed on the target Access build — see
`async_architecture.md` for the autopsy.

---

## Custom headers, auth, base URL

Module-level defaults on `modRest` apply to every call (sync and async)
unless overridden per-request:

```vba
modRest.BaseUrl          = "https://api.example.com"
modRest.DefaultTimeoutMs = 10000
modRest.AddDefaultHeader "Accept", "application/json"
modRest.AddDefaultHeader "Authorization", "Bearer " & token

' Now relative paths work:
Dim r As clsHttpResponse
Set r = modRest.HttpGet("/users/42")
```

Per-call headers (override or extend defaults):

```vba
Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
h.Add "X-Request-Id", "abc123"
h.Add "Authorization", "Bearer " & otherToken

Set r = modRest.HttpGet("/users/42", h)
```

For anything the verb facade doesn't cover, drop down to the builder:

```vba
With modRest.NewRequest
    .Method = "HEAD"
    .Url = "/users/42"
    .TimeoutMs = 2000
    .SetHeader "If-None-Match", etag
    Set r = .Send
End With
```

### Auth examples

```vba
' Bearer:
modRest.AddDefaultHeader "Authorization", "Bearer " & token

' Basic:
Dim creds As String
creds = Base64Encode(user & ":" & pass)         ' bring your own helper
modRest.AddDefaultHeader "Authorization", "Basic " & creds

' API key in custom header:
modRest.AddDefaultHeader "X-API-Key", apiKey
```

---

## JSON parsing

`modJson.JsonParse(text [, numbersAsStrings])` returns:

| JSON       | VBA                                         |
|------------|---------------------------------------------|
| object     | `Scripting.Dictionary`                      |
| array      | `VBA.Collection` (1-based)                  |
| string     | `String`                                    |
| number     | `Long` if integral and fits, else `Double`  |
| true/false | `Boolean`                                   |
| null       | `Null` (VBA `Null`, not `Nothing`)          |

```vba
Dim v As Variant
v = modJson.JsonParse("{""items"":[{""id"":1,""ok"":true}],""next"":null}")
Debug.Print v("items")(1)("id")                ' 1
Debug.Print v("items")(1)("ok")                ' True
Debug.Print IsNull(v("next"))                  ' True
```

Watch out for `Double` precision on integer IDs over 2^53. If you
deal with those, pass `True` for the second argument and the parser
returns numbers as their literal string:

```vba
v = modJson.JsonParse("{""id"":9007199254740993}", True)
Debug.Print v("id")                            ' "9007199254740993"
```

### Set vs `=` when the result is an object

When the JSON top-level (or a property you assign into a Variant) is an
**object or array**, you must use `Set`:

```vba
Dim parsed As Variant
Set parsed = modJson.JsonParse("[1,2,3]")        ' array -> Collection
Set parsed = modJson.JsonParse("{""a"":1}")      ' object -> Dictionary
```

For primitives (string / number / boolean / null) use plain `=`:

```vba
parsed = modJson.JsonParse("42")                 ' Long
parsed = modJson.JsonParse("""hello""")          ' String
parsed = modJson.JsonParse("null")               ' Null
```

Reason: when VBA's `Let` (`=`) assignment sees a Variant containing an
object, it tries to call the object's *default member*. `Collection`'s
default is `Item(Index)` and `Dictionary`'s is `Item(Key)` — both
require an argument, so a bare `parsed = modJson.JsonParse("[]")`
raises **Runtime error 450, "Wrong number of arguments"**. Using `Set`
sidesteps the default-member call.

`modJson.JsonStringify(value [, pretty])` reverses it:

```vba
Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
d.Add "x", 1
d.Add "msg", "he said ""hi"""
Debug.Print modJson.JsonStringify(d)
' {"x":1,"msg":"he said \"hi\""}

Debug.Print modJson.JsonStringify(d, True)
' (pretty-printed, tab-indented)
```

---

## Errors

Every call returns a `clsHttpResponse`. Transport errors don't raise —
they populate `.ErrorMessage` and leave `.Status` at 0. HTTP non-2xx
responses set `.Status` and `.IsSuccess = False` but no error message.

```vba
Dim r As clsHttpResponse
Set r = modRest.HttpGet("https://no.such.host.example/")
Debug.Print r.Status                            ' 0
Debug.Print r.ErrorMessage                      ' "Transport error: ..."
```

JSON parse failures (lazy, on first `.Json` access) likewise populate
`.ErrorMessage` rather than raising.

---

## Limitations (known, by design)

- **Access 64-bit only**, VBA7. `PtrSafe` declares in `modRestPump`
  require VBA7 (Office 2010+).
- **No gzip**. The library sends `Accept-Encoding: identity` so MSXML
  doesn't get a compressed body it can't decode.
- **No streaming**. Whole response is buffered. Fine for JSON,
  unsuitable for large file downloads.
- **No proxy auto-config for sync**. `ServerXMLHTTP` uses WinHTTP
  proxy settings (`netsh winhttp set proxy …`), not IE settings.
- **64-bit JSON integer IDs lose precision** unless you pass
  `numbersAsStrings:=True` to `JsonParse`.
- **No automatic retries / no rate limiting.** Both belong above this
  library; add them per endpoint as needed.
- **`End` in user code** kills in-flight async requests without
  cleanup. Unavoidable in VBA. Use `modRest.CancelAll` on `Form_Unload`.

---

## Running the tests

`modRestTest.Test_All` runs five suites:

- **Test_Json_All** (offline) — nulls, bools, numbers (incl.
  `numbersAsStrings`), strings with escapes and surrogate pairs, empty
  `{}` / `[]`, duplicate keys (last-wins), nested structures, stringify
  round-trips (incl. pretty-print and cycle detection), malformed input
  raising.
- **Test_Request_All** (offline) — `clsHttpRequest` builder mechanics:
  defaults, property setters, `Configure` chain, case-insensitive
  `SetHeader`, `MergeHeaders` with dict / `Nothing` / `Missing`,
  `SetJsonBody` serialize + Content-Type behaviour.
- **Test_Defaults_All** (offline) — `modRest.JoinUrl` permutations,
  `EffectiveUrl` with `BaseUrl`, per-request `TimeoutMs` override,
  `EffectiveHeaders` merge semantics, default-header lifecycle.
- **Test_Sync_All** (online) — GET / POST-JSON / PUT / PATCH / DELETE
  against `httpbin.org`, UTF-8 body fetch, `IsSuccess` true/false, 404
  vs transport-error paths, `BaseUrl` + relative path end-to-end,
  default-header server echo.
- **Test_Async_All** (online) — three concurrent `/delay/N` requests
  via `clsResponseCollector`, cancel test, plus a late-bound
  `SendAsyncTo` test via `clsLateBoundReceiver`.

### Manual run from the VBE

```vba
modRestTest.Test_All
```

Each assertion prints PASS/FAIL with a tally at the end. Expect 102
passes if everything is healthy and the network is reachable.

### Headless run (`Run-Tests.ps1`)

For automated runs on a Windows machine with Access installed, use the
included PowerShell script. From the repo directory:

```powershell
powershell -ExecutionPolicy Bypass -File .\Run-Tests.ps1
```

It creates a throwaway `.accdb` next to the script, imports every
module, adds the MSXML 6.0 reference, runs the test suite with output
mirrored to a log file, prints PASS/FAIL with colour, and exits with
the failure count as the exit code.

Flags:

- `-JsonOnly` — only the offline JSON tests (no network required).
- `-KeepDatabase` — keep the temp `.accdb` and log on disk; the paths
  are printed at the end. Useful for poking at a failed run in the VBE.

The script also keeps the `.accdb` automatically on any failure.

Prereqs the script checks for:
- Access installed (any modern version; **64-bit recommended** to match
  the library target).
- Trust Center → Macro Settings → **"Trust access to the VBA project
  object model"** enabled. Without it, the script errors out with a
  pointer to the setting.

Exit code = failure count, so it slots into CI cleanly later if you ever
want to wire it up.

---

## File reference

| File                        | Role                                            |
|-----------------------------|-------------------------------------------------|
| `modJson.bas`               | JSON parse / stringify                          |
| `IHttpCallback.cls`         | Async completion interface                      |
| `clsHttpResponse.cls`       | Response value object (sync & async)            |
| `clsHttpRequest.cls`        | Request builder; `Send` / `SendAsync` / `SendAsyncTo` |
| `clsAsyncCall.cls`          | Internal in-flight async worker                 |
| `modRestPump.bas`           | Pending-call registry + SetTimer polling pump   |
| `modRest.bas`               | Public facade — what callers normally use       |
| `clsResponseCollector.cls`  | Test-only `IHttpCallback` fixture               |
| `clsLateBoundReceiver.cls`  | Test-only late-bound (`SendAsyncTo`) fixture    |
| `modRestTest.bas`           | Test suites (offline + online)                  |
| `Install.ps1`               | Scripted installer                              |
| `Run-Tests.ps1`             | Headless PowerShell test runner                 |
| `async_architecture.md`     | Polling design + history of failed event sinks  |
