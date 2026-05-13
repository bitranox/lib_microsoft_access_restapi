# Async architecture: Win32 SetTimer + AddressOf polling

How `HttpGetAsync` and friends actually finish. Read alongside
`modRestPump.bas`, `clsHttpRequest.cls`, and `clsAsyncCall.cls`.

This isn't the design we *wanted* — events would have been cleaner —
but four event-sink approaches all failed on this Access / MSXML build
(see *History* below), so the shipping design is polling.

## Idea in one sentence

`modRestPump` runs a Win32 `SetTimer` whose `WM_TIMER` messages are
dispatched by Access's normal message pump (`DoEvents` in a wait loop
pumps it too). The `TimerProc` iterates the in-flight registry, polls
each call's `readyState`, and dispatches the ones at state 4.

## Public-surface impact

None. `SendAsync` returns immediately. Tests' existing
`Do While ... DoEvents Loop` waits keep working — `DoEvents` pumps
`WM_TIMER` along with everything else. No explicit pump call is
required from caller code.

## How it's wired

### modRestPump.bas

```vba
Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long

Private Const POLL_INTERVAL_MS As Long = 25
Private g_timerId As LongPtr

Public Sub EnsurePump()
    If g_timerId <> 0 Then Exit Sub
    g_timerId = SetTimer(0&, 0&, POLL_INTERVAL_MS, AddressOf TimerProc)
End Sub

Public Sub TimerProc(ByVal hWnd As LongPtr, _
                     ByVal uMsg As Long, _
                     ByVal idEvent As LongPtr, _
                     ByVal dwTime As Long)
    On Error Resume Next                       ' MUST not let an error escape
    ' ... iterate g_pending, call asyncCall.PollAndMaybeDispatch on each ...
    ' If g_pending empty, KillTimer + g_timerId = 0.
End Sub
```

Important VBA-isms baked into the file:

- `Declare` statements MUST come at the top of the module, before any
  `Sub`/`Function`. Putting them anywhere else is a compile error.
- `TimerProc` MUST be `Public` so `AddressOf` can resolve it.
- `On Error Resume Next` at the top of `TimerProc` is load-bearing —
  any unhandled VBA error inside a Win32 callback crashes Access.

### clsHttpRequest.cls and clsAsyncCall.cls

`clsHttpRequest` is the user-facing builder; `clsAsyncCall` is the
internal in-flight worker.

- `clsHttpRequest.SendAsync(callback, [tag])` constructs a
  `clsAsyncCall`, registers it with `modRestPump`, and tells it to
  start. It returns the request id straight away.
- `clsAsyncCall.Start*` creates an `MSXML2.XMLHTTP60`, opens it
  with `async=True`, applies headers, sends, and calls
  `modRestPump.EnsurePump` to make sure the polling timer is running.
- `clsAsyncCall.PollAndMaybeDispatch` is what the timer calls on each
  tick. When `readyState = 4` it loads the response, removes itself
  from the registry, releases MSXML, and fires the user callback.

A few VBA landmines stepped on while wiring this up:

- A method named `Release` raises a duplicate-symbol compile error in
  some Access builds (collides with COM `IUnknown::Release`). The
  local helper is named `ReleaseHttp`.
- `Call` is a reserved VBA keyword (the `Call procName` statement);
  it can't be used as a parameter or local-variable name. The
  registry's local is named `asyncCall`.
- VBA forbids the colon-separated `If c Then a: b` form *and* multi-
  statement `Dim x: x = ...` patterns in some configurations. Use
  plain multi-line `If/End If` blocks and put `Dim` on its own line.

## Tradeoffs

| Concern | What we get |
|---|---|
| Latency | Up to `POLL_INTERVAL_MS` (25ms). Invisible against real HTTP. |
| CPU at idle | Zero — timer self-stops when registry empty. |
| CPU under load | One `readyState` COM read per pending call per tick. |
| `AddressOf` risk | `TimerProc` must not let any error escape — disciplined `On Error Resume Next`. |
| Win32 dependency | Two `PtrSafe` declares; requires VBA7 (Office 2010+). |
| Cross-Access version risk | `SetTimer`/`KillTimer` are stable Win32 since forever. |

## History — what didn't work and why

Four event-sink mechanisms were attempted before settling on polling:

1. **`WithEvents` on `MSXML2.XMLHTTP60`** — raises "no automation
   source". MSXML's typelib doesn't declare an event interface VBA
   can sink.
2. **`VB_UserMemId = 0` directly on the async request class** —
   confused VBA's method resolution; Error 438 on unrelated calls
   into the class.
3. **`VB_UserMemId = 0` on a separate sink class (`clsXhrSink`)** —
   the wire itself failed: `Set m_http.onreadystatechange = sink`
   raised 438 on this Access/MSXML build, regardless of binding
   (late- or early-bound) or MSXML variant (XMLHTTP vs ServerXMLHTTP).
4. **Reading Status / responseBody from inside the sink callback** —
   even if (3) had compiled, MSXML raises 438 on property reads
   invoked from inside its own `onreadystatechange` callback. Reading
   from a fresh stack (which `SetTimer` naturally provides) sidesteps
   that too.

The fossil `clsXhrSink.cls` was deleted in an earlier refactor; git
history keeps the implementation if anyone ever wants to retry the
sink path on a different Access build.
