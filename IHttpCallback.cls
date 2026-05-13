VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' IHttpCallback - the typed async completion contract.
'
' Any class or form that wants to receive REST completions implements
' this interface and provides a single method:
'
'   Implements IHttpCallback
'
'   Private Sub IHttpCallback_OnResponse(ByVal response As clsHttpResponse, _
'                                        ByVal tag As Variant)
'       ' ... handle response here ...
'   End Sub
'
' `tag` is whatever opaque value the caller passed to HttpGetAsync etc.
' Use it to correlate a response to its originating request when one
' subscriber handles many.
'
' Standard modules cannot Implements interfaces. They use the late-bound
' path instead: modRest.NewRequest.SendAsyncTo target, methodName, tag.

Public Sub OnResponse(ByVal response As clsHttpResponse, ByVal tag As Variant)
End Sub
