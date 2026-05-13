VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLateBoundReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsLateBoundReceiver - test-only fixture for the SendAsyncTo path.
'
' Used by Test_Async_All to verify that clsHttpRequest.SendAsyncTo
' (`CallByName` late-bound dispatch) finds and invokes a public method
' on a non-IHttpCallback class. Standard modules use this same path
' because they can't `Implements` an interface.

Private m_count      As Long
Private m_lastTag    As Variant
Private m_lastStatus As Long

Public Property Get Count() As Long
    Count = m_count
End Property

Public Property Get LastTag() As Variant
    If IsObject(m_lastTag) Then
        Set LastTag = m_lastTag
    Else
        LastTag = m_lastTag
    End If
End Property

Public Property Get LastStatus() As Long
    LastStatus = m_lastStatus
End Property

' Public so CallByName can find it. Signature must match what
' clsAsyncCall.Dispatch invokes on the late-bound target:
' (response As clsHttpResponse, tag As Variant).
Public Sub OnHttpResponse(ByVal response As clsHttpResponse, ByVal tag As Variant)
    m_count = m_count + 1
    If IsObject(tag) Then
        Set m_lastTag = tag
    Else
        m_lastTag = tag
    End If
    m_lastStatus = response.Status
End Sub
