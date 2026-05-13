VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResponseCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsResponseCollector - test-only IHttpCallback that records every
' completion. Used by modRestTest to await N concurrent async requests.
'
'   Count    : how many completions have arrived
'   Results  : Dictionary keyed by tag -> clsHttpResponse

Implements IHttpCallback

Private m_count    As Long
Private m_results  As Object             ' Dictionary keyed by tag -> response

Private Sub Class_Initialize()
    Set m_results = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get Count() As Long
    Count = m_count
End Property

Public Property Get Results() As Object
    Set Results = m_results
End Property

Private Sub IHttpCallback_OnResponse(ByVal response As clsHttpResponse, _
                                     ByVal tag As Variant)
    m_count = m_count + 1
    Dim key As String
    If IsObject(tag) Then
        key = "obj#" & ObjPtr(tag)
    Else
        key = CStr(tag)
    End If
    If m_results.Exists(key) Then m_results.Remove key
    m_results.Add key, response
End Sub
