VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHttpResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsHttpResponse - return value of every REST call (sync or async).
'
' Status        : HTTP status (e.g. 200, 404). 0 if the request never completed.
' StatusText    : HTTP reason phrase.
' Headers       : Scripting.Dictionary, case-insensitive header names -> raw values.
' Body          : Response body as a VBA String. Always decoded as UTF-8.
' Json          : Lazily parsed body. Dictionary / Collection / primitive.
'                 Returns Empty when the body is empty.
' IsSuccess     : 2xx status and no transport error.
' ErrorMessage  : Populated on transport errors or JSON parse failures.

Private m_status        As Long
Private m_statusText    As String
Private m_headers       As Object       ' Scripting.Dictionary
Private m_body          As String
Private m_jsonParsed    As Boolean
Private m_json          As Variant
Private m_errorMessage  As String

Private Sub Class_Initialize()
    Set m_headers = CreateObject("Scripting.Dictionary")
    m_headers.CompareMode = 1            ' TextCompare = case-insensitive
End Sub

'==========================================================================
'  Public properties
'==========================================================================

Public Property Get Status() As Long
    Status = m_status
End Property

Public Property Get StatusText() As String
    StatusText = m_statusText
End Property

Public Property Get Headers() As Object
    Set Headers = m_headers
End Property

Public Property Get Body() As String
    Body = m_body
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = m_errorMessage
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = (m_status >= 200 And m_status < 300 And Len(m_errorMessage) = 0)
End Property

' Lazy JSON parse. Returns the parsed value (Dictionary, Collection, or
' primitive) on first access; cached thereafter. Failure to parse
' populates ErrorMessage rather than raising.
Public Property Get Json() As Variant
    If Not m_jsonParsed Then
        m_jsonParsed = True
        If Len(m_body) > 0 Then
            ' modJson.JsonParse returns a Variant that may contain an
            ' object (Dictionary / Collection) or a primitive. We can't
            ' use plain `=` because Let-assigning a Variant-containing-
            ' object triggers VBA's default-member trap (Dictionary.Item
            ' / Collection.Item, both of which need an argument) and
            ' raises runtime 450/13. Try Set first; if the result is a
            ' primitive Set raises 424 "Object required" and we retry
            ' with Let.
            On Error Resume Next
            Err.Clear
            Set m_json = modJson.JsonParse(m_body)
            Select Case Err.Number
                Case 0
                    ' Got an object - already stored.
                Case 424
                    Err.Clear
                    m_json = modJson.JsonParse(m_body)
                    If Err.Number <> 0 Then
                        m_errorMessage = "JSON parse failed: " & Err.Description
                        Err.Clear
                    End If
                Case Else
                    m_errorMessage = "JSON parse failed: " & Err.Description
                    Err.Clear
            End Select
            On Error GoTo 0
        End If
    End If
    If IsObject(m_json) Then
        Set Json = m_json
    Else
        Json = m_json
    End If
End Property

'==========================================================================
'  Internal population (called by the request classes)
'==========================================================================

' Populate the response from a completed MSXML2.XMLHTTP60 or ServerXMLHTTP60.
' Reads responseBody as a byte array and decodes UTF-8 via ADODB.Stream
' (sidesteps MSXML's charset-guessing of responseText). Late-bound so
' the library doesn't require an MSXML type-library reference.
Public Sub LoadFromMSXML(ByVal http As Object)
    On Error Resume Next
    m_status = http.Status
    m_statusText = http.statusText
    ParseHeaders http.getAllResponseHeaders
    m_body = DecodeUtf8(http.responseBody)
    If Err.Number <> 0 Then
        m_errorMessage = "Failed to read response: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub SetTransportError(ByVal message As String)
    m_errorMessage = message
End Sub

'==========================================================================
'  Internals
'==========================================================================

Private Sub ParseHeaders(ByVal raw As String)
    If Len(raw) = 0 Then Exit Sub
    Dim lines() As String
    lines = Split(Replace(raw, vbCrLf, vbLf), vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = lines(i)
        If Len(line) > 0 Then
            Dim colonAt As Long
            colonAt = InStr(line, ":")
            If colonAt > 1 Then
                Dim k As String, v As String
                k = Trim$(Left$(line, colonAt - 1))
                v = Trim$(Mid$(line, colonAt + 1))
                If m_headers.Exists(k) Then
                    m_headers(k) = m_headers(k) & ", " & v
                Else
                    m_headers.Add k, v
                End If
            End If
        End If
    Next i
End Sub

' Decode a byte array (from MSXML.responseBody) as UTF-8.
' Empty body returns "".
Private Function DecodeUtf8(ByVal bytes As Variant) As String
    If IsEmpty(bytes) Then Exit Function
    On Error Resume Next
    Dim ub As Long
    ub = UBound(bytes)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    If ub < 0 Then Exit Function

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                         ' adTypeBinary
    stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2                         ' adTypeText
    stm.Charset = "utf-8"
    DecodeUtf8 = stm.ReadText
    stm.Close
End Function
