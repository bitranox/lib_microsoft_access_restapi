Attribute VB_Name = "modJson"
Option Explicit
Option Compare Binary

' modJson - hand-rolled JSON parser and serializer for VBA 7 / Access 64-bit.
'
' Public API:
'   JsonParse(text, [numbersAsStrings])  -> Variant
'   JsonStringify(value, [pretty])       -> String
'
' Mapping (parse):
'   JSON object  -> Scripting.Dictionary  (late-bound via CreateObject)
'   JSON array   -> VBA.Collection         (1-based)
'   JSON string  -> String
'   JSON number  -> Long if integral and fits, otherwise Double
'   JSON true/false -> Boolean
'   JSON null    -> Null  (VBA Null, not Nothing)
'
' Limits:
'   - Max nesting depth = 100; raises a clean error rather than stack overflow.
'   - 64-bit integers lose precision past 2^53 (Double mantissa). Pass
'     numbersAsStrings:=True to keep all numbers as their literal text.
'   - Duplicate keys: last-wins.

Private Const MAX_DEPTH As Long = 100
Private Const ERR_BASE  As Long = vbObjectError + 5000

Private Type ParseState
    text             As String
    pos              As Long
    length           As Long
    numbersAsStrings As Boolean
End Type

'==========================================================================
'  PARSE
'==========================================================================

Public Function JsonParse(ByVal text As String, _
                          Optional ByVal numbersAsStrings As Boolean = False) As Variant
    Dim s As ParseState
    s.text = text
    s.length = Len(text)
    s.pos = 1
    s.numbersAsStrings = numbersAsStrings

    ' Skip UTF-8 / UTF-16 BOM if the caller fed us a string that still has it.
    If s.length >= 1 Then
        If AscW(Mid$(s.text, 1, 1)) = &HFEFF Then s.pos = 2
    End If

    SkipWs s
    Dim v As Variant
    AssignAny v, ParseValue(s, 0)
    SkipWs s
    If s.pos <= s.length Then
        Err.Raise ERR_BASE + 1, "JsonParse", _
                  "Trailing data at position " & s.pos
    End If
    AssignAny JsonParse, v
End Function

Private Function ParseValue(ByRef s As ParseState, ByVal depth As Long) As Variant
    If depth > MAX_DEPTH Then
        Err.Raise ERR_BASE + 2, "JsonParse", _
                  "Maximum nesting depth (" & MAX_DEPTH & ") exceeded"
    End If
    SkipWs s
    If s.pos > s.length Then
        Err.Raise ERR_BASE + 3, "JsonParse", "Unexpected end of input"
    End If

    Dim c As String
    c = Mid$(s.text, s.pos, 1)
    Select Case c
        Case "{":  AssignAny ParseValue, ParseObject(s, depth + 1)
        Case "[":  AssignAny ParseValue, ParseArray(s, depth + 1)
        Case """": ParseValue = ParseString(s)
        Case "t", "f": ParseValue = ParseBool(s)
        Case "n":  ParseValue = ParseNull(s)
        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            AssignAny ParseValue, ParseNumber(s)
        Case Else
            Err.Raise ERR_BASE + 4, "JsonParse", _
                      "Unexpected character '" & c & "' at position " & s.pos
    End Select
End Function

Private Function ParseObject(ByRef s As ParseState, ByVal depth As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    s.pos = s.pos + 1                      ' consume '{'
    SkipWs s
    If s.pos <= s.length Then
        If Mid$(s.text, s.pos, 1) = "}" Then
            s.pos = s.pos + 1
            Set ParseObject = d
            Exit Function
        End If
    End If

    Do
        SkipWs s
        If s.pos > s.length Then ErrJsonAt s, 5, "Unterminated object"
        If Mid$(s.text, s.pos, 1) <> """" Then ErrJsonAt s, 6, "Expected string key"
        Dim key As String
        key = ParseString(s)
        SkipWs s
        If s.pos > s.length Then ErrJsonAt s, 7, "Unterminated object"
        If Mid$(s.text, s.pos, 1) <> ":" Then ErrJsonAt s, 8, "Expected ':'"
        s.pos = s.pos + 1

        Dim v As Variant
        AssignAny v, ParseValue(s, depth)
        If d.Exists(key) Then d.Remove key  ' last-wins
        If IsObject(v) Then
            d.Add key, v
        Else
            d.Add key, v
        End If

        SkipWs s
        If s.pos > s.length Then ErrJsonAt s, 9, "Unterminated object"
        Dim nc As String
        nc = Mid$(s.text, s.pos, 1)
        s.pos = s.pos + 1
        If nc = "}" Then Exit Do
        If nc <> "," Then ErrJsonAt s, 10, "Expected ',' or '}'"
    Loop

    Set ParseObject = d
End Function

Private Function ParseArray(ByRef s As ParseState, ByVal depth As Long) As Object
    Dim col As New Collection
    s.pos = s.pos + 1                      ' consume '['
    SkipWs s
    If s.pos <= s.length Then
        If Mid$(s.text, s.pos, 1) = "]" Then
            s.pos = s.pos + 1
            Set ParseArray = col
            Exit Function
        End If
    End If

    Do
        Dim v As Variant
        AssignAny v, ParseValue(s, depth)
        If IsObject(v) Then
            col.Add v
        Else
            col.Add v
        End If

        SkipWs s
        If s.pos > s.length Then ErrJsonAt s, 11, "Unterminated array"
        Dim nc As String
        nc = Mid$(s.text, s.pos, 1)
        s.pos = s.pos + 1
        If nc = "]" Then Exit Do
        If nc <> "," Then ErrJsonAt s, 12, "Expected ',' or ']'"
    Loop

    Set ParseArray = col
End Function

Private Function ParseString(ByRef s As ParseState) As String
    ' Consumes a JSON string, including opening and closing quotes.
    s.pos = s.pos + 1                      ' consume opening "
    Dim sb As String
    Dim chunkStart As Long
    chunkStart = s.pos

    Do
        If s.pos > s.length Then
            Err.Raise ERR_BASE + 13, "JsonParse", "Unterminated string"
        End If
        Dim code As Long
        code = AscW(Mid$(s.text, s.pos, 1))

        If code = 34 Then                  ' "
            If s.pos > chunkStart Then sb = sb & Mid$(s.text, chunkStart, s.pos - chunkStart)
            s.pos = s.pos + 1
            ParseString = sb
            Exit Function
        ElseIf code = 92 Then              ' \
            If s.pos > chunkStart Then sb = sb & Mid$(s.text, chunkStart, s.pos - chunkStart)
            s.pos = s.pos + 1
            If s.pos > s.length Then
                Err.Raise ERR_BASE + 14, "JsonParse", "Unterminated escape"
            End If
            Dim esc As String
            esc = Mid$(s.text, s.pos, 1)
            Select Case esc
                Case """": sb = sb & """": s.pos = s.pos + 1
                Case "\":  sb = sb & "\":  s.pos = s.pos + 1
                Case "/":  sb = sb & "/":  s.pos = s.pos + 1
                Case "b":  sb = sb & Chr$(8):  s.pos = s.pos + 1
                Case "f":  sb = sb & Chr$(12): s.pos = s.pos + 1
                Case "n":  sb = sb & vbLf:     s.pos = s.pos + 1
                Case "r":  sb = sb & vbCr:     s.pos = s.pos + 1
                Case "t":  sb = sb & vbTab:    s.pos = s.pos + 1
                Case "u"
                    s.pos = s.pos + 1
                    If s.pos + 3 > s.length Then
                        Err.Raise ERR_BASE + 15, "JsonParse", "Truncated \uXXXX escape"
                    End If
                    Dim cu1 As Long
                    cu1 = HexQuad(s.text, s.pos)
                    s.pos = s.pos + 4
                    ' Surrogate pair?
                    If cu1 >= &HD800 And cu1 <= &HDBFF Then
                        If s.pos + 5 > s.length Then
                            Err.Raise ERR_BASE + 16, "JsonParse", _
                                      "Unpaired high surrogate"
                        End If
                        If Mid$(s.text, s.pos, 2) <> "\u" Then
                            Err.Raise ERR_BASE + 17, "JsonParse", _
                                      "Unpaired high surrogate"
                        End If
                        s.pos = s.pos + 2
                        Dim cu2 As Long
                        cu2 = HexQuad(s.text, s.pos)
                        s.pos = s.pos + 4
                        If cu2 < &HDC00 Or cu2 > &HDFFF Then
                            Err.Raise ERR_BASE + 18, "JsonParse", _
                                      "Invalid low surrogate"
                        End If
                        ' VBA String is UTF-16 internally; just join the two code units.
                        sb = sb & ChrW$(cu1) & ChrW$(cu2)
                    Else
                        sb = sb & ChrW$(cu1)
                    End If
                Case Else
                    Err.Raise ERR_BASE + 19, "JsonParse", _
                              "Unknown escape '\" & esc & "'"
            End Select
            chunkStart = s.pos
        ElseIf code < 32 Then
            Err.Raise ERR_BASE + 20, "JsonParse", _
                      "Unescaped control character (0x" & Hex$(code) & ")"
        Else
            s.pos = s.pos + 1
        End If
    Loop
End Function

Private Function HexQuad(ByRef text As String, ByVal at As Long) As Long
    Dim i As Long, ch As Long, acc As Long
    For i = 0 To 3
        ch = AscW(Mid$(text, at + i, 1))
        Select Case ch
            Case 48 To 57:  acc = acc * 16 + (ch - 48)
            Case 65 To 70:  acc = acc * 16 + (ch - 55)
            Case 97 To 102: acc = acc * 16 + (ch - 87)
            Case Else
                Err.Raise ERR_BASE + 21, "JsonParse", _
                          "Invalid hex digit in \uXXXX escape"
        End Select
    Next i
    HexQuad = acc
End Function

Private Function ParseNumber(ByRef s As ParseState) As Variant
    Dim start As Long
    start = s.pos
    Dim hasFrac As Boolean
    Dim hasExp  As Boolean
    Dim c       As String

    If Mid$(s.text, s.pos, 1) = "-" Then s.pos = s.pos + 1

    ' Integer part
    If s.pos > s.length Then ErrJsonAt s, 22, "Invalid number"
    c = Mid$(s.text, s.pos, 1)
    If c = "0" Then
        s.pos = s.pos + 1
    ElseIf c >= "1" And c <= "9" Then
        Do While s.pos <= s.length
            c = Mid$(s.text, s.pos, 1)
            If c < "0" Or c > "9" Then Exit Do
            s.pos = s.pos + 1
        Loop
    Else
        ErrJsonAt s, 23, "Invalid number"
    End If

    ' Fraction
    If s.pos <= s.length Then
        If Mid$(s.text, s.pos, 1) = "." Then
            hasFrac = True
            s.pos = s.pos + 1
            If s.pos > s.length Then ErrJsonAt s, 24, "Invalid number"
            c = Mid$(s.text, s.pos, 1)
            If c < "0" Or c > "9" Then ErrJsonAt s, 25, "Invalid number"
            Do While s.pos <= s.length
                c = Mid$(s.text, s.pos, 1)
                If c < "0" Or c > "9" Then Exit Do
                s.pos = s.pos + 1
            Loop
        End If
    End If

    ' Exponent
    If s.pos <= s.length Then
        c = Mid$(s.text, s.pos, 1)
        If c = "e" Or c = "E" Then
            hasExp = True
            s.pos = s.pos + 1
            If s.pos > s.length Then ErrJsonAt s, 26, "Invalid number"
            c = Mid$(s.text, s.pos, 1)
            If c = "+" Or c = "-" Then s.pos = s.pos + 1
            If s.pos > s.length Then ErrJsonAt s, 27, "Invalid number"
            c = Mid$(s.text, s.pos, 1)
            If c < "0" Or c > "9" Then ErrJsonAt s, 28, "Invalid number"
            Do While s.pos <= s.length
                c = Mid$(s.text, s.pos, 1)
                If c < "0" Or c > "9" Then Exit Do
                s.pos = s.pos + 1
            Loop
        End If
    End If

    Dim lit As String
    lit = Mid$(s.text, start, s.pos - start)

    If s.numbersAsStrings Then
        ParseNumber = lit
        Exit Function
    End If

    If hasFrac Or hasExp Then
        ' Val is locale-invariant: always reads '.' as the decimal separator
        ' and 'e'/'E' as the exponent. CDbl would misread "3.14" as 314 in
        ' locales where '.' is a thousands separator.
        ParseNumber = Val(lit)
    Else
        ' Try Long first; fall back to Double for out-of-range integers.
        On Error Resume Next
        Dim tmp As Long
        tmp = CLng(lit)
        If Err.Number = 0 Then
            ParseNumber = tmp
        Else
            Err.Clear
            ParseNumber = Val(lit)
        End If
        On Error GoTo 0
    End If
End Function

Private Function ParseBool(ByRef s As ParseState) As Boolean
    ' Mid$ tolerates over-reads (returns short string), so equality
    ' alone is a safe bounds check.
    If Mid$(s.text, s.pos, 4) = "true" Then
        s.pos = s.pos + 4
        ParseBool = True
        Exit Function
    End If
    If Mid$(s.text, s.pos, 5) = "false" Then
        s.pos = s.pos + 5
        ParseBool = False
        Exit Function
    End If
    ErrJsonAt s, 29, "Invalid literal (expected true/false)"
End Function

Private Function ParseNull(ByRef s As ParseState) As Variant
    If Mid$(s.text, s.pos, 4) = "null" Then
        s.pos = s.pos + 4
        ParseNull = Null
        Exit Function
    End If
    ErrJsonAt s, 30, "Invalid literal (expected null)"
End Function

Private Sub SkipWs(ByRef s As ParseState)
    Do While s.pos <= s.length
        Select Case AscW(Mid$(s.text, s.pos, 1))
            Case 32, 9, 10, 13: s.pos = s.pos + 1
            Case Else: Exit Sub
        End Select
    Loop
End Sub

Private Sub ErrJsonAt(ByRef s As ParseState, ByVal code As Long, ByVal msg As String)
    Err.Raise ERR_BASE + code, "JsonParse", msg & " at position " & s.pos
End Sub

Private Sub AssignAny(ByRef dest As Variant, ByRef src As Variant)
    If IsObject(src) Then
        Set dest = src
    Else
        dest = src
    End If
End Sub

'==========================================================================
'  STRINGIFY
'==========================================================================

Public Function JsonStringify(ByVal value As Variant, _
                              Optional ByVal pretty As Boolean = False) As String
    Dim visited As New Collection         ' tracks ObjPtr of nested objects/arrays
    JsonStringify = SerializeValue(value, pretty, 0, visited)
End Function

Private Function SerializeValue(ByVal value As Variant, _
                                ByVal pretty As Boolean, _
                                ByVal indent As Long, _
                                ByRef visited As Collection) As String
    If IsObject(value) Then
        If value Is Nothing Then
            SerializeValue = "null"
            Exit Function
        End If
        ' Cycle check
        Dim ptrKey As String
        ptrKey = "p" & CStr(ObjPtr(value))
        Dim probe As Variant
        On Error Resume Next
        probe = visited(ptrKey)
        Dim wasFound As Boolean
        wasFound = (Err.Number = 0)
        On Error GoTo 0
        If wasFound Then
            Err.Raise ERR_BASE + 40, "JsonStringify", _
                      "Circular reference detected during serialization"
        End If
        visited.Add True, ptrKey

        If TypeOf value Is Collection Then
            SerializeValue = SerializeArray(value, pretty, indent, visited)
        Else
            ' Treat as Dictionary-like (.Keys, .Items). Scripting.Dictionary
            ' is the supported case; anything else duck-typed will work too.
            SerializeValue = SerializeObject(value, pretty, indent, visited)
        End If

        visited.Remove ptrKey
        Exit Function
    End If

    Select Case VarType(value)
        Case vbNull, vbEmpty
            SerializeValue = "null"
        Case vbBoolean
            If value Then SerializeValue = "true" Else SerializeValue = "false"
        Case vbString
            SerializeValue = EncodeString(CStr(value))
        Case vbByte, vbInteger, vbLong, vbLongLong
            SerializeValue = CStr(value)
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            SerializeValue = FormatNumber(value)
        Case vbDate
            ' ISO 8601, UTC-naive (caller decides what date means).
            SerializeValue = """" & Format$(value, "yyyy-mm-dd\THh:Nn:Ss") & """"
        Case Else
            Err.Raise ERR_BASE + 41, "JsonStringify", _
                      "Unsupported value type " & VarType(value)
    End Select
End Function

Private Function SerializeObject(ByVal d As Object, _
                                 ByVal pretty As Boolean, _
                                 ByVal indent As Long, _
                                 ByRef visited As Collection) As String
    If d.Count = 0 Then
        SerializeObject = "{}"
        Exit Function
    End If
    Dim keys As Variant
    keys = d.Keys
    Dim parts() As String
    ReDim parts(0 To UBound(keys))
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim k As String
        k = CStr(keys(i))
        Dim v As Variant
        AssignAny v, d.Item(keys(i))
        parts(i) = EncodeString(k) & IIf(pretty, ": ", ":") & _
                   SerializeValue(v, pretty, indent + 1, visited)
    Next i
    SerializeObject = WrapBlock(parts, "{", "}", pretty, indent)
End Function

Private Function SerializeArray(ByVal col As Collection, _
                                ByVal pretty As Boolean, _
                                ByVal indent As Long, _
                                ByRef visited As Collection) As String
    If col.Count = 0 Then
        SerializeArray = "[]"
        Exit Function
    End If
    Dim parts() As String
    ReDim parts(0 To col.Count - 1)
    Dim i As Long
    For i = 1 To col.Count
        Dim v As Variant
        AssignAny v, col.Item(i)
        parts(i - 1) = SerializeValue(v, pretty, indent + 1, visited)
    Next i
    SerializeArray = WrapBlock(parts, "[", "]", pretty, indent)
End Function

Private Function WrapBlock(ByRef parts() As String, _
                           ByVal openTok As String, _
                           ByVal closeTok As String, _
                           ByVal pretty As Boolean, _
                           ByVal indent As Long) As String
    If pretty Then
        Dim inner As String, outer As String
        inner = String$(indent + 1, vbTab)
        outer = String$(indent, vbTab)
        WrapBlock = openTok & vbCrLf & inner & _
                    Join(parts, "," & vbCrLf & inner) & vbCrLf & outer & closeTok
    Else
        WrapBlock = openTok & Join(parts, ",") & closeTok
    End If
End Function

Private Function EncodeString(ByVal s As String) As String
    Dim sb As String
    Dim i As Long, code As Long, ch As String
    sb = """"
    Dim n As Long
    n = Len(s)
    For i = 1 To n
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 34:  sb = sb & "\"""
            Case 92:  sb = sb & "\\"
            Case 8:   sb = sb & "\b"
            Case 9:   sb = sb & "\t"
            Case 10:  sb = sb & "\n"
            Case 12:  sb = sb & "\f"
            Case 13:  sb = sb & "\r"
            Case 0 To 31
                sb = sb & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else
                sb = sb & ch
        End Select
    Next i
    EncodeString = sb & """"
End Function

Private Function FormatNumber(ByVal n As Variant) As String
    ' Locale-independent number formatting. CStr can emit ',' for decimals
    ' under German/European locales; force '.' and avoid thousand separators.
    Dim s As String
    s = Trim$(Str$(CDbl(n)))
    ' Str$ uses '.' regardless of locale; that's what we want for JSON.
    FormatNumber = s
End Function
