VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSON_RoutineClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome          : JSON_RoutineClass
'> Data e Hora   : 27/07/2019 11:08
'> Autor         : Diógenes Dias de Souza Júnior                          <
'> Descrição     :
'> Modificada em : 27/07/2019 11:08
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

Option Explicit

Private Const INVALID_JSON As Long = 1
Private Const INVALID_OBJECT As Long = 2
Private Const INVALID_ARRAY As Long = 3
Private Const INVALID_BOOLEAN As Long = 4
Private Const INVALID_NULL As Long = 5
Private Const INVALID_KEY As Long = 6
Private Const INVALID_RPC_CALL As Long = 7

Private Sub Class_Initialize()
  psErrors = ""
End Sub

Public Function ClearParserErrors() As String
  psErrors = ""
End Function

Private Function Encode(STR) As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim SB       As New StringBuilder
  Dim i        As Long
  Dim j        As Long
  Dim aL1      As Variant
  Dim aL2      As Variant
  Dim c        As String
  Dim p        As Boolean
  Dim a        As Variant

  aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
  aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
  For i = 1 To Len(STR)
    p = True
    c = Mid(STR, i, 1)
    For j = 0 To 7
      If c = Chr(aL1(j)) Then
        SB.Append "\" & Chr(aL2(j))
        p = False
        Exit For
      End If
    Next

    If p Then
      a = AscW(c)
      If a > 31 And a < 127 Then
        SB.Append c
      ElseIf a > -1 Or a < 65535 Then
        SB.Append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
      End If
    End If
  Next

  Encode = SB.toString
  Set SB = Nothing

End Function

Public Function GetParserErrors() As String
  GetParserErrors = psErrors
End Function

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim iDU      As Long
  Dim iDL      As Long
  Dim i        As Long
  Dim SB       As New StringBuilder
  Dim sPB1, sPB2   ' String PointBuffer1, String PointBuffer2
  '
  On Error Resume Next
  '
  iDL = LBound(aBD, iBC)
  iDU = UBound(aBD, iBC)

  If Err.Number = 9 Then
    sPB1 = sPT & sPS
    For i = 1 To Len(sPB1)
      If i <> 1 Then sPB2 = sPB2 & ","
      sPB2 = sPB2 & Mid(sPB1, i, 1)
    Next
    '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
    SB.Append toString(aBD(sPB2))
  Else
    sPT = sPT & sPS
    SB.Append "["
    For i = iDL To iDU
      SB.Append multiArray(aBD, iBC + 1, i, sPT)
      If i < iDU Then SB.Append ","
    Next
    SB.Append "]"
    sPT = Left(sPT, iBC - 2)
  End If
  Err.Clear
  multiArray = SB.toString

  Set SB = Nothing
End Function

'
'   parse string and create JSON object
'
Public Function parse(ByRef STR As String) As Object
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim Index    As Long
  '
  Index = 1
  psErrors = ""
  On Error Resume Next
  Call skipChar(STR, Index)
  Select Case Mid(STR, Index, 1)
    Case "{"
      Set parse = parseObject(STR, Index)
    Case "["
      Set parse = parseArray(STR, Index)
    Case Else
      psErrors = "Invalid JSON"
      Debug.Print psErrors
  End Select


End Function

'
'   parse list
'
Private Function parseArray(ByRef STR As String, ByRef Index As Long) As Collection

  Set parseArray = New Collection

  ' "["
  Call skipChar(STR, Index)
  If Mid(STR, Index, 1) <> "[" Then
    psErrors = psErrors & "Invalid Array at position " & Index & " : " + Mid(STR, Index, 20) & vbCrLf
    Debug.Print psErrors
    Exit Function
  End If

  Index = Index + 1

  Do

    Call skipChar(STR, Index)
    If "]" = Mid(STR, Index, 1) Then
      Index = Index + 1
      Exit Do
    ElseIf "," = Mid(STR, Index, 1) Then
      Index = Index + 1
      Call skipChar(STR, Index)
    ElseIf Index > Len(STR) Then
      psErrors = psErrors & "Missing ']': " & Right(STR, 20) & vbCrLf
      Debug.Print psErrors
      Exit Do
    End If

    ' add value
    On Error Resume Next
    parseArray.Add parseValue(STR, Index)
    If Err.Number <> 0 Then
      psErrors = psErrors & Err.Description & ": " & Mid(STR, Index, 20) & vbCrLf
      Debug.Print psErrors
      Exit Do
    End If
  Loop

End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef STR As String, ByRef Index As Long) As Boolean

  Call skipChar(STR, Index)
  If Mid(STR, Index, 4) = "true" Then
    parseBoolean = True
    Index = Index + 4
  ElseIf Mid(STR, Index, 5) = "false" Then
    parseBoolean = False
    Index = Index + 5
  Else
    psErrors = psErrors & "Invalid Boolean at position " & Index & " : " & Mid(STR, Index) & vbCrLf
    Debug.Print psErrors
  End If

End Function

Private Function parseKey(ByRef STR As String, ByRef Index As Long) As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim dquote   As Boolean
  Dim squote   As Boolean
  Dim Char     As String

  Call skipChar(STR, Index)
  Do While Index > 0 And Index <= Len(STR)
    Char = Mid(STR, Index, 1)
    Select Case (Char)
      Case """"
        dquote = Not dquote
        Index = Index + 1
        If Not dquote Then
          Call skipChar(STR, Index)
          If Mid(STR, Index, 1) <> ":" Then
            psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
            Debug.Print psErrors
            Exit Do
          End If
        End If
      Case "'"
        squote = Not squote
        Index = Index + 1
        If Not squote Then
          Call skipChar(STR, Index)
          If Mid(STR, Index, 1) <> ":" Then
            psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
            Debug.Print psErrors
            Exit Do
          End If
        End If
      Case ":"
        Index = Index + 1
        If Not dquote And Not squote Then
          Exit Do
        Else
          parseKey = parseKey & Char
        End If
      Case Else
        If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
        Else
          parseKey = parseKey & Char
        End If
        Index = Index + 1
    End Select
  Loop

End Function

'
'   parse null
'
Private Function parseNull(ByRef STR As String, ByRef Index As Long)

  Call skipChar(STR, Index)
  If Mid(STR, Index, 4) = "null" Then
    parseNull = Null
    Index = Index + 4
  Else
    psErrors = psErrors & "Invalid null value at position " & Index & " : " & Mid(STR, Index) & vbCrLf
    Debug.Print psErrors
  End If

End Function

'
'   parse number
'
Private Function parseNumber(ByRef STR As String, ByRef Index As Long)
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim Value    As String
  Dim Char     As String

  Call skipChar(STR, Index)
  Do While Index > 0 And Index <= Len(STR)
    Char = Mid(STR, Index, 1)
    If InStr("+-0123456789.eE", Char) Then
      Value = Value & Char
      Index = Index + 1
    Else
      parseNumber = Value  ' CDec(Value)
      Exit Function
    End If
  Loop
End Function

'
'   parse collection of key/value
'
Private Function parseObject(ByRef STR As String, ByRef Index As Long) As Object
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim sKey     As String
  '
  Set parseObject = CreateObject("Scripting.Dictionary")
  ' "{"
  Call skipChar(STR, Index)
  If Mid(STR, Index, 1) <> "{" Then
    psErrors = psErrors & "Invalid Object at position " & Index & " : " & Mid(STR, Index) & vbCrLf
    Debug.Print psErrors
    Exit Function
  End If

  Index = Index + 1

  Do
    Call skipChar(STR, Index)
    If "}" = Mid(STR, Index, 1) Then
      Index = Index + 1
      Exit Do
    ElseIf "," = Mid(STR, Index, 1) Then
      Index = Index + 1
      Call skipChar(STR, Index)
    ElseIf Index > Len(STR) Then
      psErrors = psErrors & "Missing '}': " & Right(STR, 20) & vbCrLf
      Debug.Print psErrors
      Exit Do
    End If


    ' add key/value pair
    sKey = LCase(parseKey(STR, Index))
    On Error Resume Next

    parseObject.Add sKey, parseValue(STR, Index)
    If Err.Number <> 0 Then
      psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
      Debug.Print psErrors
      Exit Do
    End If
  Loop
eh:

End Function

'
'   parse string
'
Private Function parseString(ByRef STR As String, ByRef Index As Long) As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim quote    As String
  Dim Char     As String
  Dim Code     As String
  Dim SB       As New StringBuilder

  Call skipChar(STR, Index)
  quote = Mid(STR, Index, 1)
  Index = Index + 1

  Do While Index > 0 And Index <= Len(STR)
    Char = Mid(STR, Index, 1)
    Select Case (Char)
      Case "\"
        Index = Index + 1
        Char = Mid(STR, Index, 1)
        Select Case (Char)
          Case """", "\", "/", "'"
            SB.Append Char
            Index = Index + 1
          Case "b"
            SB.Append vbBack
            Index = Index + 1
          Case "f"
            SB.Append vbFormFeed
            Index = Index + 1
          Case "n"
            SB.Append vbLf
            Index = Index + 1
          Case "r"
            SB.Append vbCr
            Index = Index + 1
          Case "t"
            SB.Append vbTab
            Index = Index + 1
          Case "u"
            Index = Index + 1
            Code = Mid(STR, Index, 4)
            SB.Append ChrW(Val("&h" + Code))
            Index = Index + 4
        End Select
      Case quote
        Index = Index + 1

        parseString = SB.toString
        Set SB = Nothing

        Exit Function

      Case Else
        SB.Append Char
        Index = Index + 1
    End Select
  Loop

  parseString = SB.toString
  Set SB = Nothing

End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef STR As String, ByRef Index As Long)

  Call skipChar(STR, Index)

  Select Case Mid(STR, Index, 1)
    Case "{"
      Set parseValue = parseObject(STR, Index)
    Case "["
      Set parseValue = parseArray(STR, Index)
    Case """", "'"
      parseValue = parseString(STR, Index)
    Case "t", "f"
      parseValue = parseBoolean(STR, Index)
    Case "n"
      parseValue = parseNull(STR, Index)
    Case Else
      parseValue = parseNumber(STR, Index)
  End Select

End Function

Public Function RStoJSON(Rs As Object) As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim sFlds    As String
  Dim sRecs    As New StringBuilder
  Dim lRecCnt  As Long
  Dim fld      As Object
  '
  On Error GoTo errHandler
  '
  Set fld = CreateObject("ADODB.Field")
  'ADODB.Recordset
  lRecCnt = 0
  If Rs.State = adStateClosed Then
    RStoJSON = "null"
  Else
    If Rs.EOF Or Rs.BOF Then
      RStoJSON = "null"
    Else
      Do While Not Rs.EOF And Not Rs.BOF
        lRecCnt = lRecCnt + 1
        sFlds = ""
        For Each fld In Rs.Fields
          sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.Value & "") & """")
        Next     'fld
        sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
        Rs.MoveNext
      Loop
      RStoJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
    End If
  End If

  Exit Function
errHandler:

End Function

'
'   skip special character
'
Private Sub skipChar(ByRef STR As String, ByRef Index As Long)
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim bComment As Boolean
  Dim bStartComment As Boolean
  Dim bLongComment As Boolean
  '
  Do While Index > 0 And Index <= Len(STR)
    Select Case Mid(STR, Index, 1)
      Case vbCr, vbLf
        If Not bLongComment Then
          bStartComment = False
          bComment = False
        End If

      Case vbTab, " ", "(", ")"

      Case "/"
        If Not bLongComment Then
          If bStartComment Then
            bStartComment = False
            bComment = True
          Else
            bStartComment = True
            bComment = False
            bLongComment = False
          End If
        Else
          If bStartComment Then
            bLongComment = False
            bStartComment = False
            bComment = False
          End If
        End If

      Case "*"
        If bStartComment Then
          bStartComment = False
          bComment = True
          bLongComment = True
        Else
          bStartComment = True
        End If

      Case Else
        If Not bComment Then
          Exit Do
        End If
    End Select

    Index = Index + 1
  Loop

End Sub

' Miscellaneous JSON functions

Public Function StringToJSON(st As String) As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Const FIELD_SEP = "~"
  Const RECORD_SEP = "|"
  '
  Dim sFlds    As String
  Dim sRecs    As New StringBuilder
  Dim lRecCnt  As Long
  Dim lFld     As Long
  Dim fld      As Variant
  Dim Rows     As Variant

  lRecCnt = 0
  If st = "" Then
    StringToJSON = "null"
  Else
    Rows = Split(st, RECORD_SEP)
    For lRecCnt = LBound(Rows) To UBound(Rows)
      sFlds = ""
      fld = Split(Rows(lRecCnt), FIELD_SEP)
      For lFld = LBound(fld) To UBound(fld) Step 2
        sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
      Next    'fld
      sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
    Next   'rec
    StringToJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
  End If
End Function

Public Function toString(ByRef Obj As Variant) As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim SB       As New StringBuilder
  Dim bFI      As Boolean
  Dim i        As Long
  Dim keys     As Variant
  Dim Key      As Variant
  Dim Value    As Variant
  Dim sEB      As Variant
  '
  Select Case VarType(Obj)
    Case vbNull
      SB.Append "null"
    Case vbDate
      SB.Append """" & CStr(Obj) & """"
    Case vbString
      SB.Append """" & Encode(Obj) & """"
    Case vbObject
      bFI = True
      '
      If TypeName(Obj) = "Dictionary" Then
        SB.Append "{"
        keys = Obj.keys
        For i = 0 To Obj.Count - 1
          If bFI Then bFI = False Else SB.Append ","
          Key = keys(i)
          SB.Append """" & Key & """:" & toString(Obj.Item(Key))
        Next i
        SB.Append "}"

      ElseIf TypeName(Obj) = "Collection" Then

        SB.Append "["
        For Each Value In Obj
          If bFI Then bFI = False Else SB.Append ","
          SB.Append toString(Value)
        Next Value
        SB.Append "]"

      End If
    Case vbBoolean
      If Obj Then SB.Append "true" Else SB.Append "false"
    Case vbVariant, vbArray, vbArray + vbVariant
      SB.Append multiArray(Obj, 1, "", sEB)
    Case Else
      SB.Append Replace(Obj, ",", ".")
  End Select

  toString = SB.toString
  Set SB = Nothing

End Function

Public Function toUnicode(STR As String) As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim x        As Long
  Dim uStr     As New StringBuilder
  Dim uChrCode As Integer

  For x = 1 To Len(STR)
    uChrCode = Asc(Mid(STR, x, 1))
    Select Case uChrCode
      Case 8:  ' backspace
        uStr.Append "\b"
      Case 9:  ' tab
        uStr.Append "\t"
      Case 10:  ' line feed
        uStr.Append "\n"
      Case 12:  ' formfeed
        uStr.Append "\f"
      Case 13:  ' carriage return
        uStr.Append "\r"
      Case 34:  ' quote
        uStr.Append "\"""
      Case 39:  ' apostrophe
        uStr.Append "\'"
      Case 92:  ' backslash
        uStr.Append "\\"
      Case 123, 125:  ' "{" and "}"
        uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
      Case Is < 32, Is > 127:  ' non-ascii characters
        uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
      Case Else
        uStr.Append Chr$(uChrCode)
    End Select
  Next
  toUnicode = uStr.toString
  Exit Function

End Function
