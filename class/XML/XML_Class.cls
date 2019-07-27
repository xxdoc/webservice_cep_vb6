VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XML_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome          : XML_Class
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
Option Explicit
Public objHTTPRequest As Variant
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : GetVersionXML
'> Data/Hora             : 19/01/2010 17:51
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 19/01/2010 17:51
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Function GetVersionXML() As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '> VERIFICO A VERSÃO DO XML INSTALADA NO PC <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '
  '
  Dim i        As Integer
  Dim Tmp      As String
  Dim Obj      As Object
  '
  GetVersionXML = ""
  '
  Err.Clear
  For i = 9 To 0 Step -1
    If i <> 0 Then
      Tmp = "Msxml2.ServerXMLHTTP." & Trim(CStr(i)) & ".0"
    Else
      Tmp = "Msxml2.ServerXMLHTTP"
    End If
    '
    On Error Resume Next
    Set Obj = CreateObject(Tmp)
    '
    If Err.Number <> 0 Then
      Err.Clear
    Else
      Exit For  'Achei a versão
    End If
    '
  Next
  '
  GetVersionXML = Tmp
  '
End Function

Private Sub Class_Initialize()
  Set objHTTPRequest = CreateObject(GetVersionXML())
End Sub


Private Sub Class_Terminate()
  Set objHTTPRequest = Nothing
End Sub
