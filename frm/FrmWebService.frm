VERSION 5.00
Begin VB.Form FrmWebService 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   Caption         =   "PESQUISA CEP"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrameTxtFields 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5685
      Left            =   150
      ScaleHeight     =   5685
      ScaleWidth      =   11385
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   11385
      Begin VB.CommandButton Command1 
         Caption         =   "Ler CEP"
         Height          =   330
         Left            =   3660
         TabIndex        =   35
         Top             =   120
         Width           =   1305
      End
      Begin VB.TextBox faixa_de_cep 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   1980
         TabIndex        =   34
         Top             =   5250
         Width           =   9150
      End
      Begin VB.TextBox ddd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   10080
         TabIndex        =   12
         Top             =   1230
         Width           =   1065
      End
      Begin VB.TextBox tempo_percurso_veiculo 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   32
         Top             =   4560
         Width           =   9150
      End
      Begin VB.TextBox distancia_da_capital 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   30
         Top             =   3870
         Width           =   9150
      End
      Begin VB.TextBox gentilico 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   23
         Top             =   2370
         Width           =   9150
      End
      Begin VB.TextBox microrregiao 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   28
         Top             =   3150
         Width           =   9150
      End
      Begin VB.TextBox mesorregiao 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   25
         Top             =   2760
         Width           =   9150
      End
      Begin VB.TextBox gia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   9480
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox ibge 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   6900
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox complemento 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   10
         Top             =   870
         Width           =   9150
      End
      Begin VB.TextBox longitude 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   6390
         TabIndex        =   15
         Top             =   1245
         Width           =   2355
      End
      Begin VB.TextBox latitude 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   13
         Top             =   1245
         Width           =   2355
      End
      Begin VB.TextBox endereco 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   8
         Top             =   495
         Width           =   9150
      End
      Begin VB.TextBox cidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   20
         Top             =   1980
         Width           =   7980
      End
      Begin VB.TextBox bairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         TabIndex        =   18
         Top             =   1620
         Width           =   9150
      End
      Begin VB.TextBox cep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   2010
         MaxLength       =   9
         TabIndex        =   4
         Top             =   120
         Width           =   1605
      End
      Begin VB.TextBox uf 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   10590
         TabIndex        =   22
         Top             =   1980
         Width           =   570
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FAIXA DE CEP"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   17
         Left            =   2010
         TabIndex        =   33
         Top             =   4950
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DDD:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   9
         Left            =   9330
         TabIndex        =   16
         Top             =   1305
         Width           =   480
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TEMPO DE PERCURSO VEÍCULO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   16
         Left            =   2010
         TabIndex        =   31
         Top             =   4260
         Width           =   3120
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DISTÂNCIA DA CAPITAL:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   15
         Left            =   2010
         TabIndex        =   29
         Top             =   3570
         Width           =   2520
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GENTÍLICO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   14
         Left            =   120
         TabIndex        =   27
         Top             =   2430
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MICROREGIÃO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   3210
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MESORREGIÃO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   2850
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GIA:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   1
         Left            =   8880
         TabIndex        =   2
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IBGE:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   0
         Left            =   6180
         TabIndex        =   1
         Top             =   180
         Width           =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   1935
         X2              =   1935
         Y1              =   0
         Y2              =   6000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   1920
         X2              =   1920
         Y1              =   0
         Y2              =   6000
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COMPLEMENTO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   930
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LONGITUDE:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   8
         Left            =   5070
         TabIndex        =   14
         Top             =   1305
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LATITUDE:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   7
         Left            =   720
         TabIndex        =   11
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ENDEREÇO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CIDADE:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BAIRRO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   11
         Left            =   120
         TabIndex        =   17
         Top             =   1710
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CEP:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   13
         Left            =   10110
         TabIndex        =   21
         Top             =   2040
         Width           =   360
      End
   End
End
Attribute VB_Name = "FrmWebService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome          : FrmWebService
'> Data e Hora   : 27/07/2019 11:01
'> Autor         : Diógenes Dias de Souza Júnior                          <
'> Descrição     :
'> Modificada em : 27/07/2019 11:01
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Option Explicit
Private CMask  As New IPAGE_MaskEdit
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : WebService_Request
'> Data/Hora             : 27/07/2019 11:01
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 11:01
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Private Function WebService_Request()
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim cXML     As New XML_Class
  Dim Ret      As String
  Dim ErrorDescription As String
  Dim Metodo   As String
  Dim LGateWay As String
  Dim cJSON_Routine As New JSON_RoutineClass
  Dim Json     As Variant
  Dim Valor    As Variant
  '
  On Error GoTo LogErrorHandler
  Screen.MousePointer = vbHourglass
  FrmBusy.Show
  Metodo = "json"

  LGateWay = "https://www.ipage.com.br/ws/v1/cep/" & cep & "/json/2e3da304a5e311e98df5289a8be9ede8/"
  '
  With cXML
    'MÉTODO GET
    .objHTTPRequest.Open "GET", LGateWay, False
    .objHTTPRequest.setRequestHeader "Content-Type", "text/plain; charset=UTF-8"
    .objHTTPRequest.setRequestHeader "Content-Type", "application/json"
    .objHTTPRequest.send (vbNull)
    '
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    '> RETORNA O NÚMERO DE DIAS PARA A DLL EXPIRAR <
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    '
    If .objHTTPRequest.ReadyState = 4 Then
      If .objHTTPRequest.Status = 200 Then  'ok
        '
        '/\/\/\/\/\/\/\/\/\/\/\/\
        '> VERIFICO SE DEU ERRO <
        '/\/\/\/\/\/\/\/\/\/\/\/\
        '
        Screen.MousePointer = vbDefault
        '
        Ret = .objHTTPRequest.ResponseText

        Set Json = cJSON_Routine.parse(Ret)
        '
        Debug.Print Ret
        If Json.Item("error") = True Then
          MsgBox "Error: " & vbNewLine & Json.Item("msg"), vbCritical
          Screen.MousePointer = vbDefault
          Unload FrmBusy
          Set cXML = Nothing
          Exit Function
        End If
        '
        For Each Valor In Json
          Select Case Valor
            Case "cep"
              cep.Text = Json.Item(Valor)
            Case "ibge"
              ibge.Text = Json.Item(Valor)
            Case "gia"
              gia.Text = Json.Item(Valor)
            Case "logradouro2"
              endereco.Text = Json.Item(Valor)
            Case "complemento"
              complemento.Text = Json.Item(Valor)
            Case "bairro"
              bairro.Text = Json.Item(Valor)
            Case "cidade"
              cidade.Text = Json.Item(Valor)
            Case "uf"
              uf.Text = UCase(Json.Item(Valor))
            Case "latitude"
              latitude.Text = UCase(Json.Item(Valor))
            Case "longitude"
              longitude.Text = UCase(Json.Item(Valor))
            Case "mesorregiao"
              mesorregiao.Text = UCase(Json.Item(Valor))
            Case "microrregiao"
              microrregiao.Text = UCase(Json.Item(Valor))
            Case "gentilico"
              gentilico.Text = UCase(Json.Item(Valor))
            Case "distancia_da_capital"
              distancia_da_capital.Text = UCase(Json.Item(Valor))
            Case "tempo_percurso_veiculo"
              tempo_percurso_veiculo.Text = UCase(Json.Item(Valor))
            Case "ddd"
              ddd.Text = Json.Item(Valor)
            Case "faixa_de_cep"
              faixa_de_cep.Text = Json.Item(Valor)
          End Select
        Next
        '
      Else
        MsgBox .objHTTPRequest.Status & " = " & .objHTTPRequest.ResponseText
        Debug.Print .objHTTPRequest.ResponseText

      End If
    End If
  End With
  '
Bye:
  Screen.MousePointer = vbDefault
  Unload FrmBusy
  Set cXML = Nothing
  Exit Function
LogErrorHandler:
  Screen.MousePointer = vbDefault
  Unload FrmBusy
  '
  If Err.Number = -2147012894 Then
    MsgBox "Esta requisição está demorando mais do que o esperado, tente mais tarde!", vbExclamation, "Servidor ocupado"
    Err.Clear
    Call cXML.objHTTPRequest.abort
    Exit Function
  End If
  '
  Call cXML.objHTTPRequest.abort
  '
  If Len(Trim(psErrors)) = 0 Then
    psErrors = Ret & vbNewLine
  Else
    psErrors = Ret & vbNewLine & "Retorno: " & psErrors
  End If
  '
  ErrorDescription = psErrors
  Err.Source = App.EXEName & ".Command1.click"
  Debug.Print Err.Number, Err.Source, Err.Description & vbNewLine & ErrorDescription & IIf(Erl > 0, " Na linha " & Erl, "")
  Err.Raise Err.Number, Err.Source, Err.Description & vbNewLine & ErrorDescription & IIf(Erl > 0, " Na linha " & Erl, "")
End Function

Private Sub cep_GotFocus()
CMask.ChangeColor cep, False
End Sub

Private Sub cep_KeyDown(KeyCode As Integer, Shift As Integer)
  If (Len(cep.Text) = 0) Then
    cep.Text = ""
    ibge.Text = ""
    gia.Text = ""
    endereco.Text = ""
    complemento.Text = ""
    bairro.Text = ""
    cidade.Text = ""
    uf.Text = UCase("")
    latitude.Text = UCase("")
    longitude.Text = UCase("")
    mesorregiao.Text = UCase("")
    microrregiao.Text = UCase("")
    gentilico.Text = UCase("")
    distancia_da_capital.Text = UCase("")
    tempo_percurso_veiculo.Text = UCase("")
    ddd.Text = ""
    faixa_de_cep.Text = ""
  End If
End Sub

Private Sub cep_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then Command1_Click
  CMask.mask cep, KeyAscii, , "99999999"
End Sub

Private Sub Command1_Click()
  If (Len(cep.Text) < 8) Then
    MsgBox "Número do cep inválido, verifique!", vbExclamation Or vbOKOnly
    cep.SetFocus
    Exit Sub
  End If
  '
  Call WebService_Request
End Sub

