VERSION 5.00
Begin VB.Form FrmRota 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   Caption         =   "Rota"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11595
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
   ScaleHeight     =   7725
   ScaleWidth      =   11595
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
      Height          =   7515
      Left            =   90
      ScaleHeight     =   7515
      ScaleWidth      =   11385
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   11385
      Begin VB.CommandButton Command1 
         Caption         =   "Calcular Rota"
         Height          =   465
         Left            =   4950
         TabIndex        =   15
         Top             =   6930
         Width           =   2415
      End
      Begin VB.TextBox total_valor_gas 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   4740
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   6420
         Width           =   3645
      End
      Begin VB.TextBox valor_gas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   4740
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   5250
         Width           =   1695
      End
      Begin VB.TextBox distancia 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   4740
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5640
         Width           =   3645
      End
      Begin VB.TextBox tempo 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   4740
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   6030
         Width           =   3645
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         ForeColor       =   &H00004000&
         Height          =   2055
         Left            =   480
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3090
         Width           =   10575
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         ForeColor       =   &H00004000&
         Height          =   2055
         Left            =   480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   10575
      End
      Begin VB.TextBox origem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   2010
         MaxLength       =   8
         TabIndex        =   2
         Top             =   120
         Width           =   1605
      End
      Begin VB.TextBox destino 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   2010
         MaxLength       =   8
         TabIndex        =   5
         Top             =   2670
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL CUSTO COMBUSTÍVEL R$:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   5
         Left            =   1320
         TabIndex        =   14
         Top             =   6480
         Width           =   3240
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR COMBUSTÍVEL POR KM RODADO R$:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   5310
         Width           =   4200
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DISTÂNCIA EM KM:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   5760
         Width           =   4080
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TEMPO GASTO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   4
         Left            =   480
         TabIndex        =   11
         Top             =   6090
         Width           =   4080
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEP DESTINO:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2730
         Width           =   1800
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CEP ORIGEM:"
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmRota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome          : FrmRota
'> Data e Hora   : 27/07/2019 10:50
'> Autor         : Diógenes Dias de Souza Júnior                          <
'> Descrição     :
'> Modificada em : 27/07/2019 10:50
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Option Explicit
Private CMask  As New IPAGE_MaskEdit
'
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : WebService_Request
'> Data/Hora             : 27/07/2019 10:50
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:50
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Private Sub WebService_Request(ByVal m_Cep As String, ByVal Lst As ListBox)
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
  '
  LGateWay = "https://www.ipage.com.br/ws/v1/cep/" & m_Cep & "/json/2e3da304a5e311e98df5289a8be9ede8/"
  '
  With cXML
    'MÉTODO GET
    .objHTTPRequest.Open "GET", LGateWay, False
    .objHTTPRequest.setRequestHeader "Content-Type", "text/plain; charset=UTF-8"
    .objHTTPRequest.send (vbNull)
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
        If Json.Item("error") = True Then
          MsgBox "Error: " & vbNewLine & Json.Item("msg"), vbCritical
          Screen.MousePointer = vbDefault
          Unload FrmBusy
          Set cXML = Nothing
          Exit Sub
        End If
        '
        Lst.Clear
        '
        For Each Valor In Json
          Lst.AddItem Replace(UCase(Valor), "_", " ") & ": " & UCase(Json.Item(Valor))
        Next
        '
      End If
    End If
  End With
  '
Bye:
  Screen.MousePointer = vbDefault
  Unload FrmBusy
  Set cXML = Nothing
  Exit Sub
LogErrorHandler:
  Screen.MousePointer = vbDefault
  Unload FrmBusy
  '
  If Err.Number = -2147012894 Then
    MsgBox "Esta requisição está demorando mais do que o esperado, tente mais tarde!", vbExclamation, "Servidor ocupado"
    Err.Clear
    Call cXML.objHTTPRequest.abort
    Exit Sub
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
End Sub

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : Distance_Request
'> Data/Hora             : 27/07/2019 10:50
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:50
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Private Sub Distance_Request(ByVal m_origem As String, ByVal m_destino As String)
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim cXML     As New XML_Class
  Dim Ret      As String
  Dim Parameters As String
  Dim ErrorDescription As String
  Dim LGateWay As String
  Dim cJSON_Routine As New JSON_RoutineClass
  Dim Json     As Variant
  '
  On Error GoTo LogErrorHandler
  Screen.MousePointer = vbHourglass
  '
  FrmBusy.Show
  '
  LGateWay = "https://www.ipage.com.br/ws/v1/rota/" & m_origem & "+" & m_destino & "/" & valor_gas & "/2e3da304a5e311e98df5289a8be9ede8/"
  '
  With cXML
    'MÉTODO GET
    Parameters = vbNull
    .objHTTPRequest.Open "GET", LGateWay, False
    .objHTTPRequest.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    .objHTTPRequest.setRequestHeader "Content-Length", Len(Parameters)
    .objHTTPRequest.send (Parameters)
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
        If Json.Item("error") = True Then
          MsgBox "Error: " & vbNewLine & Json.Item("msg"), vbCritical
          Screen.MousePointer = vbDefault
          Unload FrmBusy
          Set cXML = Nothing
          Exit Sub
        End If
        '
        Me.distancia.Text = Json.Item("ab").Item("distance")
        Me.tempo.Text = Json.Item("ab").Item("travel")
        Me.total_valor_gas.Text = Json.Item("ab").Item("total")

      ElseIf .objHTTPRequest.Status = 404 Then  'ok
        MsgBox "Página não encontrada, verifique se o endereço ou se todos os parâmetros estão corretos!", vbCritical
      End If
    End If
  End With
  '
Bye:
  Screen.MousePointer = vbDefault
  Unload FrmBusy
  Set cXML = Nothing
  Exit Sub
LogErrorHandler:
  Screen.MousePointer = vbDefault
  Unload FrmBusy
  '
  If Err.Number = -2147012894 Then
    MsgBox "Esta requisição está demorando mais do que o esperado, tente mais tarde!", vbExclamation, "Servidor ocupado"
    Err.Clear
    Call cXML.objHTTPRequest.abort
    Exit Sub
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
End Sub

Private Sub Command1_Click()
  Command1.Enabled = False
  Call Distance_Request(origem.Text, destino.Text)
  Command1.Enabled = True
End Sub

Private Sub destino_GotFocus()
  CMask.ChangeColor destino, False
End Sub

Private Sub destino_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then valor_gas.SetFocus
  CMask.mask destino, KeyAscii, , "99999999"
End Sub

Private Sub destino_KeyUp(KeyCode As Integer, Shift As Integer)
  If (Len(destino.Text) = 0) Then
    List2.Clear
  End If
End Sub

Private Sub destino_LostFocus()
  CMask.ChangeColor destino, True
  '
  If List2.ListCount > 0 Then
    Exit Sub
  End If
  '
  If (Len(destino.Text) < 8) Then
    Exit Sub
  End If
  '
  Command1.Enabled = False
  Call WebService_Request(destino, List2)
  Command1.Enabled = True
  valor_gas.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set CMask = Nothing
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  With FrameTxtFields
    .Move (Me.Width - .Width) / 2, 0
  End With
End Sub

Private Sub origem_GotFocus()
  CMask.ChangeColor origem, False
  CMask.SelText origem
End Sub

Private Sub origem_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then destino.SetFocus
  CMask.mask origem, KeyAscii, , "99999999"
End Sub

Private Sub origem_KeyUp(KeyCode As Integer, Shift As Integer)
  If (Len(origem.Text) = 0) Then
    List1.Clear
  End If
End Sub

Private Sub origem_LostFocus()
  CMask.ChangeColor origem, True
  If List1.ListCount > 0 Then Exit Sub
  If (Len(origem.Text) < 8) Then
    Exit Sub
  End If
  Command1.Enabled = False
  Call WebService_Request(origem, List1)
  Command1.Enabled = True
  destino.SetFocus
End Sub

Private Sub valor_gas_GotFocus()
  CMask.ChangeColor valor_gas, False
End Sub

Private Sub valor_gas_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then Command1_Click
  CMask.mask valor_gas, KeyAscii, "C"  ',  "$."
End Sub

Private Sub valor_gas_LostFocus()
  If IsNumeric(valor_gas) = False Then
    valor_gas.Text = "1,00"
  End If

  CMask.ChangeColor valor_gas, True
End Sub
