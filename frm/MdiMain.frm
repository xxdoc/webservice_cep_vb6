VERSION 5.00
Begin VB.MDIForm MdiMain 
   BackColor       =   &H80000002&
   Caption         =   "Ipage Webservice"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   11970
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuPrincipal 
      Caption         =   "Principal"
      Begin VB.Menu mnuWS 
         Caption         =   "Pesquisar Endereço por CEP"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuRota 
         Caption         =   "Calcular Rota"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu MnuJanelas 
      Caption         =   "Janelas"
      Begin VB.Menu MnuCloseAll 
         Caption         =   "Fechar todas as janelas"
      End
      Begin VB.Menu MnuMHorizontal 
         Caption         =   "Em Mosaico Horizontal"
      End
      Begin VB.Menu MnuMVertical 
         Caption         =   "Em Mosaico Vertical"
      End
      Begin VB.Menu MnuCascata 
         Caption         =   "Em Cascata"
      End
      Begin VB.Menu MnuBarra 
         Caption         =   "-"
      End
      Begin VB.Menu MnuWindowList 
         Caption         =   "Janelas Ativas"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Ajuda"
      Begin VB.Menu MnuVisite 
         Caption         =   "Visite a Ipage"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "Sobre"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome          : MdiMain
'> Data e Hora   : 27/07/2019 10:28
'> Autor         : Diógenes Dias de Souza Júnior                          <
'> Descrição     :
'> Modificada em : 27/07/2019 10:28
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Option Explicit
'
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    PopupMenu mnuPrincipal
  End If
End Sub

'
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> 1 Then
    StringLoc = ""
    mnuSair_Click
    If StringLoc <> "OK" Then Cancel = True
  End If
End Sub

Private Sub mnuConfig_Click()

End Sub

Private Sub MnuCascata_Click()
  Me.Arrange vbCascade
End Sub

Private Sub MnuCloseAll_Click()
UnloadAll
End Sub

Private Sub MnuMHorizontal_Click()
   Me.Arrange vbTileHorizontal
End Sub

Private Sub MnuMVertical_Click()
 Me.Arrange vbTileVertical
End Sub

Private Sub mnuRota_Click()
  Call SelMenu("mnuRota")
  FrmRota.Show
End Sub

Public Sub UnloadAll()
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim F        As Integer
  '
  F = Forms.Count
  Do While F > 0
    '*** Evita o fechamento do Form MDI ***
    If Forms(F - 1).Name = "MdiMain" Then GoTo LP
    Unload Forms(F - 1)
LP:
    If F = Forms.Count Then Exit Do
    F = F - 1
  Loop
End Sub

Private Sub mnuSair_Click()
  If MsgBox("Desejar finalizar este aplicativo?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Sair") = vbNo Then
    StringLoc = ""
    Exit Sub
  End If
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim F        As Object
  '
  StringLoc = "OK"
  '
  For Each F In Forms
    On Error Resume Next
    If F.MDIChild Then
      If Err.Number = False Then
        Unload F
      End If
    End If
    Err.Clear
  Next
  '
  Unload Me
  '
End Sub

Private Sub SelMenu(MenuName As String)
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim ctrl     As Control
  '
  For Each ctrl In Controls
    If (TypeOf ctrl Is Menu) Then
      If (UCase(ctrl.Name) = UCase(MenuName)) Then
        ctrl.Checked = True
      Else
        ctrl.Checked = False
      End If
    End If
  Next
  '
End Sub

Private Sub mnuSobre_Click()
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim Versao   As String
  '
  Call SelMenu("MnuSobre")
  '
  Versao = App.Major & "." & App.Minor & "." & App.Revision
  ShellAbout Me.hwnd, App.Title, App.CompanyName & " - Ver. " & Versao, ByVal 0&
End Sub

Private Sub MnuVisite_Click()
  Call SelMenu("MnuVisite")
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim Inet     As New Inet
  '
  Inet.GoWebPage
  Set Inet = Nothing
End Sub

Private Sub mnuWS_Click()
  Call SelMenu("mnuWS")
  FrmWebService.Show
End Sub
