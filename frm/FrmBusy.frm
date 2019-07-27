VERSION 5.00
Begin VB.Form FrmBusy 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde um momento...."
      Height          =   195
      Left            =   930
      TabIndex        =   0
      Top             =   210
      Width           =   1755
   End
End
Attribute VB_Name = "FrmBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome          : FrmBusy
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

Private Sub Form_Activate()
  DoEvents
End Sub

Private Sub Form_Load()
  DoEvents
End Sub

Private Sub Form_Resize()
  Label1.Move (Me.ScaleWidth - Label1.Width) / 2, (Me.ScaleHeight - Label1.Height) / 2
End Sub
