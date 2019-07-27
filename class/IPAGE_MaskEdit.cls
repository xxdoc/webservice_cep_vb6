VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPAGE_MaskEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private Sub TxtFields_KeyPress(Index As Integer, KeyAscii As Integer)
'  '
'  Select Case Index
'    Case 0
'      'DATA
'      Cmask.mask TxtFields(Index), KeyAscii, "D"
'    Case 1
'      'CEP
'      Cmask.mask TxtFields(Index), KeyAscii, , "99999-999"
'    Case 2
'      'MOEDA
'      Cmask.mask TxtFields(Index), KeyAscii, "C", "€$."
'    Case 3
'      'SOMENTE TEXTO
'      Cmask.mask TxtFields(Index), KeyAscii, "p", String(TxtFields(Index).MaxLength, "@")
'    Case 4
'      'NÚMERO INTEIRO
'      Cmask.mask TxtFields(Index), KeyAscii, , String(TxtFields(Index).MaxLength, "9")
'    Case 5
'      'TELEFONE
'      Cmask.mask TxtFields(Index), KeyAscii, , "(99) 9999-9999"
'    Case 6
'      'CPF
'      Cmask.mask TxtFields(Index), KeyAscii, , "999.999.999.99"
'    Case 7
'      'CNPJ
'      Cmask.mask TxtFields(Index), KeyAscii, , "99.999.999/9999-99"
'    Case 8
'      'HORA
'      Cmask.mask TxtFields(Index), KeyAscii, , "99:99:99"
'    Case 9
'      'BOOL
'      Cmask.mask TxtFields(Index), KeyAscii, "u", "01"
'    Case 10
'      'SIM/NÃO
'      Cmask.mask TxtFields(Index), KeyAscii, "u", "SN"
'    Case 11
'      'VERDADEIRO/FALSO
'      Cmask.mask TxtFields(Index), KeyAscii, "u", "VF"
'End Select
'  '
'  If KeyAscii = vbKeyReturn Then
'    Call CMask_AutoTabStart(TxtFields(IIf((Index < (TxtFields.Count - 1)), Index + 1, Index)))
'  ElseIf KeyAscii <> 8 Then
'    Cmask.AutoTab TxtFields(Index), TxtFields(IIf((Index < (TxtFields.Count - 1)), Index + 1, Index))
'  End If
'End Sub

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome         : IPAGE_MaskEdit                                          <
'> Data e Hora  : 29/10/2006 11:36                                        <
'> Autor        : Diógenes Dias de Souza Júnior                           <
'> Descrição    : Classe para formatação de caixas de entrada através de  <
'>                máscaras.                                               <
'> Modificada em         : 29/10/2006 11:36                               <
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Option Explicit
Public Event AutoTabStart(Obj As Object)
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : AutoTab                                        <
'> Data/Hora             : 29/10/2006 11:37                               <
'> Autor                 : Diógenes Dias de Souza Júnior                  <
'> Descrição             : <
'> Parâmetros Passados   : <
'> Parâmetros Retornados : <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 29/10/2006 11:37                               <
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Sub AutoTab(m_Ctrl As Object, Optional m_CtrlReceiveFocus As Object)
  '
  '              \!!!!!!!/
  '               ( õ õ )
  '+----------oOOO--(_)------------IPAGE--+
  '|                                      |
  '| Definição das variáveis              |
  '|                                      |
  '+-IPAGE---------------oOOO-------------+
  '               |__| |__|
  '                ||   ||
  '               ooO  Ooo
  '
  Dim L        As Long
  '
  On Error GoTo LogErrorHandler
  L = Len(m_Ctrl.Text)
  '
  If L >= (m_Ctrl.MaxLength - 1) Then
    '
    '> Verifico se o segundo controle foi definido
    '
    L = Len(m_CtrlReceiveFocus.Name)
    '
    '> Antes de passar o foco eu verifico se o mesmo pode aceitar.
    '
    If m_CtrlReceiveFocus.Enabled = True And m_CtrlReceiveFocus.Visible = True Then
      RaiseEvent AutoTabStart(m_CtrlReceiveFocus)
      m_CtrlReceiveFocus.SetFocus
    Else
      RaiseEvent AutoTabStart(m_CtrlReceiveFocus)
      SendKeys "{TAB}"
    End If
    '
  End If
  '
  Exit Sub
LogErrorHandler:
  If Err.Number = 91 Then
    'A variável do objeto ou a variável do bloco 'With' não foi definida
    Err.Clear
    RaiseEvent AutoTabStart(m_CtrlReceiveFocus)
    SendKeys "{TAB}"
  Else
    'Qualquer outro erro
    MsgBox "Ocorreu um excessão de n.º " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & _
           "Na sub: AutoTab da Classe IPAGE_MaskEdit", vbCritical, EMPRESA
    Err.Clear
  End If
End Sub
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : ChangeColor                                    <
'> Data/Hora             : 29/10/2006 11:37                               <
'> Autor                 : Diógenes Dias de Souza Júnior                  <
'> Descrição             : Modifica a cor da caixa de texto quando entra  <
'>                         em foco e restaura a cor padrão quando sai do  <
'>                         foco.                                          <
'> Parâmetros Passados   : O objeto textbox, um valor booleano indicando  <
'>                         o tipo de operação, a cor do foco e a cor do   <
'>                         lostfocus.                                     <
'> Parâmetros Retornados : <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 29/10/2006 11:37                               <
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Sub ChangeColor(m_Ctrl As Object, _
                Optional m_Value As Boolean = True, _
                Optional m_ColorFocus As Long = &HFFFF00, _
                Optional m_ColorLostFocus As Long = vbWindowBackground)
  '
  '*** Frm é o nome do Form ***
  '*** m_Ctrl é o nome do obejto ***
  '*** m_Value é false para evento gotfocus e true para Lostfocus ***
  '
  Select Case m_Value
    Case False
      m_Ctrl.BackColor = m_ColorFocus
    Case True
      m_Ctrl.BackColor = m_ColorLostFocus
  End Select
End Sub
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : Mask                                           <
'> Data/Hora             : 29/10/2006 11:35                               <
'> Autor                 : Diógenes Dias de Souza Júnior                  <
'> Descrição             : <
'> Parâmetros Passados   : <
'> Parâmetros Retornados : <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 29/10/2006 11:35                               <
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Sub mask(m_Ctrl As Object, m_KeyAscii As Integer, Optional m_FormatText As String = "", Optional m_Mask As String = "")

  On Error GoTo LogErrorHandler
  '
  '              \!!!!!!!/
  '               ( õ õ )
  '+----------oOOO--(_)------------IPAGE--+
  '|                                      |
  '| Definição das variáveis              |
  '|                                      |
  '+-IPAGE---------------oOOO-------------+
  '               |__| |__|
  '                ||   ||
  '               ooO  Ooo
  '
  '
  '
  'U = Ucase, L = Lcase, P = ProperCase, D = Data, C = Currency
  '
  Const M_F = "UuLuPpDdCc"
  Const ONLY_WORD = "0123456789"
  Dim Flag_Format As Integer
  Dim Char1    As String
  Dim Char2    As String
  Dim L        As Integer
  Dim Tmp      As String
  Dim mFormatData As String
  Dim FormatCurrency As String
  Dim Slash    As Variant
  Dim Slash2   As Variant
  Dim m_Decimal As Integer
  '
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '> Habilita as teclas de atalho <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Select Case m_KeyAscii
    Case 3, 22, 24, 26
      'Permite Ctrl+C, Ctrl+V e Ctrl+X
      Exit Sub
    Case vbKeyReturn
      SendKeys "{TAB}"
      Exit Sub
  End Select
  '
Begin:
  '
  L = Len(m_Mask)
  '
  Flag_Format = InStr(1, M_F, m_FormatText, vbTextCompare)
  '
  If Flag_Format > 0 Or Len(Flag_Format) < Len(M_F) Then
    '
    Select Case UCase(Mid(m_FormatText, 1, 1))
      Case "C"
        If Len(Trim(m_Mask)) = False Then
          FormatCurrency = ""
          Slash = ","
          Slash2 = "."
        Else
          If InStr(1, m_Mask, ".", vbTextCompare) <> 0 Then
            m_Decimal = InStr(1, m_Mask, ".", vbTextCompare) - 1
            FormatCurrency = Mid(m_Mask, 1, m_Decimal) + Mid(m_Mask, m_Decimal + 2)
            Slash = ","
            Slash2 = "."
          ElseIf InStr(1, m_Mask, ".", vbTextCompare) <> 0 Then
            m_Decimal = InStr(1, m_Mask, ".", vbTextCompare) - 1
            FormatCurrency = Mid(m_Mask, 1, m_Decimal) + Mid(m_Mask, m_Decimal + 2)
            Slash = ","
            Slash2 = "."
          Else
            FormatCurrency = ""
            Slash = ","
            Slash2 = "."
          End If
          '
          Tmp = m_Ctrl.Text
          '
          If InStr(1, Tmp, FormatCurrency, vbTextCompare) = False Then
            Tmp = FormatCurrency & " " & Tmp
            m_Ctrl.Text = Tmp
            m_Ctrl.SelStart = Len(Tmp)
          End If
        End If
        '
        'Currency
        '*** Permite só números com máscara para Moeda ***
        '
        Select Case m_KeyAscii
          Case 3, 22, 24, 26
            'Permite Ctrl+C, Ctrl+V e Ctrl+X
            Exit Sub
          Case 8
            Exit Sub
          Case Asc(Slash2)
            '*** Evita a repetição da vírgula ***
            If InStr(m_Ctrl, Slash) <> 0 Then
              m_KeyAscii = False
              Exit Sub
            End If
            '
            m_KeyAscii = Asc(Slash)
          Case Asc(Slash)
            '*** Evita a repetição da vírgula ***
            If InStr(m_Ctrl, Slash) <> 0 Then
              m_KeyAscii = False
            End If
          Case 48 To 57, Asc(Slash)
          Case Asc("-")
            '*** Evita a repetição do sinal de menos ***
            If InStr(m_Ctrl, "-") <> 0 Then
              m_KeyAscii = False
              Exit Sub
            End If
            '
          Case Else
            m_KeyAscii = False
        End Select
        '
        Exit Sub
        '
      Case "D"
        If Len(Trim(m_Mask)) = False Then
          mFormatData = "DD/MM/YYYY"
          m_Mask = "99/99/9999"
          Slash = Split(m_Mask, "/", , vbTextCompare)
        Else

          If InStr(1, m_Mask, "-", vbTextCompare) <> 0 Then
            Slash = Split(m_Mask, "-", , vbTextCompare)
            mFormatData = "YYYY-MM-DD"                'Formato do MYSQL
          ElseIf InStr(1, m_Mask, "/", vbTextCompare) <> 0 Then
            Slash = Split(m_Mask, "/", , vbTextCompare)
            mFormatData = "MM/DD/YYYY"                'Formato americano
          End If
          '
          If UBound(Slash) < 2 Then
            Slash = Split("99/99/9999", "/", , vbTextCompare)
          End If
          '
        End If
        '
        '*** Permite só números com máscara para data ***
        '
        Select Case m_KeyAscii
          Case vbKeyReturn
            m_KeyAscii = False
            m_Ctrl = Format(m_Ctrl, mFormatData)
          Case 8
            Exit Sub
          Case Asc("/")
            Exit Sub
          Case 3, 22, 24, 26
            'Permite Ctrl+C, Ctrl+V e Ctrl+X
            Exit Sub

          Case 48 To 57
            If Len(m_Ctrl.Text) = Len(Slash(0)) Then
              m_Ctrl.Text = m_Ctrl.Text & "/"
              m_Ctrl.SelStart = Len(m_Ctrl.Text)
            ElseIf Len(m_Ctrl.Text) = (Len(Slash(0)) + Len(Slash(1)) + 1) Then
              m_Ctrl.Text = m_Ctrl.Text & "/"
              m_Ctrl.SelStart = Len(m_Ctrl.Text)
            End If
            '
          Case Else
            m_KeyAscii = False
        End Select
        '
        Exit Sub
        '
      Case "U"                                        ' maiúscula
        m_KeyAscii = Asc(UCase(Chr(m_KeyAscii)))
      Case "L"                                        ' Minúscula
        m_KeyAscii = Asc(LCase(Chr(m_KeyAscii)))
      Case "P"                                        'Nome próprio
        Tmp = m_Ctrl.Text + Chr(m_KeyAscii)
        '
        '> Não avalia os espaços <
        '
        If m_KeyAscii <> 32 Then
          Tmp = ProperCase(Tmp)
        End If
        '
        m_Ctrl.Text = Left(Tmp, Len(Tmp) - 1)
        m_Ctrl.SelStart = Len(m_Ctrl)
        m_KeyAscii = Asc(Right(Tmp, 1))
        '
    End Select
    '
  End If
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '> Caractere pode gerar erros na consulta <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '
  If m_KeyAscii = Asc("'") Then
    m_KeyAscii = False
    Exit Sub
  ElseIf m_KeyAscii = 8 Then                          'BackSpace
    Exit Sub
  End If
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '> Não possui nenhuma formatação <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '
  If L = 0 Then
    Exit Sub
  End If
  '
  Tmp = Left(m_Ctrl.Text + Chr(m_KeyAscii), Len(m_Mask))
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '> Não permite que o texto seja maior do que a máscara <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '
  If Len(Tmp) > Len(m_Mask) Then
    m_Ctrl.Text = Tmp
    m_Ctrl.SelStart = Len(Tmp)
    m_KeyAscii = False
    Exit Sub
  End If
  '
  Char1 = Mid(m_Mask, Len(Tmp), 1)
  Char2 = Mid(Tmp, Len(Tmp), 1)
  '
  Select Case UCase(Char1)
    Case "9"
      'Apenas número
      If IsNumeric(Char2) = False Then
        m_KeyAscii = False
      End If
    Case "@"
      'Apenas letras
      If InStr(1, ONLY_WORD, UCase(Chr(m_KeyAscii)), vbTextCompare) <> 0 Then
        m_KeyAscii = False
      End If
      '
    Case "-", "/", "(", ")", ".", "+", "*", " ", ":"
      Tmp = m_Ctrl.Text & Char1
      m_Ctrl.Text = Tmp
      m_Ctrl.SelStart = Len(Tmp)
      GoTo Begin
    Case Else
      If InStr(1, m_Mask, Chr(m_KeyAscii), vbTextCompare) = False Then
        m_KeyAscii = False
      End If
  End Select
  '
  Exit Sub
LogErrorHandler:
  m_KeyAscii = False

  Err.Clear
End Sub
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : ProperCase                                     <
'> Data/Hora             : 29/10/2006 21:25                               <
'> Autor                 : Diógenes Dias de Souza Júnior                  <
'> Descrição             : Converte as primeiras letra dos nomes em       <
'>                         maiúsculas menos: de, da, do, das, dos, a, e e.<
'>                         Retira os espaços em branco excedentes entre   <
'>                         as m_Words.                                    <
'> Parâmetros Passados   : Uma String.                                    <
'> Parâmetros Retornados : Uma string no formato nome próprio.            <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 29/10/2006 21:25                               <
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Function ProperCase(m_Word As String) As String
  '
  '              \!!!!!!!/
  '               ( õ õ )
  '+----------oOOO--(_)------------IPAGE--+
  '|                                      |
  '| Definição das variáveis              |
  '|                                      |
  '+-IPAGE---------------oOOO-------------+
  '               |__| |__|
  '                ||   ||
  '               ooO  Ooo
  '
  Dim Resto    As String
  Dim Parte    As String
  Dim Posic    As Byte
  '
  Resto = Trim(LCase(m_Word)) & " "
  '
  Do Until Resto = " "
    Posic = InStr(1, Resto, " ")
    Parte = Mid(Resto, 1, Posic)
    '
    If Trim(Parte) <> "e" And _
       Trim(Parte) <> "a" And _
       Trim(Parte) <> "do" And _
       Trim(Parte) <> "dos" And _
       Trim(Parte) <> "da" And _
       Trim(Parte) <> "das" And _
       Trim(Parte) <> "de" Then
      '
      Parte = UCase(Left(Parte, 1)) & Mid(Parte, 2)
      '
    End If
    '
    ProperCase = ProperCase & Parte
    Resto = Trim(Mid(Resto, Posic)) & " "
  Loop
  '
  ProperCase = Trim(ProperCase)
  '
End Function
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : SelText                                   <
'> Data/Hora             : 29/10/2006 21:27                               <
'> Autor                 : Diógenes Dias de Souza Júnior                  <
'> Descrição             : Seleciona o m_Ctrl quando entra em foco em uma <
'>                         caixa de controle.                             <
'> Parâmetros Passados   : O objeto m_Ctrl, a posição inicial de seleção  <
'>                         a posição final de seleção.                    <
'> Parâmetros Retornados : Nenhum                                         <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 29/10/2006 21:27                               <
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Sub SelText(m_Ctrl As Object, _
            Optional m_Start As Integer = False, _
            Optional m_MaxLength As Integer = 255)
  '
  m_Ctrl.SelStart = m_Start
  m_Ctrl.SelLength = m_MaxLength
  '
End Sub
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : Version                                        <
'> Data/Hora             : 24/10/2006 12:51                               <
'> Autor                 : Diógenes Dias de Souza Júnior                  <
'> Descrição             : Momstra a versão deste componente.             <
'> Parâmetros Passados   : <
'> Parâmetros Retornados : <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 24/10/2006 12:51                               <
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Function Version() As String
  Version = App.Major & "." & App.Minor & "." & App.Revision
End Function

