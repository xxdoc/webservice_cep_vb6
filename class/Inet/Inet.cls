VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome          : Inet
'> Data e Hora   : 27/07/2019 10:05
'> Autor         : Diógenes Dias de Souza Júnior                          <
'> Descrição     :
'> Modificada em : 27/07/2019 10:05
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'

Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const ERROR_SUCCESS As Long = 0
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1

Private Type HOSTENT
  hName        As Long
  hAliases     As Long
  hAddrType    As Integer
  hLen         As Integer
  hAddrList    As Long
End Type

Private Type WSADATA
  wVersion     As Integer
  wHighVersion As Integer
  szDescription(0 To MAX_WSADescription) As Byte
  szSystemStatus(0 To MAX_WSASYSStatus) As Byte
  wMaxSockets  As Integer
  wMaxUDPDG    As Integer
  dwVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "wsock32" () As Long

Private Declare Function WSAStartup Lib "wsock32" _
                                    (ByVal wVersionRequired As Long, _
                                     lpWSADATA As WSADATA) As Long

Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function gethostname Lib "wsock32" _
                                     (ByVal szHost As String, _
                                      ByVal dwHostLen As Long) As Long

Private Declare Function gethostbyname Lib "wsock32" _
                                       (ByVal szHost As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
                               Alias "RtlMoveMemory" _
                               (hpvDest As Any, _
                                ByVal hpvSource As Long, _
                                ByVal cbCopy As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                      (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
                                       ByVal lpParameters As String, ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : GetIPAddress
'> Data/Hora             : 27/07/2019 10:07
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:07
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Function GetIPAddress() As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim sHostName As String * 256
  Dim lpHost   As Long
  Dim HOST     As HOSTENT
  Dim dwIPAddr As Long
  Dim tmpIPAddr() As Byte
  Dim i        As Integer
  Dim sIPAddr  As String

  If Not SocketsInitialize() Then
    GetIPAddress = ""
    Exit Function
  End If

  'gethostname returns the name of the local host into
  'the buffer specified by the name parameter. The host
  'name is returned as a null-terminated string. The
  'form of the host name is dependent on the Windows
  'Sockets provider - it can be a simple host name, or
  'it can be a fully qualified domain name. However, it
  'is guaranteed that the name returned will be successfully
  'parsed by gethostbyname and WSAAsyncGetHostByName.

  'In actual application, if no local host name has been
  'configured, gethostname must succeed and return a token
  'host name that gethostbyname or WSAAsyncGetHostByName
  'can resolve.
  If gethostname(sHostName, 256) = SOCKET_ERROR Then
    GetIPAddress = ""
    MsgBox "Windows Sockets error " & STR$(WSAGetLastError()) & _
           " has occurred. Unable to successfully get Host Name."
    SocketsCleanup
    Exit Function
  End If

  'gethostbyname returns a pointer to a HOSTENT structure
  '- a structure allocated by Windows Sockets. The HOSTENT
  'structure contains the results of a successful search
  'for the host specified in the name parameter.

  'The application must never attempt to modify this
  'structure or to free any of its components. Furthermore,
  'only one copy of this structure is allocated per thread,
  'so the application should copy any information it needs
  'before issuing any other Windows Sockets function calls.

  'gethostbyname function cannot resolve IP address strings
  'passed to it. Such a request is treated exactly as if an
  'unknown host name were passed. Use inet_addr to convert
  'an IP address string the string to an actual IP address,
  'then use another function, gethostbyaddr, to obtain the
  'contents of the HOSTENT structure.
  sHostName = Trim$(sHostName)
  lpHost = gethostbyname(sHostName)

  If lpHost = 0 Then
    GetIPAddress = ""
    MsgBox "Windows Sockets are not responding. " & _
           "Unable to successfully get Host Name."
    SocketsCleanup
    Exit Function
  End If

  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
  CopyMemory HOST, lpHost, Len(HOST)
  CopyMemory dwIPAddr, HOST.hAddrList, 4

  'create an array to hold the result
  ReDim tmpIPAddr(1 To HOST.hLen)
  CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen

  'and with the array, build the actual address,
  'appending a period between members
  For i = 1 To HOST.hLen
    sIPAddr = sIPAddr & tmpIPAddr(i) & "."
  Next

  'the routine adds a period to the end of the
  'string, so remove it here
  GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)

  SocketsCleanup

End Function

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : GetIPHostName
'> Data/Hora             : 27/07/2019 10:02
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:02
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Function GetIPHostName() As String
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim sHostName As String * 256

  If Not SocketsInitialize() Then
    GetIPHostName = ""
    Exit Function
  End If

  If gethostname(sHostName, 256) = SOCKET_ERROR Then
    GetIPHostName = ""
    MsgBox "Windows Sockets error " & STR$(WSAGetLastError()) & _
           " has occurred.  Unable to successfully get Host Name."
    SocketsCleanup
    Exit Function
  End If

  GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
  SocketsCleanup

End Function

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : HiByte
'> Data/Hora             : 27/07/2019 10:02
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:02
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Function HiByte(ByVal wParam As Integer) As Byte
  'note: VB4-32 users should declare this function As Integer
  HiByte = (wParam And &HFF00&) \ (&H100)
End Function

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : LoByte
'> Data/Hora             : 27/07/2019 10:02
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:02
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Function LoByte(ByVal wParam As Integer) As Byte
  'note: VB4-32 users should declare this function As Integer
  LoByte = wParam And &HFF&
End Function

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : SocketsCleanup
'> Data/Hora             : 27/07/2019 10:02
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:02
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Sub SocketsCleanup()
  If WSACleanup() <> ERROR_SUCCESS Then
    MsgBox "Socket error occurred in Cleanup."
  End If
End Sub

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : SocketsInitialize
'> Data/Hora             : 27/07/2019 10:02
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:02
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Function SocketsInitialize() As Boolean
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim WSAD     As WSADATA
  Dim sLoByte  As String
  Dim sHiByte  As String

  If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
    MsgBox "The 32-bit Windows Socket is not responding."
    SocketsInitialize = False
    Exit Function
  End If

  If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
    MsgBox "This application requires a minimum of " & _
           CStr(MIN_SOCKETS_REQD) & " supported sockets."

    SocketsInitialize = False
    Exit Function
  End If

  If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then

    sHiByte = CStr(HiByte(WSAD.wVersion))
    sLoByte = CStr(LoByte(WSAD.wVersion))

    MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
           " is not supported by 32-bit Windows Sockets."

    SocketsInitialize = False
    Exit Function

  End If

  'must be OK, so lets do it
  SocketsInitialize = True

End Function

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : GoWebPage
'> Data/Hora             : 27/07/2019 10:02
'> Autor                 : Diógenes Dias de Souza Júnior
'> Descrição             :
'> Parâmetros Passados   :
'> Parâmetros Retornados :
'> Dependências          :
'> Categoria             :
'> Modificada em         : 27/07/2019 10:02
'>                                                                        <
'> © Copyright IPAGE - Automação Comercial, Cursos e Soluções para WEB    <
'> email: diogenesdias@hotmail.com                                        <
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Sub GoWebPage()
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim lRetVal  As Long
  Dim lWindow  As Long
  '
  lRetVal = ShellExecute(lWindow, "open", App.Path & "\MyWebPage.Url", "", "", 1)
End Sub
