Attribute VB_Name = "basMain"
'
'                                ################
'                      ####################################
'                  #############################################
'            ###########                                      #######
'          ##########                                              #####
'        #########                                                     ###
'     #########                                                            #
'   ########
'  ########
' #######  ######   #######       ####        #########     #########
' ######     ##     ##   ###     ##  ##       ##            ##
' #####      ##     #######     ########      ##  #####     ######
' #####      ##     ##         ##      ##     ##     ##     ##
'  ####    ######   ##        ##        ##    #########     #########
'   ####
'     ###
'       #
'
'                                    ##
' #      # #      # #      #              # #####    #####   #####    ######        ####    ####    ##  ##     ##
' #      # #      # #      #       ####   ##    ##        # #     ## #     ##     ##      ##    ## ## ## ##     #       # ####
' #  ##  # #  ##  # #  ##  #         ##   #######    ######  ####### ######       #       #      # #  ##  #     ######   #    #
' ## ## ## ## ## ## ## ## ##  ##     ##   #        ##     #        # #        ##  ##      ##    ## #      # ##  #    ##  #
'  ##  ##   ##  ##   ##  ##   ##   ###### ###       ##### #   #####   ######  ##    ####    ####   #      # ## #######  #####
'
'
'
' [Nomeclatura utilizada para as variáveis, nome de sub-rotinas e funções]

' · Todas as variáveis definidas de sub-rotinas e funções começaram por m_
' · Todas as variáveis de banco de dados começaram por Db_
' · Todos os parâmetros passados a Subs e ou funções começam por p_
' · Variáveis de LOOP começam por I, J, K, L...
' . Todas as Constantes são em maiúsculas
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'
Option Explicit
Global noExecute As Integer
Global StringLog As String
Global gConnectionString As String
Global Const EMPRESA As String = "IPAGE SOFTWARE"

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : Main
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
Sub Main()
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '> DECLARAÇÃO VARIÁVEIS <
  '/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Dim F        As New MdiMain
  '
  StringLoc = ""
  F.Show
End Sub

