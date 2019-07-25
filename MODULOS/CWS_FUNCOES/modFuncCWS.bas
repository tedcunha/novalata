Attribute VB_Name = "modFuncCWS"
Option Explicit

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS As Long = 0
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1

Public Type HOSTENT
hName As Long
hAliases As Long
hAddrType As Integer
hLen As Integer
hAddrList As Long
End Type

Public Type WSADATA
wVersion As Integer
wHighVersion As Integer
szDescription(0 To MAX_WSADescription) As Byte
szSystemStatus(0 To MAX_WSASYSStatus) As Byte
wMaxSockets As Integer
wMaxUDPDG As Integer
dwVendorInfo As Long
End Type
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
(ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" _
(ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
(ByVal szHost As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)


Public Sub AbBanco(strConecxao As String)

On Error GoTo err_sys
  
  If Len(Trim(strConecxao)) = 0 Then Exit Sub
  
  Dim Conecxao() As String
  Conecxao = Split(strConecxao, "^")
  
  '"Microsoft.Jet.OLEDB.4.0"
  '------------------------------------------------------
  
  If Right(Conecxao(0), 1) = 1 Then
     BD.Provider = Mid(Conecxao(2), 11, 50) '' Drive do banco em Uso
     BD.ConnectionString = Mid(Conecxao(1), 9, 50) & Mid(Conecxao(3), 11, 50) '' Caminho do Banco
     BD.CommandTimeout = 1200 '' 20 minutos
  End If
  
  If Right(Conecxao(0), 1) = 2 Then
     If BD.State = 0 Then BD.ConnectionString = Conecxao(1) & Conecxao(2) & strPASSWORD
  End If
  
  If BD.State = 0 Then BD.Open
    
  Exit Sub
err_sys:
  
  If Err.Number = 53 Then
     MsgBox "ATENÇÃO - Erro : " & Err.Number & " Arquivo de Configuração não encontrado" & vbCrLf & _
            "Função = Banco_Dados" & vbCrLf & _
            "Biblioteca = BLBCWS.clsFuncoes" & vbCrLf & _
            "Caminho = " & App.Path, vbOKOnly + vbCritical, "Aviso"
  Else
     MsgBox "ATENÇÃO - " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Aviso"
  End If
  
End Sub


Public Sub FcBanco()

On Error GoTo err_sys
  
  If BD.State = 1 Then BD.Close
    
  Exit Sub
err_sys:
  
  If Err.Number = 53 Then
     MsgBox "ATENÇÃO - Erro : " & Err.Number & " Arquivo de Configuração não encontrado" & vbCrLf & _
            "Função = Banco_Dados" & vbCrLf & _
            "Biblioteca = BLBCWS.clsFuncoes" & vbCrLf & _
            "Caminho = " & App.Path, vbOKOnly + vbCritical, "Aviso"
  Else
     MsgBox "ATENÇÃO - " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Aviso"
  End If
  
End Sub

Public Sub SelecionaCampos(nomecampo As String, frmfornmulario As Form)
  
  Dim i As Integer
  
  For i = 0 To frmfornmulario.Count - 1
      If UCase(frmfornmulario.Controls(i).Name) = UCase(nomecampo) Then
         frmfornmulario.Controls(i).SelStart = 0
         frmfornmulario.Controls(i).SelLength = Len(frmfornmulario.Controls(i).Text)
      End If
  Next
  
End Sub

Public Function Maiuscula(CodAscII As Integer) As Integer
    Maiuscula = Asc(UCase(Chr(CodAscII)))
End Function

Public Function Crypt(Text As String) As String

    Dim i           As Integer
    Dim strTempChar As String
    
    For i = 1 To Len(Text)

        If Asc(Mid$(Text, i, 1)) < 128 Then
           strTempChar = Asc(Mid$(Text, i, 1)) + 128
        ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
           strTempChar = Asc(Mid$(Text, i, 1)) - 128
        End If

        Mid$(Text, i, 1) = Chr(strTempChar)

    Next i

    Crypt = Text

End Function

Public Function Gera_Codigo(sModulo As String, intFilial As Integer, Linha As Variant) As Long

    Gera_Codigo = 1
    
    Dim strCAMPO As String
    strCAMPO = Linha
    
    Call AbBanco(strCAMPO)
    
    BGRV.ActiveConnection = BD
    BD.BeginTrans
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql + "       (Max(SGI_NUMERO) + 1) As SGI_NUMERO " & vbCrLf
    sSql = sSql + "  From " & vbCrLf
    sSql = sSql + "       SGI_NUMERO " & vbCrLf
    sSql = sSql + " Where " & vbCrLf
    sSql = sSql + "       SGI_MODULO = '" & sModulo & "'"
    sSql = sSql + "   And SGI_FILIAL = " & intFilial
    
    BREC.Open sSql, BD, adOpenDynamic
    
    If Not BREC.EOF Then
    
       If IsNull(BREC!SGI_NUMERO) = True Then
          
          Gera_Codigo = 1
          
          sSql = "Insert into SGI_NUMERO (SGI_FILIAL,SGI_NUMERO,SGI_MODULO) Values(" & vbCrLf
          sSql = sSql + "                                              " & intFilial & vbCrLf
          sSql = sSql + "                                            ,1" & vbCrLf
          sSql = sSql + "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql + "                                          )" & vbCrLf
          
       ElseIf BREC!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql + "           SGI_NUMERO = " & BREC!SGI_NUMERO & vbCrLf
          sSql = sSql + "         Where " & vbCrLf
          sSql = sSql + "               SGI_MODULO = '" & sModulo & "'" & vbCrLf
          sSql = sSql + "           And SGI_FILIAL =  " & intFilial
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       BD.CommitTrans
       
       
    End If
    
    BREC.Close
    
    Call FcBanco
    
End Function


Public Function CabecForm(strCabec As String) As String
  CabecForm = strCabec
End Function

Public Function GravaLogSaida(Linha As String, intFilial As Integer, lngCODACESSO As Long) As Boolean

On Error GoTo err_Trans

     '' Inicia transação
     
     Call AbBanco(Linha)
     
     BD.BeginTrans
     BGRV.ActiveConnection = BD
     
     sSql = ""

     sSql = "Update SGI_LOGACESSO Set " & vbCrLf
     sSql = sSql & "                  SGI_DATSAIDA = '" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
     sSql = sSql & "Where" & vbCrLf
     sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
     sSql = sSql & "  And SGI_CODIGO = " & lngCODACESSO

     BGRV.CommandText = sSql
     BGRV.Execute

     BD.CommitTrans

     GravaLogSaida = True

     Call FcBanco
     
     Exit Function

err_Trans:
     
     If BD.State = 1 Then BD.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing


End Function

Public Sub LimpaCampos(frmForm As Variant)
  
  Dim vControl

  'Apagando todos os text boxes de um form:
  '----------------------------------------

  For Each vControl In frmForm.Controls
      If TypeOf vControl Is TextBox Then
         vControl.Text = ""
      'ElseIf TypeOf vControl Is MaskEdBox Then
      '   vControl.Text = "__/__/____"
      ElseIf TypeOf vControl Is OptionButton Then
         vControl.Value = 0
      ElseIf TypeOf vControl Is ComboBox Then
         If vControl.Style = 0 Then
            vControl.Text = ""
            vControl.Clear
         ElseIf vControl.Style = 2 Then
            vControl.Clear
         End If
      ElseIf TypeOf vControl Is ListBox Then
         vControl.Clear
      End If
   Next

End Sub

Public Function ChecaAcesso2(strTIPOPERACAO As String, strACESSO As String, Optional boolExibeMens As Boolean) As Boolean

  Dim i As Integer
  
  ChecaAcesso2 = False
  
  For i = 1 To Len(Trim(strACESSO))
      If Mid(strACESSO, i, 1) = strTIPOPERACAO Then ChecaAcesso2 = True
  Next i
  
  If boolExibeMens = True Then Exit Function
  
  If ChecaAcesso2 = False And strTIPOPERACAO = "I" Then
     MsgBox "Você não tem permissão para incluir !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "A" Then
     MsgBox "Você não tem permissão para alterar !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "E" Then
     MsgBox "Você não tem permissão para excluir !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "C" Then
     MsgBox "Você não tem permissão para consultar !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "R" Then
     MsgBox "Você não tem permissão para tirar relatórios !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "P" Then
     MsgBox "Você não tem permissão para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "L" Then
     MsgBox "Você não tem permissão para liberar !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "B" Then
     MsgBox "Você não tem permissão para bloquear !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "V" Then
     MsgBox "Você não tem permissão para reprovar !!!", vbOKOnly + vbExclamation, "Aviso"
  End If

End Function


Public Function GetIPAddress() As String
Dim sHostName As String * 256
Dim lpHost As Long
Dim HOST As HOSTENT
Dim dwIPAddr As Long
Dim tmpIPAddr() As Byte
Dim i As Integer
Dim sIPAddr As String
If Not SocketsInitialize() Then
GetIPAddress = ""
Exit Function
End If
If gethostname(sHostName, 256) = SOCKET_ERROR Then
GetIPAddress = ""
MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
" has occurred. Unable to successfully get Host Name."
SocketsCleanup
Exit Function
End If
sHostName = Trim$(sHostName)
lpHost = gethostbyname(sHostName)
If lpHost = 0 Then
GetIPAddress = ""
MsgBox "Windows Sockets are not responding. " & _
"Unable to successfully get Host Name."
SocketsCleanup
Exit Function
End If
CopyMemory HOST, lpHost, Len(HOST)
CopyMemory dwIPAddr, HOST.hAddrList, 4
ReDim tmpIPAddr(1 To HOST.hLen)
CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
For i = 1 To HOST.hLen
sIPAddr = sIPAddr & tmpIPAddr(i) & "."
Next
GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
SocketsCleanup
End Function

Public Function GetIPHostName() As String
Dim sHostName As String * 256
If Not SocketsInitialize() Then
GetIPHostName = ""
Exit Function
End If
If gethostname(sHostName, 256) = SOCKET_ERROR Then
GetIPHostName = ""
MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
" has occurred. Unable to successfully get Host Name."
SocketsCleanup
Exit Function
End If
GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
SocketsCleanup
End Function


Public Function HiByte(ByVal wParam As Integer)
HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
LoByte = wParam And &HFF&
End Function

Public Sub SocketsCleanup()
If WSACleanup() <> ERROR_SUCCESS Then
MsgBox "Socket error occurred in Cleanup."
End If
End Sub

Public Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
Dim sLoByte As String
Dim sHiByte As String
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
SocketsInitialize = True
End Function


