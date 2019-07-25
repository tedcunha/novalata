Attribute VB_Name = "mdlVariaveis"
Option Explicit
Public adoBanco_Dados           As New ADODB.Connection
Public adoBanco_Dados_Imagem    As New ADODB.Connection
Public adoBanco_Externo         As Object
Public cmExecuta                As New ADODB.Command
Public V_Sql                    As String
Public mResultado               As Variant
Public V_Usuario2               As String
Public V_UsuarioId              As Long
Public V_Usuario                As String
Public strNOVACONECT            As String


Public Const CB_FINDSTRING = &H14C

Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
( _
ByVal lpRootPathName As String, _
ByVal IpVolumeNameBuffer As String, _
ByVal nVolumeNameSize As Long, _
ipVolumeSerialNumber As Long, _
ipMaximumComponentLength As Long, _
IpFileSystemFlags As Long, _
ByVal IpFileSystemBuffer As String, _
ByVal nFileSystemNameSize As Long _
) As Long

Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type


Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_ALWAYS = 2
Const CREATE_NEW = 1
Const OPEN_ALWAYS = 4
Const OPEN_EXISTING = 3
Const TRUNCATE_EXISTING = 5
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Const FILE_FLAG_NO_BUFFERING = &H20000000
Const FILE_FLAG_OVERLAPPED = &H40000000
Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Const FILE_FLAG_WRITE_THROUGH = &H80000000


Function Get_Number_Serie(Unid As String) As String
    
    Dim IVSN As Long
    Dim n    As Long
    Dim s1   As String
    Dim s2   As String
    Dim sTmp As String
    s1 = String$(255, Chr$(0))
    s2 = String$(255, Chr$(0))
    n = GetVolumeInformation(Unid, s1, Len(s1), IVSN, 0, 0, s2, Len(s2))
    sTmp = Hex$(IVSN)
    Get_Number_Serie = Left$(sTmp, 4) & "-" & Right$(sTmp, 4)
    
End Function

Public Function UltNumero(strModulo As String) As Integer

    ' ---------------------------------------
    ' Pega Ultimo Código
    
    Dim L_Rst As New ADODB.Recordset
    
    V_Sql = "Select * " & vbCrLf
    V_Sql = V_Sql & "  From " & vbCrLf
    V_Sql = V_Sql & "       GERCODIGO " & vbCrLf
    V_Sql = V_Sql & " Where " & vbCrLf
    V_Sql = V_Sql & "       Modulo = '" & UCase(strModulo) & "'" & vbCrLf
    V_Sql = V_Sql & " ORDER BY " & vbCrLf
    V_Sql = V_Sql & "          Modulo," & vbCrLf
    V_Sql = V_Sql & "          CodigoID" & vbCrLf
    
    L_Rst.Open V_Sql, adoBanco_Dados, adOpenDynamic
    
    UltNumero = 1
    
    If Not L_Rst.EOF Then
       L_Rst.MoveLast
       UltNumero = L_Rst.Fields(1) + 1
    End If
    
    L_Rst.Close
    
    ' ---------------------------------------
    ' Inseri
    Set cmExecuta = New ADODB.Command
    cmExecuta.ActiveConnection = adoBanco_Dados
    
    V_Sql = "Insert Into GERCODIGO (" & vbCrLf
    V_Sql = V_Sql & "                       Modulo" & vbCrLf
    V_Sql = V_Sql & "                      ,CodigoID)" & vbCrLf
       
    V_Sql = V_Sql & "      Values ('" & UCase(strModulo) & "'" & vbCrLf
    V_Sql = V_Sql & "              ," & UltNumero & ")"
    
    cmExecuta.CommandText = V_Sql
    cmExecuta.Execute
    ' ---------------------------------------

End Function



Public Function AcessoSenha(wForm, wBotao) As Boolean

    AcessoSenha = False
    
    If V_Usuario = "CWS" Then
       AcessoSenha = True
       Exit Function
    End If
    
    Dim L_Rst As New ADODB.Recordset
    V_Sql = "Select ModuloId from Modulo where ModuloNom = '" & wForm & "' AND ModuloBot = '" & wBotao & "'"
    L_Rst.Open V_Sql, adoBanco_Dados
    If L_Rst.EOF Then
       MsgBox "Modulo nao encontrado. Acesso nao permitido", vbCritical
       Exit Function
    Else
       
       If V_Usuario <> "CWS" Then
           V_Sql = "select * from Acesso where UsuarioId = " & V_UsuarioId & " and ModuloId = " & L_Rst(0)
           L_Rst.Close
           L_Rst.Open V_Sql, adoBanco_Dados
           If L_Rst.EOF Then
              MsgBox "Acesso nao permitido", vbCritical
              Exit Function
           Else
              AcessoSenha = True
          End If
       End If
    End If
    AcessoSenha = True
    
End Function



Public Function AcessoTab(wForm, wBotao) As Boolean
    
    AcessoTab = False
    If V_Usuario = "CWS" Then
       AcessoTab = True
       Exit Function
    End If
    
    Dim L_Rst As New ADODB.Recordset
    V_Sql = "Select ModuloId from Modulo where ModuloNom = '" & wForm & "' AND ModuloBot = '" & wBotao & "'"
    L_Rst.Open V_Sql, adoBanco_Dados
    If L_Rst.EOF Then
        MsgBox "Modulo nao encontrado. Acesso nao permitido", vbCritical
        Exit Function
    Else
        V_Sql = "select * from Acesso where UsuarioId = " & V_UsuarioId & " and ModuloId = " & L_Rst(0)
        L_Rst.Close
        L_Rst.Open V_Sql, adoBanco_Dados
        If L_Rst.EOF Then
           Exit Function
        Else
           AcessoTab = True
        End If
    End If
    AcessoTab = True
    
End Function


Public Function EncriptaPW(vgSt As String) As String
   '***********************************
   'Funcao para criptografar palavra
   '***********************************
   
   Dim x As String
   x$ = Trim$(Cript$(RPad$(vgSt$, 25, "+"), "TINTURAR"))
   While Right$(x$, 1) = "+"
      x$ = Left$(x$, Len(x$) - 1)
   Wend
   EncriptaPW$ = x$
   
End Function

Public Function RPad(St As String, Tm As Integer, Ch As String) As String
   
   '***********************************
   'sub-Funcao para criptografar palavra
   '***********************************
   Dim x As String
   If VarType(St) = vbString Then
      x$ = St
   Else
      x$ = Str$(St)
   End If
   RPad$ = Left$(LTrim$(x$) + String$(Tm, Ch$), Tm)
   
End Function

Public Function Cript(St As String, Pw As String) As String
  '***********************************
  'SUB-Funcao para criptografar palavra
  '***********************************
   Dim x As String, i As Integer, n As Integer, _
       p As Integer, j As Integer, n0 As Integer
   p = 0
   For i = 1 To Len(St$)
      p = p + 1
      If p > Len(Pw$) Then p = 1
      j = Asc(Mid$(Pw$, p, 1)) Or 128
      n = Asc(Mid$(St$, i))

DeNovo:
      n = n Xor j
      If n < 31 Then
         n = (128 + n)
         GoTo DeNovo
      ElseIf n > 127 And n < 159 Then
         n = n - 128
         GoTo DeNovo
      End If
      x$ = x$ + Chr$(n)
   Next
   Cript$ = x$
   
End Function

Public Function GetFileAccessDate(fString As String) As Date
    Dim hFile As Long
    Dim ctime As FILETIME
    Dim atime As FILETIME
    Dim mtime As FILETIME
    Dim thetime As SYSTEMTIME
    Dim retval As Long
    hFile = CreateFile(fString, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
        If hFile = -1 Then
            MsgBox "Error Getting Date on " & fString
            Exit Function
        End If
    retval = GetFileTime(hFile, ctime, atime, mtime)
    retval = FileTimeToLocalFileTime(atime, atime)
    retval = FileTimeToSystemTime(atime, thetime)
    retval = CloseHandle(hFile)
    GetFileAccessDate = CDate(thetime.wMonth & "/" & thetime.wDay & "/" & thetime.wYear)
End Function
Public Function GetFileModifyDate(fString As String) As Date
    Dim hFile As Long
    Dim ctime As FILETIME
    Dim atime As FILETIME
    Dim mtime As FILETIME
    Dim thetime As SYSTEMTIME
    Dim retval As Long
    hFile = CreateFile(fString, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
        If hFile = -1 Then
            MsgBox "Error Getting Date on " & fString
            Exit Function
        End If
    retval = GetFileTime(hFile, ctime, atime, mtime)
    retval = FileTimeToLocalFileTime(mtime, mtime)
    retval = FileTimeToSystemTime(mtime, thetime)
    retval = CloseHandle(hFile)
    GetFileModifyDate = CDate(thetime.wMonth & "/" & thetime.wDay & "/" & thetime.wYear)
End Function
Public Function GetFileCreatedDate(fString As String) As Date
    Dim hFile As Long
    Dim ctime As FILETIME
    Dim atime As FILETIME
    Dim mtime As FILETIME
    Dim thetime As SYSTEMTIME
    Dim retval As Long
    hFile = CreateFile(fString, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
        If hFile = -1 Then
            MsgBox "Error Getting Date on " & fString
            Exit Function
        End If
    retval = GetFileTime(hFile, ctime, atime, mtime)
    retval = FileTimeToLocalFileTime(ctime, ctime)
    retval = FileTimeToSystemTime(ctime, thetime)
    retval = CloseHandle(hFile)
    GetFileCreatedDate = CDate(thetime.wMonth & "/" & thetime.wDay & "/" & thetime.wYear)
End Function

Public Function SetFileDates(fString As String, sCdate As Date, sAdate As Date, sMdate As Date) As Boolean
    Dim hFile As Long
    Dim cctime As SYSTEMTIME
    Dim aatime As SYSTEMTIME
    Dim mmtime As SYSTEMTIME
    Dim ctime As FILETIME
    Dim atime As FILETIME
    Dim mtime As FILETIME
    Dim retval As Long
    
    cctime.wDay = CInt(Format(sCdate, "dd"))
    cctime.wMonth = CInt(Format(sCdate, "mm"))
    cctime.wYear = CInt(Format(sCdate, "yyyy"))
    
    aatime.wDay = CInt(Format(sAdate, "dd"))
    aatime.wMonth = CInt(Format(sAdate, "mm"))
    aatime.wYear = CInt(Format(sAdate, "yyyy"))
    
    mmtime.wDay = CInt(Format(sMdate, "dd"))
    mmtime.wMonth = CInt(Format(sMdate, "mm"))
    mmtime.wYear = CInt(Format(sMdate, "yyyy"))
    
    retval = SystemTimeToFileTime(cctime, ctime)
    retval = SystemTimeToFileTime(aatime, atime)
    retval = SystemTimeToFileTime(mmtime, mtime)
    
    hFile = CreateFile(fString, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
        If hFile = -1 Then
            MsgBox "Could Not Set File Dates On " & fString
            SetFileDates = False
            Exit Function
        End If
    retval = SetFileTime(hFile, ctime, atime, mtime)
    retval = CloseHandle(hFile)
    SetFileDates = True
End Function
