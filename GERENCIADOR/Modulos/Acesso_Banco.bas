Attribute VB_Name = "Acesso_Banco"
Option Explicit
Dim Linha()         As String
Dim intregs         As String
Dim strNOVACONECT   As String

Public Function CarregaStrConect() As String

    CarregaStrConect = ""

On Error GoTo err_CarregaStrConect

    'Abre aquivo de Configuração
    Open App.Path & "\" & "SIGE.txt" For Input As #1
    
        intregs = 0
        strNOVACONECT = ""
        Do While Not EOF(1)
           
           ReDim Preserve Linha(intregs)
           Input #1, Linha(intregs)
           strNOVACONECT = strNOVACONECT & Linha(intregs) & "^"
           intregs = intregs + 1
           
        Loop
    Close #1
    
    CarregaStrConect = strNOVACONECT

    Exit Function

err_CarregaStrConect:

    MsgBox "ATENÇÃO" & vbCrLf & _
           "Erro Numero  : " & Err.Number & vbCrLf & _
           "Erro Dcrição : " & Err.Description, vbOKOnly, vbExclamation, "Aviso"
    

End Function

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
     If BD.State = 0 Then BD.ConnectionString = Conecxao(1) & Conecxao(2) & Conecxao(3)
  End If
  
  If BD.State = 0 Then BD.Open
    
  Exit Sub
err_sys:
  
  If Err.Number = 53 Then
     MsgBox "ATENÇÃO - Erro : " & Err.Number & " Arquivo de Configuração não encontrado" & vbCrLf & _
            "Função = Banco_Dados" & vbCrLf & _
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
            "Caminho = " & App.Path, vbOKOnly + vbCritical, "Aviso"
  Else
     MsgBox "ATENÇÃO - " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Aviso"
  End If
  
End Sub


