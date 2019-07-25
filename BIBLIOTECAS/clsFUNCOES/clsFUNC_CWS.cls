VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFUNC_CWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AbreBanco(Conecxao As Variant)

On Error GoTo err_sys
  
  Dim Linha()     As Variant
  Dim intregs     As Integer
  Dim I           As Integer
  
  Dim strNoBanco  As String
  Dim strProvider As String
  Dim strConnt    As String
  Dim strLinReg()   As String
    
  '"Microsoft.Jet.OLEDB.4.0"
  '------------------------------------------------------
  
  If Right(Conecxao(0), 1) = 1 Then
     BD.Provider = Mid(Conecxao(2), 11, 50) '' Drive do banco em Uso
     BD.ConnectionString = Mid(Conecxao(1), 9, 50) & Mid(Conecxao(3), 11, 50) '' Caminho do Banco
     BD.CommandTimeout = 1200 '' 20 minutos
  End If
  
  If Right(Conecxao(0), 1) = 2 Then
     BD.ConnectionString = Conecxao(1) & Conecxao(2) & Conecxao(3)
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

Public Sub FechaBanco()

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

Private Sub SelecionaCampos(nomecampo As String, frmfornmulario As Variant)
  
  Dim I As Integer
  
  For I = 0 To frmfornmulario.Count - 1
      
      If UCase(frmfornmulario.Controls(I).Name) = UCase(nomecampo) Then
         frmfornmulario.Controls(I).SelStart = 0
         frmfornmulario.Controls(I).SelLength = Len(frmfornmulario.Controls(I).Text)
      End If
  Next
  
End Sub

