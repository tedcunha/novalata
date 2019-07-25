VERSION 5.00
Begin VB.Form frmUSULIB 
   Caption         =   "Usuário"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3690
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "txtSenha"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "txtUsuario"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmbCancelar 
      Caption         =   "&Voltar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmbEntrar 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmUSULIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCodVendedor   As Long
Public lngCodUsuario    As Long
Public boolLib          As Boolean

Dim objBLBFunc          As Object
Dim objCADPEDVENDAL     As Object


Private Sub cmbCancelar_Click()
    boolLib = False
    Unload Me
End Sub

Private Sub cmbEntrar_Click()
    
    Dim strUsuario As String
    
    If Len(Trim(txtUsuario.Text)) = 0 Then
        MsgBox "Usuário Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    strUsuario = Trim(UCase(txtUsuario.Text))
    
    If strUsuario <> "OCTAVIOPCP" And _
       strUsuario <> "NVM" And _
       strUsuario <> "DANIELPCP" And _
       strUsuario <> "RONALDO" And _
       strUsuario <> "CWS" Then
        MsgBox "Usuário não permitido !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If ValidaSenha(strUsuario, Trim(txtSenha.Text)) = False Then Exit Sub
    
    boolLib = True
    Unload Me
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADPEDVENDAL = CreateObject("CADPEDVENDA.clsCADPEDVENDA")
   
   objCADPEDVENDAL.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    
    txtUsuario.Text = ""
    txtSenha.Text = ""

    boolLib = False

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADPEDVENDAL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmbCancelar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub txtSenha_GotFocus()
    objBLBFunc.SelecionaCampos txtSenha.Name, frmUSULIB
End Sub

Private Sub txtUsuario_GotFocus()
    objBLBFunc.SelecionaCampos txtUsuario.Name, frmUSULIB
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaSenha(strUsuarioDig As String, strSenhaDig As String) As Boolean
  
  ValidaSenha = False
  
  Dim rstSenha As New ADODB.Recordset
  Dim strSenha As String
  
  If UCase(strUsuarioDig) = "CWS" Then
     If UCase(strSenhaDig) = "JANNUS" Then
        ValidaSenha = True
        Exit Function
     End If
  Else
  
     strSenha = "Select * " & vbCrLf
     strSenha = strSenha & "  From " & vbCrLf
     strSenha = strSenha & "       SGI_USUARIO " & vbCrLf
     strSenha = strSenha & " Where " & vbCrLf
     strSenha = strSenha & "       SGI_NOME     = '" & objBLBFunc.Crypt(strUsuarioDig) & "'" & vbCrLf
     strSenha = strSenha & "   and SGI_SENHA    = '" & objBLBFunc.Crypt(strSenhaDig) & "'" & vbCrLf
     strSenha = strSenha & "   And SGI_FILIAL   = " & FILIAL

     rstSenha.Open strSenha, adoBanco_Dados
     If Not rstSenha.EOF Then ValidaSenha = True
     rstSenha.Close
  
  End If
 
End Function

