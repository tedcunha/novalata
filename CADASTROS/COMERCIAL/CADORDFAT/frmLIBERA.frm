VERSION 5.00
Begin VB.Form frmLIBERA 
   Caption         =   "Liberação"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4785
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmbCancelar 
      Caption         =   "&Cancelar -  F5"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmbEntrar 
      Caption         =   "&Libera"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
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
      Left            =   2040
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "txtSenha"
      Top             =   1320
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
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "txtUsuario"
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label lblUsuáruo 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   825
   End
   Begin VB.Image imgPW 
      Height          =   720
      Left            =   0
      Picture         =   "frmLIBERA.frx":0000
      Top             =   0
      Width           =   720
   End
   Begin VB.Label labSenha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Controle de acesso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4770
   End
End
Attribute VB_Name = "frmLIBERA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Linha            As Variant
Public FILIAL           As Integer
Dim objBLBFunc          As Object

Private Sub cmbCancelar_Click()
    frmCADORDFAT.intLIBERASN = 0
    frmCADORDFAT.intLIB10PORC = 0
    Unload Me
End Sub

Private Sub cmbEntrar_Click()
    frmCADORDFAT.intLIB10PORC = PegaUsuarioLib10Porc(Trim(txtUsuario.Text), Trim(txtSenha.Text))
    frmCADORDFAT.intLIBERASN = PegaUsuarioLibSN(Trim(txtUsuario.Text), Trim(txtSenha.Text))
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmbCancelar_Click
    If KeyCode = vbKeyF15 Then cmbCancelar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If

    Call LimpaCampos

End Sub

Private Sub DestroyObjeto()
    Set objBLBFunc = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyObjeto
End Sub

Private Sub LimpaCampos()
    txtUsuario.Text = ""
    txtSenha.Text = ""
End Sub

Private Function PegaUsuarioLib10Porc(strUSUARIO As String, strSENHA As String) As Integer
    
    PegaUsuarioLib10Porc = 0
    
    Dim strUSUCRIPT As String
    Dim strSENCRIPT As String
    
    strUSUCRIPT = ""
    strSENCRIPT = ""
    
    strUSUCRIPT = "'" & objBLBFunc.Crypt(strUSUARIO) & "'"
    strSENCRIPT = "'" & objBLBFunc.Crypt(strSENHA) & "'"
    
    sSql = ""
    
    sSql = "  Select" & vbCrLf
    sSql = sSql & "         SGI_PERMFAT10POR" & vbCrLf
    sSql = sSql & "    From" & vbCrLf
    sSql = sSql & "         SGI_USUARIO" & vbCrLf
    sSql = sSql & "   Where" & vbCrLf
    sSql = sSql & "         SGI_NOME  = " & strUSUCRIPT & vbCrLf
    sSql = sSql & "     And SGI_SENHA = " & strSENCRIPT
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaUsuarioLib10Porc = BREC!SGI_PERMFAT10POR
    BREC.Close

End Function

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function PegaUsuarioLibSN(strUSUARIO As String, strSENHA As String) As Integer
    
    PegaUsuarioLibSN = 0
    
    Dim strUSUCRIPT As String
    Dim strSENCRIPT As String
    
    strUSUCRIPT = ""
    strSENCRIPT = ""
    
    strUSUCRIPT = "'" & objBLBFunc.Crypt(strUSUARIO) & "'"
    strSENCRIPT = "'" & objBLBFunc.Crypt(strSENHA) & "'"
    
    sSql = ""
    
    sSql = "  Select" & vbCrLf
    sSql = sSql & "         SGI_PERMFAT10POR" & vbCrLf
    sSql = sSql & "    From" & vbCrLf
    sSql = sSql & "         SGI_USUARIO" & vbCrLf
    sSql = sSql & "   Where" & vbCrLf
    sSql = sSql & "         SGI_NOME  = " & strUSUCRIPT & vbCrLf
    sSql = sSql & "     And SGI_SENHA = " & strSENCRIPT
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaUsuarioLibSN = BREC!SGI_PERMFAT10POR
    BREC.Close

End Function

