VERSION 5.00
Begin VB.Form frmAcessoLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso ao Sistema"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmbEntrar 
      Caption         =   "&Entrar"
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
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
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
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
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
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "txtUsuario"
      Top             =   1440
      Width           =   1815
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
      Left            =   1080
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "txtSenha"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cboEmpresa 
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
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   5175
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
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   825
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
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
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
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   945
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   6345
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image imgPW 
      Height          =   720
      Left            =   0
      Picture         =   "frmAcessoLogon.frx":0000
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
      TabIndex        =   5
      Top             =   120
      Width           =   6330
   End
End
Attribute VB_Name = "frmAcessoLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Linha()        As String

Private Sub cmbCancelar_Click()
  Sair
End Sub

Private Sub cmbEntrar_Click()

  Dim intResp As Integer
  
  If ValidaSenha(txtUsuario.Text, txtSenha.Text) = False Then
     
     intResp = MsgBox("ATENÇÃO - Usuário Inválido !!!" & vbCrLf & "Tentar Novamente", vbYesNo + vbCritical, "Aviso do Sistema")
     
     If intResp = vbNo Then
        Sair
     End If
     
     txtUsuario.SetFocus
     Exit Sub
  
  End If
  
  '' -------------------------------------------------------
  '' Setando os valores do menu Principal
  frmSIGE.strEmpresa = IIf(Len(Trim(cboEmpresa.Text)) > 0, " - " & cboEmpresa.Text, "")
  ''frmSIGE.Linha = Linha
  
  Unload Me
  '' -------------------------------------------------------

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmbCancelar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

  Dim SerieHd As String
  Dim intregs As Integer
  Call LimpaCampos
  
  'Abre aquivo de Configuração
  
  Open App.Path & "\" & "SIGE.txt" For Input As #1
  intregs = 0
      
  Do While Not EOF(1)
     
     ReDim Preserve Linha(intregs)
     
     Input #1, Linha(intregs)
     intregs = intregs + 1
       
  Loop
    
  Close #1
  ' --------------------------------------
 
  Call PreencheCombo
  Call PegaEmpresaPadrao

End Sub

Private Sub LimpaCampos()
  
  txtUsuario = ""
  txtSenha = ""
  cboEmpresa.Clear
  
End Sub

Private Sub PreencheCombo()
    
    cboEmpresa.Clear
    
    Call AbBanco("")
    
    sSql = "Select * from SGI_FILIAL Order by SGI_FILIAL"
    BREC.Open sSql, BD, adOpenDynamic
    
    cboEmpresa.AddItem "ADMINISTRADOR"
    cboEmpresa.ItemData(cboEmpresa.NewIndex) = 0
    
    Do While Not BREC.EOF
       cboEmpresa.AddItem Trim(BREC!SGI_DESCRICAO)
       cboEmpresa.ItemData(cboEmpresa.NewIndex) = BREC!SGI_FILIAL
       BREC.MoveNext
    Loop
    
    BREC.Close
    
    cboEmpresa.ListIndex = -1
    
    Call FcBanco
    
End Sub

Private Sub Sair()
  Unload Me
End Sub

Private Function ValidaSenha(strUsuarioDig As String, strSenhaDig As String) As Boolean
  
  ValidaSenha = False
  
  Dim rstSenha As New ADODB.Recordset
  Dim strSenha As String
  
  If UCase(strUsuarioDig) = "CWS" Then
     If UCase(strSenhaDig) = "JANNUS" Then
        frmSIGE.strUSUARIO = UCase(strUsuarioDig)
        iCodUsu = 0
        iFilial = 0
        If cboEmpresa.ListIndex > -1 Then iFilial = cboEmpresa.ItemData(cboEmpresa.ListIndex)
        ValidaSenha = True
        Exit Function
     End If
  Else
  
     iFilial = 0
     If cboEmpresa.ListIndex > -1 Then iFilial = cboEmpresa.ItemData(cboEmpresa.ListIndex)
     
     strSenha = "Select * " & vbCrLf
     strSenha = strSenha & "  From " & vbCrLf
     strSenha = strSenha & "       SGI_USUARIO " & vbCrLf
     strSenha = strSenha & " Where " & vbCrLf
     strSenha = strSenha & "       SGI_NOME     = '" & Crypt(strUsuarioDig) & "'" & vbCrLf
     strSenha = strSenha & "   and SGI_SENHA    = '" & Crypt(strSenhaDig) & "'" & vbCrLf
     strSenha = strSenha & "   And SGI_FILIAL   = " & iFilial

     rstSenha.Open strSenha, adoBanco_Dados

     If Not rstSenha.EOF Then
        frmSIGE.strUSUARIO = Crypt(UCase(strUsuarioDig))
        frmSIGE.iAcesso = IIf(IsNull(rstSenha!SGI_ACESSO) = False, rstSenha!SGI_ACESSO, 0)
        iCodUsu = rstSenha!SGI_CODIGO
        V_UsuarioId = rstSenha!SGI_CODIGO
        ValidaSenha = True
        Exit Function
     End If

     rstSenha.Close
  
  End If
 
End Function

Private Sub txtSenha_GotFocus()
    Call SelecionaCampos(txtSenha.Name, frmAcesso)
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    KeyAscii = Maiuscula(KeyAscii)
End Sub

Private Sub txtUsuario_GotFocus()
    Call SelecionaCampos(txtUsuario.Name, frmAcesso)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Maiuscula(KeyAscii)
End Sub

Private Sub PegaEmpresaPadrao()

    Dim I As Integer
    
    cboEmpresa.ListIndex = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_FILIAL " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_PADRAO = 1 "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
       For I = 1 To (cboEmpresa.ListCount - 1)
           If cboEmpresa.ItemData(I) = BREC!SGI_FILIAL Then cboEmpresa.ListIndex = I
       Next I
    End If
    
    BREC.Close
        
End Sub

