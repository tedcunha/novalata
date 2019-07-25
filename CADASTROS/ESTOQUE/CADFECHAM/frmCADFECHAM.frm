VERSION 5.00
Begin VB.Form frmCADFECHAM 
   Caption         =   "Cadastro de Fechamento"
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   10305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   10215
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   8295
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Volta <ESC>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "frmCADFECHAM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         Picture         =   "frmCADFECHAM.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAltera 
         Caption         =   "&Altera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         Picture         =   "frmCADFECHAM.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADFECHAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCODUSUARIO    As Long

Dim lngCodLog           As Long
Dim strVALOR            As String
Dim strCAPTION          As String

Dim objBLBFunc          As Object
Dim objCADFECHAM        As Object
Dim objPESQPADRAO       As Object

Private Sub cmdAltera_Click()
    
    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
       
    If cTipOper = "I" Then objCADFECHAM.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
       
    objCADFECHAM.DESCRI = "'" & Trim(txtDescricao.Text) & "'"
       
    If objCADFECHAM.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O Fechamento foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADFECHAM = CreateObject("CADFECHAM.clsCADFECHAM")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objCADFECHAM.FILIAL = FILIAL
   
    strCAPTION = "Cadastro de Fechamento - "
   
    Call IniciaForm

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub


Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, Me
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
    End If
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição do Tipo de Apontamento Inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function

Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    If cTipOper = "I" Then iCodigo = 0
    objCADFECHAM.CODIGO = iCodigo
    
    Call CarregaCampos
    
End Sub

Private Sub CarregaCampos()
    
    If objCADFECHAM.Carrega_Campos = False Then Exit Sub
    
    txtCodigo.Text = objCADFECHAM.CODIGO
    txtDescricao.Text = objCADFECHAM.DESCRI

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADFECHAM = Nothing
    Set objPESQPADRAO = Nothing
End Sub
