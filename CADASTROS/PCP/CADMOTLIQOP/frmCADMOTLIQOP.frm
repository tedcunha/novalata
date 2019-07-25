VERSION 5.00
Begin VB.Form frmCADMOTLIQOP 
   Caption         =   "Cadastro de Motivos de Liquidação de OP"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   10290
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10215
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
         Picture         =   "frmCADMOTLIQOP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
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
         Picture         =   "frmCADMOTLIQOP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
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
         Picture         =   "frmCADMOTLIQOP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10215
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   8295
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
         TabIndex        =   4
         Top             =   240
         Width           =   585
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
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADMOTLIQOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho             As String
Public Linha                As Variant
Public cTipOper             As String
Public iCodigo              As Integer
Public FILIAL               As Integer
Public strAcesso            As String
Public strMODPAI            As String
Public strUsuario           As String
Public lngCODUSUARIO        As Long

Dim lngCodLog               As Long
Dim strVALOR                As String
Dim strCAPTION              As String

Dim objBLBFunc              As Object
Dim objCADLIQOP             As Object

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
       
    If cTipOper = "I" Then objCADLIQOP.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
       
    objCADLIQOP.DESCRI = "'" & Trim(txtDescricao.Text) & "'"
       
    If objCADLIQOP.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O Motivo de Liguidação da OP foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADLIQOP = CreateObject("CADMOTLIQOP.clsCADMOTLIQOP")
   
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objCADLIQOP.FILIAL = FILIAL
   
    strCAPTION = "Cadastro de Motivos de Liquidação de OP - "
   
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

Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    If cTipOper = "I" Then iCodigo = 0
    objCADLIQOP.CODIGO = iCodigo
    
    Call CarregaCampos
    
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
    End If
End Sub

Private Sub CarregaCampos()
    
    If objCADLIQOP.Carrega_Campos = False Then Exit Sub
    
    txtCodigo.Text = objCADLIQOP.CODIGO
    txtDescricao.Text = objCADLIQOP.DESCRI

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição de Motivo de Liquidação de Pedido Inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADLIQOP = Nothing
End Sub

