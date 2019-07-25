VERSION 5.00
Begin VB.Form frmCADDIMENCORTE 
   Caption         =   "Cadastro de Dimensões de Corte"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   10350
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   10215
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   8295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
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
         Left            =   840
         Picture         =   "frmCADDIMENCORTE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
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
         Picture         =   "frmCADDIMENCORTE.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Volta"
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
         Picture         =   "frmCADDIMENCORTE.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADDIMENCORTE"
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
Public lngCodUsuario    As Long

Dim objBLBFunc          As Object
Dim objCADDIMENCORTE    As Object
Dim strCAPTION          As String

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(frmCADDIMENCORTE, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
       
    If cTipOper = "I" Then objCADDIMENCORTE.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
       
    objCADDIMENCORTE.DESCRI = "'" & Trim(txtDescricao.Text) & "'"
       
    If objCADDIMENCORTE.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A dimensão de corte foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Call Destroy_Objeto
    Unload Me
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADDIMENCORTE = Nothing
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADDIMENCORTE = CreateObject("CADDIMENCORTE.clsCADDIMENCORTE")
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strCAPTION = "Cadastro de Agrupamento de Operações de Caixa - "
   
    objCADDIMENCORTE.FILIAL = FILIAL
   
    Call IniciaForm

End Sub


Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Dimensões de Corte - [ INCLUSÃO ] / " & Me.Name
    
    objBLBFunc.LimpaCampos frmCADDIMENCORTE
    
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADDIMENCORTE
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição da dimensão Inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function


Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(frmCADDIMENCORTE, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(frmCADDIMENCORTE)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    If cTipOper = "I" Then iCodigo = 0
    objCADDIMENCORTE.CODIGO = iCodigo
    
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
    If objCADDIMENCORTE.Carrega_Campos = True Then
        txtCodigo.Text = objCADDIMENCORTE.CODIGO
        txtDescricao.Text = objCADDIMENCORTE.DESCRI
    End If
End Sub

