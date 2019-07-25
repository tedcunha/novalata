VERSION 5.00
Begin VB.Form frmCADSTATUSPPL 
   Caption         =   "Cadastro de Status de Pipeline"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   10260
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
         Picture         =   "frmCADSTATUSPPL.frx":0000
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
         Picture         =   "frmCADSTATUSPPL.frx":0532
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
         Picture         =   "frmCADSTATUSPPL.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10215
      Begin VB.TextBox txtPorc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Text            =   "txtPorc"
         Top             =   960
         Width           =   1095
      End
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
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   8295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   150
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
Attribute VB_Name = "frmCADSTATUSPPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Dim objCADSTATUSPPL     As Object

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub CmdSalva_Click()

    Dim sValor As String

    If ValidaCampos = False Then Exit Sub
       
    If cTipOper = "I" Then objCADSTATUSPPL.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
       
    objCADSTATUSPPL.DESCRI = "'" & Trim(txtDescricao.Text) & "'"
    
    sValor = "Null"
    If Len(Trim(txtPorc.Text)) > 0 Then
       sValor = Replace(txtPorc.Text, ".", "")
       sValor = Replace(sValor, ",", ".")
    End If
    objCADSTATUSPPL.PORC = sValor
       
    If objCADSTATUSPPL.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O Status de Pipe Line foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Public Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADSTATUSPPL = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADSTATUSPPL = CreateObject("CADSTATUSPPL.clsCADSTATUSPPL")
   
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objCADSTATUSPPL.FILIAL = FILIAL
   
    strCAPTION = "Cadastro de Status de Pipeline - "
   
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
    objCADSTATUSPPL.CODIGO = iCodigo
    
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
    
    If objCADSTATUSPPL.Carrega_Campos = False Then Exit Sub
    
    txtCodigo.Text = objCADSTATUSPPL.CODIGO
    txtDescricao.Text = objCADSTATUSPPL.DESCRI
    txtPorc.Text = objCADSTATUSPPL.PORC
    
End Sub

Private Sub txtPorc_GotFocus()
    objBLBFunc.SelecionaCampos txtPorc.Name, Me
End Sub

Private Sub txtPorc_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPorc.Text
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição do Tipo de Apontamento Inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtPorc.Text)) = 0 Then
        MsgBox "Campo de Porcentagem deve ser informado !!!", vbOKOnly + vbCritical, "Aviso"
        txtPorc.SetFocus
        Exit Function
     End If
     
     If CCur(txtPorc.Text) > 100 Then
        MsgBox "Campo de Porcentagem não pode ser maior que 100% !!!", vbOKOnly + vbCritical, "Aviso"
        txtPorc.SetFocus
        Exit Function
     End If
     
     If CCur(txtPorc.Text) <= 0 Then
        MsgBox "Campo de Porcentagem não pode ser menor ou igual a 0% !!!", vbOKOnly + vbCritical, "Aviso"
        txtPorc.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function

