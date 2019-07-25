VERSION 5.00
Begin VB.Form frmCADPARADAS 
   Caption         =   "Cadastro de Paradas"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   10275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   10215
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
         Begin VB.OptionButton optAtivoSN 
            Caption         =   "SIM"
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
            Index           =   1
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optAtivoSN 
            Caption         =   "NÃO"
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
            Left            =   840
            TabIndex        =   4
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.TextBox txtCODINT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "txtCODINT"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtDescricao"
         Top             =   960
         Width           =   8295
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ativo"
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
         TabIndex        =   13
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label3 
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
         TabIndex        =   12
         Top             =   600
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
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
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
         TabIndex        =   10
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
         Picture         =   "frmCADPARADAS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmCADPARADAS.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADPARADAS.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADPARADAS"
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
Dim objCADAPARADAS      As Object
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
       
    If cTipOper = "I" Then objCADAPARADAS.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
       
    objCADAPARADAS.CODINT = "'" & Trim(txtCODINT.Text) & "'"
    objCADAPARADAS.DESCRI = "'" & Trim(txtDescricao.Text) & "'"
       
    If optAtivoSN(0).Value = True Then objCADAPARADAS.ATIVO = 0
    If optAtivoSN(1).Value = True Then objCADAPARADAS.ATIVO = 1
       
    If objCADAPARADAS.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A Parada foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADAPARADAS = CreateObject("CADPARADAS.clsCADPARADAS")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objCADAPARADAS.FILIAL = FILIAL
   
    strCAPTION = "Cadastro de Paradas - "
   
    Call IniciaForm

End Sub

Private Sub Destroy_Objetos()
    Set objBLBFunc = Nothing
    Set objCADAPARADAS = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objetos
End Sub

Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    If cTipOper = "I" Then iCodigo = 0
    objCADAPARADAS.CODIGO = iCodigo
    
    optAtivoSN(1).Value = True
    
    Call CarregaCampos
    
End Sub


Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
    End If
End Sub

Private Sub txtCODINT_GotFocus()
    objBLBFunc.SelecionaCampos txtCODINT.Name, Me
End Sub

Private Sub txtCODINT_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, Me
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtCODINT.Text)) = 0 Then
        MsgBox "Código Inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtCODINT.SetFocus
        Exit Function
     ElseIf Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição da Parada Inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function


Private Sub CarregaCampos()
    If objCADAPARADAS.Carrega_Campos = True Then
        txtCodigo.Text = objCADAPARADAS.CODIGO
        txtCODINT.Text = objCADAPARADAS.CODINT
        txtDescricao.Text = objCADAPARADAS.DESCRI
        optAtivoSN(objCADAPARADAS.ATIVO).Value = True
    End If
End Sub

