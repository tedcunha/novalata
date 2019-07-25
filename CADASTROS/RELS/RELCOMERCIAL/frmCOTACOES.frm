VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCOTACOES 
   Caption         =   "Relatório de Cotação de Vendas"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   42
      Top             =   960
      Width           =   9975
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2880
         TabIndex        =   43
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "[ Ordem ]"
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
      Left            =   0
      TabIndex        =   41
      Top             =   4800
      Width           =   5055
      Begin VB.OptionButton optOrdem 
         Caption         =   "Tipo Orçamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optOrdem 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optOrdem 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[ Periodo ]"
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
      Left            =   5160
      TabIndex        =   40
      Top             =   3600
      Width           =   4815
      Begin VB.OptionButton optDiaMesAno 
         Caption         =   "Ano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optDiaMesAno 
         Caption         =   "Mês"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optDiaMesAno 
         Caption         =   "Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "[ Cotações Atendidas ]"
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
      Left            =   0
      TabIndex        =   39
      Top             =   4200
      Width           =   5055
      Begin VB.OptionButton optCtAtend 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optCtAtend 
         Caption         =   "Integral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optCtAtend 
         Caption         =   "Parcial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Relatório ]"
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
      Left            =   5160
      TabIndex        =   38
      Top             =   4200
      Width           =   4815
      Begin VB.OptionButton optRELCOTAANSIN 
         Caption         =   "Sintético"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optRELCOTAANSIN 
         Caption         =   "Análitico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Cotações ]"
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
      Left            =   0
      TabIndex        =   37
      Top             =   3600
      Width           =   5055
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Aberto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Atendidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   0
      TabIndex        =   34
      Top             =   2520
      Width           =   9975
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3120
         Picture         =   "frmCOTACOES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cboTIPORCINI 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Text            =   "cboTIPORCINI"
         Top             =   255
         Width           =   6375
      End
      Begin VB.TextBox txtCODTIPORCINI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Text            =   "txtCODTIPORCINI"
         Top             =   255
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   3120
         Picture         =   "frmCOTACOES.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox cboTIPORCFIN 
         Height          =   315
         Left            =   3480
         TabIndex        =   12
         Text            =   "cboTIPORCFIN"
         Top             =   615
         Width           =   6375
      End
      Begin VB.TextBox txtCODTIPORCFIN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Text            =   "txtCODTIPORCFIN"
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo Orçamento Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Orçamento Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   31
      Top             =   1560
      Width           =   9975
      Begin VB.CommandButton cmdPesqCLI 
         Height          =   315
         Left            =   2640
         Picture         =   "frmCOTACOES.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cboCLIEINI 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Text            =   "cboCLIEINI"
         Top             =   255
         Width           =   6855
      End
      Begin VB.TextBox txtCODCLIINI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "txtCODCLIINI"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPesqCLIFIN 
         Height          =   315
         Left            =   2640
         Picture         =   "frmCOTACOES.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox cboCLIEFIN 
         Height          =   315
         Left            =   3000
         TabIndex        =   6
         Text            =   "cboCLIEFIN"
         Top             =   615
         Width           =   6855
      End
      Begin VB.TextBox txtCODCLIFIN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Text            =   "txtCODCLIFIN"
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         Picture         =   "frmCOTACOES.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
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
         Picture         =   "frmCOTACOES.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCOTACOES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Dim objBLBFunc      As Object
Dim objRELCOTACAO   As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object
Dim strCABEC1       As String
Dim strCABEC2       As String

Private Sub cboCLIEFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCLIEFIN, KeyAscii
End Sub

Private Sub cboCLIEFIN_Validate(Cancel As Boolean)
    If cboCLIEFIN.ListIndex > -1 Then txtCODCLIFIN.Text = cboCLIEFIN.ItemData(cboCLIEFIN.ListIndex)
End Sub

Private Sub cboCLIEINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCLIEINI, KeyAscii
End Sub

Private Sub cboCLIEINI_Validate(Cancel As Boolean)
    If cboCLIEINI.ListIndex > -1 Then txtCODCLIINI.Text = cboCLIEINI.ItemData(cboCLIEINI.ListIndex)
End Sub

Private Sub cboTIPORCFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTIPORCFIN, KeyAscii
End Sub

Private Sub cboTIPORCFIN_Validate(Cancel As Boolean)
    If cboTIPORCFIN.ListIndex > -1 Then txtCODTIPORCFIN.Text = cboTIPORCFIN.ItemData(cboTIPORCFIN.ListIndex)
End Sub

Private Sub cboTIPORCINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTIPORCINI, KeyAscii
End Sub

Private Sub cboTIPORCINI_Validate(Cancel As Boolean)
    If cboTIPORCINI.ListIndex > -1 Then txtCODTIPORCINI.Text = cboTIPORCINI.ItemData(cboTIPORCINI.ListIndex)
End Sub

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    
    Call InpRelCotAbertasAn
    
End Sub

Private Sub cmdPesqCLI_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLIINI.Text = varRETORNO
    
    cboCLIEINI.ListIndex = -1
    txtCODCLIINI.SetFocus

End Sub

Private Sub cmdPesqCLIFIN_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLIFIN.Text = varRETORNO
    
    cboCLIEFIN.ListIndex = -1
    txtCODCLIFIN.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELCOTACAO = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADESPORCA Where SGI_FILIAL = " & FILIAL
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "4000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Especie de Orçamentos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODTIPORCFIN.Text = varRETORNO
    
    cboTIPORCFIN.ListIndex = -1
    txtCODTIPORCFIN.SetFocus

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADESPORCA Where SGI_FILIAL = " & FILIAL
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "4000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Especie de Orçamentos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODTIPORCINI.Text = varRETORNO
    
    cboTIPORCINI.ListIndex = -1
    txtCODTIPORCINI.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELCOTACAO = CreateObject("RELCOMERCIAL.clsCOTACOES")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmCOTACOES
    
    objRELCOTACAO.FILIAL = FILIAL

    objRELCOTACAO.PreencheComboClientes cboCLIEINI
    objRELCOTACAO.PreencheComboClientes cboCLIEFIN
    
    objRELCOTACAO.PreencheComboEspOrc cboTIPORCINI
    objRELCOTACAO.PreencheComboEspOrc cboTIPORCFIN
    
    optTdSomSemPed(2).Value = True
    optCtAtend(0).Value = True
    optRELCOTAANSIN(0).Value = True
    optDiaMesAno(0).Value = True
    optOrdem(0).Value = True
    
    mskDTINI.Text = Format(Date, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Date + 30, "DD/MM/YYYY")

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmCOTACOES
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmCOTACOES
End Sub

Private Sub txtCODCLIFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLIFIN.Name, frmCOTACOES
End Sub

Private Sub txtCODCLIFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLIFIN.Text
End Sub

Private Sub txtCODCLIFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCLIFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLIFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCLIEFIN.ListIndex = -1
    For I = 0 To (cboCLIEFIN.ListCount - 1)
        If cboCLIEFIN.ItemData(I) = Str(Val(txtCODCLIFIN.Text)) Then cboCLIEFIN.ListIndex = I
    Next I
    
    If cboCLIEFIN.ListIndex = -1 Then
       MsgBox "Este Cliente não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODCLIINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLIINI.Name, frmCOTACOES
End Sub

Private Sub txtCODCLIINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLIINI.Text
End Sub

Private Sub txtCODCLIINI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCLIINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLIINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCLIEINI.ListIndex = -1
    For I = 0 To (cboCLIEINI.ListCount - 1)
        If cboCLIEINI.ItemData(I) = Str(Val(txtCODCLIINI.Text)) Then cboCLIEINI.ListIndex = I
    Next I
    
    If cboCLIEINI.ListIndex = -1 Then
       MsgBox "Este Cliente não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub
Private Sub txtCODTIPORCFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPORCFIN.Name, frmCOTACOES
End Sub

Private Sub txtCODTIPORCFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODTIPORCFIN.Text
End Sub

Private Sub txtCODTIPORCFIN_Validate(Cancel As Boolean)
    
    Dim I As Integer
    
    If Len(Trim(txtCODTIPORCFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTIPORCFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPORCFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTIPORCFIN.ListIndex = -1
    For I = 0 To (cboTIPORCFIN.ListCount - 1)
        If cboTIPORCFIN.ItemData(I) = Str(Val(txtCODTIPORCFIN.Text)) Then cboTIPORCFIN.ListIndex = I
    Next I
    
    If cboTIPORCFIN.ListIndex = -1 Then
       MsgBox "Esta Espécie de Orçamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPORCFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODTIPORCINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPORCINI.Name, frmCOTACOES
End Sub

Private Sub txtCODTIPORCINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODTIPORCINI.Text
End Sub

Private Sub txtCODTIPORCINI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODTIPORCINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTIPORCINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPORCINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTIPORCINI.ListIndex = -1
    For I = 0 To (cboTIPORCINI.ListCount - 1)
        If cboTIPORCINI.ItemData(I) = Str(Val(txtCODTIPORCINI.Text)) Then cboTIPORCINI.ListIndex = I
    Next I
    
    If cboTIPORCINI.ListIndex = -1 Then
       MsgBox "Esta Espécie de Orçamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPORCINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
        If Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) > 0 Then
           If CLng(txtCODCLIINI.Text) > CLng(txtCODCLIFIN.Text) Then
              MsgBox "Cliente Inicial não pode ser maior que Cliente Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODCLIINI.SetFocus
              Exit Function
           End If
        End If
        If Len(Trim(txtCODTIPORCINI.Text)) > 0 And Len(Trim(txtCODTIPORCFIN.Text)) > 0 Then
           If CLng(txtCODTIPORCINI.Text) > CLng(txtCODTIPORCFIN.Text) Then
              MsgBox "Espécie de Orçamentos Inicial não pode ser maior que Espécie de Orçamento Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODTIPORCINI.SetFocus
              Exit Function
           End If
        End If
        
        If Not IsDate(mskDTINI.Text) Then
            MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
        If Not IsDate(mskDTFIN.Text) Then
            MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTFIN.SetFocus
            Exit Function
        End If
        
        If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
            MsgBox "Data Inicial não pode ser maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
    
    ConfereCampos = True

End Function

Private Sub InpRelCotAbertasAn()

    sSql = "Select "
    sSql = sSql & "       SGI_CADCOTAVENDH.SGI_DATACOTA "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_CODIGO "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_CODCLI "
    sSql = sSql & "     , SGI_CADCLIENTE.SGI_RAZAOSOC "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_VALPROP "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_QTDITENS "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_VLTOT "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_STATUS "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_CODTIPORC "
    sSql = sSql & "     , SGI_CADESPORCA.SGI_DESCRICAO "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADCLIENTE  SGI_CADCLIENTE "
    sSql = sSql & "     , SGI_CADPEDVENDH SGI_CADPEDVENDH "
    sSql = sSql & "     , SGI_CADVENDEDOR SGI_CADVENDEDOR "
    
    sSql = sSql & " Where "
    
    sSql = sSql & "        SGI_CADPEDVENDH.SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL "
    sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO "
    sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_FILIAL = SGI_CADVENDEDOR.SGI_FILIAL "
    sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_CODVEND = SGI_CADVENDEDOR.SGI_CODIGO "
    
    If Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) > 0 Then
       sSql = sSql & "  And   (SGI_CADPEDVENDH.SGI_CODCLI >= " & Trim(txtCODCLIINI.Text) & " And SGI_CADCOTAVENDH.SGI_CODCLI <= " & Trim(txtCODCLIFIN.Text) & ")"
    ElseIf Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) = 0 Then
       sSql = sSql & "  And   SGI_CADCOTAVENDH.SGI_CODCLI = " & Trim(txtCODCLIINI.Text)
    End If
    
    If Len(Trim(txtCODTIPORCINI.Text)) > 0 And Len(Trim(txtCODTIPORCFIN.Text)) > 0 Then
       sSql = sSql & "  And   (SGI_CADCOTAVENDH.SGI_CODTIPORC >= " & Trim(txtCODTIPORCINI.Text) & " And SGI_CADCOTAVENDH.SGI_CODTIPORC <= " & Trim(txtCODTIPORCFIN.Text) & ")"
    ElseIf Len(Trim(txtCODTIPORCINI.Text)) > 0 And Len(Trim(txtCODTIPORCFIN.Text)) = 0 Then
       sSql = sSql & "  And   SGI_CADCOTAVENDH.SGI_CODTIPORC = " & Trim(txtCODTIPORCINI.Text)
    End If
    
    sSql = sSql & "  And   SGI_CADCOTAVENDH.SGI_FILIAL = " & FILIAL
    sSql = sSql & "  And  (SGI_CADCOTAVENDH.SGI_DATACOTA >= '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And SGI_CADCOTAVENDH.SGI_DATACOTA <= '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "')"
    
    If optTdSomSemPed(1).Value = True Then
       If optCtAtend(0).Value = True Then
          sSql = sSql & "  And   SGI_CADCOTAVENDH.SGI_STATUS = 'B'"
       ElseIf optCtAtend(1).Value = True Then
          sSql = sSql & "  And   SGI_CADCOTAVENDH.SGI_STATUS = 'P'"
       End If
    ElseIf optTdSomSemPed(2).Value = True Then
       sSql = sSql & "  And   SGI_CADCOTAVENDH.SGI_STATUS = 'A'"
    End If
    
    If optOrdem(0).Value = True Then
       sSql = sSql & " Order by SGI_CADCOTAVENDH.SGI_DATACOTA "
    ElseIf optOrdem(1).Value = True Then
       sSql = sSql & " Order by SGI_CADCOTAVENDH.SGI_CODCLI "
    ElseIf optOrdem(2).Value = True Then
       sSql = sSql & " Order by SGI_CADCOTAVENDH.SGI_CODTIPORC "
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Relatório de Cotações "
    
    If optTdSomSemPed(1).Value = True Then
       If optCtAtend(0).Value = True Then strCABEC1 = strCABEC1 & " Atendidas Integral "
    End If
    If optTdSomSemPed(2).Value = True Then strCABEC1 = strCABEC1 & " em Aberto "
    If optOrdem(0).Value = True Then strCABEC1 = strCABEC1 & " por Ordem de Data "
    If optRELCOTAANSIN(0).Value = True Then strCABEC1 = strCABEC1 & " [ Análitico ]"
    If optRELCOTAANSIN(1).Value = True Then strCABEC1 = strCABEC1 & " [ Sintético ]"
    
    If optDiaMesAno(0).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text
    ElseIf optDiaMesAno(1).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Mês " & Format(Month(CDate(mskDTINI.Text)), "##00") & "/" & Year(CDate(mskDTINI.Text)) & " ao Mês " & Format(Month(CDate(mskDTFIN.Text)), "##00") & "/" & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "No Mês " & Format(Month(CDate(mskDTINI.Text)), "##00") & "/" & Year(CDate(mskDTFIN.Text))
    ElseIf optDiaMesAno(2).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "Do Ano " & Year(CDate(mskDTINI.Text)) & " ao Ano " & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "No Ano " & Year(CDate(mskDTFIN.Text))
    End If
    
    '' Chamada do Relatório
    If optOrdem(0).Value = True Then
        If optDiaMesAno(0).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCOTVENDAN.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCOTVENDSI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        ElseIf optDiaMesAno(1).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAMDA.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAMDASI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        ElseIf optDiaMesAno(2).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAADA.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAADASI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        End If
    ElseIf optOrdem(1).Value = True Then
        If optDiaMesAno(0).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCOTVENDANCLI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCOTVENDSICLI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        ElseIf optDiaMesAno(1).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAMDACLI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAMDASICLI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        ElseIf optDiaMesAno(2).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAADACLI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAADASICLI.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        End If
    ElseIf optOrdem(2).Value = True Then
        If optDiaMesAno(0).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCOTVENDANORC.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCOTVENDSIORC.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        ElseIf optDiaMesAno(1).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAMDAORC.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAMDASIORC.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        ElseIf optDiaMesAno(2).Value = True Then
           If optRELCOTAANSIN(0).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAADAORC.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           ElseIf optRELCOTAANSIN(1).Value = True Then
              Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELCVAADASIORC.rpt", Linha, 1, strCABEC1, strCABEC2, True)
           End If
        End If
    End If
End Sub
