VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRELFORNEC 
   Caption         =   "Relatório de Fornecedores"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stFornec 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Normal"
      TabPicture(0)   =   "frmRELFORNEC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Grupo de Risco"
      TabPicture(1)   =   "frmRELFORNEC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "IQF"
      TabPicture(2)   =   "frmRELFORNEC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Não Conformidade"
      TabPicture(3)   =   "frmRELFORNEC.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   9375
         Begin VB.TextBox txtIQFINI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   30
            Text            =   "txtIQFINI"
            Top             =   255
            Width           =   975
         End
         Begin VB.TextBox txtIQFFIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            TabIndex        =   29
            Text            =   "txtIQFFIN"
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "IQF Inicial"
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
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "IQF Final"
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
            Left            =   2520
            TabIndex        =   31
            Top             =   255
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
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
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   3855
         Begin VB.OptionButton optOrdemFornec 
            Caption         =   "Razão Social"
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
            Index           =   1
            Left            =   1800
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optOrdemFornec 
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
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
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   -74880
         TabIndex        =   22
         Top             =   1560
         Width           =   2895
         Begin VB.OptionButton optORDEMRISCO 
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
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optORDEMRISCO 
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   9375
         Begin VB.TextBox txtCODRISCOFIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   19
            Text            =   "txtCODRISCOFIN"
            Top             =   615
            Width           =   975
         End
         Begin VB.ComboBox cboRiscoFIN 
            Height          =   315
            Left            =   3120
            TabIndex        =   18
            Text            =   "cboRiscoFIN"
            Top             =   615
            Width           =   6135
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2760
            Picture         =   "frmRELFORNEC.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCODRISCOINI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   16
            Text            =   "txtCODRISCOINI"
            Top             =   255
            Width           =   975
         End
         Begin VB.ComboBox cboRiscoINI 
            Height          =   315
            Left            =   3120
            TabIndex        =   15
            Text            =   "cboRiscoINI"
            Top             =   255
            Width           =   6135
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2760
            Picture         =   "frmRELFORNEC.frx":0172
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Grupo Risco Final"
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
            TabIndex        =   21
            Top             =   615
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Grupo Risco Inicial"
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
            TabIndex        =   20
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   9375
         Begin VB.CommandButton cmdPesqFor 
            Height          =   315
            Left            =   2760
            Picture         =   "frmRELFORNEC.frx":0274
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboFornecINI 
            Height          =   315
            Left            =   3120
            TabIndex        =   9
            Text            =   "cboFornecINI"
            Top             =   255
            Width           =   6135
         End
         Begin VB.TextBox txtCODFORNECINI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   8
            Text            =   "txtCODFORNECINI"
            Top             =   255
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2760
            Picture         =   "frmRELFORNEC.frx":0376
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   600
            Width           =   375
         End
         Begin VB.ComboBox cboFornecFIN 
            Height          =   315
            Left            =   3120
            TabIndex        =   6
            Text            =   "cboFornecFIN"
            Top             =   615
            Width           =   6135
         End
         Begin VB.TextBox txtCODFORNECFIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Text            =   "txtCODFORNECFIN"
            Top             =   615
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Fornecedor Inicial"
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
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Fornecedor Final"
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
            TabIndex        =   11
            Top             =   615
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
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
         Picture         =   "frmRELFORNEC.frx":0478
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmRELFORNEC.frx":057A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELFORNEC"
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
Dim objRELSUPRI     As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Private Sub cboFornecFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornecFIN, KeyAscii
End Sub

Private Sub cboFornecFIN_Validate(Cancel As Boolean)
    If cboFornecFIN.ListIndex > -1 Then txtCODFORNECFIN.Text = cboFornecFIN.ItemData(cboFornecFIN.ListIndex)
End Sub

Private Sub cboFornecINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornecINI, KeyAscii
End Sub

Private Sub cboFornecINI_Validate(Cancel As Boolean)
    If cboFornecINI.ListIndex > -1 Then txtCODFORNECINI.Text = cboFornecINI.ItemData(cboFornecINI.ListIndex)
End Sub

Private Sub cboRiscoFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboRiscoFIN, KeyAscii
End Sub

Private Sub cboRiscoFIN_Validate(Cancel As Boolean)
    If cboRiscoFIN.ListIndex > -1 Then txtCODRISCOFIN.Text = cboRiscoFIN.ItemData(cboRiscoFIN.ListIndex)
End Sub

Private Sub cboRiscoINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboRiscoINI, KeyAscii
End Sub

Private Sub cboRiscoINI_Validate(Cancel As Boolean)
    If cboRiscoINI.ListIndex > -1 Then txtCODRISCOINI.Text = cboRiscoINI.ItemData(cboRiscoINI.ListIndex)
End Sub

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    If stFornec.Tab = 0 Then Call ImprimirForn
    If stFornec.Tab = 1 Then Call ImprimirRiscoForn
    If stFornec.Tab = 2 Then Call ImprimirFornIQF
End Sub

Private Sub cmdPesqFor_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADFORNEC"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNECINI.Text = varRETORNO
    
    cboFornecINI.ListIndex = -1
    txtCODFORNECINI.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELSUPRI = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADFORNEC"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNECFIN.Text = varRETORNO
    
    cboFornecFIN.ListIndex = -1
    txtCODFORNECFIN.SetFocus

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADRISCO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Risco")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODRISCOINI.Text = varRETORNO
    
    cboRiscoINI.ListIndex = -1
    txtCODRISCOINI.SetFocus

End Sub

Private Sub Command3_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADRISCO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Risco")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODRISCOFIN.Text = varRETORNO
    
    cboRiscoFIN.ListIndex = -1
    txtCODRISCOFIN.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    
    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELSUPRI = CreateObject("RELSUPRI.clsRELFORNEC")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELFORNEC
    
    objRELSUPRI.FILIAL = FILIAL
    
    objRELSUPRI.PreencheComboFornec cboFornecINI
    objRELSUPRI.PreencheComboFornec cboFornecFIN
    
    objRELSUPRI.PreencheComboRiscoFornec cboRiscoINI
    objRELSUPRI.PreencheComboRiscoFornec cboRiscoFIN
    
    stFornec.Tab = 0
    
    optOrdemFornec(0).Value = True
    optORDEMRISCO(0).Value = True
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
End Sub

Private Sub txtCODFORNECFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNECFIN.Name, frmRELFORNEC
End Sub

Private Sub txtCODFORNECFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFORNECFIN.Text
End Sub

Private Sub txtCODFORNECFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNECFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNECFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFornecFIN.ListIndex = -1
    For I = 0 To (cboFornecFIN.ListCount - 1)
        If cboFornecFIN.ItemData(I) = Str(Val(txtCODFORNECFIN.Text)) Then cboFornecFIN.ListIndex = I
    Next I
    
    If cboFornecFIN.ListIndex = -1 Then
       MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODFORNECINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNECINI.Name, frmRELFORNEC
End Sub

Private Sub txtCODFORNECINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFORNECINI.Text
End Sub

Private Sub txtCODFORNECINI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNECINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNECINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFornecINI.ListIndex = -1
    For I = 0 To (cboFornecINI.ListCount - 1)
        If cboFornecINI.ItemData(I) = Str(Val(txtCODFORNECINI.Text)) Then cboFornecINI.ListIndex = I
    Next I
    
    If cboFornecINI.ListIndex = -1 Then
       MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODRISCOFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODRISCOFIN.Name, frmRELFORNEC
End Sub

Private Sub txtCODRISCOFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODRISCOFIN.Text
End Sub

Private Sub txtCODRISCOFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODRISCOFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODRISCOFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODRISCOFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    cboRiscoFIN.ListIndex = -1
    For I = 0 To (cboRiscoFIN.ListCount - 1)
        If cboRiscoFIN.ItemData(I) = Str(Val(txtCODRISCOFIN.Text)) Then cboRiscoFIN.ListIndex = I
    Next I
    
    If cboRiscoFIN.ListIndex = -1 Then
       MsgBox "Este risco não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODRISCOFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODRISCOINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODRISCOINI.Name, frmRELFORNEC
End Sub

Private Sub txtCODRISCOINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODRISCOINI.Text
End Sub

Private Sub txtCODRISCOINI_Validate(Cancel As Boolean)
    
    Dim I As Integer
    
    If Len(Trim(txtCODRISCOINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODRISCOINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODRISCOINI.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    cboRiscoINI.ListIndex = -1
    For I = 0 To (cboRiscoINI.ListCount - 1)
        If cboRiscoINI.ItemData(I) = Str(Val(txtCODRISCOINI.Text)) Then cboRiscoINI.ListIndex = I
    Next I
    
    If cboRiscoINI.ListIndex = -1 Then
       MsgBox "Este risco não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODRISCOINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub ImprimirRiscoForn()
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADRISCOFORNEC.SGI_CODIGO "
    sSql = sSql & "     , SGI_CADRISCO.SGI_DESCRICAO "
    sSql = sSql & "     , SGI_CADRISCOFORNEC.SGI_CODFORNEC "
    sSql = sSql & "     , SGI_CADFORNEC.SGI_RAZAOSOC "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & "     , SGI_CADRISCO SGI_CADRISCO "
    sSql = sSql & "     , SGI_CADRISCOFORNEC SGI_CADRISCOFORNEC "
    sSql = sSql & " Where "
    
    sSql = sSql & "       SGI_CADFORNEC.SGI_FILIAL = SGI_CADRISCOFORNEC.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADFORNEC.SGI_CODIGO = SGI_CADRISCOFORNEC.SGI_CODFORNEC "
    sSql = sSql & "   And SGI_CADRISCOFORNEC.SGI_FILIAL = SGI_CADRISCO.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADRISCOFORNEC.SGI_CODIGO = SGI_CADRISCO.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_CADRISCOFORNEC.SGI_FILIAL = " & FILIAL
    If Len(Trim(txtCODRISCOINI.Text)) > 0 And Len(Trim(txtCODRISCOFIN.Text)) = 0 Then
        sSql = sSql & "   And SGI_CADRISCOFORNEC.SGI_CODIGO = " & txtCODRISCOINI.Text
    ElseIf Len(Trim(txtCODRISCOINI.Text)) > 0 And Len(Trim(txtCODRISCOFIN.Text)) > 0 Then
        sSql = sSql & "   And (SGI_CADRISCOFORNEC.SGI_CODIGO >= " & txtCODRISCOINI.Text & " And SGI_CADRISCOFORNEC.SGI_CODIGO <= " & txtCODRISCOFIN.Text & ")"
    End If
    
    sSql = sSql & " Order by "
    If optORDEMRISCO(0).Value = True Then
       sSql = sSql & "          SGI_CADRISCOFORNEC.SGI_CODIGO "
       strCABEC2 = "Por Ordem de Código"
    ElseIf optORDEMRISCO(1).Value = True Then
       sSql = sSql & "          SGI_CADRISCO.SGI_DESCRICAO "
       strCABEC2 = "Por Ordem de Descrição"
    End If
    
    strCABEC1 = "Relatório de Risco de Fornecedores"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    '' Chamada do Relatório
    Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELRISCOFORN.rpt", Linha, 1, strCABEC1, strCABEC2, True)
   
End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
    If stFornec.Tab = 0 Then
        If Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
           If CLng(txtCODFORNECINI.Text) > CLng(txtCODFORNECFIN.Text) Then
              MsgBox "Fornecedor Inicial não pode ser maior que Fornecedor Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODFORNECINI.SetFocus
              Exit Function
           End If
        End If
    ElseIf stFornec.Tab = 1 Then
        If Len(Trim(txtCODRISCOINI.Text)) > 0 And Len(Trim(txtCODRISCOFIN.Text)) > 0 Then
           If CLng(txtCODRISCOINI.Text) > CLng(txtCODRISCOFIN.Text) Then
              MsgBox "Risco Inicial não pode ser maior que Risco Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODRISCOINI.SetFocus
              Exit Function
           End If
        End If
    ElseIf stFornec.Tab = 2 Then
        If Len(Trim(txtIQFINI.Text)) > 0 And Len(Trim(txtIQFFIN.Text)) > 0 Then
           If CLng(txtIQFINI.Text) > CLng(txtIQFFIN.Text) Then
              MsgBox "IQF Inicial não pode ser maior que IQF Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtIQFINI.SetFocus
              Exit Function
           End If
        End If
    End If
    ConfereCampos = True

End Function


Private Sub ImprimirForn()
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADFORNEC.SGI_CODIGO "
    sSql = sSql & "     , SGI_CADFORNEC.SGI_RAZAOSOC "
    sSql = sSql & "     , SGI_CADFORNEC.SGI_IQF "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADFORNEC.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) = 0 Then
        sSql = sSql & "   And SGI_CADFORNEC.SGI_CODIGO = " & txtCODFORNECINI.Text
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
        sSql = sSql & "   And (SGI_CADFORNEC.SGI_CODIGO >= " & txtCODFORNECINI.Text & " And SGI_CADFORNEC.SGI_CODIGO <= " & txtCODFORNECFIN.Text & ")"
    End If
    
    sSql = sSql & " Order by "
    If optOrdemFornec(0).Value = True Then
       sSql = sSql & "          SGI_CADFORNEC.SGI_CODIGO "
       strCABEC2 = "Por Ordem de Código"
    ElseIf optOrdemFornec(1).Value = True Then
       sSql = sSql & "          SGI_CADFORNEC.SGI_RAZAOSOC "
       strCABEC2 = "Por Ordem de Razão Social"
    End If
    
    strCABEC1 = "Relatório de Fornecedores"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    '' Chamada do Relatório
    Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELFORNEC.rpt", Linha, 1, strCABEC1, strCABEC2, False)
   
End Sub

Private Sub txtIQFFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtIQFFIN.Name, frmRELFORNEC
End Sub

Private Sub txtIQFFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtIQFFIN.Text
End Sub

Private Sub txtIQFINI_GotFocus()
    objBLBFunc.SelecionaCampos txtIQFINI.Name, frmRELFORNEC
End Sub

Private Sub txtIQFINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtIQFINI.Text
End Sub

Private Sub ImprimirFornIQF()
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADFORNEC.SGI_IQF "
    sSql = sSql & "     , SGI_CADFORNEC.SGI_CODIGO "
    sSql = sSql & "     , SGI_CADFORNEC.SGI_RAZAOSOC "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADFORNEC.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtIQFINI.Text)) > 0 And Len(Trim(txtIQFFIN.Text)) = 0 Then
        sSql = sSql & "   And SGI_CADFORNEC.SGI_IQF = " & txtIQFINI.Text
    ElseIf Len(Trim(txtIQFINI.Text)) > 0 And Len(Trim(txtIQFFIN.Text)) > 0 Then
        sSql = sSql & "   And (SGI_CADFORNEC.SGI_IQF >= " & txtIQFINI.Text & " And SGI_CADFORNEC.SGI_IQF <= " & txtIQFFIN.Text & ")"
    End If
    
    strCABEC1 = "Relatório de IQF dos Fornecedores"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    '' Chamada do Relatório
    Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELFORNECIQF.rpt", Linha, 1, strCABEC1, strCABEC2, True)
   
End Sub

