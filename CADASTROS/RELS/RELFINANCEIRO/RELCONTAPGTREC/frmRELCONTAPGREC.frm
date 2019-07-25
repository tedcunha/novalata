VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELCONTAPGREC 
   Caption         =   "Relatório de contas a pagar no periodo"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7050
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   6975
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   5160
         TabIndex        =   23
         Top             =   1920
         Width           =   1695
         Begin VB.OptionButton optTipRel 
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
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optTipRel 
            Caption         =   "Analitico"
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
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   5055
         Begin VB.OptionButton optABERTOPAGO 
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
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   680
            Width           =   1215
         End
         Begin VB.ComboBox cboTipos 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   240
            Width           =   2895
         End
         Begin VB.OptionButton optABERTOPAGO 
            Caption         =   "Pago"
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
            Left            =   120
            TabIndex        =   22
            Top             =   400
            Width           =   1215
         End
         Begin VB.OptionButton optABERTOPAGO 
            Caption         =   "Em Aberto"
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
            TabIndex        =   21
            Top             =   150
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Agrupamento"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   6735
         Begin VB.OptionButton optDataForGrup 
            Caption         =   "Grupo de despesa"
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
            Index           =   2
            Left            =   2760
            TabIndex        =   19
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optDataForGrup 
            Caption         =   "Fornecedor"
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
            Left            =   1200
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optDataForGrup 
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtCODGRUPDESP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "txtCODGRUPDESP"
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboGrupDesp 
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Text            =   "cboGrupDesp"
         Top             =   960
         Width           =   4455
      End
      Begin VB.CommandButton cmdPesqGrupDesp 
         Height          =   315
         Left            =   2040
         Picture         =   "frmRELCONTAPGREC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCODFORNEC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCODFORNEC"
         Top             =   615
         Width           =   735
      End
      Begin VB.ComboBox cboFornec 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Text            =   "cboFornec"
         Top             =   615
         Width           =   4455
      End
      Begin VB.CommandButton cmdPesqFor 
         Height          =   315
         Left            =   2040
         Picture         =   "frmRELCONTAPGREC.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDtFinal 
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDtInicial 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Desp:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   12
         Top             =   650
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Finial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   11
         Top             =   280
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   280
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6975
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
         Picture         =   "frmRELCONTAPGREC.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmRELCONTAPGREC.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELCONTAPGREC"
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
Dim objCADCONTASAPG As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object
''Dim cCamRel         As String

Private Sub cboFornec_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornec, KeyAscii
End Sub

Private Sub cboFornec_Validate(Cancel As Boolean)
    If cboFornec.ListIndex > -1 Then txtCODFORNEC.Text = cboFornec.ItemData(cboFornec.ListIndex)
End Sub

Private Sub cboGrupDesp_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboGrupDesp, KeyAscii
End Sub

Private Sub cboGrupDesp_Validate(Cancel As Boolean)
    If cboGrupDesp.ListIndex > -1 Then txtCODGRUPDESP.Text = cboGrupDesp.ItemData(cboGrupDesp.ListIndex)
End Sub

Private Sub cmdImpressao_Click()

    Dim strTitulo As String
    
    If Valida_Campos = False Then Exit Sub
    
    If CDate(mskDtInicial.Text) <> CDate(mskDtFinal.Text) Then
       strTitulo = "Relatório de contas a pagar no periodo de " & mskDtInicial.Text & " á " & mskDtFinal.Text
    Else
       strTitulo = "Relatório de contas a pagar no dia " & mskDtInicial.Text
    End If
   
    If optABERTOPAGO(0).Value = True And optTipRel(0).Value = True Then Abertos strTitulo
    If optABERTOPAGO(0).Value = True And optTipRel(1).Value = True Then AbertosAna strTitulo
    
    If optABERTOPAGO(1).Value = True Then Pagos strTitulo
    
    mskDtInicial.SetFocus
    
End Sub

Private Sub cmdPesqFor_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADFORNEC"
    
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
    arrCAMPOS(3, 4) = "3000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNEC.Text = varRETORNO
    
    cboFornec.ListIndex = -1
    txtCODFORNEC.SetFocus

End Sub

Private Sub cmdPesqGrupDesp_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADGRUPDESP"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de despesas")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODGRUPDESP.Text = varRETORNO
    
    cboGrupDesp.ListIndex = -1
    txtCODGRUPDESP.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADCONTASAPG = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
    Unload Me
End Sub

Private Sub Form_Activate()
    mskDtInicial.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADCONTASAPG = CreateObject("RELCONTAPGTREC.clsRELCONTAPGTREC")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    mskDtInicial.Text = Format(Now, "DD/MM/YYYY")
    mskDtFinal.Text = Format(Now, "DD/MM/YYYY")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objCADCONTASAPG.FILIAL = FILIAL
    
    objBLBFunc.LimpaCampos frmRELCONTAPGREC
    
    objCADCONTASAPG.PreencheComboFornec cboFornec
    objCADCONTASAPG.PreencheCombo cboGrupDesp, "SGI_CADGRUPDESP"
    
    optDataForGrup(0).Value = True
    optABERTOPAGO(0).Value = True
    optTipRel(0).Value = True
    
    '' --------------------------------------
    '' Configura Combo
    cboTipos.Clear

    cboTipos.AddItem "Todos"
    cboTipos.ItemData(cboTipos.NewIndex) = 1
    
    cboTipos.AddItem "Só Vencidos"
    cboTipos.ItemData(cboTipos.NewIndex) = 2
    
    cboTipos.AddItem "Não Vencidos"
    cboTipos.ItemData(cboTipos.NewIndex) = 3

    cboTipos.ListIndex = 0
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    '' --------------------------------------
    ''cCamRel = "\\pc6\HD\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\CONTASAPG\"
    ''cCamRel = "C:\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\CONTASAPG\"
End Sub


Private Sub mskDtFinal_GotFocus()
    objBLBFunc.SelecionaCampos mskDtFinal.Name, frmRELCONTAPGREC
End Sub

Private Sub mskDtInicial_GotFocus()
    objBLBFunc.SelecionaCampos mskDtInicial.Name, frmRELCONTAPGREC
End Sub


Private Sub optABERTOPAGO_Click(Index As Integer)

    If Index = 0 Then
       
       '' Configura Combo
       cboTipos.Enabled = True
       cboTipos.Clear
    
       cboTipos.AddItem "Todos"
       cboTipos.ItemData(cboTipos.NewIndex) = 1
        
       cboTipos.AddItem "Só Vencidos"
       cboTipos.ItemData(cboTipos.NewIndex) = 2
        
       cboTipos.AddItem "Não Vencidos"
       cboTipos.ItemData(cboTipos.NewIndex) = 3
    
       cboTipos.ListIndex = 0
       
    End If

    If Index = 1 Then
       
       '' Configura Combo
       cboTipos.Enabled = True
       cboTipos.Clear
    
       cboTipos.AddItem "Todos"
       cboTipos.ItemData(cboTipos.NewIndex) = 1
        
       cboTipos.AddItem "Com Atrazo"
       cboTipos.ItemData(cboTipos.NewIndex) = 2
        
       cboTipos.AddItem "Adiantado"
       cboTipos.ItemData(cboTipos.NewIndex) = 3
    
       cboTipos.ListIndex = 0
       
    End If
    
    If Index = 2 Then
       
       cboTipos.Enabled = False
       cboTipos.Clear
    
    End If

End Sub

Private Sub txtCODFORNEC_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNEC.Name, frmRELCONTAPGREC
End Sub

Private Sub txtCODFORNEC_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNEC.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNEC.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNEC.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFornec.ListIndex = -1
    For I = 0 To (cboFornec.ListCount - 1)
        If cboFornec.ItemData(I) = Str(Val(txtCODFORNEC.Text)) Then cboFornec.ListIndex = I
    Next I
    
    If cboFornec.ListIndex = -1 Then
       MsgBox "Esta fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNEC.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODGRUPDESP_GotFocus()
    objBLBFunc.SelecionaCampos txtCODGRUPDESP.Name, frmRELCONTAPGREC
End Sub

Private Sub txtCODGRUPDESP_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODGRUPDESP.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODGRUPDESP.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUPDESP.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboGrupDesp.ListIndex = -1
    For I = 0 To (cboGrupDesp.ListCount - 1)
        If cboGrupDesp.ItemData(I) = Str(Val(txtCODGRUPDESP.Text)) Then cboGrupDesp.ListIndex = I
    Next I
    
    If cboGrupDesp.ListIndex = -1 Then
       MsgBox "Esta Condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUPDESP.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub Pagos(strTitulo As String)

   Dim I         As Integer
   Dim j         As Integer
   Dim strCABEC2 As String
     
   sSql = "Select "
   sSql = sSql & "       SGI_CONTASIAPG.SGI_DTPGTO     "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_DATAVENC   "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_NUMDOC     "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_PARCELA    "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLDOC      "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLPAGO     "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLDESC     "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLACRES    "
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_CODIGO     "
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_FILIAL     "
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_QTDPARC    "
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_CODFOR     "
   sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC    "
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_GRPDESP    "
   sSql = sSql & "      ,SGI_CADGRUPDESP.SGI_DESCRICAO "
   sSql = sSql & "  From "
   sSql = sSql & "       SGI_CONTASIAPG  "
   sSql = sSql & "      ,SGI_CONTASHAPG  "
   sSql = sSql & "      ,SGI_CADFORNEC   "
   sSql = sSql & "      ,SGI_CADGRUPDESP "
   sSql = sSql & " Where "
   sSql = sSql & "       SGI_CONTASIAPG.SGI_STATUS = 'B' "
   sSql = sSql & "   And SGI_CONTASIAPG.SGI_FILIAL = " & FILIAL
   sSql = sSql & "   And SGI_CONTASIAPG.SGI_DTPGTO Between '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' "
   sSql = sSql & "   And SGI_CONTASHAPG.SGI_FILIAL  = SGI_CONTASIAPG.SGI_FILIAL  "
   sSql = sSql & "   And SGI_CONTASHAPG.SGI_CODIGO  = SGI_CONTASIAPG.SGI_CODIGO  "

   If Len(Trim(txtCODFORNEC.Text)) > 0 Then
      sSql = sSql & "  And SGI_CONTASHAPG.SGI_CODFOR     = " & txtCODFORNEC.Text & vbCrLf
   End If
   If Len(Trim(txtCODGRUPDESP.Text)) > 0 Then
      sSql = sSql & "  And SGI_CONTASHAPG.SGI_GRPDESP    = " & txtCODGRUPDESP.Text & vbCrLf
   End If

   sSql = sSql & "  And SGI_CADFORNEC.SGI_FILIAL      = SGI_CONTASHAPG.SGI_FILIAL  "
   sSql = sSql & "  And SGI_CADFORNEC.SGI_CODIGO      = SGI_CONTASHAPG.SGI_CODFOR  "
   sSql = sSql & "  And SGI_CADGRUPDESP.SGI_FILIAL    = SGI_CONTASHAPG.SGI_FILIAL  "
   sSql = sSql & "  And SGI_CADGRUPDESP.SGI_CODIGO    = SGI_CONTASHAPG.SGI_GRPDESP "

   If cboTipos.ItemData(cboTipos.ListIndex) = 2 Then
      sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIAPG.SGI_DTPGTO, GETDATE()) * -1) < 0 "
   End If
   If cboTipos.ItemData(cboTipos.ListIndex) = 3 Then
      sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIAPG.SGI_DTPGTO, GETDATE()) * -1) >= 0 "
   End If

   sSql = sSql & " Order by SGI_CONTASIAPG.SGI_DTPGTO "

   BREC.Open sSql, adoBanco_Dados, adOpenDynamic

   If BREC.EOF Then
      BREC.Close
      MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
      Exit Sub
   End If
   
   BREC.Close
   
   If optDataForGrup(0).Value = True Then
      strCABEC2 = "Pagos por agrupamento de data de vencimento - (" & cboTipos.Text & ")"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELPGDTP2.rpt", Linha, 2, strTitulo, strCABEC2, True
   End If
   If optDataForGrup(1).Value = True Then
      strCABEC2 = "Pagos por agrupamento de fornecedores - (" & cboTipos.Text & ")"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELPGFORN2.rpt", Linha, 2, strTitulo, strCABEC2, True
   End If
   If optDataForGrup(2).Value = True Then
      strCABEC2 = "Pagos por agrupamento de grupo de despesas - (" & cboTipos.Text & ")"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELPGDSP2.rpt", Linha, 2, strTitulo, strCABEC2, True
   End If

End Sub

Private Function Valida_Campos() As Boolean

    Valida_Campos = False
    
    If Not IsDate(mskDtInicial.Text) Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDtInicial.SetFocus
       Exit Function
    End If
    If Not IsDate(mskDtFinal.Text) Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDtFinal.SetFocus
       Exit Function
    End If
    If CDate(mskDtInicial.Text) > CDate(mskDtFinal.Text) Then
       MsgBox "Data inicial não pode ser maior que data final !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDtFinal.SetFocus
       Exit Function
    End If
    
    Valida_Campos = True

End Function

Private Sub Abertos(strTitulo As String)

   Dim I         As Integer
   Dim j         As Integer
   Dim strCABEC2 As String
   
   sSql = "Select "
   
   sSql = sSql & "       SGI_CONTASIAPG.SGI_DATAVENC "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_NUMDOC "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_PARCELA "
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_QTDPARC "
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLDOC "
   sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC "
   sSql = sSql & "      ,SGI_CADGRUPDESP.SGI_DESCRICAO "
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_CODIGO "
   
   sSql = sSql & "  From "
   sSql = sSql & "        SGI_CADFORNEC SGI_CADFORNEC "
   sSql = sSql & "      , SGI_CADGRUPDESP SGI_CADGRUPDESP "
   sSql = sSql & "      , SGI_CONTASHAPG SGI_CONTASHAPG "
   sSql = sSql & "      , SGI_CONTASIAPG SGI_CONTASIAPG "
   
   sSql = sSql & " Where "
   sSql = sSql & "       SGI_CONTASHAPG.SGI_FILIAL = SGI_CADGRUPDESP.SGI_FILIAL "
   sSql = sSql & "   And SGI_CONTASHAPG.SGI_GRPDESP = SGI_CADGRUPDESP.SGI_CODIGO "
   sSql = sSql & "   And SGI_CONTASHAPG.SGI_FILIAL = SGI_CADFORNEC.SGI_FILIAL "
   sSql = sSql & "   And SGI_CONTASHAPG.SGI_CODFOR = SGI_CADFORNEC.SGI_CODIGO "
   sSql = sSql & "   And SGI_CONTASIAPG.SGI_FILIAL = SGI_CONTASHAPG.SGI_FILIAL "
   sSql = sSql & "   And SGI_CONTASIAPG.SGI_CODIGO = SGI_CONTASHAPG.SGI_CODIGO "
   
   sSql = sSql & "   And SGI_CONTASIAPG.SGI_STATUS = 'A' "
   sSql = sSql & "   And SGI_CONTASIAPG.SGI_FILIAL = " & FILIAL
   sSql = sSql & "   And SGI_CONTASIAPG.SGI_DATAVENC >= '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And SGI_CONTASIAPG.SGI_DATAVENC <= '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' "
   
   If Len(Trim(txtCODFORNEC.Text)) > 0 Then
      sSql = sSql & "  And SGI_CONTASHAPG.SGI_CODFOR = " & txtCODFORNEC.Text & vbCrLf
   End If
   If Len(Trim(txtCODGRUPDESP.Text)) > 0 Then
      sSql = sSql & "  And SGI_CONTASHAPG.SGI_GRPDESP = " & txtCODGRUPDESP.Text & vbCrLf
   End If
   
   If cboTipos.ItemData(cboTipos.ListIndex) = 2 Then
      sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIAPG.SGI_DATAVENC, GETDATE()) * -1) < 0 "
   End If
   If cboTipos.ItemData(cboTipos.ListIndex) = 3 Then
      sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIAPG.SGI_DATAVENC, GETDATE()) * -1) >= 0 "
   End If
   
   If optDataForGrup(0).Value = True Then
      sSql = sSql & " Order by SGI_CONTASIAPG.SGI_DATAVENC "
   End If
   If optDataForGrup(1).Value = True Then
      sSql = sSql & " Order by SGI_CADFORNEC.SGI_RAZAOSOC,SGI_CONTASIAPG.SGI_DATAVENC "
   End If
   If optDataForGrup(2).Value = True Then
      sSql = sSql & " Order by SGI_CADGRUPDESP.SGI_DESCRICAO,SGI_CONTASIAPG.SGI_DATAVENC "
   End If
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If BREC.EOF Then
      BREC.Close
      MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
      Exit Sub
   End If
   
   BREC.Close
   
   If optDataForGrup(0).Value = True Then
      strCABEC2 = "Em aberto por agrupamento de data de vencimento - (" & cboTipos.Text & ") - Análitico"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELAPGDTV2.rpt", Linha, 1, strTitulo, strCABEC2, True
   End If
   If optDataForGrup(1).Value = True Then
      strCABEC2 = "Em aberto por agrupamento de fornecedores - (" & cboTipos.Text & ") - Análitico"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELAPGFORN2.rpt", Linha, 1, strTitulo, strCABEC2, True
   End If
   If optDataForGrup(2).Value = True Then
      strCABEC2 = "Em aberto por agrupamento de grupo de despesas - (" & cboTipos.Text & ") - Análitico"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELAPGDSP2.rpt", Linha, 1, strTitulo, strCABEC2, True
   End If

End Sub

Private Sub AbertosAna(strTitulo As String)

   Dim I         As Integer
   Dim j         As Integer
   Dim strCABEC2 As String
   
   If optDataForGrup(0).Value = True Then
      sSql = "Select "
      sSql = sSql & "       SGI_DATAVENC   "
      sSql = sSql & "      ,SGI_VLDOC      "
      sSql = sSql & "  From "
      sSql = sSql & "       SGI_CONTASIAPG  "
      sSql = sSql & " Where "
      sSql = sSql & "       SGI_DATAVENC >= '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And SGI_DATAVENC <= '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' "
      sSql = sSql & "   And SGI_STATUS = 'A' "
      sSql = sSql & "   And SGI_FILIAL = " & FILIAL
   End If
   
   If optDataForGrup(1).Value = True Then
      sSql = "Select "
      sSql = sSql & "       SGI_CONTASIAPG.SGI_DATAVENC   "
      sSql = sSql & "      ,SGI_CONTASIAPG.SGI_NUMDOC     "
      sSql = sSql & "      ,SGI_CONTASIAPG.SGI_PARCELA    "
      sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLDOC      "
      sSql = sSql & "      ,SGI_CONTASHAPG.SGI_CODIGO     "
      sSql = sSql & "      ,SGI_CONTASHAPG.SGI_FILIAL     "
      sSql = sSql & "      ,SGI_CONTASHAPG.SGI_QTDPARC    "
      sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC    "
      sSql = sSql & "  From "
      sSql = sSql & "       SGI_CONTASIAPG  "
      sSql = sSql & "      ,SGI_CONTASHAPG  "
      sSql = sSql & "      ,SGI_CADFORNEC   "
      sSql = sSql & " Where "
      sSql = sSql & "       SGI_CONTASIAPG.SGI_STATUS = 'A' "
      sSql = sSql & "   And SGI_CONTASIAPG.SGI_FILIAL = " & FILIAL
      sSql = sSql & "   And SGI_CONTASIAPG.SGI_DATAVENC Between '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' "
      sSql = sSql & "   And SGI_CONTASHAPG.SGI_FILIAL  = SGI_CONTASIAPG.SGI_FILIAL  "
      sSql = sSql & "   And SGI_CONTASHAPG.SGI_CODIGO  = SGI_CONTASIAPG.SGI_CODIGO  "
   
      If Len(Trim(txtCODFORNEC.Text)) > 0 Then
         sSql = sSql & "  And SGI_CONTASHAPG.SGI_CODFOR     = " & txtCODFORNEC.Text & vbCrLf
      End If
   
      sSql = sSql & "  And SGI_CADFORNEC.SGI_FILIAL      = SGI_CONTASHAPG.SGI_FILIAL  "
      sSql = sSql & "  And SGI_CADFORNEC.SGI_CODIGO      = SGI_CONTASHAPG.SGI_CODFOR  "
   End If
   
   If optDataForGrup(2).Value = True Then
   
      sSql = "Select "
      sSql = sSql & "       SGI_CONTASIAPG.SGI_DATAVENC   "
      sSql = sSql & "      ,SGI_CONTASIAPG.SGI_NUMDOC     "
      sSql = sSql & "      ,SGI_CONTASIAPG.SGI_PARCELA    "
      sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLDOC      "
      sSql = sSql & "      ,SGI_CONTASHAPG.SGI_CODIGO     "
      sSql = sSql & "      ,SGI_CONTASHAPG.SGI_FILIAL     "
      sSql = sSql & "      ,SGI_CONTASHAPG.SGI_QTDPARC    "
      sSql = sSql & "      ,SGI_CADGRUPDESP.SGI_DESCRICAO "
      sSql = sSql & "  From "
      sSql = sSql & "       SGI_CONTASIAPG  "
      sSql = sSql & "      ,SGI_CONTASHAPG  "
      sSql = sSql & "      ,SGI_CADGRUPDESP "
      sSql = sSql & " Where "
      sSql = sSql & "       SGI_CONTASIAPG.SGI_STATUS = 'A' "
      sSql = sSql & "   And SGI_CONTASIAPG.SGI_FILIAL = " & FILIAL
      sSql = sSql & "   And SGI_CONTASIAPG.SGI_DATAVENC Between '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' "
      sSql = sSql & "   And SGI_CONTASHAPG.SGI_FILIAL  = SGI_CONTASIAPG.SGI_FILIAL  "
      sSql = sSql & "   And SGI_CONTASHAPG.SGI_CODIGO  = SGI_CONTASIAPG.SGI_CODIGO  "
   
      If Len(Trim(txtCODGRUPDESP.Text)) > 0 Then
         sSql = sSql & "  And SGI_CONTASHAPG.SGI_GRPDESP    = " & txtCODGRUPDESP.Text & vbCrLf
      End If
   
      sSql = sSql & "  And SGI_CADGRUPDESP.SGI_FILIAL    = SGI_CONTASHAPG.SGI_FILIAL  "
      sSql = sSql & "  And SGI_CADGRUPDESP.SGI_CODIGO    = SGI_CONTASHAPG.SGI_GRPDESP "
   
   End If
   
   If cboTipos.ItemData(cboTipos.ListIndex) = 2 Then
      sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIAPG.SGI_DATAVENC, GETDATE()) * -1) < 0 "
   End If
   If cboTipos.ItemData(cboTipos.ListIndex) = 3 Then
      sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIAPG.SGI_DATAVENC, GETDATE()) * -1) >= 0 "
   End If
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If BREC.EOF Then
      BREC.Close
      MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
      Exit Sub
   End If
   
   BREC.Close
   
   If optDataForGrup(0).Value = True Then
      strCABEC2 = "Em aberto por agrupamento de data de vencimento - (" & cboTipos.Text & ") - Sintético"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELAPGANDTV2.rpt", Linha, 1, strTitulo, strCABEC2, True
   End If
   If optDataForGrup(1).Value = True Then
      strCABEC2 = "Em aberto por agrupamento de fornecedores - (" & cboTipos.Text & ") - Sintético"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELAPGANFORN2.rpt", Linha, 1, strTitulo, strCABEC2, True
   End If
   If optDataForGrup(2).Value = True Then
      strCABEC2 = "Em aberto por agrupamento de grupo de despesas - (" & cboTipos.Text & ") - Sintético"
      objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & "RELAPGANDSP2.rpt", Linha, 1, strTitulo, strCABEC2, True
   End If

End Sub
