VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELCONTAREC 
   Caption         =   "Relatório de contas a receber"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   6495
      Begin VB.TextBox txtCODGRPRECEB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1305
         TabIndex        =   6
         Text            =   "txtCODGRPRECEB"
         Top             =   975
         Width           =   735
      End
      Begin VB.ComboBox cboGRPRECEB 
         Height          =   315
         Left            =   2385
         TabIndex        =   7
         Text            =   "cboGRPRECEB"
         Top             =   975
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2025
         Picture         =   "frmRELCONTAREC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   960
         Width           =   375
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de relatório"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   3255
         Begin VB.OptionButton Option3 
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
            Left            =   1560
            TabIndex        =   22
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3480
         TabIndex        =   17
         Top             =   2160
         Width           =   2895
         Begin VB.OptionButton Option2 
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
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboTipos 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Baixado"
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
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   1095
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
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   6255
         Begin VB.OptionButton Option1 
            Caption         =   "Gruipo de Recebimento"
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
            Left            =   3120
            TabIndex        =   27
            Top             =   360
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
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
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
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
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdPesqFor 
         Height          =   315
         Left            =   2040
         Picture         =   "frmRELCONTAREC.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Text            =   "cboCliente"
         Top             =   615
         Width           =   3975
      End
      Begin VB.TextBox txtCODCLI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "txtCODCLI"
         Top             =   615
         Width           =   735
      End
      Begin MSMask.MaskEdBox mskDtFinal 
         Height          =   285
         Left            =   3960
         TabIndex        =   3
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
         TabIndex        =   2
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
         Caption         =   "Grp. Receb:"
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
         TabIndex        =   26
         Top             =   1005
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   500
         TabIndex        =   13
         Top             =   650
         Width           =   660
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
         TabIndex        =   11
         Top             =   285
         Width           =   1050
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
         TabIndex        =   10
         Top             =   285
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
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
         Picture         =   "frmRELCONTAREC.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
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
         Picture         =   "frmRELCONTAREC.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELCONTAREC"
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
Dim objPESQPADRAO   As Object
Dim objRELCONTAREC  As Object
Dim strTitulo       As String
Dim objREL          As Object
''Dim cCamRel         As String
Dim strCABEC2       As String
Dim lngCODOPERACAO  As Long

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCliente, KeyAscii
End Sub

Private Sub cboCliente_Validate(Cancel As Boolean)
    If cboCliente.ListIndex > -1 Then txtCODCLI.Text = cboCliente.ItemData(cboCliente.ListIndex)
End Sub

Private Sub cboGRPRECEB_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboGRPRECEB, KeyAscii
End Sub

Private Sub cboGRPRECEB_Validate(Cancel As Boolean)
    If cboGRPRECEB.ListIndex > -1 Then txtCODGRPRECEB.Text = cboGRPRECEB.ItemData(cboGRPRECEB.ListIndex)
End Sub

Private Sub cmdImpressao_Click()
    
    If Verif_Campos = False Then Exit Sub
    
    If CDate(mskDtInicial.Text) <> CDate(mskDtFinal.Text) Then
       strTitulo = "Relatório de contas a receber no periodo de " & mskDtInicial.Text & " á " & mskDtFinal.Text
    Else
       strTitulo = "Relatório de contas a receber no dia " & mskDtInicial.Text
    End If
    
    If Option2(0).Value = True Then ARECABERTO strTitulo
    If Option2(1).Value = True Then ARECBAIXADO strTitulo
    If Option2(2).Value = True Then ARECTODOS strTitulo
    

End Sub

Private Sub cmdPesqFor_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADCLIENTE"
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLI.Text = varRETORNO
    
    cboCliente.ListIndex = -1
    txtCODCLI.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()


    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPREC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de Recebimento")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODGRPRECEB.Text = varRETORNO
    
    cboGRPRECEB.ListIndex = -1
    txtCODGRPRECEB.SetFocus

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
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objRELCONTAREC = CreateObject("RELCONTAPGTREC.clsRELCONTAREC")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")

    mskDtInicial.Text = Format(Now, "DD/MM/YYYY")
    mskDtFinal.Text = Format(Now, "DD/MM/YYYY")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELCONTAREC
    
    objRELCONTAREC.FILIAL = FILIAL
    
    objRELCONTAREC.PreencheComboCliente cboCliente
    objRELCONTAREC.PreenchComboGrpReceb cboGRPRECEB
    
    
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
    '' --------------------------------------
    
    ''cCamRel = "C:\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\CONTASARC\"
    ''cCamRel = "\\pc6\HD\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\CONTASARC\"
    
    Option1(0).Value = True
    Option2(0).Value = True
    Option3(0).Value = True
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub



Private Sub mskDtFinal_GotFocus()
    objBLBFunc.SelecionaCampos mskDtFinal.Name, frmRELCONTAREC
End Sub

Private Sub mskDtInicial_GotFocus()
    objBLBFunc.SelecionaCampos mskDtInicial.Name, frmRELCONTAREC
End Sub

Private Sub Option2_Click(Index As Integer)

    If Index = 0 Then
       
       '' Configura Combo
       cboTipos.Clear
       cboTipos.Enabled = True
    
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
       cboTipos.Clear
       cboTipos.Enabled = True
    
       cboTipos.AddItem "Todos"
       cboTipos.ItemData(cboTipos.NewIndex) = 1
        
       cboTipos.AddItem "Com Atrazo"
       cboTipos.ItemData(cboTipos.NewIndex) = 2
        
       cboTipos.AddItem "Adiantado"
       cboTipos.ItemData(cboTipos.NewIndex) = 3
    
       cboTipos.ListIndex = 0
       
    End If
    
    If Index = 2 Then
       cboTipos.Clear
       cboTipos.Enabled = False
    End If

End Sub

Private Sub txtCODCLI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLI.Name, frmRELCONTAREC
End Sub

Private Sub txtCODCLI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLI.Text
End Sub

Private Sub txtCODCLI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCLI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLI.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    cboCliente.ListIndex = -1
    For I = 0 To (cboCliente.ListCount - 1)
        If cboCliente.ItemData(I) = Str(Val(txtCODCLI.Text)) Then cboCliente.ListIndex = I
    Next I
    
    If cboCliente.ListIndex = -1 Then
       MsgBox "Esta fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub ARECABERTO(strTitulo As String)
    
    
    sSql = "Select "
    
    sSql = sSql & "       SGI_CONTASIARC.SGI_DATAVENC "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_NUMDOC   "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_PARCELA  "
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_QTDPARC  "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLDOC"
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_FILIAL "
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC "
    sSql = sSql & "      ,SGI_CADGRUPREC.SGI_DESCRI "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE"
    sSql = sSql & "      ,SGI_CADGRUPREC SGI_CADGRUPREC"
    sSql = sSql & "      ,SGI_CONTASHARC SGI_CONTASHARC"
    sSql = sSql & "      ,SGI_CONTASIARC SGI_CONTASIARC"
    
    sSql = sSql & " Where "
    
    sSql = sSql & "       SGI_CONTASIARC.SGI_CODIGO = SGI_CONTASHARC.SGI_CODIGO "
    sSql = sSql & "   And SGI_CONTASIARC.SGI_FILIAL = SGI_CONTASHARC.SGI_FILIAL "
    sSql = sSql & "   And SGI_CONTASIARC.SGI_NUMDOC = SGI_CONTASHARC.SGI_DOCPAI "
    
    sSql = sSql & "   And SGI_CADGRUPREC.SGI_FILIAL = SGI_CONTASHARC.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADGRUPREC.SGI_CODIGO = SGI_CONTASHARC.SGI_CODGRPRECEB "
    
    sSql = sSql & "   And SGI_CONTASHARC.SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL "
    sSql = sSql & "   And SGI_CONTASHARC.SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_CONTASIARC.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And (SGI_CONTASIARC.SGI_DATAVENC >= '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And SGI_CONTASIARC.SGI_DATAVENC <= '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "')"
    sSql = sSql & "   And SGI_CONTASIARC.SGI_VLPAGO IS NULL "
    If cboTipos.ItemData(cboTipos.ListIndex) = 2 Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIARC.SGI_DATAVENC, GETDATE()) * -1) < 0 "
    End If
    If cboTipos.ItemData(cboTipos.ListIndex) = 3 Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIARC.SGI_DATAVENC, GETDATE()) * -1) >= 0 "
    End If
    
    If Len(Trim(txtCODCLI.Text)) > 0 Then
       sSql = sSql & "   And SGI_CONTASHARC.SGI_CODCLI   = " & txtCODCLI.Text & " "
    End If
    
    If Len(Trim(txtCODGRPRECEB.Text)) > 0 Then
       sSql = sSql & "   And SGI_CONTASHARC.SGI_CODGRPRECEB = " & txtCODGRPRECEB.Text & " "
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados há imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close

    If Option1(0).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCABSDTV2.rpt", Linha, 1, strTitulo, "", True
    End If
    If Option1(1).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCABSCLI2.rpt", Linha, 1, strTitulo, strTitulo, True
    End If
    
    If Option1(0).Value = True And Option3(1).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCABNDTV2.rpt", Linha, 1, strTitulo, strTitulo, True
    End If
    If Option1(1).Value = True And Option3(1).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCABNCLI2.rpt", Linha, 1, strTitulo, strTitulo, True
    End If
    
    If Option1(2).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCGRPREC.rpt", Linha, 1, strTitulo, "( Análitico - Abertos )", True
    End If
    If Option1(2).Value = True And Option3(1).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCGRPRECSINT.rpt", Linha, 1, strTitulo, "(Sintético - Abertos)", True
    End If
    
End Sub

Private Function Verif_Campos() As Boolean

    Verif_Campos = False

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
    
    Verif_Campos = True

End Function

Private Sub ARECBAIXADO(strTitulo As String)

    sSql = "Select "
    
    sSql = sSql & "       SGI_CONTASIARC.SGI_NUMDOC   "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_DTPGTO   "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_DATAVENC "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_PARCELA  "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLDOC    "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLPAGO   "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLDESC   "
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLACRES  "
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODIGO   "
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_FILIAL   "
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_QTDPARC  "
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODCLI   "
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_NOMFANTA "
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC "
    sSql = sSql & "      ,SGI_CADGRUPREC.SGI_DESCRI   "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE "
    sSql = sSql & "      ,SGI_CADGRUPREC SGI_CADGRUPREC "
    sSql = sSql & "      ,SGI_CONTASHARC SGI_CONTASHARC "
    sSql = sSql & "      ,SGI_CONTASIARC SGI_CONTASIARC "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CONTASIARC.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And SGI_CONTASIARC.SGI_DTPGTO >= '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And SGI_CONTASIARC.SGI_DTPGTO <= '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' "
    sSql = sSql & "   And SGI_CONTASIARC.SGI_VLPAGO IS NOT NULL "
    
    If Len(Trim(txtCODCLI.Text)) > 0 Then
       sSql = sSql & "   And SGI_CONTASHARC.SGI_CODCLI   = " & txtCODCLI.Text & " "
    End If
    If Len(Trim(txtCODGRPRECEB.Text)) > 0 Then
       sSql = sSql & "   And SGI_CONTASHARC.SGI_CODGRPRECEB = " & txtCODGRPRECEB.Text & " "
    End If
    
    sSql = sSql & "   And SGI_CONTASHARC.SGI_FILIAL   = SGI_CONTASIARC.SGI_FILIAL "
    sSql = sSql & "   And SGI_CONTASHARC.SGI_CODIGO   = SGI_CONTASIARC.SGI_CODIGO "
    sSql = sSql & "   And SGI_CONTASHARC.SGI_DOCPAI   = SGI_CONTASIARC.SGI_NUMDOC "
    
    sSql = sSql & "   And SGI_CADGRUPREC.SGI_FILIAL   = SGI_CONTASHARC.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADGRUPREC.SGI_CODIGO   = SGI_CONTASHARC.SGI_CODGRPRECEB "
    
    sSql = sSql & "   And SGI_CADCLIENTE.SGI_FILIAL   = SGI_CONTASHARC.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADCLIENTE.SGI_CODIGO   = SGI_CONTASHARC.SGI_CODCLI "
    
    If cboTipos.ItemData(cboTipos.ListIndex) = 2 Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIARC.SGI_DATAVENC, SGI_CONTASIARC.SGI_DTPGTO)) > 0 "
    End If
    If cboTipos.ItemData(cboTipos.ListIndex) = 3 Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_CONTASIARC.SGI_DATAVENC, SGI_CONTASIARC.SGI_DTPGTO)) <= 0 "
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados há imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If Option1(0).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCPADTV2.rpt", Linha, 1, strTitulo, "", True
    End If
    If Option1(1).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCPACLI2.rpt", Linha, 1, strTitulo, "", True
    End If
    
    If Option1(0).Value = True And Option3(1).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCPSDTV2.rpt", Linha, 1, strTitulo, "", True
    End If
    If Option1(1).Value = True And Option3(1).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCPSCLI2.rpt", Linha, 1, strTitulo, "", True
    End If
    
    If Option1(2).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCGRPREC.rpt", Linha, 1, strTitulo, "( Análitico - Baixados )", True
    End If
    If Option1(2).Value = True And Option3(1).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCGRPRECSINT.rpt", Linha, 1, strTitulo, "( Sintético - Baixados )", True
    End If
    

End Sub

Private Sub ARECTODOS(strTitulo As String)

On Error GoTo err_TODOS

    Dim sValor As String
    
    lngCODOPERACAO = objRELCONTAREC.Gera_Codigo(Me.Name)
    
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    '' ----------------------------------------------------------------
    '' Pegando Titulos em aberto
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CONTASIARC.SGI_NUMDOC   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_DATAVENC " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_PARCELA  " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLDOC    " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODIGO   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_FILIAL   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_QTDPARC  " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODCLI   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODGRPRECEB " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CONTASHARC SGI_CONTASHARC" & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC SGI_CONTASIARC" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CONTASIARC.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CONTASIARC.SGI_DATAVENC >= '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And SGI_CONTASIARC.SGI_DATAVENC <= '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' " & vbCrLf
    sSql = sSql & "   And SGI_CONTASIARC.SGI_VLPAGO IS NULL " & vbCrLf
    If Len(Trim(txtCODCLI.Text)) > 0 Then
       sSql = sSql & "   And SGI_CONTASHARC.SGI_CODCLI   = " & txtCODCLI.Text & " " & vbCrLf
    End If
    If Len(Trim(txtCODGRPRECEB.Text)) > 0 Then
       sSql = sSql & "   And SGI_CONTASHARC.SGI_CODGRPRECEB  = " & txtCODGRPRECEB.Text & " " & vbCrLf
    End If
    sSql = sSql & "   And SGI_CONTASHARC.SGI_FILIAL   = SGI_CONTASIARC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CONTASHARC.SGI_CODIGO   = SGI_CONTASIARC.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       sSql = " Insert into SGI_TEMPCONTAPGREC( " & vbCrLf
       sSql = sSql & "                                SGI_FILIAL" & vbCrLf
       sSql = sSql & "                               ,SGI_OPERACAO" & vbCrLf
       sSql = sSql & "                               ,SGI_NUMDOC" & vbCrLf
       sSql = sSql & "                               ,SGI_DATA" & vbCrLf
       sSql = sSql & "                               ,SGI_DATAVENC" & vbCrLf
       sSql = sSql & "                               ,SGI_DATAPGTO" & vbCrLf
       sSql = sSql & "                               ,SGI_CODFORNEC" & vbCrLf
       sSql = sSql & "                               ,SGI_CODCLI" & vbCrLf
       sSql = sSql & "                               ,SGI_CODGRPDSP" & vbCrLf
       sSql = sSql & "                               ,SGI_PARCELA" & vbCrLf
       sSql = sSql & "                               ,SGI_TOTPARC" & vbCrLf
       sSql = sSql & "                               ,SGI_VLDOC" & vbCrLf
       sSql = sSql & "                               ,SGI_VLPAGO" & vbCrLf
       sSql = sSql & "                               ,SGI_VLDESC" & vbCrLf
       sSql = sSql & "                               ,SGI_VLACRESC" & vbCrLf
       sSql = sSql & "                               ,SGI_STATUS" & vbCrLf
       sSql = sSql & "                               ,SGI_TIPREL)" & vbCrLf
       sSql = sSql & "                        Values ( " & vbCrLf
       sSql = sSql & "                                 " & FILIAL & vbCrLf
       sSql = sSql & "                                ," & lngCODOPERACAO & vbCrLf
       sSql = sSql & "                                ,'" & BREC!SGI_NUMDOC & "'" & vbCrLf
       sSql = sSql & "                                ,'" & Format(BREC!SGI_DATAVENC, "MM/DD/YYYY") & "'" & vbCrLf
       sSql = sSql & "                                ,'" & Format(BREC!SGI_DATAVENC, "MM/DD/YYYY") & "'" & vbCrLf
       sSql = sSql & "                                ,Null" & vbCrLf
       sSql = sSql & "                                ,Null" & vbCrLf
       sSql = sSql & "                                ," & BREC!SGI_CODCLI & vbCrLf
       sSql = sSql & "                                ," & BREC!SGI_CODGRPRECEB & vbCrLf
       sSql = sSql & "                                ," & BREC!SGI_PARCELA & vbCrLf
       sSql = sSql & "                                ," & BREC!SGI_QTDPARC & vbCrLf
       
       sValor = Replace(BREC!SGI_VLDOC, ".", "")
       sValor = Replace(Trim(sValor), ",", ".")
       
       sSql = sSql & "                                ," & sValor & vbCrLf
       sSql = sSql & "                                ,Null" & vbCrLf
       sSql = sSql & "                                ,Null" & vbCrLf
       sSql = sSql & "                                ,Null" & vbCrLf
       sSql = sSql & "                                ,'A'" & vbCrLf
       sSql = sSql & "                                ,2)"
       
       BGRV.CommandText = sSql
       BGRV.Execute
        
       BREC.MoveNext
    Loop
    
    BREC.Close
    
    
    '' ----------------------------------------------------------------
    '' Pegando Titulos Baixados
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CONTASIARC.SGI_NUMDOC   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_DTPGTO   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_DATAVENC " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_PARCELA  " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLDOC    " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLPAGO   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLDESC   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLACRES  " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODIGO   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_FILIAL   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_QTDPARC  " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODCLI   " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODGRPRECEB " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CONTASIARC " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CONTASIARC.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CONTASIARC.SGI_DTPGTO >= '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And SGI_CONTASIARC.SGI_DTPGTO <= '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' " & vbCrLf
    sSql = sSql & "   And SGI_CONTASIARC.SGI_VLPAGO IS NOT NULL " & vbCrLf
    If Len(Trim(txtCODCLI.Text)) > 0 Then
       sSql = sSql & "   And SGI_CONTASHARC.SGI_CODCLI   = " & txtCODCLI.Text & " " & vbCrLf
    End If
    If Len(Trim(txtCODGRPRECEB.Text)) > 0 Then
       sSql = sSql & "   And SGI_CONTASHARC.SGI_CODGRPRECEB  = " & txtCODGRPRECEB.Text & " " & vbCrLf
    End If
    sSql = sSql & "   And SGI_CONTASHARC.SGI_FILIAL   = SGI_CONTASIARC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CONTASHARC.SGI_CODIGO   = SGI_CONTASIARC.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       sSql = " Insert into SGI_TEMPCONTAPGREC( " & vbCrLf
       sSql = sSql & "                                SGI_FILIAL" & vbCrLf
       sSql = sSql & "                               ,SGI_OPERACAO" & vbCrLf
       sSql = sSql & "                               ,SGI_NUMDOC" & vbCrLf
       sSql = sSql & "                               ,SGI_DATA" & vbCrLf
       sSql = sSql & "                               ,SGI_DATAVENC" & vbCrLf
       sSql = sSql & "                               ,SGI_DATAPGTO" & vbCrLf
       sSql = sSql & "                               ,SGI_CODFORNEC" & vbCrLf
       sSql = sSql & "                               ,SGI_CODCLI" & vbCrLf
       sSql = sSql & "                               ,SGI_CODGRPDSP" & vbCrLf
       sSql = sSql & "                               ,SGI_PARCELA" & vbCrLf
       sSql = sSql & "                               ,SGI_TOTPARC" & vbCrLf
       sSql = sSql & "                               ,SGI_VLDOC" & vbCrLf
       sSql = sSql & "                               ,SGI_VLPAGO" & vbCrLf
       sSql = sSql & "                               ,SGI_VLDESC" & vbCrLf
       sSql = sSql & "                               ,SGI_VLACRESC" & vbCrLf
       sSql = sSql & "                               ,SGI_STATUS" & vbCrLf
       sSql = sSql & "                               ,SGI_TIPREL)" & vbCrLf
       sSql = sSql & "                        Values ( " & vbCrLf
       sSql = sSql & "                                 " & FILIAL & vbCrLf
       sSql = sSql & "                                ," & lngCODOPERACAO & vbCrLf
       sSql = sSql & "                                ,'" & BREC!SGI_NUMDOC & "'" & vbCrLf
       sSql = sSql & "                                ,'" & Format(BREC!SGI_DTPGTO, "MM/DD/YYYY") & "'" & vbCrLf
       sSql = sSql & "                                ,'" & Format(BREC!SGI_DATAVENC, "MM/DD/YYYY") & "'" & vbCrLf
       sSql = sSql & "                                ,'" & Format(BREC!SGI_DTPGTO, "MM/DD/YYYY") & "'" & vbCrLf
       sSql = sSql & "                                ,Null" & vbCrLf
       sSql = sSql & "                                ," & BREC!SGI_CODCLI & vbCrLf
       sSql = sSql & "                                ," & BREC!SGI_CODGRPRECEB & vbCrLf
       sSql = sSql & "                                ," & BREC!SGI_PARCELA & vbCrLf
       sSql = sSql & "                                ," & BREC!SGI_QTDPARC & vbCrLf
       
       sValor = Replace(BREC!SGI_VLDOC, ".", "")
       sValor = Replace(Trim(sValor), ",", ".")
       sSql = sSql & "                                ," & sValor & vbCrLf
       
       sValor = Replace(BREC!SGI_VLPAGO, ".", "")
       sValor = Replace(Trim(sValor), ",", ".")
       sSql = sSql & "                                ," & sValor & vbCrLf
       
       If Not IsNull(BREC!SGI_VLDESC) Then
          sValor = Replace(BREC!SGI_VLDESC, ".", "")
          sValor = Replace(Trim(sValor), ",", ".")
          sSql = sSql & "                                ," & sValor & vbCrLf
       Else
          sSql = sSql & "                                ,Null" & vbCrLf
       End If
       If Not IsNull(BREC!SGI_VLACRES) Then
          sValor = Replace(BREC!SGI_VLACRES, ".", "")
          sValor = Replace(Trim(sValor), ",", ".")
          sSql = sSql & "                                ," & sValor & vbCrLf
       Else
          sSql = sSql & "                                ,Null" & vbCrLf
       End If
       sSql = sSql & "                                ,'B'" & vbCrLf
       sSql = sSql & "                                ,2)"
       
       BGRV.CommandText = sSql
       BGRV.Execute
        
       BREC.MoveNext
    Loop
    
    BREC.Close
    
    adoBanco_Dados.CommitTrans
    
    
    '' Chamando o Relatório
    sSql = "Select "
    sSql = sSql & "       * "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_TEMPCONTAPGREC "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL
    sSql = sSql & "   And SGI_OPERACAO = " & lngCODOPERACAO
    sSql = sSql & "Order by SGI_DATA "
    
    BREC.Open sSql, adoBanco_Dados, adOpenStatic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If Option1(0).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCANTDTV2.rpt", Linha, 1, strTitulo, "", True
    End If
    If Option1(1).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCANTFOR2.rpt", Linha, 1, strTitulo, "", True
    End If
    If Option1(0).Value = True And Option3(1).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCSITDTV2.rpt", Linha, 1, strTitulo, "", True
    End If
    
    If Option1(2).Value = True And Option3(0).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCGRPRECTOD.rpt", Linha, 1, strTitulo, "(Análitico - Todos)", True
    End If
    If Option1(2).Value = True And Option3(1).Value = True Then
       objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasARC & "RELARCGRPRECSINTOD.rpt", Linha, 1, strTitulo, "(Sintético - Todos)", True
    End If
    
    
    '' -------------------------------------------------------------------
    '' Apagando a base de dados
    sSql = "Delete From SGI_TEMPCONTAPGREC " & vbCrLf
    sSql = sSql & "      Where " & vbCrLf
    sSql = sSql & "            SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "        And SGI_OPERACAO = " & lngCODOPERACAO
    
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    BGRV.CommandText = sSql
    BGRV.Execute
    
    adoBanco_Dados.CommitTrans
    '' -------------------------------------------------------------------
    
    Exit Sub
    
err_TODOS:

    MsgBox "Erro Nº: " & Err.Number & " ]- Dewscrição : " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    adoBanco_Dados.RollbackTrans
    If BREC.State = 1 Then BREC.Close
    
End Sub
Private Sub txtCODGRPRECEB_GotFocus()
    objBLBFunc.SelecionaCampos txtCODGRPRECEB.Name, frmRELCONTAREC
End Sub

Private Sub txtCODGRPRECEB_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLI.Text
End Sub

Private Sub txtCODGRPRECEB_Validate(Cancel As Boolean)
    Dim I As Integer
    
    If Len(Trim(txtCODGRPRECEB.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODGRPRECEB.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRPRECEB.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    cboGRPRECEB.ListIndex = -1
    For I = 0 To (cboGRPRECEB.ListCount - 1)
        If cboGRPRECEB.ItemData(I) = Str(Val(txtCODGRPRECEB.Text)) Then cboGRPRECEB.ListIndex = I
    Next I
    
    If cboGRPRECEB.ListIndex = -1 Then
       MsgBox "EstE  Grupo de Recebimento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRPRECEB.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub
