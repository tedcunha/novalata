VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADTABPRECO 
   Caption         =   "Cadastro de Tabela de Preços"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   2655
      Left            =   0
      TabIndex        =   18
      Top             =   3120
      Width           =   7815
      Begin MSFlexGridLib.MSFlexGrid flxTABPRECO 
         Height          =   2415
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4260
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   7815
      Begin VB.TextBox txtVLVENDA 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   7
         Text            =   "txtVLVENDA"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtPORCACRE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Text            =   "txtPORCACRE"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtVALOR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "txtVALOR"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmbGravPagto 
         Height          =   315
         Left            =   7320
         Picture         =   "frmCADTABPRECO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCONDPGTO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "txtCONDPGTO"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2400
         Picture         =   "frmCADTABPRECO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   375
      End
      Begin VB.ComboBox cboCONDPGTO 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Text            =   "cboCONDPGTO"
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vl. Venda:"
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
         Index           =   5
         Left            =   5160
         TabIndex        =   23
         Top             =   510
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   490
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "% Acrecimo:"
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
         Index           =   3
         Left            =   2640
         TabIndex        =   21
         Top             =   510
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cond. Pgto:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   160
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   7815
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   975
         Width           =   1815
         Begin VB.OptionButton optVIGSIMNAO 
            Caption         =   "Nâo"
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
            Left            =   960
            TabIndex        =   30
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optVIGSIMNAO 
            Caption         =   "Sim"
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
            TabIndex        =   29
            Top             =   0
            Width           =   735
         End
      End
      Begin MSMask.MaskEdBox mskDATATAB 
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPRODUTO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtPRODUTO"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdFornec 
         Height          =   315
         Left            =   2400
         Picture         =   "frmCADTABPRECO.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox cboPRODUTO 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Text            =   "cboPRODUTO"
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtCODIGO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Text            =   "txtCODIGO"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vigente:"
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
         Index           =   7
         Left            =   2640
         TabIndex        =   27
         Top             =   975
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         TabIndex        =   15
         Top             =   600
         Width           =   735
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7815
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
         Left            =   960
         Picture         =   "frmCADTABPRECO.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   855
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
         Left            =   1800
         MaskColor       =   &H8000000F&
         Picture         =   "frmCADTABPRECO.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmCADTABPRECO.frx":093A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADTABPRECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public cTipOper     As String
Public iCodigo      As Integer
Public FILIAL       As Integer
Public strAcesso    As String
Public strMODPAI    As String
Public strUSUARIO   As String
Dim arrTABPRECOS    As Variant
Dim objBLBFunc      As Object
Dim objCADTABPRECO  As Object
Dim objPESQPADRAO   As Object
Private Sub cboCONDPGTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCONDPGTO, KeyAscii
End Sub

Private Sub cboCONDPGTO_Validate(Cancel As Boolean)
   If cboCONDPGTO.ListIndex > -1 Then txtCONDPGTO.Text = cboCONDPGTO.ItemData(cboCONDPGTO.ListIndex)
End Sub

Private Sub cboPRODUTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboPRODUTO, KeyAscii
End Sub

Private Sub cboPRODUTO_Validate(Cancel As Boolean)
    If cboPRODUTO.ListIndex > -1 Then
       txtPRODUTO.Text = Mid(cboPRODUTO.List(cboPRODUTO.ListIndex), 1, 10)
       If PegaValorProd(txtPRODUTO.Text) > 0 Then txtVALOR.Enabled = False
       If PegaValorProd(txtPRODUTO.Text) = 0 Then txtVALOR.Enabled = True
    End If
End Sub

Private Sub cmbGravPagto_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncTabPreco
End Sub

Private Sub cmdAltera_Click()
    
    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Tabela de Preços - [ ALTERACAO ]"
    cTipOper = "A"
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    txtCODIGO.Enabled = False
    txtPRODUTO.Enabled = False
    cmdFornec.Enabled = False
    cboPRODUTO.Enabled = False
    mskDATATAB.Enabled = False
    
    
End Sub

Private Sub cmdFornec_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "SELECT " & vbCrLf
    sSql = sSql & "       PRODUTO.SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,PRODUTO.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  FROM " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO     PRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRODUTO.SGI_FILIAL = " & FILIAL & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "PRODUTO.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "PRODUTO.SGI_DESCRICAO"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
        
    If Len(Trim(varRETORNO)) > 0 Then txtPRODUTO.Text = varRETORNO
    
    cboPRODUTO.ListIndex = -1
    txtPRODUTO.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADTABPRECO.CODIGO = objCADTABPRECO.Gera_Codigo(Me.Name)
    
    objCADTABPRECO.CODPROD = Trim(txtPRODUTO.Text)
    objCADTABPRECO.DATATAB = CDate(mskDATATAB.Text)
    If optVIGSIMNAO(0).Value = True Then objCADTABPRECO.VIGSIMNAO = "S"
    If optVIGSIMNAO(1).Value = True Then objCADTABPRECO.VIGSIMNAO = "N"
    
    If flxTABPRECO.Rows > 1 Then
       ReDim arrTABPRECOS(1 To (flxTABPRECO.Rows - 1), 1 To 4) As String
       For I = 1 To (flxTABPRECO.Rows - 1)
           arrTABPRECOS(I, 1) = flxTABPRECO.TextMatrix(I, 1)
           arrTABPRECOS(I, 2) = flxTABPRECO.TextMatrix(I, 3)
           arrTABPRECOS(I, 3) = flxTABPRECO.TextMatrix(I, 4)
           arrTABPRECOS(I, 4) = flxTABPRECO.TextMatrix(I, 5)
       Next I
       objCADTABPRECO.TABPRECO = arrTABPRECOS
    End If
    
    If objCADTABPRECO.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A tabela de preços foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADTABPRECO = Nothing
       Set objPESQPADRAO = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADTABPRECO = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADCONDPGTO " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Condição de Pagamento")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCONDPGTO.Text = varRETORNO
        
    cboCONDPGTO.ListIndex = -1
    txtCONDPGTO.SetFocus

End Sub

Private Sub flxTABPRECO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxTABPRECO.Rows = 1 Then Exit Sub
       If flxTABPRECO.Rows = 2 Then flxTABPRECO.Rows = 1
       If flxTABPRECO.Rows > 2 Then flxTABPRECO.RemoveItem flxTABPRECO.Row
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADTABPRECO = CreateObject("CADTABPRECO.clsCADTABPRECO")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
      
   objCADTABPRECO.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Tabela de Preços - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTABPRECO
    
    ConfGridPrecos
    
    objCADTABPRECO.PreencheComboProd cboPRODUTO
    objCADTABPRECO.PreencheComboCondPgto cboCONDPGTO
    
    mskDATATAB.Text = Format(Now, "DD/MM/YYYY")
    optVIGSIMNAO(0).Value = True
   
End Sub


Private Sub ConfGridPrecos()

    flxTABPRECO.Rows = 1
    flxTABPRECO.Cols = 6
    
    flxTABPRECO.TextMatrix(0, 0) = ""
    flxTABPRECO.TextMatrix(0, 1) = "Código"
    flxTABPRECO.TextMatrix(0, 2) = "Descrição"
    flxTABPRECO.TextMatrix(0, 3) = "Valor"
    flxTABPRECO.TextMatrix(0, 4) = "% Acr."
    flxTABPRECO.TextMatrix(0, 5) = "Vl. Venda"
    
    flxTABPRECO.ColWidth(0) = 0
    flxTABPRECO.ColWidth(1) = 1000
    flxTABPRECO.ColWidth(2) = 3000
    flxTABPRECO.ColWidth(3) = 1000
    flxTABPRECO.ColWidth(4) = 1000
    flxTABPRECO.ColWidth(5) = 1000
    
    flxTABPRECO.ColAlignment(2) = vbLeftJustify
    
End Sub

Private Sub mskDATATAB_GotFocus()
    objBLBFunc.SelecionaCampos mskDATATAB.Name, frmCADTABPRECO
End Sub

Private Sub txtCONDPGTO_GotFocus()
    objBLBFunc.SelecionaCampos txtCONDPGTO.Name, frmCADTABPRECO
End Sub

Private Sub txtCONDPGTO_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCONDPGTO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCONDPGTO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCONDPGTO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCONDPGTO.ListIndex = -1
    For I = 0 To (cboCONDPGTO.ListCount - 1)
        If cboCONDPGTO.ItemData(I) = CInt(txtCONDPGTO.Text) Then cboCONDPGTO.ListIndex = I
    Next I
    
    If cboCONDPGTO.ListIndex = -1 Then
       MsgBox "Esta condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCONDPGTO.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtPORCACRE_GotFocus()
    objBLBFunc.SelecionaCampos txtPORCACRE.Name, frmCADTABPRECO
End Sub

Private Sub txtPORCACRE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPORCACRE.Text
End Sub

Private Sub txtPORCACRE_Validate(Cancel As Boolean)

    If Len(Trim(txtPORCACRE.Text)) = 0 Then
       CalcValorVenda
       Exit Sub
    End If

    If Not IsNumeric(txtPORCACRE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If CCur(txtPORCACRE.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtPORCACRE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPORCACRE.Text = Format(txtPORCACRE.Text, "#,##0.00")
    CalcValorVenda

End Sub

Private Sub txtPRODUTO_GotFocus()
    objBLBFunc.SelecionaCampos txtPRODUTO.Name, frmCADTABPRECO
End Sub

Private Sub txtPRODUTO_Validate(Cancel As Boolean)

   Dim I As Integer

   txtVALOR.Text = ""
   If Len(Trim(txtPRODUTO.Text)) = 0 Then Exit Sub
   
   cboPRODUTO.ListIndex = -1
   For I = 0 To (cboPRODUTO.ListCount - 1)
       If Trim(Mid(cboPRODUTO.List(I), 1, 10)) = Trim(txtPRODUTO.Text) Then cboPRODUTO.ListIndex = I
   Next I
    
   If cboPRODUTO.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtPRODUTO.Text = ""
      Cancel = True
      Exit Sub
   End If
   
   If PegaValorProd(txtPRODUTO.Text) > 0 Then txtVALOR.Enabled = False
   If PegaValorProd(txtPRODUTO.Text) = 0 Then txtVALOR.Enabled = True
   
End Sub

Private Function PegaValorProd(strPRODUTO As String) As Currency

    PegaValorProd = 0
    
    txtVALOR.Text = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & txtPRODUTO.Text & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaValorProd = BREC!SGI_PRECOPROD
    BREC.Close
    
    If PegaValorProd > 0 Then txtVALOR.Text = Format(PegaValorProd, "#,##0.00")

End Function

Private Sub txtVALOR_GotFocus()
    objBLBFunc.SelecionaCampos txtVALOR.Name, frmCADTABPRECO
End Sub

Private Sub txtVALOR_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtVALOR.Text
End Sub

Private Sub txtVALOR_Validate(Cancel As Boolean)

    If Len(Trim(txtVALOR.Text)) = 0 Then
       CalcValorVenda
       Exit Sub
    End If

    If Not IsNumeric(txtVALOR.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If CCur(txtVALOR.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtVALOR.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtVALOR.Text = Format(txtVALOR.Text, "#,##0.00")
    CalcValorVenda

End Sub

Private Sub CalcValorVenda()

    Dim curVALOR   As Currency
    Dim curPORCAC  As Currency
    Dim curVLACRES As Currency
    Dim curVLVENDA As Currency
    
    If Len(Trim(txtVALOR.Text)) > 0 Then curVALOR = CCur(txtVALOR.Text)
    If Len(Trim(txtPORCACRE.Text)) > 0 Then curPORCAC = CCur(txtPORCACRE.Text)
    
    curVLACRES = ((curPORCAC * curVALOR) / 100)
    curVLVENDA = curVALOR + curVLACRES
    
    txtVLVENDA.Text = Format(curVLVENDA, "#,##0.00")

End Sub

Private Sub IncTabPreco()

    Dim I As Integer
    
    If Len(Trim(txtCONDPGTO.Text)) = 0 Then
       MsgBox "Informe a condição de pagamento !!!", vbOKOnly + vbCritical, "Aviso"
       txtCONDPGTO.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtVALOR.Text)) = 0 Then
       MsgBox "Informe o campo valor !!!", vbOKOnly + vbExclamation, "Aviso"
       txtVALOR.SetFocus
       Exit Sub
    End If
    If CCur(txtVALOR.Text) = 0 Then
       MsgBox "Informe o campo valor !!!", vbOKOnly + vbExclamation, "Aviso"
       txtVALOR.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtPORCACRE.Text)) = 0 Then
       MsgBox "Informe a porcentagem de acrescimo !!!", vbOKOnly + vbExclamation, "Aviso"
       txtPORCACRE.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxTABPRECO.Rows - 1)
        If CInt(txtCONDPGTO.Text) = CInt(flxTABPRECO.TextMatrix(I, 1)) Then
           MsgBox "Está condição de Pgto já foi relacionada !!!", vbOKOnly + vbExclamation, "Aviso"
           txtCONDPGTO.SetFocus
           Exit Sub
        End If
    Next I
    
    flxTABPRECO.AddItem "" & vbTab & _
                        txtCONDPGTO.Text & vbTab & _
                        cboCONDPGTO.Text & vbTab & _
                        txtVALOR.Text & vbTab & _
                        txtPORCACRE.Text & vbTab & _
                        txtVLVENDA.Text
                        
    txtCONDPGTO.Text = ""
    cboCONDPGTO.ListIndex = -1
    txtPORCACRE.Text = ""
    txtVLVENDA.Text = ""
    
    txtCONDPGTO.SetFocus
    
End Sub


Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     Dim strVIGSIMNAO As String
     
     If optVIGSIMNAO(0).Value = True Then strVIGSIMNAO = "S"
     If optVIGSIMNAO(1).Value = True Then strVIGSIMNAO = "N"
     
     If Trim(Len(txtPRODUTO.Text)) = 0 Then
        MsgBox "Código do Produto inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtPRODUTO.SetFocus
        Exit Function
     End If
     If flxTABPRECO.Rows = 1 Then
        MsgBox "Informe a tabela de preços !!!", vbOKOnly + vbCritical, "aviso"
        txtCONDPGTO.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
     
        If optVIGSIMNAO(0).Value = True Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_TABPRECO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODPROD = '" & Trim(txtPRODUTO.Text) & "'" & vbCrLf
           sSql = sSql & "   And SGI_VIGENTE = 'S'"
        
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
           If Not BREC.EOF Then
              MsgBox "Já existe tabela vigente favor desmarcar a tabela vigente atual !!!", vbOKOnly + vbExclamation, "Aviso"
              BREC.Close
              Exit Function
           End If
           BREC.Close
           
        End If
     
     End If
     
     If cTipOper = "A" Then
     
        If objCADTABPRECO.VIGSIMNAO <> strVIGSIMNAO Then
        
           If strVIGSIMNAO = "S" Then
           
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_TABPRECO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODPROD = '" & Trim(txtPRODUTO.Text) & "'" & vbCrLf
              sSql = sSql & "   And SGI_VIGENTE = 'S'"
        
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
              If Not BREC.EOF Then
                 MsgBox "Já existe tabela vigente favor desmarcar a tabela vigente atual !!!", vbOKOnly + vbExclamation, "Aviso"
              optVIGSIMNAO(1).Value = True
              BREC.Close
              Exit Function
           End If
           BREC.Close
           
           End If
           
        End If
        
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    Frame2.Enabled = False
    Frame3.Enabled = False
   
    Me.Caption = "Cadastro de Tabela de Preços - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADTABPRECO
    
    ConfGridPrecos
    
    objCADTABPRECO.PreencheComboProd cboPRODUTO
    objCADTABPRECO.PreencheComboCondPgto cboCONDPGTO
    
    objCADTABPRECO.CODIGO = iCodigo
    
    If objCADTABPRECO.Carrega_campos = True Then
    
       txtCODIGO.Text = objCADTABPRECO.CODIGO
       txtPRODUTO.Text = objCADTABPRECO.CODPROD
       For I = 0 To cboPRODUTO.ListCount
           If Trim(Mid(cboPRODUTO.List(I), 1, 10)) = Trim(txtPRODUTO.Text) Then cboPRODUTO.ListIndex = I
       Next I
       
       mskDATATAB.Text = Format(objCADTABPRECO.DATATAB, "DD/MM/YYYY")
       If objCADTABPRECO.VIGSIMNAO = "S" Then optVIGSIMNAO(0).Value = True
       If objCADTABPRECO.VIGSIMNAO = "N" Then optVIGSIMNAO(1).Value = True
       
       arrTABPRECOS = objCADTABPRECO.TABPRECO
       
       If IsArray(arrTABPRECOS) = True Then
          For I = 1 To UBound(arrTABPRECOS)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADCONDPGTO " & vbCrLf
              sSql = sSql & "  Where " & vbCrLf
              sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "    And SGI_CODIGO = " & arrTABPRECOS(I, 1)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          
              If Not BREC.EOF Then
                 flxTABPRECO.AddItem "" & vbTab & _
                                     arrTABPRECOS(I, 1) & vbTab & _
                                     BREC!SGI_DESCRICAO & vbTab & _
                                     Format(CCur(arrTABPRECOS(I, 2)), "#,##0.00") & vbTab & _
                                     Format(CCur(arrTABPRECOS(I, 3)), "#,##0.00") & vbTab & _
                                     Format(CCur(arrTABPRECOS(I, 4)), "#,##0.00")
              End If
              BREC.Close
              
          Next I
       End If
       
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de Tabela de Preços - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTABPRECO
    
    ConfGridPrecos
    
    objCADTABPRECO.PreencheComboProd cboPRODUTO
    objCADTABPRECO.PreencheComboCondPgto cboCONDPGTO
    
    objCADTABPRECO.CODIGO = iCodigo
    
    txtPRODUTO.Enabled = False
    cmdFornec.Enabled = False
    cboPRODUTO.Enabled = False
    mskDATATAB.Enabled = False
    
    
    If objCADTABPRECO.Carrega_campos = True Then
    
       txtCODIGO.Text = objCADTABPRECO.CODIGO
       txtPRODUTO.Text = objCADTABPRECO.CODPROD
       For I = 0 To cboPRODUTO.ListCount
           If Trim(Mid(cboPRODUTO.List(I), 1, 10)) = Trim(txtPRODUTO.Text) Then cboPRODUTO.ListIndex = I
       Next I
       
       mskDATATAB.Text = Format(objCADTABPRECO.DATATAB, "DD/MM/YYYY")
       If objCADTABPRECO.VIGSIMNAO = "S" Then optVIGSIMNAO(0).Value = True
       If objCADTABPRECO.VIGSIMNAO = "N" Then optVIGSIMNAO(1).Value = True
       
       arrTABPRECOS = objCADTABPRECO.TABPRECO
       
       If IsArray(arrTABPRECOS) = True Then
          For I = 1 To UBound(arrTABPRECOS)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADCONDPGTO " & vbCrLf
              sSql = sSql & "  Where " & vbCrLf
              sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "    And SGI_CODIGO = " & arrTABPRECOS(I, 1)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          
              If Not BREC.EOF Then
                 flxTABPRECO.AddItem "" & vbTab & _
                                     arrTABPRECOS(I, 1) & vbTab & _
                                     BREC!SGI_DESCRICAO & vbTab & _
                                     Format(CCur(arrTABPRECOS(I, 2)), "#,##0.00") & vbTab & _
                                     Format(CCur(arrTABPRECOS(I, 3)), "#,##0.00") & vbTab & _
                                     Format(CCur(arrTABPRECOS(I, 4)), "#,##0.00")
              End If
              BREC.Close
              
          Next I
       End If
       
    End If

End Sub


