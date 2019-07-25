VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADLOTE 
   Caption         =   "Geração de Lotes"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11190
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   0
      TabIndex        =   12
      Top             =   4800
      Width           =   11055
      Begin VB.TextBox txtBanco 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtBanco"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   2760
         TabIndex        =   6
         Text            =   "cboBanco"
         Top             =   840
         Width           =   4815
      End
      Begin VB.CommandButton cmdPesq 
         Height          =   315
         Left            =   2400
         Picture         =   "frmCADLOTE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdTipoPgto 
         Height          =   315
         Left            =   2415
         Picture         =   "frmCADLOTE.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cboTIPOPGTO 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Text            =   "cboTIPOPGTO"
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txtCODTIPPGTO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "txtCODFORNEC"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtNDOC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtNDOC"
         Top             =   120
         Width           =   1815
      End
      Begin MSMask.MaskEdBox mskDTLOTE 
         Height          =   285
         Left            =   4800
         TabIndex        =   2
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   7
         Left            =   10320
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.Titulos.:"
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
         Index           =   6
         Left            =   9120
         TabIndex        =   21
         Top             =   165
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
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
         Index           =   5
         Left            =   960
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pagamento:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data.:"
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
         Index           =   0
         Left            =   4080
         TabIndex        =   16
         Top             =   165
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor.:"
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
         Index           =   1
         Left            =   6360
         TabIndex        =   15
         Top             =   165
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   7080
         TabIndex        =   14
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   165
         Width           =   1305
      End
   End
   Begin VB.Frame frmLote 
      Height          =   3855
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   11055
      Begin MSFlexGridLib.MSFlexGrid flxTitulos 
         Height          =   1695
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   2990
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxLote 
         Height          =   1935
         Left            =   120
         TabIndex        =   0
         Top             =   1800
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11055
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
         Picture         =   "frmCADLOTE.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmCADLOTE.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADLOTE.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADLOTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public cTipOper      As String
Public iCodigo       As Integer
Public iParcela      As Integer
Public FILIAL        As Integer
Public strAcesso     As String
Public strMODPAI     As String
Public strUSUARIO    As String
Dim objBLBFunc       As Object
Dim objCADCONTASAPG  As Object
Dim objPESQPADRAO    As Object
Dim arrGRIDPGTOS     As Variant

Private Sub cboBanco_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboBanco, KeyAscii
End Sub

Private Sub cboBanco_Validate(Cancel As Boolean)
    If cboBanco.ListIndex > -1 Then txtBanco.Text = cboBanco.ItemData(cboBanco.ListIndex)
End Sub

Private Sub cboTIPOPGTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTIPOPGTO, KeyAscii
End Sub

Private Sub cboTIPOPGTO_Validate(Cancel As Boolean)
    If cboTIPOPGTO.ListIndex > -1 Then txtCODTIPPGTO.Text = cboTIPOPGTO.ItemData(cboTIPOPGTO.ListIndex)
End Sub

Private Sub cmdAltera_Click()
    
    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    If objCADCONTASAPG.STATUS = "L" Then
       MsgBox "Lote já liberado não pode ser alterado deslibere !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    If objCADCONTASAPG.STATUS = "B" Then
       MsgBox "Lote já baixa não pode ser alterado extorne !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    cmdVoltar.Enabled = True
    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    Frame3.Enabled = True
    
    Me.Caption = "Geração de Lotes - [ ALTERA ]"
    
    PreencheGrid
    
    cTipOper = "A"

End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 4, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    
    arrCAMPOS(2, 1) = "SGI_AGENCIA"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Agência"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_CC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "C/C"
    arrCAMPOS(3, 4) = "1500"
    
    arrCAMPOS(4, 1) = "SGI_DESCRICAO"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Banco"
    arrCAMPOS(4, 4) = "3000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Bancos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtBanco.Text = varRETORNO
    
    cboBanco.ListIndex = -1
    txtBanco.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I        As Integer
    Dim intResp  As Integer
    Dim intLinha As Integer
    
    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADCONTASAPG.CODLOTE = objCADCONTASAPG.Gera_Codigo("LOTE")
    
    objCADCONTASAPG.NUMDOC = txtNDOC.Text
    objCADCONTASAPG.DATALOTE = CDate(mskDTLOTE.Text)
    objCADCONTASAPG.TIPOPGTO = txtCODTIPPGTO.Text
    If Len(Trim(txtBanco.Text)) > 0 Then objCADCONTASAPG.CODBANCO = txtBanco.Text
    objCADCONTASAPG.VLTOTLCTO = CCur(Label1(2).Caption)
    
    If (flxLote.Rows - 1) > 0 Then
       ReDim arrGRIDPGTOS(1 To CInt(Label1(7).Caption), 1 To 5) As String
       intLinha = 1
       For I = 1 To (flxLote.Rows - 1)
           If flxLote.TextMatrix(I, 9) = "S" Then
              arrGRIDPGTOS(intLinha, 1) = flxLote.TextMatrix(I, 1)
              arrGRIDPGTOS(intLinha, 2) = flxLote.TextMatrix(I, 2)
              arrGRIDPGTOS(intLinha, 3) = flxLote.TextMatrix(I, 3)
              arrGRIDPGTOS(intLinha, 4) = flxLote.TextMatrix(I, 8)
              If intLinha = CInt(Label1(7).Caption) Then Exit For
              intLinha = intLinha + 1
           End If
       Next I
    End If
    
    objCADCONTASAPG.DOCPGTO = arrGRIDPGTOS
    
    If objCADCONTASAPG.GRAVALOTE(cTipOper) = False Then Exit Sub
          
    MsgBox "O lote foi " & IIf(cTipOper = "I", "gerado", IIf(cTipOper = "A", "alterado", IIf(cTipOper = "B", "baixado", IIf(cTipOper = "AB", "alterado", "")))) & " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    If objCADCONTASAPG.Atualiza(cTipOper, Str(objCADCONTASAPG.CODLOTE), FILIAL, "LOTE") = False Then Exit Sub
          
    If cTipOper = "I" Then
       intResp = MsgBox("Deseja incluir novo lote ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
       If intResp = 7 Then
          Set objBLBFunc = Nothing
          Set objCADCONTASAPG = Nothing
          Set objPESQPADRAO = Nothing
          Unload Me
       Else
          Inclui
          txtNDOC.SetFocus
       End If
    ElseIf cTipOper = "B" Or cTipOper = "AB" Then
       Set objBLBFunc = Nothing
       Set objCADCONTASAPG = Nothing
       Set objPESQPADRAO = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdTipoPgto_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPOPGTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_OPERACAO = 1"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    
    arrCAMPOS(3, 1) = "SGI_SINAL"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Sinal"
    arrCAMPOS(3, 4) = "500"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Tipo de Pagamento")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODTIPPGTO.Text = varRETORNO
    
    cboTIPOPGTO.ListIndex = -1
    txtCODTIPPGTO.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADCONTASAPG = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub flxLote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And (cTipOper = "I" Or cTipOper = "A") Then MudaTitulo
End Sub

Private Sub flxTitulos_DblClick()
    If cTipOper = "I" Or cTipOper = "A" Then AddTitulo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADCONTASAPG = CreateObject("CADCONTASAPG.clsCADCONTASAPG")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADCONTASAPG.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    cmdVoltar.Enabled = True
    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    Frame3.Enabled = True
    
    Me.Caption = "Geração de Lotes - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADLOTE
    
    Label1(2).Caption = ""
    Label1(7).Caption = "000"
    mskDTLOTE.Text = Format(Now, "DD/MM/YYYY")
    
    objCADCONTASAPG.PreenchComboBancos cboBanco
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGrid
    PreencheGrid

End Sub

Private Sub ConfGrid()
    
    flxTitulos.Rows = 1
    flxTitulos.Cols = 11
    
    flxTitulos.TextMatrix(0, 0) = ""
    flxTitulos.TextMatrix(0, 1) = "Código"
    flxTitulos.TextMatrix(0, 2) = "Nº Doc."
    flxTitulos.TextMatrix(0, 3) = "Vencto"
    flxTitulos.TextMatrix(0, 4) = "Parcela"
    flxTitulos.TextMatrix(0, 5) = "Valor"
    flxTitulos.TextMatrix(0, 6) = "Cod. Forn"
    flxTitulos.TextMatrix(0, 7) = "Razão Social"
    flxTitulos.TextMatrix(0, 8) = "Parcela"
    flxTitulos.TextMatrix(0, 9) = "Paga"
    flxTitulos.TextMatrix(0, 10) = "Pago"
    
    flxTitulos.ColWidth(0) = 0
    flxTitulos.ColWidth(1) = 1000
    flxTitulos.ColWidth(2) = 1000
    flxTitulos.ColWidth(3) = 1000
    flxTitulos.ColWidth(4) = 1000
    flxTitulos.ColWidth(5) = 1000
    flxTitulos.ColWidth(6) = 1000
    flxTitulos.ColWidth(7) = 3000
    flxTitulos.ColWidth(8) = 0
    flxTitulos.ColWidth(9) = 0
    flxTitulos.ColWidth(10) = 500
    
    '' ---------------------------------------
    
    flxLote.Rows = 1
    flxLote.Cols = 11
    
    flxLote.TextMatrix(0, 0) = ""
    flxLote.TextMatrix(0, 1) = "Código"
    flxLote.TextMatrix(0, 2) = "Nº Doc."
    flxLote.TextMatrix(0, 3) = "Vencto"
    flxLote.TextMatrix(0, 4) = "Parcela"
    flxLote.TextMatrix(0, 5) = "Valor"
    flxLote.TextMatrix(0, 6) = "Cod. Forn"
    flxLote.TextMatrix(0, 7) = "Razão Social"
    flxLote.TextMatrix(0, 8) = "Parcela"
    flxLote.TextMatrix(0, 9) = "Paga"
    flxLote.TextMatrix(0, 10) = "Pago"
    
    flxLote.ColWidth(0) = 0
    flxLote.ColWidth(1) = 1000
    flxLote.ColWidth(2) = 1000
    flxLote.ColWidth(3) = 1000
    flxLote.ColWidth(4) = 1000
    flxLote.ColWidth(5) = 1000
    flxLote.ColWidth(6) = 1000
    flxLote.ColWidth(7) = 3000
    flxLote.ColWidth(8) = 0
    flxLote.ColWidth(9) = 0
    flxLote.ColWidth(10) = 500
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
    sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
    sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
    sSql = sSql & "  And ITENS.SGI_STATUS = 'A'" & vbCrLf
    sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "       ,CABEC.SGI_CODFOR" & vbCrLf
    sSql = sSql & "       ,ITENS.SGI_PARCELA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       
       flxTitulos.AddItem "" & vbTab & _
                          BREC!SGI_CODIGO & vbTab & _
                          BREC!SGI_NUMDOC & vbTab & _
                          Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                          Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                          Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                          BREC!SGI_CODFOR & vbTab & _
                          BREC!SGI_RAZAOSOC & vbTab & _
                          BREC!SGI_PARCELA & vbTab & _
                          "N" & vbTab & _
                          "Não"
                               
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Private Sub MudaTitulo()
    If flxLote.Rows = 1 Then Exit Sub
       
    flxTitulos.AddItem flxLote.TextMatrix(flxLote.RowSel, 0) & vbTab & _
                       flxLote.TextMatrix(flxLote.RowSel, 1) & vbTab & _
                       flxLote.TextMatrix(flxLote.RowSel, 2) & vbTab & _
                       flxLote.TextMatrix(flxLote.RowSel, 3) & vbTab & _
                       flxLote.TextMatrix(flxLote.RowSel, 4) & vbTab & _
                       flxLote.TextMatrix(flxLote.RowSel, 5) & vbTab & _
                       flxLote.TextMatrix(flxLote.RowSel, 6) & vbTab & _
                       flxLote.TextMatrix(flxLote.RowSel, 7) & vbTab & _
                       flxLote.TextMatrix(flxLote.RowSel, 8) & vbTab & _
                       "N" & vbTab & _
                       "Não"
                    
    If flxLote.Rows = 2 Then flxLote.Rows = 1
    If flxLote.Rows > 2 Then flxLote.RemoveItem flxLote.RowSel

    Label1(7).Caption = Format(CLng(Label1(7).Caption) - 1, "###000")
    SomaLote
       
End Sub

Private Sub SomaLote()
    
    Dim I        As Integer
    Dim curTOTAL As Currency

    Label1(2).Caption = ""
    curTOTAL = 0
    For I = 1 To (flxLote.Rows - 1)
        If flxLote.TextMatrix(I, 9) = "S" Then curTOTAL = curTOTAL + CCur(flxLote.TextMatrix(I, 5))
    Next I
    
    If curTOTAL > 0 Then Label1(2).Caption = Format(curTOTAL, "#,##0.00")

End Sub


Private Sub mskDTLOTE_GotFocus()
    objBLBFunc.SelecionaCampos mskDTLOTE.Name, frmCADCONTASAPG
End Sub

Private Sub mskDTLOTE_Validate(Cancel As Boolean)

    If Not IsDate(mskDTLOTE.Text) Then
       MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If
    
    If VerifCaixa(CDate(mskDTLOTE.Text)) = True Then Cancel = True

End Sub

Private Sub txtBanco_GotFocus()
    objBLBFunc.SelecionaCampos txtBanco.Name, frmCADLOTE
End Sub

Private Sub txtBanco_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blACHOU As Boolean
    
    If Len(Trim(txtBanco.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtBanco.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtBanco.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    blACHOU = False
    For I = 0 To (cboBanco.ListCount - 1)
        If CInt(txtBanco.Text) = cboBanco.ItemData(I) Then
           blACHOU = True
           cboBanco.ListIndex = I
        End If
    Next I
    
    If blACHOU = False Then
       MsgBox "Este banco não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtBanco.Text = ""
       cboBanco.ListIndex = -1
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODTIPPGTO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPPGTO.Name, frmCADCONTASAPG
End Sub

Private Sub txtCODTIPPGTO_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODTIPPGTO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTIPPGTO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPPGTO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTIPOPGTO.ListIndex = -1
    For I = 0 To (cboTIPOPGTO.ListCount - 1)
        If cboTIPOPGTO.ItemData(I) = Str(Val(txtCODTIPPGTO.Text)) Then cboTIPOPGTO.ListIndex = I
    Next I
    
    If cboTIPOPGTO.ListIndex = -1 Then
       MsgBox "Esta Condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPPGTO.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtNDOC_GotFocus()
    objBLBFunc.SelecionaCampos txtNDOC.Name, frmCADCONTASAPG
End Sub

Private Function Valida_Campos() As Boolean

    Valida_Campos = False
    
    If Len(Trim(txtNDOC.Text)) = 0 Then
       MsgBox "Nº do documento não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
       txtNDOC.SetFocus
       Exit Function
    End If
    If IsDate(mskDTLOTE.Text) = False Then
       MsgBox "Data do lote inválido !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTLOTE.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODTIPPGTO.Text)) = 0 Then
       MsgBox "Tipo de pagamento não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCODTIPPGTO.SetFocus
       Exit Function
    End If
    
    Valida_Campos = True

End Function

Private Sub Consulta()

    Dim I As Integer
    
    cmdVoltar.Enabled = True
    cmdAltera.Enabled = True
    CmdSalva.Enabled = False
    
    Frame3.Enabled = False
    
    Me.Caption = "Geração de Lotes - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADLOTE
    
    Label1(2).Caption = ""
    Label1(7).Caption = "000"
    
    objCADCONTASAPG.PreenchComboBancos cboBanco
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGrid
    
    objCADCONTASAPG.CODLOTE = iCodigo
    
    If objCADCONTASAPG.Carrega_CamposLote = True Then
       
       txtNDOC.Text = objCADCONTASAPG.NUMDOC
       mskDTLOTE.Text = Format(objCADCONTASAPG.DATALOTE, "DD/MM/YYYY")
       Label1(2).Caption = Format(objCADCONTASAPG.VLTOTLCTO, "#,##0.00")
       txtCODTIPPGTO.Text = objCADCONTASAPG.TIPOPGTO
       txtBanco.Text = objCADCONTASAPG.CODBANCO
       
       arrGRIDPGTOS = objCADCONTASAPG.DOCPGTO
       
       '' Tipo pagamento
       For I = 0 To (cboTIPOPGTO.ListCount - 1)
           If objCADCONTASAPG.TIPOPGTO = cboTIPOPGTO.ItemData(I) Then cboTIPOPGTO.ListIndex = I
       Next I
       
       '' Tipo Bancos
       For I = 0 To (cboBanco.ListCount - 1)
           If objCADCONTASAPG.CODBANCO = cboBanco.ItemData(I) Then cboBanco.ListIndex = I
       Next I
       
       If IsArray(arrGRIDPGTOS) Then
       
          Label1(7).Caption = Format(UBound(arrGRIDPGTOS), "###000")
          
          For I = 1 To UBound(arrGRIDPGTOS)
          
              sSql = "Select" & vbCrLf
              sSql = sSql & "      ITENS.* " & vbCrLf
              sSql = sSql & "     ,HEADE.* " & vbCrLf
              sSql = sSql & "     ,FORNE.SGI_RAZAOSOC " & vbCrLf
              sSql = sSql & "  From" & vbCrLf
              sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
              sSql = sSql & "     ,SGI_CONTASHAPG HEADE" & vbCrLf
              sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
              sSql = sSql & " Where" & vbCrLf
              sSql = sSql & "      ITENS.SGI_FILIAL  = " & FILIAL & vbCrLf
              sSql = sSql & "  And ITENS.SGI_CODIGO  = " & arrGRIDPGTOS(I, 1) & vbCrLf
              sSql = sSql & "  And ITENS.SGI_PARCELA = " & arrGRIDPGTOS(I, 2) & vbCrLf
              sSql = sSql & "  And HEADE.SGI_FILIAL  = ITENS.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And HEADE.SGI_CODIGO  = ITENS.SGI_CODIGO " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_FILIAL  = HEADE.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_CODIGO  = HEADE.SGI_CODFOR "
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              
              If Not BREC.EOF Then
                 flxLote.AddItem "" & vbTab & _
                                 BREC!SGI_CODIGO & vbTab & _
                                 BREC!SGI_NUMDOC & vbTab & _
                                 Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                 Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                 Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                 BREC!SGI_CODFOR & vbTab & _
                                 BREC!SGI_RAZAOSOC & vbTab & _
                                 BREC!SGI_PARCELA & vbTab & _
                                 "S" & vbTab & _
                                 "Sim"
                                 
              End If
              
              BREC.Close
              
          Next I
       End If
       
    End If
    

End Sub

Private Sub AddTitulo()
    
    If flxTitulos.Rows = 1 Then Exit Sub
    
    flxLote.AddItem flxTitulos.TextMatrix(flxTitulos.RowSel, 0) & vbTab & _
                    flxTitulos.TextMatrix(flxTitulos.RowSel, 1) & vbTab & _
                    flxTitulos.TextMatrix(flxTitulos.RowSel, 2) & vbTab & _
                    flxTitulos.TextMatrix(flxTitulos.RowSel, 3) & vbTab & _
                    flxTitulos.TextMatrix(flxTitulos.RowSel, 4) & vbTab & _
                    flxTitulos.TextMatrix(flxTitulos.RowSel, 5) & vbTab & _
                    flxTitulos.TextMatrix(flxTitulos.RowSel, 6) & vbTab & _
                    flxTitulos.TextMatrix(flxTitulos.RowSel, 7) & vbTab & _
                    flxTitulos.TextMatrix(flxTitulos.RowSel, 8) & vbTab & _
                    "S" & vbTab & _
                    "Sim"
                    
    If flxTitulos.Rows = 2 Then flxTitulos.Rows = 1
    If flxTitulos.Rows > 2 Then flxTitulos.RemoveItem flxTitulos.RowSel
    
    Label1(7).Caption = Format(CLng(Label1(7).Caption) + 1, "###000")
    SomaLote
    
    
End Sub

Private Sub Altera()

    Dim I As Integer
    
    cmdVoltar.Enabled = True
    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    Frame3.Enabled = True
    
    Me.Caption = "Geração de Lotes - [ ALTERA ]"
    
    objBLBFunc.LimpaCampos frmCADLOTE
    
    Label1(2).Caption = ""
    Label1(7).Caption = "000"
    
    objCADCONTASAPG.PreenchComboBancos cboBanco
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGrid
    PreencheGrid
    
    objCADCONTASAPG.CODLOTE = iCodigo
    
    If objCADCONTASAPG.Carrega_CamposLote = True Then
       
       txtNDOC.Text = objCADCONTASAPG.NUMDOC
       mskDTLOTE.Text = Format(objCADCONTASAPG.DATALOTE, "DD/MM/YYYY")
       Label1(2).Caption = Format(objCADCONTASAPG.VLTOTLCTO, "#,##0.00")
       txtCODTIPPGTO.Text = objCADCONTASAPG.TIPOPGTO
       txtBanco.Text = objCADCONTASAPG.CODBANCO
       
       arrGRIDPGTOS = objCADCONTASAPG.DOCPGTO
       
       '' Tipo pagamento
       For I = 0 To (cboTIPOPGTO.ListCount - 1)
           If objCADCONTASAPG.TIPOPGTO = cboTIPOPGTO.ItemData(I) Then cboTIPOPGTO.ListIndex = I
       Next I
       
       '' Tipo Bancos
       For I = 0 To (cboBanco.ListCount - 1)
           If objCADCONTASAPG.CODBANCO = cboBanco.ItemData(I) Then cboBanco.ListIndex = I
       Next I
       
       If IsArray(arrGRIDPGTOS) Then
          
          Label1(7).Caption = Format(UBound(arrGRIDPGTOS), "###000")
       
          For I = 1 To UBound(arrGRIDPGTOS)
          
              sSql = "Select" & vbCrLf
              sSql = sSql & "      ITENS.* " & vbCrLf
              sSql = sSql & "     ,HEADE.* " & vbCrLf
              sSql = sSql & "     ,FORNE.SGI_RAZAOSOC " & vbCrLf
              sSql = sSql & "  From" & vbCrLf
              sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
              sSql = sSql & "     ,SGI_CONTASHAPG HEADE" & vbCrLf
              sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
              sSql = sSql & " Where" & vbCrLf
              sSql = sSql & "      ITENS.SGI_FILIAL  = " & FILIAL & vbCrLf
              sSql = sSql & "  And ITENS.SGI_CODIGO  = " & arrGRIDPGTOS(I, 1) & vbCrLf
              sSql = sSql & "  And ITENS.SGI_PARCELA = " & arrGRIDPGTOS(I, 2) & vbCrLf
              sSql = sSql & "  And HEADE.SGI_FILIAL  = ITENS.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And HEADE.SGI_CODIGO  = ITENS.SGI_CODIGO " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_FILIAL  = HEADE.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_CODIGO  = HEADE.SGI_CODFOR "
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              
              If Not BREC.EOF Then
                 flxLote.AddItem "" & vbTab & _
                                 BREC!SGI_CODIGO & vbTab & _
                                 BREC!SGI_NUMDOC & vbTab & _
                                 Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                 Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                 Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                 BREC!SGI_CODFOR & vbTab & _
                                 BREC!SGI_RAZAOSOC & vbTab & _
                                 BREC!SGI_PARCELA & vbTab & _
                                 "S" & vbTab & _
                                 "Sim"
                                 
              End If
              
              BREC.Close
              
          Next I
       End If
       
    End If
    

End Sub


Private Function VerifCaixa(dtData As Date) As Boolean

    VerifCaixa = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFLXCXHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_DATA   = '" & Format(dtData, "MM/DD/YYYY") & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Existe fluxo de caixa criado !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTLOTE.Text = Format(Now, "DD/MM/YYYY")
       VerifCaixa = True
    End If
    BREC.Close
    
    If VerifCaixa = True Then Exit Function

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFLXCXHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "Order by SGI_DATA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       If dtData < CDate(Format(BREC!SGI_DATA, "DD/MM/YYYY")) Then
          MsgBox "Data de lançamento menor que data do fluxo de caixa !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTLOTE.Text = Format(Now, "DD/MM/YYYY")
          VerifCaixa = True
       End If
    End If
    BREC.Close

End Function

