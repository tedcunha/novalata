VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADCONFFATP 
   Caption         =   "Cadastro de Confirmação de Faturamento"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   13005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   6495
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   12975
      Begin VSFlex8LCtl.VSFlexGrid grdORDCONF 
         Height          =   6135
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   12735
         _cx             =   22463
         _cy             =   10821
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   12975
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   10800
         Top             =   120
      End
      Begin VB.CommandButton cmdOrden 
         Caption         =   "Ordem"
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
         Left            =   12120
         Picture         =   "frmCADCONFFATP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ordena os Registros"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdCanFiltro 
         Caption         =   "Desfas"
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
         Left            =   11400
         Picture         =   "frmCADCONFFATP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Desfas Ultima Pesqusa"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdExclui 
         Caption         =   "&Exclui"
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
         Left            =   2640
         Picture         =   "frmCADCONFFATP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exclui Registro"
         Top             =   120
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
         Left            =   1800
         Picture         =   "frmCADCONFFATP.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Altera Registro"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdInclui 
         Caption         =   "&Inclui"
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
         Picture         =   "frmCADCONFFATP.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Inclui um novo registro"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar"
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
         Picture         =   "frmCADCONFFATP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         Picture         =   "frmCADCONFFATP.frx":129C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   9015
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   200
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo:"
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
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtro:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCADCONFFATP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public strUsuario       As String
Public lngCodUsuaro     As Long
Public intFILIALPED     As Integer

Dim lngCodVendedor      As Long
Dim objFuncoes          As Object
Dim objCADCONFFAT       As Object
Dim objRel              As Object
Dim iCodigo             As Long
Dim boolComAcao         As Boolean
Dim strNOMTABELA1       As String
Dim strNOMTABELA2       As String
Dim strNOMTABELA3       As String
Dim strNOMTABELA4       As String
Dim strNomModulo        As String

Const conCOL_OrdConf_Codigo                    As Integer = 0
Const conCOL_OrdConf_DataOrdem                 As Integer = 1
Const conCOL_OrdConf_Cliente                   As Integer = 2
Const conCOL_OrdConf_Pedido                    As Integer = 3
Const conCOL_OrdConf_CodEmp                    As Integer = 4
Const conCOL_OrdConf_NomeEmp                   As Integer = 5
Const conCOL_OrdConf_FormatString              As String = "=Código|Data Conf|Cliente|Pedido|CodEmp|Empresa"
Const conColumnsIn_OrdConf                     As Integer = 6

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    Call AbilitaCampos
    Call ConfGridOrdConf
    Call PreencheGrid
End Sub

Private Sub cmdExclui_Click()

    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
    If (grdORDCONF.Row) = 0 Then
       MsgBox "Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    Dim iResp     As Integer
    Dim lngCodLog As Long
    
    iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
    If iResp <> 6 Then Exit Sub
    
    
    objCADCONFFAT.CODCONF = CLng(grdORDCONF.Cell(flexcpText, grdORDCONF.Row, conCOL_OrdConf_Codigo))
    objCADCONFFAT.Carrega_Campos
    
    objCADCONFFAT.SALDOFECHADO = CalcSaldoPed
    
    If objCADCONFFAT.GRAVA("E") = False Then Exit Sub
    ''If objFuncoes.Atualiza("E", Str(objCADCONFFAT.CODCONF), FILIAL, "frmCADCONFFAT", Linha) = False Then Exit Sub
    
    lngCodLog = objFuncoes.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objFuncoes.GravaLogModulo(FILIAL, lngCodLog, "frmCADCONFFAT", "E", lngCodUsuaro, Str(objCADCONFFAT.CODCONF), Linha)
    
    
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
    ''Call Atualiza_Grid
    Call ConfGridOrdConf
    Call AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
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

   Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
   Set objCADCONFFAT = CreateObject("CADCONFFAT.clsCADCONFFAT")
   Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
   objCADCONFFAT.FILIAL = FILIAL
   objFuncoes.LimpaCampos frmCADCONFFATP
    
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
    
   cboFiltro.AddItem "Nº Confirmação"
   cboFiltro.AddItem "Data Confirmação"
   cboFiltro.AddItem "Cliente"
   cboFiltro.AddItem "Pedido"
   
   cboFiltro.ListIndex = 0
   
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    boolComAcao = False

    If intFILIALPED = 0 Then
       Me.Caption = Me.Caption & " / NOVALATA"
       strNOMTABELA1 = "SGI_ORDEMPROD"
       strNOMTABELA2 = "SGI_CADPEDVENDH"
       strNOMTABELA3 = "SGI_CADORDCONFH"
       strNOMTABELA4 = "SGI_CADORDFATH"
       strNomModulo = "frmCADCONFFAT"
    ElseIf intFILIALPED = 1 Then
       Me.Caption = Me.Caption & " / STEEL ROLL"
       strNOMTABELA1 = "SGI_ORDEMPROD_STEEL"
       strNOMTABELA2 = "SGI_CADPEDVENDH_STEEL"
       strNOMTABELA3 = "SGI_CADORDCONFH_STEEL"
       strNOMTABELA4 = "SGI_CADORDFATH_STEEL"
       strNomModulo = "frmCADCONFFAT_STEEL"
    End If

   Call AbilitaCampos
   Call ConfGridOrdConf


End Sub

Private Sub Operacao(strOperacao As String)
  
    If (grdORDCONF.Rows - 1) > 0 And grdORDCONF.Row > 0 Then iCodigo = CLng(grdORDCONF.Cell(flexcpText, grdORDCONF.Row, conCOL_OrdConf_Codigo))
    
    boolComAcao = True
    
    frmCADCONFFAT.cCaminho = cCaminho
    frmCADCONFFAT.Linha = Linha
    frmCADCONFFAT.iCodigo = iCodigo
    frmCADCONFFAT.cTipOper = strOperacao
    frmCADCONFFAT.FILIAL = FILIAL
    frmCADCONFFAT.strAcesso = strAcesso
    frmCADCONFFAT.strMODPAI = Me.Name
    frmCADCONFFAT.strUsuario = strUsuario
    frmCADCONFFAT.lngCodVendedor = lngCodVendedor
    frmCADCONFFAT.lngCodUsuario = lngCodUsuaro
    frmCADCONFFAT.intFILIALPED = intFILIALPED
    
    frmCADCONFFAT.Show vbModal
    
    boolComAcao = False
    
    Call ConfGridOrdConf
    Call AbilitaCampos
    
End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADCONFFAT.AtivoDesativo(Trim(strNOMTABELA3))
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdImpressao.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo
    Frame3.Enabled = boolAtivoDesativo

End Sub


Private Sub ConfGridOrdConf()

    With grdORDCONF
    
       .Cols = conColumnsIn_OrdConf
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdConf_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_OrdConf_Codigo) = ""
       .ColDataType(conCOL_OrdConf_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdConf_DataOrdem) = ""
       .ColDataType(conCOL_OrdConf_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdConf_Cliente) = ""
       .ColDataType(conCOL_OrdConf_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdConf_Pedido) = ""
       .ColDataType(conCOL_OrdConf_Pedido) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdConf_CodEmp) = ""
       .ColDataType(conCOL_OrdConf_CodEmp) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdConf_NomeEmp) = ""
       .ColDataType(conCOL_OrdConf_NomeEmp) = flexDTString
       
       .ColWidth(conCOL_OrdConf_Codigo) = 1500
       .ColWidth(conCOL_OrdConf_DataOrdem) = 1000
       .ColWidth(conCOL_OrdConf_Cliente) = 5000
       .ColWidth(conCOL_OrdConf_Pedido) = 1000
       .ColWidth(conCOL_OrdConf_CodEmp) = 0
       .ColWidth(conCOL_OrdConf_NomeEmp) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub PreencheGrid()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       CONF.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDCONFH CONF " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATH  FAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       CONF.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And FAT.SGI_FILIAL  = CONF.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FAT.SGI_CODORD  = CONF.SGI_CODORD " & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL  = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO  = FAT.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL  = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO  = PED.SGI_CODCLI "
    
    sSql = sSql & "Order by CONF.SGI_CODCONF "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       grdORDCONF.AddItem BREC!SGI_CODCONF & vbTab & _
                          Format(BREC!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                          Trim(BREC!SGI_RAZAOSOC) & vbTab & _
                          BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub

Private Sub grdORDCONF_Click()
   If (grdORDCONF.Rows - 1) > 0 And grdORDCONF.Row > 0 Then objCADCONFFAT.CODCONF = CLng(grdORDCONF.Cell(flexcpText, grdORDCONF.Row, conCOL_OrdConf_Codigo))
End Sub

Private Sub grdORDCONF_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdORDCONF.Rows - 1) > 0 Then Operacao "C"
End Sub

Private Sub grdORDCONF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdORDCONF.Rows - 1) > 0 Then Operacao "C"
    End If
End Sub

Private Sub grdORDCONF_RowColChange()
   If (grdORDCONF.Rows - 1) > 0 And grdORDCONF.Row > 0 Then objCADCONFFAT.CODCONF = CLng(grdORDCONF.Cell(flexcpText, grdORDCONF.Row, conCOL_OrdConf_Codigo))
End Sub

Private Sub Atualiza_Grid()
    
     Dim I         As Integer
     Dim bolAchou  As Boolean
     Dim lngCODIGO As Long
     
     bolAchou = False
      
     With grdORDCONF
     
         sSql = "Select" & vbCrLf
         sSql = sSql & "      * " & vbCrLf
         sSql = sSql & "  From" & vbCrLf
         sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
         sSql = sSql & " Where" & vbCrLf
         sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
         sSql = sSql & "   And SGI_MODULO = 'frmCADCONFFAT'" & vbCrLf
    
         BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
         If Not BREC.EOF Then
            I = .FindRow(BREC!SGI_CODIGO, , conCOL_OrdConf_Codigo)
            If I > 0 Then
               If Trim(BREC!SGI_ACAO) = "E" Then
                  If .Rows = 2 Then .Rows = 1
                  If .Rows > 2 Then .RemoveItem I
               ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
                  bolAchou = True
               End If
            End If
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       CONF.* " & vbCrLf
            sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
            sSql = sSql & "      ,PED.SGI_CODIGO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADORDCONFH CONF " & vbCrLf
            sSql = sSql & "      ,SGI_CADORDFATH  FAT " & vbCrLf
            sSql = sSql & "      ,SGI_CADPEDVENDH PED " & vbCrLf
            sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            
            sSql = sSql & "       CONF.SGI_FILIAL  = " & FILIAL & vbCrLf
            sSql = sSql & "   AND CONF.SGI_CODCONF = " & BREC!SGI_CODIGO & vbCrLf
            sSql = sSql & "   And FAT.SGI_FILIAL   = CONF.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And FAT.SGI_CODORD   = CONF.SGI_CODORD " & vbCrLf
            sSql = sSql & "   And PED.SGI_FILIAL   = FAT.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And PED.SGI_CODIGO   = FAT.SGI_CODPED " & vbCrLf
            sSql = sSql & "   And CLI.SGI_FILIAL   = PED.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And CLI.SGI_CODIGO   = PED.SGI_CODCLI "
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF Then
                If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
                   
                      .AddItem BREC2!SGI_CODCONF & vbTab & _
                               Format(BREC2!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                               BREC2!SGI_RAZAOSOC & vbTab & _
                               BREC2!SGI_CODIGO
                
                ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
                
                      .Cell(flexcpText, I, conCOL_OrdConf_Codigo) = BREC2!SGI_CODCONF
                      .Cell(flexcpText, I, conCOL_OrdConf_DataOrdem) = Format(BREC2!SGI_DATACONF, "DD/MM/YYYY")
                      .Cell(flexcpText, I, conCOL_OrdConf_Cliente) = BREC2!SGI_RAZAOSOC
                      .Cell(flexcpText, I, conCOL_OrdConf_Pedido) = BREC2!SGI_CODIGO
                
                End If
            End If
            BREC2.Close

         End If
         BREC.Close
      
     End With
      
End Sub

Private Sub Timer1_Timer()
    Call AbilitaCampos
End Sub

Private Sub Ordem()

    Dim strCampos As String
    
    Call ConfGridOrdConf
    
    txtCampos.Text = ""
    
    sSql = ""
  
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CONF.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA3 & " CONF " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA4 & " FAT " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       CONF.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And FAT.SGI_FILIAL  = CONF.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FAT.SGI_CODORD  = CONF.SGI_CODORD " & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL  = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO  = FAT.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL  = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO  = PED.SGI_CODCLI " & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by CONF.SGI_CODCONF "
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by CONF.SGI_DATACONF "
    If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order by CLI.SGI_RAZAOSOC "
    If cboFiltro.ListIndex = 3 Then sSql = sSql & "Order by PED.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados
    
    strCampos = ""
    Do While Not BREC.EOF
    
        strCampos = BREC!SGI_CODCONF & vbTab & _
                    Format(BREC!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                    BREC!SGI_RAZAOSOC & vbTab & _
                    BREC!SGI_CODIGO & vbTab & _
                    "" & vbTab & _
                    ""
    
        grdORDCONF.AddItem strCampos
       
       BREC.MoveNext
    Loop
    BREC.Close

End Sub
Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADCONFFATP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim strCampos As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CONF.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA3 & " CONF " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA4 & " FAT " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       CONF.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And FAT.SGI_FILIAL  = CONF.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FAT.SGI_CODORD  = CONF.SGI_CODORD " & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL  = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO  = FAT.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL  = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO  = PED.SGI_CODCLI "
    
    If cboFiltro.ListIndex = 0 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And CONF.SGI_CODCONF = " & Trim(txtCampos.Text) & vbCrLf
        sSql = sSql & "Order by CONF.SGI_CODCONF " & vbCrLf
    ElseIf cboFiltro.ListIndex = 1 Then
        If IsDate(txtCampos.Text) = False Then
           MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And CONF.SGI_DATACONF = '" & Format(txtCampos.Text, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "Order by CONF.SGI_DATACONF " & vbCrLf
    ElseIf cboFiltro.ListIndex = 2 Then
       sSql = sSql & "     And CLI.SGI_RAZAOSOC LIKE '" & Trim(txtCampos.Text) & "%'" & vbCrLf
       sSql = sSql & "Order by CLI.SGI_RAZAOSOC " & vbCrLf
    ElseIf cboFiltro.ListIndex = 3 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And PED.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
        sSql = sSql & "Order by PED.SGI_CODIGO " & vbCrLf
    End If
        
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        
        Call ConfGridOrdConf
        
        Do While Not BREC.EOF()
            
            strCampos = BREC!SGI_CODCONF & vbTab & _
                        Format(BREC!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                        BREC!SGI_RAZAOSOC & vbTab & _
                        BREC!SGI_CODIGO & vbTab & _
                        BREC!SGI_FILIALPED & vbTab & _
                        IIf(BREC!SGI_FILIALPED = 0, "NOVALATA", "STEEL ROW")
        
            grdORDCONF.AddItem strCampos
            BREC.MoveNext
        Loop
    Else
        MsgBox "Este Registro não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

End Sub

Private Function CalcSaldoPed() As Boolean

    sSql = "Select " & vbCrLf
    sSql = sSql & "       ORDFA.SGI_CODPED  " & vbCrLf
    sSql = sSql & "      ,CONFH.SGI_CODCONF " & vbCrLf
    sSql = sSql & "      ,ORDFA.SGI_CODORD  " & vbCrLf
    sSql = sSql & "      ,PEDV.SGI_QTDEITENSFATURADOS " & vbCrLf
    sSql = sSql & "      ,CONFH.SGI_QTDETOTFAT " & vbCrLf
    sSql = sSql & "      ,(PEDV.SGI_QTDEITENSFATURADOS - CONFH.SGI_QTDETOTFAT) As SGI_SALDO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA3 & " CONFH " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA4 & " ORDFA " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & " PEDV  " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CONFH.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And CONFH.SGI_CODCONF = " & objCADCONFFAT.CODCONF & vbCrLf
    sSql = sSql & "   And ORDFA.SGI_FILIAL  = CONFH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORDFA.SGI_CODORD  = CONFH.SGI_CODORD " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_FILIAL   = ORDFA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_CODIGO   = ORDFA.SGI_CODPED "
    
    BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC12.EOF() Then
       objCADCONFFAT.CODPED = BREC12!SGI_CODPED
       objCADCONFFAT.CODORD = BREC12!SGI_CODORD
       If BREC12!SGI_SALDO > 0 Then
          CalcSaldoPed = False
          objCADCONFFAT.QTDEATENDPED = BREC12!SGI_SALDO
       ElseIf BREC12!SGI_SALDO <= 0 Then
          CalcSaldoPed = True
          objCADCONFFAT.QTDEATENDPED = BREC12!SGI_SALDO
       End If
    End If
    BREC12.Close

End Function

Private Sub DestroiObjeto()
    If adoBanco_Dados.State = 1 Then adoBanco_Dados.Close
    Set objFuncoes = Nothing
    Set objCADCONFFAT = Nothing
    Set objRel = Nothing
End Sub
