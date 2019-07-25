VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADDADOSOPP 
   Caption         =   "Cadastra Dados da OP"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   13020
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   5775
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   12975
      Begin VSFlex8LCtl.VSFlexGrid grdDADOSOP 
         Height          =   5535
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   12735
         _cx             =   22463
         _cy             =   9763
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
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12975
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "txtCampos"
         Top             =   200
         Width           =   9015
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   9
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
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
         Picture         =   "frmCADDADOSOPP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADDADOSOPP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADDADOSOPP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADDADOSOPP.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCADDADOSOPP.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADDADOSOPP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADDADOSOPP.frx":129C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADDADOSOPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public FILIAL        As Integer
Public strAcesso     As String
Public strUsuario    As String
Public lngCodUsuaro  As Long
Dim lngCodVendedor   As Long
Dim objFuncoes       As Object
Dim objCADDADOSOPP   As Object
Dim objRel           As Object
Dim iCodigo          As Long
Dim boolComAcao      As Boolean

Const conCOL_OrdProd_Codigo                    As Integer = 0
Const conCOL_OrdProd_DataOrdem                 As Integer = 1
Const conCOL_OrdProd_CodOP                     As Integer = 2
Const conCOL_OrdProd_CodPed                    As Integer = 3
Const conCOL_OrdProd_Cliente                   As Integer = 4
Const conCOL_OrdProd_FormatString              As String = "=Código|Data|Cod.OP|Cod.Ped|Cliente"
Const conColumnsIn_OrdProd                     As Integer = 5

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    Call AbilitaCampos
    Call ConfGridOP
    Call PreencheGrid
End Sub

Private Sub cmdExclui_Click()

    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
    If (grdDADOSOP.Row) = 0 Then
       MsgBox "Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    Dim iResp     As Integer
    Dim lngCodLog As Long
    
    iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
    If iResp <> 6 Then Exit Sub
    
    objCADDADOSOPP.CODIGO = CLng(grdDADOSOP.Cell(flexcpText, grdDADOSOP.Row, conCOL_OrdProd_Codigo))
    
    If objCADDADOSOPP.GRAVA("E") = False Then Exit Sub
    If objFuncoes.Atualiza("E", Str(objCADDADOSOPP.CODIGO), FILIAL, "frmCADDADOSOP", Linha) = False Then Exit Sub
    
    lngCodLog = objFuncoes.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objFuncoes.GravaLogModulo(FILIAL, lngCodLog, "frmCADDADOSOP", "E", lngCodUsuaro, Str(objCADDADOSOPP.CODIGO), Linha)
    
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
    Call Atualiza_Grid
    Call AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub FechaOBJ()
    Set objFuncoes = Nothing
    Set objCADDADOSOPP = Nothing
    Set objRel = Nothing
End Sub

Private Sub Form_Load()

   Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
   Set objCADDADOSOPP = CreateObject("CADDADOSOP.clsCADDADOSOP")
   Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
   objCADDADOSOPP.FILIAL = FILIAL
   objFuncoes.LimpaCampos frmCADDADOSOPP
    
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
    
   Call AbilitaCampos
   Call ConfGridOP
   Call PreencheGrid
    
   Call ConfFiltro
   
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    boolComAcao = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FechaOBJ
End Sub

Private Sub ConfFiltro()

   cboFiltro.Clear
   
   cboFiltro.AddItem "Código"
   cboFiltro.AddItem "Data"
   cboFiltro.AddItem "Cód.OP"
   cboFiltro.AddItem "Cód.Ped"
   cboFiltro.AddItem "Cliente"
   
   cboFiltro.ListIndex = 0

End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADDADOSOPP.AtivoDesativo
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdImpressao.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo
    Frame3.Enabled = boolAtivoDesativo

End Sub

Private Sub ConfGridOP()

    With grdDADOSOP
    
       .Cols = conColumnsIn_OrdProd
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdProd_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_OrdProd_Codigo) = ""
       .ColDataType(conCOL_OrdProd_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdProd_DataOrdem) = ""
       .ColDataType(conCOL_OrdProd_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdProd_CodOP) = ""
       .ColDataType(conCOL_OrdProd_CodOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdProd_CodPed) = ""
       .ColDataType(conCOL_OrdProd_CodPed) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdProd_Cliente) = ""
       .ColDataType(conCOL_OrdProd_Cliente) = flexDTString
       
       .ColWidth(conCOL_OrdProd_Codigo) = 1000
       .ColWidth(conCOL_OrdProd_DataOrdem) = 1000
       .ColWidth(conCOL_OrdProd_CodOP) = 1000
       .ColWidth(conCOL_OrdProd_CodPed) = 1000
       .ColWidth(conCOL_OrdProd_Cliente) = 5000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub Operacao(strOperacao As String)
  
    With grdDADOSOP
        If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_OrdProd_Codigo))
    End With
    
    boolComAcao = True
    
    frmCADDADOSOP.cCaminho = cCaminho
    frmCADDADOSOP.Linha = Linha
    frmCADDADOSOP.iCodigo = iCodigo
    frmCADDADOSOP.cTipOper = strOperacao
    frmCADDADOSOP.FILIAL = FILIAL
    frmCADDADOSOP.strAcesso = strAcesso
    frmCADDADOSOP.strMODPAI = Me.Name
    frmCADDADOSOP.strUsuario = strUsuario
    frmCADDADOSOP.lngCodVendedor = lngCodVendedor
    frmCADDADOSOP.lngCodUsuario = lngCodUsuaro
    frmCADDADOSOP.Show vbModal
    
    boolComAcao = False
    
    Call Atualiza_Grid
    Call AbilitaCampos

End Sub

Private Sub PreencheGrid()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       DADOSOP.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,DADOSOP.SGI_DATALOTE " & vbCrLf
    sSql = sSql & "      ,DADOSOP.SGI_CODOP " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDADOSOP DADOSOP " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD ORDP " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH PEDV " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE CLIE " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       DADOSOP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_FILIAL = DADOSOP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO = DADOSOP.SGI_CODOP " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_FILIAL = ORDP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_CODIGO = ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = PEDV.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = PEDV.SGI_CODCLI " & vbCrLf
    
    sSql = sSql & "Order by DADOSOP.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    With grdDADOSOP
    
        Do While Not BREC.EOF
           .AddItem BREC!SGI_CODIGO & vbTab & _
                    Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                    BREC!SGI_CODOP & vbTab & _
                    BREC!SGI_CODPED & vbTab & _
                    Trim(BREC!SGI_RAZAOSOC)
                    
           BREC.MoveNext
        Loop
    
    End With
    BREC.Close
    
End Sub

Private Sub grdDADOSOP_Click()
   If (grdDADOSOP.Rows - 1) > 0 And grdDADOSOP.Row > 0 Then objCADDADOSOPP.CODIGO = CLng(grdDADOSOP.Cell(flexcpText, grdDADOSOP.Row, conCOL_OrdProd_Codigo))
End Sub

Private Sub grdDADOSOP_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdDADOSOP.Rows - 1) > 0 And grdDADOSOP.Row > 0 Then Operacao "C"
End Sub

Private Sub grdDADOSOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdDADOSOP.Rows - 1) > 0 And grdDADOSOP.Row > 0 Then Operacao "C"
    End If
End Sub

Private Sub grdDADOSOP_RowColChange()
   If (grdDADOSOP.Rows - 1) > 0 And grdDADOSOP.Row > 0 Then objCADDADOSOPP.CODIGO = CLng(grdDADOSOP.Cell(flexcpText, grdDADOSOP.Row, conCOL_OrdProd_Codigo))
End Sub

Private Sub Timer1_Timer()
    Call AbilitaCampos
    Call Atualiza_Grid
End Sub

Private Sub Atualiza_Grid()
    
    Dim I              As Long
    Dim bolAchou       As Boolean
    Dim lngCODIGO      As Long
    Dim strACAO        As String
    Dim lngCOL         As Long
    Dim grdGENERICA    As VSFlexGrid
    Dim strCampos      As String
     
    If boolComAcao = True Then Exit Sub
    
    If BRECATU.State = 1 Then BRECATU.Close
     
    lngCOL = conCOL_OrdProd_Codigo
    Set grdGENERICA = grdDADOSOP
     
    bolAchou = False
      
    With grdGENERICA
     
        sSql = "Select" & vbCrLf
        sSql = sSql & "      * " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_MODULO = 'frmCADDADOSOP'" & vbCrLf
        
        BRECATU.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BRECATU.EOF Then
           lngCODIGO = BRECATU!SGI_CODIGO
           strACAO = Trim(BRECATU!SGI_ACAO)
        End If
        BRECATU.Close
         
        I = .FindRow(lngCODIGO, , lngCOL)
        If I > 0 Then
           If strACAO = "E" Then
              If .Rows = 2 Then .Rows = 1
              If .Rows > 2 Then .RemoveItem I
           ElseIf strACAO = "I" Or strACAO = "A" Then
              bolAchou = True
           End If
        End If
            
        sSql = "Select " & vbCrLf
        sSql = sSql & "       DADOSOP.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,DADOSOP.SGI_DATALOTE " & vbCrLf
        sSql = sSql & "      ,DADOSOP.SGI_CODOP " & vbCrLf
        sSql = sSql & "      ,ORDP.SGI_CODPED " & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_RAZAOSOC " & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADDADOSOP DADOSOP " & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD ORDP " & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH PEDV " & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE CLIE " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       DADOSOP.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And DADOSOP.SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "   And ORDP.SGI_FILIAL = DADOSOP.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ORDP.SGI_CODIGO = DADOSOP.SGI_CODOP " & vbCrLf
        sSql = sSql & "   And PEDV.SGI_FILIAL = ORDP.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PEDV.SGI_CODIGO = ORDP.SGI_CODPED " & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL = PEDV.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO = PEDV.SGI_CODCLI " & vbCrLf
        sSql = sSql & "Order by DADOSOP.SGI_CODIGO "
            
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF Then
            If bolAchou = False And Trim(strACAO) = "I" Then
               
                    strCampos = BREC2!SGI_CODIGO & vbTab & _
                                Format(BREC2!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                BREC2!SGI_CODOP & vbTab & _
                                BREC2!SGI_CODPED & vbTab & _
                                Trim(BREC2!SGI_RAZAOSOC)
                    .AddItem strCampos
            
            ElseIf bolAchou = True And Trim(strACAO) = "A" Then
                    .Cell(flexcpText, I, conCOL_OrdProd_Codigo) = BREC2!SGI_CODIGO
                    .Cell(flexcpText, I, conCOL_OrdProd_DataOrdem) = BREC2!SGI_DATALOTE
                    .Cell(flexcpText, I, conCOL_OrdProd_CodOP) = BREC2!SGI_CODOP
                    .Cell(flexcpText, I, conCOL_OrdProd_CodPed) = BREC2!SGI_CODPED
                    .Cell(flexcpText, I, conCOL_OrdProd_Cliente) = BREC2!SGI_RAZAOSOC
            End If
        End If
        BREC2.Close

     End With
      
End Sub

Private Sub Ordem()

    Dim strCampos As String
    
    Call ConfGridOP
    
    txtCampos.Text = ""
    
    sSql = ""
  
    sSql = "Select " & vbCrLf
    sSql = sSql & "       DADOSOP.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,DADOSOP.SGI_DATALOTE " & vbCrLf
    sSql = sSql & "      ,DADOSOP.SGI_CODOP " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDADOSOP DADOSOP " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD ORDP " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH PEDV " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE CLIE " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       DADOSOP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_FILIAL = DADOSOP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO = DADOSOP.SGI_CODOP " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_FILIAL = ORDP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_CODIGO = ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = PEDV.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = PEDV.SGI_CODCLI " & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by DADOSOP.SGI_CODIGO "
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by DADOSOP.SGI_DATALOTE "
    If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order by DADOSOP.SGI_CODOP "
    If cboFiltro.ListIndex = 3 Then sSql = sSql & "Order by ORDP.SGI_CODPED "
    If cboFiltro.ListIndex = 4 Then sSql = sSql & "Order by CLIE.SGI_RAZAOSOC "
    
    BREC.Open sSql, adoBanco_Dados
    
    strCampos = ""
    Do While Not BREC.EOF
    
        strCampos = BREC!SGI_CODIGO & vbTab & _
                    Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                    BREC!SGI_CODOP & vbTab & _
                    BREC!SGI_CODPED & vbTab & _
                    Trim(BREC!SGI_RAZAOSOC)
    
        grdDADOSOP.AddItem strCampos
       
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADDADOSOP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim strCampos As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       DADOSOP.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,DADOSOP.SGI_DATALOTE " & vbCrLf
    sSql = sSql & "      ,DADOSOP.SGI_CODOP " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDADOSOP DADOSOP " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD ORDP " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH PEDV " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE CLIE " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       DADOSOP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_FILIAL = DADOSOP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO = DADOSOP.SGI_CODOP " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_FILIAL = ORDP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_CODIGO = ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = PEDV.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = PEDV.SGI_CODCLI " & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And DADOSOP.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
        sSql = sSql & "Order by DADOSOP.SGI_CODIGO " & vbCrLf
    ElseIf cboFiltro.ListIndex = 1 Then
        If IsDate(txtCampos.Text) = False Then
           MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And DADOSOP.SGI_DATALOTE = '" & Format(txtCampos.Text, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "Order by DADOSOP.SGI_DATALOTE " & vbCrLf
    ElseIf cboFiltro.ListIndex = 2 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And DADOSOP.SGI_CODOP = " & Trim(txtCampos.Text) & vbCrLf
        sSql = sSql & "Order by DADOSOP.SGI_CODOP " & vbCrLf
    ElseIf cboFiltro.ListIndex = 3 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And ORDP.SGI_CODPED = " & Trim(txtCampos.Text) & vbCrLf
        sSql = sSql & "Order by ORDP.SGI_CODPED " & vbCrLf
    ElseIf cboFiltro.ListIndex = 4 Then
        sSql = sSql & "     And CLIE.SGI_RAZAOSOC Like '" & Trim(txtCampos.Text) & "%'" & vbCrLf
        sSql = sSql & "Order by CLIE.SGI_RAZAOSOC " & vbCrLf
    End If
        
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        
        Call ConfGridOP
        
        Do While Not BREC.EOF()
            
            strCampos = BREC!SGI_CODIGO & vbTab & _
                        Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                        BREC!SGI_CODOP & vbTab & _
                        BREC!SGI_CODPED & vbTab & _
                        Trim(BREC!SGI_RAZAOSOC)
        
            grdDADOSOP.AddItem strCampos
            BREC.MoveNext
        Loop
    Else
        MsgBox "Este Registro não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCampos.Text = ""
        txtCampos.SetFocus
    End If
    BREC.Close

End Sub
