VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADPRODOPP 
   Caption         =   "Controle de OP's Programadas"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   12435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   5895
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   12375
      Begin VSFlex8LCtl.VSFlexGrid grdOPENVIADAS 
         Height          =   5535
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   12135
         _cx             =   21405
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
      Width           =   12375
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   200
         Width           =   2175
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "txtCampos"
         Top             =   200
         Width           =   8415
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
         TabIndex        =   12
         Top             =   240
         Width           =   495
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
         TabIndex        =   11
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   12375
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
         Picture         =   "frmCADPRODOPP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Width           =   735
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
         Picture         =   "frmCADPRODOPP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
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
         Picture         =   "frmCADPRODOPP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Inclui um novo registro"
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
         Picture         =   "frmCADPRODOPP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Altera Registro"
         Top             =   120
         Width           =   855
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
         Picture         =   "frmCADPRODOPP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Registro"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCanFiltro 
         Caption         =   "Limpa"
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
         Left            =   10800
         Picture         =   "frmCADPRODOPP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Desfas Ultima Pesqusa"
         Top             =   120
         Width           =   735
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
         Left            =   11520
         Picture         =   "frmCADPRODOPP.frx":129C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ordena os Registros"
         Top             =   120
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   10200
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmCADPRODOPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public strUSUARIO       As String
Public lngCodUsuaro     As Long
Public lngEMPRESA       As Integer

Dim lngCodVendedor      As Long

Dim objFuncoes          As New clsFuncoes
Dim objCADPRODOPP       As New clsCADPRODOP
Dim objRel              As Object

Dim iCodigo             As Long
Dim RecSet_CADPRODOPP   As New ADODB.Recordset
Dim strEMPRESA          As String
Dim strTABELA           As String

Const conCOL_OPENVIADAS_Codigo                          As Integer = 0
Const conCOL_OPENVIADAS_DatMov                          As Integer = 1
Const conCOL_OPENVIADAS_FormatString                    As String = "=Cód.Doc|Dt.Programação"
Const conColumnsIn_OPENVIADAS                           As Integer = 2


Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    With grdOPENVIADAS
        If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("A")
    End With
End Sub

Private Sub cmdCanFiltro_Click()
   txtCampos.Text = ""
   Call ConfGrid
   Call AbilitaCampos
End Sub

Private Sub cmdExclui_Click()

    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
    
    With grdOPENVIADAS
        If (.Rows - 1) = 0 Or (.Row = 0) Then Exit Sub
        objCADPRODOPP.Codigo = CLng(.Cell(flexcpText, .Row, conCOL_OPENVIADAS_Codigo))
    End With
    
    Dim iResp     As Integer
    Dim lngCodLog As Long
    
    Beep
    iResp = MsgBox("Confirma a exclusão do registro ? [ " & objCADPRODOPP.Codigo & " ]", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
    If iResp <> 6 Then Exit Sub
    
    objCADPRODOPP.Carrega_Campos (strTABELA)
  
    If objCADPRODOPP.GRAVA("E", strTABELA) = False Then Exit Sub
    If objFuncoes.Atualiza("E", Str(objCADPRODOPP.Codigo), FILIAL, "frmCADPRODOP" & strTABELA, Linha) = False Then Exit Sub
    
    lngCodLog = objFuncoes.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objFuncoes.GravaLogModulo(FILIAL, lngCodLog, "frmCADPRODOP", "E", lngCodUsuaro, Str(objCADPRODOPP.Codigo), Linha)
    
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
    Call Atualiza_Grid
    Call AbilitaCampos
    objCADPRODOPP.Codigo = 0

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

Private Sub Destroy_Objeto()
    Set objFuncoes = Nothing
    Set objCADPRODOPP = Nothing
    Set objRel = Nothing
    Set RecSet_CADPRODOPP = Nothing
End Sub

Private Sub Form_Load()

    ''Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    ''Set objCADPRODOPP = CreateObject("CADPRODOP.clsCADPRODOP")
    Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
     
    objCADPRODOPP.FILIAL = FILIAL
    objFuncoes.LimpaCampos frmCADPRODOPP
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If

    strEMPRESA = "NOVALATA"
    strTABELA = ""
    If lngEMPRESA = 1 Then
        strEMPRESA = "STEEL"
        strTABELA = "_STEEL"
    End If
   
   
    Call AbilitaCampos
    Call ConfGrid
     
    cboFiltro.AddItem "Nº Doc."
    cboFiltro.AddItem "Dt.Programação"
    
    cboFiltro.ListIndex = 0
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
     
     Me.Caption = Me.Caption & " / " & strEMPRESA

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADPRODOPP.AtivoDesativo(strTABELA)
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdImpressao.Enabled = boolAtivoDesativo
    cmdCanFiltro.Enabled = boolAtivoDesativo
    cmdOrden.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo
    Frame3.Enabled = boolAtivoDesativo

End Sub

Private Sub ConfGrid()

    With grdOPENVIADAS
    
       .Cols = conColumnsIn_OPENVIADAS
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OPENVIADAS_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_OPENVIADAS_Codigo) = ""
       .ColDataType(conCOL_OPENVIADAS_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OPENVIADAS_DatMov) = ""
       .ColDataType(conCOL_OPENVIADAS_DatMov) = flexDTDate
       
       .ColWidth(conCOL_OPENVIADAS_Codigo) = 0
       .ColWidth(conCOL_OPENVIADAS_DatMov) = 1300
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub Operacao(strOperacao As String)
  
    With grdOPENVIADAS
        If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_OPENVIADAS_Codigo))
    End With
    
    If strOperacao = "I" Then iCodigo = 0
    
    frmCADPRODOP.cCaminho = cCaminho
    frmCADPRODOP.Linha = Linha
    frmCADPRODOP.iCodigo = iCodigo
    frmCADPRODOP.cTipOper = strOperacao
    frmCADPRODOP.FILIAL = FILIAL
    frmCADPRODOP.strAcesso = strAcesso
    frmCADPRODOP.strMODPAI = Me.Name
    frmCADPRODOP.strUSUARIO = strUSUARIO
    frmCADPRODOP.lngCodVendedor = lngCodVendedor
    frmCADPRODOP.lngCodUsuario = lngCodUsuaro
    frmCADPRODOP.strEMPRESA = strEMPRESA
    frmCADPRODOP.strTABELA = strTABELA
    frmCADPRODOP.Show vbModal
    
    ''Call Atualiza_Grid
    Call ConfGrid
    Call AbilitaCampos

End Sub


Private Sub PreencheGrid()
    
    If RecSet_CADPRODOPP.State = 1 Then RecSet_CADPRODOPP.Close
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADOPENVIADAH " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    RecSet_CADPRODOPP.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not RecSet_CADPRODOPP.EOF() Then
       With grdOPENVIADAS
            Do While Not RecSet_CADPRODOPP.EOF()
                .AddItem RecSet_CADPRODOPP!SGI_CODIGO & vbTab & _
                         Format(RecSet_CADPRODOPP!SGI_DTLANCTO, "DD/MM/YYYY") & vbTab & _
                         RecSet_CADPRODOPP!SGI_QTDEOP
                RecSet_CADPRODOPP.MoveNext
            Loop
       End With
    End If
    RecSet_CADPRODOPP.Close

End Sub

Private Sub grdOPENVIADAS_Click()
   With grdOPENVIADAS
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPRODOPP.Codigo = CLng(.Cell(flexcpText, .Row, conCOL_OPENVIADAS_Codigo))
   End With
End Sub

Private Sub grdOPENVIADAS_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdOPENVIADAS.Rows - 1) > 0 And grdOPENVIADAS.Row > 0 Then Call Operacao("C")
End Sub

Private Sub grdOPENVIADAS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdOPENVIADAS.Rows - 1) > 0 And grdOPENVIADAS.Row > 0 Then Call Operacao("C")
    End If
End Sub

Private Sub grdOPENVIADAS_RowColChange()
   With grdOPENVIADAS
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPRODOPP.Codigo = CLng(.Cell(flexcpText, .Row, conCOL_OPENVIADAS_Codigo))
   End With
End Sub

Private Sub Timer1_Timer()
    Call Atualiza_Grid
    Call AbilitaCampos
End Sub


Private Sub Atualiza_Grid()
    
     Dim I              As Long
     Dim bolAchou       As Boolean
     Dim lngCODIGO      As Long
     Dim strACAO        As String
     Dim lngCOL         As Long
     Dim grdGenerica    As VSFlexGrid
     Dim strCAMPOS      As String
     
     If BRECATU.State = 1 Then BRECATU.Close
     
     lngCOL = conCOL_OPENVIADAS_Codigo
     Set grdGenerica = grdOPENVIADAS
     bolAchou = False
      
     With grdGenerica
     
        sSql = "Select" & vbCrLf
        sSql = sSql & "      * " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_MODULO = 'frmCADMOVPCP'" & vbCrLf
        
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
           ElseIf strACAO = "A" Then
           ''ElseIf strACAO = "I" Or strACAO = "A" Then
              bolAchou = True
           End If
        End If
        
        If bolAchou = True And Trim(strACAO) = "A" Then
            
            sSql = ""
        
            sSql = "Select Distinct" & vbCrLf
            sSql = sSql & "       OPENV.SGI_CODIGO " & vbCrLf
            sSql = sSql & "      ,OPENV.SGI_DATAPROG " & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADMOVPCP" & strTABELA & "  OPENV" & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            
            sSql = sSql & "       OPENV.SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "   And OPENV.SGI_CODIGO = " & lngCODIGO
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF Then
                .Cell(flexcpText, I, conCOL_OPENVIADAS_Codigo) = BREC2!SGI_CODIGO
                .Cell(flexcpText, I, conCOL_OPENVIADAS_DatMov) = Format(BREC2!SGI_DATAPROG, "DD/MM/YYYY")
            End If
            BREC2.Close
        End If
     
     End With
      
End Sub


Private Sub Ordem()

    Dim strCAMPOS As String
    
    Call ConfGrid
    
    txtCampos.Text = ""
    
    sSql = ""
  
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       OPENV.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,OPENV.SGI_DATAPONT" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strTABELA & " OPENV " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       OPENV.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by OPENV.SGI_CODIGO"
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by OPENV.SGI_DATAPONT"
    
    BREC.Open sSql, adoBanco_Dados
    
    strCAMPOS = ""
    Do While Not BREC.EOF
    
        strCAMPOS = BREC!SGI_CODIGO & vbTab & _
                    Format(BREC!SGI_DATAPONT, "DD/MM/YYYY")
    
        grdOPENVIADAS.AddItem strCAMPOS
       
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADPRODOPP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim strCAMPOS As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    Call ConfGrid
    
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       OPENV.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,OPENV.SGI_DATAPROG" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strTABELA & "  OPENV " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       OPENV.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And OPENV.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
    ElseIf cboFiltro.ListIndex = 1 Then
        If IsDate(txtCampos.Text) = False Then
           MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And OPENV.SGI_DATAPROG = '" & Format(txtCampos.Text, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "Order by OPENV.SGI_DATAPROG " & vbCrLf
    End If
        
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then
        
        Do While Not BREC2.EOF()
            
            strCAMPOS = BREC2!SGI_CODIGO & vbTab & _
                        Format(BREC2!SGI_DATAPROG, "DD/MM/YYYY")
        
            grdOPENVIADAS.AddItem strCAMPOS
            
            BREC2.MoveNext
        Loop
    Else
        MsgBox "Este Registro não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC2.Close

End Sub

