VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADAPONTPRODP 
   Caption         =   "Apontamento de Produção"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   12660
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   6015
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   12615
      Begin TabDlg.SSTab stApont 
         Height          =   5655
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   9975
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Em Aberto"
         TabPicture(0)   =   "frmCADAPONTPRODP.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdAPONTPROD"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Fechado"
         TabPicture(1)   =   "frmCADAPONTPRODP.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdFECHADO"
         Tab(1).ControlCount=   1
         Begin VSFlex8LCtl.VSFlexGrid grdFECHADO 
            Height          =   4935
            Left            =   -74880
            TabIndex        =   16
            Top             =   480
            Width           =   12135
            _cx             =   21405
            _cy             =   8705
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
         Begin VSFlex8LCtl.VSFlexGrid grdAPONTPROD 
            Height          =   5055
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   12135
            _cx             =   21405
            _cy             =   8916
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
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12615
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   200
         Width           =   1695
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "txtCampos"
         Top             =   200
         Width           =   9135
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
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   12615
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
         Picture         =   "frmCADAPONTPRODP.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprime Registro"
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
         Left            =   11640
         Picture         =   "frmCADAPONTPRODP.frx":013A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   855
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
         Left            =   10800
         Picture         =   "frmCADAPONTPRODP.frx":023C
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADAPONTPRODP.frx":076E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmCADAPONTPRODP.frx":0870
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Altera Empresa "
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
         Picture         =   "frmCADAPONTPRODP.frx":0972
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Inclui uma nova empresa"
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
         Picture         =   "frmCADAPONTPRODP.frx":0EA4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   6840
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmCADAPONTPRODP"
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
Public lngCodUsuario    As Long

Dim objFuncoes          As Object
Dim objCADAPONTPRODP    As Object
Dim objRel              As Object
Dim iCodigo             As Long

Const conCOL_APONT_Codigo                          As Integer = 0
Const conCOL_APONT_CodOP                           As Integer = 1
Const conCOL_APONT_DatMov                          As Integer = 2
Const conCOL_APONT_FormatString                    As String = "=Cód.Doc|Cod.OP|Dt.Doc"
Const conColumnsIn_APONT                           As Integer = 3

Const conCOL_APONTFC_Codigo                        As Integer = 0
Const conCOL_APONTFC_CodOP                         As Integer = 1
Const conCOL_APONTFC_DatMov                        As Integer = 2
Const conCOL_APONTFC_FormatString                  As String = "=Cód.Doc|Cod.OP|Dt.Doc"
Const conColumnsIn_APONTFC                         As Integer = 3

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
   txtCampos.Text = ""
   Call ConfGrid
   Call ConfGridFC
   Call AbilitaCampos
End Sub

Private Sub cmdExclui_Click()

    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
    
    If stApont.Tab = 0 Then
        With grdAPONTPROD
            If (.Rows - 1) = 0 Or (.Row = 0) Then Exit Sub
            objCADAPONTPRODP.CODIGO = CLng(.Cell(flexcpText, .Row, conCOL_APONT_Codigo))
        End With
    ElseIf stApont.Tab = 1 Then
        With grdFECHADO
            If (.Rows - 1) = 0 Or (.Row = 0) Then Exit Sub
            objCADAPONTPRODP.CODIGO = CLng(.Cell(flexcpText, .Row, conCOL_APONTFC_Codigo))
        End With
    End If
    
    Dim iResp     As Integer
    Dim lngCodLog As Long
    
    iResp = MsgBox("Confirma a exclusão do registro ? [ " & objCADAPONTPRODP.CODIGO & " ]", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
    If iResp <> 6 Then Exit Sub
    
    objCADAPONTPRODP.Carrega_Campos
  
    If objCADAPONTPRODP.GRAVA("E") = False Then Exit Sub
    If objFuncoes.Atualiza("E", Str(objCADAPONTPRODP.CODIGO), FILIAL, "frmCADAPONTPROD", Linha) = False Then Exit Sub
    
    lngCodLog = objFuncoes.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objFuncoes.GravaLogModulo(FILIAL, lngCodLog, "frmCADAPONTPROD", "E", lngCodUsuario, Str(objCADAPONTPRODP.CODIGO), Linha)
    
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
    Call ConfGrid
    Call ConfGridFC
    Call AbilitaCampos
    objCADAPONTPRODP.CODIGO = 0
    
    stApont.Tab = 0

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    If stApont.Tab = 1 Then Exit Sub
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

Private Sub Destroy_Objeto()
    Set objFuncoes = Nothing
    Set objCADAPONTPRODP = Nothing
    Set objRel = Nothing
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADAPONTPRODP = CreateObject("CADAPONTPROD.clsCADAPONTPROD")
    Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    objCADAPONTPRODP.FILIAL = FILIAL
    objFuncoes.LimpaCampos frmCADAPONTPRODP
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If

    Call AbilitaCampos
    Call ConfGrid
    Call ConfGridFC
    
    cboFiltro.AddItem "Nº Apontamento"
    cboFiltro.AddItem "Cód.OP"
    cboFiltro.AddItem "Dt.Apontamento"
    cboFiltro.ListIndex = 0
   
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    stApont.Tab = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADAPONTPRODP.AtivoDesativo
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdImpressao.Enabled = boolAtivoDesativo
    cmdCanFiltro.Enabled = boolAtivoDesativo
    cmdOrden.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo
    Frame3.Enabled = boolAtivoDesativo

End Sub

Private Sub ConfGrid()

    With grdAPONTPROD
    
       .Cols = conColumnsIn_APONT
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_APONT_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_APONT_Codigo) = ""
       .ColDataType(conCOL_APONT_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_APONT_CodOP) = ""
       .ColDataType(conCOL_APONT_CodOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_APONT_DatMov) = ""
       .ColDataType(conCOL_APONT_DatMov) = flexDTDate
       
       .ColWidth(conCOL_APONT_Codigo) = 1200
       .ColWidth(conCOL_APONT_CodOP) = 1200
       .ColWidth(conCOL_APONT_DatMov) = 1200
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub Operacao(strOperacao As String)
  
    If stApont.Tab = 0 Then
        With grdAPONTPROD
            If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_APONT_Codigo))
        End With
    ElseIf stApont.Tab = 1 Then
        With grdFECHADO
            If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_APONTFC_Codigo))
        End With
    End If
    
    If strOperacao = "I" Then iCodigo = 0
    
    frmCADAPONTPROD.cCaminho = cCaminho
    frmCADAPONTPROD.Linha = Linha
    frmCADAPONTPROD.iCodigo = iCodigo
    frmCADAPONTPROD.cTipOper = strOperacao
    frmCADAPONTPROD.FILIAL = FILIAL
    frmCADAPONTPROD.strAcesso = strAcesso
    frmCADAPONTPROD.strMODPAI = Me.Name
    frmCADAPONTPROD.strUsuario = strUsuario
    frmCADAPONTPROD.lngCodUsuario = lngCodUsuario
    frmCADAPONTPROD.Show vbModal
    
    Call ConfGrid
    Call ConfGridFC
    Call AbilitaCampos

End Sub


Private Sub Ordem()

    Dim strCampos As String
    
    If stApont.Tab = 0 Then Call ConfGrid
    If stApont.Tab = 1 Then Call ConfGridFC
    
    txtCampos.Text = ""
    
    sSql = ""
  
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRDH.* " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPRDH PRDH" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRDH.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If stApont.Tab = 0 Then sSql = sSql & "   And PRDH.SGI_STATUS = 0" & vbCrLf
    If stApont.Tab = 1 Then sSql = sSql & "   And PRDH.SGI_STATUS = 1" & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by PRDH.SGI_CODIGO "
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by PRDH.SGI_CODOP "
    If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order by PRDH.SGI_DTDOC "
    
    BREC.Open sSql, adoBanco_Dados
    
    strCampos = ""
    Do While Not BREC.EOF
    
        strCampos = BREC!SGI_CODIGO & vbTab & _
                    BREC!SGI_CODOP & vbTab & _
                    Format(BREC!SGI_DTDOC, "DD/MM/YYYY")
    
        If stApont.Tab = 0 Then grdAPONTPROD.AddItem strCampos
        If stApont.Tab = 1 Then grdFECHADO.AddItem strCampos
       
       BREC.MoveNext
    Loop
    BREC.Close

End Sub


Private Sub grdAPONTPROD_Click()
   With grdAPONTPROD
        If (.Rows - 1) > 0 And .Row > 0 Then objCADAPONTPRODP.CODIGO = CLng(.Cell(flexcpText, .Row, conCOL_APONT_Codigo))
   End With
End Sub

Private Sub grdAPONTPROD_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdAPONTPROD.Rows - 1) > 0 And grdAPONTPROD.Row > 0 Then Call Operacao("C")
End Sub

Private Sub grdAPONTPROD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdAPONTPROD.Rows - 1) > 0 And grdAPONTPROD.Row > 0 Then Call Operacao("C")
    End If
End Sub

Private Sub grdAPONTPROD_RowColChange()
   With grdAPONTPROD
        If (.Rows - 1) > 0 And .Row > 0 Then objCADAPONTPRODP.CODIGO = CLng(.Cell(flexcpText, .Row, conCOL_APONT_Codigo))
   End With
End Sub

Private Sub grdFECHADO_Click()
   With grdFECHADO
        If (.Rows - 1) > 0 And .Row > 0 Then objCADAPONTPRODP.CODIGO = CLng(.Cell(flexcpText, .Row, conCOL_APONTFC_Codigo))
   End With
End Sub

Private Sub grdFECHADO_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdFECHADO.Rows - 1) > 0 And grdFECHADO.Row > 0 Then Call Operacao("C")
End Sub

Private Sub grdFECHADO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdFECHADO.Rows - 1) > 0 And grdFECHADO.Row > 0 Then Call Operacao("C")
    End If
End Sub

Private Sub grdFECHADO_RowColChange()
   With grdFECHADO
        If (.Rows - 1) > 0 And .Row > 0 Then objCADAPONTPRODP.CODIGO = CLng(.Cell(flexcpText, .Row, conCOL_APONTFC_Codigo))
   End With
End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADAPONTPRODP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim strCampos As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    Call ConfGrid
    Call ConfGridFC
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRDH.* " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPRDH  PRDH " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRDH.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If stApont.Tab = 0 Then sSql = sSql & "   And PRDH.SGI_STATUS = 0" & vbCrLf
    If stApont.Tab = 1 Then sSql = sSql & "   And PRDH.SGI_STATUS = 1" & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And PRDH.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
    ElseIf cboFiltro.ListIndex = 1 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And PRDH.SGI_CODOP = " & Trim(txtCampos.Text) & vbCrLf
    ElseIf cboFiltro.ListIndex = 2 Then
        If IsDate(txtCampos.Text) = False Then
           MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And PRDH.SGI_DTDOC = '" & Format(txtCampos.Text, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "Order by PRDH.SGI_DTDOC " & vbCrLf
    End If
        
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then
        
        Do While Not BREC2.EOF()
            
            strCampos = BREC2!SGI_CODIGO & vbTab & _
                        BREC2!SGI_CODOP & vbTab & _
                        Format(BREC2!SGI_DTDOC, "DD/MM/YYYY")

            If stApont.Tab = 0 Then grdAPONTPROD.AddItem strCampos
            If stApont.Tab = 1 Then grdFECHADO.AddItem strCampos
            
            BREC2.MoveNext
        Loop
    Else
        MsgBox "Este Registro não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC2.Close

End Sub

Private Sub ConfGridFC()

    With grdFECHADO
    
       .Cols = conColumnsIn_APONTFC
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_APONTFC_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_APONTFC_Codigo) = ""
       .ColDataType(conCOL_APONTFC_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_APONTFC_CodOP) = ""
       .ColDataType(conCOL_APONTFC_CodOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_APONTFC_DatMov) = ""
       .ColDataType(conCOL_APONTFC_DatMov) = flexDTDate
       
       .ColWidth(conCOL_APONTFC_Codigo) = 1200
       .ColWidth(conCOL_APONTFC_CodOP) = 1200
       .ColWidth(conCOL_APONTFC_DatMov) = 1200
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

