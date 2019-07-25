VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADRECMATLITP 
   Caption         =   "Recebimento de Material (Folhas Litografadas)"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   16005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   0
      TabIndex        =   23
      Top             =   2280
      Width           =   15975
      Begin VSFlex8LCtl.VSFlexGrid grdENTLIT 
         Height          =   5055
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   15735
         _cx             =   27755
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
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   2895
      Begin VB.ListBox lstFiltro 
         Appearance      =   0  'Flat
         Height          =   1155
         ItemData        =   "frmCADRECMATLITP.frx":0000
         Left            =   120
         List            =   "frmCADRECMATLITP.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   2880
      TabIndex        =   9
      Top             =   0
      Width           =   13095
      Begin MSMask.MaskEdBox mskDTENTRADA 
         Height          =   285
         Left            =   4680
         TabIndex        =   25
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCODOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Text            =   "txtCODOP"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtCODPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   13
         Text            =   "txtCODPED"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtNOMEMPENT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   12
         Text            =   "txtNOMEMPENT"
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox txtCODEMPENT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "txtCODEMPENT"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "txtCodigo"
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cód.Pedido"
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
         Left            =   2640
         TabIndex        =   20
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cód.OP"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1005
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social Empresa"
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
         TabIndex        =   18
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cód.Empresa"
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
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dt.Entrada"
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
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame cmdFECHA 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   15975
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar <ESC>"
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
         Picture         =   "frmCADRECMATLITP.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdInclui 
         Caption         =   "&Incluir <F5>"
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
         Left            =   1440
         Picture         =   "frmCADRECMATLITP.frx":0536
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdAltera 
         Caption         =   "&Alterar <F6>"
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
         Picture         =   "frmCADRECMATLITP.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Altera Empresa "
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdExclui 
         Caption         =   "&Excluir"
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
         Left            =   3840
         Picture         =   "frmCADRECMATLITP.frx":0B6A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCanFiltro 
         Caption         =   "&Desfas"
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
         Left            =   14160
         Picture         =   "frmCADRECMATLITP.frx":0C6C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdOrden 
         Caption         =   "&Ordem"
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
         Left            =   15000
         Picture         =   "frmCADRECMATLITP.frx":119E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "&Pesquisa"
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
         Left            =   12840
         Picture         =   "frmCADRECMATLITP.frx":12A0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
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
         Left            =   4680
         Picture         =   "frmCADRECMATLITP.frx":13A2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADRECMATLITP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Public strUsuario   As String
Public lngCodUsuaro As Long

Dim objFuncoes          As Object
Dim objCADRECMATLITP    As Object
Dim objRel              As Object
Dim iCodigo             As Long

Const conCOL_Mov_Codigo                          As Integer = 0
Const conCOL_Mov_DtEntrada                       As Integer = 1
Const conCOL_Mov_CodClie                         As Integer = 2
Const conCOL_Mov_RazSoc                          As Integer = 3
Const conCOL_Mov_CodClieDest                     As Integer = 4
Const conCOL_Mov_RazSocDest                      As Integer = 5
Const conCOL_Mov_CodOP                           As Integer = 6
Const conCOL_Mov_CodPed                          As Integer = 7
Const conCOL_Mov_CodENV                          As Integer = 8
Const conCOL_Mov_FormatString                    As String = "=Código|Dt.Entrada|Empresa Orig.|Razão Social Empresa Origem|Empresa Dest.|Razão Social Empresa Destino|Cód.OP|Cód.Pedido|Cod.Env"
Const conCOL_Mov_Campos                          As String = "ITEN.SGI_CODIGO|CABE.SGI_DTENTRADA|CABE.SGI_CODCLIE|CLIE.SGI_RAZAOSOC|CABE.SGI_CODCLIEDEST|CLIEDEST.SGI_RAZAOSOCDEST|ITEN.SGI_CODOP|ITEN.SGI_CODPED"
Const conColumnsIn_Mov                           As Integer = 9

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
    Call ConfFiltro
    Call ConfGrid
End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  If (grdENTLIT.Rows - 1) <= 0 Then Exit Sub
  If (grdENTLIT.Row) <= 0 Then
    MsgBox "ATENÇÂO" & vbCrLf & _
           "Selecioone um registro !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
  End If
  
  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  objCADRECMATLITP.STATUS = "'ENV'"
  objCADRECMATLITP.CODIGOENV = CLng(grdENTLIT.Cell(flexcpText, grdENTLIT.Row, conCOL_Mov_CodENV))
  
  If objCADRECMATLITP.GRAVA("E") = False Then Exit Sub
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call AbilitaCampos
  Call ConfGrid

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdPesquisa_Click()
    If ValidaCampos = False Then Exit Sub
    Call Pesquisa
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
    Set objCADRECMATLITP = Nothing
    Set objRel = Nothing
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADRECMATLITP = CreateObject("CADRECMATLIT.clsCADRECMATLIT")
    Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    objCADRECMATLITP.FILIAL = FILIAL
    Call objFuncoes.LimpaCampos(Me)
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
        MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
        Exit Sub
    End If
    
    Call ConfFiltro
    Call AbilitaCampos
    Call ConfGrid
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    Me.Caption = "Recebimento de Material (Folhas Litografadas)"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub ConfFiltro()
   
   Dim i As Integer
   Dim arrCAMPOS()   As String
   
   With lstFiltro
   
        .Clear
        arrCAMPOS = Split(Trim(Replace(conCOL_Mov_FormatString, "=", "")), "|")
        
        For i = 0 To UBound(arrCAMPOS) - 1
            .AddItem arrCAMPOS(i)
            .ItemData(.NewIndex) = i
            .ListIndex = 0
        Next i
   
        For i = 0 To (.ListCount - 1)
            .Selected(i) = False
        Next i
    
    End With
   
End Sub

Private Sub ConfGrid()

    With grdENTLIT
    
       .Cols = conColumnsIn_Mov
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Mov_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_Mov_Codigo) = ""
       .ColDataType(conCOL_Mov_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_DtEntrada) = ""
       .ColDataType(conCOL_Mov_DtEntrada) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_Mov_CodClie) = ""
       .ColDataType(conCOL_Mov_CodClie) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_RazSoc) = ""
       .ColDataType(conCOL_Mov_RazSoc) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Mov_CodClieDest) = ""
       .ColDataType(conCOL_Mov_CodClieDest) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_RazSocDest) = ""
       .ColDataType(conCOL_Mov_RazSocDest) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Mov_CodOP) = ""
       .ColDataType(conCOL_Mov_CodOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_CodPed) = ""
       .ColDataType(conCOL_Mov_CodPed) = flexDTLong
       
       
       .Cell(flexcpData, 0, conCOL_Mov_CodENV) = ""
       .ColDataType(conCOL_Mov_CodENV) = flexDTLong
       
       .ColWidth(conCOL_Mov_Codigo) = 1000
       .ColWidth(conCOL_Mov_DtEntrada) = 1000
       .ColWidth(conCOL_Mov_CodClie) = 1100
       .ColWidth(conCOL_Mov_CodClieDest) = 1100
       .ColWidth(conCOL_Mov_RazSoc) = 5000
       .ColWidth(conCOL_Mov_RazSocDest) = 5000
       .ColWidth(conCOL_Mov_CodOP) = 1000
       .ColWidth(conCOL_Mov_CodPed) = 1000
       .ColWidth(conCOL_Mov_CodENV) = 1000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADRECMATLITP.AtivoDesativo("SGI_CADRECROTLIT_IT")
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdCanFiltro.Enabled = boolAtivoDesativo
    cmdOrden.Enabled = boolAtivoDesativo
    cmdPesquisa.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo
    lstFiltro.Enabled = boolAtivoDesativo

End Sub


Private Sub Operacao(strOperacao As String)
  
    With grdENTLIT
        If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
    End With
    
    frmCADRECMATLIT.cCaminho = cCaminho
    frmCADRECMATLIT.Linha = Linha
    frmCADRECMATLIT.iCodigo = iCodigo
    frmCADRECMATLIT.cTipOper = strOperacao
    frmCADRECMATLIT.FILIAL = FILIAL
    frmCADRECMATLIT.strAcesso = strAcesso
    frmCADRECMATLIT.strMODPAI = Me.Name
    frmCADRECMATLIT.strUsuario = strUsuario
    frmCADRECMATLIT.lngCodUsuario = lngCodUsuaro
    frmCADRECMATLIT.Show vbModal
    
    Call AbilitaCampos
    Call ConfGrid

End Sub

Private Sub grdENTLIT_Click()
    With grdENTLIT
        If (.Rows - 1) > 0 And .Row > 0 Then objCADRECMATLITP.CODIGO = CLng(grdENTLIT.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
    End With
End Sub

Private Sub grdENTLIT_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    With grdENTLIT
        If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("C")
    End With
End Sub

Private Sub grdENTLIT_RowColChange()
    With grdENTLIT
        If (.Rows - 1) > 0 And .Row > 0 Then objCADRECMATLITP.CODIGO = CLng(grdENTLIT.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
    End With
End Sub

Private Sub Ordem()

  Call ConfGrid
  
  Dim strCAMPO As String
  
  sSql = ""
  
  sSql = " Select " & vbCrLf
  sSql = sSql & "        CABE.SGI_CODIGO " & vbCrLf
  sSql = sSql & "       ,CABE.SGI_DTENTRADA" & vbCrLf
  sSql = sSql & "       ,CABE.SGI_CODCLIE" & vbCrLf
  sSql = sSql & "       ,CABE.SGI_CODCLIEDEST" & vbCrLf
  sSql = sSql & "       ,CABE.SGI_CODENV" & vbCrLf
  sSql = sSql & "       ,CLIE.SGI_RAZAOSOC" & vbCrLf
  sSql = sSql & "       ,CLIEDEST.SGI_RAZAOSOC As SGI_RAZAOSOCDEST" & vbCrLf
  sSql = sSql & "       ,CLIEDEST.SGI_CONFENTREST" & vbCrLf
  sSql = sSql & "       ,ITEN.SGI_CODOP" & vbCrLf
  sSql = sSql & "       ,ITEN.SGI_CODPED" & vbCrLf
 
  sSql = sSql & "   from " & vbCrLf
  sSql = sSql & "        SGI_CADRECROTLIT_IT ITEN" & vbCrLf
  sSql = sSql & "       ,SGI_CADRECROTLIT    CABE" & vbCrLf
  sSql = sSql & "       ,SGI_CADCLIENTE      CLIE" & vbCrLf
  sSql = sSql & "       ,SGI_CADCLIENTE      CLIEDEST" & vbCrLf
  
  sSql = sSql & " Where " & vbCrLf
  sSql = sSql & "        ITEN.SGI_FILIAL        = " & FILIAL & vbCrLf
  sSql = sSql & "   And  CABE.SGI_FILIAL        = ITEN.SGI_FILIAL" & vbCrLf
  sSql = sSql & "   And  CABE.SGI_CODIGO        = ITEN.SGI_CODIGO" & vbCrLf
  sSql = sSql & "   And  CLIE.SGI_FILIAL        = CABE.SGI_FILIAL" & vbCrLf
  sSql = sSql & "   And  CLIE.SGI_CODIGO        = CABE.SGI_CODCLIE" & vbCrLf
  sSql = sSql & "   And  CLIEDEST.SGI_FILIAL    = CABE.SGI_FILIAL" & vbCrLf
  sSql = sSql & "   And  CLIEDEST.SGI_CODIGO    = CABE.SGI_CODCLIEDEST" & vbCrLf
  
  sSql = sSql & MontaOrderBy(conCOL_Mov_Campos, lstFiltro)
  
  BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
  If Not BREC.EOF() Then
    With grdENTLIT
        Do While Not BREC.EOF()
            
            strCAMPO = BREC!SGI_CODIGO & vbTab & _
                        Format(BREC!SGI_DTENTRADA, "DD/MM/YYYY") & vbTab & _
                        BREC!SGI_CODCLIE & vbTab & _
                        BREC!SGI_RAZAOSOC & vbTab & _
                        BREC!SGI_CODCLIEDEST & vbTab & _
                        BREC!SGI_RAZAOSOCDEST & vbTab & _
                        BREC!SGI_CODOP & vbTab & _
                        BREC!SGI_CODPED & vbTab & _
                        BREC!SGI_CODENV
            
            .AddItem strCAMPO
            
            BREC.MoveNext
        Loop
    End With
  End If
  
  BREC.Close

End Sub


Private Function MontaOrderBy(strCAMPOS As String, lstGENERICO As Variant) As String

  MontaOrderBy = ""

  Dim i                   As Integer
  Dim intCODIGO           As Integer
  Dim arrCAMPOS_BD()      As String
  Dim intINDICE           As Integer
  
  intINDICE = 0
  arrCAMPOS_BD = Split(strCAMPOS, "|")
  
  If lstGENERICO.SelCount > 0 Then
    MontaOrderBy = MontaOrderBy & "Order By " & vbCrLf
     
    intCODIGO = 1
    For i = 0 To (lstGENERICO.ListCount - 1)
        If lstGENERICO.Selected(i) = True Then
           
           If lstGENERICO.ItemData(i) = intCODIGO Then MontaOrderBy = MontaOrderBy & "         " & arrCAMPOS_BD(intCODIGO)
           
           intINDICE = intINDICE + 1
           If intINDICE < lstGENERICO.SelCount Then MontaOrderBy = MontaOrderBy & "," & vbCrLf
        End If
        intCODIGO = intCODIGO + 1
    Next i
  
  End If

End Function

Private Sub mskDTENTRADA_GotFocus()
    objFuncoes.SelecionaCampos mskDTENTRADA.Name, Me
End Sub

Private Sub txtCODEMPENT_GotFocus()
    objFuncoes.SelecionaCampos txtCODEMPENT.Name, Me
End Sub

Private Sub txtCODEMPENT_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODEMPENT.Text
End Sub

Private Sub txtCodigo_GotFocus()
    objFuncoes.SelecionaCampos txtCodigo.Name, Me
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCodigo.Text
End Sub

Private Sub txtCODOP_GotFocus()
    objFuncoes.SelecionaCampos txtCODOP.Name, Me
End Sub

Private Sub txtCODOP_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODOP.Text
End Sub

Private Sub txtCODPED_GotFocus()
    objFuncoes.SelecionaCampos txtCODPED.Name, Me
End Sub

Private Sub txtCODPED_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODPED.Text
End Sub

Private Sub txtNOMEMPENT_GotFocus()
    objFuncoes.SelecionaCampos txtNOMEMPENT.Name, Me
End Sub

Private Function ValidaCampos() As Boolean

        ValidaCampos = False
     
        If Len(Trim(Replace(Replace(mskDTENTRADA.Text, "/", ""), "_", ""))) = 0 Then
            ValidaCampos = True
            Exit Function
        End If
     
        If Not IsDate(mskDTENTRADA.Text) Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Campo Data inválido !!!", vbOKOnly + vbExclamation, "Acviso"
                   mskDTENTRADA.SetFocus
                   Exit Function
        End If
        
        ValidaCampos = True
     
End Function



Private Function Pesquisa() As Boolean

    Pesquisa = False
    
    Call ConfGrid
    
    Dim strMOV As String
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "        CABE.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      , CABE.SGI_DTENTRADA" & vbCrLf
    sSql = sSql & "      , CABE.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "      , CABE.SGI_CODCLIEDEST" & vbCrLf
    sSql = sSql & "      , CABE.SGI_CODENV" & vbCrLf
    sSql = sSql & "      , CLIE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      , CLIEDEST.SGI_RAZAOSOC As SGI_RAZAOSOCDEST" & vbCrLf
    sSql = sSql & "      , CLIEDEST.SGI_CONFENTREST" & vbCrLf
    sSql = sSql & "      , ITEN.SGI_CODOP" & vbCrLf
    sSql = sSql & "      , ITEN.SGI_CODPED" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADRECROTLIT_IT ITEN" & vbCrLf
    sSql = sSql & "      ,SGI_CADRECROTLIT    CABE" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE      CLIEDEST" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       ITEN.SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And ITEN.SGI_FILIAL        = CABE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ITEN.SGI_CODIGO        = CABE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And CABE.SGI_FILIAL        = CLIE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CABE.SGI_CODCLIE       = CLIE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And CLIEDEST.SGI_FILIAL    = CABE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIEDEST.SGI_CODIGO    = CABE.SGI_CODCLIEDEST" & vbCrLf

    If Len(Trim(txtCodigo.Text)) > 0 Then
        sSql = sSql & "   And ITEN.SGI_CODIGO = " & Trim(txtCodigo.Text) & vbCrLf
    End If
    
    If Len(Trim(Replace(Replace(mskDTENTRADA.Text, "/", ""), "_", ""))) > 0 Then
        sSql = sSql & "   And CABE.SGI_DTENTRADA = '" & Format(CDate(mskDTENTRADA.Text), "MM/DD/YYYY") & "'" & vbCrLf
    End If
    
    If Len(Trim(txtCODEMPENT.Text)) > 0 Then
        sSql = sSql & "   And CABE.SGI_CODCLIE = " & Trim(txtCODEMPENT.Text) & vbCrLf
    End If
    
    If Len(Trim(txtNOMEMPENT.Text)) > 0 Then
        sSql = sSql & "   And CLIE.SGI_RAZAOSOC Like '" & Trim(txtNOMEMPENT.Text) & "%'" & vbCrLf
    End If
    
    If Len(Trim(txtCODOP.Text)) > 0 Then
       sSql = sSql & "    And ITEN.SGI_CODOP = " & Trim(txtCODOP.Text) & vbCrLf
    End If

    If Len(Trim(txtCODPED.Text)) > 0 Then
        sSql = sSql & "   And ITEN.SGI_CODPED = " & Trim(txtCODPED.Text) & vbCrLf
    End If

    sSql = sSql & MontaOrderBy(conCOL_Mov_Campos, lstFiltro)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
       Pesquisa = True
    
        With grdENTLIT
            Do While Not BREC.EOF()
                
                strMOV = BREC!SGI_CODIGO & vbTab & _
                         Format(BREC!SGI_DTENTRADA, "DD/MM/YYYY") & vbTab & _
                         BREC!SGI_CODCLIE & vbTab & _
                         BREC!SGI_RAZAOSOC & vbTab & _
                         BREC!SGI_CODCLIEDEST & vbTab & _
                         BREC!SGI_RAZAOSOCDEST & vbTab & _
                         BREC!SGI_CODOP & vbTab & _
                         BREC!SGI_CODPED & vbTab & _
                         BREC!SGI_CODENV
                
                .AddItem strMOV
                
                BREC.MoveNext
            Loop
        End With
    
    End If
    BREC.Close
    
    
    If Pesquisa = False Then
        MsgBox "Não foi encontrado nenhum registro !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If

End Function


