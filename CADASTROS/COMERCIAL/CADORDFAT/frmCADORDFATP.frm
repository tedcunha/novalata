VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADORDFATP 
   Caption         =   "Cadastro de Ordem de Faturamento"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13065
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stOrdFat 
      Height          =   6495
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
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
      TabPicture(0)   =   "frmCADORDFATP.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdORDFAT"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Confirmado NF"
      TabPicture(1)   =   "frmCADORDFATP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdORDFATCONF"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Geral"
      TabPicture(2)   =   "frmCADORDFATP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdGERAL"
      Tab(2).ControlCount=   1
      Begin VSFlex8LCtl.VSFlexGrid grdGERAL 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   12735
         _cx             =   22463
         _cy             =   10610
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
      Begin VSFlex8LCtl.VSFlexGrid grdORDFATCONF 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   12735
         _cx             =   22463
         _cy             =   10610
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
      Begin VSFlex8LCtl.VSFlexGrid grdORDFAT 
         Height          =   6015
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   12735
         _cx             =   22463
         _cy             =   10610
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
         Width           =   9015
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
      Width           =   12975
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
         Picture         =   "frmCADORDFATP.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprime Registro"
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
         Picture         =   "frmCADORDFATP.frx":0156
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
         Picture         =   "frmCADORDFATP.frx":0688
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
         Picture         =   "frmCADORDFATP.frx":0BBA
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
         Picture         =   "frmCADORDFATP.frx":0CBC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Registro"
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
         Left            =   11400
         Picture         =   "frmCADORDFATP.frx":0DBE
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
         Left            =   12120
         Picture         =   "frmCADORDFATP.frx":12F0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ordena os Registros"
         Top             =   120
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   10800
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmCADORDFATP"
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
Public intFILIALPED     As Integer

Dim lngCodVendedor      As Long
Dim objFuncoes          As Object
Dim objCADORDFAT        As Object
Dim objRel              As Object
Dim iCodigo             As Long
Dim boolComAcao         As Boolean
Dim strNOMTABELA1       As String
Dim strNOMTABELA2       As String
Dim strNOMTABELA3       As String
Dim strNOMTABELA4       As String
Dim strNOMMODULO        As String
Dim strLOGMODULO        As String

Const conCOL_OrdFat_Selecionado                     As Integer = 0
Const conCOL_OrdFat_Codigo                          As Integer = 1
Const conCOL_OrdFat_DataOrdem                       As Integer = 2
Const conCOL_OrdFat_Cliente                         As Integer = 3
Const conCOL_OrdFat_CodEmp                          As Integer = 4
Const conCOL_OrdFat_NomeEmp                         As Integer = 5
Const conCOL_OrdFat_FormatString                    As String = "=  |Código|Data Ordem|Cliente|CodEmp|Empresa"
Const conColumnsIn_OrdFat                           As Integer = 6

Const conCOL_OrdFatConf_Selecionado                 As Integer = 0
Const conCOL_OrdFatConf_Codigo                      As Integer = 1
Const conCOL_OrdFatConf_DataOrdem                   As Integer = 2
Const conCOL_OrdFatConf_Cliente                     As Integer = 3
Const conCOL_OrdFatConf_CodEmp                      As Integer = 4
Const conCOL_OrdFatConf_NomeEmp                     As Integer = 5
Const conCOL_OrdFatConf_FormatString                As String = "=  |Código|Data Ordem|Cliente|CodEmp|Empresa"
Const conColumnsIn_OrdFatConf                       As Integer = 6

Const conCOL_OrdFatGeral_Selecionado                 As Integer = 0
Const conCOL_OrdFatGeral_Codigo                      As Integer = 1
Const conCOL_OrdFatGeral_DataOrdem                   As Integer = 2
Const conCOL_OrdFatGeral_Cliente                     As Integer = 3
Const conCOL_OrdFatGeral_CodEmp                      As Integer = 4
Const conCOL_OrdFatGeral_NomeEmp                     As Integer = 5
Const conCOL_OrdFatGeral_FormatString                As String = "=  |Código|Data Ordem|Cliente|CodEmp|Empresa"
Const conColumnsIn_OrdFatGeral                       As Integer = 6

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    If Not VerifStatusOrdem Then
        MsgBox "Ordem de Faturamento já Confirmada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    ElseIf Pega_QtdeOrdens > 1 Then
        MsgBox "Existem mais ordens de faturamento atrelados !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
   txtCampos.Text = ""
   Call AbilitaCampos
   Call ConfGridOrdFat
   Call ConfGridOrdFatConf
   Call ConfGridOrdFatGeral
End Sub

Private Sub cmdExclui_Click()

    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
    If stOrdFat.Tab = 1 Then
       MsgBox "Ordem já Confirmada Não pode ser Excluida !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    Dim iResp     As Integer
    Dim lngCodLog As Long
    
    iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
    If iResp <> 6 Then Exit Sub
  
    If PegaDadosParaExclusao = True Then
        MsgBox "Existe ordem de faturamento que já foi confirmada não pode ser excluida !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If objCADORDFAT.GRAVA("E", intFILIALPED) = False Then Exit Sub
    If objFuncoes.Atualiza("E", Str(objCADORDFAT.CODORD), FILIAL, "frmCADORDFAT", Linha, Str(intFILIALPED)) = False Then Exit Sub
    
    
    lngCodLog = objFuncoes.Gera_Codigo(strLOGMODULO, FILIAL, Linha)
    Call objFuncoes.GravaLogModulo(FILIAL, lngCodLog, strNOMMODULO, "E", lngCodUsuaro, Str(objCADORDFAT.CODORD), Linha)
    
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
    If stOrdFat.Tab = 0 Then Call ConfGridOrdFat
    If stOrdFat.Tab = 1 Then Call ConfGridOrdFatConf
    If stOrdFat.Tab = 2 Then Call ConfGridOrdFatGeral
    
    Call AbilitaCampos

End Sub

Private Sub cmdImpressao_Click()
    Call ImpOrdNovo
    Exit Sub
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
    Set objCADORDFAT = CreateObject("CADORDFAT.clsCADORDFAT")
    Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    objCADORDFAT.FILIAL = FILIAL
    objFuncoes.LimpaCampos frmCADORDFATP
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call AbilitaCampos
    Call ConfGridOrdFat
    Call ConfGridOrdFatConf
    Call ConfGridOrdFatGeral
    
    stOrdFat.Tab = 0
    
    Call ConfFiltro
    
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    boolComAcao = False

    If intFILIALPED = 0 Then
       Me.Caption = Me.Caption & " / NOVALATA"
       strNOMTABELA1 = "SGI_CADORDFATH"
       strNOMTABELA2 = "SGI_CADPEDVENDH"
       strNOMTABELA3 = "SGI_CADORDFATI"
       strNOMTABELA4 = "SGI_CADORDCONFH"
       strNOMMODULO = "frmCADORDFAT"
       strLOGMODULO = "SGI_LOGMODULO"
    ElseIf intFILIALPED = 1 Then
       Me.Caption = Me.Caption & " / STEEL ROLL"
       strNOMTABELA1 = "SGI_CADORDFATH_STEEL"
       strNOMTABELA2 = "SGI_CADPEDVENDH_STEEL"
       strNOMTABELA3 = "SGI_CADORDFATI_STEEL"
       strNOMTABELA4 = "SGI_CADORDCONFH_STEEL"
       strNOMMODULO = "frmCADORDFAT_STEEL"
       strLOGMODULO = "SGI_LOGMODULO_STEEL"
    End If

    Me.Caption = Me.Caption & " - Versão : " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub AbilitaCampos()

    If boolComAcao = True Then Exit Sub
    
    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADORDFAT.AtivoDesativo(intFILIALPED)
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdImpressao.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo

End Sub

Private Sub ConfGridOrdFat()

    With grdORDFAT
    
       .Cols = conColumnsIn_OrdFat
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdFat_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_OrdFat_Selecionado) = ""
       .ColDataType(conCOL_OrdFat_Selecionado) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_OrdFat_Codigo) = ""
       .ColDataType(conCOL_OrdFat_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFat_DataOrdem) = ""
       .ColDataType(conCOL_OrdFat_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdFat_Cliente) = ""
       .ColDataType(conCOL_OrdFat_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFat_CodEmp) = ""
       .ColDataType(conCOL_OrdFat_CodEmp) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFat_NomeEmp) = ""
       .ColDataType(conCOL_OrdFat_NomeEmp) = flexDTString
       
       .ColWidth(conCOL_OrdFat_Selecionado) = 300
       .ColWidth(conCOL_OrdFat_Codigo) = 1500
       .ColWidth(conCOL_OrdFat_DataOrdem) = 1000
       .ColWidth(conCOL_OrdFat_Cliente) = 5000
       .ColWidth(conCOL_OrdFat_CodEmp) = 0
       .ColWidth(conCOL_OrdFat_NomeEmp) = 0
       
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
    sSql = sSql & "       FAT.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA1 & " FAT " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       FAT.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And FAT.SGI_STATUS = 0" & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO = FAT.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI "
    
    sSql = sSql & "Order by FAT.SGI_CODORD "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       grdORDFAT.AddItem BREC!SGI_CODORD & vbTab & _
                         Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                         Trim(BREC!SGI_RAZAOSOC)
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub


Private Sub Operacao(strOperacao As String)
  
    If stOrdFat.Tab = 0 Then
        If (grdORDFAT.Rows - 1) > 0 And grdORDFAT.RowSel > 0 Then iCodigo = CLng(grdORDFAT.Cell(flexcpText, grdORDFAT.Row, conCOL_OrdFat_Codigo))
    ElseIf stOrdFat.Tab = 1 Then
        If (grdORDFATCONF.Rows - 1) > 0 And grdORDFATCONF.RowSel > 0 Then iCodigo = CLng(grdORDFATCONF.Cell(flexcpText, grdORDFATCONF.Row, conCOL_OrdFatConf_Codigo))
    ElseIf stOrdFat.Tab = 2 Then
        If (grdGERAL.Rows - 1) > 0 And grdGERAL.RowSel > 0 Then iCodigo = CLng(grdGERAL.Cell(flexcpText, grdGERAL.Row, conCOL_OrdFatGeral_Codigo))
    End If
    
    boolComAcao = True
    
    frmCADORDFAT.cCaminho = cCaminho
    frmCADORDFAT.Linha = Linha
    frmCADORDFAT.iCodigo = iCodigo
    frmCADORDFAT.cTipOper = strOperacao
    frmCADORDFAT.FILIAL = FILIAL
    frmCADORDFAT.strAcesso = strAcesso
    frmCADORDFAT.strMODPAI = Me.Name
    frmCADORDFAT.strUSUARIO = strUSUARIO
    frmCADORDFAT.lngCodVendedor = lngCodVendedor
    frmCADORDFAT.lngCodUsuario = lngCodUsuaro
    frmCADORDFAT.intFILIALPED = intFILIALPED
    frmCADORDFAT.strNOMMODULO = strNOMMODULO
    frmCADORDFAT.Show vbModal
    
    boolComAcao = False
    
    Call AbilitaCampos
    Call ConfGridOrdFat
    Call ConfGridOrdFatConf
    Call ConfGridOrdFatGeral

End Sub


Private Sub Atualiza_Grid()
    
     Dim I              As Long
     Dim bolAchou       As Boolean
     Dim lngCODIGO      As Long
     Dim strACAO        As String
     Dim lngCOL         As Long
     Dim grdGENERICA    As VSFlexGrid
     Dim strCampos      As String
     
     If BRECATU.State = 1 Then BRECATU.Close
     
     If stOrdFat.Tab = 0 Then
        lngCOL = conCOL_OrdFat_Codigo
        Set grdGENERICA = grdORDFAT
     ElseIf stOrdFat.Tab = 1 Then
        lngCOL = conCOL_OrdFatConf_Codigo
        Set grdGENERICA = grdORDFATCONF
     ElseIf stOrdFat.Tab = 2 Then
        lngCOL = conCOL_OrdFatGeral_Codigo
        Set grdGENERICA = grdGERAL
     End If
     
     bolAchou = False
      
     With grdGENERICA
     
        sSql = "Select" & vbCrLf
        sSql = sSql & "      * " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_MODULO    = 'frmCADORDFAT'" & vbCrLf
        sSql = sSql & "   And SGI_FILIALPED = " & intFILIALPED & vbCrLf
        
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
        sSql = sSql & "       FAT.* " & vbCrLf
        sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADORDFATH  FAT " & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH PED " & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        
        sSql = sSql & "       FAT.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And FAT.SGI_CODORD = " & lngCODIGO
        
        If stOrdFat.Tab = 0 Then
            sSql = sSql & "   And FAT.SGI_STATUS = 0" & vbCrLf
        ElseIf stOrdFat.Tab = 1 Then
            sSql = sSql & "   And FAT.SGI_STATUS = 1" & vbCrLf
        End If
        
        sSql = sSql & "   And PED.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO = FAT.SGI_CODPED " & vbCrLf
        sSql = sSql & "   And CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI "
            
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF Then
            If bolAchou = False And Trim(strACAO) = "I" Then
               
                strCampos = 0 & vbTab & _
                            BREC2!SGI_CODORD & vbTab & _
                            Format(BREC2!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                            BREC2!SGI_RAZAOSOC & vbTab & _
                            BREC2!SGI_FILIALPED & vbTab & _
                            IIf(BREC2!SGI_FILIALPED = 0, "NOVALATA", "STEEL ROW")
                      
                .AddItem strCampos
            
            ElseIf bolAchou = True And Trim(strACAO) = "A" Then
                    If stOrdFat.Tab = 0 Then
                        .Cell(flexcpText, I, conCOL_OrdFat_Codigo) = BREC2!SGI_CODORD
                        .Cell(flexcpText, I, conCOL_OrdFat_DataOrdem) = Format(BREC2!SGI_DATAORDEM, "DD/MM/YYYY")
                        .Cell(flexcpText, I, conCOL_OrdFat_Cliente) = BREC2!SGI_RAZAOSOC
                        .Cell(flexcpText, I, conCOL_OrdFat_CodEmp) = BREC2!SGI_FILIALPED
                        .Cell(flexcpText, I, conCOL_OrdFat_NomeEmp) = IIf(BREC2!SGI_FILIALPED = 0, "NOVALATA", "STEEL ROW")
                    ElseIf stOrdFat.Tab = 1 Then
                        .Cell(flexcpText, I, conCOL_OrdFatConf_Codigo) = BREC2!SGI_CODORD
                        .Cell(flexcpText, I, conCOL_OrdFatConf_DataOrdem) = Format(BREC2!SGI_DATAORDEM, "DD/MM/YYYY")
                        .Cell(flexcpText, I, conCOL_OrdFatConf_Cliente) = BREC2!SGI_RAZAOSOC
                        .Cell(flexcpText, I, conCOL_OrdFatConf_CodEmp) = BREC2!SGI_FILIALPED
                        .Cell(flexcpText, I, conCOL_OrdFatConf_NomeEmp) = IIf(BREC2!SGI_FILIALPED = 0, "NOVALATA", "STEEL ROW")
                    ElseIf stOrdFat.Tab = 2 Then
                        .Cell(flexcpText, I, conCOL_OrdFatGeral_Codigo) = BREC2!SGI_CODORD
                        .Cell(flexcpText, I, conCOL_OrdFatGeral_DataOrdem) = Format(BREC2!SGI_DATAORDEM, "DD/MM/YYYY")
                        .Cell(flexcpText, I, conCOL_OrdFatGeral_Cliente) = BREC2!SGI_RAZAOSOC
                        .Cell(flexcpText, I, conCOL_OrdFatGeral_CodEmp) = BREC2!SGI_FILIALPED
                        .Cell(flexcpText, I, conCOL_OrdFatGeral_NomeEmp) = IIf(BREC2!SGI_FILIALPED = 0, "NOVALATA", "STEEL ROW")
                    
                    End If
            End If
        End If
        BREC2.Close

     End With
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub

Private Sub grdGERAL_Click()
   If (grdGERAL.Rows - 1) > 0 And grdGERAL.RowSel > 0 Then objCADORDFAT.CODORD = CLng(grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_OrdFatGeral_Codigo))
End Sub

Private Sub grdGERAL_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdGERAL.Rows - 1) > 0 And grdGERAL.RowSel > 0 Then Operacao "C"
End Sub

Private Sub grdGERAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdGERAL.Rows - 1) > 0 And grdGERAL.RowSel > 0 Then Operacao "C"
    End If
End Sub

Private Sub grdGERAL_RowColChange()
   If (grdGERAL.Rows - 1) > 0 And grdGERAL.RowSel > 0 Then objCADORDFAT.CODORD = CLng(grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_OrdFatGeral_Codigo))
End Sub

Private Sub grdORDFAT_Click()
   If (grdORDFAT.Rows - 1) > 0 And grdORDFAT.RowSel > 0 Then objCADORDFAT.CODORD = CLng(grdORDFAT.Cell(flexcpText, grdORDFAT.RowSel, conCOL_OrdFat_Codigo))
End Sub

Private Sub grdORDFAT_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdORDFAT.Rows - 1) > 0 And grdORDFAT.RowSel > 0 Then Operacao "C"
End Sub

Private Sub grdORDFAT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdORDFAT.Rows - 1) > 0 And grdORDFAT.RowSel > 0 Then Operacao "C"
    End If
End Sub


Private Sub grdORDFAT_RowColChange()
   If (grdORDFAT.Rows - 1) > 0 And grdORDFAT.RowSel > 0 Then objCADORDFAT.CODORD = CLng(grdORDFAT.Cell(flexcpText, grdORDFAT.RowSel, conCOL_OrdFat_Codigo))
End Sub

Private Sub grdORDFATCONF_Click()
   If (grdORDFATCONF.Rows - 1) > 0 And grdORDFATCONF.RowSel > 0 Then objCADORDFAT.CODORD = CLng(grdORDFATCONF.Cell(flexcpText, grdORDFATCONF.RowSel, conCOL_OrdFatConf_Codigo))
End Sub

Private Sub grdORDFATCONF_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdORDFATCONF.Rows - 1) > 0 And grdORDFATCONF.RowSel > 0 Then Operacao "C"
End Sub

Private Sub grdORDFATCONF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdORDFATCONF.Rows - 1) > 0 And grdORDFATCONF.RowSel > 0 Then Operacao "C"
    End If
End Sub

Private Sub grdORDFATCONF_RowColChange()
   If (grdORDFATCONF.Rows - 1) > 0 And grdORDFATCONF.RowSel > 0 Then objCADORDFAT.CODORD = CLng(grdORDFATCONF.Cell(flexcpText, grdORDFATCONF.RowSel, conCOL_OrdFatConf_Codigo))
End Sub

Private Sub stOrdFat_Click(PreviousTab As Integer)
    Call ConfFiltro
End Sub

Private Sub Timer1_Timer()
    Call Atualiza_Grid
    Call AbilitaCampos
End Sub


Private Sub Ordem()

    Dim strCampos As String
    
    Call ConfGridOrdFat
    Call ConfGridOrdFatConf
    Call ConfGridOrdFatGeral
    
    txtCampos.Text = ""
    
    sSql = ""
  
    sSql = "Select " & vbCrLf
    sSql = sSql & "       FAT.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
    sSql = sSql & "      ,PED.*" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA1 & " FAT " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       FAT.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If stOrdFat.Tab = 0 Then
        sSql = sSql & "   And FAT.SGI_STATUS = 0" & vbCrLf
    ElseIf stOrdFat.Tab = 1 Then
        sSql = sSql & "   And FAT.SGI_STATUS = 1" & vbCrLf
    End If
    
    sSql = sSql & "   And PED.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO = FAT.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI "
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by FAT.SGI_CODORD "
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by FAT.SGI_DATAORDEM "
    If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order by CLI.SGI_RAZAOSOC "
    
    BREC.Open sSql, adoBanco_Dados
    
    strCampos = ""
    Do While Not BREC.EOF
        
        strCampos = 0 & vbTab & _
                    BREC!SGI_CODORD & vbTab & _
                    Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                    BREC!SGI_RAZAOSOC & vbTab & _
                    "" & vbTab & _
                    ""
                        
    
        If stOrdFat.Tab = 0 Then grdORDFAT.AddItem strCampos
        If stOrdFat.Tab = 1 Then grdORDFATCONF.AddItem strCampos
        If stOrdFat.Tab = 2 Then grdGERAL.AddItem strCampos
       
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADORDFATP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub


Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim strCampos As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    If stOrdFat.Tab = 0 Then Call ConfGridOrdFat
    If stOrdFat.Tab = 1 Then Call ConfGridOrdFatConf
    If stOrdFat.Tab = 2 Then Call ConfGridOrdFatGeral
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       FAT.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA1 & " FAT " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       FAT.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If stOrdFat.Tab = 0 Then
        sSql = sSql & "   and FAT.SGI_STATUS = 0 " & vbCrLf
    ElseIf stOrdFat.Tab = 1 Then
        sSql = sSql & "   and FAT.SGI_STATUS = 1 " & vbCrLf
    End If
    
    sSql = sSql & "   And PED.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO = FAT.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And FAT.SGI_CODORD = " & Trim(txtCampos.Text) & vbCrLf
        sSql = sSql & "Order by FAT.SGI_CODORD " & vbCrLf
    ElseIf cboFiltro.ListIndex = 1 Then
        If IsDate(txtCampos.Text) = False Then
           MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
        sSql = sSql & "     And FAT.SGI_DATAORDEM = '" & Format(txtCampos.Text, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "Order by FAT.SGI_DATAORDEM " & vbCrLf
    ElseIf cboFiltro.ListIndex = 2 Then
       sSql = sSql & "     And CLI.SGI_RAZAOSOC LIKE '" & Trim(txtCampos.Text) & "%'" & vbCrLf
       sSql = sSql & "Order by CLI.SGI_RAZAOSOC " & vbCrLf
    ''ElseIf cboFiltro.ListIndex = 4 Then
    ''    sSql = sSql & "     And FAT.SGI_CODORD = " & Trim(txtCampos.Text) & vbCrLf
    ''    sSql = sSql & "Order by FAT.SGI_CODORD " & vbCrLf
    End If
        
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then
        
        Do While Not BREC2.EOF()
            
            strCampos = 0 & vbTab & _
                        BREC2!SGI_CODORD & vbTab & _
                        Format(BREC2!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                        BREC2!SGI_RAZAOSOC & vbTab & _
                        "" & vbTab & _
                        ""
                            
        
            If stOrdFat.Tab = 0 Then grdORDFAT.AddItem strCampos
            If stOrdFat.Tab = 1 Then grdORDFATCONF.AddItem strCampos
            If stOrdFat.Tab = 2 Then grdGERAL.AddItem strCampos
            
            BREC2.MoveNext
        Loop
    Else
        MsgBox "Este Registro não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC2.Close

End Sub

Private Function PegaDadosParaExclusao() As Boolean
    
    PegaDadosParaExclusao = False
    
    Dim I As Integer
    
    objCADORDFAT.QTDETOTALFAT = 0
    objCADORDFAT.QTDEATENDPED = 0
    objCADORDFAT.ORDENS = Empty
    
    Dim qtdREGS As Long
    Dim arrITENS() As Long
    
    '' Pegando dados da Ordem de Faturamento
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA1 & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODORD = " & objCADORDFAT.CODORD
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then
        objCADORDFAT.CODPED = BREC8!SGI_CODPED
    End If
    BREC8.Close
    
    
    '' Pegando as Outras Confirmações Atreladas
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA1 & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED = " & objCADORDFAT.CODPED
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then
        qtdREGS = 0
        Do While Not BREC8.EOF()
        
           qtdREGS = (qtdREGS + 1)
           ReDim Preserve arrITENS(1 To qtdREGS) As Long
           arrITENS(qtdREGS) = BREC8!SGI_CODORD
           
           BREC8.MoveNext
        Loop
        objCADORDFAT.ORDENS = arrITENS
    End If
    BREC8.Close
    
    
    '' Verificando se Existe Confirmação
    If IsArray(objCADORDFAT.ORDENS) Then
        For I = 1 To UBound(arrITENS)
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       " & strNOMTABELA4 & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODORD = " & arrITENS(I)
            
            BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC12.EOF() Then PegaDadosParaExclusao = True
            BREC12.Close
            
        Next I
    End If

End Function

Private Function Pega_QtdeOrdens() As Long

    Pega_QtdeOrdens = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA1 & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODORD = " & objCADORDFAT.CODORD
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       Count(SGI_CODPED) As Qtde " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "        " & strNOMTABELA1 & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODPED = " & BREC6!SGI_CODPED
    
        BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC7.EOF() Then Pega_QtdeOrdens = BREC7!Qtde
        BREC7.Close
        
    End If
    BREC6.Close

End Function

Private Sub ConfGridOrdFatConf()

    With grdORDFATCONF
    
       .Cols = conColumnsIn_OrdFatConf
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdFatConf_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       
       .Cell(flexcpData, 0, conCOL_OrdFatConf_Selecionado) = ""
       .ColDataType(conCOL_OrdFatConf_Selecionado) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_OrdFatConf_Codigo) = ""
       .ColDataType(conCOL_OrdFatConf_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFatConf_DataOrdem) = ""
       .ColDataType(conCOL_OrdFatConf_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdFatConf_Cliente) = ""
       .ColDataType(conCOL_OrdFatConf_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFatConf_CodEmp) = ""
       .ColDataType(conCOL_OrdFatConf_CodEmp) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFatConf_NomeEmp) = ""
       .ColDataType(conCOL_OrdFatConf_NomeEmp) = flexDTString
       
       .ColWidth(conCOL_OrdFatConf_Selecionado) = 300
       .ColWidth(conCOL_OrdFatConf_Codigo) = 1500
       .ColWidth(conCOL_OrdFatConf_DataOrdem) = 1000
       .ColWidth(conCOL_OrdFatConf_Cliente) = 5000
       .ColWidth(conCOL_OrdFatConf_CodEmp) = 0
       .ColWidth(conCOL_OrdFatConf_NomeEmp) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PreencheGridConf()

    With grdORDFATCONF
        
        sSql = "Select " & vbCrLf
        
        sSql = sSql & "       FAT.* " & vbCrLf
        sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
        sSql = sSql & "      ,CONF.SGI_CODFATURA " & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       " & strNOMTABELA1 & " FAT " & vbCrLf
        sSql = sSql & "      ," & strNOMTABELA2 & " PED " & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
        sSql = sSql & "      ,SGI_CADORDCONFH CONF " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        
        sSql = sSql & "       FAT.SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And FAT.SGI_STATUS  = 1" & vbCrLf
        sSql = sSql & "   And PED.SGI_FILIAL  = FAT.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO  = FAT.SGI_CODPED " & vbCrLf
        sSql = sSql & "   And CLI.SGI_FILIAL  = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLI.SGI_CODIGO  = PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "   And CONF.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CONF.SGI_CODORD = FAT.SGI_CODORD " & vbCrLf
        
        sSql = sSql & "Order by FAT.SGI_CODORD "
        
        BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC7.EOF
                .AddItem BREC7!SGI_CODORD & vbTab & _
                         Format(BREC7!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                         Trim(BREC7!SGI_RAZAOSOC) & vbTab & _
                         BREC7!SGI_CODPED & vbTab & _
                         BREC7!SGI_CODFATURA
                                 
           BREC7.MoveNext
        Loop
        BREC7.Close
        
    End With
    
End Sub



Private Function VerifStatusOrdem() As Boolean
    
    VerifStatusOrdem = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA1 & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODORD = " & objCADORDFAT.CODORD & vbCrLf
    sSql = sSql & "   And SGI_STATUS = 0 "
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then VerifStatusOrdem = True
    BREC10.Close
    
End Function

Private Sub ImpOrd()

On Error GoTo Err_Imp

    Dim strNOMARQ       As String
    Dim strFILIALORD    As String
        
    strFILIALORD = ""
    If intFILIALPED = 1 Then strFILIALORD = "_STEEL"
        
    sSql = ""
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "       SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORD" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_CODPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_QTDFAT" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_VLUNIT" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_VLFATURADO" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_VLDOPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORDFAB" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_CODFORN" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_VLTOTAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_FECHTPFU" & vbCrLf
    
    sSql = sSql & "      ," & strNOMTABELA1 & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA1 & ".SGI_CODORD" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA1 & ".SGI_CODPED" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA1 & ".SGI_DATAORDEM" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA1 & ".SGI_OBS" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODCLIE" & vbCrLf
    
    sSql = sSql & "      ," & strNOMTABELA2 & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & ".SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & ".SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & ".SGI_CODCONDPGT" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & ".SGI_CODTRANSP" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & ".SGI_EMAIL" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RGCGC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ESTNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ENDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_BAINROM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CEPNORM" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADTRANSP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADTRANSP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADTRANSP.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "      ,SGI_ORDEMPROD" & strFILIALORD & ".SGI_FECHTPFU" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    
    sSql = sSql & "       SGI_CADCLIENTE  SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "      ,SGI_CADCONDPGTO SGI_CADCONDPGTO" & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA1 & " " & strNOMTABELA1 & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & " SGI_CADORDFATI" & strFILIALORD & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & " " & strNOMTABELA2 & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADTRANSP   SGI_CADTRANSP" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strFILIALORD & "   SGI_ORDEMPROD" & strFILIALORD & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORD      = " & objCADORDFAT.CODORD & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_QTDFAT      > 0" & vbCrLf
    
    sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_FILIAL      = SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_CODORD      = SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORD " & vbCrLf

    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL      = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_IDPRODUTO   = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_FILIAL      = " & strNOMTABELA2 & ".SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_CODPED      = " & strNOMTABELA2 & ".SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_FILIAL     = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_CODCLI     = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_FILIAL     = SGI_CADCONDPGTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_CODCONDPGT = SGI_CADCONDPGTO.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And " & strNOMTABELA2 & " .SGI_FILIAL    = SGI_CADTRANSP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_CODTRANSP  = SGI_CADTRANSP.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL      = SGI_ORDEMPROD" & strFILIALORD & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORDFAB   = SGI_ORDEMPROD" & strFILIALORD & ".SGI_CODIGO" & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não Há dados Para Imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If intFILIALPED = 0 Then strNOMARQ = "RELORDFAT.RPT"
    If intFILIALPED = 1 Then strNOMARQ = "RELORDFATSTEEL.RPT"
    
   Call objRel.REL(FILIAL, _
                   sSql, _
                   strCamRelNovo & cCamRelPedidoVendas & Trim(strNOMARQ), _
                   Linha, _
                   1, _
                   "Ordem de Faturamento", _
                   "Ordem de Faturamento", _
                   False, _
                   strAcesso, _
                   True)
    
   Exit Sub
   
Err_Imp:

    MsgBox "Erro : " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Erro"

End Sub

Private Sub ConfFiltro()

    cboFiltro.Clear
    
    If stOrdFat.Tab = 0 Then
    
        cboFiltro.AddItem "Nº Ordem"
        cboFiltro.AddItem "Data Ordem"
        cboFiltro.AddItem "Cliente"
   
    Else
    
        cboFiltro.AddItem "Nº Ordem"
        cboFiltro.AddItem "Data Ordem"
        cboFiltro.AddItem "Cliente"
        cboFiltro.AddItem "Cód.Pedido"
        cboFiltro.AddItem "Cód.OP"
        cboFiltro.AddItem "Cód.NF"
    
    End If
    
    cboFiltro.ListIndex = 0

End Sub

Private Sub DestroiObjeto()
    If adoBanco_Dados.State = 1 Then adoBanco_Dados.Close
    Set objFuncoes = Nothing
    Set objCADORDFAT = Nothing
    Set objRel = Nothing
End Sub


Private Sub ImpOrdNovo()

On Error GoTo Err_Imp

    Dim strNOMARQ       As String
    Dim strFILIALORD    As String
        
    strFILIALORD = ""
    If intFILIALPED = 1 Then strFILIALORD = "_STEEL"
        
    sSql = ""
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "       SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORD" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_CODPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_QTDFAT" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_VLUNIT" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_VLFATURADO" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_VLDOPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORDFAB" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_CODFORN" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & ".SGI_VLTOTAL" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADORDFATH" & strFILIALORD & ".SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATH" & strFILIALORD & ".SGI_DATAORDEM" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATH" & strFILIALORD & ".SGI_OBS" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strFILIALORD & ".SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strFILIALORD & ".SGI_EMAIL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strFILIALORD & ".SGI_CODCONDPGT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strFILIALORD & ".SGI_CODTRANSP" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RGCGC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ENDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_BAINROM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ESTNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CEPNORM" & vbCrLf
    
    sSql = sSql & "      ,SGI_ORDEMPROD" & strFILIALORD & ".SGI_FECHTPFU" & vbCrLf
    
''    sSql = sSql & "      ,SGI_CADFECHAM.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADFECHAM.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "      ,SGI_CADCONDPGTO SGI_CADCONDPGTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATH" & strFILIALORD & " SGI_CADORDFATH" & strFILIALORD & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strFILIALORD & " SGI_CADORDFATI" & strFILIALORD & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strFILIALORD & " SGI_CADPEDVENDH" & strFILIALORD & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADTRANSP SGI_CADTRANSP" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strFILIALORD & " SGI_ORDEMPROD" & strFILIALORD & vbCrLf
    sSql = sSql & "      ,SGI_CADFECHAM SGI_CADFECHAM" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORD      = " & objCADORDFAT.CODORD & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_QTDFAT      > 0" & vbCrLf
    
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL      = SGI_CADORDFATH" & strFILIALORD & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORD      = SGI_CADORDFATH" & strFILIALORD & ".SGI_CODORD" & vbCrLf
    
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL      = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_IDPRODUTO   = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_FILIAL      = SGI_ORDEMPROD" & strFILIALORD & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATI" & strFILIALORD & ".SGI_CODORDFAB   = SGI_ORDEMPROD" & strFILIALORD & ".SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADORDFATH" & strFILIALORD & ".SGI_FILIAL      = SGI_CADPEDVENDH" & strFILIALORD & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADORDFATH" & strFILIALORD & ".SGI_CODPED      = SGI_CADPEDVENDH" & strFILIALORD & ".SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH" & strFILIALORD & ".SGI_FILIAL     = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH" & strFILIALORD & ".SGI_CODCLI     = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH" & strFILIALORD & ".SGI_FILIAL     = SGI_CADCONDPGTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH" & strFILIALORD & ".SGI_CODCONDPGT = SGI_CADCONDPGTO.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH" & strFILIALORD & ".SGI_FILIAL     = SGI_CADTRANSP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH" & strFILIALORD & ".SGI_CODTRANSP  = SGI_CADTRANSP.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD" & strFILIALORD & ".SGI_FILIAL = SGI_CADFECHAM.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strFILIALORD & ".SGI_FECHTPFU = SGI_CADFECHAM.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não Há dados Para Imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If intFILIALPED = 0 Then strNOMARQ = "RELORDFATNOVALATA.RPT"
    If intFILIALPED = 1 Then strNOMARQ = "RELORDFATSTEELNOVO.RPT"
    
   Call objRel.REL(FILIAL, _
                   sSql, _
                   strCamRelNovo & cCamRelPedidoVendas & Trim(strNOMARQ), _
                   Linha, _
                   1, _
                   "Ordem de Faturamento", _
                   "Ordem de Faturamento", _
                   False, _
                   strAcesso, _
                   True)
    
   Exit Sub
   
Err_Imp:

    MsgBox "Erro : " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Erro"

End Sub



Private Sub ConfGridOrdFatGeral()

    With grdGERAL
    
       .Cols = conColumnsIn_OrdFatGeral
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdFatGeral_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_OrdFatGeral_Selecionado) = ""
       .ColDataType(conCOL_OrdFatGeral_Selecionado) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_OrdFatGeral_Codigo) = ""
       .ColDataType(conCOL_OrdFatGeral_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFatGeral_DataOrdem) = ""
       .ColDataType(conCOL_OrdFatGeral_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdFatGeral_Cliente) = ""
       .ColDataType(conCOL_OrdFatGeral_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFatGeral_CodEmp) = ""
       .ColDataType(conCOL_OrdFatGeral_CodEmp) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFatGeral_NomeEmp) = ""
       .ColDataType(conCOL_OrdFatGeral_NomeEmp) = flexDTString
       
       .ColWidth(conCOL_OrdFatGeral_Selecionado) = 300
       .ColWidth(conCOL_OrdFatGeral_Codigo) = 1500
       .ColWidth(conCOL_OrdFatGeral_DataOrdem) = 1000
       .ColWidth(conCOL_OrdFatGeral_Cliente) = 5000
       .ColWidth(conCOL_OrdFatGeral_CodEmp) = 0
       .ColWidth(conCOL_OrdFatGeral_NomeEmp) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

