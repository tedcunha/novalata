VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRELMAPAPROD 
   Caption         =   "Relatório de Mapa de Produção"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   12855
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   "[ OP's ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4695
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   12735
      Begin VSFlex8LCtl.VSFlexGrid VSFlexGrid1 
         Height          =   4335
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   12495
         _cx             =   22040
         _cy             =   7646
         Appearance      =   1
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
   Begin VB.Frame Frame6 
      Caption         =   "[ Carga do Banco ]"
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
      Height          =   615
      Left            =   8640
      TabIndex        =   16
      Top             =   960
      Width           =   2775
      Begin VB.OptionButton optCargBco 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCargBco 
         Caption         =   "Por Periodo"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Linhas ]"
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
      Height          =   2055
      Left            =   7200
      TabIndex        =   14
      Top             =   1560
      Width           =   5535
      Begin VSFlex8LCtl.VSFlexGrid grdLINHA 
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5295
         _cx             =   9340
         _cy             =   2990
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
   Begin VB.Frame Frame4 
      Caption         =   "[ Capacidade ]"
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
      Height          =   2055
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   7095
      Begin VSFlex8LCtl.VSFlexGrid grdCAPAC 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6855
         _cx             =   12091
         _cy             =   2990
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
      Caption         =   "[ Filial ]"
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
      Height          =   615
      Left            =   5520
      TabIndex        =   8
      Top             =   960
      Width           =   3015
      Begin VB.OptionButton optFilial 
         Caption         =   "NOVALATA"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optFilial 
         Caption         =   "STEEL"
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
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   5415
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
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
         Caption         =   "Data Inicial"
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final"
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
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmdCarrega 
         Caption         =   "&Carrega"
         Height          =   615
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
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
         Picture         =   "frmRELMAPAPROD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmRELMAPAPROD.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELMAPAPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public FILIAL        As Integer
Public strAcesso     As String
Public lngCodUsuario As Long

Dim objBLBFunc       As Object
Dim objRELMAPAPROD   As New clsRELMAPAPROD
Dim objPESQPADRAO    As Object
Dim objREL           As Object
Dim objFUNC_CWS      As New clsFUNC_CWS


Dim strCABEC1        As String
Dim strCABEC2        As String

Dim lngPORC          As Long

Const conCOL_MapaCapc_Codigo                        As Integer = 0
Const conCOL_MapaCapc_CodLinha                      As Integer = 1
Const conCOL_MapaCapc_DescCapc                      As Integer = 2
Const conCOL_MapaCapc_FormatString                  As String = "=Codigo|Cód.linha|Descrição da Capacidade"
Const conColumnsIn_MapaCapc                         As Integer = 3

Const conCOL_MapaLinha_Codigo                       As Integer = 0
Const conCOL_MapaLinha_DescLinha                    As Integer = 1
Const conCOL_Mapalinha_FormatString                 As String = "=Codigo|Descrição da Linha"
Const conColumnsIn_MapaLinha                        As Integer = 2

Const conCOL_OP_Codigo                              As Integer = 0
Const conCOL_OP_DatOP                               As Integer = 1
Const conCOL_OP_FormatString                        As String = "=Codigo OP|Data.OP"
Const conColumnsIn_OP                               As Integer = 2


Private Sub cmdCarrega_Click()
    Call CarregaCapacidade
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

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    ''Set objRELPREPARA = CreateObject("RELPCP.clsRELPREPARA")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    
    ''Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    ''If adoBanco_Dados.State = 0 Then
    ''   MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
    ''   Exit Sub
    ''End If
    
    objBLBFunc.LimpaCampos Me
    objRELMAPAPROD.FILIAL = FILIAL

    Call ConfGrdCapacidade
    Call ConfGrdLinha
    
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Now, "DD/MM/YYYY")

    optFilial(0).value = True
    optCargBco(1).value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Campos
End Sub

Private Sub grdCAPAC_Click()
    If (grdCAPAC.Rows - 1) = 0 Then Exit Sub
    If grdCAPAC.Row = 0 Then Exit Sub
    Call Carregalinhas(grdCAPAC.Cell(flexcpText, grdCAPAC.Row, conCOL_MapaCapc_Codigo))
End Sub

Private Sub grdCAPAC_RowColChange()
    If (grdCAPAC.Rows - 1) = 0 Then Exit Sub
    If grdCAPAC.Row = 0 Then Exit Sub
    Call Carregalinhas(grdCAPAC.Cell(flexcpText, grdCAPAC.Row, conCOL_MapaCapc_Codigo))
End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, Me
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, Me
End Sub

Private Sub Destroy_Campos()
    Set objBLBFunc = Nothing
    Set objRELMAPAPROD = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
    Set objFUNC_CWS = Nothing
End Sub

Private Sub ConfGrdCapacidade()

    With grdCAPAC

       .Cols = conColumnsIn_MapaCapc
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_MapaCapc_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_MapaCapc_Codigo) = ""
       .ColDataType(conCOL_MapaCapc_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_MapaCapc_CodLinha) = ""
       .ColDataType(conCOL_MapaCapc_CodLinha) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_MapaCapc_DescCapc) = ""
       .ColDataType(conCOL_MapaCapc_DescCapc) = flexDTString
       
       .COLWIDTH(conCOL_MapaCapc_Codigo) = 0
       .COLWIDTH(conCOL_MapaCapc_CodLinha) = 1000
       .COLWIDTH(conCOL_MapaCapc_DescCapc) = 5000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack

    End With
    
End Sub


Private Sub CarregaCapacidade()

    Call ConfGrdCapacidade

    sSql = "Select" & vbCrLf
    sSql = sSql & "      LI.SGI_CODIGO" & vbCrLf
    sSql = sSql & "     ,LI.SGI_CODLIN" & vbCrLf
    sSql = sSql & "     ,LI.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_ORDEMPROD_STEEL      OP" & vbCrLf
    sSql = sSql & "     ,SGI_CADPRODUTO           CP" & vbCrLf
    sSql = sSql & "     ,SGI_CADLINHAPRODUTO      LI" & vbCrLf
    sSql = sSql & "     ,SGI_CADGRUPLINHAIT_STEEL GR" & vbCrLf
 
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      OP.SGI_FILIAL   = " & FILIAL & vbCrLf
    ''--   And OP.SGI_DATENTREGA Between '2010-01-01' And '2013-09-20'
    sSql = sSql & "   And (OP.SGI_STATUS    = 0 Or OP.SGI_STATUS = 2)" & vbCrLf
    sSql = sSql & "   And CP.SGI_FILIAL     = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CP.SGI_IDPRODUTO  = OP.SGI_IDPRODUTO" & vbCrLf
    ''--   And CP.SGI_CODLINPROD = 900
    sSql = sSql & "   And LI.SGI_FILIAL     = CP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LI.SGI_CODLIN     = CP.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "   And GR.SGI_FILIAL     = LI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GR.SGI_CODLIN     = LI.SGI_CODIGO" & vbCrLf
    sSql = sSql & "Group By" & vbCrLf
    sSql = sSql & "      LI.SGI_CODIGO" & vbCrLf
    sSql = sSql & "     ,LI.SGI_CODLIN" & vbCrLf
    sSql = sSql & "     ,LI.SGI_DESCRI"

    
    Call objFUNC_CWS.AbreBanco(Linha)
    
    BREC.Open sSql, BD, adOpenDynamic
    With grdCAPAC
        Do While Not BREC.EOF()
            .AddItem BREC!SGI_CODIGO & vbTab & _
                     BREC!SGI_CODLIN & vbTab & _
                     BREC!SGI_DESCRI
                     
            BREC.MoveNext
        Loop
    End With
    BREC.Close

    Call objFUNC_CWS.FechaBanco


End Sub


Private Sub ConfGrdLinha()

    With grdLINHA

       .Cols = conColumnsIn_MapaLinha
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Mapalinha_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_MapaLinha_Codigo) = ""
       .ColDataType(conCOL_MapaLinha_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_MapaLinha_DescLinha) = ""
       .ColDataType(conCOL_MapaLinha_DescLinha) = flexDTString
       
       .COLWIDTH(conCOL_MapaLinha_Codigo) = 1000
       .COLWIDTH(conCOL_MapaLinha_DescLinha) = 5000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack

    End With
    
End Sub

Private Sub Carregalinhas(strCODLIN As String)

    If Len(Trim(strCODLIN)) = 0 Then Exit Sub
    
    Call ConfGrdLinha
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       CB.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,CB.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADGRUPLINHAIT_STEEL IT" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHA_STEEL   CB" & vbCrLf
  
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       IT.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And IT.SGI_CODLIN = " & Trim(strCODLIN) & vbCrLf
    sSql = sSql & "   And CB.SGI_FILIAL = IT.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CB.SGI_CODIGO = IT.SGI_CODIGO"

    Call objFUNC_CWS.AbreBanco(Linha)
    
    BREC10.Open sSql, BD, adOpenDynamic
    With grdLINHA
        Do While Not BREC10.EOF()
            .AddItem BREC10!SGI_CODIGO & vbTab & _
                     BREC10!SGI_DESCRI
            BREC10.MoveNext
        Loop
    End With
    BREC10.Close
    
    Call objFUNC_CWS.FechaBanco
    
End Sub

