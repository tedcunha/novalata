VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADOORDFAB 
   Caption         =   "Cadastro de Ordem de Fabricação"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   12840
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame27 
      Caption         =   "[ Motivo da Liquidação da OP ]"
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
      Height          =   2175
      Left            =   6000
      TabIndex        =   22
      Top             =   5280
      Width           =   6735
      Begin VB.Frame Frame28 
         Caption         =   "[ Motivo ]"
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   6495
         Begin VB.TextBox txtCODMOTLIQ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   27
            Text            =   "txtCODMOTL"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   1200
            Picture         =   "frmCADOORDFAB.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblDescMotLiq 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescMotLiq"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1560
            TabIndex        =   28
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "[ Observação ]"
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
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   6495
         Begin VB.TextBox txtOBS_MotLiq 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   24
            Text            =   "frmCADOORDFAB.frx":0102
            Top             =   240
            Width           =   6255
         End
      End
   End
   Begin VB.Frame Frame26 
      Caption         =   "[ Log ]"
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
      Height          =   2655
      Left            =   6000
      TabIndex        =   20
      Top             =   2640
      Width           =   6735
      Begin VSFlex8LCtl.VSFlexGrid grdLogPed 
         Height          =   2295
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   6495
         _cx             =   11456
         _cy             =   4048
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
   Begin VB.CommandButton cmdMostraOP 
      Caption         =   "Mostra OP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   12735
      Begin VB.TextBox txtPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Text            =   "txtPedido"
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
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
         Height          =   255
         Index           =   3
         Left            =   9840
         TabIndex        =   18
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblSTATUS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSTATUS"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10560
         TabIndex        =   17
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblDataPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDataPedido"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8640
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Dt.Pedido"
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
         Height          =   255
         Index           =   2
         Left            =   7680
         TabIndex        =   15
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCliente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2640
         TabIndex        =   14
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.OP"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Rótulo ]"
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
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   12735
      Begin VSFlex8LCtl.VSFlexGrid grdProdutos 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   12495
         _cx             =   22040
         _cy             =   1720
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
      Caption         =   "[ OP - Gerada ]"
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
      Height          =   3015
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   5895
      Begin VB.CommandButton Command4 
         Height          =   300
         Left            =   5520
         Picture         =   "frmCADOORDFAB.frx":0110
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exclui a linha da Gride Selecionada"
         Top             =   600
         Width           =   300
      End
      Begin VB.CommandButton Command5 
         Height          =   300
         Left            =   5520
         Picture         =   "frmCADOORDFAB.frx":025A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   240
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdProgEntrega 
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5295
         _cx             =   9340
         _cy             =   4683
         Appearance      =   2
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
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12735
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
         Picture         =   "frmCADOORDFAB.frx":03A4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
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
         Picture         =   "frmCADOORDFAB.frx":08D6
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
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
         Picture         =   "frmCADOORDFAB.frx":09D8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADOORDFAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCodVendedor   As Long
Public lngCodUsuario    As Long
Public intFILIALPED     As Integer

Dim arrPRODUTOS         As Variant
Dim arrORDFATGERADAS    As Variant
Dim objBLBFunc          As Object
Dim objCADOORDFAB       As Object
Dim objPESQPADRAO       As Object
Dim strTabela1          As String
Dim strTabela2          As String
Dim strTabela3          As String
Dim strNomeMod          As String
Dim strNOMFILIAL        As String

'' -----------------------------------------------------------------------------------
Const conCOL_Produto_IDProduto                  As Integer = 0
Const conCOL_Produto_CodProduto                 As Integer = 1
Const conCOL_Produto_DescProduto                As Integer = 2
Const conCOL_Produto_QtdePedido                 As Integer = 3
Const conCOL_Produto_QtdeJaProgramada           As Integer = 4
Const conCOL_Produto_Saldo                      As Integer = 5
Const conCOL_Produto_DataOrdem                  As Integer = 6
Const conCOL_Produto_DataEntrega                As Integer = 7
Const conCOL_Produto_Action2Do                  As Integer = 8
Const conCOL_Produto_Codped                     As Integer = 9

Const conCOL_Produto_FormatString               As String = "=IDProduto|Produto|Descrição|Qtd.Pedido|Qtd.Já.Prog.|Saldo|Dt.Ordem|Dt.Entrega|Action2Do|Cod.Ped"
Const conColumnsIn_Produto                      As Integer = 10

'' ========================================================================================
Const conCOL_SonProgEntr_IdProduto              As Integer = 0
Const conCOL_SonProgEntr_QtdProd                As Integer = 1
Const conCOL_SonProgEntr_DataEntrega            As Integer = 2
Const conCOL_SonProgEntr_CodOP                  As Integer = 3
Const conCOL_SonProgEntr_Action2Do              As Integer = 4
Const conCOL_SonProgEntr_FormatString           As String = "=Cod|Quantidade|Dt.Entrega|Cod.OP|Action2Do"
Const conColumnsIn_SonProgEntr                  As Integer = 5

'' ========================================================================================
Const conCOL_SonLogOP_Data                     As Integer = 0
Const conCOL_SonLogOP_Hora                     As Integer = 1
Const conCOL_SonLogOP_CodUsuario               As Integer = 2
Const conCOL_SonLogOP_Usuario                  As Integer = 3
Const conCOL_SonLogOP_CodAcao                  As Integer = 4
Const conCOL_SonLogOP_Acao                     As Integer = 5
Const conCOL_SonLogOP_Tipo                     As Integer = 6
Const conCOL_SonLogOP_FormatString             As String = "=Data|Hora|CodUsuario|Usuário|CodAcao|Ação|Tipo"
Const conColumnsIn_SonLogOP                    As Integer = 7


Private Sub cmdAltera_Click()
    If objCADOORDFAB.STATUS = 1 Or objCADOORDFAB.STATUS = 2 Then
        MsgBox "Esta Ordem de Produção não ser Alterada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If objCADOORDFAB.TIPO = 1 Then
        MsgBox "Esta Ordem Não pode ser Alterada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
End Sub

Private Sub cmdMostraOP_Click()
    frmMOSTRAOP.Show vbModal
End Sub

Private Sub CmdSalva_Click()

On Error GoTo err_grava
    
    If Not VerifCampos Then Exit Sub
    
    Dim I           As Integer
    Dim intRESP     As Integer
    Dim lngCodLog   As Long
    Dim strNOMFORM  As String
    
    ''objCADOORDFAB.CODPEDIDO = CLng(txtPedido.Text)
    ''objCADOORDFAB.DATAPED = CDate(lblDataPedido.Caption)
    If cTipOper = "I" Then objCADOORDFAB.TIPO = 0
    
    Call VerifOrdFatAberto(Trim(Str(objCADOORDFAB.CODPEDIDO)))
    
    If cTipOper = "BX" Then
        objCADOORDFAB.OBSLIQ = "'" & Trim(Replace(txtOBS_MotLiq.Text, ",", " ")) & "'"
        objCADOORDFAB.CODMOTLIQOP = Trim(txtCODMOTLIQ.Text)
    End If
    
    '' Itens da Ordem de Fabricação
    objCADOORDFAB.PRODUTOS = Empty
    With grdProgEntrega
        If (.Rows - 1) > 0 Then
            ReDim arrPRODUTOS(1 To (.Rows - 1), 0 To 11) As String
            
            For I = 1 To (.Rows - 1)
                
                arrPRODUTOS(I, conCOL_SonProgEntr_IdProduto) = .Cell(flexcpText, I, conCOL_SonProgEntr_IdProduto)
                arrPRODUTOS(I, conCOL_SonProgEntr_QtdProd) = .Cell(flexcpText, I, conCOL_SonProgEntr_QtdProd)
                arrPRODUTOS(I, conCOL_SonProgEntr_DataEntrega) = .Cell(flexcpText, I, conCOL_SonProgEntr_DataEntrega)
                arrPRODUTOS(I, conCOL_SonProgEntr_Action2Do) = .Cell(flexcpText, I, conCOL_SonProgEntr_Action2Do)
                
                arrPRODUTOS(I, 4) = objBLBFunc.Gera_Codigo("frmCADOORDFAB", FILIAL, Linha) & Year(Now)
                arrPRODUTOS(I, 5) = Now
                
                arrPRODUTOS(I, 6) = grdProdutos.Cell(flexcpText, grdProdutos.Row, conCOL_Produto_CodProduto)
                arrPRODUTOS(I, 7) = 0
                arrPRODUTOS(I, 8) = 0
                arrPRODUTOS(I, 9) = 1
                arrPRODUTOS(I, 10) = grdProdutos.Cell(flexcpText, grdProdutos.Row, conCOL_Produto_Codped)
                arrPRODUTOS(I, 11) = Trim(txtPedido.Text)
                
                '' ===============================
                '' Pega os Dados do Produto
                ''For j = 1 To (grdProduto.Rows - 1)
                ''    If CLng(grdProduto.Cell(flexcpText, j, conCOL_SonProd_IdProduto)) = CLng(.Cell(flexcpText, I, conCOL_SonProgEntr_IdProduto)) Then
                ''
                ''    End If
                ''Next j
                '' ===============================
            Next I
            
            objCADOORDFAB.PRODUTOS = arrPRODUTOS
        End If
    End With
    
    '' Gravando a Ordem de Fabricação
    If objCADOORDFAB.GRAVA(cTipOper, Trim(strTabela1), strNOMFILIAL) = False Then Exit Sub
    
    '' Gerando Log de Sistema
    For I = 1 To UBound(arrPRODUTOS)
        If Len(Trim(arrPRODUTOS(I, 2))) > 0 Then
            lngCodLog = objBLBFunc.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
            Call objBLBFunc.GravaLogModulo(FILIAL, lngCodLog, Trim(strNomeMod), cTipOper, lngCodUsuario, Str(arrPRODUTOS(I, 1)), Linha)
        End If
    Next I
    
    Call objBLBFunc.GravaLogForm(FILIAL, objCADOORDFAB.CODORDEM, lngCodUsuario, cTipOper, strNomeMod)
    
    
    MsgBox "A Ordem de Fabricação foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", IIf(cTipOper = "BX", "Liquidada", ""))) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    
    If cTipOper = "I" Then
       intRESP = MsgBox("Deseja gerar outra Ordem de Fabricação ?", vbYesNo + vbQuestion, "Aviso")
       
       If intRESP = 6 Then
          Call Inclui
       Else
          Set objBLBFunc = Nothing
          Set objCADOORDFAB = Nothing
          Set objPESQPADRAO = Nothing
          Unload Me
       End If
    ElseIf cTipOper = "BX" Then
          Unload Me
    End If
    
    Exit Sub
    
err_grava:
    
    MsgBox "Erro nº  : " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, vbOKOnly + vbCritical, "Aviso"

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    If cTipOper = "I" Or cTipOper = "A" Then
       Call ExcLinhaGrd(grdProgEntrega, grdProgEntrega.Row, conCOL_SonProgEntr_Action2Do)
       With grdProdutos
            Call objBLBFunc.CarregaDadosGrdFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, .Cell(flexcpText, .Row, conCOL_Produto_IDProduto))
       End With
    End If
End Sub

Private Sub Command5_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGridProg
End Sub

Private Sub Command6_Click()

On Error GoTo Err_Command6_Click

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRI " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADMOTLIQOP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Motivos de Liquidação do OP")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODMOTLIQ.Text = varRETORNO
    
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMOTLIQOP", varRETORNO, lblDescMotLiq)
    If Len(Trim(lblDescMotLiq.Caption)) = 0 Then txtCODMOTLIQ.Text = ""
    
    If txtOBS_MotLiq.Enabled = True Then txtOBS_MotLiq.SetFocus

    Exit Sub
    
Err_Command6_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command6_Click()", Me.Name, "Command6_Click()", strCAMARQERRO)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADOORDFAB = CreateObject("CADOORDFAB.clsCADOORDFAB")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADOORDFAB.FILIAL = FILIAL
   
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   
    strNOMFILIAL = ""
    If intFILIALPED = 0 Then
        strTabela1 = "SGI_ORDEMPROD"
        strTabela2 = "SGI_CADPEDVENDH"
        strTabela3 = "SGI_CADPEDVENDI"
        strNomeMod = Trim(Me.Name)
    ElseIf intFILIALPED = 1 Then
        strTabela1 = "SGI_ORDEMPROD_STEEL"
        strTabela2 = "SGI_CADPEDVENDH_STEEL"
        strTabela3 = "SGI_CADPEDVENDI_STEEL"
        strNomeMod = Trim(Me.Name) & "_STEEL"
        strNOMFILIAL = "_STEEL"
    End If
   
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta
   If cTipOper = "BX" Then Liquida

End Sub


Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame28.Enabled = False
    txtOBS_MotLiq.Locked = True
    
    
    Me.Caption = "Cadastro de Confirmação de Faturamento - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos Me
    
    Call LimpaCamposLabel
    Call ConfGridProdutos
    Call InitGridProg
    Call InitGridLogOP
    
    objCADOORDFAB.STATUS = 0
    lblSTATUS.Caption = "Em Aberto"
    
End Sub

Private Sub LimpaCamposLabel()
    lblCliente.Caption = ""
    lblDataPedido.Caption = ""
    lblSTATUS.Caption = ""
    lblDescMotLiq.Caption = ""
End Sub

Private Sub ConfGridProdutos()

    With grdProdutos
    
       .Cols = conColumnsIn_Produto
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Produto_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_Produto_IDProduto) = ""
       .ColDataType(conCOL_Produto_IDProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_CodProduto) = ""
       .ColDataType(conCOL_Produto_CodProduto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Produto_DescProduto) = ""
       .ColDataType(conCOL_Produto_DescProduto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Produto_QtdePedido) = ""
       .ColDataType(conCOL_Produto_QtdePedido) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_QtdeJaProgramada) = ""
       .ColDataType(conCOL_Produto_QtdeJaProgramada) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_Saldo) = ""
       .ColDataType(conCOL_Produto_Saldo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_DataOrdem) = ""
       .ColDataType(conCOL_Produto_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_Produto_Action2Do) = ""
       .ColDataType(conCOL_Produto_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_DataEntrega) = ""
       .ColDataType(conCOL_Produto_DataEntrega) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_Produto_Codped) = ""
       .ColDataType(conCOL_Produto_Codped) = flexDTLong
       
       .ColWidth(conCOL_Produto_IDProduto) = 0
       .ColWidth(conCOL_Produto_CodProduto) = 1200
       .ColWidth(conCOL_Produto_DescProduto) = 4000
       .ColWidth(conCOL_Produto_QtdePedido) = 1000
       .ColWidth(conCOL_Produto_QtdeJaProgramada) = 1000
       .ColWidth(conCOL_Produto_Saldo) = 1000
       .ColWidth(conCOL_Produto_DataOrdem) = 1000
       .ColWidth(conCOL_Produto_DataEntrega) = 1000
       .ColWidth(conCOL_Produto_Action2Do) = 0
       .ColWidth(conCOL_Produto_Codped) = 1000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub


Private Sub grdProgEntrega_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdProgEntrega
          Select Case Col
                 Case conCOL_SonProgEntr_QtdProd
                    If CalcTotEntregas(CCur(grdProdutos.Cell(flexcpText, grdProdutos.Row, conCOL_Produto_Saldo)), CLng(.Cell(flexcpText, Row, conCOL_SonProgEntr_IdProduto))) Then
                        ''.Cell(flexcpText, Row, Col) = Format(.Cell(flexcpText, Row, Col), "#,##0.00")
                        If Col = conCOL_SonProgEntr_QtdProd Then
                           .Col = (Col + 1)
                           .EditCell
                        End If
                    Else
                        .Cell(flexcpText, Row, Col) = Empty
                    End If
          End Select
     End With
End Sub

Private Sub grdProgEntrega_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonProgEntr_QtdProd, _
         conCOL_SonProgEntr_DataEntrega
         If cTipOper = "C" Then Cancel = True
    Case conCOL_SonProgEntr_CodOP
         Cancel = True
    Case Else
        grdProgEntrega.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdProgEntrega_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdProgEntrega
          Select Case Col
                    Case conCOL_SonProgEntr_QtdProd
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_SonProgEntr_DataEntrega
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          End Select
     End With
End Sub

Private Sub grdProgEntrega_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdProgEntrega
          Select Case Col
                 Case conCOL_SonProgEntr_QtdProd, conCOL_SonProgEntr_DataEntrega
                        If .EditText = Empty Then Exit Sub
                        If Col = conCOL_SonProgEntr_DataEntrega Then
                            If Not IsDate(.EditText) Then
                                MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                                Cancel = True
                            Else
                                If Not IsDate(.EditText) Then
                                    MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                                    Cancel = True
                                    Exit Sub
                                End If
                                If CDate(grdProdutos.Cell(flexcpText, grdProdutos.Row, conCOL_Produto_DataEntrega)) > CDate(.EditText) Then
                                    MsgBox "A data de entrega não pode ser menor que a data original de entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                                    Cancel = True
                                End If
                            End If
                        End If
          End Select
     End With
End Sub

Private Sub txtCODMOTLIQ_GotFocus()

On Error GoTo Err_txtCODMOTLIQ_GotFocus
    
    objBLBFunc.SelecionaCampos txtCODMOTLIQ.Name, Me

    Exit Sub
    
Err_txtCODMOTLIQ_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODMOTLIQ_GotFocus()", Me.Name, "txtCODMOTLIQ_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCODMOTLIQ_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCODMOTLIQ_KeyPress

    objBLBFunc.SoNumeroPonto KeyAscii, txtCODMOTLIQ.Text

    Exit Sub
    
Err_txtCODMOTLIQ_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODMOTLIQ_KeyPress()", Me.Name, "txtCODMOTLIQ_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCODMOTLIQ_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODMOTLIQ_Validate

    Dim I As Integer
    
    If Len(Trim(txtCODMOTLIQ.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODMOTLIQ.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODMOTLIQ.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMOTLIQOP", txtCODMOTLIQ.Text, lblDescMotLiq)
    If Len(Trim(lblDescMotLiq.Caption)) = 0 Then
       txtCODMOTLIQ.Text = ""
       Cancel = True
    Else
        txtOBS_MotLiq.SetFocus
    End If
    
    Exit Sub
    
Err_txtCODMOTLIQ_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODMOTLIQ_Validate()", Me.Name, "txtCODMOTLIQ_Validate()", strCAMARQERRO)

End Sub

Private Sub txtPedido_GotFocus()
    objBLBFunc.SelecionaCampos txtPedido.Name, frmCADOORDFAB
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPedido.Text
End Sub

Private Sub txtPedido_Validate(Cancel As Boolean)

    If Len(Trim(txtPedido.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPedido.Text) Then
       MsgBox "Somente é Permitido Numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtPedido.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Cancel = PegaDadosDoPedido(Trim(txtPedido.Text))

End Sub

Private Function PegaDadosDoPedido(strCODPEDIDO As String) As Boolean

    PegaDadosDoPedido = True
    
    Call LimpaCamposLabel
    Call ConfGridProdutos
    
    Dim curTotItensPed  As Currency
    Dim curTotOFjaProg  As Currency
    Dim curSALDO        As Currency
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO AS SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLI.* " & vbCrLf
    sSql = sSql & "      ,ORDP.*" & vbCrLf
    sSql = sSql & "      ,PROD.*" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strTabela1 & " ORDP" & vbCrLf
    sSql = sSql & "      ," & strTabela2 & " PED" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  PROD" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO    = " & Trim(strCODPEDIDO) & vbCrLf
    sSql = sSql & "   And (ORDP.SGI_STATUS   = 0 Or ORDP.SGI_STATUS = 1)"
    sSql = sSql & "   And PED.SGI_FILIAL     = ORDP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO     = ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL     = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO     = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL    = ORDP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = ORDP.SGI_IDPRODUTO "

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        
        objCADOORDFAB.CODPEDIDO = CLng(strCODPEDIDO)
        lblCliente.Caption = Trim(BREC!SGI_CODIGO) & " - " & Trim(BREC!SGI_RAZAOSOC)
        lblDataPedido.Caption = Format(BREC!SGI_DATAPED, "DD/MM/YYYY")
        
        With grdProdutos
                 
             curTotItensPed = 0
             curTotOFjaProg = 0
             curSALDO = 0
             
             '' --------------------------------------------
             If Not IsNull(BREC!SGI_QTDEPED) Then curTotItensPed = BREC!SGI_QTDEPED
             If Not IsNull(BREC!SGI_QTDFAT) Then curTotOFjaProg = BREC!SGI_QTDFAT
             curSALDO = (curTotItensPed - curTotOFjaProg)
             
             '' --------------------------------------------
            
            .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                     Trim(BREC!SGI_CODPROD) & vbTab & _
                     Trim(BREC!SGI_DESCRICAO) & vbTab & _
                     curTotItensPed & vbTab & _
                     curTotOFjaProg & vbTab & _
                     curSALDO & vbTab & _
                     Format(Now, "DD/MM/YYYY") & vbTab & _
                     IIf(IsNull(BREC!SGI_DATENTREGA), Format(Now, "DD/MM/YYYY"), Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY")) & vbTab & _
                     dacEnumUpdateAction_Insert & vbTab & _
                     BREC!SGI_CODPED
                              
        End With
        
    Else
        MsgBox "Esta OP Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
        txtPedido.Text = ""
        txtPedido.SetFocus
        BREC.Close
        Exit Function
    End If
    BREC.Close
    
    PegaDadosDoPedido = False
    
End Function


Private Function CalcSaldoIten(curQtdeDigit As Currency, lngLinha As Long) As Boolean
    
    CalcSaldoIten = False
    
    ''If curQtdeDigit = 0 Then Exit Function
    
    Dim curQtdePedido       As Currency
    Dim curQtdeJaProgramada As Currency
    Dim curQtde             As Currency
    Dim curSALDO            As Currency

    With grdProdutos
        If Len(Trim(.Cell(flexcpText, lngLinha, conCOL_Produto_QtdePedido))) > 0 Then curQtdePedido = CCur(.Cell(flexcpText, lngLinha, conCOL_Produto_QtdePedido))
        If Len(Trim(.Cell(flexcpText, lngLinha, conCOL_Produto_QtdeJaProgramada))) > 0 Then curQtdeJaProgramada = CCur(.Cell(flexcpText, lngLinha, conCOL_Produto_QtdeJaProgramada))
    
        curQtde = (curQtdePedido - curQtdeJaProgramada)
    
        If curQtdeDigit > curQtde Then
            MsgBox "ATENÇÃO - A quantidade programada não pode ser maior que o saldo !!!", vbOKOnly + vbExclamation, "Aviso"
            CalcSaldoIten = True
            Exit Function
        End If
        
        curSALDO = (curQtde - curQtdeDigit)
        .Cell(flexcpText, lngLinha, conCOL_Produto_Saldo) = Format(curSALDO, "#,##0.00")
    End With
    
End Function

Private Function VerifCampos() As Boolean
    
    VerifCampos = False
    
    Dim I               As Integer
    Dim curQTD_QTDEPROG As Currency
    Dim curQTD_REAL     As Currency
    
    If cTipOper = "I" Or cTipOper = "A" Then
        If Len(Trim(txtPedido.Text)) = 0 Then
           MsgBox "O Código da OP Não foi Informado !!!", vbOKOnly + vbExclamation, "Aviso"
           txtPedido.SetFocus
           Exit Function
        End If
        
        If (grdProgEntrega.Rows - 1) = 0 Then
           MsgBox "Não foi informado a nova OP para Gerar !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
        
        curQTD_REAL = CLng(grdProdutos.Cell(flexcpText, grdProdutos.Row, conCOL_Produto_Saldo))
        With grdProgEntrega
            curQTD_QTDEPROG = 0
            For I = 1 To (.Rows - 1)
                curQTD_QTDEPROG = curQTD_QTDEPROG + CCur(.Cell(flexcpText, I, conCOL_SonProgEntr_QtdProd))
            Next I
        End With
        
        If curQTD_QTDEPROG < curQTD_REAL Then
            MsgBox "A quantidade programada não está igual a quantidade da op original !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        '' ------------------------------------------
        
    Else
        If Len(Trim(txtCODMOTLIQ.Text)) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Não foi inrformado um motivo para liquidação da OP !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If Len(Trim(txtOBS_MotLiq.Text)) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Não foi inrformado uma Observação para liquidação da OP !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    
        If VerificaPedido(Str(objCADOORDFAB.CODPEDIDO), Str(objCADOORDFAB.IdProduto)) = False Then Exit Function
        
    End If
    
    VerifCampos = True
    
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    Me.Caption = "Cadastro de Ordem de Fabricação - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos Me

    Frame2.Enabled = False
    Frame3.Enabled = True
    Frame28.Enabled = False
    txtOBS_MotLiq.Locked = True
    
    objCADOORDFAB.CODORDEM = iCodigo
    
    Call ConfGridProdutos
    Call LimpaCamposLabel
    Call InitGridProg
    Call InitGridLogOP
    
    Call CarregaCampos
    
End Sub

Private Sub CarregaCampos()

    If intFILIALPED = 0 Then '' Novalata
        If objCADOORDFAB.Carrega_Campos = False Then Exit Sub
    ElseIf intFILIALPED = 1 Then '' Steel Roll
        If objCADOORDFAB.Carrega_CamposSteel = False Then Exit Sub
    End If
    
    If objCADOORDFAB.STATUS = 0 Then lblSTATUS.Caption = "Em Aberto"
    If objCADOORDFAB.STATUS = 1 Then lblSTATUS.Caption = "Parcial"
    If objCADOORDFAB.STATUS = 2 Then lblSTATUS.Caption = "Finalizado"
    If objCADOORDFAB.STATUS = 4 Then lblSTATUS.Caption = "Bloqueado"
    If objCADOORDFAB.STATUS = 6 Then lblSTATUS.Caption = "P.Cota"
    If objCADOORDFAB.STATUS = 9 Then lblSTATUS.Caption = "Liquidada Manualmente"
    
    txtPedido.Text = objCADOORDFAB.CODORDEM
    lblDataPedido.Caption = Format(objCADOORDFAB.DATAPED, "DD/MM/YYYY")
    
    If Len(Trim(objCADOORDFAB.CODMOTLIQOP)) > 0 Then
        txtCODMOTLIQ.Text = objCADOORDFAB.CODMOTLIQOP
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMOTLIQOP", txtCODMOTLIQ.Text, lblDescMotLiq)
        txtOBS_MotLiq.Text = objCADOORDFAB.OBSLIQ
    End If
    
    Call CarregaDadosCliente(txtPedido.Text)
    Call PopGrdItens
    Call PopGrdOrdemFilhas
    Call PopLogOP

End Sub

Private Sub CarregaDadosCliente(strCODPED As String)

    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO AS SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLI.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strTabela1 & " ORDP" & vbCrLf
    sSql = sSql & "      ," & strTabela2 & " PED" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO = " & strCODPED & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL = ORDP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO = ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        lblCliente.Caption = Trim(BREC!SGI_CODCLI) & " - " & Trim(BREC!SGI_RAZAOSOC)
    End If
    BREC.Close

End Sub

Private Sub PopGrdItens()

    Dim I As Integer
    
    arrPRODUTOS = objCADOORDFAB.PRODUTOS

    If IsArray(arrPRODUTOS) Then
        With grdProdutos
            For I = 1 To UBound(arrPRODUTOS)
            
                .AddItem arrPRODUTOS(I, 2) & vbTab & _
                         arrPRODUTOS(I, 3) & vbTab & _
                         PegaDadosProduto(Str(arrPRODUTOS(I, 2))) & vbTab & _
                         arrPRODUTOS(I, 4) & vbTab & _
                         arrPRODUTOS(I, 5) & vbTab & _
                         IIf(Len(Trim(arrPRODUTOS(I, 7))) = 0, "", arrPRODUTOS(I, 7)) & vbTab & _
                         Format(arrPRODUTOS(I, 8), "DD/MM/YYYY") & vbTab & _
                         Format(arrPRODUTOS(I, 10), "DD/MM/YYYY") & vbTab & _
                         arrPRODUTOS(I, 9) & vbTab & _
                         arrPRODUTOS(I, 13)
                         
            Next I
            
        End With
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Ordem de Fabricação - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADOORDFAB

    Frame2.Enabled = False
    Frame3.Enabled = True
    Frame28.Enabled = False
    txtOBS_MotLiq.Locked = True
    
    objCADOORDFAB.CODORDEM = iCodigo
    
    Call ConfGridProdutos
    Call LimpaCamposLabel
    Call InitGridProg
    Call InitGridLogOP
    
    Call CarregaCampos
    
End Sub


Private Function SomaTotalItensPedido(strCODPEDIDO As String) As Currency

    SomaTotalItensPedido = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(SGI_QTDE) As SGI_TOTPEDIDO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       " & strTabela3 & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCODPEDIDO
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then SomaTotalItensPedido = BREC10!SGI_TOTPEDIDO
    BREC10.Close

End Function

Private Function SomaSaldoOF(strCODPEDIDO As String) As Currency

    SomaSaldoOF = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(SGI_QTDE) As SGI_SALDO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strTabela1 & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED = " & strCODPEDIDO

    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
        If Not IsNull(BREC10!SGI_SALDO) Then SomaSaldoOF = BREC10!SGI_SALDO
    End If
    BREC10.Close

End Function

Private Function PegaDadosProduto(strIDPRODUTO As String) As String

    PegaDadosProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "    And SGI_IDPRODUTO = " & Trim(strIDPRODUTO)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaDadosProduto = BREC!SGI_DESCRICAO
    BREC.Close
    
End Function

Private Sub InitGridProg()

    With grdProgEntrega
    
       .Cols = conColumnsIn_SonProgEntr
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProgEntr_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_IdProduto) = ""
       .ColDataType(conCOL_SonProgEntr_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DataEntrega) = ""
       .ColDataType(conCOL_SonProgEntr_DataEntrega) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_QtdProd) = ""
       .ColDataType(conCOL_SonProgEntr_QtdProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_CodOP) = ""
       .ColDataType(conCOL_SonProgEntr_CodOP) = flexDTLong
       
       .ColWidth(conCOL_SonProgEntr_IdProduto) = 0
       .ColWidth(conCOL_SonProgEntr_QtdProd) = 1500
       .ColWidth(conCOL_SonProgEntr_DataEntrega) = 1500
       .ColWidth(conCOL_SonProgEntr_CodOP) = 1500
       .ColWidth(conCOL_SonProgEntr_Action2Do) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGridProg()
    
    Dim curQtdProdInf   As Currency
    Dim curPegaSaldo    As Currency
    Dim I               As Integer
    
    
    With grdProdutos
        If .Row <= 0 Then
            MsgBox "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If Len(Trim(.Cell(flexcpText, .Row, conCOL_Produto_Saldo))) = 0 Then
            MsgBox "Primeiro Informe a Qtde do Produto !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If TravaEntregas(CCur(.Cell(flexcpText, .Row, conCOL_Produto_Saldo)), CLng(.Cell(flexcpText, .Row, conCOL_Produto_IDProduto))) Then Exit Sub
        
        '' ======================================
        '' Verificando Campos em Branco
        If objBLBFunc.FcExisteLinhaVaziaFilho(grdProgEntrega, conCOL_SonProgEntr_QtdProd, conCOL_SonProgEntr_IdProduto, conCOL_SonProgEntr_Action2Do, .Cell(flexcpText, .Row, conCOL_Produto_IDProduto)) = False Then
            Exit Sub
        End If
        If objBLBFunc.FcExisteLinhaVaziaFilho(grdProgEntrega, conCOL_SonProgEntr_DataEntrega, conCOL_SonProgEntr_IdProduto, conCOL_SonProgEntr_Action2Do, .Cell(flexcpText, .Row, conCOL_Produto_IDProduto)) = False Then
            Exit Sub
        End If
        '' ======================================
        
        curPegaSaldo = 0
        If Len(Trim(.Cell(flexcpText, .Row, conCOL_Produto_Saldo))) > 0 Then
           curPegaSaldo = CCur(.Cell(flexcpText, .Row, conCOL_Produto_Saldo))
           For I = 1 To (grdProgEntrega.Rows - 1)
               If .Cell(flexcpText, .Row, conCOL_Produto_IDProduto) = grdProgEntrega.Cell(flexcpText, I, conCOL_SonProgEntr_IdProduto) And _
                  grdProgEntrega.Cell(flexcpText, I, conCOL_SonProgEntr_Action2Do) <> dacEnumUpdateAction_delete Then
                    If Len(Trim(grdProgEntrega.Cell(flexcpText, I, conCOL_SonProgEntr_QtdProd))) > 0 Then
                         curQtdProdInf = curQtdProdInf + CCur(grdProgEntrega.Cell(flexcpText, I, conCOL_SonProgEntr_QtdProd))
                    End If
               End If
           Next I
           curPegaSaldo = (curPegaSaldo - curQtdProdInf)
        End If
        
        grdProgEntrega.AddItem .Cell(flexcpText, .Row, conCOL_Produto_IDProduto) & vbTab & _
                               curPegaSaldo & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               dacEnumUpdateAction_Insert
                               
        
    End With
End Sub


Private Function TravaEntregas(curQtdTotal As Currency, IdProduto As Long) As Boolean

    TravaEntregas = False
    
    Dim curQtdGrd   As Currency
    Dim I           As Integer

    curQtdGrd = 0
    
    With grdProgEntrega
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_SonProgEntr_IdProduto) = IdProduto And _
               Len(Trim(.Cell(flexcpText, I, conCOL_SonProgEntr_QtdProd))) > 0 And _
                (.Cell(flexcpText, I, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Insert Or _
                .Cell(flexcpText, I, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update Or _
                .Cell(flexcpText, I, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Ignore) Then
                curQtdGrd = curQtdGrd + .Cell(flexcpText, I, conCOL_SonProgEntr_QtdProd)
            End If
        Next I
    End With
    
    If curQtdGrd >= curQtdTotal Then TravaEntregas = True
    
End Function


Private Function CalcTotEntregas(curQtdTotal As Currency, IdProduto As Long) As Boolean

    CalcTotEntregas = True
    
    Dim curQtdGrd   As Currency
    Dim I           As Integer

    curQtdGrd = 0
    
    With grdProgEntrega
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_SonProgEntr_IdProduto) = IdProduto And _
               Len(Trim(.Cell(flexcpText, I, conCOL_SonProgEntr_QtdProd))) > 0 And _
                (.Cell(flexcpText, I, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Insert Or _
                .Cell(flexcpText, I, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update Or _
                .Cell(flexcpText, I, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Ignore) Then
                curQtdGrd = curQtdGrd + .Cell(flexcpText, I, conCOL_SonProgEntr_QtdProd)
            End If
        Next I
    End With
    
    If curQtdGrd > curQtdTotal Then
        MsgBox "A soma dos itens não pode ser maior que a qtde total do Produto !!!", vbOKOnly + vbExclamation, "Aviso"
        CalcTotEntregas = False
    End If
    
End Function


Private Sub ExcLinhaGrd(grdGENERICO As VSFlexGrid, lngRol As Long, lngColAction2Do As Long)

    With grdGENERICO
        If .Cell(flexcpText, lngRol, lngColAction2Do) = dacEnumUpdateAction_Ignore Or _
           .Cell(flexcpText, lngRol, lngColAction2Do) = dacEnumUpdateAction_update Then
           .Cell(flexcpText, lngRol, lngColAction2Do) = dacEnumUpdateAction_delete
        ElseIf .Cell(flexcpText, lngRol, lngColAction2Do) = dacEnumUpdateAction_Insert Then
            If (.Rows - 1) = 1 Then .Rows = 1
            If (.Rows - 1) > 1 Then .RemoveItem lngRol
        End If
    End With
    
End Sub



Private Sub PopGrdOrdemFilhas()

    Dim I As Integer
    
    arrORDFATGERADAS = objCADOORDFAB.ORDFATGERADAS
    If IsArray(arrORDFATGERADAS) Then
        With grdProgEntrega
            For I = 1 To UBound(arrORDFATGERADAS)
                
                .AddItem arrORDFATGERADAS(I, 1) & vbTab & _
                         arrORDFATGERADAS(I, 2) & vbTab & _
                         arrORDFATGERADAS(I, 3) & vbTab & _
                         arrORDFATGERADAS(I, 4) & vbTab & _
                         arrORDFATGERADAS(I, 5)
            Next I
        End With
    End If
    

End Sub

Private Sub DestroiObjeto()
    Set objBLBFunc = Nothing
    Set objCADOORDFAB = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub InitGridLogOP()

    With grdLogPed
    
       .Cols = conColumnsIn_SonLogOP
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonLogOP_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       
       .Cell(flexcpData, 0, conCOL_SonLogOP_Data) = ""
       .ColDataType(conCOL_SonLogOP_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonLogOP_Hora) = ""
       .ColDataType(conCOL_SonLogOP_Hora) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogOP_CodUsuario) = ""
       .ColDataType(conCOL_SonLogOP_CodUsuario) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonLogOP_Usuario) = ""
       .ColDataType(conCOL_SonLogOP_Usuario) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogOP_CodAcao) = ""
       .ColDataType(conCOL_SonLogOP_CodAcao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogOP_Acao) = ""
       .ColDataType(conCOL_SonLogOP_Acao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogOP_Tipo) = ""
       .ColDataType(conCOL_SonLogOP_Tipo) = flexDTString
       
       .ColWidth(conCOL_SonLogOP_Data) = 1000
       .ColWidth(conCOL_SonLogOP_Hora) = 1000
       .ColWidth(conCOL_SonLogOP_CodUsuario) = 0
       .ColWidth(conCOL_SonLogOP_Usuario) = 1000
       .ColWidth(conCOL_SonLogOP_CodAcao) = 0
       .ColWidth(conCOL_SonLogOP_Acao) = 3000
       .ColWidth(conCOL_SonLogOP_Tipo) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub Liquida()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Ordem de Fabricação - [ LIQUIDA ORDEM DE FABRICAÇÃO ]"
    
    objBLBFunc.LimpaCampos Me

    Frame2.Enabled = False
    Frame3.Enabled = True
    Frame28.Enabled = True
    txtOBS_MotLiq.Locked = False
    
    objCADOORDFAB.CODORDEM = iCodigo
    
    Call ConfGridProdutos
    Call LimpaCamposLabel
    Call InitGridProg
    Call InitGridLogOP
    
    Call CarregaCampos
    
End Sub


Private Sub VerifOrdFatAberto(strCODPED As String)

    Dim intLinha    As Integer
    Dim arrORDFAT() As String
    
    objCADOORDFAB.ORDFAT = Empty
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED = " & Trim(strCODPED) & vbCrLf
    sSql = sSql & "   And SGI_STATUS = 0"

    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC7.EOF() Then
        Do While Not BREC7.EOF()
            intLinha = intLinha + 1
            ReDim Preserve arrORDFAT(1 To intLinha) As String
            arrORDFAT(intLinha) = Trim(Str(BREC7!SGI_CODORD))
            BREC7.MoveNext
        Loop
        objCADOORDFAB.ORDFAT = arrORDFAT
    End If
    BREC7.Close
    
End Sub

Private Function VerificaPedido(strCODPED As String, strIDPRODUTO As String) As Boolean

    VerificaPedido = False
    
    Dim lngQTDETOTALPEDIDO  As Long
    Dim lngQTDEOPS          As Long
    Dim lngQTDETORULOS      As Long
    Dim intRESP             As Integer
    
    lngQTDETOTALPEDIDO = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Sum(PEDI.SGI_QTDE) As SGI_QTDE" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & " PEDI" & vbCrLf
    sSql = sSql & "  Where" & vbCrLf
    sSql = sSql & "       PEDI.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDI.SGI_CODIGO    = " & Trim(strCODPED) & vbCrLf
    ''sSql = sSql & "   And PEDI.SGI_IDPRODUTO = " & Trim(strIDPRODUTO)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDE) Then lngQTDETOTALPEDIDO = BREC!SGI_QTDE
    End If
    BREC.Close


    '' Pega Quantas OP's Existem para este pedido e Produto
    lngQTDEOPS = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Count(*) As SGI_QTDEOP" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(strIDPRODUTO) & vbCrLf
    sSql = sSql & "   And SGI_CODPED    = " & Trim(strCODPED)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDEOP) Then lngQTDEOPS = BREC!SGI_QTDEOP
    End If
    BREC.Close

    If lngQTDEOPS > 1 Then
       intRESP = MsgBox("ATENÇÃO" & vbCrLf & _
                        "Para este rótulo existe(m) " & Format(lngQTDEOPS, "##00") & " OP's." & vbCrLf & _
                        "Deseja realmente baixar esta OP. ?", vbExclamation + vbYesNo + vbDefaultButton2, "Aviso")
        
       If intRESP = vbNo Then Exit Function
    End If


    '' Verificando se existe Rótulos Diferentes para este pedido
    lngQTDETORULOS = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Count(*) As SGI_QTDEROTULOS" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & " PEDI" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PEDI.SGI_FILIAL    =  " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDI.SGI_CODIGO    =  " & Trim(strCODPED) & vbCrLf
    sSql = sSql & "   And PEDI.SGI_IDPRODUTO <> " & Trim(strIDPRODUTO)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDEROTULOS) Then lngQTDETORULOS = BREC!SGI_QTDEROTULOS
    End If
    BREC.Close
    
    If lngQTDETORULOS <> 0 Then
       intRESP = MsgBox("ATENÇÃO" & vbCrLf & _
                        "Para esta OP existe(m) " & Format(lngQTDETORULOS, "##00") & " Rótulo(s) que diferem do Rótulo que vc esta querendo baixar." & vbCrLf & _
                        "Deseja realmente baixar esta OP. ?", vbExclamation + vbYesNo + vbDefaultButton2, "Aviso")
        
       If intRESP = vbNo Then Exit Function
    End If


    
    VerificaPedido = True
    
    '' Pegando as OPs para poder Baixar

End Function



Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label)

On Error GoTo Err_PegaDescTabelas

    lblLabel.Caption = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(Replace(Replace(strCODIGO, ".", ""), ",", ""))) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCODIGO)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
    Exit Sub
    
Err_PegaDescTabelas:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub


Private Sub PopLogOP()

    Dim I           As Integer
    Dim strNNOMUSU  As String
    Dim arrLOG      As Variant
    
    arrLOG = objCADOORDFAB.LOG
    
    If IsArray(arrLOG) Then
    
        With grdLogPed
        
            For I = 1 To UBound(arrLOG)
                
                strNNOMUSU = objBLBFunc.PegaUsuario(CLng(arrLOG(I, 3)), Linha, FILIAL)
                
                .AddItem arrLOG(I, 1) & vbTab & _
                         arrLOG(I, 2) & vbTab & _
                         arrLOG(I, 3) & vbTab & _
                         objBLBFunc.PegaUsuario(CLng(arrLOG(I, 3)), Linha, FILIAL) & vbTab & _
                         arrLOG(I, 4) & vbTab & _
                         "" & vbTab & _
                         ""
                         
                .Cell(flexcpText, (.Rows - 1), conCOL_SonLogOP_Acao) = DescAcao(.Cell(flexcpText, (.Rows - 1), conCOL_SonLogOP_CodAcao))
                         
            Next I
        
        End With
    
    End If

End Sub


Private Function DescAcao(strACAO As String) As String

    DescAcao = ""
    
    If strACAO = "I" Then DescAcao = "Incluso"
    If strACAO = "A" Then DescAcao = "Alterado"
    If strACAO = "BX" Then DescAcao = "Liquidado Manualmente"
        
End Function

