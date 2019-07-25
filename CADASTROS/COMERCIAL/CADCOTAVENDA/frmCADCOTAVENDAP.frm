VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADCOTAVENDAP 
   Caption         =   "Cadastro de orçamento de vendas"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   12285
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stPedidos 
      Height          =   5535
      Left            =   0
      TabIndex        =   19
      Top             =   1680
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cotações Em Aberto"
      TabPicture(0)   =   "frmCADCOTAVENDAP.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cotações Baixadas Parcial"
      TabPicture(1)   =   "frmCADCOTAVENDAP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cotações Baixadas Total"
      TabPicture(2)   =   "frmCADCOTAVENDAP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Não Atendidas"
      TabPicture(3)   =   "frmCADCOTAVENDAP.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraNAOATEND"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraNAOATEND 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   12015
         Begin VSFlex8LCtl.VSFlexGrid grdNAOATEND 
            Height          =   4695
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   11775
            _cx             =   20770
            _cy             =   8281
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
      Begin VB.Frame Frame5 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   24
         Top             =   360
         Width           =   12015
         Begin VSFlex8LCtl.VSFlexGrid flxCOTACAOBAIXTOTAL 
            Height          =   4815
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   11775
            _cx             =   20770
            _cy             =   8493
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
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
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
         Height          =   5055
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   12015
         Begin VSFlex8LCtl.VSFlexGrid flxCOTACAOPARCIAL 
            Height          =   4815
            Left            =   120
            TabIndex        =   23
            Top             =   195
            Width           =   11775
            _cx             =   20770
            _cy             =   8493
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
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
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
         Height          =   5055
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   12015
         Begin VSFlex8LCtl.VSFlexGrid flxCotaVend 
            Height          =   4815
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   11775
            _cx             =   20770
            _cy             =   8493
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
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
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
      Width           =   12255
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
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "txtCampos"
         Top             =   200
         Width           =   8175
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
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   12255
      Begin VB.Timer Timer1 
         Interval        =   2500
         Left            =   4080
         Top             =   240
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
         Left            =   3000
         Picture         =   "frmCADCOTAVENDAP.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
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
         Picture         =   "frmCADCOTAVENDAP.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   240
         Width           =   735
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
         Left            =   840
         Picture         =   "frmCADCOTAVENDAP.frx":06A4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   240
         Width           =   735
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
         Left            =   1560
         Picture         =   "frmCADCOTAVENDAP.frx":0BD6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Altera Empresa "
         Top             =   240
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
         Left            =   2280
         Picture         =   "frmCADCOTAVENDAP.frx":0CD8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
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
         Height          =   735
         Left            =   10320
         Picture         =   "frmCADCOTAVENDAP.frx":0DDA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   855
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
         Height          =   735
         Left            =   11280
         Picture         =   "frmCADCOTAVENDAP.frx":130C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Parcial"
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
      Left            =   2040
      TabIndex        =   18
      Top             =   7320
      Width           =   600
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Pedido Total"
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
      Left            =   3600
      TabIndex        =   16
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Aberto"
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
      Left            =   600
      TabIndex        =   14
      Top             =   7320
      Width           =   570
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7320
      Width           =   375
   End
End
Attribute VB_Name = "frmCADCOTAVENDAP"
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
Dim objFuncoes      As Object
Dim objCADCOTAVENDA As Object
Dim objRel          As Object
Dim iCodigo         As Long
Dim lngCodVendAtua  As Long

Const conCOL_Codigo     As Integer = 0
Const conCOL_Data       As Integer = 1
Const conCOL_Cliente    As Integer = 2
Const conCOL_Tipo       As Integer = 3
Const conCOL_TipOrca    As Integer = 4
Const conCOL_Pedido     As Integer = 5
Const conCOL_Status     As Integer = 6


Const conCOL_SonNaoAtend_Codigo                    As Integer = 0
Const conCOL_SonNaoAtend_Data                      As Integer = 1
Const conCOL_SonNaoAtend_Cliente                   As Integer = 2
Const conCOL_SonNaoAtend_Tipo                      As Integer = 3
Const conCOL_SonNaoAtend_TipoOrca                  As Integer = 4
Const conCOL_SonNaoAtend_Pedido                    As Integer = 5
Const conCOL_SonNaoAtend_Status                    As Integer = 6
Const conCOL_SonNaoAtend_FormatString              As String = "=Código|Data|Cliente|Tipo|Tipo de Orçamento|Pedido|Status"
Const conColumnsIn_SonNaoAtend                     As Integer = 7
Private Sub cmdAltera_Click()
  
  If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
  
  If (flxCOTACAOBAIXTOTAL.Rows - 1) > 1 And flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, 5) = "Sim" And stPedidos.Tab = 2 Then
      MsgBox "Não pose ser alterada já existe pedidos !!!", vbOKOnly + vbExclamation, "Aviso"
      Exit Sub
  ElseIf (flxCOTACAOPARCIAL.Rows - 1) > 1 And flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, 5) = "Sim" And stPedidos.Tab = 1 Then
      MsgBox "Não pose ser alterada já existe pedidos !!!", vbOKOnly + vbExclamation, "Aviso"
      Exit Sub
  End If
  
  Operacao "A"
  
End Sub

Private Sub cmdCanFiltro_Click()
    AbilitaCampos
    ConfGrid
    ConfGridCotacaoParcial
    ConfGridCotacaoTotal
    ConfGrdNaoAtend
    PreencheGrid
    PreencheGridBaixadoParcial
    PreencheGridBaixadoTotal
End Sub

Private Sub cmdExclui_Click()
  
  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  If (flxCOTACAOBAIXTOTAL.Rows - 1) > 1 And flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, 5) = "Sim" And stPedidos.Tab = 2 Then
     MsgBox "Não pose ser excluso já existe pedidos !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  ElseIf (flxCOTACAOPARCIAL.Rows - 1) > 1 And flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, 5) = "Sim" And stPedidos.Tab = 1 Then
     MsgBox "Não pose ser excluso já existe pedidos !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  End If

  Dim iResp  As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADCOTAVENDA.GRAVA("E") = False Then Exit Sub
  If objCADCOTAVENDA.Atualiza("E", Str(objCADCOTAVENDA.CODIGO), FILIAL, "frmCADCOTAVENDA") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  AbilitaCampos
  Atualiza_Grid

End Sub

Private Sub cmdImpressao_Click()
       Call ImpCota
End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxCotaVend.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
   Set objFuncoes = Nothing
   Set objCADCOTAVENDA = Nothing
   Unload Me
End Sub

Private Sub flxCOTACAOBAIXTOTAL_CellChanged(ByVal Row As Long, ByVal Col As Long)
        flxCOTACAOBAIXTOTAL.Cell(flexcpForeColor, Row, Col) = &H8000&
End Sub

Private Sub flxCOTACAOBAIXTOTAL_Click()
    If (flxCOTACAOBAIXTOTAL.Rows - 1) > 0 And (flxCOTACAOBAIXTOTAL.Row) > 0 Then objCADCOTAVENDA.CODIGO = CLng(Trim(Replace(flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, conCOL_Codigo), "/", "")))
End Sub

Private Sub flxCOTACAOBAIXTOTAL_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCOTACAOBAIXTOTAL.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCOTACAOPARCIAL_CellChanged(ByVal Row As Long, ByVal Col As Long)
        flxCOTACAOPARCIAL.Cell(flexcpForeColor, Row, Col) = &H8080&
End Sub

Private Sub flxCOTACAOPARCIAL_Click()
    If (flxCOTACAOPARCIAL.Rows - 1) > 0 And (flxCOTACAOPARCIAL.Row) > 0 Then objCADCOTAVENDA.CODIGO = CLng(Trim(Replace(flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, conCOL_Codigo), "/", "")))
End Sub

Private Sub flxCOTACAOPARCIAL_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCOTACAOPARCIAL.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCotaVend_CellChanged(ByVal Row As Long, ByVal Col As Long)
        flxCotaVend.Cell(flexcpForeColor, Row, Col) = &HFF&
End Sub

Private Sub flxCotaVend_Click()
    If (flxCotaVend.Rows - 1) > 0 And (flxCotaVend.Row) > 0 Then objCADCOTAVENDA.CODIGO = CLng(Trim(Replace(flxCotaVend.Cell(flexcpText, flxCotaVend.Row, conCOL_Codigo), "/", "")))
End Sub

Private Sub flxCotaVend_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCotaVend.Rows > 1 Then Operacao "C"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADCOTAVENDA = CreateObject("CADCOTAVENDA.clsCADCOTAVENDA")
    Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    objCADCOTAVENDA.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADCOTAVENDAP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    txtCampos.Visible = True
    
    lngCodVendAtua = PegaCodVendedor(strUsuario)
    
    AbilitaCampos
    ConfGrid
    ConfGridCotacaoParcial
    ConfGridCotacaoTotal
    ConfGrdNaoAtend
    PreencheGrid
    PreencheGridBaixadoParcial
    PreencheGridBaixadoTotal
    PreencheGridNaoAtend
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Data"
    cboFiltro.AddItem "Cliente"
    cboFiltro.AddItem "Tipo de Cotação"
    cboFiltro.AddItem "Status"
    
    cboFiltro.ListIndex = 0
    
    stPedidos.Tab = 0
    If flxCotaVend.Enabled = True Then
       If flxCotaVend.Rows - 1 > 0 Then flxCotaVend.Row = 1
    End If
    

    Timer1.Interval = 2500
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
End Sub

Private Sub AbilitaCampos()
    If objCADCOTAVENDA.Pesq_CadOrcVenda = False Then
       cmdAltera.Enabled = False
       cmdExclui.Enabled = False
       Frame3.Enabled = False
       Frame4.Enabled = False
       Frame5.Enabled = False
       fraNAOATEND.Enabled = False
    Else
       cmdAltera.Enabled = True
       cmdExclui.Enabled = True
       Frame3.Enabled = True
       Frame4.Enabled = True
       Frame5.Enabled = True
       fraNAOATEND.Enabled = True
    End If
End Sub

Private Sub ConfGrid()
   
    
    flxCotaVend.Rows = 1
    flxCotaVend.Cols = 7
    flxCotaVend.FixedCols = 0
    flxCotaVend.AllowBigSelection = False
    
    flxCotaVend.Editable = flexEDNone
    
    flxCotaVend.FormatString = "Código|Data|Cliente|Tipo|TIPORCA|Pedido|Status"
    
    flxCotaVend.ColWidth(conCOL_Codigo) = 1000
    flxCotaVend.ColWidth(conCOL_Data) = 1000
    flxCotaVend.ColWidth(conCOL_Cliente) = 5000
    flxCotaVend.ColWidth(conCOL_Tipo) = 2000
    flxCotaVend.ColWidth(conCOL_TipOrca) = 0
    flxCotaVend.ColWidth(conCOL_Pedido) = 600
    flxCotaVend.ColWidth(conCOL_Status) = 700
    
End Sub

Private Sub PreencheGrid()

    Dim strSTATUS       As String
    Dim I               As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       COTA.* " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If lngCodVendAtua > 0 Then
        sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua
    End If
    
    sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
    sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
    sSql = sSql & " Order by COTA.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       If BREC!SGI_STATUS = "A" Then strSTATUS = "Não"
       
       '' Vendo se existe pedido
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strSTATUS = "Sim"
       BREC2.Close
       '' ---------------------------
       
       flxCotaVend.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                           Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                           BREC!SGI_RAZAOSOC & vbTab & _
                           BREC!SGI_DESCRICAO & vbTab & _
                           BREC!SGI_CODTIPORC & vbTab & _
                           strSTATUS & vbTab & _
                           IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
                           
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        
        For I = 1 To (flxCotaVend.Rows - 1)
             
            If CLng(Trim(Replace(flxCotaVend.TextMatrix(I, 1), "/", ""))) = iCodigo Then
               flxCotaVend.Row = I
               flxCotaVend.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
    
  If stPedidos.Tab = 0 Then
     ''If flxCotaVend.Row = 0 Then Exit Sub
     If flxCotaVend.Rows > 1 Then iCodigo = CLng(Trim(Replace(flxCotaVend.Cell(flexcpText, flxCotaVend.Row, conCOL_Codigo), "/", "")))
  ElseIf stPedidos.Tab = 1 Then
     ''If flxCOTACAOPARCIAL.Row = 0 Then Exit Sub
     If flxCOTACAOPARCIAL.Rows > 1 Then iCodigo = CLng(Trim(Replace(flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, conCOL_Codigo), "/", "")))
  ElseIf stPedidos.Tab = 2 Then
     ''If flxCOTACAOBAIXTOTAL.Row = 0 Then Exit Sub
     If flxCOTACAOBAIXTOTAL.Rows > 1 Then iCodigo = CLng(Trim(Replace(flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, conCOL_Codigo), "/", "")))
  ElseIf stPedidos.Tab = 3 Then
     ''If grdNAOATEND.Row = 0 Then Exit Sub
     If grdNAOATEND.Rows > 1 Then iCodigo = CLng(Trim(Replace(grdNAOATEND.Cell(flexcpText, grdNAOATEND.Row, conCOL_SonNaoAtend_Codigo), "/", "")))
  End If
  
  frmCADCOTAVENDA.cCaminho = cCaminho
  frmCADCOTAVENDA.Linha = Linha
  frmCADCOTAVENDA.iCodigo = iCodigo
  frmCADCOTAVENDA.cTipOper = strOperacao
  frmCADCOTAVENDA.FILIAL = FILIAL
  frmCADCOTAVENDA.strAcesso = strAcesso
  frmCADCOTAVENDA.strMODPAI = Me.Name
  frmCADCOTAVENDA.strUsuario = strUsuario
  frmCADCOTAVENDA.lngCodUsuario = lngCodVendAtua
  frmCADCOTAVENDA.Show vbModal
  
  AbilitaCampos
  Atualiza_Grid

End Sub


Private Sub ImpCota()

    If stPedidos.Tab = 0 Then
        If (flxCotaVend.Rows - 1) = 1 Then
           MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
    
        If flxCotaVend.Cell(flexcpText, flxCotaVend.Row, 4) = 1 Then ImpCotaVend     '' Orçamento de Venda
        If flxCotaVend.Cell(flexcpText, flxCotaVend.Row, 4) = 2 Then ImpCotaAssTec   '' Orçamento de Ass.Técnica
        If flxCotaVend.Cell(flexcpText, flxCotaVend.Row, 4) = 3 Then ImpCotaManut    '' Orçamento de Manutenção
        If flxCotaVend.Cell(flexcpText, flxCotaVend.Row, 4) = 4 Then ImpCotaRevenda  '' Orçamento de Revenda
        If flxCotaVend.Cell(flexcpText, flxCotaVend.Row, 4) = 5 Then ImpCotaAssTec   '' Orçamento de Calibração
    
    ElseIf stPedidos.Tab = 1 Then
        If (flxCOTACAOPARCIAL.Rows - 1) = 1 Then
           MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
    
        If flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, 4) = 1 Then ImpCotaVend     '' Orçamento de Venda
        If flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, 4) = 2 Then ImpCotaAssTec   '' Orçamento de Ass.Técnica
        If flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, 4) = 3 Then ImpCotaManut    '' Orçamento de Manutenção
        If flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, 4) = 4 Then ImpCotaRevenda  '' Orçamento de Revenda
        If flxCOTACAOPARCIAL.Cell(flexcpText, flxCOTACAOPARCIAL.Row, 4) = 5 Then ImpCotaAssTec   '' Orçamento de Calibração
    
    ElseIf stPedidos.Tab = 2 Then
        If (flxCOTACAOBAIXTOTAL.Rows - 1) = 1 Then
           MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
    
        If flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, 4) = 1 Then ImpCotaVend     '' Orçamento de Venda
        If flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, 4) = 2 Then ImpCotaAssTec   '' Orçamento de Ass.Técnica
        If flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, 4) = 3 Then ImpCotaManut    '' Orçamento de Manutenção
        If flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, 4) = 4 Then ImpCotaRevenda  '' Orçamento de Revenda
        If flxCOTACAOBAIXTOTAL.Cell(flexcpText, flxCOTACAOBAIXTOTAL.Row, 4) = 5 Then ImpCotaAssTec   '' Orçamento de Calibração
    ElseIf stPedidos.Tab = 3 Then
        If (grdNAOATEND.Rows - 1) = 1 Then
           MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
    
        If grdNAOATEND.Cell(flexcpText, grdNAOATEND.Row, 4) = 1 Then ImpCotaVend     '' Orçamento de Venda
        If grdNAOATEND.Cell(flexcpText, grdNAOATEND.Row, 4) = 2 Then ImpCotaAssTec   '' Orçamento de Ass.Técnica
        If grdNAOATEND.Cell(flexcpText, grdNAOATEND.Row, 4) = 3 Then ImpCotaManut    '' Orçamento de Manutenção
        If grdNAOATEND.Cell(flexcpText, grdNAOATEND.Row, 4) = 4 Then ImpCotaRevenda  '' Orçamento de Revenda
        If grdNAOATEND.Cell(flexcpText, grdNAOATEND.Row, 4) = 5 Then ImpCotaAssTec   '' Orçamento de Calibração
    
    End If
    

End Sub

Private Sub ImpCotaVend()

On Error GoTo Imp_ImpCotaVend

    sSql = "Select "
    sSql = sSql & "       SGI_CADCOTAVENDH.SGI_CODIGO     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_CODPROD    "
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC     "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_CONTATO    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DEPTO      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_EMAIL      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DATACOTA   "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_DESCRICAO   "
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO   "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO    "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_ESPTEC     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_QTDE       "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_VLUNIT     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_PRCDESC    "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_PRCIPI     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_VLTOT      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_PRZENTR    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VALPROP    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLIPI      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLICMS     "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_FRETETOT   "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLDESCTO   "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_OUTRASTOT  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_TEL        "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_EMAIL       "
    sSql = sSql & "      ,SGI_CADESPORCA.SGI_DESCRICAO    "
     
    sSql = sSql & "  From "
    
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE  "
    sSql = sSql & "      ,SGI_CADCONDPGTO SGI_CADCONDPGTO "
    sSql = sSql & "      ,SGI_CADCOTAVENDH SGI_CADCOTAVENDH "
    sSql = sSql & "      ,SGI_CADCOTAVENDI SGI_CADCOTAVENDI "
    sSql = sSql & "      ,SGI_CADESPORCA SGI_CADESPORCA "
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & "      ,SGI_CADVENDEDOR SGI_CADVENDEDOR "
    ''sSql = sSql & "      ,SGI_FILIAL SGI_FILIAL "
    
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADCLIENTE.SGI_FILIAL      = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCLIENTE.SGI_CODIGO      = SGI_CADCOTAVENDH.SGI_CODCLI  "
    
    sSql = sSql & "   And SGI_CADCONDPGTO.SGI_FILIAL     = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCONDPGTO.SGI_CODIGO     = SGI_CADCOTAVENDH.SGI_CODCONDPGT "
    
    sSql = sSql & "   And SGI_CADVENDEDOR.SGI_FILIAL     = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADVENDEDOR.SGI_CODIGO     = SGI_CADCOTAVENDH.SGI_CODVEND "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL    = SGI_CADCOTAVENDI.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODIGO    = SGI_CADCOTAVENDI.SGI_CODIGO  "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL    = SGI_CADESPORCA.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODTIPORC = SGI_CADESPORCA.SGI_CODIGO  "
    
    ''sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL    = SGI_FILIAL.SGI_FILIAL  "
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL      = SGI_CADCOTAVENDI.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODIGO      = SGI_CADCOTAVENDI.SGI_CODPROD "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL    = " & FILIAL
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODIGO    = " & objCADCOTAVENDA.CODIGO

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Esta cotação não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    objRel.REL FILIAL, sSql, strCamRelNovo & cCamRelCotacaoVendas & "RELCOTAVENDVDA.rpt", Linha, 1, "", "Cotação de vendas (Vendas)"
    
    Exit Sub
    
Imp_ImpCotaVend:

    MsgBox "Erro nº : " & Err.Number & vbCrLf & " Erro Descrição : " & Err.Description, vbOKOnly + vbCritical, "Erro"
    
End Sub

Private Sub ImpCotaAssTec()

On Error GoTo Imp_ImpCotaVend

    
    sSql = "Select "
    
    sSql = sSql & "       SGI_CADCOTADESP.SGI_CODIGO  "
    sSql = sSql & "      ,SGI_CADCOTADESP.SGI_CODDESP  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_CODIGO    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DATACOTA  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_TEL  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_CONTATO  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DEPTO  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_EMAIL  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_CODCLI "
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC "
    sSql = sSql & "      ,SGI_CADTIPDESP.SGI_DESCRICAO "
    sSql = sSql & "      ,SGI_CADCOTADESP.SGI_QTDE  "
    sSql = sSql & "      ,SGI_CADCOTADESP.SGI_VALOR  "
    sSql = sSql & "      ,SGI_CADCOTADESP.SGI_VALORTOTAL  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_CODCONDPGT "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DTCALIBRA "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VALPROP "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLTOT "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_CODVEND "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_DESCRICAO "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_IMAGEM "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_EMAIL "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_SKYPE "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_MSN "
    sSql = sSql & "      ,SGI_CADESPORCA.SGI_DESCRICAO "
    
    sSql = sSql & "  From "
    sSql = sSql & "        SGI_CADCLIENTE SGI_CADCLIENTE "
    sSql = sSql & "       ,SGI_CADCONDPGTO SGI_CADCONDPGTO "
    sSql = sSql & "       ,SGI_CADCOTADESP SGI_CADCOTADESP "
    sSql = sSql & "       ,SGI_CADCOTAVENDH SGI_CADCOTAVENDH "
    sSql = sSql & "       ,SGI_CADESPORCA SGI_CADESPORCA "
    sSql = sSql & "       ,SGI_CADTIPDESP SGI_CADTIPDESP "
    sSql = sSql & "       ,SGI_CADVENDEDOR SGI_CADVENDEDOR "
    ''sSql = sSql & "       ,SGI_FILIAL SGI_FILIAL "
    
    sSql = sSql & " Where "
    
    sSql = sSql & "       SGI_CADCOTADESP.SGI_FILIAL =  " & FILIAL
    sSql = sSql & "   And SGI_CADCOTADESP.SGI_CODIGO =  " & objCADCOTAVENDA.CODIGO
    
    sSql = sSql & "   And SGI_CADCOTADESP.SGI_FILIAL = SGI_CADCOTAVENDH.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADCOTADESP.SGI_CODIGO = SGI_CADCOTAVENDH.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_CADCOTADESP.SGI_FILIAL  = SGI_CADTIPDESP.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADCOTADESP.SGI_CODDESP = SGI_CADTIPDESP.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL = SGI_CADCONDPGTO.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODCONDPGT = SGI_CADCONDPGTO.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL = SGI_CADVENDEDOR.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODVEND = SGI_CADVENDEDOR.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL = SGI_CADESPORCA.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODTIPORC = SGI_CADESPORCA.SGI_CODIGO "
    
    ''sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL = SGI_FILIAL.SGI_FILIAL "

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Esta cotação não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
   
    objRel.REL FILIAL, sSql, strCamRelNovo & strCamRelNovo & cCamRelCotacaoVendas & "RELCOTAVENDAST3.rpt", Linha, 1, "", "Cotação de vendas (Assitência técnica)", False

    Exit Sub
    
Imp_ImpCotaVend:

    MsgBox "Erro nº : " & Err.Number & vbCrLf & " Erro Descrição : " & Err.Description, vbOKOnly + vbCritical, "Erro"

End Sub

Private Sub ImpCotaManut()

On Error GoTo Imp_ImpCotaVend

    sSql = "Select "
    sSql = sSql & "       SGI_CADCOTAVENDH.SGI_CODIGO     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_CODPROD    "
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC     "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_CONTATO    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DEPTO      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_EMAIL      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DATACOTA   "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_DESCRICAO   "
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO   "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO    "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_ESPTEC     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_QTDE       "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_VLUNIT     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_PRCDESC    "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_PRCIPI     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_VLTOT      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_PRZENTR    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VALPROP    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLIPI      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLICMS     "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_FRETETOT   "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLDESCTO   "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_OUTRASTOT  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_TEL        "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_EMAIL       "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_SKYPE       "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADESPORCA SGI_CADESPORCA "
    sSql = sSql & "      ,SGI_CADCLIENTE SGI_CADCLIENTE "
    sSql = sSql & "      ,SGI_CADCONDPGTO SGI_CADCONDPGTO "
    sSql = sSql & "      ,SGI_CADCOTAVENDH SGI_CADCOTAVENDH "
    sSql = sSql & "      ,SGI_CADCOTAVENDI SGI_CADCOTAVENDI "
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & "      ,SGI_CADVENDEDOR SGI_CADVENDEDOR "
 
    sSql = sSql & " Where "
    
    sSql = sSql & "       SGI_CADCLIENTE.SGI_FILIAL    = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCLIENTE.SGI_CODIGO    = SGI_CADCOTAVENDH.SGI_CODCLI  "
  
    sSql = sSql & "   And SGI_CADCONDPGTO.SGI_FILIAL   = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCONDPGTO.SGI_CODIGO   = SGI_CADCOTAVENDH.SGI_CODCONDPGT "
    
    sSql = sSql & "   And SGI_CADVENDEDOR.SGI_FILIAL   = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADVENDEDOR.SGI_CODIGO   = SGI_CADCOTAVENDH.SGI_CODVEND "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL  = SGI_CADCOTAVENDI.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODIGO  = SGI_CADCOTAVENDI.SGI_CODIGO  "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL     = SGI_CADESPORCA.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODTIPORC  = SGI_CADESPORCA.SGI_CODIGO  "
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODIGO       = SGI_CADCOTAVENDI.SGI_CODPROD "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL  = " & FILIAL
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODIGO  = " & objCADCOTAVENDA.CODIGO
    
    

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Esta cotação não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
  
    objRel.REL FILIAL, sSql, strCamRelNovo & cCamRelCotacaoVendas & "RELCOTAVENDMAN.rpt", Linha, 1, "", "Cotação de vendas (Manutenção)"
    
    Exit Sub
    
Imp_ImpCotaVend:

    MsgBox "Erro nº : " & Err.Number & vbCrLf & " Erro Descrição : " & Err.Description, vbOKOnly + vbCritical, "Erro"

End Sub

Private Sub ImpCotaRevenda()

On Error GoTo Imp_ImpCotaVend
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADCOTAVENDH.SGI_CODIGO     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_CODPROD    "
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC     "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_CONTATO    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DEPTO      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_EMAIL      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_DATACOTA   "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_DESCRICAO   "
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO   "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO    "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_ESPTEC     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_QTDE       "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_VLUNIT     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_PRCDESC    "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_PRCIPI     "
    sSql = sSql & "      ,SGI_CADCOTAVENDI.SGI_VLTOT      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_PRZENTR    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VALPROP    "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLIPI      "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLICMS     "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_FRETETOT   "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_VLDESCTO   "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_OUTRASTOT  "
    sSql = sSql & "      ,SGI_CADCOTAVENDH.SGI_TEL        "
    sSql = sSql & "      ,SGI_CADVENDEDOR.SGI_EMAIL       "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_PROCEDENCIA  "
    sSql = sSql & "      ,SGI_CADESPORCA.SGI_DESCRICAO    "
    
    sSql = sSql & "  From "
    
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE "
    sSql = sSql & "      ,SGI_CADCONDPGTO SGI_CADCONDPGTO "
    sSql = sSql & "      ,SGI_CADCOTAVENDH SGI_CADCOTAVENDH "
    sSql = sSql & "      ,SGI_CADCOTAVENDI SGI_CADCOTAVENDI "
    sSql = sSql & "      ,SGI_CADESPORCA SGI_CADESPORCA "
    sSql = sSql & "      ,SGI_CADVENDEDOR SGI_CADVENDEDOR "
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO "
    ''sSql = sSql & "      ,SGI_FILIAL SGI_FILIAL "
    
    sSql = sSql & " Where "
    
    sSql = sSql & "       SGI_CADPRODUTO.SGI_FILIAL       = SGI_CADCOTAVENDI.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCLIENTE.SGI_FILIAL       = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCLIENTE.SGI_CODIGO       = SGI_CADCOTAVENDH.SGI_CODCLI  "
    
    sSql = sSql & "   And SGI_CADCONDPGTO.SGI_FILIAL      = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCONDPGTO.SGI_CODIGO      = SGI_CADCOTAVENDH.SGI_CODCONDPGT "
    
    sSql = sSql & "   And SGI_CADVENDEDOR.SGI_FILIAL      = SGI_CADCOTAVENDH.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADVENDEDOR.SGI_CODIGO      = SGI_CADCOTAVENDH.SGI_CODVEND "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL     = SGI_CADCOTAVENDI.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODIGO     = SGI_CADCOTAVENDI.SGI_CODIGO  "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL     = SGI_CADESPORCA.SGI_FILIAL  "
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODTIPORC  = SGI_CADESPORCA.SGI_CODIGO  "
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODIGO       = SGI_CADCOTAVENDI.SGI_CODPROD "
    ''sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL     =  SGI_FILIAL.SGI_FILIAL "
    
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_FILIAL     = " & FILIAL
    sSql = sSql & "   And SGI_CADCOTAVENDH.SGI_CODIGO     = " & objCADCOTAVENDA.CODIGO
    

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Esta cotação não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
  
    objRel.REL FILIAL, sSql, strCamRelNovo & cCamRelCotacaoVendas & "RELCOTAVENDREV.rpt", Linha, 1, "", "Cotação de vendas (Revendas)"
    
    Exit Sub
    
Imp_ImpCotaVend:

    MsgBox "Erro nº : " & Err.Number & vbCrLf & " Erro Descrição : " & Err.Description, vbOKOnly + vbCritical, "Erro"


End Sub


Private Sub grdNAOATEND_CellChanged(ByVal Row As Long, ByVal Col As Long)
    grdNAOATEND.Cell(flexcpForeColor, Row, Col) = &HFF&
End Sub

Private Sub grdNAOATEND_Click()
    If (grdNAOATEND.Rows - 1) > 0 And (grdNAOATEND.Row) > 0 Then objCADCOTAVENDA.CODIGO = CLng(Trim(Replace(grdNAOATEND.Cell(flexcpText, grdNAOATEND.Row, conCOL_SonNaoAtend_Codigo), "/", "")))
End Sub

Private Sub grdNAOATEND_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If grdNAOATEND.Rows > 1 Then Operacao "C"
End Sub

Private Sub grdNAOATEND_RowColChange()
    If (grdNAOATEND.Rows - 1) > 0 And (grdNAOATEND.Row) > 0 Then objCADCOTAVENDA.CODIGO = CLng(Trim(Replace(grdNAOATEND.Cell(flexcpText, grdNAOATEND.Row, conCOL_SonNaoAtend_Codigo), "/", "")))
End Sub


Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub

Private Sub Atualiza_Grid()
    
     Dim I          As Integer
     Dim j          As Integer
     Dim bolAchou   As Boolean
     Dim strSTATUS  As String
     Dim intSTATUS  As Integer
      
     bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADCOTAVENDA'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        If stPedidos.Tab = 0 Then
            For I = 1 To (flxCotaVend.Rows - 1)
                If Trim(BREC!SGI_ACAO) = "E" Then
                   If Trim(Replace(flxCotaVend.Cell(flexcpText, I, conCOL_Codigo), "/", "")) = Trim(BREC!SGI_CODIGO) Then
                      If flxCotaVend.Rows = 2 Then flxCotaVend.Rows = 1
                      If flxCotaVend.Rows > 2 Then flxCotaVend.RemoveItem I
                      Exit For
                   End If
                ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
                   If Trim(BREC!SGI_CODIGO) = Trim(Replace(flxCotaVend.Cell(flexcpText, I, conCOL_Codigo), "/", "")) Then
                      bolAchou = True
                      Exit For
                   End If
                End If
            Next I
        ElseIf stPedidos.Tab = 1 Then
            For I = 1 To (flxCOTACAOPARCIAL.Rows - 1)
                If Trim(BREC!SGI_ACAO) = "E" Then
                   If Trim(Replace(flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_Codigo), "/", "")) = Trim(BREC!SGI_CODIGO) Then
                      If flxCOTACAOPARCIAL.Rows = 2 Then flxCOTACAOPARCIAL.Rows = 1
                      If flxCOTACAOPARCIAL.Rows > 2 Then flxCOTACAOPARCIAL.RemoveItem I
                      Exit For
                   End If
                ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
                   If Trim(BREC!SGI_CODIGO) = Trim(Replace(flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_Codigo), "/", "")) Then
                      bolAchou = True
                      Exit For
                   End If
                End If
            Next I
        ElseIf stPedidos.Tab = 2 Then
            For I = 1 To (flxCOTACAOBAIXTOTAL.Rows - 1)
                If Trim(BREC!SGI_ACAO) = "E" Then
                   If Trim(Replace(flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_Codigo), "/", "")) = Trim(BREC!SGI_CODIGO) Then
                      If flxCOTACAOBAIXTOTAL.Rows = 2 Then flxCOTACAOBAIXTOTAL.Rows = 1
                      If flxCOTACAOBAIXTOTAL.Rows > 2 Then flxCOTACAOBAIXTOTAL.RemoveItem I
                      Exit For
                   End If
                ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
                   If Trim(BREC!SGI_CODIGO) = Trim(Replace(flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_Codigo), "/", "")) Then
                      bolAchou = True
                      Exit For
                   End If
                End If
            Next I
        ElseIf stPedidos.Tab = 3 Then
            For I = 1 To (grdNAOATEND.Rows - 1)
                If Trim(BREC!SGI_ACAO) = "E" Then
                   If Trim(Replace(grdNAOATEND.Cell(flexcpText, I, conCOL_Codigo), "/", "")) = Trim(BREC!SGI_CODIGO) Then
                      If grdNAOATEND.Rows = 2 Then grdNAOATEND.Rows = 1
                      If grdNAOATEND.Rows > 2 Then grdNAOATEND.RemoveItem I
                      Exit For
                   End If
                ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
                   If Trim(BREC!SGI_CODIGO) = Trim(Replace(grdNAOATEND.Cell(flexcpText, I, conCOL_Codigo), "/", "")) Then
                      bolAchou = True
                      Exit For
                   End If
                End If
            Next I
        End If
        
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       COTA.* " & vbCrLf
            sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
            sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
            sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
            sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "   And COTA.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
            
            If lngCodVendAtua > 0 Then
               sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
            End If
            
            sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
            sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
            If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
            If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
            If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
            If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
            sSql = sSql & " Order by COTA.SGI_CODIGO "
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF Then
            
                If BREC2!SGI_STATUS = "A" Then strSTATUS = "Não"
                If BREC2!SGI_STATUS = "N" Then intSTATUS = 0
                
                '' Vendo se existe pedido
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODCOTA = " & BREC2!SGI_CODIGO
                
                BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC3.EOF Then strSTATUS = "Sim"
                BREC3.Close
                '' ---------------------------
                
                If stPedidos.Tab = 0 Then
                    flxCotaVend.AddItem Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4) & vbTab & _
                                        Format(BREC2!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                        BREC2!SGI_RAZAOSOC & vbTab & _
                                        BREC2!SGI_DESCRICAO & vbTab & _
                                        BREC2!SGI_CODTIPORC & vbTab & _
                                        strSTATUS & vbTab & _
                                        IIf(Trim(BREC2!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC2!SGI_STATUS) = "B", "Baixado", "Parcial"))
                ElseIf stPedidos.Tab = 1 Then
                    flxCOTACAOPARCIAL.AddItem Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4) & vbTab & _
                                              Format(BREC2!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                              BREC2!SGI_RAZAOSOC & vbTab & _
                                              BREC2!SGI_DESCRICAO & vbTab & _
                                              BREC2!SGI_CODTIPORC & vbTab & _
                                              strSTATUS & vbTab & _
                                              IIf(Trim(BREC2!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC2!SGI_STATUS) = "B", "Baixado", "Parcial"))
                ElseIf stPedidos.Tab = 2 Then
                    flxCOTACAOBAIXTOTAL.AddItem Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4) & vbTab & _
                                                Format(BREC2!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                                BREC2!SGI_RAZAOSOC & vbTab & _
                                                BREC2!SGI_DESCRICAO & vbTab & _
                                                BREC2!SGI_CODTIPORC & vbTab & _
                                                strSTATUS & vbTab & _
                                                IIf(Trim(BREC2!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC2!SGI_STATUS) = "B", "Baixado", "Parcial"))
                ElseIf stPedidos.Tab = 3 Then
                    grdNAOATEND.AddItem Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4) & vbTab & _
                                                Format(BREC2!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                                BREC2!SGI_RAZAOSOC & vbTab & _
                                                BREC2!SGI_DESCRICAO & vbTab & _
                                                BREC2!SGI_CODTIPORC & vbTab & _
                                                intSTATUS & vbTab & _
                                                IIf(Trim(BREC2!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC2!SGI_STATUS) = "B", "Baixado", "Parcial"))
                End If
            End If
            BREC2.Close
            
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "       COTA.* " & vbCrLf
            sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
            sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
            sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
            sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "   And COTA.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
            
            If lngCodVendAtua > 0 Then
               sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
            End If
            
            sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
            sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
            If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
            If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
            If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
            If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
            sSql = sSql & " Order by COTA.SGI_CODIGO "
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF Then
            
                If BREC2!SGI_STATUS = "A" Then strSTATUS = "Não"
                If BREC2!SGI_STATUS = "N" Then intSTATUS = 0
                
                '' Vendo se existe pedido
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODCOTA = " & BREC2!SGI_CODIGO
                
                BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC3.EOF Then strSTATUS = "Sim"
                BREC3.Close
                '' ---------------------------
                
                If stPedidos.Tab = 0 Then
                   flxCotaVend.Cell(flexcpText, I, conCOL_Codigo) = Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4)
                   flxCotaVend.Cell(flexcpText, I, conCOL_Data) = Format(BREC2!SGI_DATACOTA, "DD/MM/YYYY")
                   flxCotaVend.Cell(flexcpText, I, conCOL_Cliente) = BREC2!SGI_RAZAOSOC
                   flxCotaVend.Cell(flexcpText, I, conCOL_Tipo) = BREC2!SGI_DESCRICAO
                   flxCotaVend.Cell(flexcpText, I, conCOL_TipOrca) = BREC2!SGI_CODTIPORC
                   flxCotaVend.Cell(flexcpText, I, conCOL_Pedido) = strSTATUS
                   flxCotaVend.Cell(flexcpText, I, conCOL_Status) = IIf(Trim(BREC2!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC2!SGI_STATUS) = "B", "Baixado", "Parcial"))
                ElseIf stPedidos.Tab = 1 Then
                   flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_Codigo) = Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4)
                   flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_Data) = Format(BREC2!SGI_DATACOTA, "DD/MM/YYYY")
                   flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_Cliente) = BREC2!SGI_RAZAOSOC
                   flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_Tipo) = BREC2!SGI_DESCRICAO
                   flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_TipOrca) = BREC2!SGI_CODTIPORC
                   flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_Pedido) = strSTATUS
                   flxCOTACAOPARCIAL.Cell(flexcpText, I, conCOL_Status) = IIf(Trim(BREC2!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC2!SGI_STATUS) = "B", "Baixado", "Parcial"))
                ElseIf stPedidos.Tab = 2 Then
                   flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_Codigo) = Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4)
                   flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_Data) = Format(BREC2!SGI_DATACOTA, "DD/MM/YYYY")
                   flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_Cliente) = BREC2!SGI_RAZAOSOC
                   flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_Tipo) = BREC2!SGI_DESCRICAO
                   flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_TipOrca) = BREC2!SGI_CODTIPORC
                   flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_Pedido) = strSTATUS
                   flxCOTACAOBAIXTOTAL.Cell(flexcpText, I, conCOL_Status) = IIf(Trim(BREC2!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC2!SGI_STATUS) = "B", "Baixado", "Parcial"))
                ElseIf stPedidos.Tab = 3 Then
                   grdNAOATEND.Cell(flexcpText, I, conCOL_SonNaoAtend_Codigo) = Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4)
                   grdNAOATEND.Cell(flexcpText, I, conCOL_SonNaoAtend_Data) = Format(BREC2!SGI_DATACOTA, "DD/MM/YYYY")
                   grdNAOATEND.Cell(flexcpText, I, conCOL_SonNaoAtend_Cliente) = BREC2!SGI_RAZAOSOC
                   grdNAOATEND.Cell(flexcpText, I, conCOL_SonNaoAtend_Tipo) = BREC2!SGI_DESCRICAO
                   grdNAOATEND.Cell(flexcpText, I, conCOL_SonNaoAtend_TipoOrca) = BREC2!SGI_CODTIPORC
                   grdNAOATEND.Cell(flexcpText, I, conCOL_SonNaoAtend_Pedido) = intSTATUS
                   grdNAOATEND.Cell(flexcpText, I, conCOL_SonNaoAtend_Status) = IIf(Trim(BREC2!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC2!SGI_STATUS) = "B", "Baixado", "Parcial"))
                End If
            End If
            BREC2.Close

        End If
        
     End If
     BREC.Close
      
End Sub


Private Sub Ordem()

  Dim strSTATUS As String
  Dim intSTATUS As Integer
  Dim I As Integer
  
  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
        sSql = "Select " & vbCrLf
        sSql = sSql & "       COTA.* " & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
        sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
        sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
        sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
        
        If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
        If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
        If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
        If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
        
        sSql = sSql & " Order by COTA.SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
        sSql = "Select " & vbCrLf
        sSql = sSql & "       COTA.* " & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
        sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
        sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
        sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
        
        If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
        If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
        If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
        If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
        
        sSql = sSql & " Order by COTA.SGI_DATACOTA "
  ElseIf cboFiltro.ListIndex = 2 Then
    sSql = "Select " & vbCrLf
    sSql = sSql & "       COTA.* " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
    
    If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
    If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
    If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
    If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
    
    sSql = sSql & " Order by CLIE.SGI_RAZAOSOC "
  ElseIf cboFiltro.ListIndex = 3 Then
    sSql = "Select " & vbCrLf
    sSql = sSql & "       COTA.* " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
    
    If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
    If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
    If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
    If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
    
    sSql = sSql & " Order by ORCA.SGI_DESCRICAO "
  ElseIf cboFiltro.ListIndex = 4 Then
    sSql = "Select " & vbCrLf
    sSql = sSql & "       COTA.* " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
    
    If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
    If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
    If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
    If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
    
    sSql = sSql & " Order by COTA.SGI_STATUS "
  End If
  
  BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
  Do While Not BREC.EOF
     
       If BREC!SGI_STATUS = "A" Then strSTATUS = "Não"
       If BREC!SGI_STATUS = "N" Then intSTATUS = 0
       
       '' Vendo se existe pedido
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strSTATUS = "Sim"
       BREC2.Close
       '' ---------------------------
       
       If stPedidos.Tab = 0 Then
            flxCotaVend.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                BREC!SGI_CODTIPORC & vbTab & _
                                strSTATUS & vbTab & _
                                IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
       ElseIf stPedidos.Tab = 1 Then
            flxCOTACAOPARCIAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                BREC!SGI_CODTIPORC & vbTab & _
                                strSTATUS & vbTab & _
                                IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
       
       ElseIf stPedidos.Tab = 2 Then
            flxCOTACAOBAIXTOTAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                BREC!SGI_CODTIPORC & vbTab & _
                                strSTATUS & vbTab & _
                                IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
       ElseIf stPedidos.Tab = 3 Then
            grdNAOATEND.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                BREC!SGI_CODTIPORC & vbTab & _
                                intSTATUS & vbTab & _
                                IIf(Trim(BREC!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
        End If
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADCOTAVENDAP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim strSTATUS As String
    Dim intSTATUS As Integer
    Dim I As Integer
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    If stPedidos.Tab = 0 Then ConfGrid
    If stPedidos.Tab = 1 Then ConfGridCotacaoParcial
    If stPedidos.Tab = 2 Then ConfGridCotacaoTotal
    If stPedidos.Tab = 3 Then ConfGrdNaoAtend
    
    If cboFiltro.ListIndex = 0 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       COTA.* " & vbCrLf
       sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
       sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
       sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
       sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       COTA.SGI_FILIAL  = " & FILIAL & vbCrLf
       sSql = sSql & "   And COTA.SGI_CODIGO  = " & txtCampos.Text & vbCrLf
       
       If lngCodVendAtua > 0 Then
          sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
       End If
       
       sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
       If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
       If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
       If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
       If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
       sSql = sSql & " Order by COTA.SGI_CODIGO "
        
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
       If Not BREC.EOF Then
            
          Do While Not BREC.EOF
              
             If BREC!SGI_4STATUS = "A" Then strSTATUS = "Não"
             If BREC!SGI_4STATUS = "N" Then intSTATUS = 0
       
             '' Vendo se existe pedido
             sSql = "Select " & vbCrLf
             sSql = sSql & "       * " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
             sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
                     
             BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
             If Not BREC2.EOF Then strSTATUS = "Sim"
             BREC2.Close
             '' ---------------------------
       
             If stPedidos.Tab = 0 Then
                flxCotaVend.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                    Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                    BREC!SGI_RAZAOSOC & vbTab & _
                                    BREC!SGI_DESCRICAO & vbTab & _
                                    BREC!SGI_CODTIPORC & vbTab & _
                                    strSTATUS & vbTab & _
                                    IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 1 Then
                flxCOTACAOPARCIAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                          Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                          BREC!SGI_RAZAOSOC & vbTab & _
                                          BREC!SGI_DESCRICAO & vbTab & _
                                          BREC!SGI_CODTIPORC & vbTab & _
                                          strSTATUS & vbTab & _
                                          IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 2 Then
                flxCOTACAOBAIXTOTAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            strSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 3 Then
                grdNAOATEND.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            intSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             End If
             
             BREC.MoveNext
          Loop
          
       End If
          
       BREC.Close
       If stPedidos.Tab = 0 Then flxCotaVend.SetFocus
       If stPedidos.Tab = 1 Then flxCOTACAOPARCIAL.SetFocus
       If stPedidos.Tab = 2 Then flxCOTACAOBAIXTOTAL.SetFocus
       If stPedidos.Tab = 3 Then grdNAOATEND.SetFocus
       Exit Sub
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       If IsDate(txtCampos.Text) = False Then
          MsgBox "Somente é permitido Data !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       COTA.* " & vbCrLf
       sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
       sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
       sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
       sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And COTA.SGI_DATACOTA = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
       
       If lngCodVendAtua > 0 Then
          sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
       End If
       
       sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
       If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
       If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
       If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
       If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
       sSql = sSql & " Order by COTA.SGI_DATACOTA "
        
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
       If Not BREC.EOF Then
            
          Do While Not BREC.EOF
              
             If BREC!SGI_STATUS = "A" Then strSTATUS = "Não"
             If BREC!SGI_STATUS = "N" Then intSTATUS = 0
       
             '' Vendo se existe pedido
             sSql = "Select " & vbCrLf
             sSql = sSql & "       * " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
             sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
                     
             BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
             If Not BREC2.EOF Then strSTATUS = "Sim"
             BREC2.Close
             '' ---------------------------
       
             If stPedidos.Tab = 0 Then
                flxCotaVend.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                    Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                    BREC!SGI_RAZAOSOC & vbTab & _
                                    BREC!SGI_DESCRICAO & vbTab & _
                                    BREC!SGI_CODTIPORC & vbTab & _
                                    strSTATUS & vbTab & _
                                    IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 1 Then
                flxCOTACAOPARCIAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                          Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                          BREC!SGI_RAZAOSOC & vbTab & _
                                          BREC!SGI_DESCRICAO & vbTab & _
                                          BREC!SGI_CODTIPORC & vbTab & _
                                          strSTATUS & vbTab & _
                                          IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 2 Then
                flxCOTACAOBAIXTOTAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            strSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 3 Then
                grdNAOATEND.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            intSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             End If
              
             BREC.MoveNext
          Loop
          
       End If
          
       BREC.Close
       If stPedidos.Tab = 0 Then flxCotaVend.SetFocus
       If stPedidos.Tab = 1 Then flxCOTACAOPARCIAL.SetFocus
       If stPedidos.Tab = 2 Then flxCOTACAOBAIXTOTAL.SetFocus
       If stPedidos.Tab = 3 Then grdNAOATEND.SetFocus
       Exit Sub
    
    
    ElseIf cboFiltro.ListIndex = 2 Then
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       COTA.* " & vbCrLf
       sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
       sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
       sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
       sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       COTA.SGI_FILIAL      = " & FILIAL & vbCrLf
       
       If lngCodVendAtua > 0 Then
          sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
       End If
       
       sSql = sSql & "   And CLIE.SGI_FILIAL      = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And CLIE.SGI_CODIGO      = COTA.SGI_CODCLI " & vbCrLf
       sSql = sSql & "   And CLIE.SGI_RAZAOSOC LIKE '" & Trim(txtCampos.Text) & "%'" & vbCrLf
       sSql = sSql & "   And ORCA.SGI_FILIAL      = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_CODIGO      = COTA.SGI_CODTIPORC " & vbCrLf
       If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
       If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
       If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
       If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
       sSql = sSql & " Order by CLIE.SGI_RAZAOSOC "
        
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
       If Not BREC.EOF Then
            
          Do While Not BREC.EOF
              
             If BREC!SGI_STATUS = "A" Then strSTATUS = "Não"
             If BREC!SGI_STATUS = "N" Then intSTATUS = 0
       
             '' Vendo se existe pedido
             sSql = "Select " & vbCrLf
             sSql = sSql & "       * " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
             sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
                     
             BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
             If Not BREC2.EOF Then strSTATUS = "Sim"
             BREC2.Close
             '' ---------------------------
       
             If stPedidos.Tab = 0 Then
                flxCotaVend.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                    Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                    BREC!SGI_RAZAOSOC & vbTab & _
                                    BREC!SGI_DESCRICAO & vbTab & _
                                    BREC!SGI_CODTIPORC & vbTab & _
                                    strSTATUS & vbTab & _
                                    IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 1 Then
                flxCOTACAOPARCIAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                          Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                          BREC!SGI_RAZAOSOC & vbTab & _
                                          BREC!SGI_DESCRICAO & vbTab & _
                                          BREC!SGI_CODTIPORC & vbTab & _
                                          strSTATUS & vbTab & _
                                          IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 2 Then
                flxCOTACAOBAIXTOTAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            strSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 3 Then
                grdNAOATEND.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            intSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             End If
     
             BREC.MoveNext
          Loop
          
       End If
          
       BREC.Close
       If stPedidos.Tab = 0 Then flxCotaVend.SetFocus
       If stPedidos.Tab = 1 Then flxCOTACAOPARCIAL.SetFocus
       If stPedidos.Tab = 2 Then flxCOTACAOBAIXTOTAL.SetFocus
       If stPedidos.Tab = 3 Then grdNAOATEND.SetFocus
       Exit Sub
    
    ElseIf cboFiltro.ListIndex = 3 Then
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       COTA.* " & vbCrLf
       sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
       sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
       sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
       sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       COTA.SGI_FILIAL      = " & FILIAL & vbCrLf
       
       If lngCodVendAtua > 0 Then
          sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
       End If
       
       sSql = sSql & "   And CLIE.SGI_FILIAL      = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And CLIE.SGI_CODIGO      = COTA.SGI_CODCLI " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_FILIAL      = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_CODIGO      = COTA.SGI_CODTIPORC " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_DESCRICAO LIKE '" & Trim(txtCampos.Text) & "%'" & vbCrLf
       If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
       If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
       If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
       If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
       
       sSql = sSql & " Order by ORCA.SGI_DESCRICAO "
        
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
       If Not BREC.EOF Then
            
          Do While Not BREC.EOF
              
             If BREC!SGI_STATUS = "A" Then strSTATUS = "Não"
             If BREC!SGI_STATUS = "N" Then intSTATUS = 0
       
             '' Vendo se existe pedido
             sSql = "Select " & vbCrLf
             sSql = sSql & "       * " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
             sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
                     
             BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
             If Not BREC2.EOF Then strSTATUS = "Sim"
             BREC2.Close
             '' ---------------------------
       
             If stPedidos.Tab = 0 Then
                flxCotaVend.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                    Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                    BREC!SGI_RAZAOSOC & vbTab & _
                                    BREC!SGI_DESCRICAO & vbTab & _
                                    BREC!SGI_CODTIPORC & vbTab & _
                                    strSTATUS & vbTab & _
                                    IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 1 Then
                flxCOTACAOPARCIAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                          Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                          BREC!SGI_RAZAOSOC & vbTab & _
                                          BREC!SGI_DESCRICAO & vbTab & _
                                          BREC!SGI_CODTIPORC & vbTab & _
                                          strSTATUS & vbTab & _
                                          IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 2 Then
                flxCOTACAOBAIXTOTAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            strSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 3 Then
                grdNAOATEND.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            intSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             End If
     
             BREC.MoveNext
          Loop
          
       End If
          
       BREC.Close
       If stPedidos.Tab = 0 Then flxCotaVend.SetFocus
       If stPedidos.Tab = 1 Then flxCOTACAOPARCIAL.SetFocus
       If stPedidos.Tab = 2 Then flxCOTACAOBAIXTOTAL.SetFocus
       If stPedidos.Tab = 3 Then grdNAOATEND.SetFocus
       Exit Sub
    
    ElseIf cboFiltro.ListIndex = 4 Then
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       COTA.* " & vbCrLf
       sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
       sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
       sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
       sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       COTA.SGI_FILIAL      = " & FILIAL & vbCrLf
       
       If lngCodVendAtua > 0 Then
          sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
       End If
       
       sSql = sSql & "   And COTA.SGI_STATUS LIKE '" & Trim(txtCampos.Text) & "%'" & vbCrLf
       sSql = sSql & "   And CLIE.SGI_FILIAL      = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And CLIE.SGI_CODIGO      = COTA.SGI_CODCLI " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_FILIAL      = COTA.SGI_FILIAL " & vbCrLf
       sSql = sSql & "   And ORCA.SGI_CODIGO      = COTA.SGI_CODTIPORC " & vbCrLf
       If stPedidos.Tab = 0 Then sSql = sSql & "   And COTA.SGI_STATUS = 'A'" & vbCrLf
       If stPedidos.Tab = 1 Then sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
       If stPedidos.Tab = 2 Then sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
       If stPedidos.Tab = 3 Then sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
       sSql = sSql & " Order by COTA.SGI_STATUS "
        
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
       If Not BREC.EOF Then
            
          Do While Not BREC.EOF
              
             If BREC!SGI_STATUS = "A" Then strSTATUS = "Não"
             If BREC!SGI_STATUS = "N" Then intSTATUS = 0
       
             '' Vendo se existe pedido
             sSql = "Select " & vbCrLf
             sSql = sSql & "       * " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
             sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
                     
             BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
             If Not BREC2.EOF Then strSTATUS = "Sim"
             BREC2.Close
             '' ---------------------------
       
             If stPedidos.Tab = 0 Then
                flxCotaVend.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                    Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                    BREC!SGI_RAZAOSOC & vbTab & _
                                    BREC!SGI_DESCRICAO & vbTab & _
                                    BREC!SGI_CODTIPORC & vbTab & _
                                    strSTATUS & vbTab & _
                                    IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 1 Then
                flxCOTACAOPARCIAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                          Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                          BREC!SGI_RAZAOSOC & vbTab & _
                                          BREC!SGI_DESCRICAO & vbTab & _
                                          BREC!SGI_CODTIPORC & vbTab & _
                                          strSTATUS & vbTab & _
                                          IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 2 Then
                flxCOTACAOBAIXTOTAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            strSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             ElseIf stPedidos.Tab = 3 Then
                grdNAOATEND.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                            Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                            BREC!SGI_RAZAOSOC & vbTab & _
                                            BREC!SGI_DESCRICAO & vbTab & _
                                            BREC!SGI_CODTIPORC & vbTab & _
                                            intSTATUS & vbTab & _
                                            IIf(Trim(BREC!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
             End If
     
             BREC.MoveNext
          Loop
          
       End If
          
       BREC.Close
       If stPedidos.Tab = 0 Then flxCotaVend.SetFocus
       If stPedidos.Tab = 1 Then flxCOTACAOPARCIAL.SetFocus
       If stPedidos.Tab = 2 Then flxCOTACAOBAIXTOTAL.SetFocus
       If stPedidos.Tab = 3 Then grdNAOATEND.SetFocus
       Exit Sub
    
    End If


End Sub

Private Sub ConfGridCotacaoParcial()
    
    flxCOTACAOPARCIAL.Rows = 1
    flxCOTACAOPARCIAL.Cols = 7
    flxCOTACAOPARCIAL.FixedCols = 0
    flxCOTACAOPARCIAL.AllowBigSelection = False
    
    flxCOTACAOPARCIAL.Editable = flexEDNone
    
    flxCOTACAOPARCIAL.FormatString = "Código|Data|Cliente|Tipo|TIPORCA|Pedido|Status"
    
    flxCOTACAOPARCIAL.ColWidth(conCOL_Codigo) = 1000
    flxCOTACAOPARCIAL.ColWidth(conCOL_Data) = 1000
    flxCOTACAOPARCIAL.ColWidth(conCOL_Cliente) = 5000
    flxCOTACAOPARCIAL.ColWidth(conCOL_Tipo) = 2000
    flxCOTACAOPARCIAL.ColWidth(conCOL_TipOrca) = 0
    flxCOTACAOPARCIAL.ColWidth(conCOL_Pedido) = 600
    flxCOTACAOPARCIAL.ColWidth(conCOL_Status) = 700
    
End Sub

Private Sub ConfGridCotacaoTotal()
    
    flxCOTACAOBAIXTOTAL.Rows = 1
    flxCOTACAOBAIXTOTAL.Cols = 7
    flxCOTACAOBAIXTOTAL.FixedCols = 0
    flxCOTACAOBAIXTOTAL.AllowBigSelection = False
    
    flxCOTACAOBAIXTOTAL.Editable = flexEDNone
    
    flxCOTACAOBAIXTOTAL.FormatString = "Código|Data|Cliente|Tipo|TIPORCA|Pedido|Status"
    
    flxCOTACAOBAIXTOTAL.ColWidth(conCOL_Codigo) = 1000
    flxCOTACAOBAIXTOTAL.ColWidth(conCOL_Data) = 1000
    flxCOTACAOBAIXTOTAL.ColWidth(conCOL_Cliente) = 5000
    flxCOTACAOBAIXTOTAL.ColWidth(conCOL_Tipo) = 2000
    flxCOTACAOBAIXTOTAL.ColWidth(conCOL_TipOrca) = 0
    flxCOTACAOBAIXTOTAL.ColWidth(conCOL_Pedido) = 600
    flxCOTACAOBAIXTOTAL.ColWidth(conCOL_Status) = 700
    
End Sub

Private Sub PreencheGridBaixadoParcial()

    Dim strSTATUS       As String
    Dim I               As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       COTA.* " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If lngCodVendAtua > 0 Then
        sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
    End If
    
    sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
    sSql = sSql & "   And COTA.SGI_STATUS = 'P'" & vbCrLf
    sSql = sSql & " Order by COTA.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       If BREC!SGI_STATUS = "A" Then strSTATUS = "Não"
       
       '' Vendo se existe pedido
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strSTATUS = "Sim"
       BREC2.Close
       '' ---------------------------
       
       flxCOTACAOPARCIAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                  Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                  BREC!SGI_RAZAOSOC & vbTab & _
                                  BREC!SGI_DESCRICAO & vbTab & _
                                  BREC!SGI_CODTIPORC & vbTab & _
                                  strSTATUS & vbTab & _
                                  IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
                           
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PreencheGridBaixadoTotal()

    Dim strSTATUS       As String
    Dim I               As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       COTA.* " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If lngCodVendAtua > 0 Then
        sSql = sSql & "   And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
    End If
    
    sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
    sSql = sSql & "   And COTA.SGI_STATUS = 'B'" & vbCrLf
    sSql = sSql & " Order by COTA.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       If BREC!SGI_STATUS = "A" Then strSTATUS = "Não"
       
       '' Vendo se existe pedido
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODCOTA = " & BREC!SGI_CODIGO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strSTATUS = "Sim"
       BREC2.Close
       '' ---------------------------
       
       flxCOTACAOBAIXTOTAL.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                                   Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                                   BREC!SGI_RAZAOSOC & vbTab & _
                                   BREC!SGI_DESCRICAO & vbTab & _
                                   BREC!SGI_CODTIPORC & vbTab & _
                                   strSTATUS & vbTab & _
                                   IIf(Trim(BREC!SGI_STATUS) = "A", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
                           
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub


Private Sub ConfGrdNaoAtend()

    With grdNAOATEND
    
       .Cols = conColumnsIn_SonNaoAtend
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonNaoAtend_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonNaoAtend_Codigo) = ""
       .ColDataType(conCOL_SonNaoAtend_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonNaoAtend_Data) = ""
       .ColDataType(conCOL_SonNaoAtend_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonNaoAtend_Cliente) = ""
       .ColDataType(conCOL_SonNaoAtend_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonNaoAtend_Tipo) = ""
       .ColDataType(conCOL_SonNaoAtend_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonNaoAtend_TipoOrca) = ""
       .ColDataType(conCOL_SonNaoAtend_TipoOrca) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonNaoAtend_Pedido) = ""
       .ColDataType(conCOL_SonNaoAtend_Pedido) = flexDTBoolean
       .ColFormat(conCOL_SonNaoAtend_Pedido) = "Sim;Não"
       
       .Cell(flexcpData, 0, conCOL_SonNaoAtend_Status) = ""
       .ColDataType(conCOL_SonNaoAtend_Status) = flexDTString
       
       .ColWidth(conCOL_SonNaoAtend_Codigo) = 1000
       .ColWidth(conCOL_SonNaoAtend_Data) = 1000
       .ColWidth(conCOL_SonNaoAtend_Cliente) = 5000
       .ColWidth(conCOL_SonNaoAtend_Tipo) = 2000
       .ColWidth(conCOL_SonNaoAtend_TipoOrca) = 0
       .ColWidth(conCOL_SonNaoAtend_Pedido) = 600
       .ColWidth(conCOL_SonNaoAtend_Status) = 700
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PreencheGridNaoAtend()

    Dim intSTATUS
    Dim I               As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       COTA.* " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ORCA.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDH COTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA   ORCA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       COTA.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If lngCodVendAtua > 0 Then
       sSql = sSql & "    And COTA.SGI_CODVEND = " & lngCodVendAtua & vbCrLf
    End If
    
    sSql = sSql & "   And CLIE.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = COTA.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_FILIAL = COTA.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORCA.SGI_CODIGO = COTA.SGI_CODTIPORC " & vbCrLf
    sSql = sSql & "   And COTA.SGI_STATUS = 'N'" & vbCrLf
    sSql = sSql & " Order by COTA.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       If BREC!SGI_STATUS = "A" Then intSTATUS = 0
       
       '' ---------------------------
       
       grdNAOATEND.AddItem Mid(Trim(Str(BREC!SGI_CODIGO)), 1, (Len(Trim(Str(BREC!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC!SGI_CODIGO), 4) & vbTab & _
                           Format(BREC!SGI_DATACOTA, "DD/MM/YYYY") & vbTab & _
                           BREC!SGI_RAZAOSOC & vbTab & _
                           BREC!SGI_DESCRICAO & vbTab & _
                           BREC!SGI_CODTIPORC & vbTab & _
                           intSTATUS & vbTab & _
                           IIf(Trim(BREC!SGI_STATUS) = "N", "Aberto", IIf(Trim(BREC!SGI_STATUS) = "B", "Baixado", "Parcial"))
                           
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub

Private Function PegaCodVendedor(strUsuario As String) As Long
    
    PegaCodVendedor = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       VEN.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO      USU" & vbCrLf
    sSql = sSql & "      ,SGI_CADVENDEDOR  VEN" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       USU.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And USU.SGI_NOME       = '" & Trim(objFuncoes.Crypt(strUsuario)) & "'" & vbCrLf
    sSql = sSql & "   And VEN.SGI_FILIAL     = USU.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And VEN.SGI_CODUSUARIO = USU.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaCodVendedor = BREC!SGI_CODIGO
    BREC.Close
    
End Function
