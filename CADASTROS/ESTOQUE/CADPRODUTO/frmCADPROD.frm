VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADPRODP 
   Caption         =   "Cadastro de Produto"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   15570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   15495
      Begin VB.CommandButton cmdAtivDesat 
         Caption         =   "&Status"
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
         Left            =   3600
         Picture         =   "frmCADPROD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Muda o Status do Produto"
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   6000
         Top             =   240
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
         Picture         =   "frmCADPROD.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmCADPROD.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Inclui um novo produto"
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
         Picture         =   "frmCADPROD.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Altera Produto"
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
         Picture         =   "frmCADPROD.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exclui Produto"
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
         Left            =   13680
         Picture         =   "frmCADPROD.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Height          =   615
         Left            =   14520
         Picture         =   "frmCADPROD.frx":129C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   15495
      Begin TabDlg.SSTab stTabProd 
         Height          =   6375
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   11245
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
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
         TabCaption(0)   =   "RÓTULOS - [ Ativos ]"
         TabPicture(0)   =   "frmCADPROD.frx":139E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdPRODFABRACAB"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "MATÉRIA PRIMA - [ Ativos ]"
         TabPicture(1)   =   "frmCADPROD.frx":13BA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdBRUTACAB"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "RÓTULOS - [ INATIVOS ]"
         TabPicture(2)   =   "frmCADPROD.frx":13D6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "grdPRODINAT"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "[ PRÉ CADASTRO RÓTULOS / AGUARDANDO LIBERAÇÃO ]"
         TabPicture(3)   =   "frmCADPROD.frx":13F2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Command6"
         Tab(3).Control(1)=   "grdPRODAGLIB"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Geral"
         TabPicture(4)   =   "frmCADPROD.frx":140E
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         Begin VB.CommandButton Command6 
            Caption         =   "Libera"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   -74880
            Picture         =   "frmCADPROD.frx":142A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Liberação das Alterações do Fotolito"
            Top             =   5640
            Width           =   1335
         End
         Begin VSFlex8LCtl.VSFlexGrid grdPRODAGLIB 
            Height          =   5175
            Left            =   -74880
            TabIndex        =   14
            Top             =   360
            Width           =   15015
            _cx             =   26485
            _cy             =   9128
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
         Begin VSFlex8LCtl.VSFlexGrid grdPRODINAT 
            Height          =   5895
            Left            =   -74880
            TabIndex        =   13
            Top             =   360
            Width           =   15015
            _cx             =   26485
            _cy             =   10398
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
         Begin VSFlex8LCtl.VSFlexGrid grdBRUTACAB 
            Height          =   5895
            Left            =   -74880
            TabIndex        =   12
            Top             =   360
            Width           =   15015
            _cx             =   26485
            _cy             =   10398
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
         Begin VSFlex8LCtl.VSFlexGrid grdPRODFABRACAB 
            Height          =   5895
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   15015
            _cx             =   26485
            _cy             =   10398
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
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      Begin VB.TextBox txtCodigio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Text            =   "txtCodigio"
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton cmdLinProd 
         Height          =   315
         Left            =   9600
         Picture         =   "frmCADPROD.frx":1857
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtLinProd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "txtLinProd"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtNomClie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         TabIndex        =   24
         Text            =   "txtNomClie"
         Top             =   120
         Width           =   5295
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "txtCIDCLIE"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   9600
         Picture         =   "frmCADPROD.frx":1959
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtComplemento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Text            =   "txtComplemento"
         Top             =   840
         Width           =   6375
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Text            =   "txtDescricao"
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label lblLinhProd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblLinhProd"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   9960
         TabIndex        =   28
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Linha"
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
         Left            =   7920
         TabIndex        =   27
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   7920
         TabIndex        =   20
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Complemento"
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
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
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
         Width           =   870
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmCADPRODP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public FILIAL        As Integer
Public strAcesso     As String
Public strUSUARIO    As String
Public lngCodUsuario As Long
Dim objFuncoes       As Object
Dim objCADPRODUTO    As Object
Dim objPESQPADRAO    As Object
Dim iCodigo          As Long
Dim boolComAcao      As Boolean


Const conCOL_ProdFrabAcab_IDProd                    As Integer = 0
Const conCOL_ProdFrabAcab_CodRotulo                 As Integer = 1
Const conCOL_ProdFrabAcab_Descricao                 As Integer = 2
Const conCOL_ProdFrabAcab_Especie                   As Integer = 3
Const conCOL_ProdFrabAcab_Complemento               As Integer = 4
Const conCOL_ProdFrabAcab_Status                    As Integer = 5
Const conCOL_ProdFrabAcab_CodStatus                 As Integer = 6
Const conCOL_ProdFrabAcab_FormatString              As String = "=IDProd|Rótulo|Descrição|Espécie|Complemento|Status|CodStatus"
Const conColumnsIn_ProdFrabAcab                     As Integer = 7

Const conCOL_ProdBrutAcab_IDProd                    As Integer = 0
Const conCOL_ProdBrutAcab_CodRotulo                 As Integer = 1
Const conCOL_ProdBrutAcab_Descricao                 As Integer = 2
Const conCOL_ProdBrutAcab_Especie                   As Integer = 3
Const conCOL_ProdBrutAcab_Complemento               As Integer = 4
Const conCOL_ProdBrutAcab_Status                    As Integer = 5
Const conCOL_ProdBrutAcab_CodStatus                 As Integer = 6
Const conCOL_ProdBrutAcab_FormatString              As String = "=IDProd|Código|Descrição|Espécie|Complemento|Status|CodStatus"
Const conColumnsIn_ProdBrutAcab                     As Integer = 7

Const conCOL_ProdFrabAcabInat_IDProd                As Integer = 0
Const conCOL_ProdFrabAcabInat_CodRotulo             As Integer = 1
Const conCOL_ProdFrabAcabInat_Descricao             As Integer = 2
Const conCOL_ProdFrabAcabInat_Especie               As Integer = 3
Const conCOL_ProdFrabAcabInat_Complemento           As Integer = 4
Const conCOL_ProdFrabAcabInat_Status                As Integer = 5
Const conCOL_ProdFrabAcabInat_CodStatus             As Integer = 6
Const conCOL_ProdFrabAcabInat_FormatString          As String = "=IDProd|Rótulo|Descrição|Espécie|Complemento|Status|CodStatus"
Const conColumnsIn_ProdFrabAcabInat                 As Integer = 7

Const conCOL_ProdFrabAgLib_IDProd                    As Integer = 0
Const conCOL_ProdFrabAgLib_CodRotulo                 As Integer = 1
Const conCOL_ProdFrabAgLib_Descricao                 As Integer = 2
Const conCOL_ProdFrabAgLib_Especie                   As Integer = 3
Const conCOL_ProdFrabAgLib_Complemento               As Integer = 4
Const conCOL_ProdFrabAgLib_Status                    As Integer = 5
Const conCOL_ProdFrabAgLib_CodStatus                 As Integer = 6
Const conCOL_ProdFrabAgLib_FormatString              As String = "=IDProd|Rótulo|Descrição|Espécie|Complemento|Status|CodStatus"
Const conColumnsIn_ProdFrabAgLib                     As Integer = 7



Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Operacao "A"
End Sub

Private Sub cmdAtivDesat_Click()

    Dim intStatus As Integer
    
    If Verif_reg = True Then Exit Sub
    
    Dim iResp As Integer
    
    iResp = MsgBox("Confirma a alteração do STATUS do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
    
    If iResp <> 6 Then Exit Sub
    
    If stTabProd.Tab = 0 Then
        With grdPRODFABRACAB
            If .Row = 0 Then
                MsgBox "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Sub
            End If
            If .Cell(flexcpText, .Row, conCOL_ProdFrabAcab_CodStatus) = 1 Then
                intStatus = 0
            ElseIf .Cell(flexcpText, .Row, conCOL_ProdFrabAcab_CodStatus) = 0 Then
                intStatus = 1
            End If
            ''If Verifica_Se_TemPedido = True Then
            ''    MsgBox "ATENÇÂO - Existe Pedidos em Aberto para este Rótulo não pode ser desativado !!!", vbOKOnly + vbExclamation, "Aviso"
            ''    Exit Sub
            ''End If
        End With
    ElseIf stTabProd.Tab = 1 Then
          With grdBRUTACAB
              If .Row = 0 Then
                 MsgBox "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
                 Exit Sub
              End If
              If .Cell(flexcpText, .Row, conCOL_ProdBrutAcab_CodStatus) = 1 Then
                 intStatus = 0
              ElseIf .Cell(flexcpText, .Row, conCOL_ProdBrutAcab_CodStatus) = 0 Then
                 intStatus = 1
              End If
          End With
    ElseIf stTabProd.Tab = 2 Then
          With grdPRODINAT
              If .Row = 0 Then
                 MsgBox "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
                 Exit Sub
              End If
              If .Cell(flexcpText, .Row, conCOL_ProdFrabAcabInat_CodStatus) = 1 Then
                 intStatus = 0
              ElseIf .Cell(flexcpText, .Row, conCOL_ProdFrabAcabInat_CodStatus) = 0 Then
                 intStatus = 1
              End If
          End With
    End If
    
    objCADPRODUTO.STATUS = intStatus
  
    If objCADPRODUTO.GRAVA("AS") = False Then Exit Sub
    If objCADPRODUTO.Atualiza("A", objCADPRODUTO.IDProduto, FILIAL, "frmCADPROD") = False Then Exit Sub
    
    MsgBox "O Status do Registro foi alterado com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
    Call AbilitaCampos

    Call ConfGrid
    Call ConfGridDiversos
    Call ConfGridInat

End Sub

Private Sub cmdCanFiltro_Click()
    
    Call AbilitaCampos
    
    Call ConfGrid
    Call ConfGridDiversos
    
    Call PreencheGrid
    Call PreencheGridDiversos
    
End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub

  If Verif_reg = True Then Exit Sub
  
  Dim iResp As Integer
  
  If Verifica_Integracao = False Then Exit Sub
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADPRODUTO.GRAVA("E") = False Then Exit Sub
  If objCADPRODUTO.Atualiza("E", objCADPRODUTO.IDProduto, FILIAL, "frmCADPROD") = False Then Exit Sub
    
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call Atualiza_Grid

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub


Private Sub cmdLinProd_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODLIN"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODLIN"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Linha de Produto", "CADLINHAPROD.clsCADLINHAPROD")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtLinProd.Text = varRETORNO
       lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
    End If
    

End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub



Private Sub Command1_Click()

On Error GoTo Err_Command1_Click


    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CLIE.* " & vbCrLf
   
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE  CLIE" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLIE.SGI_FILIAL = " & FILIAL & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "CLIE.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "CLIE.SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "3000"
    arrCAMPOS(3, 5) = "CLIE.SGI_RAZAOSOC"
    
    arrCAMPOS(4, 1) = "SGI_NOMFANTA"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Nome Fantasia"
    arrCAMPOS(4, 4) = "2000"
    arrCAMPOS(4, 5) = "CLIE.SGI_NOMFANTA"
    
    arrCAMPOS(5, 1) = "SGI_CIDNORM"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Cidade"
    arrCAMPOS(5, 4) = "1500"
    arrCAMPOS(5, 5) = "CLIE.SGI_CIDNORM"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCIDCLIE.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, txtNomClie)

    Exit Sub
    
Err_Command1_Click:
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : Command1_Click()", Me.Name, "Command1_Click()", strCAMARQERRO)

End Sub

Private Sub Command6_Click()
    
    Dim intStatus As Integer
    If Verif_reg = True Then Exit Sub
    
    Dim iResp As Integer
    
    iResp = MsgBox("Confirma a liberação do Produto ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
    If iResp <> 6 Then Exit Sub
    
    If stTabProd.Tab = 3 Then
          With grdPRODAGLIB
              If .Row = 0 Then
                 MsgBox "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
                 Exit Sub
              End If
              intStatus = 1
          End With
    End If
    
    If ConsisteCampos = False Then
        MsgBox "ATENÇÂO - Falta Dados para Incluir Impossivel Liberar !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    Else
        objCADPRODUTO.STATUS = intStatus
        If objCADPRODUTO.PRODNOVO = 1 Then objCADPRODUTO.PRODNOVO = 0
        If objCADPRODUTO.FOTALTSN = 1 Then objCADPRODUTO.FOTALTSN = 0
    End If
    
    If objCADPRODUTO.GRAVA("AS") = False Then Exit Sub
    If objCADPRODUTO.Atualiza("A", objCADPRODUTO.IDProduto, FILIAL, "frmCADPROD") = False Then Exit Sub
    
    MsgBox "O Produto Foi Liberado com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
    Call AbilitaCampos

    Call ConfGrid
    Call ConfGridDiversos
    Call ConfGridInat
  

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADPRODUTO = CreateObject("CADPRODU.clsCADPRODU")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    
    objCADPRODUTO.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADPRODP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call AbilitaCampos
    Call ConfGrid
    Call ConfGridDiversos
    Call ConfGridInat
    Call ConfGridAgLib
    
    stTabProd.Tab = 0
      
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    strCamImgRotulos = Right(Linha(8), Len(Trim(Linha(8))) - 7)
    
    boolComAcao = False

    Call AtivaDesativaBotoes

    Call LimpaCamposLabel
    
End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
  
  If strOperacao = "A" Or strOperacao = "C" Then
     If Verif_reg = True Then Exit Sub
  End If
  
  If stTabProd.Tab = 0 Then
     With grdPRODFABRACAB
        If (.Rows - 1) > 0 And (.Row > 0) Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAcab_IDProd))
     End With
  ElseIf stTabProd.Tab = 1 Then
     With grdBRUTACAB
        If (.Rows - 1) > 0 And (.Row > 0) Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_ProdBrutAcab_IDProd))
     End With
  ElseIf stTabProd.Tab = 2 Then
     With grdPRODINAT
        If (.Rows - 1) > 0 And (.Row > 0) Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAcabInat_IDProd))
     End With
  ElseIf stTabProd.Tab = 3 Then
     With grdPRODAGLIB
        If (.Rows - 1) > 0 And (.Row > 0) Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAgLib_IDProd))
     End With
  End If
  
  boolComAcao = True
  
  frmCADPROD.cCaminho = cCaminho
  frmCADPROD.Linha = Linha
  frmCADPROD.iCodigo = iCodigo
  frmCADPROD.cTipOper = strOperacao
  frmCADPROD.FILIAL = FILIAL
  frmCADPROD.strAcesso = strAcesso
  frmCADPROD.lngCodUsuario = lngCodUsuario
  frmCADPROD.Show vbModal
  
  boolComAcao = False
  
  Atualiza_Grid

End Sub

Private Sub AbilitaCampos()
    
    If objCADPRODUTO.Pesq_CadProduto = False Then
       cmdAltera.Enabled = False
       cmdExclui.Enabled = False
       Frame1.Enabled = False
       Frame3.Enabled = False
    Else
       cmdAltera.Enabled = True
       cmdExclui.Enabled = True
       Frame1.Enabled = True
       Frame3.Enabled = True
    End If

End Sub

Private Sub ConfGrid()
    
    With grdPRODFABRACAB
    
       .Cols = conColumnsIn_ProdFrabAcab
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdFrabAcab_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcab_IDProd) = ""
       .ColDataType(conCOL_ProdFrabAcab_IDProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcab_CodRotulo) = ""
       .ColDataType(conCOL_ProdFrabAcab_CodRotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcab_Descricao) = ""
       .ColDataType(conCOL_ProdFrabAcab_Descricao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcab_Especie) = ""
       .ColDataType(conCOL_ProdFrabAcab_Especie) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcab_Complemento) = ""
       .ColDataType(conCOL_ProdFrabAcab_Complemento) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcab_Status) = ""
       .ColDataType(conCOL_ProdFrabAcab_Status) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcab_CodStatus) = ""
       .ColDataType(conCOL_ProdFrabAcab_CodStatus) = flexDTLong
       
       .ColWidth(conCOL_ProdFrabAcab_IDProd) = 0
       .ColWidth(conCOL_ProdFrabAcab_CodRotulo) = 1500
       .ColWidth(conCOL_ProdFrabAcab_Descricao) = 5000
       .ColWidth(conCOL_ProdFrabAcab_Especie) = 2500
       .ColWidth(conCOL_ProdFrabAcab_Complemento) = 2500
       .ColWidth(conCOL_ProdFrabAcab_Status) = 1000
       .ColWidth(conCOL_ProdFrabAcab_CodStatus) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With


End Sub

Private Sub PreencheGrid()

    If BREC.State = 1 Then BREC.Close
    
    With grdPRODFABRACAB
    
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       PRO.* " & vbCrLf
        sSql = sSql & "      ,TIP.SGI_DESCRICAO as DESC_TIPO " & vbCrLf
        sSql = sSql & "      ,ESP.SGI_DESCRICAO as DESC_ESP " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO PRO " & vbCrLf
        sSql = sSql & "      ,SGI_CADESPPROD ESP " & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TIP " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       PRO.SGI_FILIAL        = " & FILIAL & vbCrLf
        sSql = sSql & "   And PRO.SGI_PRODUTOTIPO   = 1" & vbCrLf
        sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
        sSql = sSql & "   And ESP.SGI_FILIAL        = PRO.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ESP.SGI_CODIGO        = PRO.SGI_CODESPECIE " & vbCrLf
        sSql = sSql & "   And TIP.SGI_FILIAL        = PRO.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And TIP.SGI_CODIGO        = PRO.SGI_CODTIPO " & vbCrLf
        sSql = sSql & " Order by PRO.SGI_CODIGO " & vbCrLf
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        Do While Not BREC.EOF
           .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                    Format(IIf(IsNull(BREC!SGI_CODLINPROD), 0, BREC!SGI_CODLINPROD), "###000") & "." & _
                    Format(IIf(IsNull(BREC!SGI_CODCLIE), 0, BREC!SGI_CODCLIE), "####0000") & "." & _
                    Format(IIf(IsNull(BREC!SGI_CODROTULO), 0, BREC!SGI_CODROTULO), "##00") & "." & _
                    Format(IIf(IsNull(BREC!SGI_DIGVERIF), 0, BREC!SGI_DIGVERIF), "#0") & vbTab & _
                    BREC!SGI_DESCRICAO & vbTab & _
                    BREC!DESC_ESP & vbTab & _
                    BREC!DESC_TIPO & vbTab & _
                    IIf(BREC!SGI_STATUS = 1, "ATIVO", "DESATIVADO") & vbTab & _
                    BREC!SGI_STATUS
                              
           BREC.MoveNext
        Loop
        BREC.Close
    
    End With
End Sub

Private Sub Ordem()

    Dim boolTemRegistro As Boolean
    
    If stTabProd.Tab = 0 Then ConfGrid
    If stTabProd.Tab = 1 Then ConfGridDiversos
    If stTabProd.Tab = 2 Then ConfGridInat
    If stTabProd.Tab = 3 Then ConfGridAgLib
    
    sSql = ""
  
    If BREC.State = 1 Then BREC.Close
    
    boolTemRegistro = False
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.* " & vbCrLf
    If stTabProd.Tab <> 1 Then sSql = sSql & "      ,CLI.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,TIP.SGI_DESCRICAO as DESC_TIPO " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO as DESC_ESP " & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD ESP" & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPPROD TIP" & vbCrLf
    If stTabProd.Tab <> 1 Then sSql = sSql & "      ,SGI_CADCLIENTE CLI" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
  
    If stTabProd.Tab = 0 Or stTabProd.Tab = 2 Or stTabProd.Tab = 3 Then
       sSql = sSql & "   And PRO.SGI_PRODUTOTIPO   = 1" & vbCrLf
       sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
    ElseIf stTabProd.Tab = 1 Then
       sSql = sSql & "   And PRO.SGI_PRODUTOTIPO    = 0" & vbCrLf
       sSql = sSql & "   And (PRO.SGI_PRODUTOESTILO = 0 Or PRO.SGI_PRODUTOESTILO = 4)" & vbCrLf
    End If
  
    If stTabProd.Tab = 0 Then sSql = sSql & "   And PRO.SGI_STATUS = 1" & vbCrLf '' Ativo
    If stTabProd.Tab = 1 Then sSql = sSql & "   And PRO.SGI_STATUS = 1" & vbCrLf '' Ativo
    If stTabProd.Tab = 2 Then sSql = sSql & "   And PRO.SGI_STATUS = 0" & vbCrLf '' Inativo
    If stTabProd.Tab = 3 Then sSql = sSql & "   And PRO.SGI_STATUS = 2" & vbCrLf '' Aguardando Liberação
    
    sSql = sSql & "   And ESP.SGI_FILIAL = PRO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PRO.SGI_CODESPECIE " & vbCrLf
    sSql = sSql & "   And TIP.SGI_FILIAL = PRO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And TIP.SGI_CODIGO = PRO.SGI_CODTIPO " & vbCrLf
       
    If stTabProd.Tab <> 1 Then
        sSql = sSql & "   And CLI.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLI.SGI_CODIGO = PRO.SGI_CODCLIE" & vbCrLf
    End If
       
    If Len(Trim(txtCodigio.Text)) > 0 Then
        sSql = sSql & "   And PRO.SGI_CODIGO Like '%" & Trim(txtCodigio.Text) & "%'" & vbCrLf
    End If
    
    If Len(Trim(txtDescricao.Text)) > 0 Then
        sSql = sSql & "   And PRO.SGI_DESCRICAO Like '%" & Trim(txtDescricao.Text) & "%'" & vbCrLf
    End If
    
    If Len(Trim(txtComplemento.Text)) > 0 Then
        sSql = sSql & "   And PRO.SGI_COMPLEMENTO Like '%" & Trim(txtDescricao.Text) & "%'" & vbCrLf
    End If
    
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        sSql = sSql & "   And PRO.SGI_CODCLIE = " & Trim(txtCIDCLIE.Text) & vbCrLf
    ElseIf Len(Trim(txtNomClie.Text)) > 0 Then
        sSql = sSql & "   And CLI.SGI_RAZAOSOC Like '%" & Trim(txtDescricao.Text) & "%'" & vbCrLf
    End If
    
    If Len(Trim(txtLinProd.Text)) > 0 Then
        sSql = sSql & "   And PRO.SGI_CODLINPROD = " & Trim(txtLinProd.Text)
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then boolTemRegistro = True
    
    Do While Not BREC.EOF
       If stTabProd.Tab = 0 Then
      
         With grdPRODFABRACAB
              .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                       Format(IIf(IsNull(BREC!SGI_CODLINPROD), 0, BREC!SGI_CODLINPROD), "###000") & "." & _
                       Format(IIf(IsNull(BREC!SGI_CODCLIE), 0, BREC!SGI_CODCLIE), "####0000") & "." & _
                       Format(IIf(IsNull(BREC!SGI_CODROTULO), 0, BREC!SGI_CODROTULO), "##00") & "." & _
                       Format(IIf(IsNull(BREC!SGI_DIGVERIF), 0, BREC!SGI_DIGVERIF), "#0") & vbTab & _
                       BREC!SGI_DESCRICAO & vbTab & _
                       BREC!DESC_ESP & vbTab & _
                       BREC!SGI_COMPLEMENTO & vbTab & _
                       IIf(BREC!SGI_STATUS = 1, "ATIVO", "DESATIVADO") & vbTab & _
                       BREC!SGI_STATUS
         End With
         
      ElseIf stTabProd.Tab = 1 Then
      
          With grdBRUTACAB
               .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                        BREC!SGI_CODIGO & vbTab & _
                        BREC!SGI_DESCRICAO & vbTab & _
                        BREC!DESC_ESP & vbTab & _
                        BREC!SGI_COMPLEMENTO & vbTab & _
                        IIf(BREC!SGI_STATUS = 1, "ATIVO", "DESATIVADO") & vbTab & _
                        BREC!SGI_STATUS
          End With
      ElseIf stTabProd.Tab = 2 Then
      
         With grdPRODINAT
              .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                            Format(IIf(IsNull(BREC!SGI_CODLINPROD), 0, BREC!SGI_CODLINPROD), "###000") & "." & _
                            Format(IIf(IsNull(BREC!SGI_CODCLIE), 0, BREC!SGI_CODCLIE), "####0000") & "." & _
                            Format(IIf(IsNull(BREC!SGI_CODROTULO), 0, BREC!SGI_CODROTULO), "##00") & "." & _
                            Format(IIf(IsNull(BREC!SGI_DIGVERIF), 0, BREC!SGI_DIGVERIF), "#0") & vbTab & _
                            BREC!SGI_DESCRICAO & vbTab & _
                            BREC!DESC_ESP & vbTab & _
                            BREC!SGI_COMPLEMENTO & vbTab & _
                            IIf(BREC!SGI_STATUS = 1, "ATIVO", "DESATIVADO") & vbTab & _
                            BREC!SGI_STATUS
      
        End With
        
      ElseIf stTabProd.Tab = 3 Then
      
         With grdPRODAGLIB
              .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                            Format(IIf(IsNull(BREC!SGI_CODLINPROD), 0, BREC!SGI_CODLINPROD), "###000") & "." & _
                            Format(IIf(IsNull(BREC!SGI_CODCLIE), 0, BREC!SGI_CODCLIE), "####0000") & "." & _
                            Format(IIf(IsNull(BREC!SGI_CODROTULO), 0, BREC!SGI_CODROTULO), "##00") & "." & _
                            Format(IIf(IsNull(BREC!SGI_DIGVERIF), 0, BREC!SGI_DIGVERIF), "#0") & vbTab & _
                            BREC!SGI_DESCRICAO & vbTab & _
                            BREC!DESC_ESP & vbTab & _
                            BREC!SGI_COMPLEMENTO & vbTab & _
                            IIf(BREC!SGI_STATUS = 2, "AG.LIB", "") & vbTab & _
                            BREC!SGI_STATUS
         
         End With
      
      End If
      
      BREC.MoveNext
    Loop
    BREC.Close
    
    Call LimpaCamposLabel
    
    If boolTemRegistro = False Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Não há dados para apresentar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub

Private Sub grdBRUTACAB_Click()
    With grdBRUTACAB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then objCADPRODUTO.IDProduto = CLng(.Cell(flexcpText, .Row, conCOL_ProdBrutAcab_IDProd))
    End With
End Sub

Private Sub grdBRUTACAB_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    With grdBRUTACAB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then Call Operacao("C")
    End With
End Sub

Private Sub grdBRUTACAB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
       With grdBRUTACAB
            If (.Rows - 1) > 0 And (.RowSel > 0) Then Call Operacao("C")
       End With
    End If
End Sub

Private Sub grdBRUTACAB_RowColChange()
    With grdBRUTACAB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then objCADPRODUTO.IDProduto = CLng(.Cell(flexcpText, .Row, conCOL_ProdBrutAcab_IDProd))
    End With
End Sub

Private Sub grdPRODAGLIB_Click()
    With grdPRODAGLIB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then objCADPRODUTO.IDProduto = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAgLib_IDProd))
    End With
End Sub

Private Sub grdPRODAGLIB_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    With grdPRODAGLIB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then Call Operacao("C")
    End With
End Sub

Private Sub grdPRODAGLIB_RowColChange()
    With grdPRODAGLIB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then objCADPRODUTO.IDProduto = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAgLib_IDProd))
    End With
End Sub

Private Sub grdPRODFABRACAB_Click()
    With grdPRODFABRACAB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then objCADPRODUTO.IDProduto = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAcab_IDProd))
    End With
End Sub

Private Sub grdPRODFABRACAB_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    With grdPRODFABRACAB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then Call Operacao("C")
    End With
End Sub

Private Sub grdPRODFABRACAB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
       With grdPRODFABRACAB
            If (.Rows - 1) > 0 And (.RowSel > 0) Then Call Operacao("C")
       End With
    End If
End Sub

Private Sub grdPRODFABRACAB_RowColChange()
    With grdPRODFABRACAB
        If (.Rows - 1) > 0 And (.RowSel > 0) Then objCADPRODUTO.IDProduto = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAcab_IDProd))
    End With
End Sub


Private Sub grdPRODINAT_Click()
    With grdPRODINAT
        If (.Rows - 1) > 0 And (.RowSel > 0) Then objCADPRODUTO.IDProduto = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAcabInat_IDProd))
    End With
End Sub

Private Sub grdPRODINAT_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    With grdPRODINAT
        If (.Rows - 1) > 0 And (.RowSel > 0) Then Call Operacao("C")
    End With
End Sub

Private Sub grdPRODINAT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
       With grdPRODINAT
            If (.Rows - 1) > 0 And (.RowSel > 0) Then Call Operacao("C")
       End With
    End If
End Sub

Private Sub grdPRODINAT_RowColChange()
    With grdPRODINAT
        If (.Rows - 1) > 0 And (.RowSel > 0) Then objCADPRODUTO.IDProduto = CLng(.Cell(flexcpText, .Row, conCOL_ProdFrabAcabInat_IDProd))
    End With
End Sub


Private Sub Timer1_Timer()
    Call Atualiza_Grid
End Sub



Private Function Verifica_Integracao() As Boolean

     Verifica_Integracao = False
     
     
     '' Verifica Maquinas Se existe para este produto
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADMAQPROD " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_CODPROD = '" & objCADPRODUTO.CodigoProd & "'"
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     If Not BREC.EOF Then
        MsgBox "Há máquinas cadastrada para este produto impossivel Excluir !!!", vbInformation + vbOKOnly, "Aviso"
        BREC.Close
        Exit Function
     End If
     BREC.Close
     
     '' Verifica se este produto faz parte de uma hierarquia
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_LISTAMAT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_PRODLST = '" & objCADPRODUTO.CodigoProd & "'"
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     If Not BREC.EOF Then
        MsgBox "Este produto está relacionado a arvore de outro produto, impossivel Excluir !!!" & vbCrLf & "Produto Relacionado : " & Trim(BREC!SGI_PRODUTO), vbInformation + vbOKOnly, "Aviso"
        BREC.Close
        Exit Function
     End If
     BREC.Close
     
     '' Verifica Se existe requisição de Entrada de Material
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADITREQENTRMAT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_PRODUTO = '" & objCADPRODUTO.CodigoProd & "'"
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     If Not BREC.EOF Then
        MsgBox "Este produto tem requisições de entrada de material, impossivel Excluir !!!", vbInformation + vbOKOnly, "Aviso"
        BREC.Close
        Exit Function
     End If
     BREC.Close
     
     '' Verifica se Existe Requisição de Material
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADITREQMAT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_PRODUTO = '" & objCADPRODUTO.CodigoProd & "'"
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        MsgBox "Este produto tem requisições de material, impossivel Excluir !!!", vbInformation + vbOKOnly, "Aviso"
        BREC.Close
        Exit Function
     End If
     BREC.Close
     
     '' Verifica Se existe requisição de Saida de Material
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADITREQSAIMAT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_PRODUTO = '" & objCADPRODUTO.CodigoProd & "'"
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        MsgBox "Este produto tem requisições de saida de material, impossivel Excluir !!!", vbInformation + vbOKOnly, "Aviso"
        BREC.Close
        Exit Function
     End If
     BREC.Close
    
     '' -------------------------------------------------
     Verifica_Integracao = True
     
     
End Function

Private Sub Atualiza_Grid()
    
     If boolComAcao = True Then Exit Sub
     
     Dim i              As Integer
     Dim bolAchou       As Boolean
     Dim grdGENERICA    As VSFlexGrid
     Dim lngCODIGO      As Long
     Dim strACAO        As String
     Dim lngCol         As Long
      
     If adoBanco_Dados.State = 0 Then Exit Sub
     If BRECATU.State = 1 Then BRECATU.Close
     
     If stTabProd.Tab = 0 Then
        Set grdGENERICA = grdPRODFABRACAB
        lngCol = conCOL_ProdFrabAcab_IDProd
     ElseIf stTabProd.Tab = 1 Then
        Set grdGENERICA = grdBRUTACAB
        lngCol = conCOL_ProdBrutAcab_IDProd
     ElseIf stTabProd.Tab = 2 Then
        Set grdGENERICA = grdPRODINAT
        lngCol = conCOL_ProdFrabAcabInat_IDProd
     ElseIf stTabProd.Tab = 3 Then
        Set grdGENERICA = grdPRODAGLIB
        lngCol = conCOL_ProdFrabAgLib_IDProd
     End If
     
     bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADPROD'" & vbCrLf

     BRECATU.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BRECATU.EOF Then
        lngCODIGO = BRECATU!SGI_CODIGO
        strACAO = Trim(BRECATU!SGI_ACAO)
     End If
     BRECATU.Close
        
     With grdGENERICA
    
            i = .FindRow(lngCODIGO, , lngCol)
            If i > 0 Then
               If strACAO = "E" Then
                  If (.Rows - 1) = 1 Then .Rows = 1
                  If (.Rows - 1) > 1 Then .RemoveItem i
               ElseIf strACAO = "I" Or strACAO = "A" Then
                  bolAchou = True
               End If
            End If
        
        If BREC2.State = 1 Then BREC2.Close
        
        If bolAchou = False And strACAO = "I" Then
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRO.* " & vbCrLf
            If stTabProd.Tab <> 1 Then sSql = sSql & "      ,CLI.SGI_RAZAOSOC" & vbCrLf
            sSql = sSql & "      ,TIP.SGI_DESCRICAO as DESC_TIPO " & vbCrLf
            sSql = sSql & "      ,ESP.SGI_DESCRICAO as DESC_ESP " & vbCrLf
            
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO PRO " & vbCrLf
            sSql = sSql & "      ,SGI_CADESPPROD ESP " & vbCrLf
            sSql = sSql & "      ,SGI_CADTIPPROD TIP " & vbCrLf
            If stTabProd.Tab <> 1 Then sSql = sSql & "      ,SGI_CADCLIENTE CLI " & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       PRO.SGI_FILIAL        = " & FILIAL & vbCrLf
            sSql = sSql & "   And PRO.SGI_IDPRODUTO     = " & lngCODIGO & vbCrLf
           
            If stTabProd.Tab = 0 Or stTabProd.Tab = 2 Or stTabProd.Tab = 3 Then
               sSql = sSql & "   And PRO.SGI_PRODUTOTIPO   = 1" & vbCrLf
               sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
            ElseIf stTabProd.Tab = 1 Then
               sSql = sSql & "   And PRO.SGI_PRODUTOTIPO   = 0" & vbCrLf
               sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
            End If
            
            If stTabProd.Tab = 0 Then sSql = sSql & "   And PRO.SGI_STATUS = 1" & vbCrLf '' Ativo
            If stTabProd.Tab = 1 Then sSql = sSql & "   And PRO.SGI_STATUS = 1" & vbCrLf '' Ativo
            If stTabProd.Tab = 2 Then sSql = sSql & "   And PRO.SGI_STATUS = 0" & vbCrLf '' Inativo
            If stTabProd.Tab = 3 Then sSql = sSql & "   And PRO.SGI_STATUS = 2" & vbCrLf '' Aguardando Liberação
           
            sSql = sSql & "   And ESP.SGI_FILIAL        = PRO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And ESP.SGI_CODIGO        = PRO.SGI_CODESPECIE " & vbCrLf
            sSql = sSql & "   And TIP.SGI_FILIAL        = PRO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And TIP.SGI_CODIGO        = PRO.SGI_CODTIPO " & vbCrLf
           
            If stTabProd.Tab <> 1 Then
                sSql = sSql & "   And CLI.SGI_FILIAL        = PRO.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO        = PRO.SGI_CODCLIE"
            End If
           
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF Then
              
                 If BREC2!SGI_PRODUTOTIPO = 1 Or BREC2!SGI_PRODUTOTIPO = 2 Then
                 
                              .AddItem BREC2!SGI_IDPRODUTO & vbTab & _
                                       Format(BREC2!SGI_CODLINPROD, "###000") & "." & _
                                       Format(BREC2!SGI_CODCLIE, "####0000") & "." & _
                                       Format(BREC2!SGI_CODROTULO, "##00") & "." & _
                                       Format(IIf(IsNull(BREC2!SGI_DIGVERIF), 0, BREC2!SGI_DIGVERIF), "#0") & vbTab & _
                                       BREC2!SGI_DESCRICAO & vbTab & _
                                       BREC2!DESC_ESP & vbTab & _
                                       BREC2!DESC_TIPO & vbTab & _
                                       IIf(BREC2!SGI_STATUS = 1, "ATIVO", "DESATIVADO") & vbTab & _
                                       BREC2!SGI_STATUS
              
                 ElseIf BREC2!SGI_PRODUTOTIPO = 0 Then
                 
                              .AddItem BREC2!SGI_IDPRODUTO & vbTab & _
                                       BREC2!SGI_CODIGO & vbTab & _
                                       BREC2!SGI_DESCRICAO & vbTab & _
                                       BREC2!DESC_ESP & vbTab & _
                                       BREC2!DESC_TIPO & vbTab & _
                                       IIf(BREC2!SGI_STATUS = 1, "ATIVO", "DESATIVADO") & vbTab & _
                                       BREC2!SGI_STATUS
                 
                 End If
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And strACAO = "A" Then
          
           sSql = "Select " & vbCrLf
           sSql = sSql & "       PRO.* " & vbCrLf
           sSql = sSql & "      ,TIP.SGI_DESCRICAO as DESC_TIPO " & vbCrLf
           sSql = sSql & "      ,ESP.SGI_DESCRICAO as DESC_ESP " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADPRODUTO PRO " & vbCrLf
           sSql = sSql & "      ,SGI_CADESPPROD ESP " & vbCrLf
           sSql = sSql & "      ,SGI_CADTIPPROD TIP " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       PRO.SGI_FILIAL        = " & FILIAL & vbCrLf
           sSql = sSql & "   And PRO.SGI_IDPRODUTO     = " & lngCODIGO & vbCrLf
           
           If stTabProd.Tab = 0 Then sSql = sSql & "   And PRO.SGI_STATUS        = 1" & vbCrLf
           If stTabProd.Tab = 1 Then sSql = sSql & "   And PRO.SGI_STATUS        = 1" & vbCrLf
           If stTabProd.Tab = 2 Then sSql = sSql & "   And PRO.SGI_STATUS        = 0" & vbCrLf
           
           sSql = sSql & "   And ESP.SGI_FILIAL        = PRO.SGI_FILIAL " & vbCrLf
           sSql = sSql & "   And ESP.SGI_CODIGO        = PRO.SGI_CODESPECIE " & vbCrLf
           sSql = sSql & "   And TIP.SGI_FILIAL        = PRO.SGI_FILIAL " & vbCrLf
           sSql = sSql & "   And TIP.SGI_CODIGO        = PRO.SGI_CODTIPO " & vbCrLf
           
           If stTabProd.Tab = 0 Or stTabProd.Tab = 2 Then
              sSql = sSql & " Order by PRO.SGI_CODLINPROD " & vbCrLf
              sSql = sSql & "         ,PRO.SGI_CODCLIE " & vbCrLf
              sSql = sSql & "         ,PRO.SGI_CODROTULO " & vbCrLf
              sSql = sSql & "         ,PRO.SGI_DIGVERIF " & vbCrLf
           ElseIf stTabProd.Tab = 1 Then
              sSql = sSql & " Order by PRO.SGI_CODIGO " & vbCrLf
           End If
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
           
              If BREC2!SGI_PRODUTOTIPO = 1 Then
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_IDProd) = BREC2!SGI_IDPRODUTO
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_CodRotulo) = Format(BREC2!SGI_CODLINPROD, "###000") & "." & Format(BREC2!SGI_CODCLIE, "####0000") & "." & Format(BREC2!SGI_CODROTULO, "##00") & "." & Format(IIf(IsNull(BREC2!SGI_DIGVERIF), 0, BREC2!SGI_DIGVERIF), "#0")
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_Descricao) = BREC2!SGI_DESCRICAO
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_Especie) = BREC2!DESC_ESP
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_Complemento) = BREC2!DESC_TIPO
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_Status) = IIf(BREC2!SGI_STATUS = 1, "ATIVO", "DESATIVADO")
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_CodStatus) = BREC2!SGI_STATUS
              ElseIf BREC2!SGI_PRODUTOTIPO = 0 Then
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_IDProd) = BREC2!SGI_IDPRODUTO
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_CodRotulo) = BREC2!SGI_CODIGO
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_Descricao) = BREC2!SGI_DESCRICAO
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_Especie) = BREC2!DESC_ESP
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_Complemento) = BREC2!DESC_TIPO
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_Status) = IIf(BREC2!SGI_STATUS = 1, "ATIVO", "DESATIVADO")
                 .Cell(flexcpText, i, conCOL_ProdFrabAcab_CodStatus) = BREC2!SGI_STATUS
              End If
           End If
           BREC2.Close
        End If
     End With
      
End Sub


Private Function Verif_reg() As Boolean

    Verif_reg = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & objCADPRODUTO.IDProduto
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Este registro foi excluso !!!", vbOKOnly + vbExclamation, "Aviso"
       Verif_reg = True
    End If
    BREC.Close

End Function

Private Sub Concerta_Unid_Cons()

    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       LST.SGI_PRODLST  " & vbCrLf
    sSql = sSql & "      ,UNI.SGI_UNIDADE  " & vbCrLf
    sSql = sSql & "      ,LST.SGI_UNIDCONS " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMAT   LST  " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PROD " & vbCrLf
    sSql = sSql & "      ,SGI_CADUNIMED  UNI  " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL = LST.SGI_FILIAL  " & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODIGO = LST.SGI_PRODLST " & vbCrLf
    sSql = sSql & "   And UNI.SGI_FILIAL  = PROD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And UNI.SGI_CODIGO  = PROD.SGI_UNIDMEDIDA " & vbCrLf
    sSql = sSql & " Order by SGI_PRODUTO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       
       If Len(Trim(BREC!SGI_UNIDCONS)) = 0 Or IsNull(BREC!SGI_UNIDCONS) = True Then
          
          sSql = "Update SGI_LISTAMAT Set " & vbCrLf
          sSql = sSql & "                        SGI_UNIDCONS = '" & Trim(BREC!SGI_UNIDADE) & "'" & vbCrLf
          sSql = sSql & "                  Where " & vbCrLf
          sSql = sSql & "                        SGI_FILIAL     =  " & FILIAL & vbCrLf
          sSql = sSql & "                    And SGI_PRODLST    = '" & Trim(BREC!SGI_PRODLST) & "'"
          
          BGRV.CommandText = sSql
          BGRV.Execute
          
       End If
       
       BREC.MoveNext
    Loop
    BREC.Close
    
    adoBanco_Dados.CommitTrans

End Sub


Private Sub ConfGridDiversos()
        
    With grdBRUTACAB
    
       .Cols = conColumnsIn_ProdBrutAcab
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdBrutAcab_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_ProdBrutAcab_IDProd) = ""
       .ColDataType(conCOL_ProdBrutAcab_IDProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdBrutAcab_CodRotulo) = ""
       .ColDataType(conCOL_ProdBrutAcab_CodRotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdBrutAcab_Descricao) = ""
       .ColDataType(conCOL_ProdBrutAcab_Descricao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdBrutAcab_Especie) = ""
       .ColDataType(conCOL_ProdBrutAcab_Especie) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdBrutAcab_Complemento) = ""
       .ColDataType(conCOL_ProdBrutAcab_Complemento) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdBrutAcab_Status) = ""
       .ColDataType(conCOL_ProdBrutAcab_Status) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdBrutAcab_CodStatus) = ""
       .ColDataType(conCOL_ProdBrutAcab_CodStatus) = flexDTLong
       
       .ColWidth(conCOL_ProdBrutAcab_IDProd) = 0
       .ColWidth(conCOL_ProdBrutAcab_CodRotulo) = 1500
       .ColWidth(conCOL_ProdBrutAcab_Descricao) = 5000
       .ColWidth(conCOL_ProdBrutAcab_Especie) = 2500
       .ColWidth(conCOL_ProdBrutAcab_Complemento) = 2500
       .ColWidth(conCOL_ProdBrutAcab_Status) = 1000
       .ColWidth(conCOL_ProdBrutAcab_CodStatus) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With


End Sub


Private Sub PreencheGridDiversos()

    If BREC.State = 1 Then BREC.Close
    
    With grdBRUTACAB
    
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       PRO.* " & vbCrLf
        sSql = sSql & "      ,TIP.SGI_DESCRICAO as DESC_TIPO " & vbCrLf
        sSql = sSql & "      ,ESP.SGI_DESCRICAO as DESC_ESP " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO PRO " & vbCrLf
        sSql = sSql & "      ,SGI_CADESPPROD ESP " & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TIP " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       PRO.SGI_FILIAL        = " & FILIAL & vbCrLf
        sSql = sSql & "   And PRO.SGI_PRODUTOTIPO   = 0" & vbCrLf
        sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
        sSql = sSql & "   And ESP.SGI_FILIAL        = PRO.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ESP.SGI_CODIGO        = PRO.SGI_CODESPECIE " & vbCrLf
        sSql = sSql & "   And TIP.SGI_FILIAL        = PRO.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And TIP.SGI_CODIGO        = PRO.SGI_CODTIPO " & vbCrLf
        sSql = sSql & " Order by PRO.SGI_CODIGO " & vbCrLf
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        Do While Not BREC.EOF
                    .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                             BREC!SGI_CODIGO & vbTab & _
                             BREC!SGI_DESCRICAO & vbTab & _
                             BREC!DESC_ESP & vbTab & _
                             BREC!DESC_TIPO & vbTab & _
                             IIf(BREC!SGI_STATUS = 1, "ATIVO", "DESATIVADO") & vbTab & _
                             BREC!SGI_STATUS
                              
           BREC.MoveNext
        Loop
        BREC.Close
    
    End With
    
End Sub

Private Function VerifArvProd() As String

    VerifArvProd = "I"
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & objCADPRODUTO.CodigoProd & "'" & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL     " & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST    " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then VerifArvProd = "C"
    BREC.Close

End Function

Private Sub DestroiObjeto()
    Set objFuncoes = Nothing
    Set objCADPRODUTO = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub ConfGridInat()
    
    With grdPRODINAT
    
       .Cols = conColumnsIn_ProdFrabAcabInat
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdFrabAcabInat_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcabInat_IDProd) = ""
       .ColDataType(conCOL_ProdFrabAcabInat_IDProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcabInat_CodRotulo) = ""
       .ColDataType(conCOL_ProdFrabAcabInat_CodRotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcabInat_Descricao) = ""
       .ColDataType(conCOL_ProdFrabAcabInat_Descricao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcabInat_Especie) = ""
       .ColDataType(conCOL_ProdFrabAcabInat_Especie) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcabInat_Complemento) = ""
       .ColDataType(conCOL_ProdFrabAcabInat_Complemento) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcabInat_Status) = ""
       .ColDataType(conCOL_ProdFrabAcabInat_Status) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAcabInat_CodStatus) = ""
       .ColDataType(conCOL_ProdFrabAcabInat_CodStatus) = flexDTLong
       
       .ColWidth(conCOL_ProdFrabAcabInat_IDProd) = 0
       .ColWidth(conCOL_ProdFrabAcabInat_CodRotulo) = 1500
       .ColWidth(conCOL_ProdFrabAcabInat_Descricao) = 5000
       .ColWidth(conCOL_ProdFrabAcabInat_Especie) = 2500
       .ColWidth(conCOL_ProdFrabAcabInat_Complemento) = 2500
       .ColWidth(conCOL_ProdFrabAcabInat_Status) = 1000
       .ColWidth(conCOL_ProdFrabAcabInat_CodStatus) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub

Private Function Verifica_Se_TemPedido() As Boolean

    Verifica_Se_TemPedido = False

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITEN.*" & vbCrLf
    sSql = sSql & "     ,CABE.SGI_STATUS" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADPEDVENDI_STEEL ITEN" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDH_STEEL CABE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITEN.SGI_IDPRODUTO = " & objCADPRODUTO.IDProduto & vbCrLf
    sSql = sSql & "  And ITEN.SGI_FILIAL    = CABE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And ITEN.SGI_CODIGO    = CABE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  And CABE.SGI_STATUS    <> 'F'" & vbCrLf
    sSql = sSql & "  And CABE.SGI_STATUS    <> 'M'"

    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC7.EOF() Then Verifica_Se_TemPedido = True
    BREC7.Close
    
    If Verifica_Se_TemPedido = True Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITEN.*" & vbCrLf
    sSql = sSql & "     ,CABE.SGI_STATUS" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADPEDVENDI ITEN" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDH CABE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITEN.SGI_IDPRODUTO = " & objCADPRODUTO.IDProduto & vbCrLf
    sSql = sSql & "  And ITEN.SGI_FILIAL    = CABE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And ITEN.SGI_CODIGO    = CABE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  And CABE.SGI_STATUS    <> 'F'" & vbCrLf
    sSql = sSql & "  And CABE.SGI_STATUS    <> 'M'"

    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC7.EOF() Then Verifica_Se_TemPedido = True
    BREC7.Close
    
    If Verifica_Se_TemPedido = True Then Exit Function
    
End Function

Private Sub ConfGridAgLib()
    
    With grdPRODAGLIB
    
       .Cols = conColumnsIn_ProdFrabAgLib
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdFrabAgLib_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAgLib_IDProd) = ""
       .ColDataType(conCOL_ProdFrabAgLib_IDProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAgLib_CodRotulo) = ""
       .ColDataType(conCOL_ProdFrabAgLib_CodRotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAgLib_Descricao) = ""
       .ColDataType(conCOL_ProdFrabAgLib_Descricao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAgLib_Especie) = ""
       .ColDataType(conCOL_ProdFrabAgLib_Especie) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAgLib_Complemento) = ""
       .ColDataType(conCOL_ProdFrabAgLib_Complemento) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAgLib_Status) = ""
       .ColDataType(conCOL_ProdFrabAgLib_Status) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdFrabAgLib_CodStatus) = ""
       .ColDataType(conCOL_ProdFrabAgLib_CodStatus) = flexDTLong
       
       .ColWidth(conCOL_ProdFrabAgLib_IDProd) = 0
       .ColWidth(conCOL_ProdFrabAgLib_CodRotulo) = 1500
       .ColWidth(conCOL_ProdFrabAgLib_Descricao) = 5000
       .ColWidth(conCOL_ProdFrabAgLib_Especie) = 2500
       .ColWidth(conCOL_ProdFrabAgLib_Complemento) = 2500
       .ColWidth(conCOL_ProdFrabAgLib_Status) = 1000
       .ColWidth(conCOL_ProdFrabAgLib_CodStatus) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub


Private Function ConsisteCampos() As Boolean

    ConsisteCampos = False
    
    Dim intVERNIZ   As Integer
    Dim intLINHA    As Integer
    Dim arrDADOS    As Variant
    Dim intFECHA    As Integer
    
    Const conCOL_SomFecha_TampaPressao             As Integer = 0
    Const conCOL_SomFecha_BatoqueRetra             As Integer = 1
    Const conCOL_SomFecha_BatoquePlast             As Integer = 2
    Const conCOL_SomFecha_TampaVisor               As Integer = 3
    
    objCADPRODUTO.TampaPressao = conCOL_SomFecha_TampaPressao
    objCADPRODUTO.BatoqueRetratil = conCOL_SomFecha_BatoqueRetra
    objCADPRODUTO.BatoquePlastico = conCOL_SomFecha_BatoquePlast
    objCADPRODUTO.TAMPAVIS = conCOL_SomFecha_TampaVisor
    
    If objCADPRODUTO.Carrega_campos = False Then Exit Function
    
    If objCADPRODUTO.CodLinProd = 0 Then Exit Function
    If objCADPRODUTO.CodClie = 0 Then Exit Function
    
    If objCADPRODUTO.NATURALSIMNAO = 0 Then
        If objCADPRODUTO.CodRotulo = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "O Código do Rótulo não pode ser 0 !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End If
    
    If Len(Trim(objCADPRODUTO.DescriProd)) = 0 Then Exit Function
    If objCADPRODUTO.CODGRUPPROD = 0 Then Exit Function
    If objCADPRODUTO.CODSUBGPROD = 0 Then Exit Function
    If objCADPRODUTO.EspProduto = 0 Then Exit Function
    If objCADPRODUTO.Unidade = 0 Then Exit Function
    
    If objCADPRODUTO.NATURALSIMNAO = 0 Then '' Natural 1 = Sim / 0 = Não
        If Not IsArray(objCADPRODUTO.CORES) Then
            MsgBox "ATENÇÃO - Não foi informado as Cores !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
''        Else
''            If UBound(objCADPRODUTO.CORES) = 0 Then
''                MsgBox "ATENÇÃO - Não foi informado as Cores !!!", vbOKOnly + vbExclamation, "Aviso"
''                Exit Function
''            End If
        End If
    End If
        
    intVERNIZ = 0
    If IsArray(objCADPRODUTO.VERNIZ) Then intVERNIZ = (intVERNIZ + 1)
    If IsArray(objCADPRODUTO.VERNIZ02) Then intVERNIZ = (intVERNIZ + 1)
    If IsArray(objCADPRODUTO.VERNIZACAB) Then intVERNIZ = (intVERNIZ + 1)
    If IsArray(objCADPRODUTO.ESMALTE) Then intVERNIZ = (intVERNIZ + 1)
    
    If intVERNIZ = 0 Then
        MsgBox "ATENÇÂO - Não foi Informado nenhum Verniz Interno ou Verniz de Acabamento ou Emalte !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
    '' Folha de Flandres
    If objCADPRODUTO.VernCorpo = 0 Then
        MsgBox "ATENÇÂO - Verniz do Corpo não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    Else
        If objCADPRODUTO.EspessCorpo = 0 Then
            MsgBox "ATENÇÂO - A Espessura do Corpo não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If objCADPRODUTO.RevestCorpo = 0 Then
            MsgBox "ATENÇÂO - Revestimento do Corpo não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If objCADPRODUTO.RevestCorpo2 = 0 Then
            MsgBox "ATENÇÂO - Revestimento do Corpo não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End If
    If objCADPRODUTO.VernTampa = 0 Then
        MsgBox "ATENÇÂO - Verniz da Tampa não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    Else
        If objCADPRODUTO.EspessTampa = 0 Then
            MsgBox "ATENÇÂO - A Espessura da Tampa não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If objCADPRODUTO.RevestTampa = 0 Then
            MsgBox "ATENÇÂO - O Revestimento da Tampa não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If objCADPRODUTO.RevestTampa2 = 0 Then
            MsgBox "ATENÇÂO - O Revestimento da Tampa não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End If
    If objCADPRODUTO.VernFundo = 0 Then
        MsgBox "ATENÇÂO - Verniz do Fundo não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    Else
        If objCADPRODUTO.EspessFundo = 0 Then
            MsgBox "ATENÇÂO - A Espessura do Fundo não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If objCADPRODUTO.RevestFundo = 0 Then
            MsgBox "ATENÇÂO - O Revestimento do Fundo não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If objCADPRODUTO.RevestFundo2 = 0 Then
            MsgBox "ATENÇÂO - O Revestimento do Fundo não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End If
        
    '' Quantidade por Folha
    If objCADPRODUTO.QTDCORPSPADRAOSN = 0 Then '' Padrão = 0 - Não
        If objCADPRODUTO.QTDEPORFOLHA = 0 Then
            MsgBox "ATENÇÂO - A Quantidade por Folha não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End If
    
    If objCADPRODUTO.QTDPASSADAS = 0 Then
        MsgBox "ATENÇÂO - A Quantidade de Passadas na Litografia não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
    '' Dimensões para Corte
    If objCADPRODUTO.DIMPADRAO = 0 Then
        If Len(Trim(objCADPRODUTO.DESENV)) = 0 Then
            MsgBox "ATENÇÂO - A Espessura da Dimensão de Corte não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If Len(Trim(objCADPRODUTO.ALTURA)) = 0 Then
            MsgBox "ATENÇÂO - A Altura da Dimensão de Corte não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End If
    
    
    '' Aba Fechamento
    If objCADPRODUTO.FechSoldaAgrafado = -1 Then
        MsgBox "ATENÇÂO - O Fechamento não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    If objCADPRODUTO.FechTampaFuro = -1 Then
        MsgBox "ATENÇÂO - O Fechamento Tampa/Furo de não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
    intFECHA = 0
    arrDADOS = objCADPRODUTO.TAMPAPRESS
    If IsArray(arrDADOS) Then
        For intLINHA = 1 To UBound(arrDADOS)
            If arrDADOS(intLINHA) = 0 Then
                MsgBox "ATENÇÂO - A Tampa de Pressão não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        Next intLINHA
        intFECHA = (intFECHA + 1)
    End If
    
    arrDADOS = objCADPRODUTO.BATRETRATI
    If IsArray(arrDADOS) Then
        For intLINHA = 1 To UBound(arrDADOS)
            If arrDADOS(intLINHA) = 0 Then
                MsgBox "ATENÇÂO - O Batoque Retratil não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        Next intLINHA
        intFECHA = (intFECHA + 1)
    End If
    
    arrDADOS = objCADPRODUTO.BATPLASTIC
    If IsArray(arrDADOS) Then
        For intLINHA = 1 To UBound(arrDADOS)
            If arrDADOS(intLINHA) = 0 Then
                MsgBox "ATENÇÂO - O Batoque de Plático não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        Next intLINHA
        intFECHA = (intFECHA + 1)
    End If
    
    arrDADOS = objCADPRODUTO.TAMPAVISOR
    If IsArray(arrDADOS) Then
        For intLINHA = 1 To UBound(arrDADOS)
            If arrDADOS(intLINHA) = 0 Then
                MsgBox "ATENÇÂO - A Tampa do Visor não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        Next intLINHA
        intFECHA = (intFECHA + 1)
    End If
    If intFECHA = 0 Then
        MsgBox "ATENÇÂO - Não Foi Informado Nenhum Tipo de Tampa de Pressão/Batoque Retratil/Batoque Plástico/Tampa Visor !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
''    If Not IsArray(objCADPRODUTO.VEDANTE) Then
''        MsgBox "ATENÇÂO - O Vedante não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
''        Exit Function
''    End If
    
    ConsisteCampos = True

End Function

Private Sub AtivaDesativaBotoes()

On Error GoTo Err_AtivaDesativaBotoes
    
    If lngCodUsuario = 0 Then Exit Sub
    
    cmdAtivDesat.Visible = PermiteLibStatus
    Command6.Visible = PermiteLibFotolito
    
    Exit Sub
    
Err_AtivaDesativaBotoes:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "A", "Função : AtivaDesativaBotoes()", Me.Name, "AtivaDesativaBotoes()", strCAMARQERRO)
    
End Sub

Private Function PermiteLibStatus() As Boolean

    PermiteLibStatus = False
    
    If lngCodUsuario = 0 Then
       PermiteLibStatus = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_DESABPROD" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCodUsuario & vbCrLf
    sSql = sSql & "   And SGI_DESABPROD = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibStatus = True
    BREC.Close

End Function

Private Function PermiteLibFotolito() As Boolean

    PermiteLibFotolito = False
    
    If lngCodUsuario = 0 Then
       PermiteLibFotolito = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_PERMLIBFOT" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO     = " & lngCodUsuario & vbCrLf
    sSql = sSql & "   And SGI_PERMLIBFOT = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibFotolito = True
    BREC.Close

End Function


Private Sub txtCIDCLIE_GotFocus()
    objFuncoes.SelecionaCampos txtCIDCLIE.Name, Me
End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCIDCLIE.Text
End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

On Error GoTo Err_txtCIDCLIE_Validate

    Dim i As Integer
    
    If Len(Trim(txtCIDCLIE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCIDCLIE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, txtNomClie)
    If Len(Trim(txtNomClie.Text)) = 0 Then
       txtCIDCLIE.Text = ""
       txtNomClie.Text = ""
       Cancel = True
       Exit Sub
    End If

    Exit Sub
    
Err_txtCIDCLIE_Validate:
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : txtCIDCLIE_Validate()", Me.Name, "txtCIDCLIE_Validate()", strCAMARQERRO)

End Sub

Private Sub txtCodigio_GotFocus()
    objFuncoes.SelecionaCampos txtCodigio.Name, Me
End Sub

Private Sub txtCodigio_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtComplemento_GotFocus()
    objFuncoes.SelecionaCampos txtComplemento.Name, Me
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtDescricao_GotFocus()
    objFuncoes.SelecionaCampos txtDescricao.Name, Me
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As TextBox)

On Error GoTo Err_PegaDescTabelas

    lblLabel.Text = ""
    
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
       lblLabel.Text = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
    Exit Sub
    
Err_PegaDescTabelas:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub


Private Sub txtLinProd_GotFocus()
    objFuncoes.SelecionaCampos txtLinProd.Name, Me
End Sub

Private Sub txtLinProd_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCIDCLIE.Text
End Sub

Private Sub txtLinProd_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtLinProd.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtLinProd.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtLinProd.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
    If Len(Trim(lblLinhProd.Caption)) = 0 Then
       MsgBox "Linha de Produto não cadastrada !!!", vbOKOnly + vbExclamation, "Aviso"
       txtLinProd.Text = ""
       Cancel = True
       Exit Sub
    End If
    


End Sub

Private Sub txtNomClie_GotFocus()
    objFuncoes.SelecionaCampos txtNomClie.Name, Me
End Sub
Private Function PegaDescLinProd(lngCodLinProd As Long) As String

    PegaDescLinProd = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "    And SGI_CODLIN = " & lngCodLinProd
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescLinProd = BREC!SGI_DESCRI
    BREC.Close
    
End Function


Private Sub LimpaCamposLabel()
    lblLinhProd.Caption = ""
End Sub
