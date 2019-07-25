VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADLINHAPROD 
   Caption         =   "Cadastro de Linha de Produto"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   16065
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados da Linha"
      TabPicture(0)   =   "frmCADLINHAPROD.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cortes"
      TabPicture(1)   =   "frmCADLINHAPROD.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Folhas Usadas na Capacidade"
      TabPicture(2)   =   "frmCADLINHAPROD.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Fechamento"
      TabPicture(3)   =   "frmCADLINHAPROD.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Armazenamento"
      TabPicture(4)   =   "frmCADLINHAPROD.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdARMZLIN"
      Tab(4).Control(1)=   "Command5"
      Tab(4).Control(2)=   "Command6"
      Tab(4).ControlCount=   3
      Begin VB.CommandButton Command6 
         Height          =   300
         Left            =   -59520
         Picture         =   "frmCADLINHAPROD.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Exclui a linha da Gride Selecionada"
         Top             =   780
         Width           =   300
      End
      Begin VB.CommandButton Command5 
         Height          =   300
         Left            =   -59520
         Picture         =   "frmCADLINHAPROD.frx":01D6
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   420
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdARMZLIN 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   41
         Top             =   420
         Width           =   15255
         _cx             =   26908
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
      Begin VB.Frame Frame9 
         Caption         =   "[ Tipo ]"
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
         Height          =   2535
         Left            =   -70560
         TabIndex        =   39
         Top             =   420
         Width           =   4575
         Begin VB.ListBox lstFech 
            Height          =   2085
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   40
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "[ Tampa/Furo ]"
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
         Height          =   4455
         Left            =   -74880
         TabIndex        =   37
         Top             =   420
         Width           =   4215
         Begin VB.ListBox lstTpFr 
            Height          =   4110
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   38
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "[ Produtos Usados ]"
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
         Height          =   6255
         Left            =   -74880
         TabIndex        =   33
         Top             =   420
         Width           =   15735
         Begin VB.CommandButton Command4 
            Height          =   300
            Left            =   15360
            Picture         =   "frmCADLINHAPROD.frx":0320
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Inclui uma nova linha na Gride"
            Top             =   240
            Width           =   300
         End
         Begin VB.CommandButton Command3 
            Height          =   300
            Left            =   15360
            Picture         =   "frmCADLINHAPROD.frx":046A
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Exclui a linha da Gride Selecionada"
            Top             =   600
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid grdPRODUTOS 
            Height          =   5895
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   15135
            _cx             =   26696
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
      Begin VB.Frame Frame6 
         Caption         =   "[ Dimensões ]"
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
         Height          =   6255
         Left            =   -72720
         TabIndex        =   27
         Top             =   420
         Width           =   13575
         Begin VB.CommandButton Command26 
            Height          =   300
            Left            =   13080
            Picture         =   "frmCADLINHAPROD.frx":05B4
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Exclui a linha da Gride Selecionada"
            Top             =   600
            Width           =   300
         End
         Begin VB.CommandButton Command27 
            Height          =   300
            Left            =   13080
            Picture         =   "frmCADLINHAPROD.frx":06FE
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Inclui uma nova linha na Gride"
            Top             =   240
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid grdCORTES 
            Height          =   5895
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   12855
            _cx             =   22675
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
      Begin VB.Frame Frame5 
         Caption         =   "[ Cortes ]"
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
         Height          =   6255
         Left            =   -74880
         TabIndex        =   23
         Top             =   420
         Width           =   2175
         Begin VB.CommandButton Command2 
            Height          =   300
            Left            =   1800
            Picture         =   "frmCADLINHAPROD.frx":0848
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Exclui a linha da Gride Selecionada"
            Top             =   600
            Width           =   300
         End
         Begin VB.CommandButton Command1 
            Height          =   300
            Left            =   1800
            Picture         =   "frmCADLINHAPROD.frx":0992
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Inclui uma nova linha na Gride"
            Top             =   240
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid grdITECORTES 
            Height          =   5895
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1575
            _cx             =   2778
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
      Begin VB.Frame Frame2 
         Height          =   6375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   15735
         Begin VB.TextBox txtQTDECORPOS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   32
            Text            =   "txtQTDECORPOS"
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   17
            Text            =   "txtCodigo"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   16
            Text            =   "txtDescricao"
            Top             =   1080
            Width           =   5055
         End
         Begin VB.TextBox txtCodLin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   15
            Text            =   "txtCodLin"
            Top             =   720
            Width           =   855
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1200
            TabIndex        =   12
            Top             =   1440
            Width           =   3975
            Begin VB.OptionButton optFILIALPED 
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
               TabIndex        =   14
               Top             =   0
               Width           =   1575
            End
            Begin VB.OptionButton optFILIALPED 
               Caption         =   "STEEL ROLL"
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
               TabIndex        =   13
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Dimensão para corte ]"
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
            Height          =   975
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Width           =   11775
            Begin VB.TextBox txtDESENV 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2040
               TabIndex        =   9
               Text            =   "txtDESENV"
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txtALTURA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2040
               TabIndex        =   8
               Text            =   "txtALTURA"
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Desenvolvimento"
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
               TabIndex        =   11
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "Altura"
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
               TabIndex        =   10
               Top             =   600
               Width           =   1575
            End
         End
         Begin VB.TextBox txtPERDPROC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   6
            Text            =   "txtPERDPROC"
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade de Corpos"
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
            TabIndex        =   31
            Top             =   3120
            Width           =   1905
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sequencial"
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
            TabIndex        =   22
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
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
            TabIndex        =   21
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            TabIndex        =   20
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Filial"
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
            TabIndex        =   19
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Label5 
            Caption         =   "Perda no Processo"
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
            TabIndex        =   18
            Top             =   2760
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15975
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
         Picture         =   "frmCADLINHAPROD.frx":0ADC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
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
         MaskColor       =   &H8000000F&
         Picture         =   "frmCADLINHAPROD.frx":100E
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADLINHAPROD.frx":1110
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADLINHAPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public cTipOper     As String
Public iCodigo      As Integer
Public FILIAL       As Integer
Public strAcesso    As String
Public strMODPAI    As String
Public strUSUARIO   As String

Dim objBLBFunc      As Object
Dim objCADINHAPROD  As Object
Dim objPESQPADRAO   As Object
Dim arrDIMCORTE     As Variant
Dim arrSEQCORTE     As Variant
Dim arrPRODUTOS     As Variant
Dim arrFECHAMENTOS  As Variant
Dim arrFECHTPFR     As Variant
Dim arrARMAZ        As Variant

Const conCOL_SonCort_Itens                      As Integer = 0
Const conCOL_SonCort_ItensBKP                   As Integer = 1
Const conCOL_SonCort_FormatString               As String = "=Sequência|BKP"
Const conColumnsIn_SonCort                      As Integer = 2

Const conCOL_SonDim_Itens                       As Integer = 0
Const conCOL_SonDim_CodDim                      As Integer = 1
Const conCOL_SonDim_PesqDim                     As Integer = 2
Const conCOL_SonDim_DescDim                     As Integer = 3
Const conCOL_SonDim_INDICE                      As Integer = 4
Const conCOL_SonDim_PADRAO                      As Integer = 5
Const conCOL_SonDim_EXPESS                      As Integer = 6
Const conCOL_SonDim_LAGURA                      As Integer = 7
Const conCOL_SonDim_COMPRI                      As Integer = 8
Const conCOL_SonDim_QTDECORPOS                  As Integer = 9
Const conCOL_SonDim_PERDPROC                    As Integer = 10
Const conCOL_SonDim_FormatString                As String = "=Seq.Corte|Código|...|Descrição das Dimensões|INDICE|Padrão|Expessura|Largura|Comprimento|Qtde Corpos|Perda no Processo"
Const conColumnsIn_SonDim                       As Integer = 11

Const conCOL_SonProd_CodID                      As Integer = 0
Const conCOL_SonProd_CodProd                    As Integer = 1
Const conCOL_SonProd_PesqProd                   As Integer = 2
Const conCOL_SonProd_DescProd                   As Integer = 3
Const conCOL_SonProd_FormatString               As String = "=IDProd|Código|...|Descrição do Produto"
Const conColumnsIn_SonProd                      As Integer = 4

Const conCOL_SonArm_Qtdelatas                   As Integer = 0
Const conCOL_SonArm_FormatString                As String = "=Qtde. Latas/Palhet"
Const conColumnsIn_SonArm                       As Integer = 1

Private Sub cmdAltera_Click()

    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    Frame2.Enabled = True
    
    txtDescricao.SetFocus
    
    cTipOper = "A"
    
    Me.Caption = "Cadastro de Linha de Produto - [ ALTERAÇÃO ]"

End Sub

Private Sub CmdSalva_Click()

    Dim i           As Integer
    Dim strValor    As String
    Dim intQTDREGS  As Integer
    
    If Verifica_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADINHAPROD.CODIGO = objCADINHAPROD.Gera_Codigo(Me.Name)
    
    objCADINHAPROD.DESCRI = txtDescricao.Text
    objCADINHAPROD.CODIGOLIN = CLng(txtCodLin.Text)
    
    If optFILIALPED(0).Value = True Then objCADINHAPROD.FILIALPED = 0 '' Novalata
    If optFILIALPED(1).Value = True Then objCADINHAPROD.FILIALPED = 1 '' Stil Row
    
    strValor = "Null"
    If Len(Trim(txtDESENV.Text)) > 0 Then
        strValor = Replace(Format(txtDESENV.Text, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
    End If
    objCADINHAPROD.DESENV = strValor
    
    strValor = "Null"
    If Len(Trim(txtALTURA.Text)) > 0 Then
        strValor = Replace(Format(txtALTURA.Text, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
    End If
    objCADINHAPROD.ALTURA = strValor
    
    strValor = "Null"
    If Len(Trim(txtPERDPROC.Text)) > 0 Then
        strValor = Replace(Format(txtPERDPROC.Text, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
    End If
    objCADINHAPROD.PERDPROC = strValor
    
    
    objCADINHAPROD.QTDECORPOS = "Null"
    If Len(Trim(txtQTDECORPOS.Text)) > 0 Then objCADINHAPROD.QTDECORPOS = txtQTDECORPOS.Text
    
    ''---------------------------------------
    '' Seguência de Corte
    arrSEQCORTE = Empty
    With grdITECORTES
        If (.Rows - 1) > 0 Then
            ReDim arrSEQCORTE(1 To (.Rows - 1), 1 To 1) As String
            For i = 1 To (.Rows - 1)
                arrSEQCORTE(i, 1) = Trim(.Cell(flexcpText, i, conCOL_SonCort_Itens))
            Next i
        End If
    End With
    objCADINHAPROD.SEQCORTE = arrSEQCORTE
    ''---------------------------------------
    
    
    ''---------------------------------------
    '' Dimensões de Corte
    arrDIMCORTE = Empty
    With grdCORTES
        If (.Rows - 1) > 0 Then
            ReDim arrDIMCORTE(1 To (.Rows - 1), 1 To 9) As String
            For i = 1 To (.Rows - 1)
                arrDIMCORTE(i, 1) = .Cell(flexcpText, i, conCOL_SonDim_Itens)
                arrDIMCORTE(i, 2) = .Cell(flexcpText, i, conCOL_SonDim_CodDim)
                arrDIMCORTE(i, 3) = .Cell(flexcpText, i, conCOL_SonDim_INDICE)
                arrDIMCORTE(i, 4) = .Cell(flexcpText, i, conCOL_SonDim_PADRAO)
                
                strValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonDim_EXPESS))) > 0 Then
                    strValor = Replace(.Cell(flexcpText, i, conCOL_SonDim_EXPESS), ".", "")
                    strValor = Replace(strValor, ",", ".")
                End If
                arrDIMCORTE(i, 5) = strValor
                
                strValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonDim_LAGURA))) > 0 Then
                    strValor = Replace(.Cell(flexcpText, i, conCOL_SonDim_LAGURA), ".", "")
                    strValor = Replace(strValor, ",", ".")
                End If
                arrDIMCORTE(i, 6) = strValor
                
                strValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonDim_COMPRI))) > 0 Then
                    strValor = Replace(.Cell(flexcpText, i, conCOL_SonDim_COMPRI), ".", "")
                    strValor = Replace(strValor, ",", ".")
                End If
                arrDIMCORTE(i, 7) = strValor
            
                strValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonDim_QTDECORPOS))) > 0 Then
                    strValor = Replace(.Cell(flexcpText, i, conCOL_SonDim_QTDECORPOS), ".", "")
                    strValor = Replace(strValor, ",", ".")
                End If
                arrDIMCORTE(i, 8) = strValor
            
                strValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonDim_PERDPROC))) > 0 Then
                    strValor = Replace(.Cell(flexcpText, i, conCOL_SonDim_PERDPROC), ".", "")
                    strValor = Replace(strValor, ",", ".")
                End If
                arrDIMCORTE(i, 9) = strValor
            
            Next i
        End If
    End With
    objCADINHAPROD.DIMCORTE = arrDIMCORTE
    ''---------------------------------------
    
    ''---------------------------------------
    '' Produtos
    Call objBLBFunc.RemoveLinhaVazia(grdPRODUTOS, conCOL_SonProd_CodID)
    arrPRODUTOS = Empty
    With grdPRODUTOS
        If (.Rows - 1) > 0 Then
            ReDim arrPRODUTOS(1 To (.Rows - 1), 1 To 1) As String
            For i = 1 To (.Rows - 1)
                arrPRODUTOS(i, 1) = .Cell(flexcpText, i, conCOL_SonProd_CodID)
            Next i
        End If
    End With
    objCADINHAPROD.PRODUTOS = arrPRODUTOS
    ''---------------------------------------
    
    '' Fechamentos Tipo
    arrFECHAMENTOS = Empty
    With lstFech
        intQTDREGS = 0
        For i = 0 To (.ListCount - 1)
            If .Selected(i) = True Then intQTDREGS = (intQTDREGS + 1)
        Next i
        
        If intQTDREGS > 0 Then
            ReDim arrFECHAMENTOS(1 To intQTDREGS) As String
            intQTDREGS = 0
            For i = 0 To (.ListCount - 1)
                If .Selected(i) = True Then
                   intQTDREGS = (intQTDREGS + 1)
                   arrFECHAMENTOS(intQTDREGS) = .ItemData(i)
                End If
            Next i
        End If
    End With
    objCADINHAPROD.FECHAMENTOS = arrFECHAMENTOS
    ''---------------------------------------
    
    
    '' Fechamentos Tampa Furo
    arrFECHTPFR = Empty
    With lstTpFr
        intQTDREGS = 0
        For i = 0 To (.ListCount - 1)
            If .Selected(i) = True Then intQTDREGS = (intQTDREGS + 1)
        Next i
        
        If intQTDREGS > 0 Then
            ReDim arrFECHTPFR(1 To intQTDREGS) As String
            intQTDREGS = 0
            For i = 0 To (.ListCount - 1)
                If .Selected(i) = True Then
                   intQTDREGS = (intQTDREGS + 1)
                   arrFECHTPFR(intQTDREGS) = .ItemData(i)
                End If
            Next i
        End If
    End With
    objCADINHAPROD.FECHTPFR = arrFECHTPFR
    ''---------------------------------------
    
    ''---------------------------------------
    Call objBLBFunc.RemoveLinhaVazia(grdARMZLIN, conCOL_SonArm_Qtdelatas)
    arrARMAZ = Empty
    With grdARMZLIN
        If (.Rows - 1) > 0 Then
            ReDim arrARMAZ(1 To (.Rows - 1)) As String
            For i = 1 To (.Rows - 1)
                arrARMAZ(i) = .Cell(flexcpText, i, conCOL_SonArm_Qtdelatas)
            Next i
        End If
    End With
    objCADINHAPROD.ARMAZ = arrARMAZ
    ''---------------------------------------
    
    '' Grava as informações
    If objCADINHAPROD.GRAVA(cTipOper) = False Then Exit Sub
    
    MsgBox "A Linha de Produto foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
    If objCADINHAPROD.Atualiza(cTipOper, Str(objCADINHAPROD.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
    
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADINHAPROD = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Call IncRegGridCortes
End Sub

Private Sub Command2_Click()
    If cTipOper = "C" Then Exit Sub
    With grdITECORTES
        If (.Rows - 1) = 0 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        Call objBLBFunc.ExcLinhaGrdFilho(grdCORTES, conCOL_SonDim_Itens, .Cell(flexcpText, .Row, conCOL_SonCort_Itens))
        
        If (.Rows - 1) = 1 Then .Rows = 1
        If (.Rows - 1) > 1 Then Call objBLBFunc.ExclLinhaGrid(grdITECORTES, grdITECORTES.Row)
        Call RefazIndiceCortes
        Call TrocaIndice
    End With
End Sub

Private Sub Command26_Click()
    If cTipOper = "C" Then Exit Sub
    If cTipOper = "I" Or cTipOper = "A" Then
        With grdCORTES
            If (.Rows - 1) = 1 Then .Rows = 1
            If (.Rows - 1) > 1 Then Call objBLBFunc.ExclLinhaGrid(grdCORTES, grdCORTES.Row)
        End With
    End If
End Sub

Private Sub Command27_Click()
    If (cTipOper = "I" Or cTipOper = "A") Then Call IncRegGrid
End Sub

Private Sub Command3_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdPRODUTOS, grdPRODUTOS.Row)
End Sub

Private Sub Command4_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGridProdtos
End Sub

Private Sub Command5_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGridArmTPFR
End Sub

Private Sub Command6_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdARMZLIN, grdARMZLIN.Row)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADINHAPROD = CreateObject("CADLINHAPROD.clsCADLINHAPROD")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADINHAPROD.FILIAL = FILIAL
   SSTab1.Tab = 0
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Linha de Produto - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADLINHAPROD
    
    optFILIALPED(0).Value = True
    
    Call InitGrid
    Call InitGridCortes
    Call InitGridProdutos
    Call InitGridArmaz
    
    Call LimpaListBox
    Call ConfLstFecham
    Call ConfLstTpFr
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub grdARMZLIN_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     
     Dim strINDICE As String
     
     With grdARMZLIN
          Select Case Col
                 Case conCOL_SonArm_Qtdelatas
                        strINDICE = Trim(.Cell(flexcpText, Row, Col))
                        If objBLBFunc.FcVerifItensRepetidos(grdARMZLIN, Row, conCOL_SonArm_Qtdelatas, strINDICE) = False Then
                           MsgBox "Esta Fechamento já esta relacionado na Gride !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGridFTPFR(Row)
                           Exit Sub
                        End If
          End Select
     End With
End Sub

Private Sub grdARMZLIN_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonArm_Qtdelatas
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdARMZLIN.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdARMZLIN_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdARMZLIN
          Select Case Col
                    Case conCOL_SonArm_Qtdelatas
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdARMZLIN_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     Dim curVLUNITARIO  As Currency
     Dim strINDICE      As String
     Dim i              As Integer
     
     With grdARMZLIN
          Select Case Col
                 Case conCOL_SonArm_Qtdelatas
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Valor Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
          End Select
     End With

End Sub


Private Sub grdCORTES_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     Dim i As Integer
     With grdCORTES
          Select Case Col
                 Case conCOL_SonDim_CodDim
                 Case conCOL_SonDim_EXPESS, _
                      conCOL_SonDim_LAGURA, _
                      conCOL_SonDim_COMPRI
                      If Len(Trim(.Cell(flexcpText, Row, Col))) > 0 Then .Cell(flexcpText, Row, Col) = Format(.Cell(flexcpText, Row, Col), "#,##0.00")
          End Select
     End With
End Sub

Private Sub grdCORTES_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonDim_DescDim, _
         conCOL_SonDim_Itens
         Cancel = True
    Case conCOL_SonDim_CodDim, _
         conCOL_SonDim_PesqDim, _
         conCOL_SonDim_PADRAO, _
         conCOL_SonDim_EXPESS, _
         conCOL_SonDim_LAGURA, _
         conCOL_SonDim_COMPRI
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdCORTES.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdCORTES_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    Dim strDESCPROD As String
    Dim strINDICE   As String
    
    If (grdCORTES.Rows - 1) = 0 Then Exit Sub
    
    Select Case Col
        Case conCOL_SonDim_PesqDim
    
            If cTipOper = "C" Then Exit Sub
            
            ReDim arrCAMPOS(1 To 2, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
            
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADDIMCORTE " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL = " & FILIAL
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "2000"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCORTE"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "SGI_DESCORTE"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Medidas de Corte")
            
            If Len(Trim(varRETORNO)) > 0 Then
                strINDICE = grdCORTES.Cell(flexcpText, Row, conCOL_SonDim_Itens) & varRETORNO
                If objBLBFunc.FcVerifItensRepetidos(grdCORTES, Row, conCOL_SonDim_INDICE, strINDICE) = False Then
                     MsgBox "Esta Medida de Corte já foi relacionado na Gride !!!", vbOKOnly + vbExclamation
                     Call LimpaCamposGrid(Row)
                     Exit Sub
                End If
               
                grdCORTES.Cell(flexcpText, Row, conCOL_SonDim_INDICE) = strINDICE
                grdCORTES.Cell(flexcpText, Row, conCOL_SonDim_CodDim) = varRETORNO
                If PegaDescrMedidaCorte(varRETORNO, Row) = False Then
                     Call LimpaCamposGrid(Row)
                     Exit Sub
                End If
            End If
    End Select

End Sub


Private Sub grdCORTES_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdCORTES
          Select Case Col
                    Case conCOL_SonDim_CodDim, _
                         conCOL_SonDim_QTDECORPOS
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_SonDim_EXPESS, _
                         conCOL_SonDim_LAGURA, _
                         conCOL_SonDim_COMPRI, _
                         conCOL_SonDim_PERDPROC
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsDouble)
          
          End Select
     End With
End Sub

Private Sub grdCORTES_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     Dim curVLUNITARIO  As Currency
     Dim strINDICE      As String
     Dim i              As Integer
     
     With grdCORTES
          Select Case Col
                 Case conCOL_SonDim_CodDim
                        If .EditText = Empty Then Exit Sub
                        
                        strINDICE = Trim(.Cell(flexcpText, Row, conCOL_SonCort_Itens) & .EditText)
                        If objBLBFunc.FcVerifItensRepetidos(grdCORTES, Row, conCOL_SonDim_INDICE, strINDICE) = False Then
                           MsgBox "Esta medida de corte já esta relacionado na Gride !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGrid(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_SonDim_INDICE) = strINDICE
                        If PegaDescrMedidaCorte(.EditText, Row) = False Then
                            Call LimpaCamposGrid(Row)
                            Cancel = True
                            Exit Sub
                        End If
                Case conCOL_SonDim_PADRAO
                        If .EditText = Empty Then Exit Sub
                        
                        strINDICE = Trim(.Cell(flexcpText, Row, conCOL_SonDim_INDICE))
                        
                        For i = 1 To (.Rows - 1)
                            If strINDICE <> Trim(.Cell(flexcpText, i, conCOL_SonDim_INDICE)) Then .Cell(flexcpText, i, conCOL_SonDim_PADRAO) = "0"
                        Next i
                Case conCOL_SonDim_EXPESS, _
                     conCOL_SonDim_LAGURA, _
                     conCOL_SonDim_COMPRI
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Valor Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
          End Select
     End With
End Sub


Private Sub grdITECORTES_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonCort_Itens, _
         conCOL_SonCort_ItensBKP
         Cancel = True
    Case Else
        grdITECORTES.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdITECORTES_Click()
    With grdITECORTES
        If (.Rows - 1) = 0 Then Exit Sub
        If .Rows = 0 Then Exit Sub
        Call MostraLinhasFilho(.Cell(flexcpText, .Row, conCOL_SonCort_Itens), grdCORTES, conCOL_SonDim_Itens)
    End With
End Sub

Private Sub grdITECORTES_RowColChange()
    With grdITECORTES
        If (.Rows - 1) = 0 Then Exit Sub
        If .Rows = 0 Then Exit Sub
        Call MostraLinhasFilho(.Cell(flexcpText, .Row, conCOL_SonCort_Itens), grdCORTES, conCOL_SonDim_Itens)
    End With
End Sub


Private Sub grdPRODUTOS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdPRODUTOS
          Select Case Col
                 Case conCOL_SonProd_CodProd
          End Select
     End With
End Sub

Private Sub grdPRODUTOS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonProd_CodID, _
         conCOL_SonProd_DescProd
         Cancel = True
    Case conCOL_SonProd_CodProd, _
         conCOL_SonProd_PesqProd
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdPRODUTOS.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdPRODUTOS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If (grdPRODUTOS.Rows - 1) = 0 Then Exit Sub
    
    Dim strINDICE As String
    
    Select Case Col
        Case conCOL_SonProd_PesqProd
    
            If cTipOper = "C" Then Exit Sub
            
            ReDim arrCAMPOS(1 To 4, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
            
            sSql = ""
            
            sSql = "Select" & vbCrLf
            sSql = sSql & "       PRO.SGI_IDPRODUTO" & vbCrLf
            
            ''sSql = sSql & "       ,Case When PRO.SGI_PRODUTOTIPO = 1 then" & vbCrLf
            ''sSql = sSql & "                  replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' +" & vbCrLf
            ''sSql = sSql & "                  replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' +" & vbCrLf
            ''sSql = sSql & "                  replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' +" & vbCrLf
            ''sSql = sSql & "                  (Case When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
            ''sSql = sSql & "                        When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End)" & vbCrLf
            ''sSql = sSql & "             Else" & vbCrLf
            ''sSql = sSql & "                  SGI_CODIGO" & vbCrLf
            ''sSql = sSql & "             End As SGI_CODIGO" & vbCrLf
            
            sSql = sSql & "       ,PRO.SGI_CODIGO" & vbCrLf
            sSql = sSql & "       ,PRO.SGI_DESCRICAO" & vbCrLf
            sSql = sSql & "       ,PRO.SGI_COMPLEMENTO" & vbCrLf
            sSql = sSql & "  From" & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
            sSql = sSql & " Where" & vbCrLf
            sSql = sSql & "       PRO.SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "   And PRO.SGI_PRODUTOTIPO = 0"
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_IDPRODUTO"
            arrCAMPOS(1, 2) = "N"
            arrCAMPOS(1, 3) = "ID"
            arrCAMPOS(1, 4) = "800"
            arrCAMPOS(1, 5) = "PRO.SGI_IDPRODUTO"
            
            arrCAMPOS(2, 1) = "SGI_CODIGO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Produto"
            arrCAMPOS(2, 4) = "1500"
            arrCAMPOS(2, 5) = "PRO.SGI_CODIGO"
            
            arrCAMPOS(3, 1) = "SGI_COMPLEMENTO"
            arrCAMPOS(3, 2) = "S"
            arrCAMPOS(3, 3) = "Complemento"
            arrCAMPOS(3, 4) = "2500"
            arrCAMPOS(3, 5) = "PRO.SGI_COMPLEMENTO"
            
            arrCAMPOS(4, 1) = "SGI_DESCRICAO"
            arrCAMPOS(4, 2) = "S"
            arrCAMPOS(4, 3) = "Descrição"
            arrCAMPOS(4, 4) = "5000"
            arrCAMPOS(4, 5) = "PRO.SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
            
            If Len(Trim(varRETORNO)) > 0 Then
               
                If objBLBFunc.FcVerifItensRepetidos(grdPRODUTOS, Row, conCOL_SonProd_CodID, varRETORNO) = False Then
                   MsgBox "Este Produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                   Call LimpaCamposGridProdutos(Row)
                   Exit Sub
                End If
               
                With grdPRODUTOS
                    .Cell(flexcpText, Row, conCOL_SonProd_CodID) = varRETORNO
                    Call PesDescProduto(varRETORNO, Row)
                End With
               
            End If
    
    End Select

End Sub

Private Sub grdPRODUTOS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdPRODUTOS
          Select Case Col
                    Case conCOL_SonProd_CodProd
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
          End Select
     End With
End Sub

Private Sub grdPRODUTOS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdPRODUTOS
          Select Case Col
                 Case conCOL_SonProd_CodProd
                        If .EditText = Empty Then Exit Sub
                        
                        .Cell(flexcpText, Row, conCOL_SonProd_CodID) = PegaIDProduto(Trim(.EditText))
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_SonProd_CodID))) = 0 Then
                           MsgBox "Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGridProdutos(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        
                        If objBLBFunc.FcVerifItensRepetidos(grdPRODUTOS, Row, conCOL_SonProd_CodID, Trim(.Cell(flexcpText, Row, conCOL_SonProd_CodID))) = False Then
                           MsgBox "Este produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGridProdutos(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        
                        Call PesDescProduto(.Cell(flexcpText, Row, conCOL_SonProd_CodID), Row)
          End Select
     End With

End Sub

Private Sub txtALTURA_GotFocus()
    objBLBFunc.SelecionaCampos txtALTURA.Name, frmCADLINHAPROD
End Sub

Private Sub txtALTURA_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtALTURA.Text
End Sub

Private Sub txtALTURA_Validate(Cancel As Boolean)

    If Len(Trim(txtALTURA.Text)) = 0 Then Exit Sub
    
    txtALTURA.Text = Format(txtALTURA.Text, "#,##0.00")

End Sub

Private Sub txtCodLin_GotFocus()
    objBLBFunc.SelecionaCampos txtCodLin.Name, frmCADLINHAPROD
End Sub

Private Sub txtCodLin_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodLin.Text
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADLINHAPROD
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function Verifica_Campos() As Boolean

    Verifica_Campos = False
    
    Dim j As Integer
    Dim i As Integer
    Dim blAchou      As Boolean
    Dim lngQTDPadrao As Long
    
    If Len(Trim(txtCodLin.Text)) = 0 Then
       MsgBox "Informe o Código da Linha de Produto !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodLin.SetFocus
       Exit Function
    End If
    If Len(Trim(txtDescricao.Text)) = 0 Then
       MsgBox "Informe a descrição da Linha de Produto !!!", vbOKOnly + vbExclamation, "Aviso"
       txtDescricao.SetFocus
       Exit Function
    End If
    If Len(Trim(txtDESENV.Text)) = 0 Then
        MsgBox "ATENÇÃO - Favor informar o valor para o campo desenvolvimento !!!", vbOKOnly + vbExclamation, "Aviso"
        txtDESENV.SetFocus
        Exit Function
    End If
    If Len(Trim(txtALTURA.Text)) = 0 Then
        MsgBox "ATENÇÃO - Favor informar o valor para o campo altura !!!", vbOKOnly + vbExclamation, "Aviso"
        txtALTURA.SetFocus
        Exit Function
    End If
    If Len(Trim(txtPERDPROC.Text)) = 0 Then
        MsgBox "ATENÇÃO - Favor informar o valor para o campo perda de processo !!!", vbOKOnly + vbExclamation, "Aviso"
        txtPERDPROC.SetFocus
        Exit Function
    End If
    
    ''Armazenamento
    ''For I = 1 To (grdARMZLIN.Rows - 1)
    ''    If lstTpFr.Selected((CLng(grdARMZLIN.Cell(flexcpText, I, conCOL_SonArm_Fecham)) - 1)) = False Then
    ''        MsgBox "ATENÇÃO" & vbCrLf & _
    ''               "O Fechamento incluso não esta relacionado em fechamento !!!", vbOKOnly + vbExclamation, "Aviso"
    ''       Exit Function
    ''  End If
    ''Next I
    
    If cTipOper = "I" Then
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Está descrição da Linha de Produto já existe !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          txtDescricao.SetFocus
          Exit Function
       End If
       BREC.Close
       
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODLIN = " & Trim(txtCodLin.Text)
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Código da Linha de Produto já existe !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          txtCodLin.SetFocus
          Exit Function
       End If
       BREC.Close
       
    End If
    
    If cTipOper = "A" Then
    
       If objCADINHAPROD.DESCRI <> txtDescricao.Text Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Está descrição da linha de produto já existe !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtDescricao.Text = objCADINHAPROD.DESCRI
             txtDescricao.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
       
       If objCADINHAPROD.CODIGOLIN <> Val(Trim(txtCodLin.Text)) Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CODLIN = " & Trim(txtCodLin)
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Código da linha de produto já existe !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtCodLin.Text = objCADINHAPROD.CODIGOLIN
             txtCodLin.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
       
       lngQTDPadrao = 0
       With grdCORTES
            For i = 1 To (.Rows - 1)
                If .Cell(flexcpText, i, conCOL_SonDim_PADRAO) = 1 Then
                    lngQTDPadrao = lngQTDPadrao + 1
                End If
            Next i
            If .Rows - 1 > 0 Then
                If lngQTDPadrao = 0 Then
                    MsgBox "Escolha uma medida como padrão !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                ElseIf lngQTDPadrao > 1 Then
                    MsgBox "Somente é permitido 1 medida como padrão !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
            End If
       End With
       
    End If
    
    Verifica_Campos = True

End Function

Private Sub Consulta()
    
    Dim i As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    
    Me.Caption = "Cadastro de Linha de Produto - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADLINHAPROD
    
    objCADINHAPROD.CODIGO = iCodigo
    
    optFILIALPED(0).Value = True
    Call InitGrid
    Call InitGridCortes
    Call InitGridProdutos
    Call InitGridArmaz
    
    Call LimpaListBox
    Call ConfLstFecham
    Call ConfLstTpFr
    
    If objCADINHAPROD.Carrega_campos = True Then
       txtCodigo.Text = objCADINHAPROD.CODIGO
       txtCodLin.Text = Format(objCADINHAPROD.CODIGOLIN, "###000")
       txtDescricao.Text = objCADINHAPROD.DESCRI
       optFILIALPED(objCADINHAPROD.FILIALPED).Value = True
       If Len(Trim(objCADINHAPROD.DESENV)) > 0 Then txtDESENV.Text = Format(objCADINHAPROD.DESENV, "#,##0.00")
       If Len(Trim(objCADINHAPROD.ALTURA)) > 0 Then txtALTURA.Text = Format(objCADINHAPROD.ALTURA, "#,##0.00")
       If Len(Trim(objCADINHAPROD.PERDPROC)) > 0 Then txtPERDPROC.Text = Format(objCADINHAPROD.PERDPROC, "#,##0.00")
       If Len(Trim(objCADINHAPROD.QTDECORPOS)) > 0 Then txtQTDECORPOS.Text = objCADINHAPROD.QTDECORPOS
       
       Call PopGrdCorte
       Call PopGrdsEQCorte
       Call PopGrdProdutos
       Call PopGrdTPFRArm
       
       Call Seleciona_Fech
       Call Seleciona_FechTPFR
   End If

End Sub

Private Sub Altera()

    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de Linha de Produto - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADLINHAPROD
    
    objCADINHAPROD.CODIGO = iCodigo
    
    optFILIALPED(0).Value = True
    Call InitGrid
    Call InitGridCortes
    Call InitGridProdutos
    Call InitGridArmaz
    
    Call LimpaListBox
    Call ConfLstFecham
    Call ConfLstTpFr
    
    If objCADINHAPROD.Carrega_campos = True Then
        txtCodigo.Text = objCADINHAPROD.CODIGO
        txtCodLin.Text = Format(objCADINHAPROD.CODIGOLIN, "###000")
        txtDescricao.Text = objCADINHAPROD.DESCRI
        optFILIALPED(objCADINHAPROD.FILIALPED).Value = True
        If Len(Trim(objCADINHAPROD.DESENV)) > 0 Then txtDESENV.Text = Format(objCADINHAPROD.DESENV, "#,##0.00")
        If Len(Trim(objCADINHAPROD.ALTURA)) > 0 Then txtALTURA.Text = Format(objCADINHAPROD.ALTURA, "#,##0.00")
        If Len(Trim(objCADINHAPROD.PERDPROC)) > 0 Then txtPERDPROC.Text = Format(objCADINHAPROD.PERDPROC, "#,##0.00")
        If Len(Trim(objCADINHAPROD.QTDECORPOS)) > 0 Then txtQTDECORPOS.Text = objCADINHAPROD.QTDECORPOS
        
        Call PopGrdCorte
        Call PopGrdsEQCorte
        Call PopGrdProdutos
        Call PopGrdTPFRArm
        
        Call Seleciona_Fech
        Call Seleciona_FechTPFR
    End If

End Sub

Private Sub txtDESENV_GotFocus()
    objBLBFunc.SelecionaCampos txtDESENV.Name, frmCADLINHAPROD
End Sub

Private Sub txtDESENV_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtDESENV.Text
End Sub

Private Sub txtDESENV_Validate(Cancel As Boolean)

    If Len(Trim(txtDESENV.Text)) = 0 Then Exit Sub
    
    txtDESENV.Text = Format(txtDESENV.Text, "#,##0.00")

End Sub

Private Sub txtPERDPROC_GotFocus()
    objBLBFunc.SelecionaCampos txtPERDPROC.Name, frmCADLINHAPROD
End Sub

Private Sub txtPERDPROC_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPERDPROC.Text
End Sub

Private Sub txtPERDPROC_Validate(Cancel As Boolean)

    If Len(Trim(txtPERDPROC.Text)) = 0 Then Exit Sub
    
    txtPERDPROC.Text = Format(txtPERDPROC.Text, "#,##0.00")

End Sub

Private Sub InitGrid()

    With grdCORTES
    
       .Cols = conColumnsIn_SonDim
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonDim_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonDim_Itens) = ""
       .ColDataType(conCOL_SonDim_Itens) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonDim_CodDim) = ""
       .ColDataType(conCOL_SonDim_CodDim) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonDim_PesqDim) = ""
       .ColDataType(conCOL_SonDim_PesqDim) = flexDTString
       .ColComboList(conCOL_SonDim_PesqDim) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonDim_DescDim) = ""
       .ColDataType(conCOL_SonDim_DescDim) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonDim_INDICE) = ""
       .ColDataType(conCOL_SonDim_INDICE) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonDim_PADRAO) = ""
       .ColDataType(conCOL_SonDim_PADRAO) = flexDTString
       .ColComboList(conCOL_SonDim_PADRAO) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_SonDim_EXPESS) = ""
       .ColDataType(conCOL_SonDim_EXPESS) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_SonDim_LAGURA) = ""
       .ColDataType(conCOL_SonDim_LAGURA) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_SonDim_COMPRI) = ""
       .ColDataType(conCOL_SonDim_COMPRI) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_SonDim_QTDECORPOS) = ""
       .ColDataType(conCOL_SonDim_QTDECORPOS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonDim_PERDPROC) = ""
       .ColDataType(conCOL_SonDim_PERDPROC) = flexDTDouble
       
       .ColWidth(conCOL_SonDim_Itens) = 1000
       .ColWidth(conCOL_SonDim_CodDim) = 1000
       .ColWidth(conCOL_SonDim_PesqDim) = 300
       .ColWidth(conCOL_SonDim_DescDim) = 3000
       .ColWidth(conCOL_SonDim_INDICE) = 0
       .ColWidth(conCOL_SonDim_PADRAO) = 700
       .ColWidth(conCOL_SonDim_EXPESS) = 900
       .ColWidth(conCOL_SonDim_LAGURA) = 800
       .ColWidth(conCOL_SonDim_COMPRI) = 1000
       .ColWidth(conCOL_SonDim_QTDECORPOS) = 1000
       .ColWidth(conCOL_SonDim_PERDPROC) = 1500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGrid()
   
    If (grdITECORTES.Rows - 1) = 0 Then
        MsgBox "Favor inserir registro na Gride de Cortes !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    ElseIf (grdITECORTES.Row) = 0 Then
        MsgBox "Favor Selecionar uma seguencia de Corte !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If objBLBFunc.FcExisteLinhaVazia(grdCORTES, conCOL_SonDim_CodDim) = False Then Exit Sub
    
    grdCORTES.AddItem grdITECORTES.Cell(flexcpText, grdITECORTES.Row, conCOL_SonCort_Itens) & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      0 & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      ""
End Sub

Private Sub RefazIndice()

    Dim i As Integer
    
    With grdCORTES
        For i = 1 To (.Rows - 1)
            .Cell(flexcpText, i, conCOL_SonDim_Itens) = i
        Next i
    End With
End Sub

Private Function PegaDescrMedidaCorte(strCodMedida As String, lngLINHA As Long) As Boolean
    
    PegaDescrMedidaCorte = False
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDIMCORTE" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodMedida
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then
       With grdCORTES
            .Cell(flexcpText, lngLINHA, conCOL_SonDim_DescDim) = BREC2!SGI_DESCORTE
            PegaDescrMedidaCorte = True
       End With
    Else
       MsgBox "Este Corte Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC2.Close
    
End Function

Private Sub LimpaCamposGrid(lngROWS As Long)
    With grdCORTES
            .Cell(flexcpText, lngROWS, conCOL_SonDim_CodDim) = Empty
            .Cell(flexcpText, lngROWS, conCOL_SonDim_DescDim) = Empty
            .Cell(flexcpText, lngROWS, conCOL_SonDim_INDICE) = Empty
    End With
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADINHAPROD = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub PopGrdCorte()
    Dim i As Long
    arrDIMCORTE = objCADINHAPROD.DIMCORTE
    If IsArray(arrDIMCORTE) Then
        With grdCORTES
            For i = 1 To UBound(arrDIMCORTE)
                .AddItem arrDIMCORTE(i, 1) & vbTab & _
                         arrDIMCORTE(i, 2) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         Trim(arrDIMCORTE(i, 1)) & Trim(arrDIMCORTE(i, 2)) & vbTab & _
                         arrDIMCORTE(i, 3) & vbTab & _
                         arrDIMCORTE(i, 4) & vbTab & _
                         arrDIMCORTE(i, 5) & vbTab & _
                         arrDIMCORTE(i, 6) & vbTab & _
                         arrDIMCORTE(i, 7) & vbTab & _
                         arrDIMCORTE(i, 8)
                         
                Call PegaDescrMedidaCorte(Trim(Str(arrDIMCORTE(i, 2))), (.Rows - 1))
            Next i
        End With
    End If
End Sub

Private Sub InitGridCortes()

    With grdITECORTES
    
       .Cols = conColumnsIn_SonCort
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonCort_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCort_Itens) = ""
       .ColDataType(conCOL_SonCort_Itens) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCort_ItensBKP) = ""
       .ColDataType(conCOL_SonCort_ItensBKP) = flexDTLong
       
       .ColWidth(conCOL_SonCort_Itens) = 1000
       .ColWidth(conCOL_SonCort_ItensBKP) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGridCortes()
   
    If cTipOper = "C" Then Exit Sub

    If objBLBFunc.FcExisteLinhaVazia(grdITECORTES, conCOL_SonCort_Itens) = False Then Exit Sub
    
    grdITECORTES.AddItem "" & vbTab & _
                         ""
                            
    Call RefazIndiceCortes
    
    grdITECORTES.Row = (grdITECORTES.Rows - 1)
    grdITECORTES.Cell(flexcpText, grdITECORTES.Row, conCOL_SonCort_ItensBKP) = (grdITECORTES.Rows - 1)
    
    Call MostraLinhasFilho(grdITECORTES.Cell(flexcpText, grdITECORTES.Row, conCOL_SonCort_Itens), grdCORTES, conCOL_SonDim_Itens)
    
End Sub

Private Sub RefazIndiceCortes()
    Dim i As Integer
    Dim j As Integer
    With grdITECORTES
        For i = 1 To (.Rows - 1)
            .Cell(flexcpText, i, conCOL_SonCort_Itens) = i
        Next i
    End With
End Sub

Private Sub MostraLinhasFilho(strINDICEPAI As String, grdFILHO As Variant, lngCOLFILHO As Long)
    Dim i As Integer
    With grdFILHO
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, lngCOLFILHO) = strINDICEPAI Then
               .RowHidden(i) = False
            Else
               .RowHidden(i) = True
            End If
        Next i
    End With
End Sub

Private Sub TrocaIndice()
    Dim i As Integer
    Dim j As Integer
    With grdITECORTES
        For i = 1 To (.Rows - 1)
            For j = 1 To (grdCORTES.Rows - 1)
                If Trim(grdCORTES.Cell(flexcpText, j, conCOL_SonDim_Itens)) = Trim(.Cell(flexcpText, i, conCOL_SonCort_ItensBKP)) Then
                    grdCORTES.Cell(flexcpText, j, conCOL_SonDim_Itens) = Trim(.Cell(flexcpText, i, conCOL_SonCort_Itens))
                    grdCORTES.Cell(flexcpText, j, conCOL_SonDim_INDICE) = Trim(grdCORTES.Cell(flexcpText, j, conCOL_SonDim_Itens)) & Trim(grdCORTES.Cell(flexcpText, j, conCOL_SonDim_CodDim))
                    .Cell(flexcpText, i, conCOL_SonCort_ItensBKP) = Trim(.Cell(flexcpText, i, conCOL_SonCort_Itens))
                End If
            Next j
        Next i
    End With
End Sub


Private Sub PopGrdsEQCorte()
    Dim i As Long
    arrSEQCORTE = objCADINHAPROD.SEQCORTE
    If IsArray(arrSEQCORTE) Then
        With grdITECORTES
            For i = 1 To UBound(arrSEQCORTE)
                .AddItem arrSEQCORTE(i, 1) & vbTab & _
                         arrSEQCORTE(i, 1)
            Next i
            If (.Rows - 1) > 0 Then
                .Row = 1
                Call MostraLinhasFilho(.Cell(flexcpText, .Row, conCOL_SonCort_Itens), grdCORTES, conCOL_SonDim_Itens)
            End If
        End With
    End If
End Sub

Private Sub txtQTDECORPOS_GotFocus()
    objBLBFunc.SelecionaCampos txtQTDECORPOS.Name, Me
End Sub

Private Sub txtQTDECORPOS_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPERDPROC.Text
End Sub

Private Sub InitGridProdutos()

    With grdPRODUTOS
    
       .Cols = conColumnsIn_SonProd
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProd_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProd_CodID) = ""
       .ColDataType(conCOL_SonProd_CodID) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_CodProd) = ""
       .ColDataType(conCOL_SonProd_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProd_PesqProd) = ""
       .ColDataType(conCOL_SonProd_PesqProd) = flexDTString
       .ColComboList(conCOL_SonProd_PesqProd) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProd_DescProd) = ""
       .ColDataType(conCOL_SonProd_DescProd) = flexDTString
       
       .ColWidth(conCOL_SonProd_CodID) = 0
       .ColWidth(conCOL_SonProd_CodProd) = 1500
       .ColWidth(conCOL_SonProd_PesqProd) = 300
       .ColWidth(conCOL_SonProd_DescProd) = 5500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub IncRegGridProdtos()
   
    If objBLBFunc.FcExisteLinhaVazia(grdPRODUTOS, conCOL_SonProd_CodID) = False Then Exit Sub
    
    grdPRODUTOS.AddItem "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        ""
End Sub

Private Function PegaIDProduto(strCODPRODUTO As String) As String

    PegaIDProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = '" & Trim(UCase(strCODPRODUTO)) & "'" & vbCrLf
    sSql = sSql & "   And SGI_FILIAL = " & FILIAL

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaIDProduto = BREC!SGI_IDPRODUTO
    BREC.Close
    
End Function

Private Sub PesDescProduto(strID As String, lngROW As Long)

    If Len(Trim(strID)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PRO.SGI_IDPRODUTO" & vbCrLf
    
    ''sSql = sSql & "       ,Case When PRO.SGI_PRODUTOTIPO = 1 then" & vbCrLf
    ''sSql = sSql & "                  replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' +" & vbCrLf
    ''sSql = sSql & "                  replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' +" & vbCrLf
    ''sSql = sSql & "                  replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' +" & vbCrLf
    ''sSql = sSql & "                  (Case When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
    ''sSql = sSql & "                        When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End)" & vbCrLf
    ''sSql = sSql & "             Else" & vbCrLf
    ''sSql = sSql & "                  SGI_CODIGO" & vbCrLf
    ''sSql = sSql & "             End As SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "       ,PRO.SGI_CODIGO" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_COMPLEMENTO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_IDPRODUTO     = " & Trim(strID)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       With grdPRODUTOS
            .Cell(flexcpText, lngROW, conCOL_SonProd_DescProd) = Trim(BREC!SGI_DESCRICAO)
            .Cell(flexcpText, lngROW, conCOL_SonProd_CodProd) = Trim(BREC!SGI_CODIGO)
       End With
    End If
    BREC.Close
    
End Sub


Private Sub LimpaCamposGridProdutos(lngROWS As Long)
    With grdPRODUTOS
            .Cell(flexcpText, lngROWS, conCOL_SonProd_CodID) = Empty
            .Cell(flexcpText, lngROWS, conCOL_SonProd_CodProd) = Empty
            .Cell(flexcpText, lngROWS, conCOL_SonProd_DescProd) = Empty
    End With
End Sub


Private Sub PopGrdProdutos()
    Dim i As Long
    arrPRODUTOS = objCADINHAPROD.PRODUTOS
    If IsArray(arrPRODUTOS) Then
        With grdPRODUTOS
            For i = 1 To UBound(arrPRODUTOS)
                .AddItem arrPRODUTOS(i, 1) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         ""
                         
                Call PesDescProduto(Str(arrPRODUTOS(i, 1)), (.Rows - 1))
            
            Next i
        End With
    End If
End Sub

Private Sub LimpaListBox()
    lstFech.Clear
    lstTpFr.Clear
End Sub

Public Sub ConfLstFecham()

    lstFech.AddItem "SOLDA"
    lstFech.ItemData(lstFech.NewIndex) = 0
    
    lstFech.AddItem "AGRAFADO"
    lstFech.ItemData(lstFech.NewIndex) = 1

    lstFech.AddItem "REPUXO"
    lstFech.ItemData(lstFech.NewIndex) = 2
End Sub

Private Sub Seleciona_Fech()

        If Not IsArray(objCADINHAPROD.FECHAMENTOS) Then Exit Sub
        If lstFech.ListCount = 0 Then Exit Sub
       
        Dim i As Integer
        Dim j As Integer
       
        arrFECHAMENTOS = objCADINHAPROD.FECHAMENTOS
        
        With lstFech
            
            For i = 0 To (.ListCount - 1)
                 For j = 1 To UBound(arrFECHAMENTOS)
                    If .ItemData(i) = CInt(arrFECHAMENTOS(j)) Then .Selected(i) = True
                 Next j
            Next i
        
        End With

End Sub

Public Sub ConfLstTpFr()

    lstTpFr.Clear
    
    If BREC8.State = 1 Then Exit Sub
    
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFECHAM" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC8.EOF()
        lstTpFr.AddItem Trim(BREC8!SGI_DESCRI)
        lstTpFr.ItemData(lstTpFr.NewIndex) = BREC8!SGI_CODIGO
        BREC8.MoveNext
    Loop
    BREC8.Close
    
End Sub

Private Sub Seleciona_FechTPFR()

        If Not IsArray(objCADINHAPROD.FECHTPFR) Then Exit Sub
        If lstTpFr.ListCount = 0 Then Exit Sub
       
        Dim i As Integer
        Dim j As Integer
       
        arrFECHTPFR = objCADINHAPROD.FECHTPFR
        
        With lstTpFr
            
            For i = 0 To (.ListCount - 1)
                 For j = 1 To UBound(arrFECHTPFR)
                    If .ItemData(i) = CInt(arrFECHTPFR(j)) Then .Selected(i) = True
                 Next j
            Next i
        
        End With

End Sub

Private Sub InitGridArmaz()

    With grdARMZLIN
    
       .Cols = conColumnsIn_SonArm
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonArm_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonArm_Qtdelatas) = ""
       .ColDataType(conCOL_SonArm_Qtdelatas) = flexDTLong
       
       .ColWidth(conCOL_SonArm_Qtdelatas) = 1500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub IncRegGridArmTPFR()
   
    If objBLBFunc.FcExisteLinhaVazia(grdARMZLIN, conCOL_SonArm_Qtdelatas) = False Then Exit Sub
    
    grdARMZLIN.AddItem ""
                       
End Sub

Private Sub LimpaCamposGridFTPFR(lngROWS As Long)
    With grdARMZLIN
            .Cell(flexcpText, lngROWS, conCOL_SonArm_Qtdelatas) = Empty
    End With
End Sub


Private Sub PopGrdTPFRArm()
    Dim i As Long
    arrARMAZ = objCADINHAPROD.ARMAZ
    If IsArray(arrARMAZ) Then
        With grdARMZLIN
            For i = 1 To UBound(arrARMAZ)
                .AddItem arrARMAZ(i)
                         
                ''Call PesDescProduto(Str(arrPRODUTOS(I, 1)), (.Rows - 1))
            
            Next i
        End With
    End If
End Sub



Private Function CalcTotPecasHora(strTOTHORAS As String, lngQTPCDEHORAS As Long) As Long

    CalcTotPecasHora = 0
    
    Dim arrHORAMIN      As Variant
    Dim lngTotPecasHora As Long

    
    '' -------------------------
    '' Qtde de Pecas
    arrHORAMIN = Split(strTOTHORAS, ":")
    lngTotPecasHora = (arrHORAMIN(0) * lngQTPCDEHORAS)
    
    CalcTotPecasHora = lngTotPecasHora
    '' -------------------------

End Function
