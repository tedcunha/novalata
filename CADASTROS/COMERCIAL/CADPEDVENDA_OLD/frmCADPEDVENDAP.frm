VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADPEDVENDAP 
   Caption         =   "Cadastro de Pedidos de Vendas"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   14685
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   6975
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   14655
      Begin TabDlg.SSTab stPEDIDOS 
         Height          =   6735
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   11880
         _Version        =   393216
         Style           =   1
         Tabs            =   8
         TabsPerRow      =   8
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
         TabCaption(0)   =   "Liberados"
         TabPicture(0)   =   "frmCADPEDVENDAP.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "flxGRIDPEDIDOS"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Aguardando Liberação"
         TabPicture(1)   =   "frmCADPEDVENDAP.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(1)=   "flxGRIDBLOQUADOS"
         Tab(1).Control(2)=   "Label3(2)"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Reprovados"
         TabPicture(2)   =   "frmCADPEDVENDAP.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "flxReprovados"
         Tab(2).Control(1)=   "Label3(3)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Faturados"
         TabPicture(3)   =   "frmCADPEDVENDAP.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "grdPEDFATURADO"
         Tab(3).Control(1)=   "Label3(0)"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Bloqueados"
         TabPicture(4)   =   "frmCADPEDVENDAP.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Command5"
         Tab(4).Control(1)=   "grdBLOQALT"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "Aguardando Liberação do Fotolito"
         TabPicture(5)   =   "frmCADPEDVENDAP.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "grdLIBLITO"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "Command6"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).ControlCount=   2
         TabCaption(6)   =   "Para Estoque"
         TabPicture(6)   =   "frmCADPEDVENDAP.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "grdPARAEST"
         Tab(6).Control(1)=   "Command7"
         Tab(6).ControlCount=   2
         TabCaption(7)   =   "Bloqueado - P.Data/P.Cota"
         TabPicture(7)   =   "frmCADPEDVENDAP.frx":00C4
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "cmdLIBPDATAPCOTA"
         Tab(7).Control(1)=   "grdLIBPDATAPCOTA"
         Tab(7).ControlCount=   2
         Begin VB.CommandButton cmdLIBPDATAPCOTA 
            Caption         =   "Libera P.Data/P.Cota"
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
            Left            =   -74880
            Picture         =   "frmCADPEDVENDAP.frx":00E0
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Liberação das alterações"
            Top             =   6000
            Width           =   1575
         End
         Begin VSFlex8LCtl.VSFlexGrid grdLIBPDATAPCOTA 
            Height          =   5535
            Left            =   -74880
            TabIndex        =   38
            Top             =   360
            Width           =   14175
            _cx             =   25003
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
         Begin VB.CommandButton Command7 
            Caption         =   "Libera Para Estoque"
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
            Left            =   -74880
            Picture         =   "frmCADPEDVENDAP.frx":050D
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Liberação das Alterações do Fotolito"
            Top             =   6000
            Width           =   1695
         End
         Begin VSFlex8LCtl.VSFlexGrid grdPARAEST 
            Height          =   5535
            Left            =   -74880
            TabIndex        =   36
            Top             =   360
            Width           =   14175
            _cx             =   25003
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
         Begin VB.CommandButton Command6 
            Caption         =   "Libera Fotolito"
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
            Left            =   -74880
            Picture         =   "frmCADPEDVENDAP.frx":093A
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Liberação das Alterações do Fotolito"
            Top             =   6000
            Width           =   1335
         End
         Begin VSFlex8LCtl.VSFlexGrid grdLIBLITO 
            Height          =   5535
            Left            =   -74880
            TabIndex        =   34
            Top             =   360
            Width           =   14175
            _cx             =   25003
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
         Begin VB.CommandButton Command5 
            Caption         =   "Libera Alteração"
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
            Left            =   -74880
            Picture         =   "frmCADPEDVENDAP.frx":0D67
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Liberação das alterações"
            Top             =   6000
            Width           =   1335
         End
         Begin VSFlex8LCtl.VSFlexGrid grdBLOQALT 
            Height          =   5535
            Left            =   -74880
            TabIndex        =   32
            Top             =   360
            Width           =   14175
            _cx             =   25003
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
         Begin VSFlex8LCtl.VSFlexGrid grdPEDFATURADO 
            Height          =   5895
            Left            =   -74880
            TabIndex        =   31
            Top             =   360
            Width           =   14175
            _cx             =   25003
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
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   255
            Left            =   -63240
            TabIndex        =   22
            Top             =   6360
            Width           =   2775
            Begin VB.OptionButton optLiberados 
               Caption         =   "Comercial"
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
               Index           =   0
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton optLiberados 
               Caption         =   "Financeiro"
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
               Index           =   1
               Left            =   1320
               TabIndex        =   23
               Top             =   0
               Width           =   1335
            End
         End
         Begin MSFlexGridLib.MSFlexGrid flxReprovados 
            Height          =   5895
            Left            =   -74880
            TabIndex        =   20
            Top             =   360
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   10398
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxGRIDBLOQUADOS 
            Height          =   5895
            Left            =   -74880
            TabIndex        =   16
            Top             =   360
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   10398
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxGRIDPEDIDOS 
            Height          =   5895
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   10398
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
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
            Left            =   -74880
            TabIndex        =   28
            Top             =   6360
            Width           =   11775
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
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
            Left            =   -74880
            TabIndex        =   27
            Top             =   6360
            Width           =   11055
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
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
            Left            =   -74880
            TabIndex        =   26
            Top             =   6360
            Width           =   8055
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
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
            Left            =   120
            TabIndex        =   25
            Top             =   6360
            Width           =   11055
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   14655
      Begin VB.CommandButton Command4 
         Caption         =   "&Liquida"
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
         Left            =   10560
         Picture         =   "frmCADPEDVENDAP.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Reprova o Pedido"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   12240
         Picture         =   "frmCADPEDVENDAP.frx":1591
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Exclui Registro"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Libera Financeiro"
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
         Left            =   5520
         Picture         =   "frmCADPEDVENDAP.frx":1693
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Liberação Financeira"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   11760
         Top             =   120
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Re&prova"
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
         Left            =   9720
         Picture         =   "frmCADPEDVENDAP.frx":1AC0
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Reprova o Pedido"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdDeslib 
         Caption         =   "&Bloqueia"
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
         Left            =   8160
         Picture         =   "frmCADPEDVENDAP.frx":1EBD
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Bloqueia o Pedido"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdLibera 
         Caption         =   "Libe&ra Comercial"
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
         Left            =   6840
         Picture         =   "frmCADPEDVENDAP.frx":22E3
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Liberação Comercial"
         Top             =   120
         Width           =   1335
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
         Left            =   13800
         Picture         =   "frmCADPEDVENDAP.frx":2710
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   13080
         Picture         =   "frmCADPEDVENDAP.frx":2812
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmCADPEDVENDAP.frx":2D44
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmCADPEDVENDAP.frx":2E46
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmCADPEDVENDAP.frx":2F48
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmCADPEDVENDAP.frx":347A
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADPEDVENDAP.frx":39AC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   10095
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   200
         Width           =   2415
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
         Left            =   3720
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCADPEDVENDAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strACESSO        As String
Public strUSUARIO       As String
Public lngCodUsuaro     As Long
Public intFILIALPED     As Integer
Public strVERSAO        As String
Public strNOMCOMP       As String

Dim boolComAcao      As Boolean
Dim lngCodVendedor  As Long
Dim objFuncoes      As Object
Dim objCADPEDVENDA  As Object
Dim objRel          As Object
Dim iCodigo         As Long
Dim strOperacao     As String
Dim arrOPS()        As String
Dim cTipOper        As String

Const conCOL_SonFat_Codigo                          As Integer = 0
Const conCOL_SonFat_Data                            As Integer = 1
Const conCOL_SonFat_Cliente                         As Integer = 2
Const conCOL_SonFat_Situacao                        As Integer = 3
Const conCOL_SonFat_Tipo                            As Integer = 4
Const conCOL_SonFat_Cotacao                         As Integer = 5
Const conCOL_SonFat_Status                          As Integer = 6
Const conCOL_SonFat_FormatString                    As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Cotação|Status"
Const conColumnsIn_SonFat                           As Integer = 7

Const conCOL_SonBloq_Codigo                         As Integer = 0
Const conCOL_SonBloq_Data                           As Integer = 1
Const conCOL_SonBloq_Cliente                        As Integer = 2
Const conCOL_SonBloq_Situacao                       As Integer = 3
Const conCOL_SonBloq_Tipo                           As Integer = 4
Const conCOL_SonBloq_Status                         As Integer = 5
Const conCOL_SonBloq_FormatString                   As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|tatus"
Const conColumnsIn_SonBloq                          As Integer = 6

Const conCOL_SonBloqLit_Codigo                      As Integer = 0
Const conCOL_SonBloqLit_Data                        As Integer = 1
Const conCOL_SonBloqLit_Cliente                     As Integer = 2
Const conCOL_SonBloqLit_Situacao                    As Integer = 3
Const conCOL_SonBloqLit_Tipo                        As Integer = 4
Const conCOL_SonBloqLit_Status                      As Integer = 5
Const conCOL_SonBloqLit_FormatString                As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|tatus"
Const conColumnsIn_SonBloqLit                       As Integer = 6

Const conCOL_SonParaEst_Codigo                      As Integer = 0
Const conCOL_SonParaEst_Data                        As Integer = 1
Const conCOL_SonParaEst_Cliente                     As Integer = 2
Const conCOL_SonParaEst_Situacao                    As Integer = 3
Const conCOL_SonParaEst_Tipo                        As Integer = 4
Const conCOL_SonParaEst_Status                      As Integer = 5
Const conCOL_SonParaEst_FormatString                As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|tatus"
Const conColumnsIn_SonParaEst                       As Integer = 6

Const conCOL_SonBloqPDPC_Codigo                     As Integer = 0
Const conCOL_SonBloqPDPC_Data                       As Integer = 1
Const conCOL_SonBloqPDPC_Cliente                    As Integer = 2
Const conCOL_SonBloqPDPC_Situacao                   As Integer = 3
Const conCOL_SonBloqPDPC_Tipo                       As Integer = 4
Const conCOL_SonBloqPDPC_Status                     As Integer = 5
Const conCOL_SonBloqPDPC_FormatString               As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|tatus"
Const conColumnsIn_SonBloqPDPC                      As Integer = 6


Private Sub cmdAltera_Click()
    
On Error GoTo Err_cmdAltera_Click
    
    If objFuncoes.ChecaAcesso2("A", strACESSO) = False Then Exit Sub
    If stPEDIDOS.Tab = 0 Then
       MsgBox "Este Pedido já está Liberado não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    If stPEDIDOS.Tab = 1 Then
       If Trim(flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.Row, 4)) = "N" Then
            MsgBox "Este Pedido já está Liberado não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
       End If
    End If
    If stPEDIDOS.Tab = 3 Then
       With grdPEDFATURADO
       If Trim(.Cell(flexcpText, .Row, conCOL_SonFat_Situacao)) = "F" Then
            MsgBox "Este Pedido já está Faturado não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
       End If
       End With
    End If
    If stPEDIDOS.Tab = 1 Or _
       stPEDIDOS.Tab = 4 Or _
       stPEDIDOS.Tab = 5 Or _
       stPEDIDOS.Tab = 7 Then Call Operacao("A")
    
    Exit Sub
    
Err_cmdAltera_Click:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdAltera_Click()", Me.Name, "cmdAltera_Click()", strCAMARQERRO)
    
    
End Sub

Private Sub cmdCanFiltro_Click()
   
On Error GoTo Err_cmdCanFiltro_Click
   
   strOperacao = ""
   Call AbilitaCampos
   Call ConfGrid
   Call ConfGridBloqueados
   Call ConfGridReprovados
   Call ConfGridFaturado
   Call ConfGridBloqAlt
   Call ConfGridBloqAltLit
   Call ConfGridParaEstoque
   
   Exit Sub
   
Err_cmdCanFiltro_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdCanFiltro_Click()", Me.Name, "cmdCanFiltro_Click()", strCAMARQERRO)
   
End Sub

Private Sub cmdDeslib_Click()
    
On Error GoTo Err_cmdDeslib_Click
    
        
    
    ''If objFuncoes.ChecaAcesso2("B", strAcesso) = False Then Exit Sub
    If stPEDIDOS.Tab = 3 Then Exit Sub
    If VerifNF = False Then Exit Sub
    If stPEDIDOS.Tab = 0 Or _
       stPEDIDOS.Tab = 2 Or _
       stPEDIDOS.Tab = 5 Then Call Operacao("D")
    If stPEDIDOS.Tab = 1 Then
       If flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.Row, 4) = "N" Then
          Call Operacao("D")
       End If
    End If
    
    Exit Sub
    
Err_cmdDeslib_Click:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdDeslib_Click()", Me.Name, "cmdDeslib_Click()", strCAMARQERRO)
    
End Sub

Private Sub cmdExclui_Click()
  
On Error GoTo Err_cmdExclui_Click
  
  If objFuncoes.ChecaAcesso2("E", strACESSO) = False Then Exit Sub
  If VerifNF = False Then Exit Sub
  
  Dim iResp As Integer
    
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADPEDVENDA.Carrega_Campos = True Then
  
    
    If intFILIALPED = 0 Then
        If objCADPEDVENDA.GRAVA("E") = False Then Exit Sub
    ElseIf intFILIALPED = 1 Then
        If objCADPEDVENDA.GRAVASTEEL("E") = False Then Exit Sub
    End If
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"

  End If
  
  Call AbilitaCampos
  Call ConfGrid
  Call ConfGridBloqueados
  Call ConfGridReprovados
  Call ConfGridFaturado
  Call ConfGridParaEstoque

  Exit Sub

Err_cmdExclui_Click:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdExclui_Click()", Me.Name, "cmdExclui_Click()", strCAMARQERRO)

End Sub

Private Sub cmdImpressao_Click()
        
On Error GoTo Err_cmdImpressao_Click

    If objFuncoes.ChecaAcesso2("R", strACESSO) = False Then Exit Sub
        
    If intFILIALPED = 0 Then Call ImpPed
    If intFILIALPED = 1 Then Call ImpPedSteel

    Exit Sub

Err_cmdImpressao_Click:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdImpressao_Click()", Me.Name, "cmdImpressao_Click()", strCAMARQERRO)

End Sub

Private Sub cmdInclui_Click()
    
On Error GoTo Err_cmdInclui_Click
    
    If objFuncoes.ChecaAcesso2("I", strACESSO) = False Then Exit Sub
    Call Operacao("I")
    
    Exit Sub
    
Err_cmdInclui_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdInclui_Click()", Me.Name, "cmdInclui_Click()", strCAMARQERRO)
    
End Sub

Private Sub cmdLibera_Click()
    
On Error GoTo Err_cmdLibera_Click
    
    ''If objFuncoes.ChecaAcesso2("L", strAcesso) = False Then Exit Sub
    If stPEDIDOS.Tab = 3 Then Exit Sub
    If (flxGRIDBLOQUADOS.Rows > 1) Or (flxReprovados.Rows - 1) Then
       Call Operacao("N")
    End If
    
    Exit Sub
    
Err_cmdLibera_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdLibera_Click()", Me.Name, "cmdLibera_Click()", strCAMARQERRO)
    
End Sub

Private Sub cmdLIBPDATAPCOTA_Click()

On Error GoTo Err_cmdLIBPDATAPCOTA_Click

    If stPEDIDOS.Tab = 7 Then
        If (grdLIBPDATAPCOTA.Rows - 1) = 0 Then
           MsgBox "ATENÇÂO - Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
        Call Operacao("LC")
    End If
    
    Exit Sub
    
Err_cmdLIBPDATAPCOTA_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdLIBPDATAPCOTA_Click()", Me.Name, "cmdLIBPDATAPCOTA_Click()", strCAMARQERRO)

End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
On Error GoTo Err_Command1_Click
    
    ''If objFuncoes.ChecaAcesso2("V", strAcesso) = False Then Exit Sub
    If stPEDIDOS.Tab = 3 Then Exit Sub
    If VerifNF = False Then Exit Sub
    If (flxGRIDBLOQUADOS.Rows - 1) > 0 Or (flxGRIDPEDIDOS.Rows - 1) > 0 Then
       Call Operacao("R")
    End If
    
    Exit Sub
    
Err_Command1_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command1_Click()", Me.Name, "Command1_Click()", strCAMARQERRO)
    
End Sub

Private Sub Command2_Click()
    
On Error GoTo Err_Command2_Click

''    If objFuncoes.ChecaAcesso2("L", strAcesso) = False Then Exit Sub
    If stPEDIDOS.Tab = 3 Then Exit Sub
    If (flxGRIDBLOQUADOS.Rows > 1) Or (flxReprovados.Rows - 1) Then
       Call Operacao("L")
    End If

    Exit Sub
    
Err_Command2_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command2_Click()", Me.Name, "Command2_Click()", strCAMARQERRO)
    
End Sub

Private Sub Command3_Click()

On Error GoTo Err_Command3_Click
    
    
    If BREC8.State = 1 Then BREC8.Close
    If BREC9.State = 1 Then BREC9.Close
    If BREC10.State = 1 Then BREC10.Close
    
    
    Dim arrPRODUTOSEXCL   As Variant
    Dim arrORDEMFAT       As Variant
    Dim arrCONFORDFAT     As Variant
    Dim lngQTDREGS        As Long
    Dim lngCODPEDINI      As Long
    Dim lngCODPEDFIN      As Long
    
    arrPRODUTOSEXCL = Empty
    arrORDEMFAT = Empty
    arrCONFORDFAT = Empty
    
    lngQTDREGS = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Count(SGI_CODIGO) As SGI_QTDE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_DATAPED < '02/01/2010'" & vbCrLf

    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then lngQTDREGS = BREC8!SGI_QTDE
    BREC8.Close
    
    If lngQTDREGS > 0 Then
    
        ReDim arrPRODUTOSEXCL(1 To lngQTDREGS) As Long
    
        lngQTDREGS = 1
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_CODIGO " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_DATAPED < '02/01/2010'" & vbCrLf
        sSql = sSql & "Order By SGI_CODIGO"
        
        BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC8.EOF()
        
            arrPRODUTOSEXCL(lngQTDREGS) = BREC8!SGI_CODIGO
        
            lngQTDREGS = (lngQTDREGS + 1)
            BREC8.MoveNext
        Loop
        BREC8.Close
        
        '' ==========================================
        lngCODPEDINI = 0
        lngCODPEDFIN = 0
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       Min(SGI_CODIGO) SGI_CODINI " & vbCrLf
        sSql = sSql & "      ,Max(SGI_CODIGO) SGI_CODFIN " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_DATAPED < '02/01/2010'"
        
        BREC9.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC9.EOF() Then
           lngCODPEDINI = BREC9!SGI_CODINI
           lngCODPEDFIN = BREC9!SGI_CODFIN
        End If
        BREC9.Close
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       Count(SGI_CODORD) As SGI_QTDE " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADORDFATH " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And (SGI_CODPED >= " & lngCODPEDINI & " And SGI_CODPED <= " & lngCODPEDFIN & ")"
        
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then lngQTDREGS = BREC10!SGI_QTDE
        BREC10.Close
        
        If lngQTDREGS > 0 Then
            ReDim arrORDEMFAT(1 To lngQTDREGS) As Long
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "       SGI_CODORD " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADORDFATH " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "  And (SGI_CODPED >= " & lngCODPEDINI & " And SGI_CODPED <= " & lngCODPEDFIN & ")" & vbCrLf
            sSql = sSql & "Order by SGI_CODORD"

            lngQTDREGS = 1
            BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
            Do While Not BREC10.EOF()
                arrORDEMFAT(lngQTDREGS) = BREC10!SGI_CODORD
                lngQTDREGS = (lngQTDREGS + 1)
                BREC10.MoveNext
            Loop
            BREC10.Close
        
        End If
        '' ==========================================
        
        
        '' ==========================================
        sSql = "Select " & vbCrLf
        sSql = sSql & "       Min(SGI_CODORD) SGI_CODINI " & vbCrLf
        sSql = sSql & "      ,Max(SGI_CODORD) SGI_CODFIN " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADORDFATH " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And (SGI_CODPED >= " & lngCODPEDINI & " And SGI_CODPED <= " & lngCODPEDFIN & ")"
        
        lngCODPEDINI = 0
        lngCODPEDFIN = 0
        
        BREC9.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC9.EOF() Then
           lngCODPEDINI = BREC9!SGI_CODINI
           lngCODPEDFIN = BREC9!SGI_CODFIN
        End If
        BREC9.Close
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       Count(SGI_CODORD) as SGI_QTDE " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADORDCONFH " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And (SGI_CODORD >= " & lngCODPEDINI & " And SGI_CODORD <= " & lngCODPEDFIN & ")"
        
        lngQTDREGS = 0
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then lngQTDREGS = BREC10!SGI_QTDE
        BREC10.Close
        
        If lngQTDREGS > 0 Then
            ReDim arrCONFORDFAT(1 To lngQTDREGS) As Long
        
            lngQTDREGS = 1
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       SGI_CODCONF " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADORDCONFH " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
            sSql = sSql & "   And (SGI_CODORD >= " & lngCODPEDINI & " And SGI_CODORD <= " & lngCODPEDFIN & ")" & vbCrLf
            sSql = sSql & "Order By SGI_CODCONF "
        
            BREC9.Open sSql, adoBanco_Dados, adOpenDynamic
            Do While Not BREC9.EOF()
                arrCONFORDFAT(lngQTDREGS) = BREC9!SGI_CODCONF
                lngQTDREGS = (lngQTDREGS + 1)
                BREC9.MoveNext
            Loop
            BREC9.Close
        
        End If
        '' ==========================================
        
        objCADPEDVENDA.PRODUTOS = arrPRODUTOSEXCL
        objCADPEDVENDA.SERVICOS = arrCONFORDFAT
        objCADPEDVENDA.PROGENTREGAS = arrORDEMFAT
        
        If objCADPEDVENDA.GRAVA("T") = False Then Exit Sub
    
    Else
        MsgBox "Não há dados para Excluir !!!", vbOKOnly + vbExclamation, "Aviso"
    End If

    Exit Sub

Err_Command3_Click:
    
    If BREC8.State = 1 Then BREC8.Close
    If BREC9.State = 1 Then BREC9.Close
    If BREC10.State = 1 Then BREC10.Close
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command3_Click()", Me.Name, "Command3_Click()", strCAMARQERRO)

End Sub

Private Sub Command4_Click()

On Error GoTo Err_Command4_Click
  
    ''If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
    If stPEDIDOS.Tab = 1 Or stPEDIDOS.Tab = 2 Then
       MsgBox "Não pode ser liquidado !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    If stPEDIDOS.Tab = 3 Then
          With grdPEDFATURADO
             If Trim(.Cell(flexcpText, .Row, conCOL_SonFat_Situacao)) <> "P" Then
                 MsgBox "Não pode ser liquidado !!!", vbOKOnly + vbCritical, "Aviso"
                 Exit Sub
             End If
          End With
    End If
  
  
    Dim iResp As Integer
    
    iResp = MsgBox("Confirma a liquidação do pedido ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
    If iResp <> 6 Then Exit Sub
  
    If stPEDIDOS.Tab = 3 Then
          With grdPEDFATURADO
                objCADPEDVENDA.STATUS = Trim(.Cell(flexcpText, .Row, conCOL_SonFat_Situacao))
          End With
    ElseIf stPEDIDOS.Tab = 0 Then
          With flxGRIDPEDIDOS
                objCADPEDVENDA.STATUS = Trim(flxGRIDPEDIDOS.TextMatrix(.Row, 4))
          End With
    End If
    Call PegaOPS(Trim(Str(objCADPEDVENDA.CODPEDIDO)))
  
    If intFILIALPED = 0 Then
      If objCADPEDVENDA.GRAVA("M") = False Then Exit Sub
    ElseIf intFILIALPED = 1 Then
      If objCADPEDVENDA.GRAVASTEEL("M") = False Then Exit Sub
    End If
  
    If objCADPEDVENDA.Atualiza("M", Str(objCADPEDVENDA.CODPEDIDO), FILIAL, "frmCADPEDVENDA") = False Then Exit Sub
    
    MsgBox "Pedido liquidado com sucesso !!!", vbOKOnly + vbInformation, "Aviso"

    Call Atualiza_Grid
    Call AbilitaCampos

    Exit Sub
    
Err_Command4_Click:
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command4_Click()", Me.Name, "Command4_Click()", strCAMARQERRO)

End Sub

Private Sub Command5_Click()
    
On Error GoTo Err_Command5_Click
    
    If stPEDIDOS.Tab = 4 Then
        If (grdBLOQALT.Rows - 1) = 0 Then
           MsgBox "ATENÇÂO = Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
        Call Operacao("S")
    End If
    
    Exit Sub
    
Err_Command5_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command5_Click()", Me.Name, "Command5_Click()", strCAMARQERRO)
    
End Sub

Private Sub Command6_Click()
    
On Error GoTo Err_Command6_Click

    If stPEDIDOS.Tab = 5 Then
        If (grdLIBLITO.Rows - 1) = 0 Then
           MsgBox "ATENÇÂO - Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
        Call Operacao("V")
    End If
    
    Exit Sub
    
Err_Command6_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command6_Click()", Me.Name, "Command6_Click()", strCAMARQERRO)
    
End Sub

Private Sub Command7_Click()

On Error GoTo Err_Command7_Click

    If lngCodUsuaro = 1 Or _
       lngCodUsuaro = 2 Or _
       lngCodUsuaro = 16 Or _
       lngCodUsuaro = 3 Or _
       lngCodUsuaro = 36 Or _
       lngCodUsuaro = 0 Then
        If objFuncoes.ChecaAcesso2("B", strACESSO) = False Then Exit Sub
        If (grdPARAEST.Rows - 1) = 0 Then
           MsgBox "ATENÇÂO - Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
        If stPEDIDOS.Tab = 6 Then Call Operacao("V")
    Else
        MsgBox "Não tem permissão para realizar esta operação !!!", vbOKOnly + vbInformation, "Aviso"
    End If
    
    Exit Sub
    
Err_Command7_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command7_Click()", Me.Name, "Command7_Click()", strCAMARQERRO)
    
End Sub


Private Sub flxGRIDBLOQUADOS_Click()
    
On Error GoTo Err_flxGRIDBLOQUADOS_Click
    
    If flxGRIDBLOQUADOS.Rows > 1 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.RowSel, 1), "/", "")))
    If flxGRIDBLOQUADOS.Rows > 1 And _
       Len(Trim(flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.Row, 6))) > 0 Then
       objCADPEDVENDA.CODCOTA = CLng(Trim(Replace(flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.Row, 6), "/", "")))
    End If
    
    Exit Sub
    
Err_flxGRIDBLOQUADOS_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : flxGRIDBLOQUADOS_Click()", Me.Name, "flxGRIDBLOQUADOS_Click()", strCAMARQERRO)
    
End Sub

Private Sub flxGRIDBLOQUADOS_DblClick()
   
On Error GoTo Err_flxGRIDBLOQUADOS_DblClick
   
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   If flxGRIDBLOQUADOS.Rows > 1 Then Operacao "C"
   
   Exit Sub
   
Err_flxGRIDBLOQUADOS_DblClick:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : flxGRIDBLOQUADOS_DblClick()", Me.Name, "flxGRIDBLOQUADOS_DblClick()", strCAMARQERRO)
   
End Sub

Private Sub flxGRIDBLOQUADOS_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo Err_flxGRIDBLOQUADOS_KeyDown
    
    With flxGRIDBLOQUADOS
        If KeyCode = vbKeyReturn Then
            If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
            If .Rows > 1 Then Operacao "C"
        ElseIf KeyCode = vbKeySpace Then
            If Len(Trim(.TextMatrix(.Row, 7))) > 0 Then
               .TextMatrix(.Row, 7) = ""
            ElseIf Len(Trim(.TextMatrix(.Row, 7))) = 0 Then
               .TextMatrix(.Row, 7) = "*"
            End If
        End If
    End With
    
    Exit Sub
    
Err_flxGRIDBLOQUADOS_KeyDown:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : flxGRIDBLOQUADOS_KeyDown()", Me.Name, "flxGRIDBLOQUADOS_KeyDown()", strCAMARQERRO)
    
End Sub


Private Sub flxGRIDBLOQUADOS_RowColChange()
    
On Error GoTo Err_flxGRIDBLOQUADOS_RowColChange
    
    If flxGRIDBLOQUADOS.Rows > 1 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.RowSel, 1), "/", "")))
    If flxGRIDBLOQUADOS.Rows > 1 And _
       Len(Trim(flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.Row, 6))) > 0 Then
       objCADPEDVENDA.CODCOTA = CLng(Trim(Replace(flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.Row, 6), "/", "")))
    End If
    
    Exit Sub
    
Err_flxGRIDBLOQUADOS_RowColChange:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : flxGRIDBLOQUADOS_RowColChange()", Me.Name, "flxGRIDBLOQUADOS_RowColChange()", strCAMARQERRO)
    
End Sub

Private Sub flxGRIDPEDIDOS_Click()
   
On Error GoTo Err_flxGRIDPEDIDOS_Click

   If flxGRIDPEDIDOS.Rows > 1 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(flxGRIDPEDIDOS.TextMatrix(flxGRIDPEDIDOS.RowSel, 1), "/", "")))
   
   If flxGRIDPEDIDOS.Rows > 1 And _
      Len(Trim(flxGRIDPEDIDOS.TextMatrix(flxGRIDPEDIDOS.Row, 6))) > 0 Then
      objCADPEDVENDA.CODCOTA = CLng(Trim(Replace(flxGRIDPEDIDOS.TextMatrix(flxGRIDPEDIDOS.Row, 6), "/", "")))
   End If
    
   Exit Sub
   
Err_flxGRIDPEDIDOS_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : flxGRIDPEDIDOS_Click()", Me.Name, "flxGRIDPEDIDOS_Click()", strCAMARQERRO)
   

End Sub

Private Sub flxGRIDPEDIDOS_DblClick()
   
On Error GoTo Err_flxGRIDPEDIDOS_DblClick
   
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   If flxGRIDPEDIDOS.Rows > 1 Then Operacao "C"
   
   Exit Sub
   
Err_flxGRIDPEDIDOS_DblClick:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : flxGRIDPEDIDOS_DblClick()", Me.Name, "flxGRIDPEDIDOS_DblClick()", strCAMARQERRO)
   
End Sub

Private Sub flxGRIDPEDIDOS_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo Err_flxGRIDPEDIDOS_KeyDown
    
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        If flxGRIDPEDIDOS.Rows > 1 Then Operacao "C"
    End If
    
    Exit Sub
    
Err_flxGRIDPEDIDOS_KeyDown:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : flxGRIDPEDIDOS_KeyDown()", Me.Name, "flxGRIDPEDIDOS_KeyDown()", strCAMARQERRO)
    
End Sub

Private Sub flxGRIDPEDIDOS_RowColChange()
   
On Error GoTo Err_flxGRIDPEDIDOS_RowColChange
   
   If flxGRIDPEDIDOS.Rows > 1 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(flxGRIDPEDIDOS.TextMatrix(flxGRIDPEDIDOS.RowSel, 1), "/", "")))
   If flxGRIDPEDIDOS.Rows > 1 And _
      Len(Trim(flxGRIDPEDIDOS.TextMatrix(flxGRIDPEDIDOS.Row, 6))) > 0 Then
      objCADPEDVENDA.CODCOTA = CLng(Trim(Replace(flxGRIDPEDIDOS.TextMatrix(flxGRIDPEDIDOS.Row, 6), "/", "")))
   End If
   
   Exit Sub
   
Err_flxGRIDPEDIDOS_RowColChange:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : flxGRIDPEDIDOS_RowColChange()", Me.Name, "flxGRIDPEDIDOS_RowColChange()", strCAMARQERRO)
   
   
End Sub

Private Sub flxReprovados_Click()
    If flxReprovados.Rows > 1 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(flxReprovados.TextMatrix(flxReprovados.RowSel, 1), "/", "")))
End Sub

Private Sub flxReprovados_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   If flxReprovados.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxReprovados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        If flxReprovados.Rows > 1 Then Operacao "C"
    End If
End Sub

Private Sub flxReprovados_RowColChange()
   If flxReprovados.Rows > 1 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(flxReprovados.TextMatrix(flxReprovados.RowSel, 1), "/", "")))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
   
On Error GoTo Err_Form_Load
   
   strCAMARQERRO = Right(Linha(9), Len(Trim(Linha(9))) - 8)
   
   Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
   Set objCADPEDVENDA = CreateObject("CADPEDVENDA.clsCADPEDVENDA")
   Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
   objCADPEDVENDA.FILIAL = FILIAL
   objFuncoes.LimpaCampos frmCADPEDVENDAP
    
   Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
    
   cTipOper = ""
   strOperacao = ""
   lngCodVendedor = PegaCodVendedor(strUSUARIO)
   
   Label3(0).Caption = "F - Faturado Total / P - Faturado Parcial / M - Manual"
   Label3(1).Caption = "L = Liberado"
   Label3(2).Caption = "N - Liberado Pelo Comercial / B = Bloqueado"
   Label3(3).Caption = "R - Reprovado"
   
   stPEDIDOS.Tab = 0
   
   Call AbilitaCampos
   Call ConfGrid
   Call ConfGridBloqueados
   Call ConfGridReprovados
   Call ConfGridFaturado
   Call ConfGridBloqAlt
   Call ConfGridBloqAltLit
   Call ConfGridParaEstoque
   Call ConfGridBloqPDataPCota
   
   Call ConfComboFiltro
   Call AtivaDesativaBotoes
   
    Command3.Visible = False
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    boolComAcao = False
    
    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO
    
    objCADPEDVENDA.FILIALPED = intFILIALPED
    
    Exit Sub
    
Err_Form_Load:
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Form_Load()", Me.Name, "Form_Load()", strCAMARQERRO)
    
End Sub

Private Sub AbilitaCampos()

On Error GoTo Err_AbilitaCampos
    
    If objCADPEDVENDA.Pesq_CadPedido = False Then
       cmdAltera.Enabled = False
       cmdExclui.Enabled = False
       Frame1.Enabled = False
    Else
       cmdAltera.Enabled = True
       cmdExclui.Enabled = True
       Frame1.Enabled = True
    End If

    Exit Sub
    
Err_AbilitaCampos:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : AbilitaCampos()", Me.Name, "AbilitaCampos()", strCAMARQERRO)
    

End Sub

Private Sub ConfGrid()
    
    
    flxGRIDPEDIDOS.Rows = 1
    flxGRIDPEDIDOS.Cols = 8
    
    flxGRIDPEDIDOS.AllowBigSelection = False
    
    flxGRIDPEDIDOS.TextMatrix(0, 0) = ""
    flxGRIDPEDIDOS.TextMatrix(0, 1) = "Cód. Ped"
    flxGRIDPEDIDOS.TextMatrix(0, 2) = "Data"
    flxGRIDPEDIDOS.TextMatrix(0, 3) = "Cliente"
    flxGRIDPEDIDOS.TextMatrix(0, 4) = "S"
    flxGRIDPEDIDOS.TextMatrix(0, 5) = "Tipo Pedido"
    flxGRIDPEDIDOS.TextMatrix(0, 6) = "Cotação"
    flxGRIDPEDIDOS.TextMatrix(0, 7) = "Cotação"
    
    flxGRIDPEDIDOS.ColWidth(0) = 0
    flxGRIDPEDIDOS.ColWidth(1) = 1300
    flxGRIDPEDIDOS.ColWidth(2) = 1000
    flxGRIDPEDIDOS.ColWidth(3) = 5500
    flxGRIDPEDIDOS.ColWidth(4) = 200
    flxGRIDPEDIDOS.ColWidth(5) = 2000
    flxGRIDPEDIDOS.ColWidth(6) = 0
    flxGRIDPEDIDOS.ColWidth(7) = 0
    
End Sub


Private Sub PreencheGrid()

On Error GoTo Err_PreencheGrid
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
    sSql = sSql & "   And PED.SGI_STATUS = 'L' " & vbCrLf
    
    If lngCodVendedor > 0 Then
        sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
    End If
    
    sSql = sSql & "Order by PED.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       flxGRIDPEDIDOS.AddItem "" & vbTab & _
                              Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                              Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                              BREC!SGI_RAZAOSOC & vbTab & _
                              BREC!SGI_STATUS & vbTab & _
                              BREC!SGI_DESCRICAO & vbTab & _
                              IIf(IsNull(BREC!SGI_CODCOTA) = False, Format(BREC!SGI_CODCOTA, "#/####"), "") & vbTab & _
                              BREC!SGI_CODTIPORC
       BREC.MoveNext
    Loop
    
    Call PosGrid
    
    BREC.Close
    
    Exit Sub

Err_PreencheGrid:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PreencheGrid()", Me.Name, "PreencheGrid()", strCAMARQERRO)
    
End Sub

Private Sub PosGrid()

On Error GoTo Err_PosGrid

    If iCodigo > 0 Then
        Dim I As Integer
        
        
        For I = 1 To (flxGRIDPEDIDOS.Rows - 1)
             
            If CLng(Trim(Replace(flxGRIDPEDIDOS.TextMatrix(I, 1), "/", ""))) = iCodigo Then
               flxGRIDPEDIDOS.Row = I
               flxGRIDPEDIDOS.Col = 1
               
               Exit For
            End If
         
        Next I
    End If
    
    Exit Sub
    
Err_PosGrid:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PosGrid()", Me.Name, "PosGrid()", strCAMARQERRO)

End Sub

Private Sub Operacao(strOperacao As String)
 
On Error GoTo Err_Operacao
  
  Dim Pesquisa As String
  
  If stPEDIDOS.Tab = 0 Then
     If flxGRIDPEDIDOS.Rows > 1 Then iCodigo = CLng(Trim(Replace(flxGRIDPEDIDOS.TextMatrix(flxGRIDPEDIDOS.RowSel, 1), "/", "")))
  
     If strOperacao = "L" Or strOperacao = "N" Then
        MsgBox "Pedido já Liberado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
  
  End If
  If stPEDIDOS.Tab = 1 Then
     If flxGRIDBLOQUADOS.Rows > 1 Then iCodigo = CLng(Trim(Replace(flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.RowSel, 1), "/", "")))
     
     If strOperacao = "L" And flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.Row, 4) = "B" Then
        MsgBox "Pedido Ainda Não Foi Liberado pelo Comercial !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     ElseIf strOperacao = "N" And flxGRIDBLOQUADOS.TextMatrix(flxGRIDBLOQUADOS.Row, 4) = "N" Then
        MsgBox "Pedido ja Foi Liberado pelo Comercial !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
  
  End If
  If stPEDIDOS.Tab = 2 Then
     If flxReprovados.Rows > 1 Then iCodigo = CLng(Trim(Replace(flxReprovados.TextMatrix(flxReprovados.RowSel, 1), "/", "")))
  End If
  If stPEDIDOS.Tab = 3 Then
     If (grdPEDFATURADO.Rows - 1) > 0 And grdPEDFATURADO.Row > 0 Then iCodigo = CLng(Trim(Replace(grdPEDFATURADO.Cell(flexcpText, grdPEDFATURADO.Row, conCOL_SonFat_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 4 Then
     If (grdBLOQALT.Rows - 1) > 0 And grdBLOQALT.Row > 0 Then iCodigo = CLng(Trim(Replace(grdBLOQALT.Cell(flexcpText, grdBLOQALT.Row, conCOL_SonBloq_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 5 Then
     If (grdLIBLITO.Rows - 1) > 0 And grdLIBLITO.Row > 0 Then iCodigo = CLng(Trim(Replace(grdLIBLITO.Cell(flexcpText, grdLIBLITO.Row, conCOL_SonBloqLit_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 6 Then
     If (grdPARAEST.Rows - 1) > 0 And grdPARAEST.Row > 0 Then iCodigo = CLng(Trim(Replace(grdPARAEST.Cell(flexcpText, grdPARAEST.Row, conCOL_SonParaEst_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 7 Then
     If (grdLIBPDATAPCOTA.Rows - 1) > 0 And grdLIBPDATAPCOTA.Row > 0 Then iCodigo = CLng(Trim(Replace(grdLIBPDATAPCOTA.Cell(flexcpText, grdLIBPDATAPCOTA.Row, conCOL_SonBloqPDPC_Codigo), "/", "")))
  End If
    
  boolComAcao = True
  
  frmCADPEDVENDA.cCaminho = cCaminho
  frmCADPEDVENDA.Linha = Linha
  frmCADPEDVENDA.iCodigo = iCodigo
  frmCADPEDVENDA.cTipOper = strOperacao
  frmCADPEDVENDA.FILIAL = FILIAL
  frmCADPEDVENDA.strACESSO = strACESSO
  frmCADPEDVENDA.strMODPAI = Me.Name
  frmCADPEDVENDA.strUSUARIO = strUSUARIO
  frmCADPEDVENDA.lngCodVendedor = lngCodVendedor
  frmCADPEDVENDA.lngCodUsuario = lngCodUsuaro
  frmCADPEDVENDA.intFILIALPED = intFILIALPED
  frmCADPEDVENDA.boolSomenteCons = False
  frmCADPEDVENDA.strVERSAO = strVERSAO
  frmCADPEDVENDA.strNOMCOMP = strNOMCOMP
  frmCADPEDVENDA.Show vbModal
  
  boolComAcao = False
  
  Call AbilitaCampos
  Call ConfGrid
  Call ConfGridBloqueados
  Call ConfGridReprovados
  Call ConfGridFaturado
  Call ConfGridBloqAlt
  Call ConfGridBloqAltLit
  Call ConfGridParaEstoque
  Call ConfGridBloqPDataPCota

  Exit Sub
  
Err_Operacao:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Operacao()", Me.Name, "Operacao()", strCAMARQERRO)

End Sub

Private Sub ConfGridBloqueados()
    
    
    flxGRIDBLOQUADOS.Rows = 1
    flxGRIDBLOQUADOS.Cols = 8
    flxGRIDBLOQUADOS.AllowBigSelection = False
    
    flxGRIDBLOQUADOS.TextMatrix(0, 0) = ""
    flxGRIDBLOQUADOS.TextMatrix(0, 1) = "Cód. Ped"
    flxGRIDBLOQUADOS.TextMatrix(0, 2) = "Data"
    flxGRIDBLOQUADOS.TextMatrix(0, 3) = "Cliente"
    flxGRIDBLOQUADOS.TextMatrix(0, 4) = "S"
    flxGRIDBLOQUADOS.TextMatrix(0, 5) = "Tipo Pedido"
    flxGRIDBLOQUADOS.TextMatrix(0, 6) = "Cotação"
    flxGRIDBLOQUADOS.TextMatrix(0, 7) = " "
    
    flxGRIDBLOQUADOS.ColWidth(0) = 0
    flxGRIDBLOQUADOS.ColWidth(1) = 1300
    flxGRIDBLOQUADOS.ColWidth(2) = 1000
    flxGRIDBLOQUADOS.ColWidth(3) = 5500
    flxGRIDBLOQUADOS.ColWidth(4) = 200
    flxGRIDBLOQUADOS.ColWidth(5) = 2000
    flxGRIDBLOQUADOS.ColWidth(6) = 0
    flxGRIDBLOQUADOS.ColWidth(7) = 0
    
End Sub

Private Sub PreencheGridBloqueado()

On Error GoTo Err_PreencheGridBloqueado
    
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL    = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO    = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL    = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO    = PED.SGI_CODTIPORC " & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
    
    If lngCodVendedor > 0 Then
        sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
    End If
    
    sSql = sSql & "   And (PED.SGI_STATUS = 'B' or PED.SGI_STATUS = 'N')" & vbCrLf
    sSql = sSql & "Order by PED.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       flxGRIDBLOQUADOS.AddItem "" & vbTab & _
                              Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                              Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                              BREC!SGI_RAZAOSOC & vbTab & _
                              BREC!SGI_STATUS & vbTab & _
                              BREC!SGI_DESCRICAO & vbTab & _
                              IIf(IsNull(BREC!SGI_CODCOTA) = False, Format(BREC!SGI_CODCOTA, "#/####"), "")
                              
       BREC.MoveNext
    Loop
    BREC.Close
    
    Exit Sub
    
Err_PreencheGridBloqueado:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PreencheGridBloqueado()", Me.Name, "PreencheGridBloqueado()", strCAMARQERRO)
    
End Sub

Private Function Verif_Cota(lngCODCOTA As Long) As Boolean

On Error GoTo Err_Verif_Cota
    
    Verif_Cota = False
    
    If BREC.State = 1 Then BREC.Close
    
    Dim intQTD As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODCOTA = " & lngCODCOTA
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    intQTD = 0
    Do While Not BREC.EOF
       intQTD = intQTD + 1
       BREC.MoveNext
    Loop
    BREC.Close
    
    If intQTD > 1 Then
       MsgBox "Impossivel alterar, Cotações desmenbradas em varios pedidos !!!", vbOKOnly + vbExclamation, "Aviso"
       Verif_Cota = True
    End If
    
    Exit Function
    
Err_Verif_Cota:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Err_Verif_Cota()", Me.Name, "Err_Verif_Cota()", strCAMARQERRO)
    
End Function

Private Sub ConfGridReprovados()
    
    flxReprovados.Rows = 1
    flxReprovados.Cols = 7
    flxReprovados.AllowBigSelection = False
    
    flxReprovados.TextMatrix(0, 0) = ""
    flxReprovados.TextMatrix(0, 1) = "Cód. Ped"
    flxReprovados.TextMatrix(0, 2) = "Data"
    flxReprovados.TextMatrix(0, 3) = "Cliente"
    flxReprovados.TextMatrix(0, 4) = "S"
    flxReprovados.TextMatrix(0, 5) = "Tipo Pedido"
    flxReprovados.TextMatrix(0, 6) = "Cotação"
    
    flxReprovados.ColWidth(0) = 0
    flxReprovados.ColWidth(1) = 1300
    flxReprovados.ColWidth(2) = 1000
    flxReprovados.ColWidth(3) = 5500
    flxReprovados.ColWidth(4) = 200
    flxReprovados.ColWidth(5) = 2000
    flxReprovados.ColWidth(6) = 0
    
End Sub

Private Sub PreencheGridReprovados()

On Error GoTo Err_PreencheGridReprovados
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
    
    If lngCodVendedor > 0 Then
        sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
    End If
    
    sSql = sSql & "   And PED.SGI_STATUS = 'R' " & vbCrLf
    sSql = sSql & "Order by PED.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       
       flxReprovados.AddItem "" & vbTab & _
                             Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                             Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                             BREC!SGI_RAZAOSOC & vbTab & _
                             BREC!SGI_STATUS & vbTab & _
                             BREC!SGI_DESCRICAO & vbTab & _
                             IIf(IsNull(BREC!SGI_CODCOTA) = False, Format(BREC!SGI_CODCOTA, "#/####"), "")
                              
       BREC.MoveNext
    Loop
    BREC.Close
    
    Exit Sub

Err_PreencheGridReprovados:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PreencheGridReprovados()", Me.Name, "PreencheGridReprovados()", strCAMARQERRO)
    
End Sub

Private Function VerifNF() As Boolean
    
On Error GoTo Err_VerifNF
    
    VerifNF = True
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED = " & objCADPEDVENDA.CODPEDIDO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Operação inválida já foi emitido ordem de faturamento para este pedido !!!", vbOKOnly + vbExclamation, "Aviso"
       VerifNF = False
    End If
    BREC.Close
    
    Exit Function
    
Err_VerifNF:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : VerifNF()", Me.Name, "VerifNF()", strCAMARQERRO)
    
End Function

Private Sub ConfGridFaturado()
        
    With grdPEDFATURADO
    
       .Cols = conColumnsIn_SonFat
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonFat_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonFat_Codigo) = ""
       .ColDataType(conCOL_SonFat_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonFat_Data) = ""
       .ColDataType(conCOL_SonFat_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonFat_Cliente) = ""
       .ColDataType(conCOL_SonFat_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonFat_Situacao) = ""
       .ColDataType(conCOL_SonFat_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonFat_Tipo) = ""
       .ColDataType(conCOL_SonFat_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonFat_Cotacao) = ""
       .ColDataType(conCOL_SonFat_Cotacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonFat_Status) = ""
       .ColDataType(conCOL_SonFat_Status) = flexDTString
       
       .ColWidth(conCOL_SonFat_Codigo) = 1300
       .ColWidth(conCOL_SonFat_Data) = 1000
       .ColWidth(conCOL_SonFat_Cliente) = 5500
       .ColWidth(conCOL_SonFat_Situacao) = 250
       .ColWidth(conCOL_SonFat_Tipo) = 2000
       .ColWidth(conCOL_SonFat_Cotacao) = 0
       .ColWidth(conCOL_SonFat_Status) = 0
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub

Private Sub PreencheGridFaturado()

On Error GoTo Err_PreencheGridFaturado
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
    
    If lngCodVendedor > 0 Then
        sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
    End If
    
    sSql = sSql & "   And (PED.SGI_STATUS = 'F' or PED.SGI_STATUS = 'P' or PED.SGI_STATUS = 'M') " & vbCrLf
    sSql = sSql & "Order by PED.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       grdPEDFATURADO.AddItem Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                              Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                              BREC!SGI_RAZAOSOC & vbTab & _
                              BREC!SGI_STATUS & vbTab & _
                              BREC!SGI_DESCRICAO & vbTab & _
                              IIf(IsNull(BREC!SGI_CODCOTA) = False, Format(BREC!SGI_CODCOTA, "#/####"), "") & vbTab & _
                              PegaNF(BREC!SGI_CODIGO)
                              
       BREC.MoveNext
    Loop
    BREC.Close
    
    Exit Sub
    
Err_PreencheGridFaturado:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PreencheGridFaturado()", Me.Name, "PreencheGridFaturado()", strCAMARQERRO)
    
End Sub

Private Function PegaNF(lngCODPED As Long) As String
    PegaNF = ""
End Function


Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub

Private Sub grdBLOQALT_Click()
    With grdBLOQALT
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloq_Codigo), "/", "")))
    End With
End Sub

Private Sub grdBLOQALT_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdBLOQALT
        If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdBLOQALT_RowColChange()
    With grdBLOQALT
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloq_Codigo), "/", "")))
    End With
End Sub

Private Sub grdLIBLITO_Click()
    With grdLIBLITO
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloqLit_Codigo), "/", "")))
    End With
End Sub

Private Sub grdLIBLITO_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdLIBLITO
        If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdLIBLITO_RowColChange()
    With grdLIBLITO
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloqLit_Codigo), "/", "")))
    End With
End Sub

Private Sub grdLIBPDATAPCOTA_Click()
    With grdLIBPDATAPCOTA
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloqPDPC_Codigo), "/", "")))
    End With
End Sub

Private Sub grdLIBPDATAPCOTA_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdLIBPDATAPCOTA
        If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdLIBPDATAPCOTA_RowColChange()
    With grdLIBPDATAPCOTA
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloqPDPC_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPARAEST_Click()
    With grdPARAEST
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonParaEst_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPARAEST_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdPARAEST
        If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdPARAEST_RowColChange()
    With grdPARAEST
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonParaEst_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPEDFATURADO_Click()
    With grdPEDFATURADO
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonFat_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPEDFATURADO_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdPEDFATURADO
        If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdPEDFATURADO_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        With grdPEDFATURADO
            If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("C")
        End With
    End If
End Sub

Private Sub grdPEDFATURADO_RowColChange()
    With grdPEDFATURADO
        If (.Rows - 1) > 0 And .Row > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonFat_Codigo), "/", "")))
    End With
End Sub

Private Sub optLiberados_Click(Index As Integer)
    
On Error GoTo Err_optLiberados_Click
    
    Dim strCAMPO  As String
    
    If BREC.State = 1 Then BREC.Close
    
    If stPEDIDOS.Tab <> 1 Then Exit Sub
    
    Call ConfGridBloqueados
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
    
    If Index = 0 Then sSql = sSql & "   And (PED.SGI_STATUS = 'B')" & vbCrLf
    If Index = 1 Then sSql = sSql & "   And (PED.SGI_STATUS = 'N')" & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by PED.SGI_CODIGO "
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by CLI.SGI_RAZAOSOC "
    If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order by PED.SGI_DATAPED "
    If cboFiltro.ListIndex = 3 Then sSql = sSql & "Order by ESP.SGI_DESCRICAO "

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
      
    Do While Not BREC.EOF
       
       strCAMPO = "" & vbTab & _
                  Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                  Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                  BREC!SGI_RAZAOSOC & vbTab & _
                  BREC!SGI_STATUS & vbTab & _
                  BREC!SGI_DESCRICAO & vbTab & _
                  IIf(IsNull(BREC!SGI_CODCOTA) = False, Format(BREC!SGI_CODCOTA, "#/####"), "")
       
       flxGRIDBLOQUADOS.AddItem strCAMPO
       
       BREC.MoveNext
    Loop
    
    BREC.Close

    Exit Sub

Err_optLiberados_Click:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : optLiberados_Click()", Me.Name, "optLiberados_Click()", strCAMARQERRO)

End Sub

Private Sub stPEDIDOS_Click(PreviousTab As Integer)
    Call ConfComboFiltro
End Sub

Private Sub Atualiza_Grid()
    
On Error GoTo Err_Atualiza_Grid
     
     If BREC.State = 1 Then BREC.Close
     If BREC2.State = 1 Then BREC2.Close
     
     If boolComAcao = True Then Exit Sub
     
     Dim I        As Integer
     Dim bolAchou As Boolean
      
     If strOperacao = "P" Then Exit Sub
     
     bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADPEDVENDA'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        If stPEDIDOS.Tab = 0 Then
            For I = 1 To (flxGRIDPEDIDOS.Rows - 1)
                If Trim(BREC!SGI_ACAO) = "E" Or Trim(BREC!SGI_ACAO) = "D" Or Trim(BREC!SGI_ACAO) = "R" Then
                   If Trim(Replace(flxGRIDPEDIDOS.TextMatrix(I, 1), "/", "")) = Trim(BREC!SGI_CODIGO) Then
                      If flxGRIDPEDIDOS.Rows = 2 Then flxGRIDPEDIDOS.Rows = 1
                      If flxGRIDPEDIDOS.Rows > 2 Then flxGRIDPEDIDOS.RemoveItem I
                      Exit For
                   End If
                ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Or Trim(BREC!SGI_ACAO) = "L" Then
                   If Trim(BREC!SGI_CODIGO) = Trim(Replace(flxGRIDPEDIDOS.TextMatrix(I, 1), "/", "")) Then
                      bolAchou = True
                      Exit For
                   End If
                End If
            Next I
        ElseIf stPEDIDOS.Tab = 1 Then
            For I = 1 To (flxGRIDBLOQUADOS.Rows - 1)
                If Trim(BREC!SGI_ACAO) = "E" Or Trim(BREC!SGI_ACAO) = "R" Or Trim(BREC!SGI_ACAO) = "L" Then
                   If Trim(Replace(flxGRIDBLOQUADOS.TextMatrix(I, 1), "/", "")) = Trim(BREC!SGI_CODIGO) Then
                      If flxGRIDBLOQUADOS.Rows = 2 Then flxGRIDBLOQUADOS.Rows = 1
                      If flxGRIDBLOQUADOS.Rows > 2 Then flxGRIDBLOQUADOS.RemoveItem I
                      Exit For
                   End If
                ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Or Trim(BREC!SGI_ACAO) = "D" Or Trim(BREC!SGI_ACAO) = "N" Then
                   If Trim(BREC!SGI_CODIGO) = Trim(Replace(flxGRIDBLOQUADOS.TextMatrix(I, 1), "/", "")) Then
                      bolAchou = True
                      Exit For
                   End If
                End If
            Next I
        ElseIf stPEDIDOS.Tab = 2 Then
            For I = 1 To (flxReprovados.Rows - 1)
                If Trim(BREC!SGI_ACAO) = "E" Or Trim(BREC!SGI_ACAO) = "L" Then
                   If Trim(Replace(flxReprovados.TextMatrix(I, 1), "/", "")) = Trim(BREC!SGI_CODIGO) Then
                      If flxReprovados.Rows = 2 Then flxReprovados.Rows = 1
                      If flxReprovados.Rows > 2 Then flxReprovados.RemoveItem I
                      Exit For
                   End If
                ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Or Trim(BREC!SGI_ACAO) = "R" Then
                   If Trim(BREC!SGI_CODIGO) = Trim(Replace(flxReprovados.TextMatrix(I, 1), "/", "")) Then
                      bolAchou = True
                      Exit For
                   End If
                End If
            Next I
        ElseIf stPEDIDOS.Tab = 3 Then
            For I = 1 To (grdPEDFATURADO.Rows - 1)
                If Trim(BREC!SGI_ACAO) = "E" Then
                   If Trim(Replace(grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Codigo), "/", "")) = Trim(BREC!SGI_CODIGO) Then
                      If grdPEDFATURADO.Rows = 2 Then grdPEDFATURADO.Rows = 1
                      If grdPEDFATURADO.Rows > 2 Then grdPEDFATURADO.RemoveItem I
                      Exit For
                   End If
                ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
                   If Trim(BREC!SGI_CODIGO) = Trim(Replace(grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Codigo), "/", "")) Then
                      bolAchou = True
                      Exit For
                   End If
                End If
            Next I
        End If
        
        If bolAchou = False And (Trim(BREC!SGI_ACAO) = "I" Or BREC!SGI_ACAO = "D" Or BREC!SGI_ACAO = "R" Or BREC!SGI_ACAO = "L" Or BREC!SGI_ACAO = "N") Then
        
            If stPEDIDOS.Tab = 0 Then
            
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PED.* " & vbCrLf
                sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
                sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
                sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
                sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
                sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
                sSql = sSql & "   And PED.SGI_STATUS = 'L' " & vbCrLf
                sSql = sSql & "   And PED.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
                
                If lngCodVendedor > 0 Then
                    sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
                End If
                
                sSql = sSql & "Order by PED.SGI_CODIGO "
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                
                Do While Not BREC2.EOF
                
                   flxGRIDPEDIDOS.AddItem "" & vbTab & _
                                          Format(BREC2!SGI_CODIGO, "#/####") & vbTab & _
                                          Format(BREC2!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                          BREC2!SGI_RAZAOSOC & vbTab & _
                                          BREC2!SGI_STATUS & vbTab & _
                                          BREC2!SGI_DESCRICAO & vbTab & _
                                          IIf(IsNull(BREC2!SGI_CODCOTA) = False, Format(BREC2!SGI_CODCOTA, "#/####"), "")
                                          
                   BREC2.MoveNext
                Loop
                BREC2.Close
            
            ElseIf stPEDIDOS.Tab = 1 Then
            
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PED.* " & vbCrLf
                sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
                sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
                sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
                sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
                sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
                sSql = sSql & "   And (PED.SGI_STATUS = 'B' or PED.SGI_STATUS = 'N')" & vbCrLf
                sSql = sSql & "   And PED.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
                
                If lngCodVendedor > 0 Then
                    sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
                End If
                
                sSql = sSql & "Order by PED.SGI_CODIGO "
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                
                Do While Not BREC2.EOF
                   
                   flxGRIDBLOQUADOS.AddItem "" & vbTab & _
                                            Format(BREC2!SGI_CODIGO, "#/####") & vbTab & _
                                            Format(BREC2!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                            BREC2!SGI_RAZAOSOC & vbTab & _
                                            BREC2!SGI_STATUS & vbTab & _
                                            BREC2!SGI_DESCRICAO & vbTab & _
                                            IIf(IsNull(BREC2!SGI_CODCOTA) = False, Format(BREC2!SGI_CODCOTA, "#/####"), "")
                                          
                   BREC2.MoveNext
                Loop
                
                BREC2.Close
            
            ElseIf stPEDIDOS.Tab = 2 Then
            
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PED.* " & vbCrLf
                sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
                sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
                sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
                sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
                sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
                sSql = sSql & "   And PED.SGI_STATUS = 'R' " & vbCrLf
                sSql = sSql & "   And PED.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
                
                If lngCodVendedor > 0 Then
                    sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
                End If
                
                sSql = sSql & "Order by PED.SGI_CODIGO "
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                
                Do While Not BREC2.EOF
                   
                   flxReprovados.AddItem "" & vbTab & _
                                         Format(BREC2!SGI_CODIGO, "#/####") & vbTab & _
                                         Format(BREC2!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                         BREC2!SGI_RAZAOSOC & vbTab & _
                                         BREC2!SGI_STATUS & vbTab & _
                                         BREC2!SGI_DESCRICAO & vbTab & _
                                         IIf(IsNull(BREC2!SGI_CODCOTA) = False, Format(BREC2!SGI_CODCOTA, "#/####"), "")
                                          
                   BREC2.MoveNext
                Loop
                
                BREC2.Close
            
            ElseIf stPEDIDOS.Tab = 3 Then
            
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PED.* " & vbCrLf
                sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
                sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
                sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
                sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
                sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
                sSql = sSql & "   And PED.SGI_STATUS = 'F' " & vbCrLf
                sSql = sSql & "   And PED.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
                
                If lngCodVendedor > 0 Then
                    sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
                End If
                
                sSql = sSql & "Order by PED.SGI_CODIGO "
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                
                Do While Not BREC2.EOF
                
                   grdPEDFATURADO.AddItem Format(BREC2!SGI_CODIGO, "#/####") & vbTab & _
                                          Format(BREC2!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                          BREC2!SGI_RAZAOSOC & vbTab & _
                                          BREC2!SGI_STATUS & vbTab & _
                                          BREC2!SGI_DESCRICAO & vbTab & _
                                          IIf(IsNull(BREC2!SGI_CODCOTA) = False, Format(BREC2!SGI_CODCOTA, "#/####"), "") & vbTab & _
                                          PegaNF(BREC2!SGI_CODIGO)
                                          
                   BREC2.MoveNext
                Loop
                BREC2.Close
            
            End If
        
        ElseIf bolAchou = True And (BREC!SGI_ACAO = "A" Or BREC!SGI_ACAO = "D" Or BREC!SGI_ACAO = "R" Or BREC!SGI_ACAO = "L" Or BREC!SGI_ACAO = "N") Then
        
           If stPEDIDOS.Tab = 0 Then
           
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PED.* " & vbCrLf
                sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
                sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
                sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
                sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
                sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
                sSql = sSql & "   And PED.SGI_STATUS = 'L' " & vbCrLf
                sSql = sSql & "   And PED.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
                
                If lngCodVendedor > 0 Then
                    sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
                End If
                
                sSql = sSql & "Order by PED.SGI_CODIGO "
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF Then
                   flxGRIDPEDIDOS.TextMatrix(I, 0) = ""
                   flxGRIDPEDIDOS.TextMatrix(I, 1) = Format(BREC2!SGI_CODIGO, "#/####")
                   flxGRIDPEDIDOS.TextMatrix(I, 2) = Format(BREC2!SGI_DATAPED, "DD/MM/YYYY")
                   flxGRIDPEDIDOS.TextMatrix(I, 3) = BREC2!SGI_RAZAOSOC
                   flxGRIDPEDIDOS.TextMatrix(I, 4) = BREC2!SGI_STATUS
                   flxGRIDPEDIDOS.TextMatrix(I, 5) = BREC2!SGI_DESCRICAO
                   flxGRIDPEDIDOS.TextMatrix(I, 6) = IIf(IsNull(BREC2!SGI_CODCOTA) = False, Format(BREC2!SGI_CODCOTA, "#/####"), "")
                End If
                BREC2.Close
            ElseIf stPEDIDOS.Tab = 1 Then
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PED.* " & vbCrLf
                sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
                sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
                sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
                sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
                sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
                sSql = sSql & "   And (PED.SGI_STATUS = 'B' or PED.SGI_STATUS = 'N') " & vbCrLf
                sSql = sSql & "   And PED.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
                
                If lngCodVendedor > 0 Then
                    sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
                End If
                
                sSql = sSql & "Order by PED.SGI_CODIGO "
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF Then
                   flxGRIDBLOQUADOS.TextMatrix(I, 0) = ""
                   flxGRIDBLOQUADOS.TextMatrix(I, 1) = Format(BREC2!SGI_CODIGO, "#/####")
                   flxGRIDBLOQUADOS.TextMatrix(I, 2) = Format(BREC2!SGI_DATAPED, "DD/MM/YYYY")
                   flxGRIDBLOQUADOS.TextMatrix(I, 3) = BREC2!SGI_RAZAOSOC
                   flxGRIDBLOQUADOS.TextMatrix(I, 4) = BREC2!SGI_STATUS
                   flxGRIDBLOQUADOS.TextMatrix(I, 5) = BREC2!SGI_DESCRICAO
                   flxGRIDBLOQUADOS.TextMatrix(I, 6) = IIf(IsNull(BREC2!SGI_CODCOTA) = False, Format(BREC2!SGI_CODCOTA, "#/####"), "")
                End If
                BREC2.Close
            ElseIf stPEDIDOS.Tab = 2 Then
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PED.* " & vbCrLf
                sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
                sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
                sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
                sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
                sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
                sSql = sSql & "   And PED.SGI_STATUS = 'R' " & vbCrLf
                sSql = sSql & "   And PED.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
                
                If lngCodVendedor > 0 Then
                    sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
                End If
                
                sSql = sSql & "Order by PED.SGI_CODIGO "
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                
                If Not BREC2.EOF Then
                   flxReprovados.TextMatrix(I, 0) = ""
                   flxReprovados.TextMatrix(I, 1) = Format(BREC2!SGI_CODIGO, "#/####")
                   flxReprovados.TextMatrix(I, 2) = Format(BREC2!SGI_DATAPED, "DD/MM/YYYY")
                   flxReprovados.TextMatrix(I, 3) = BREC2!SGI_RAZAOSOC
                   flxReprovados.TextMatrix(I, 4) = BREC2!SGI_STATUS
                   flxReprovados.TextMatrix(I, 5) = BREC2!SGI_DESCRICAO
                   flxReprovados.TextMatrix(I, 6) = IIf(IsNull(BREC2!SGI_CODCOTA) = False, Format(BREC2!SGI_CODCOTA, "#/####"), "")
                End If
                BREC2.Close
            ElseIf stPEDIDOS.Tab = 3 Then
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PED.* " & vbCrLf
                sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
                sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH PED " & vbCrLf
                sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
                sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
                sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
                sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
                sSql = sSql & "   And PED.SGI_STATUS = 'F' " & vbCrLf
                sSql = sSql & "   And PED.SGI_CODIGO = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
                
                If lngCodVendedor > 0 Then
                    sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
                End If
                
                sSql = sSql & "Order by PED.SGI_CODIGO "
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                
                If Not BREC2.EOF Then
                   grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Codigo) = Format(BREC2!SGI_CODIGO, "#/####")
                   grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Data) = Format(BREC2!SGI_DATAPED, "DD/MM/YYYY")
                   grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Cliente) = BREC2!SGI_RAZAOSOC
                   grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Situacao) = BREC2!SGI_STATUS
                   grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Tipo) = BREC2!SGI_DESCRICAO
                   grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Cotacao) = IIf(IsNull(BREC2!SGI_CODCOTA) = False, Format(BREC2!SGI_CODCOTA, "#/####"), "")
                   grdPEDFATURADO.Cell(flexcpText, I, conCOL_SonFat_Status) = PegaNF(BREC2!SGI_CODIGO)
                End If
                BREC2.Close
            End If
        End If
        
     End If
     BREC.Close
      
     Exit Sub
     
Err_Atualiza_Grid:
     
     If BREC.State = 1 Then BREC.Close
     If BREC2.State = 1 Then BREC2.Close
     Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Atualiza_Grid()", Me.Name, "Atualiza_Grid()", strCAMARQERRO)
      
End Sub

Private Sub Ordem()

On Error GoTo Err_Ordem
    
    If BREC.State = 1 Then BREC.Close
    
    Call ConfGrid
    Call ConfGridBloqueados
    Call ConfGridReprovados
    Call ConfGridFaturado
    Call ConfGridBloqAlt
    Call ConfGridBloqAltLit
    Call ConfGridParaEstoque
    Call ConfGridBloqPDataPCota
     
    Dim strCAMPO As String
    Dim strNOMETABELA As String
    
    If intFILIALPED = 0 Then strNOMETABELA = "SGI_CADPEDVENDH"
    If intFILIALPED = 1 Then strNOMETABELA = "SGI_CADPEDVENDH_STEEL"
    
    
    txtCampos.Text = ""
    
    sSql = ""
  
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMETABELA & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
    
    If stPEDIDOS.Tab = 0 Then sSql = sSql & "   And PED.SGI_STATUS = 'L' " & vbCrLf
    If stPEDIDOS.Tab = 1 Then sSql = sSql & "   And (PED.SGI_STATUS = 'B' or PED.SGI_STATUS = 'N')" & vbCrLf
    If stPEDIDOS.Tab = 2 Then sSql = sSql & "   And PED.SGI_STATUS = 'R' " & vbCrLf
    If stPEDIDOS.Tab = 3 Then sSql = sSql & "   And (PED.SGI_STATUS = 'F' or PED.SGI_STATUS = 'P' or PED.SGI_STATUS = 'M') " & vbCrLf
    If stPEDIDOS.Tab = 4 Then sSql = sSql & "   And PED.SGI_STATUS = 'S' " & vbCrLf
    If stPEDIDOS.Tab = 5 Then sSql = sSql & "   And PED.SGI_STATUS = 'V' " & vbCrLf
    If stPEDIDOS.Tab = 6 Then sSql = sSql & "   And PED.SGI_STATUS = 'X' " & vbCrLf
    If stPEDIDOS.Tab = 7 Then sSql = sSql & "   And (PED.SGI_STATUS = 'C' or PED.SGI_STATUS = '4') " & vbCrLf
    
    If lngCodVendedor > 0 Then
        sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor
    End If
    
    sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by PED.SGI_CODIGO "
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by CLI.SGI_RAZAOSOC "
    If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order by PED.SGI_DATAPED "
    If cboFiltro.ListIndex = 3 Then sSql = sSql & "Order by ESP.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados
      
    Do While Not BREC.EOF
       strCAMPO = "" & vbTab & _
                  Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                  Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                  BREC!SGI_RAZAOSOC & vbTab & _
                  BREC!SGI_STATUS & vbTab & _
                  BREC!SGI_DESCRICAO & vbTab & _
                  IIf(IsNull(BREC!SGI_CODCOTA) = False, Format(BREC!SGI_CODCOTA, "#/####"), "")
       
       
       If stPEDIDOS.Tab = 0 Then
          flxGRIDPEDIDOS.AddItem strCAMPO
       ElseIf stPEDIDOS.Tab = 1 Then
         flxGRIDBLOQUADOS.AddItem strCAMPO
       ElseIf stPEDIDOS.Tab = 2 Then
         flxReprovados.AddItem strCAMPO
       ElseIf stPEDIDOS.Tab = 3 Then
         
        grdPEDFATURADO.AddItem Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                               Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                               BREC!SGI_RAZAOSOC & vbTab & _
                               BREC!SGI_STATUS & vbTab & _
                               BREC!SGI_DESCRICAO & vbTab & _
                               IIf(IsNull(BREC!SGI_CODCOTA) = False, Format(BREC!SGI_CODCOTA, "#/####"), "") & vbTab & _
                               PegaNF(BREC!SGI_CODIGO)
       ElseIf stPEDIDOS.Tab = 4 Then
       
        grdBLOQALT.AddItem Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                           Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                           BREC!SGI_RAZAOSOC & vbTab & _
                           BREC!SGI_STATUS & vbTab & _
                           BREC!SGI_DESCRICAO & vbTab & _
                           ""
       ElseIf stPEDIDOS.Tab = 5 Then
       
        grdLIBLITO.AddItem Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                           Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                           BREC!SGI_RAZAOSOC & vbTab & _
                           BREC!SGI_STATUS & vbTab & _
                           BREC!SGI_DESCRICAO & vbTab & _
                           ""
       
       ElseIf stPEDIDOS.Tab = 6 Then
       
        grdPARAEST.AddItem Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                           Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                           BREC!SGI_RAZAOSOC & vbTab & _
                           BREC!SGI_STATUS & vbTab & _
                           BREC!SGI_DESCRICAO & vbTab & _
                           ""
       ElseIf stPEDIDOS.Tab = 7 Then
       
        grdLIBPDATAPCOTA.AddItem Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                                 Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                 BREC!SGI_RAZAOSOC & vbTab & _
                                 IIf(BREC!SGI_STATUS = "C", "C", "D") & vbTab & _
                                 BREC!SGI_DESCRICAO & vbTab & _
                                 ""
       
       End If
       
       BREC.MoveNext
    Loop
    BREC.Close

    Call PintaGride

    Exit Sub

Err_Ordem:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Atualiza_Grid()", Me.Name, "Atualiza_Grid()", strCAMARQERRO)


End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADPEDVENDAP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

On Error GoTo Err_txtCampos_Validate
    
    Dim strCampos   As String
    Dim strNOMETAB  As String
    Dim strEMPRESA  As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    If BREC.State = 1 Then BREC.Close
    
    Call AbilitaCampos
    
    
    If stPEDIDOS.Tab = 0 Then ConfGrid
    If stPEDIDOS.Tab = 1 Then ConfGridBloqueados
    If stPEDIDOS.Tab = 2 Then ConfGridReprovados
    If stPEDIDOS.Tab = 3 Then ConfGridFaturado
    If stPEDIDOS.Tab = 4 Then ConfGridBloqAlt
    If stPEDIDOS.Tab = 5 Then ConfGridBloqAltLit
    If stPEDIDOS.Tab = 6 Then ConfGridParaEstoque
    If stPEDIDOS.Tab = 7 Then ConfGridBloqPDataPCota
    
    strOperacao = "P"
    sSql = ""
    
    strEMPRESA = ""
    If intFILIALPED = 0 Then strNOMETAB = "SGI_CADPEDVENDH"
    If intFILIALPED = 1 Then
       strNOMETAB = "SGI_CADPEDVENDH_STEEL"
       strEMPRESA = "_STEEL"
    End If
    
    If cboFiltro.ListIndex = 0 Or cboFiltro.ListIndex = 3 Or cboFiltro.ListIndex = 4 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
    ElseIf cboFiltro.ListIndex = 2 Then
        If IsDate(txtCampos.Text) = False Then
           MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
    End If
    
    sSql = ""
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMETAB & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
    
    If stPEDIDOS.Tab = 0 Then
        If cboFiltro.ListIndex = 4 Then
            sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & " OP" & vbCrLf
        End If
    End If
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    
    
    
    sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
    
    If stPEDIDOS.Tab = 0 Then sSql = sSql & "   And PED.SGI_STATUS = 'L' " & vbCrLf
    If stPEDIDOS.Tab = 1 Then sSql = sSql & "   And (PED.SGI_STATUS = 'B' or PED.SGI_STATUS = 'N')" & vbCrLf
    If stPEDIDOS.Tab = 2 Then sSql = sSql & "   And PED.SGI_STATUS = 'R' " & vbCrLf
    If stPEDIDOS.Tab = 3 Then sSql = sSql & "   And (PED.SGI_STATUS = 'F' or PED.SGI_STATUS = 'P' or PED.SGI_STATUS = 'M') " & vbCrLf
    If stPEDIDOS.Tab = 4 Then sSql = sSql & "   And PED.SGI_STATUS = 'S' " & vbCrLf
    If stPEDIDOS.Tab = 5 Then sSql = sSql & "   And PED.SGI_STATUS = 'V' " & vbCrLf
    If stPEDIDOS.Tab = 6 Then sSql = sSql & "   And PED.SGI_STATUS = 'X' " & vbCrLf
    If stPEDIDOS.Tab = 7 Then sSql = sSql & "   And (PED.SGI_STATUS = 'C' or PED.SGI_STATUS = '4') " & vbCrLf
    
    If lngCodVendedor > 0 Then
        sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor & vbCrLf
    End If
    
    sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
       sSql = sSql & "     And PED.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
       sSql = sSql & "Order by PED.SGI_CODIGO " & vbCrLf
    ElseIf cboFiltro.ListIndex = 1 Then
       sSql = sSql & "     And CLI.SGI_RAZAOSOC LIKE '" & Trim(txtCampos.Text) & "%'" & vbCrLf
       sSql = sSql & "Order by CLI.SGI_RAZAOSOC " & vbCrLf
    ElseIf cboFiltro.ListIndex = 2 Then
       sSql = sSql & "     And PED.SGI_DATAPED = '" & Format(txtCampos.Text, "MM/DD/YYYY") & "'" & vbCrLf
       sSql = sSql & "Order by PED.SGI_DATAPED " & vbCrLf
    ElseIf cboFiltro.ListIndex = 3 Then
       sSql = sSql & "     And PED.SGI_CODCLI = " & Trim(txtCampos.Text) & vbCrLf
       sSql = sSql & "Order by PED.SGI_CODCLI " & vbCrLf
    End If
    
    If stPEDIDOS.Tab = 0 Then
        If cboFiltro.ListIndex = 4 Then
            sSql = sSql & "     And OP.SGI_FILIAL = PED.SGI_FILIAL" & vbCrLf
            sSql = sSql & "     And OP.SGI_CODPED = PED.SGI_CODIGO" & vbCrLf
            sSql = sSql & "     And OP.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
            sSql = sSql & "Order By PED.SGI_CODIGO"
        End If
    End If
    
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
    If Not BREC.EOF Then
         
           
       Do While Not BREC.EOF
       
          strCampos = "" & vbTab & _
                      Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                      Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                      BREC!SGI_RAZAOSOC & vbTab & _
                      BREC!SGI_STATUS & vbTab & _
                      BREC!SGI_DESCRICAO & vbTab & _
                      IIf(IsNull(BREC!SGI_CODCOTA) = False, Format(BREC!SGI_CODCOTA, "#/####"), "")
       
          If stPEDIDOS.Tab = 0 Then flxGRIDPEDIDOS.AddItem strCampos
          If stPEDIDOS.Tab = 1 Then flxGRIDBLOQUADOS.AddItem strCampos
          If stPEDIDOS.Tab = 2 Then flxReprovados.AddItem strCampos
          
          If stPEDIDOS.Tab = 3 Then
          
             grdPEDFATURADO.AddItem BREC!SGI_CODIGO & vbTab & _
                                    Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                    BREC!SGI_RAZAOSOC & vbTab & _
                                    BREC!SGI_STATUS & vbTab & _
                                    BREC!SGI_DESCRICAO & vbTab & _
                                    IIf(IsNull(BREC!SGI_CODCOTA) = False, BREC!SGI_CODCOTA, "") & vbTab & _
                                    PegaNF(BREC!SGI_CODIGO)
          ElseIf stPEDIDOS.Tab = 4 Then
             grdBLOQALT.AddItem BREC!SGI_CODIGO & vbTab & _
                                Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_STATUS & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                IIf(IsNull(BREC!SGI_CODCOTA) = False, BREC!SGI_CODCOTA, "") & vbTab & _
                                PegaNF(BREC!SGI_CODIGO)
          ElseIf stPEDIDOS.Tab = 5 Then
             grdLIBLITO.AddItem BREC!SGI_CODIGO & vbTab & _
                                Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_STATUS & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                IIf(IsNull(BREC!SGI_CODCOTA) = False, BREC!SGI_CODCOTA, "") & vbTab & _
                                PegaNF(BREC!SGI_CODIGO)
          ElseIf stPEDIDOS.Tab = 6 Then
             grdPARAEST.AddItem BREC!SGI_CODIGO & vbTab & _
                                Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_STATUS & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                IIf(IsNull(BREC!SGI_CODCOTA) = False, BREC!SGI_CODCOTA, "") & vbTab & _
                                PegaNF(BREC!SGI_CODIGO)
          ElseIf stPEDIDOS.Tab = 7 Then
                grdLIBPDATAPCOTA.AddItem Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                                         Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                         BREC!SGI_RAZAOSOC & vbTab & _
                                         IIf(BREC!SGI_STATUS = "C", "C", "D") & vbTab & _
                                         BREC!SGI_DESCRICAO & vbTab & _
                                         ""
          End If
          
          BREC.MoveNext
       Loop
           
       BREC.Close
       flxGRIDPEDIDOS.SetFocus
       Exit Sub
        
    Else
        MsgBox "Este Pedido Não Existe !!! - Favor Verificar em outra aba !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close
    
    Exit Sub
    
Err_txtCampos_Validate:

    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCampos_Validate()", Me.Name, "txtCampos_Validate()", strCAMARQERRO)
    

End Sub

Private Sub ImpPed()

On Error GoTo Err_Imp
    
    If BREC.State = 1 Then BREC.Close
    
    Dim strNOMREL As String
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "       SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ENDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RGCGC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_BAINROM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ESTNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CEPNORM" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_DATAPED" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_ORDCOCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODCOTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_ENDCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_BAICOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CIDCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_ESTCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CEPCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_ENDENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_BAIENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CIDENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_ESTENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CEPENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_EMAIL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_DTENTREGA" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_OBS" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CONTATO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODTRANSP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODCONDPGT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODVEND" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_QTDE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_VLUNIT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_PRCIPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_VLTOT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_VLIPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI.SGI_FECHTPFU" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADFECHAM.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH SGI_CADPEDVENDH" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI SGI_CADPEDVENDI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADFECHAM SGI_CADFECHAM" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI.SGI_CODIGO    = " & objCADPEDVENDA.CODPEDIDO & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI.SGI_FILIAL    = " & FILIAL & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDI.SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI.SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDI.SGI_FILIAL    = SGI_CADPEDVENDH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI.SGI_CODIGO    = SGI_CADPEDVENDH.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDI.SGI_FILIAL    = SGI_CADFECHAM.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI.SGI_FECHTPFU  = SGI_CADFECHAM.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_FILIAL    = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODCLI    = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não Há dados Para Imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
   If stPEDIDOS.Tab = 0 Then
        strNOMREL = "RELPVNOVA.RPT"
   Else
        strNOMREL = "RELPVNOVA2.RPT"
   End If
   
   Call objRel.REL(FILIAL, sSql, strCamRelNovo & cCamRelPedidoVendas & strNOMREL, Linha, 1, "Pedido de Vendas", "Pedido de Vendas", False, strACESSO, True)
    
   Exit Sub
   
Err_Imp:

    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : ImpPed()", Me.Name, "ImpPed()", strCAMARQERRO)

End Sub


Private Function PegaCodVendedor(strUSUARIO As String) As Long
    
On Error GoTo Err_PegaCodVendedor
    
    PegaCodVendedor = 0
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       VEN.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO      USU" & vbCrLf
    sSql = sSql & "      ,SGI_CADVENDEDOR  VEN" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       USU.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And USU.SGI_NOME       = '" & Trim(objFuncoes.Crypt(strUSUARIO)) & "'" & vbCrLf
    sSql = sSql & "   And VEN.SGI_FILIAL     = USU.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And VEN.SGI_CODUSUARIO = USU.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaCodVendedor = BREC!SGI_CODIGO
    BREC.Close
    
    Exit Function
    
Err_PegaCodVendedor:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaCodVendedor()", Me.Name, "PegaCodVendedor()", strCAMARQERRO)
    
End Function


Private Sub AtivaDesativaBotoes()

On Error GoTo Err_AtivaDesativaBotoes
    
    If lngCodUsuaro = 0 Then Exit Sub
    
    Command1.Visible = True
    Command4.Visible = False
    
    cmdLibera.Visible = PermiteLibComercial
    Command2.Visible = PermiteLibFinanceiro
    cmdDeslib.Visible = PermiteBloqSN
    Command1.Visible = PermiteReprova
    Command4.Visible = PermiteLiqPedido
    Command5.Visible = PermiteLibPedBloq
    Command6.Visible = PermiteLibPedFotolito
    cmdLIBPDATAPCOTA.Visible = PermiteLibPDataPCota

    Exit Sub
    
Err_AtivaDesativaBotoes:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : AtivaDesativaBotoes()", Me.Name, "AtivaDesativaBotoes()", strCAMARQERRO)
    
End Sub

Private Sub ConfComboFiltro()

On Error GoTo Err_ConfComboFiltro
   
   cboFiltro.Clear
   
   cboFiltro.AddItem "Nº Pedido"
   cboFiltro.AddItem "Cliente"
   cboFiltro.AddItem "Data"
   cboFiltro.AddItem "Cod.Cliente"

   If stPEDIDOS.Tab = 1 Then
        cboFiltro.AddItem "N - Lib/Comercial"
        cboFiltro.AddItem "B = Bloqueado"
   ElseIf stPEDIDOS.Tab = 3 Then
        cboFiltro.AddItem "F - Faturado Total"
        cboFiltro.AddItem "P - Faturado Parcial"
   ElseIf stPEDIDOS.Tab = 0 Then
        cboFiltro.AddItem "Cod.OP"
   End If
   
   
   
   cboFiltro.ListIndex = 0

    Exit Sub
    
Err_ConfComboFiltro:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : ConfComboFiltro()", Me.Name, "ConfComboFiltro()", strCAMARQERRO)

End Sub

Private Sub PegaOPS(strCODPED As String)

On Error GoTo Err_PegaOPS
    
    If Len(Trim(strCODPED)) = 0 Then Exit Sub
    
    If BREC10.State = 1 Then BREC.Close
    
    objCADPEDVENDA.OPS = Empty
    
    Dim lngQTDREGS  As Long
    Dim strFILIAL   As String
    
    strFILIAL = ""
    If intFILIALPED = 1 Then strFILIAL = "_STEEL"
    
    sSql = "Select * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_ORDEMPROD" & strFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And  SGI_CODPED = " & strCODPED & vbCrLf
    sSql = sSql & "   And  (SGI_STATUS = 0 Or SGI_STATUS = 1)"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    
    lngQTDREGS = 0
    
    If Not BREC10.EOF() Then
        Do While Not BREC10.EOF()
            lngQTDREGS = (lngQTDREGS + 1)
            ReDim Preserve arrOPS(1 To lngQTDREGS) As String
            arrOPS(lngQTDREGS) = BREC10!SGI_CODIGO
        
            BREC10.MoveNext
        Loop
        objCADPEDVENDA.OPS = arrOPS
    End If
    BREC10.Close
    
    Exit Sub
    
Err_PegaOPS:
    If BREC10.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaOPS()", Me.Name, "PegaOPS()", strCAMARQERRO)
    
End Sub

Private Sub DestroiObjeto()
    Set objFuncoes = Nothing
    Set objCADPEDVENDA = Nothing
    Set objRel = Nothing
End Sub

Private Sub PintaGride()
    Dim I As Long
    With grdPEDFATURADO
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_SonFat_Situacao) = "F" Then .Cell(flexcpBackColor, I, conCOL_SonFat_Codigo, I, conCOL_SonFat_Status) = &HC000&
            If .Cell(flexcpText, I, conCOL_SonFat_Situacao) = "P" Then .Cell(flexcpBackColor, I, conCOL_SonFat_Codigo, I, conCOL_SonFat_Status) = &HC0C0&
            If .Cell(flexcpText, I, conCOL_SonFat_Situacao) = "M" Then .Cell(flexcpBackColor, I, conCOL_SonFat_Codigo, I, conCOL_SonFat_Status) = &HC0C000
        Next I
    End With
End Sub


Private Sub ImpPedSteel()

On Error GoTo Err_ImpSteel
    
    Dim strNOMREL As String
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "       SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ENDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RGCGC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_BAINROM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ESTNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CEPNORM" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_DATAPED" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ORDCOCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODCOTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ENDCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_BAICOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CIDCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ESTCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CEPCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ENDENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_BAIENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CIDENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ESTENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CEPENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_EMAIL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_DTENTREGA" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_OBS" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CONTATO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODTRANSP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODCONDPGT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODVEND" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_QTDE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_VLUNIT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_PRCIPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_VLTOT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_VLIPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_FECHTPFU" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechSoldaAgrafado" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernCorpo" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernTampa" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernFundo" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernArgola" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechTampaFuro" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADFECHAM.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL SGI_CADPEDVENDH_STEEL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL SGI_CADPEDVENDI_STEEL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADFECHAM SGI_CADFECHAM" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI_STEEL.SGI_CODIGO    = " & objCADPEDVENDA.CODPEDIDO & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_FILIAL    = " & FILIAL & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_FILIAL    = SGI_CADPEDVENDH_STEEL.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_CODIGO    = SGI_CADPEDVENDH_STEEL.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_FILIAL    = SGI_CADFECHAM.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_FECHTPFU  = SGI_CADFECHAM.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH_STEEL.SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH_STEEL.SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    
    ''Call Teste
    
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não Há dados Para Imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
   If stPEDIDOS.Tab = 0 Then
        strNOMREL = "RELPVSTEEL.RPT"
   Else
        strNOMREL = "RELPVSTEEL2.RPT"
   End If
   
   Call objRel.REL(FILIAL, sSql, strCamRelNovo & cCamRelPedidoVendas & strNOMREL, Linha, 1, "Pedido de Vendas", "Pedido de Vendas", False, strACESSO, True)
    
   Exit Sub
   
Err_ImpSteel:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : ImpPedSteel()", Me.Name, "ImpPedSteel()", strCAMARQERRO)

End Sub


Private Sub ConfGridBloqAlt()
        
    With grdBLOQALT
    
       .Cols = conColumnsIn_SonBloq
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonBloq_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonBloq_Codigo) = ""
       .ColDataType(conCOL_SonBloq_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloq_Data) = ""
       .ColDataType(conCOL_SonBloq_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonBloq_Cliente) = ""
       .ColDataType(conCOL_SonBloq_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloq_Situacao) = ""
       .ColDataType(conCOL_SonBloq_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloq_Tipo) = ""
       .ColDataType(conCOL_SonBloq_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloq_Status) = ""
       .ColDataType(conCOL_SonBloq_Status) = flexDTString
       
       .ColWidth(conCOL_SonBloq_Codigo) = 1300
       .ColWidth(conCOL_SonBloq_Data) = 1000
       .ColWidth(conCOL_SonBloq_Cliente) = 5500
       .ColWidth(conCOL_SonBloq_Situacao) = 250
       .ColWidth(conCOL_SonBloq_Tipo) = 2000
       .ColWidth(conCOL_SonBloq_Status) = 0
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub

Private Sub ConfGridBloqAltLit()
        
    With grdLIBLITO
    
       .Cols = conColumnsIn_SonBloqLit
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonBloqLit_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonBloqLit_Codigo) = ""
       .ColDataType(conCOL_SonBloqLit_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloqLit_Data) = ""
       .ColDataType(conCOL_SonBloqLit_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonBloqLit_Cliente) = ""
       .ColDataType(conCOL_SonBloqLit_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloqLit_Situacao) = ""
       .ColDataType(conCOL_SonBloqLit_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloqLit_Tipo) = ""
       .ColDataType(conCOL_SonBloqLit_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloqLit_Status) = ""
       .ColDataType(conCOL_SonBloqLit_Status) = flexDTString
       
       .ColWidth(conCOL_SonBloqLit_Codigo) = 1300
       .ColWidth(conCOL_SonBloqLit_Data) = 1000
       .ColWidth(conCOL_SonBloqLit_Cliente) = 5500
       .ColWidth(conCOL_SonBloqLit_Situacao) = 250
       .ColWidth(conCOL_SonBloqLit_Tipo) = 2000
       .ColWidth(conCOL_SonBloqLit_Status) = 0
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub


Private Sub ConfGridParaEstoque()
        
    With grdPARAEST
    
       .Cols = conColumnsIn_SonParaEst
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonParaEst_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonParaEst_Codigo) = ""
       .ColDataType(conCOL_SonParaEst_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonParaEst_Data) = ""
       .ColDataType(conCOL_SonParaEst_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonParaEst_Cliente) = ""
       .ColDataType(conCOL_SonParaEst_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonParaEst_Situacao) = ""
       .ColDataType(conCOL_SonParaEst_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonParaEst_Tipo) = ""
       .ColDataType(conCOL_SonParaEst_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonParaEst_Status) = ""
       .ColDataType(conCOL_SonParaEst_Status) = flexDTString
       
       .ColWidth(conCOL_SonParaEst_Codigo) = 1300
       .ColWidth(conCOL_SonParaEst_Data) = 1000
       .ColWidth(conCOL_SonParaEst_Cliente) = 5500
       .ColWidth(conCOL_SonParaEst_Situacao) = 250
       .ColWidth(conCOL_SonParaEst_Tipo) = 2000
       .ColWidth(conCOL_SonParaEst_Status) = 0
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub


Private Sub Refaz_Indice_Pai()

    Dim strNOMTABELA As String
    Dim arrOPES      As Variant
    Dim lngQTDREGS   As Long
    Dim I            As Long

    strNOMTABELA = ""
    If intFILIALPED = 1 Then strNOMTABELA = "_STEEL"

    '' Contando o Total de Registros
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Count(*) As QtdeRegs " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_PROGENTRPROD" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then lngQTDREGS = BREC!QtdeRegs
    BREC.Close

    If lngQTDREGS > 0 Then
        ReDim arrOPES(1 To lngQTDREGS, 1 To 7) As String
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       *" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_PROGENTRPROD" & strNOMTABELA & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "Order By SGI_CODPED" & vbCrLf
        sSql = sSql & "        ,SGI_INDICE"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        lngQTDREGS = 0
        Do While Not BREC.EOF()
            lngQTDREGS = (lngQTDREGS + 1)
            
            arrOPES(lngQTDREGS, 1) = Trim(Str(BREC!SGI_CODPED))
            arrOPES(lngQTDREGS, 2) = Trim(Str(BREC!SGI_IDPRODUTO))
            arrOPES(lngQTDREGS, 3) = Trim(Str(BREC!SGI_INDICE))
            arrOPES(lngQTDREGS, 4) = Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY")
            arrOPES(lngQTDREGS, 5) = lngQTDREGS
            arrOPES(lngQTDREGS, 7) = BREC!SGI_QTDE
            
            '' OP
            sSql = ""
            
            sSql = "Select" & vbCrLf
            sSql = sSql & "       *" & vbCrLf
            sSql = sSql & "  From" & vbCrLf
            sSql = sSql & "       SGI_ORDEMPROD" & strNOMTABELA & vbCrLf
            sSql = sSql & " Where" & vbCrLf
            sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODPED    = " & BREC!SGI_CODPED & vbCrLf
            sSql = sSql & "   And SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And SGI_QTDEPED   = " & BREC!SGI_QTDE
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then arrOPES(lngQTDREGS, 6) = BREC2!SGI_CODIGO
            BREC2.Close
            
            BREC.MoveNext
        Loop
        BREC.Close
        
        '' Realizando Transação
        '' Inicia transação
        adoBanco_Dados.BeginTrans
        BGRV.ActiveConnection = adoBanco_Dados
        
        For I = 1 To UBound(arrOPES)
            
            sSql = ""
            
            '' SGI_PROGENTRPROD
            sSql = "Update SGI_PROGENTRPROD" & strNOMTABELA & " Set " & vbCrLf
            sSql = sSql & "                           SGI_IDINTERNO = " & arrOPES(I, 5) & vbCrLf
            sSql = sSql & "                     Where " & vbCrLf
            sSql = sSql & "                           SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "                       And SGI_CODPED      = " & arrOPES(I, 1) & vbCrLf
            sSql = sSql & "                       And SGI_IDPRODUTO   = " & arrOPES(I, 2) & vbCrLf
            sSql = sSql & "                       And SGI_INDICE      = " & arrOPES(I, 3) & vbCrLf
            sSql = sSql & "                       And SGI_DATENTREGA  = '" & Format(CDate(arrOPES(I, 4)), "MM/DD/YYYY") & "'"
            
            BGRV.CommandText = sSql
            BGRV.Execute
        
            sSql = ""
            
            sSql = "Update SGI_ORDEMPROD" & strNOMTABELA & " Set" & vbCrLf
            sSql = sSql & "                     SGI_IDPAI = " & arrOPES(I, 5) & vbCrLf
            sSql = sSql & "               Where " & vbCrLf
            sSql = sSql & "                    SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "                And SGI_CODPED     = " & arrOPES(I, 1) & vbCrLf
            sSql = sSql & "                And SGI_IDPRODUTO  = " & arrOPES(I, 2) & vbCrLf
            sSql = sSql & "                And SGI_QTDEPED    = " & arrOPES(I, 7)
        
            BGRV.CommandText = sSql
            BGRV.Execute
        
        Next I
        
        sSql = ""
        
        sSql = "Delete From SGI_NUMERO " & vbCrLf
        sSql = sSql & "                 Where " & vbCrLf
        sSql = sSql & "                       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                   And SGI_MODULO = 'frmCADPEDVENDA_PROGENTR" & strNOMTABELA & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        sSql = ""
        
        sSql = "Insert Into SGI_NUMERO (" & vbCrLf
        sSql = sSql & "                        SGI_FILIAL" & vbCrLf
        sSql = sSql & "                       ,SGI_MODULO" & vbCrLf
        sSql = sSql & "                       ,SGI_NUMERO" & vbCrLf
        sSql = sSql & "              ) Values (" & vbCrLf
        sSql = sSql & "                        " & FILIAL & vbCrLf
        sSql = sSql & "                       ,'frmCADPEDVENDA_PROGENTR" & strNOMTABELA & "'" & vbCrLf
        sSql = sSql & "                       ," & UBound(arrOPES) & vbCrLf
        sSql = sSql & "                       )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        adoBanco_Dados.CommitTrans
        
        MsgBox "Dados Processados com Exito !!!", vbOKOnly + vbExclamation, "Aviso"
        
    Else
        MsgBox "Não há dados para processar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If

End Sub

Private Function PermiteBloqSN() As Boolean

    PermiteBloqSN = False
    
    If lngCodUsuaro = 0 Then
       PermiteBloqSN = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_PERMBLOQPED" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuaro

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If BREC!SGI_PERMBLOQPED = 1 Then
           PermiteBloqSN = True
        ElseIf BREC!SGI_PERMBLOQPED = 0 Then
           PermiteBloqSN = False
        End If
    End If
    BREC.Close

End Function
 
Private Function PermiteLibFinanceiro() As Boolean

    PermiteLibFinanceiro = False
    
    If lngCodUsuaro = 0 Then
       PermiteLibFinanceiro = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBFINSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO   = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_LIBFINSN = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibFinanceiro = True
    BREC.Close

End Function
 
Private Function PermiteLibComercial() As Boolean

    PermiteLibComercial = False
    
    If lngCodUsuaro = 0 Then
       PermiteLibComercial = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBCOMSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO   = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_LIBCOMSN = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibComercial = True
    BREC.Close

End Function
 
 
Private Function PermiteReprova() As Boolean

    PermiteReprova = False
    
    If lngCodUsuaro = 0 Then
       PermiteReprova = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBCOMSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO   = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_REPEDSN  = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteReprova = True
    BREC.Close

End Function
 
 
Private Function PermiteLiqPedido() As Boolean

    PermiteLiqPedido = False
    
    If lngCodUsuaro = 0 Then
       PermiteLiqPedido = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBCOMSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO   = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_LIQPEDSN = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLiqPedido = True
    BREC.Close

End Function
 
 
Private Function PermiteLibPedBloq() As Boolean

    PermiteLibPedBloq = False
    
    If lngCodUsuaro = 0 Then
       PermiteLibPedBloq = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBCOMSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO       = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_LIBPEDBLOQSN = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibPedBloq = True
    BREC.Close

End Function
 
Private Function PermiteLibPedFotolito() As Boolean

    PermiteLibPedFotolito = False
    
    If lngCodUsuaro = 0 Then
       PermiteLibPedFotolito = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBCOMSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO       = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_LIBPEDFOTSN  = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibPedFotolito = True
    BREC.Close

End Function

Private Sub ConfGridBloqPDataPCota()
        
    With grdLIBPDATAPCOTA
    
       .Cols = conColumnsIn_SonBloqPDPC
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonBloqPDPC_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonBloqPDPC_Codigo) = ""
       .ColDataType(conCOL_SonBloqPDPC_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloqPDPC_Data) = ""
       .ColDataType(conCOL_SonBloqPDPC_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonBloqPDPC_Cliente) = ""
       .ColDataType(conCOL_SonBloqPDPC_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloqPDPC_Situacao) = ""
       .ColDataType(conCOL_SonBloqPDPC_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloqPDPC_Tipo) = ""
       .ColDataType(conCOL_SonBloqPDPC_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonBloqPDPC_Status) = ""
       .ColDataType(conCOL_SonBloqPDPC_Status) = flexDTString
       
       .ColWidth(conCOL_SonBloqPDPC_Codigo) = 1300
       .ColWidth(conCOL_SonBloqPDPC_Data) = 1000
       .ColWidth(conCOL_SonBloqPDPC_Cliente) = 5500
       .ColWidth(conCOL_SonBloqPDPC_Situacao) = 250
       .ColWidth(conCOL_SonBloqPDPC_Tipo) = 2000
       .ColWidth(conCOL_SonBloqPDPC_Status) = 0
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub


Private Function PermiteLibPDataPCota() As Boolean

    PermiteLibPDataPCota = False
    
    If lngCodUsuaro = 0 Then
       PermiteLibPDataPCota = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBPDATAPCOTA" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO         = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_LIBPDATAPCOTA  = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibPDataPCota = True
    BREC.Close

End Function


Private Sub Teste()

    sSql = ""

    sSql = "Select " & vbCrLf
    
    sSql = sSql & "       SGI_CADPEDVENDI_STEEL.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_QTDE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_VLUNIT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_PRCIPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_VLTOT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI_STEEL.SGI_VLIPI" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechSoldaAgrafado" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernCorpo" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernTampa" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernFundo" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernArgola" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechTampaFuro" & vbCrLf

    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_FILIAL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_DATAPED" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ORDCOCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODCOTA" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ENDCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_BAICOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CIDCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ESTCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CEPCOBR" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ENDENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_BAIENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CIDENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_ESTENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CEPENTRE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_EMAIL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_DTENTREGA" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_OBS" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CONTATO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODTRANSP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODCONDPGT" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL.SGI_CODVEND" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ENDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RGCGC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_BAINROM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ESTNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CEPNORM" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    
    sSql = sSql & "       SGI_CADPEDVENDI_STEEL SGI_CADPEDVENDI_STEEL" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL SGI_CADPEDVENDH_STEEL" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI_STEEL.SGI_CODIGO    = " & objCADPEDVENDA.CODPEDIDO & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_FILIAL    = " & FILIAL & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_FILIAL    = SGI_CADPEDVENDH_STEEL.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDI_STEEL.SGI_CODIGO    = SGI_CADPEDVENDH_STEEL.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADCLIENTE.SGI_FILIAL = SGI_CADPEDVENDH_STEEL.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADCLIENTE.SGI_CODIGO = SGI_CADPEDVENDH_STEEL.SGI_CODCLI " & vbCrLf

End Sub
