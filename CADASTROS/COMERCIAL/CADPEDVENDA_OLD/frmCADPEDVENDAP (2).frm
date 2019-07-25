VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADPEDVENDAP 
   Caption         =   "Cadastro de Pedidos de Vendas"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   14820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   6495
      Left            =   0
      TabIndex        =   17
      Top             =   3120
      Width           =   14775
      Begin TabDlg.SSTab stPEDIDOS 
         Height          =   6255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   11033
         _Version        =   393216
         Style           =   1
         Tabs            =   9
         TabsPerRow      =   9
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
         Tab(0).Control(1)=   "lblRegsRelac(7)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(7)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "grdPEDIDOS"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Aguardando Liberação"
         TabPicture(1)   =   "frmCADPEDVENDAP.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdGRIDBLOQUADOS"
         Tab(1).Control(1)=   "Frame4"
         Tab(1).Control(2)=   "Label1(6)"
         Tab(1).Control(3)=   "lblRegsRelac(6)"
         Tab(1).Control(4)=   "Label3(2)"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Reprovados"
         TabPicture(2)   =   "frmCADPEDVENDAP.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label3(3)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "lblRegsRelac(5)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label1(5)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "grdReprovados"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Faturados"
         TabPicture(3)   =   "frmCADPEDVENDAP.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "grdPEDFATURADO"
         Tab(3).Control(1)=   "lblRegsRelac(1)"
         Tab(3).Control(2)=   "Label1(1)"
         Tab(3).Control(3)=   "Label3(0)"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Bloqueados"
         TabPicture(4)   =   "frmCADPEDVENDAP.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Command5"
         Tab(4).Control(1)=   "grdBLOQALT"
         Tab(4).Control(2)=   "Label1(4)"
         Tab(4).Control(3)=   "lblRegsRelac(4)"
         Tab(4).ControlCount=   4
         TabCaption(5)   =   "Aguardando Liberação de Artes"
         TabPicture(5)   =   "frmCADPEDVENDAP.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label1(0)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "lblRegsRelac(0)"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "grdLIBLITO"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).Control(3)=   "Command6"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).ControlCount=   4
         TabCaption(6)   =   "Para Estoque"
         TabPicture(6)   =   "frmCADPEDVENDAP.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "grdPARAEST"
         Tab(6).Control(1)=   "Command7"
         Tab(6).ControlCount=   2
         TabCaption(7)   =   "Bloqueado - P.Data/P.Cota"
         TabPicture(7)   =   "frmCADPEDVENDAP.frx":00C4
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Label1(3)"
         Tab(7).Control(1)=   "lblRegsRelac(3)"
         Tab(7).Control(2)=   "grdLIBPDATAPCOTA"
         Tab(7).Control(3)=   "cmdLIBPDATAPCOTA"
         Tab(7).ControlCount=   4
         TabCaption(8)   =   "Geral"
         TabPicture(8)   =   "frmCADPEDVENDAP.frx":00E0
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "lblRegsRelac(2)"
         Tab(8).Control(1)=   "Label1(2)"
         Tab(8).Control(2)=   "grdGERAL"
         Tab(8).ControlCount=   3
         Begin VSFlex8LCtl.VSFlexGrid grdGERAL 
            Height          =   5295
            Left            =   -74880
            TabIndex        =   54
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   9340
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
         Begin VSFlex8LCtl.VSFlexGrid grdReprovados 
            Height          =   5415
            Left            =   -74880
            TabIndex        =   51
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   9551
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
         Begin VSFlex8LCtl.VSFlexGrid grdPEDIDOS 
            Height          =   5415
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   9551
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
         Begin VSFlex8LCtl.VSFlexGrid grdGRIDBLOQUADOS 
            Height          =   5415
            Left            =   -74880
            TabIndex        =   49
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   9551
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Picture         =   "frmCADPEDVENDAP.frx":00FC
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Liberação das alterações"
            Top             =   5520
            Width           =   1695
         End
         Begin VSFlex8LCtl.VSFlexGrid grdLIBPDATAPCOTA 
            Height          =   5055
            Left            =   -74880
            TabIndex        =   39
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   8916
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Picture         =   "frmCADPEDVENDAP.frx":0529
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Liberação das Alterações do Fotolito"
            Top             =   5520
            Width           =   1695
         End
         Begin VSFlex8LCtl.VSFlexGrid grdPARAEST 
            Height          =   5055
            Left            =   -74880
            TabIndex        =   37
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   8916
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Picture         =   "frmCADPEDVENDAP.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Liberação das Alterações do Fotolito"
            Top             =   5520
            Width           =   1335
         End
         Begin VSFlex8LCtl.VSFlexGrid grdLIBLITO 
            Height          =   5055
            Left            =   -74880
            TabIndex        =   35
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   8916
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Picture         =   "frmCADPEDVENDAP.frx":0D83
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Liberação das alterações"
            Top             =   5520
            Width           =   1335
         End
         Begin VSFlex8LCtl.VSFlexGrid grdBLOQALT 
            Height          =   5055
            Left            =   -74880
            TabIndex        =   33
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   8916
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Height          =   5415
            Left            =   -74880
            TabIndex        =   32
            Top             =   360
            Width           =   14295
            _cx             =   25215
            _cy             =   9551
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Left            =   -67800
            TabIndex        =   23
            Top             =   5880
            Width           =   2535
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
               TabIndex        =   25
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
               TabIndex        =   24
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Itens Relacionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   7
            Left            =   10200
            TabIndex        =   74
            Top             =   5880
            Width           =   2025
         End
         Begin VB.Label lblRegsRelac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblRegsRelac"
            Height          =   255
            Index           =   7
            Left            =   12600
            TabIndex        =   73
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Itens Relacionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   6
            Left            =   -64800
            TabIndex        =   72
            Top             =   5880
            Width           =   2025
         End
         Begin VB.Label lblRegsRelac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblRegsRelac"
            Height          =   255
            Index           =   6
            Left            =   -62400
            TabIndex        =   71
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Itens Relacionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   5
            Left            =   -64800
            TabIndex        =   70
            Top             =   5880
            Width           =   2025
         End
         Begin VB.Label lblRegsRelac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblRegsRelac"
            Height          =   255
            Index           =   5
            Left            =   -62400
            TabIndex        =   69
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Itens Relacionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   4
            Left            =   -64800
            TabIndex        =   68
            Top             =   5760
            Width           =   2025
         End
         Begin VB.Label lblRegsRelac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblRegsRelac"
            Height          =   255
            Index           =   4
            Left            =   -62400
            TabIndex        =   67
            Top             =   5760
            Width           =   1815
         End
         Begin VB.Label lblRegsRelac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblRegsRelac"
            Height          =   255
            Index           =   3
            Left            =   -62400
            TabIndex        =   66
            Top             =   5760
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Itens Relacionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   3
            Left            =   -64800
            TabIndex        =   65
            Top             =   5760
            Width           =   2025
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Itens Relacionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   2
            Left            =   -64800
            TabIndex        =   64
            Top             =   5760
            Width           =   2025
         End
         Begin VB.Label lblRegsRelac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblRegsRelac"
            Height          =   255
            Index           =   2
            Left            =   -62400
            TabIndex        =   63
            Top             =   5760
            Width           =   1815
         End
         Begin VB.Label lblRegsRelac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblRegsRelac"
            Height          =   255
            Index           =   1
            Left            =   -62400
            TabIndex        =   62
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Itens Relacionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   1
            Left            =   -64800
            TabIndex        =   61
            Top             =   5880
            Width           =   2025
         End
         Begin VB.Label lblRegsRelac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblRegsRelac"
            Height          =   255
            Index           =   0
            Left            =   -62400
            TabIndex        =   60
            Top             =   5760
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Itens Relacionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   -64800
            TabIndex        =   59
            Top             =   5760
            Width           =   2025
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
            TabIndex        =   29
            Top             =   5880
            Width           =   6615
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
            TabIndex        =   28
            Top             =   5880
            Width           =   7695
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
            TabIndex        =   27
            Top             =   5880
            Width           =   6855
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
            TabIndex        =   26
            Top             =   5880
            Width           =   6015
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   14775
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
         Height          =   735
         Left            =   10560
         Picture         =   "frmCADPEDVENDAP.frx":11B0
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Height          =   735
         Left            =   12360
         Picture         =   "frmCADPEDVENDAP.frx":15AD
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Height          =   735
         Left            =   5520
         Picture         =   "frmCADPEDVENDAP.frx":16AF
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Liberação Financeira"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   11640
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
         Height          =   735
         Left            =   9720
         Picture         =   "frmCADPEDVENDAP.frx":1ADC
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Height          =   735
         Left            =   8400
         Picture         =   "frmCADPEDVENDAP.frx":1ED9
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Height          =   735
         Left            =   6960
         Picture         =   "frmCADPEDVENDAP.frx":22FF
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Liberação Comercial"
         Top             =   120
         Width           =   1455
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
         Left            =   13920
         Picture         =   "frmCADPEDVENDAP.frx":272C
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Height          =   735
         Left            =   13200
         Picture         =   "frmCADPEDVENDAP.frx":282E
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Height          =   735
         Left            =   3000
         Picture         =   "frmCADPEDVENDAP.frx":2D60
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Exclui Registro"
         Top             =   120
         Width           =   975
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
         Height          =   735
         Left            =   2040
         Picture         =   "frmCADPEDVENDAP.frx":2E62
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Altera Registro"
         Top             =   120
         Width           =   975
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
         Height          =   735
         Left            =   1080
         Picture         =   "frmCADPEDVENDAP.frx":2F64
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Inclui um novo registro"
         Top             =   120
         Width           =   975
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
         Height          =   735
         Left            =   120
         Picture         =   "frmCADPEDVENDAP.frx":3496
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Volta ao Menu Principal"
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
         Height          =   735
         Left            =   3960
         Picture         =   "frmCADPEDVENDAP.frx":39C8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Tag             =   "2"
      Top             =   0
      Width           =   14775
      Begin VB.Frame fraOrdem 
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
         Height          =   1215
         Left            =   6600
         TabIndex        =   57
         Top             =   120
         Width           =   3735
         Begin VB.ListBox lstOrdem 
            Appearance      =   0  'Flat
            Height          =   810
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame fraStatus 
         Caption         =   "[ Status ]"
         DragMode        =   1  'Automatic
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
         Height          =   1935
         Left            =   10440
         TabIndex        =   55
         Tag             =   "PED.SGI_STATUS"
         Top             =   120
         Width           =   4215
         Begin VB.ListBox lstStatus 
            Appearance      =   0  'Flat
            Height          =   1605
            ItemData        =   "frmCADPEDVENDAP.frx":3ACA
            Left            =   120
            List            =   "frmCADPEDVENDAP.frx":3ACC
            Style           =   1  'Checkbox
            TabIndex        =   56
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   9960
         Picture         =   "frmCADPEDVENDAP.frx":3ACE
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   9960
         Picture         =   "frmCADPEDVENDAP.frx":3BD0
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtCODOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   3
         Text            =   "txtCODOP"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtNomVend 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "txtNomVend"
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data do Pedido"
         DragMode        =   1  'Automatic
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
         TabIndex        =   45
         Tag             =   "PED.SGI_DATAPED"
         Top             =   600
         Width           =   3135
         Begin MSMask.MaskEdBox mskDataI 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataF 
            Height          =   255
            Left            =   1920
            TabIndex        =   5
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblCampo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   1560
            TabIndex        =   46
            Top             =   240
            Width           =   120
         End
      End
      Begin VB.TextBox txtCODVENDEDOR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Text            =   "txtCODVENDEDOR"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtNomClie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "txtNomClie"
         Top             =   1800
         Width           =   4815
      End
      Begin VB.TextBox txtCODCLIE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Text            =   "txtCODCLIE"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtCODPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Text            =   "txtCODPED"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCampo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Código da OP"
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
         Left            =   3480
         TabIndex        =   48
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblCampo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nome do Vendedor"
         DragMode        =   1  'Automatic
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
         Left            =   3360
         TabIndex        =   47
         Tag             =   "VEN.SGI_NOMVEND"
         Top             =   1440
         Width           =   1650
      End
      Begin VB.Label lblCampo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cód Vendedor"
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
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label lblCampo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nome do Cliente"
         DragMode        =   1  'Automatic
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
         Left            =   3360
         TabIndex        =   43
         Tag             =   "CLI.SGI_RAZAOSOC"
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label lblCampo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cód Cliente"
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
         TabIndex        =   42
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label lblCampo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "N° Pedido"
         DragMode        =   1  'Automatic
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
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Tag             =   "PED.SGI_CODIGO"
         Top             =   240
         Width           =   870
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

Dim boolComAcao         As Boolean
Dim boolEVendedor       As Boolean
Dim lngCodVendedor      As Long
Dim iCodigo             As Long
Dim strOperacao         As String
Dim arrOPS()            As String
Dim cTipOper            As String
Dim arrORDEM()          As String

Dim objFuncoes          As Object
Dim objCADPEDVENDA      As Object
Dim objRel              As Object
Dim objPESQPADRAO       As Object
Dim boolTelaAberta      As Boolean

Const conCOL_SonPed_Codigo                              As Integer = 0
Const conCOL_SonPed_Data                                As Integer = 1
Const conCOL_SonPed_Cliente                             As Integer = 2
Const conCOL_SonPed_Situacao                            As Integer = 3
Const conCOL_SonPed_Tipo                                As Integer = 4
Const conCOL_SonPed_Vendedor                            As Integer = 5
Const conCOL_SonPed_FormatString                        As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonPed                               As Integer = 6

Const conCOL_SonAgLib_Codigo                            As Integer = 0
Const conCOL_SonAgLib_Data                              As Integer = 1
Const conCOL_SonAgLib_Cliente                           As Integer = 2
Const conCOL_SonAgLib_Situacao                          As Integer = 3
Const conCOL_SonAgLib_Tipo                              As Integer = 4
Const conCOL_SonAgLib_Vendedor                          As Integer = 5
Const conCOL_SonAgLib_FormatString                      As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonAgLib                             As Integer = 6

Const conCOL_SonFat_Codigo                              As Integer = 0
Const conCOL_SonFat_Data                                As Integer = 1
Const conCOL_SonFat_Cliente                             As Integer = 2
Const conCOL_SonFat_Situacao                            As Integer = 3
Const conCOL_SonFat_Tipo                                As Integer = 4
Const conCOL_SonFat_Vendedor                            As Integer = 5
Const conCOL_SonFat_FormatString                        As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonFat                               As Integer = 6

Const conCOL_SonRep_Codigo                              As Integer = 0
Const conCOL_SonRep_Data                                As Integer = 1
Const conCOL_SonRep_Cliente                             As Integer = 2
Const conCOL_SonRep_Situacao                            As Integer = 3
Const conCOL_SonRep_Tipo                                As Integer = 4
Const conCOL_SonRep_Vendedor                            As Integer = 5
Const conCOL_SonRep_FormatString                        As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonRep                               As Integer = 6

Const conCOL_SonBloq_Codigo                             As Integer = 0
Const conCOL_SonBloq_Data                               As Integer = 1
Const conCOL_SonBloq_Cliente                            As Integer = 2
Const conCOL_SonBloq_Situacao                           As Integer = 3
Const conCOL_SonBloq_Tipo                               As Integer = 4
Const conCOL_SonBloq_Vendedor                           As Integer = 5
Const conCOL_SonBloq_FormatString                       As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonBloq                              As Integer = 6

Const conCOL_SonBloqLit_Codigo                          As Integer = 0
Const conCOL_SonBloqLit_Data                            As Integer = 1
Const conCOL_SonBloqLit_Cliente                         As Integer = 2
Const conCOL_SonBloqLit_Situacao                        As Integer = 3
Const conCOL_SonBloqLit_Tipo                            As Integer = 4
Const conCOL_SonBloqLit_Vendedor                        As Integer = 5
Const conCOL_SonBloqLit_FormatString                    As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonBloqLit                           As Integer = 6

Const conCOL_SonParaEst_Codigo                          As Integer = 0
Const conCOL_SonParaEst_Data                            As Integer = 1
Const conCOL_SonParaEst_Cliente                         As Integer = 2
Const conCOL_SonParaEst_Situacao                        As Integer = 3
Const conCOL_SonParaEst_Tipo                            As Integer = 4
Const conCOL_SonParaEst_Vendedor                        As Integer = 5
Const conCOL_SonParaEst_FormatString                    As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonParaEst                           As Integer = 6

Const conCOL_SonBloqPDPC_Codigo                         As Integer = 0
Const conCOL_SonBloqPDPC_Data                           As Integer = 1
Const conCOL_SonBloqPDPC_Cliente                        As Integer = 2
Const conCOL_SonBloqPDPC_Situacao                       As Integer = 3
Const conCOL_SonBloqPDPC_Tipo                           As Integer = 4
Const conCOL_SonBloqPDPC_Vendedor                       As Integer = 5
Const conCOL_SonBloqPDPC_FormatString                   As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonBloqPDPC                          As Integer = 6

Const conCOL_SonGeral_Codigo                            As Integer = 0
Const conCOL_SonGeral_Data                              As Integer = 1
Const conCOL_SonGeral_Cliente                           As Integer = 2
Const conCOL_SonGeral_Situacao                          As Integer = 3
Const conCOL_SonGeral_Tipo                              As Integer = 4
Const conCOL_SonGeral_Vendedor                          As Integer = 5
Const conCOL_SonGeral_FormatString                      As String = "=Cód. Ped|Data|Cliente|S|Tipo Pedido|Vendedor"
Const conColumnsIn_SonGeral                             As Integer = 6


Private Sub cmdAltera_Click()
    
On Error GoTo Err_cmdAltera_Click
    
    Dim strSITUACAO As String
    
    If objFuncoes.ChecaAcesso2("A", strACESSO) = False Then Exit Sub
    If stPEDIDOS.Tab = 0 Then
       MsgBox "Este Pedido já está Liberado não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    If stPEDIDOS.Tab = 1 Then
       If Trim(grdGRIDBLOQUADOS.Cell(flexcpText, grdGRIDBLOQUADOS.Row, conCOL_SonAgLib_Situacao)) = "N" Then
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
    
    If stPEDIDOS.Tab = 5 Then
        MsgBox "Este não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If stPEDIDOS.Tab = 8 Then
       strSITUACAO = Trim(grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao))
       If strSITUACAO = "N" Then
            MsgBox "Este Pedido já está Liberado não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
       ElseIf (strSITUACAO = "F" Or strSITUACAO = "M" Or strSITUACAO = "P") Then
            MsgBox "Este Pedido já está Faturado não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
       ElseIf strSITUACAO = "L" Then
            MsgBox "Este Pedido já está Liberado não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
       ElseIf strSITUACAO = "V" Or _
              strSITUACAO = "R" Then
            MsgBox "Este pedido não pode ser Alterado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
       End If
    End If
    
    If stPEDIDOS.Tab = 1 Or _
       stPEDIDOS.Tab = 4 Or _
       stPEDIDOS.Tab = 7 Or _
       stPEDIDOS.Tab = 8 Then Call Operacao("A")
    Exit Sub
    
Err_cmdAltera_Click:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdAltera_Click()", Me.Name, "cmdAltera_Click()", strCAMARQERRO)
    
    
End Sub

Private Sub cmdCanFiltro_Click()
   
On Error GoTo Err_cmdCanFiltro_Click
   
   strOperacao = ""
   Call AbilitaCampos
   
   Call ConfGridPedidos
   Call ConfGridAgLib
   Call ConfGridReprovados
   Call ConfGridFaturado
   Call ConfGridBloqAlt
   Call ConfGridBloqAltLit
   Call ConfGridParaEstoque
   Call ConfGridBloqPDataPCota
   
   objFuncoes.LimpaCampos Me
   Call OrdemPadrao
   
   Exit Sub
   
Err_cmdCanFiltro_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdCanFiltro_Click()", Me.Name, "cmdCanFiltro_Click()", strCAMARQERRO)
   
End Sub

Private Sub cmdDeslib_Click()
    
On Error GoTo Err_cmdDeslib_Click
    
    Dim strSITUACAO As String
    
    ''If objFuncoes.ChecaAcesso2("B", strAcesso) = False Then Exit Sub
    If stPEDIDOS.Tab = 3 Then Exit Sub
    If VerifNF = False Then Exit Sub
    
    If stPEDIDOS.Tab = 0 Or _
       stPEDIDOS.Tab = 2 Then Call Operacao("D")
    
    If stPEDIDOS.Tab = 1 Then
       If grdGRIDBLOQUADOS.Cell(flexcpText, grdGRIDBLOQUADOS.RowSel, conCOL_SonAgLib_Situacao) = "N" Then
          Call Operacao("D")
       End If
    End If
    
    If stPEDIDOS.Tab = 8 Then
        If Verif_Linha_Sel(grdGERAL, "D") = False Then Exit Sub
        strSITUACAO = Trim(grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao))
        If strSITUACAO = "N" Or _
           strSITUACAO = "R" Or _
           strSITUACAO = "L" Then
           Call Operacao("D")
        ElseIf strSITUACAO = "M" Or _
               strSITUACAO = "F" Or _
               strSITUACAO = "P" Or _
               strSITUACAO = "S" Or _
               strSITUACAO = "V" Then
            MsgBox "ATENÇÃO" & vbCrLf & "Não é possivel realisar esta ação !!!", vbOKOnly + vbExclamation, "Aviso"
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
      
    objCADPEDVENDA.FILIALPED = intFILIALPED
    If objCADPEDVENDA.GRAVASTEEL("E") = False Then Exit Sub
    If objCADPEDVENDA.Atualiza("E", objCADPEDVENDA.CODPEDIDO, FILIAL, "frmCADPEDVENDA", intFILIALPED) = False Then Exit Sub
    
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"

  End If
  
  Call AbilitaCampos
  Call FuncaoAtualiza

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
    
    Dim strSITUACAO As String
    
    ''If objFuncoes.ChecaAcesso2("L", strAcesso) = False Then Exit Sub
    If stPEDIDOS.Tab = 3 Then Exit Sub
    
    If stPEDIDOS.Tab = 8 Then
       
        If Verif_Linha_Sel(grdGERAL, "N") = False Then Exit Sub
        
        strSITUACAO = Trim(grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao))
        If strSITUACAO = "M" Or _
           strSITUACAO = "F" Or _
           strSITUACAO = "P" Or _
           strSITUACAO = "S" Or _
           strSITUACAO = "V" Then
           MsgBox "ATENÇÃO" & vbCrLf & _
                  "Não é possivel realizar esta ação !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    
    If ((grdGRIDBLOQUADOS.Rows - 1) > 0) Or ((grdReprovados.Rows - 1) > 0) Or ((grdGERAL.Rows - 1) > 0) Then
       Call Operacao("LN")
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
    If ConsisteCampos = False Then Exit Sub
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
On Error GoTo Err_Command1_Click
    
    If stPEDIDOS.Tab = 0 Then If Verif_Linha_Sel(grdPEDIDOS, "R") = False Then Exit Sub
    If stPEDIDOS.Tab = 1 Then If Verif_Linha_Sel(grdGRIDBLOQUADOS, "R") = False Then Exit Sub
    If stPEDIDOS.Tab = 8 Then If Verif_Linha_Sel(grdGERAL, "R") = False Then Exit Sub
    
    Dim strSITUACAO As String
    
    ''If objFuncoes.ChecaAcesso2("V", strAcesso) = False Then Exit Sub
    If stPEDIDOS.Tab = 3 Then Exit Sub
    
    If stPEDIDOS.Tab = 8 Then
        strSITUACAO = grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao)
        If strSITUACAO = "R" Then
            Exit Sub
        ElseIf strSITUACAO = "M" Or _
               strSITUACAO = "F" Or _
               strSITUACAO = "P" Or _
               strSITUACAO = "S" Or _
               strSITUACAO = "V" Then
               MsgBox "ATENÇÃO" & vbCrLf & _
                      "Não é possivel realizar esta ação !!!", vbOKOnly + vbExclamation, "Aviso"
               Exit Sub
        End If
    End If
    
    If VerifNF = False Then Exit Sub
    If (grdGRIDBLOQUADOS.Rows - 1) > 0 Or (grdPEDIDOS.Rows - 1) > 0 Or (grdGERAL.Rows - 1) > 0 Then
       Call Operacao("R")
    End If
    
    Exit Sub
    
Err_Command1_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command1_Click()", Me.Name, "Command1_Click()", strCAMARQERRO)
    
End Sub

Private Sub Command2_Click()
    
On Error GoTo Err_Command2_Click

    Dim strSITUACAO As String

    If stPEDIDOS.Tab = 2 Then
       MsgBox "ATENÇÃO" & vbCrLf & _
              "Não é possivel realizar esta ação !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If stPEDIDOS.Tab = 3 Then Exit Sub
    
    If stPEDIDOS.Tab = 8 Then
       
        If Verif_Linha_Sel(grdGERAL, "L") = False Then Exit Sub
        
        strSITUACAO = Trim(grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao))
        If strSITUACAO = "M" Or _
           strSITUACAO = "F" Or _
           strSITUACAO = "P" Or _
           strSITUACAO = "S" Or _
           strSITUACAO = "R" Or _
           strSITUACAO = "V" Then
           MsgBox "ATENÇÃO" & vbCrLf & _
                  "Não é possivel realizar esta ação !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    If ((grdGRIDBLOQUADOS.Rows - 1) > 0) Or ((grdReprovados.Rows - 1) > 0) Or ((grdGERAL.Rows - 1) > 0) Then
       Call Operacao("LF")
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
  
    Dim strSITUACAO As String
  
    ''If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
    If stPEDIDOS.Tab = 1 Or stPEDIDOS.Tab = 2 Or stPEDIDOS.Tab = 5 Then
       MsgBox "Não pode ser liquidado !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    If stPEDIDOS.Tab = 3 Then
          With grdPEDFATURADO
             If Trim(.Cell(flexcpText, .RowSel, conCOL_SonFat_Situacao)) <> "P" Then
                 MsgBox "Não pode ser liquidado !!!", vbOKOnly + vbCritical, "Aviso"
                 Exit Sub
             End If
          End With
    End If
  
    If stPEDIDOS.Tab = 8 Then
          With grdGERAL
             strSITUACAO = Trim(.Cell(flexcpText, .RowSel, conCOL_SonFat_Situacao))
             If strSITUACAO = "B" Or _
                strSITUACAO = "N" Or _
                strSITUACAO = "F" Or _
                strSITUACAO = "M" Or _
                strSITUACAO = "R" Or _
                strSITUACAO = "V" Then
                 MsgBox "Não pode ser liquidado !!!", vbOKOnly + vbExclamation, "Aviso"
                 Exit Sub
             End If
          End With
    End If
  
  
    Dim iResp As Integer
    
    iResp = MsgBox("Confirma a liquidação do pedido ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
    If iResp <> 6 Then Exit Sub
  
    Call Operacao("M")
    If objCADPEDVENDA.Atualiza("M", Str(objCADPEDVENDA.CODPEDIDO), FILIAL, "frmCADPEDVENDA", intFILIALPED) = False Then Exit Sub
    
    Call AbilitaCampos
    Call FuncaoAtualiza

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
        Call Operacao("LS")
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
        Call Operacao("LV")
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

Private Sub Command8_Click()

On Error GoTo Err_Command8_Click

    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    If boolEVendedor = True Then
         
         sSql = "Select " & vbCrLf
         sSql = sSql & "       CLIE.* " & vbCrLf
        
         sSql = sSql & "  from " & vbCrLf
         sSql = sSql & "       SGI_CADCLIEVEND CVEN" & vbCrLf
         sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
         
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       CVEN.SGI_FILIAL = " & FILIAL & vbCrLf
         sSql = sSql & "   And CVEN.SGI_CODIGO = " & lngCodVendedor & vbCrLf
         sSql = sSql & "   And CLIE.SGI_FILIAL = CVEN.SGI_FILIAL" & vbCrLf
         sSql = sSql & "   And CLIE.SGI_CODIGO = CVEN.SGI_CODCLI"
         
    Else
        
         If Len(Trim(txtCODVENDEDOR.Text)) > 0 Then
         
            sSql = "Select " & vbCrLf
            sSql = sSql & "       CLIE.* " & vbCrLf
            
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADCLIEVEND CVEN" & vbCrLf
            sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       CVEN.SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "   And CVEN.SGI_CODIGO = " & Trim(txtCODVENDEDOR.Text) & vbCrLf
            sSql = sSql & "   And CLIE.SGI_FILIAL = CVEN.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And CLIE.SGI_CODIGO = CVEN.SGI_CODCLI"
         
         Else
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       CLIE.* " & vbCrLf
            
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADCLIENTE  CLIE" & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       CLIE.SGI_FILIAL = " & FILIAL
    
        End If
    End If
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLIE.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, txtNomClie, "Command8_Click()")
    If Len(Trim(txtNomClie.Text)) = 0 Then txtCODCLIE.Text = ""
    
    txtCODCLIE.SetFocus

    Exit Sub
    
Err_Command8_Click:
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command8_Click()", Me.Name, "Command8_Click()", strCAMARQERRO)

End Sub

Private Sub Command9_Click()

On Error GoTo Err_Command9_Click

    Dim strCodInVendedores  As String
    Dim strCodVendedores    As String

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL
    
    If lngCodVendedor > 0 Then
        strCodInVendedores = PegaVendedoresConjulgados(Str(lngCodVendedor))
        
        If Len(Trim(strCodInVendedores)) > 0 Then
            strCodVendedores = Trim(strCodInVendedores)
            sSql = sSql & "   And SGI_CODIGO In(" & strCodVendedores & ")"
        End If
        
    End If
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Venderores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVENDEDOR.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, txtNomVend, "Command9_Click()")
    If Len(Trim(txtNomVend.Text)) = 0 Then txtCODVENDEDOR.Text = ""
    
    If txtCODVENDEDOR.Enabled = True Then txtCODVENDEDOR.SetFocus

    Exit Sub
    
Err_Command9_Click:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command9_Click()", Me.Name, "Command9_Click()", strCAMARQERRO)

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
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    
    objCADPEDVENDA.FILIAL = FILIAL
    objFuncoes.LimpaCampos Me
    
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
   
    Call ConfGridPedidos
    Call ConfGridAgLib
    Call ConfGridReprovados
    Call ConfGridFaturado
    Call ConfGridBloqAlt
    Call ConfGridBloqAltLit
    Call ConfGridParaEstoque
    Call ConfGridBloqPDataPCota
    Call ConfGridGeral
    Call ConfLstStatus
   
    Call AtivaDesativaBotoes
   
    Command3.Visible = False
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    boolComAcao = False
    
    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO
    
    objCADPEDVENDA.FILIALPED = intFILIALPED
    
    mskDataI.Text = "__/__/____"
    mskDataF.Text = "__/__/____"
    
    Call SelecionaStatus(stPEDIDOS.Tab)
    
    lstOrdem.Clear
    Call OrdemPadrao
    
    '' Para Estoque
    stPEDIDOS.TabVisible(6) = False
    
    Call LimpaLabelQtdeRelac
    
    boolTelaAberta = False
    
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

Private Sub Operacao(strOperacao As String)
 
On Error GoTo Err_Operacao
  
  If stPEDIDOS.Tab = 0 Then
     If Verif_Linha_Sel(grdPEDIDOS, strOperacao) = False Then Exit Sub
  ElseIf stPEDIDOS.Tab = 1 Then
     If Verif_Linha_Sel(grdGRIDBLOQUADOS, strOperacao) = False Then Exit Sub
  ElseIf stPEDIDOS.Tab = 2 Then
     If Verif_Linha_Sel(grdReprovados, strOperacao) = False Then Exit Sub
  ElseIf stPEDIDOS.Tab = 3 Then
     If Verif_Linha_Sel(grdPEDFATURADO, strOperacao) = False Then Exit Sub
  ElseIf stPEDIDOS.Tab = 4 Then
     If Verif_Linha_Sel(grdBLOQALT, strOperacao) = False Then Exit Sub
  ElseIf stPEDIDOS.Tab = 5 Then
     If Verif_Linha_Sel(grdLIBLITO, strOperacao) = False Then Exit Sub
  ElseIf stPEDIDOS.Tab = 6 Then
     If Verif_Linha_Sel(grdPARAEST, strOperacao) = False Then Exit Sub
  ElseIf stPEDIDOS.Tab = 7 Then
     If Verif_Linha_Sel(grdLIBPDATAPCOTA, strOperacao) = False Then Exit Sub
  ElseIf stPEDIDOS.Tab = 8 Then
     If Verif_Linha_Sel(grdGERAL, strOperacao) = False Then Exit Sub
  End If
  
  Dim Pesquisa As String
  
  If stPEDIDOS.Tab = 0 Then
     If (grdPEDIDOS.Rows - 1) > 0 And grdPEDIDOS.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdPEDIDOS.Cell(flexcpText, grdPEDIDOS.RowSel, conCOL_SonPed_Codigo), "/", "")))
     If strOperacao = "L" Or strOperacao = "N" Then
        MsgBox "Pedido já Liberado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
  End If
  If stPEDIDOS.Tab = 1 Then
     If (grdGRIDBLOQUADOS.Rows - 1) > 0 And grdGRIDBLOQUADOS.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdGRIDBLOQUADOS.Cell(flexcpText, grdGRIDBLOQUADOS.RowSel, conCOL_SonAgLib_Codigo), "/", "")))
     
     If strOperacao = "LF" And grdGRIDBLOQUADOS.Cell(flexcpText, grdGRIDBLOQUADOS.RowSel, conCOL_SonAgLib_Situacao) = "B" Then
        MsgBox "Pedido Ainda Não Foi Liberado pelo Comercial !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     ElseIf strOperacao = "LN" And grdGRIDBLOQUADOS.Cell(flexcpText, grdGRIDBLOQUADOS.RowSel, conCOL_SonAgLib_Situacao) = "N" Then
        MsgBox "Pedido ja Foi Liberado pelo Comercial !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
  End If
  
  If stPEDIDOS.Tab = 2 Then
     If (grdReprovados.Rows - 1) > 0 And grdReprovados.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdReprovados.Cell(flexcpText, grdReprovados.RowSel, conCOL_SonRep_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 3 Then
     If (grdPEDFATURADO.Rows - 1) > 0 And grdPEDFATURADO.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdPEDFATURADO.Cell(flexcpText, grdPEDFATURADO.RowSel, conCOL_SonFat_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 4 Then
     If (grdBLOQALT.Rows - 1) > 0 And grdBLOQALT.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdBLOQALT.Cell(flexcpText, grdBLOQALT.RowSel, conCOL_SonBloq_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 5 Then
     If (grdLIBLITO.Rows - 1) > 0 And grdLIBLITO.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdLIBLITO.Cell(flexcpText, grdLIBLITO.RowSel, conCOL_SonBloqLit_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 6 Then
     If (grdPARAEST.Rows - 1) > 0 And grdPARAEST.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdPARAEST.Cell(flexcpText, grdPARAEST.RowSel, conCOL_SonParaEst_Codigo), "/", "")))
  End If
  If stPEDIDOS.Tab = 7 Then
     If (grdLIBPDATAPCOTA.Rows - 1) > 0 And grdLIBPDATAPCOTA.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdLIBPDATAPCOTA.Cell(flexcpText, grdLIBPDATAPCOTA.RowSel, conCOL_SonBloqPDPC_Codigo), "/", "")))
  End If
  
  If stPEDIDOS.Tab = 8 Then
     If (grdGERAL.Rows - 1) > 0 And grdGERAL.RowSel > 0 Then iCodigo = CLng(Trim(Replace(grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Codigo), "/", "")))
  
     If strOperacao = "LF" And grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao) = "B" Then
        MsgBox "Pedido Ainda Não Foi Liberado pelo Comercial !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     ElseIf strOperacao = "LN" And grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao) = "N" Then
        MsgBox "Pedido ja Foi Liberado pelo Comercial !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     ElseIf (strOperacao = "LN" And grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao) = "L") Or _
            (strOperacao = "LF" And grdGERAL.Cell(flexcpText, grdGERAL.RowSel, conCOL_SonGeral_Situacao) = "L") Then
        MsgBox "Pedido ja Foi Liberado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
  
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
  Call FuncaoAtualiza

  Exit Sub
  
Err_Operacao:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Operacao()", Me.Name, "Operacao()", strCAMARQERRO)

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
                              
       grdGRIDBLOQUADOS.AddItem Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                                Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_STATUS & vbTab & _
                                BREC!SGI_DESCRICAO
                                
                              
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
       
       .Cell(flexcpData, 0, conCOL_SonFat_Vendedor) = ""
       .ColDataType(conCOL_SonFat_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonFat_Codigo) = 1300
       .ColWidth(conCOL_SonFat_Data) = 1300
       .ColWidth(conCOL_SonFat_Cliente) = 6500
       .ColWidth(conCOL_SonFat_Situacao) = 250
       .ColWidth(conCOL_SonFat_Tipo) = 0
       .ColWidth(conCOL_SonFat_Vendedor) = 3500
    
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
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloq_Codigo), "/", "")))
    End With
End Sub

Private Sub grdBLOQALT_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdBLOQALT
        If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdBLOQALT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        With grdBLOQALT
            If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
        End With
    End If
End Sub

Private Sub grdBLOQALT_RowColChange()
    With grdBLOQALT
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloq_Codigo), "/", "")))
    End With
End Sub

Private Sub grdGERAL_Click()

On Error GoTo Err_grdGERAL_Click

    With grdGERAL
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .RowSel, conCOL_SonGeral_Codigo), "/", "")))
    End With
    
   Exit Sub
   
Err_grdGERAL_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdGERAL_Click()", Me.Name, "grdGERAL_Click()", strCAMARQERRO)

End Sub

Private Sub grdGERAL_DblClick()

On Error GoTo Err_grdGERAL_DblClick
   
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   If (grdGERAL.Rows - 1) > 0 And grdGERAL.RowSel > 0 Then Call Operacao("C")
   
   Exit Sub
   
Err_grdGERAL_DblClick:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdGERAL_DblClick()", Me.Name, "grdGERAL_DblClick()", strCAMARQERRO)

End Sub

Private Sub grdGERAL_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Err_grdGERAL_KeyDown
    
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        If (grdGERAL.Rows - 1) > 0 And grdGERAL.RowSel > 0 Then Call Operacao("C")
    End If
    
    Exit Sub
    
Err_grdGERAL_KeyDown:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdGERAL_KeyDown()", Me.Name, "grdGERAL_KeyDown()", strCAMARQERRO)

End Sub

Private Sub grdGERAL_RowColChange()

On Error GoTo Err_grdGERAL_RowColChange

    With grdGERAL
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .RowSel, conCOL_SonGeral_Codigo), "/", "")))
    End With
    
   Exit Sub
   
Err_grdGERAL_RowColChange:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdGERAL_RowColChange()", Me.Name, "grdGERAL_RowColChange()", strCAMARQERRO)

End Sub

Private Sub grdGRIDBLOQUADOS_Click()

On Error GoTo Err_grdGRIDBLOQUADOS_Click
    
    With grdGRIDBLOQUADOS
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonAgLib_Codigo), "/", "")))
    End With
    
    Exit Sub
    
Err_grdGRIDBLOQUADOS_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdGRIDBLOQUADOS_Click()", Me.Name, "grdGRIDBLOQUADOS_Click()", strCAMARQERRO)

End Sub

Private Sub grdGRIDBLOQUADOS_DblClick()

On Error GoTo Err_grdGRIDBLOQUADOS_DblClick
   
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   If (grdGRIDBLOQUADOS.Rows - 1) > 0 And grdGRIDBLOQUADOS.RowSel > 0 Then Call Operacao("C")
   
   Exit Sub
   
Err_grdGRIDBLOQUADOS_DblClick:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdGRIDBLOQUADOS_DblClick()", Me.Name, "grdGRIDBLOQUADOS_DblClick()", strCAMARQERRO)

End Sub

Private Sub grdGRIDBLOQUADOS_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Err_grdGRIDBLOQUADOS_KeyDown
    
    With grdGRIDBLOQUADOS
        If KeyCode = vbKeyReturn Then
            If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
            If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
        ElseIf KeyCode = vbKeySpace Then
        
        End If
    End With
    
    Exit Sub
    
Err_grdGRIDBLOQUADOS_KeyDown:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdGRIDBLOQUADOS_KeyDown()", Me.Name, "grdGRIDBLOQUADOS_KeyDown()", strCAMARQERRO)

End Sub

Private Sub grdGRIDBLOQUADOS_RowColChange()

On Error GoTo Err_grdGRIDBLOQUADOS_RowColChange
    
    With grdGRIDBLOQUADOS
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonAgLib_Codigo), "/", "")))
    End With
    
    Exit Sub
    
Err_grdGRIDBLOQUADOS_RowColChange:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdGRIDBLOQUADOS_RowColChange()", Me.Name, "grdGRIDBLOQUADOS_RowColChange()", strCAMARQERRO)

End Sub

Private Sub grdLIBLITO_Click()
    With grdLIBLITO
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloqLit_Codigo), "/", "")))
    End With
End Sub

Private Sub grdLIBLITO_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdLIBLITO
        If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdLIBLITO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        With grdLIBLITO
             If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
        End With
    End If
End Sub

Private Sub grdLIBLITO_RowColChange()
    With grdLIBLITO
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloqLit_Codigo), "/", "")))
    End With
End Sub

Private Sub grdLIBPDATAPCOTA_Click()
    With grdLIBPDATAPCOTA
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloqPDPC_Codigo), "/", "")))
    End With
End Sub

Private Sub grdLIBPDATAPCOTA_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdLIBPDATAPCOTA
        If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdLIBPDATAPCOTA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        With grdLIBPDATAPCOTA
             If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
        End With
    End If
End Sub

Private Sub grdLIBPDATAPCOTA_RowColChange()
    With grdLIBPDATAPCOTA
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonBloqPDPC_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPARAEST_Click()
    With grdPARAEST
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonParaEst_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPARAEST_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdPARAEST
        If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdPARAEST_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        With grdPARAEST
             If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
        End With
    End If
End Sub

Private Sub grdPARAEST_RowColChange()
    With grdPARAEST
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonParaEst_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPEDFATURADO_Click()
    With grdPEDFATURADO
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonFat_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPEDFATURADO_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdPEDFATURADO
        If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdPEDFATURADO_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        With grdPEDFATURADO
            If (.Rows - 1) > 0 And .RowSel > 0 Then Call Operacao("C")
        End With
    End If
End Sub

Private Sub grdPEDFATURADO_RowColChange()
    With grdPEDFATURADO
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonFat_Codigo), "/", "")))
    End With
End Sub

Private Sub grdPEDIDOS_Click()

On Error GoTo Err_grdPEDIDOS_Click

    With grdPEDIDOS
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .RowSel, conCOL_SonPed_Codigo), "/", "")))
    End With
    
   Exit Sub
   
Err_grdPEDIDOS_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPEDIDOS_Click()", Me.Name, "grdPEDIDOS_Click()", strCAMARQERRO)

End Sub

Private Sub grdPEDIDOS_DblClick()

On Error GoTo Err_grdPEDIDOS_DblClick
   
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   If (grdPEDIDOS.Rows - 1) > 0 And grdPEDIDOS.RowSel > 0 Then Call Operacao("C")
   
   Exit Sub
   
Err_grdPEDIDOS_DblClick:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPEDIDOS_DblClick()", Me.Name, "grdPEDIDOS_DblClick()", strCAMARQERRO)

End Sub

Private Sub grdPEDIDOS_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Err_grdPEDIDOS_KeyDown
    
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        If (grdPEDIDOS.Rows - 1) > 0 And grdPEDIDOS.RowSel > 0 Then Call Operacao("C")
    End If
    
    Exit Sub
    
Err_grdPEDIDOS_KeyDown:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPEDIDOS_KeyDown()", Me.Name, "grdPEDIDOS_KeyDown()", strCAMARQERRO)

End Sub

Private Sub grdPEDIDOS_RowColChange()

On Error GoTo Err_grdPEDIDOS_RowColChange

    With grdPEDIDOS
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .RowSel, conCOL_SonPed_Codigo), "/", "")))
    End With
    
   Exit Sub
   
Err_grdPEDIDOS_RowColChange:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPEDIDOS_RowColChange()", Me.Name, "grdPEDIDOS_RowColChange)", strCAMARQERRO)

End Sub

Private Sub grdReprovados_Click()

On Error GoTo Err_grdReprovados_Click
    
    With grdReprovados
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .Row, conCOL_SonRep_Codigo), "/", "")))
    End With
    
    Exit Sub
    
Err_grdReprovados_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdReprovados_Click()", Me.Name, "grdReprovados_Click()", strCAMARQERRO)

End Sub

Private Sub grdReprovados_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   If (grdReprovados.Rows - 1) > 0 And grdReprovados.RowSel > 0 Then Call Operacao("C")
End Sub

Private Sub grdReprovados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
        If (grdReprovados.Rows - 1) > 0 And grdReprovados.RowSel > 0 Then Call Operacao("C")
    End If
End Sub

Private Sub grdReprovados_RowColChange()

On Error GoTo Err_grdReprovados_RowColChange
    
    With grdReprovados
        If (.Rows - 1) > 0 And .RowSel > 0 Then objCADPEDVENDA.CODPEDIDO = CLng(Trim(Replace(.Cell(flexcpText, .RowSel, conCOL_SonRep_Codigo), "/", "")))
    End With
    
    Exit Sub
    
Err_grdReprovados_RowColChange:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdReprovados_RowColChange()", Me.Name, "grdReprovados_RowColChange()", strCAMARQERRO)

End Sub


Private Sub lstOrdem_DragDrop(Source As Control, x As Single, Y As Single)
    Dim i As Integer
    For i = 0 To (lstOrdem.ListCount - 1)
        If lstOrdem.ItemData(i) = Source.TabIndex Then Exit Sub
    Next i
    
    lstOrdem.AddItem Source
    lstOrdem.ItemData(lstOrdem.NewIndex) = Source.TabIndex
End Sub


Private Sub lstOrdem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If lstOrdem.ListCount = 0 Then Exit Sub
        If lstOrdem.ListIndex = -1 Then Exit Sub
        lstOrdem.RemoveItem lstOrdem.ListIndex
    End If
End Sub

Private Sub lstStatus_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
''    lstStatus.DragMode = 0
End Sub

Private Sub lstStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''    If Button = 2 Then lstStatus.DragMode = 1
End Sub

Private Sub mskDataF_GotFocus()
    objFuncoes.SelecionaCampos mskDataF.Name, Me
End Sub

Private Sub mskDataI_GotFocus()
    objFuncoes.SelecionaCampos mskDataI.Name, Me
End Sub

Private Sub optLiberados_Click(Index As Integer)
    
On Error GoTo Err_optLiberados_Click
    
    Dim strCAMPO  As String
    
    If BREC.State = 1 Then BREC.Close
    
    If stPEDIDOS.Tab <> 1 Then Exit Sub
    
    Call ConfGridAgLib
    
    If ConsisteCampos = False Then Exit Sub
    
    Dim strNOMETABELA As String
    
    If intFILIALPED = 0 Then strNOMETABELA = "SGI_CADPEDVENDH"
    If intFILIALPED = 1 Then strNOMETABELA = "SGI_CADPEDVENDH_STEEL"
    
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
    
    If Index = 0 Then sSql = sSql & "   And (PED.SGI_STATUS = 'B')" & vbCrLf
    If Index = 1 Then sSql = sSql & "   And (PED.SGI_STATUS = 'N')" & vbCrLf
    
    If lngCodVendedor > 0 Then
        sSql = sSql & "   And PED.SGI_CODVEND = " & lngCodVendedor & vbCrLf
    End If
    
    sSql = sSql & "   And PED.SGI_FILIALPED = " & intFILIALPED & vbCrLf
    
    If Len(Trim(txtCODPED.Text)) > 0 Then sSql = sSql & "   And PED.SGI_CODIGO = " & Trim(txtCODPED.Text) & vbCrLf
    If Len(Trim(txtCODVENDEDOR.Text)) > 0 Then sSql = sSql & "   And PED.SGI_CODVEND = " & Trim(txtCODVENDEDOR.Text) & vbCrLf
    If Len(Trim(txtCODCLIE.Text)) > 0 Then sSql = sSql & "   And PED.SGI_CODCLI = " & Trim(txtCODCLIE.Text) & vbCrLf
    
    If Len(Trim(Replace(Replace(mskDataI.Text, "/", ""), "_", ""))) > 0 And Len(Trim(Replace(Replace(mskDataF.Text, "/", ""), "_", ""))) > 0 Then
        sSql = sSql & "   And PED.SGI_DATAPED Between '" & Format(CDate(mskDataI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDataF.Text), "MM/DD/YYYY") & "'"
    End If
    
    ''If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by PED.SGI_CODIGO "
    ''If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by CLI.SGI_RAZAOSOC "
    ''If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order by PED.SGI_DATAPED "
    ''If cboFiltro.ListIndex = 3 Then sSql = sSql & "Order by ESP.SGI_DESCRICAO "

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
      
    Do While Not BREC.EOF
       
       strCAMPO = Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                  Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                  BREC!SGI_RAZAOSOC & vbTab & _
                  BREC!SGI_STATUS & vbTab & _
                  BREC!SGI_DESCRICAO
       
       grdGRIDBLOQUADOS.AddItem strCAMPO
       
       BREC.MoveNext
    Loop
    
    BREC.Close

    Exit Sub

Err_optLiberados_Click:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : optLiberados_Click()", Me.Name, "optLiberados_Click()", strCAMARQERRO)

End Sub

Private Sub Atualiza_Grid(msfGRID As VSFlexGrid, lngCOL As Long)
    
On Error GoTo Err_Atualiza_Grid
     
    If boolComAcao = True Then Exit Sub
    
    Dim i                  As Integer
    Dim bolAchou           As Boolean
    Dim lngCOL0            As Long
    Dim lngCOL1            As Long
    Dim lngCOL2            As Long
    Dim lngCOL3            As Long
    Dim lngCOL4            As Long
    Dim lngCOL5            As Long
    Dim strACAO            As String
    Dim strSTATUS          As String
    Dim strCODIGO          As String
    Dim lngLINHA           As Long
    
    Dim strEMPRESA_TAB     As String
    Dim strCodInVendedores As String
    Dim strCodVendedores   As String
     
    strEMPRESA_TAB = ""
    If intFILIALPED = 1 Then strEMPRESA_TAB = "_STEEL"
    
    If strOperacao = "P" Then Exit Sub
    
    bolAchou = False
    
    strACAO = ""
    strCODIGO = ""
    strSTATUS = ""
    lngLINHA = -1
      
    sSql = ""
      
    sSql = "Select" & vbCrLf
    sSql = sSql & "       ATU.SGI_ACAO " & vbCrLf
    sSql = sSql & "     , ATU.SGI_CODIGO" & vbCrLf
    sSql = sSql & "     , PED.SGI_STATUS" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ATUALIZA              ATU" & vbCrLf
    sSql = sSql & "     , SGI_CADPEDVENDH" & strEMPRESA_TAB & "     PED" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       ATU.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And ATU.SGI_MODULO    = 'frmCADPEDVENDA'" & vbCrLf
    sSql = sSql & "   And ATU.SGI_FILIALPED = " & intFILIALPED & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL    = ATU.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO    = ATU.SGI_CODIGO"
    
    BRECATU.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BRECATU.EOF() Then
       strACAO = Trim(BRECATU!SGI_ACAO)
       strSTATUS = Trim(BRECATU!SGI_STATUS)
       strCODIGO = Trim(BRECATU!SGI_CODIGO)
    End If
    BRECATU.Close
       
    If Len(Trim(strACAO)) = 0 Then Exit Sub
        
    With msfGRID
        If stPEDIDOS.Tab = 0 Then
            
            lngCOL0 = conCOL_SonPed_Codigo
            lngCOL1 = conCOL_SonPed_Data
            lngCOL2 = conCOL_SonPed_Cliente
            lngCOL3 = conCOL_SonPed_Situacao
            lngCOL4 = conCOL_SonPed_Tipo
            lngCOL5 = conCOL_SonPed_Vendedor
            
            lngLINHA = .FindRow(Format(strCODIGO, "#/####"), , lngCOL0)
            If lngLINHA > -1 Then
                If (Trim(strACAO) = "E" Or Trim(strACAO) = "D" Or Trim(strACAO) = "R" Or Trim(strACAO) = "M") Then
                    If (.Rows - 1) = 1 Then .Rows = 1
                    If (.Rows - 1) > 1 Then .RemoveItem lngLINHA
                ElseIf Trim(strACAO) = "I" Or Trim(strACAO) = "A" Or Trim(strACAO) = "L" Or Trim(strACAO) = "LC" Or Trim(strACAO) = "LF" Then
                    bolAchou = True
                End If
            End If
            
        ElseIf stPEDIDOS.Tab = 1 Then
            
            lngCOL0 = conCOL_SonAgLib_Codigo
            lngCOL1 = conCOL_SonAgLib_Data
            lngCOL2 = conCOL_SonAgLib_Cliente
            lngCOL3 = conCOL_SonAgLib_Situacao
            lngCOL4 = conCOL_SonAgLib_Tipo
            lngCOL5 = conCOL_SonAgLib_Vendedor
            
            lngLINHA = .FindRow(Format(strCODIGO, "#/####"), , lngCOL0)
            If lngLINHA > -1 Then
                If (Trim(strACAO) = "E" Or Trim(strACAO) = "LF" Or Trim(strSTATUS) = "V" Or Trim(strACAO) = "D" Or Trim(strACAO) = "R") Then
                    If (.Rows - 1) = 1 Then .Rows = 1
                    If (.Rows - 1) > 1 Then .RemoveItem lngLINHA
                ElseIf (Trim(strACAO) = "I" Or Trim(strACAO) = "A" Or Trim(strACAO) = "LN" Or Trim(strACAO) = "LV" Or Trim(strACAO) = "LS") Then
                    bolAchou = True
                End If
                
            End If
            
        
        ElseIf stPEDIDOS.Tab = 2 Then
        
            lngCOL0 = conCOL_SonRep_Codigo
            lngCOL1 = conCOL_SonRep_Data
            lngCOL2 = conCOL_SonRep_Cliente
            lngCOL3 = conCOL_SonRep_Situacao
            lngCOL4 = conCOL_SonRep_Tipo
            lngCOL5 = conCOL_SonRep_Vendedor
            
            lngLINHA = .FindRow(Format(strCODIGO, "#/####"), , lngCOL0)
            If lngLINHA > -1 Then
                If Trim(strACAO) = "E" Or Trim(strACAO) = "L" Or Trim(strACAO) = "N" Or Trim(strACAO) = "LN" Or Trim(strACAO) = "LF" Or Trim(strACAO) = "D" Then
                      If (.Rows - 1) = 1 Then .Rows = 1
                      If (.Rows - 1) > 1 Then .RemoveItem lngLINHA
                ElseIf Trim(strACAO) = "I" Or Trim(strACAO) = "A" Or Trim(strACAO) = "R" Then
                      bolAchou = True
                End If
            End If
        
        ElseIf stPEDIDOS.Tab = 3 Then
            
            lngCOL0 = conCOL_SonFat_Codigo
            lngCOL1 = conCOL_SonFat_Data
            lngCOL2 = conCOL_SonFat_Cliente
            lngCOL3 = conCOL_SonFat_Situacao
            lngCOL4 = conCOL_SonFat_Tipo
            lngCOL5 = conCOL_SonFat_Vendedor
            
            lngLINHA = .FindRow(Format(strCODIGO, "#/####"), , lngCOL0)
            
            If lngLINHA > -1 Then
                If Trim(strACAO) = "E" Then
                      If (.Rows - 1) = 1 Then .Rows = 1
                      If (.Rows - 1) > 1 Then .RemoveItem lngLINHA
                ElseIf Trim(strACAO) = "I" Or Trim(strACAO) = "A" Or Trim(strACAO) = "M" Or Trim(strACAO) = "P" Or Trim(strACAO) = "F" Or Trim(strSTATUS) = "P" Or Trim(strSTATUS) = "F" Then
                      bolAchou = True
                End If
            End If
        
        ElseIf stPEDIDOS.Tab = 4 Then
            
            lngCOL0 = conCOL_SonBloq_Codigo
            lngCOL1 = conCOL_SonBloq_Data
            lngCOL2 = conCOL_SonBloq_Cliente
            lngCOL3 = conCOL_SonBloq_Situacao
            lngCOL4 = conCOL_SonBloq_Tipo
            lngCOL5 = conCOL_SonBloq_Vendedor
            
            lngLINHA = .FindRow(Format(strCODIGO, "#/####"), , lngCOL0)
            If lngLINHA > -1 Then
                If Trim(strACAO) = "E" Or Trim(strACAO) = "LS" Or strSTATUS = "V" Or strSTATUS = "C" Then
                      If (.Rows - 1) = 1 Then .Rows = 1
                      If (.Rows - 1) > 1 Then .RemoveItem lngLINHA
                ElseIf Trim(strACAO) = "I" Or Trim(strACAO) = "A" Or Trim(strACAO) = "D" Then
                      bolAchou = True
                End If
             End If
        
        ElseIf stPEDIDOS.Tab = 5 Then
            
            lngCOL0 = conCOL_SonBloqLit_Codigo
            lngCOL1 = conCOL_SonBloqLit_Data
            lngCOL2 = conCOL_SonBloqLit_Cliente
            lngCOL3 = conCOL_SonBloqLit_Situacao
            lngCOL4 = conCOL_SonBloqLit_Tipo
            lngCOL5 = conCOL_SonBloqLit_Vendedor
            
            lngLINHA = .FindRow(Format(strCODIGO, "#/####"), , lngCOL0)
            If lngLINHA > -1 Then
                If Trim(strACAO) = "E" Or Trim(strACAO) = "V" Or Trim(strACAO) = "LV" Then
                      If (.Rows - 1) = 1 Then .Rows = 1
                      If (.Rows - 1) > 1 Then .RemoveItem lngLINHA
                ElseIf Trim(strACAO) = "I" Or Trim(strACAO) = "A" Or Trim(strACAO) = "D" Or Trim(strACAO) = "LS" Then
                      bolAchou = True
                End If
            End If
        
        ElseIf stPEDIDOS.Tab = 7 Then
            
            lngCOL0 = conCOL_SonBloqPDPC_Codigo
            lngCOL1 = conCOL_SonBloqPDPC_Data
            lngCOL2 = conCOL_SonBloqPDPC_Cliente
            lngCOL3 = conCOL_SonBloqPDPC_Situacao
            lngCOL4 = conCOL_SonBloqPDPC_Tipo
            lngCOL5 = conCOL_SonBloqPDPC_Vendedor
            
            lngLINHA = .FindRow(Format(strCODIGO, "#/####"), , lngCOL0)
            If lngLINHA > -1 Then
                If (Trim(strACAO) = "E" Or Trim(strACAO) = "LC" Or Trim(strACAO) = "M") Then
                      If (.Rows - 1) = 1 Then .Rows = 1
                      If (.Rows - 1) > 1 Then .RemoveItem lngLINHA
                ElseIf (Trim(strACAO) = "I" Or Trim(strACAO) = "A" Or Trim(strACAO) = "L" Or Trim(strACAO) = "LF" Or Trim(strACAO) = "LS") Then
                      bolAchou = True
                End If
            End If
        
        ElseIf stPEDIDOS.Tab = 8 Then
            
            lngCOL0 = conCOL_SonGeral_Codigo
            lngCOL1 = conCOL_SonGeral_Data
            lngCOL2 = conCOL_SonGeral_Cliente
            lngCOL3 = conCOL_SonGeral_Situacao
            lngCOL4 = conCOL_SonGeral_Tipo
            lngCOL5 = conCOL_SonGeral_Vendedor
            
            lngLINHA = .FindRow(Format(strCODIGO, "#/####"), , lngCOL0)
            If lngLINHA > -1 Then
                If (Trim(strACAO) = "E") Then
                    If (.Rows - 1) = 1 Then .Rows = 1
                    If (.Rows - 1) > 1 Then .RemoveItem lngLINHA
                Else
                    bolAchou = True
                End If
            End If
        
        End If
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       PED.* " & vbCrLf
        sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
        sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "      ,VEN.SGI_DESCRICAO As SGI_NOMVEND" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA_TAB & "  PED " & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE   CLI " & vbCrLf
        sSql = sSql & "      ,SGI_CADESPORCA   ESP " & vbCrLf
        sSql = sSql & "      ,SGI_CADVENDEDOR  VEN " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
        sSql = sSql & "   And VEN.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And VEN.SGI_CODIGO = PED.SGI_CODVEND " & vbCrLf
        
        If stPEDIDOS.Tab = 0 Then
            sSql = sSql & "   And PED.SGI_STATUS = 'L' " & vbCrLf
        ElseIf stPEDIDOS.Tab = 1 Then
            sSql = sSql & "   And (PED.SGI_STATUS = 'B' or PED.SGI_STATUS = 'N')" & vbCrLf
        ElseIf stPEDIDOS.Tab = 2 Then
            sSql = sSql & "   And PED.SGI_STATUS = 'R' " & vbCrLf
        ElseIf stPEDIDOS.Tab = 3 Then
            sSql = sSql & "   And (PED.SGI_STATUS = 'F' or PED.SGI_STATUS = 'M' or PED.SGI_STATUS = 'P' )" & vbCrLf
        ElseIf stPEDIDOS.Tab = 4 Then
            sSql = sSql & "   And PED.SGI_STATUS = 'S' " & vbCrLf
        ElseIf stPEDIDOS.Tab = 5 Then
            sSql = sSql & "   And PED.SGI_STATUS = 'V' " & vbCrLf
        ElseIf stPEDIDOS.Tab = 7 Then
            sSql = sSql & "   And (PED.SGI_STATUS = 'C' or PED.SGI_STATUS = '4')" & vbCrLf
        End If
        
        sSql = sSql & "   And PED.SGI_CODIGO = " & Trim(strCODIGO) & vbCrLf
        
        If lngCodVendedor > 0 Then
            strCodInVendedores = PegaVendedoresConjulgados(Str(lngCodVendedor))
            
            If Len(Trim(strCodInVendedores)) > 0 Then
                strCodVendedores = Trim(Str(lngCodVendedor)) & "," & Trim(strCodInVendedores)
            Else
                strCodVendedores = Trim(Str(lngCodVendedor))
            End If
            
            sSql = sSql & "   And PED.SGI_CODVEND In(" & Trim(strCodVendedores) & ")" & vbCrLf
        End If
        BRECATU.Open sSql, adoBanco_Dados, adOpenDynamic
        
        If bolAchou = False Then
            Do While Not BRECATU.EOF()
                .AddItem Format(BRECATU!SGI_CODIGO, "#/####") & vbTab & _
                                Format(BRECATU!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                                BRECATU!SGI_RAZAOSOC & vbTab & _
                                IIf(BRECATU!SGI_STATUS = "4", "D", BRECATU!SGI_STATUS) & vbTab & _
                                BRECATU!SGI_DESCRICAO & vbTab & _
                                BRECATU!SGI_NOMVEND
                BRECATU.MoveNext
            Loop
        ElseIf bolAchou = True Then
            If Not BRECATU.EOF() Then
                .Cell(flexcpText, lngLINHA, lngCOL0) = Format(BRECATU!SGI_CODIGO, "#/####")
                .Cell(flexcpText, lngLINHA, lngCOL1) = Format(BRECATU!SGI_DATAPED, "DD/MM/YYYY")
                .Cell(flexcpText, lngLINHA, lngCOL2) = BRECATU!SGI_RAZAOSOC
                .Cell(flexcpText, lngLINHA, lngCOL3) = IIf(BRECATU!SGI_STATUS = "4", "D", BRECATU!SGI_STATUS)
                .Cell(flexcpText, lngLINHA, lngCOL4) = BRECATU!SGI_DESCRICAO
                .Cell(flexcpText, lngLINHA, lngCOL5) = BRECATU!SGI_NOMVEND
            End If
        End If
        
        BRECATU.Close
    End With
      
    Exit Sub
     
Err_Atualiza_Grid:
     
    If BRECATU.State = 1 Then BRECATU.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Atualiza_Grid()", Me.Name, "Atualiza_Grid()", strCAMARQERRO)
      
End Sub

Private Sub Ordem()

On Error GoTo Err_Ordem
    
    Call FechaTableSpace
    
    Call ConfGridPedidos
    Call ConfGridAgLib
    Call ConfGridReprovados
    Call ConfGridFaturado
    Call ConfGridBloqAlt
    Call ConfGridBloqAltLit
    Call ConfGridParaEstoque
    Call ConfGridBloqPDataPCota
    Call ConfGridGeral
     
    Dim strCAMPO            As String
    Dim strNOMETABELA       As String
    Dim strEMPTABELA        As String
    Dim strCodVendedores    As String
    Dim strCodInVendedores  As String
    Dim boolTEMCAMPOPREENCH As Boolean
    
    boolTEMCAMPOPREENCH = False
    
    strEMPTABELA = ""
    If intFILIALPED = 0 Then strNOMETABELA = "SGI_CADPEDVENDH"
    If intFILIALPED = 1 Then
       strNOMETABELA = "SGI_CADPEDVENDH_STEEL"
       strEMPTABELA = "_STEEL"
    End If
    
    sSql = ""
  
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,VEN.SGI_DESCRICAO As SGI_NOMVEND" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMETABELA & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  ESP " & vbCrLf
    sSql = sSql & "      ,SGI_CADVENDEDOR VEN " & vbCrLf
    
    If Len(Trim(txtCODOP.Text)) > 0 Then
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPTABELA & " ORP" & vbCrLf
        boolTEMCAMPOPREENCH = True
    End If
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    
    If Len(Trim(txtNomClie.Text)) > 0 Then
        sSql = sSql & "   And CLI.SGI_RAZAOSOC Like '" & Trim(txtNomClie.Text) & "%'" & vbCrLf
        boolTEMCAMPOPREENCH = True
    End If
    
    If Len(Trim(txtCODOP.Text)) > 0 Then
        sSql = sSql & "   And ORP.SGI_FILIAL = CLI.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ORP.SGI_CODIGO = " & Trim(txtCODOP.Text) & vbCrLf
        sSql = sSql & "   And PED.SGI_FILIAL = ORP.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO = ORP.SGI_CODPED " & vbCrLf
    End If
    
    sSql = sSql & "   And ESP.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PED.SGI_CODTIPORC " & vbCrLf
    sSql = sSql & "   And VEN.SGI_FILIAL = PED.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And VEN.SGI_CODIGO = PED.SGI_CODVEND" & vbCrLf
    
    If Len(Trim(txtNomVend.Text)) > 0 Then
        sSql = sSql & "   And VEN.SGI_DESCRICAO Like '" & Trim(txtNomVend.Text) & "%'" & vbCrLf
        boolTEMCAMPOPREENCH = True
    End If
    
    If Len(Trim(Replace(Replace(mskDataI.Text, "/", ""), "_", ""))) > 0 And Len(Trim(Replace(Replace(mskDataF.Text, "/", ""), "_", ""))) > 0 Then
        sSql = sSql & "   And PED.SGI_DATAPED Between '" & Format(CDate(mskDataI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDataF.Text), "MM/DD/YYYY") & "'"
        boolTEMCAMPOPREENCH = True
    End If
    
    If Len(Trim(txtCODPED.Text)) > 0 Then
        sSql = sSql & "   And PED.SGI_CODIGO = " & Trim(txtCODPED.Text) & vbCrLf
        boolTEMCAMPOPREENCH = True
    End If
    If Len(Trim(txtCODVENDEDOR.Text)) > 0 Then
        sSql = sSql & "   And PED.SGI_CODVEND = " & Trim(txtCODVENDEDOR.Text) & vbCrLf
        boolTEMCAMPOPREENCH = True
    End If
    If Len(Trim(txtCODCLIE.Text)) > 0 Then
        sSql = sSql & "   And PED.SGI_CODCLI = " & Trim(txtCODCLIE.Text) & vbCrLf
        boolTEMCAMPOPREENCH = True
    End If
    
    If stPEDIDOS.Tab = 0 Then sSql = sSql & "   And PED.SGI_STATUS = 'L' " & vbCrLf
    If stPEDIDOS.Tab = 1 Then sSql = sSql & "   And (PED.SGI_STATUS = 'B' or PED.SGI_STATUS = 'N')" & vbCrLf
    If stPEDIDOS.Tab = 2 Then sSql = sSql & "   And PED.SGI_STATUS = 'R' " & vbCrLf
    If stPEDIDOS.Tab = 3 Then sSql = sSql & "   And (PED.SGI_STATUS = 'F' or PED.SGI_STATUS = 'P' or PED.SGI_STATUS = 'M') " & vbCrLf
    If stPEDIDOS.Tab = 4 Then sSql = sSql & "   And PED.SGI_STATUS = 'S' " & vbCrLf
    If stPEDIDOS.Tab = 5 Then sSql = sSql & "   And PED.SGI_STATUS = 'V' " & vbCrLf
    If stPEDIDOS.Tab = 6 Then sSql = sSql & "   And PED.SGI_STATUS = 'X' " & vbCrLf
    If stPEDIDOS.Tab = 7 Then sSql = sSql & "   And (PED.SGI_STATUS = 'C' or PED.SGI_STATUS = '4') " & vbCrLf
    
    If stPEDIDOS.Tab = 8 Then
        
        If boolTEMCAMPOPREENCH = False Then
            
            Dim intIndice           As Integer
            Dim arrCAMPO()          As String
            Dim strCAMPOARRAY       As String
            Dim intItenSelecionado  As Integer
            
            strCAMPOARRAY = ""
            intItenSelecionado = 0
            For intIndice = 0 To (lstStatus.ListCount - 1)
                If lstStatus.Selected(intIndice) = True Then
                    intItenSelecionado = intItenSelecionado + 1
                    
                    arrCAMPO = Split(lstStatus.List(intIndice), "-")
                    strCAMPOARRAY = strCAMPOARRAY & "'" & Trim(IIf(Trim(arrCAMPO(0)) = "D", "4", arrCAMPO(0))) & "'"
                    
                    If intItenSelecionado < lstStatus.SelCount Then strCAMPOARRAY = strCAMPOARRAY & ","
                End If
            Next intIndice
            
            If Len(Trim(strCAMPOARRAY)) > 0 Then
               strCAMPOARRAY = "In(" & strCAMPOARRAY & ")"
               sSql = sSql & "   And PED.SGI_STATUS " & strCAMPOARRAY & vbCrLf
            End If
        
        End If
        
    End If
    
    If lngCodVendedor > 0 Then
        strCodInVendedores = PegaVendedoresConjulgados(Str(lngCodVendedor))
        
        If Len(Trim(strCodInVendedores)) > 0 Then
            strCodVendedores = Trim(Str(lngCodVendedor)) & "," & Trim(strCodInVendedores)
        Else
            strCodVendedores = Trim(Str(lngCodVendedor))
        End If
        
        sSql = sSql & "   And PED.SGI_CODVEND In(" & Trim(strCodVendedores) & ")" & vbCrLf
    End If
    
    If lstOrdem.ListCount > 0 Then sSql = sSql & ConfCamposOrdem
    
    BREC.Open sSql, adoBanco_Dados
      
    If Not BREC.EOF Then
        Do While Not BREC.EOF
            strCAMPO = Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                       Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                       BREC!SGI_RAZAOSOC & vbTab & _
                       IIf(Trim(BREC!SGI_STATUS) = "4", "D", Trim(BREC!SGI_STATUS)) & vbTab & _
                       BREC!SGI_DESCRICAO & vbTab & _
                       BREC!SGI_NOMVEND
           
           
            If stPEDIDOS.Tab = 0 Then grdPEDIDOS.AddItem strCAMPO
            If stPEDIDOS.Tab = 1 Then grdGRIDBLOQUADOS.AddItem strCAMPO
            If stPEDIDOS.Tab = 2 Then grdReprovados.AddItem strCAMPO
            If stPEDIDOS.Tab = 3 Then grdPEDFATURADO.AddItem strCAMPO
            If stPEDIDOS.Tab = 4 Then grdBLOQALT.AddItem strCAMPO
            If stPEDIDOS.Tab = 5 Then grdLIBLITO.AddItem strCAMPO
            If stPEDIDOS.Tab = 6 Then grdPARAEST.AddItem strCAMPO
            If stPEDIDOS.Tab = 7 Then grdLIBPDATAPCOTA.AddItem strCAMPO
            If stPEDIDOS.Tab = 8 Then grdGERAL.AddItem strCAMPO
            
            BREC.MoveNext
        Loop
    Else
        MsgBox "ATENÇÂO" & vbCrLf & "Não há dados para pesquisar !!!", vbOKOnly + vbOKOnly, "Aviso"
    End If
    BREC.Close
    

    If stPEDIDOS.Tab = 0 Then
        Call PintaGride(grdPEDIDOS, conCOL_SonPed_Situacao, conCOL_SonPed_Situacao, conCOL_SonPed_Situacao)
        Call ContaItensRelacionados(grdPEDIDOS, 7)
    ElseIf stPEDIDOS.Tab = 1 Then
        Call PintaGride(grdGRIDBLOQUADOS, conCOL_SonAgLib_Situacao, conCOL_SonAgLib_Situacao, conCOL_SonAgLib_Situacao)
        Call ContaItensRelacionados(grdGRIDBLOQUADOS, 6)
    ElseIf stPEDIDOS.Tab = 2 Then
        Call PintaGride(grdReprovados, conCOL_SonRep_Situacao, conCOL_SonRep_Situacao, conCOL_SonRep_Situacao)
        Call ContaItensRelacionados(grdReprovados, 5)
    ElseIf stPEDIDOS.Tab = 3 Then
        Call PintaGride(grdPEDFATURADO, conCOL_SonFat_Situacao, conCOL_SonFat_Situacao, conCOL_SonFat_Situacao)
        Call ContaItensRelacionados(grdPEDFATURADO, 1)
    ElseIf stPEDIDOS.Tab = 4 Then
        Call ContaItensRelacionados(grdBLOQALT, 4)
    ElseIf stPEDIDOS.Tab = 5 Then
        Call ContaItensRelacionados(grdLIBLITO, 0)
    ElseIf stPEDIDOS.Tab = 7 Then
        Call ContaItensRelacionados(grdLIBPDATAPCOTA, 3)
    ElseIf stPEDIDOS.Tab = 8 Then
        Call PintaGride(grdGERAL, conCOL_SonGeral_Situacao, conCOL_SonGeral_Situacao, conCOL_SonGeral_Situacao)
        Call ContaItensRelacionados(grdGERAL, 2)
    End If
    
    objFuncoes.LimpaCampos Me
    Call SelecionaStatus(stPEDIDOS.Tab)
    Call OrdemPadrao

    Exit Sub

Err_Ordem:
    
    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Ordem()", Me.Name, "Ordem()", strCAMARQERRO)


End Sub

Private Sub stPEDIDOS_Click(PreviousTab As Integer)
    Call SelecionaStatus(stPEDIDOS.Tab)
End Sub

Private Sub Timer1_Timer()
    Call FuncaoAtualiza
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
    
    Dim boolAtDeBotoesPesqVend As Boolean
    
    
    cmdLibera.Visible = objCADPEDVENDA.PermiteLibComercial(lngCodUsuaro)
    Command2.Visible = objCADPEDVENDA.PermiteLibFinanceiro(lngCodUsuaro)
    cmdDeslib.Visible = PermiteBloqSN
    Command1.Visible = PermiteReprova
    Command4.Visible = PermiteLiqPedido
    Command5.Visible = PermiteLibPedBloq
    Command6.Visible = PermiteLibPedFotolito
    cmdLIBPDATAPCOTA.Visible = PermiteLibPDataPCota
    
    '' Permições para vendedores
    boolEVendedor = PermiteEVendedor
    If boolEVendedor = False Then
        boolAtDeBotoesPesqVend = True
    Else
        boolAtDeBotoesPesqVend = PermiteConsultarOutroVendedor
    End If
    
    Command9.Visible = boolAtDeBotoesPesqVend
    txtNomVend.Visible = boolAtDeBotoesPesqVend
    lblCampo(3).Visible = boolAtDeBotoesPesqVend
    txtCODVENDEDOR.Visible = boolAtDeBotoesPesqVend
    lblCampo(4).Visible = boolAtDeBotoesPesqVend
    
    Exit Sub
    
Err_AtivaDesativaBotoes:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : AtivaDesativaBotoes()", Me.Name, "AtivaDesativaBotoes()", strCAMARQERRO)
    
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
    Set objPESQPADRAO = Nothing
End Sub

Private Sub PintaGride(grdGenerica As VSFlexGrid, lngColSituacao As Long, lngColI As Long, lngColF As Long)
    Dim i As Long
    With grdGenerica
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, lngColSituacao) = "F" Then .Cell(flexcpBackColor, i, lngColI, i, lngColF) = &HC000&
            If .Cell(flexcpText, i, lngColSituacao) = "P" Then .Cell(flexcpBackColor, i, lngColI, i, lngColF) = &HC0C0&
            If .Cell(flexcpText, i, lngColSituacao) = "M" Then .Cell(flexcpBackColor, i, lngColI, i, lngColF) = &HC0C000
            If .Cell(flexcpText, i, lngColSituacao) = "L" Then .Cell(flexcpBackColor, i, lngColI, i, lngColF) = &H80FF80
            If .Cell(flexcpText, i, lngColSituacao) = "B" Then .Cell(flexcpBackColor, i, lngColI, i, lngColF) = &H8080FF
            If .Cell(flexcpText, i, lngColSituacao) = "N" Then .Cell(flexcpBackColor, i, lngColI, i, lngColF) = &H80FFFF
            If .Cell(flexcpText, i, lngColSituacao) = "R" Then .Cell(flexcpBackColor, i, lngColI, i, lngColF) = &H80FF&
        Next i
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
       
       .Cell(flexcpData, 0, conCOL_SonBloq_Vendedor) = ""
       .ColDataType(conCOL_SonBloq_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonBloq_Codigo) = 1300
       .ColWidth(conCOL_SonBloq_Data) = 1300
       .ColWidth(conCOL_SonBloq_Cliente) = 6500
       .ColWidth(conCOL_SonBloq_Situacao) = 250
       .ColWidth(conCOL_SonBloq_Tipo) = 0
       .ColWidth(conCOL_SonBloq_Vendedor) = 3500
    
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
       
       .Cell(flexcpData, 0, conCOL_SonBloqLit_Vendedor) = ""
       .ColDataType(conCOL_SonBloqLit_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonBloqLit_Codigo) = 1300
       .ColWidth(conCOL_SonBloqLit_Data) = 1300
       .ColWidth(conCOL_SonBloqLit_Cliente) = 6500
       .ColWidth(conCOL_SonBloqLit_Situacao) = 250
       .ColWidth(conCOL_SonBloqLit_Tipo) = 0
       .ColWidth(conCOL_SonBloqLit_Vendedor) = 3500
    
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
       
       .Cell(flexcpData, 0, conCOL_SonParaEst_Vendedor) = ""
       .ColDataType(conCOL_SonParaEst_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonParaEst_Codigo) = 1300
       .ColWidth(conCOL_SonParaEst_Data) = 1300
       .ColWidth(conCOL_SonParaEst_Cliente) = 6500
       .ColWidth(conCOL_SonParaEst_Situacao) = 250
       .ColWidth(conCOL_SonParaEst_Tipo) = 0
       .ColWidth(conCOL_SonParaEst_Vendedor) = 3500
    
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
    Dim i            As Long

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
        
        For i = 1 To UBound(arrOPES)
            
            sSql = ""
            
            '' SGI_PROGENTRPROD
            sSql = "Update SGI_PROGENTRPROD" & strNOMTABELA & " Set " & vbCrLf
            sSql = sSql & "                           SGI_IDINTERNO = " & arrOPES(i, 5) & vbCrLf
            sSql = sSql & "                     Where " & vbCrLf
            sSql = sSql & "                           SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "                       And SGI_CODPED      = " & arrOPES(i, 1) & vbCrLf
            sSql = sSql & "                       And SGI_IDPRODUTO   = " & arrOPES(i, 2) & vbCrLf
            sSql = sSql & "                       And SGI_INDICE      = " & arrOPES(i, 3) & vbCrLf
            sSql = sSql & "                       And SGI_DATENTREGA  = '" & Format(CDate(arrOPES(i, 4)), "MM/DD/YYYY") & "'"
            
            BGRV.CommandText = sSql
            BGRV.Execute
        
            sSql = ""
            
            sSql = "Update SGI_ORDEMPROD" & strNOMTABELA & " Set" & vbCrLf
            sSql = sSql & "                     SGI_IDPAI = " & arrOPES(i, 5) & vbCrLf
            sSql = sSql & "               Where " & vbCrLf
            sSql = sSql & "                    SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "                And SGI_CODPED     = " & arrOPES(i, 1) & vbCrLf
            sSql = sSql & "                And SGI_IDPRODUTO  = " & arrOPES(i, 2) & vbCrLf
            sSql = sSql & "                And SGI_QTDEPED    = " & arrOPES(i, 7)
        
            BGRV.CommandText = sSql
            BGRV.Execute
        
        Next i
        
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
       
       .Cell(flexcpData, 0, conCOL_SonBloqPDPC_Vendedor) = ""
       .ColDataType(conCOL_SonBloqPDPC_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonBloqPDPC_Codigo) = 1300
       .ColWidth(conCOL_SonBloqPDPC_Data) = 1300
       .ColWidth(conCOL_SonBloqPDPC_Cliente) = 6500
       .ColWidth(conCOL_SonBloqPDPC_Situacao) = 250
       .ColWidth(conCOL_SonBloqPDPC_Tipo) = 0
       .ColWidth(conCOL_SonBloqPDPC_Vendedor) = 3500
    
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

Private Sub FuncaoAtualiza()
    If stPEDIDOS.Tab = 0 Then
        Call Atualiza_Grid(grdPEDIDOS, conCOL_SonPed_Codigo)
        Call PintaGride(grdPEDIDOS, conCOL_SonPed_Situacao, conCOL_SonPed_Situacao, conCOL_SonPed_Situacao)
    ElseIf stPEDIDOS.Tab = 1 Then
        Call Atualiza_Grid(grdGRIDBLOQUADOS, conCOL_SonAgLib_Codigo)
        Call PintaGride(grdGRIDBLOQUADOS, conCOL_SonAgLib_Situacao, conCOL_SonAgLib_Situacao, conCOL_SonAgLib_Situacao)
    ElseIf stPEDIDOS.Tab = 2 Then
        Call Atualiza_Grid(grdReprovados, conCOL_SonRep_Codigo)
        Call PintaGride(grdReprovados, conCOL_SonRep_Situacao, conCOL_SonRep_Situacao, conCOL_SonRep_Situacao)
    ElseIf stPEDIDOS.Tab = 3 Then
        Call Atualiza_Grid(grdPEDFATURADO, conCOL_SonFat_Codigo)
        Call PintaGride(grdPEDFATURADO, conCOL_SonFat_Situacao, conCOL_SonFat_Situacao, conCOL_SonFat_Situacao)
    ElseIf stPEDIDOS.Tab = 4 Then
        Call Atualiza_Grid(grdBLOQALT, conCOL_SonBloq_Codigo)
        Call PintaGride(grdBLOQALT, conCOL_SonBloq_Situacao, conCOL_SonBloq_Situacao, conCOL_SonBloq_Situacao)
    ElseIf stPEDIDOS.Tab = 5 Then
        Call Atualiza_Grid(grdLIBLITO, conCOL_SonBloqLit_Codigo)
        Call PintaGride(grdLIBLITO, conCOL_SonBloqLit_Situacao, conCOL_SonBloqLit_Situacao, conCOL_SonBloqLit_Situacao)
    ElseIf stPEDIDOS.Tab = 6 Then
        Call Atualiza_Grid(grdPARAEST, conCOL_SonParaEst_Codigo)
        Call PintaGride(grdPARAEST, conCOL_SonParaEst_Situacao, conCOL_SonParaEst_Situacao, conCOL_SonParaEst_Situacao)
    ElseIf stPEDIDOS.Tab = 7 Then
        Call Atualiza_Grid(grdLIBPDATAPCOTA, conCOL_SonBloqPDPC_Codigo)
        Call PintaGride(grdLIBPDATAPCOTA, conCOL_SonBloqPDPC_Situacao, conCOL_SonBloqPDPC_Situacao, conCOL_SonBloqPDPC_Situacao)
    ElseIf stPEDIDOS.Tab = 8 Then
        Call Atualiza_Grid(grdGERAL, conCOL_SonGeral_Codigo)
        Call PintaGride(grdGERAL, conCOL_SonGeral_Situacao, conCOL_SonGeral_Situacao, conCOL_SonGeral_Situacao)
    End If
End Sub

Private Sub ConfGridAgLib()
        
    With grdGRIDBLOQUADOS
    
       .Cols = conColumnsIn_SonAgLib
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonAgLib_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonAgLib_Codigo) = ""
       .ColDataType(conCOL_SonAgLib_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonAgLib_Data) = ""
       .ColDataType(conCOL_SonAgLib_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonAgLib_Cliente) = ""
       .ColDataType(conCOL_SonAgLib_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonAgLib_Situacao) = ""
       .ColDataType(conCOL_SonBloq_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonAgLib_Tipo) = ""
       .ColDataType(conCOL_SonAgLib_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonAgLib_Vendedor) = ""
       .ColDataType(conCOL_SonAgLib_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonAgLib_Codigo) = 1300
       .ColWidth(conCOL_SonAgLib_Data) = 1300
       .ColWidth(conCOL_SonAgLib_Cliente) = 6500
       .ColWidth(conCOL_SonAgLib_Situacao) = 250
       .ColWidth(conCOL_SonAgLib_Tipo) = 0
       .ColWidth(conCOL_SonAgLib_Vendedor) = 3500
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub

Private Sub ConfGridPedidos()
        
    With grdPEDIDOS
    
       .Cols = conColumnsIn_SonPed
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonPed_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonPed_Codigo) = ""
       .ColDataType(conCOL_SonPed_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPed_Data) = ""
       .ColDataType(conCOL_SonPed_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonPed_Cliente) = ""
       .ColDataType(conCOL_SonPed_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPed_Situacao) = ""
       .ColDataType(conCOL_SonPed_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPed_Tipo) = ""
       .ColDataType(conCOL_SonPed_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPed_Vendedor) = ""
       .ColDataType(conCOL_SonPed_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonPed_Codigo) = 1300
       .ColWidth(conCOL_SonPed_Data) = 1300
       .ColWidth(conCOL_SonPed_Cliente) = 6500
       .ColWidth(conCOL_SonPed_Situacao) = 250
       .ColWidth(conCOL_SonPed_Tipo) = 0
       .ColWidth(conCOL_SonPed_Vendedor) = 3500
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub

Private Sub ConfGridReprovados()
        
    With grdReprovados
    
       .Cols = conColumnsIn_SonRep
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonRep_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonRep_Codigo) = ""
       .ColDataType(conCOL_SonRep_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonRep_Data) = ""
       .ColDataType(conCOL_SonRep_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonRep_Cliente) = ""
       .ColDataType(conCOL_SonRep_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonRep_Situacao) = ""
       .ColDataType(conCOL_SonRep_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonRep_Tipo) = ""
       .ColDataType(conCOL_SonRep_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonRep_Vendedor) = ""
       .ColDataType(conCOL_SonRep_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonRep_Codigo) = 1300
       .ColWidth(conCOL_SonRep_Data) = 1300
       .ColWidth(conCOL_SonRep_Cliente) = 6500
       .ColWidth(conCOL_SonRep_Situacao) = 250
       .ColWidth(conCOL_SonRep_Tipo) = 0
       .ColWidth(conCOL_SonRep_Vendedor) = 3500
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub


Private Sub txtCODCLIE_GotFocus()
    objFuncoes.SelecionaCampos txtCODCLIE.Name, Me
End Sub

Private Sub txtCODCLIE_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODCLIE.Text
End Sub

Private Sub txtCODCLIE_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODCLIE_Validate

    Dim i As Integer
    
    If Len(Trim(txtCODCLIE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLIE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If boolEVendedor = True Then
        If ConfereCliente(txtCODCLIE.Text, Str(lngCodVendedor)) = False Then
           txtCODCLIE.Text = ""
           txtNomClie.Text = ""
           Cancel = True
           Exit Sub
        End If
    Else
        If ConfereCliente(txtCODCLIE.Text, txtCODVENDEDOR.Text) = False Then
           txtCODCLIE.Text = ""
           txtNomClie.Text = ""
           Cancel = True
           Exit Sub
        End If
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCODCLIE.Text, txtNomClie, "txtCODCLIE_Validate()")
    If Len(Trim(txtNomClie.Text)) = 0 Then
       txtCODCLIE.Text = ""
       txtNomClie.Text = ""
       Cancel = True
       Exit Sub
    End If

    Exit Sub
    
Err_txtCODCLIE_Validate:
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODCLIE_Validate()", Me.Name, "txtCODCLIE_Validate()", strCAMARQERRO)

End Sub

Private Sub txtCODOP_GotFocus()
    objFuncoes.SelecionaCampos txtCODOP.Name, Me
End Sub

Private Sub txtCODOP_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODOP.Text
End Sub

Private Sub txtCODOP_Validate(Cancel As Boolean)

    If Len(Trim(txtCODOP.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODOP.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Somente é permitido numeros !!!"
        txtCODOP.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtCODPED_GotFocus()
    objFuncoes.SelecionaCampos txtCODPED.Name, Me
End Sub

Private Sub txtCODPED_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODPED.Text
End Sub

Private Sub txtCODPED_Validate(Cancel As Boolean)

    If Len(Trim(txtCODPED.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODPED.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Somente é permitido numeros !!!"
        txtCODPED.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtCODVENDEDOR_GotFocus()
    objFuncoes.SelecionaCampos txtCODVENDEDOR.Name, Me
End Sub

Private Sub txtCODVENDEDOR_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODVENDEDOR.Text
End Sub

Private Sub txtCODVENDEDOR_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODVENDEDOR_Validate

    Dim i As Integer
    
    If Len(Trim(txtCODVENDEDOR.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVENDEDOR.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVENDEDOR.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    
    If lngCodVendedor > 0 Then
        If ConfereVendedor(Str(lngCodVendedor), txtCODVENDEDOR.Text) = False Then
            MsgBox "ATENÇÃO" & vbCrLf & "Este vendedor não faz parte da consulta para o vendedor " & objFuncoes.Crypt(strUSUARIO) & " !!!", vbOKOnly + vbExclamation, "Aviso"
            Cancel = True
            Exit Sub
        End If
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVENDEDOR.Text, txtNomVend, "txtCODVENDEDOR_Validate()")
    If Len(Trim(txtNomVend.Text)) = 0 Then
       txtCODVENDEDOR.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCODVENDEDOR_Validate:
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDEDOR_Validate()", Me.Name, "txtCODVENDEDOR_Validate()", strCAMARQERRO)

End Sub

Private Sub txtNomClie_GotFocus()
    objFuncoes.SelecionaCampos txtNomClie.Name, Me
End Sub

Private Sub txtNomClie_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtNomVend_GotFocus()
    objFuncoes.SelecionaCampos txtNomVend.Name, Me
End Sub

Private Sub txtNomVend_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub


Private Sub ConfGridGeral()
        
    With grdGERAL
    
       .Cols = conColumnsIn_SonGeral
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonGeral_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonGeral_Codigo) = ""
       .ColDataType(conCOL_SonGeral_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonGeral_Data) = ""
       .ColDataType(conCOL_SonGeral_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonGeral_Cliente) = ""
       .ColDataType(conCOL_SonGeral_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonGeral_Situacao) = ""
       .ColDataType(conCOL_SonGeral_Situacao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonGeral_Tipo) = ""
       .ColDataType(conCOL_SonGeral_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonGeral_Vendedor) = ""
       .ColDataType(conCOL_SonGeral_Vendedor) = flexDTString
       
       .ColWidth(conCOL_SonGeral_Codigo) = 1300
       .ColWidth(conCOL_SonGeral_Data) = 1300
       .ColWidth(conCOL_SonGeral_Cliente) = 6500
       .ColWidth(conCOL_SonGeral_Situacao) = 250
       .ColWidth(conCOL_SonGeral_Tipo) = 0
       .ColWidth(conCOL_SonGeral_Vendedor) = 3500
    
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With

End Sub


Private Sub ConfLstStatus()
    With lstStatus
        .Clear
        .AddItem "L - Liberado  (Para Produção)"
        .AddItem "B - Bloqueado (Incluso)"
        .AddItem "N - Liberado  (Comercial)"
        .AddItem "R - Reprovado"
        .AddItem "F - Faturado  (Total)"
        .AddItem "P - Faturado  (Parcial)"
        .AddItem "M - Faturado  (Liquidado Manualmente)"
        .AddItem "S - Bloqueado"
        .AddItem "V - Aguardando Liberção de Artes"
        .AddItem "X - Para Estoque"
        .AddItem "C - Bloqueado P/Cota"
        .AddItem "D - Bloqueado P/Data"
    End With
End Sub

Private Function ConsisteCampos() As Boolean

    Dim strDATAI As String
    Dim strDATAF As String
    
    ConsisteCampos = True
    
    strDATAI = Replace(Replace(mskDataI.Text, "/", ""), "_", "")
    strDATAF = Replace(Replace(mskDataF.Text, "/", ""), "_", "")
    
    If (Len(Trim(strDATAI)) = 0 And Len(Trim(strDATAF)) = 0) Then Exit Function
    
    
    If (Len(Trim(strDATAI)) = 0 And Len(Trim(strDATAF)) > 0) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data inicial não pode ser vázia !!!"
        ConsisteCampos = False
        Exit Function
    End If
    If (Len(Trim(strDATAI)) > 0 And Len(Trim(strDATAF)) = 0) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data final não pode ser vázia !!!"
        ConsisteCampos = False
        Exit Function
    End If
    
    If Not IsDate(mskDataI.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data inicial inválida !!!"
        mskDataI.SetFocus
        ConsisteCampos = False
        Exit Function
    End If
    If Not IsDate(mskDataF.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data final inválida !!!"
        mskDataF.SetFocus
        ConsisteCampos = False
        Exit Function
    End If
    
    
    If CDate(mskDataI.Text) > CDate(mskDataF.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data inicial não pode ser maior que data final !!!"
        mskDataI.SetFocus
        ConsisteCampos = False
        Exit Function
    End If
    
    
    
    
End Function

Private Function PermiteEVendedor() As Boolean

    PermiteEVendedor = False
    
    If lngCodUsuaro = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_PVCLIE" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO     = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_EVENDEDOR  = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteEVendedor = True
    BREC.Close

End Function


Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, txtGeral As TextBox, strFUNCAOPAI As String)

On Error GoTo Err_PegaDescTabelas

    txtGeral.Text = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(Replace(strCODIGO, ",", ""))
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       txtGeral.Text = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
    Exit Sub
    
Err_PegaDescTabelas:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDescTabelas()" & vbCrLf & "Campo Nome : " & txtGeral.Name & vbCrLf & "Função Pai : " & strFUNCAOPAI, Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub


Private Function ConfereCliente(strCODIGO As String, strCODVEND As String) As Boolean

On Error GoTo Err_PegaDescTabelas

    If BREC10.State = 1 Then BREC10.Close
    
    ConfereCliente = False
    
    If lngCodUsuaro = 0 Then
       ConfereCliente = True
       Exit Function
    End If
    
    Dim boolDadosInv        As Boolean
    Dim boolPermitePesqClie As Boolean
    
    If Len(Trim(strCODVEND)) > 0 Then
    
    
        If Len(Trim(strCODIGO)) = 0 Then Exit Function
    
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       CLIE.*" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADCLIEVEND CVEN" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       CVEN.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And CVEN.SGI_CODIGO = " & Trim(strCODVEND) & vbCrLf
        sSql = sSql & "   And CVEN.SGI_CODCLI = " & strCODIGO & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL = CVEN.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO = CVEN.SGI_CODCLI"
    
    Else
    
        If Len(Trim(strCODIGO)) = 0 Then Exit Function
    
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       CLIE.*" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADCLIENTE  CLIE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       CLIE.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO = " & strCODIGO
    
    End If
    
    boolDadosInv = True
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC10.EOF() Then
       MsgBox "Este Cliente não pertence a este vendedor !!!", vbOKOnly + vbExclamation, "Aviso"
       boolDadosInv = False
    End If
    BREC10.Close
    
    If boolDadosInv = False Then Exit Function
    
    ConfereCliente = True
    
    Exit Function
    
Err_PegaDescTabelas:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Function

Private Sub SelecionaStatus(intTAB As Integer)
    
    Dim i As Integer
    
    With lstStatus
        
        fraStatus.Visible = False
        .Enabled = False
        Call ConfLstStatus
        
        If intTAB <> 8 Then Exit Sub
               
        For i = 0 To (.ListCount - 1)
            .Selected(i) = False
        Next i
        
        If intTAB = 0 Then .Selected(0) = True
        
        If intTAB = 1 Then
           .Selected(1) = True
           .Selected(2) = True
        End If
        
        If intTAB = 2 Then .Selected(3) = True
    
        If intTAB = 3 Then
           .Selected(4) = True
           .Selected(5) = True
           .Selected(6) = True
        End If
    
        If intTAB = 4 Then .Selected(7) = True
        If intTAB = 5 Then .Selected(8) = True
        If intTAB = 6 Then .Selected(9) = True
    
        If intTAB = 7 Then
            .Selected(10) = True
            .Selected(11) = True
        End If
    
        If intTAB = 8 Then
            fraStatus.Visible = True
            .Enabled = True
            .Selected(0) = True
            .Selected(1) = True
            .Selected(2) = True
        End If
    
    End With
End Sub

Private Function ConfCamposOrdem() As String

    ConfCamposOrdem = ""
    
    If lstOrdem.ListCount = 0 Then Exit Function
    
    Dim strCAMPOOrd         As String
    Dim i                   As Integer
    Dim intTABINDEX         As Integer
    Dim vControl
    
    ConfCamposOrdem = "Order By" & vbCrLf
    
    For i = 0 To (lstOrdem.ListCount - 1)
        For Each vControl In Me.Controls
            If TypeOf vControl Is TextBox Then
            ElseIf TypeOf vControl Is OptionButton Then
            ElseIf TypeOf vControl Is ComboBox Then
            ElseIf TypeOf vControl Is ListBox Then
            ElseIf TypeOf vControl Is Label Then
                If vControl.TabIndex = lstOrdem.ItemData(i) Then
                    ConfCamposOrdem = ConfCamposOrdem & Trim(vControl.Tag)
                    If i < (lstOrdem.ListCount - 1) Then ConfCamposOrdem = ConfCamposOrdem & ","
                End If
            ElseIf TypeOf vControl Is Frame Then
                If vControl.TabIndex = lstOrdem.ItemData(i) Then
                    ConfCamposOrdem = ConfCamposOrdem & Trim(vControl.Tag)
                    If i < (lstOrdem.ListCount - 1) Then ConfCamposOrdem = ConfCamposOrdem & ","
                End If
            End If
            
        Next
    Next i

End Function

Private Sub OrdemPadrao()
    lstOrdem.Clear
    lstOrdem.AddItem Frame5.Caption
    lstOrdem.ItemData(lstOrdem.NewIndex) = Frame5.TabIndex
End Sub

Private Function Verif_Linha_Sel(msGRGEN As VSFlexGrid, strACAO As String) As Boolean
    Verif_Linha_Sel = False
    
    If strACAO = "I" Then
        Verif_Linha_Sel = True
        Exit Function
    End If
    
    If msGRGEN.RowSel = 0 Then
       MsgBox "ATENÇÃO" & vbCrLf & _
              "Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
              Exit Function
    End If
    
    Verif_Linha_Sel = True
End Function

Private Sub LimpaLabelQtdeRelac()
    Dim i As Integer
    For i = 0 To (lblRegsRelac.Count - 1)
        lblRegsRelac(i).Caption = ""
    Next i
End Sub

Private Sub ContaItensRelacionados(grdGEN As VSFlexGrid, lngINDLABEL As Integer)
    lblRegsRelac(lngINDLABEL).Caption = (grdGEN.Rows - 1)
End Sub

Private Function PegaVendedoresConjulgados(strCODVEND As String) As String

    If BREC10.State = 1 Then BREC10.Close
    
    PegaVendedoresConjulgados = ""
    
    Dim strSQL As String
    
    strSQL = ""
    
    strSQL = "Select" & vbCrLf
    strSQL = strSQL & "      *" & vbCrLf
    
    strSQL = strSQL & "  From" & vbCrLf
    strSQL = strSQL & "      SGI_VENDTOVEND" & vbCrLf
    
    strSQL = strSQL & " Where" & vbCrLf
    strSQL = strSQL & "       SGI_FILIAL = " & FILIAL & vbCrLf
    strSQL = strSQL & "   And SGI_CODIGO = " & Trim(strCODVEND)
    
    BREC10.Open strSQL, adoBanco_Dados, adOpenDynamic
    Do While Not BREC10.EOF()
        PegaVendedoresConjulgados = PegaVendedoresConjulgados & Trim(Str(BREC10!SGI_CODVEND))
        BREC10.MoveNext
        If Not BREC10.EOF() Then PegaVendedoresConjulgados = PegaVendedoresConjulgados & ","
    Loop
    
End Function

Private Function PermiteConsultarOutroVendedor() As Boolean

    PermiteConsultarOutroVendedor = False
    
    If lngCodUsuaro = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_PVCLIE" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO     = " & lngCodUsuaro & vbCrLf
    sSql = sSql & "   And SGI_PVCLIE     = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteConsultarOutroVendedor = True
    BREC.Close

End Function

Private Function ConfereVendedor(strCODVEND As String, strCodDig As String) As Boolean

    ConfereVendedor = False
    
    If BREC11.State = 1 Then BREC11.Close

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    
    sSql = sSql & "  From"
    sSql = sSql & "       SGI_VENDTOVEND" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO  = " & Trim(strCODVEND) & vbCrLf
    sSql = sSql & "   And SGI_CODVEND = " & Trim(strCodDig)
    
    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then ConfereVendedor = True
    BREC11.Close
    
End Function

Private Sub FechaTableSpace()
    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    If BREC3.State = 1 Then BREC3.Close
    If BREC4.State = 1 Then BREC4.Close
    If BREC5.State = 1 Then BREC5.Close
    If BREC6.State = 1 Then BREC6.Close
    If BREC7.State = 1 Then BREC7.Close
    If BREC8.State = 1 Then BREC8.Close
    If BREC9.State = 1 Then BREC9.Close
    If BREC10.State = 1 Then BREC10.Close
    If BREC11.State = 1 Then BREC11.Close
    If BREC12.State = 1 Then BREC12.Close
End Sub
