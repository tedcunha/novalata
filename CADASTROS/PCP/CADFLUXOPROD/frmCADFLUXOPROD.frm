VERSION 5.00
Object = "{69ECBBD3-5C2A-4A84-ABEC-23937DBF1B54}#1.4#0"; "FlowChartPro.dll"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADFLUXOPROD 
   Caption         =   "Cadastro de Fluxo Produtivo"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ForeColor       =   &H8000000D&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stSenarios 
      Height          =   8775
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15478
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "Processos"
      TabPicture(0)   =   "frmCADFLUXOPROD.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdPIOR"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grdRESUMO"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraFulxoProd"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Timer1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Dados das Maquinas"
      TabPicture(1)   =   "frmCADFLUXOPROD.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Dados de Produção"
      TabPicture(2)   =   "frmCADFLUXOPROD.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   38
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
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
         TabCaption(0)   =   "Gargalos"
         TabPicture(0)   =   "frmCADFLUXOPROD.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdRitimo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "VSFlexGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VSFlex8LCtl.VSFlexGrid VSFlexGrid1 
            Height          =   1695
            Left            =   120
            TabIndex        =   54
            Top             =   2400
            Width           =   14535
            _cx             =   25638
            _cy             =   2990
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
         Begin VSFlex8LCtl.VSFlexGrid grdRitimo 
            Height          =   1935
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   14535
            _cx             =   25638
            _cy             =   3413
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
      Begin VB.Timer Timer1 
         Interval        =   7200
         Left            =   4800
         Top             =   8040
      End
      Begin VB.Frame Frame5 
         Caption         =   "[ Carteira de Máquinas do Processo ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -74880
         TabIndex        =   30
         Top             =   1680
         Width           =   14775
         Begin VB.TextBox txtDemanda 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6000
            TabIndex        =   41
            Text            =   "txtDemanda"
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtQtde 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   11400
            TabIndex        =   37
            Text            =   "txtQtde"
            Top             =   240
            Width           =   1815
         End
         Begin MSMask.MaskEdBox mskDtInicial 
            Height          =   285
            Left            =   2280
            TabIndex        =   34
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   4335
            Left            =   120
            TabIndex        =   31
            Top             =   2520
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   7646
            _Version        =   393216
            Style           =   1
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
            TabCaption(0)   =   "Scenário Otimista"
            TabPicture(0)   =   "frmCADFLUXOPROD.frx":0070
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame6"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame7"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Frame12"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Scenário Ideal"
            TabPicture(1)   =   "frmCADFLUXOPROD.frx":008C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame13"
            Tab(1).Control(1)=   "Frame10"
            Tab(1).Control(2)=   "Frame8"
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Scenário Pessimista"
            TabPicture(2)   =   "frmCADFLUXOPROD.frx":00A8
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame14"
            Tab(2).Control(1)=   "Frame11"
            Tab(2).Control(2)=   "Frame9"
            Tab(2).ControlCount=   3
            Begin VB.Frame Frame14 
               Caption         =   "[ Produção Diária da Maquina ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1695
               Left            =   -67680
               TabIndex        =   59
               Top             =   2520
               Width           =   7335
               Begin VSFlex8LCtl.VSFlexGrid grdProdMaqDiaPess 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   60
                  Top             =   240
                  Width           =   7095
                  _cx             =   12515
                  _cy             =   2355
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
            Begin VB.Frame Frame13 
               Caption         =   "[ Produção Diária da Maquina ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1695
               Left            =   -67680
               TabIndex        =   57
               Top             =   2520
               Width           =   7095
               Begin VSFlex8LCtl.VSFlexGrid grdProdMaqDiaIdeal 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   58
                  Top             =   240
                  Width           =   6855
                  _cx             =   12091
                  _cy             =   2355
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
            Begin VB.Frame Frame12 
               Caption         =   "[ Produção Diária da Maquina ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1695
               Left            =   7920
               TabIndex        =   55
               Top             =   2520
               Width           =   6495
               Begin VSFlex8LCtl.VSFlexGrid grdProdMaqDiaOtim 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   56
                  Top             =   240
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   2355
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
            Begin VB.Frame Frame11 
               Caption         =   "[ Capacidade Produtiva Diária do Processo ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2175
               Left            =   -67680
               TabIndex        =   52
               Top             =   360
               Width           =   7335
               Begin VSFlex8LCtl.VSFlexGrid grdPRODDIAPESS 
                  Height          =   1815
                  Left            =   120
                  TabIndex        =   53
                  Top             =   240
                  Width           =   7095
                  _cx             =   12515
                  _cy             =   3201
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
            Begin VB.Frame Frame10 
               Caption         =   "[ Capacidade Produtiva Diária do Processo ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2175
               Left            =   -67680
               TabIndex        =   50
               Top             =   360
               Width           =   7095
               Begin VSFlex8LCtl.VSFlexGrid grdPRODDIAIDEAL 
                  Height          =   1815
                  Left            =   120
                  TabIndex        =   51
                  Top             =   240
                  Width           =   6855
                  _cx             =   12091
                  _cy             =   3201
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
            Begin VB.Frame Frame9 
               Caption         =   "[ Velocidade de Produção - Pc/Hr ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3855
               Left            =   -74880
               TabIndex        =   48
               Top             =   360
               Width           =   7095
               Begin VSFlex8LCtl.VSFlexGrid grdCARTMAQPIOR 
                  Height          =   3495
                  Left            =   120
                  TabIndex        =   49
                  Top             =   240
                  Width           =   6855
                  _cx             =   12091
                  _cy             =   6165
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
            Begin VB.Frame Frame8 
               Caption         =   "[ Velocidade de Produção - Pc/Hr ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3855
               Left            =   -74880
               TabIndex        =   46
               Top             =   360
               Width           =   7095
               Begin VSFlex8LCtl.VSFlexGrid grdIdeal 
                  Height          =   3495
                  Left            =   120
                  TabIndex        =   47
                  Top             =   240
                  Width           =   6855
                  _cx             =   12091
                  _cy             =   6165
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
            Begin VB.Frame Frame7 
               Caption         =   "[ Capacidade Produtiva Diária do Processo ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2175
               Left            =   7920
               TabIndex        =   44
               Top             =   360
               Width           =   6495
               Begin VSFlex8LCtl.VSFlexGrid grdPRODDIAOTIM 
                  Height          =   1815
                  Left            =   120
                  TabIndex        =   45
                  Top             =   240
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   3201
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
               Caption         =   "[ Velocidade de Produção - Pc/Hr ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3855
               Left            =   120
               TabIndex        =   42
               Top             =   360
               Width           =   7695
               Begin VSFlex8LCtl.VSFlexGrid grdCARTMAQMELHOR 
                  Height          =   3495
                  Left            =   120
                  TabIndex        =   43
                  Top             =   240
                  Width           =   7335
                  _cx             =   12938
                  _cy             =   6165
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
         End
         Begin VSFlex8LCtl.VSFlexGrid grdPROCESSOS 
            Height          =   1815
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   14535
            _cx             =   25638
            _cy             =   3201
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
         Begin MSMask.MaskEdBox mskDtFinal 
            Height          =   285
            Left            =   3600
            TabIndex        =   35
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            Caption         =   "Demanda"
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
            Left            =   5040
            TabIndex        =   40
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Valocidade de Produção por Hora"
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
            Left            =   8280
            TabIndex        =   36
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label4 
            Caption         =   "Periodo"
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
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   14775
         Begin VSFlex8LCtl.VSFlexGrid grdMelhorPiorMedio 
            Height          =   975
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   14535
            _cx             =   25638
            _cy             =   1720
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
      Begin VB.Frame fraFulxoProd 
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   7800
         Width           =   2175
         Begin VB.PictureBox pb1 
            DragIcon        =   "frmCADFLUXOPROD.frx":00C4
            Height          =   550
            Left            =   120
            Picture         =   "frmCADFLUXOPROD.frx":098E
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   18
            Top             =   180
            Width           =   550
         End
         Begin VB.Label Label1 
            Caption         =   "Processo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "A&justa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   16680
         Picture         =   "frmCADFLUXOPROD.frx":1658
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7800
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "[ Legenda ]"
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
         Height          =   855
         Left            =   2400
         TabIndex        =   11
         Top             =   7800
         Width           =   2175
         Begin VB.Label lblEntrada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblSaida 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Entrada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Saida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   12
            Top             =   480
            Width           =   1455
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid grdRESUMO 
         Height          =   255
         Left            =   5400
         TabIndex        =   21
         Top             =   7920
         Visible         =   0   'False
         Width           =   255
         _cx             =   450
         _cy             =   450
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
      Begin VSFlex8LCtl.VSFlexGrid grdPIOR 
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   8280
         Visible         =   0   'False
         Width           =   255
         _cx             =   450
         _cy             =   450
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
      Begin FLOWCHARTLibCtl.FlowChart fc 
         Height          =   7335
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   14775
         _cx             =   26061
         _cy             =   12938
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BoxFillColor    =   16777215
         BoxFrameColor   =   0
         ArrowColor      =   16711680
         AlignToGrid     =   -1  'True
         ShowGrid        =   0   'False
         GridColor       =   9866380
         GridSize        =   16
         BoxStyle        =   2
         BackColor       =   13150890
         ShowShadows     =   -1  'True
         ShadowColor     =   9203310
         TextStyle       =   37
         PicturePos      =   0
         TextColor       =   0
         ArrowStyle      =   1
         ArrowSegments   =   1
         StaticMode      =   0   'False
         ScrollX         =   0
         ScrollY         =   0
         ZoomFactor      =   100
         ShowToolTips    =   0   'False
         ToolTipStyle    =   0
         AllowRefLinks   =   -1  'True
         PenStyle        =   0
         PenWidth        =   1
         Behavior        =   0
         ArrowHead       =   4
         SelectAfterCreate=   -1  'True
         BoxCustomDraw   =   0
         ShadowOffsetX   =   3
         ShadowOffsetY   =   3
         RestrObjsToDoc  =   1
         DynamicArrows   =   0   'False
         AutoScroll      =   -1  'True
         TableFillColor  =   10526900
         TableFrameColor =   0
         TableRowsCount  =   4
         TableColumnsCount=   2
         MultiSelStyle   =   0
         ModificationStart=   0
         TableColWidth   =   50
         TableRowHeight  =   22
         TableCaptionHeight=   22
         TableCaption    =   "Table"
         PrpArrowStartOrnt=   0
         TableCellBorders=   2
         UndoDepth       =   0
         SelectAfterPaste=   -1  'True
         DragDropMode    =   3
         KbdActive       =   -1  'True
         ActiveMnpColor  =   16777215
         SelMnpColor     =   11184810
         DisabledMnpColor=   200
         BoxSelStyle     =   1
         TableSelStyle   =   2
         ArrowHeadSize   =   14
         ShowDisabledHandles=   -1  'True
         ExpandOnIncoming=   0   'False
         BoxesExpandable =   -1  'True
         TablesScrollable=   0   'False
         RecursiveExpand =   -1  'True
         ShadowsStyle    =   0
         ArrowEndsMovable=   0   'False
         BoxIncmAnchor   =   0
         BoxOutgAnchor   =   0
         ArrowBase       =   0
         ArrowBaseSize   =   20
         IntermArrowHead =   0
         IntermHeadSize  =   12
         ExtrnDragDrop   =   0   'False
         SnapStyle       =   0
         SnapDistance    =   20
         BoxFillStyle    =   0
         BoxFillColor2   =   16768220
         ArrowFillColor  =   12632064
         AutoSizeDoc     =   0   'False
         FeedbackColor   =   25800
         FeedbackPenStyle=   0
         FeedbackPenWidth=   3
         LayoutGap       =   14
         LayoutStyle     =   0
         FeedbackOnDragOver=   -1  'True
         InplaceEditAllowed=   0   'False
         ShowFocusFrame  =   0   'False
         ExpandBtnPos    =   0
         ArrowsSplittable=   0   'False
         KbdBehavior     =   0
         FireMouseMove   =   0   'False
         AxControlId     =   ""
         AllowMultiSel   =   -1  'True
         AllowLinksRepeat=   -1  'True
         GridStyle       =   0
         ShapeOrientation=   0
         ShowAnchors     =   2
         SnapToAnchor    =   0
         IconTextWidth   =   100
         ArrowCrossings  =   0
         CrossRadius     =   8
         TableLinkStyle  =   0
         RouteArrows     =   0   'False
         IconTextHeight  =   50
         HighSpeedRouting=   0   'False
         HostedAxActivation=   0
         EnableStyledText=   0   'False
         AllowUnconnectedArrows=   0   'False
         ScrollRate      =   1
         BoxFillColorAlpha=   255
         ArrowFillColorAlpha=   255
         TableFillColorAlpha=   255
         ShadowColorAlpha=   150
         TableStyle      =   0
         RerouteArrows   =   1
         RoutingGridSize =   16
         MinimizeRouteSegments=   0   'False
         ArrowText       =   ""
         BoxText         =   ""
         ArrowsSnapToNodeBorders=   0   'False
         ArrowTextStyle  =   0
         ScrollZoneSize  =   0
         HitTestPriority =   1
         ArrowSelStyle   =   1
         BoxWindowFrame  =   2
         AllowUnanchoredArrows=   -1  'True
         MeasureUnit     =   2
         SelHandleSize   =   9
         SelectionOnTop  =   -1  'True
         ToolTipDelay    =   500
         BoxPicturePos   =   2
         MergeThreshold  =   0
         ControlPadding  =   4
         MiddleButtonAction=   0
         RoundedArrows   =   0   'False
         RoundedArrowsRadius=   40
         DisableNoScroll =   -1  'True
      End
   End
   Begin VB.Frame fraProcesso 
      Caption         =   "[ Processo ]"
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
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   15015
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   12360
         TabIndex        =   24
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton optSimNao 
            Caption         =   "Sim"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optSimNao 
            Caption         =   "Não"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   25
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   4440
         Picture         =   "frmCADFLUXOPROD.frx":1962
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCODPROD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Text            =   "txtCODPROD"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblDescProd 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescProd"
         Height          =   285
         Left            =   4800
         TabIndex        =   61
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Padrão"
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
         Height          =   195
         Index           =   3
         Left            =   11520
         TabIndex        =   23
         Top             =   240
         Width           =   615
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   660
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   14160
         Picture         =   "frmCADFLUXOPROD.frx":1A64
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
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
         Left            =   960
         Picture         =   "frmCADFLUXOPROD.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADFLUXOPROD.frx":2098
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADFLUXOPROD.frx":219A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Menu mnuNode 
      Caption         =   "Node menu"
      Visible         =   0   'False
      Begin VB.Menu miAddChild 
         Caption         =   "Add a child node"
      End
   End
End
Attribute VB_Name = "frmCADFLUXOPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho     As String
Public Linha        As Variant
Public cTipOper     As String
Public iCodigo      As Integer
Public FILIAL       As Integer
Public strAcesso    As String

Dim objBLBFunc          As Object
Dim objCADFLUXOPROD     As Object
Dim objPESQPADRAO       As Object

Dim currNode            As box
Dim root                As box
Dim company             As Group
Dim uniqueID            As Long
Dim newEntrada          As FLOWCHARTLibCtl.box
Dim newSaida            As FLOWCHARTLibCtl.box

Dim strMAQNESSOTIM          As String
Dim strMAQNESSOTIMGERAL     As String

Dim strMAQNESSPESS          As String
Dim strMAQNESSPESSGERAL     As String

Dim strMAQNESSIDEAL         As String
Dim strMAQNESSIDEALGERAL    As String


Const conCOL_SonRes_CodProc                         As Integer = 0
Const conCOL_SonRes_CodProd                         As Integer = 1
Const conCOL_SonRes_CodMaq                          As Integer = 2
Const conCOL_SonRes_TempoProd                       As Integer = 3
Const conCOL_SonRes_Indice                          As Integer = 4
Const conCOL_strSonRes_FormatString                 As String = "=Cod. Proc|Cod. Prod.|Cod. Maquina|Tempo Prod. Min.|Indice"
Const conColumnsIn_SonRes                           As Integer = 5

Const conCOL_SonPior_CodProc                        As Integer = 0
Const conCOL_SonPior_CodProd                        As Integer = 1
Const conCOL_SonPior_CodMaq                         As Integer = 2
Const conCOL_SonPior_TempoProd                      As Integer = 3
Const conCOL_SonPior_Indice                         As Integer = 4
Const conCOL_strSonPior_FormatString                As String = "=Cod. Proc|Cod. Prod.|Cod. Maquina|Tempo Prod. Min.|Indice"
Const conColumnsIn_SonPior                          As Integer = 5

Const conCOL_SonPMM_Descricao                       As Integer = 0
Const conCOL_SonPMM_Valor                           As Integer = 1
Const conCOL_strSonPMM_FormatString                 As String = "=Scenários|Valor"
Const conColumnsIn_SonPMM                           As Integer = 2

Const conCOL_SonProcessos_Processo                  As Integer = 0
Const conCOL_SonProcessos_CapPcHora                 As Integer = 1
Const conCOL_SonProcessos_QtdeMaq                   As Integer = 2
Const conCOL_SonProcessos_QtdePcsOtim               As Integer = 3
Const conCOL_SonProcessos_QtdeMaqOtim               As Integer = 4
Const conCOL_SonProcessos_DtPrevOtim                As Integer = 5

Const conCOL_SonProcessos_QtdePcsIdeal              As Integer = 6
Const conCOL_SonProcessos_QtdeMaqIdeal              As Integer = 7
Const conCOL_SonProcessos_DtPrevIdeal               As Integer = 8

Const conCOL_SonProcessos_QtdePcsPess               As Integer = 9
Const conCOL_SonProcessos_QtdeMaqPess               As Integer = 10
Const conCOL_SonProcessos_DtPrevPess                As Integer = 11

Const conCOL_SonProcessos_Indice                    As Integer = 12
Const conCOL_strSonProcessos_FormatString           As String = "=Processos|Cap. Pc. Hora|Qtde. Máq.|Qtde. Pcs.|Qtde. Maq.|Dt. Prv.|Qtde. Pcs.|Qtde. Maq.|Dt. Prv.|Qtde. Pcs.|Qtde. Maq.|Dt. Prv.|Indice"
Const conColumnsIn_SonProcessos                     As Integer = 13

Const conCOL_SonCARMAQMELHSEN_CodMaq                As Integer = 0
Const conCOL_SonCARMAQMELHSEN_EficMed               As Integer = 1
Const conCOL_SonCARMAQMELHSEN_ProdTeor              As Integer = 2
Const conCOL_SonCARMAQMELHSEN_ProdReal              As Integer = 3
Const conCOL_SonCARMAQMELHSEN_Soma                  As Integer = 4
Const conCOL_SonCARMAQMELHSEN_Pai                   As Integer = 5
Const conCOL_SonCARMAQMELHSEN_Indice                As Integer = 6
Const conCOL_strSonCARMAQMELHSEN_FormatString       As String = "=Código Máquina|Eficiência Média %|Produção Teórica|Prod. Real Pc/Hr|Soma/Peça|Pai|Indice"
Const conColumnsIn_SonCARMAQMELHSEN                 As Integer = 7

Const conCOL_SonIdeal_CodMaq                        As Integer = 0
Const conCOL_SonIdeal_EficMed                       As Integer = 1
Const conCOL_SonIdeal_ProdTeor                      As Integer = 2
Const conCOL_SonIdeal_ProdReal                      As Integer = 3
Const conCOL_SonIdeal_Soma                          As Integer = 4
Const conCOL_SonIdeal_Pai                           As Integer = 5
Const conCOL_SonIdeal_Indice                        As Integer = 6
Const conCOL_strSonIdeal_FormatString               As String = "=Código Máquina|Eficiência Média %|Produção Teórica|Prod. Real Pc/Hr|Soma/Peça|Pai|Indice"
Const conColumnsIn_SonIdeal                         As Integer = 5

Const conCOL_SonCARMAQPIORSEN_CodMaq                As Integer = 0
Const conCOL_SonCARMAQPIORSEN_EficMed               As Integer = 1
Const conCOL_SonCARMAQPIORSEN_ProdTeor              As Integer = 2
Const conCOL_SonCARMAQPIORSEN_ProdReal              As Integer = 3
Const conCOL_SonCARMAQPIORSEN_Soma                  As Integer = 4
Const conCOL_SonCARMAQPIORSEN_Pai                   As Integer = 5
Const conCOL_SonCARMAQPIORSEN_Indice                As Integer = 6
Const conCOL_strSonCARMAQPIORSEN_FormatString       As String = "=Código Máquina|Eficiência Média %|Produção Teórica|Prod. Real Pc/Hr|Soma/Peça|Pai|Indice"
Const conColumnsIn_SonCARMAQPIORSEN                 As Integer = 7

Const conCOL_SonCampo1                              As Integer = 0
Const conCOL_SonCampo2                              As Integer = 1
Const conCOL_strSonCampo1_FormatString              As String = "=Processos|Peças/Hora"
Const conColumnsIn_SonCampo1                        As Integer = 2

Const conCOL_SonProdDiaOtim_Data                    As Integer = 0
Const conCOL_SonProdDiaOtim_DiaSem                  As Integer = 1
Const conCOL_SonProdDiaOtim_HorDisp                 As Integer = 2
Const conCOL_SonProdDiaOtim_TotProdHora             As Integer = 3
Const conCOL_SonProdDiaOtim_TotProdTurn             As Integer = 4
Const conCOL_SonProdDiaOtim_Saldo                   As Integer = 5
Const conCOL_SonProdDiaOtim_Pai                     As Integer = 6
Const conCOL_strSonProdDiaOtim_FormatString         As String = "=Data|Dia da Semana|Hr. Disp.|Prod. Hora|Prod. Turno|Saldo|Pai"
Const conColumnsIn_SonProdDiaOtim                   As Integer = 3

Const conCOL_SonProdMaqDiaOtim_Data                    As Integer = 0
Const conCOL_SonProdMaqDiaOtim_DiaSem                  As Integer = 1
Const conCOL_SonProdMaqDiaOtim_HorDisp                 As Integer = 2
Const conCOL_SonProdMaqDiaOtim_TotProdHora             As Integer = 3
Const conCOL_SonProdMaqDiaOtim_TotProdTurn             As Integer = 4
Const conCOL_SonProdMaqDiaOtim_Saldo                   As Integer = 5
Const conCOL_SonProdMaqDiaOtim_Pai                     As Integer = 6
Const conCOL_strSonProdMaqDiaOtim_FormatString         As String = "=Data|Dia da Semana|Hr. Disp.|Prod. Hora|Prod. Turno|Saldo|Pai"
Const conColumnsIn_SonProdMaqDiaOtim                   As Integer = 3

Const conCOL_SonProdDiaIdeal_Data                    As Integer = 0
Const conCOL_SonProdDiaIdeal_DiaSem                  As Integer = 1
Const conCOL_SonProdDiaIdeal_HorDisp                 As Integer = 2
Const conCOL_SonProdDiaIdeal_TotProdHora             As Integer = 3
Const conCOL_SonProdDiaIdeal_TotProdTurn             As Integer = 4
Const conCOL_SonProdDiaIdeal_Saldo                   As Integer = 5
Const conCOL_SonProdDiaIdeal_Pai                     As Integer = 6
Const conCOL_strSonProdDiaIdeal_FormatString         As String = "=Data|Dia da Semana|Hr. Disp.|Prod. Hora|Prod. Turno|Saldo|Pai"
Const conColumnsIn_SonProdDiaIdeal                   As Integer = 3

Const conCOL_SonProdMaqDiaIdeal_Data                    As Integer = 0
Const conCOL_SonProdMaqDiaIdeal_DiaSem                  As Integer = 1
Const conCOL_SonProdMaqDiaIdeal_HorDisp                 As Integer = 2
Const conCOL_SonProdMaqDiaIdeal_TotProdHora             As Integer = 3
Const conCOL_SonProdMaqDiaIdeal_TotProdTurn             As Integer = 4
Const conCOL_SonProdMaqDiaIdeal_Saldo                   As Integer = 5
Const conCOL_SonProdMaqDiaIdeal_Pai                     As Integer = 6
Const conCOL_strSonProdMaqDiaIdeal_FormatString         As String = "=Data|Dia da Semana|Hr. Disp.|Prod. Hora|Prod. Turno|Saldo|Pai"
Const conColumnsIn_SonProdMaqDiaIdeal                   As Integer = 3

Const conCOL_SonProdDiaPess_Data                    As Integer = 0
Const conCOL_SonProdDiaPess_DiaSem                  As Integer = 1
Const conCOL_SonProdDiaPess_HorDisp                 As Integer = 2
Const conCOL_SonProdDiaPess_TotProdHora             As Integer = 3
Const conCOL_SonProdDiaPess_TotProdTurn             As Integer = 4
Const conCOL_SonProdDiaPess_Saldo                   As Integer = 5
Const conCOL_SonProdDiaPess_Pai                     As Integer = 6
Const conCOL_strSonProdDiaPess_FormatString         As String = "=Data|Dia da Semana|Hr. Disp.|Prod. Hora|Prod. Turno|Saldo|Pai"
Const conColumnsIn_SonProdDiaPess                   As Integer = 3

Const conCOL_SonProdMaqDiaPess_Data                    As Integer = 0
Const conCOL_SonProdMaqDiaPess_DiaSem                  As Integer = 1
Const conCOL_SonProdMaqDiaPess_HorDisp                 As Integer = 2
Const conCOL_SonProdMaqDiaPess_TotProdHora             As Integer = 3
Const conCOL_SonProdMaqDiaPess_TotProdTurn             As Integer = 4
Const conCOL_SonProdMaqDiaPess_Saldo                   As Integer = 5
Const conCOL_SonProdMaqDiaPess_Pai                     As Integer = 6
Const conCOL_strSonProdMaqDiaPess_FormatString         As String = "=Data|Dia da Semana|Hr. Disp.|Prod. Hora|Prod. Turno|Saldo|Pai"
Const conColumnsIn_SonProdMaqDiaPess                   As Integer = 3

Private Sub layoutTree(dir As ETreeLayoutDirection)
    Dim tl As TreeLayout
    Set tl = New TreeLayout
    
    tl.root = root
    tl.Type = tltCentralized
    tl.Direction = dir
    tl.ArrowStyle = tlaStraight
    tl.NodeSpacing = 15
    tl.LevelSpacing = 30
    tl.KeepRootPos = True
    tl.reversedArrows = False
    tl.KeepGroupLayout = False
    
    fc.ArrangeDiagram tl
End Sub



Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    fraProcesso.Enabled = True
    fraFulxoProd.Enabled = True
   
    Me.Caption = "Cadastro de Fluxo Produtivo - [ ALTERAÇÃO ]"
    cTipOper = "A"

End Sub

Private Sub cmdImpressao_Click()
    fc.PreviewDiagram
End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADFLUXOPROD.CODIGO = objCADFLUXOPROD.Gera_Codigo(Me.Name)

    objCADFLUXOPROD.IDPRODUTO = Trim(txtCODPROD.Tag)
    objCADFLUXOPROD.CODPROD = Trim(txtCODPROD.Text)
    objCADFLUXOPROD.MELHOR = CCur(grdMelhorPiorMedio.Cell(flexcpText, 1, conCOL_SonPMM_Valor))
    objCADFLUXOPROD.PIOR = CCur(grdMelhorPiorMedio.Cell(flexcpText, 2, conCOL_SonPMM_Valor))
    
    objCADFLUXOPROD.CORRMELHOR = CCur(grdMelhorPiorMedio.Cell(flexcpText, 1, conCOL_SonPMM_Valor))
    objCADFLUXOPROD.CORRPIOR = CCur(grdMelhorPiorMedio.Cell(flexcpText, 2, conCOL_SonPMM_Valor))
    
    If objCADFLUXOPROD.GRAVA(cTipOper) = False Then Exit Sub
    If objCADFLUXOPROD.Atualiza(cTipOper, Str(objCADFLUXOPROD.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
          
    MsgBox "O Processo foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command8_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    
    sSql = "" & vbCrLf
    sSql = sSql & "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    ''sSql = "Select Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    ''sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    ''sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    ''sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    ''sSql = sSql & "                        (Case " & vbCrLf
    ''sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    ''sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    ''sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
    ''sSql = sSql & ",SGI_DESCRICAO " & vbCrLf
    ''sSql = sSql & "  From " & vbCrLf
    ''sSql = sSql & "         SGI_CADPRODUTO PRO " & vbCrLf
    ''sSql = sSql & " Where " & vbCrLf
    ''sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    ''sSql = sSql & "   And SGI_STATUS     = 1"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODLIN"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Cod.Linha"
    arrCAMPOS(1, 4) = "2000"
    arrCAMPOS(1, 5) = "SGI_CODLIN"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODPROD.Text = varRETORNO
       Call PegaProduto(varRETORNO)
       lblDescProd.Caption = PegaDescProd(txtCODPROD.Tag)
    End If
    txtCODPROD.SetFocus

End Sub

Private Sub fc_BoxCollapsed(ByVal box As FLOWCHARTLibCtl.IBoxItem)
    layoutTree tldLeftToRight
End Sub

Private Sub fc_BoxDblClicked(ByVal box As FLOWCHARTLibCtl.IBoxItem, ByVal button As FLOWCHARTLibCtl.EMouseButton, ByVal x As Long, ByVal Y As Long)
    If box.Text <> "E" And box.Text <> "S" Then
        If Len(Trim(txtCODPROD.Text)) = 0 Then
           MsgBox "Primero Informe o Produto !!!", vbOKOnly + vbExclamation, "Aviso"
           txtCODPROD.SetFocus
           Exit Sub
        End If
            
        frmCADPROCESSO.cCaminho = cCaminho
        frmCADPROCESSO.Linha = Linha
        frmCADPROCESSO.iCodigo = iCodigo
        frmCADPROCESSO.cTipOper = cTipOper
        frmCADPROCESSO.FILIAL = FILIAL
        frmCADPROCESSO.strAcesso = strAcesso
        frmCADPROCESSO.strPRODUTO = Trim(txtCODPROD.Text)
        frmCADPROCESSO.lngIndice = arrPROCESSO(box.Tag).intORCEM
        
        frmCADPROCESSO.Show vbModal    '' Processo
        
        box.Text = arrPROCESSO(frmCADPROCESSO.lngIndice).strDESCRI
    ElseIf box.Text = "E" Then
        If ConsisteEntrSaida(box.Tag) = False Then Exit Sub
        
        frmCADPRODENTRADA.cCaminho = cCaminho
        frmCADPRODENTRADA.Linha = Linha
        frmCADPRODENTRADA.iCodigo = iCodigo
        frmCADPRODENTRADA.cTipOper = cTipOper
        frmCADPRODENTRADA.FILIAL = FILIAL
        frmCADPRODENTRADA.strAcesso = strAcesso
        frmCADPRODENTRADA.strPRODUTO = Trim(txtCODPROD.Text)
        frmCADPRODENTRADA.strIDProduto = Trim(txtCODPROD.Tag)
        frmCADPRODENTRADA.lngIndice = arrPROCESSO(box.Tag).intORCEM
        
        frmCADPRODENTRADA.Show vbModal    '' Processo
        
        box.Text = "E"
    ElseIf box.Text = "S" Then
        If ConsisteEntrSaida(box.Tag) = False Then Exit Sub
        
        frmCADPRODSAIDA.cCaminho = cCaminho
        frmCADPRODSAIDA.Linha = Linha
        frmCADPRODSAIDA.iCodigo = iCodigo
        frmCADPRODSAIDA.cTipOper = cTipOper
        frmCADPRODSAIDA.FILIAL = FILIAL
        frmCADPRODSAIDA.strAcesso = strAcesso
        frmCADPRODSAIDA.strPRODUTO = Trim(txtCODPROD.Text)
        frmCADPRODSAIDA.lngIndice = arrPROCESSO(box.Tag).intORCEM
        
        frmCADPRODSAIDA.Show vbModal    '' Processo
        
        box.Text = "S"
    End If
    
    Call SomaTempos
    Call CalcCorre
    Call PopGrdProcesso
    Call PopGrdRitimo
    
    If (grdPROCESSOS.Rows - 1) > 0 Then
       Call InitGridCARTMAQMELHOR
       Call InitGridCARMAQPIORSEN
       Call InitGridIdeal
       
       Call PopScearios
    End If

End Sub

Private Sub fc_BoxExpanded(ByVal box As FLOWCHARTLibCtl.IBoxItem)
    layoutTree tldLeftToRight
End Sub

Private Sub fc_DragOverBoxVB(ByVal box As FLOWCHARTLibCtl.IBoxItem, ByVal dataObj As FLOWCHARTLibCtl.IVBDataObject, ByVal docX As Long, ByVal docY As Long, ByVal keyState As Long, effect As Long)
    If box.Visible Then
        effect = vbDropEffectCopy
    Else
        effect = vbDropEffectNone
    End If
End Sub

Private Sub fc_DragOverDocVB(ByVal dataObj As FLOWCHARTLibCtl.IVBDataObject, ByVal docX As Long, ByVal docY As Long, ByVal keyState As Long, effect As Long)
    effect = vbDropEffectNone
End Sub

Private Sub fc_DropInBoxVB(ByVal box As FLOWCHARTLibCtl.IBoxItem, ByVal dataObj As FLOWCHARTLibCtl.IVBDataObject, ByVal docX As Long, ByVal docY As Long, ByVal keyState As Long, effect As Long)
    
    If TemSetasSaidaInsBox(box) = True Then Exit Sub
    
    Dim newNode     As FLOWCHARTLibCtl.box
    Dim newEntrada  As FLOWCHARTLibCtl.box
    Dim newSaida    As FLOWCHARTLibCtl.box
    
    Set newNode = addChild(box)
    If newNode Is Nothing Then Exit Sub
    
    'show the node tag
    newNode.Text = ""
    newNode.TextStyle = tsCenter
    
    newNode.PicturePos = picCenterLeft
    newNode.Picture = pb1.Picture
    
    '' Adiciona Entradas
    ''Set newEntrada = addChildEntrada(newNode)
    ''Set newSaida = addChildSaida(newNode)
    '' Para a Nova Lata Não Vai precisar
        
End Sub


Private Sub fc_DropInDocVB(ByVal dataObj As FLOWCHARTLibCtl.IVBDataObject, ByVal docX As Long, ByVal docY As Long, ByVal keyState As Long, effect As Long)
    If Not fc.ActiveBox Is Nothing Then
        addChild fc.ActiveBox
    End If
End Sub

Private Sub fc_RequestDeleteArrow(ByVal arrow As FLOWCHARTLibCtl.IArrowItem, pbDelete As Boolean)
    'don't allow the user to delete arrows
    pbDelete = False
End Sub

Private Sub fc_RequestDeleteBox(ByVal box As FLOWCHARTLibCtl.IBoxItem, pbDelete As Boolean)
    'don't allow the user to delete boxes
    If box.Text = "S" Or box.Text = "E" Then
       pbDelete = False
       Exit Sub
    End If
    If box.Tag = 0 Then pbDelete = False
    If box.Tag > 0 Then pbDelete = TemSetasSaida(box)
End Sub

Private Sub fc_RequestSelectArrow(ByVal arrow As FLOWCHARTLibCtl.IArrowItem, pbSelect As Boolean)
    'don't allow users to select arrows
    pbSelect = False
End Sub

Private Sub Form_Activate()
    If (grdPROCESSOS.Rows - 1) > 0 Then
        grdPROCESSOS.Row = 1
        Call grdPROCESSOS_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    
   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADFLUXOPROD = CreateObject("CADFLUXOPROD.clsCADFLUXOPROD")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADFLUXOPROD.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'cleanup
    fc.RevokeDragDrop
    fc.Graphics.ShutDown
    Call DestroyObjeto
End Sub

Private Sub grdCARTMAQMELHOR_Click()
    With grdCARTMAQMELHOR
        If (.Rows - 1) > 0 And .Row > 0 Then
           Call PosRegGrdMaqProd(.Cell(flexcpText, .Row, conCOL_SonCARMAQMELHSEN_Indice))
        End If
    End With
End Sub

Private Sub grdCARTMAQMELHOR_RowColChange()
    With grdCARTMAQMELHOR
        If (.Rows - 1) > 0 And .Row > 0 Then
           Call PosRegGrdMaqProd(.Cell(flexcpText, .Row, conCOL_SonCARMAQMELHSEN_Indice))
        End If
    End With
End Sub

Private Sub grdCARTMAQPIOR_Click()
    With grdCARTMAQPIOR
        If (.Rows - 1) > 0 And .Row > 0 Then
           Call PosRegGrdMaqProd(.Cell(flexcpText, .Row, conCOL_SonCARMAQPIORSEN_Indice))
        End If
    End With
End Sub

Private Sub grdCARTMAQPIOR_RowColChange()
    With grdCARTMAQPIOR
        If (.Rows - 1) > 0 And .Row > 0 Then
           Call PosRegGrdMaqProd(.Cell(flexcpText, .Row, conCOL_SonCARMAQPIORSEN_Indice))
        End If
    End With
End Sub

Private Sub grdIdeal_Click()
    With grdIdeal
        If (.Rows - 1) > 0 And .Row > 0 Then
           Call PosRegGrdMaqProd(.Cell(flexcpText, .Row, conCOL_SonIdeal_Indice))
        End If
    End With
End Sub

Private Sub grdIdeal_RowColChange()
    With grdIdeal
        If (.Rows - 1) > 0 And .Row > 0 Then
           Call PosRegGrdMaqProd(.Cell(flexcpText, .Row, conCOL_SonIdeal_Indice))
        End If
    End With
End Sub

Private Sub grdPROCESSOS_Click()
    Dim intPESQ As Integer
    
    If (grdPROCESSOS.Rows - 1) > 0 And grdPROCESSOS.Row > 0 Then
       Call PosRegGrdProcessos(grdPROCESSOS.Cell(flexcpText, grdPROCESSOS.Row, conCOL_SonProcessos_Indice))
       Call PosRegGrdCarteira(grdPROCESSOS.Cell(flexcpText, grdPROCESSOS.Row, conCOL_SonProcessos_Indice))
    End If

    With grdCARTMAQMELHOR
        If (.Rows - 1) > 0 Then
           intPESQ = .FindRow(grdPROCESSOS.Cell(flexcpText, grdPROCESSOS.Row, conCOL_SonProcessos_Indice), , conCOL_SonCARMAQMELHSEN_Pai)
           If intPESQ <> -1 Then
              Call PosRegGrdMaqProd(.Cell(flexcpText, intPESQ, conCOL_SonCARMAQMELHSEN_Indice))
           End If
        End If
    End With

End Sub

Private Sub grdPROCESSOS_RowColChange()
    Dim intPESQ As Integer
    
    If (grdPROCESSOS.Rows - 1) > 0 And grdPROCESSOS.Row > 0 Then
       Call PosRegGrdProcessos(grdPROCESSOS.Cell(flexcpText, grdPROCESSOS.Row, conCOL_SonProcessos_Indice))
       Call PosRegGrdCarteira(grdPROCESSOS.Cell(flexcpText, grdPROCESSOS.Row, conCOL_SonProcessos_Indice))
    End If

    With grdCARTMAQMELHOR
        If (.Rows - 1) > 0 Then
           intPESQ = .FindRow(grdPROCESSOS.Cell(flexcpText, grdPROCESSOS.Row, conCOL_SonProcessos_Indice), , conCOL_SonCARMAQMELHSEN_Pai)
           If intPESQ <> -1 Then
              Call PosRegGrdMaqProd(.Cell(flexcpText, intPESQ, conCOL_SonCARMAQMELHSEN_Indice))
           End If
        End If
    End With
End Sub

Private Sub grdRitimo_DblClick()
    If Len(Trim(grdRitimo.Cell(flexcpText, grdRitimo.Row, conCOL_SonCampo1))) > 0 Then
       txtQtde.Text = grdRitimo.Cell(flexcpText, grdRitimo.Row, conCOL_SonCampo2)
    
       Call InitGridCARTMAQMELHOR
       Call InitGridCARMAQPIORSEN
       Call InitGridIdeal
       
       Call PopScearios
       
       Call PrevisaoOtim(strMAQNESSOTIMGERAL, Date, txtDemanda.Text)
       Call PrevisaoPess(strMAQNESSPESSGERAL, Date, txtDemanda.Text)
       Call PrevisaoIdeal(strMAQNESSIDEALGERAL, Date, txtDemanda.Text)
       
       If (grdPROCESSOS.Rows - 1) > 0 Then
           grdPROCESSOS.Row = 1
           Call grdPROCESSOS_Click
       End If
       
    End If
End Sub

Private Sub miAddChild_Click()
    addChild currNode
    Set currNode = Nothing
End Sub

Private Sub mskDtFinal_Validate(Cancel As Boolean)
    Cancel = ValidaPeriodo
End Sub

Private Sub mskDtInicial_Validate(Cancel As Boolean)
    Cancel = ValidaPeriodo
End Sub

Private Sub pb1_MouseDown(button As Integer, Shift As Integer, x As Single, Y As Single)
    pb1.OLEDrag
End Sub

Private Sub pb1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectCopy
    Data.SetData pb1.Image, vbCFDIB
End Sub

Private Function addChild(node As FLOWCHARTLibCtl.box) As box
    
    Dim child       As box
    Dim link        As arrow
    Dim division    As Group
    Dim parentDiv   As Group
    
    Set parentDiv = node.SubordinateGroup
    
    Set child = fc.CreateBox(0, 0, 150, 40)
    parentDiv.AttachToCorner child, 0
    
    Set link = fc.CreateArrow(node, child)
    
    uniqueID = (UBound(arrPROCESSO) + 1)
    child.Tag = uniqueID
    
    ReDim Preserve arrPROCESSO(0 To uniqueID) As Processos
    arrPROCESSO(uniqueID).intORCEM = child.Tag
    arrPROCESSO(uniqueID).intTipo = 0
    arrPROCESSO(uniqueID).lngCODIGO = Empty
    arrPROCESSO(uniqueID).strDESCRI = Empty
    
    Set division = fc.CreateGroup(child)
    
    layoutTree tldLeftToRight
        
    fc.ClearSelection
    fc.AddToSelection child
    
    Set addChild = child
    
End Function

Private Function TemSetasSaida(ByVal box As FLOWCHARTLibCtl.IBoxItem) As Boolean
    Dim I As Integer
    Dim intDestino As Integer
    
    Dim SetaDeSaida As FLOWCHARTLibCtl.IArrows
    Set SetaDeSaida = box.OutgoingArrows
    If SetaDeSaida.Count <= 2 Then
       TemSetasSaida = True
       
VOLTA:
       For I = 0 To SetaDeSaida.Count - 1
           fc.DeleteItem SetaDeSaida.Item(I).DestinationBox
           GoTo VOLTA
       Next I
       
       uniqueID = (UBound(arrPROCESSO) - 1)
       ReDim Preserve arrPROCESSO(0 To uniqueID) As Processos
       Call SomaTempos
       
    End If
    If SetaDeSaida.Count >= 3 Then TemSetasSaida = False
End Function

Private Function TemSetasEntrada(ByVal box As FLOWCHARTLibCtl.IBoxItem) As Boolean
    Dim SetaDeEntrada As FLOWCHARTLibCtl.IArrows
    Set SetaDeEntrada = box.IncomingArrows
    If SetaDeEntrada.Count = 0 Then TemSetasEntrada = False
End Function

Private Function addChildEntrada(node As FLOWCHARTLibCtl.box) As box
    
    Dim child       As box
    Dim link        As arrow
    Dim division    As Group
    Dim parentDiv   As Group
    
    Set parentDiv = node.SubordinateGroup
    
    Set child = fc.CreateBox(0, 0, 20, 20)
    child.FillColor = vbBlue
    parentDiv.AttachToCorner child, 0
    
    Set link = fc.CreateArrow(node, child)
    
    child.Tag = uniqueID
    child.Text = "E"
    
    Set division = fc.CreateGroup(child)
    
    layoutTree tldLeftToRight
        
    fc.ClearSelection
    fc.AddToSelection child
    
    Set addChildEntrada = child
    
End Function

Private Function addChildSaida(node As FLOWCHARTLibCtl.box) As box
    
    Dim child       As box
    Dim link        As arrow
    Dim division    As Group
    Dim parentDiv   As Group
    
    Set parentDiv = node.SubordinateGroup
    
    Set child = fc.CreateBox(0, 0, 20, 20)
    child.FillColor = vbGreen
    parentDiv.AttachToCorner child, 0
    
    Set link = fc.CreateArrow(node, child)
    
    child.Tag = uniqueID
    child.Text = "S"
    
    Set division = fc.CreateGroup(child)
    
    layoutTree tldLeftToRight
        
    fc.ClearSelection
    fc.AddToSelection child
    
    Set addChildSaida = child
    
End Function


Private Sub CriaBoxINI()

    fc.Graphics.StartUp geGdiPlus
    
    'enable drag'n'drop
    fc.RegisterDragDrop
    
    uniqueID = 0
    
    'create the root of the tree
    Set root = fc.CreateBox(100, 200, 150, 40)
    root.Tag = uniqueID
    root.Text = ""
    
    '' ==============================================================
    '' Criando Array de Estrutura de Processo
    ReDim Preserve arrPROCESSO(0 To 0) As Processos
    arrPROCESSO(0).intORCEM = root.Tag
    arrPROCESSO(0).intTipo = 0
    arrPROCESSO(0).lngCODIGO = Empty
    arrPROCESSO(0).strDESCRI = Empty
    '' ==============================================================
    
    root.PicturePos = picCenterLeft
    root.Picture = pb1.Picture
    root.TextStyle = tsCenter
    
    'create an hierarchical group, so that when
    'the root is moved all the children will move too
    Set company = fc.CreateGroup(root)
    
    
    '' Adiciona Entradas/Saidas
    ''Set newEntrada = addChildEntrada(root)
    ''Set newSaida = addChildSaida(root)
    '' Para a Novalata Não Vai Precisar
    
    'deselect the root
    fc.ClearSelection
    
    Call layoutTree(tldLeftToRight)
    
    lblEntrada.BackColor = vbGreen
    lblSaida.BackColor = vbBlue

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    fraFulxoProd.Enabled = True
   
    Me.Caption = "Cadastro de Fluxo Produtivo - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFLUXOPROD
    
    Call CriaBoxINI
    
    stSenarios.Tab = 0
    
    Call InitGridResumo
    Call InitGridPior
    Call InitGridPMM
    Call InitGridProcessos
    Call InitGridCARTMAQMELHOR
    Call InitGridCARMAQPIORSEN
    Call InitGridIdeal
    Call InitGridCampo
    Call InitGridProdDiaOtim
    Call InitGridProdDiaIdeal
    
    Call InitGridProdMaqDiaOtim
    Call InitGridProdMaqDiaIdeal
    Call InitGridProdMaqDiaPess
    
    mskDtInicial.Text = Format(Date, "DD/MM/YYYY")
    mskDtFinal.Text = Format(Date, "DD/MM/YYYY")
    
    Call LimpaCamposLabel
    Call DeabilitaCampos
    
End Sub


Private Sub Timer1_Timer()
'    If (grdPROCESSOS.Rows - 1) > 0 And grdPROCESSOS.Row > 0 Then
'       Call PopGrdMELHORSEN(grdPROCESSOS.Cell(flexcpText, grdPROCESSOS.Row, conCOL_SonProcessos_Indice))
'       Call PopGrdPIORSEN(grdPROCESSOS.Cell(flexcpText, grdPROCESSOS.Row, conCOL_SonProcessos_Indice))
'    End If
End Sub

Private Sub txtCODPROD_GotFocus()
    objBLBFunc.SelecionaCampos txtCODPROD.Name, frmCADFLUXOPROD
End Sub

Private Sub txtCODPROD_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCODPROD.Text)) = 0 Then Exit Sub
   
   Call PegaProduto(txtCODPROD.Text)
   lblDescProd.Caption = PegaDescProd(txtCODPROD.Tag)
   If Len(Trim(lblDescProd.Caption)) = 0 Then
      MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
      Cancel = True
      Exit Sub
   End If
   
End Sub

Private Function TemSetasSaidaInsBox(ByVal box As FLOWCHARTLibCtl.IBoxItem) As Boolean
    Dim SetaDeSaida As FLOWCHARTLibCtl.IArrows
    
    If Len(Trim(arrPROCESSO(box.Tag).strDESCRI)) = 0 Then
       MsgBox "Primeiro Defina o Processo para depois inserir o proximo processo !!!", vbOKOnly + vbExclamation, "Aviso"
       TemSetasSaidaInsBox = True
       Exit Function
    End If
    
    If arrPROCESSO(box.Tag).lngQTDPRODUTOS = 0 Then
       MsgBox "Primeiro Defina os Produtos do Processo para depois inserir o proximo processo !!!", vbOKOnly + vbExclamation, "Aviso"
       TemSetasSaidaInsBox = True
       Exit Function
    End If
    
    Set SetaDeSaida = box.OutgoingArrows
    If SetaDeSaida.Count > 2 Then TemSetasSaidaInsBox = True
    If SetaDeSaida.Count = 2 Then TemSetasSaidaInsBox = False
End Function

Private Sub InitGridResumo()

    With grdRESUMO
    
       .Cols = conColumnsIn_SonRes
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonRes_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonRes_CodProc) = ""
       .ColDataType(conCOL_SonRes_CodProc) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonRes_CodProd) = ""
       .ColDataType(conCOL_SonRes_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonRes_CodMaq) = ""
       .ColDataType(conCOL_SonRes_CodMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonRes_TempoProd) = ""
       .ColDataType(conCOL_SonRes_TempoProd) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonRes_Indice) = ""
       .ColDataType(conCOL_SonRes_Indice) = flexDTString
       
       .ColWidth(conCOL_SonRes_CodProc) = 2000
       .ColWidth(conCOL_SonRes_CodProd) = 2000
       .ColWidth(conCOL_SonRes_CodMaq) = 2000
       .ColWidth(conCOL_SonRes_TempoProd) = 2500
       .ColWidth(conCOL_SonRes_Indice) = 2500
       
       .ColHidden(conCOL_SonRes_Indice) = False
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub InitGridPior()

    With grdPIOR
    
       .Cols = conColumnsIn_SonPior
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonPior_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonPior_CodProc) = ""
       .ColDataType(conCOL_SonPior_CodProc) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonPior_CodProd) = ""
       .ColDataType(conCOL_SonPior_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPior_CodMaq) = ""
       .ColDataType(conCOL_SonPior_CodMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonPior_TempoProd) = ""
       .ColDataType(conCOL_SonPior_TempoProd) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonPior_Indice) = ""
       .ColDataType(conCOL_SonPior_Indice) = flexDTString
       
       .ColWidth(conCOL_SonPior_CodProc) = 2000
       .ColWidth(conCOL_SonPior_CodProd) = 2000
       .ColWidth(conCOL_SonPior_CodMaq) = 2000
       .ColWidth(conCOL_SonPior_TempoProd) = 2500
       .ColWidth(conCOL_SonPior_Indice) = 2500
       
       .ColHidden(conCOL_SonPior_Indice) = False
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub InitGridPMM()

    With grdMelhorPiorMedio
    
       .Cols = conColumnsIn_SonPMM
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonPMM_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonPMM_Descricao) = ""
       .ColDataType(conCOL_SonPMM_Descricao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPMM_Valor) = ""
       .ColDataType(conCOL_SonPMM_Valor) = flexDTCurrency
       
       .ColWidth(conCOL_SonPMM_Descricao) = 2000
       .ColWidth(conCOL_SonPMM_Valor) = 1500
       
       .RowHidden(0) = True
        
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Public Function SomaTempos()

    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim curESTINTEM As Currency
    Dim curTEMPPROD As Currency
    Dim curMELHOR   As Currency
    Dim curPIOR     As Currency
    
    Call InitGridResumo
    Call InitGridPior
    Call InitGridPMM
    
    curMELHOR = 0
    curPIOR = 0
    For I = 0 To UBound(arrPROCESSO) '' Processo
        For j = 1 To arrPROCESSO(I).lngQTDPRODUTOS
            '' ==========================================================
            '' Pegando Estoque Intemediário
            curESTINTEM = 0
            For k = 1 To arrPROCESSO(I).typProdutos(j).lngTOTPRODENTRADA
                If arrPROCESSO(I).typProdutos(j).typProdEntrada(k).intCADENCIA = 1 Then curESTINTEM = arrPROCESSO(I).typProdutos(j).typProdEntrada(k).curQTDESTOQUE
            Next k
            '' ==========================================================
            
            Call InitGridResumo
            Call InitGridPior
            For k = 1 To arrPROCESSO(I).typProdutos(j).lngTOTMAQUINAS
                If curESTINTEM > 0 Then
                   If arrPROCESSO(I).typProdutos(j).typMaquinas(k).lngQTDPCMIN > 0 Then
                      curTEMPPROD = (curESTINTEM / arrPROCESSO(I).typProdutos(j).typMaquinas(k).lngQTDPCMIN)
                      arrPROCESSO(I).typProdutos(j).typMaquinas(k).curTEMPPROD = curTEMPPROD
                   End If
                Else
                   arrPROCESSO(I).typProdutos(j).typMaquinas(k).curTEMPPROD = 0
                End If
                grdRESUMO.AddItem arrPROCESSO(I).lngCODIGO & vbTab & _
                                  arrPROCESSO(I).typProdutos(j).strPRODUTO & vbTab & _
                                  arrPROCESSO(I).typProdutos(j).typMaquinas(k).lngCODMAQ & vbTab & _
                                  Format(arrPROCESSO(I).typProdutos(j).typMaquinas(k).curTEMPPROD, "#,##0.00") & vbTab & _
                                  arrPROCESSO(I).lngCODIGO & arrPROCESSO(I).typProdutos(j).strPRODUTO
                                  
                grdPIOR.AddItem arrPROCESSO(I).lngCODIGO & vbTab & _
                                 arrPROCESSO(I).typProdutos(j).strPRODUTO & vbTab & _
                                 arrPROCESSO(I).typProdutos(j).typMaquinas(k).lngCODMAQ & vbTab & _
                                 Format(arrPROCESSO(I).typProdutos(j).typMaquinas(k).curTEMPPROD, "#,##0.00") & vbTab & _
                                 arrPROCESSO(I).lngCODIGO & arrPROCESSO(I).typProdutos(j).strPRODUTO
                                  
            Next k
            '' Melhor Cenario
            If grdRESUMO.Rows - 1 > 0 Then
               grdRESUMO.Subtotal flexSTMin, , conCOL_SonRes_TempoProd, "#,##0.00", , , , "Melhor"
               arrPROCESSO(I).typProdutos(j).curMELHORCENARIO = grdRESUMO.Cell(flexcpText, 1, conCOL_SonRes_TempoProd)
               curMELHOR = (curMELHOR + arrPROCESSO(I).typProdutos(j).curMELHORCENARIO)
            End If
            '' Pior Cenario
            If grdPIOR.Rows - 1 > 0 Then
               grdPIOR.Subtotal flexSTMax, , conCOL_SonPior_TempoProd, "#,##0.00", , , , "Pior"
               arrPROCESSO(I).typProdutos(j).curPIORCENARIO = grdPIOR.Cell(flexcpText, 1, conCOL_SonPior_TempoProd)
               curPIOR = (curPIOR + arrPROCESSO(I).typProdutos(j).curPIORCENARIO)
            End If
        
        Next j
        arrPROCESSO(I).curMELHORCENARIO = curMELHOR
        arrPROCESSO(I).curPIORCENARIO = curPIOR
    Next I
    
    Call InitGridResumo
    Call InitGridPior
    Call InitGridPMM
    
    grdMelhorPiorMedio.AddItem "Melhor Cenário " & vbTab & Format(curMELHOR, "#,##0.00")
    grdMelhorPiorMedio.AddItem "Pior Cenário" & vbTab & Format(curPIOR, "#,##0.00")
    

End Function

Private Function ConsisteEntrSaida(lngTag As Long) As Boolean

            ConsisteEntrSaida = False
            
            If Len(Trim(txtCODPROD.Text)) = 0 Then
               MsgBox "Primeiro Informe o Produtos  !!!", vbExclamation + vbOKOnly, "Aviso"
               txtCODPROD.SetFocus
               Exit Function
            End If
            If arrPROCESSO(lngTag).lngQTDPRODUTOS = 0 Then
               MsgBox "Primeiro Informe os Produtos a Serem Produzidos !!!", vbExclamation + vbOKOnly, "Aviso"
               Exit Function
            End If
            
            ConsisteEntrSaida = True

End Function


Private Function ValidaCampos() As Boolean

     Dim I As Integer
     Dim j As Integer
     ValidaCampos = False
     
     If Trim(Len(txtCODPROD.Text)) = 0 Then
        MsgBox "Informe o código do produto !!!", vbOKOnly + vbCritical, "Aviso"
        txtCODPROD.SetFocus
        Exit Function
     End If
     
     For I = 0 To (UBound(arrPROCESSO))
        If arrPROCESSO(I).lngCODIGO = 0 Then
           MsgBox "Faltam processos a serem informados !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
        If arrPROCESSO(I).lngQTDPRODUTOS = 0 Then
           MsgBox "Não foi informado produtos para o Processo : " & arrPROCESSO(I).strDESCRI & " !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
        '' Para a Novalata Não Vai usar
        '' Produtos de Entrada
        ''For j = 1 To arrPROCESSO(I).lngQTDPRODUTOS
        ''   If arrPROCESSO(I).typProdutos(j).lngTOTPRODENTRADA = 0 Then
        ''      MsgBox "Não foi informado produtos de entrada o Processo : " & arrPROCESSO(I).strDESCRI & " !!!", vbOKOnly + vbExclamation, "Aviso"
        ''      Exit Function
        ''   End If
        ''Next j
        '' Produtos de Saida
        ''For j = 1 To arrPROCESSO(I).lngQTDPRODUTOS
        ''   If arrPROCESSO(I).typProdutos(j).lngTOTPRODSAIDA = 0 Then
        ''      MsgBox "Não foi informado produtos de saida o Processo : " & arrPROCESSO(I).strDESCRI & " !!!", vbOKOnly + vbExclamation, "Aviso"
        ''      Exit Function
        ''   End If
        ''Next j
        '' Maquinas
        For j = 1 To arrPROCESSO(I).lngQTDPRODUTOS
           If arrPROCESSO(I).typProdutos(j).lngTOTMAQUINAS = 0 Then
              MsgBox "Não foi informado as máquinas para o Processo : " & arrPROCESSO(I).strDESCRI & " !!!", vbOKOnly + vbExclamation, "Aviso"
              Exit Function
           End If
        Next j
     Next I
     
     ValidaCampos = True
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    fraProcesso.Enabled = False
    fraFulxoProd.Enabled = False
   
    Me.Caption = "Cadastro de Fluxo Produtivo - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADFLUXOPROD
    
    objCADFLUXOPROD.CODIGO = iCodigo
    
    Call InitGridPior
    Call InitGridPMM
    Call InitGridResumo
    Call InitGridProcessos
    Call InitGridCARTMAQMELHOR
    Call InitGridCARMAQPIORSEN
    Call InitGridIdeal
    Call InitGridCampo
    Call InitGridProdDiaOtim
    Call InitGridProdDiaIdeal
    Call InitGridProdDiaPess
    
    
    Call InitGridProdMaqDiaOtim
    Call InitGridProdMaqDiaIdeal
    Call InitGridProdMaqDiaPess
    
    mskDtInicial.Text = Format(Date, "DD/MM/YYYY")
    mskDtFinal.Text = Format(Date, "DD/MM/YYYY")
    
    uniqueID = 0
    
    Call LimpaCamposLabel
    Call DeabilitaCampos
    
    If objCADFLUXOPROD.Carrega_campos = True Then
        
       txtCodigo.Text = objCADFLUXOPROD.CODIGO
       txtCODPROD.Text = objCADFLUXOPROD.CODPROD
       txtCODPROD.Tag = objCADFLUXOPROD.IDPRODUTO
       lblDescProd.Caption = PegaDescProd(txtCODPROD.Tag)
       
       Call CriaBoxINIConsAlt
       Call SomaTempos
       
       Call PegaProduto(txtCodigo.Text)
       Call CalcCorre
       Call PopGrdProcesso
       Call PopGrdRitimo
       
       '' Gerando escolhendo conforme o ritimo de produção
       If Len(Trim(grdRitimo.Cell(flexcpText, grdRitimo.Row, conCOL_SonCampo1))) > 0 Then
          grdRitimo.Row = 1
          grdRitimo.Cell(flexcpBackColor, grdRitimo.Row, conCOL_SonCampo1, grdRitimo.Row, conCOL_SonCampo2) = vbRed
          
          txtQtde.Text = grdRitimo.Cell(flexcpText, grdRitimo.Row, conCOL_SonCampo2)
          Call PopScearios
          
       End If
    
    End If

End Sub

Private Sub CriaBoxINIConsAlt()

       Dim newNode     As FLOWCHARTLibCtl.box
       Dim newEntrada  As FLOWCHARTLibCtl.box
       Dim newSaida    As FLOWCHARTLibCtl.box
       
       '' ==============================================================
       fc.Graphics.StartUp geGdiPlus
        
       'enable drag'n'drop
       fc.RegisterDragDrop
        
       'create the root of the tree
       Set root = fc.CreateBox(100, 200, 150, 40)
       root.Tag = uniqueID
       root.Text = arrPROCESSO(uniqueID).strDESCRI
        
       root.PicturePos = picCenterLeft
       root.Picture = pb1.Picture
       root.TextStyle = tsCenter
        
       'create an hierarchical group, so that when
       'the root is moved all the children will move too
       Set company = fc.CreateGroup(root)
        
       '' Adiciona Entradas/Saidas
       ''Set newEntrada = addChildEntrada(root)
       ''Set newSaida = addChildSaida(root)
       '' Para a Nova Lata Nao precisa
        
       'deselect the root
       fc.ClearSelection
       
       '' Construindo o Flow
       For I = 1 To UBound(arrPROCESSO)
           
                uniqueID = (uniqueID + 1)
                
                Set newNode = addChildAltCon(root)
                If newNode Is Nothing Then Exit Sub
                
                Set root = newNode
                
                'show the node tag
                newNode.TextStyle = tsCenter
                newNode.PicturePos = picCenterLeft
                newNode.Picture = pb1.Picture
                
                '' Adiciona Entradas
                ''Set newEntrada = addChildEntrada(newNode)
                ''Set newSaida = addChildSaida(newNode)
                '' Para a Novalata nao vai precisar
           
       Next I
    
       Call layoutTree(tldLeftToRight)
    
       lblEntrada.BackColor = vbGreen
       lblSaida.BackColor = vbBlue

End Sub


Private Function addChildAltCon(node As FLOWCHARTLibCtl.box) As box
    
    Dim child       As box
    Dim link        As arrow
    Dim division    As Group
    Dim parentDiv   As Group
    
    Set parentDiv = node.SubordinateGroup
    
    Set child = fc.CreateBox(0, 0, 150, 40)
    parentDiv.AttachToCorner child, 0
    
    Set link = fc.CreateArrow(node, child)
    
    child.Tag = uniqueID
    child.Text = arrPROCESSO(uniqueID).strDESCRI
    
    arrPROCESSO(uniqueID).intORCEM = arrPROCESSO(uniqueID).intORCEM
    arrPROCESSO(uniqueID).intTipo = 0
    arrPROCESSO(uniqueID).lngCODIGO = arrPROCESSO(uniqueID).lngCODIGO
    arrPROCESSO(uniqueID).strDESCRI = arrPROCESSO(uniqueID).strDESCRI
    
    Set division = fc.CreateGroup(child)
    
    layoutTree tldLeftToRight
        
    fc.ClearSelection
    fc.AddToSelection child
    
    Set addChildAltCon = child
    
End Function

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    fraProcesso.Enabled = True
    fraFulxoProd.Enabled = True
   
    Me.Caption = "Cadastro de Fluxo Produtivo - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFLUXOPROD
    
    objCADFLUXOPROD.CODIGO = iCodigo
    
    Call InitGridPior
    Call InitGridPMM
    Call InitGridResumo
    Call InitGridProcessos
    Call InitGridCARTMAQMELHOR
    Call InitGridCARMAQPIORSEN
    Call InitGridIdeal
    Call InitGridCampo
    Call InitGridProdDiaOtim
    Call InitGridProdDiaIdeal
    Call InitGridProdDiaPess
    
    Call InitGridProdMaqDiaOtim
    Call InitGridProdMaqDiaIdeal
    Call InitGridProdMaqDiaPess
    
    mskDtInicial.Text = Format(Date, "DD/MM/YYYY")
    mskDtFinal.Text = Format(Date, "DD/MM/YYYY")
    
    uniqueID = 0
    
    Call LimpaCamposLabel
    Call DeabilitaCampos
    
    If objCADFLUXOPROD.Carrega_campos = True Then
        
       txtCodigo.Text = objCADFLUXOPROD.CODIGO
       txtCODPROD.Text = objCADFLUXOPROD.CODPROD
       txtCODPROD.Tag = objCADFLUXOPROD.IDPRODUTO
       lblDescProd.Caption = PegaDescProd(txtCODPROD.Tag)
       
       Call CriaBoxINIConsAlt
       Call SomaTempos
       
       Call PegaProduto(txtCodigo.Text)
       Call CalcCorre
       Call PopGrdProcesso
       Call PopGrdRitimo
       
       '' Gerando escolhendo conforme o ritimo de produção
       If Len(Trim(grdRitimo.Cell(flexcpText, grdRitimo.Row, conCOL_SonCampo1))) > 0 Then
          txtQtde.Text = grdRitimo.Cell(flexcpText, grdRitimo.Row, conCOL_SonCampo2)
          Call PopScearios
       End If
    
    End If

End Sub

Private Sub PegaProduto(strCODPRODUTO As String)

    optSimNao(0).Value = True
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODLIN = " & strCODPRODUTO
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF() Then
        txtCODPROD.Tag = BREC4!SGI_CODIGO
    End If
    BREC4.Close
    
End Sub

Private Sub CalcCorre()
    
    If optSimNao(1).Value = True Then Exit Sub
    
    Dim curMELHOR   As Currency
    Dim curPIOR     As Currency
    
    curMELHOR = 0
    curPIOR = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        FLX.SGI_MELHOR" & vbCrLf
    sSql = sSql & "       ,FLX.SGI_PIOR" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADPRODUTO  PRD " & vbCrLf
    sSql = sSql & "       ,SGI_CADFLUXPROD FLX " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        PRD.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "    And PRD.SGI_PADRAO  = 1" & vbCrLf
    sSql = sSql & "    And FLX.SGI_FILIAL  = PRD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And FLX.SGI_CODPROD = PRD.SGI_CODIGO " & vbCrLf
    
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       curMELHOR = (CCur(grdMelhorPiorMedio.Cell(flexcpText, 1, conCOL_SonPMM_Valor)) / BREC!SGI_MELHOR)
       curPIOR = (CCur(grdMelhorPiorMedio.Cell(flexcpText, 2, conCOL_SonPMM_Valor)) / BREC!SGI_PIOR)
    End If
    BREC.Close
    
    grdMelhorPiorMedio.AddItem "" & vbTab & ""
    
    grdMelhorPiorMedio.AddItem "Correlação Melhor" & vbTab & Format(curMELHOR, "#,##0.00")
    grdMelhorPiorMedio.AddItem "Correlação Pior" & vbTab & Format(curPIOR, "#,##0.00")
    
End Sub

Private Sub InitGridProcessos()

    With grdPROCESSOS
    
       .Cols = conColumnsIn_SonProcessos
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonProcessos_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_Processo) = ""
       .ColDataType(conCOL_SonProcessos_Processo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_CapPcHora) = ""
       .ColDataType(conCOL_SonProcessos_CapPcHora) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_QtdeMaq) = ""
       .ColDataType(conCOL_SonProcessos_QtdeMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_QtdeMaqOtim) = ""
       .ColDataType(conCOL_SonProcessos_QtdeMaqOtim) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_QtdePcsOtim) = ""
       .ColDataType(conCOL_SonProcessos_QtdePcsOtim) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_DtPrevOtim) = ""
       .ColDataType(conCOL_SonProcessos_DtPrevOtim) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_QtdeMaqIdeal) = ""
       .ColDataType(conCOL_SonProcessos_QtdeMaqIdeal) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_QtdePcsIdeal) = ""
       .ColDataType(conCOL_SonProcessos_QtdePcsIdeal) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_DtPrevIdeal) = ""
       .ColDataType(conCOL_SonProcessos_DtPrevIdeal) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_QtdeMaqPess) = ""
       .ColDataType(conCOL_SonProcessos_QtdeMaqPess) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_QtdePcsPess) = ""
       .ColDataType(conCOL_SonProcessos_QtdePcsPess) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_DtPrevPess) = ""
       .ColDataType(conCOL_SonProcessos_DtPrevPess) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProcessos_Indice) = ""
       .ColDataType(conCOL_SonProcessos_Indice) = flexDTLong
       
       .ColWidth(conCOL_SonProcessos_Processo) = 2500
       .ColWidth(conCOL_SonProcessos_CapPcHora) = 1100
       .ColWidth(conCOL_SonProcessos_QtdeMaq) = 1000
       .ColWidth(conCOL_SonProcessos_QtdeMaqOtim) = 1500
       .ColWidth(conCOL_SonProcessos_QtdePcsOtim) = 1500
       .ColWidth(conCOL_SonProcessos_QtdeMaqPess) = 1500
       .ColWidth(conCOL_SonProcessos_QtdePcsPess) = 1500
       .ColWidth(conCOL_SonProcessos_Indice) = 1000
       
       .ColHidden(conCOL_SonProcessos_Indice) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub InitGridCARTMAQMELHOR()

    With grdCARTMAQMELHOR
    
       .Cols = conColumnsIn_SonCARMAQMELHSEN
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonCARMAQMELHSEN_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQMELHSEN_CodMaq) = ""
       .ColDataType(conCOL_SonCARMAQMELHSEN_CodMaq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQMELHSEN_EficMed) = ""
       .ColDataType(conCOL_SonCARMAQMELHSEN_EficMed) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQMELHSEN_ProdTeor) = ""
       .ColDataType(conCOL_SonCARMAQMELHSEN_ProdTeor) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQMELHSEN_ProdReal) = ""
       .ColDataType(conCOL_SonCARMAQMELHSEN_ProdReal) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQMELHSEN_Soma) = ""
       .ColDataType(conCOL_SonCARMAQMELHSEN_Soma) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQMELHSEN_Pai) = ""
       .ColDataType(conCOL_SonCARMAQMELHSEN_Pai) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQMELHSEN_Indice) = ""
       .ColDataType(conCOL_SonCARMAQMELHSEN_Indice) = flexDTLong
       
       .ColWidth(conCOL_SonCARMAQMELHSEN_CodMaq) = 2000
       .ColWidth(conCOL_SonCARMAQMELHSEN_EficMed) = 1500
       .ColWidth(conCOL_SonCARMAQMELHSEN_ProdTeor) = 1500
       .ColWidth(conCOL_SonCARMAQMELHSEN_ProdReal) = 1500
       .ColWidth(conCOL_SonCARMAQMELHSEN_Pai) = 1000
       
       .ColHidden(conCOL_SonCARMAQMELHSEN_Pai) = True
       .ColHidden(conCOL_SonCARMAQMELHSEN_Indice) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub InitGridCARMAQPIORSEN()

    With grdCARTMAQPIOR
    
       .Cols = conColumnsIn_SonCARMAQPIORSEN
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonCARMAQPIORSEN_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQPIORSEN_CodMaq) = ""
       .ColDataType(conCOL_SonCARMAQPIORSEN_CodMaq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQPIORSEN_EficMed) = ""
       .ColDataType(conCOL_SonCARMAQPIORSEN_EficMed) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQPIORSEN_ProdTeor) = ""
       .ColDataType(conCOL_SonCARMAQPIORSEN_ProdTeor) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQPIORSEN_ProdReal) = ""
       .ColDataType(conCOL_SonCARMAQPIORSEN_ProdReal) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQPIORSEN_Soma) = ""
       .ColDataType(conCOL_SonCARMAQPIORSEN_Soma) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQPIORSEN_Pai) = ""
       .ColDataType(conCOL_SonCARMAQPIORSEN_Pai) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCARMAQPIORSEN_Indice) = ""
       .ColDataType(conCOL_SonCARMAQPIORSEN_Indice) = flexDTLong
       
       .ColWidth(conCOL_SonCARMAQPIORSEN_CodMaq) = 2000
       .ColWidth(conCOL_SonCARMAQPIORSEN_EficMed) = 1500
       .ColWidth(conCOL_SonCARMAQPIORSEN_ProdTeor) = 1500
       .ColWidth(conCOL_SonCARMAQPIORSEN_ProdReal) = 1500
       .ColWidth(conCOL_SonCARMAQPIORSEN_Pai) = 1500
       
       .ColHidden(conCOL_SonCARMAQPIORSEN_Pai) = True
       .ColHidden(conCOL_SonCARMAQPIORSEN_Indice) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub InitGridIdeal()

    With grdIdeal
    
       .Cols = conColumnsIn_SonIdeal
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonIdeal_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonIdeal_CodMaq) = ""
       .ColDataType(conCOL_SonIdeal_CodMaq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonIdeal_EficMed) = ""
       .ColDataType(conCOL_SonIdeal_EficMed) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonIdeal_ProdTeor) = ""
       .ColDataType(conCOL_SonIdeal_ProdTeor) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonIdeal_ProdReal) = ""
       .ColDataType(conCOL_SonIdeal_ProdReal) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonIdeal_Soma) = ""
       .ColDataType(conCOL_SonIdeal_Soma) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonIdeal_Pai) = ""
       .ColDataType(conCOL_SonIdeal_Pai) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonIdeal_Indice) = ""
       .ColDataType(conCOL_SonIdeal_Indice) = flexDTLong
       
       .ColWidth(conCOL_SonIdeal_CodMaq) = 2000
       .ColWidth(conCOL_SonIdeal_EficMed) = 1500
       .ColWidth(conCOL_SonIdeal_ProdTeor) = 1500
       .ColWidth(conCOL_SonIdeal_ProdReal) = 1500
       .ColWidth(conCOL_SonIdeal_Pai) = 2000
       
       .ColHidden(conCOL_SonIdeal_Pai) = True
       .ColHidden(conCOL_SonIdeal_Indice) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub PopGrdProcesso()

    Dim I As Integer
    
    Call InitGridProcessos
    Call InitGridCARTMAQMELHOR
    Call InitGridCARMAQPIORSEN
    Call InitGridIdeal
    
    For I = 0 To UBound(arrPROCESSO)
        If Len(Trim(arrPROCESSO(I).strDESCRI)) > 0 Then
           grdPROCESSOS.AddItem arrPROCESSO(I).strDESCRI & vbTab & _
                                Format(TotCapProdutiva(arrPROCESSO(I).lngCODIGO), "#,##0.00") & vbTab & _
                                QtdeMaqDisponivel(arrPROCESSO(I).lngCODIGO) & vbTab & _
                                "0,00" & vbTab & _
                                "0" & vbTab & _
                                "" & vbTab & _
                                "0,00" & vbTab & _
                                "0" & vbTab & _
                                "" & vbTab & _
                                "0,00" & vbTab & _
                                "0" & vbTab & _
                                "" & vbTab & _
                                arrPROCESSO(I).lngCODIGO
        End If
    Next I
    
    If (grdPROCESSOS.Rows - 1) > 0 Then
        grdPROCESSOS.Row = 1
        grdPROCESSOS.Cell(flexcpBackColor, 1, conCOL_SonProcessos_QtdePcsOtim, (grdPROCESSOS.Rows - 1), conCOL_SonProcessos_DtPrevOtim) = RGB(138, 255, 138)   '' Verde
        grdPROCESSOS.Cell(flexcpBackColor, 1, conCOL_SonProcessos_QtdePcsIdeal, (grdPROCESSOS.Rows - 1), conCOL_SonProcessos_DtPrevIdeal) = RGB(255, 255, 140) '' Amarelo
        grdPROCESSOS.Cell(flexcpBackColor, 1, conCOL_SonProcessos_QtdePcsPess, (grdPROCESSOS.Rows - 1), conCOL_SonProcessos_DtPrevPess) = RGB(255, 174, 174)   '' Vermelho
    End If

End Sub

Private Sub PopGrdMELHORSEN(lngCodProcesso As Long)

    Dim curSoma As Currency
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAQUI.SGI_CODIGO      " & vbCrLf
    sSql = sSql & "      ,MAQUI.SGI_DESCRI      " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_EFCMEDIA    " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_PRODPECTEOR " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_PRODPECREAL " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO     PROCE " & vbCrLf
    sSql = sSql & "      ,SGI_CADFICHATECHEAD FICHA " & vbCrLf
    sSql = sSql & "      ,SGI_CADMAQUINA      MAQUI " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "       PROCE.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROCE.SGI_CODIGO    = " & lngCodProcesso & vbCrLf
    sSql = sSql & "   And FICHA.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FICHA.SGI_CADFAMMAQ = PROCE.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_CODIGO    = FICHA.SGI_CODMAQ " & vbCrLf
    sSql = sSql & " Order by FICHA.SGI_PRODPECREAL DESC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    curSoma = 0
    Do While Not BREC.EOF
       curSoma = curSoma + BREC!SGI_PRODPECREAL
       grdCARTMAQMELHOR.AddItem BREC!SGI_DESCRI & vbTab & _
                                Format(BREC!SGI_EFCMEDIA, "#,##0.00") & vbTab & _
                                Format(BREC!SGI_PRODPECTEOR, "#,##0.00") & vbTab & _
                                Format(BREC!SGI_PRODPECREAL, "#,##0.00") & vbTab & _
                                Format(curSoma, "#,##0.00") & vbTab & _
                                lngCodProcesso & vbTab & _
                                BREC!SGI_CODIGO

       BREC.MoveNext
    Loop
    BREC.Close

End Sub


Private Sub PopGrdIdeal(lngCodProcesso As Long)

    
    Dim intMEIO                 As Integer
    Dim intCIMA                 As Integer
    Dim intBAIXO                As Integer
    Dim curVALOR                As Currency
    Dim arrIDEAL()              As IDEAL
    Dim intLinha                As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAQUI.SGI_CODIGO      " & vbCrLf
    sSql = sSql & "      ,MAQUI.SGI_DESCRI      " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_EFCMEDIA    " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_PRODPECTEOR " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_PRODPECREAL " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO PROCE     " & vbCrLf
    sSql = sSql & "      ,SGI_CADFICHATECHEAD FICHA " & vbCrLf
    sSql = sSql & "      ,SGI_CADMAQUINA      MAQUI " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "       PROCE.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROCE.SGI_CODIGO    = " & lngCodProcesso & vbCrLf
    sSql = sSql & "   And FICHA.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FICHA.SGI_CADFAMMAQ = PROCE.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_CODIGO    = FICHA.SGI_CODMAQ " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       BREC.Close
       Exit Sub
    End If
    
    intLinha = 1
    Do While Not BREC.EOF
       ReDim Preserve arrIDEAL(1 To intLinha) As IDEAL
    
       arrIDEAL(intLinha).SGI_DESCRI = BREC!SGI_DESCRI
       arrIDEAL(intLinha).SGI_EFCMEDIA = BREC!SGI_EFCMEDIA
       arrIDEAL(intLinha).SGI_PRODPECTEOR = BREC!SGI_PRODPECTEOR
       arrIDEAL(intLinha).SGI_PRODPECREAL = BREC!SGI_PRODPECREAL
       arrIDEAL(intLinha).SGI_PRODDIA = 0
       arrIDEAL(intLinha).lngPROCESSO = lngCodProcesso
       arrIDEAL(intLinha).lngCODMAQUINA = BREC!SGI_CODIGO
       
       intLinha = intLinha + 1
       BREC.MoveNext
    Loop
    BREC.Close

    '' ==============================================================================================
    '' Scenário Ideal
    intMEIO = (UBound(arrIDEAL) / 2)
    intCIMA = (intMEIO - 1)
    intBAIXO = (intMEIO + 1)

    curVALOR = arrIDEAL(intMEIO).SGI_PRODPECREAL
    intCampos = 1
    
    arrIDEAL(intCampos).SGI_PRODDIA = curVALOR
    intCampos = intCampos + 1
    For I = 1 To (intMEIO)
        If intCIMA > 0 Then
           curVALOR = curVALOR + arrIDEAL(intCIMA).SGI_PRODPECREAL
           arrIDEAL(intCampos).SGI_PRODDIA = curVALOR
           intCampos = (intCampos + 1)
        End If
        
        If intBAIXO <= UBound(arrIDEAL) Then
           curVALOR = curVALOR + arrIDEAL(intBAIXO).SGI_PRODPECREAL
           arrIDEAL(intCampos).SGI_PRODDIA = curVALOR
           intCampos = (intCampos + 1)
        End If
        intCIMA = (intCIMA - 1)
        intBAIXO = (intBAIXO + 1)
    Next I


    For I = 1 To UBound(arrIDEAL)
       grdIdeal.AddItem arrIDEAL(I).SGI_DESCRI & vbTab & _
                        Format(arrIDEAL(I).SGI_EFCMEDIA, "#,##0.00") & vbTab & _
                        Format(arrIDEAL(I).SGI_PRODPECTEOR, "#,##0.00") & vbTab & _
                        Format(arrIDEAL(I).SGI_PRODPECREAL, "#,##0.00") & vbTab & _
                        Format(arrIDEAL(I).SGI_PRODDIA, "#,##0.00") & vbTab & _
                        arrIDEAL(I).lngPROCESSO & vbTab & _
                        arrIDEAL(I).lngCODMAQUINA
    Next I

End Sub


Private Sub PopGrdPIORSEN(lngCodProcesso As Long)

    Dim curSoma As Currency
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAQUI.SGI_CODIGO      " & vbCrLf
    sSql = sSql & "      ,MAQUI.SGI_DESCRI      " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_EFCMEDIA    " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_PRODPECTEOR " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_PRODPECREAL " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO PROCE     " & vbCrLf
    sSql = sSql & "      ,SGI_CADFICHATECHEAD FICHA " & vbCrLf
    sSql = sSql & "      ,SGI_CADMAQUINA      MAQUI " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "       PROCE.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROCE.SGI_CODIGO    = " & lngCodProcesso & vbCrLf
    sSql = sSql & "   And FICHA.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FICHA.SGI_CADFAMMAQ = PROCE.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_CODIGO    = FICHA.SGI_CODMAQ " & vbCrLf
    sSql = sSql & " Order by FICHA.SGI_PRODPECREAL ASC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    curSoma = 0
    Do While Not BREC.EOF
       curSoma = curSoma + BREC!SGI_PRODPECREAL
       grdCARTMAQPIOR.AddItem BREC!SGI_DESCRI & vbTab & _
                              Format(BREC!SGI_EFCMEDIA, "#,##0.00") & vbTab & _
                              Format(BREC!SGI_PRODPECTEOR, "#,##0.00") & vbTab & _
                              Format(BREC!SGI_PRODPECREAL, "#,##0.00") & vbTab & _
                              Format(curSoma, "#,##0.00") & vbTab & _
                              lngCodProcesso & vbTab & _
                              BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

Public Function TotCapProdutiva(lngCodProcesso As Long) As Currency

    TotCapProdutiva = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(FICHA.SGI_PRODPECREAL) As SGI_PRODPECREAL " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO PROCE     " & vbCrLf
    sSql = sSql & "      ,SGI_CADFICHATECHEAD FICHA " & vbCrLf
    sSql = sSql & "      ,SGI_CADMAQUINA      MAQUI " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "       PROCE.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROCE.SGI_CODIGO    = " & lngCodProcesso & vbCrLf
    sSql = sSql & "   And FICHA.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FICHA.SGI_CADFAMMAQ = PROCE.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_CODIGO    = FICHA.SGI_CODMAQ " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not IsNull(BREC!SGI_PRODPECREAL) Then TotCapProdutiva = BREC!SGI_PRODPECREAL
    BREC.Close

End Function

Public Function QtdeMaqDisponivel(lngCodProcesso As Long) As Long

    QtdeMaqDisponivel = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Count(FICHA.SGI_CODMAQ) As SGI_QTDMAQUINA " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO PROCE     " & vbCrLf
    sSql = sSql & "      ,SGI_CADFICHATECHEAD FICHA " & vbCrLf
    sSql = sSql & "      ,SGI_CADMAQUINA      MAQUI " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "       PROCE.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROCE.SGI_CODIGO    = " & lngCodProcesso & vbCrLf
    sSql = sSql & "   And FICHA.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FICHA.SGI_CADFAMMAQ = PROCE.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_CODIGO    = FICHA.SGI_CODMAQ " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not IsNull(BREC!SGI_QTDMAQUINA) Then QtdeMaqDisponivel = BREC!SGI_QTDMAQUINA
    BREC.Close

End Function



Private Function ValidaPeriodo() As Boolean

    ValidaPeriodo = False
    
    If mskDtInicial.Text = "__/__/____" Or mskDtFinal.Text = "__/__/____" Then Exit Function
    
    If Not IsDate(mskDtInicial.Text) Then
       MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       ValidaPeriodo = True
       Exit Function
    End If
    If Not IsDate(mskDtFinal.Text) Then
       MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       ValidaPeriodo = True
       Exit Function
    End If
    
    If mskDtInicial.Text > mskDtFinal.Text Then
       MsgBox "Data Inicial Maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
       ValidaPeriodo = True
       Exit Function
    End If

    txtDemanda.Text = PegaPedidos(CDate(mskDtInicial.Text), CDate(mskDtFinal.Text), txtCODPROD.Text)
    Call PopScearios

End Function

Private Sub txtDemanda_GotFocus()
    objBLBFunc.SelecionaCampos txtDemanda.Name, frmCADFLUXOPROD
End Sub

Private Sub txtDemanda_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtDemanda.Text
End Sub

Private Sub txtDemanda_Validate(Cancel As Boolean)
    Call PrevisaoOtim(strMAQNESSOTIMGERAL, Date, txtDemanda.Text)
    Call PrevisaoPess(strMAQNESSPESSGERAL, Date, txtDemanda.Text)
    Call PrevisaoIdeal(strMAQNESSIDEALGERAL, Date, txtDemanda.Text)
End Sub


Private Sub ReservaMaquinas(curTOTDEMANDA As Currency, lngPROCESSO As Long)

    Dim I                       As Integer
    Dim intRowPesq              As Integer
    Dim intQtdMaqLocMelh        As Integer
    Dim curQtdProdMelh          As Currency
    Dim intQtdMaqLocIdeal       As Integer
    Dim curQtdProdIdeal         As Currency
    Dim intQtdMaqLocPior        As Integer
    Dim curQtdProdPior          As Currency
    
    
    '' ==============================================================================================
    '' Scenário Otimista
    intQtdMaqLocMelh = 0
    curQtdProdMelh = 0
    strMAQNESSOTIM = ""
    With grdCARTMAQMELHOR
         For I = 1 To (.Rows - 1)
             If lngPROCESSO = CLng(.Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_Pai)) Then
                If CCur(.Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_Soma)) < curTOTDEMANDA Then
                   .Cell(flexcpBackColor, I, conCOL_SonCARMAQMELHSEN_CodMaq, I, conCOL_SonCARMAQMELHSEN_Pai) = RGB(138, 255, 138) '' Verde
                   intQtdMaqLocMelh = (intQtdMaqLocMelh + 1)
                   curQtdProdMelh = curQtdProdMelh + CCur(.Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_ProdReal))
                   
                   strMAQNESSOTIM = strMAQNESSOTIM & .Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_Indice) & ";" & .Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_ProdReal) & ";" & .Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_Pai) & "|"
                   
                ElseIf CCur(.Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_Soma)) >= curTOTDEMANDA Then
                   .Cell(flexcpBackColor, I, conCOL_SonCARMAQMELHSEN_CodMaq, I, conCOL_SonCARMAQMELHSEN_Pai) = RGB(138, 255, 138) '' Verde
                   intQtdMaqLocMelh = (intQtdMaqLocMelh + 1)
                   curQtdProdMelh = curQtdProdMelh + CCur(.Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_ProdReal))
                   
                   strMAQNESSOTIM = strMAQNESSOTIM & .Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_Indice) & ";" & .Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_ProdReal) & ";" & .Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_Pai) & "|"
                   
                   Exit For
                End If
             End If
         Next I
    End With
    
    intRowPesq = grdPROCESSOS.FindRow(lngPROCESSO, , conCOL_SonProcessos_Indice)
    If intRowPesq <> -1 Then
       grdPROCESSOS.Cell(flexcpText, intRowPesq, conCOL_SonProcessos_QtdeMaqOtim) = intQtdMaqLocMelh
       grdPROCESSOS.Cell(flexcpText, intRowPesq, conCOL_SonProcessos_QtdePcsOtim) = Format(curQtdProdMelh, "#,##0.00")
    End If
    strMAQNESSOTIMGERAL = strMAQNESSOTIMGERAL & strMAQNESSOTIM & "#"
    
    
    '' ==============================================================================================
    '' Scenário Ideal
    intQtdMaqLocIdeal = 0
    curQtdProdIdeal = 0
    strMAQNESSIDEAL = ""
    With grdIdeal
         For I = 1 To (.Rows - 1)
             If lngPROCESSO = CLng(.Cell(flexcpText, I, conCOL_SonIdeal_Pai)) Then
                If CCur(.Cell(flexcpText, I, conCOL_SonIdeal_Soma)) < curTOTDEMANDA Then
                   .Cell(flexcpBackColor, I, conCOL_SonIdeal_CodMaq, I, conCOL_SonIdeal_Pai) = RGB(255, 255, 140) '' Amarelo
                   intQtdMaqLocIdeal = (intQtdMaqLocIdeal + 1)
                   curQtdProdIdeal = CCur(.Cell(flexcpText, I, conCOL_SonIdeal_Soma))
                   
                   strMAQNESSIDEAL = strMAQNESSIDEAL & .Cell(flexcpText, I, conCOL_SonIdeal_Indice) & ";" & .Cell(flexcpText, I, conCOL_SonIdeal_ProdReal) & ";" & .Cell(flexcpText, I, conCOL_SonIdeal_Pai) & "|"
                ElseIf CCur(.Cell(flexcpText, I, conCOL_SonIdeal_Soma)) >= curTOTDEMANDA Then
                   .Cell(flexcpBackColor, I, conCOL_SonIdeal_CodMaq, I, conCOL_SonIdeal_Pai) = RGB(255, 255, 140) '' Amarelo
                   intQtdMaqLocIdeal = (intQtdMaqLocIdeal + 1)
                   curQtdProdIdeal = CCur(.Cell(flexcpText, I, conCOL_SonIdeal_Soma))
                   
                   strMAQNESSIDEAL = strMAQNESSIDEAL & .Cell(flexcpText, I, conCOL_SonIdeal_Indice) & ";" & .Cell(flexcpText, I, conCOL_SonIdeal_ProdReal) & ";" & .Cell(flexcpText, I, conCOL_SonIdeal_Pai) & "|"
                   Exit For
                End If
             End If
         Next I
    End With
    If intRowPesq <> -1 Then
      grdPROCESSOS.Cell(flexcpText, intRowPesq, conCOL_SonProcessos_QtdeMaqIdeal) = intQtdMaqLocIdeal
      grdPROCESSOS.Cell(flexcpText, intRowPesq, conCOL_SonProcessos_QtdePcsIdeal) = Format(curQtdProdIdeal, "#,##0.00")
    End If
    strMAQNESSIDEALGERAL = strMAQNESSIDEALGERAL & strMAQNESSIDEAL & "#"
    
    '' ==============================================================================================
    '' Scenário Pessimista
    intQtdMaqLocPior = 0
    curQtdProdPior = 0
    strMAQNESSPESS = ""
    With grdCARTMAQPIOR
         For I = 1 To (.Rows - 1)
             If lngPROCESSO = CLng(.Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_Pai)) Then
                If CCur(.Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_Soma)) < curTOTDEMANDA Then
                   .Cell(flexcpBackColor, I, conCOL_SonCARMAQPIORSEN_CodMaq, I, conCOL_SonCARMAQPIORSEN_Pai) = RGB(255, 174, 174) '' Vermelho
                   intQtdMaqLocPior = (intQtdMaqLocPior + 1)
                   curQtdProdPior = curQtdProdPior + CCur(.Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_ProdReal))
                   
                   strMAQNESSPESS = strMAQNESSPESS & .Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_Indice) & ";" & .Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_ProdReal) & ";" & .Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_Pai) & "|"
                ElseIf CCur(.Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_Soma)) >= curTOTDEMANDA Then
                   .Cell(flexcpBackColor, I, conCOL_SonCARMAQPIORSEN_CodMaq, I, conCOL_SonCARMAQPIORSEN_Pai) = RGB(255, 174, 174) '' Vermelho
                   intQtdMaqLocPior = (intQtdMaqLocPior + 1)
                   curQtdProdPior = curQtdProdPior + CCur(.Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_ProdReal))
                   
                   strMAQNESSPESS = strMAQNESSPESS & .Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_Indice) & ";" & .Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_ProdReal) & ";" & .Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_Pai) & "|"
                   Exit For
                End If
             End If
         Next I
    End With
    
    If intRowPesq <> -1 Then
       grdPROCESSOS.Cell(flexcpText, intRowPesq, conCOL_SonProcessos_QtdeMaqPess) = intQtdMaqLocPior
       grdPROCESSOS.Cell(flexcpText, intRowPesq, conCOL_SonProcessos_QtdePcsPess) = Format(curQtdProdPior, "#,##0.00")
    End If
    strMAQNESSPESSGERAL = strMAQNESSPESSGERAL & strMAQNESSPESS & "#"
End Sub


Private Sub MaquinasNecess(lngRow As Long)

    If (grdPROCESSOS.Rows - 1) > 0 And lngRow > 0 Then
       Call PopGrdMELHORSEN(grdPROCESSOS.Cell(flexcpText, lngRow, conCOL_SonProcessos_Indice))
       Call PopGrdPIORSEN(grdPROCESSOS.Cell(flexcpText, lngRow, conCOL_SonProcessos_Indice))
       Call PopGrdIdeal(grdPROCESSOS.Cell(flexcpText, lngRow, conCOL_SonProcessos_Indice))
       
       If Len(Trim(txtQtde.Text)) = 0 Then Exit Sub
       Call ReservaMaquinas(CCur(txtQtde.Text), CLng(grdPROCESSOS.Cell(flexcpText, lngRow, conCOL_SonProcessos_Indice)))
    End If

End Sub

Public Function QtdeMaqOtimista(lngCodProcesso As Long) As Long

    QtdeMaqOtimista = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "      ,FICHA.SGI_PRODPECREAL " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO PROCE     " & vbCrLf
    sSql = sSql & "      ,SGI_CADFICHATECHEAD FICHA " & vbCrLf
    sSql = sSql & "      ,SGI_CADMAQUINA      MAQUI " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "       PROCE.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROCE.SGI_CODIGO    = " & lngCodProcesso & vbCrLf
    sSql = sSql & "   And FICHA.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FICHA.SGI_CADFAMMAQ = PROCE.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_FILIAL    = PROCE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And MAQUI.SGI_CODIGO    = FICHA.SGI_CODMAQ " & vbCrLf
    sSql = sSql & " Order by FICHA.SGI_PRODPECREAL DESC "
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC2.EOF()
        
       BREC2.MoveNext
    Loop
    BREC2.Close

End Function


Private Function PegaPedidos(dtDataIni As Date, dtDataFin As Date, strPRODUTO As String) As Currency

    PegaPedidos = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(ITENS.SGI_QTDE) As SGI_QTDE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH CABEC " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI ITENS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CABEC.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CABEC.SGI_DATAPED Between '" & Format(dtDataIni, "MM/DD/YYYY") & "' And '" & Format(dtDataFin, "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And CABEC.SGI_STATUS  = 'L'" & vbCrLf
    sSql = sSql & "   And ITENS.SGI_FILIAL  = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ITENS.SGI_CODIGO  = CABEC.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And ITENS.SGI_CODPROD = '" & Trim(strPRODUTO) & "'"

    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not IsNull(BREC3!SGI_QTDE) Then PegaPedidos = BREC3!SGI_QTDE
    BREC3.Close
    
End Function

Private Sub PopScearios()
    Dim I As Long
    If (grdPROCESSOS.Rows - 1) > 0 Then
        strMAQNESSOTIMGERAL = ""
        strMAQNESSPESSGERAL = ""
        strMAQNESSIDEALGERAL = ""
        For I = 1 To (grdPROCESSOS.Rows - 1)
            Call MaquinasNecess(I)
        Next I
    End If
End Sub

Private Sub InitGridCampo()

    With grdRitimo
    
       .Cols = conColumnsIn_SonCampo1
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonCampo1_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCampo1) = ""
       .ColDataType(conCOL_SonCampo1) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCampo2) = ""
       .ColDataType(conCOL_SonCampo2) = flexDTCurrency
       
       .ColWidth(conCOL_SonCampo1) = 4000
       .ColWidth(conCOL_SonCampo2) = 1000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PopGrdRitimo()
    
    Dim I As Integer
    
    Call InitGridCampo
    
    For I = 1 To (grdPROCESSOS.Rows - 1)
        grdRitimo.AddItem grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_Processo) & vbTab & grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_CapPcHora)
    Next I
    
    grdRitimo.Col = conCOL_SonCampo2
    grdRitimo.Sort = flexSortNumericAscending
    
    If (grdRitimo.Rows - 1) Then
       grdRitimo.Row = 1
       grdRitimo.Cell(flexcpBackColor, grdRitimo.Row, conCOL_SonCampo1, grdRitimo.Row, conCOL_SonCampo2) = vbRed
    End If
    
End Sub


Private Sub InitGridProdDiaOtim()

    With grdPRODDIAOTIM
    
       .Cols = conColumnsIn_SonProdDiaOtim
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonProdDiaOtim_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaOtim_Data) = ""
       .ColDataType(conCOL_SonProdDiaOtim_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaOtim_DiaSem) = ""
       .ColDataType(conCOL_SonProdDiaOtim_DiaSem) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaOtim_HorDisp) = ""
       .ColDataType(conCOL_SonProdDiaOtim_HorDisp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaOtim_TotProdHora) = ""
       .ColDataType(conCOL_SonProdDiaOtim_TotProdHora) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaOtim_TotProdTurn) = ""
       .ColDataType(conCOL_SonProdDiaOtim_TotProdTurn) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaOtim_Saldo) = ""
       .ColDataType(conCOL_SonProdDiaOtim_Saldo) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaOtim_Pai) = ""
       .ColDataType(conCOL_SonProdDiaOtim_Pai) = flexDTLong
       
       .ColWidth(conCOL_SonProdDiaOtim_Data) = 1000
       .ColWidth(conCOL_SonProdDiaOtim_DiaSem) = 1500
       .ColWidth(conCOL_SonProdDiaOtim_HorDisp) = 1000
       .ColWidth(conCOL_SonProdDiaOtim_TotProdHora) = 1000
       .ColWidth(conCOL_SonProdDiaOtim_TotProdTurn) = 1000
       .ColWidth(conCOL_SonProdDiaOtim_Saldo) = 1000
       .ColWidth(conCOL_SonProdDiaOtim_Pai) = 1000
       
       .ColHidden(conCOL_SonProdDiaOtim_Pai) = True
       .ColHidden(conCOL_SonProdDiaOtim_HorDisp) = True
       .ColHidden(conCOL_SonProdDiaOtim_TotProdHora) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub InitGridProdDiaPess()

    With grdPRODDIAPESS
    
       .Cols = conColumnsIn_SonProdDiaPess
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonProdDiaPess_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaPess_Data) = ""
       .ColDataType(conCOL_SonProdDiaPess_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaPess_DiaSem) = ""
       .ColDataType(conCOL_SonProdDiaPess_DiaSem) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaPess_HorDisp) = ""
       .ColDataType(conCOL_SonProdDiaPess_HorDisp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaPess_TotProdHora) = ""
       .ColDataType(conCOL_SonProdDiaPess_TotProdHora) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaPess_TotProdTurn) = ""
       .ColDataType(conCOL_SonProdDiaPess_TotProdTurn) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaPess_Saldo) = ""
       .ColDataType(conCOL_SonProdDiaPess_Saldo) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaPess_Pai) = ""
       .ColDataType(conCOL_SonProdDiaPess_Pai) = flexDTLong
       
       .ColWidth(conCOL_SonProdDiaPess_Data) = 1000
       .ColWidth(conCOL_SonProdDiaPess_DiaSem) = 1500
       .ColWidth(conCOL_SonProdDiaPess_HorDisp) = 1000
       .ColWidth(conCOL_SonProdDiaPess_TotProdHora) = 1000
       .ColWidth(conCOL_SonProdDiaPess_TotProdTurn) = 1000
       .ColWidth(conCOL_SonProdDiaPess_Saldo) = 1000
       .ColWidth(conCOL_SonProdDiaPess_Pai) = 1000
       
       .ColHidden(conCOL_SonProdDiaPess_Pai) = True
       .ColHidden(conCOL_SonProdDiaPess_HorDisp) = True
       .ColHidden(conCOL_SonProdDiaPess_TotProdHora) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Function PegaQtdeTurno(lngCodProcesso As Long) As Long

    PegaQtdeTurno = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Count(TUR.SGI_CODIGO) As QtdeTurn " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO      PRC " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMMAQUINAS   FAM " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMTURNO      FAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN      TUR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRC.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRC.SGI_CODIGO = " & lngCodProcesso & vbCrLf
    sSql = sSql & "   And FAM.SGI_FILIAL = PRC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FAM.SGI_CODIGO = PRC.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "   And FAT.SGI_FILIAL = FAM.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FAT.SGI_CODIGO = FAM.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And TUR.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And TUR.SGI_CODIGO = FAT.SGI_CODTURNO "
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaQtdeTurno = BREC!QtdeTurn
    BREC.Close
    
End Function

Private Sub PopGrdProdDiaOtim(lngCODPROC As Long, dtDate As Date, curDemanda As Currency)

    Dim I                     As Integer
    Dim lngDiaSemana          As Long
    Dim lngTOTHORAS           As Long
    Dim lngTOTGERAL           As Long
    Dim strTOTHORAS           As String
    Dim TOTMINUTOS            As Double
    Dim curTOTPRODHORAOTIM    As Currency
    Dim curTOTPCTURNO         As Currency
    Dim curTOTDEMANDA         As Currency
    
    Dim arrDIASSEMANA(1 To 7) As String
    arrDIASSEMANA(1) = "Domingo"
    arrDIASSEMANA(2) = "Segunda"
    arrDIASSEMANA(3) = "Terça"
    arrDIASSEMANA(4) = "Quarta"
    arrDIASSEMANA(5) = "Quinta"
    arrDIASSEMANA(6) = "Sexta"
    arrDIASSEMANA(7) = "Sabado"
    
    curTOTDEMANDA = curDemanda
    
    curTOTPRODHORAOTIM = PegaProdDia(lngCODPROC, conCOL_SonProcessos_QtdePcsOtim)
    
Retorno:
    lngDiaSemana = Weekday(dtDate)
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SEM.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO      PRC " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMMAQUINAS   FAM " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMTURNO      FAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN      TUR " & vbCrLf
    sSql = sSql & "      ,SGI_CADTURNSEM       SEM " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        PRC.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "    And PRC.SGI_CODIGO = " & lngCODPROC & vbCrLf
    sSql = sSql & "    And FAM.SGI_FILIAL = PRC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And FAM.SGI_CODIGO = PRC.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "    And FAT.SGI_FILIAL = FAM.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And FAT.SGI_CODIGO = FAM.SGI_CODIGO " & vbCrLf
    sSql = sSql & "    And TUR.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And TUR.SGI_CODIGO = FAT.SGI_CODTURNO " & vbCrLf
    sSql = sSql & "    And SEM.SGI_FILIAL = TUR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And SEM.SGI_CODIGO = TUR.SGI_CODIGO " & vbCrLf
    sSql = sSql & "    And SEM.SGI_DIASEM = " & lngDiaSemana & vbCrLf
    sSql = sSql & " Order by SGI_DIASEM "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    lngTOTHORAS = 0
    lngTOTGERAL = 0
    If Not BREC.EOF Then
       Do While Not BREC.EOF()
          lngTOTHORAS = CONVHRMIN(BREC!SGI_HORALIQ)
          lngTOTGERAL = (lngTOTGERAL + lngTOTHORAS)
          BREC.MoveNext
       Loop
    
       With grdPRODDIAOTIM
            strTOTHORAS = CONVMINHR(lngTOTGERAL)
            
            
            
            
            
            '' ======================================
            TOTMINUTOS = Round((lngTOTGERAL / 60), 2)
            curTOTPCTURNO = (TOTMINUTOS * curTOTPRODHORAOTIM)
            curTOTDEMANDA = (curTOTDEMANDA - curTOTPCTURNO)
            '' ======================================
       
            grdPRODDIAOTIM.AddItem Format(dtDate, "DD/MM/YYYY") & vbTab & _
                                   arrDIASSEMANA(lngDiaSemana) & vbTab & _
                                   strTOTHORAS & vbTab & _
                                   Format(curTOTPRODHORAOTIM, "#,##0.00") & vbTab & _
                                   Format(curTOTPCTURNO, "#,##0.00") & vbTab & _
                                   Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                   lngCODPROC
                                   
                                   
            If curTOTDEMANDA > 0 Then
               BREC.Close
               dtDate = (dtDate + 1)
               GoTo Retorno
            End If
       End With
    Else
       BREC.Close
       dtDate = (dtDate + 1)
       GoTo Retorno
    End If
    BREC.Close
    
    
    '' ===================================
    If grdPROCESSOS.Rows - 1 Then
       For I = 1 To (grdPROCESSOS.Rows - 1)
           If grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_Indice) = lngCODPROC Then
              grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_DtPrevOtim) = grdPRODDIAOTIM.Cell(flexcpText, grdPRODDIAOTIM.Rows - 1, conCOL_SonProdDiaOtim_Data)
              Exit For
           End If
       Next I
    End If
    '' ===================================
End Sub

Private Sub PopGrdProdDiaIdeal(lngCODPROC As Long, dtDate As Date, curDemanda As Currency)

    Dim lngDiaSemana          As Long
    Dim lngTOTHORAS           As Long
    Dim lngTOTGERAL           As Long
    Dim strTOTHORAS           As String
    Dim TOTMINUTOS            As Double
    Dim curTOTPRODHORAIDEAL   As Currency
    Dim curTOTPCTURNO         As Currency
    Dim curTOTDEMANDA         As Currency
    
    Dim arrDIASSEMANA(1 To 7) As String
    arrDIASSEMANA(1) = "Domingo"
    arrDIASSEMANA(2) = "Segunda"
    arrDIASSEMANA(3) = "Terça"
    arrDIASSEMANA(4) = "Quarta"
    arrDIASSEMANA(5) = "Quinta"
    arrDIASSEMANA(6) = "Sexta"
    arrDIASSEMANA(7) = "Sabado"
    
    curTOTDEMANDA = curDemanda
    
    curTOTPRODHORAIDEAL = PegaProdDia(lngCODPROC, conCOL_SonProcessos_QtdePcsIdeal)
    
Retorno:
    lngDiaSemana = Weekday(dtDate)
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SEM.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO      PRC " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMMAQUINAS   FAM " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMTURNO      FAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN      TUR " & vbCrLf
    sSql = sSql & "      ,SGI_CADTURNSEM       SEM " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRC.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "    And PRC.SGI_CODIGO = " & lngCODPROC & vbCrLf
    sSql = sSql & "    And FAM.SGI_FILIAL = PRC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And FAM.SGI_CODIGO = PRC.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "    And FAT.SGI_FILIAL = FAM.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And FAT.SGI_CODIGO = FAM.SGI_CODIGO " & vbCrLf
    sSql = sSql & "    And TUR.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And TUR.SGI_CODIGO = FAT.SGI_CODTURNO " & vbCrLf
    sSql = sSql & "    And SEM.SGI_FILIAL = TUR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And SEM.SGI_CODIGO = TUR.SGI_CODIGO " & vbCrLf
    sSql = sSql & "    And SEM.SGI_DIASEM = " & lngDiaSemana & vbCrLf
    sSql = sSql & " Order by SGI_DIASEM "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    lngTOTHORAS = 0
    lngTOTGERAL = 0
    If Not BREC.EOF Then
       Do While Not BREC.EOF()
          lngTOTHORAS = CONVHRMIN(BREC!SGI_HORALIQ)
          lngTOTGERAL = (lngTOTGERAL + lngTOTHORAS)
          BREC.MoveNext
       Loop
    
       With grdPRODDIAOTIM
            strTOTHORAS = CONVMINHR(lngTOTGERAL)
            
            '' ======================================
            TOTMINUTOS = Round((lngTOTGERAL / 60), 2)
            curTOTPCTURNO = (TOTMINUTOS * curTOTPRODHORAIDEAL)
            curTOTDEMANDA = (curTOTDEMANDA - curTOTPCTURNO)
            '' ======================================
       
            grdPRODDIAIDEAL.AddItem Format(dtDate, "DD/MM/YYYY") & vbTab & _
                                    arrDIASSEMANA(lngDiaSemana) & vbTab & _
                                    strTOTHORAS & vbTab & _
                                    Format(curTOTPRODHORAIDEAL, "#,##0.00") & vbTab & _
                                    Format(curTOTPCTURNO, "#,##0.00") & vbTab & _
                                    Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                    lngCODPROC
                                   
                                   
            If curTOTDEMANDA > 0 Then
               BREC.Close
               dtDate = (dtDate + 1)
               GoTo Retorno
            End If
       End With
    Else
       BREC.Close
       dtDate = (dtDate + 1)
       GoTo Retorno
    End If
    BREC.Close
    
    '' ===================================
    If grdPROCESSOS.Rows - 1 Then
       For I = 1 To (grdPROCESSOS.Rows - 1)
           If grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_Indice) = lngCODPROC Then
              grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_DtPrevIdeal) = grdPRODDIAIDEAL.Cell(flexcpText, grdPRODDIAIDEAL.Rows - 1, conCOL_SonProdDiaIdeal_Data)
              Exit For
           End If
       Next I
    End If
    '' ===================================
    
    
End Sub

Private Sub PopGrdProdDiaPess(lngCODPROC As Long, dtDate As Date, curDemanda As Currency)

    Dim lngDiaSemana          As Long
    Dim lngTOTHORAS           As Long
    Dim lngTOTGERAL           As Long
    Dim strTOTHORAS           As String
    Dim TOTMINUTOS            As Double
    Dim curTOTPRODHORAPESS    As Currency
    Dim curTOTPCTURNO         As Currency
    Dim curTOTDEMANDA         As Currency
    
    Dim arrDIASSEMANA(1 To 7) As String
    arrDIASSEMANA(1) = "Domingo"
    arrDIASSEMANA(2) = "Segunda"
    arrDIASSEMANA(3) = "Terça"
    arrDIASSEMANA(4) = "Quarta"
    arrDIASSEMANA(5) = "Quinta"
    arrDIASSEMANA(6) = "Sexta"
    arrDIASSEMANA(7) = "Sabado"
    
    curTOTDEMANDA = curDemanda
    
    curTOTPRODHORAPESS = PegaProdDia(lngCODPROC, conCOL_SonProcessos_QtdePcsPess)
    
Retorno:
    lngDiaSemana = Weekday(dtDate)
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SEM.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO      PRC " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMMAQUINAS   FAM " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMTURNO      FAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN      TUR " & vbCrLf
    sSql = sSql & "      ,SGI_CADTURNSEM       SEM " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRC.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "    And PRC.SGI_CODIGO = " & lngCODPROC & vbCrLf
    sSql = sSql & "    And FAM.SGI_FILIAL = PRC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And FAM.SGI_CODIGO = PRC.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "    And FAT.SGI_FILIAL = FAM.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And FAT.SGI_CODIGO = FAM.SGI_CODIGO " & vbCrLf
    sSql = sSql & "    And TUR.SGI_FILIAL = FAT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And TUR.SGI_CODIGO = FAT.SGI_CODTURNO " & vbCrLf
    sSql = sSql & "    And SEM.SGI_FILIAL = TUR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "    And SEM.SGI_CODIGO = TUR.SGI_CODIGO " & vbCrLf
    sSql = sSql & "    And SEM.SGI_DIASEM = " & lngDiaSemana & vbCrLf
    sSql = sSql & " Order by SGI_DIASEM "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    lngTOTHORAS = 0
    lngTOTGERAL = 0
    If Not BREC.EOF Then
       Do While Not BREC.EOF()
          lngTOTHORAS = CONVHRMIN(BREC!SGI_HORALIQ)
          lngTOTGERAL = (lngTOTGERAL + lngTOTHORAS)
          BREC.MoveNext
       Loop
    
       With grdPRODDIAOTIM
            strTOTHORAS = CONVMINHR(lngTOTGERAL)
            
            '' ======================================
            TOTMINUTOS = Round((lngTOTGERAL / 60), 2)
            curTOTPCTURNO = (TOTMINUTOS * curTOTPRODHORAPESS)
            curTOTDEMANDA = (curTOTDEMANDA - curTOTPCTURNO)
            '' ======================================
       
            grdPRODDIAPESS.AddItem Format(dtDate, "DD/MM/YYYY") & vbTab & _
                                    arrDIASSEMANA(lngDiaSemana) & vbTab & _
                                    strTOTHORAS & vbTab & _
                                    Format(curTOTPRODHORAPESS, "#,##0.00") & vbTab & _
                                    Format(curTOTPCTURNO, "#,##0.00") & vbTab & _
                                    Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                    lngCODPROC
                                   
                                   
            If curTOTDEMANDA > 0 Then
               BREC.Close
               dtDate = (dtDate + 1)
               GoTo Retorno
            End If
       End With
    Else
       BREC.Close
       dtDate = (dtDate + 1)
       GoTo Retorno
    End If
    BREC.Close
    
    '' ===================================
    If grdPROCESSOS.Rows - 1 Then
       For I = 1 To (grdPROCESSOS.Rows - 1)
           If grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_Indice) = lngCODPROC Then
              grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_DtPrevPess) = grdPRODDIAPESS.Cell(flexcpText, grdPRODDIAPESS.Rows - 1, conCOL_SonProdDiaPess_Data)
              Exit For
           End If
       Next I
    End If
    '' ===================================
    
End Sub


Private Function PegaProdDia(lngCODPROC As Long, lngCol As Long) As Currency
    
    PegaProdDia = 0
    
    Dim I As Integer
    For I = 1 To (grdPROCESSOS.Rows - 1)
        If lngCODPROC = grdPROCESSOS.Cell(flexcpText, I, conCOL_SonProcessos_Indice) Then
           PegaProdDia = CCur(grdPROCESSOS.Cell(flexcpText, I, lngCol))
        End If
    Next I
End Function


Private Function CONVHRMIN(strHORA As String) As Long
    
    CONVHRMIN = 0
    
    If Len(Trim(strHORA)) = 0 Then Exit Function
    
    Dim HORAS       As Long
    Dim MINUTOS     As Long
    Dim TOTMINUTOS  As Long
    
    HORAS = Hour(CDate(strHORA))
    MINUTOS = Minute(CDate(strHORA))
    TOTMINUTOS = ((HORAS * 60) + MINUTOS)
    
    CONVHRMIN = TOTMINUTOS

End Function

Private Function CONVMINHR(lngMINUTOS As Long) As String
    
    CONVMINHR = ""
    
    If lngMINUTOS = 0 Then Exit Function
    
    Dim TOTMINUTOS  As Double
    Dim HORA        As Long
    Dim MINUTO      As Long
    Dim strHORAS    As String
    Dim arrHRMN()   As String
    
    TOTMINUTOS = Round((lngMINUTOS / 60), 2)
    strHORAS = Format(TOTMINUTOS, "###,##000.00")
    arrHRMN = Split(strHORAS, ",")
    
    HORA = CLng(arrHRMN(0))
    MINUTO = (CLng(arrHRMN(1)) * (0.6))
    
    CONVMINHR = Trim(Format(HORA, "##00") & ":" & Format(MINUTO, "##00") & ":" & "00")

End Function

Private Sub PosRegGrdProcessos(strCODPROD As String)
    Dim I As Integer
    
    With grdPRODDIAOTIM
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonProdDiaOtim_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

    With grdPRODDIAIDEAL
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonProdDiaIdeal_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

    With grdPRODDIAPESS
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonProdDiaPess_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

End Sub


Private Sub PosRegGrdCarteira(strCODPROD As String)
    Dim I As Integer
    
    With grdCARTMAQMELHOR
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonCARMAQMELHSEN_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

    With grdIdeal
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonIdeal_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

    With grdCARTMAQPIOR
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonCARMAQPIORSEN_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

End Sub

Private Sub PosRegGrdMaqProd(strCODPROD As String)
    Dim I As Integer
    
    With grdProdMaqDiaOtim
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonProdMaqDiaOtim_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

    With grdProdMaqDiaIdeal
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonProdMaqDiaIdeal_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

    With grdProdMaqDiaPess
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonProdMaqDiaPess_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With

End Sub

Private Sub InitGridProdDiaIdeal()

    With grdPRODDIAIDEAL
    
       .Cols = conColumnsIn_SonProdDiaIdeal
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonProdDiaIdeal_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaIdeal_Data) = ""
       .ColDataType(conCOL_SonProdDiaIdeal_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaIdeal_DiaSem) = ""
       .ColDataType(conCOL_SonProdDiaIdeal_DiaSem) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaIdeal_HorDisp) = ""
       .ColDataType(conCOL_SonProdDiaIdeal_HorDisp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaIdeal_TotProdHora) = ""
       .ColDataType(conCOL_SonProdDiaIdeal_TotProdHora) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaIdeal_TotProdTurn) = ""
       .ColDataType(conCOL_SonProdDiaIdeal_TotProdTurn) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaIdeal_Saldo) = ""
       .ColDataType(conCOL_SonProdDiaIdeal_Saldo) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdDiaIdeal_Pai) = ""
       .ColDataType(conCOL_SonProdDiaIdeal_Pai) = flexDTLong
       
       .ColWidth(conCOL_SonProdDiaIdeal_Data) = 1000
       .ColWidth(conCOL_SonProdDiaIdeal_DiaSem) = 1500
       .ColWidth(conCOL_SonProdDiaIdeal_HorDisp) = 1000
       .ColWidth(conCOL_SonProdDiaIdeal_TotProdHora) = 1000
       .ColWidth(conCOL_SonProdDiaIdeal_TotProdTurn) = 1000
       .ColWidth(conCOL_SonProdDiaIdeal_Saldo) = 1000
       .ColWidth(conCOL_SonProdDiaIdeal_Pai) = 1000
       
       .ColHidden(conCOL_SonProdDiaIdeal_Pai) = True
       .ColHidden(conCOL_SonProdDiaIdeal_HorDisp) = True
       .ColHidden(conCOL_SonProdDiaIdeal_TotProdHora) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub PrevisaoOtim(strCODMAQS As String, dtData As Date, strDemanda As String)
    
    Dim intDiaSemana  As Integer
    Dim dtDataAtual   As Date
    Dim lngTotHora    As Long
    Dim TOTMINUTOS    As Double
    Dim ccurPRODDIA   As Currency
    Dim curTOTPCTURNO As Currency
    Dim curTOTDEMANDA As Currency
    Dim curTOTGERMAQ  As Currency
    
    Dim arrGRPMAQUINAS() As String
    Dim arrMAQUINAS()    As String
    Dim arrDADOS()       As String
    Dim I                As Integer
    Dim j                As Integer
    Dim intINDICE        As Integer
    Dim curDemanda       As Currency
    
    Dim arrDIASSEMANA(1 To 7) As String
    arrDIASSEMANA(1) = "Domingo"
    arrDIASSEMANA(2) = "Segunda"
    arrDIASSEMANA(3) = "Terça"
    arrDIASSEMANA(4) = "Quarta"
    arrDIASSEMANA(5) = "Quinta"
    arrDIASSEMANA(6) = "Sexta"
    arrDIASSEMANA(7) = "Sabado"
    
    Call InitGridProdDiaOtim
    Call InitGridProdMaqDiaOtim
    
    If Len(Trim(strDemanda)) = 0 Then Exit Sub
    
    arrGRPMAQUINAS = Split(Trim(strCODMAQS), "#")
    dtDataAtual = dtData
        
    For I = 0 To (UBound(arrGRPMAQUINAS) - 1)
        curTOTDEMANDA = CCur(strDemanda)
        dtDataAtual = dtData
        intDiaSemana = Weekday(dtDataAtual)
        
        Do While curTOTDEMANDA > 0
           
           arrMAQUINAS = Split(arrGRPMAQUINAS(I), "|")
           curTOTGERMAQ = 0
           For j = 0 To UBound(arrMAQUINAS) - 1
        
               arrDADOS = Split(arrMAQUINAS(j), ";")
               ccurPRODDIA = CCur(arrDADOS(1))
        
               sSql = "Select " & vbCrLf
               sSql = sSql & "       TUR.* " & vbCrLf
               sSql = sSql & "  From " & vbCrLf
               sSql = sSql & "       SGI_CADMAQUINA     MAQ " & vbCrLf
               sSql = sSql & "      ,SGI_CADFAMTURNO    FAM " & vbCrLf
               sSql = sSql & "      ,SGI_CADTURNSEM     TUR " & vbCrLf
               sSql = sSql & " Where " & vbCrLf
               sSql = sSql & "       MAQ.SGI_FILIAL     = " & FILIAL & vbCrLf
               sSql = sSql & "   And MAQ.SGI_CODIGO     = " & arrDADOS(0) & vbCrLf
               sSql = sSql & "   And FAM.SGI_FILIAL     = MAQ.SGI_FILIAL     " & vbCrLf
               sSql = sSql & "   And FAM.SGI_CODIGO     = MAQ.SGI_CODFAMILIA " & vbCrLf
               sSql = sSql & "   And TUR.SGI_FILIAL     = FAM.SGI_FILIAL     " & vbCrLf
               sSql = sSql & "   And TUR.SGI_CODIGO     = FAM.SGI_CODTURNO   " & vbCrLf
               sSql = sSql & "   And TUR.SGI_DIASEM     = " & intDiaSemana
        
               BREC.Open sSql, adoBanco_Dados, adOpenDynamic
               lngTotHora = 0
               If Not BREC.EOF() Then
                  Do While Not BREC.EOF()
                     lngTotHora = lngTotHora + CONVHRMIN(BREC!SGI_HORALIQ)
                     BREC.MoveNext
                  Loop
               
                  lngTotHora = (lngTotHora - PegaDescHoras(CLng(arrDADOS(0)), dtDataAtual))
               
                  '' ======================================
                  TOTMINUTOS = Round((lngTotHora / 60), 2)
                  curTOTPCTURNO = (TOTMINUTOS * ccurPRODDIA)
                  curTOTGERMAQ = (curTOTGERMAQ + curTOTPCTURNO)
                  '' ======================================
                  
                  If curTOTGERMAQ > 0 Then
                     grdProdMaqDiaOtim.AddItem Format(dtDataAtual, "DD/MM/YYYY") & vbTab & _
                                               arrDIASSEMANA(intDiaSemana) & vbTab & _
                                               CONVMINHR(lngTotHora) & vbTab & _
                                               Format(ccurPRODDIA, "#,##0.00") & vbTab & _
                                               Format(curTOTPCTURNO, "#,##0.00") & vbTab & _
                                               Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                               arrDADOS(0)
                  End If
                  
               End If
               BREC.Close
            
           Next j
           
           curTOTDEMANDA = (curTOTDEMANDA - curTOTGERMAQ)
           If curTOTGERMAQ > 0 Then
              grdPRODDIAOTIM.AddItem Format(dtDataAtual, "DD/MM/YYYY") & vbTab & _
                                     arrDIASSEMANA(intDiaSemana) & vbTab & _
                                     CONVMINHR(lngTotHora) & vbTab & _
                                     Format(0, "#,##0.00") & vbTab & _
                                     Format(curTOTGERMAQ, "#,##0.00") & vbTab & _
                                     Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                     arrDADOS(2)
                                     
              
              intINDICE = grdPROCESSOS.FindRow(arrDADOS(2), , conCOL_SonProcessos_Indice)
              If intINDICE <> -1 Then
                 grdPROCESSOS.Cell(flexcpText, intINDICE, conCOL_SonProcessos_DtPrevOtim) = Format(dtDataAtual, "DD/MM/YYYY")
              End If
              
           End If
           
           dtDataAtual = (dtDataAtual + 1)
           intDiaSemana = Weekday(dtDataAtual)
           
        Loop
    
    Next I
    
    Call grdPROCESSOS_Click
    Call grdCARTMAQMELHOR_Click
    
End Sub

Private Sub PrevisaoPess(strCODMAQS As String, dtData As Date, strDemanda As String)
    
    Dim intDiaSemana  As Integer
    Dim dtDataAtual   As Date
    Dim lngTotHora    As Long
    Dim TOTMINUTOS    As Double
    Dim ccurPRODDIA   As Currency
    Dim curTOTPCTURNO As Currency
    Dim curTOTDEMANDA As Currency
    Dim curTOTGERMAQ  As Currency
    
    Dim arrGRPMAQUINAS() As String
    Dim arrMAQUINAS()    As String
    Dim arrDADOS()       As String
    Dim I                As Integer
    Dim j                As Integer
    Dim intINDICE        As Integer
    Dim curDemanda       As Currency
    
    Dim arrDIASSEMANA(1 To 7) As String
    arrDIASSEMANA(1) = "Domingo"
    arrDIASSEMANA(2) = "Segunda"
    arrDIASSEMANA(3) = "Terça"
    arrDIASSEMANA(4) = "Quarta"
    arrDIASSEMANA(5) = "Quinta"
    arrDIASSEMANA(6) = "Sexta"
    arrDIASSEMANA(7) = "Sabado"
    
    Call InitGridProdDiaPess
    Call InitGridProdMaqDiaPess
    
    If Len(Trim(strDemanda)) = 0 Then Exit Sub
    
    arrGRPMAQUINAS = Split(Trim(strCODMAQS), "#")
    dtDataAtual = dtData
        
    For I = 0 To (UBound(arrGRPMAQUINAS) - 1)
        curTOTDEMANDA = CCur(strDemanda)
        dtDataAtual = dtData
        intDiaSemana = Weekday(dtDataAtual)
        
        Do While curTOTDEMANDA > 0
           
           arrMAQUINAS = Split(arrGRPMAQUINAS(I), "|")
           curTOTGERMAQ = 0
           For j = 0 To UBound(arrMAQUINAS) - 1
        
               arrDADOS = Split(arrMAQUINAS(j), ";")
               ccurPRODDIA = CCur(arrDADOS(1))
        
               sSql = "Select " & vbCrLf
               sSql = sSql & "       TUR.* " & vbCrLf
               sSql = sSql & "  From " & vbCrLf
               sSql = sSql & "       SGI_CADMAQUINA     MAQ " & vbCrLf
               sSql = sSql & "      ,SGI_CADFAMTURNO    FAM " & vbCrLf
               sSql = sSql & "      ,SGI_CADTURNSEM     TUR " & vbCrLf
               sSql = sSql & " Where " & vbCrLf
               sSql = sSql & "       MAQ.SGI_FILIAL     = " & FILIAL & vbCrLf
               sSql = sSql & "   And MAQ.SGI_CODIGO     = " & arrDADOS(0) & vbCrLf
               sSql = sSql & "   And FAM.SGI_FILIAL     = MAQ.SGI_FILIAL     " & vbCrLf
               sSql = sSql & "   And FAM.SGI_CODIGO     = MAQ.SGI_CODFAMILIA " & vbCrLf
               sSql = sSql & "   And TUR.SGI_FILIAL     = FAM.SGI_FILIAL     " & vbCrLf
               sSql = sSql & "   And TUR.SGI_CODIGO     = FAM.SGI_CODTURNO   " & vbCrLf
               sSql = sSql & "   And TUR.SGI_DIASEM     = " & intDiaSemana
        
               BREC.Open sSql, adoBanco_Dados, adOpenDynamic
               lngTotHora = 0
               If Not BREC.EOF() Then
                  Do While Not BREC.EOF()
                     lngTotHora = lngTotHora + CONVHRMIN(BREC!SGI_HORALIQ)
                     BREC.MoveNext
                  Loop
               
                  lngTotHora = (lngTotHora - PegaDescHoras(CLng(arrDADOS(0)), dtDataAtual))
               
                  '' ======================================
                  TOTMINUTOS = Round((lngTotHora / 60), 2)
                  curTOTPCTURNO = (TOTMINUTOS * ccurPRODDIA)
                  curTOTGERMAQ = (curTOTGERMAQ + curTOTPCTURNO)
                  '' ======================================
                  
                  If curTOTGERMAQ > 0 Then
                     grdProdMaqDiaPess.AddItem Format(dtDataAtual, "DD/MM/YYYY") & vbTab & _
                                               arrDIASSEMANA(intDiaSemana) & vbTab & _
                                               CONVMINHR(lngTotHora) & vbTab & _
                                               Format(ccurPRODDIA, "#,##0.00") & vbTab & _
                                               Format(curTOTPCTURNO, "#,##0.00") & vbTab & _
                                               Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                               arrDADOS(0)
                  End If
               
               End If
               BREC.Close
            
           Next j
           
           curTOTDEMANDA = (curTOTDEMANDA - curTOTGERMAQ)
           If curTOTGERMAQ > 0 Then
              grdPRODDIAPESS.AddItem Format(dtDataAtual, "DD/MM/YYYY") & vbTab & _
                                     arrDIASSEMANA(intDiaSemana) & vbTab & _
                                     CONVMINHR(lngTotHora) & vbTab & _
                                     Format(0, "#,##0.00") & vbTab & _
                                     Format(curTOTGERMAQ, "#,##0.00") & vbTab & _
                                     Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                     arrDADOS(2)
                                     
              
              intINDICE = grdPROCESSOS.FindRow(arrDADOS(2), , conCOL_SonProcessos_Indice)
              If intINDICE <> -1 Then
                 grdPROCESSOS.Cell(flexcpText, intINDICE, conCOL_SonProcessos_DtPrevPess) = Format(dtDataAtual, "DD/MM/YYYY")
              End If
              
           End If
           
           dtDataAtual = (dtDataAtual + 1)
           intDiaSemana = Weekday(dtDataAtual)
           
        Loop
    
    Next I
    
    Call grdPROCESSOS_Click
End Sub

Private Sub PrevisaoIdeal(strCODMAQS As String, dtData As Date, strDemanda As String)
    
    Dim intDiaSemana  As Integer
    Dim dtDataAtual   As Date
    Dim lngTotHora    As Long
    Dim TOTMINUTOS    As Double
    Dim ccurPRODDIA   As Currency
    Dim curTOTPCTURNO As Currency
    Dim curTOTDEMANDA As Currency
    Dim curTOTGERMAQ  As Currency
    
    Dim arrGRPMAQUINAS() As String
    Dim arrMAQUINAS()    As String
    Dim arrDADOS()       As String
    Dim I                As Integer
    Dim j                As Integer
    Dim intINDICE        As Integer
    Dim curDemanda       As Currency
    
    Dim arrDIASSEMANA(1 To 7) As String
    arrDIASSEMANA(1) = "Domingo"
    arrDIASSEMANA(2) = "Segunda"
    arrDIASSEMANA(3) = "Terça"
    arrDIASSEMANA(4) = "Quarta"
    arrDIASSEMANA(5) = "Quinta"
    arrDIASSEMANA(6) = "Sexta"
    arrDIASSEMANA(7) = "Sabado"
    
    Call InitGridProdDiaIdeal
    Call InitGridProdMaqDiaIdeal
    
    If Len(Trim(strDemanda)) = 0 Then Exit Sub
    
    arrGRPMAQUINAS = Split(Trim(strCODMAQS), "#")
    dtDataAtual = dtData
        
    For I = 0 To (UBound(arrGRPMAQUINAS) - 1)
        curTOTDEMANDA = CCur(strDemanda)
        dtDataAtual = dtData
        intDiaSemana = Weekday(dtDataAtual)
        
        Do While curTOTDEMANDA > 0
           
           arrMAQUINAS = Split(arrGRPMAQUINAS(I), "|")
           curTOTGERMAQ = 0
           For j = 0 To UBound(arrMAQUINAS) - 1
        
               arrDADOS = Split(arrMAQUINAS(j), ";")
               ccurPRODDIA = CCur(arrDADOS(1))
        
               sSql = "Select " & vbCrLf
               sSql = sSql & "       TUR.* " & vbCrLf
               sSql = sSql & "  From " & vbCrLf
               sSql = sSql & "       SGI_CADMAQUINA     MAQ " & vbCrLf
               sSql = sSql & "      ,SGI_CADFAMTURNO    FAM " & vbCrLf
               sSql = sSql & "      ,SGI_CADTURNSEM     TUR " & vbCrLf
               sSql = sSql & " Where " & vbCrLf
               sSql = sSql & "       MAQ.SGI_FILIAL     = " & FILIAL & vbCrLf
               sSql = sSql & "   And MAQ.SGI_CODIGO     = " & arrDADOS(0) & vbCrLf
               sSql = sSql & "   And FAM.SGI_FILIAL     = MAQ.SGI_FILIAL     " & vbCrLf
               sSql = sSql & "   And FAM.SGI_CODIGO     = MAQ.SGI_CODFAMILIA " & vbCrLf
               sSql = sSql & "   And TUR.SGI_FILIAL     = FAM.SGI_FILIAL     " & vbCrLf
               sSql = sSql & "   And TUR.SGI_CODIGO     = FAM.SGI_CODTURNO   " & vbCrLf
               sSql = sSql & "   And TUR.SGI_DIASEM     = " & intDiaSemana
        
               BREC.Open sSql, adoBanco_Dados, adOpenDynamic
               lngTotHora = 0
               If Not BREC.EOF() Then
                  Do While Not BREC.EOF()
                     lngTotHora = lngTotHora + CONVHRMIN(BREC!SGI_HORALIQ)
                     BREC.MoveNext
                  Loop
               
                  lngTotHora = (lngTotHora - PegaDescHoras(CLng(arrDADOS(0)), dtDataAtual))
               
                  '' ======================================
                  TOTMINUTOS = Round((lngTotHora / 60), 2)
                  curTOTPCTURNO = (TOTMINUTOS * ccurPRODDIA)
                  curTOTGERMAQ = (curTOTGERMAQ + curTOTPCTURNO)
                  '' ======================================
                  
                  If curTOTGERMAQ > 0 Then
                     grdProdMaqDiaIdeal.AddItem Format(dtDataAtual, "DD/MM/YYYY") & vbTab & _
                                                arrDIASSEMANA(intDiaSemana) & vbTab & _
                                                CONVMINHR(lngTotHora) & vbTab & _
                                                Format(ccurPRODDIA, "#,##0.00") & vbTab & _
                                                Format(curTOTPCTURNO, "#,##0.00") & vbTab & _
                                                Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                                arrDADOS(0)
                  End If
               
               End If
               BREC.Close
            
           Next j
           
           curTOTDEMANDA = (curTOTDEMANDA - curTOTGERMAQ)
           If curTOTGERMAQ > 0 Then
              grdPRODDIAIDEAL.AddItem Format(dtDataAtual, "DD/MM/YYYY") & vbTab & _
                                      arrDIASSEMANA(intDiaSemana) & vbTab & _
                                      CONVMINHR(lngTotHora) & vbTab & _
                                      Format(0, "#,##0.00") & vbTab & _
                                      Format(curTOTGERMAQ, "#,##0.00") & vbTab & _
                                      Format(curTOTDEMANDA, "#,##0.00") & vbTab & _
                                      arrDADOS(2)
                                     
              
              intINDICE = grdPROCESSOS.FindRow(arrDADOS(2), , conCOL_SonProcessos_Indice)
              If intINDICE <> -1 Then
                 grdPROCESSOS.Cell(flexcpText, intINDICE, conCOL_SonProcessos_DtPrevIdeal) = Format(dtDataAtual, "DD/MM/YYYY")
              End If
              
           End If
           
           dtDataAtual = (dtDataAtual + 1)
           intDiaSemana = Weekday(dtDataAtual)
           
        Loop
    
    Next I
    
    Call grdPROCESSOS_Click
End Sub


Private Function PegaDescHoras(lngCODMAQ As Long, dtData As Date) As Long

    PegaDescHoras = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_HORAINI    " & vbCrLf
    sSql = sSql & "      ,SGI_HORAFIN    " & vbCrLf
    sSql = sSql & "      ,SGI_TOTALUSADO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPERIODOMANUT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODMAQUINA = " & lngCODMAQ & vbCrLf
    sSql = sSql & "   And SGI_DATAMANUT  = '" & Format(dtData, "MM/DD/YYYY") & "'"

    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC2.EOF()
       PegaDescHoras = PegaDescHoras + CONVHRMIN(Format(BREC2!SGI_TOTALUSADO, "HH:MM"))
       BREC2.MoveNext
    Loop
    BREC2.Close
    
End Function

Private Sub InitGridProdMaqDiaOtim()

    With grdProdMaqDiaOtim
    
       .Cols = conColumnsIn_SonProdMaqDiaOtim
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonProdMaqDiaOtim_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaOtim_Data) = ""
       .ColDataType(conCOL_SonProdMaqDiaOtim_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaOtim_DiaSem) = ""
       .ColDataType(conCOL_SonProdMaqDiaOtim_DiaSem) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaOtim_HorDisp) = ""
       .ColDataType(conCOL_SonProdMaqDiaOtim_HorDisp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaOtim_TotProdHora) = ""
       .ColDataType(conCOL_SonProdMaqDiaOtim_TotProdHora) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaOtim_TotProdTurn) = ""
       .ColDataType(conCOL_SonProdMaqDiaOtim_TotProdTurn) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaOtim_Saldo) = ""
       .ColDataType(conCOL_SonProdMaqDiaOtim_Saldo) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaOtim_Pai) = ""
       .ColDataType(conCOL_SonProdMaqDiaOtim_Pai) = flexDTLong
       
       .ColWidth(conCOL_SonProdMaqDiaOtim_Data) = 1000
       .ColWidth(conCOL_SonProdMaqDiaOtim_DiaSem) = 1500
       .ColWidth(conCOL_SonProdMaqDiaOtim_HorDisp) = 1000
       .ColWidth(conCOL_SonProdMaqDiaOtim_TotProdHora) = 1000
       .ColWidth(conCOL_SonProdMaqDiaOtim_TotProdTurn) = 1000
       .ColWidth(conCOL_SonProdMaqDiaOtim_Saldo) = 1000
       .ColWidth(conCOL_SonProdMaqDiaOtim_Pai) = 1000
       
       .ColHidden(conCOL_SonProdMaqDiaOtim_Pai) = True
       .ColHidden(conCOL_SonProdMaqDiaOtim_Saldo) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub InitGridProdMaqDiaIdeal()

    With grdProdMaqDiaIdeal
    
       .Cols = conColumnsIn_SonProdMaqDiaIdeal
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonProdMaqDiaIdeal_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaIdeal_Data) = ""
       .ColDataType(conCOL_SonProdMaqDiaIdeal_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaIdeal_DiaSem) = ""
       .ColDataType(conCOL_SonProdMaqDiaIdeal_DiaSem) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaIdeal_HorDisp) = ""
       .ColDataType(conCOL_SonProdMaqDiaIdeal_HorDisp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaIdeal_TotProdHora) = ""
       .ColDataType(conCOL_SonProdMaqDiaIdeal_TotProdHora) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaIdeal_TotProdTurn) = ""
       .ColDataType(conCOL_SonProdMaqDiaIdeal_TotProdTurn) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaIdeal_Saldo) = ""
       .ColDataType(conCOL_SonProdMaqDiaIdeal_Saldo) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaIdeal_Pai) = ""
       .ColDataType(conCOL_SonProdMaqDiaIdeal_Pai) = flexDTLong
       
       .ColWidth(conCOL_SonProdMaqDiaIdeal_Data) = 1000
       .ColWidth(conCOL_SonProdMaqDiaIdeal_DiaSem) = 1500
       .ColWidth(conCOL_SonProdMaqDiaIdeal_HorDisp) = 1000
       .ColWidth(conCOL_SonProdMaqDiaIdeal_TotProdHora) = 1000
       .ColWidth(conCOL_SonProdMaqDiaIdeal_TotProdTurn) = 1000
       .ColWidth(conCOL_SonProdMaqDiaIdeal_Saldo) = 1000
       .ColWidth(conCOL_SonProdMaqDiaIdeal_Pai) = 1000
       
       .ColHidden(conCOL_SonProdMaqDiaIdeal_Pai) = True
       .ColHidden(conCOL_SonProdMaqDiaIdeal_Saldo) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub InitGridProdMaqDiaPess()

    With grdProdMaqDiaPess
    
       .Cols = conColumnsIn_SonProdMaqDiaPess
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_strSonProdMaqDiaPess_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaPess_Data) = ""
       .ColDataType(conCOL_SonProdMaqDiaPess_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaPess_DiaSem) = ""
       .ColDataType(conCOL_SonProdMaqDiaPess_DiaSem) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaPess_HorDisp) = ""
       .ColDataType(conCOL_SonProdMaqDiaPess_HorDisp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaPess_TotProdHora) = ""
       .ColDataType(conCOL_SonProdMaqDiaPess_TotProdHora) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaPess_TotProdTurn) = ""
       .ColDataType(conCOL_SonProdMaqDiaPess_TotProdTurn) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaPess_Saldo) = ""
       .ColDataType(conCOL_SonProdMaqDiaPess_Saldo) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdMaqDiaPess_Pai) = ""
       .ColDataType(conCOL_SonProdMaqDiaPess_Pai) = flexDTLong
       
       .ColWidth(conCOL_SonProdMaqDiaPess_Data) = 1000
       .ColWidth(conCOL_SonProdMaqDiaPess_DiaSem) = 1500
       .ColWidth(conCOL_SonProdMaqDiaPess_HorDisp) = 1000
       .ColWidth(conCOL_SonProdMaqDiaPess_TotProdHora) = 1000
       .ColWidth(conCOL_SonProdMaqDiaPess_TotProdTurn) = 1000
       .ColWidth(conCOL_SonProdMaqDiaPess_Saldo) = 1000
       .ColWidth(conCOL_SonProdMaqDiaPess_Pai) = 1000
       
       .ColHidden(conCOL_SonProdMaqDiaPess_Pai) = True
       .ColHidden(conCOL_SonProdMaqDiaPess_Saldo) = True
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Function PegaDescProd(strProdutoID As String) As String
    PegaDescProd = ""
    
    If Len(Trim(strProdutoID)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_DESCRI " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO  = " & strProdutoID
    
    BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC5.EOF() Then PegaDescProd = Trim(BREC5!SGI_DESCRI)
    BREC5.Close
    
End Function

Private Sub LimpaCamposLabel()
    lblDescProd.Caption = ""
End Sub

Private Sub DestroyObjeto()
    Set objBLBFunc = Nothing
    Set objCADFLUXOPROD = Nothing
End Sub

Private Sub DeabilitaCampos()
    '' Desativando estes Controles
    stSenarios.TabVisible(1) = False
    stSenarios.TabVisible(2) = False
    Frame2.Visible = False
    Frame4.Visible = False
    Label1(3).Visible = False
    '' ---------------------------
End Sub
