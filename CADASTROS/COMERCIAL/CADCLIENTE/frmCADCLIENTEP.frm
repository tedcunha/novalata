VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADCLIENTE 
   Caption         =   "Cadastro de clientes"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   13515
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   178
      Top             =   120
      Width           =   13455
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
         Picture         =   "frmCADCLIENTEP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   181
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
         Picture         =   "frmCADCLIENTEP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   180
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
         Picture         =   "frmCADCLIENTEP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   240
         Width           =   855
      End
   End
   Begin TabDlg.SSTab stCLIENTE 
      Height          =   6495
      Left            =   0
      TabIndex        =   63
      Top             =   1080
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   16
      TabsPerRow      =   10
      TabHeight       =   520
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cadastrais"
      TabPicture(0)   =   "frmCADCLIENTEP.frx":0736
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cobrança/Entrega"
      TabPicture(1)   =   "frmCADCLIENTEP.frx":0752
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Informações"
      TabPicture(2)   =   "frmCADCLIENTEP.frx":076E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "stFin"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Crédito"
      TabPicture(3)   =   "frmCADCLIENTEP.frx":078A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Frame24"
      Tab(3).Control(2)=   "Frame7"
      Tab(3).Control(3)=   "Frame25"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Financeiro"
      TabPicture(4)   =   "frmCADCLIENTEP.frx":07A6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "stDuplicatas"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Folhow UP"
      TabPicture(5)   =   "frmCADCLIENTEP.frx":07C2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame26"
      Tab(5).Control(1)=   "Frame27"
      Tab(5).Control(2)=   "Frame28"
      Tab(5).Control(3)=   "Frame29"
      Tab(5).Control(4)=   "Frame30"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Obs - Ass. Tec."
      TabPicture(6)   =   "frmCADCLIENTEP.frx":07DE
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame31"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Despesas"
      TabPicture(7)   =   "frmCADCLIENTEP.frx":07FA
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Transportadora"
      TabPicture(8)   =   "frmCADCLIENTEP.frx":0816
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame40"
      Tab(8).Control(1)=   "Frame41"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Técnicos"
      TabPicture(9)   =   "frmCADCLIENTEP.frx":0832
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Vendedores"
      TabPicture(10)  =   "frmCADCLIENTEP.frx":084E
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame43"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Obs - Comercial"
      TabPicture(11)  =   "frmCADCLIENTEP.frx":086A
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame44"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "Produtos"
      TabPicture(12)  =   "frmCADCLIENTEP.frx":0886
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "grdPRODUTOS"
      Tab(12).Control(1)=   "Command26"
      Tab(12).Control(2)=   "Command27"
      Tab(12).Control(3)=   "grdULTFAT"
      Tab(12).Control(4)=   "Frame45"
      Tab(12).Control(5)=   "Frame46"
      Tab(12).ControlCount=   6
      TabCaption(13)  =   "Faturamento"
      TabPicture(13)  =   "frmCADCLIENTEP.frx":08A2
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "SSTab1"
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "Curva ABC"
      TabPicture(14)  =   "frmCADCLIENTEP.frx":08BE
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "grdPRODABC"
      Tab(14).Control(1)=   "Frame47"
      Tab(14).Control(2)=   "Frame50"
      Tab(14).Control(3)=   "Frame54"
      Tab(14).Control(4)=   "Frame55"
      Tab(14).ControlCount=   5
      TabCaption(15)  =   "Condições de Pagamento"
      TabPicture(15)  =   "frmCADCLIENTEP.frx":08DA
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "grdCONDPGTO"
      Tab(15).Control(1)=   "Command11"
      Tab(15).Control(2)=   "Command12"
      Tab(15).ControlCount=   3
      Begin VB.CommandButton Command12 
         Height          =   300
         Left            =   -66360
         Picture         =   "frmCADCLIENTEP.frx":08F6
         Style           =   1  'Graphical
         TabIndex        =   273
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   720
         Width           =   300
      End
      Begin VB.CommandButton Command11 
         Height          =   300
         Left            =   -66360
         Picture         =   "frmCADCLIENTEP.frx":0A40
         Style           =   1  'Graphical
         TabIndex        =   272
         ToolTipText     =   "Exclui a linha da Gride Selecionada"
         Top             =   1080
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdCONDPGTO 
         Height          =   5535
         Left            =   -74760
         TabIndex        =   271
         Top             =   720
         Width           =   8295
         _cx             =   14631
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
      Begin VB.Frame Frame55 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74880
         TabIndex        =   269
         Top             =   2880
         Width           =   13095
         Begin VSFlex8LCtl.VSFlexGrid grdCURVAABC 
            Height          =   1095
            Left            =   120
            TabIndex        =   270
            Top             =   240
            Width           =   12855
            _cx             =   22675
            _cy             =   1931
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
      Begin VB.Frame Frame54 
         Caption         =   "[ Faturado ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -68280
         TabIndex        =   266
         Top             =   4440
         Width           =   6495
         Begin VSFlex8LCtl.VSFlexGrid grdFaturado 
            Height          =   1335
            Left            =   120
            TabIndex        =   268
            Top             =   360
            Width           =   6255
            _cx             =   11033
            _cy             =   2355
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
      Begin VB.Frame Frame50 
         Caption         =   "[ Pedidos ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74880
         TabIndex        =   265
         Top             =   4440
         Width           =   6495
         Begin VSFlex8LCtl.VSFlexGrid grdPedidos 
            Height          =   1335
            Left            =   120
            TabIndex        =   267
            Top             =   360
            Width           =   6255
            _cx             =   11033
            _cy             =   2355
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
      Begin VB.Frame Frame47 
         Caption         =   "[ Empresa ]"
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
         Left            =   -65040
         TabIndex        =   262
         Top             =   600
         Width           =   3255
         Begin VB.OptionButton optEMP 
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
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   264
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optEMP 
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   263
            Top             =   240
            Width           =   1335
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid grdPRODABC 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   261
         Top             =   720
         Width           =   9735
         _cx             =   17171
         _cy             =   3625
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
      Begin VB.Frame Frame46 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -70320
         TabIndex        =   258
         Top             =   720
         Width           =   3975
         Begin VB.OptionButton optSTATUS 
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
            Left            =   1440
            TabIndex        =   260
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optSTATUS 
            Caption         =   "Faturados"
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
            TabIndex        =   259
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Frame Frame45 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -75000
         TabIndex        =   255
         Top             =   720
         Width           =   3615
         Begin VB.OptionButton optFILIALPRODFAT 
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
            Left            =   1800
            TabIndex        =   257
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optFILIALPRODFAT 
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
            TabIndex        =   256
            Top             =   0
            Width           =   1695
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid grdULTFAT 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   254
         Top             =   3360
         Width           =   12615
         _cx             =   22251
         _cy             =   4683
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
      Begin VB.CommandButton Command27 
         Height          =   300
         Left            =   -62160
         Picture         =   "frmCADCLIENTEP.frx":0B8A
         Style           =   1  'Graphical
         TabIndex        =   253
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton Command26 
         Height          =   300
         Left            =   -62160
         Picture         =   "frmCADCLIENTEP.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   252
         ToolTipText     =   "Exclui a linha da Gride Selecionada"
         Top             =   1440
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdPRODUTOS 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   251
         Top             =   1080
         Width           =   12615
         _cx             =   22251
         _cy             =   3836
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   233
         Top             =   720
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   9340
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         Tab             =   4
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
         TabCaption(0)   =   "Cotações"
         TabPicture(0)   =   "frmCADCLIENTEP.frx":0E1E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Pedidos"
         TabPicture(1)   =   "frmCADCLIENTEP.frx":0E3A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame48"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Faturamento"
         TabPicture(2)   =   "frmCADCLIENTEP.frx":0E56
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame49"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Resumo"
         TabPicture(3)   =   "frmCADCLIENTEP.frx":0E72
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame53"
         Tab(3).Control(1)=   "Frame52"
         Tab(3).Control(2)=   "Frame51"
         Tab(3).ControlCount=   3
         TabCaption(4)   =   "Produtos Faturados"
         TabPicture(4)   =   "frmCADCLIENTEP.frx":0E8E
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "grdITENSPEDIDO"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "Frame42"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).ControlCount=   2
         Begin VB.Frame Frame42 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            TabIndex        =   248
            Top             =   360
            Width           =   3975
            Begin VB.OptionButton optFATNOVSTEEL 
               Caption         =   "STEEL ROL"
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
               Left            =   1560
               TabIndex        =   250
               Top             =   0
               Width           =   1575
            End
            Begin VB.OptionButton optFATNOVSTEEL 
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
               Index           =   1
               Left            =   0
               TabIndex        =   249
               Top             =   0
               Width           =   1455
            End
         End
         Begin VSFlex8LCtl.VSFlexGrid grdITENSPEDIDO 
            Height          =   4455
            Left            =   120
            TabIndex        =   244
            Top             =   720
            Width           =   12855
            _cx             =   22675
            _cy             =   7858
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
         Begin VB.Frame Frame53 
            Caption         =   "[ Total Faturado ]"
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
            Height          =   4815
            Left            =   -67320
            TabIndex        =   240
            Top             =   360
            Width           =   3975
         End
         Begin VB.Frame Frame52 
            Caption         =   "[ Total de Pedidos ]"
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
            Height          =   4815
            Left            =   -71400
            TabIndex        =   239
            Top             =   360
            Width           =   3975
         End
         Begin VB.Frame Frame51 
            Caption         =   "[ Total de Cotações ]"
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
            Height          =   4815
            Left            =   -74880
            TabIndex        =   238
            Top             =   360
            Width           =   3375
         End
         Begin VB.Frame Frame49 
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
            Height          =   4815
            Left            =   -74880
            TabIndex        =   236
            Top             =   360
            Width           =   12855
            Begin MSFlexGridLib.MSFlexGrid flxFaturamento 
               Height          =   4095
               Left            =   120
               TabIndex        =   237
               Top             =   600
               Width           =   12615
               _ExtentX        =   22251
               _ExtentY        =   7223
               _Version        =   393216
               FixedCols       =   0
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame48 
            Caption         =   "[ Classificação por Ordem de Pedidos ]"
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
            Height          =   4815
            Left            =   -74880
            TabIndex        =   234
            Top             =   360
            Width           =   12855
            Begin VB.Frame Frame34 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   120
               TabIndex        =   245
               Top             =   240
               Width           =   3255
               Begin VB.OptionButton optTipoNOVASTEL 
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
                  Index           =   1
                  Left            =   0
                  TabIndex        =   247
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.OptionButton optTipoNOVASTEL 
                  Caption         =   "STEEL ROL"
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
                  Left            =   1440
                  TabIndex        =   246
                  Top             =   0
                  Width           =   1455
               End
            End
            Begin MSFlexGridLib.MSFlexGrid flxPedidos 
               Height          =   4095
               Left            =   120
               TabIndex        =   235
               Top             =   600
               Width           =   12615
               _ExtentX        =   22251
               _ExtentY        =   7223
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
      End
      Begin VB.Frame Frame44 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   231
         Top             =   720
         Width           =   9015
         Begin VB.TextBox txtObsCom 
            Height          =   4095
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   232
            Top             =   240
            Width           =   8775
         End
      End
      Begin VB.Frame Frame43 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   227
         Top             =   720
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxVENDEDORES 
            Height          =   4095
            Left            =   120
            TabIndex        =   228
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   7223
            _Version        =   393216
            FixedCols       =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame41 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   225
         Top             =   1440
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxTRANSP 
            Height          =   3375
            Left            =   120
            TabIndex        =   226
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   5953
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame40 
         Height          =   735
         Left            =   -74880
         TabIndex        =   219
         Top             =   720
         Width           =   9015
         Begin VB.CommandButton cmdIncTransp 
            Height          =   315
            Left            =   8160
            Picture         =   "frmCADCLIENTEP.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   223
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdTRANSP 
            Height          =   315
            Left            =   2760
            Picture         =   "frmCADCLIENTEP.frx":0FAC
            Style           =   1  'Graphical
            TabIndex        =   222
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboTRANSP 
            Height          =   315
            Left            =   3120
            TabIndex        =   221
            Text            =   "cboTRANSP"
            Top             =   240
            Width           =   5055
         End
         Begin VB.TextBox txtCODTRANSP 
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   220
            Text            =   "txtCODTRAN"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Transportadoras:"
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
            Index           =   2
            Left            =   240
            TabIndex        =   224
            Top             =   240
            Width           =   1455
         End
      End
      Begin TabDlg.SSTab stDuplicatas 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   182
         Top             =   720
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7858
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   -2147483638
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
         TabCaption(0)   =   "Duplicatas"
         TabPicture(0)   =   "frmCADCLIENTEP.frx":10AE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "StHistDupl"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Resumo"
         TabPicture(1)   =   "frmCADCLIENTEP.frx":10CA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame37"
         Tab(1).ControlCount=   1
         Begin TabDlg.SSTab StHistDupl 
            Height          =   3855
            Left            =   120
            TabIndex        =   194
            Top             =   480
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   6800
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
            TabCaption(0)   =   "Duplicatas á Vencer"
            TabPicture(0)   =   "frmCADCLIENTEP.frx":10E6
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame20"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Duplicatas Vencidas"
            TabPicture(1)   =   "frmCADCLIENTEP.frx":1102
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame35"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Duplicatas Pagas"
            TabPicture(2)   =   "frmCADCLIENTEP.frx":111E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "StHistDupli"
            Tab(2).ControlCount=   1
            Begin TabDlg.SSTab StHistDupli 
               Height          =   3375
               Left            =   -74880
               TabIndex        =   201
               Top             =   360
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   5953
               _Version        =   393216
               Style           =   1
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
               TabCaption(0)   =   "Pagas Antecipado"
               TabPicture(0)   =   "frmCADCLIENTEP.frx":113A
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Frame36"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Pagas no Prazo"
               TabPicture(1)   =   "frmCADCLIENTEP.frx":1156
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Frame39"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "Pagas com Atrazo"
               TabPicture(2)   =   "frmCADCLIENTEP.frx":1172
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Frame38"
               Tab(2).ControlCount=   1
               Begin VB.Frame Frame39 
                  Height          =   2895
                  Left            =   -74880
                  TabIndex        =   210
                  Top             =   360
                  Width           =   8295
                  Begin MSFlexGridLib.MSFlexGrid flxTitPgtoPrazo 
                     Height          =   2535
                     Left            =   120
                     TabIndex        =   211
                     Top             =   240
                     Width           =   8055
                     _ExtentX        =   14208
                     _ExtentY        =   4471
                     _Version        =   393216
                     FixedCols       =   0
                     HighLight       =   2
                     SelectionMode   =   1
                     Appearance      =   0
                  End
               End
               Begin VB.Frame Frame38 
                  Height          =   2895
                  Left            =   -74880
                  TabIndex        =   204
                  Top             =   360
                  Width           =   8295
                  Begin MSFlexGridLib.MSFlexGrid flxDpPgtAtrazo 
                     Height          =   2535
                     Left            =   120
                     TabIndex        =   205
                     Top             =   240
                     Width           =   8055
                     _ExtentX        =   14208
                     _ExtentY        =   4471
                     _Version        =   393216
                     FixedCols       =   0
                     HighLight       =   2
                     SelectionMode   =   1
                     Appearance      =   0
                  End
               End
               Begin VB.Frame Frame36 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2895
                  Left            =   120
                  TabIndex        =   202
                  Top             =   360
                  Width           =   8295
                  Begin MSFlexGridLib.MSFlexGrid flxPgtAntecipado 
                     Height          =   2535
                     Left            =   120
                     TabIndex        =   203
                     Top             =   240
                     Width           =   8055
                     _ExtentX        =   14208
                     _ExtentY        =   4471
                     _Version        =   393216
                     FixedCols       =   0
                     HighLight       =   2
                     SelectionMode   =   1
                     Appearance      =   0
                  End
               End
            End
            Begin VB.Frame Frame35 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3375
               Left            =   -74880
               TabIndex        =   197
               Top             =   360
               Width           =   8535
               Begin MSFlexGridLib.MSFlexGrid flxDuplVencidas 
                  Height          =   3015
                  Left            =   120
                  TabIndex        =   198
                  Top             =   240
                  Width           =   8295
                  _ExtentX        =   14631
                  _ExtentY        =   5318
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
                  Appearance      =   0
               End
            End
            Begin VB.Frame Frame20 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3375
               Left            =   120
               TabIndex        =   195
               Top             =   360
               Width           =   8535
               Begin MSFlexGridLib.MSFlexGrid flxDuplApgt 
                  Height          =   3015
                  Left            =   120
                  TabIndex        =   196
                  Top             =   240
                  Width           =   8295
                  _ExtentX        =   14631
                  _ExtentY        =   5318
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
                  Appearance      =   0
               End
            End
         End
         Begin VB.Frame Frame37 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   183
            Top             =   360
            Width           =   8775
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   14
               Left            =   5760
               TabIndex        =   218
               Top             =   3240
               Width           =   930
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   13
               Left            =   5760
               TabIndex        =   217
               Top             =   2880
               Width           =   930
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   12
               Left            =   5760
               TabIndex        =   216
               Top             =   2520
               Width           =   930
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   11
               Left            =   5760
               TabIndex        =   215
               Top             =   1440
               Width           =   930
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   10
               Left            =   5760
               TabIndex        =   214
               Top             =   1080
               Width           =   930
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   5760
               TabIndex        =   213
               Top             =   2160
               Width           =   930
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   8
               Left            =   5760
               TabIndex        =   212
               Top             =   720
               Width           =   930
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   7
               Left            =   3600
               TabIndex        =   209
               Top             =   2880
               Width           =   2010
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Total de duplicatas pagas no prazo:"
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
               Index           =   7
               Left            =   440
               TabIndex        =   208
               Top             =   2900
               Width           =   3075
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   6
               Left            =   3600
               TabIndex        =   207
               Top             =   3240
               Width           =   2010
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Total de duplicatas pagas atrazado:"
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
               Index           =   6
               Left            =   435
               TabIndex        =   206
               Top             =   3260
               Width           =   3075
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   5
               Left            =   3600
               TabIndex        =   200
               Top             =   240
               Width           =   2010
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Toptal faturado:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   5
               Left            =   2050
               TabIndex        =   199
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   3600
               TabIndex        =   193
               Top             =   2160
               Width           =   2010
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Toptal de duplicatas Pagas:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   1080
               TabIndex        =   192
               Top             =   2190
               Width           =   2400
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   3
               Left            =   3600
               TabIndex        =   191
               Top             =   2520
               Width           =   2010
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Total de duplicatas pagas antecipado:"
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
               Left            =   240
               TabIndex        =   190
               Top             =   2540
               Width           =   3285
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   2
               Left            =   3600
               TabIndex        =   189
               Top             =   1440
               Width           =   2010
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Total de duplicatas vencidas:"
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
               Left            =   915
               TabIndex        =   188
               Top             =   1455
               Width           =   2535
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   0
               Left            =   3600
               TabIndex        =   187
               Top             =   1080
               Width           =   2010
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Total de duplicatas a vencer:"
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
               Left            =   915
               TabIndex        =   186
               Top             =   1110
               Width           =   2520
            End
            Begin VB.Label lblValoresDupl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   3600
               TabIndex        =   185
               Top             =   720
               Width           =   2010
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Toptal de duplicatas em Aberto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   184
               Top             =   720
               Width           =   2730
            End
         End
      End
      Begin VB.Frame Frame31 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   168
         Top             =   720
         Width           =   9015
         Begin VB.TextBox txtOBS 
            Height          =   4095
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   169
            Top             =   240
            Width           =   8775
         End
      End
      Begin VB.Frame Frame30 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   163
         Top             =   3720
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxATENDIDO 
            Height          =   1095
            Left            =   120
            TabIndex        =   164
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   1931
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame29 
         Height          =   615
         Left            =   -74880
         TabIndex        =   160
         Top             =   3120
         Width           =   9015
         Begin VB.CommandButton Command9 
            Height          =   315
            Left            =   8160
            Picture         =   "frmCADCLIENTEP.frx":118E
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboSEGMENTO2 
            Height          =   315
            Left            =   5520
            TabIndex        =   61
            Text            =   "cboSEGMENTO2"
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtATENDIDO 
            Height          =   285
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   60
            Text            =   "txtATENDIDO"
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Segmento :"
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
            Left            =   4440
            TabIndex        =   162
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Atendido :"
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
            Left            =   120
            TabIndex        =   161
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame28 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   158
         Top             =   1920
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxSISTSERTFIC 
            Height          =   855
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   1508
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame27 
         Height          =   615
         Left            =   -74880
         TabIndex        =   156
         Top             =   1320
         Width           =   9015
         Begin VB.CommandButton Command8 
            Height          =   315
            Left            =   7320
            Picture         =   "frmCADCLIENTEP.frx":16C0
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtSISTCERTIFI 
            Height          =   285
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   58
            Text            =   "txtSISTCERTIFI"
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Sistemas de certificação :"
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
            Left            =   120
            TabIndex        =   157
            Top             =   240
            Width           =   2220
         End
      End
      Begin VB.Frame Frame26 
         Height          =   615
         Left            =   -74880
         TabIndex        =   153
         Top             =   720
         Width           =   9015
         Begin VB.CommandButton Command10 
            Height          =   315
            Left            =   1920
            Picture         =   "frmCADCLIENTEP.frx":1BF2
            Style           =   1  'Graphical
            TabIndex        =   167
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton optTEMCERTINAO 
            Caption         =   "Nâo"
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
            Left            =   7800
            TabIndex        =   166
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optTEMCERTISIM 
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
            Left            =   7080
            TabIndex        =   165
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboSEGMENTO 
            Height          =   315
            Left            =   2280
            TabIndex        =   57
            Text            =   "cboSEGMENTO"
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtCODSEQ 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   56
            Text            =   "txtCODSEQ"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "É certificado :"
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
            Left            =   5880
            TabIndex        =   155
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Segmento :"
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
            Left            =   120
            TabIndex        =   154
            Top             =   260
            Width           =   975
         End
      End
      Begin VB.Frame Frame25 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   150
         Top             =   3240
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxSTATUSAVALIACAO 
            Height          =   1575
            Left            =   2280
            TabIndex        =   152
            Top             =   240
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            Appearance      =   0
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Histórico de avaliações :"
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
            Left            =   120
            TabIndex        =   151
            Top             =   240
            Width           =   2130
         End
      End
      Begin VB.Frame Frame7 
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
         Left            =   -74880
         TabIndex        =   145
         Top             =   2400
         Width           =   9015
         Begin VB.OptionButton optBloqNao 
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
            Left            =   7560
            TabIndex        =   54
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optBloqSim 
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
            Left            =   6720
            TabIndex        =   53
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtBLOQPEDSALDO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            TabIndex        =   55
            Text            =   "txtBLOQPEDSALDO"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtSALDOACIMA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            TabIndex        =   52
            Text            =   "txtSALDOACIMA"
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Bloquear ? :"
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
            Left            =   5520
            TabIndex        =   148
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Bloqueia com saldo a cima de :"
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
            Left            =   120
            TabIndex        =   147
            Top             =   525
            Width           =   2670
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Avisar quando o saldo estiver a cima de :"
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
            Left            =   120
            TabIndex        =   146
            Top             =   225
            Width           =   3540
         End
      End
      Begin VB.Frame Frame24 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   143
         Top             =   1320
         Width           =   9015
         Begin VB.OptionButton optSempNAO 
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
            Left            =   3720
            TabIndex        =   51
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optSempSIM 
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
            Left            =   2520
            TabIndex        =   50
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtMeseReavali 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   49
            Text            =   "txtMeseReavali"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Sempre Bloquear pedido :"
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
            Left            =   120
            TabIndex        =   149
            Top             =   720
            Width           =   2205
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Meses p/reavaliação :"
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
            Left            =   120
            TabIndex        =   144
            Top             =   285
            Width           =   1920
         End
      End
      Begin TabDlg.SSTab stFin 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   96
         Top             =   720
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7858
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
         TabCaption(0)   =   "Bancos"
         TabPicture(0)   =   "frmCADCLIENTEP.frx":1CF4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame23"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Referências"
         TabPicture(1)   =   "frmCADCLIENTEP.frx":1D10
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "stREFDIVER"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Restrições"
         TabPicture(2)   =   "frmCADCLIENTEP.frx":1D2C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame19"
         Tab(2).Control(1)=   "Frame18"
         Tab(2).Control(2)=   "Frame17"
         Tab(2).Control(3)=   "Frame16"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Fornecedores"
         TabPicture(3)   =   "frmCADCLIENTEP.frx":1D48
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame8"
         Tab(3).Control(1)=   "Frame21"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Clientes"
         TabPicture(4)   =   "frmCADCLIENTEP.frx":1D64
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame9"
         Tab(4).Control(1)=   "Frame22"
         Tab(4).ControlCount=   2
         Begin VB.Frame Frame9 
            Height          =   615
            Left            =   -74880
            TabIndex        =   140
            Top             =   360
            Width           =   8775
            Begin VB.TextBox txtCLIENTE 
               Height          =   285
               Left            =   960
               TabIndex        =   45
               Text            =   "txtCLIENTE"
               Top             =   220
               Width           =   4575
            End
            Begin VB.CommandButton Command2 
               Height          =   315
               Left            =   7560
               Picture         =   "frmCADCLIENTEP.frx":1D80
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   240
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskDTCLIENTE 
               Height          =   285
               Left            =   6360
               TabIndex        =   46
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               Caption         =   "Desde:"
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
               Left            =   5640
               TabIndex        =   142
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "Empresa:"
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
               Left            =   120
               TabIndex        =   141
               Top             =   240
               Width           =   795
            End
         End
         Begin VB.Frame Frame22 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   138
            Top             =   960
            Width           =   8775
            Begin MSFlexGridLib.MSFlexGrid flxEMPRESA 
               Height          =   3015
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   5318
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame8 
            Height          =   615
            Left            =   -74880
            TabIndex        =   135
            Top             =   360
            Width           =   8775
            Begin VB.TextBox txtFORNEC 
               Height          =   285
               Left            =   960
               TabIndex        =   42
               Text            =   "txtFORNEC"
               Top             =   220
               Width           =   4575
            End
            Begin VB.CommandButton Command1 
               Height          =   315
               Left            =   7560
               Picture         =   "frmCADCLIENTEP.frx":22B2
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   240
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskDTEMPRESA 
               Height          =   285
               Left            =   6360
               TabIndex        =   43
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "Empresa:"
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
               Left            =   120
               TabIndex        =   137
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               Caption         =   "Desde:"
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
               Left            =   5640
               TabIndex        =   136
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame21 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   133
            Top             =   960
            Width           =   8775
            Begin MSFlexGridLib.MSFlexGrid flxFORNEC 
               Height          =   3015
               Left            =   120
               TabIndex        =   134
               Top             =   240
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   5318
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame23 
            Height          =   3375
            Left            =   120
            TabIndex        =   129
            Top             =   960
            Width           =   8775
            Begin MSFlexGridLib.MSFlexGrid flxBANCOS 
               Height          =   3015
               Left            =   120
               TabIndex        =   130
               Top             =   240
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   5318
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame19 
            Height          =   615
            Left            =   -69960
            TabIndex        =   125
            Top             =   3720
            Width           =   3855
            Begin VB.OptionButton optAVISANAO 
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
               Left            =   3000
               TabIndex        =   128
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optAVISASIM 
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
               Left            =   2040
               TabIndex        =   127
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "Avisar sobre restrições :"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   126
               Top             =   240
               Width           =   1680
            End
         End
         Begin VB.Frame Frame18 
            Height          =   615
            Left            =   -74880
            TabIndex        =   121
            Top             =   3720
            Width           =   4095
            Begin VB.OptionButton optBLPEDNAO 
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
               Left            =   3000
               TabIndex        =   124
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optBLPEDSIM 
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
               Left            =   1920
               TabIndex        =   123
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Bloquear pedidos:"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   122
               Top             =   240
               Width           =   1275
            End
         End
         Begin VB.Frame Frame17 
            Height          =   615
            Left            =   -74880
            TabIndex        =   119
            Top             =   360
            Width           =   8775
            Begin VB.CommandButton Command7 
               Height          =   315
               Left            =   6600
               Picture         =   "frmCADCLIENTEP.frx":27E4
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtRESTRICOES 
               Height          =   285
               Left            =   1080
               TabIndex        =   40
               Text            =   "txtRESTRICOES"
               Top             =   240
               Width           =   5535
            End
            Begin VB.Label Label40 
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
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   120
               Top             =   240
               Width           =   930
            End
         End
         Begin VB.Frame Frame16 
            Height          =   2655
            Left            =   -74880
            TabIndex        =   117
            Top             =   960
            Width           =   8775
            Begin MSFlexGridLib.MSFlexGrid flxRESTRICOES 
               Height          =   2295
               Left            =   120
               TabIndex        =   118
               Top             =   240
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   4048
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin TabDlg.SSTab stREFDIVER 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   99
            Top             =   360
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   7011
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
            TabCaption(0)   =   "Bancária"
            TabPicture(0)   =   "frmCADCLIENTEP.frx":2D16
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame10"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame11"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Comercial"
            TabPicture(1)   =   "frmCADCLIENTEP.frx":2D32
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame13"
            Tab(1).Control(1)=   "Frame12"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Pessoal"
            TabPicture(2)   =   "frmCADCLIENTEP.frx":2D4E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame15"
            Tab(2).Control(1)=   "Frame14"
            Tab(2).ControlCount=   2
            Begin VB.Frame Frame15 
               Height          =   615
               Left            =   -74880
               TabIndex        =   114
               Top             =   360
               Width           =   8535
               Begin VB.CommandButton Command5 
                  Height          =   315
                  Left            =   8040
                  Picture         =   "frmCADCLIENTEP.frx":2D6A
                  Style           =   1  'Graphical
                  TabIndex        =   39
                  Top             =   200
                  Width           =   375
               End
               Begin VB.TextBox txtREFTELPESSOAL 
                  Height          =   285
                  Left            =   6480
                  TabIndex        =   38
                  Text            =   "txtREFTELPESSOAL"
                  Top             =   200
                  Width           =   1575
               End
               Begin VB.TextBox txtREFPESSOAL 
                  Height          =   285
                  Left            =   720
                  TabIndex        =   37
                  Text            =   "txtREFPESSOAL"
                  Top             =   200
                  Width           =   4695
               End
               Begin VB.Label Label39 
                  AutoSize        =   -1  'True
                  Caption         =   "Telefone:"
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
                  Left            =   5520
                  TabIndex        =   116
                  Top             =   195
                  Width           =   825
               End
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "Nome:"
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
                  Left            =   120
                  TabIndex        =   115
                  Top             =   200
                  Width           =   555
               End
            End
            Begin VB.Frame Frame14 
               Height          =   2895
               Left            =   -74880
               TabIndex        =   112
               Top             =   960
               Width           =   8535
               Begin MSFlexGridLib.MSFlexGrid flxREFPESSOAL 
                  Height          =   2535
                  Left            =   120
                  TabIndex        =   113
                  Top             =   240
                  Width           =   8295
                  _ExtentX        =   14631
                  _ExtentY        =   4471
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
                  Appearance      =   0
               End
            End
            Begin VB.Frame Frame13 
               Height          =   735
               Left            =   -74880
               TabIndex        =   108
               Top             =   360
               Width           =   8535
               Begin VB.CommandButton Command4 
                  Height          =   315
                  Left            =   7320
                  Picture         =   "frmCADCLIENTEP.frx":329C
                  Style           =   1  'Graphical
                  TabIndex        =   36
                  Top             =   360
                  Width           =   375
               End
               Begin VB.TextBox txtREFCOM 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   33
                  Text            =   "txtREFCOM"
                  Top             =   360
                  Width           =   4095
               End
               Begin VB.TextBox txtREFCOMNOME 
                  Height          =   285
                  Left            =   4320
                  TabIndex        =   34
                  Text            =   "txtREFCOMNOME"
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txtTELREFCOM 
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   35
                  Text            =   "txtTELREFCOM"
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  Caption         =   "Empresa:"
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
                  Left            =   120
                  TabIndex        =   111
                  Top             =   120
                  Width           =   795
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  Caption         =   "Nome:"
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
                  Left            =   4320
                  TabIndex        =   110
                  Top             =   120
                  Width           =   555
               End
               Begin VB.Label Label33 
                  AutoSize        =   -1  'True
                  Caption         =   "Telefone:"
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
                  Left            =   5880
                  TabIndex        =   109
                  Top             =   120
                  Width           =   825
               End
            End
            Begin VB.Frame Frame12 
               Height          =   2775
               Left            =   -74880
               TabIndex        =   106
               Top             =   1080
               Width           =   8535
               Begin MSFlexGridLib.MSFlexGrid flxREFEMPRESA 
                  Height          =   2415
                  Left            =   120
                  TabIndex        =   107
                  Top             =   240
                  Width           =   8295
                  _ExtentX        =   14631
                  _ExtentY        =   4260
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
                  Appearance      =   0
               End
            End
            Begin VB.Frame Frame11 
               Height          =   2775
               Left            =   120
               TabIndex        =   104
               Top             =   1080
               Width           =   8535
               Begin MSFlexGridLib.MSFlexGrid flxREFBANCARIA 
                  Height          =   2415
                  Left            =   120
                  TabIndex        =   105
                  Top             =   240
                  Width           =   8415
                  _ExtentX        =   14843
                  _ExtentY        =   4260
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
                  Appearance      =   0
               End
            End
            Begin VB.Frame Frame10 
               Height          =   735
               Left            =   120
               TabIndex        =   100
               Top             =   360
               Width           =   8535
               Begin VB.CommandButton Command3 
                  Height          =   315
                  Left            =   7320
                  Picture         =   "frmCADCLIENTEP.frx":37CE
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  Top             =   360
                  Width           =   375
               End
               Begin VB.TextBox txtTelREF 
                  Height          =   285
                  Left            =   5760
                  MaxLength       =   15
                  TabIndex        =   31
                  Text            =   "txtTelREF"
                  Top             =   360
                  Width           =   1575
               End
               Begin VB.TextBox txtRefNome 
                  Height          =   285
                  Left            =   4200
                  MaxLength       =   20
                  TabIndex        =   30
                  Text            =   "txtRefNome"
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txtRefBanco 
                  Height          =   285
                  Left            =   120
                  MaxLength       =   30
                  TabIndex        =   29
                  Text            =   "txtRefBanco"
                  Top             =   360
                  Width           =   3975
               End
               Begin VB.Label Label32 
                  AutoSize        =   -1  'True
                  Caption         =   "Telefone:"
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
                  Left            =   5760
                  TabIndex        =   103
                  Top             =   120
                  Width           =   825
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  Caption         =   "Nome:"
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
                  Left            =   4200
                  TabIndex        =   102
                  Top             =   120
                  Width           =   555
               End
               Begin VB.Label Label30 
                  AutoSize        =   -1  'True
                  Caption         =   "Banco:"
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
                  Left            =   120
                  TabIndex        =   101
                  Top             =   120
                  Width           =   615
               End
            End
         End
         Begin VB.Frame Frame6 
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
            TabIndex        =   97
            Top             =   360
            Width           =   8775
            Begin VB.CommandButton cmdPesq 
               Height          =   315
               Left            =   1800
               Picture         =   "frmCADCLIENTEP.frx":3D00
               Style           =   1  'Graphical
               TabIndex        =   131
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton cmbGravPagto 
               Height          =   315
               Left            =   8280
               Picture         =   "frmCADCLIENTEP.frx":3E02
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtCODBANCO 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   840
               MaxLength       =   10
               TabIndex        =   26
               Text            =   "txtCODBANCO"
               Top             =   240
               Width           =   975
            End
            Begin VB.ComboBox cboBANCOS 
               Height          =   315
               Left            =   2160
               TabIndex        =   27
               Text            =   "cboBANCOS"
               Top             =   240
               Width           =   6135
            End
            Begin VB.Label Label29 
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
               Left            =   120
               TabIndex        =   98
               Top             =   240
               Width           =   660
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Limite de Crédito"
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
         Left            =   -74880
         TabIndex        =   90
         Top             =   720
         Width           =   9015
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   7800
            Picture         =   "frmCADCLIENTEP.frx":4334
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtVLLIMCRED 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   48
            Text            =   "txtVLLIMCRED"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtAPROVADO 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6240
            TabIndex        =   91
            Text            =   "txtAPROVADO"
            Top             =   240
            Width           =   1575
         End
         Begin MSMask.MaskEdBox mskDTAPROV 
            Height          =   285
            Left            =   3960
            TabIndex        =   92
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Limite de compras:"
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
            Left            =   120
            TabIndex        =   95
            Top             =   255
            Width           =   1605
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
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
            Left            =   3360
            TabIndex        =   94
            Top             =   270
            Width           =   480
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Aprovado:"
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
            Left            =   5280
            TabIndex        =   93
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Endereço para Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   66
         Top             =   2760
         Width           =   9015
         Begin VB.TextBox txtCEPENTR 
            Height          =   285
            Left            =   3840
            MaxLength       =   9
            TabIndex        =   25
            Text            =   "txtCEPENTR"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox cboESTENTR 
            Height          =   315
            Left            =   1200
            TabIndex        =   24
            Text            =   "cboESTENTR"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtCIDENTR 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   23
            Text            =   "txtCIDENTR"
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtBAIENTR 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   22
            Text            =   "txtBAIENTR"
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtENDENTR 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   21
            Text            =   "txtENDENTR"
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Left            =   3240
            TabIndex        =   87
            Top             =   1500
            Width           =   435
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Left            =   300
            TabIndex        =   86
            Top             =   1500
            Width           =   660
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   300
            TabIndex        =   85
            Top             =   1080
            Width           =   660
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            Left            =   405
            TabIndex        =   84
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Endereço para Cobrança"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         TabIndex        =   65
         Top             =   720
         Width           =   9015
         Begin VB.TextBox txtCEPCOBR 
            Height          =   285
            Left            =   3840
            MaxLength       =   9
            TabIndex        =   20
            Text            =   "txtCEPCOBR"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox cboESTCOBR 
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Text            =   "cboESTCOBR"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtCIDCOBR 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   18
            Text            =   "txtCIDCOBR"
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtBAICOBR 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   17
            Text            =   "txtBAICOBR"
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtENDCOBR 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   16
            Text            =   "txtENDCOBR"
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Index           =   0
            Left            =   3240
            TabIndex        =   82
            Top             =   1500
            Width           =   435
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Index           =   330
            Left            =   300
            TabIndex        =   81
            Top             =   1500
            Width           =   660
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   300
            TabIndex        =   80
            Top             =   1080
            Width           =   660
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            Left            =   400
            TabIndex        =   79
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            Left            =   120
            TabIndex        =   78
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   120
         TabIndex        =   64
         Top             =   720
         Width           =   13095
         Begin VB.Frame Frame56 
            Caption         =   "[ Permite Fechamento de OP com 10% ]"
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
            Left            =   8760
            TabIndex        =   295
            Top             =   4920
            Width           =   4215
            Begin VB.OptionButton optPermFecOP 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   297
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optPermFecOP 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   296
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame59 
            Caption         =   "[ Pemite Faturar Rotulos Separados ]"
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
            Index           =   4
            Left            =   120
            TabIndex        =   290
            Top             =   4920
            Width           =   3495
            Begin VB.OptionButton optPermFatSepSN 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   292
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton optPermFatSepSN 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   291
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame59 
            Caption         =   "[ Ultimo Faturamento ]"
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
            Index           =   3
            Left            =   10320
            TabIndex        =   285
            Top             =   3720
            Width           =   2655
            Begin MSMask.MaskEdBox mskDTULTFATNOVA 
               Height          =   285
               Left            =   1320
               TabIndex        =   286
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDTULTFATSTEEL 
               Height          =   285
               Left            =   1320
               TabIndex        =   287
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Novalata"
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
               TabIndex        =   289
               Top             =   360
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Steel"
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
               TabIndex        =   288
               Top             =   720
               Width           =   450
            End
         End
         Begin VB.Frame Frame59 
            Caption         =   "[ Necessita Confirmação de Estoque ]"
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
            Index           =   2
            Left            =   120
            TabIndex        =   282
            Top             =   4320
            Width           =   3495
            Begin VB.OptionButton optNECCONFEST 
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
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   1080
               TabIndex        =   284
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optNECCONFEST 
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
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   283
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame59 
            Caption         =   "Cliente Habilitado ]"
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
            Index           =   1
            Left            =   3720
            TabIndex        =   281
            Top             =   4920
            Width           =   4935
            Begin VB.OptionButton optDesbClie 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   294
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optDesbClie 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   2760
               TabIndex        =   293
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame59 
            Caption         =   "[ Será visualizada na Tela de Entradada de Estoque ]"
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
            Index           =   0
            Left            =   3720
            TabIndex        =   276
            Top             =   4320
            Width           =   4935
            Begin VB.OptionButton optVISTELENT 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   280
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optVISTELENT 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   2760
               TabIndex        =   279
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox txtCodRef 
            Height          =   285
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   15
            Text            =   "txtCodRef"
            Top             =   3720
            Width           =   1695
         End
         Begin VB.TextBox txtZonaGeo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6480
            TabIndex        =   230
            Text            =   "txtZonaGeo"
            Top             =   2040
            Width           =   3135
         End
         Begin VB.Frame Frame33 
            Height          =   495
            Left            =   8640
            TabIndex        =   174
            Top             =   1080
            Width           =   3015
            Begin VB.OptionButton optECLIENAO 
               Caption         =   "Não"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   1920
               TabIndex        =   177
               Top             =   220
               Width           =   735
            End
            Begin VB.OptionButton optECLIESIM 
               Caption         =   "Sim"
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   1035
               TabIndex        =   176
               Top             =   200
               Width           =   615
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               Caption         =   "É Cliente:"
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
               Left            =   120
               TabIndex        =   175
               Top             =   200
               Width           =   840
            End
         End
         Begin VB.Frame Frame32 
            Height          =   495
            Left            =   8640
            TabIndex        =   170
            Top             =   520
            Width           =   3015
            Begin VB.OptionButton optFisica 
               Caption         =   "Fisica"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   1035
               TabIndex        =   172
               Top             =   165
               Width           =   735
            End
            Begin VB.OptionButton opfJuridica 
               Caption         =   "Juridica"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   1920
               TabIndex        =   171
               Top             =   190
               Width           =   855
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               Caption         =   "Pessoa:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   173
               Top             =   165
               Width           =   690
            End
         End
         Begin VB.TextBox txtSTATUS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   89
            Text            =   "txtCodigo"
            Top             =   240
            Width           =   1335
         End
         Begin MSMask.MaskEdBox mskDTCADASTRO 
            Height          =   285
            Left            =   10440
            TabIndex        =   14
            Top             =   2820
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboSITENORM 
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Text            =   "cboSITENORM"
            Top             =   3300
            Width           =   4575
         End
         Begin VB.ComboBox cboEMAILNORM 
            Height          =   315
            Left            =   1440
            TabIndex        =   12
            Text            =   "cboEMAILNORM"
            Top             =   2880
            Width           =   4575
         End
         Begin VB.ComboBox cboCONTNORM 
            Height          =   315
            Left            =   6120
            TabIndex        =   11
            Text            =   "cboCONTNORM"
            Top             =   2450
            Width           =   5535
         End
         Begin VB.ComboBox cboTELNORM 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Text            =   "cboTELNORM"
            Top             =   2450
            Width           =   3135
         End
         Begin VB.TextBox txtCEPNORM 
            Height          =   285
            Left            =   10320
            MaxLength       =   9
            TabIndex        =   9
            Text            =   "txtCEPNORM"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.ComboBox cboESTNORM 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            Text            =   "cboESTNORM"
            Top             =   2040
            Width           =   750
         End
         Begin VB.TextBox txtCIDNORM 
            Height          =   285
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "txtCIDNORM"
            Top             =   1680
            Width           =   5175
         End
         Begin VB.TextBox txtBAINOM 
            Height          =   285
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   6
            Text            =   "txtBAINOM"
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtENDNORM 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "txtENDNORM"
            Top             =   1320
            Width           =   7095
         End
         Begin VB.TextBox txtNOMFANTA 
            Height          =   285
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   4
            Text            =   "txtNOMFANTA"
            Top             =   960
            Width           =   6975
         End
         Begin VB.TextBox txRGCGC 
            Height          =   285
            Left            =   9600
            MaxLength       =   20
            TabIndex        =   2
            Text            =   "txRGCGC"
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtCPFCNPJ 
            Height          =   285
            Left            =   6120
            MaxLength       =   18
            TabIndex        =   1
            Text            =   "txtCPFCNPJ"
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtRAZAOSOC 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "txtRAZAOSOC"
            Top             =   600
            Width           =   6975
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   0
            Text            =   "txtCodigo"
            Top             =   240
            Width           =   855
         End
         Begin MSMask.MaskEdBox dtNasc 
            Height          =   285
            Left            =   7200
            TabIndex        =   242
            Top             =   2880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Telefone/Fax:"
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
            Index           =   4
            Left            =   120
            TabIndex        =   278
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Left            =   720
            TabIndex        =   277
            Top             =   2040
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            Left            =   720
            TabIndex        =   275
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            Left            =   480
            TabIndex        =   274
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Ref:"
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
            Left            =   525
            TabIndex        =   243
            Top             =   3720
            Width           =   825
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Data Nasc."
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
            Left            =   6120
            TabIndex        =   241
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Zona Geografica:"
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
            Left            =   4800
            TabIndex        =   229
            Top             =   2100
            Width           =   1500
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Left            =   2520
            TabIndex        =   88
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Data Cadastro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   9000
            TabIndex        =   77
            Top             =   2880
            Width           =   1290
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Site:"
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
            Left            =   900
            TabIndex        =   76
            Top             =   3330
            Width           =   405
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "E-MAIL:"
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
            Left            =   650
            TabIndex        =   75
            Top             =   2880
            Width           =   690
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
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
            Left            =   4800
            TabIndex        =   74
            Top             =   2475
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
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
            Left            =   9840
            TabIndex        =   73
            Top             =   2100
            Width           =   405
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   4800
            TabIndex        =   72
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nome fantasia:"
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
            Index           =   0
            Left            =   45
            TabIndex        =   71
            Top             =   960
            Width           =   1290
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "RG/I.EST:"
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
            Left            =   8640
            TabIndex        =   70
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   4980
            TabIndex        =   69
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   1200
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
            Index           =   0
            Left            =   655
            TabIndex        =   67
            Top             =   270
            Width           =   660
         End
      End
   End
End
Attribute VB_Name = "frmCADCLIENTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public cTipOper     As String
Public iCodigo      As Long
Public FILIAL       As Integer
Public strAcesso    As String
Public strMODPAI    As String
Public strUSUARIO   As String
Public lngIDUsuario As Long

Dim objBLBFunc     As New clsFuncoes
Dim objCADCLIENTE  As New clsCADCLIENTE
Dim objPESQPADRAO  As New clsPESQPADRAO
Dim objCADCONFAT   As Object
Dim objCADPEDIDO   As Object


Dim arrCAMPOS           As Variant
Dim arrTABELA           As Variant
Dim arrTELEFONE         As Variant
Dim arrCONTATO          As Variant
Dim arrEMAIL            As Variant
Dim arrSITE             As Variant
Dim arrBANCOS           As Variant
Dim arrREFBANCARIA      As Variant
Dim arrREFCOMERC        As Variant
Dim arrREFPESSOAL       As Variant
Dim arrRESTRICOES       As Variant
Dim arrFORNECEDOR       As Variant
Dim arrCLIENTES         As Variant
Dim arrSISTCERTI        As Variant
Dim arrEMPRESATEND      As Variant
Dim arrTRANSP           As Variant
Dim arrTECNICOS         As Variant
Dim arrVENDEDORES       As Variant
Dim arrPRODUTO          As Variant
Dim arrPRODUTOSCLIE     As Variant
Dim arrCONDPGTOCLIE     As Variant

Const conCOL_Produto_IdProduto                  As Integer = 0
Const conCOL_Produto_DataEmissao                As Integer = 1
Const conCOL_Produto_NumeroFatura               As Integer = 2
Const conCOL_Produto_Vendmento                  As Integer = 3
Const conCOL_Produto_ValorUnit                  As Integer = 4
Const conCOL_Produto_Quantidade                 As Integer = 5
Const conCOL_Produto_CodProd                    As Integer = 6
Const conCOL_Produto_Descricao                  As Integer = 7
Const conCOL_Produto_FormatString               As String = "=Código|Data Emissão|Numero Fatura|Vencimento|Valor Unitário|Quantidade|Cod. Produto|Descrição Produto"
Const conColumnsIn_Produto                      As Integer = 8

Const conCOL_ProdutoClie_IdProduto              As Integer = 0
Const conCOL_ProdutoClie_Rotulo                 As Integer = 1
Const conCOL_ProdutoClie_PesqRot                As Integer = 2
Const conCOL_ProdutoClie_Descricao              As Integer = 3
Const conCOL_ProdutoClie_FormatString           As String = "=IdProduto|Rótulo|...|Descrição"
Const conColumnsIn_ProdutoClie                  As Integer = 4

Const conCOL_UltFat_IdProduto                       As Integer = 0
Const conCOL_UltFat_DtFat                           As Integer = 1
Const conCOL_UltFat_DtEntrega                       As Integer = 2
Const conCOL_UltFat_QtdFat                          As Integer = 3
Const conCOL_UltFat_QtdPed                          As Integer = 4
Const conCOL_UltFat_Saldo                           As Integer = 5
Const conCOL_UltFat_Unit                            As Integer = 6
Const conCOL_UltFat_Valor                           As Integer = 7
Const conCOL_UltFat_CodConf                         As Integer = 8
Const conCOL_UltFat_CodPed                          As Integer = 9
Const conCOL_UltFat_CodOP                           As Integer = 10
Const conCOL_UltFat_CodNF                           As Integer = 11
Const conCOL_UltFat_Status                          As Integer = 12
Const conCOL_UltFat_FormatString                    As String = "=IdProduto|Data.Fat|Dt.Entrega|Qtd.Fat|Qtd.Ped|Saldo|Vl.Unit|Valor.Tot|Cod.Conf.Fat|Cod.Pedido|Cod.OP|Cod.NF|Status"
Const conColumnsIn_UltFat                           As Integer = 13

Const conCOL_ProdutoClieABC_IdProduto               As Integer = 0
Const conCOL_ProdutoClieABC_Rotulo                  As Integer = 1
Const conCOL_ProdutoClieABC_PesqRot                 As Integer = 2
Const conCOL_ProdutoClieABC_Descricao               As Integer = 3
Const conCOL_ProdutoClieABC_FormatString            As String = "=IdProduto|Rótulo|...|Descrição"
Const conColumnsIn_ProdutoClieABC                   As Integer = 4

Const conCOL_ProdutoPedidos_MesAno                  As Integer = 0
Const conCOL_ProdutoPedidos_Codigo                  As Integer = 1
Const conCOL_ProdutoPedidos_Qtde                    As Integer = 2
Const conCOL_ProdutoPedidos_FormatString            As String = "=MesAno|Pedido|Qtde"
Const conColumnsIn_ProdutoPedidos                   As Integer = 3

Const conCOL_ProdutoFaturado_MesAno                 As Integer = 0
Const conCOL_ProdutoFaturado_Codigo                 As Integer = 1
Const conCOL_ProdutoFaturado_Qtde                   As Integer = 2
Const conCOL_ProdutoFaturado_FormatString           As String = "=MesAno|Pedido|Qtde"
Const conColumnsIn_ProdutoFaturado                  As Integer = 3

Dim ProdutoCurvABC_FormatString                     As String
Dim ColumnsIn_ProdutoCurvABC                        As Long

Const conCOL_Cliente_CodCOndPgto                    As Integer = 0
Const conCOL_Cliente_Pesq                           As Integer = 1
Const conCOL_Cliente_DescCondPgto                   As Integer = 2
Const conCOL_Cliente_FormatString                   As String = "=Código|...|Descrição da Condição de Pagamento"
Const conColumnsIn_Cliente                          As Integer = 3



Private Sub cboBANCOS_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboBANCOS, KeyAscii
End Sub

Private Sub cboBANCOS_Validate(Cancel As Boolean)
    If cboBANCOS.ListIndex > -1 Then txtCODBANCO.Text = cboBANCOS.ItemData(cboBANCOS.ListIndex)
End Sub

Private Sub cboCONTNORM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       
       If cboCONTNORM.ListIndex = -1 Then Exit Sub
       
       cboCONTNORM.RemoveItem cboCONTNORM.ListIndex
       cboCONTNORM.Text = ""
       
    End If
End Sub

Private Sub cboCONTNORM_Validate(Cancel As Boolean)

   If Len(Trim(cboCONTNORM.Text)) > 40 Then
      MsgBox "Somente é permitido 40 Digitos !!!", vbOKOnly + vbCritical, "aviso"
      cboCONTNORM.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If Len(Trim(cboCONTNORM.Text)) = 0 Then Exit Sub
   
   cboCONTNORM.AddItem cboCONTNORM.Text
   cboCONTNORM.Text = ""

End Sub

Private Sub cboEMAILNORM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       
       If cboEMAILNORM.ListIndex = -1 Then Exit Sub
       
       cboEMAILNORM.RemoveItem cboEMAILNORM.ListIndex
       cboEMAILNORM.Text = ""
       
    End If
End Sub


Private Sub cboEMAILNORM_Validate(Cancel As Boolean)

   If Len(Trim(cboEMAILNORM.Text)) > 50 Then
      MsgBox "Somente é permitido 50 Digitos !!!", vbOKOnly + vbCritical, "aviso"
      cboEMAILNORM.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If Len(Trim(cboEMAILNORM.Text)) = 0 Then Exit Sub
   
   cboEMAILNORM.AddItem cboEMAILNORM.Text
   cboEMAILNORM.Text = ""

End Sub

Private Sub cboESTCOBR_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboESTCOBR, KeyAscii
End Sub
Private Sub cboESTENTR_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboESTCOBR, KeyAscii
End Sub

Private Sub cboESTNORM_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboESTNORM, KeyAscii
End Sub

Private Sub cboESTNORM_Validate(Cancel As Boolean)
    If cboESTNORM.ListIndex > -1 Then txtZonaGeo.Text = BuscaAreaGeo(CLng(cboESTNORM.ItemData(cboESTNORM.ListIndex)))
End Sub


Private Sub cboSEGMENTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSEGMENTO, KeyAscii
End Sub

Private Sub cboSEGMENTO_Validate(Cancel As Boolean)
        If cboSEGMENTO.ListIndex > -1 Then txtCODSEQ.Text = cboSEGMENTO.ItemData(cboSEGMENTO.ListIndex)
End Sub

Private Sub cboSEGMENTO2_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSEGMENTO2, KeyAscii
End Sub

Private Sub cboSITENORM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cboSITENORM.ListIndex = -1 Then Exit Sub
       cboSITENORM.RemoveItem cboSITENORM.ListIndex
       cboSITENORM.Text = ""
    End If
End Sub

Private Sub cboSITENORM_Validate(Cancel As Boolean)

   If Len(Trim(cboSITENORM.Text)) > 50 Then
      MsgBox "Somente é permitido 50 Digitos !!!", vbOKOnly + vbCritical, "aviso"
      cboSITENORM.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If Len(Trim(cboSITENORM.Text)) = 0 Then Exit Sub
   
   cboSITENORM.AddItem cboSITENORM.Text
   cboSITENORM.Text = ""

End Sub

Private Sub cboTELNORM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       
       If cboTELNORM.ListIndex = -1 Then Exit Sub
       
       cboTELNORM.RemoveItem cboTELNORM.ListIndex
       cboTELNORM.Text = ""
       
    End If
End Sub

Private Sub cboTELNORM_Validate(Cancel As Boolean)

   If Len(Trim(cboTELNORM.Text)) > 13 Then
      MsgBox "Somente é permitido 13 Digitos !!!", vbOKOnly + vbCritical, "aviso"
      cboTELNORM.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If Len(Trim(cboTELNORM.Text)) = 0 Then Exit Sub
   
   cboTELNORM.AddItem cboTELNORM.Text
   cboTELNORM.Text = ""

End Sub

Private Sub cboTRANSP_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTRANSP, KeyAscii
End Sub

Private Sub cboTRANSP_Validate(Cancel As Boolean)
    If cboTRANSP.ListIndex > -1 Then txtCODTRANSP.Text = cboTRANSP.ItemData(cboTRANSP.ListIndex)
End Sub

Private Sub cmbGravPagto_Click()
    If cTipOper = "I" Then InclBancos
    If cTipOper = "A" Then InclBancos
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    '' Dados Cadastrais **
    Frame2.Enabled = True
    txtCPFCNPJ.Enabled = True
    txtRAZAOSOC.Enabled = True
    txRGCGC.Enabled = True
    txtNOMFANTA.Enabled = True
    optFisica.Enabled = True
    opfJuridica.Enabled = True
    txtENDNORM.Enabled = True
    txtBAINOM.Enabled = True
    txtCIDNORM.Enabled = True
    cboESTNORM.Enabled = True
    txtCEPNORM.Enabled = True
    
    cboTELNORM.Locked = False
    cboCONTNORM.Locked = False
    cboEMAILNORM.Locked = False
    cboEMAILNORM.Locked = False
    
    mskDTCADASTRO.Enabled = False
    
    txtOBS.Locked = False
    txtObsCom.Locked = False
    
    '' -----------------------
    
    '' Cobrança /  Entrega
    Frame3.Enabled = True
    Frame4.Enabled = True
    '' -----------------------
    
    '' Financeiro
    Frame6.Enabled = True
    Frame10.Enabled = True
    Frame13.Enabled = True
    Frame15.Enabled = True
    Frame17.Enabled = True
    Frame18.Enabled = True
    Frame19.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True
    '' -----------------------
    
    '' Crédito **
    Frame5.Enabled = True
    Frame24.Enabled = True
    Frame7.Enabled = False
    '' -----------------------
    
    '' Folhow UP
    Frame26.Enabled = True
    Frame27.Enabled = True
    Frame28.Enabled = True
    Frame29.Enabled = True
    Frame30.Enabled = True
    '' -----------------------
    
    '' Transportadora
    Frame40.Enabled = True
    Frame41.Enabled = True
    
    
    Me.Caption = "Cadastro de clientes - [ ALTERAÇÃO ]"

    cTipOper = "A"

End Sub


Private Sub cmdIncTransp_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridTransp
End Sub

Private Sub cmdPesq_Click()
    
    ReDim arrCAMPOS(1 To 4, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADBANCOS"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_AGENCIA"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Agência"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_AGENCIA"
    
    arrCAMPOS(3, 1) = "SGI_CC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "C/C"
    arrCAMPOS(3, 4) = "1500"
    arrCAMPOS(3, 5) = "SGI_CC"
    
    arrCAMPOS(4, 1) = "SGI_DESCRICAO"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Banco"
    arrCAMPOS(4, 4) = "3000"
    arrCAMPOS(4, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Bancos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODBANCO.Text = varRETORNO
    
    cboBANCOS.ListIndex = -1
    txtCODBANCO.SetFocus
    
End Sub

Private Sub CmdSalva_Click()
    
    Dim I As Integer
    
    If Verifica_Campos = False Then Exit Sub
    
    If optECLIESIM.Value = True Then
       If cTipOper = "I" Then objCADCLIENTE.CLIECODIGO = objCADCLIENTE.Gera_Codigo(Me.Name, True)
    End If
    If optECLIENAO.Value = True Then
       If cTipOper = "I" Then objCADCLIENTE.CLIECODIGO = objCADCLIENTE.Gera_Codigo(Me.Name, False)
    End If
    
    objCADCLIENTE.Modulo = Me.Name
    
    objCADCLIENTE.CLISTATUS = "EM APROVAÇÂO"
    objCADCLIENTE.CLIERZAOSO = txtRAZAOSOC.Text
    objCADCLIENTE.CODREF = txtCodRef.Text
    objCADCLIENTE.CLICPFCNPJ = Trim(Replace(Replace(Replace(txtCPFCNPJ.Text, ".", ""), "/", ""), "-", ""))
    objCADCLIENTE.CLIRGCGC = txRGCGC.Text
    objCADCLIENTE.CLINOMFANT = txtNOMFANTA.Text
    
    If optFisica.Value = True Then objCADCLIENTE.CLIPESSOA = "F"
    If opfJuridica.Value = True Then objCADCLIENTE.CLIPESSOA = "J"
    
    If optNECCONFEST(0).Value = True Then objCADCLIENTE.NECCONFEST = 0
    If optNECCONFEST(1).Value = True Then objCADCLIENTE.NECCONFEST = 1
    
    
    If optVISTELENT(0).Value = True Then objCADCLIENTE.VISTELEST = 0
    If optVISTELENT(1).Value = True Then objCADCLIENTE.VISTELEST = 1
    
    objCADCLIENTE.CLIENDEREC = txtENDNORM.Text
    objCADCLIENTE.CLIBAIRRO = txtBAINOM.Text
    objCADCLIENTE.CLICIDADE = txtCIDNORM.Text
    If cboESTNORM.ListIndex > -1 Then objCADCLIENTE.CLIESTADO = cboESTNORM.ItemData(cboESTNORM.ListIndex)
    objCADCLIENTE.CLICEP = txtCEPNORM.Text
    
    If optPermFatSepSN(0).Value = True Then objCADCLIENTE.PERMFATSSEPSN = 0
    If optPermFatSepSN(1).Value = True Then objCADCLIENTE.PERMFATSSEPSN = 1
    
    If optDesbClie(0).Value = True Then objCADCLIENTE.DESBCLIE = 0
    If optDesbClie(1).Value = True Then objCADCLIENTE.DESBCLIE = 1
    
    If (cboTELNORM.ListCount) > 0 Then '' Telefone
       ReDim arrTELEFONE(0 To (cboTELNORM.ListCount - 1), 1 To 2) As String
       For I = 0 To (cboTELNORM.ListCount - 1)
           arrTELEFONE(I, 1) = cboTELNORM.List(I)
           arrTELEFONE(I, 2) = cboTELNORM.Name
       Next I
       objCADCLIENTE.CLITELNORM = arrTELEFONE
    End If
    If (cboCONTNORM.ListCount) > 0 Then '' Contato
       ReDim arrCONTATO(0 To (cboCONTNORM.ListCount - 1), 1 To 2) As String
       For I = 0 To (cboCONTNORM.ListCount - 1)
           arrCONTATO(I, 1) = cboCONTNORM.List(I)
           arrCONTATO(I, 2) = cboCONTNORM.Name
       Next I
       objCADCLIENTE.CLICONTATO = arrCONTATO
    End If
    If (cboEMAILNORM.ListCount) > 0 Then '' E-Mail
       ReDim arrEMAIL(0 To (cboEMAILNORM.ListCount - 1), 1 To 2) As String
       For I = 0 To (cboEMAILNORM.ListCount - 1)
           arrEMAIL(I, 1) = cboEMAILNORM.List(I)
           arrEMAIL(I, 2) = cboEMAILNORM.Name
       Next I
       objCADCLIENTE.CLIEMAIL = arrEMAIL
    End If
    If (cboSITENORM.ListCount) > 0 Then '' Site
       ReDim arrSITE(0 To (cboSITENORM.ListCount - 1), 1 To 2) As String
       For I = 0 To (cboSITENORM.ListCount - 1)
           arrSITE(I, 1) = cboSITENORM.List(I)
           arrSITE(I, 2) = cboSITENORM.Name
       Next I
       objCADCLIENTE.CLISITE = arrSITE
    End If
    '' Transportadoras
    If flxTRANSP.Rows > 1 Then
       ReDim arrTRANSP(1 To (flxTRANSP.Rows - 1)) As Long
       For I = 1 To (flxTRANSP.Rows - 1)
           arrTRANSP(I) = flxTRANSP.TextMatrix(I, 1)
       Next I
       objCADCLIENTE.TRANSP = arrTRANSP
    Else
       ReDim arrTRANSP(0) As String
       objCADCLIENTE.TRANSP = arrTRANSP
    End If

    arrPRODUTO = Empty
    objCADCLIENTE.PRODUTO = arrPRODUTO
    
    objCADCLIENTE.CADASTRO = CDate(mskDTCADASTRO.Text)
    
    objCADCLIENTE.ENDCOBRA = txtENDCOBR.Text
    objCADCLIENTE.BAICOBRA = txtBAICOBR.Text
    objCADCLIENTE.CIDCOBRA = txtCIDCOBR.Text
    If cboESTCOBR.ListIndex > -1 Then objCADCLIENTE.ESTCOBRA = cboESTCOBR.ItemData(cboESTCOBR.ListIndex)
    objCADCLIENTE.CEPCOBRA = txtCEPCOBR.Text
    
    objCADCLIENTE.ENDENTREGA = txtENDENTR.Text
    objCADCLIENTE.BAIENTREGA = txtBAIENTR.Text
    objCADCLIENTE.CIDENTREGA = txtCIDENTR.Text
    If cboESTENTR.ListIndex > -1 Then objCADCLIENTE.ESTENTREGA = cboESTENTR.ItemData(cboESTENTR.ListIndex)
    objCADCLIENTE.CEPENTREGA = txtCEPENTR.Text
    
    If optBLPEDSIM.Value = True Then objCADCLIENTE.BLOQPEDREST = "S"
    If optBLPEDNAO.Value = True Then objCADCLIENTE.BLOQPEDREST = "N"
    
    If optAVISASIM.Value = True Then objCADCLIENTE.AVISARRESTR = "S"
    If optAVISANAO.Value = True Then objCADCLIENTE.AVISARRESTR = "N"
    
    If optSempSIM.Value = True Then objCADCLIENTE.SEMPRBLOQPE = "S"
    If optSempNAO.Value = True Then objCADCLIENTE.SEMPRBLOQPE = "N"
    
    If optBloqSim.Value = True Then objCADCLIENTE.BLOQAVISALD = "S"
    If optBloqNao.Value = True Then objCADCLIENTE.BLOQAVISALD = "N"
    
    '' pERMITE fECHAR A op COM 10%
    If optPermFecOP(0).Value = True Then objCADCLIENTE.PERMFECHOP = 0
    If optPermFecOP(1).Value = True Then objCADCLIENTE.PERMFECHOP = 1
    
    If (flxBANCOS.Rows - 1) > 0 Then '' Bancos
       ReDim arrBANCOS(1 To (flxBANCOS.Rows - 1)) As String
       For I = 1 To (flxBANCOS.Rows - 1)
           arrBANCOS(I) = flxBANCOS.TextMatrix(I, 1)
       Next I
       objCADCLIENTE.BANCOS = arrBANCOS
    End If
    If (flxREFBANCARIA.Rows - 1) > 0 Then '' Referencia bancaria
       ReDim arrREFBANCARIA(1 To (flxREFBANCARIA.Rows - 1), 1 To 3) As String
       For I = 1 To (flxREFBANCARIA.Rows - 1)
           arrREFBANCARIA(I, 1) = flxREFBANCARIA.TextMatrix(I, 1)
           arrREFBANCARIA(I, 2) = flxREFBANCARIA.TextMatrix(I, 2)
           arrREFBANCARIA(I, 3) = flxREFBANCARIA.TextMatrix(I, 3)
       Next I
       objCADCLIENTE.REFBANCARIA = arrREFBANCARIA
    End If
    If (flxREFEMPRESA.Rows - 1) > 0 Then '' Referencia Comercial
       ReDim arrREFCOMERC(1 To (flxREFEMPRESA.Rows - 1), 1 To 3) As String
       For I = 1 To (flxREFEMPRESA.Rows - 1)
           arrREFCOMERC(I, 1) = flxREFEMPRESA.TextMatrix(I, 1)
           arrREFCOMERC(I, 2) = flxREFEMPRESA.TextMatrix(I, 2)
           arrREFCOMERC(I, 3) = flxREFEMPRESA.TextMatrix(I, 3)
       Next I
       objCADCLIENTE.REFCOMERC = arrREFCOMERC
    End If
    If (flxREFPESSOAL.Rows - 1) > 0 Then '' Referencia Pessoal
       ReDim arrREFPESSOAL(1 To (flxREFPESSOAL.Rows - 1), 1 To 2) As String
       For I = 1 To (flxREFPESSOAL.Rows - 1)
           arrREFPESSOAL(I, 1) = flxREFPESSOAL.TextMatrix(I, 1)
           arrREFPESSOAL(I, 2) = flxREFPESSOAL.TextMatrix(I, 2)
       Next I
       objCADCLIENTE.REFPESSOAL = arrREFPESSOAL
    End If
    If (flxRESTRICOES.Rows - 1) > 0 Then '' Restrições
       ReDim arrRESTRICOES(1 To (flxRESTRICOES.Rows - 1)) As String
       For I = 1 To (flxRESTRICOES.Rows - 1)
           arrRESTRICOES(I) = flxRESTRICOES.TextMatrix(I, 1)
       Next I
       objCADCLIENTE.RESTRICOES = arrRESTRICOES
    End If
    If (flxFORNEC.Rows - 1) > 0 Then '' Fornecedores
       ReDim arrFORNECEDOR(1 To (flxFORNEC.Rows - 1), 1 To 2) As String
       For I = 1 To (flxFORNEC.Rows - 1)
           arrFORNECEDOR(I, 1) = flxFORNEC.TextMatrix(I, 1)
           arrFORNECEDOR(I, 2) = flxFORNEC.TextMatrix(I, 2)
       Next I
       objCADCLIENTE.FORNECEDORE = arrFORNECEDOR
    End If
    If (flxEMPRESA.Rows - 1) > 0 Then '' Clientes
       ReDim arrCLIENTES(1 To (flxEMPRESA.Rows - 1), 1 To 2) As String
       For I = 1 To (flxEMPRESA.Rows - 1)
           arrCLIENTES(I, 1) = flxEMPRESA.TextMatrix(I, 1)
           arrCLIENTES(I, 2) = flxEMPRESA.TextMatrix(I, 2)
       Next I
       objCADCLIENTE.CLIENTES = arrCLIENTES
    End If
    
    If Len(Trim(txtVLLIMCRED.Text)) > 0 Then objCADCLIENTE.VALLIMCRED = CDbl(txtVLLIMCRED.Text)
    If Len(Trim(txtSALDOACIMA.Text)) > 0 Then objCADCLIENTE.AVISASALAC = CDbl(txtSALDOACIMA.Text)
    If Len(Trim(txtBLOQPEDSALDO.Text)) > 0 Then objCADCLIENTE.BLQSALACIM = CDbl(txtBLOQPEDSALDO.Text)
    If Len(Trim(txtMeseReavali.Text)) > 0 Then objCADCLIENTE.MESEREAVAL = CInt(txtMeseReavali.Text)
    
    If Len(Trim(txtCODSEQ.Text)) > 0 Then objCADCLIENTE.SEGMENTO = CInt(txtCODSEQ.Text)
    If optTEMCERTISIM.Value = True Then objCADCLIENTE.SERTIFICADO = "S"
    If optTEMCERTINAO.Value = True Then objCADCLIENTE.SERTIFICADO = "N"
    
    If (flxSISTSERTFIC.Rows - 1) > 0 Then
       ReDim arrSISTCERTI(1 To (flxSISTSERTFIC.Rows - 1)) As String
       For I = 1 To (flxSISTSERTFIC.Rows - 1)
           arrSISTCERTI(I) = flxSISTSERTFIC.TextMatrix(I, 1)
       Next I
       objCADCLIENTE.SISTCERTIF = arrSISTCERTI
    End If
    If (flxATENDIDO.Rows - 1) > 0 Then
       ReDim EMPRESATEND(1 To (flxATENDIDO.Rows - 1), 1 To 2) As String
       For I = 1 To (flxATENDIDO.Rows - 1)
           EMPRESATEND(I, 1) = flxATENDIDO.TextMatrix(I, 1)
           EMPRESATEND(I, 2) = flxATENDIDO.TextMatrix(I, 3)
       Next I
       objCADCLIENTE.EMPREATENDE = EMPRESATEND
    End If
    
    objCADCLIENTE.OBS = txtOBS.Text
    objCADCLIENTE.OBSCOM = txtObsCom.Text
    
    If optECLIESIM.Value = True Then objCADCLIENTE.CLIEJAECLI = "S"
    If optECLIENAO.Value = True Then objCADCLIENTE.CLIEJAECLI = "N"
    
    '' --------------------------------
    '' Produtos Do Cliente
    arrPRODUTOSCLIE = Empty
    With grdPRODUTOS
        If (.Rows - 1) > 0 Then
            ReDim arrPRODUTOSCLIE(1 To (.Rows - 1), 1 To 1) As String
            For I = 1 To (.Rows - 1)
                arrPRODUTOSCLIE(I, 1) = .Cell(flexcpText, I, conCOL_ProdutoClie_IdProduto)
            Next I
        End If
    End With
    objCADCLIENTE.PRODUTOS = arrPRODUTOSCLIE
    '' --------------------------------
    
    arrCONDPGTOCLIE = Empty
    With grdCONDPGTO
        If (.Rows - 1) > 0 Then
            Call objBLBFunc.RemoveLinhaVazia(grdCONDPGTO, conCOL_Cliente_CodCOndPgto)
            ReDim arrCONDPGTOCLIE(1 To (.Rows - 1), 1 To 1) As String
            For I = 1 To (.Rows - 1)
                arrCONDPGTOCLIE(I, 1) = Trim(.Cell(flexcpText, I, conCOL_Cliente_CodCOndPgto))
            Next I
        End If
    End With
    objCADCLIENTE.CONDPGTOCLIE = arrCONDPGTOCLIE
    
    '' Grava as informações
    If objCADCLIENTE.GRAVA(cTipOper) = True Then
          
       MsgBox "O Cliente foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
       If objCADCLIENTE.Atualiza(cTipOper, objCADCLIENTE.CLIECODIGO, FILIAL, Me.Name) = False Then Exit Sub
       
       If cTipOper = "I" Then
          Set objBLBFunc = Nothing
          Set objCADCLIENTE = Nothing
          Unload Me
       End If
          
    End If
   
End Sub

Private Sub cmdTRANSP_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTRANSP" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Transportadoras")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODTRANSP.Text = varRETORNO
    
    cboTRANSP.ListIndex = -1
    txtCODTRANSP.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub


Private Sub Command1_Click()
    AddGridFornec
End Sub

Private Sub Command10_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADSEGCLI"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "4000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Segmento")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODSEQ.Text = varRETORNO
    
    cboSEGMENTO.ListIndex = -1
    txtCODSEQ.SetFocus

End Sub


Private Sub Command11_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdCONDPGTO, grdCONDPGTO.Row)
End Sub

Private Sub Command12_Click()
    If (cTipOper = "I" Or cTipOper = "A") Then Call IncRegGridCondPgto
End Sub

Private Sub Command2_Click()
    AddGridClientes
End Sub

Private Sub Command26_Click()
    If cTipOper = "C" Then Exit Sub
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdPRODUTOS, grdPRODUTOS.Row)
End Sub

Private Sub Command27_Click()
    If (cTipOper = "I" Or cTipOper = "A") Then Call IncRegGridProdtos
End Sub

Private Sub Command3_Click()
    AddGridRefBancaria
End Sub

Private Sub Command4_Click()
    AddGridRefEmpresa
End Sub

Private Sub Command5_Click()
    addGridRefPessoal
End Sub

Private Sub Command7_Click()
    AddGridRestricoes
End Sub

Private Sub Command8_Click()
    If cTipOper = "I" Then InsertSistCertif
    If cTipOper = "A" Then InsertSistCertif
End Sub

Private Sub Command9_Click()
    If cTipOper = "I" Then UnseriEmpreAtende
    If cTipOper = "A" Then UnseriEmpreAtende
End Sub

Private Sub flxATENDIDO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxATENDIDO.Rows = 2 Then flxATENDIDO.Rows = 1
       If flxATENDIDO.Rows > 2 Then flxATENDIDO.RemoveItem flxATENDIDO.RowSel
    End If
End Sub

Private Sub flxBANCOS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxBANCOS.Rows = 2 Then flxBANCOS.Rows = 1
       If flxBANCOS.Rows > 2 Then flxBANCOS.RemoveItem (flxBANCOS.RowSel)
    End If
End Sub


Private Sub flxEMPRESA_KeyDown(KeyCode As Integer, Shift As Integer)

    If cTipOper = "C" Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
       If flxEMPRESA.Rows = 2 Then flxEMPRESA.Rows = 1
       If flxEMPRESA.Rows > 2 Then flxEMPRESA.RemoveItem flxEMPRESA.RowSel
    End If

End Sub

Private Sub flxFORNEC_KeyDown(KeyCode As Integer, Shift As Integer)

    If cTipOper = "C" Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
       If flxFORNEC.Rows = 2 Then flxFORNEC.Rows = 1
       If flxFORNEC.Rows > 2 Then flxFORNEC.RemoveItem flxFORNEC.RowSel
    End If

End Sub


Private Sub flxREFBANCARIA_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If cTipOper = "C" Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
       If flxREFBANCARIA.Rows = 2 Then flxREFBANCARIA.Rows = 1
       If flxREFBANCARIA.Rows > 2 Then flxREFBANCARIA.RemoveItem flxREFBANCARIA.RowSel
    End If
    
End Sub

Private Sub flxREFEMPRESA_KeyDown(KeyCode As Integer, Shift As Integer)

    If cTipOper = "C" Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
       If flxREFEMPRESA.Rows = 2 Then flxREFEMPRESA.Rows = 1
       If flxREFEMPRESA.Rows > 2 Then flxREFEMPRESA.RemoveItem flxREFEMPRESA.RowSel
    End If

End Sub


Private Sub flxREFPESSOAL_KeyDown(KeyCode As Integer, Shift As Integer)

    If cTipOper = "C" Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
       If flxREFPESSOAL.Rows = 2 Then flxREFPESSOAL.Rows = 1
       If flxREFPESSOAL.Rows > 2 Then flxREFPESSOAL.RemoveItem flxREFPESSOAL.RowSel
    End If

End Sub
Private Sub flxRESTRICOES_KeyDown(KeyCode As Integer, Shift As Integer)

    If cTipOper = "C" Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
       If flxRESTRICOES.Rows = 2 Then flxRESTRICOES.Rows = 1
       If flxRESTRICOES.Rows > 2 Then flxRESTRICOES.RemoveItem flxRESTRICOES.RowSel
    End If

End Sub

Private Sub flxSISTSERTFIC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxSISTSERTFIC.Rows = 2 Then flxSISTSERTFIC.Rows = 1
       If flxSISTSERTFIC.Rows > 2 Then flxSISTSERTFIC.RemoveItem flxSISTSERTFIC.RowSel
    End If
End Sub

Private Sub flxTRANSP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxTRANSP.Rows = 2 Then flxTRANSP.Rows = 1
       If flxTRANSP.Rows > 2 Then flxTRANSP.RemoveItem flxTRANSP.RowSel
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   ''Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   ''Set objCADCLIENTE = CreateObject("CADCLIENTE.clsCADCLIENTE")
   ''Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   Set objCADCONFAT = CreateObject("CADCONFFAT.clsCADCONFFAT")
   Set objCADPEDIDO = CreateObject("CADPEDVENDA.clsCADPEDVENDA")
   
   objCADCLIENTE.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   Call DesabilitaCampos
       
   
   ColumnsIn_ProdutoCurvABC = 1
   ProdutoCurvABC_FormatString = "=IdProduto|Descrição"
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    '' Dados Cadastrais **
    optECLIESIM.Enabled = True
    optECLIENAO.Enabled = True
    Frame2.Enabled = True
    '' -----------------------
    
    '' Cobrança /  Entrega
    Frame3.Enabled = True
    Frame4.Enabled = True
    '' -----------------------
    
    '' Financeiro
    Frame6.Enabled = True
    Frame10.Enabled = True
    Frame13.Enabled = True
    Frame15.Enabled = True
    Frame17.Enabled = True
    Frame18.Enabled = True
    Frame19.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True
    '' -----------------------
    
    '' Crédito **
    Frame5.Enabled = True
    Frame24.Enabled = True
    Frame7.Enabled = False
    '' -----------------------
    
    '' Folhow UP
    Frame26.Enabled = True
    Frame27.Enabled = True
    Frame28.Enabled = True
    Frame29.Enabled = True
    Frame30.Enabled = True
    '' -----------------------
    
    '' Transportadora
    Frame40.Enabled = True
    Frame41.Enabled = True
    
    SSTab1.TabVisible(0) = False
    
    Me.Caption = "Cadastro de clientes - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCLIENTE
    
    objBLBFunc.Preenche_Estado cboESTNORM
    objBLBFunc.Preenche_Estado cboESTCOBR
    objBLBFunc.Preenche_Estado cboESTENTR
    
    txtCodigo.Text = ""
    mskDTCADASTRO.Text = Format(Date, "DD/MM/YYYY")
    
    '' --------------------------
    optTipoNOVASTEL(1).Value = True
    optNECCONFEST(0).Value = True
    
    ConfGridBancos
    ConfGridRefBancaria
    CondGridREfCOMERCIAL
    ConfGriRefPessoal
    ConfGridRestricoes
    ConfGridFornec
    ConfGridClientes
    ConfGridHistAvaliacao
    ConfGridCertificado
    ConfGridEmpresaAtende
    ConfGridDuplicatas
    ConfGridTransp
    ConfGridVendedores
    ConfGrdPedido
    ConfGrdFat
    Call ConfGridProdutosClie
    Call ConfGridProdutosABC
    Call ConfGridCurvaABC(ColumnsIn_ProdutoCurvABC, ProdutoCurvABC_FormatString)
    Call ConfGridClienteCondPgto
    
    Call ConfGridProdutosPedidos
    Call ConfGridProdutosFaturado
    
    '' --------------------------
    optBLPEDNAO.Value = True
    optAVISASIM.Value = True
    optSempNAO.Value = True
    optBloqNao.Value = True
    optTEMCERTINAO.Value = True
    optECLIESIM.Value = True
    '' --------------------------
    
    objCADCLIENTE.PreenchComboBancos cboBANCOS
    objCADCLIENTE.PreenchComboSegmento cboSEGMENTO
    objCADCLIENTE.PreenchComboSegmento cboSEGMENTO2
    objCADCLIENTE.PreenchComboTransportadoras cboTRANSP
    
    cboBANCOS.ListIndex = -1
    cboSEGMENTO.ListIndex = -1
    cboSEGMENTO2.ListIndex = -1
    
    txtOBS.Locked = False
    txtObsCom.Locked = False
    optFATNOVSTEEL(1).Value = True
    optSTATUS(0).Value = True
    optEMP(0).Value = True
    optPermFatSepSN(0).Value = True
    optVISTELENT(0).Value = True
    optDesbClie(1).Value = True

End Sub


Private Sub ConfGridBancos()

    flxBANCOS.Rows = 1
    flxBANCOS.Cols = 5
    
    flxBANCOS.TextMatrix(0, 0) = ""
    flxBANCOS.TextMatrix(0, 1) = "Código"
    flxBANCOS.TextMatrix(0, 2) = "Banco"
    flxBANCOS.TextMatrix(0, 3) = "Agência"
    flxBANCOS.TextMatrix(0, 4) = "C/C"
    
    flxBANCOS.ColWidth(0) = 0
    flxBANCOS.ColWidth(1) = 700
    flxBANCOS.ColWidth(2) = 3000
    flxBANCOS.ColWidth(3) = 1500
    flxBANCOS.ColWidth(4) = 1500
    
    optPermFecOP(0).Value = True
    
End Sub

Private Sub ConfGridRefBancaria()

    flxREFBANCARIA.Rows = 1
    flxREFBANCARIA.Cols = 4
    
    flxREFBANCARIA.TextMatrix(0, 0) = ""
    flxREFBANCARIA.TextMatrix(0, 1) = "Banco"
    flxREFBANCARIA.TextMatrix(0, 2) = "Nome"
    flxREFBANCARIA.TextMatrix(0, 3) = "Telefone"
    
    flxREFBANCARIA.ColWidth(0) = 0
    flxREFBANCARIA.ColWidth(1) = 3000
    flxREFBANCARIA.ColWidth(2) = 1500
    flxREFBANCARIA.ColWidth(3) = 1500
    
End Sub

Private Sub CondGridREfCOMERCIAL()

    flxREFEMPRESA.Rows = 1
    flxREFEMPRESA.Cols = 4
    
    flxREFEMPRESA.TextMatrix(0, 0) = ""
    flxREFEMPRESA.TextMatrix(0, 1) = "Empresa"
    flxREFEMPRESA.TextMatrix(0, 2) = "Nome"
    flxREFEMPRESA.TextMatrix(0, 3) = "Telefone"
    
    flxREFEMPRESA.ColWidth(0) = 0
    flxREFEMPRESA.ColWidth(1) = 3000
    flxREFEMPRESA.ColWidth(2) = 1500
    flxREFEMPRESA.ColWidth(3) = 1500
    
End Sub

Private Sub ConfGriRefPessoal()

    flxREFPESSOAL.Rows = 1
    flxREFPESSOAL.Cols = 3
    
    flxREFPESSOAL.TextMatrix(0, 0) = ""
    flxREFPESSOAL.TextMatrix(0, 1) = "Nome"
    flxREFPESSOAL.TextMatrix(0, 2) = "Telefone"
    
    flxREFPESSOAL.ColWidth(0) = 0
    flxREFPESSOAL.ColWidth(1) = 5000
    flxREFPESSOAL.ColWidth(2) = 1500
    
End Sub

Private Sub ConfGridRestricoes()

    flxRESTRICOES.Rows = 1
    flxRESTRICOES.Cols = 2
    
    flxRESTRICOES.TextMatrix(0, 0) = ""
    flxRESTRICOES.TextMatrix(0, 1) = "Descrição da restrição"
    
    flxRESTRICOES.ColWidth(0) = 0
    flxRESTRICOES.ColWidth(1) = 6000
    
End Sub

Private Sub ConfGridFornec()

    flxFORNEC.Rows = 1
    flxFORNEC.Cols = 3
    
    flxFORNEC.TextMatrix(0, 0) = ""
    flxFORNEC.TextMatrix(0, 1) = "Fornecedor"
    flxFORNEC.TextMatrix(0, 2) = "Desde"
    
    flxFORNEC.ColWidth(0) = 0
    flxFORNEC.ColWidth(1) = 5000
    flxFORNEC.ColWidth(2) = 1500
    
End Sub

Private Sub ConfGridClientes()

    flxEMPRESA.Rows = 1
    flxEMPRESA.Cols = 3
    
    flxEMPRESA.TextMatrix(0, 0) = ""
    flxEMPRESA.TextMatrix(0, 1) = "Cliente"
    flxEMPRESA.TextMatrix(0, 2) = "Desde"
    
    flxEMPRESA.ColWidth(0) = 0
    flxEMPRESA.ColWidth(1) = 5000
    flxEMPRESA.ColWidth(2) = 1500

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub grdCONDPGTO_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdCONDPGTO
          Select Case Col
                 Case conCOL_Cliente_CodCOndPgto
          End Select
     End With
End Sub

Private Sub grdCONDPGTO_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_Cliente_DescCondPgto
         Cancel = True
    Case conCOL_Cliente_CodCOndPgto, _
         conCOL_Cliente_Pesq
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdCONDPGTO.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdCONDPGTO_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If (grdCONDPGTO.Rows - 1) = 0 Then Exit Sub
    
    Dim strINDICE As String
    
    Select Case Col
        Case conCOL_Cliente_Pesq
    
            If cTipOper = "C" Then Exit Sub
            
            ReDim arrCAMPOS(1 To 2, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
            
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       SGI_CODIGO" & vbCrLf
            sSql = sSql & "      ,SGI_DESCRICAO" & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADCONDPGTO" & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL = " & FILIAL
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "N"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "800"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição da Condição de Pagamento"
            arrCAMPOS(2, 4) = "3000"
            arrCAMPOS(2, 5) = "SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Condições de Pagamento")
            
            If Len(Trim(varRETORNO)) > 0 Then
               
                If objBLBFunc.FcVerifItensRepetidos(grdCONDPGTO, Row, conCOL_Cliente_CodCOndPgto, varRETORNO) = False Then
                   MsgBox "Esta Condição de Pagamento ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                   Call LimpaCamposGridCodCondPgto(Row)
                   Exit Sub
                End If
               
                With grdCONDPGTO
                    .Cell(flexcpText, Row, conCOL_Cliente_CodCOndPgto) = varRETORNO
                    .Cell(flexcpText, Row, conCOL_Cliente_DescCondPgto) = PesDescCondPgto(varRETORNO, Row)
                End With
               
            End If
    
    End Select

End Sub

Private Sub grdCONDPGTO_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

On Error GoTo Err_grdCONDPGTO_KeyPressEdit
     
     With grdCONDPGTO
          Select Case Col
                    Case conCOL_Cliente_CodCOndPgto
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsCurrency)
          End Select
     End With
     
     Exit Sub
     
Err_grdCONDPGTO_KeyPressEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdCONDPGTO_KeyPressEdit()", Me.Name, "grdCONDPGTO_KeyPressEdit()", strCAMARQERRO)

End Sub

Private Sub grdCONDPGTO_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

On Error GoTo Err_grdCONDPGTO_ValidateEdit
     
     Dim curVLUNITARIO As Currency
     Dim I As Integer
     Dim strDESCPROD As String
     
     
     With grdCONDPGTO
          Select Case Col
                 Case conCOL_Cliente_CodCOndPgto
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdCONDPGTO, Row, conCOL_Cliente_CodCOndPgto, .EditText) = False Then
                           MsgBox "Esta Condição de Pagamento ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGridCodCondPgto(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        
                        grdCONDPGTO.Cell(flexcpText, Row, conCOL_Cliente_DescCondPgto) = Trim(PesDescCondPgto(.EditText, Row))
                        
                        If Len(grdCONDPGTO.Cell(flexcpText, Row, conCOL_Cliente_DescCondPgto)) = 0 Then
                           MsgBox "Esta Condição de Pagamento não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGridCodCondPgto(Row)
                           Cancel = True
                           Exit Sub
                        End If
                
          End Select
     End With
     
     Exit Sub
     
Err_grdCONDPGTO_ValidateEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProduto_ValidateEdit()", Me.Name, "grdProduto_ValidateEdit()", strCAMARQERRO)
     

End Sub

Private Sub grdCURVAABC_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case 0, _
         1, _
         (grdCURVAABC.Cols - 2), _
         (grdCURVAABC.Cols - 1)
         Cancel = True
    Case Else
        grdCURVAABC.ComboList = ""
    End Select
    Exit Sub

End Sub

Private Sub grdPRODABC_Click()
    Call PopCurvaABC
End Sub

Private Sub grdPRODABC_RowColChange()
    Call PopCurvaABC
End Sub

Private Sub grdPRODUTOS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     Dim I As Integer
     With grdPRODUTOS
          Select Case Col
                 Case conCOL_ProdutoClie_IdProduto
          End Select
     End With
End Sub


Private Sub grdPRODUTOS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_ProdutoClie_IdProduto, _
         conCOL_ProdutoClie_Descricao
         Cancel = True
    Case conCOL_ProdutoClie_Rotulo, _
         conCOL_ProdutoClie_PesqRot
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdPRODUTOS.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdPRODUTOS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    Dim strDESCPROD As String
    
    If (grdPRODUTOS.Rows - 1) = 0 Then Exit Sub
    
    Select Case Col
        Case conCOL_ProdutoClie_PesqRot
    
            If cTipOper = "C" Then Exit Sub
            
            ReDim arrCAMPOS(1 To 4, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
            
            sSql = ""
            
            sSql = "Select Case PRO.SGI_PRODUTOTIPO " & vbCrLf
            sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
            sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
            sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
            sSql = sSql & "                        (Case " & vbCrLf
            sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
            sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
            sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
            sSql = sSql & ",PRO.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & ",PRO.SGI_COMPLEMENTO " & vbCrLf
            sSql = sSql & ",PRO.SGI_CODCLIE " & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO  PRO " & vbCrLf
            sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINHA " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       PRO.SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "   And PRO.SGI_CODCLIE     = " & objCADCLIENTE.CLIECODIGO & vbCrLf
            sSql = sSql & "   And (PRO.SGI_STATUS     = 1 or PRO.SGI_STATUS      = 2)" & vbCrLf
            sSql = sSql & "   And LINHA.SGI_FILIAL    = PRO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And LINHA.SGI_CODLIN    = PRO.SGI_CODLINPROD " & vbCrLf
            
            sSql = sSql & "Union" & vbCrLf
            
            ''If VerificaProdPai(sSql) = False Then
            
            sSql = sSql & "Select Case PRO.SGI_PRODUTOTIPO" & vbCrLf
            sSql = sSql & "             When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
            sSql = sSql & "                         replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
            sSql = sSql & "                         replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
            sSql = sSql & "                         (Case" & vbCrLf
            sSql = sSql & "                               When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
            sSql = sSql & "                               When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End)" & vbCrLf
            sSql = sSql & "             When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO" & vbCrLf
            sSql = sSql & "  ,PRO.SGI_DESCRICAO" & vbCrLf
            sSql = sSql & "  ,PRO.SGI_COMPLEMENTO" & vbCrLf
            sSql = sSql & "  ,PRO.SGI_CODCLIE " & vbCrLf
            
            sSql = sSql & "    From" & vbCrLf
            sSql = sSql & "         SGI_CADPRODUTO   PRO" & vbCrLf
            sSql = sSql & "        ,SGI_PRODATECLIE  PCL" & vbCrLf
            sSql = sSql & "        ,SGI_CADLINHAPRODUTO LINHA " & vbCrLf
            sSql = sSql & "   Where" & vbCrLf
            sSql = sSql & "         PCL.SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "     And PCL.SGI_IDCLIENTE   = " & objCADCLIENTE.CLIECODIGO & vbCrLf
            sSql = sSql & "     And PCL.SGI_FILIAL      = PRO.SGI_FILIAL" & vbCrLf
            sSql = sSql & "     And PCL.SGI_IDPRODUTO   = PRO.SGI_IDPRODUTO" & vbCrLf
            sSql = sSql & "     And (PRO.SGI_STATUS     = 1 Or PRO.SGI_STATUS      = 2)" & vbCrLf
            sSql = sSql & "     And LINHA.SGI_FILIAL    = PRO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "     And LINHA.SGI_CODLIN    = PRO.SGI_CODLINPROD " & vbCrLf
            
            ''End If
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "2000"
            arrCAMPOS(1, 5) = "PRO.SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "PRO.SGI_DESCRICAO"
            
            arrCAMPOS(3, 1) = "SGI_COMPLEMENTO"
            arrCAMPOS(3, 2) = "S"
            arrCAMPOS(3, 3) = "Complemento"
            arrCAMPOS(3, 4) = "3000"
            arrCAMPOS(3, 5) = "PRO.SGI_COMPLEMENTO"
            
            arrCAMPOS(4, 1) = "SGI_CODCLIE"
            arrCAMPOS(4, 2) = "N"
            arrCAMPOS(4, 3) = "Cod.Cliente"
            arrCAMPOS(4, 4) = "1500"
            arrCAMPOS(4, 5) = "PRO.SGI_CODCLIE"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
            
            If Len(Trim(varRETORNO)) > 0 Then
            
               If objBLBFunc.FcVerifItensRepetidos(grdPRODUTOS, Row, conCOL_ProdutoClie_Rotulo, varRETORNO) = False Then
                    MsgBox "Este Produto já foi relacionado na Gride !!!", vbOKOnly + vbExclamation
                    Call LimpaCamposGrid(Row)
                    Exit Sub
               End If
            
               grdPRODUTOS.Cell(flexcpText, Row, conCOL_ProdutoClie_Rotulo) = Trim(varRETORNO)
               grdPRODUTOS.Cell(flexcpText, Row, conCOL_ProdutoClie_IdProduto) = PegaIDProduto(Trim(varRETORNO))
               
               strDESCPROD = PegaDescrProduto(grdPRODUTOS.Cell(flexcpText, Row, conCOL_ProdutoClie_IdProduto))
               If Len(Trim(strDESCPROD)) = 0 Then
                    Call LimpaCamposGrid(Row)
                    Exit Sub
               End If
               grdPRODUTOS.Cell(flexcpText, Row, conCOL_ProdutoClie_Descricao) = strDESCPROD
                
            End If
    
    End Select

End Sub

Private Sub grdPRODUTOS_Click()
    Call MostraItens
End Sub

Private Sub grdPRODUTOS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdPRODUTOS
          Select Case Col
                    Case conCOL_ProdutoClie_Rotulo
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
          End Select
     End With
End Sub

Private Sub grdPRODUTOS_RowColChange()
    Call MostraItens
End Sub

Private Sub grdPRODUTOS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdPRODUTOS
          Select Case Col
                 Case conCOL_ProdutoClie_Rotulo
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdPRODUTOS, Row, conCOL_ProdutoClie_Rotulo, .EditText) = False Then
                           MsgBox "Este produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGrid(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_ProdutoClie_IdProduto) = PegaIDProduto(Trim(.EditText))
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_Produto_IdProduto))) = 0 Then
                           MsgBox "Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
                           Cancel = True
                           Exit Sub
                        End If
                        
                        If Len(Trim(PegaDescrProduto(.Cell(flexcpText, Row, conCOL_ProdutoClie_IdProduto)))) = 0 Then
                           MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_ProdutoClie_Descricao) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_ProdutoClie_IdProduto) = PegaDescrProduto(.Cell(flexcpText, Row, conCOL_ProdutoClie_IdProduto))
          
          End Select
     End With

End Sub

Private Sub grdULTFAT_DblClick()
    Call AbreTelas
End Sub

Private Sub mskDTCADASTRO_GotFocus()
    objBLBFunc.SelecionaCampos mskDTCADASTRO.Name, frmCADCLIENTE
End Sub

Private Sub mskDTCLIENTE_GotFocus()
    objBLBFunc.SelecionaCampos mskDTCLIENTE.Name, frmCADCLIENTE
End Sub

Private Sub mskDTEMPRESA_GotFocus()
    objBLBFunc.SelecionaCampos mskDTEMPRESA.Name, frmCADCLIENTE
End Sub

Private Sub Option1_Click()
    Frame27.Enabled = True
End Sub

Private Sub Option2_Click()
    ConfGridCertificado
    Frame27.Enabled = False
End Sub

Private Sub optEMP_Click(Index As Integer)
    If (grdPRODABC.Rows - 1) = 0 Then Exit Sub
    Call PopCurvaABC
End Sub

Private Sub optFATNOVSTEEL_Click(Index As Integer)
    Call PegaDadosFatItens(Index)
End Sub

Private Sub optFILIALPRODFAT_Click(Index As Integer)
    Call ConfGrdUltFat
    If optSTATUS(0).Value = True Then Call PopGrdUltFat
    If optSTATUS(1).Value = True Then Call PopTodosPed
End Sub

Private Sub optSempNAO_Click()
    If cTipOper <> "C" Then Frame7.Enabled = True
End Sub

Private Sub optSempSIM_Click()
    Frame7.Enabled = False
    txtSALDOACIMA.Text = ""
    txtBLOQPEDSALDO.Text = ""
    optBloqNao.Value = True
End Sub

Private Sub optSTATUS_Click(Index As Integer)
    Call ConfGrdUltFat
    If Index = 0 Then Call PopGrdUltFat
    If Index = 1 Then Call PopTodosPed
End Sub

Private Sub optTipoNOVASTEL_Click(Index As Integer)
    Call PopPedidos(Index)
End Sub

Private Sub stCLIENTE_Click(PreviousTab As Integer)
    
    If cTipOper = "C" Then Exit Sub
    
    If stCLIENTE.Tab = 0 Then
       If txtCPFCNPJ.Enabled And txtCPFCNPJ.Visible Then txtCPFCNPJ.SetFocus
    ElseIf stCLIENTE.Tab = 1 Then
       If Frame3.Enabled = True Then txtENDCOBR.SetFocus
    ElseIf stCLIENTE.Tab = 2 Then
       stFin.Tab = 0
       If Frame6.Enabled = True Then txtCODBANCO.SetFocus
    ElseIf stCLIENTE.Tab = 3 Then
       If Frame5.Enabled = True Then txtVLLIMCRED.SetFocus
    ElseIf stCLIENTE.Tab = 5 Then
       If Frame26.Enabled = True Then txtCODSEQ.SetFocus
    ElseIf stCLIENTE.Tab = 6 Then
       If Frame31.Enabled = True Then txtOBS.SetFocus
    End If
    
End Sub

Private Sub stFin_Click(PreviousTab As Integer)
    If stFin.Tab = 0 Then
       If Frame6.Enabled = True Then txtCODBANCO.SetFocus
    ElseIf stFin.Tab = 1 Then
       If Frame10.Enabled = True Then txtRefBanco.SetFocus
    ElseIf stFin.Tab = 2 Then
       If Frame17.Enabled = True Then txtRESTRICOES.SetFocus
    ElseIf stFin.Tab = 3 Then
       If Frame8.Enabled = True Then txtFORNEC.SetFocus
    ElseIf stFin.Tab = 4 Then
       If Frame9.Enabled = True Then txtCLIENTE.SetFocus
    End If
End Sub

Private Sub stREFDIVER_Click(PreviousTab As Integer)
    If stREFDIVER.Tab = 0 Then
       If Frame10.Enabled = True Then txtRefBanco.SetFocus
    ElseIf stREFDIVER.Tab = 1 Then
       If Frame13.Enabled = True Then txtREFCOM.SetFocus
    ElseIf stREFDIVER.Tab = 2 Then
       If Frame15.Enabled = True Then txtREFPESSOAL.SetFocus
    End If
End Sub


Private Sub AddGridRefBancaria()

    Dim I As Integer
    
    If Len(Trim(txtRefBanco.Text)) = 0 Then
       MsgBox "Informe o banco de referência !!!", vbOKOnly + vbCritical, "Aviso"
       txtRefBanco.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxREFBANCARIA.Rows - 1)
        If flxREFBANCARIA.TextMatrix(I, 1) = txtRefBanco.Text Then
           MsgBox "Este banco já foi cadastrado !!!", vbOKOnly + vbCritical, "Aviso"
           txtRefBanco.Text = ""
           txtRefNome.Text = ""
           txtTelREF.Text = ""
           txtRefBanco.SetFocus
           Exit Sub
        End If
    Next I
    
    flxREFBANCARIA.AddItem "" & vbTab & txtRefBanco.Text & vbTab & txtRefNome.Text & vbTab & txtTelREF.Text
    
    txtRefBanco.Text = ""
    txtRefNome.Text = ""
    txtTelREF.Text = ""
    txtRefBanco.SetFocus

End Sub

Private Sub txRGCGC_GotFocus()
    objBLBFunc.SelecionaCampos txRGCGC.Name, frmCADCLIENTE
End Sub

Private Sub txRGCGC_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtATENDIDO_GotFocus()
    objBLBFunc.SelecionaCampos txtATENDIDO.Name, frmCADCLIENTE
End Sub

Private Sub txtATENDIDO_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtBAICOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtBAICOBR.Name, frmCADCLIENTE
End Sub

Private Sub txtBAICOBR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtBAIENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtBAIENTR.Name, frmCADCLIENTE
End Sub

Private Sub txtBAIENTR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtBAINOM_GotFocus()
    objBLBFunc.SelecionaCampos txtBAINOM.Name, frmCADCLIENTE
End Sub

Private Sub txtBAINOM_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtBLOQPEDSALDO_GotFocus()
    objBLBFunc.SelecionaCampos txtBLOQPEDSALDO.Name, frmCADCLIENTE
End Sub

Private Sub txtBLOQPEDSALDO_Validate(Cancel As Boolean)

    If Len(Trim(txtBLOQPEDSALDO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtBLOQPEDSALDO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtBLOQPEDSALDO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtBLOQPEDSALDO.Text = Format(txtBLOQPEDSALDO.Text, "#,##0.00")

End Sub

Private Sub txtCEPCOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtCEPCOBR.Name, frmCADCLIENTE
End Sub

Private Sub txtCEPCOBR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCEPENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtCEPENTR.Name, frmCADCLIENTE
End Sub

Private Sub txtCEPENTR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCEPNORM_GotFocus()
    objBLBFunc.SelecionaCampos txtCEPNORM.Name, frmCADCLIENTE
End Sub

Private Sub txtCEPNORM_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCIDCOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCOBR.Name, frmCADCLIENTE
End Sub

Private Sub txtCIDCOBR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCIDENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDENTR.Name, frmCADCLIENTE
End Sub

Private Sub txtCIDENTR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCIDNORM_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDNORM.Name, frmCADCLIENTE
End Sub

Private Sub txtCIDNORM_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCLIENTE_GotFocus()
    objBLBFunc.SelecionaCampos txtCLIENTE.Name, frmCADCLIENTE
End Sub

Private Sub txtCLIENTE_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCODBANCO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODBANCO.Name, frmCADCLIENTE
End Sub

Private Sub txtCODBANCO_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blACHOU As Boolean
    
    If Len(Trim(txtCODBANCO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODBANCO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODBANCO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    blACHOU = False
    For I = 0 To (cboBANCOS.ListCount - 1)
        If CInt(txtCODBANCO.Text) = cboBANCOS.ItemData(I) Then
           blACHOU = True
           cboBANCOS.ListIndex = I
        End If
    Next I
    
    If blACHOU = False Then
       MsgBox "Este banco não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCODBANCO.Text = ""
       cboBANCOS.ListIndex = -1
       Cancel = True
       Exit Sub
    End If

End Sub



Private Sub txtCODSEQ_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODSEQ.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODSEQ.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSEQ.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    For I = 0 To (cboSEGMENTO.ListCount - 1)
        If cboSEGMENTO.ItemData(I) = CInt(txtCODSEQ.Text) Then cboSEGMENTO.ListIndex = I
    Next I
       
End Sub

Private Sub txtCODSUFRAMA_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub
Private Sub txtCODTRANSP_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTRANSP.Name, frmCADCLIENTE
End Sub

Private Sub txtCODTRANSP_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blACHOU As Boolean
    
    If Len(Trim(txtCODTRANSP.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTRANSP.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODTRANSP.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    blACHOU = False
    For I = 0 To (cboTRANSP.ListCount - 1)
        If CInt(txtCODTRANSP.Text) = cboTRANSP.ItemData(I) Then
           blACHOU = True
           cboTRANSP.ListIndex = I
        End If
    Next I
    
    If blACHOU = False Then
       MsgBox "Este banco não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCODTRANSP.Text = ""
       cboTRANSP.ListIndex = -1
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCPFCNPJ_GotFocus()
    objBLBFunc.SelecionaCampos txtCPFCNPJ.Name, frmCADCLIENTE
End Sub

Private Sub txtCPFCNPJ_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtENDCOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtENDCOBR.Name, frmCADCLIENTE
End Sub

Private Sub txtENDCOBR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtENDENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtENDENTR.Name, frmCADCLIENTE
End Sub

Private Sub txtENDENTR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtENDNORM_GotFocus()
    objBLBFunc.SelecionaCampos txtENDNORM.Name, frmCADCLIENTE
End Sub

Private Sub txtENDNORM_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFORNEC_GotFocus()
    objBLBFunc.SelecionaCampos txtFORNEC.Name, frmCADCLIENTE
End Sub

Private Sub txtFORNEC_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtMeseReavali_GotFocus()
    objBLBFunc.SelecionaCampos txtMeseReavali.Name, frmCADCLIENTE
End Sub

Private Sub txtMeseReavali_Validate(Cancel As Boolean)
    
    If Len(Trim(txtMeseReavali.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtMeseReavali.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtMeseReavali.Text = ""
       Cancel = True
       Exit Sub
    End If
    
End Sub

Private Sub txtNOMFANTA_GotFocus()
    objBLBFunc.SelecionaCampos txtNOMFANTA.Name, frmCADCLIENTE
End Sub

Private Sub txtNOMFANTA_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub






Private Sub txtRAZAOSOC_GotFocus()
    objBLBFunc.SelecionaCampos txtRAZAOSOC.Name, frmCADCLIENTE
End Sub

Private Sub txtRAZAOSOC_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtRefBanco_GotFocus()
    objBLBFunc.SelecionaCampos txtRefBanco.Name, frmCADCLIENTE
End Sub

Private Sub txtRefBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtREFCOM_GotFocus()
    objBLBFunc.SelecionaCampos txtREFCOM.Name, frmCADCLIENTE
End Sub

Private Sub txtREFCOM_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtREFCOMNOME_GotFocus()
    objBLBFunc.SelecionaCampos txtREFCOMNOME.Name, frmCADCLIENTE
End Sub

Private Sub txtREFCOMNOME_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtRefNome_GotFocus()
    objBLBFunc.SelecionaCampos txtRefNome.Name, frmCADCLIENTE
End Sub

Private Sub txtRefNome_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtREFPESSOAL_GotFocus()
    objBLBFunc.SelecionaCampos txtREFPESSOAL.Name, frmCADCLIENTE
End Sub

Private Sub txtREFPESSOAL_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtREFTELPESSOAL_GotFocus()
    objBLBFunc.SelecionaCampos txtREFTELPESSOAL.Name, frmCADCLIENTE
End Sub

Private Sub txtREFTELPESSOAL_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtRESTRICOES_GotFocus()
    objBLBFunc.SelecionaCampos txtRESTRICOES.Name, frmCADCLIENTE
End Sub

Private Sub txtRESTRICOES_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtSALDOACIMA_GotFocus()
    objBLBFunc.SelecionaCampos txtSALDOACIMA.Name, frmCADCLIENTE
End Sub

Private Sub txtSALDOACIMA_Validate(Cancel As Boolean)

    If Len(Trim(txtSALDOACIMA.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtSALDOACIMA.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtSALDOACIMA.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtSALDOACIMA.Text = Format(txtSALDOACIMA.Text, "#,##0.00")

End Sub

Private Sub txtSISTCERTIFI_GotFocus()
    objBLBFunc.SelecionaCampos txtSISTCERTIFI.Name, frmCADCLIENTE
End Sub

Private Sub txtSISTCERTIFI_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtTelREF_GotFocus()
    objBLBFunc.SelecionaCampos txtTelREF.Name, frmCADCLIENTE
End Sub

Private Sub txtTelREF_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtTELREFCOM_GotFocus()
    objBLBFunc.SelecionaCampos txtTELREFCOM.Name, frmCADCLIENTE
End Sub

Private Sub txtTELREFCOM_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtVLLIMCRED_GotFocus()
    objBLBFunc.SelecionaCampos txtVLLIMCRED.Name, frmCADCLIENTE
End Sub

Private Sub txtVLLIMCRED_Validate(Cancel As Boolean)
    
    If Len(Trim(txtVLLIMCRED.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtVLLIMCRED.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical
       txtVLLIMCRED.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtVLLIMCRED.Text = Format(txtVLLIMCRED.Text, "#,##0.00")
    
End Sub

Private Sub AddGridRefEmpresa()

    Dim I As Integer
    
    If Len(Trim(txtREFCOM.Text)) = 0 Then
       MsgBox "Informe a empresa da referência !!!", vbOKOnly + vbCritical, "Aviso"
       txtREFCOM.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxREFEMPRESA.Rows - 1)
        If flxREFEMPRESA.TextMatrix(I, 1) = txtREFCOM.Text Then
           MsgBox "Este empresa já foi cadastrada !!!", vbOKOnly + vbCritical, "Aviso"
           txtREFCOM.Text = ""
           txtREFCOMNOME.Text = ""
           txtTELREFCOM.Text = ""
           txtREFCOM.SetFocus
           Exit Sub
        End If
    Next I
    
    flxREFEMPRESA.AddItem "" & vbTab & txtREFCOM.Text & vbTab & txtREFCOMNOME.Text & vbTab & txtTELREFCOM.Text
    
    txtREFCOM.Text = ""
    txtREFCOMNOME.Text = ""
    txtTELREFCOM.Text = ""
    txtREFCOM.SetFocus

End Sub

Private Sub addGridRefPessoal()

    Dim I As Integer
    
    If Len(Trim(txtREFPESSOAL.Text)) = 0 Then
       MsgBox "Informe o nome da referência !!!", vbOKOnly + vbCritical, "Aviso"
       txtREFPESSOAL.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxREFPESSOAL.Rows - 1)
        If flxREFPESSOAL.TextMatrix(I, 1) = txtREFPESSOAL.Text Then
           MsgBox "Este nome já foi cadastrado !!!", vbOKOnly + vbCritical, "Aviso"
           txtREFPESSOAL.Text = ""
           txtREFTELPESSOAL.Text = ""
           txtREFPESSOAL.SetFocus
           Exit Sub
        End If
    Next I
    
    flxREFPESSOAL.AddItem "" & vbTab & txtREFPESSOAL.Text & vbTab & txtREFTELPESSOAL.Text
    
    txtREFPESSOAL.Text = ""
    txtREFTELPESSOAL.Text = ""
    txtREFPESSOAL.SetFocus

End Sub

Public Sub AddGridRestricoes()

    Dim I As Integer
    
    If Len(Trim(txtRESTRICOES.Text)) = 0 Then
       MsgBox "Informe a restrição !!!", vbOKOnly + vbCritical, "Aviso"
       txtRESTRICOES.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxRESTRICOES.Rows - 1)
        If flxRESTRICOES.TextMatrix(I, 1) = txtRESTRICOES.Text Then
           MsgBox "Esta restrição já foi cadastrada !!!", vbOKOnly + vbCritical, "Aviso"
           txtRESTRICOES.Text = ""
           txtRESTRICOES.SetFocus
           Exit Sub
        End If
    Next I
    
    flxRESTRICOES.AddItem "" & vbTab & txtRESTRICOES.Text
    
    txtRESTRICOES.Text = ""
    txtRESTRICOES.SetFocus

End Sub

Private Sub AddGridFornec()

    Dim I As Integer
    
    If Len(Trim(txtFORNEC.Text)) = 0 Then
       MsgBox "Informe o fornecedor !!!", vbOKOnly + vbCritical, "Aviso"
       txtFORNEC.SetFocus
       Exit Sub
    End If
    
    If mskDTEMPRESA.Text <> "__/__/____" Then
       If Not IsDate(mskDTEMPRESA.Text) Then
          MsgBox "Data Inválida !!!", vbOKOnly + vbCritical, "aviso"
          mskDTEMPRESA.Text = "__/__/____"
          mskDTEMPRESA.SetFocus
          Exit Sub
       End If
    End If
    
    For I = 1 To (flxFORNEC.Rows - 1)
        If flxFORNEC.TextMatrix(I, 1) = txtFORNEC.Text Then
           MsgBox "Este fornecedor já foi cadastrado !!!", vbOKOnly + vbCritical, "Aviso"
           txtFORNEC.Text = ""
           mskDTEMPRESA.Text = "__/__/____"
           txtFORNEC.SetFocus
           Exit Sub
        End If
    Next I
    
    If mskDTEMPRESA.Text = "__/__/____" Then
       flxFORNEC.AddItem "" & vbTab & txtFORNEC.Text & vbTab & ""
    Else
       flxFORNEC.AddItem "" & vbTab & txtFORNEC.Text & vbTab & "" & mskDTEMPRESA.Text
    End If
    
    txtFORNEC.Text = ""
    mskDTEMPRESA.Text = "__/__/____"
    
    txtFORNEC.SetFocus

End Sub

Private Sub AddGridClientes()

    Dim I As Integer
    
    If Len(Trim(txtCLIENTE.Text)) = 0 Then
       MsgBox "Informe o cliente !!!", vbOKOnly + vbCritical, "Aviso"
       txtCLIENTE.SetFocus
       Exit Sub
    End If
    
    If mskDTCLIENTE.Text <> "__/__/____" Then
       If Not IsDate(mskDTCLIENTE.Text) Then
          MsgBox "Data Inválida !!!", vbOKOnly + vbCritical, "aviso"
          mskDTCLIENTE.Text = "__/__/____"
          mskDTCLIENTE.SetFocus
          Exit Sub
       End If
    End If
    
    For I = 1 To (flxEMPRESA.Rows - 1)
        If flxEMPRESA.TextMatrix(I, 1) = txtCLIENTE.Text Then
           MsgBox "Este cliente já foi cadastrado !!!", vbOKOnly + vbCritical, "Aviso"
           txtCLIENTE.Text = ""
           mskDTCLIENTE.Text = "__/__/____"
           txtCLIENTE.SetFocus
           Exit Sub
        End If
    Next I
    
    If mskDTCLIENTE.Text = "__/__/____" Then
       flxEMPRESA.AddItem "" & vbTab & txtCLIENTE.Text & vbTab & ""
    Else
       flxEMPRESA.AddItem "" & vbTab & txtCLIENTE.Text & vbTab & "" & mskDTCLIENTE.Text
    End If
    
    txtCLIENTE.Text = ""
    mskDTCLIENTE.Text = "__/__/____"
    
    txtCLIENTE.SetFocus

End Sub

Private Sub ConfGridHistAvaliacao()

    flxSTATUSAVALIACAO.Rows = 0
    flxSTATUSAVALIACAO.Cols = 0
    
End Sub

Private Function Verifica_Campos() As Boolean

    Verifica_Campos = False
    
    Dim dblVLLIMCRED As Double
    Dim dblSALDACIMA As Double
    Dim dblBLQSALDO  As Double
    Dim strCNPJ      As String
    
    If Len(Trim(txtCPFCNPJ.Text)) > 0 Then
    
       strCNPJ = Trim(Replace(Replace(Replace(txtCPFCNPJ.Text, ".", ""), "-", ""), "/", ""))
       strCNPJ = Trim(Replace(strCNPJ, ",", ""))
       
       If Not IsNumeric(strCNPJ) Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "aviso"
          txtCPFCNPJ.Text = ""
          stCLIENTE.Tab = 0
          txtCPFCNPJ.SetFocus
          Exit Function
       End If
       
       If Len(Trim(strCNPJ)) < 14 And Len(Trim(strCNPJ)) > 11 Then
          MsgBox "CPF/CNPJ Inválido !!!", vbOKOnly + vbCritical, "Aviso"
          txtCPFCNPJ.Text = ""
          stCLIENTE.Tab = 0
          txtCPFCNPJ.SetFocus
          Exit Function
       End If
       
       If Len(Trim(strCNPJ)) < 11 Then
          MsgBox "CPF/CNPJ Inválido !!!", vbOKOnly + vbCritical, "Aviso"
          txtCPFCNPJ.Text = ""
          stCLIENTE.Tab = 0
          txtCPFCNPJ.SetFocus
          Exit Function
       End If
       
       If Len(Trim(strCNPJ)) = 11 Then
          If objBLBFunc.ViewCPF(strCNPJ) = False Then
             MsgBox "CPF Inválido !!!", vbOKOnly + vbCritical, "Aviso"
             txtCPFCNPJ.Text = ""
             stCLIENTE.Tab = 0
             txtCPFCNPJ.SetFocus
             Exit Function
          End If
       End If
       
       If Len(Trim(strCNPJ)) = 14 Then
          If objBLBFunc.ViewCGC(strCNPJ) = False Then
             MsgBox "CNPJ Inválido !!!", vbOKOnly + vbCritical, "aviso"
             txtCPFCNPJ.Text = ""
             stCLIENTE.Tab = 0
             txtCPFCNPJ.SetFocus
             Exit Function
          End If
       End If
       
    Else
       '' MsgBox "Informe o CPF/CNPJ !!!", vbOKOnly + vbExclamation, "Aviso"
       '' stCLIENTE.Tab = 0
       '' txtCPFCNPJ.SetFocus
        ''Exit Function
    End If
    
    If Len(Trim(txtRAZAOSOC.Text)) = 0 Then
       MsgBox "Razão social inválida !!!", vbOKOnly + vbCritical, "aviso"
       stCLIENTE.Tab = 0
       txtRAZAOSOC.SetFocus
       Exit Function
    End If
    
    If (optFisica.Value = False And opfJuridica.Value = False) Then
       MsgBox "Informar se pessoa é Fisica ou Juridica !!!", vbOKOnly + vbCritical, "aviso"
       stCLIENTE.Tab = 0
       optFisica.SetFocus
       Exit Function
    End If
    
    If Not IsDate(mskDTCADASTRO.Text) Then
       MsgBox "Data de cadastro inválida !!!", vbOKOnly + vbCritical, "aviso"
       stCLIENTE.Tab = 0
       mskDTCADASTRO.SetFocus
       Exit Function
    End If
    ''If mshDTNIRC.Text <> "__/__/____" Then
    ''   If Not IsDate(mshDTNIRC.Text) Then
    ''      MsgBox "Data do NIRC inválida !!!", vbOKOnly + vbCritical, "aviso"
    ''      stCLIENTE.Tab = 0
    ''      mshDTNIRC.SetFocus
    ''      Exit Function
    ''   End If
    ''End If
    
    If cTipOper = "I" Then
       
       If Len(Trim(txtCPFCNPJ.Text)) > 0 Then
          
          strCNPJ = Trim(Replace(Replace(Replace(txtCPFCNPJ.Text, ".", ""), "-", ""), "/", ""))
          strCNPJ = Trim(Replace(strCNPJ, ",", ""))
          
          '' Verifica se existe CPF/CNPJ
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CPFCNPJ = '" & Trim(strCNPJ) & "'"
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          
          If Not BREC.EOF Then
             MsgBox "Este CPF ou CNPJ já existe !!!", vbOKOnly + vbCritical, "aviso"
             BREC.Close
             stCLIENTE.Tab = 0
             txtCPFCNPJ.SetFocus
             Exit Function
          End If
          
          BREC.Close
          
       End If
       
       If Len(Trim(txRGCGC.Text)) > 0 Then
       
          '' Verifica se existe RG/INNSC.ESTADUAL
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_RGCGC   = '" & txRGCGC.Text & "'"
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          
          If Not BREC.EOF Then
             MsgBox "Este RG ou Insc. Estadual já existe !!!", vbOKOnly + vbCritical, "aviso"
             BREC.Close
             stCLIENTE.Tab = 0
             txRGCGC.SetFocus
             Exit Function
          End If
          
          BREC.Close
       
       End If
       
       If Len(Trim(txtRAZAOSOC.Text)) > 0 Then
          
          '' Verifica se existe a razão social
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_RAZAOSOC = '" & txtRAZAOSOC.Text & "'"
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          
          If Not BREC.EOF Then
             MsgBox "Esta razão social já existe !!!", vbOKOnly + vbCritical, "aviso"
             BREC.Close
             stCLIENTE.Tab = 0
             txtRAZAOSOC.SetFocus
             Exit Function
          End If
          
          BREC.Close
       
       End If
       
       If Len(Trim(txtCodRef.Text)) > 0 Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CODREF = '" & txtCodRef.Text & "'"
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Este Código de Referência já existe !!!", vbOKOnly + vbCritical, "aviso"
             BREC.Close
             stCLIENTE.Tab = 0
             txtCodRef.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
       
    End If
    
    If cTipOper = "A" Then
    
       strCNPJ = Trim(Replace(Replace(Replace(txtCPFCNPJ.Text, ".", ""), "-", ""), "/", ""))
       strCNPJ = Trim(Replace(strCNPJ, ",", ""))
       
       If (Len(Trim(strCNPJ)) > 0) And (strCNPJ <> objCADCLIENTE.CLICPFCNPJ) Then
          
          '' Verifica se existe CPF/CNPJ
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CPFCNPJ = '" & strCNPJ & "'"
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          
          If Not BREC.EOF Then
             MsgBox "Este CPF ou CNPJ já existe !!!", vbOKOnly + vbCritical, "aviso"
             BREC.Close
             stCLIENTE.Tab = 0
             txtCPFCNPJ.Text = objBLBFunc.FormataCnpj(objCADCLIENTE.CLICPFCNPJ)
             txtCPFCNPJ.SetFocus
             Exit Function
          End If
          
          BREC.Close
       
       End If
       
       If (Len(Trim(txRGCGC.Text)) > 0) And (txRGCGC.Text <> objCADCLIENTE.CLIRGCGC) Then
       
          '' Verifica se existe RG/INNSC.ESTADUAL
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_RGCGC   = '" & txRGCGC.Text & "'"
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          
          If Not BREC.EOF Then
             MsgBox "Este RG ou Insc. Estadual já existe !!!", vbOKOnly + vbCritical, "aviso"
             BREC.Close
             stCLIENTE.Tab = 0
             txRGCGC.SetFocus
             txRGCGC.Text = objCADCLIENTE.CLIRGCGC
             Exit Function
          End If
          
          BREC.Close
       
       End If
       
       If (Len(Trim(txtRAZAOSOC.Text)) > 0) And (txtRAZAOSOC.Text <> objCADCLIENTE.CLIERZAOSO) Then
          
          '' Verifica se existe a razão social
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_RAZAOSOC = '" & txtRAZAOSOC.Text & "'"
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          
          If Not BREC.EOF Then
             MsgBox "Esta razão social já existe !!!", vbOKOnly + vbCritical, "aviso"
             BREC.Close
             stCLIENTE.Tab = 0
             txtRAZAOSOC.Text = objCADCLIENTE.CLIERZAOSO
             txtRAZAOSOC.SetFocus
             Exit Function
          End If
          
          BREC.Close
       
       End If
       
       
       If objCADCLIENTE.CODREF <> txtCodRef.Text Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CODREF = '" & txtCodRef.Text & "'"
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Este Código de Referência já existe !!!", vbOKOnly + vbCritical, "aviso"
             BREC.Close
             stCLIENTE.Tab = 0
             txtCodRef.Text = objCADCLIENTE.CODREF
             txtCodRef.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
    
    End If
    
    If Len(Trim(txtVLLIMCRED.Text)) > 0 Then dblVLLIMCRED = CDbl(txtVLLIMCRED.Text)
    If Len(Trim(txtSALDOACIMA.Text)) > 0 Then dblSALDACIMA = CDbl(txtSALDOACIMA.Text)
    If Len(Trim(txtBLOQPEDSALDO.Text)) > 0 Then dblBLQSALDO = CDbl(txtBLOQPEDSALDO.Text)
    
    If dblSALDACIMA >= dblVLLIMCRED And dblSALDACIMA > 0 Then
       MsgBox "O campo <Avisar quando o saldo estiver a cima de :> não pode ser maior que o campo limite de compras !!!", vbOKOnly + vbCritical, "aviso"
       txtSALDOACIMA.SetFocus
       Exit Function
    End If
    
    If dblBLQSALDO >= dblVLLIMCRED And dblBLQSALDO > 0 Then
       MsgBox "O campo <Bloqueia com saldo a cima de :> não pode ser maior que o campo Limite de Crédito !!!", vbOKOnly + vbCritical, "aviso"
       txtBLOQPEDSALDO.SetFocus
       Exit Function
    End If
    
    Verifica_Campos = True

End Function

Private Sub ConfGridCertificado()
    
    flxSISTSERTFIC.Rows = 1
    flxSISTSERTFIC.Cols = 2
    
    flxSISTSERTFIC.TextMatrix(0, 0) = ""
    flxSISTSERTFIC.TextMatrix(0, 1) = "Certificação"
    
    flxSISTSERTFIC.ColWidth(0) = 0
    flxSISTSERTFIC.ColWidth(1) = 5000
    
End Sub

Private Sub ConfGridEmpresaAtende()

    flxATENDIDO.Rows = 1
    flxATENDIDO.Cols = 4
    
    flxATENDIDO.TextMatrix(0, 0) = ""
    flxATENDIDO.TextMatrix(0, 1) = "Empresa"
    flxATENDIDO.TextMatrix(0, 2) = "Ramo de atividade"
    flxATENDIDO.TextMatrix(0, 3) = ""
    
    flxATENDIDO.ColWidth(0) = 0
    flxATENDIDO.ColWidth(1) = 4000
    flxATENDIDO.ColWidth(2) = 3000
    flxATENDIDO.ColWidth(3) = 0
    
End Sub

Private Sub InsertSistCertif()

    Dim I As Integer
    
    If Len(Trim(txtSISTCERTIFI.Text)) = 0 Then
       MsgBox "Informe o sistema de certificação !!!", vbOKOnly + vbCritical, "aviso"
       txtSISTCERTIFI.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxSISTSERTFIC.Rows - 1)
       If flxSISTSERTFIC.TextMatrix(I, 1) = txtSISTCERTIFI.Text Then
          MsgBox "Este sistema de certificação já existe !!!", vbOKOnly + vbCritical, "aviso"
          txtSISTCERTIFI.SetFocus
          Exit Sub
       End If
    Next I
    
    
    flxSISTSERTFIC.AddItem "" & vbTab & txtSISTCERTIFI.Text
    txtSISTCERTIFI.Text = ""
    txtSISTCERTIFI.SetFocus

End Sub

Private Sub UnseriEmpreAtende()

    Dim I As Integer
    
    If Len(Trim(txtATENDIDO.Text)) = 0 Then
       MsgBox "informe a empresa !!!", vbOKOnly + vbCritical, "aviso"
       txtATENDIDO.SetFocus
       Exit Sub
    End If
    
    If cboSEGMENTO2.ListIndex = -1 Then
       MsgBox "informe o segmento !!!", vbOKOnly + vbCritical, "aviso"
       cboSEGMENTO2.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxATENDIDO.Rows - 1)
        If flxATENDIDO.TextMatrix(0, 1) = txtATENDIDO.Text Then
           MsgBox "Esta empresa já existe !!!", vbOKOnly + vbCritical, "aviso"
           txtATENDIDO.Text = ""
           cboSEGMENTO2.Text = ""
           cboSEGMENTO2.ListIndex = -1
           txtATENDIDO.SetFocus
           Exit Sub
        End If
    Next I
    
    flxATENDIDO.AddItem "" & vbTab & txtATENDIDO.Text & vbTab & cboSEGMENTO2.Text & vbTab & cboSEGMENTO2.ItemData(cboSEGMENTO2.ListIndex)
    txtATENDIDO.Text = ""
    cboSEGMENTO2.Text = ""
    cboSEGMENTO2.ListIndex = -1
    txtATENDIDO.SetFocus
       

End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    '' Dados Cadastrais **
    Frame2.Enabled = True
    txtCPFCNPJ.Enabled = False
    txtRAZAOSOC.Enabled = False
    txRGCGC.Enabled = False
    txtNOMFANTA.Enabled = False
    optFisica.Enabled = False
    opfJuridica.Enabled = False
    optECLIESIM.Enabled = False
    optECLIENAO.Enabled = False
    txtENDNORM.Enabled = False
    txtBAINOM.Enabled = False
    txtCIDNORM.Enabled = False
    cboESTNORM.Enabled = False
    txtCEPNORM.Enabled = False
    
    cboTELNORM.Locked = True
    cboCONTNORM.Locked = True
    cboEMAILNORM.Locked = True
    cboEMAILNORM.Locked = True
    
    mskDTCADASTRO.Enabled = False
    
    txtOBS.Locked = True
    txtObsCom.Locked = True
    
    Frame33.Enabled = False
    '' -----------------------
    
    '' Cobrança /  Entrega
    Frame3.Enabled = False
    Frame4.Enabled = False
    '' -----------------------
    
    '' Financeiro
    Frame6.Enabled = False
    Frame10.Enabled = False
    Frame13.Enabled = False
    Frame15.Enabled = False
    Frame17.Enabled = False
    Frame18.Enabled = False
    Frame19.Enabled = False
    Frame8.Enabled = False
    Frame9.Enabled = False
    '' -----------------------
    
    '' Crédito **
    Frame5.Enabled = False
    Frame24.Enabled = False
    Frame7.Enabled = False
    '' -----------------------
    
    '' Folhow UP
    Frame26.Enabled = False
    Frame27.Enabled = False
    Frame28.Enabled = True
    Frame29.Enabled = False
    Frame30.Enabled = True
    '' -----------------------
    
    '' Transportadora
    Frame40.Enabled = False
    Frame41.Enabled = True
    
    Me.Caption = "Cadastro de clientes - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADCLIENTE
    
    objBLBFunc.Preenche_Estado cboESTNORM
    objBLBFunc.Preenche_Estado cboESTCOBR
    objBLBFunc.Preenche_Estado cboESTENTR
    
    '' --------------------------
    
    ConfGridBancos
    ConfGridRefBancaria
    CondGridREfCOMERCIAL
    ConfGriRefPessoal
    ConfGridRestricoes
    ConfGridFornec
    ConfGridClientes
    ConfGridHistAvaliacao
    ConfGridCertificado
    ConfGridEmpresaAtende
    ConfGridDuplicatas
    ConfGridTransp
    ConfGridVendedores
    ConfGrdPedido
    ConfGrdFat
    ConfGridProdutos
    ConfGridProdutosABC
    Call ConfGridCurvaABC(ColumnsIn_ProdutoCurvABC, ProdutoCurvABC_FormatString)
    
    Call ConfGridProdutosClie
    Call ConfGrdUltFat
    Call ConfGridProdutosPedidos
    Call ConfGridProdutosFaturado
    Call ConfGridClienteCondPgto
    
    
    objCADCLIENTE.PreenchComboBancos cboBANCOS
    objCADCLIENTE.PreenchComboSegmento cboSEGMENTO
    objCADCLIENTE.PreenchComboSegmento cboSEGMENTO2
    objCADCLIENTE.PreenchComboTransportadoras cboTRANSP
    
    cboBANCOS.ListIndex = -1
    cboSEGMENTO.ListIndex = -1
    cboSEGMENTO2.ListIndex = -1
    
    objCADCLIENTE.Modulo = Me.Name
    objCADCLIENTE.CLIECODIGO = iCodigo
    
    SSTab1.TabVisible(0) = False
    optTipoNOVASTEL(1).Value = True
    optFATNOVSTEEL(1).Value = True
    optFILIALPRODFAT(0).Value = True
    optSTATUS(0).Value = True
    optEMP(0).Value = True
    optPermFatSepSN(0).Value = True
    optNECCONFEST(0).Value = True
    optVISTELENT(0).Value = True
    optDesbClie(0).Value = True
    optPermFecOP(0).Value = True
   
    If objCADCLIENTE.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADCLIENTE.CLIECODIGO)
       txtCPFCNPJ.Text = objBLBFunc.FormataCnpj(objCADCLIENTE.CLICPFCNPJ)
       txtRAZAOSOC.Text = objCADCLIENTE.CLIERZAOSO
       txRGCGC.Text = objCADCLIENTE.CLIRGCGC
       txtNOMFANTA.Text = objCADCLIENTE.CLINOMFANT
       txtCodRef.Text = objCADCLIENTE.CODREF

        If Len(Trim(objCADCLIENTE.DTULTFATNOVA)) > 0 Then mskDTULTFATNOVA.Text = objCADCLIENTE.DTULTFATNOVA
        If Len(Trim(objCADCLIENTE.DTULTFATSTEEL)) > 0 Then mskDTULTFATSTEEL.Text = objCADCLIENTE.DTULTFATSTEEL
       
       If objCADCLIENTE.CLIPESSOA = "F" Then optFisica.Value = True
       If objCADCLIENTE.CLIPESSOA = "J" Then opfJuridica.Value = True
       
       txtENDNORM.Text = objCADCLIENTE.CLIENDEREC
       txtBAINOM.Text = objCADCLIENTE.CLIBAIRRO
       txtCIDNORM.Text = objCADCLIENTE.CLICIDADE
       
       For I = 0 To (cboESTNORM.ListCount - 1)
           If cboESTNORM.ItemData(I) = objCADCLIENTE.CLIESTADO Then cboESTNORM.ListIndex = I
       Next I
       
       If objCADCLIENTE.CLIESTADO > 0 Then
          txtZonaGeo.Text = BuscaAreaGeo(objCADCLIENTE.CLIESTADO)
       End If
       
       txtCEPNORM.Text = objCADCLIENTE.CLICEP
       
       mskDTCADASTRO.Text = Format(objCADCLIENTE.CADASTRO, "DD/MM/YYYY")
       
       arrTELEFONE = objCADCLIENTE.CLITELNORM
       arrCONTATO = objCADCLIENTE.CLICONTATO
       arrEMAIL = objCADCLIENTE.CLIEMAIL
       arrSITE = objCADCLIENTE.CLISITE
       
       If IsArray(arrTELEFONE) = True Then
          For I = 1 To UBound(arrTELEFONE)
              cboTELNORM.AddItem arrTELEFONE(I)
          Next I
       End If
       If IsArray(arrCONTATO) = True Then
          For I = 1 To UBound(arrCONTATO)
              cboCONTNORM.AddItem arrCONTATO(I)
          Next I
       End If
       If IsArray(arrEMAIL) = True Then
          For I = 1 To UBound(arrEMAIL)
              cboEMAILNORM.AddItem arrEMAIL(I)
          Next I
       End If
       If IsArray(arrSITE) = True Then
          For I = 1 To UBound(arrSITE)
              cboSITENORM.AddItem arrSITE(I)
          Next I
       End If
       
       txtENDCOBR.Text = objCADCLIENTE.ENDCOBRA
       txtBAICOBR.Text = objCADCLIENTE.BAICOBRA
       txtCIDCOBR.Text = objCADCLIENTE.CIDCOBRA
       For I = 0 To (cboESTCOBR.ListCount - 1)
           If cboESTCOBR.ItemData(I) = objCADCLIENTE.ESTCOBRA Then cboESTCOBR.ListIndex = I
       Next I
       txtCEPCOBR.Text = objCADCLIENTE.CEPCOBRA
       
       txtENDENTR.Text = objCADCLIENTE.ENDENTREGA
       txtBAIENTR.Text = objCADCLIENTE.BAIENTREGA
       txtCIDENTR.Text = objCADCLIENTE.CIDENTREGA
       For I = 0 To (cboESTENTR.ListCount - 1)
           If cboESTENTR.ItemData(I) = objCADCLIENTE.ESTENTREGA Then cboESTENTR.ListIndex = I
       Next I
       txtCEPENTR.Text = objCADCLIENTE.CEPENTREGA
       
       If objCADCLIENTE.BLOQPEDREST = "S" Then optBLPEDSIM.Value = True
       If objCADCLIENTE.BLOQPEDREST = "N" Then optBLPEDNAO.Value = True
       
       If objCADCLIENTE.AVISARRESTR = "S" Then optAVISASIM.Value = True
       If objCADCLIENTE.AVISARRESTR = "N" Then optAVISANAO.Value = True
       
       txtVLLIMCRED.Text = Format(objCADCLIENTE.VALLIMCRED, "#,##0.00")
       txtMeseReavali.Text = Str(objCADCLIENTE.MESEREAVAL)
       
       If objCADCLIENTE.SEMPRBLOQPE = "S" Then optSempSIM.Value = True
       If objCADCLIENTE.SEMPRBLOQPE = "N" Then optSempNAO.Value = True
       
       txtSALDOACIMA.Text = Format(objCADCLIENTE.AVISASALAC, "#,##0.00")
       txtBLOQPEDSALDO.Text = Format(objCADCLIENTE.BLQSALACIM, "#,##0.00")
       
       If objCADCLIENTE.BLOQAVISALD = "S" Then optBloqSim.Value = True
       If objCADCLIENTE.BLOQAVISALD = "N" Then optBloqNao.Value = True
       
       
       optNECCONFEST(objCADCLIENTE.NECCONFEST).Value = True
       optVISTELENT(objCADCLIENTE.VISTELEST).Value = True
       optDesbClie(objCADCLIENTE.DESBCLIE).Value = True
       optPermFecOP(objCADCLIENTE.PERMFECHOP).Value = True
       
       arrBANCOS = objCADCLIENTE.BANCOS
       arrREFBANCARIA = objCADCLIENTE.REFBANCARIA
       arrREFCOMERC = objCADCLIENTE.REFCOMERC
       arrREFPESSOAL = objCADCLIENTE.REFPESSOAL
       
       If IsArray(arrBANCOS) = True Then '' Bancos
          For I = 1 To UBound(arrBANCOS)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrBANCOS(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxBANCOS.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_AGENCIA & vbTab & BREC!SGI_CC
              BREC.Close
              
          Next I
       End If
       If IsArray(arrREFBANCARIA) = True Then '' Ref. Bancaria
          For I = 1 To UBound(arrREFBANCARIA)
              flxREFBANCARIA.AddItem "" & vbTab & arrREFBANCARIA(I, 1) & vbTab & arrREFBANCARIA(I, 2) & vbTab & arrREFBANCARIA(I, 3)
          Next I
       End If
       If IsArray(arrREFCOMERC) = True Then '' Ref. Comercial
          For I = 1 To UBound(arrREFCOMERC)
              flxREFEMPRESA.AddItem "" & vbTab & arrREFCOMERC(I, 1) & vbTab & arrREFCOMERC(I, 2) & vbTab & arrREFCOMERC(I, 3)
          Next I
       End If
       If IsArray(arrREFPESSOAL) = True Then '' Ref. Pessoal
          For I = 1 To UBound(arrREFPESSOAL)
              flxREFPESSOAL.AddItem "" & vbTab & arrREFPESSOAL(I, 1) & vbTab & arrREFPESSOAL(I, 2)
          Next I
       End If
       
       arrRESTRICOES = objCADCLIENTE.RESTRICOES
       
       If IsArray(arrRESTRICOES) = True Then '' Restrições
          For I = 1 To UBound(arrRESTRICOES)
              flxRESTRICOES.AddItem "" & vbTab & arrRESTRICOES(I)
          Next I
       End If
       
       arrFORNECEDOR = objCADCLIENTE.FORNECEDORE
       arrCLIENTES = objCADCLIENTE.CLIENTES
       
       If IsArray(arrFORNECEDOR) = True Then '' Restrições
          For I = 1 To UBound(arrFORNECEDOR)
              flxFORNEC.AddItem "" & vbTab & arrFORNECEDOR(I, 1) & vbTab & arrFORNECEDOR(I, 2)
          Next I
       End If
       If IsArray(arrCLIENTES) = True Then   '' Clientes
          For I = 1 To UBound(arrCLIENTES)
              flxEMPRESA.AddItem "" & vbTab & arrCLIENTES(I, 1) & vbTab & arrCLIENTES(I, 2)
          Next I
       End If
       
       If objCADCLIENTE.SEGMENTO > 0 Then txtCODSEQ.Text = Str(objCADCLIENTE.SEGMENTO)
       For I = 0 To (cboSEGMENTO.ListCount - 1)
           If objCADCLIENTE.SEGMENTO = cboSEGMENTO.ItemData(I) Then cboSEGMENTO.ListIndex = I
       Next I
       
       If objCADCLIENTE.SERTIFICADO = "S" Then optTEMCERTISIM.Value = True
       If objCADCLIENTE.SERTIFICADO = "N" Then optTEMCERTINAO.Value = True
       
       arrSISTCERTI = objCADCLIENTE.SISTCERTIF    '' Sistema de certificação
       arrEMPRESATEND = objCADCLIENTE.EMPREATENDE '' Empresas que atende
       
       If IsArray(arrSISTCERTI) = True Then   '' Sistema de certificação
          For I = 1 To UBound(arrSISTCERTI)
              flxSISTSERTFIC.AddItem "" & vbTab & arrSISTCERTI(I)
          Next I
       End If
       If IsArray(arrEMPRESATEND) = True Then '' Empresas atendidas
          For I = 1 To UBound(arrEMPRESATEND)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSEGCLI " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL
              sSql = sSql & "   And SGI_CODIGO = " & arrEMPRESATEND(I, 2)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
               
              If Not BREC.EOF Then flxATENDIDO.AddItem "" & vbTab & arrEMPRESATEND(I, 1) & vbTab & BREC!SGI_DESCRICAO & vbTab & arrEMPRESATEND(I, 2)
              If BREC.EOF Then flxATENDIDO.AddItem "" & vbTab & arrEMPRESATEND(I, 1) & vbTab & vbTab & arrEMPRESATEND(I, 2)
              
              BREC.Close
              
          Next I
       End If
       
       txtOBS.Text = objCADCLIENTE.OBS
       txtObsCom.Text = objCADCLIENTE.OBSCOM
       
       If objCADCLIENTE.CLIEJAECLI = "S" Then optECLIESIM.Value = True
       If objCADCLIENTE.CLIEJAECLI = "N" Then optECLIENAO.Value = True
       
       '' Transportadoras
       arrTRANSP = objCADCLIENTE.TRANSP
       If IsArray(arrTRANSP) = True Then
          For I = 1 To UBound(arrTRANSP)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrTRANSP(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxTRANSP.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
              BREC.Close
              
          Next I
       End If
       
       
        '' Vendedores
        arrVENDEDORES = objCADCLIENTE.VENDEDORES
        If IsArray(arrVENDEDORES) = True Then
          For I = 1 To UBound(arrVENDEDORES)
              flxVENDEDORES.AddItem arrVENDEDORES(I) & vbTab & _
                                    arrVENDEDORES(I) & vbTab & _
                                    PegaNomeVendedor(arrVENDEDORES(I))
          Next I
        End If
       
       
       Call PopGridDuplicatas
       
       arrPRODUTOSCLIE = objCADCLIENTE.PRODUTOS
       Call PopGrdProdClie
       Call PopGrdProdClieABC
       Call PopGrdUltFat
       
       With grdPRODUTOS
            If (.Rows - 1) > 0 Then
                .Row = 1
                .Col = 1
            End If
       End With
    
       optPermFatSepSN(objCADCLIENTE.PERMFATSSEPSN).Value = True
    
    
       Call PopGrdCondPgto
    
    End If
    
    stCLIENTE.Tab = 0

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    '' Dados Cadastrais **
    Frame2.Enabled = True
    txtCPFCNPJ.Enabled = True
    txtRAZAOSOC.Enabled = True
    txRGCGC.Enabled = True
    txtNOMFANTA.Enabled = True
    optFisica.Enabled = True
    opfJuridica.Enabled = True
    optECLIESIM.Enabled = False
    optECLIENAO.Enabled = False
    txtENDNORM.Enabled = True
    txtBAINOM.Enabled = True
    txtCIDNORM.Enabled = True
    cboESTNORM.Enabled = True
    txtCEPNORM.Enabled = True
    
    cboTELNORM.Locked = False
    cboCONTNORM.Locked = False
    cboEMAILNORM.Locked = False
    cboEMAILNORM.Locked = False
    
    mskDTCADASTRO.Enabled = False
    
    txtOBS.Locked = False
    txtObsCom.Locked = False
    
    '' -----------------------
    Frame33.Enabled = False
    
    '' Cobrança /  Entrega
    Frame3.Enabled = True
    Frame4.Enabled = True
    '' -----------------------
    
    '' Financeiro
    Frame6.Enabled = True
    Frame10.Enabled = True
    Frame13.Enabled = True
    Frame15.Enabled = True
    Frame17.Enabled = True
    Frame18.Enabled = True
    Frame19.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True
    '' -----------------------
    
    '' Crédito **
    Frame5.Enabled = True
    Frame24.Enabled = True
    Frame7.Enabled = False
    '' -----------------------
    
    '' Folhow UP
    Frame26.Enabled = True
    Frame27.Enabled = True
    Frame28.Enabled = True
    Frame29.Enabled = True
    Frame30.Enabled = True
    '' -----------------------
    
    '' Transportadora
    Frame40.Enabled = True
    Frame41.Enabled = True
    
    
    Me.Caption = "Cadastro de clientes - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCLIENTE
    
    objBLBFunc.Preenche_Estado cboESTNORM
    objBLBFunc.Preenche_Estado cboESTCOBR
    objBLBFunc.Preenche_Estado cboESTENTR
    
    '' --------------------------
    
    ConfGridBancos
    ConfGridRefBancaria
    CondGridREfCOMERCIAL
    ConfGriRefPessoal
    ConfGridRestricoes
    ConfGridFornec
    ConfGridClientes
    ConfGridHistAvaliacao
    ConfGridCertificado
    ConfGridEmpresaAtende
    ConfGridDuplicatas
    ConfGridTransp
    ConfGridVendedores
    ConfGrdPedido
    ConfGrdFat
    ConfGridProdutos
    Call ConfGridProdutosClie
    Call ConfGridProdutosABC
    Call ConfGridCurvaABC(ColumnsIn_ProdutoCurvABC, ProdutoCurvABC_FormatString)
    
    Call ConfGridProdutosPedidos
    Call ConfGridProdutosFaturado
    Call ConfGridClienteCondPgto
    
    objCADCLIENTE.PreenchComboBancos cboBANCOS
    objCADCLIENTE.PreenchComboSegmento cboSEGMENTO
    objCADCLIENTE.PreenchComboSegmento cboSEGMENTO2
    objCADCLIENTE.PreenchComboTransportadoras cboTRANSP
    
    cboBANCOS.ListIndex = -1
    cboSEGMENTO.ListIndex = -1
    cboSEGMENTO2.ListIndex = -1
    
    objCADCLIENTE.Modulo = Me.Name
    objCADCLIENTE.CLIECODIGO = iCodigo
    
    SSTab1.TabVisible(0) = False
    optTipoNOVASTEL(1).Value = True
    optFATNOVSTEEL(1).Value = True
    optSTATUS(0).Value = True
    optEMP(0).Value = True
    optPermFatSepSN(0).Value = True
    optNECCONFEST(0).Value = True
    optVISTELENT(0).Value = True
    optDesbClie(0).Value = True
    optPermFecOP(0).Value = True
       
    If objCADCLIENTE.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADCLIENTE.CLIECODIGO)
       txtCPFCNPJ.Text = objBLBFunc.FormataCnpj(objCADCLIENTE.CLICPFCNPJ)
       txtRAZAOSOC.Text = objCADCLIENTE.CLIERZAOSO
       txRGCGC.Text = objCADCLIENTE.CLIRGCGC
       txtNOMFANTA.Text = objCADCLIENTE.CLINOMFANT
       txtCodRef.Text = objCADCLIENTE.CODREF
       
       If objCADCLIENTE.CLIPESSOA = "F" Then optFisica.Value = True
       If objCADCLIENTE.CLIPESSOA = "J" Then opfJuridica.Value = True
       
       txtENDNORM.Text = objCADCLIENTE.CLIENDEREC
       txtBAINOM.Text = objCADCLIENTE.CLIBAIRRO
       txtCIDNORM.Text = objCADCLIENTE.CLICIDADE
       
       For I = 0 To (cboESTNORM.ListCount - 1)
           If cboESTNORM.ItemData(I) = objCADCLIENTE.CLIESTADO Then cboESTNORM.ListIndex = I
       Next I
       
       If objCADCLIENTE.CLIESTADO > 0 Then
          txtZonaGeo.Text = BuscaAreaGeo(objCADCLIENTE.CLIESTADO)
       End If
       
       txtCEPNORM.Text = objCADCLIENTE.CLICEP
       
       mskDTCADASTRO.Text = Format(objCADCLIENTE.CADASTRO, "DD/MM/YYYY")
       
       arrTELEFONE = objCADCLIENTE.CLITELNORM
       arrCONTATO = objCADCLIENTE.CLICONTATO
       arrEMAIL = objCADCLIENTE.CLIEMAIL
       arrSITE = objCADCLIENTE.CLISITE
       
       If IsArray(arrTELEFONE) = True Then
          For I = 1 To UBound(arrTELEFONE)
              cboTELNORM.AddItem arrTELEFONE(I)
          Next I
       End If
       If IsArray(arrCONTATO) = True Then
          For I = 1 To UBound(arrCONTATO)
              cboCONTNORM.AddItem arrCONTATO(I)
          Next I
       End If
       If IsArray(arrEMAIL) = True Then
          For I = 1 To UBound(arrEMAIL)
              cboEMAILNORM.AddItem arrEMAIL(I)
          Next I
       End If
       If IsArray(arrSITE) = True Then
          For I = 1 To UBound(arrSITE)
              cboSITENORM.AddItem arrSITE(I)
          Next I
       End If
       
       txtENDCOBR.Text = objCADCLIENTE.ENDCOBRA
       txtBAICOBR.Text = objCADCLIENTE.BAICOBRA
       txtCIDCOBR.Text = objCADCLIENTE.CIDCOBRA
       For I = 0 To (cboESTCOBR.ListCount - 1)
           If cboESTCOBR.ItemData(I) = objCADCLIENTE.ESTCOBRA Then cboESTCOBR.ListIndex = I
       Next I
       txtCEPCOBR.Text = objCADCLIENTE.CEPCOBRA
       
       txtENDENTR.Text = objCADCLIENTE.ENDENTREGA
       txtBAIENTR.Text = objCADCLIENTE.BAIENTREGA
       txtCIDENTR.Text = objCADCLIENTE.CIDENTREGA
       For I = 0 To (cboESTENTR.ListCount - 1)
           If cboESTENTR.ItemData(I) = objCADCLIENTE.ESTENTREGA Then cboESTENTR.ListIndex = I
       Next I
       txtCEPENTR.Text = objCADCLIENTE.CEPENTREGA
       
       If objCADCLIENTE.BLOQPEDREST = "S" Then optBLPEDSIM.Value = True
       If objCADCLIENTE.BLOQPEDREST = "N" Then optBLPEDNAO.Value = True
       
       If objCADCLIENTE.AVISARRESTR = "S" Then optAVISASIM.Value = True
       If objCADCLIENTE.AVISARRESTR = "N" Then optAVISANAO.Value = True
       
       txtVLLIMCRED.Text = Format(objCADCLIENTE.VALLIMCRED, "#,##0.00")
       txtMeseReavali.Text = Str(objCADCLIENTE.MESEREAVAL)
       
       If objCADCLIENTE.SEMPRBLOQPE = "S" Then optSempSIM.Value = True
       If objCADCLIENTE.SEMPRBLOQPE = "N" Then optSempNAO.Value = True
       
       txtSALDOACIMA.Text = Format(objCADCLIENTE.AVISASALAC, "#,##0.00")
       txtBLOQPEDSALDO.Text = Format(objCADCLIENTE.BLQSALACIM, "#,##0.00")
       
       If objCADCLIENTE.BLOQAVISALD = "S" Then optBloqSim.Value = True
       If objCADCLIENTE.BLOQAVISALD = "N" Then optBloqNao.Value = True
       
       optNECCONFEST(objCADCLIENTE.NECCONFEST).Value = True
       optVISTELENT(objCADCLIENTE.VISTELEST).Value = True
       optDesbClie(objCADCLIENTE.DESBCLIE).Value = True
       optPermFecOP(objCADCLIENTE.PERMFECHOP).Value = True
       
      
       arrBANCOS = objCADCLIENTE.BANCOS
       arrREFBANCARIA = objCADCLIENTE.REFBANCARIA
       arrREFCOMERC = objCADCLIENTE.REFCOMERC
       arrREFPESSOAL = objCADCLIENTE.REFPESSOAL
       
       If Len(Trim(objCADCLIENTE.DTULTFATNOVA)) > 0 Then mskDTULTFATNOVA.Text = objCADCLIENTE.DTULTFATNOVA
       If Len(Trim(objCADCLIENTE.DTULTFATSTEEL)) > 0 Then mskDTULTFATSTEEL.Text = objCADCLIENTE.DTULTFATSTEEL
       
       If IsArray(arrBANCOS) = True Then '' Bancos
          For I = 1 To UBound(arrBANCOS)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrBANCOS(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxBANCOS.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_AGENCIA & vbTab & BREC!SGI_CC
              BREC.Close
              
          Next I
       End If
       If IsArray(arrREFBANCARIA) = True Then '' Ref. Bancaria
          For I = 1 To UBound(arrREFBANCARIA)
              flxREFBANCARIA.AddItem "" & vbTab & arrREFBANCARIA(I, 1) & vbTab & arrREFBANCARIA(I, 2) & vbTab & arrREFBANCARIA(I, 3)
          Next I
       End If
       If IsArray(arrREFCOMERC) = True Then '' Ref. Comercial
          For I = 1 To UBound(arrREFCOMERC)
              flxREFEMPRESA.AddItem "" & vbTab & arrREFCOMERC(I, 1) & vbTab & arrREFCOMERC(I, 2) & vbTab & arrREFCOMERC(I, 3)
          Next I
       End If
       If IsArray(arrREFPESSOAL) = True Then '' Ref. Pessoal
          For I = 1 To UBound(arrREFPESSOAL)
              flxREFPESSOAL.AddItem "" & vbTab & arrREFPESSOAL(I, 1) & vbTab & arrREFPESSOAL(I, 2)
          Next I
       End If
       
       arrRESTRICOES = objCADCLIENTE.RESTRICOES
       
       If IsArray(arrRESTRICOES) = True Then '' Restrições
          For I = 1 To UBound(arrRESTRICOES)
              flxRESTRICOES.AddItem "" & vbTab & arrRESTRICOES(I)
          Next I
       End If
       
       arrFORNECEDOR = objCADCLIENTE.FORNECEDORE
       arrCLIENTES = objCADCLIENTE.CLIENTES
       
       If IsArray(arrFORNECEDOR) = True Then '' Restrições
          For I = 1 To UBound(arrFORNECEDOR)
              flxFORNEC.AddItem "" & vbTab & arrFORNECEDOR(I, 1) & vbTab & arrFORNECEDOR(I, 2)
          Next I
       End If
       If IsArray(arrCLIENTES) = True Then   '' Clientes
          For I = 1 To UBound(arrCLIENTES)
              flxEMPRESA.AddItem "" & vbTab & arrCLIENTES(I, 1) & vbTab & arrCLIENTES(I, 2)
          Next I
       End If
       
       If objCADCLIENTE.SEGMENTO > 0 Then txtCODSEQ.Text = Str(objCADCLIENTE.SEGMENTO)
       For I = 0 To (cboSEGMENTO.ListCount - 1)
           If objCADCLIENTE.SEGMENTO = cboSEGMENTO.ItemData(I) Then cboSEGMENTO.ListIndex = I
       Next I
       
       If objCADCLIENTE.SERTIFICADO = "S" Then optTEMCERTISIM.Value = True
       If objCADCLIENTE.SERTIFICADO = "N" Then optTEMCERTINAO.Value = True
       
       arrSISTCERTI = objCADCLIENTE.SISTCERTIF    '' Sistema de certificação
       arrEMPRESATEND = objCADCLIENTE.EMPREATENDE '' Empresas que atende
       
       If IsArray(arrSISTCERTI) = True Then   '' Sistema de certificação
          For I = 1 To UBound(arrSISTCERTI)
              flxSISTSERTFIC.AddItem "" & vbTab & arrSISTCERTI(I)
          Next I
       End If
       If IsArray(arrEMPRESATEND) = True Then '' Empresas atendidas
          For I = 1 To UBound(arrEMPRESATEND)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSEGCLI " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL
              sSql = sSql & "   And SGI_CODIGO = " & arrEMPRESATEND(I, 2)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
               
              If Not BREC.EOF Then flxATENDIDO.AddItem "" & vbTab & arrEMPRESATEND(I, 1) & vbTab & BREC!SGI_DESCRICAO & vbTab & arrEMPRESATEND(I, 2)
              If BREC.EOF Then flxATENDIDO.AddItem "" & vbTab & arrEMPRESATEND(I, 1) & vbTab & vbTab & arrEMPRESATEND(I, 2)
              
              BREC.Close
              
          Next I
       End If
       
       txtOBS.Text = objCADCLIENTE.OBS
       txtObsCom.Text = objCADCLIENTE.OBSCOM
       
       If objCADCLIENTE.CLIEJAECLI = "S" Then optECLIESIM.Value = True
       If objCADCLIENTE.CLIEJAECLI = "N" Then optECLIENAO.Value = True
       
       '' Transportadoras
       arrTRANSP = objCADCLIENTE.TRANSP
       If IsArray(arrTRANSP) = True Then
          For I = 1 To UBound(arrTRANSP)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrTRANSP(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxTRANSP.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
              BREC.Close
              
          Next I
       End If
       
        '' Vendedores
        arrVENDEDORES = objCADCLIENTE.VENDEDORES
        If IsArray(arrVENDEDORES) = True Then
          For I = 1 To UBound(arrVENDEDORES)
              flxVENDEDORES.AddItem arrVENDEDORES(I) & vbTab & _
                                    arrVENDEDORES(I) & vbTab & _
                                    PegaNomeVendedor(arrVENDEDORES(I))
          Next I
        End If
       
       PopGridDuplicatas
       
       arrPRODUTOSCLIE = objCADCLIENTE.PRODUTOS
        Call PopGrdProdClie
       
    
        optPermFatSepSN(objCADCLIENTE.PERMFATSSEPSN).Value = True
    
        Call PopGrdCondPgto
    
    End If

End Sub

Private Sub InclBancos()

    Dim I As Integer
    
    If Len(Trim(txtCODBANCO.Text)) = 0 Or cboBANCOS.ListIndex = -1 Then
       MsgBox "Informe o banco !!!", vbOKOnly + vbCritical, "aviso"
       txtCODBANCO.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxBANCOS.Rows - 1)
        If txtCODBANCO.Text = flxBANCOS.TextMatrix(I, 1) Then
           MsgBox "Este banco já foi incluso !!!", vbOKOnly + vbCritical, "aviso"
           txtCODBANCO.Text = ""
           cboBANCOS.ListIndex = -1
           txtCODBANCO.SetFocus
           Exit Sub
        End If
    Next I
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & txtCODBANCO.Text
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then flxBANCOS.AddItem "" & vbTab & txtCODBANCO.Text & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_AGENCIA & vbTab & BREC!SGI_CC
    BREC.Close
    
    cboBANCOS.ListIndex = -1
    txtCODBANCO.Text = ""
    txtCODBANCO.SetFocus
    
End Sub

Private Sub ConfGridDuplicatas()

    '' Duplicatas a vencar
    flxDuplApgt.Rows = 1
    flxDuplApgt.Cols = 6
    
    flxDuplApgt.TextMatrix(0, 0) = ""
    flxDuplApgt.TextMatrix(0, 1) = "Doc. Nº"
    flxDuplApgt.TextMatrix(0, 2) = "Dt. Venc."
    flxDuplApgt.TextMatrix(0, 3) = "Valor"
    flxDuplApgt.TextMatrix(0, 4) = "Parc."
    flxDuplApgt.TextMatrix(0, 5) = "Dias"
    
    flxDuplApgt.ColWidth(0) = 0
    flxDuplApgt.ColWidth(1) = 1000
    flxDuplApgt.ColWidth(2) = 1000
    flxDuplApgt.ColWidth(3) = 1000
    flxDuplApgt.ColWidth(4) = 600
    flxDuplApgt.ColWidth(5) = 500
    
    '' Duplicatas Vencidas
    flxDuplVencidas.Rows = 1
    flxDuplVencidas.Cols = 6
    
    flxDuplVencidas.TextMatrix(0, 0) = ""
    flxDuplVencidas.TextMatrix(0, 1) = "Doc. Nº"
    flxDuplVencidas.TextMatrix(0, 2) = "Dt. Venc."
    flxDuplVencidas.TextMatrix(0, 3) = "Valor"
    flxDuplVencidas.TextMatrix(0, 4) = "Parc."
    flxDuplVencidas.TextMatrix(0, 5) = "Dias"
    
    flxDuplVencidas.ColWidth(0) = 0
    flxDuplVencidas.ColWidth(1) = 1000
    flxDuplVencidas.ColWidth(2) = 1000
    flxDuplVencidas.ColWidth(3) = 1000
    flxDuplVencidas.ColWidth(4) = 600
    flxDuplVencidas.ColWidth(5) = 500
    
    '' Pagamento Antecipado
    flxPgtAntecipado.Rows = 1
    flxPgtAntecipado.Cols = 8
    
    flxPgtAntecipado.TextMatrix(0, 0) = ""
    flxPgtAntecipado.TextMatrix(0, 1) = "Doc. Nº"
    flxPgtAntecipado.TextMatrix(0, 2) = "Dt. Venc."
    flxPgtAntecipado.TextMatrix(0, 3) = "Dt. Pgto."
    flxPgtAntecipado.TextMatrix(0, 4) = "Valor"
    flxPgtAntecipado.TextMatrix(0, 5) = "Valor Pgto."
    flxPgtAntecipado.TextMatrix(0, 6) = "Parc."
    flxPgtAntecipado.TextMatrix(0, 7) = "Dias"

    flxPgtAntecipado.ColWidth(0) = 0
    flxPgtAntecipado.ColWidth(1) = 1000
    flxPgtAntecipado.ColWidth(2) = 1000
    flxPgtAntecipado.ColWidth(3) = 1000
    flxPgtAntecipado.ColWidth(4) = 1000
    flxPgtAntecipado.ColWidth(5) = 1000
    flxPgtAntecipado.ColWidth(6) = 600
    flxPgtAntecipado.ColWidth(7) = 500
    
    '' Pagamento no Prazo
    flxTitPgtoPrazo.Rows = 1
    flxTitPgtoPrazo.Cols = 8
    
    flxTitPgtoPrazo.TextMatrix(0, 0) = ""
    flxTitPgtoPrazo.TextMatrix(0, 1) = "Doc. Nº"
    flxTitPgtoPrazo.TextMatrix(0, 2) = "Dt. Venc."
    flxTitPgtoPrazo.TextMatrix(0, 3) = "Dt. Pgto."
    flxTitPgtoPrazo.TextMatrix(0, 4) = "Valor"
    flxTitPgtoPrazo.TextMatrix(0, 5) = "Valor Pgto."
    flxTitPgtoPrazo.TextMatrix(0, 6) = "Parc."
    flxTitPgtoPrazo.TextMatrix(0, 7) = "Dias"

    flxTitPgtoPrazo.ColWidth(0) = 0
    flxTitPgtoPrazo.ColWidth(1) = 1000
    flxTitPgtoPrazo.ColWidth(2) = 1000
    flxTitPgtoPrazo.ColWidth(3) = 1000
    flxTitPgtoPrazo.ColWidth(4) = 1000
    flxTitPgtoPrazo.ColWidth(5) = 1000
    flxTitPgtoPrazo.ColWidth(6) = 600
    flxTitPgtoPrazo.ColWidth(7) = 500
    
    '' Pagamento com Atrazo
    flxDpPgtAtrazo.Rows = 1
    flxDpPgtAtrazo.Cols = 8
    
    flxDpPgtAtrazo.TextMatrix(0, 0) = ""
    flxDpPgtAtrazo.TextMatrix(0, 1) = "Doc. Nº"
    flxDpPgtAtrazo.TextMatrix(0, 2) = "Dt. Venc."
    flxDpPgtAtrazo.TextMatrix(0, 3) = "Dt. Pgto."
    flxDpPgtAtrazo.TextMatrix(0, 4) = "Valor"
    flxDpPgtAtrazo.TextMatrix(0, 5) = "Valor Pgto."
    flxDpPgtAtrazo.TextMatrix(0, 6) = "Parc."
    flxDpPgtAtrazo.TextMatrix(0, 7) = "Dias"

    flxDpPgtAtrazo.ColWidth(0) = 0
    flxDpPgtAtrazo.ColWidth(1) = 1000
    flxDpPgtAtrazo.ColWidth(2) = 1000
    flxDpPgtAtrazo.ColWidth(3) = 1000
    flxDpPgtAtrazo.ColWidth(4) = 1000
    flxDpPgtAtrazo.ColWidth(5) = 1000
    flxDpPgtAtrazo.ColWidth(6) = 600
    flxDpPgtAtrazo.ColWidth(7) = 500
    
    

End Sub

Private Sub PopGridDuplicatas()

    stDuplicatas.Tab = 0
    StHistDupl.Tab = 0
    StHistDupli.Tab = 0
    
    Dim curTOTTITABERTO     As Currency
    Dim curTOTTITABERTOVENC As Currency
    Dim curTOTTITPAGOS      As Currency
    Dim curTOTTITPGANTECIP  As Currency
    Dim curTOTTITPGNOPRAZO  As Currency
    Dim curTOTTITPGATRAZADO As Currency
    
    '' Titulos em Aberto e Não Vencidos
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_NUMDOC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_VLPAGO  IS NULL" & vbCrLf
    sSql = sSql & "  And ITENS.SGI_DATAVENC > '" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODCLI   = " & objCADCLIENTE.CLIECODIGO & vbCrLf
    sSql = sSql & "  And CABEC.SGI_FILIAL   = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO   = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "       ,ITENS.SGI_PARCELA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    curTOTTITABERTO = 0
    Do While Not BREC.EOF
        
       flxDuplApgt.AddItem "" & vbTab & _
                           BREC!SGI_NUMDOC & vbTab & _
                           Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                           Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                           Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                           Format((BREC!SGI_DATAVENC - Now), "###000")
                           
       curTOTTITABERTO = curTOTTITABERTO + BREC!SGI_VLDOC
       
       BREC.MoveNext
       If BREC.EOF Then
          flxDuplApgt.AddItem "" & vbTab & _
                              "Total" & vbTab & _
                              "" & vbTab & _
                              Format(curTOTTITABERTO, "#,##0.00")
       End If
       
    Loop
    
    BREC.Close
    
    '' Titulos em Aberto e Vencidos
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_NUMDOC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_VLPAGO  IS NULL" & vbCrLf
    sSql = sSql & "  And ITENS.SGI_DATAVENC <= '" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODCLI   = " & objCADCLIENTE.CLIECODIGO & vbCrLf
    sSql = sSql & "  And CABEC.SGI_FILIAL   = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO   = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "       ,ITENS.SGI_PARCELA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    curTOTTITABERTOVENC = 0
    Do While Not BREC.EOF
        
       flxDuplVencidas.AddItem "" & vbTab & _
                               BREC!SGI_NUMDOC & vbTab & _
                               Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                               Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                               Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                               Format((Now - BREC!SGI_DATAVENC), "###000")
                           
       curTOTTITABERTOVENC = curTOTTITABERTOVENC + BREC!SGI_VLDOC
       
       BREC.MoveNext
       If BREC.EOF Then
          flxDuplVencidas.AddItem "" & vbTab & _
                                   "Total" & vbTab & _
                                   "" & vbTab & _
                                   Format(curTOTTITABERTOVENC, "#,##0.00")
       End If
    Loop
    
    BREC.Close
    
    
    '' Baixadas Pagas
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_NUMDOC  " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DTPGTO " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLPAGO  " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODCLI   = " & objCADCLIENTE.CLIECODIGO & vbCrLf
    sSql = sSql & "  And CABEC.SGI_FILIAL   = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO   = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "       ,ITENS.SGI_PARCELA"
    
    
    curTOTTITPAGOS = 0
    curTOTTITPGANTECIP = 0
    curTOTTITPGNOPRAZO = 0
    curTOTTITPGATRAZADO = 0
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
    
       '' Antecipado
       If BREC!SGI_DTPGTO < BREC!SGI_DATAVENC Then
          flxPgtAntecipado.AddItem "" & vbTab & _
                                   BREC!SGI_NUMDOC & vbTab & _
                                   Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                   Format(BREC!SGI_VLPAGO, "#,##0.00") & vbTab & _
                                   Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                   Format((BREC!SGI_DATAVENC - BREC!SGI_DTPGTO), "###000")
          curTOTTITPGANTECIP = curTOTTITPGANTECIP + BREC!SGI_VLDOC
       
       End If
       
       '' No Prazo
       If BREC!SGI_DTPGTO = BREC!SGI_DATAVENC Then
          flxTitPgtoPrazo.AddItem "" & vbTab & _
                                   BREC!SGI_NUMDOC & vbTab & _
                                   Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                   Format(BREC!SGI_VLPAGO, "#,##0.00") & vbTab & _
                                   Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00")
          curTOTTITPGNOPRAZO = curTOTTITPGNOPRAZO + BREC!SGI_VLDOC
       End If
       
       '' Atrazado
       If BREC!SGI_DTPGTO > BREC!SGI_DATAVENC Then
          flxDpPgtAtrazo.AddItem "" & vbTab & _
                                   BREC!SGI_NUMDOC & vbTab & _
                                   Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                   Format(BREC!SGI_VLPAGO, "#,##0.00") & vbTab & _
                                   Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                   Format((BREC!SGI_DTPGTO - BREC!SGI_DATAVENC), "###000")
          curTOTTITPGATRAZADO = curTOTTITPGATRAZADO + BREC!SGI_VLDOC
       End If
    
       curTOTTITPAGOS = curTOTTITPAGOS + BREC!SGI_VLDOC
       BREC.MoveNext
    Loop
    BREC.Close
    
    '' Labels de Resumo
    '' Totais de Titulos em  Aberto
    lblValoresDupl(1).Caption = Format((curTOTTITABERTO + curTOTTITABERTOVENC), "#,##0.00")
    lblValoresDupl(4).Caption = Format(curTOTTITPAGOS, "#,##0.00")
    lblValoresDupl(5).Caption = Format((curTOTTITABERTO + curTOTTITABERTOVENC + curTOTTITPAGOS), "#,##0.00")
    
    '' Titulos em Aberto e não vencido
    lblValoresDupl(0).Caption = Format(curTOTTITABERTO, "#,##0.00")
    '' Titulos em Aberto e vencido
    lblValoresDupl(2).Caption = Format(curTOTTITABERTOVENC, "#,##0.00")
    
    '' Pago Antecipado
    lblValoresDupl(3).Caption = Format(curTOTTITPGANTECIP, "#,##0.00")
    '' Pagas no Prazo
    lblValoresDupl(7).Caption = Format(curTOTTITPGNOPRAZO, "#,##0.00")
    '' Pago Com Atrazo
    lblValoresDupl(6).Caption = Format(curTOTTITPGATRAZADO, "#,##0.00")
    
    '' % Desempenhos Totais
    If curTOTTITABERTO > 0 Or _
       curTOTTITABERTOVENC > 0 Or _
       curTOTTITPAGOS Then
       lblValoresDupl(8).Caption = Format(((curTOTTITABERTO + curTOTTITABERTOVENC) / (curTOTTITABERTO + curTOTTITABERTOVENC + curTOTTITPAGOS)) * 100, "#,##0.00") & "%"
    End If
       
    If curTOTTITABERTO > 0 Or _
       curTOTTITABERTOVENC > 0 Or _
       curTOTTITPAGOS > 0 Then
       lblValoresDupl(9).Caption = Format((curTOTTITPAGOS / (curTOTTITABERTO + curTOTTITABERTOVENC + curTOTTITPAGOS)) * 100, "#,##0.00") & "%"
    End If
    
    '' Sub Totais
    If curTOTTITABERTO > 0 Or curTOTTITABERTOVENC > 0 Or curTOTTITPAGOS > 0 Then lblValoresDupl(10).Caption = Format((curTOTTITABERTO / (curTOTTITABERTO + curTOTTITABERTOVENC + curTOTTITPAGOS)) * 100, "#,##0.00") & "%"
    If curTOTTITABERTOVENC > 0 Or curTOTTITABERTO > 0 Or curTOTTITABERTOVENC > 0 Or curTOTTITPAGOS > 0 Then lblValoresDupl(11).Caption = Format((curTOTTITABERTOVENC / (curTOTTITABERTO + curTOTTITABERTOVENC + curTOTTITPAGOS)) * 100, "#,##0.00") & "%"
    
    
    If curTOTTITPGANTECIP > 0 Or curTOTTITABERTO > 0 Or curTOTTITABERTOVENC > 0 Or curTOTTITPAGOS > 0 Then lblValoresDupl(12).Caption = Format((curTOTTITPGANTECIP / (curTOTTITABERTO + curTOTTITABERTOVENC + curTOTTITPAGOS)) * 100, "#,##0.00") & "%"
    If curTOTTITPGNOPRAZO > 0 Or curTOTTITABERTO > 0 Or curTOTTITABERTOVENC > 0 Or curTOTTITPAGOS > 0 Then lblValoresDupl(13).Caption = Format((curTOTTITPGNOPRAZO / (curTOTTITABERTO + curTOTTITABERTOVENC + curTOTTITPAGOS)) * 100, "#,##0.00") & "%"
    If curTOTTITPGATRAZADO > 0 Or curTOTTITABERTO > 0 Or curTOTTITABERTOVENC > 0 Or curTOTTITPAGOS > 0 Then lblValoresDupl(14).Caption = Format((curTOTTITPGATRAZADO / (curTOTTITABERTO + curTOTTITABERTOVENC + curTOTTITPAGOS)) * 100, "#,##0.00") & "%"
    
    
    SomDuplPagas
    
End Sub

Private Sub SomDuplPagas()
    
    Dim I        As Integer
    Dim curTotal As Currency
    
    '' ---------------------------
    '' Grid Antecipado
    curTotal = 0
    For I = 1 To (flxPgtAntecipado.Rows - 1)
        curTotal = curTotal + CCur(flxPgtAntecipado.TextMatrix(I, 4))
    Next I
    If (flxPgtAntecipado.Rows - 1) > 0 Then
       flxPgtAntecipado.AddItem "" & vbTab & _
                                "Total" & vbTab & _
                                "" & vbTab & _
                                "" & vbTab & _
                                Format(curTotal, "#,##0.00")
    End If
    
    '' ---------------------------
    '' Grid no Prazo
    curTotal = 0
    For I = 1 To (flxTitPgtoPrazo.Rows - 1)
        curTotal = curTotal + CCur(flxTitPgtoPrazo.TextMatrix(I, 4))
    Next I
    If (flxTitPgtoPrazo.Rows - 1) > 0 Then
        flxTitPgtoPrazo.AddItem "" & vbTab & _
                                "Total" & vbTab & _
                                "" & vbTab & _
                                "" & vbTab & _
                                Format(curTotal, "#,##0.00")
    End If
    
    '' ---------------------------
    '' Grid Atrazado
    curTotal = 0
    For I = 1 To (flxDpPgtAtrazo.Rows - 1)
        curTotal = curTotal + CCur(flxDpPgtAtrazo.TextMatrix(I, 4))
    Next I
    If (flxDpPgtAtrazo.Rows - 1) > 0 Then
        flxDpPgtAtrazo.AddItem "" & vbTab & _
                                "Total" & vbTab & _
                                "" & vbTab & _
                                "" & vbTab & _
                                Format(curTotal, "#,##0.00")
    End If
    
End Sub

Private Sub ConfGridTransp()

    flxTRANSP.Rows = 1
    flxTRANSP.Cols = 3
    
    flxTRANSP.TextMatrix(0, 0) = ""
    flxTRANSP.TextMatrix(0, 1) = "Código"
    flxTRANSP.TextMatrix(0, 2) = "Descrição"
    
    flxTRANSP.ColWidth(0) = 0
    flxTRANSP.ColWidth(1) = 1000
    flxTRANSP.ColWidth(2) = 5000

End Sub

Private Sub IncGridTransp()

   Dim I As Integer
   
   If (Len(Trim(txtCODTRANSP.Text)) = 0) Or (cboTRANSP.ListIndex = -1) Then
      MsgBox "Informe o código da transportadora !!!", vbOKOnly + vbExclamation, "aviso"
      txtCODTRANSP.SetFocus
      Exit Sub
   End If
      
   For I = 1 To (flxTRANSP.Rows - 1)
       If flxTRANSP.TextMatrix(I, 0) = txtCODTRANSP.Text Then
          MsgBox "Esta transportadora já esta relacionada !!!", vbOKOnly + vbExclamation, "aviso"
          txtCODTRANSP.Text = ""
          cboTRANSP.ListIndex = -1
          txtCODTRANSP.SetFocus
          Exit Sub
       End If
   Next I
   
   flxTRANSP.AddItem "" & vbTab & _
                     txtCODTRANSP.Text & vbTab & _
                     Trim(cboTRANSP.Text)
   
   txtCODTRANSP.Text = ""
   cboTRANSP.ListIndex = -1
   
   txtCODTRANSP.SetFocus

End Sub

Private Function PegaNomeTecnico(varCODIGOTECNICO As Variant) As String

    PegaNomeTecnico = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTECNICO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & varCODIGOTECNICO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaNomeTecnico = BREC!SGI_NOME
    BREC.Close

End Function

Private Sub ConfGridVendedores()

    flxVENDEDORES.Rows = 1
    flxVENDEDORES.Cols = 3
    
    flxVENDEDORES.TextMatrix(0, 0) = ""
    flxVENDEDORES.TextMatrix(0, 1) = "Código"
    flxVENDEDORES.TextMatrix(0, 2) = "Descrição"

    flxVENDEDORES.ColWidth(0) = 0
    flxVENDEDORES.ColWidth(1) = 1000
    flxVENDEDORES.ColWidth(2) = 5000

End Sub

Private Function PegaNomeVendedor(varCODIGOVENDEDOR As Variant) As String

    PegaNomeVendedor = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & varCODIGOVENDEDOR
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaNomeVendedor = BREC!SGI_DESCRICAO
    BREC.Close

End Function

Private Function BuscaAreaGeo(lngCODESTADO As Long) As String
    
    BuscaAreaGeo = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ZON.SGI_DESCRI " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CLIZONAGEO CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADZONAGEO ZON " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLI.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODESTADO = " & lngCODESTADO & vbCrLf
    sSql = sSql & "   And ZON.SGI_FILIAL    = CLI.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ZON.SGI_CODIGO    = CLI.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then BuscaAreaGeo = Trim(BREC!SGI_DESCRI)
    BREC.Close
    
End Function

Private Function PegaQtdeCotacao(lngCODIGO As Long) As Currency

    PegaQtdeCotacao = 0

    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(SGI_QTDE) as SGI_QTDE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDI " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF Then PegaQtdeCotacao = BREC3!SGI_QTDE
    BREC3.Close
    
End Function

Private Function PegaQtdePedido(lngCODIGO As Long) As Currency

    PegaQtdePedido = 0

    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(SGI_QTDPED) as SGI_QTDPED " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDI " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF Then PegaQtdePedido = BREC4!SGI_QTDPED
    BREC4.Close
    
End Function


Private Sub ConfGrdPedido()
    
        flxPedidos.Rows = 1
        flxPedidos.Cols = 9
        
        flxPedidos.TextMatrix(0, 0) = ""
        flxPedidos.TextMatrix(0, 1) = "Código"
        flxPedidos.TextMatrix(0, 2) = "Data"
        flxPedidos.TextMatrix(0, 3) = "Valor"
        flxPedidos.TextMatrix(0, 4) = "Qtde"
        flxPedidos.TextMatrix(0, 5) = "Faturado"
        flxPedidos.TextMatrix(0, 6) = "Saldo"
        flxPedidos.TextMatrix(0, 7) = "Tipo"
        flxPedidos.TextMatrix(0, 8) = "Status"
        
        flxPedidos.ColWidth(0) = 0
        flxPedidos.ColWidth(1) = 1000
        flxPedidos.ColWidth(2) = 1000
        flxPedidos.ColWidth(3) = 1000
        flxPedidos.ColWidth(4) = 1000
        flxPedidos.ColWidth(5) = 1000
        flxPedidos.ColWidth(6) = 1000
        flxPedidos.ColWidth(7) = 3000
        flxPedidos.ColWidth(8) = 1000
        
End Sub

Private Sub PegaPedidos(lngCODCOTA As Long)

    Dim I As Integer
    Call ConfGrdPedido

    sSql = "Select " & vbCrLf
    sSql = sSql & "       PEDI.* " & vbCrLf
    sSql = sSql & "      ,TIPI.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,VEND.SGI_DESCRICAO AS NOME" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH PEDI" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA  TIPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADVENDEDOR VEND"
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PEDI.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDI.SGI_CODCLI  = " & objCADCLIENTE.CLIECODIGO & vbCrLf
    sSql = sSql & "   And PEDI.SGI_CODCOTA = " & lngCODCOTA & vbCrLf
    sSql = sSql & "   And TIPI.SGI_FILIAL  = PEDI.SGI_FILIAL    " & vbCrLf
    sSql = sSql & "   And TIPI.SGI_CODIGO  = PEDI.SGI_CODTIPORC " & vbCrLf
    sSql = sSql & "   And VEND.SGI_FILIAL = PEDI.SGI_FILIAL    " & vbCrLf
    sSql = sSql & "   And VEND.SGI_CODIGO = PEDI.SGI_CODVEND   " & vbCrLf
    sSql = sSql & " Order by PEDI.SGI_DATAPED "
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC2.EOF
       
       flxPedidos.AddItem "" & vbTab & _
                          Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4) & vbTab & _
                          Format(BREC2!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                          Format(BREC2!SGI_VLTOT, "#,##0.00") & vbTab & _
                          Format(PegaQtdePedido(BREC2!SGI_CODCOTA), "#,###0.000") & vbTab & _
                          Format(0, "#,###0.000") & vbTab & _
                          Trim(BREC2!SGI_DESCRICAO) & vbTab & _
                          IIf(Trim(BREC2!SGI_STATUS) = "B", "Bloqueado", IIf(Trim(BREC2!SGI_STATUS) = "L", "Liberado", ""))

        For I = 1 To (flxPedidos.Cols - 1)
            flxPedidos.Row = (flxPedidos.Rows - 1)
            flxPedidos.Col = I
            If Trim(BREC2!SGI_STATUS) = "B" Then
               flxPedidos.CellForeColor = &HFF&
            ElseIf Trim(BREC2!SGI_STATUS) = "L" Or Trim(BREC2!SGI_STATUS) = "N" Then
               flxPedidos.CellForeColor = &H8000&
            End If
        Next I
    
       BREC2.MoveNext
    Loop
    BREC2.Close
    
End Sub

Public Sub ConfGrdFat()

    flxFaturamento.Rows = 1
    flxFaturamento.Cols = 7
    
    flxFaturamento.TextMatrix(0, 0) = ""
    flxFaturamento.TextMatrix(0, 1) = "Cod.Nf"
    flxFaturamento.TextMatrix(0, 2) = "Dt.Emissão"
    flxFaturamento.TextMatrix(0, 3) = "Dt.Saida"
    flxFaturamento.TextMatrix(0, 4) = "Cod.Romaneio"
    flxFaturamento.TextMatrix(0, 5) = "Status"
    flxFaturamento.TextMatrix(0, 6) = "Valor Nf"
    
    flxFaturamento.ColWidth(0) = 0
    flxFaturamento.ColWidth(1) = 1000
    flxFaturamento.ColWidth(2) = 1000
    flxFaturamento.ColWidth(3) = 1000
    flxFaturamento.ColWidth(4) = 1000
    flxFaturamento.ColWidth(5) = 1000
    flxFaturamento.ColWidth(6) = 1000

End Sub

Private Sub DesabilitaCampos()

   stCLIENTE.TabVisible(4) = False
   stCLIENTE.TabVisible(5) = False
   stCLIENTE.TabVisible(6) = False
   stCLIENTE.TabVisible(7) = False
   ''stCLIENTE.TabVisible(8) = False
   stCLIENTE.TabVisible(9) = False
   stCLIENTE.TabVisible(13) = True
   Label60.Visible = False
   dtNasc.Visible = False
   
   stCLIENTE.TabVisible(3) = False
   
   sSql = ""
   
   sSql = "Select" & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_USUARIO" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & lngIDUsuario & vbCrLf
   sSql = sSql & "   And SGI_BLOQCRED = 1"
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC.EOF() Then stCLIENTE.TabVisible(3) = True
   BREC.Close

   If lngIDUsuario = 0 Then
        stCLIENTE.TabVisible(3) = True
   End If

End Sub

Private Sub PopPedidos(intINDEX As Integer)

    Dim curSaldo    As Currency
    Dim curQtdPed   As Currency
    Dim curQtdFat   As Currency
    Dim strEMPTIPO  As String
    
    Call ConfGrdPedido
    
    strEMPTIPO = ""
    If intINDEX = 0 Then strEMPTIPO = "_STEEL"

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PEDI.* " & vbCrLf
    sSql = sSql & "      ,TIPI.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH" & strEMPTIPO & " PEDI" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPORCA TIPI" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PEDI.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDI.SGI_CODCLI = " & objCADCLIENTE.CLIECODIGO & vbCrLf
    sSql = sSql & "   And TIPI.SGI_FILIAL = PEDI.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And TIPI.SGI_CODIGO = PEDI.SGI_CODTIPORC "
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC2.EOF() Then
        With flxPedidos
            Do While Not BREC2.EOF()
            
                curSaldo = 0
                curQtdPed = 0
                
                If Not IsNull(BREC2!SGI_QTDEITENSPEDIDO) Then curQtdPed = BREC2!SGI_QTDEITENSPEDIDO
                curQtdFat = PegaQtdeFaturada(BREC2!SGI_CODIGO)
                
                curSaldo = (curQtdPed - curQtdFat)
                
                .AddItem "" & vbTab & _
                         Mid(Trim(Str(BREC2!SGI_CODIGO)), 1, (Len(Trim(Str(BREC2!SGI_CODIGO))) - 4)) & "/" & Right(Str(BREC2!SGI_CODIGO), 4) & vbTab & _
                         Format(BREC2!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                         Format(BREC2!SGI_VLTOT, "#,##0.00") & vbTab & _
                         Format(BREC2!SGI_QTDEITENSPEDIDO, "#,###0.000") & vbTab & _
                         Format(curQtdFat, "#,###0.000") & vbTab & _
                         Format(curSaldo, " #,###0.000") & vbTab & _
                         Trim(BREC2!SGI_DESCRICAO) & vbTab & _
                         IIf(Trim(BREC2!SGI_STATUS) = "B", "Bloqueado", IIf(Trim(BREC2!SGI_STATUS) = "L", "Liberado", IIf(Trim(BREC2!SGI_STATUS) = "N", "Liberado", IIf(Trim(BREC2!SGI_STATUS) = "R", "Reprovado", IIf(Trim(BREC2!SGI_STATUS) = "F", "Faturado", "")))))
            
                BREC2.MoveNext
            Loop
        End With
    End If
    BREC2.Close

End Sub

Private Function PegaQtdeFaturada(lngCODPEDIDO As Long) As Currency

    PegaQtdeFaturada = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(ORDI.SGI_QTDFAT) As SGI_QTDFAT " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH ORDF" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI ORDI" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDF.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDF.SGI_CODPED = " & lngCODPEDIDO & vbCrLf
    sSql = sSql & "   And ORDI.SGI_FILIAL = ORDF.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ORDI.SGI_CODORD = ORDF.SGI_CODORD "
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF() Then
       If Not IsNull(BREC3!SGI_QTDFAT) Then PegaQtdeFaturada = BREC3!SGI_QTDFAT
    End If
    BREC3.Close
    
End Function

Private Sub PegaDadosFatItens(intINDEX As Integer)

    Call ConfGridProdutos
    
    Dim strEMPRESA As String
    strEMPRESA = ""
    If intINDEX = 0 Then strEMPRESA = "_STEEL"
    
    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       CONF.SGI_DATACONF  " & vbCrLf
    sSql = sSql & "      ,CONF.SGI_CODFATURA " & vbCrLf
    sSql = sSql & "      ,PGTO.SGI_DESCRICAO As SGI_DESCPGTO " & vbCrLf
    sSql = sSql & "      ,CONFI.SGI_VLUNIT   " & vbCrLf
    sSql = sSql & "      ,CONFI.SGI_QTDREAL  " & vbCrLf
    sSql = sSql & "      ,CONFI.SGI_CODPRODUTO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO  " & vbCrLf
    sSql = sSql & "      ,CONFI.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & " From " & vbCrLf
    sSql = sSql & "      SGI_CADPEDVENDH" & strEMPRESA & " PEDVH " & vbCrLf
    sSql = sSql & "     ,SGI_CADCONDPGTO PGTO  " & vbCrLf
    sSql = sSql & "     ,SGI_CADORDFATH" & strEMPRESA & "  ORDF  " & vbCrLf
    sSql = sSql & "     ,SGI_CADORDCONFH" & strEMPRESA & " CONF  " & vbCrLf
    sSql = sSql & "     ,SGI_CADORDCONFI" & strEMPRESA & " CONFI " & vbCrLf
    sSql = sSql & "     ,SGI_CADPRODUTO  PROD " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      PEDVH.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & " And  PEDVH.SGI_CODCLI   = " & objCADCLIENTE.CLIECODIGO & vbCrLf
    sSql = sSql & " And  PGTO.SGI_FILIAL    = PEDVH.SGI_FILIAL " & vbCrLf
    sSql = sSql & " And  PGTO.SGI_CODIGO    = PEDVH.SGI_CODCONDPGT " & vbCrLf
    sSql = sSql & " And  ORDF.SGI_FILIAL    = PEDVH.SGI_FILIAL " & vbCrLf
    sSql = sSql & " And  ORDF.SGI_CODPED    = PEDVH.SGI_CODIGO " & vbCrLf
    sSql = sSql & " And  CONF.SGI_FILIAL    = ORDF.SGI_FILIAL  " & vbCrLf
    sSql = sSql & " And  CONF.SGI_CODORD    = ORDF.SGI_CODORD  " & vbCrLf
    sSql = sSql & " And  CONFI.SGI_FILIAL   = CONF.SGI_FILIAL  " & vbCrLf
    sSql = sSql & " And  CONFI.SGI_CODCONF  = CONF.SGI_CODCONF " & vbCrLf
    sSql = sSql & " And  PROD.SGI_FILIAL    = CONFI.SGI_FILIAL " & vbCrLf
    sSql = sSql & " And  PROD.SGI_IDPRODUTO = CONFI.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "Order By CONF.SGI_DATACONF " & vbCrLf
    sSql = sSql & "        ,CONF.SGI_CODFATURA " & vbCrLf
    sSql = sSql & "        ,CONFI.SGI_CODPRODUTO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
    
        With grdITENSPEDIDO
        
             .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                      Format(BREC!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                      BREC!SGI_CODFATURA & vbTab & _
                      Trim(BREC!SGI_DESCPGTO) & vbTab & _
                      Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                      BREC!SGI_QTDREAL & vbTab & _
                      Trim(BREC!SGI_CODPRODUTO) & vbTab & _
                      Trim(BREC!SGI_DESCRICAO)
        End With
    
        BREC.MoveNext
    Loop
    BREC.Close
    
End Sub


Private Sub ConfGridProdutos()

    With grdITENSPEDIDO
    
       .Cols = conColumnsIn_Produto
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Produto_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_Produto_IdProduto) = ""
       .ColDataType(conCOL_Produto_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_DataEmissao) = ""
       .ColDataType(conCOL_Produto_DataEmissao) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_Produto_NumeroFatura) = ""
       .ColDataType(conCOL_Produto_NumeroFatura) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_Vendmento) = ""
       .ColDataType(conCOL_Produto_Vendmento) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Produto_ValorUnit) = ""
       .ColDataType(conCOL_Produto_ValorUnit) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_Quantidade) = ""
       .ColDataType(conCOL_Produto_Quantidade) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_CodProd) = ""
       .ColDataType(conCOL_Produto_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Produto_Descricao) = ""
       .ColDataType(conCOL_Produto_Descricao) = flexDTString
       
       .ColWidth(conCOL_Produto_IdProduto) = 0
       .ColWidth(conCOL_Produto_DataEmissao) = 1200
       .ColWidth(conCOL_Produto_NumeroFatura) = 1200
       .ColWidth(conCOL_Produto_Vendmento) = 3000
       .ColWidth(conCOL_Produto_ValorUnit) = 1200
       .ColWidth(conCOL_Produto_Quantidade) = 1200
       .ColWidth(conCOL_Produto_CodProd) = 1200
       .ColWidth(conCOL_Produto_Descricao) = 3500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub ConfGridProdutosClie()

    With grdPRODUTOS
    
       .Cols = conColumnsIn_ProdutoClie
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdutoClie_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_ProdutoClie_IdProduto) = ""
       .ColDataType(conCOL_ProdutoClie_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdutoClie_Rotulo) = ""
       .ColDataType(conCOL_ProdutoClie_Rotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdutoClie_PesqRot) = ""
       .ColDataType(conCOL_ProdutoClie_PesqRot) = flexDTString
       .ColComboList(conCOL_ProdutoClie_PesqRot) = "..."
       
       .Cell(flexcpData, 0, conCOL_ProdutoClie_Descricao) = ""
       .ColDataType(conCOL_ProdutoClie_Descricao) = flexDTString
       
       .ColWidth(conCOL_ProdutoClie_IdProduto) = 1000
       .ColWidth(conCOL_ProdutoClie_Rotulo) = 1200
       .ColWidth(conCOL_ProdutoClie_PesqRot) = 300
       .ColWidth(conCOL_ProdutoClie_Descricao) = 6000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub LimpaCamposGrid(lngROW As Long)
    With grdPRODUTOS
        .Cell(flexcpText, lngROW, conCOL_ProdutoClie_IdProduto) = ""
        .Cell(flexcpText, lngROW, conCOL_ProdutoClie_Rotulo) = ""
        .Cell(flexcpText, lngROW, conCOL_ProdutoClie_Descricao) = ""
    End With
End Sub

Private Function PegaIDProduto(strCodProduto As String) As String

    PegaIDProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And (PRO.SGI_STATUS = 1 Or PRO.SGI_STATUS = 2)" & vbCrLf

    sSql = sSql & "   And (Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End) = '" & Trim(strCodProduto) & "'"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If BREC!SGI_STATUS = 0 Then
           MsgBox "ATENÇÃO !!!" & vbCrLf & "O Produto " & Trim(strCodProduto) & " - " & Trim(BREC!SGI_DESCRICAO) & vbCrLf & "Não pode ser Utilizado está Desativado !!!", vbOKOnly + vbExclamation, "Aviso"
        Else
           PegaIDProduto = BREC!SGI_IDPRODUTO
        End If
    End If
    BREC.Close
    
End Function


Private Function PegaDescrProduto(strCodProduto As String) As String
    
    PegaDescrProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.* " & vbCrLf
    sSql = sSql & "      ,LINHA.SGI_FILIALPED As FILIALPED_2" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINHA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_IDPRODUTO   = " & strCodProduto & vbCrLf
    sSql = sSql & "   And (PRO.SGI_STATUS     = 1 or PRO.SGI_STATUS      = 2)" & vbCrLf
    sSql = sSql & "   And LINHA.SGI_FILIAL    = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINHA.SGI_CODLIN    = PRO.SGI_CODLINPROD" & vbCrLf
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then
        If BREC2!SGI_CODTIPO = 2 Then
            PegaDescrProduto = BREC2!SGI_DESCRICAO
        Else
            PegaDescrProduto = BREC2!SGI_DESCRICAO
        End If
    End If
    BREC2.Close
    
End Function

Private Sub IncRegGridProdtos()
   
    If objBLBFunc.FcExisteLinhaVazia(grdPRODUTOS, conCOL_ProdutoClie_IdProduto) = False Then Exit Sub
    
    grdPRODUTOS.AddItem "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        ""
                            
End Sub

Private Sub PopGrdProdClie()

    Dim I As Integer
    
    If IsArray(arrPRODUTOSCLIE) Then
        With grdPRODUTOS
            For I = 1 To UBound(arrPRODUTOSCLIE)
                .AddItem arrPRODUTOSCLIE(I, 1) & vbTab & _
                         arrPRODUTOSCLIE(I, 2) & vbTab & _
                         "" & vbTab & _
                         arrPRODUTOSCLIE(I, 3)
            Next I
        End With
    End If

End Sub



Private Sub ConfGrdUltFat()

    With grdULTFAT
    
        .Cols = conColumnsIn_UltFat
        .Rows = 1
        .FixedCols = 0
        .FormatString = conCOL_UltFat_FormatString
       
        .AutoSizeMouse = False
        .AllowUserResizing = flexResizeNone
       
        .Cell(flexcpData, 0, conCOL_UltFat_IdProduto) = ""
        .ColDataType(conCOL_UltFat_IdProduto) = flexDTLong
        
        .Cell(flexcpData, 0, conCOL_UltFat_DtFat) = ""
        .ColDataType(conCOL_UltFat_DtFat) = flexDTDate
        
        .Cell(flexcpData, 0, conCOL_UltFat_DtEntrega) = ""
        .ColDataType(conCOL_UltFat_DtEntrega) = flexDTDate
        
        .Cell(flexcpData, 0, conCOL_UltFat_QtdFat) = ""
        .ColDataType(conCOL_UltFat_QtdFat) = flexDTLong
        
        .Cell(flexcpData, 0, conCOL_UltFat_QtdPed) = ""
        .ColDataType(conCOL_UltFat_QtdPed) = flexDTLong
        
        .Cell(flexcpData, 0, conCOL_UltFat_Saldo) = ""
        .ColDataType(conCOL_UltFat_Saldo) = flexDTLong
        
        .Cell(flexcpData, 0, conCOL_UltFat_Unit) = ""
        .ColDataType(conCOL_UltFat_Unit) = flexDTCurrency
        
        .Cell(flexcpData, 0, conCOL_UltFat_Valor) = ""
        .ColDataType(conCOL_UltFat_Valor) = flexDTCurrency
       
        .Cell(flexcpData, 0, conCOL_UltFat_CodConf) = ""
        .ColDataType(conCOL_UltFat_CodConf) = flexDTLong
       
        .Cell(flexcpData, 0, conCOL_UltFat_CodOP) = ""
        .ColDataType(conCOL_UltFat_CodOP) = flexDTLong
       
        .Cell(flexcpData, 0, conCOL_UltFat_CodPed) = ""
        .ColDataType(conCOL_UltFat_CodPed) = flexDTLong
       
        .Cell(flexcpData, 0, conCOL_UltFat_CodNF) = ""
        .ColDataType(conCOL_UltFat_CodNF) = flexDTLong
       
        .Cell(flexcpData, 0, conCOL_UltFat_Status) = ""
        .ColDataType(conCOL_UltFat_Status) = flexDTString
       
        .ColWidth(conCOL_UltFat_IdProduto) = 0
        .ColWidth(conCOL_UltFat_DtFat) = 1200
        .ColWidth(conCOL_UltFat_DtEntrega) = 1200
        .ColWidth(conCOL_UltFat_QtdFat) = 1200
        .ColWidth(conCOL_UltFat_QtdPed) = 1200
        .ColWidth(conCOL_UltFat_Saldo) = 1200
        .ColWidth(conCOL_UltFat_Unit) = 1200
        .ColWidth(conCOL_UltFat_Valor) = 1200
        .ColWidth(conCOL_UltFat_CodConf) = 1200
        .ColWidth(conCOL_UltFat_CodOP) = 1200
        .ColWidth(conCOL_UltFat_CodPed) = 1200
        .ColWidth(conCOL_UltFat_CodNF) = 1200
        .ColWidth(conCOL_UltFat_Status) = 1200
       
        .Editable = flexEDNone
        .AllowSelection = False
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .BackColor = &H80000018
        .ForeColor = vbBlack
    
    End With

End Sub

Private Sub PopGrdUltFat()

    
    Dim strEMPRESA As String
    
    strEMPRESA = ""
    If optFILIALPRODFAT(1).Value = True Then strEMPRESA = "_STEEL"
    
    sSql = ""

    sSql = "Select" & vbCrLf
    sSql = sSql & "      CABEC.SGI_DATACONF" & vbCrLf
    sSql = sSql & "     ,ITEN.*" & vbCrLf
    sSql = sSql & "     ,OP.SGI_CODIGO As SGI_CODOP" & vbCrLf
    sSql = sSql & "     ,OP.SGI_CODPED" & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_CODFATURA" & vbCrLf
    sSql = sSql & "     ,OP.SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "     ,OP.SGI_QTDEPED" & vbCrLf
    sSql = sSql & "     ,OP.SGI_QTDFAT As SGI_OPQTDFAT" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_PRODCLIE          PRODCL" & vbCrLf
    sSql = sSql & "     ,SGI_CADORDCONFI" & strEMPRESA & " ITEN" & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD" & strEMPRESA & "   OP" & vbCrLf
    sSql = sSql & "     ,SGI_CADORDCONFH" & strEMPRESA & " CABEC" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      PRODCL.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "  And PRODCL.SGI_CODIGO    = " & objCADCLIENTE.CLIECODIGO & vbCrLf
    sSql = sSql & "  And PRODCL.SGI_FILIAL    = ITEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And PRODCL.SGI_IDPRODUTO = ITEN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "  And ITEN.SGI_FILIAL      = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And ITEN.SGI_CODORDPROD  = OP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  And ITEN.SGI_IDPRODUTO   = OP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "  And ITEN.SGI_FILIAL      = CABEC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And ITEN.SGI_CODCONF     = CABEC.SGI_CODCONF" & vbCrLf
    sSql = sSql & "Order BY CABEC.SGI_DATACONF DESC" & vbCrLf
    sSql = sSql & "        ,ITEN.SGI_IDPRODUTO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        With grdULTFAT
            Do While Not BREC.EOF()
            
                .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                         Format(BREC!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                         Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                         Format(BREC!SGI_QTDREAL, "#0") & vbTab & _
                         Format(BREC!SGI_QTDEPED, "#0") & vbTab & _
                         Format((BREC!SGI_QTDEPED - BREC!SGI_QTDREAL), "#0") & vbTab & _
                         Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                         Format(BREC!SGI_VLTOTAL, "#,##0.00") & vbTab & _
                         BREC!SGI_CODCONF & vbTab & _
                         BREC!SGI_CODPED & vbTab & _
                         BREC!SGI_CODOP & vbTab & _
                         BREC!SGI_CODFATURA
                         
                BREC.MoveNext
            Loop
        End With
    End If
    BREC.Close
    
    Call MostraItens

End Sub

Private Sub Mostra_Itens(strIDPRODUTO As String)
    Dim I As Long
    With grdULTFAT
        For I = 1 To (.Rows - 1)
            .RowHidden(I) = True
            If Trim(.Cell(flexcpText, I, conCOL_UltFat_IdProduto)) = Trim(strIDPRODUTO) Then .RowHidden(I) = False
        Next I
    End With
    
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADCLIENTE = Nothing
    Set objPESQPADRAO = Nothing
    Set objCADCONFAT = Nothing
    Set objCADPEDIDO = Nothing
End Sub

Private Sub AbreTelas()
        Dim I              As Integer
        Dim intFILIALPED   As Integer
     
        intFILIALPED = 0
        If optFILIALPRODFAT(1).Value = True Then intFILIALPED = 1
     
        With grdULTFAT
             Select Case .Col
                    Case conCOL_UltFat_CodConf
                           If Len(Trim(.Cell(flexcpText, .Row, conCOL_UltFat_CodConf))) = 0 Then Exit Sub
                           objCADCONFAT.CODCONF = .Cell(flexcpText, .Row, conCOL_UltFat_CodConf)
                           Call objCADCONFAT.cConnectPesq(cCaminho, Linha, FILIAL, strAcesso, strUSUARIO, lngIDUsuario, intFILIALPED, True)
                    Case conCOL_UltFat_CodPed
                           If Len(Trim(.Cell(flexcpText, .Row, conCOL_UltFat_CodPed))) = 0 Then Exit Sub
                           objCADPEDIDO.CODPEDIDO = .Cell(flexcpText, .Row, conCOL_UltFat_CodPed)
                           Call objCADPEDIDO.cConnectPesq(cCaminho, Linha, FILIAL, strAcesso, strUSUARIO, lngIDUsuario, intFILIALPED, True)
             End Select
        End With

End Sub

Private Sub MostraItens()
    With grdPRODUTOS
        If .Row > 0 Then
            Call Mostra_Itens(Trim(.Cell(flexcpText, .Row, conCOL_ProdutoClie_IdProduto)))
        End If
    End With
End Sub

Private Sub PopTodosPed()

    Dim strEMPRESA As String
    Dim strCODOP   As String
    
    strEMPRESA = ""
    If optFILIALPRODFAT(1).Value = True Then strEMPRESA = "_STEEL"
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       ITEPED.*" & vbCrLf
    sSql = sSql & "     , PROGEN.*" & vbCrLf

    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_PRODCLIE PRODCL" & vbCrLf
    sSql = sSql & "     , SGI_CADPEDVENDH" & strEMPRESA & "  HEAPED" & vbCrLf
    sSql = sSql & "     , SGI_CADPEDVENDI" & strEMPRESA & "  ITEPED" & vbCrLf
    sSql = sSql & "     , SGI_PROGENTRPROD" & strEMPRESA & " PROGEN" & vbCrLf

    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRODCL.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRODCL.SGI_CODIGO     = " & objCADCLIENTE.CLIECODIGO & vbCrLf
    sSql = sSql & "   And PRODCL.SGI_FILIAL     = HEAPED.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PRODCL.SGI_CODIGO     = HEAPED.SGI_CODCLI" & vbCrLf
    sSql = sSql & "   And PRODCL.SGI_FILIAL     = ITEPED.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PRODCL.SGI_IDPRODUTO  = ITEPED.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And HEAPED.SGI_CODIGO     = ITEPED.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And ITEPED.SGI_FILIAL     = PROGEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ITEPED.SGI_CODIGO     = PROGEN.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And ITEPED.SGI_IDPRODUTO  = PROGEN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "Order By ITEPED.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "        ,PROGEN.SGI_DATENTREGA DESC"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
        With grdULTFAT
            Do While Not BREC.EOF()
                
                strCODOP = ""
                If Not IsNull(BREC!SGI_IDINTERNO) Then strCODOP = PegaOPPesq(Trim(Str(BREC!SGI_IDINTERNO)), Str(BREC!SGI_IDPRODUTO), Str(BREC!SGI_CODIGO))
                
                If Len(Trim(strCODOP)) > 0 Then
                
                    sSql = ""
                    
                    sSql = "Select" & vbCrLf
                    sSql = sSql & "       ITEN.*" & vbCrLf
                    sSql = sSql & "      ,CABEC.SGI_DATACONF " & vbCrLf
                    sSql = sSql & "      ,CABEC.SGI_CODFATURA" & vbCrLf

                    sSql = sSql & "  From" & vbCrLf
                    sSql = sSql & "       SGI_CADORDCONFI" & strEMPRESA & " ITEN" & vbCrLf
                    sSql = sSql & "      ,SGI_CADORDCONFH" & strEMPRESA & " CABEC" & vbCrLf
                    
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       ITEN.SGI_FILIAL     = " & FILIAL & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_CODORDPROD = " & Trim(strCODOP) & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                
                    sSql = sSql & "   And ITEN.SGI_FILIAL     = CABEC.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_CODCONF    = CABEC.SGI_CODCONF" & vbCrLf
                    
                    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                
                    If Not BREC3.EOF() Then
                       Do While Not BREC3.EOF()
                       
                            .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                                     Format(BREC3!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                                     Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                                     Format(BREC3!SGI_QTDREAL, "#,##0.00") & vbTab & _
                                     Format(BREC!SGI_QTDE, "#0") & vbTab & _
                                     Format((BREC!SGI_QTDE - BREC3!SGI_QTDREAL), "#0") & vbTab & _
                                     Format(BREC3!SGI_VLUNIT, "#,##0.00") & vbTab & _
                                     Format(BREC3!SGI_VLTOTAL, "#,##0.00") & vbTab & _
                                     BREC3!SGI_CODCONF & vbTab & _
                                     BREC!SGI_CODIGO & vbTab & _
                                     strCODOP & vbTab & _
                                     BREC3!SGI_CODFATURA & vbTab & _
                                     ""
                          
                          BREC3.MoveNext
                       Loop
                    Else
                    
                        .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                                 "" & vbTab & _
                                 Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                                 "" & vbTab & _
                                 Format(BREC!SGI_QTDE, "#0") & vbTab & _
                                 "" & vbTab & _
                                 Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                                 Format(BREC!SGI_VLTOT, "#,##0.00") & vbTab & _
                                 "" & vbTab & _
                                 BREC!SGI_CODIGO & vbTab & _
                                 strCODOP & vbTab & _
                                 "" & vbTab & _
                                 ""
                    
                    End If
                    BREC3.Close
                
                Else
                    
                    .AddItem BREC!SGI_IDPRODUTO & vbTab & _
                             "" & vbTab & _
                             Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                             "" & vbTab & _
                             Format(BREC!SGI_QTDE, "#0") & vbTab & _
                             "" & vbTab & _
                             Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                             Format(BREC!SGI_VLTOT, "#,##0.00") & vbTab & _
                             "" & vbTab & _
                             BREC!SGI_CODIGO & vbTab & _
                             strCODOP & vbTab & _
                             "" & vbTab & _
                             ""
                End If
                BREC.MoveNext
            Loop
        End With
    End If
    BREC.Close
    
    Call MostraItens

End Sub

Private Function PegaOPPesq(strIDINTERNO As String, strIDPRODUTO As String, strCODPED As String) As String

    PegaOPPesq = ""

    If Len(Trim(strIDINTERNO)) = 0 Then Exit Function
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
    If Len(Trim(strCODPED)) = 0 Then Exit Function
    
    Dim strEMPRESA As String
    
    strEMPRESA = ""
    If optFILIALPRODFAT(1).Value = True Then strEMPRESA = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strEMPRESA & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPAI     = " & Trim(strIDINTERNO)
    sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(strIDPRODUTO)
    sSql = sSql & "   And SGI_CODPED    = " & Trim(strCODPED)
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then PegaOPPesq = Trim(Str(BREC2!SGI_CODIGO))
    BREC2.Close

End Function

Private Sub ConfGridProdutosABC()

    With grdPRODABC
    
       .Cols = conColumnsIn_ProdutoClieABC
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdutoClieABC_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_ProdutoClieABC_IdProduto) = ""
       .ColDataType(conCOL_ProdutoClieABC_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdutoClieABC_Rotulo) = ""
       .ColDataType(conCOL_ProdutoClieABC_Rotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdutoClieABC_PesqRot) = ""
       .ColDataType(conCOL_ProdutoClieABC_PesqRot) = flexDTString
       .ColComboList(conCOL_ProdutoClieABC_PesqRot) = "..."
       
       .Cell(flexcpData, 0, conCOL_ProdutoClieABC_Descricao) = ""
       .ColDataType(conCOL_ProdutoClieABC_Descricao) = flexDTString
       
       .ColWidth(conCOL_ProdutoClieABC_IdProduto) = 1000
       .ColWidth(conCOL_ProdutoClieABC_Rotulo) = 1200
       .ColWidth(conCOL_ProdutoClieABC_PesqRot) = 300
       .ColWidth(conCOL_ProdutoClieABC_Descricao) = 6000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub PopGrdProdClieABC()

    Dim I As Integer
    
    If IsArray(arrPRODUTOSCLIE) Then
        With grdPRODABC
            For I = 1 To UBound(arrPRODUTOSCLIE)
                .AddItem arrPRODUTOSCLIE(I, 1) & vbTab & _
                         arrPRODUTOSCLIE(I, 2) & vbTab & _
                         "" & vbTab & _
                         arrPRODUTOSCLIE(I, 3)
            Next I
        End With
    End If

End Sub

Private Sub ConfGridCurvaABC(ColumnsIn_ProdutoCurvABC As Long, ProdutoCurvABC_FormatString As String)

    With grdCURVAABC
    
       Dim I As Long
       
       .Cols = ColumnsIn_ProdutoCurvABC
       .Rows = 1
       .FixedCols = 0
       .FormatString = ProdutoCurvABC_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, 0) = ""
       .ColDataType(0) = flexDTLong
       
       .Cell(flexcpData, 0, 1) = ""
       .ColDataType(1) = flexDTString
       
       For I = 0 To (.Cols - 1)
            .ColWidth(I) = 1500
       Next I
       
       .ColWidth(0) = 0
       .ColWidth(1) = 1500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub CarregaCurvaABC(strCODCLIE As String, strIDPRODUTO As String)
    
    ProdutoCurvABC_FormatString = "=IdProduto|Descrição"
    ColumnsIn_ProdutoCurvABC = 1
    
    Call ConfGridCurvaABC(ColumnsIn_ProdutoCurvABC, ProdutoCurvABC_FormatString)
    
    If Len(Trim(strCODCLIE)) = 0 Then Exit Sub
    If Len(Trim(strIDPRODUTO)) = 0 Then
        MsgBox "Selecione um Produto !!!", vbExclamation + vbOKOnly, "Aviso"
        Exit Sub
    End If
    
    Frame55.Caption = "[ " & grdPRODABC.Cell(flexcpText, grdPRODABC.Row, conCOL_ProdutoClieABC_Rotulo) & " - " & grdPRODABC.Cell(flexcpText, grdPRODABC.Row, conCOL_ProdutoClieABC_Descricao) & " ]"
    
    
    Dim strMESANO       As String
    Dim strEMPRESA      As String
    Dim arrCOLS()       As String
    Dim lngCOL          As Long
    Dim lngTOTQTDPED    As Long
    Dim lngQTDEMESES    As Long
    Dim dblMEDIA        As Double
    Dim I               As Long
    
    strEMPRESA = ""
    If optEMP(1).Value = True Then strEMPRESA = "_STEEL"
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       ITEN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,Month(PED.SGI_DATAPED)   As SGI_MES" & vbCrLf
    sSql = sSql & "      ,Year(PED.SGI_DATAPED)    As SGI_ANO" & vbCrLf
    sSql = sSql & "      ,Sum(ITEN.SGI_QTDE)       As SGI_QTDE" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & " PED" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & " ITEN" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PED.SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And PED.SGI_CODCLI       = " & Trim(strCODCLIE) & vbCrLf
    sSql = sSql & "   And ITEN.SGI_FILIAL      = PED.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ITEN.SGI_CODIGO      = PED.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And ITEN.SGI_IDPRODUTO   = " & Trim(strIDPRODUTO) & vbCrLf
    
    sSql = sSql & "Group By" & vbCrLf
    sSql = sSql & "         ITEN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "        ,Month(PED.SGI_DATAPED)" & vbCrLf
    sSql = sSql & "        ,Year(PED.SGI_DATAPED)" & vbCrLf
    
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "         ITEN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "        ,Year(PED.SGI_DATAPED)  Desc" & vbCrLf
    sSql = sSql & "        ,Month(PED.SGI_DATAPED) Desc"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
        Do While Not BREC10.EOF()
            strMESANO = Format(BREC10!SGI_MES, "##00") & "/" & Trim(Str(BREC10!SGI_ANO))
            ProdutoCurvABC_FormatString = ProdutoCurvABC_FormatString & "|" & Trim(strMESANO)
            BREC10.MoveNext
        Loop
        ProdutoCurvABC_FormatString = ProdutoCurvABC_FormatString & "|Total|Média"
        arrCOLS = Split(ProdutoCurvABC_FormatString, "|")
    
        ColumnsIn_ProdutoCurvABC = UBound(arrCOLS)
        Call ConfGridCurvaABC(ColumnsIn_ProdutoCurvABC, ProdutoCurvABC_FormatString)
        
        With grdCURVAABC
            BREC10.MoveFirst
            .AddItem BREC10!SGI_IDPRODUTO & vbTab & "Pedidos" & vbTab & ""
            .AddItem BREC10!SGI_IDPRODUTO & vbTab & "Faturados" & vbTab & ""
        
            BREC10.MoveFirst
            strMESANO = ""
            lngTOTQTDPED = 0
            lngQTDEMESES = 0
            Do While Not BREC10.EOF()
                strMESANO = Format(BREC10!SGI_MES, "##00") & "/" & Trim(Str(BREC10!SGI_ANO))
                lngTOTQTDPED = (lngTOTQTDPED + BREC10!SGI_QTDE)
                
                For I = 0 To (.Cols - 1)
                    If .Cell(flexcpText, 0, I) = strMESANO Then
                        lngCOL = I
                        lngQTDEMESES = (lngQTDEMESES + 1)
                        Exit For
                    End If
                Next I
            
                .Cell(flexcpText, 1, lngCOL) = BREC10!SGI_QTDE
                
                BREC10.MoveNext
            Loop
            If lngTOTQTDPED > 0 Then .Cell(flexcpText, 1, (.Cols - 2)) = lngTOTQTDPED
            If lngQTDEMESES > 0 Then
                dblMEDIA = (lngTOTQTDPED / lngQTDEMESES)
                .Cell(flexcpText, 1, (.Cols - 1)) = Format(dblMEDIA, "#0")
            End If
        End With
    
    End If
    BREC10.Close
  
  
    If (grdCURVAABC.Rows - 1) = 0 Then Exit Sub
  
    '' -------------------------
    '' Pegando o Faturmento
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       CONFI.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,Month(ORDH.SGI_DATAORDEM) As SGI_MES" & vbCrLf
    sSql = sSql & "      ,Year(ORDH.SGI_DATAORDEM)  As SGI_ANO" & vbCrLf
    sSql = sSql & "      ,Sum(CONFI.SGI_QTDREAL)    As SGI_QTDREAL" & vbCrLf
    
    sSql = sSql & " From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & " PEDH" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATH" & strEMPRESA & "  ORDH" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFH" & strEMPRESA & " CONF" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFI" & strEMPRESA & " CONFI" & vbCrLf

    sSql = sSql & "Where" & vbCrLf
    sSql = sSql & "       PEDH.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDH.SGI_CODCLI     = " & Trim(strCODCLIE) & vbCrLf
    sSql = sSql & "   And ORDH.SGI_FILIAL     = PEDH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ORDH.SGI_CODPED     = PEDH.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And CONF.SGI_FILIAL     = ORDH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CONF.SGI_CODORD     = ORDH.SGI_CODORD" & vbCrLf
    sSql = sSql & "   And CONFI.SGI_FILIAL    = CONF.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CONFI.SGI_CODCONF   = CONF.SGI_CODCONF" & vbCrLf
    sSql = sSql & "   And CONFI.SGI_IDPRODUTO = " & Trim(strIDPRODUTO) & vbCrLf

    sSql = sSql & "Group By" & vbCrLf
    sSql = sSql & "         CONFI.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "        ,Month(ORDH.SGI_DATAORDEM)" & vbCrLf
    sSql = sSql & "        ,Year(ORDH.SGI_DATAORDEM)" & vbCrLf

    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "         CONFI.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "        ,Month(ORDH.SGI_DATAORDEM) Desc" & vbCrLf
    sSql = sSql & "        ,Year(ORDH.SGI_DATAORDEM)  Desc"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
    
        With grdCURVAABC
            strMESANO = ""
            lngTOTQTDPED = 0
            lngQTDEMESES = 0
            Do While Not BREC10.EOF()
                strMESANO = Format(BREC10!SGI_MES, "##00") & "/" & Trim(Str(BREC10!SGI_ANO))
                lngTOTQTDPED = (lngTOTQTDPED + BREC10!SGI_QTDREAL)
                
                For I = 0 To (.Cols - 1)
                    If Trim(.Cell(flexcpText, 0, I)) = Trim(strMESANO) Then
                        lngCOL = I
                        lngQTDEMESES = (lngQTDEMESES + 1)
                        Exit For
                    End If
                Next I
            
                .Cell(flexcpText, 2, lngCOL) = BREC10!SGI_QTDREAL
                
                BREC10.MoveNext
            Loop
            If lngTOTQTDPED > 0 Then .Cell(flexcpText, 2, (.Cols - 2)) = lngTOTQTDPED
            If lngQTDEMESES > 0 Then
                dblMEDIA = (lngTOTQTDPED / lngQTDEMESES)
                .Cell(flexcpText, 2, (.Cols - 1)) = Format(dblMEDIA, "#0")
            End If
        End With
    
    End If
    BREC10.Close

End Sub

Private Sub PopCurvaABC()
    If (grdPRODABC.Rows - 1) = 0 Then Exit Sub
    If grdPRODABC.Row = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & "Selecione um Produto !!!", vbExclamation + vbOKOnly, "Aviso"
        Exit Sub
    End If
    Call CarregaCurvaABC(Str(objCADCLIENTE.CLIECODIGO), grdPRODABC.Cell(flexcpText, grdPRODABC.Row, conCOL_ProdutoClieABC_IdProduto))
End Sub

Private Sub ConfGridProdutosPedidos()

    With grdPedidos
    
       .Cols = conColumnsIn_ProdutoPedidos
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdutoPedidos_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_ProdutoPedidos_MesAno) = ""
       .ColDataType(conCOL_ProdutoPedidos_MesAno) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdutoPedidos_Codigo) = ""
       .ColDataType(conCOL_ProdutoPedidos_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdutoPedidos_Qtde) = ""
       .ColDataType(conCOL_ProdutoPedidos_Qtde) = flexDTLong
       
       .ColWidth(conCOL_ProdutoPedidos_MesAno) = 1000
       .ColWidth(conCOL_ProdutoPedidos_Codigo) = 1000
       .ColWidth(conCOL_ProdutoPedidos_Qtde) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub ConfGridProdutosFaturado()

    With grdFaturado
    
       .Cols = conColumnsIn_ProdutoFaturado
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdutoFaturado_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_ProdutoFaturado_MesAno) = ""
       .ColDataType(conCOL_ProdutoFaturado_MesAno) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdutoFaturado_Codigo) = ""
       .ColDataType(conCOL_ProdutoFaturado_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdutoFaturado_Qtde) = ""
       .ColDataType(conCOL_ProdutoFaturado_Qtde) = flexDTLong
       
       .ColWidth(conCOL_ProdutoFaturado_MesAno) = 1000
       .ColWidth(conCOL_ProdutoFaturado_Codigo) = 1000
       .ColWidth(conCOL_ProdutoFaturado_Qtde) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub



Private Sub ConfGridClienteCondPgto()

    With grdCONDPGTO
    
       .Cols = conColumnsIn_Cliente
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Cliente_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_Cliente_CodCOndPgto) = ""
       .ColDataType(conCOL_Cliente_CodCOndPgto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Cliente_Pesq) = ""
       .ColDataType(conCOL_Cliente_Pesq) = flexDTString
       .ColComboList(conCOL_Cliente_Pesq) = "..."
       
       .Cell(flexcpData, 0, conCOL_Cliente_DescCondPgto) = ""
       .ColDataType(conCOL_Cliente_DescCondPgto) = flexDTString
       
       .ColWidth(conCOL_Cliente_CodCOndPgto) = 1000
       .ColWidth(conCOL_Cliente_Pesq) = 300
       .ColWidth(conCOL_Cliente_DescCondPgto) = 6000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub IncRegGridCondPgto()
   
On Error GoTo Err_IncRegGridCondPgto
    
    Dim strCampos01 As String
    
    
    If objBLBFunc.FcExisteLinhaVazia(grdCONDPGTO, conCOL_Cliente_CodCOndPgto) = False Then Exit Sub
    
    strCampos01 = "" & vbTab & _
                  "" & vbTab & _
                  ""
    
    
                       
    grdCONDPGTO.AddItem strCampos01
    
    Exit Sub

Err_IncRegGridCondPgto:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : IncRegGridCondPgto()", Me.Name, "IncRegGridCondPgto()", strCAMARQERRO)
                            
End Sub

Private Sub LimpaCamposGridCodCondPgto(lngROW As Long)
    With grdCONDPGTO
        .Cell(flexcpText, lngROW, conCOL_Cliente_CodCOndPgto) = Empty
        .Cell(flexcpText, lngROW, conCOL_Cliente_DescCondPgto) = Empty
    End With
End Sub

Private Function PesDescCondPgto(strCodCondPgto, lngROW As Long) As String

    PesDescCondPgto = ""
    
    If Len(Trim(strCodCondPgto)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCONDPGTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCodCondPgto)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PesDescCondPgto = Trim(BREC!SGI_DESCRICAO)
    BREC.Close

End Function

Private Sub PopGrdCondPgto()

    Dim I As Long
    
    Call ConfGridClienteCondPgto

    arrCONDPGTOCLIE = objCADCLIENTE.CONDPGTOCLIE

    If IsArray(arrCONDPGTOCLIE) Then
    
        With grdCONDPGTO
            For I = 1 To UBound(arrCONDPGTOCLIE)
                .AddItem arrCONDPGTOCLIE(I, 1) & vbTab & _
                         "" & vbTab & _
                         ""
                         
                .Cell(flexcpText, I, conCOL_Cliente_DescCondPgto) = Trim(PesDescCondPgto(arrCONDPGTOCLIE(I, 1), I))
            Next I
        End With
    
    End If

End Sub
