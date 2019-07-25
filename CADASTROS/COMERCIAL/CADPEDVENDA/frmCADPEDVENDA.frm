VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCADPEDVENDA 
   Caption         =   "Cadastro de Pedidos de Venda"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   17580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   17535
      Begin VB.CommandButton cmdCancLibFin 
         Caption         =   "&Cancela"
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
         Left            =   15720
         Picture         =   "frmCADPEDVENDA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   170
         ToolTipText     =   "Bloqueia o Pedido"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdLiberaFinanceiro 
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
         Left            =   14040
         Picture         =   "frmCADPEDVENDA.frx":0426
         Style           =   1  'Graphical
         TabIndex        =   169
         ToolTipText     =   "Liberação Financeira"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancLibCom 
         Caption         =   "&Cancela"
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
         Left            =   15720
         Picture         =   "frmCADPEDVENDA.frx":0853
         Style           =   1  'Graphical
         TabIndex        =   168
         ToolTipText     =   "Bloqueia o Pedido"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdLiberaCom 
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
         Left            =   14040
         Picture         =   "frmCADPEDVENDA.frx":0C79
         Style           =   1  'Graphical
         TabIndex        =   167
         ToolTipText     =   "Liberação Comercial"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancLibFot 
         Caption         =   "&Cancela"
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
         Left            =   15720
         Picture         =   "frmCADPEDVENDA.frx":10A6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Bloqueia o Pedido"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdLibFot 
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
         Height          =   735
         Left            =   14040
         Picture         =   "frmCADPEDVENDA.frx":14CC
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Liberação das Alterações do Fotolito"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancAlteracao 
         Caption         =   "&Cancela"
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
         Left            =   15720
         Picture         =   "frmCADPEDVENDA.frx":18F9
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Bloqueia o Pedido"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdLibAlteracao 
         Caption         =   "&Libera Alteração"
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
         Left            =   14040
         Picture         =   "frmCADPEDVENDA.frx":1D1F
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Liberação das alterações"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancLibPcotaPdata 
         Caption         =   "&Cancela"
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
         Left            =   15720
         Picture         =   "frmCADPEDVENDA.frx":214C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Bloqueia o Pedido"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdLibPcotaPData 
         Caption         =   "&Libera P.Data/P.Cota"
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
         Left            =   14040
         Picture         =   "frmCADPEDVENDA.frx":2572
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Liberação das alterações"
         Top             =   120
         Width           =   1695
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
         Picture         =   "frmCADPEDVENDA.frx":299F
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmCADPEDVENDA.frx":2AA1
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
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
         Left            =   960
         Picture         =   "frmCADPEDVENDA.frx":2BA3
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin TabDlg.SSTab stCAMPOSVENDA 
      Height          =   9255
      Left            =   0
      TabIndex        =   23
      Top             =   1005
      Width           =   17535
      _ExtentX        =   30930
      _ExtentY        =   16325
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
      TabCaption(0)   =   "Dados do Pedido"
      TabPicture(0)   =   "frmCADPEDVENDA.frx":30D5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdLIBPDATAPCOTA"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame25"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame30"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Local de Entrega\Local de Cobrança"
      TabPicture(1)   =   "frmCADPEDVENDA.frx":30F1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(3)=   "Frame14"
      Tab(1).Control(4)=   "Frame11"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Reprovação\Log de Ações"
      TabPicture(2)   =   "frmCADPEDVENDA.frx":310D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame27"
      Tab(2).Control(1)=   "Frame26"
      Tab(2).Control(2)=   "Frame21"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "[ Produção ]"
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
         Height          =   1095
         Left            =   5280
         TabIndex        =   185
         Top             =   8100
         Width           =   3495
         Begin VSFlex8LCtl.VSFlexGrid grdPRODUCAO 
            Height          =   735
            Left            =   120
            TabIndex        =   186
            Top             =   240
            Width           =   3255
            _cx             =   5741
            _cy             =   1296
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
      Begin VB.Frame Frame27 
         Caption         =   "[ Motivo da Liquidação do Pedido ]"
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
         Height          =   2295
         Left            =   -74880
         TabIndex        =   178
         Top             =   6000
         Width           =   12615
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
            TabIndex        =   181
            Top             =   240
            Width           =   9615
            Begin VB.TextBox txtCODMOTLIQ 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   10
               TabIndex        =   183
               Text            =   "txtCODMOTL"
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command6 
               Height          =   315
               Left            =   1200
               Picture         =   "frmCADPEDVENDA.frx":3129
               Style           =   1  'Graphical
               TabIndex        =   182
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
               TabIndex        =   184
               Top             =   240
               Width           =   7935
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
            TabIndex        =   179
            Top             =   840
            Width           =   9615
            Begin VB.TextBox txtOBS_MotLiq 
               Appearance      =   0  'Flat
               Height          =   855
               Left            =   120
               MaxLength       =   500
               MultiLine       =   -1  'True
               TabIndex        =   180
               Text            =   "frmCADPEDVENDA.frx":322B
               Top             =   240
               Width           =   9375
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
         Left            =   -74880
         TabIndex        =   176
         Top             =   3360
         Width           =   12615
         Begin VSFlex8LCtl.VSFlexGrid grdLogPed 
            Height          =   2295
            Left            =   120
            TabIndex        =   177
            Top             =   240
            Width           =   12375
            _cx             =   21828
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
      Begin VB.Frame Frame11 
         Caption         =   "[ Confirmação de Faturamento ]"
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
         Height          =   1455
         Left            =   -74880
         TabIndex        =   161
         Top             =   5280
         Visible         =   0   'False
         Width           =   6615
         Begin VSFlex8LCtl.VSFlexGrid grdConfFat 
            Height          =   1095
            Left            =   120
            TabIndex        =   162
            Top             =   240
            Width           =   6375
            _cx             =   11245
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
      Begin VB.Frame Frame14 
         Caption         =   "[ Resumo - Pedido ]"
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
         Height          =   1455
         Left            =   -68160
         TabIndex        =   154
         Top             =   5280
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Label Label37 
            Caption         =   "Tot.Geral Itens"
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
            TabIndex        =   160
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "Tot.Geral Faturado"
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
            TabIndex        =   159
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label37 
            Caption         =   "Saldo do Pedido"
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
            Left            =   120
            TabIndex        =   158
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblTotGer 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TotGeralIten"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   157
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblTotGer 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TotGeraFat"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   156
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblTotGer 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TotSaldoProd"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   155
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "[ Faturamento ]"
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
         Height          =   2655
         Left            =   8880
         TabIndex        =   152
         Top             =   6540
         Width           =   8535
         Begin VSFlex8LCtl.VSFlexGrid grdOrdFat 
            Height          =   2295
            Left            =   120
            TabIndex        =   153
            Top             =   240
            Width           =   8295
            _cx             =   14631
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
      Begin VB.Frame Frame21 
         Caption         =   "[ Reprovação do Pedido ]"
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
         Height          =   2895
         Left            =   -74880
         TabIndex        =   151
         Top             =   420
         Width           =   12615
         Begin VB.TextBox txtOBS 
            Appearance      =   0  'Flat
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   173
            Text            =   "frmCADPEDVENDA.frx":3239
            Top             =   2040
            Width           =   9375
         End
         Begin VB.CommandButton Command11 
            Height          =   300
            Left            =   9360
            Picture         =   "frmCADPEDVENDA.frx":3240
            Style           =   1  'Graphical
            TabIndex        =   172
            ToolTipText     =   "Inclui uma nova linha na Gride"
            Top             =   240
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid grdTIPREPROV 
            Height          =   1695
            Left            =   120
            TabIndex        =   171
            Top             =   240
            Width           =   9135
            _cx             =   16113
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
      Begin VB.Frame Frame9 
         Caption         =   "[ Totais ]"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   123
         Top             =   3420
         Width           =   13455
         Begin VB.TextBox txtFRETE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5400
            TabIndex        =   134
            Text            =   "txtFRETE"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtALIQICMS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5400
            TabIndex        =   133
            Text            =   "txtALIQICMS"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtOutrDesp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   132
            Text            =   "txtOutrDesp"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtVLDESCTOTOT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9120
            TabIndex        =   131
            Text            =   "txtVLDESCTOTOT"
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtPDESCTOTAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5400
            TabIndex        =   130
            Text            =   "txtPDESCTOTAL"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Frame Frame16 
            Caption         =   "[ Epecial ]"
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
            Height          =   495
            Left            =   10920
            TabIndex        =   127
            Top             =   720
            Visible         =   0   'False
            Width           =   1815
            Begin VB.OptionButton optESPECIAL 
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
               Left            =   960
               TabIndex        =   129
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optESPECIAL 
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
               TabIndex        =   128
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "[ Para Estoque ]"
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
            Height          =   495
            Left            =   10920
            TabIndex        =   124
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
            Begin VB.OptionButton optPARAESTOQUE 
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
               Left            =   960
               TabIndex        =   126
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optPARAESTOQUE 
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
               TabIndex        =   125
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Vl. Desconto nos Itens"
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
            TabIndex        =   150
            Top             =   960
            Width           =   1950
         End
         Begin VB.Label lblVLIPI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLIPI"
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
            Height          =   285
            Left            =   9120
            TabIndex        =   149
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblVLTOTAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLTOTAL"
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
            Height          =   285
            Left            =   2160
            TabIndex        =   148
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Total do Pedido"
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
            Left            =   120
            TabIndex        =   147
            Top             =   1320
            Width           =   1365
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor do IPI"
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
            Left            =   7560
            TabIndex        =   146
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
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
            Left            =   4200
            TabIndex        =   145
            Top             =   630
            Width           =   450
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Outras despesas"
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
            Left            =   120
            TabIndex        =   144
            Top             =   645
            Width           =   1425
         End
         Begin VB.Label lblVLICMS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLICMS"
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
            Height          =   285
            Left            =   9120
            TabIndex        =   143
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Valor do ICMS"
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
            Left            =   7560
            TabIndex        =   142
            Top             =   285
            Width           =   1230
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Aliq ICMS"
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
            Left            =   4200
            TabIndex        =   141
            Top             =   285
            Width           =   840
         End
         Begin VB.Label lblBASICMS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblBASICMS"
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
            Height          =   285
            Left            =   2160
            TabIndex        =   140
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Base Calculo ICMS"
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
            TabIndex        =   139
            Top             =   285
            Width           =   1635
         End
         Begin VB.Label lblVLDESCONTO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLDESCONTO"
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
            Height          =   285
            Left            =   2160
            TabIndex        =   138
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Vl.Desconto"
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
            Left            =   7560
            TabIndex        =   137
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Desconto%:"
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
            Left            =   4200
            TabIndex        =   136
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label Label38 
            Caption         =   "Label38"
            ForeColor       =   &H0080FF80&
            Height          =   495
            Left            =   6360
            TabIndex        =   135
            Top             =   2640
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "[ Local de Cobrança ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74880
         TabIndex        =   108
         Top             =   1920
         Width           =   9135
         Begin VB.TextBox txtFAXCOBR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            MaxLength       =   30
            TabIndex        =   115
            Text            =   "txtFAXCOBR"
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox txtTELCOBR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   114
            Text            =   "txtTELCOBR"
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtENDCOBR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   113
            Text            =   "txtENDCOBR"
            Top             =   240
            Width           =   4095
         End
         Begin VB.TextBox txtBAICOBR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   112
            Text            =   "txtBAICOBR"
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtCIDCOBR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   111
            Text            =   "txtCIDCOBR"
            Top             =   600
            Width           =   2535
         End
         Begin VB.ComboBox cboESTCOBR 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4560
            TabIndex        =   110
            Text            =   "cboESTCOBR"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtCEPCOBR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6480
            MaxLength       =   9
            TabIndex        =   109
            Text            =   "txtCEPCOB"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
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
            Left            =   4080
            TabIndex        =   122
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label25 
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   121
            Top             =   960
            Width           =   825
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   885
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5685
            TabIndex        =   119
            Top             =   240
            Width           =   570
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   360
            TabIndex        =   118
            Top             =   600
            Width           =   660
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   330
            Left            =   3780
            TabIndex        =   117
            Top             =   660
            Width           =   660
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   5880
            TabIndex        =   116
            Top             =   660
            Width           =   435
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "[ Local de Entrega ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74880
         TabIndex        =   93
         Top             =   420
         Width           =   9135
         Begin VB.TextBox txtFAXENTRE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            MaxLength       =   30
            TabIndex        =   100
            Text            =   "txtFAXENTRE"
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtTELENTR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   99
            Text            =   "txtTELENTR"
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtENDENTR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   98
            Text            =   "txtENDENTR"
            Top             =   360
            Width           =   4095
         End
         Begin VB.TextBox txtBAIENTR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6480
            MaxLength       =   20
            TabIndex        =   97
            Text            =   "txtBAIENTR"
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtCIDENTR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   96
            Text            =   "txtCIDENTR"
            Top             =   720
            Width           =   2535
         End
         Begin VB.ComboBox cboESTENTR 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4560
            TabIndex        =   95
            Text            =   "cboESTENTR"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtCEPENTR 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6480
            MaxLength       =   9
            TabIndex        =   94
            Text            =   "txtCEPENTR"
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
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
            Left            =   4080
            TabIndex        =   107
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label25 
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   106
            Top             =   1080
            Width           =   825
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   885
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5685
            TabIndex        =   104
            Top             =   360
            Width           =   570
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   103
            Top             =   720
            Width           =   660
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3780
            TabIndex        =   102
            Top             =   780
            Width           =   660
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   5880
            TabIndex        =   101
            Top             =   780
            Width           =   435
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "[ Observação do Iten ]"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   87
         Top             =   8100
         Width           =   5055
         Begin VB.TextBox txtOBSROT 
            Appearance      =   0  'Flat
            Height          =   735
            Left            =   120
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Text            =   "frmCADPEDVENDA.frx":338A
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   120
         TabIndex        =   82
         Top             =   360
         Width           =   17295
         Begin MSMask.MaskEdBox mskDATAPED 
            Height          =   285
            Left            =   3480
            TabIndex        =   1
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
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
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   150
            Width           =   660
         End
         Begin VB.Label lblCODIGO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCODIGO"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1560
            TabIndex        =   0
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   85
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   4920
            TabIndex        =   84
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblSTATUS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblSTATUS"
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
            Height          =   255
            Left            =   5670
            TabIndex        =   83
            Top             =   120
            Width           =   4515
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   120
         TabIndex        =   65
         Top             =   780
         Width           =   17295
         Begin VB.TextBox txtCONTATO 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12480
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "txtCONTATO"
            Top             =   420
            Width           =   4695
         End
         Begin VB.TextBox txtDEPARTAMENTO 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12480
            MaxLength       =   30
            TabIndex        =   12
            Text            =   "txtDEPARTAMENTO"
            Top             =   720
            Width           =   4695
         End
         Begin VB.TextBox txtEMAIL 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12480
            MaxLength       =   30
            TabIndex        =   13
            Text            =   "txtEMAIL"
            Top             =   1020
            Width           =   4695
         End
         Begin VB.TextBox txtORDCOMPCLI 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12480
            MaxLength       =   10
            TabIndex        =   88
            Text            =   "txtORDCOMPCLI"
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton cmdCondPgto 
            Height          =   315
            Left            =   2760
            Picture         =   "frmCADPEDVENDA.frx":3394
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   1020
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2760
            Picture         =   "frmCADPEDVENDA.frx":3496
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtCodCondPgto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   5
            Text            =   "txtCodCondPgto"
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtCIDCLIE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "txtCIDCLIE"
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2760
            Picture         =   "frmCADPEDVENDA.frx":3598
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtCODVEND 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txtCODVEND"
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2760
            Picture         =   "frmCADPEDVENDA.frx":369A
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   420
            Width           =   375
         End
         Begin VB.TextBox txtTIPPED 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "txtTIPPED"
            Top             =   420
            Width           =   1215
         End
         Begin VB.CheckBox chkVerificado 
            Caption         =   "Conferido"
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
            Left            =   9960
            TabIndex        =   67
            Top             =   1340
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command10 
            Height          =   315
            Left            =   2760
            Picture         =   "frmCADPEDVENDA.frx":379C
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   1340
            Width           =   375
         End
         Begin VB.TextBox txtCODTRANSP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   6
            Text            =   "txtCODTRANSP"
            Top             =   1340
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Comprador"
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
            Left            =   9960
            TabIndex        =   92
            Top             =   420
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
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
            Left            =   9960
            TabIndex        =   91
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail"
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
            Index           =   7
            Left            =   9960
            TabIndex        =   90
            Top             =   1020
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Ordem de Compra do Cliente"
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
            Left            =   9960
            TabIndex        =   89
            Top             =   120
            Width           =   2430
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
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
            TabIndex        =   81
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cond. Pagto:"
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
            TabIndex        =   80
            Top             =   1020
            Width           =   1125
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
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
            Left            =   120
            TabIndex        =   79
            Top             =   165
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Pedido:"
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
            TabIndex        =   78
            Top             =   420
            Width           =   1095
         End
         Begin VB.Label lblDescVendedor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescVendedor"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3120
            TabIndex        =   77
            Top             =   120
            Width           =   6735
         End
         Begin VB.Label lblDescTpPed 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescTpPed"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3120
            TabIndex        =   76
            Top             =   420
            Width           =   6735
         End
         Begin VB.Label lblDescCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescCliente"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3120
            TabIndex        =   75
            Top             =   720
            Width           =   6735
         End
         Begin VB.Label lblDescCondPgto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescCondPgto"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3120
            TabIndex        =   74
            Top             =   1020
            Width           =   6735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Transportdora"
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
            TabIndex        =   73
            Top             =   1340
            Width           =   1200
         End
         Begin VB.Label lblDescTransp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescTransp"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3120
            TabIndex        =   72
            Top             =   1340
            Width           =   6735
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "[ Observação 1 ]"
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
         Height          =   855
         Left            =   120
         TabIndex        =   64
         Top             =   2460
         Width           =   8055
         Begin VB.TextBox txtOBSERVACAO 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Text            =   "frmCADPEDVENDA.frx":389E
            Top             =   240
            Width           =   7815
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "[ Observação 2 ]"
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
         Height          =   855
         Left            =   8160
         TabIndex        =   63
         Top             =   2460
         Width           =   9255
         Begin VB.TextBox txtOBS2 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Text            =   "frmCADPEDVENDA.frx":38AE
            Top             =   240
            Width           =   9015
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "[Itens do Pedido ]"
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
         Height          =   2295
         Left            =   120
         TabIndex        =   48
         Top             =   3300
         Width           =   17295
         Begin VB.ComboBox cboFechTPFR 
            Height          =   315
            ItemData        =   "frmCADPEDVENDA.frx":38B6
            Left            =   6000
            List            =   "frmCADPEDVENDA.frx":38B8
            TabIndex        =   49
            Text            =   "cboFechTPFR"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Frame Frame13 
            BorderStyle     =   0  'None
            Caption         =   "Frame13"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1950
            Width           =   5415
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "B.Manual"
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
               Left            =   4320
               TabIndex        =   175
               Top             =   0
               Width           =   810
            End
            Begin VB.Label Label3 
               BackColor       =   &H000080FF&
               Caption         =   "Label31"
               ForeColor       =   &H000080FF&
               Height          =   255
               Left            =   3960
               TabIndex        =   174
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label31 
               BackColor       =   &H008080FF&
               Caption         =   "Label31"
               ForeColor       =   &H008080FF&
               Height          =   255
               Left            =   0
               TabIndex        =   59
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label32 
               BackColor       =   &H0080FFFF&
               Caption         =   "Label31"
               ForeColor       =   &H0080FFFF&
               Height          =   255
               Left            =   1440
               TabIndex        =   58
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label33 
               BackColor       =   &H0080FF80&
               Caption         =   "Label31"
               ForeColor       =   &H0080FF80&
               Height          =   255
               Left            =   2760
               TabIndex        =   57
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Em Aberto"
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
               Left            =   360
               TabIndex        =   56
               Top             =   0
               Width           =   885
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "B.Parcial"
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
               Left            =   1800
               TabIndex        =   55
               Top             =   0
               Width           =   780
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "B.Total"
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
               TabIndex        =   54
               Top             =   0
               Width           =   630
            End
         End
         Begin VB.CommandButton Command26 
            Height          =   300
            Left            =   16920
            Picture         =   "frmCADPEDVENDA.frx":38BA
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Exclui a linha da Gride Selecionada"
            Top             =   600
            Width           =   300
         End
         Begin VB.CommandButton Command27 
            Height          =   300
            Left            =   16920
            Picture         =   "frmCADPEDVENDA.frx":3A04
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Inclui uma nova linha na Gride"
            Top             =   240
            Width           =   300
         End
         Begin VB.ComboBox cboQtdePorPalhet 
            Height          =   315
            ItemData        =   "frmCADPEDVENDA.frx":3B4E
            Left            =   6000
            List            =   "frmCADPEDVENDA.frx":3B50
            TabIndex        =   50
            Text            =   "cboQtdePorPalhet"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VSFlex8LCtl.VSFlexGrid grdProduto 
            Height          =   1695
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   16695
            _cx             =   29448
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
         Begin VB.Label Label30 
            Caption         =   "Já faturado"
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
            Left            =   7800
            TabIndex        =   166
            Top             =   1950
            Width           =   1095
         End
         Begin VB.Label lblSaldoJaFat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblSaldoJaFat"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9000
            TabIndex        =   165
            Top             =   1950
            Width           =   1575
         End
         Begin VB.Label Label29 
            Caption         =   "Saldo do Rótulo"
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
            Left            =   10800
            TabIndex        =   164
            Top             =   1950
            Width           =   1455
         End
         Begin VB.Label lblSaldRot 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblSaldRot"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   12360
            TabIndex        =   163
            Top             =   1950
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vl. Total"
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
            Index           =   12
            Left            =   14280
            TabIndex        =   62
            Top             =   1950
            Width           =   735
         End
         Begin VB.Label lblTotalItens 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblTotalItens"
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
            Height          =   285
            Left            =   15120
            TabIndex        =   61
            Top             =   1950
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "[Programação de Entregas ]"
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
         Left            =   120
         TabIndex        =   44
         Top             =   5580
         Width           =   8655
         Begin MSComCtl2.MonthView MonthView1 
            Height          =   2310
            Left            =   5160
            TabIndex        =   187
            Top             =   120
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   4075
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            BorderStyle     =   1
            Appearance      =   0
            StartOfWeek     =   126091265
            CurrentDate     =   42968
         End
         Begin VB.CommandButton Command5 
            Height          =   300
            Left            =   4680
            Picture         =   "frmCADPEDVENDA.frx":3B52
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Inclui uma nova linha na Gride"
            Top             =   240
            Width           =   300
         End
         Begin VB.CommandButton Command4 
            Height          =   300
            Left            =   4680
            Picture         =   "frmCADPEDVENDA.frx":3C9C
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Exclui a linha da Gride Selecionada"
            Top             =   600
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid grdProgEntrega 
            Height          =   1335
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   4455
            _cx             =   7858
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
      Begin VB.Frame Frame12 
         Height          =   975
         Left            =   8880
         TabIndex        =   25
         Top             =   5580
         Width           =   7335
         Begin VB.Label Label4 
            Caption         =   "Fech."
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
            TabIndex        =   43
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Corpo"
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
            TabIndex        =   42
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Tampa"
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
            TabIndex        =   41
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Fundo"
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
            Left            =   2640
            TabIndex        =   40
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Argola"
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
            Left            =   2640
            TabIndex        =   39
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Fech.(T/F)"
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
            Left            =   2640
            TabIndex        =   38
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "Alt.Filme"
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
            Left            =   5640
            TabIndex        =   37
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "Fot.Novo"
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
            Left            =   5640
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label28 
            Caption         =   "Repetição"
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
            Left            =   5640
            TabIndex        =   35
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblDescFecham 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescFecham"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   720
            TabIndex        =   34
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lblDescCorpo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescCorpo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   720
            TabIndex        =   33
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblDescTampa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescTampa"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   720
            TabIndex        =   32
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblDescFundo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescFundo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3600
            TabIndex        =   31
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lblDescArgola 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescArgola"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3600
            TabIndex        =   30
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblDescFechTPFURO 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescFechTPFURO"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3600
            TabIndex        =   29
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblDescAltFilme 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescAltFilme"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6600
            TabIndex        =   28
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblFotNovo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblFotNovo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6600
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblDescRepet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescRepet"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6600
            TabIndex        =   26
            Top             =   600
            Width           =   615
         End
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
         Height          =   855
         Left            =   16320
         Picture         =   "frmCADPEDVENDA.frx":3DE6
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Liberação das alterações"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCADPEDVENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho             As String
Public Linha                As Variant
Public cTipOper             As String
Public iCodigo              As Long
Public FILIAL               As Integer
Public strACESSO            As String
Public strMODPAI            As String
Public strUSUARIO           As String
Public lngCodVendedor       As Long
Public lngCodUsuario        As Long
Public intFILIALPED         As Integer
Public boolSomenteCons      As Boolean
Public strVERSAO            As String
Public strNOMCOMP           As String

Dim objBLBFunc              As New clsFuncoes
Dim objCADPEDVENDA          As New clsCADPEDVENDA
Dim objPESQPADRAO           As Object

Dim strDTENTREGAFOT         As String
Dim lngQTDDIASLDTIMF        As Long

Dim arrItensPedido          As Variant
Dim arrPRODUTOS             As Variant
Dim arrDIASCOTAS            As Variant

Dim arrPRODCOPIA            As Variant
Dim arrTIPREPROVA           As Variant
Dim arrENTREGAS             As Variant
Dim curTOTPEDFRETE          As Currency
Dim strFINANCEIRO           As String
Dim strSTATUSPED            As String
Dim booAltorizado           As Boolean
Dim boolALteradoITEN        As Boolean
Dim lngIDPROGENTREGA        As Long      '' Criação de um Código interno da Programação de Entrega
Dim strNOMFILIAL            As String

'' ========================================================================================
Const conCOL_SonProd_IdProduto                  As Integer = 0
Const conCOL_SonProd_Codigo                     As Integer = 1
Const conCOL_SonProd_PesqProd                   As Integer = 2
Const conCOL_SonProd_DescProd                   As Integer = 3
Const conCOL_SonProd_QtdProd                    As Integer = 4
Const conCOL_SonProd_VlUniProd                  As Integer = 5
Const conCOL_SonProd_PorcDesc                   As Integer = 6
Const conCOL_SonProd_PorcIPI                    As Integer = 7
Const conCOL_SonProd_VlTotal                    As Integer = 8
Const conCOL_SonProd_VlDesc                     As Integer = 9
Const conCOL_SonProd_VlIPI                      As Integer = 10
Const conCOL_SonProd_VlItens                    As Integer = 11
Const conCOL_SonProd_Fechamento                 As Integer = 12
Const conCOL_SonProd_Corpo                      As Integer = 13
Const conCOL_SonProd_Tampa                      As Integer = 14
Const conCOL_SonProd_FornTampa                  As Integer = 15
Const conCOL_SonProd_PesqForn                   As Integer = 16
Const conCOL_SonProd_Fundo                      As Integer = 17
Const conCOL_SonProd_Argola                     As Integer = 18
Const conCOL_SonProd_FechTpFr                   As Integer = 19
Const conCOL_SonProd_Desenho                    As Integer = 20
Const conCOL_SonProd_AltFilme                   As Integer = 21
Const conCOL_SonProd_FotNovo                    As Integer = 22
Const conCOL_SonProd_Repeticao                  As Integer = 23
Const conCOL_SonProd_CodLinProd                 As Integer = 24
Const conCOL_SonProd_OBSOP                      As Integer = 25
Const conCOL_SonProd_IDBKP                      As Integer = 26
Const conCOL_SonProd_PRECOBKP                   As Integer = 27
Const conCOL_SonProd_QTDBKP                     As Integer = 28
Const conCOL_SonProd_FechTpFrBKP                As Integer = 29
Const conCOL_SonProd_AltFilmeBKP                As Integer = 30
Const conCOL_SonProd_FotNovoBKP                 As Integer = 31
Const conCOL_SonProd_RepeticaoBKP               As Integer = 32
Const conCOL_SonProd_Action2Do                  As Integer = 33
Const conCOL_SonProd_TemOP                      As Integer = 34
Const conCOL_SonProd_StatusProd                 As Integer = 35
Const conCOL_SonProd_GrpPlanMestre              As Integer = 36
Const conCOL_SonProd_CodCapacidade              As Integer = 37
Const conCOL_SonProd_NECKIN                     As Integer = 38
Const conCOL_SonProd_HOMOLOGADO                 As Integer = 39
Const conCOL_SonProd_QTDELATASPALLETS           As Integer = 40
Const conCOL_SonProd_PALLETS                    As Integer = 41
Const conCOL_SonProd_Conferido                  As Integer = 42
Const conCOL_SonProd_PalhetPadrao               As Integer = 43
Const conCOL_SonProd_OS_Artes                   As Integer = 44
Const conCOL_SonProd_FormatString               As String = "=Cod|Código|...|Descrição|Qtde.|Vl. Unit.|% Desc.|% IPI|Vl. Total|Vl.Desc|Vl.IPI.|Vl.Itens|Fechamento|Corpo|Tampa|Fornecedor|...|Fundo|Argola|Fech.(TP/FURO)|...|Alt.Filme|Fot.Novo|Repetição|CodLinProd|OBSOP|IDBKP|PRECOBKP|QTDBKP|FechTpFrBKP|AltFilmeBKP|FotNovoBKP|RepeticaoBKP|Action2Do|TemOP|StatusProd|GrpPMestre|CodCapacidade|NECKIN|HOMOLOGADO|Latas por Palhets|Qrde de Palhets|Conferido|PalhetPadrao|..."
Const conColumnsIn_SonProd                      As Integer = 45

'' ========================================================================================
Const conCOL_SonProgEntr_IdProduto              As Integer = 0
Const conCOL_SonProgEntr_QtdProd                As Integer = 1
Const conCOL_SonProgEntr_DataEntrega            As Integer = 2
Const conCOL_SonProgEntr_Action2Do              As Integer = 3
Const conCOL_SonProgEntr_OBSOP                  As Integer = 4
Const conCOL_SonProgEntr_CodOP                  As Integer = 5
Const conCOL_SonProgEntr_StatusOP               As Integer = 6
Const conCOL_SonProgEntr_FechTpFr               As Integer = 7
Const conCOL_SonProgEntr_INDICE                 As Integer = 8
Const conCOL_SonProgEntr_INDICEBKP              As Integer = 9
Const conCOL_SonProgEntr_DataEntregaBKP         As Integer = 10
Const conCOL_SonProgEntr_IDINTERNO              As Integer = 11
Const conCOL_SonProgEntr_DescStatusOP           As Integer = 12
Const conCOL_SonProgEntr_GrpPlanMestre          As Integer = 13
Const conCOL_SonProgEntr_PegaPlanMestre         As Integer = 14
Const conCOL_SonProgEntr_QTDENOPALHET           As Integer = 15
Const conCOL_SonProgEntr_PALHET                 As Integer = 16
Const conCOL_SonProgEntr_Action2DoDtEntrega     As Integer = 17
Const conCOL_SonProgEntr_DataPrevLito           As Integer = 18
Const conCOL_SonProgEntr_DataPrevProd           As Integer = 19
Const conCOL_SonProgEntr_CODIDPROG              As Integer = 20
Const conCOL_SonProgEntr_CODSTATAPONT           As Integer = 21
Const conCOL_SonProgEntr_DESCSTATUSAPONT        As Integer = 22
Const conCOL_SonProgEntr_FormatString           As String = "=Cod|Quant.OP|Dt.Entrega|Action2Do|OBSOP|Cod.OP|Status|FechTF|INDICE|INDICEBKP|DataEntregaBKP|IDINTERNO|Status.OP|GrPMestre|...|Qtde por Palhet|Qtde de Palhet|Action2DoDtEntrega|Previsão de Lito|Previsão de Montagem|CODIDPROG|CODSTATAPONT|Status Apontamento"
Const conColumnsIn_SonProgEntr                  As Integer = 23

'' ========================================================================================
Const conCOL_SonOrdemFat_IdProduto              As Integer = 0
Const conCOL_SonOrdemFat_VlUnit                 As Integer = 1
Const conCOL_SonOrdemFat_QtdOP                  As Integer = 2
Const conCOL_SonOrdemFat_QtdProd                As Integer = 3
Const conCOL_SonOrdemFat_Saldo                  As Integer = 4
Const conCOL_SonOrdemFat_CodOrdem               As Integer = 5
Const conCOL_SonOrdemFat_DatOrdem               As Integer = 6
Const conCOL_SonOrdemFat_Action2Do              As Integer = 7
Const conCOL_SonOrdemFat_CodOP                  As Integer = 8
Const conCOL_SonOrdemFat_SaldoPed               As Integer = 9
Const conCOL_SonOrdemFat_NF                     As Integer = 10
Const conCOL_SonOrdemFat_DataNF                 As Integer = 11
Const conCOL_SonOrdemFat_CodORDFAT              As Integer = 12
Const conCOL_SonOrdemFat_FormatString           As String = "=IDProd|Vl.Unit|Qtd.OP|Qtd.Fat|Saldo.OP|Cod. Ord. Fat.|Data Ord. Fat|Action2Do|Cod.OP|Saldo.Item|Nota Fiscal|Data Fat.|Ord. Fat"
Const conColumnsIn_SonOrdemFat                  As Integer = 13

'' ========================================================================================
Const conCOL_SonConfFat_IdProduto              As Integer = 0
Const conCOL_SonConfFat_CodOrdem               As Integer = 1
Const conCOL_SonConfFat_CodConf                As Integer = 2
Const conCOL_SonConfFat_QtdProd                As Integer = 3
Const conCOL_SonConfFat_VlUnit                 As Integer = 4
Const conCOL_SonConfFat_NF                     As Integer = 5
Const conCOL_SonConfFat_FormatString           As String = "=IDProd|Cod.Ordem|Cod.Conf|Qtd.Fat|V.Unit|Cod.NF|Data.Conf"
Const conColumnsIn_SonConfFat                  As Integer = 6


'' ========================================================================================
Const conCOL_SonLogPed_Data                     As Integer = 0
Const conCOL_SonLogPed_Hora                     As Integer = 1
Const conCOL_SonLogPed_CodUsuario               As Integer = 2
Const conCOL_SonLogPed_Usuario                  As Integer = 3
Const conCOL_SonLogPed_CodAcao                  As Integer = 4
Const conCOL_SonLogPed_Acao                     As Integer = 5
Const conCOL_SonLogPed_Tipo                     As Integer = 6
Const conCOL_SonLogPed_FormatString             As String = "=Data|Hora|CodUsuario|Usuário|CodAcao|Ação|Tipo"
Const conColumnsIn_SonLogPed                    As Integer = 7

'' ========================================================================================
Const conCOL_SonRep_Codigo                     As Integer = 0
Const conCOL_SonRep_Pesq                       As Integer = 1
Const conCOL_SonRep_Desc                       As Integer = 2
Const conCOL_SonRep_FormatString               As String = "=Código|...|Descrição"
Const conColumnsIn_SonRep                      As Integer = 3

'' ========================================================================================
Const conCOL_SonProducao_IDPROG                     As Integer = 0
Const conCOL_SonProducao_IDITERNO                   As Integer = 1
Const conCOL_SonProducao_CODOP                      As Integer = 2
Const conCOL_SonProducao_IDOP                       As Integer = 3
Const conCOL_SonProducao_IDPROD                     As Integer = 4
Const conCOL_SonProducao_DTPROG                     As Integer = 5
Const conCOL_SonProducao_QTDPROG                    As Integer = 6
Const conCOL_SonProducao_CODSTATUS                  As Integer = 7
Const conCOL_SonProducao_DESCSTATUS                 As Integer = 8
Const conCOL_SonProducao_FormatString               As String = "=IDPROG|IDINTERNO|CODOP|IDOP|IDPROD|Prev.Mont.|Qtde.Progr.|CODSTATUS|Desc.Status"
Const conColumnsIn_SonProducao                      As Integer = 9

Private Sub cboFechTPFR_LostFocus()
    ' hide date picker when user is done with it
    cboFechTPFR.Visible = False
End Sub

Private Sub cboFechTPFR_Validate(Cancel As Boolean)

    grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_FechTpFr) = Empty
    grdProduto.Cell(flexcpData, grdProduto.Row, conCOL_SonProd_FechTpFr) = Empty
    
    Dim i As Integer

    With grdProduto
        For i = 1 To (grdProgEntrega.Rows - 1)
            If grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) = grdProduto.Cell(flexcpText, .Row, conCOL_SonProd_IdProduto) Then
               If cboFechTPFR.ListIndex > -1 Then
                  grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_FechTpFr) = cboFechTPFR.ItemData(cboFechTPFR.ListIndex)
               Else
                  grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_FechTpFr) = ""
               End If
            End If
        Next i
        Call objBLBFunc.TrocaAction2Do(grdProduto, .Row, conCOL_SonProd_Action2Do, .Cell(flexcpTextDisplay, .Row, conCOL_SonProd_FechTpFr), cboFechTPFR.Text)
        Call MudaActio2DoFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, .Cell(flexcpText, .Row, conCOL_SonProd_IdProduto))
    End With

    If Len(Trim(cboFechTPFR.Text)) > 0 Then
    
        With grdProduto
            grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_FechTpFr) = cboFechTPFR.Text
            grdProduto.Cell(flexcpData, grdProduto.Row, conCOL_SonProd_FechTpFr) = cboFechTPFR.ItemData(cboFechTPFR.ListIndex)
        End With
    
    End If

End Sub

Private Sub cboQtdePorPalhet_LostFocus()
    
    ' hide date picker when user is done with it
    cboQtdePorPalhet.Visible = False

End Sub

Private Sub cboQtdePorPalhet_Validate(Cancel As Boolean)

    grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_QTDELATASPALLETS) = Empty
    
    Dim i As Integer

    With grdProduto
        For i = 1 To (grdProgEntrega.Rows - 1)
            If grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) = grdProduto.Cell(flexcpText, .Row, conCOL_SonProd_IdProduto) Then
               If cboQtdePorPalhet.ListIndex > -1 Then
                  grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_PALHET) = cboQtdePorPalhet.Text
               Else
                  grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_PALHET) = ""
               End If
            End If
        Next i
        Call objBLBFunc.TrocaAction2Do(grdProduto, .Row, conCOL_SonProd_Action2Do, .Cell(flexcpTextDisplay, .Row, conCOL_SonProd_QTDELATASPALLETS), cboQtdePorPalhet.Text)
        Call MudaActio2DoFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, .Cell(flexcpText, .Row, conCOL_SonProd_IdProduto))
    End With

    If Len(Trim(cboQtdePorPalhet.Text)) > 0 Then
    
        With grdProduto
            grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_QTDELATASPALLETS) = cboQtdePorPalhet.Text
        
            If Len(Trim(grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_QtdProd))) > 0 Then
                '' Depois voltar
                ''If ConferePalhets(grdProduto.Row, CLng(grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_QtdProd))) = False Then
                ''   Cancel = True
                ''   Exit Sub
                ''End If
            End If
        
        End With
    
    End If

End Sub

Private Sub cmdAltera_Click()
    
On Error GoTo cmdAltera_Click
    
    
    If objBLBFunc.ChecaAcesso2("A", strACESSO) = False Then Exit Sub
    
    If PermiteAltPedidoFatParc = False Then
       MsgBox "O usuário não tem permição para alterar o Pedido !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    
    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Pedido de Venda - [ ALTERAÇÃO ]"
    
    stCAMPOSVENDA.Tab = 0
    
    Frame3.Enabled = True
    Frame4.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True
   
    Frame13.Visible = True
    
    txtCIDCLIE.Enabled = False
    Command1.Enabled = False
    
    Call DesativasCampos
    
    cTipOper = "A"
    
    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes
    
    Call VisualizaBotoesLibAlteracao(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibComercial(objCADPEDVENDA.STATUS, cTipOper)
    
    Exit Sub
    
cmdAltera_Click:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdAltera_Click()", Me.Name, "cmdAltera_Click()", strCAMARQERRO)
    
End Sub

Private Sub cmdCancAlteracao_Click()

    cTipOper = "C"
    
    Call Consulta

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes

End Sub

Private Sub cmdCancLibCom_Click()

    cTipOper = "C"
    
    Call Consulta

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes

End Sub

Private Sub cmdCancLibFin_Click()

    cTipOper = "C"
    
    Call Consulta

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes

End Sub

Private Sub cmdCancLibFot_Click()

    cTipOper = "C"
    
    Call Consulta

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes

End Sub

Private Sub cmdCancLibPcotaPdata_Click()

    cTipOper = "C"
    
    Call Consulta

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes

End Sub

Private Sub cmdCondPgto_Click()

On Error GoTo Err_cmdCondPgto_Click

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADCONDPGTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Condição de Pagamento")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCodCondPgto.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADCONDPGTO", varRETORNO, lblDescCondPgto, "cmdCondPgto_Click()")
    If Len(Trim(lblDescCondPgto.Caption)) = 0 Then txtCodCondPgto.Text = ""
    
    txtCodCondPgto.SetFocus

    Exit Sub
    
Err_cmdCondPgto_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : cmdCondPgto_Click()", Me.Name, "cmdCondPgto_Click()", strCAMARQERRO)

End Sub

Private Sub cmdLibAlteracao_Click()

    cTipOper = "LS"
    
    Call Libera

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes

End Sub

Private Sub cmdLiberaCom_Click()

    cTipOper = "LN"
    
    Call Libera

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

End Sub

Private Sub cmdLiberaFinanceiro_Click()

    cTipOper = "LF"
    
    Call Libera

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

End Sub

Private Sub cmdLibFot_Click()

    cTipOper = "LV"
    
    Call Libera

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes

End Sub

Private Sub cmdLibPcotaPData_Click()
    
    cTipOper = "LC"
    
    Call Libera

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO

    If boolSomenteCons = True Then cmdAltera.Enabled = False

    Call AtivaDesativaBotoes

End Sub

Private Sub cmdLIBPDATAPCOTA_Click()
    If (grdProgEntrega.Rows - 1) = 0 Then Exit Sub
    If grdProgEntrega.Row = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & "Selecione uma Entrega !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    With grdProgEntrega
        If .Cell(flexcpText, .Row, conCOL_SonProgEntr_StatusOP) <> 0 Then
           .Cell(flexcpText, .Row, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update
           .Cell(flexcpText, .Row, conCOL_SonProgEntr_StatusOP) = 0
           .Cell(flexcpText, .Row, conCOL_SonProgEntr_DescStatusOP) = PegaStatusOP(CLng(.Cell(flexcpText, .Row, conCOL_SonProgEntr_StatusOP)))
        End If
    End With
End Sub

Private Sub CmdSalva_Click()
    
On Error GoTo err_grava

    Dim i                   As Integer
    Dim j                   As Integer
    Dim intRESP             As Integer
    Dim lngITENS            As Integer
    Dim sValor              As String
    Dim dtENTREGA           As Date
    Dim boolTemPCotaPData   As Boolean
    Dim intAction2Do        As Integer
    Dim strDATA             As String
    Dim dtDTLIDTIME         As String
    Dim intStatusOP         As Integer
    Dim strCODLIN           As String
    Dim boolTemPData        As Boolean
    Dim intLINHAPROD        As Integer
    Dim boolTemOP           As Boolean
    
    objCADPEDVENDA.CODUSUARIO = lngCodUsuario
    
    Call objBLBFunc.RemoveLinhaVazia(grdProduto, conCOL_SonProd_Codigo)
    Call objBLBFunc.RemoveLinhaVazia(grdProgEntrega, conCOL_SonProgEntr_IdProduto)
    Call objBLBFunc.RemoveLinhaVazia(grdProgEntrega, conCOL_SonProgEntr_QtdProd)
    Call objBLBFunc.RemoveLinhaVazia(grdTIPREPROV, conCOL_SonRep_Codigo)
    
    If Valida_Campos = False Then Exit Sub
    
    '' Nome do Computador
    objCADPEDVENDA.NOMCOMP = "Null"
    If Len(Trim(strNOMCOMP)) > 0 Then objCADPEDVENDA.NOMCOMP = "'" & strNOMCOMP & "'"
    
    '' Versão
    objCADPEDVENDA.VERSAO = "Null"
    If Len(Trim(strVERSAO)) > 0 Then objCADPEDVENDA.VERSAO = "'" & strVERSAO & "'"
    
    '' Verificando Crédito
    ''If cTipOper = "I" Then objCADPEDVENDA.STATUS = Verifica_Credito
    If cTipOper = "I" Then objCADPEDVENDA.STATUS = "B"
    If cTipOper = "A" Then objCADPEDVENDA.PRODCOPIA = arrPRODCOPIA
        
    If intFILIALPED = 0 Then        '' Nova Lata
        If cTipOper = "I" Then objCADPEDVENDA.CODPEDIDO = objCADPEDVENDA.Gera_Codigo(Trim(Me.Name & Format(intFILIALPED, "##00"))) & Year(Now)
    ElseIf intFILIALPED = 1 Then    ''Steel Row
        If cTipOper = "I" Then objCADPEDVENDA.CODPEDIDO = objCADPEDVENDA.Gera_Codigo(Trim(Me.Name & "_STEEL" & Format(intFILIALPED, "##00"))) & Year(Now)
    End If
    
    objCADPEDVENDA.DATAPED = CDate(mskDATAPED.Text)
    objCADPEDVENDA.CODCLIE = CInt(txtCIDCLIE.Text)
    objCADPEDVENDA.CODCONDPGTO = CInt(txtCodCondPgto.Text)
    objCADPEDVENDA.CODVEND = CInt(txtCODVEND.Text)
    objCADPEDVENDA.TIPPED = CInt(txtTIPPED.Text)
    
    objCADPEDVENDA.ENDENTR = txtENDENTR.Text
    objCADPEDVENDA.BAIENTR = txtBAIENTR.Text
    objCADPEDVENDA.CIDENTR = txtCIDENTR.Text
    If cboESTENTR.ListIndex > -1 Then objCADPEDVENDA.ESTENTREGA = cboESTENTR.ItemData(cboESTENTR.ListIndex)
    objCADPEDVENDA.CEPENTREGA = txtCEPENTR.Text
    objCADPEDVENDA.TELENTR = txtTELENTR.Text
    objCADPEDVENDA.FAXENTR = txtFAXENTRE.Text
    
    objCADPEDVENDA.ENDCOBRA = txtENDCOBR.Text
    objCADPEDVENDA.BAICOBRA = txtBAICOBR.Text
    objCADPEDVENDA.CIDCOBRA = txtCIDCOBR.Text
    If cboESTCOBR.ListIndex > -1 Then objCADPEDVENDA.ESTCOBRA = cboESTCOBR.ItemData(cboESTCOBR.ListIndex)
    objCADPEDVENDA.CEPCOBRA = txtCEPCOBR.Text
    objCADPEDVENDA.TELCOBRA = txtTELCOBR.Text
    objCADPEDVENDA.FAXCOBRA = txtFAXCOBR.Text
    
    objCADPEDVENDA.CODTRANSP = txtCODTRANSP.Text
    objCADPEDVENDA.ORDCOMPCLI = txtORDCOMPCLI.Text
    objCADPEDVENDA.CONTATO = txtCONTATO.Text
    objCADPEDVENDA.DEPARTAMENTO = txtDEPARTAMENTO.Text
    objCADPEDVENDA.EMAIL = txtEMAIL.Text
    
    objCADPEDVENDA.OBSERVACAO = Trim(Replace(Replace(txtOBSERVACAO.Text, "'", ""), ",", ""))
        
    objCADPEDVENDA.OBS2 = "'" & Trim(Replace(Replace(txtOBS2.Text, ",", ""), "'", "")) & "'"
        
    '' Pedido Especial
    If optESPECIAL(0).Value = True Then objCADPEDVENDA.ESPECIAL = 0
    If optESPECIAL(1).Value = True Then objCADPEDVENDA.ESPECIAL = 1
        
    '' Para Estoque
    If optPARAESTOQUE(0).Value = True Then objCADPEDVENDA.PARAESTOQUE = 0
    If optPARAESTOQUE(1).Value = True Then objCADPEDVENDA.PARAESTOQUE = 1
        
    '' Conferido
    objCADPEDVENDA.CONFERIDO = chkVerificado.Value
        
    '' ----------------------------------------
    '' Produtos
    objCADPEDVENDA.PRODUTOS = Empty
    objCADPEDVENDA.TOTALITENS = 0
    objCADPEDVENDA.QTDPEDATEND = 0
    objCADPEDVENDA.QTDITENSPROD = 0
    
    '' =====================================
    '' Liberando Fotolito
    '' Ao Liberar Fotolito joga para ser liberado Comewrcial
    If cTipOper = "LV" Then objCADPEDVENDA.STATUS = "B"
    '' =====================================
    
    '' =====================================
    '' Liberando Comercial
    '' Ao Liberar Comercial joga Para Liberar Financeiro
    If cTipOper = "LN" Then objCADPEDVENDA.STATUS = "N"
    '' =====================================
    
    '' =====================================
    '' Liberando Financeiro
    '' Ao Liberar Financeiro joga Para Liberados
    If cTipOper = "LF" Then objCADPEDVENDA.STATUS = "L"
    '' =====================================
    
    '' =====================================
    '' Bloquendo o Pedido
    If cTipOper = "D" Then objCADPEDVENDA.STATUS = "S"
    '' =====================================
    
    '' =====================================
    '' Liberar o Pedido Bloqueado
    If cTipOper = "LS" Then objCADPEDVENDA.STATUS = "B"
    '' =====================================
    
    '' =====================================
    '' Reprova Pedido
    If cTipOper = "R" Then objCADPEDVENDA.STATUS = "R"
    '' =====================================
    
    '' =====================================
    '' Liberar o Pedido P.Cota/P.Data
    If cTipOper = "LC" Then objCADPEDVENDA.STATUS = "L"
    '' =====================================
    
    '' =====================================
    '' Pedido Alterado
    If cTipOper = "A" Then
       If objCADPEDVENDA.STATUS = "S" Then objCADPEDVENDA.STATUS = "S"
       If (objCADPEDVENDA.STATUS = "C" Or objCADPEDVENDA.STATUS = "4") Then objCADPEDVENDA.STATUS = "C"
    End If
    '' =====================================
    
    Dim boolAquardLibFot As Boolean
    boolAquardLibFot = False
    boolTemOP = False
    
    
    '' Itens do Pedido
    With grdProduto
        arrItensPedido = Empty
        If (.Rows - 1) > 0 Then
            ReDim arrItensPedido(1 To (.Rows - 1), 1 To 28) As String
            lngITENS = 0
            For i = 1 To (.Rows - 1)
                arrItensPedido(i, 1) = .Cell(flexcpText, i, conCOL_SonProd_Codigo)
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_QtdProd))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, i, conCOL_SonProd_QtdProd), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrItensPedido(i, 2) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_VlUniProd))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, i, conCOL_SonProd_VlUniProd), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrItensPedido(i, 3) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_PorcDesc))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, i, conCOL_SonProd_PorcDesc), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrItensPedido(i, 4) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_PorcIPI))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, i, conCOL_SonProd_PorcIPI), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrItensPedido(i, 5) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_VlTotal))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, i, conCOL_SonProd_VlTotal), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrItensPedido(i, 6) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_VlIPI))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, i, conCOL_SonProd_VlIPI), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrItensPedido(i, 7) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_VlDesc))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, i, conCOL_SonProd_VlDesc), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrItensPedido(i, 8) = sValor
                
                arrItensPedido(i, 9) = .Cell(flexcpText, i, conCOL_SonProd_IdProduto)
                
                arrItensPedido(i, 10) = .Cell(flexcpText, i, conCOL_SonProd_AltFilme)
                arrItensPedido(i, 11) = .Cell(flexcpText, i, conCOL_SonProd_FotNovo)
                arrItensPedido(i, 12) = .Cell(flexcpText, i, conCOL_SonProd_Repeticao)
                arrItensPedido(i, 13) = .Cell(flexcpText, i, conCOL_SonProd_CodLinProd)
                
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_FornTampa))) > 0 Then
                    arrItensPedido(i, 14) = .Cell(flexcpText, i, conCOL_SonProd_FornTampa)
                Else
                    arrItensPedido(i, 14) = "Null"
                End If
                
                arrItensPedido(i, 15) = Trim(Replace(Replace(.Cell(flexcpText, i, conCOL_SonProd_OBSOP), vbCrLf, ""), vbTab, ""))
                arrItensPedido(i, 16) = .Cell(flexcpData, i, conCOL_SonProd_FechTpFr)
                arrItensPedido(i, 17) = .Cell(flexcpText, i, conCOL_SonProd_Action2Do)
                arrItensPedido(i, 18) = .Cell(flexcpText, i, conCOL_SonProd_TemOP)
                
                '' Palhets
                arrItensPedido(i, 19) = "Null"
                arrItensPedido(i, 20) = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_QTDELATASPALLETS))) > 0 Then arrItensPedido(i, 19) = .Cell(flexcpText, i, conCOL_SonProd_QTDELATASPALLETS)
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_PALLETS))) > 0 Then arrItensPedido(i, 20) = .Cell(flexcpText, i, conCOL_SonProd_PALLETS)
                '' -------------------------------
                
                '' Atualizando as OBS das Prog de Entrega
                For j = 1 To (grdProgEntrega.Rows - 1)
                    If Trim(grdProgEntrega.Cell(flexcpText, j, conCOL_SonProgEntr_IdProduto)) = Trim(.Cell(flexcpText, i, conCOL_SonProd_IdProduto)) Then
                       grdProgEntrega.Cell(flexcpText, j, conCOL_SonProgEntr_OBSOP) = Trim(Replace(Replace(.Cell(flexcpText, i, conCOL_SonProd_OBSOP), vbTab, ""), vbCrLf, ""))
                    End If
                Next j
                '' =====================================
                
                objCADPEDVENDA.QTDPEDATEND = objCADPEDVENDA.QTDPEDATEND + CCur(.Cell(flexcpText, i, conCOL_SonProd_QtdProd))
                objCADPEDVENDA.QTDITENSPROD = objCADPEDVENDA.QTDITENSPROD + CCur(.Cell(flexcpText, i, conCOL_SonProd_QtdProd))
                
                
                If .Cell(flexcpText, i, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_delete Then lngITENS = (lngITENS + 1)
            
                '' ========================
                '' Muda o Status do Pedido
                If (.Cell(flexcpText, i, conCOL_SonProd_Action2Do) = dacEnumUpdateAction_Insert Or _
                   .Cell(flexcpText, i, conCOL_SonProd_Action2Do) = dacEnumUpdateAction_update) Then
                    If .Cell(flexcpText, i, conCOL_SonProd_StatusProd) = 1 Then
                        If cTipOper = "I" Or cTipOper = "A" Then
                            If objCADPEDVENDA.STATUS <> "P" Then objCADPEDVENDA.STATUS = "B"
                        End If
                    ElseIf .Cell(flexcpText, i, conCOL_SonProd_StatusProd) = 2 Then
                        If cTipOper = "I" Or cTipOper = "A" Then
                           If objCADPEDVENDA.STATUS <> "S" Then objCADPEDVENDA.STATUS = "V"
                        End If
                    End If
                Else
                    If .Cell(flexcpText, i, conCOL_SonProd_Action2Do) = dacEnumUpdateAction_Ignore Then
                        If .Cell(flexcpText, i, conCOL_SonProd_StatusProd) = 1 Then
                            If cTipOper = "I" Or cTipOper = "A" Then
                                If objCADPEDVENDA.STATUS = "S" Then objCADPEDVENDA.STATUS = "S"
                                If objCADPEDVENDA.STATUS = "B" Then objCADPEDVENDA.STATUS = "B"
                            ElseIf cTipOper = "R" Then
                                objCADPEDVENDA.STATUS = "R"
                            End If
                        ElseIf .Cell(flexcpText, i, conCOL_SonProd_StatusProd) = 2 Then
                            objCADPEDVENDA.STATUS = "V"
                        End If
                    End If
                End If
            
                arrItensPedido(i, 21) = .Cell(flexcpText, i, conCOL_SonProd_Conferido)
                
                arrItensPedido(i, 22) = 0
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_PalhetPadrao))) > 0 Then arrItensPedido(i, 22) = .Cell(flexcpText, i, conCOL_SonProd_PalhetPadrao)
            
            
                arrItensPedido(i, 23) = "'" & Trim(.Cell(flexcpText, i, conCOL_SonProd_DescProd)) & "'"
            
                arrItensPedido(i, 24) = .Cell(flexcpText, i, conCOL_SonProd_Fechamento)
                arrItensPedido(i, 25) = .Cell(flexcpText, i, conCOL_SonProd_Corpo)
                arrItensPedido(i, 26) = .Cell(flexcpText, i, conCOL_SonProd_Tampa)
                arrItensPedido(i, 27) = .Cell(flexcpText, i, conCOL_SonProd_Fundo)
                arrItensPedido(i, 28) = .Cell(flexcpText, i, conCOL_SonProd_Argola)
                
                '' Se tiver OP
                If .Cell(flexcpText, i, conCOL_SonProd_TemOP) = "S" Then boolTemOP = True
            
            Next i
            objCADPEDVENDA.TOTALITENS = lngITENS
        End If
        objCADPEDVENDA.PRODUTOS = arrItensPedido
    End With
    
    '' ----------------------------------------
    
    '' ----------------------------------------
    '' Tipos de Reprovação
    arrTIPREPROVA = Empty
    With grdTIPREPROV
        If (.Rows - 1) > 0 Then
           ReDim arrTIPREPROVA(1 To (.Rows - 1)) As String
           For i = 1 To (.Rows - 1)
               arrTIPREPROVA(i) = .Cell(flexcpText, i, conCOL_SonRep_Codigo)
           Next i
        End If
    End With
    objCADPEDVENDA.TIPREPROVA = arrTIPREPROVA
    
       
    If Len(Trim(lblBASICMS.Caption)) > 0 Then objCADPEDVENDA.VALBASICMS = CCur(lblBASICMS.Caption)
    If Len(Trim(txtALIQICMS.Text)) > 0 Then objCADPEDVENDA.ALIQICMS = CCur(txtALIQICMS.Text)
    If Len(Trim(lblVLICMS.Caption)) > 0 Then objCADPEDVENDA.VLICMS = CCur(lblVLICMS.Caption)
    If Len(Trim(txtOutrDesp.Text)) > 0 Then objCADPEDVENDA.OUTRDESPESAS = CCur(txtOutrDesp.Text)
    If Len(Trim(txtFRETE.Text)) > 0 Then objCADPEDVENDA.VLFRETE = CCur(txtFRETE.Text)
    If Len(Trim(lblVLIPI.Caption)) > 0 Then objCADPEDVENDA.VLIPI = CCur(lblVLIPI.Caption)
    If Len(Trim(lblVLDESCONTO.Caption)) > 0 Then objCADPEDVENDA.VLDESCTO = CCur(lblVLDESCONTO.Caption)
    If Len(Trim(txtPDESCTOTAL.Text)) > 0 Then objCADPEDVENDA.VALDESC = CCur(txtPDESCTOTAL.Text)
    If Len(Trim(txtVLDESCTOTOT.Text)) > 0 Then objCADPEDVENDA.PORDESC = CCur(txtVLDESCTOTOT.Text)
    If Len(Trim(lblVLTOTAL.Caption)) > 0 Then objCADPEDVENDA.TOTORCTO = CCur(lblVLTOTAL.Caption)
    
    If cTipOper = "LF" Then
       objCADPEDVENDA.LIBERADO = objBLBFunc.Crypt(strUSUARIO)
       objCADPEDVENDA.DTHORALIB = CDate(Format(Now, "DD/MM/YYYY HH:MM:SS"))
       objCADPEDVENDA.OBSCOMERCIAL = txtOBS.Text
    End If
    If cTipOper = "R" Or cTipOper = "D" Then
       '' R - Reprovado / D - Bloqueia
       objCADPEDVENDA.LIBERADO = objBLBFunc.Crypt(strUSUARIO)
       objCADPEDVENDA.DTHORALIB = CDate(Format(Now, "DD/MM/YYYY HH:MM:SS"))
       objCADPEDVENDA.OBSCOMERCIAL = txtOBS.Text
    End If
    
    If cTipOper = "I" Then objCADPEDVENDA.QTDPEDATEND = objCADPEDVENDA.QTDPEDATEND + objCADPEDVENDA.PegaQtdTotItPedido(objCADPEDVENDA.CODCOTA)
    objCADPEDVENDA.QTDTOTCOTA = objCADPEDVENDA.PegaQtdTotItCota(objCADPEDVENDA.CODCOTA)
    
    '' =====================================
    '' Quantidades/Data de Entrega
    arrENTREGAS = Empty
    boolTemPCotaPData = False
    Dim intPEGAFECH As Integer
    
    '' Programação de Entrega
    With grdProgEntrega
        If (.Rows - 1) > 0 Then
            ReDim arrENTREGAS(1 To (.Rows - 1), 0 To 22) As String
            For i = 1 To (.Rows - 1)
                arrENTREGAS(i, conCOL_SonProgEntr_IdProduto) = .Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto)
                arrENTREGAS(i, conCOL_SonProgEntr_QtdProd) = .Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd)
                
                arrENTREGAS(i, conCOL_SonProgEntr_DataEntrega) = "Null"
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega))) > 0 Then
                   arrENTREGAS(i, conCOL_SonProgEntr_DataEntrega) = "'" & Format(CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)), "MM/DD/YYYY") & "'"
                End If
                
                arrENTREGAS(i, conCOL_SonProgEntr_Action2Do) = .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do)
                arrENTREGAS(i, conCOL_SonProgEntr_OBSOP) = Trim(Replace(Replace(.Cell(flexcpText, i, conCOL_SonProgEntr_OBSOP), "'", ""), ",", ""))
                
                arrENTREGAS(i, 5) = .Cell(flexcpText, i, conCOL_SonProgEntr_FechTpFr)
                intPEGAFECH = grdProduto.FindRow(.Cell(flexcpText, i, conCOL_SonProgEntr_FechTpFr), , conCOL_SonProd_FechTpFr)
                If intPEGAFECH > -1 Then
                    arrENTREGAS(i, 5) = grdProduto.Cell(flexcpData, intPEGAFECH, conCOL_SonProd_FechTpFr)
                End If
                
                arrENTREGAS(i, 6) = Empty
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP))) > 0 Then arrENTREGAS(i, 6) = .Cell(flexcpText, i, conCOL_SonProgEntr_CodOP)
                
                arrENTREGAS(i, 7) = .Cell(flexcpText, i, conCOL_SonProgEntr_INDICE)
                
                arrENTREGAS(i, 8) = Empty
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_INDICEBKP))) > 0 Then arrENTREGAS(i, 8) = .Cell(flexcpText, i, conCOL_SonProgEntr_INDICEBKP)
            
                arrENTREGAS(i, 9) = Empty
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP))) > 0 Then arrENTREGAS(i, 9) = Trim(Replace(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP), "/", ""))
            
                arrENTREGAS(i, 10) = Empty
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntregaBKP))) > 0 Then arrENTREGAS(i, 10) = "'" & Format(CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntregaBKP)), "MM/DD/YYYY") & "'"
            
                '' ===============================
                '' Pega os Dados do Produto
                intLINHAPROD = grdProduto.FindRow(.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto), , conCOL_SonProd_IdProduto)
                If intLINHAPROD > -1 Then
                    arrENTREGAS(i, 11) = grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_Codigo)
                    arrENTREGAS(i, 12) = grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_AltFilme)
                    arrENTREGAS(i, 13) = grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_FotNovo)
                    arrENTREGAS(i, 14) = grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_Repeticao)
                End If
                '' ===============================
            
                If .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Insert Then
                    lngIDPROGENTREGA = objBLBFunc.Gera_Codigo(Me.Name & "_PROGENTR" & strNOMFILIAL, FILIAL, Linha)
                    arrENTREGAS(i, 16) = lngIDPROGENTREGA
                Else
                    arrENTREGAS(i, 16) = .Cell(flexcpText, i, conCOL_SonProgEntr_IDINTERNO)
                End If
                
                
                arrENTREGAS(i, 17) = .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP)
                
                If .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) <> dacEnumUpdateAction_delete And _
                   .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) <> dacEnumUpdateAction_Ignore Then
                    If .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 6 Or _
                       .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 7 Then boolTemPCotaPData = True
                    If .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 4 Then .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 0
                Else
                    If .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Ignore Then
                        If .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 6 Or _
                           .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 7 Then boolTemPCotaPData = True
                        If .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 4 Then .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 0
                    End If
                End If
                
                '' Palhets
                arrENTREGAS(i, 18) = "Null"
                arrENTREGAS(i, 19) = "Null"
            
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_QTDENOPALHET))) Then arrENTREGAS(i, 18) = CLng(.Cell(flexcpText, i, conCOL_SonProgEntr_QTDENOPALHET))
                If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_PALHET))) Then arrENTREGAS(i, 19) = CLng(.Cell(flexcpText, i, conCOL_SonProgEntr_PALHET))
                
                arrENTREGAS(i, 20) = .Cell(flexcpText, i, conCOL_SonProgEntr_Action2DoDtEntrega)
                
                If .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update Then
                    If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega))) > 0 And Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntregaBKP))) > 0 Then
                       If CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)) <> CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntregaBKP)) Then
                            arrENTREGAS(i, 21) = .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)
                            arrENTREGAS(i, 22) = .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntregaBKP)
                       End If
                    Else
                       arrENTREGAS(i, 21) = .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)
                       arrENTREGAS(i, 22) = .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)
                    End If
                End If
                
            Next i
        End If
    End With
    objCADPEDVENDA.PROGENTREGAS = arrENTREGAS
    
    '' =====================================
    '' Tem P.Cota e P.Data
    If boolTemPCotaPData = True Then
       If cTipOper = "LF" Then
          objCADPEDVENDA.STATUS = "C"
       ElseIf cTipOper = "LS" Then
          If boolTemOP = True Then objCADPEDVENDA.STATUS = "C"
       End If
   End If
    '' =====================================
    
    '' =====================================
    '' Liberação Financeiro
    '' Pegando Dados Para Gerar a Ordem de Fabricacao
    If cTipOper = "LF" Then
        
        '' =====================================
        '' Itens da Orden de Fabricação
        '' As Ordens de Fabricação OP, São Geradas conforme as Datas de Entrega
        '' =====================================
        '' Quantidades/Data de Entrega
        
        arrENTREGAS = Empty
        With grdProgEntrega
            If (.Rows - 1) > 0 Then
                ReDim arrENTREGAS(1 To (.Rows - 1), 0 To 18) As String
                For i = 1 To (.Rows - 1)
                    If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP))) = 0 Then
                        
                        arrENTREGAS(i, conCOL_SonProgEntr_IdProduto) = .Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto)
                        arrENTREGAS(i, conCOL_SonProgEntr_QtdProd) = .Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd)
                        arrENTREGAS(i, conCOL_SonProgEntr_DataEntrega) = .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)
                        arrENTREGAS(i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Insert
                        
                        If intFILIALPED = 0 Then
                            arrENTREGAS(i, 4) = objBLBFunc.Gera_Codigo("frmCADOORDFAB", FILIAL, Linha) & Year(Now)
                        ElseIf intFILIALPED = 1 Then
                            arrENTREGAS(i, 4) = objBLBFunc.Gera_Codigo("frmCADOORDFAB_STELL", FILIAL, Linha) & Year(Now)
                        End If
                        
                        arrENTREGAS(i, 5) = CDate(objCADPEDVENDA.DTHORALIB)
                        
                        '' ===============================
                        '' Pega os Dados do Produto
                        intLINHAPROD = grdProduto.FindRow(.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto), , conCOL_SonProd_IdProduto)
                        If intLINHAPROD > -1 Then
                            arrENTREGAS(i, 6) = grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_Codigo)
                            arrENTREGAS(i, 7) = grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_AltFilme)
                            arrENTREGAS(i, 8) = grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_FotNovo)
                            arrENTREGAS(i, 9) = grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_Repeticao)
                            arrENTREGAS(i, 11) = grdProduto.Cell(flexcpData, intLINHAPROD, conCOL_SonProd_FechTpFr)
                        End If
                        '' ===============================
                        
                        arrENTREGAS(i, 10) = Trim(Replace(Replace(.Cell(flexcpText, i, conCOL_SonProgEntr_OBSOP), "'", ""), ",", ""))
                        arrENTREGAS(i, 12) = .Cell(flexcpText, i, conCOL_SonProgEntr_INDICE)
                    
                        arrENTREGAS(i, 13) = "Null"
                        If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP))) > 0 Then arrENTREGAS(i, 13) = .Cell(flexcpText, i, conCOL_SonProgEntr_CodOP)
                        
                        arrENTREGAS(i, 14) = .Cell(flexcpText, i, conCOL_SonProgEntr_IDINTERNO)
                    
                        arrENTREGAS(i, 15) = "Null"
                        arrENTREGAS(i, 16) = "Null"
                    
                        If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_QTDENOPALHET))) > 0 Then arrENTREGAS(i, 15) = .Cell(flexcpText, i, conCOL_SonProgEntr_QTDENOPALHET)
                        If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_PALHET))) > 0 Then arrENTREGAS(i, 16) = .Cell(flexcpText, i, conCOL_SonProgEntr_PALHET)
                        
                        If objCADPEDVENDA.STATUS <> "C" Then .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = "0"
                        
                        arrENTREGAS(i, 17) = .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP)
                    
                    Else
                        
                        arrENTREGAS(i, 9) = Empty
                        If objCADPEDVENDA.FOIREPROVADO = True Then
                            arrENTREGAS(i, 9) = Empty
                            If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP))) > 0 Then arrENTREGAS(i, 9) = Trim(Replace(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP), "/", ""))
                        End If
                        
                        If cTipOper = "LF" Then
                            arrENTREGAS(i, 18) = .Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto)
                            arrENTREGAS(i, 6) = Trim(Replace(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP), "/", ""))
                            arrENTREGAS(i, 16) = .Cell(flexcpText, i, conCOL_SonProgEntr_IDINTERNO)
                            arrENTREGAS(i, 17) = .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP)
                        End If
                    
                    End If
                Next i
            End If
        End With
        objCADPEDVENDA.PROGENTREGAS = arrENTREGAS
        
    End If
    '' =====================================
    
    objCADPEDVENDA.FILIALPED = intFILIALPED
    
    If optPARAESTOQUE(1).Value = True Then objCADPEDVENDA.STATUS = "X" '' Para Estoque
    
    
    '' Liquida Pedido
    If cTipOper = "M" Then
        objCADPEDVENDA.CODMOTLIQ = Trim(txtCODMOTLIQ.Text)
        objCADPEDVENDA.OBSLIQ = "'" & Replace(txtOBS_MotLiq.Text, ",", " ") & "'"
        objCADPEDVENDA.STATUS = "L"
        Call PegaOPS(Str(objCADPEDVENDA.CODPEDIDO))
    End If
    
    
    '' Grava
    If objCADPEDVENDA.GRAVASTEEL(cTipOper) = False Then Exit Sub
    MsgBox "O Pedido de venda ( nº " & objCADPEDVENDA.CODPEDIDO & " ) - foi " & PegaStatus(cTipOper) & " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
        
    Call GeraLog
    
    If objCADPEDVENDA.Atualiza(cTipOper, objCADPEDVENDA.CODPEDIDO, FILIAL, Me.Name, intFILIALPED) = False Then Exit Sub
    
    
    If cTipOper = "I" Then
       intRESP = MsgBox("Deseja gerar outro Pedido ?", vbYesNo + vbQuestion, "Aviso")
       
       If intRESP = 6 Then
          Call Inclui
       Else
          Call DestroiObjeto
          Unload Me
       End If
    ElseIf cTipOper = "A" Then
        If objCADPEDVENDA.STATUS = "B" Or objCADPEDVENDA.STATUS = "C" Or objCADPEDVENDA.STATUS = "S" Then
            cTipOper = "C"
            Call Consulta
            If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
            If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO
            Exit Sub
        End If
        
        Call DestroiObjeto
        Unload Me
        
    Else
        If cTipOper <> "A" Then
           Call DestroiObjeto
           Unload Me
        End If
    End If
    
    If objCADPEDVENDA.STATUS = "S" Then Call VisualizaBotoesLibAlteracao("G", cTipOper)
    
    Exit Sub
    
err_grava:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CmdSalva_Click()", Me.Name, "CmdSalva_Click()", strCAMARQERRO)
    
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

On Error GoTo Err_Command1_Click

    If Len(Trim(txtCODVEND.Text)) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Favor Informar o Vendedor !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If

    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CLIE.* " & vbCrLf
   
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIEVEND CVEN" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CVEN.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And CVEN.SGI_CODIGO   = " & txtCODVEND.Text & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL   = CVEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO   = CVEN.SGI_CODCLI" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_DESBCLIE = 1"
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    If Len(Trim(varRETORNO)) = 0 Then
        txtCIDCLIE.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(varRETORNO)) > 0 Then txtCIDCLIE.Text = varRETORNO
    
    If Verifica_Credito = "N" Then
        txtCIDCLIE.Text = ""
        Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescCliente, "Command1_Click()")
    If Len(Trim(lblDescCliente.Caption)) = 0 Then txtCIDCLIE.Text = ""
    
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        objCADPEDVENDA.PERMITEFECHOP = PermiteFechamOP(txtCIDCLIE.Text)
    End If
    
    txtCodCondPgto.SetFocus

    Exit Sub
    
Err_Command1_Click:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command1_Click()", Me.Name, "Command1_Click()", strCAMARQERRO)


End Sub

Private Sub Command10_Click()
    
On Error GoTo Err_Command10_Click
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Transportadoras")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODTRANSP.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTRANSP", varRETORNO, lblDescTransp, "Command10_Click()")
    If Len(Trim(lblDescTransp.Caption)) = 0 Then txtCODTRANSP.Text = ""
    
    txtCODTRANSP.SetFocus
    
    Exit Sub

Err_Command10_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command10_Click()", Me.Name, "Command10_Click()", strCAMARQERRO)

End Sub

Private Sub Command11_Click()
    If cTipOper = "R" Then Call IncRegGridRep
End Sub

Private Sub Command2_Click()

On Error GoTo Err_Command2_Click


    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ATIVO      = 1"
    
    If lngCodVendedor > 0 Then
        sSql = sSql & "   And SGI_CODIGO = " & lngCodVendedor
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVEND.Text = varRETORNO
    
    Call PegaDescTabelasVend("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedor, "Command2_Click()")
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then txtCODVEND.Text = ""
    
    If txtCODVEND.Enabled = True Then txtCODVEND.SetFocus

    Exit Sub
    
Err_Command2_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command2_Click()", Me.Name, "Command2_Click()", strCAMARQERRO)


End Sub

Private Sub Command26_Click()
    
On Error GoTo Err_Command26_Click
    
    If cTipOper = "C" Then Exit Sub
    If cTipOper = "I" Or cTipOper = "A" Then
        With grdProduto
            If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonProgEntr_IdProduto))) = 0 Then Exit Sub
            If .Cell(flexcpText, .Row, conCOL_SonProd_Action2Do) = dacEnumUpdateAction_Insert Then
                If (.Rows - 1) = 1 Then .Rows = 1
                If (.Rows - 1) > 1 Then
                   Call objBLBFunc.ExcLinhaGrdFilho(grdProgEntrega, conCOL_SonProgEntr_IdProduto, grdProduto.Cell(flexcpText, .Row, conCOL_SonProd_IdProduto))
                   Call objBLBFunc.ExclLinhaGrid(grdProduto, grdProduto.Row)
                End If
            Else
                .Cell(flexcpText, .Row, conCOL_SonProd_Action2Do) = dacEnumUpdateAction_delete
                Call objBLBFunc.ExcLinhaGrdFilhoAct2Do(grdProgEntrega, conCOL_SonProgEntr_IdProduto, grdProduto.Cell(flexcpText, .Row, conCOL_SonProd_IdProduto), conCOL_SonProgEntr_Action2Do)
                Call objBLBFunc.ExclLinhaGridAction2Do(grdProduto, .Row, conCOL_SonProd_Action2Do)
            End If
            Call CalcTotPedido
        End With
    End If
    
    Exit Sub
    
Err_Command26_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command26_Click()", Me.Name, "Command26_Click()", strCAMARQERRO)
    
End Sub

Private Sub Command27_Click()
    If (cTipOper = "I" Or cTipOper = "A") Then
        Dim boolTRAVA As Boolean
        boolTRAVA = False
        With grdProduto
            If (.Rows - 1) = 1 Then
                MsgBox "Somente e permitido inserir um rótulo !!!", vbOKOnly + vbExclamation, "Aviso"
                boolTRAVA = True
            End If
        End With
        If boolTRAVA = True Then Exit Sub
        Call IncRegGridProdtos
    End If
End Sub

Private Sub Command3_Click()

On Error GoTo Err_Command3_Click

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADESPORCA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Tipo de Pedido", "CADESPORCA.clsCADESPORCA")
    
    If Len(Trim(varRETORNO)) > 0 Then txtTIPPED.Text = varRETORNO
    
    If Len(Trim(txtTIPPED.Text)) = 0 Then Exit Sub
        
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADESPORCA", varRETORNO, lblDescTpPed, "Command3_Click()")
    If Len(Trim(lblDescTpPed.Caption)) = 0 Then txtTIPPED.Text = ""
    
    txtTIPPED.SetFocus


    Exit Sub
    
Err_Command3_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command3_Click()", Me.Name, "Command3_Click()", strCAMARQERRO)

End Sub


Private Sub Command4_Click()
    
On Error GoTo Err_Command4_Click
    
    If cTipOper = "C" Then Exit Sub
    If cTipOper = "I" Or cTipOper = "A" Then
        With grdProgEntrega
            If .Cell(flexcpText, .Row, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Ignore Then
                .Cell(flexcpText, .Row, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_delete
                Call objBLBFunc.ExclLinhaGridAction2Do(grdProgEntrega, .Row, conCOL_SonProgEntr_Action2Do)
            ElseIf .Cell(flexcpText, .Row, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Insert Then
                If (.Rows - 1) = 1 Then .Rows = 1
                If (.Rows - 1) > 1 Then Call objBLBFunc.ExclLinhaGrid(grdProgEntrega, .Row)
            End If
            Call RefazIndice
        End With
    End If
    
    Exit Sub
    
Err_Command4_Click:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command4_Click()", Me.Name, "Command4_Click()", strCAMARQERRO)
    
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
    sSql = sSql & "        SGI_CADMOTLIQ " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Motivos de Liquidação do Pedido")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODMOTLIQ.Text = varRETORNO
    
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMOTLIQ", varRETORNO, lblDescMotLiq, "Command6_Click()")
    If Len(Trim(lblDescMotLiq.Caption)) = 0 Then txtCODMOTLIQ.Text = ""
    
    If txtOBS_MotLiq.Enabled = True Then txtOBS_MotLiq.SetFocus

    Exit Sub
    
Err_Command6_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command6_Click()", Me.Name, "Command6_Click()", strCAMARQERRO)

End Sub




Private Sub Form_Activate()
    
On Error GoTo Err_Form_Activate
    
    If cTipOper = "L" Then txtOBS.SetFocus
    
    Exit Sub

Err_Form_Activate:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Form_Activate()", Me.Name, "Form_Activate()", strCAMARQERRO)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo Err_Form_KeyDown
    
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
    If KeyCode = vbKeyF5 Then booAltorizado = ChamaSenhaUsuario

    Exit Sub

Err_Form_KeyDown:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Form_KeyDown()", Me.Name, "Form_KeyDown()", strCAMARQERRO)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
On Error GoTo Err_Form_KeyPress
    
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"

    Exit Sub
    
Err_Form_KeyPress:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Form_KeyPress()", Me.Name, "Form_KeyPress()", strCAMARQERRO)

End Sub

Private Sub Form_Load()

On Error GoTo Form_Load

    strCAMARQERRO = Right(Linha(9), Len(Trim(Linha(9))) - 8)
    
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
    objCADPEDVENDA.FILIAL = FILIAL
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    lngQTDDIASLDTIMF = 20
    
    boolALteradoITEN = False
    
    strNOMFILIAL = ""
    If intFILIALPED = 1 Then strNOMFILIAL = "_STEEL"
   
    '' Depois Retirar
    chkVerificado.Visible = True
   
    If cTipOper = "I" Then Inclui
    If cTipOper = "A" Then Altera
    If cTipOper = "C" Then Consulta
   
    If cTipOper = "LF" Or _
       cTipOper = "LN" Or _
       cTipOper = "LS" Or _
       cTipOper = "LV" Or _
       cTipOper = "X" Or _
       cTipOper = "LC" Then Libera
   
    If cTipOper = "R" Then Reprova
    If cTipOper = "M" Then Liquida
    If cTipOper = "D" Then Deslibera

    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA - Versão " & strVERSAO
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROL - Versão " & strVERSAO
    
    If boolSomenteCons = True Then cmdAltera.Enabled = False
    
    Call AtivaDesativaBotoes
    
    Exit Sub
    
Form_Load:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Form_Load", Me.Name, "Form_Load", strCAMARQERRO)


End Sub

Private Sub Inclui()

On Error GoTo Err_Inclui
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Pedido de Venda - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos Me
    
    stCAMPOSVENDA.Tab = 0
    stCAMPOSVENDA.TabVisible(2) = False
        
    Frame3.Enabled = True
    Frame4.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True
    
    Frame13.Visible = False
    txtOBS2.Locked = False
    
    
    lblCODIGO.Caption = ""
    lblSTATUS.Caption = "ABERTO"
    objCADPEDVENDA.STATUS = ""
    mskDATAPED.Text = Format(Now, "DD/MM/YYYY")
    lblTotalItens.Caption = ""
    
    objBLBFunc.Preenche_Estado cboESTENTR
    objBLBFunc.Preenche_Estado cboESTCOBR
    
    Call InitGridReprovacao
    Call InitGridProd
    Call InitGridProg
    Call InitGridOrdemFat
    Call InitGridLogPed
    Call InitGridProducao
    
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    txtOutrDesp.Text = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    
    '' --------------------
    '' Desconto
    lblVLDESCONTO.Caption = ""
    '' --------------------
    
    Call LimpaCamposLabel
    Call LimpaCamposDadosAdicionais
    Call LimpaCampoSaldoRot
    Call LimpaSaldoPedido

    Call Pega_Vendedor(lngCodVendedor)
    
    txtOBS2.Text = "As datas de entrega serão consideradas no prazo, 4 dias antes e 4 dias após a data acordada"
    
    objCADPEDVENDA.FILIALPED = intFILIALPED

    optESPECIAL(0).Value = True
    optPARAESTOQUE(0).Value = True

    Call AbilDesConferido(False, 0)

    Call VisualizaBotoesPCD(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibAlteracao(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibFotolito(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibComercial(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibFinanceira(objCADPEDVENDA.STATUS, cTipOper)

    Exit Sub

Err_Inclui:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Inclui()", Me.Name, "Inclui()", strCAMARQERRO)

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub


Private Sub grdProduto_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     
On Error GoTo Err_grdProduto_AfterEdit
     
     Dim i As Integer
     With grdProduto
          Select Case Col
                 Case conCOL_SonProd_Codigo
                 Case conCOL_SonProd_QtdProd, _
                      conCOL_SonProd_VlUniProd, _
                      conCOL_SonProd_PorcDesc, _
                      conCOL_SonProd_PorcIPI
                      
                      If Col = conCOL_SonProd_VlUniProd Or _
                         Col = conCOL_SonProd_PorcDesc Or _
                         Col = conCOL_SonProd_PorcIPI Then
                         If Len(Trim(.Cell(flexcpText, Row, Col))) > 0 Then .Cell(flexcpText, Row, Col) = Format(CDbl(.Cell(flexcpText, Row, Col)), "#,##0.00")
                      End If
                      .Cell(flexcpText, Row, conCOL_SonProd_VlTotal) = Format(CalcItenGrid(Row), "#,##0.00")
                      Call CalcTotPedido
                      
                 Case conCOL_SonProd_AltFilme, _
                      conCOL_SonProd_FotNovo, _
                      conCOL_SonProd_Repeticao, _
                      conCOL_SonProd_FechTpFr
                        
                        Call LimpaCamposDadosAdicionais
                        Call PegadadosGrid(Row)
                        For i = 1 To (grdProgEntrega.Rows - 1)
                            If grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) = grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto) Then
                               grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_FechTpFr) = grdProduto.Cell(flexcpText, Row, conCOL_SonProd_FechTpFr)
                            End If
                        Next i
                        
                        
                        
                        If Col = conCOL_SonProd_AltFilme Then
                            If .Cell(flexcpText, Row, Col) = 1 Then
                               .Cell(flexcpText, Row, conCOL_SonProd_FotNovo) = 0
                               .Cell(flexcpText, Row, conCOL_SonProd_Repeticao) = 0
                               
                               If .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_delete And _
                                  .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_Ignore Then .Cell(flexcpText, Row, conCOL_SonProd_StatusProd) = 2
                                Call LimpaCamposProEntr(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                            Else
                               .Cell(flexcpText, Row, conCOL_SonProd_FotNovo) = 0
                               .Cell(flexcpText, Row, conCOL_SonProd_Repeticao) = 0
                               .Cell(flexcpText, Row, conCOL_SonProd_StatusProd) = 1
                            End If
                        ElseIf Col = conCOL_SonProd_FotNovo Then
                            If .Cell(flexcpText, Row, Col) = 1 Then
                               Call LimpaCamposProEntr(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                               .Cell(flexcpText, Row, conCOL_SonProd_AltFilme) = 0
                               .Cell(flexcpText, Row, conCOL_SonProd_Repeticao) = 0
                               
                               If .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_delete And _
                                  .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_Ignore Then .Cell(flexcpText, Row, conCOL_SonProd_StatusProd) = 2
                                Call LimpaCamposProEntr(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                            Else
                               .Cell(flexcpText, Row, conCOL_SonProd_AltFilme) = 0
                               .Cell(flexcpText, Row, conCOL_SonProd_Repeticao) = 0
                               .Cell(flexcpText, Row, conCOL_SonProd_StatusProd) = 1
                            End If
                        ElseIf Col = conCOL_SonProd_Repeticao Then
                            If .Cell(flexcpText, Row, Col) = 1 Then
                               .Cell(flexcpText, Row, conCOL_SonProd_AltFilme) = 0
                               .Cell(flexcpText, Row, conCOL_SonProd_FotNovo) = 0
                               
                               If .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_delete And _
                                  .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_Ignore Then .Cell(flexcpText, Row, conCOL_SonProd_StatusProd) = 1
                            Else
                               .Cell(flexcpText, Row, conCOL_SonProd_AltFilme) = 0
                               .Cell(flexcpText, Row, conCOL_SonProd_FotNovo) = 0
                               
                               If .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_delete And _
                                  .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_Ignore Then .Cell(flexcpText, Row, conCOL_SonProd_StatusProd) = 1
                            End If
                        End If
                        
                        If .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_delete And _
                            .Cell(flexcpText, Row, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_Ignore Then Call CorRotulo(CLng(Str(Row)))

          End Select
          
     End With
     
     Exit Sub
     
Err_grdProduto_AfterEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProduto_AfterEdit()", Me.Name, "grdProduto_AfterEdit()", strCAMARQERRO)
     
End Sub


Private Sub grdProduto_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
On Error GoTo Err_grdProduto_BeforeEdit
    
    Select Case Col
    Case conCOL_SonProd_DescProd, _
         conCOL_SonProd_VlTotal, _
         conCOL_SonProd_Fechamento, _
         conCOL_SonProd_Corpo, _
         conCOL_SonProd_Tampa, _
         conCOL_SonProd_Fundo, _
         conCOL_SonProd_Argola, _
         conCOL_SonProd_StatusProd, _
         conCOL_SonProd_GrpPlanMestre, _
         conCOL_SonProd_CodCapacidade, _
         conCOL_SonProd_NECKIN, _
         conCOL_SonProd_HOMOLOGADO, _
         conCOL_SonProd_PALLETS
         Cancel = True
    Case conCOL_SonProd_Conferido
         If cTipOper = "I" Or _
            cTipOper = "A" Or _
            cTipOper = "C" Or _
            cTipOper = "S" Or _
            cTipOper = "LS" Or _
            cTipOper = "V" Or _
            cTipOper = "X" Or _
            cTipOper = "LC" Or _
            cTipOper = "LV" Or _
            cTipOper = "R" Or _
            cTipOper = "D" Then Cancel = True
    Case conCOL_SonProd_Codigo, _
         conCOL_SonProd_PesqProd, _
         conCOL_SonProd_QtdProd
         If cTipOper = "C" Or _
            cTipOper = "D" Or _
            cTipOper = "LF" Or _
            cTipOper = "LN" Or _
            cTipOper = "LV" Or _
            cTipOper = "LS" Or _
            cTipOper = "R" Or _
            cTipOper = "S" Then
            Cancel = True
            Exit Sub
         End If
         If objCADPEDVENDA.STATUS = "C" Or objCADPEDVENDA.STATUS = "4" Or objCADPEDVENDA.STATUS = "P" Then Cancel = True
    Case conCOL_SonProd_QTDELATASPALLETS
         If cTipOper = "C" Then
            Cancel = True
         Else
            If grdProduto.Cell(flexcpText, Row, conCOL_SonProd_PalhetPadrao) = 1 Then Cancel = True
         End If
    Case conCOL_SonProd_VlUniProd, _
         conCOL_SonProd_PorcDesc, _
         conCOL_SonProd_PorcIPI
         If cTipOper = "C" Or _
            cTipOper = "D" Or _
            cTipOper = "LF" Or _
            cTipOper = "LN" Or _
            cTipOper = "LV" Or _
            cTipOper = "LS" Or _
            cTipOper = "R" Or _
            cTipOper = "S" Then
            Cancel = True
            Exit Sub
         End If
    Case conCOL_SonProd_AltFilme, _
         conCOL_SonProd_FotNovo, _
         conCOL_SonProd_Repeticao, _
         conCOL_SonProd_FornTampa, _
         conCOL_SonProd_PesqForn, _
         conCOL_SonProd_FechTpFr, _
         conCOL_SonProd_Conferido
         If cTipOper = "C" Or _
            cTipOper = "D" Or _
            cTipOper = "LF" Or _
            cTipOper = "LN" Or _
            cTipOper = "LV" Or _
            cTipOper = "LS" Or _
            cTipOper = "R" Or _
            cTipOper = "S" Then
            Cancel = True
            Exit Sub
         Else
            grdProduto.Editable = flexEDKbdMouse
         End If
         If (objCADPEDVENDA.STATUS = "C" Or objCADPEDVENDA.STATUS = "4" Or objCADPEDVENDA.STATUS = "P") Then Cancel = True
         
         '' travando colunas se
         '' Fotolito Novo       = Sim
         '' Alteração Fotolito  = Sim
         If (grdProduto.Cell(flexcpText, Row, conCOL_SonProd_StatusProd) = "2" And _
            (objCADPEDVENDA.PegaStatusFotolitoNovo(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto)) = True Or _
            objCADPEDVENDA.PegaStatusFotolitoAlteracao(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto)) = True)) Then
            If Col = conCOL_SonProd_FotNovo Then Cancel = True
            If Col = conCOL_SonProd_AltFilme Then Cancel = True
            If Col = conCOL_SonProd_Repeticao Then Cancel = True
         Else
            If objCADPEDVENDA.PegaStatusFotolitoNovo(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto)) = False Then If Col = conCOL_SonProd_FotNovo Then Cancel = True
         End If
         
    Case Else
        grdProduto.ComboList = ""
    End Select
    
    Exit Sub
    
Err_grdProduto_BeforeEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProduto_BeforeEdit()", Me.Name, "grdProduto_BeforeEdit()", strCAMARQERRO)
    
End Sub

Private Sub grdProduto_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    ' don't resize columns while editing dates
    If cboFechTPFR.Visible Then Cancel = True

    ' don't resize columns while editing dates
    If cboQtdePorPalhet.Visible Then Cancel = True

End Sub

Private Sub grdProduto_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
      With grdProduto
        If .RowSel > 0 And (.Rows - 1) > 0 Then
            Call Volta_Cres_Grid(OldRowSel)
            Call CorRotulo(CInt(Str(OldRowSel)))
        End If
      End With
End Sub

Private Sub grdProduto_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' don't resize columns while editing dates
    If cboFechTPFR.Visible Then Cancel = True

    ' don't resize columns while editing dates
    If cboQtdePorPalhet.Visible Then Cancel = True

End Sub

Private Sub grdProduto_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
On Error GoTo Err_grdProduto_CellButtonClick
    
    Dim strDESCPROD As String
    
    If (grdProduto.Rows - 1) = 0 Then Exit Sub
    If Len(Trim(txtCIDCLIE.Text)) = 0 Then
       MsgBox "O Cliente não foi Informado !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    
    Select Case Col
        Case conCOL_SonProd_Desenho
        
            If Len(Trim(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))) = 0 Then Exit Sub
            
            frmDesenhoPedido.cCaminho = cCaminho
            frmDesenhoPedido.Linha = Linha
            frmDesenhoPedido.iCodigo = iCodigo
            frmDesenhoPedido.cTipOper = cTipOper
            frmDesenhoPedido.FILIAL = FILIAL
            frmDesenhoPedido.strACESSO = strACESSO
            frmDesenhoPedido.strMODPAI = Me.Name
            frmDesenhoPedido.strUSUARIO = strUSUARIO
            frmDesenhoPedido.strDescProduto = Trim(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_Codigo)) & " - " & Trim(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_DescProd))
            frmDesenhoPedido.lngIDProduto = CLng(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
            frmDesenhoPedido.Show vbModal
            
        Case conCOL_SonProd_PesqProd
    
            If cTipOper = "C" Then Exit Sub
            
            ReDim arrCAMPOS(1 To 5, 1 To 6) As String
            ReDim arrTABELA(1 To 1) As String
            ReDim arrTABELA2(1 To 1) As String
            
            Dim strIDPRODUTO As String
            
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
            sSql = sSql & ",LINHA.SGI_DESCRI" & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO  PRO " & vbCrLf
            sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINHA " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       PRO.SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "   And PRO.SGI_CODCLIE     = " & Trim(txtCIDCLIE.Text) & vbCrLf
            sSql = sSql & "   And (PRO.SGI_STATUS     = 1 or PRO.SGI_STATUS      = 2)" & vbCrLf
            sSql = sSql & "   And LINHA.SGI_FILIAL    = PRO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And LINHA.SGI_CODLIN    = PRO.SGI_CODLINPROD " & vbCrLf
            sSql = sSql & "   And PRO.SGI_FILIALPED   = " & intFILIALPED & vbCrLf
            
            arrTABELA(1) = sSql
            
            sSql = ""
            
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
            sSql = sSql & "  ,LINHA.SGI_DESCRI" & vbCrLf
            
            sSql = sSql & "    From" & vbCrLf
            sSql = sSql & "         SGI_CADPRODUTO   PRO" & vbCrLf
            sSql = sSql & "        ,SGI_PRODATECLIE  PCL" & vbCrLf
            sSql = sSql & "        ,SGI_CADLINHAPRODUTO LINHA " & vbCrLf
            sSql = sSql & "   Where" & vbCrLf
            sSql = sSql & "         PCL.SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "     And PCL.SGI_IDCLIENTE   = " & Trim(txtCIDCLIE.Text) & vbCrLf
            sSql = sSql & "     And PCL.SGI_FILIAL      = PRO.SGI_FILIAL" & vbCrLf
            sSql = sSql & "     And PCL.SGI_IDPRODUTO   = PRO.SGI_IDPRODUTO" & vbCrLf
            sSql = sSql & "     And (PRO.SGI_STATUS     = 1 Or PRO.SGI_STATUS      = 2)" & vbCrLf
            sSql = sSql & "     And LINHA.SGI_FILIAL    = PRO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "     And LINHA.SGI_CODLIN    = PRO.SGI_CODLINPROD " & vbCrLf
            sSql = sSql & "     And PRO.SGI_FILIALPED   = " & intFILIALPED & vbCrLf
            
            arrTABELA2(1) = sSql
            
            '' ------------------------------
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "2000"
            arrCAMPOS(1, 5) = "PRO.SGI_CODIGO"
            arrCAMPOS(1, 6) = ""
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "PRO.SGI_DESCRICAO"
            arrCAMPOS(2, 6) = ""
            
            arrCAMPOS(3, 1) = "SGI_COMPLEMENTO"
            arrCAMPOS(3, 2) = "S"
            arrCAMPOS(3, 3) = "Complemento"
            arrCAMPOS(3, 4) = "3000"
            arrCAMPOS(3, 5) = "PRO.SGI_COMPLEMENTO"
            arrCAMPOS(3, 6) = ""
            
            arrCAMPOS(4, 1) = "SGI_CODCLIE"
            arrCAMPOS(4, 2) = "N"
            arrCAMPOS(4, 3) = "Cod.Cliente"
            arrCAMPOS(4, 4) = "1500"
            arrCAMPOS(4, 5) = "PRO.SGI_CODCLIE"
            arrCAMPOS(4, 6) = ""
            
            arrCAMPOS(5, 1) = "SGI_DESCRI"
            arrCAMPOS(5, 2) = "C"
            arrCAMPOS(5, 3) = "Capacidade"
            arrCAMPOS(5, 4) = "2000"
            arrCAMPOS(5, 5) = "LINHA.SGI_CODIGO"
            arrCAMPOS(5, 6) = "SGI_CODIGO|SGI_DESCRI|SGI_CADLINHAPRODUTO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos", "", arrTABELA2)
            
            If Len(Trim(varRETORNO)) > 0 Then
            
                If objBLBFunc.FcVerifItensRepetidosAct2Do(grdProduto, Row, conCOL_SonProd_Codigo, varRETORNO, conCOL_SonProd_Action2Do) = False Then
                     MsgBox "Este Produto já foi relacionado na Gride !!!", vbOKOnly + vbExclamation
                     Exit Sub
                End If
            
                strIDPRODUTO = PegaIDProduto(Trim(varRETORNO))
                If Len(Trim(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))) > 0 And grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto) <> strIDPRODUTO Then
                    Call objBLBFunc.ExcLinhaGrdFilho(grdProgEntrega, conCOL_SonProgEntr_IdProduto, grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                End If
               
               
               grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto) = strIDPRODUTO
               If Len(Trim(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))) = 0 Then
                  Exit Sub
               End If
               
               '' verifica se existe pedidos pendentes para este produto
               ''If PegaPedidosAberto(Trim(txtCIDCLIE.Text), Trim(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))) = False Then
               ''   Exit Sub
               ''End If
               
               Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, grdProduto.Cell(flexcpText, Row, conCOL_SonProd_Codigo), varRETORNO)
               grdProduto.Cell(flexcpText, Row, conCOL_SonProd_Codigo) = varRETORNO
               
               strDESCPROD = PegaDescrProduto(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
               If Len(Trim(strDESCPROD)) = 0 Then
                    Call LimpaCamposGrid(Row)
                    Exit Sub
               End If
               
               grdProduto.Cell(flexcpText, Row, conCOL_SonProd_DescProd) = strDESCPROD
               Call PosCol(conCOL_SonProd_QtdProd, Row)
               
               grdProduto.Cell(flexcpText, Row, conCOL_SonProd_NECKIN) = objCADPEDVENDA.PegaNECKIN(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
               grdProduto.Cell(flexcpText, Row, conCOL_SonProd_HOMOLOGADO) = objCADPEDVENDA.PegaHOMOLOGADO(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
               grdProduto.Cell(flexcpText, Row, conCOL_SonProd_GrpPlanMestre) = PegaGrdPMestre(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto), CLng(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_NECKIN)))
               grdProduto.Cell(flexcpText, Row, conCOL_SonProd_CodCapacidade) = PegaGrdCodCapac(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
               grdProduto.Cell(flexcpText, Row, conCOL_SonProd_QTDELATASPALLETS) = PegaQtdeLT_4_Palhets(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
           
               Call PreenchCboFechTPFR(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
               Call PreenchCboPallet(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
            
               ''If Len(Trim(grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_QtdProd))) > 0 Then
               ''    Call ConferePalhets(grdProduto.Row, CLng(grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_QtdProd)))
               ''End If
            
                Call CorRotulo(CInt(Str(Row)))
            
            End If
            
        Case conCOL_SonProd_PesqForn
    
            If cTipOper = "C" Then Exit Sub
    
            ReDim arrCAMPOS(1 To 4, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
    
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       FORN.SGI_CODIGO " & vbCrLf
            sSql = sSql & "      ,FORN.SGI_CPFCNPJ " & vbCrLf
            sSql = sSql & "      ,FORN.SGI_RAZAOSOC " & vbCrLf
            sSql = sSql & "      ,FORN.SGI_NOMFANTA " & vbCrLf
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADFORNEC FORN " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       FORN.SGI_FILIAL = " & FILIAL
    
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1000"
            arrCAMPOS(1, 5) = "FORN.SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "CNPJ/CPF"
            arrCAMPOS(2, 4) = "1500"
            arrCAMPOS(2, 5) = "FORN.SGI_CPFCNPJ"
            
            arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
            arrCAMPOS(3, 2) = "S"
            arrCAMPOS(3, 3) = "Razão Social"
            arrCAMPOS(3, 4) = "3500"
            arrCAMPOS(3, 5) = "FORN.SGI_RAZAOSOC"
            
            arrCAMPOS(4, 1) = "SGI_NOMFANTA"
            arrCAMPOS(4, 2) = "S"
            arrCAMPOS(4, 3) = "Nome Fantasia"
            arrCAMPOS(4, 4) = "3500"
            arrCAMPOS(4, 5) = "FORN.SGI_NOMFANTA"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Fornecedores")
    
            If Len(Trim(varRETORNO)) > 0 Then
                grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_FornTampa) = varRETORNO
            End If
    
        Case conCOL_SonProd_OS_Artes
        
            If Len(Trim(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))) = 0 Then Exit Sub
            
            frmCADOSARTES.FILIAL = FILIAL
            frmCADOSARTES.strNOMFILIAL = strNOMFILIAL
            frmCADOSARTES.strIDPRODUTO = grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto)
            frmCADOSARTES.cTipOper = cTipOper
            frmCADOSARTES.strRETORNO = ""
            frmCADOSARTES.mskDTPED = mskDATAPED.Text
            frmCADOSARTES.intALTFILME = grdProduto.Cell(flexcpText, Row, conCOL_SonProd_AltFilme)
            frmCADOSARTES.intFOTNOVO = grdProduto.Cell(flexcpText, Row, conCOL_SonProd_FotNovo)
            frmCADOSARTES.lngCODPED = objCADPEDVENDA.CODPEDIDO
            frmCADOSARTES.Show vbModal
    
    End Select

    Exit Sub
    
Err_grdProduto_CellButtonClick:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProduto_CellButtonClick()", Me.Name, "grdProduto_CellButtonClick()", strCAMARQERRO)
    

End Sub

Private Sub grdProduto_Click()
    Call MostraDados
End Sub

Private Sub grdProduto_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     
On Error GoTo Err_grdProduto_KeyPressEdit
     
     With grdProduto
          Select Case Col
                    Case conCOL_SonProd_Codigo
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_SonProd_VlUniProd, _
                         conCOL_SonProd_PorcDesc, _
                         conCOL_SonProd_PorcIPI
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsCurrency)
                    Case conCOL_SonProd_QtdProd
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
     
     Exit Sub
     
Err_grdProduto_KeyPressEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProduto_KeyPressEdit()", Me.Name, "grdProduto_KeyPressEdit()", strCAMARQERRO)
     
End Sub

Private Sub grdProduto_RowColChange()
    Call MostraDados
End Sub

Private Sub grdProduto_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With grdProduto

        ' if this is a date column, edit it with the date picker control
        If .Col = conCOL_SonProd_FechTpFr Then
            
            Call PreenchCboFechTPFR(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
            
            ' we'll handle the editing ourselves
            Cancel = True
            
            ' position date picker control over cell
            cboFechTPFR.Left = .Cell(flexcpLeft, Row, Col) + 100
            cboFechTPFR.Top = .Cell(flexcpTop, Row, Col) + 250
            cboFechTPFR.Width = .Cell(flexcpWidth, Row, Col)
            
            ' initialize value, save original in tag in case user hits escape
            ''cboFechTPFR.Value = cboFechTPFR
            ''cboFechTPFR.Tag = cboFechTPFR
            
            ' show and activate date picker control
            cboFechTPFR.Visible = True
            cboFechTPFR.SetFocus
            
            ' make it drop down the calendar
            ''SendKeys "{f4}"
            
        ' if this is a date column, edit it with the date picker control
        ElseIf .Col = conCOL_SonProd_QTDELATASPALLETS Then
            
            Call PreenchCboPallet(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
            
            ' we'll handle the editing ourselves
            Cancel = True
            
            ' position date picker control over cell
            cboQtdePorPalhet.Left = .Cell(flexcpLeft, Row, Col) + 250
            cboQtdePorPalhet.Top = .Cell(flexcpTop, Row, Col) + 600
            cboQtdePorPalhet.Width = .Cell(flexcpWidth, Row, Col)
            
            ' initialize value, save original in tag in case user hits escape
            ''cboFechTPFR.Value = cboFechTPFR
            ''cboFechTPFR.Tag = cboFechTPFR
            
            ' show and activate date picker control
            cboQtdePorPalhet.Visible = True
            cboQtdePorPalhet.SetFocus
            
            ' make it drop down the calendar
            ''SendKeys "{f4}"
        
        End If

    End With


End Sub

Private Sub grdProduto_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     
On Error GoTo Err_grdProduto_ValidateEdit
     
     Dim curVLUNITARIO As Currency
     Dim strIDPRODUTO As String
     Dim i As Integer
     
     
     
     With grdProduto
          Select Case Col
                 Case conCOL_SonProd_Codigo
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdProduto, Row, conCOL_SonProd_Codigo, .EditText) = False Then
                           MsgBox "Este produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGrid(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        strIDPRODUTO = Trim(PegaIDProduto(Trim(.EditText)))
                        
                        If grdProduto.Cell(flexcpText, .Row, conCOL_SonProd_IdProduto) <> strIDPRODUTO Then
                            Call objBLBFunc.ExcLinhaGrdFilho(grdProgEntrega, conCOL_SonProgEntr_IdProduto, grdProduto.Cell(flexcpText, .Row, conCOL_SonProd_IdProduto))
                        End If
                                            
                        .Cell(flexcpText, Row, conCOL_SonProd_IdProduto) = strIDPRODUTO
                        
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))) = 0 Then
                           MsgBox "Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
                           Cancel = True
                           Exit Sub
                        End If
                        
                        If Len(Trim(PegaDescrProduto(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto)))) = 0 Then
                           MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonProd_DescProd) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_SonProd_DescProd) = PegaDescrProduto(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                        Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpText, Row, conCOL_SonProd_Codigo), .EditText)
                        Call PosCol(conCOL_SonProd_QtdProd, Row)
                        
                        grdProduto.Cell(flexcpText, Row, conCOL_SonProd_NECKIN) = objCADPEDVENDA.PegaNECKIN(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                        grdProduto.Cell(flexcpText, Row, conCOL_SonProd_HOMOLOGADO) = objCADPEDVENDA.PegaHOMOLOGADO(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                        grdProduto.Cell(flexcpText, Row, conCOL_SonProd_GrpPlanMestre) = PegaGrdPMestre(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto), CLng(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_NECKIN)))
                        grdProduto.Cell(flexcpText, Row, conCOL_SonProd_CodCapacidade) = PegaGrdCodCapac(.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                
                        Call PreenchCboFechTPFR(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                        Call PreenchCboPallet(grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
               
                Case conCOL_SonProd_QtdProd
                    If .EditText = Empty Then
                       Call PosCol(conCOL_SonProd_VlUniProd, Row)
                       Exit Sub
                    End If
                    If CLng(Replace(Replace(.EditText, ",", ""), ".", "")) = 0 Then
                       MsgBox "ATENÇÃO" & vbCrLf & "Não é permitido valores iqual a 0 !!!", vbOKOnly + vbExclamation, "Aviso"
                       Cancel = True
                       Exit Sub
                    End If
                    curVLUNITARIO = 0
                    If Len(Trim(.Cell(flexcpText, Row, conCOL_SonProd_VlUniProd))) > 0 Then curVLUNITARIO = CCur(.Cell(flexcpText, Row, conCOL_SonProd_VlUniProd))
                    
                    Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpText, Row, conCOL_SonProd_QtdProd), .EditText)
                    Call MudaActio2DoFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, .Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                    Call PosCol(conCOL_SonProd_VlUniProd, Row)
                
                    '' Depois voltar
                    ''If ConferePalhets(Row, CLng(.EditText)) = False Then
                    ''   Cancel = True
                    ''   Exit Sub
                    ''End If
                
                Case conCOL_SonProd_VlUniProd
                    If .EditText = Empty Then
                        Call PosCol(conCOL_SonProd_PorcDesc, Row)
                        Exit Sub
                    End If
                    If Not IsNumeric(.EditText) Then
                        MsgBox "Valor Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                        Cancel = True
                        Exit Sub
                    End If
                    Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpText, Row, conCOL_SonProd_VlUniProd), .EditText)
                    Call PosCol(conCOL_SonProd_PorcDesc, Row)
                Case conCOL_SonProd_PorcDesc
                    If .EditText = Empty Then
                       Call PosCol(conCOL_SonProd_PorcIPI, Row)
                       Exit Sub
                    End If
                    If Not IsNumeric(.EditText) Then
                        MsgBox "Valor Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                        Cancel = True
                        Exit Sub
                    End If
                    Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpText, Row, conCOL_SonProd_PorcDesc), .EditText)
                    Call PosCol(conCOL_SonProd_PorcIPI, Row)
                Case conCOL_SonProd_PorcIPI
                    If .EditText = Empty Then
                       Call IncRegGridProdtos
                       Call PosCol(conCOL_SonProd_Codigo, (Row - 1))
                       Exit Sub
                    End If
                    If Not IsNumeric(.EditText) Then
                        MsgBox "Valor Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                        Cancel = True
                        Exit Sub
                    End If
                    Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpText, Row, conCOL_SonProd_PorcIPI), .EditText)
                    Call PosCol(conCOL_SonProd_Codigo, (.Rows - 1))
                Case conCOL_SonProd_FornTampa
                    If .EditText = Empty Then Exit Sub
                    Cancel = ValidaFornecedor(Str(.EditText))
                Case conCOL_SonProd_FechTpFr
                    For i = 1 To (grdProgEntrega.Rows - 1)
                        If grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) = grdProduto.Cell(flexcpText, Row, conCOL_SonProd_IdProduto) Then
                           grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_FechTpFr) = grdProduto.Cell(flexcpData, Row, conCOL_SonProd_FechTpFr)
                        End If
                    Next i
                    Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpTextDisplay, Row, conCOL_SonProd_FechTpFr), .EditText)
                    Call MudaActio2DoFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, .Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                Case conCOL_SonProd_AltFilme
                    Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpTextDisplay, Row, conCOL_SonProd_AltFilme), .EditText)
                    Call MudaActio2DoFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, .Cell(flexcpText, Row, conCOL_SonProd_IdProduto))

                Case conCOL_SonProd_FotNovo
                    Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpTextDisplay, Row, conCOL_SonProd_FotNovo), .EditText)
                    Call MudaActio2DoFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, .Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
                
                Case conCOL_SonProd_Repeticao
                    Call objBLBFunc.TrocaAction2Do(grdProduto, Row, conCOL_SonProd_Action2Do, .Cell(flexcpTextDisplay, Row, conCOL_SonProd_Repeticao), .EditText)
                    Call MudaActio2DoFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, .Cell(flexcpText, Row, conCOL_SonProd_IdProduto))
          End Select
     End With
     
     Exit Sub
     
Err_grdProduto_ValidateEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProduto_ValidateEdit()", Me.Name, "grdProduto_ValidateEdit()", strCAMARQERRO)
     
     
End Sub


Private Sub grdProgEntrega_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
On Error GoTo Err_grdProgEntrega_BeforeEdit

    Dim lngLIN As Long

    Select Case Col
    Case conCOL_SonProgEntr_QtdProd, _
         conCOL_SonProgEntr_PegaPlanMestre
         If cTipOper = "C" Or _
            cTipOper = "D" Or _
            cTipOper = "LF" Or _
            cTipOper = "LN" Or _
            cTipOper = "R" Or _
            cTipOper = "S" Or _
            cTipOper = "LS" Or _
            cTipOper = "LC" Then
            
            If Col = conCOL_SonProgEntr_PegaPlanMestre Then
                If objCADPEDVENDA.STATUS <> "R" Then Cancel = True
            Else
                Cancel = True
            End If
            Exit Sub
         End If
         
         If Col = conCOL_SonProgEntr_QtdProd Then
            If objCADPEDVENDA.STATUS = "C" Or objCADPEDVENDA.STATUS = "4" Or objCADPEDVENDA.STATUS = "P" Then Cancel = True
         End If
         
         If Col = conCOL_SonProgEntr_PegaPlanMestre Then
            lngLIN = grdProduto.FindRow(grdProgEntrega.Cell(flexcpText, Row, conCOL_SonProgEntr_IdProduto), , conCOL_SonProd_IdProduto)
            If lngLIN = -1 Then Exit Sub
            If grdProduto.Cell(flexcpText, lngLIN, conCOL_SonProd_StatusProd) = 2 Then Cancel = True
         End If
    Case conCOL_SonProgEntr_CodOP, _
         conCOL_SonProgEntr_IDINTERNO, _
         conCOL_SonProgEntr_DataEntrega, _
         conCOL_SonProgEntr_DataPrevLito, _
         conCOL_SonProgEntr_DataPrevProd, _
         conCOL_SonProgEntr_DescStatusOP, _
         conCOL_SonProgEntr_StatusOP, _
         conCOL_SonProgEntr_GrpPlanMestre, _
         conCOL_SonProgEntr_QTDENOPALHET, _
         conCOL_SonProgEntr_PALHET, _
         conCOL_SonProgEntr_CODIDPROG, _
         conCOL_SonProgEntr_CODSTATAPONT, _
         conCOL_SonProgEntr_DESCSTATUSAPONT
         Cancel = True
    Case Else
        grdProgEntrega.ComboList = ""
    End Select
    
    Exit Sub

Err_grdProgEntrega_BeforeEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProgEntrega_BeforeEdit()", Me.Name, "grdProgEntrega_BeforeEdit()", strCAMARQERRO)

End Sub

Private Sub grdProgEntrega_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

On Error GoTo Err_grdProgEntrega_CellButtonClick
    
    Dim strDESCPROD     As String
    Dim intLINHA        As Integer
    
    intLINHA = 1
    If (grdProduto.Rows - 1) = 0 Then
       MsgBox "ATENÇÃO" & vbCrLf & _
              "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    intLINHA = 2
    If (grdProduto.RowSel) = 0 Then
       MsgBox "ATENÇÃO" & vbCrLf & _
              "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    
    Select Case Col
        Case conCOL_SonProgEntr_PegaPlanMestre
        
            With grdProgEntrega
                
                intLINHA = 3
                If .RowSel = 0 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "Selecione uma Programação de Entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                           Exit Sub
                End If
                
                intLINHA = 4
                If (.Rows - 1) = 0 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "Selecione uma Programação de Entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                           Exit Sub
                End If
                
                intLINHA = 5
                If grdProduto.RowSel = 0 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
                           Exit Sub
                End If
                
                intLINHA = 6
                If (grdProduto.Rows - 1) = 0 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
                           Exit Sub
                End If
                
                
                Dim lngSALDOQTDENTR As Long
                Dim strGRPCOD       As String
                Dim strCODLIN       As String
                Dim lngNECKIN       As Long
                Dim arrARRGRPLIN()  As String
                Dim i               As Long
                Dim strDTENTREGA    As String
                
                strGRPCOD = ""
                intLINHA = 7
                strCODLIN = grdProduto.Cell(flexcpText, grdProduto.RowSel, conCOL_SonProd_CodCapacidade)
                
                intLINHA = 8
                lngNECKIN = grdProduto.Cell(flexcpText, grdProduto.RowSel, conCOL_SonProd_NECKIN)
                
                
                '' =========================
                intLINHA = 9
                sSql = ""
                
                sSql = "Select Distinct" & vbCrLf
                sSql = sSql & "       GRPI.*" & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADGRUPLINHAIT" & strNOMFILIAL & "  GRPI" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       GRPI.SGI_FILIAL         = " & FILIAL & vbCrLf
                sSql = sSql & "   And GRPI.SGI_CODLIN         = " & strCODLIN & vbCrLf
                sSql = sSql & "   And GRPI.SGI_OPTCOMNECKINSN = " & lngNECKIN & vbCrLf
                
                BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
                Do While Not BREC7.EOF()
                    strGRPCOD = strGRPCOD & BREC7!SGI_CODIGO
                    BREC7.MoveNext
                    If Not BREC7.EOF() Then strGRPCOD = strGRPCOD & ","
                Loop
                BREC7.Close
                
                '' Depois dar uma olhada
                '' Pega as Quantidades já Reservadas no Dia
                With grdProgEntrega
                    arrDIASCOTAS = Empty
                    lngSALDOQTDENTR = 0
                    If (.Rows - 1) Then
                        ReDim arrDIASCOTAS(1 To (.Rows - 1), 1 To 5) As String
                        For i = 1 To (.Rows - 1)
                            intLINHA = 10
                            arrDIASCOTAS(i, 1) = .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)
                            intLINHA = 11
                            arrDIASCOTAS(i, 2) = .Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd)
                            intLINHA = 12
                            arrDIASCOTAS(i, 3) = .Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto)
                            intLINHA = 13
                            arrDIASCOTAS(i, 4) = .Cell(flexcpText, i, conCOL_SonProgEntr_GrpPlanMestre)
                            intLINHA = 14
                            arrDIASCOTAS(i, 5) = .Cell(flexcpText, i, conCOL_SonProgEntr_IDINTERNO)
                        Next i
                    End If
                End With
                frmCADCOTAS.arrDIASCOTAS = arrDIASCOTAS
                
                intLINHA = 15
                frmCADCOTAS.FILIAL = FILIAL
                intLINHA = 16
                frmCADCOTAS.strNOMFILIAL = strNOMFILIAL
                intLINHA = 17
                frmCADCOTAS.strIDPRODUTO = .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_IdProduto)
                intLINHA = 18
                frmCADCOTAS.cTipOper = cTipOper
                intLINHA = 19
                frmCADCOTAS.strRETORNO = ""
                intLINHA = 20
                frmCADCOTAS.mskDTPED = mskDATAPED.Text
                intLINHA = 21
                frmCADCOTAS.strPRODCODLIN = grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_CodLinProd)
                intLINHA = 22
                frmCADCOTAS.intALTFILME = grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_AltFilme)
                intLINHA = 23
                frmCADCOTAS.intFOTNOVO = grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_FotNovo)
                intLINHA = 24
                frmCADCOTAS.lngCODPED = objCADPEDVENDA.CODPEDIDO
                intLINHA = 25
                frmCADCOTAS.strGRPCOD = Trim(Replace(strGRPCOD, ",", ""))
                intLINHA = 26
                frmCADCOTAS.lngSALDOQTDENTR = CLng(.Cell(flexcpText, .RowSel, conCOL_SonProgEntr_QtdProd))
                intLINHA = 27
                frmCADCOTAS.intHOMOLOGADO = grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_HOMOLOGADO)
                intLINHA = 28
                frmCADCOTAS.intAction2Do = .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_Action2Do)
                intLINHA = 29
                frmCADCOTAS.strIDINTERNO = .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_IDINTERNO)
                intLINHA = 30
                frmCADCOTAS.Show vbModal
                
                If Len(Trim(frmCADCOTAS.strRETORNO)) > 0 Then
                    intLINHA = 31
                    .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_Action2Do) = frmCADCOTAS.intAction2Do
                    intLINHA = 32
                    .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntrega) = frmCADCOTAS.strRETORNO
                    intLINHA = 33
                    If objCADPEDVENDA.STATUS <> "P" Then
                        .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_StatusOP) = frmCADCOTAS.intStatusOP
                    End If
                    intLINHA = 34
                    .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DescStatusOP) = PegaStatusOP(.Cell(flexcpText, .RowSel, conCOL_SonProgEntr_StatusOP))
                    If .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntrega) <> .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntregaBKP) Then
                        Call objBLBFunc.TrocaAction2Do(grdProgEntrega, Row, conCOL_SonProgEntr_Action2DoDtEntrega, .Cell(flexcpText, Row, conCOL_SonProgEntr_DataEntrega), .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntregaBKP))
                    End If
                End If
        
            End With
    
    End Select

    Exit Sub
    
Err_grdProgEntrega_CellButtonClick:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description & vbCrLf & "Linha : " & Format(intLINHA, "##00"), cTipOper, "Função : grdProgEntrega_CellButtonClick()", Me.Name, "grdProgEntrega_CellButtonClick()", strCAMARQERRO)
    
End Sub

Private Sub grdProgEntrega_Click()
    Call MostraDadosProgEntr
End Sub

Private Sub grdProgEntrega_GotFocus()
    If (grdProduto.Rows - 1) > 0 And grdProduto.RowSel > 0 Then
        grdProduto.Cell(flexcpBackColor, grdProduto.RowSel, conCOL_SonProd_IdProduto, grdProduto.RowSel, conCOL_SonProd_OS_Artes) = &H8000000D
        grdProduto.Cell(flexcpForeColor, grdProduto.RowSel, conCOL_SonProd_IdProduto, grdProduto.RowSel, conCOL_SonProd_OS_Artes) = &H8000000E
    End If
End Sub

Private Sub grdProgEntrega_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     
On Error GoTo Err_grdProgEntrega_KeyPressEdit
     
     With grdProgEntrega
          Select Case Col
                    Case conCOL_SonProgEntr_QtdProd
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_SonProgEntr_DataEntrega
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          End Select
     End With
     
     Exit Sub
     
Err_grdProgEntrega_KeyPressEdit:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProgEntrega_KeyPressEdit()", Me.Name, "grdProgEntrega_KeyPressEdit()", strCAMARQERRO)
     
End Sub

Private Sub grdProgEntrega_RowColChange()
    Call MostraDadosProgEntr
End Sub

Private Sub grdProgEntrega_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     
On Error GoTo Err_grdProgEntrega_ValidateEdit


     With grdProgEntrega
     
          Select Case Col
                 Case conCOL_SonProgEntr_QtdProd, conCOL_SonProgEntr_DataEntrega
                        If .EditText = Empty Then Exit Sub
                        Call objBLBFunc.TrocaAction2Do(grdProgEntrega, Row, conCOL_SonProgEntr_Action2Do, .Cell(flexcpText, Row, Col), .EditText)
                        
                        If Col = conCOL_SonProgEntr_QtdProd Then
                            ''If ConferePalhetsProgEntrg(Row, CLng(.EditText)) = False Then
                            ''    Cancel = True
                            ''    Exit Sub
                            ''End If
                        ElseIf Col = conCOL_SonProgEntr_DataEntrega Then
                            If Not IsDate(.EditText) Then
                                MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                                Cancel = True
                            Else
                                If Weekday(CDate(.EditText)) = 1 Then
                                    MsgBox "ATENÇÃO - Não é permitido date de Entrega no DOMINGO !!!", vbOKOnly + vbExclamation, "Aviso"
                                    Cancel = True
                                    Exit Sub
                                End If
                                If CDate(mskDATAPED.Text) > CDate(.EditText) Then
                                    MsgBox "A data de entrega não pode ser menor que a data do pedido !!!", vbOKOnly + vbExclamation, "Aviso"
                                    Cancel = True
                                    Exit Sub
                                ElseIf CDate(mskDATAPED.Text) = CDate(.EditText) Then
                                    MsgBox "ATENÇÂO - Data de entrega deve ser de 3 dias da data atual !!!", vbOKOnly + vbExclamation, "Aviso"
                                    Cancel = True
                                    Exit Sub
                                Else
                                    Dim lngQTDIAS As Long
                                    lngQTDIAS = (CDate(.EditText) - CDate(mskDATAPED.Text))
                                    If lngQTDIAS < 7 Then
                                        MsgBox "ATENÇÂO - Data de entrega deve ser de 3 dias da data atual !!!", vbOKOnly + vbExclamation, "Aviso"
                                        Cancel = True
                                        Exit Sub
                                    End If
                                End If
                                Call objBLBFunc.TrocaAction2Do(grdProgEntrega, Row, conCOL_SonProgEntr_Action2DoDtEntrega, .Cell(flexcpText, Row, Col), .EditText)
                            End If
                        End If
          End Select
     End With
    
    Exit Sub
    
Err_grdProgEntrega_ValidateEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdProgEntrega_ValidateEdit()", Me.Name, "grdProgEntrega_ValidateEdit()", strCAMARQERRO)

End Sub


Private Sub grdTIPREPROV_AfterEdit(ByVal Row As Long, ByVal Col As Long)

On Error GoTo Err_grdTIPREPROV_AfterEdit
     
     Dim i As Integer
     With grdTIPREPROV
          Select Case Col
                 Case conCOL_SonRep_Codigo
          End Select
          
     End With
     
     Exit Sub
     
Err_grdTIPREPROV_AfterEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdTIPREPROV_AfterEdit()", Me.Name, "grdTIPREPROV_AfterEdit()", strCAMARQERRO)

End Sub

Private Sub grdTIPREPROV_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

On Error GoTo Err_grdTIPREPROV_BeforeEdit
    
    Select Case Col
        Case conCOL_SonRep_Desc
             Cancel = True
        Case conCOL_SonRep_Codigo, _
             conCOL_SonRep_Pesq
             If cTipOper = "C" Then Cancel = True
        Case Else
            grdTIPREPROV.ComboList = ""
    End Select
    
    Exit Sub
    
Err_grdTIPREPROV_BeforeEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdTIPREPROV_BeforeEdit()", Me.Name, "grdTIPREPROV_BeforeEdit()", strCAMARQERRO)

End Sub

Private Sub grdTIPREPROV_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

On Error GoTo Err_grdTIPREPROV_CellButtonClick

    
    Select Case Col
        Case conCOL_SonRep_Pesq
    
            ReDim arrCAMPOS(1 To 2, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "        SGI_CODIGO " & vbCrLf
            sSql = sSql & "       ,SGI_DESCRI " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "        SGI_CADTIPREP " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "        SGI_FILIAL = " & FILIAL
            
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
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Tipo de Reprovação")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdTIPREPROV.Cell(flexcpText, Row, conCOL_SonRep_Codigo) = varRETORNO
               grdTIPREPROV.Cell(flexcpText, Row, conCOL_SonRep_Desc) = PegaDescrReprovacao(CLng(varRETORNO))
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(grdTIPREPROV, Row, conCOL_SonRep_Codigo, varRETORNO) = False Then
               MsgBox "Este tipo de reprovação já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdTIPREPROV.Cell(flexcpText, Row, conCOL_SonRep_Codigo) = ""
               grdTIPREPROV.Cell(flexcpText, Row, conCOL_SonRep_Desc) = ""
               Exit Sub
            End If
        
    End Select
    
    Exit Sub
    
    
Err_grdTIPREPROV_CellButtonClick:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdTIPREPROV_CellButtonClick()", Me.Name, "grdTIPREPROV_CellButtonClick()", strCAMARQERRO)

End Sub

Private Sub grdTIPREPROV_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

On Error GoTo Err_grdTIPREPROV_KeyPressEdit
     
     With grdTIPREPROV
          Select Case Col
                    Case conCOL_SonRep_Codigo
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
     
     Exit Sub
     
Err_grdTIPREPROV_KeyPressEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdTIPREPROV_KeyPressEdit()", Me.Name, "grdTIPREPROV_KeyPressEdit()", strCAMARQERRO)

End Sub

Private Sub grdTIPREPROV_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     
On Error GoTo Err_grdTIPREPROV_ValidateEdit
     
     With grdTIPREPROV
          Select Case Col
                 Case conCOL_SonRep_Codigo
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdTIPREPROV, Row, conCOL_SonRep_Codigo, .EditText) = False Then
                           MsgBox "Este tipo de reprovação já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonRep_Codigo) = ""
                           .Cell(flexcpText, Row, conCOL_SonRep_Desc) = ""
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrReprovacao(CLng(.EditText)))) = 0 Then
                           MsgBox "Esta reprovação não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonRep_Desc) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonRep_Codigo) = .EditText
                        
                        .EditText = Trim(Replace(Replace(.EditText, ",", ""), ".", ""))
                        .Cell(flexcpText, Row, conCOL_SonRep_Desc) = PegaDescrReprovacao(CLng(.EditText))
          End Select
     End With

     Exit Sub

Err_grdTIPREPROV_ValidateEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdTIPREPROV_ValidateEdit()", Me.Name, "grdTIPREPROV_ValidateEdit()", strCAMARQERRO)

End Sub

Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)

    Dim lngLIN As Long
    
    If cTipOper = "C" Or _
       cTipOper = "D" Or _
       cTipOper = "LF" Or _
       cTipOper = "LN" Or _
       cTipOper = "R" Or _
       cTipOper = "S" Or _
       cTipOper = "LS" Or _
       cTipOper = "LC" Then
       
        lngLIN = grdProduto.FindRow(grdProgEntrega.Cell(flexcpText, grdProgEntrega.RowSel, conCOL_SonProgEntr_IdProduto), , conCOL_SonProd_IdProduto)
        If lngLIN = -1 Then Exit Sub
        If grdProduto.Cell(flexcpText, lngLIN, conCOL_SonProd_StatusProd) = 2 Then
           Exit Sub
        End If
        
       If objCADPEDVENDA.STATUS <> "R" Then Exit Sub
    End If
    
    If (grdProgEntrega.Rows - 1) = 0 Then
        MsgBox "ATENÇÃO - A gride de programação de entrega não foi informado item !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If (grdProgEntrega.RowSel < 1) Then
        MsgBox "ATENÇÃO - Selecione uma linha da gride !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If Month(DateDblClicked) < Month(CDate(mskDATAPED.Text)) Then
        MsgBox "ATENÇÃO - O Mês Selecionado não pode ser memor que o Mês do Pedido !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If DateDblClicked < CDate(Format(Now, "DD/MM/YYYY")) Then
        MsgBox "ATENÇÃO - O dia selecionado é diferente da data vigente !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    With grdProgEntrega
    
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       PMDIA.SGI_TOTALPECAS" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO           PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO      LIMP" & vbCrLf
        sSql = sSql & "      ,SGI_CADGRUPLINHAIT" & strNOMFILIAL & " GRPI" & vbCrLf
        sSql = sSql & "      ,SGI_CADGRUPLINHA" & strNOMFILIAL & "   GLIN" & vbCrLf
        sSql = sSql & "      ,SGI_MAQULIN_MESANO" & strNOMFILIAL & " PM" & vbCrLf
        sSql = sSql & "      ,SGI_MAQULIN_CAPAC" & strNOMFILIAL & "  PMDIA" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO  = " & Trim(.Cell(flexcpText, .RowSel, conCOL_SonProgEntr_IdProduto)) & vbCrLf
        
        sSql = sSql & "   And LIMP.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And LIMP.SGI_CODLIN     = PROD.SGI_CODLINPROD" & vbCrLf
        
        sSql = sSql & "   And GRPI.SGI_FILIAL     = LIMP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And GRPI.SGI_CODLIN     = LIMP.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And GRPI.SGI_HOMOLOGSN  = " & grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_HOMOLOGADO) & vbCrLf
        
        sSql = sSql & "   And GLIN.SGI_FILIAL     = GRPI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And GLIN.SGI_CODIGO     = GRPI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And GLIN.SGI_ATIVO      = 1" & vbCrLf
        
        sSql = sSql & "   And PM.SGI_FILIAL       = GRPI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PM.SGI_CODIGO       = GRPI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PM.SGI_Mes          = " & Month(DateDblClicked) & vbCrLf
        sSql = sSql & "   And PM.SGI_Ano          = " & Year(DateDblClicked) & vbCrLf
        
        sSql = sSql & "   And PMDIA.SGI_FILIAL    = PM.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PMDIA.SGI_CODIGO    = PM.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PMDIA.SGI_CODINDICE = PM.SGI_INDICE" & vbCrLf
        sSql = sSql & "   And PMDIA.SGI_ATIVO     = 1" & vbCrLf
        sSql = sSql & "   And PMDIA.SGI_DATA      = '" & Format(DateDblClicked, "MM/DD/YYYY") & "'"
        
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
            .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntrega) = DateDblClicked
            
            If .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntrega) <> .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntregaBKP) Then
                If .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Ignore Then .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update
                Call objBLBFunc.TrocaAction2Do(grdProgEntrega, .RowSel, conCOL_SonProgEntr_Action2DoDtEntrega, .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntrega), .Cell(flexcpText, .RowSel, conCOL_SonProgEntr_DataEntregaBKP))
            End If
        Else
            MsgBox "ATENÇÂO - O dia " & DateDblClicked & " não está disponivel no calendário favor informar o PCP !!!", vbOKOnly + vbExclamation, "Aviso"
        End If
        BREC10.Close
        
    End With
    
End Sub

Private Sub mskDATAPED_GotFocus()
    
On Error GoTo Err_mskDATAPED_GotFocus
    
    objBLBFunc.SelecionaCampos mskDATAPED.Name, frmCADPEDVENDA

    Exit Sub
    
Err_mskDATAPED_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : mskDATAPED_GotFocus()", Me.Name, "mskDATAPED_GotFocus()", strCAMARQERRO)

End Sub


Private Sub mskDATAPED_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Not IsDate(mskDATAPED.Text) Then
        MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        Cancel = True
        Exit Sub
    End If
End Sub


Private Sub txtALIQICMS_GotFocus()
    objBLBFunc.SelecionaCampos txtALIQICMS.Name, frmCADPEDVENDA
End Sub

Private Sub txtALIQICMS_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtALIQICMS.Text
End Sub

Private Sub txtALIQICMS_Validate(Cancel As Boolean)
    If Len(Trim(txtALIQICMS.Text)) > 0 Then txtALIQICMS.Text = Format(txtALIQICMS.Text, "#,##0.00")
    Call CalcTotPedido
End Sub

Private Sub txtBAICOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtBAICOBR.Name, frmCADPEDVENDA
End Sub

Private Sub txtBAICOBR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtBAIENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtBAIENTR.Name, frmCADPEDVENDA
End Sub

Private Sub txtBAIENTR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCEPCOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtCEPCOBR.Name, frmCADPEDVENDA
End Sub

Private Sub txtCEPCOBR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCEPENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtCEPENTR.Name, frmCADPEDVENDA
End Sub

Private Sub txtCEPENTR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCIDCLIE_GotFocus()

On Error GoTo Err_txtCIDCLIE_GotFocus

    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, frmCADPEDVENDA

    Exit Sub
    
Err_txtCIDCLIE_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIE_GotFocus()", Me.Name, "txtCIDCLIE_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)
    
On Error GoTo Err_txtCIDCLIE_KeyPress
    
    objBLBFunc.SoNumeroPonto KeyAscii, txtCIDCLIE.Text

    Exit Sub
    
Err_txtCIDCLIE_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIE_KeyPress()", Me.Name, "txtCIDCLIE_KeyPress()", strCAMARQERRO)

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
    
    txtCIDCLIE.Text = Trim(Replace(Replace(txtCIDCLIE.Text, ",", ""), ".", ""))
    If Verifica_Credito = "N" Then
        txtCIDCLIE.Text = ""
        Exit Sub
    End If
    
    txtCODVEND.Text = Trim(Replace(Replace(txtCODVEND.Text, ",", ""), ".", ""))
    
    If ConfereCliente(txtCIDCLIE.Text, txtCODVEND.Text) = False Then
       txtCIDCLIE.Text = ""
       lblDescCliente.Caption = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente, "txtCIDCLIE_Validate()")
    If Len(Trim(lblDescCliente.Caption)) = 0 Then
       txtCIDCLIE.Text = ""
       lblDescCliente.Caption = ""
       Cancel = True
       Exit Sub
    End If
    Call DadosCliente(CInt(txtCIDCLIE.Text))

    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        objCADPEDVENDA.PERMITEFECHOP = PermiteFechamOP(txtCIDCLIE.Text)
    End If
    
    Exit Sub
    
Err_txtCIDCLIE_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIE_Validate()", Me.Name, "txtCIDCLIE_Validate()", strCAMARQERRO)

End Sub

Private Sub txtCIDCOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCOBR.Name, frmCADPEDVENDA
End Sub

Private Sub txtCIDCOBR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCIDENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDENTR.Name, frmCADPEDVENDA
End Sub

Private Sub txtCIDENTR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCodCondPgto_GotFocus()
    
On Error GoTo Err_txtCodCondPgto_GotFocus
    
    objBLBFunc.SelecionaCampos txtCodCondPgto.Name, frmCADPEDVENDA

    Exit Sub
    
Err_txtCodCondPgto_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCodCondPgto_GotFocus()", Me.Name, "txtCodCondPgto_GotFocus()", strCAMARQERRO)


End Sub

Private Sub txtCodCondPgto_KeyPress(KeyAscii As Integer)
    
On Error GoTo Err_txtCodCondPgto_KeyPress

    objBLBFunc.SoNumeroPonto KeyAscii, txtCodCondPgto.Text

    Exit Sub

Err_txtCodCondPgto_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCodCondPgto_KeyPress()", Me.Name, "txtCodCondPgto_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCodCondPgto_Validate(Cancel As Boolean)

On Error GoTo Err_txtCodCondPgto_Validate

    Dim i As Integer
    
    If Len(Trim(txtCodCondPgto.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodCondPgto.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCodCondPgto.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCodCondPgto.Text = Trim(Replace(Replace(txtCodCondPgto.Text, ",", ""), ".", ""))
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADCONDPGTO", txtCodCondPgto.Text, lblDescCondPgto, "txtCodCondPgto_Validate()")
    If Len(Trim(lblDescCondPgto.Caption)) = 0 Then
       txtCodCondPgto.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCodCondPgto_Validate:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCodCondPgto_Validate()", Me.Name, "txtCodCondPgto_Validate()", strCAMARQERRO)
    
End Sub

Private Sub DadosCliente(lngCODCLI As Long)

On Error GoTo Err_DadosCliente

    
    If BREC2.State = 1 Then BREC2.Close
    
    Dim i As Integer
    
    txtENDENTR.Text = ""
    txtBAIENTR.Text = ""
    txtCIDENTR.Text = ""
    cboESTENTR.ListIndex = -1
    txtCEPENTR.Text = ""
    
    txtENDCOBR.Text = ""
    txtBAICOBR.Text = ""
    txtCIDCOBR.Text = ""
    cboESTCOBR.ListIndex = -1
    txtCEPCOBR.Text = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO   = " & lngCODCLI & vbCrLf
    sSql = sSql & "   And SGI_DESBCLIE = 1"
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then
    
       If Not IsNull(BREC2!SGI_ENDENTREGA) Then txtENDENTR.Text = BREC2!SGI_ENDENTREGA
       If Not IsNull(BREC2!SGI_BAIENTREGA) Then txtBAIENTR.Text = BREC2!SGI_BAIENTREGA
       If Not IsNull(BREC2!SGI_CIDENTREGA) Then txtCIDENTR.Text = BREC2!SGI_CIDENTREGA
       If Not IsNull(BREC2!SGI_ESTENTREGA) Then cboESTENTR.ListIndex = (BREC2!SGI_ESTENTREGA - 1)
       If Not IsNull(BREC2!SGI_CEPENTREGA) Then txtCEPENTR.Text = BREC2!SGI_CEPENTREGA
       
       If Not IsNull(BREC2!SGI_ENDCOBRA) Then txtENDCOBR.Text = BREC2!SGI_ENDCOBRA
       If Not IsNull(BREC2!SGI_BAICOBRA) Then txtBAICOBR.Text = BREC2!SGI_BAICOBRA
       If Not IsNull(BREC2!SGI_CIDCOBRA) Then txtCIDCOBR.Text = BREC2!SGI_CIDCOBRA
       If Not IsNull(BREC2!SGI_ESTCOBRA) Then cboESTCOBR.ListIndex = (BREC2!SGI_ESTCOBRA - 1)
       If Not IsNull(BREC2!SGI_CEPCOBRA) Then txtCEPCOBR.Text = BREC2!SGI_CEPCOBRA
       
       objCADPEDVENDA.CODCLIE = BREC2!SGI_CODIGO
    
    End If
    BREC2.Close

    Exit Sub
    
Err_DadosCliente:

    If BREC2.State = 1 Then BREC2.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : DadosCliente()", Me.Name, "DadosCliente()", strCAMARQERRO)

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

    Dim i As Integer
    
    If Len(Trim(txtCODMOTLIQ.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODMOTLIQ.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODMOTLIQ.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCODMOTLIQ.Text = Trim(Replace(Replace(txtCODMOTLIQ.Text, ",", ""), ".", ""))
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMOTLIQ", txtCODMOTLIQ.Text, lblDescMotLiq, "txtCODMOTLIQ_Validate()")
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

Private Sub txtCODTRANSP_GotFocus()
   objBLBFunc.SelecionaCampos txtCODTRANSP.Name, frmCADPEDVENDA
End Sub

Private Sub txtCODTRANSP_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODTRANSP.Text
End Sub

Private Sub txtCODTRANSP_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODTRANSP_Validate

    Dim i As Integer
    
    If Len(Trim(txtCODTRANSP.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTRANSP.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTRANSP.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    txtCODTRANSP.Text = Trim(Replace(Replace(txtCODTRANSP.Text, ",", ""), ".", ""))
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTRANSP", txtCODTRANSP.Text, lblDescTransp, "txtCODTRANSP_Validate()")
    If Len(Trim(lblDescTransp.Caption)) = 0 Then
       txtCODTRANSP.Text = ""
       Cancel = True
    End If
        
    Exit Sub
    
Err_txtCODTRANSP_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODTRANSP_Validate()", Me.Name, "txtCODTRANSP_Validate()", strCAMARQERRO)
    
End Sub

Private Sub txtCODVEND_GotFocus()
    
On Error GoTo Err_txtCODVEND_GotFocus
    
    objBLBFunc.SelecionaCampos txtCODVEND.Name, frmCADPEDVENDA

    Exit Sub
    
Err_txtCODVEND_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVEND_GotFocus()", Me.Name, "txtCODVEND_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCODVEND_KeyPress(KeyAscii As Integer)
    
On Error GoTo Err_txtCODVEND_KeyPress

    objBLBFunc.SoNumeroPonto KeyAscii, txtCODVEND.Text

    Exit Sub
    
Err_txtCODVEND_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVEND_KeyPress()", Me.Name, "txtCODVEND_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCODVEND_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODVEND_Validate

    Dim i As Integer
    
    If Len(Trim(txtCODVEND.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVEND.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVEND.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCODVEND.Text = Trim(Replace(Replace(txtCODVEND.Text, ",", ""), ".", ""))
    
    Call PegaDescTabelasVend("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVEND.Text, lblDescVendedor, "txtCODVEND_Validate()")
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then
       txtCODVEND.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCODVEND_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVEND_Validate()", Me.Name, "txtCODVEND_Validate()", strCAMARQERRO)
    
    
End Sub

Private Sub txtCONTATO_GotFocus()
    objBLBFunc.SelecionaCampos txtCONTATO.Name, frmCADPEDVENDA
End Sub

Private Sub txtDEPARTAMENTO_GotFocus()
    objBLBFunc.SelecionaCampos txtDEPARTAMENTO.Name, frmCADPEDVENDA
End Sub

Private Sub txtEMAIL_GotFocus()
    objBLBFunc.SelecionaCampos txtDEPARTAMENTO.Name, frmCADPEDVENDA
End Sub

Private Sub txtENDCOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtENDCOBR.Name, frmCADPEDVENDA
End Sub

Private Sub txtENDCOBR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtENDENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtENDENTR.Name, frmCADPEDVENDA
End Sub

Private Sub txtENDENTR_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFAXCOBR_GotFocus()
   objBLBFunc.SelecionaCampos txtFAXCOBR.Name, frmCADPEDVENDA
End Sub

Private Sub txtFAXENTRE_GotFocus()
   objBLBFunc.SelecionaCampos txtFAXENTRE.Name, frmCADPEDVENDA
End Sub

Private Sub txtFRETE_GotFocus()
    objBLBFunc.SelecionaCampos txtFRETE.Name, frmCADPEDVENDA
End Sub

Private Sub txtFRETE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFRETE.Text
End Sub

Private Sub txtFRETE_Validate(Cancel As Boolean)
    If Len(Trim(txtFRETE.Text)) > 0 Then txtFRETE.Text = Format(txtFRETE.Text, "#,##0.00")
    Call CalcTotPedido
End Sub


Private Sub txtOBSROT_GotFocus()

On Error GoTo Err_txtOBSROT_GotFocus
    
    objBLBFunc.SelecionaCampos txtOBSROT.Name, Me

    Exit Sub
    
Err_txtOBSROT_GotFocus:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtOBSROT_GotFocus()", Me.Name, "txtOBSROT_GotFocus()", strCAMARQERRO)
    
End Sub

Private Sub txtOBSROT_Validate(Cancel As Boolean)
    
On Error GoTo Err_txtOBSROT_Validate
    
    Dim K As Integer
    
    With grdProduto
        If .RowSel = 0 Or (.Rows - 1) = 0 Then Exit Sub
        
        .Cell(flexcpText, .Row, conCOL_SonProd_OBSOP) = Trim(Replace(txtOBSROT.Text, ",", ""))
        For K = 1 To (grdProgEntrega.Rows - 1)
            If Trim(grdProgEntrega.Cell(flexcpText, K, conCOL_SonProgEntr_IdProduto)) = Trim(.Cell(flexcpText, .Row, conCOL_SonProd_IdProduto)) Then
                grdProgEntrega.Cell(flexcpText, K, conCOL_SonProgEntr_OBSOP) = Trim(Replace(txtOBSROT.Text, ",", ""))
            End If
        Next K
        
    End With
    
    Exit Sub
    
Err_txtOBSROT_Validate:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtOBSROT_Validate()", Me.Name, "txtOBSROT_Validate()", strCAMARQERRO)
    
    
End Sub

Private Sub txtORDCOMPCLI_GotFocus()
    objBLBFunc.SelecionaCampos txtORDCOMPCLI.Name, frmCADPEDVENDA
End Sub

Private Sub txtOutrDesp_GotFocus()
    objBLBFunc.SelecionaCampos txtFRETE.Name, frmCADPEDVENDA
End Sub

Private Sub txtOutrDesp_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFRETE.Text
End Sub

Private Sub txtOutrDesp_Validate(Cancel As Boolean)
    If Len(Trim(txtOutrDesp.Text)) > 0 Then txtOutrDesp.Text = Format(txtOutrDesp.Text, "#,##0.00")
    Call CalcTotPedido
End Sub

Private Sub txtPDESCTOTAL_GotFocus()
    objBLBFunc.SelecionaCampos txtPDESCTOTAL.Name, frmCADPEDVENDA
End Sub

Private Sub txtPDESCTOTAL_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPDESCTOTAL.Text
End Sub

Private Sub txtPDESCTOTAL_Validate(Cancel As Boolean)
    If Len(Trim(txtPDESCTOTAL.Text)) > 0 Then txtPDESCTOTAL.Text = Format(txtPDESCTOTAL.Text, "#,##0.00")
    Call CalcTotPedido
End Sub




Private Sub txtTELCOBR_GotFocus()
    objBLBFunc.SelecionaCampos txtTELCOBR.Name, frmCADPEDVENDA
End Sub

Private Sub txtTELENTR_GotFocus()
    objBLBFunc.SelecionaCampos txtTELENTR.Name, frmCADPEDVENDA
End Sub


Private Sub txtTIPPED_GotFocus()
    
On Error GoTo Err_txtTIPPED_GotFocus
    
    objBLBFunc.SelecionaCampos txtTIPPED.Name, frmCADPEDVENDA
    
    Exit Sub
    
Err_txtTIPPED_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtTIPPED_GotFocus()", Me.Name, "txtTIPPED_GotFocus()", strCAMARQERRO)

End Sub

Private Sub LimpaCamposCliente()

On Error GoTo Err_LimpaCamposCliente
    
    '' Limpando Campos
    txtCIDCLIE.Text = ""
    txtCodCondPgto.Text = ""
    If lngCodVendedor = 0 Then
        txtCODVEND.Text = ""
        txtTIPPED.Text = ""
    End If
    
    '' ----------------------------
    
    ''txtPRZENTREGA.Text = ""
    txtFRETE.Text = ""
    
    txtALIQICMS.Text = ""
    lblVLICMS.Caption = ""
    
    Exit Sub

Err_LimpaCamposCliente:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : LimpaCamposCliente()", Me.Name, "LimpaCamposCliente()", strCAMARQERRO)

End Sub

Private Sub txtTIPPED_KeyPress(KeyAscii As Integer)
    
On Error GoTo Err_txtTIPPED_KeyPress
    
    objBLBFunc.SoNumeroPonto KeyAscii, txtTIPPED.Text

    Exit Sub
    
Err_txtTIPPED_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtTIPPED_KeyPress()", Me.Name, "txtTIPPED_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtTIPPED_Validate(Cancel As Boolean)

On Error GoTo Err_txtTIPPED_Validate
    
    
    Dim i As Integer
    
    If Len(Trim(txtTIPPED.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtTIPPED.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtTIPPED.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtTIPPED.Text = Trim(Replace(Replace(txtTIPPED.Text, ",", ""), ".", ""))
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADESPORCA", txtTIPPED.Text, lblDescTpPed, "txtTIPPED_Validate()")
    If Len(Trim(lblDescTpPed.Caption)) = 0 Then
       txtTIPPED.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtTIPPED_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtTIPPED_Validate()", Me.Name, "txtTIPPED_Validate()", strCAMARQERRO)
    
End Sub

Private Function PegaPreco(strCODPROD As String) As Double

On Error GoTo Err_PegaPreco

    PegaPreco = 0
    
    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    
    If Len(Trim(txtCodCondPgto.Text)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_TABPRECO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPROD = '" & strCODPROD & "'" & vbCrLf
    sSql = sSql & "   And SGI_CODPGTO = " & txtCodCondPgto.Text & vbCrLf
    sSql = sSql & "   And SGI_VIGENTE = 'S' "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
       PegaPreco = BREC!SGI_VLVENDA
    Else
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = '" & strCODPROD & "'"
    
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          If Not IsNull(BREC2!SGI_PRECOPROD) Then PegaPreco = BREC2!SGI_PRECOPROD
       End If
       BREC2.Close
    
    End If
    BREC.Close
    
    Exit Function
    
Err_PegaPreco:

    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaPreco()", Me.Name, "PegaPreco()", strCAMARQERRO)

End Function
Private Function Valida_Campos() As Boolean

On Error GoTo Err_Valida_Campos
     
     Valida_Campos = False
     
     Dim i                  As Integer
     Dim j                  As Integer
     Dim intQtdItens        As Integer
     Dim lngIDProduto       As Long
     Dim curQtdTotal        As Currency
     Dim curQtdItens        As Currency
     Dim intQTDEXCLUIDOS    As Long
     Dim intTOTREGS         As Long
     Dim intALTFILME        As Integer
     Dim intNORMAL          As Integer
     Dim intFOTNOVO         As Integer
     Dim intFOTNORMAL       As Integer
     Dim intREPETICAO       As Integer
     Dim boolTODOSCONF      As Boolean
     Dim intLINHA           As Integer
     
     If cTipOper = "I" Or cTipOper = "A" Then
        If Len(Trim(txtCIDCLIE.Text)) = 0 Then
           MsgBox "O cliente não pode ser vázio !!!", vbOKOnly + vbCritical, "Aviso"
           txtCIDCLIE.SetFocus
           Exit Function
        End If
        If Len(Trim(txtCodCondPgto.Text)) = 0 Then
           MsgBox "A condição de pagamento não pode ser vázio !!!", vbOKOnly + vbCritical, "Aviso"
           txtCodCondPgto.SetFocus
           Exit Function
        End If
        If Len(Trim(txtCODVEND.Text)) = 0 Then
           MsgBox "O vendedor não pode ser vázio !!!", vbOKOnly + vbCritical, "Aviso"
           txtCODVEND.SetFocus
           Exit Function
        End If
        If Len(Trim(txtTIPPED.Text)) = 0 Then
           MsgBox "O tipo de pedido não pode ser vázio !!!", vbOKOnly + vbCritical, "Aviso"
           txtTIPPED.SetFocus
           Exit Function
        End If
        If (grdProduto.Rows - 1) = 0 Then
           MsgBox "Não Foi Informado Itens para o pedido !!!", vbOKOnly + vbCritical, "Aviso"
           Exit Function
        Else
           With grdProduto
                intQTDEXCLUIDOS = 0
                intTOTREGS = 0
                For i = 1 To (.Rows - 1)
                    If .Cell(flexcpText, i, conCOL_SonProd_Action2Do) = dacEnumUpdateAction_delete Then
                        intQTDEXCLUIDOS = (intQTDEXCLUIDOS + 1)
                    End If
                Next i
                If intQTDEXCLUIDOS > 0 Then
                    intTOTREGS = ((.Rows - 1) - intQTDEXCLUIDOS)
                    If intTOTREGS <= 0 Then
                        MsgBox "Não foi informado nenhum item no pedido !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                End If
           End With
        End If
        
        
        If Not IsDate(mskDATAPED.Text) Then
           MsgBox "Data do pedido inválido !!!", vbOKOnly + vbExclamation, "Aviso"
           mskDATAPED.SetFocus
           Exit Function
        End If
        If Len(Trim(txtCODTRANSP.Text)) = 0 Then
           MsgBox "A transportadora deve ser informada!!!", vbOKOnly + vbExclamation, "Aviso"
           txtCODTRANSP.SetFocus
           Exit Function
        End If
        
        '' ======================================================
        '' Validando  a Quantidade de Entrega e os prazos
        With grdProduto
            intALTFILME = 0
            intFOTNOVO = 0
            intNORMAL = 0
            intFOTNORMAL = 0
            intREPETICAO = 0
            For i = 1 To (.Rows - 1)
                If .Cell(flexcpText, i, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_delete Then
                    
                    If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_VlItens))) = 0 Then
                        MsgBox "Informe o valor unitário do Produto [ " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & " ] !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                    
                    If CLng(Replace(.Cell(flexcpText, i, conCOL_SonProd_VlItens), ",", "")) = 0 Then
                        MsgBox "Informe o valor unitário do Produto [ " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & " ] !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                    
                    If .Cell(flexcpData, i, conCOL_SonProd_FechTpFr) = 0 Then
                        MsgBox "Informe o Fechamento do Produto [ " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & " ] !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                    If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_AltFilme))) = 0 Then
                        MsgBox "Informe se ouve alteração no filme do rótulo [ " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & " ] !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                    If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_FotNovo))) = 0 Then
                        MsgBox "Informe se o fotolito é Novo Sim ou Não para o rótulo [ " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & " ] !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                    If Len(Trim(.Cell(flexcpText, i, conCOL_SonProd_Repeticao))) = 0 Then
                        MsgBox "Informe se a repetição no fotolito do rótulo [ " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & " ] !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                    
                    If .Cell(flexcpText, i, conCOL_SonProd_AltFilme) = 0 And .Cell(flexcpText, i, conCOL_SonProd_FotNovo) = 0 And .Cell(flexcpText, i, conCOL_SonProd_Repeticao) = 0 Then
                        MsgBox "ATENÇÂO - Não é permitido Alteração no Filme,Fotolito Novo,Repetição iqual a NÂO !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                    
                    If .Cell(flexcpText, i, conCOL_SonProd_AltFilme) = 1 Then intALTFILME = intALTFILME + 1
                    If .Cell(flexcpText, i, conCOL_SonProd_FotNovo) = 1 Then intFOTNOVO = intFOTNOVO + 1
                    If .Cell(flexcpText, i, conCOL_SonProd_Repeticao) = 1 Then intREPETICAO = intREPETICAO + 1
                                        
                    lngIDProduto = CLng(.Cell(flexcpText, i, conCOL_SonProd_IdProduto))
                    curQtdTotal = CCur(.Cell(flexcpText, i, conCOL_SonProd_QtdProd))
                    
                    intQtdItens = 0
                    curQtdItens = 0
                    For j = 1 To (grdProgEntrega.Rows - 1)
                        If grdProgEntrega.Cell(flexcpText, j, conCOL_SonProgEntr_IdProduto) = lngIDProduto And _
                           grdProgEntrega.Cell(flexcpText, j, conCOL_SonProgEntr_Action2Do) <> dacEnumUpdateAction_delete Then
                            intQtdItens = (intQtdItens + 1)
                            If Len(Trim(grdProgEntrega.Cell(flexcpText, j, conCOL_SonProgEntr_QtdProd))) = 0 Then
                                MsgBox "Para o produto - " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & ", não foi informado a qtde de entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                                Exit Function
                            End If
                            If .Cell(flexcpText, i, conCOL_SonProd_Repeticao) = 1 Then
                                If Len(Trim(grdProgEntrega.Cell(flexcpText, j, conCOL_SonProgEntr_DataEntrega))) = 0 Then
                                    MsgBox "Para o produto - " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & ", não foi informado a data de entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                                    Exit Function
                                End If
                            End If
                            curQtdItens = curQtdItens + CCur(grdProgEntrega.Cell(flexcpText, j, conCOL_SonProgEntr_QtdProd))
                        End If
                    Next j
                    
                    If intQtdItens = 0 Then
                        MsgBox "Para o produto - " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & ", não foi informado prazo de entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                    If (curQtdItens < curQtdTotal) Or _
                       (curQtdItens > curQtdTotal) Then
                        MsgBox "Para o produto - " & .Cell(flexcpText, i, conCOL_SonProd_Codigo) & ", a soma das qtde`s da entrega não e iqual a quantidade do produto !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Function
                    End If
                End If
            Next i
            If intALTFILME > 0 Then
                If intREPETICAO > 0 Then
                    MsgBox "ATENÇÃO - Não é permitido gravar Rótulos cujo filme será alterado junto com Rótulos que não irá sofrer alteração no filme !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
            End If
            If intFOTNOVO > 0 Then
                If intREPETICAO > 0 Then
                    MsgBox "ATENÇÃO - Não é permitido gravar Rótulos cujo Fotolito é novo junto com Rótulos que o Fotolito não é novo !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
            End If
        
        End With
     
        Dim intALTFILME2 As Integer
        Dim intFOTNOVO2  As Integer
        Dim lngPROD      As Long
        
        '' Programação de Entrega
        For i = 1 To (grdProgEntrega.Rows - 1)
           If grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) <> dacEnumUpdateAction_Ignore Then
                If grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 0 Then
                    lngPROD = grdProduto.FindRow(grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto), , conCOL_SonProd_IdProduto)
                    intALTFILME2 = 0
                    intFOTNOVO2 = 0
                    If lngPROD > -1 Then
                        intALTFILME2 = grdProduto.Cell(flexcpText, lngPROD, conCOL_SonProd_AltFilme)
                        intFOTNOVO2 = grdProduto.Cell(flexcpText, lngPROD, conCOL_SonProd_FotNovo)
                    End If
                    If intALTFILME = 0 And intFOTNOVO = 0 Then
                       ''If ConfereCotas(grdProgEntrega.Cell(flexcpText, I, conCOL_SonProgEntr_IdProduto), grdProgEntrega.Cell(flexcpText, I, conCOL_SonProgEntr_DataEntrega)) = True Then Exit Function
                    End If
                End If
           End If
        Next i
     
     End If
     
     '' Para liberação Comercial e Financeira
     If objCADPEDVENDA.STATUS = "R" Or _
        objCADPEDVENDA.STATUS = "B" Or _
        objCADPEDVENDA.STATUS = "N" Then
        If cTipOper = "LF" Or cTipOper = "LN" Then
            With grdProgEntrega
                For i = 1 To (.Rows - 1)
                    If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_CodOP))) = 0 Then
                        If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega))) = 0 Then
                            MsgBox "ATENÇÃO" & vbCrLf & "Não é possivel liberar o pedido pois a data de entrega esta vázia !!!", vbOKOnly + vbExclamation, "Aviso"
                            Exit Function
                        End If
                        If CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)) < Date Then
                            MsgBox "ATENÇÃO" & vbCrLf & "Não é possivel liberar o pedido pois a data de entrega será menor que a data de criação da OP !!!", vbOKOnly + vbExclamation, "Aviso"
                            Exit Function
                        End If
                    End If
                Next i
            End With
        End If
     End If
     
     
     If cTipOper = "R" Then
        If (grdTIPREPROV.Rows - 1) = 0 Then
           MsgBox "Informe pelo menos 1 tipo de reprovação !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
     ElseIf cTipOper = "LV" Then
        For i = 1 To (grdProduto.Rows - 1)
            If grdProduto.Cell(flexcpText, i, conCOL_SonProd_StatusProd) = 2 Then
                MsgBox "ATENÇÃO - Este produto ainda não foi liberado pelo departamento de Artes impossivel Liberar !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        Next i
        
        '' Verificando se existe Data em Branco
        For i = 1 To (grdProgEntrega.Rows - 1)
            If Not IsDate(grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)) Then
                intLINHA = grdProduto.FindRow(grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto), , conCOL_SonProd_IdProduto)
                If intLINHA > -1 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "O Rótulo : " & Trim(grdProduto.Cell(flexcpText, intLINHA, conCOL_SonProd_DescProd)) & vbCrLf & _
                           "Não foi informado data de entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                End If
                Exit Function
            End If
        Next i
     End If
     
     '' Depois retornar
     '' Verificando o Chek-List
     If cTipOper = "LN" Or cTipOper = "LF" Then
        If chkVerificado.Value = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Não foi conferido a Condição de Pagamento !!!", vbOKOnly + vbExclamation, "Aviso"
            stCAMPOSVENDA.Tab = 0
            Exit Function
        End If
     
        For i = 1 To (grdProduto.Rows - 1)
            If grdProduto.Cell(flexcpText, i, conCOL_SonProd_Conferido) = 0 Then
                MsgBox "ATENÇÃO" & vbCrLf & _
                       "Atenção, o Rótulo [ " & grdProduto.Cell(flexcpText, i, conCOL_SonProd_Codigo) & " ] não foi conferido !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        Next i
     End If
     
     '' Verifica P.Cota e P.Data
     If cTipOper = "LC" Then
        Dim intLINHAPROD As Integer
        With grdProgEntrega
            For i = 1 To (.Rows - 1)
                intLINHAPROD = grdProduto.FindRow(.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto), , conCOL_SonProd_IdProduto)
                If .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 7 Then
                    MsgBox "ATENÇÃO" & vbCrLf & "O Rótulo [ " & grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_Codigo) & " ]  ainda não foi liberado do P.Data !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                ElseIf .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 6 Then
                    MsgBox "ATENÇÃO" & vbCrLf & "O Rótulo [ " & grdProduto.Cell(flexcpText, intLINHAPROD, conCOL_SonProd_Codigo) & " ]  ainda não foi liberado do P.Cota !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
            Next i
        End With
     End If
     
     
    '' Liquidação de Pedido
    If cTipOper = "M" Then
        If Len(Trim(txtCODMOTLIQ.Text)) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & "O Código de Motivo de Liquidação de Pedido é inválido !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        If Len(Trim(txtOBS_MotLiq.Text)) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & "A Observação de Motivo de Liquidação de Pedido é inválido !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End If
     
     
     Valida_Campos = True

    Exit Function
    
Err_Valida_Campos:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Valida_Campos()", Me.Name, "Valida_Campos()", strCAMARQERRO)

End Function

Private Sub Consulta()

On Error GoTo Err_Consulta
    
    Dim i As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = False
    
    If BREC.State = 1 Then BREC.Close
    
    Me.Caption = "Cadastro de Pedido de Venda - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos Me

    stCAMPOSVENDA.Tab = 0
   
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame8.Enabled = True
    Frame9.Enabled = False
    Frame27.Visible = True
    Frame28.Enabled = False
    
    Frame13.Visible = True
    txtOBS2.Locked = True
   
    objBLBFunc.Preenche_Estado cboESTENTR
    objBLBFunc.Preenche_Estado cboESTCOBR
    
    objCADPEDVENDA.CODPEDIDO = iCodigo
    
    Call InitGridReprovacao
    Call InitGridProd
    Call InitGridProg
    Call InitGridOrdemFat
    Call InitGridConfFat
    Call InitGridLogPed
    Call InitGridProducao
    
    
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    txtOutrDesp.Text = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    
    '' --------------------
    lblTotalItens.Caption = ""
    
    '' --------------------
    '' Desconto
    lblVLDESCONTO.Caption = ""
    '' --------------------
    
    lblDescMotLiq.Caption = ""
    txtOBS_MotLiq.Locked = True
    
    Call LimpaCamposLabel
    Call LimpaCamposDadosAdicionais
    Call LimpaCampoSaldoRot
    Call LimpaSaldoPedido
    
    objCADPEDVENDA.FILIALPED = intFILIALPED
    
    optESPECIAL(0).Value = True
    
    Call AbilDesConferido(False, 0)
    
    If objCADPEDVENDA.Carrega_Campos = True Then
    
       If (objCADPEDVENDA.STATUS = "L" Or objCADPEDVENDA.STATUS = "N") Then
          If objCADPEDVENDA.STATUS = "L" Then lblSTATUS.Caption = "LIBERADO FINANCEIRO"
          If objCADPEDVENDA.STATUS = "N" Then lblSTATUS.Caption = "LIBERADO COMERCIAL"
          cmdAltera.Enabled = False
       End If
       If objCADPEDVENDA.STATUS = "B" Or objCADPEDVENDA.STATUS = "S" Then
          lblSTATUS.Caption = "BLOQUEADO"
          cmdAltera.Enabled = True
       End If
       If objCADPEDVENDA.STATUS = "R" Then
          lblSTATUS.Caption = "REPROVADO"
          cmdAltera.Enabled = False
          stCAMPOSVENDA.TabVisible(2) = True
       End If
       If objCADPEDVENDA.STATUS = "F" Then
          lblSTATUS.Caption = "FATURADO TOTAL"
          cmdAltera.Enabled = False
       End If
       If objCADPEDVENDA.STATUS = "P" Then
          lblSTATUS.Caption = "FATURADO PARCIAL"
          cmdAltera.Enabled = True
       End If
       If objCADPEDVENDA.STATUS = "V" Then
          lblSTATUS.Caption = "AGUARDANDO LIBERAÇÃO DE ARTES"
          cmdAltera.Enabled = False
       End If
       If objCADPEDVENDA.STATUS = "X" Then
          lblSTATUS.Caption = "PARA ESTOQUE"
          cmdAltera.Enabled = False
       End If
       If objCADPEDVENDA.STATUS = "C" Then
          lblSTATUS.Caption = "BLOQUEADO POR P.COTA/P.DATA"
          cmdAltera.Enabled = True
       End If
       If objCADPEDVENDA.STATUS = "4" Then
          lblSTATUS.Caption = "BLOQUEADO POR P.COTA/P.DATA"
          cmdAltera.Enabled = True
       End If
       If objCADPEDVENDA.STATUS = "M" Then
          lblSTATUS.Caption = "LIQUIDADO MANUALMENTE"
          cmdAltera.Enabled = False
       End If
       
       lblCODIGO.Caption = objCADPEDVENDA.CODPEDIDO
       mskDATAPED.Text = Format(objCADPEDVENDA.DATAPED, "DD/MM/YYYY")
       If mskDATAPED.Text = "30/12/1899" Then mskDATAPED.Text = "__/__/____"
       txtCIDCLIE.Text = objCADPEDVENDA.CODCLIE
       
       txtCodCondPgto.Text = objCADPEDVENDA.CODCONDPGTO
       
       txtCODVEND.Text = objCADPEDVENDA.CODVEND
       
       txtTIPPED.Text = objCADPEDVENDA.TIPPED
       
       txtOBSERVACAO.Text = objCADPEDVENDA.OBSERVACAO
       txtOBS2.Text = objCADPEDVENDA.OBS2
       
       '' Dados de Entrega
       txtENDENTR.Text = objCADPEDVENDA.ENDENTR
       txtBAIENTR.Text = objCADPEDVENDA.BAIENTR
       txtCIDENTR.Text = objCADPEDVENDA.CIDENTR
       If objCADPEDVENDA.ESTENTREGA > 0 Then cboESTENTR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTENTREGA)
       txtCEPENTR.Text = objCADPEDVENDA.CEPENTREGA
       txtTELENTR.Text = objCADPEDVENDA.TELENTR
       txtFAXENTRE.Text = objCADPEDVENDA.FAXENTR
       
       '' Dados de Cobrança
       txtENDCOBR.Text = objCADPEDVENDA.ENDCOBRA
       txtBAICOBR.Text = objCADPEDVENDA.BAICOBRA
       txtCIDCOBR.Text = objCADPEDVENDA.CIDCOBRA
       If objCADPEDVENDA.ESTCOBRA > 0 Then cboESTCOBR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTCOBRA)
       txtCEPCOBR.Text = objCADPEDVENDA.CEPCOBRA
       txtTELCOBR.Text = objCADPEDVENDA.TELCOBRA
       txtFAXCOBR.Text = objCADPEDVENDA.FAXCOBRA
       
       '' Diversos
       ''If objCADPEDVENDA.PRZENTREGA > 0 Then txtPRZENTREGA.Text = objCADPEDVENDA.PRZENTREGA
       
       txtCODTRANSP.Text = objCADPEDVENDA.CODTRANSP
       txtORDCOMPCLI.Text = objCADPEDVENDA.ORDCOMPCLI
       txtCONTATO.Text = objCADPEDVENDA.CONTATO
       txtDEPARTAMENTO.Text = objCADPEDVENDA.DEPARTAMENTO
       txtEMAIL.Text = objCADPEDVENDA.EMAIL
       
       
       '' LOQ ( Motivos de Liquidação )
       txtCODMOTLIQ.Text = objCADPEDVENDA.CODMOTLIQ
       txtOBS_MotLiq.Text = objCADPEDVENDA.OBSLIQ
       
       optESPECIAL(objCADPEDVENDA.ESPECIAL).Value = True
       optPARAESTOQUE(objCADPEDVENDA.PARAESTOQUE).Value = True
       
       '' Totais
       If objCADPEDVENDA.VALBASICMS > 0 Then lblBASICMS.Caption = Format(objCADPEDVENDA.VALBASICMS, "#,##0.00")
       If objCADPEDVENDA.ALIQICMS > 0 Then txtALIQICMS.Text = Format(objCADPEDVENDA.ALIQICMS, "#,##0.00")
       If objCADPEDVENDA.VLICMS > 0 Then lblVLICMS.Caption = Format(objCADPEDVENDA.VLICMS, "#,##0.00")
       If objCADPEDVENDA.OUTRDESPESAS > 0 Then txtOutrDesp.Text = Format(objCADPEDVENDA.OUTRDESPESAS, "#,##0.00")
       If objCADPEDVENDA.VLFRETE > 0 Then txtFRETE.Text = Format(objCADPEDVENDA.VLFRETE, "#,##0.00")
       If objCADPEDVENDA.VLIPI > 0 Then lblVLIPI.Caption = Format(objCADPEDVENDA.VLIPI, "#,##0.00")
       If objCADPEDVENDA.VLDESCTO > 0 Then lblVLDESCONTO.Caption = Format(objCADPEDVENDA.VLDESCTO, "#,##0.00")
       If objCADPEDVENDA.VALDESC > 0 Then txtPDESCTOTAL.Text = Format(objCADPEDVENDA.VALDESC, "#,##0.00")
       If objCADPEDVENDA.PORDESC > 0 Then txtVLDESCTOTOT.Text = Format(objCADPEDVENDA.PORDESC, "#,##0.00")
       If objCADPEDVENDA.TOTORCTO > 0 Then lblVLTOTAL.Caption = Format(objCADPEDVENDA.TOTORCTO, "#,##0.00")
       
       objCADPEDVENDA.PERMITEFECHOP = objCADPEDVENDA.PERMITEFECHOP
       
       chkVerificado.Value = objCADPEDVENDA.CONFERIDO
       
       Call PopGrdProdutos
       Call PegaDadosLabel
       Call PopLogPedidos
       
       
       Call MostraDadosReprovacao
       
       Call CarregaPlanoEntrega
       Call GeraSaldoPedido(Str(objCADPEDVENDA.CODPEDIDO))
       
       
       '' Verifica se Existe Ordem de Faturamento para a tivar a Tab.
       
       grdProduto.Row = 1
       grdProduto.Col = 1
    
       Call VisualizaBotoesPCD(objCADPEDVENDA.STATUS, cTipOper)
       Call VisualizaBotoesLibAlteracao(objCADPEDVENDA.STATUS, cTipOper)
       Call VisualizaBotoesLibFotolito(objCADPEDVENDA.STATUS, cTipOper)
       Call VisualizaBotoesLibComercial(objCADPEDVENDA.STATUS, cTipOper)
       Call VisualizaBotoesLibFinanceira(objCADPEDVENDA.STATUS, cTipOper)
       
    End If
        
    Exit Sub
    
Err_Consulta:

    If BREC.State = 1 Then BREC.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Consulta()", Me.Name, "Consulta()", strCAMARQERRO)
        
End Sub

Private Sub Altera()

On Error GoTo Err_Altera
    
    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Pedido de Venda - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos Me

    stCAMPOSVENDA.Tab = 0
    stCAMPOSVENDA.TabEnabled(2) = False
        
    Frame3.Enabled = True
    Frame4.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True
    txtOBS2.Locked = False
    Frame13.Visible = True
    Frame27.Visible = False
    
    txtCIDCLIE.Enabled = False
    Command1.Enabled = False
    
    txtTIPPED.Enabled = False
    Command3.Enabled = False
    
    
    objBLBFunc.Preenche_Estado cboESTENTR
    objBLBFunc.Preenche_Estado cboESTCOBR
    
    objCADPEDVENDA.CODPEDIDO = iCodigo
    
    Call InitGridReprovacao
    Call InitGridProd
    Call InitGridProg
    Call InitGridOrdemFat
    Call InitGridLogPed
    Call InitGridProducao
    
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    txtOutrDesp.Text = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    lblTotalItens.Caption = ""
   
    '' --------------------
    '' Desconto
    lblVLDESCONTO.Caption = ""
    '' --------------------
    
    Call LimpaCamposLabel
    Call LimpaCamposDadosAdicionais
    Call LimpaCampoSaldoRot
    Call LimpaSaldoPedido
    
    Call Pega_Vendedor(lngCodVendedor)
    Call DesativasCampos
    
    Call AbilDesConferido(False, 0)
    
    txtCIDCLIE.Enabled = False
    Command1.Enabled = False
    
    objCADPEDVENDA.FILIALPED = intFILIALPED
    
    optESPECIAL(0).Value = True
    
    If objCADPEDVENDA.Carrega_Campos = False Then Exit Sub
    
    If objCADPEDVENDA.STATUS = "L" Or objCADPEDVENDA.STATUS = "N" Then lblSTATUS.Caption = "LIBERADO"
    If objCADPEDVENDA.STATUS = "B" Or objCADPEDVENDA.STATUS = "S" Then
       lblSTATUS.Caption = "BLOQUEADO"
    End If
    If objCADPEDVENDA.STATUS = "V" Then lblSTATUS.Caption = "AGUARDANDO LIBERAÇÃO DE ARTES"
    If objCADPEDVENDA.STATUS = "C" Then lblSTATUS.Caption = "BLOQUEADO POR COTA"
    If objCADPEDVENDA.STATUS = "4" Then lblSTATUS.Caption = "BLOQUEADO POR DATA"
    If objCADPEDVENDA.STATUS = "P" Then
        lblSTATUS.Caption = "FATURADO PARCIAL"
        cmdAltera.Enabled = False
    End If
        
    lblCODIGO.Caption = objCADPEDVENDA.CODPEDIDO
    mskDATAPED.Text = Format(objCADPEDVENDA.DATAPED, "DD/MM/YYYY")
    
    txtCIDCLIE.Text = objCADPEDVENDA.CODCLIE
    ''mskDTENTREGA.Text = Format(objCADPEDVENDA.DATAENTREGA, "DD/MM/YYYY")
    
    txtCodCondPgto.Text = objCADPEDVENDA.CODCONDPGTO
    
    txtCODVEND.Text = objCADPEDVENDA.CODVEND
    
    txtOBSERVACAO.Text = objCADPEDVENDA.OBSERVACAO
    txtOBS2.Text = objCADPEDVENDA.OBS2
    
    txtTIPPED.Text = objCADPEDVENDA.TIPPED
    
    '' Dados de Entrega
    txtENDENTR.Text = objCADPEDVENDA.ENDENTR
    txtBAIENTR.Text = objCADPEDVENDA.BAIENTR
    txtCIDENTR.Text = objCADPEDVENDA.CIDENTR
    If objCADPEDVENDA.ESTENTREGA > 0 Then cboESTENTR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTENTREGA)
    txtCEPENTR.Text = objCADPEDVENDA.CEPENTREGA
    txtTELENTR.Text = objCADPEDVENDA.TELENTR
    txtFAXENTRE.Text = objCADPEDVENDA.FAXENTR
    
    '' Dados de Cobrança
    txtENDCOBR.Text = objCADPEDVENDA.ENDCOBRA
    txtBAICOBR.Text = objCADPEDVENDA.BAICOBRA
    txtCIDCOBR.Text = objCADPEDVENDA.CIDCOBRA
    If objCADPEDVENDA.ESTCOBRA > 0 Then cboESTCOBR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTCOBRA)
    txtCEPCOBR.Text = objCADPEDVENDA.CEPCOBRA
    txtTELCOBR.Text = objCADPEDVENDA.TELCOBRA
    txtFAXCOBR.Text = objCADPEDVENDA.FAXCOBRA
    
    '' Diversos
    ''If objCADPEDVENDA.PRZENTREGA > 0 Then txtPRZENTREGA.Text = objCADPEDVENDA.PRZENTREGA
    
    txtCODTRANSP.Text = objCADPEDVENDA.CODTRANSP
    txtORDCOMPCLI.Text = objCADPEDVENDA.ORDCOMPCLI
    txtCONTATO.Text = objCADPEDVENDA.CONTATO
    txtDEPARTAMENTO.Text = objCADPEDVENDA.DEPARTAMENTO
    txtEMAIL.Text = objCADPEDVENDA.EMAIL
    
    optESPECIAL(objCADPEDVENDA.ESPECIAL).Value = True
    optPARAESTOQUE(objCADPEDVENDA.PARAESTOQUE).Value = True
    
    '' Totais
    If objCADPEDVENDA.VALBASICMS > 0 Then lblBASICMS.Caption = Format(objCADPEDVENDA.VALBASICMS, "#,##0.00")
    If objCADPEDVENDA.ALIQICMS > 0 Then txtALIQICMS.Text = Format(objCADPEDVENDA.ALIQICMS, "#,##0.00")
    If objCADPEDVENDA.VLICMS > 0 Then lblVLICMS.Caption = Format(objCADPEDVENDA.VLICMS, "#,##0.00")
    If objCADPEDVENDA.OUTRDESPESAS > 0 Then txtOutrDesp.Text = Format(objCADPEDVENDA.OUTRDESPESAS, "#,##0.00")
    If objCADPEDVENDA.VLFRETE > 0 Then txtFRETE.Text = Format(objCADPEDVENDA.VLFRETE, "#,##0.00")
    If objCADPEDVENDA.VLIPI > 0 Then lblVLIPI.Caption = Format(objCADPEDVENDA.VLIPI, "#,##0.00")
    If objCADPEDVENDA.VLDESCTO > 0 Then lblVLDESCONTO.Caption = Format(objCADPEDVENDA.VLDESCTO, "#,##0.00")
    If objCADPEDVENDA.VALDESC > 0 Then txtPDESCTOTAL.Text = Format(objCADPEDVENDA.VALDESC, "#,##0.00")
    If objCADPEDVENDA.PORDESC > 0 Then txtVLDESCTOTOT.Text = Format(objCADPEDVENDA.PORDESC, "#,##0.00")
    If objCADPEDVENDA.TOTORCTO > 0 Then lblVLTOTAL.Caption = Format(objCADPEDVENDA.TOTORCTO, "#,##0.00")
    
    chkVerificado.Value = objCADPEDVENDA.CONFERIDO
    
    objCADPEDVENDA.PERMITEFECHOP = objCADPEDVENDA.PERMITEFECHOP
    
    Call PopGrdProdutos
    Call PopLogPedidos
    
    Call CarregaPlanoEntrega
    Call GeraSaldoPedido(Str(objCADPEDVENDA.CODPEDIDO))
    Call PegaDadosLabel


    If objCADPEDVENDA.STATUS = "V" Then
        stCAMPOSVENDA.Tab = 0
        With grdProgEntrega
            For i = 1 To (.Rows - 1)
                .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega) = Empty
            Next i
        End With
    End If


    Call VisualizaBotoesPCD(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibAlteracao(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibFotolito(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibComercial(objCADPEDVENDA.STATUS, cTipOper)
    Call VisualizaBotoesLibFinanceira(objCADPEDVENDA.STATUS, cTipOper)

    Exit Sub
    
Err_Altera:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Altera()", Me.Name, "Altera()", strCAMARQERRO)

End Sub


Private Function Verifica_Credito() As String
    
On Error GoTo Err_Verifica_Credito
    
    If BREC.State = 1 Then BREC.Close
    
    Dim curLIMCRED     As Currency
    Dim strBLOQPED     As String
    Dim curVLTITABERTO As Currency
    Dim intRESP        As Integer
    
    curLIMCRED = 0
    curVLTITABERTO = 0
    strBLOQPED = "N"
    
    ''Verifica_Credito = "B"
    Verifica_Credito = "S"
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_VLLIMCRED    " & vbCrLf
    sSql = sSql & "      ,SGI_CREDSBLQPED  " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & txtCIDCLIE.Text
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       curLIMCRED = BREC!SGI_VLLIMCRED
       strBLOQPED = BREC!SGI_CREDSBLQPED
    End If
    BREC.Close
    
    '' -----------------------------------------------------
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SUM(ITENS.SGI_VLDOC) As SGI_SALDO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CONTASHARC CABEC " & vbCrLf
    sSql = sSql & "       ,SGI_CONTASIARC ITENS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CABEC.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And CABEC.SGI_CODCLI = " & txtCIDCLIE.Text
    sSql = sSql & "   And ITENS.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ITENS.SGI_CODIGO = CABEC.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And ITENS.SGI_VLPAGO IS NULL "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       If Not IsNull(BREC!SGI_SALDO) Then curVLTITABERTO = BREC!SGI_SALDO
    End If
    BREC.Close
    
    '' Regras para para crédito
    ''If (curVLTITABERTO >= curLIMCRED) And curVLTITABERTO > 0 And curLIMCRED > 0 Then
    ''   If strBLOQPED = "N" Then
    ''      intRESP = MsgBox("Atenção este Cliente está com o limite estourado, bloqueia para análize ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
    ''      If intRESP = 6 Then Verifica_Credito = "B"
    ''      If intRESP = 7 Then Verifica_Credito = "L"
    ''   End If
    ''   If strBLOQPED = "S" Then
    ''      MsgBox "Atenção este Cliente está com o limite estourado, pedido será bloqueado para análize", vbOKOnly + vbInformation, "Aviso"
    ''      Verifica_Credito = "B"
    ''   End If
    ''End If
    '' ------------------------
    
    If strBLOQPED = "S" Then
       MsgBox "Atenção este Cliente está bloqueado por motivos financeiros !!!", vbOKOnly + vbInformation, "Aviso"
       Verifica_Credito = "N"
    End If
    
    ''Verifica_Credito = "B"
    
    Exit Function
    
Err_Verifica_Credito:
    
    If BREC.State = 1 Then BREC.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Altera()", Me.Name, "Altera()", strCAMARQERRO)
    
End Function

Private Sub Libera()

On Error GoTo Err_Libera

    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Pedido de Venda - [ LIBERAÇÃO ]"
    If cTipOper = "LV" Then Me.Caption = "Cadastro de Pedido de Venda - [ LIBERAÇÃO DE FOTOLITO ]"
    If cTipOper = "LN" Then Me.Caption = "Cadastro de Pedido de Venda - [ LIBERAÇÃO COMERCIAL ]"
    If cTipOper = "LF" Then Me.Caption = "Cadastro de Pedido de Venda - [ LIBERAÇÃO FINANCEIRA ]"
    
    objBLBFunc.LimpaCampos Me

    stCAMPOSVENDA.Tab = 0
    stCAMPOSVENDA.TabVisible(2) = True
    
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame8.Enabled = True
    Frame9.Enabled = False
    Frame27.Visible = False
    
    Frame13.Visible = True
    txtOBS2.Locked = True
    
    objBLBFunc.Preenche_Estado cboESTENTR
    objBLBFunc.Preenche_Estado cboESTCOBR
    
    objCADPEDVENDA.CODPEDIDO = iCodigo
    
    Call InitGridReprovacao
    Call InitGridProd
    Call InitGridProg
    Call InitGridOrdemFat
    Call InitGridLogPed
    Call InitGridProducao
    
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    txtOutrDesp.Text = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    
    Call LimpaCamposLabel
    Call LimpaCamposDadosAdicionais
    Call LimpaCampoSaldoRot
    Call LimpaSaldoPedido
    
    '' --------------------
    '' Desconto
    lblVLDESCONTO.Caption = ""
    '' --------------------
   
    objCADPEDVENDA.FILIALPED = intFILIALPED
   
    optESPECIAL(0).Value = True
   
    Call AbilDesConferido(False, 0)
   
    If objCADPEDVENDA.Carrega_Campos = True Then
    
       If objCADPEDVENDA.STATUS = "L" Then lblSTATUS.Caption = "LIBERADO"
       If objCADPEDVENDA.STATUS = "N" Then lblSTATUS.Caption = "LIBERADO COMERCIAL"
       If objCADPEDVENDA.STATUS = "B" Then lblSTATUS.Caption = "BLOQUEADO"
       If objCADPEDVENDA.STATUS = "S" Then lblSTATUS.Caption = "BLOQUEADO"
       If objCADPEDVENDA.STATUS = "V" Then
          lblSTATUS.Caption = "AGUARDANDO LIBERAÇÃO DE ARTES"
       End If
       If objCADPEDVENDA.STATUS = "C" Then lblSTATUS.Caption = "BLOQUEADO POR P.COTA ou P.DATA"
       If objCADPEDVENDA.STATUS = "4" Then lblSTATUS.Caption = "BLOQUEADO POR DATA"
       If objCADPEDVENDA.STATUS = "R" Then lblSTATUS.Caption = "REPROVADO"
       
       lblCODIGO.Caption = objCADPEDVENDA.CODPEDIDO
       mskDATAPED.Text = Format(objCADPEDVENDA.DATAPED, "DD/MM/YYYY")
       
       txtCIDCLIE.Text = objCADPEDVENDA.CODCLIE
       
       txtCodCondPgto.Text = objCADPEDVENDA.CODCONDPGTO
       txtCODVEND.Text = objCADPEDVENDA.CODVEND
       txtTIPPED.Text = objCADPEDVENDA.TIPPED
       txtOBSERVACAO.Text = objCADPEDVENDA.OBSERVACAO
       
       '' Dados de Entrega
       txtENDENTR.Text = objCADPEDVENDA.ENDENTR
       txtBAIENTR.Text = objCADPEDVENDA.BAIENTR
       txtCIDENTR.Text = objCADPEDVENDA.CIDENTR
       If objCADPEDVENDA.ESTENTREGA > 0 Then cboESTENTR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTENTREGA)
       txtCEPENTR.Text = objCADPEDVENDA.CEPENTREGA
       txtTELENTR.Text = objCADPEDVENDA.TELENTR
       txtFAXENTRE.Text = objCADPEDVENDA.FAXENTR
       
       '' Dados de Cobrança
       txtENDCOBR.Text = objCADPEDVENDA.ENDCOBRA
       txtBAICOBR.Text = objCADPEDVENDA.BAICOBRA
       txtCIDCOBR.Text = objCADPEDVENDA.CIDCOBRA
       If objCADPEDVENDA.ESTCOBRA > 0 Then cboESTCOBR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTCOBRA)
       txtCEPCOBR.Text = objCADPEDVENDA.CEPCOBRA
       txtTELCOBR.Text = objCADPEDVENDA.TELCOBRA
       txtFAXCOBR.Text = objCADPEDVENDA.FAXCOBRA
       
       txtOBS2.Text = objCADPEDVENDA.OBS2
       
       txtCODTRANSP.Text = objCADPEDVENDA.CODTRANSP
       txtORDCOMPCLI.Text = objCADPEDVENDA.ORDCOMPCLI
       txtCONTATO.Text = objCADPEDVENDA.CONTATO
       txtDEPARTAMENTO.Text = objCADPEDVENDA.DEPARTAMENTO
       
       optESPECIAL(objCADPEDVENDA.ESPECIAL).Value = True
       optPARAESTOQUE(objCADPEDVENDA.PARAESTOQUE).Value = True
       
       '' Totais
       If objCADPEDVENDA.VALBASICMS > 0 Then lblBASICMS.Caption = Format(objCADPEDVENDA.VALBASICMS, "#,##0.00")
       If objCADPEDVENDA.ALIQICMS > 0 Then txtALIQICMS.Text = Format(objCADPEDVENDA.ALIQICMS, "#,##0.00")
       If objCADPEDVENDA.VLICMS > 0 Then lblVLICMS.Caption = Format(objCADPEDVENDA.VLICMS, "#,##0.00")
       If objCADPEDVENDA.OUTRDESPESAS > 0 Then txtOutrDesp.Text = Format(objCADPEDVENDA.OUTRDESPESAS, "#,##0.00")
       If objCADPEDVENDA.VLFRETE > 0 Then txtFRETE.Text = Format(objCADPEDVENDA.VLFRETE, "#,##0.00")
       If objCADPEDVENDA.VLIPI > 0 Then lblVLIPI.Caption = Format(objCADPEDVENDA.VLIPI, "#,##0.00")
       If objCADPEDVENDA.VLDESCTO > 0 Then lblVLDESCONTO.Caption = Format(objCADPEDVENDA.VLDESCTO, "#,##0.00")
       If objCADPEDVENDA.TOTORCTO > 0 Then lblVLTOTAL.Caption = Format(objCADPEDVENDA.TOTORCTO, "#,##0.00")
       
        objCADPEDVENDA.PERMITEFECHOP = objCADPEDVENDA.PERMITEFECHOP
        
        chkVerificado.Value = objCADPEDVENDA.CONFERIDO
       
        If cTipOper = "LN" Or cTipOper = "LF" Then Call AbilConferido
       
        Call PopGrdProdutos
        Call PopLogPedidos
       
        stCAMPOSVENDA.Tab = 0
        If objCADPEDVENDA.STATUS = "C" Or objCADPEDVENDA.STATUS = "4" Then stCAMPOSVENDA.Tab = 0
       
        Call CarregaPlanoEntrega
        Call GeraSaldoPedido(Str(objCADPEDVENDA.CODPEDIDO))
        Call PegaDadosLabel
    
        Call MostraDadosReprovacao
    
        '' Se precisar voltar a trava
        If objCADPEDVENDA.STATUS = "V" Then
            With grdProgEntrega
                For i = 1 To (.Rows - 1)
                    .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega) = Empty
                Next i
            End With
        End If
    
        'If cTipOper = "LC" Then
        '    Call MudaStatusOP_PDPC
        'End If
    
        Call VisualizaBotoesPCD(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibAlteracao(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibFotolito(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibComercial(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibFinanceira(objCADPEDVENDA.STATUS, cTipOper)
        
        If objCADPEDVENDA.STATUS = "R" Then
            '' Limpa da Datas de Entrega
            With grdProgEntrega
                For i = 1 To (.Rows - 1)
                    .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega) = Empty
                Next i
            End With
        End If
    
    End If
        
    Exit Sub
    
Err_Libera:
        
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Libera()", Me.Name, "Libera()", strCAMARQERRO)
        
End Sub

Private Sub Reprova()

On Error GoTo Err_Reprova
    
    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Pedido de Venda - [ REPROVAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPEDVENDA

    stCAMPOSVENDA.Tab = 0
    stCAMPOSVENDA.TabVisible(2) = False
        
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame8.Enabled = True
    Frame9.Enabled = False
    Frame13.Visible = True
    
    txtOBS2.Locked = True
    
    objBLBFunc.Preenche_Estado cboESTENTR
    objBLBFunc.Preenche_Estado cboESTCOBR
    
    objCADPEDVENDA.CODPEDIDO = iCodigo
    
    Call InitGridReprovacao
    Call InitGridProd
    Call InitGridProg
    Call InitGridOrdemFat
    Call InitGridLogPed
    Call InitGridProducao
    
    
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    txtOutrDesp.Text = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    
    
    '' --------------------
    '' Desconto
    lblVLDESCONTO.Caption = ""
    '' --------------------
   
    Call LimpaCamposLabel
    Call LimpaCamposDadosAdicionais
    Call LimpaCampoSaldoRot
    Call LimpaSaldoPedido
    
    objCADPEDVENDA.FILIALPED = intFILIALPED
    
    optESPECIAL(0).Value = True
    
    Call AbilDesConferido(False, 0)
    
    If objCADPEDVENDA.Carrega_Campos = True Then
    
       If objCADPEDVENDA.STATUS = "B" Then lblSTATUS.Caption = "BLOQUEADO"
       If (objCADPEDVENDA.STATUS = "L" Or objCADPEDVENDA.STATUS = "N") Then
          lblSTATUS.Caption = "LIBERADO"
          cmdAltera.Enabled = False
       End If
       
       lblCODIGO.Caption = objCADPEDVENDA.CODPEDIDO
       mskDATAPED.Text = Format(objCADPEDVENDA.DATAPED, "DD/MM/YYYY")
       ''mskDTENTREGA.Text = Format(objCADPEDVENDA.DATAENTREGA, "DD/MM/YYYY")
       
       txtCIDCLIE.Text = objCADPEDVENDA.CODCLIE
       txtCodCondPgto.Text = objCADPEDVENDA.CODCONDPGTO
       txtCODVEND.Text = objCADPEDVENDA.CODVEND
       txtTIPPED.Text = objCADPEDVENDA.TIPPED
       txtOBSERVACAO.Text = objCADPEDVENDA.OBSERVACAO
       
       '' Dados de Entrega
       txtENDENTR.Text = objCADPEDVENDA.ENDENTR
       txtBAIENTR.Text = objCADPEDVENDA.BAIENTR
       txtCIDENTR.Text = objCADPEDVENDA.CIDENTR
       If objCADPEDVENDA.ESTENTREGA > 0 Then cboESTENTR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTENTREGA)
       txtCEPENTR.Text = objCADPEDVENDA.CEPENTREGA
       txtTELENTR.Text = objCADPEDVENDA.TELENTR
       txtFAXENTRE.Text = objCADPEDVENDA.FAXENTR
       
       '' Dados de Cobrança
       txtENDCOBR.Text = objCADPEDVENDA.ENDCOBRA
       txtBAICOBR.Text = objCADPEDVENDA.BAICOBRA
       txtCIDCOBR.Text = objCADPEDVENDA.CIDCOBRA
       If objCADPEDVENDA.ESTCOBRA > 0 Then cboESTCOBR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTCOBRA)
       txtCEPCOBR.Text = objCADPEDVENDA.CEPCOBRA
       txtTELCOBR.Text = objCADPEDVENDA.TELCOBRA
       txtFAXCOBR.Text = objCADPEDVENDA.FAXCOBRA
       
       txtOBSERVACAO.Text = objCADPEDVENDA.OBSERVACAO
       txtOBS2.Text = objCADPEDVENDA.OBS2
       
       '' Diversos
       ''If objCADPEDVENDA.PRZENTREGA > 0 Then txtPRZENTREGA.Text = objCADPEDVENDA.PRZENTREGA
       
       txtCODTRANSP.Text = objCADPEDVENDA.CODTRANSP
       txtORDCOMPCLI.Text = objCADPEDVENDA.ORDCOMPCLI
       txtCONTATO.Text = objCADPEDVENDA.CONTATO
       txtDEPARTAMENTO.Text = objCADPEDVENDA.DEPARTAMENTO
       
       optESPECIAL(objCADPEDVENDA.ESPECIAL).Value = True
       optPARAESTOQUE(objCADPEDVENDA.PARAESTOQUE).Value = True
       
       '' Totais
       If objCADPEDVENDA.VALBASICMS > 0 Then lblBASICMS.Caption = Format(objCADPEDVENDA.VALBASICMS, "#,##0.00")
       If objCADPEDVENDA.ALIQICMS > 0 Then txtALIQICMS.Text = Format(objCADPEDVENDA.ALIQICMS, "#,##0.00")
       If objCADPEDVENDA.VLICMS > 0 Then lblVLICMS.Caption = Format(objCADPEDVENDA.VLICMS, "#,##0.00")
       If objCADPEDVENDA.OUTRDESPESAS > 0 Then txtOutrDesp.Text = Format(objCADPEDVENDA.OUTRDESPESAS, "#,##0.00")
       If objCADPEDVENDA.VLFRETE > 0 Then txtFRETE.Text = Format(objCADPEDVENDA.VLFRETE, "#,##0.00")
       If objCADPEDVENDA.VLIPI > 0 Then lblVLIPI.Caption = Format(objCADPEDVENDA.VLIPI, "#,##0.00")
       If objCADPEDVENDA.VLDESCTO > 0 Then lblVLDESCONTO.Caption = Format(objCADPEDVENDA.VLDESCTO, "#,##0.00")
       If objCADPEDVENDA.TOTORCTO > 0 Then lblVLTOTAL.Caption = Format(objCADPEDVENDA.TOTORCTO, "#,##0.00")
       
        chkVerificado.Value = objCADPEDVENDA.CONFERIDO
        objCADPEDVENDA.PERMITEFECHOP = objCADPEDVENDA.PERMITEFECHOP
       
        Call PopGrdProdutos
        stCAMPOSVENDA.TabVisible(2) = True
        stCAMPOSVENDA.Tab = 2
       
        Call CarregaPlanoEntrega
        Call GeraSaldoPedido(Str(objCADPEDVENDA.CODPEDIDO))
        Call PegaDadosLabel
        Call PopLogPedidos
    
        Call VisualizaBotoesPCD(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibAlteracao(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibFotolito(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibComercial(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibFinanceira(objCADPEDVENDA.STATUS, cTipOper)
    
    End If
        
    Exit Sub
    
Err_Reprova:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Reprova()", Me.Name, "Reprova()", strCAMARQERRO)
    
        
End Sub

Private Sub Deslibera()

On Error GoTo Err_Deslibera
    
    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Pedido de Venda - [ DESLIBERA ]"
    
    objBLBFunc.LimpaCampos Me

    stCAMPOSVENDA.Tab = 0
    stCAMPOSVENDA.TabVisible(2) = False
    
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame8.Enabled = True
    Frame9.Enabled = False
    
    Frame13.Visible = True
    
    txtOBS2.Locked = True
    
    objBLBFunc.Preenche_Estado cboESTENTR
    objBLBFunc.Preenche_Estado cboESTCOBR
    
    objCADPEDVENDA.CODPEDIDO = iCodigo
    
    Call InitGridReprovacao
    Call InitGridProd
    Call InitGridProg
    Call InitGridOrdemFat
    Call InitGridLogPed
    Call InitGridProducao
    
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    txtOutrDesp.Text = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    
    Call LimpaCamposLabel
    Call LimpaCamposDadosAdicionais
    Call LimpaCampoSaldoRot
    Call LimpaSaldoPedido
    
    '' --------------------
    '' Desconto
    lblVLDESCONTO.Caption = ""
    '' --------------------
   
    objCADPEDVENDA.FILIALPED = intFILIALPED
   
    optESPECIAL(0).Value = True
    
    Call AbilDesConferido(False, 0)
    
    If objCADPEDVENDA.Carrega_Campos = True Then
    
       If objCADPEDVENDA.STATUS = "L" Then lblSTATUS.Caption = "LIBERADO"
       If objCADPEDVENDA.STATUS = "B" Or _
          objCADPEDVENDA.STATUS = "N" Or _
          objCADPEDVENDA.STATUS = "T" Or _
          objCADPEDVENDA.STATUS = "R" Then lblSTATUS.Caption = "BLOQUEADO"
       
       lblCODIGO.Caption = objCADPEDVENDA.CODPEDIDO
       mskDATAPED.Text = Format(objCADPEDVENDA.DATAPED, "DD/MM/YYYY")
       
       txtCIDCLIE.Text = objCADPEDVENDA.CODCLIE
       
       txtCodCondPgto.Text = objCADPEDVENDA.CODCONDPGTO
       txtCODVEND.Text = objCADPEDVENDA.CODVEND
       txtTIPPED.Text = objCADPEDVENDA.TIPPED
       
       '' Dados de Entrega
       txtENDENTR.Text = objCADPEDVENDA.ENDENTR
       txtBAIENTR.Text = objCADPEDVENDA.BAIENTR
       txtCIDENTR.Text = objCADPEDVENDA.CIDENTR
       If objCADPEDVENDA.ESTENTREGA > 0 Then cboESTENTR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTENTREGA)
       txtCEPENTR.Text = objCADPEDVENDA.CEPENTREGA
       txtTELENTR.Text = objCADPEDVENDA.TELENTR
       txtFAXENTRE.Text = objCADPEDVENDA.FAXENTR
       
       '' Dados de Cobrança
       txtENDCOBR.Text = objCADPEDVENDA.ENDCOBRA
       txtBAICOBR.Text = objCADPEDVENDA.BAICOBRA
       txtCIDCOBR.Text = objCADPEDVENDA.CIDCOBRA
       If objCADPEDVENDA.ESTCOBRA > 0 Then cboESTCOBR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTCOBRA)
       txtCEPCOBR.Text = objCADPEDVENDA.CEPCOBRA
       txtTELCOBR.Text = objCADPEDVENDA.TELCOBRA
       txtFAXCOBR.Text = objCADPEDVENDA.FAXCOBRA
       
       txtOBSERVACAO.Text = objCADPEDVENDA.OBSERVACAO
       txtOBS2.Text = objCADPEDVENDA.OBS2
       
       '' Diversos
       ''If objCADPEDVENDA.PRZENTREGA > 0 Then txtPRZENTREGA.Text = objCADPEDVENDA.PRZENTREGA
       
       txtCODTRANSP.Text = objCADPEDVENDA.CODTRANSP
       txtORDCOMPCLI.Text = objCADPEDVENDA.ORDCOMPCLI
       txtCONTATO.Text = objCADPEDVENDA.CONTATO
       txtDEPARTAMENTO.Text = objCADPEDVENDA.DEPARTAMENTO
       
       optESPECIAL(objCADPEDVENDA.ESPECIAL).Value = True
       optPARAESTOQUE(objCADPEDVENDA.PARAESTOQUE).Value = True
       
       '' Totais
       If objCADPEDVENDA.VALBASICMS > 0 Then lblBASICMS.Caption = Format(objCADPEDVENDA.VALBASICMS, "#,##0.00")
       If objCADPEDVENDA.ALIQICMS > 0 Then txtALIQICMS.Text = Format(objCADPEDVENDA.ALIQICMS, "#,##0.00")
       If objCADPEDVENDA.VLICMS > 0 Then lblVLICMS.Caption = Format(objCADPEDVENDA.VLICMS, "#,##0.00")
       If objCADPEDVENDA.OUTRDESPESAS > 0 Then txtOutrDesp.Text = Format(objCADPEDVENDA.OUTRDESPESAS, "#,##0.00")
       If objCADPEDVENDA.VLFRETE > 0 Then txtFRETE.Text = Format(objCADPEDVENDA.VLFRETE, "#,##0.00")
       If objCADPEDVENDA.VLIPI > 0 Then lblVLIPI.Caption = Format(objCADPEDVENDA.VLIPI, "#,##0.00")
       If objCADPEDVENDA.VLDESCTO > 0 Then lblVLDESCONTO.Caption = Format(objCADPEDVENDA.VLDESCTO, "#,##0.00")
       If objCADPEDVENDA.TOTORCTO > 0 Then lblVLTOTAL.Caption = Format(objCADPEDVENDA.TOTORCTO, "#,##0.00")
       
        chkVerificado.Value = objCADPEDVENDA.CONFERIDO
        objCADPEDVENDA.PERMITEFECHOP = objCADPEDVENDA.PERMITEFECHOP
        
        Call PopGrdProdutos
        Call CarregaPlanoEntrega
        Call GeraSaldoPedido(Str(objCADPEDVENDA.CODPEDIDO))
        Call PegaDadosLabel
        Call PopLogPedidos
    
        Call VisualizaBotoesPCD(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibAlteracao(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibFotolito(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibComercial(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibFinanceira(objCADPEDVENDA.STATUS, cTipOper)
    
    End If
        
    Exit Sub
    
Err_Deslibera:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Deslibera()", Me.Name, "Deslibera()", strCAMARQERRO)
        
        
End Sub


Private Sub Pega_Vendedor(lngCodUsuario As Long)

On Error GoTo Err_Pega_Vendedor
    
    Dim i As Integer
    
    If BREC2.State = 1 Then BREC2.Close
    
    txtCODVEND.Text = Str(lngCodUsuario)
    Call AtivaDesativacampos(True)
       
    '' ===========================================
    '' Pega os Tipos de Orcamentos para o Vendedor
    If Len(Trim(txtCODVEND.Text)) > 0 Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       ESPO.* " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_VENDTIPORCA VEND " & vbCrLf
        sSql = sSql & "      ,SGI_CADESPORCA  ESPO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       VEND.SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And VEND.SGI_CODVEND = " & Trim(txtCODVEND.Text) & vbCrLf
        sSql = sSql & "   And ESPO.SGI_FILIAL  = VEND.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ESPO.SGI_CODIGO  = VEND.SGI_CODTIPORCA "
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF() Then
           txtTIPPED.Text = BREC2!SGI_CODIGO
           Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVEND.Text, lblDescVendedor, "Pega_Vendedor()")
           Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADESPORCA", txtTIPPED.Text, lblDescTpPed, "Pega_Vendedor()")
           Call AtivaDesativacampos(False)
        Else
           txtCODVEND.Text = ""
           txtTIPPED.Text = ""
        End If
        BREC2.Close
        
    End If
    '' ===========================================
       
    Exit Sub
       
Err_Pega_Vendedor:

    If BREC2.State = 1 Then BREC2.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Pega_Vendedor()", Me.Name, "Pega_Vendedor()", strCAMARQERRO)
    
End Sub

Private Sub InitGridProd()

    With grdProduto
    
       .Cols = conColumnsIn_SonProd
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProd_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProd_IdProduto) = ""
       .ColDataType(conCOL_SonProd_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_Codigo) = ""
       .ColDataType(conCOL_SonProd_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProd_PesqProd) = ""
       .ColDataType(conCOL_SonProd_PesqProd) = flexDTString
       .ColComboList(conCOL_SonProd_PesqProd) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProd_DescProd) = ""
       .ColDataType(conCOL_SonProd_DescProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProd_QtdProd) = ""
       .ColDataType(conCOL_SonProd_QtdProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_VlUniProd) = ""
       .ColDataType(conCOL_SonProd_VlUniProd) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProd_PorcDesc) = ""
       .ColDataType(conCOL_SonProd_PorcDesc) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProd_PorcIPI) = ""
       .ColDataType(conCOL_SonProd_PorcIPI) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProd_VlTotal) = ""
       .ColDataType(conCOL_SonProd_VlTotal) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProd_VlDesc) = ""
       .ColDataType(conCOL_SonProd_VlDesc) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProd_VlIPI) = ""
       .ColDataType(conCOL_SonProd_VlIPI) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProd_VlItens) = ""
       .ColDataType(conCOL_SonProd_VlItens) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProd_Fechamento) = ""
       .ColDataType(conCOL_SonProd_Fechamento) = flexDTString
       .ColComboList(conCOL_SonProd_Fechamento) = objCADPEDVENDA.PreenchComboFechamentoGrdSA
       
       .Cell(flexcpData, 0, conCOL_SonProd_Corpo) = ""
       .ColDataType(conCOL_SonProd_Corpo) = flexDTString
       .ColComboList(conCOL_SonProd_Corpo) = objCADPEDVENDA.PreenchComboFechamentoGrd
       
       .Cell(flexcpData, 0, conCOL_SonProd_Tampa) = ""
       .ColDataType(conCOL_SonProd_Tampa) = flexDTString
       .ColComboList(conCOL_SonProd_Tampa) = objCADPEDVENDA.PreenchComboFechamentoGrd
       
       .Cell(flexcpData, 0, conCOL_SonProd_FornTampa) = ""
       .ColDataType(conCOL_SonProd_FornTampa) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_PesqForn) = ""
       .ColDataType(conCOL_SonProd_PesqForn) = flexDTString
       .ColComboList(conCOL_SonProd_PesqForn) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProd_Fundo) = ""
       .ColDataType(conCOL_SonProd_Fundo) = flexDTString
       .ColComboList(conCOL_SonProd_Fundo) = objCADPEDVENDA.PreenchComboFechamentoGrd
       
       .Cell(flexcpData, 0, conCOL_SonProd_Argola) = ""
       .ColDataType(conCOL_SonProd_Argola) = flexDTString
       .ColComboList(conCOL_SonProd_Argola) = objCADPEDVENDA.PreenchComboFechamentoGrd
       
       .Cell(flexcpData, 0, conCOL_SonProd_FechTpFr) = ""
       .ColDataType(conCOL_SonProd_FechTpFr) = flexDTString
       ''.ColComboList(conCOL_SonProd_FechTpFr) = objCADPEDVENDA.PreenchComboFechamentoTampaFuro

       .Cell(flexcpData, 0, conCOL_SonProd_Desenho) = ""
       .ColDataType(conCOL_SonProd_Desenho) = flexDTString
       .ColComboList(conCOL_SonProd_Desenho) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProd_AltFilme) = ""
       .ColDataType(conCOL_SonProd_AltFilme) = flexDTString
       .ColComboList(conCOL_SonProd_AltFilme) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_SonProd_FotNovo) = ""
       .ColDataType(conCOL_SonProd_FotNovo) = flexDTString
       .ColComboList(conCOL_SonProd_FotNovo) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_SonProd_Repeticao) = ""
       .ColDataType(conCOL_SonProd_Repeticao) = flexDTString
       .ColComboList(conCOL_SonProd_Repeticao) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_SonProd_CodLinProd) = ""
       .ColDataType(conCOL_SonProd_CodLinProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_OBSOP) = ""
       .ColDataType(conCOL_SonProd_OBSOP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProd_IDBKP) = ""
       .ColDataType(conCOL_SonProd_IDBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_QTDBKP) = ""
       .ColDataType(conCOL_SonProd_QTDBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_PRECOBKP) = ""
       .ColDataType(conCOL_SonProd_PRECOBKP) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProd_FechTpFrBKP) = ""
       .ColDataType(conCOL_SonProd_FechTpFrBKP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProd_AltFilmeBKP) = ""
       .ColDataType(conCOL_SonProd_AltFilmeBKP) = flexDTString

       .Cell(flexcpData, 0, conCOL_SonProd_FotNovoBKP) = ""
       .ColDataType(conCOL_SonProd_FotNovoBKP) = flexDTString

       .Cell(flexcpData, 0, conCOL_SonProd_RepeticaoBKP) = ""
       .ColDataType(conCOL_SonProd_RepeticaoBKP) = flexDTString

       .Cell(flexcpData, 0, conCOL_SonProd_Action2Do) = ""
       .ColDataType(conCOL_SonProd_Action2Do) = flexDTLong

       .Cell(flexcpData, 0, conCOL_SonProd_TemOP) = ""
       .ColDataType(conCOL_SonProd_TemOP) = flexDTString

       .Cell(flexcpData, 0, conCOL_SonProd_StatusProd) = ""
       .ColDataType(conCOL_SonProd_StatusProd) = flexDTLong

       .Cell(flexcpData, 0, conCOL_SonProd_GrpPlanMestre) = ""
       .ColDataType(conCOL_SonProd_GrpPlanMestre) = flexDTLong

       .Cell(flexcpData, 0, conCOL_SonProd_CodCapacidade) = ""
       .ColDataType(conCOL_SonProd_CodCapacidade) = flexDTLong

       .Cell(flexcpData, 0, conCOL_SonProd_NECKIN) = ""
       .ColDataType(conCOL_SonProd_NECKIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_HOMOLOGADO) = ""
       .ColDataType(conCOL_SonProd_HOMOLOGADO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_QTDELATASPALLETS) = ""
       .ColDataType(conCOL_SonProd_QTDELATASPALLETS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_PALLETS) = ""
       .ColDataType(conCOL_SonProd_QTDELATASPALLETS) = flexDTLong
       
       ''.Cell(flexcpData, 0, conCOL_SonProd_Conferido) = ""
       ''.ColDataType(conCOL_SonProd_Conferido) = flexDTString
       ''.ColComboList(conCOL_SonProd_Conferido) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_SonProd_Conferido) = ""
       .ColDataType(conCOL_SonProd_Conferido) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_SonProd_PalhetPadrao) = ""
       .ColDataType(conCOL_SonProd_PalhetPadrao) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProd_OS_Artes) = ""
       .ColDataType(conCOL_SonProd_OS_Artes) = flexDTString
       .ColComboList(conCOL_SonProd_OS_Artes) = "..."
       
       .ColWidth(conCOL_SonProd_IdProduto) = 0
       .ColWidth(conCOL_SonProd_Codigo) = 1200
       .ColWidth(conCOL_SonProd_PesqProd) = 300
       .ColWidth(conCOL_SonProd_DescProd) = 5000
       
       .ColWidth(conCOL_SonProd_QtdProd) = 800
       .ColWidth(conCOL_SonProd_VlUniProd) = 800
       .ColWidth(conCOL_SonProd_PorcDesc) = 700
       .ColWidth(conCOL_SonProd_PorcIPI) = 600
       .ColWidth(conCOL_SonProd_VlTotal) = 1000
       
       .ColWidth(conCOL_SonProd_VlDesc) = 0
       .ColWidth(conCOL_SonProd_VlIPI) = 0
       .ColWidth(conCOL_SonProd_VlItens) = 0
       
       .ColWidth(conCOL_SonProd_Fechamento) = 0
       .ColWidth(conCOL_SonProd_Corpo) = 0
       .ColWidth(conCOL_SonProd_Tampa) = 0
       .ColWidth(conCOL_SonProd_FornTampa) = 0
       .ColWidth(conCOL_SonProd_PesqForn) = 0
       
       .ColWidth(conCOL_SonProd_Fundo) = 0
       .ColWidth(conCOL_SonProd_Argola) = 0
       .ColWidth(conCOL_SonProd_FechTpFr) = 2000
       .ColWidth(conCOL_SonProd_Desenho) = 300
       
       .ColWidth(conCOL_SonProd_AltFilme) = 800
       .ColWidth(conCOL_SonProd_FotNovo) = 800
       .ColWidth(conCOL_SonProd_Repeticao) = 900
       .ColWidth(conCOL_SonProd_CodLinProd) = 0
       .ColWidth(conCOL_SonProd_OBSOP) = 0
       
       .ColWidth(conCOL_SonProd_IDBKP) = 0
       .ColWidth(conCOL_SonProd_PRECOBKP) = 0
       .ColWidth(conCOL_SonProd_QTDBKP) = 0
       .ColWidth(conCOL_SonProd_FechTpFrBKP) = 0
       .ColWidth(conCOL_SonProd_AltFilmeBKP) = 0
       .ColWidth(conCOL_SonProd_FotNovoBKP) = 0
       .ColWidth(conCOL_SonProd_RepeticaoBKP) = 0
       .ColWidth(conCOL_SonProd_Action2Do) = 0
       .ColWidth(conCOL_SonProd_TemOP) = 0
       .ColWidth(conCOL_SonProd_StatusProd) = 0
       .ColWidth(conCOL_SonProd_GrpPlanMestre) = 0
       .ColWidth(conCOL_SonProd_CodCapacidade) = 0
       .ColWidth(conCOL_SonProd_NECKIN) = 0
       .ColWidth(conCOL_SonProd_HOMOLOGADO) = 0
       
       '' Depois Voltar
       .ColWidth(conCOL_SonProd_QTDELATASPALLETS) = 0
       .ColWidth(conCOL_SonProd_PALLETS) = 0
       .ColWidth(conCOL_SonProd_Conferido) = 900
       .ColWidth(conCOL_SonProd_PalhetPadrao) = 0
       
       .ColWidth(conCOL_SonProd_OS_Artes) = 300
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGridProdtos()
   
On Error GoTo Err_IncRegGridProdtos
    
    Dim strCampos01 As String
    
    
    If objBLBFunc.FcExisteLinhaVazia(grdProduto, conCOL_SonProd_IdProduto) = False Then Exit Sub
    
    strCampos01 = "" & vbTab & _
                  ""
    
    
                       
    grdProduto.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       dacEnumUpdateAction_Insert & vbTab & "N" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                       "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & 0 & vbTab & 0 & vbTab & ""
    
    
    grdProduto.Cell(flexcpText, (grdProduto.Rows - 1), conCOL_SonProd_AltFilme) = 0
    grdProduto.Cell(flexcpText, (grdProduto.Rows - 1), conCOL_SonProd_FotNovo) = 0
    grdProduto.Cell(flexcpText, (grdProduto.Rows - 1), conCOL_SonProd_Repeticao) = 0
    
    
    
    Exit Sub

Err_IncRegGridProdtos:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : IncRegGridProdtos()", Me.Name, "IncRegGridProdtos()", strCAMARQERRO)
                            
End Sub

Private Function PegaIDProduto(strCODPRODUTO As String) As String

On Error GoTo Err_PegaIDProduto
    
    PegaIDProduto = ""
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And (PRO.SGI_STATUS = 1 Or PRO.SGI_STATUS = 2)" & vbCrLf
    sSql = sSql & "   And PRO.SGI_CODIGO  = '" & strCODPRODUTO & "'" & vbCrLf

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If BREC!SGI_STATUS = 0 Then
           MsgBox "ATENÇÃO !!!" & vbCrLf & "O Produto " & Trim(strCODPRODUTO) & " - " & Trim(BREC!SGI_DESCRICAO) & vbCrLf & "Não pode ser Utilizado está Desativado !!!", vbOKOnly + vbExclamation, "Aviso"
        Else
           PegaIDProduto = BREC!SGI_IDPRODUTO
        End If
    End If
    BREC.Close
    
    Exit Function
    
Err_PegaIDProduto:

    If BREC.State = 1 Then BREC.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaIDProduto()", Me.Name, "PegaIDProduto()", strCAMARQERRO)
    
End Function

Private Function PegaDescrProduto(strCODPRODUTO As String) As String
    
On Error GoTo Err_PegaDescrProduto
    
    PegaDescrProduto = ""
    
    If BREC2.State = 1 Then BREC2.Close
    
    Dim i As Integer
    
    For i = 1 To 2
    
        If i = 1 Then
        
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRO.*" & vbCrLf
            sSql = sSql & "      ,LINHA.SGI_FILIALPED As FILIALPED_2" & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADCLIEVEND CLIV" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO  PRO" & vbCrLf
            sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINHA" & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       CLIV.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And CLIV.SGI_CODIGO     = " & Trim(txtCODVEND.Text) & vbCrLf
            sSql = sSql & "   And CLIV.SGI_CODCLI     = " & Trim(txtCIDCLIE.Text) & vbCrLf
            sSql = sSql & "   And CLIV.SGI_FILIAL     = PRO.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And CLIV.SGI_CODCLI     = PRO.SGI_CODCLIE" & vbCrLf
            
            sSql = sSql & "   And PRO.SGI_IDPRODUTO   = " & strCODPRODUTO & vbCrLf
            sSql = sSql & "   And (PRO.SGI_STATUS     = 1 Or PRO.SGI_STATUS = 2)" & vbCrLf
            sSql = sSql & "   And PRO.SGI_FILIALPED   = " & intFILIALPED & vbCrLf
            
            sSql = sSql & "   And LINHA.SGI_FILIAL    = PRO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And LINHA.SGI_CODLIN    = PRO.SGI_CODLINPROD " & vbCrLf
        
        ElseIf i = 2 Then
        
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRO.* " & vbCrLf
            sSql = sSql & "      ,LINHA.SGI_FILIALPED As FILIALPED_2" & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO      PRO" & vbCrLf
            sSql = sSql & "      ,SGI_PRODATECLIE     PCL" & vbCrLf
            sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINHA" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "         PCL.SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "     And PCL.SGI_IDCLIENTE   = " & Trim(txtCIDCLIE.Text) & vbCrLf
            sSql = sSql & "     And PCL.SGI_IDPRODUTO   = " & strCODPRODUTO & vbCrLf
            sSql = sSql & "     And PCL.SGI_FILIAL      = PRO.SGI_FILIAL" & vbCrLf
            sSql = sSql & "     And PCL.SGI_IDPRODUTO   = PRO.SGI_IDPRODUTO" & vbCrLf
            sSql = sSql & "     And (PRO.SGI_STATUS     = 1 Or PRO.SGI_STATUS      = 2)" & vbCrLf
            sSql = sSql & "     And LINHA.SGI_FILIAL    = PRO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "     And LINHA.SGI_CODLIN    = PRO.SGI_CODLINPROD " & vbCrLf
            sSql = sSql & "     And PRO.SGI_FILIALPED   = " & intFILIALPED & vbCrLf
        
        End If
        
        ''sSql = sSql & "       PRO.SGI_FILIAL      = " & FILIAL & vbCrLf
        ''sSql = sSql & "   And PRO.SGI_IDPRODUTO   = " & strCodProduto & vbCrLf
        ''sSql = sSql & "   And (PRO.SGI_STATUS     = 1 or PRO.SGI_STATUS      = 2)" & vbCrLf
        ''sSql = sSql & "   And LINHA.SGI_FILIAL    = PRO.SGI_FILIAL" & vbCrLf
        ''sSql = sSql & "   And LINHA.SGI_CODLIN    = PRO.SGI_CODLINPROD" & vbCrLf
        ''sSql = sSql & "   And PRO.SGI_FILIALPED   = " & intFILIALPED
        
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF Then
            If BREC2!SGI_CODTIPO = 2 Then
            
                PegaDescrProduto = BREC2!SGI_DESCRICAO
            
                With grdProduto
                     
                     .Cell(flexcpText, .Row, conCOL_SonProd_Fechamento) = 0
                     .Cell(flexcpText, .Row, conCOL_SonProd_Corpo) = -1
                     .Cell(flexcpText, .Row, conCOL_SonProd_Tampa) = -1
                     .Cell(flexcpText, .Row, conCOL_SonProd_Fundo) = -1
                     .Cell(flexcpText, .Row, conCOL_SonProd_Argola) = -1
                     .Cell(flexcpText, .Row, conCOL_SonProd_FechTpFr) = Empty
                     .Cell(flexcpText, .Row, conCOL_SonProd_CodLinProd) = Empty
                     .Cell(flexcpText, .Row, conCOL_SonProd_StatusProd) = Empty
                     .Cell(flexcpText, .Row, conCOL_SonProd_GrpPlanMestre) = Empty
                     .Cell(flexcpText, .Row, conCOL_SonProd_FotNovo) = 0
                     .Cell(flexcpText, .Row, conCOL_SonProd_AltFilme) = 0
                     .Cell(flexcpText, .Row, conCOL_SonProd_Repeticao) = 0
                     
                     If Not IsNull(BREC2!SGI_FechSoldaAgrafado) Then .Cell(flexcpText, .Row, conCOL_SonProd_Fechamento) = BREC2!SGI_FechSoldaAgrafado
                     If Not IsNull(BREC2!SGI_VernCorpo) Then .Cell(flexcpText, .Row, conCOL_SonProd_Corpo) = BREC2!SGI_VernCorpo
                     If Not IsNull(BREC2!SGI_VernTampa) Then .Cell(flexcpText, .Row, conCOL_SonProd_Tampa) = BREC2!SGI_VernTampa
                     If Not IsNull(BREC2!SGI_VernFundo) Then .Cell(flexcpText, .Row, conCOL_SonProd_Fundo) = BREC2!SGI_VernFundo
                     If Not IsNull(BREC2!SGI_VernArgola) Then .Cell(flexcpText, .Row, conCOL_SonProd_Argola) = BREC2!SGI_VernArgola
                     If Not IsNull(BREC2!SGI_STATUS) Then .Cell(flexcpText, .Row, conCOL_SonProd_StatusProd) = BREC2!SGI_STATUS
                     
                     If Not IsNull(BREC2!SGI_PRODNOVO) Then .Cell(flexcpText, .Row, conCOL_SonProd_FotNovo) = BREC2!SGI_PRODNOVO
                     If Not IsNull(BREC2!SGI_ALTFOTOLIT) Then .Cell(flexcpText, .Row, conCOL_SonProd_AltFilme) = BREC2!SGI_ALTFOTOLIT
                     
                     ''If Not IsNull(BREC2!SGI_FechTampaFuro) Then .Cell(flexcpText, .Row, conCOL_SonProd_FechTpFr) = BREC2!SGI_FechTampaFuro
                     
                     .Cell(flexcpText, .Row, conCOL_SonProd_CodLinProd) = BREC2!SGI_CODLINPROD
                     
                     .Cell(flexcpText, .Row, conCOL_SonProd_PorcIPI) = Empty
                     If Not IsNull(BREC2!SGI_IPI) Then .Cell(flexcpText, .Row, conCOL_SonProd_PorcIPI) = BREC2!SGI_IPI
                
                
                End With
            
            Else
                ''If BREC2!FILIALPED_2 = 0 Then
                ''    MsgBox "Somente é permitido incluir produtos homologados !!!", vbOKOnly + vbExclamation, "Aviso"
                ''ElseIf BREC2!FILIALPED_2 = 1 Then
                
                    PegaDescrProduto = BREC2!SGI_DESCRICAO
                
                    With grdProduto
                         
                         .Cell(flexcpText, .Row, conCOL_SonProd_Fechamento) = 0
                         .Cell(flexcpText, .Row, conCOL_SonProd_Corpo) = -1
                         .Cell(flexcpText, .Row, conCOL_SonProd_Tampa) = -1
                         .Cell(flexcpText, .Row, conCOL_SonProd_Fundo) = -1
                         .Cell(flexcpText, .Row, conCOL_SonProd_Argola) = -1
                         .Cell(flexcpText, .Row, conCOL_SonProd_FechTpFr) = Empty
                         .Cell(flexcpText, .Row, conCOL_SonProd_CodLinProd) = Empty
                         .Cell(flexcpText, .Row, conCOL_SonProd_StatusProd) = Empty
                         .Cell(flexcpText, .Row, conCOL_SonProd_FotNovo) = 0
                         .Cell(flexcpText, .Row, conCOL_SonProd_AltFilme) = 0
                         .Cell(flexcpText, .Row, conCOL_SonProd_Repeticao) = 0
                         
                         If Not IsNull(BREC2!SGI_FechSoldaAgrafado) Then .Cell(flexcpText, .Row, conCOL_SonProd_Fechamento) = BREC2!SGI_FechSoldaAgrafado
                         If Not IsNull(BREC2!SGI_VernCorpo) Then .Cell(flexcpText, .Row, conCOL_SonProd_Corpo) = BREC2!SGI_VernCorpo
                         If Not IsNull(BREC2!SGI_VernTampa) Then .Cell(flexcpText, .Row, conCOL_SonProd_Tampa) = BREC2!SGI_VernTampa
                         If Not IsNull(BREC2!SGI_VernFundo) Then .Cell(flexcpText, .Row, conCOL_SonProd_Fundo) = BREC2!SGI_VernFundo
                         If Not IsNull(BREC2!SGI_VernArgola) Then .Cell(flexcpText, .Row, conCOL_SonProd_Argola) = BREC2!SGI_VernArgola
                         If Not IsNull(BREC2!SGI_STATUS) Then .Cell(flexcpText, .Row, conCOL_SonProd_StatusProd) = BREC2!SGI_STATUS
                         
                         If Not IsNull(BREC2!SGI_PRODNOVO) Then .Cell(flexcpText, .Row, conCOL_SonProd_FotNovo) = BREC2!SGI_PRODNOVO
                         If Not IsNull(BREC2!SGI_ALTFOTOLIT) Then .Cell(flexcpText, .Row, conCOL_SonProd_AltFilme) = BREC2!SGI_ALTFOTOLIT
                         
                         
                         ''If Not IsNull(BREC2!SGI_FechTampaFuro) Then .Cell(flexcpText, .Row, conCOL_SonProd_FechTpFr) = BREC2!SGI_FechTampaFuro
                         
                         .Cell(flexcpText, .Row, conCOL_SonProd_CodLinProd) = BREC2!SGI_CODLINPROD
                    
                        .Cell(flexcpText, .Row, conCOL_SonProd_PorcIPI) = Empty
                        If Not IsNull(BREC2!SGI_IPI) Then .Cell(flexcpText, .Row, conCOL_SonProd_PorcIPI) = BREC2!SGI_IPI
                    
                    End With
                
                ''End If
            End If
        End If
        BREC2.Close
    Next i
    
    Exit Function
    
Err_PegaDescrProduto:

    If BREC2.State = 1 Then BREC2.Close

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDescrProduto()", Me.Name, "PegaDescrProduto()", strCAMARQERRO)
    
    
End Function

Private Sub CalcTotPedido()

On Error GoTo Err_CalcTotPedido
    
    Dim i                   As Integer
    Dim curBaseCalculo      As Currency
    Dim curAliqICMS         As Currency
    Dim curValICMS          As Currency
    Dim curValorIPI         As Currency
    Dim curTotalDescIten    As Currency
    Dim curValOutrDesp      As Currency
    Dim curValFrete         As Currency
    Dim curPercDescPedido   As Currency
    Dim curDescPedido       As Currency
    Dim curTotalPedido      As Currency
    Dim curVLITens          As Currency
    Dim curVLDesc           As Currency
    Dim curVLIPI            As Currency
    Dim curVTOT             As Currency
    
    
    lblTotalItens.Caption = ""
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    lblVLDESCONTO.Caption = ""
    lblVLTOTAL.Caption = ""
    txtVLDESCTOTOT.Text = ""
    
    curBaseCalculo = 0
    curAliqICMS = 0
    curValICMS = 0
    curValorIPI = 0
    curVLITens = 0
    curVLDesc = 0
    curTotalDescIten = 0
    curTotalPedido = 0
    curValOutrDesp = 0
    curValFrete = 0
    curPercDescPedido = 0
    curVLIPI = 0
    curVTOT = 0
    
    
    For i = 1 To (grdProduto.Rows - 1)
        If Len(Trim(grdProduto.Cell(flexcpText, i, conCOL_SonProd_Codigo))) > 0 Then
            
            curVLITens = 0
            If Len(Trim(grdProduto.Cell(flexcpText, i, conCOL_SonProd_VlItens))) > 0 Then curVLITens = CCur(grdProduto.Cell(flexcpText, i, conCOL_SonProd_VlItens))
            curBaseCalculo = curBaseCalculo + curVLITens
            
            curVLDesc = 0
            If Len(Trim(grdProduto.Cell(flexcpText, i, conCOL_SonProd_VlDesc))) > 0 Then curVLDesc = CCur(grdProduto.Cell(flexcpText, i, conCOL_SonProd_VlDesc))
            curTotalDescIten = curTotalDescIten + curVLDesc
            
            curVLIPI = 0
            If Len(Trim(grdProduto.Cell(flexcpText, i, conCOL_SonProd_VlIPI))) > 0 Then curVLIPI = CCur(grdProduto.Cell(flexcpText, i, conCOL_SonProd_VlIPI))
            curValorIPI = curValorIPI + curVLIPI
            
            curVTOT = 0
            If Len(Trim(grdProduto.Cell(flexcpText, i, conCOL_SonProd_VlTotal))) > 0 Then curVTOT = CCur(grdProduto.Cell(flexcpText, i, conCOL_SonProd_VlTotal))
            curTotalPedido = curTotalPedido + curVTOT
        
        End If
    Next i
    lblTotalItens.Caption = Format(curTotalPedido, "#,##0.00")
    
    lblBASICMS.Caption = Format(curBaseCalculo, "#,##0.00")
    
    If Len(Trim(txtALIQICMS.Text)) > 0 Then
        curAliqICMS = CCur(txtALIQICMS.Text)
        curValICMS = ((curBaseCalculo * curAliqICMS) / 100)
        lblVLICMS.Caption = Format(curValICMS, "#,##0.00")
    End If
    If Len(Trim(txtPDESCTOTAL.Text)) > 0 Then
        curPercDescPedido = CCur(txtPDESCTOTAL.Text)
        curDescPedido = ((curBaseCalculo * curPercDescPedido) / 100)
        txtVLDESCTOTOT.Text = Format(curDescPedido, "#,##0.00")
    End If
    
    lblVLDESCONTO.Caption = Format(curTotalDescIten, "#,##0.00")
    lblVLIPI.Caption = Format(curValorIPI, "#,##0.00")
    
    If Len(Trim(txtOutrDesp.Text)) > 0 Then curValOutrDesp = CCur(txtOutrDesp.Text)
    If Len(Trim(txtFRETE.Text)) > 0 Then curValFrete = CCur(txtFRETE.Text)
    curTotalPedido = (curTotalPedido + (curValOutrDesp + curValFrete) - curDescPedido)
    
    lblVLTOTAL.Caption = Format(curTotalPedido, "#,##0.00")
    
    Exit Sub
    
Err_CalcTotPedido:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CalcTotPedido()", Me.Name, "CalcTotPedido()", strCAMARQERRO)
    
End Sub

Private Function CalcItenGrid(lngROW As Long) As Currency
    
On Error GoTo Err_CalcItenGrid

    CalcItenGrid = 0
    
    Dim curQtd_do_Item        As Currency
    Dim curVlUn_do_Item       As Currency
    Dim curDESC_do_Item       As Currency
    Dim curVLDESC_do_Item     As Currency
    Dim curIPI_do_Item        As Currency
    Dim curVlIPI_do_Item      As Currency
    Dim curTotal_do_Iten      As Currency
    Dim curTot_Latas_4_Palhet As Currency
    Dim curTot_Latas_Palhet   As Currency
    Dim lngRestoPAlhets       As Long
    
    curQtd_do_Item = 0
    curVlUn_do_Item = 0
    curDESC_do_Item = 0
    curVLDESC_do_Item = 0
    curIPI_do_Item = 0
    curVlIPI_do_Item = 0
    curTotal_do_Iten = 0
    curTot_Latas_4_Palhet = 0
    curTot_Latas_Palhet = 0
    
    With grdProduto
         If Len(Trim(.Cell(flexcpText, lngROW, conCOL_SonProd_QtdProd))) > 0 Then curQtd_do_Item = CCur(.Cell(flexcpText, lngROW, conCOL_SonProd_QtdProd))
         If Len(Trim(.Cell(flexcpText, lngROW, conCOL_SonProd_VlUniProd))) > 0 Then curVlUn_do_Item = CCur(.Cell(flexcpText, lngROW, conCOL_SonProd_VlUniProd))
         If Len(Trim(.Cell(flexcpText, lngROW, conCOL_SonProd_PorcDesc))) > 0 Then curDESC_do_Item = CCur(.Cell(flexcpText, lngROW, conCOL_SonProd_PorcDesc))
         If Len(Trim(.Cell(flexcpText, lngROW, conCOL_SonProd_PorcIPI))) > 0 Then curIPI_do_Item = CCur(.Cell(flexcpText, lngROW, conCOL_SonProd_PorcIPI))
         If Len(Trim(.Cell(flexcpText, lngROW, conCOL_SonProd_QTDELATASPALLETS))) > 0 Then curTot_Latas_4_Palhet = CCur(.Cell(flexcpText, lngROW, conCOL_SonProd_QTDELATASPALLETS))
         
         curTotal_do_Iten = (curQtd_do_Item * curVlUn_do_Item)
         .Cell(flexcpText, lngROW, conCOL_SonProd_VlItens) = Format(curTotal_do_Iten, "#,##0.00")
         
         '' Desconto
         curVLDESC_do_Item = ((curTotal_do_Iten * curDESC_do_Item) / 100)
         curTotal_do_Iten = (curTotal_do_Iten - curVLDESC_do_Item)
         
         '' IPI
         curVlIPI_do_Item = ((curTotal_do_Iten * curIPI_do_Item) / 100)
         curTotal_do_Iten = (curTotal_do_Iten + curVlIPI_do_Item)
         
         .Cell(flexcpText, lngROW, conCOL_SonProd_VlDesc) = Format(curVLDESC_do_Item, "#,##0.00")
         .Cell(flexcpText, lngROW, conCOL_SonProd_VlIPI) = Format(curVlIPI_do_Item, "#,##0.00")
    
''         If curTot_Latas_4_Palhet > 0 Then
''            curTot_Latas_Palhet = Round((curQtd_do_Item / curTot_Latas_4_Palhet), 2)
''            lngRestoPAlhets = (curQtd_do_Item Mod curTot_Latas_4_Palhet)
''         End If
            
''         If lngRestoPAlhets = 0 Then
''            If curTot_Latas_Palhet > 0 Then .Cell(flexcpText, lngROW, conCOL_SonProd_PALLETS) = curTot_Latas_Palhet
''         Else
''            MsgBox "ATENÇÃO" & vbCrLf & _
''                   "Esta Quantidade de Latas, irá gerar : " & Fix(curTot_Latas_Palhet) & " Palhet(s) Inteiros , " & vbCrLf & _
''                   "e ira gerar um resto de " & lngRestoPAlhets & " lata(s), quantidade sugerida : " & (curQtd_do_Item - lngRestoPAlhets) & ".", vbOKOnly + vbExclamation, "Aviso"
''         End If
    
    End With
    
    CalcItenGrid = curTotal_do_Iten

    Exit Function
    
Err_CalcItenGrid:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CalcItenGrid()", Me.Name, "CalcItenGrid()", strCAMARQERRO)

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
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DataPrevLito) = ""
       .ColDataType(conCOL_SonProgEntr_DataPrevLito) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DataPrevProd) = ""
       .ColDataType(conCOL_SonProgEntr_DataPrevProd) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DataEntrega) = ""
       .ColDataType(conCOL_SonProgEntr_DataEntrega) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_QtdProd) = ""
       .ColDataType(conCOL_SonProgEntr_QtdProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_Action2Do) = ""
       .ColDataType(conCOL_SonProgEntr_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_OBSOP) = ""
       .ColDataType(conCOL_SonProgEntr_OBSOP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_CodOP) = ""
       .ColDataType(conCOL_SonProgEntr_CodOP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_StatusOP) = ""
       .ColDataType(conCOL_SonProgEntr_StatusOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_FechTpFr) = ""
       .ColDataType(conCOL_SonProgEntr_FechTpFr) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_INDICE) = ""
       .ColDataType(conCOL_SonProgEntr_INDICE) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DataEntregaBKP) = ""
       .ColDataType(conCOL_SonProgEntr_DataEntregaBKP) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_IDINTERNO) = ""
       .ColDataType(conCOL_SonProgEntr_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DescStatusOP) = ""
       .ColDataType(conCOL_SonProgEntr_DescStatusOP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_GrpPlanMestre) = ""
       .ColDataType(conCOL_SonProgEntr_GrpPlanMestre) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_PegaPlanMestre) = ""
       .ColDataType(conCOL_SonProgEntr_PegaPlanMestre) = flexDTString
       .ColComboList(conCOL_SonProgEntr_PegaPlanMestre) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_QTDENOPALHET) = ""
       .ColDataType(conCOL_SonProgEntr_QTDENOPALHET) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_PALHET) = ""
       .ColDataType(conCOL_SonProgEntr_PALHET) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_Action2DoDtEntrega) = ""
       .ColDataType(conCOL_SonProgEntr_Action2DoDtEntrega) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DataPrevProd) = ""
       .ColDataType(conCOL_SonProgEntr_DataPrevProd) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DataPrevLito) = ""
       .ColDataType(conCOL_SonProgEntr_DataPrevLito) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_CODIDPROG) = ""
       .ColDataType(conCOL_SonProgEntr_CODIDPROG) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_CODSTATAPONT) = ""
       .ColDataType(conCOL_SonProgEntr_CODSTATAPONT) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProgEntr_DESCSTATUSAPONT) = ""
       .ColDataType(conCOL_SonProgEntr_DESCSTATUSAPONT) = flexDTString
       
       .ColWidth(conCOL_SonProgEntr_IdProduto) = 0
       .ColWidth(conCOL_SonProgEntr_QtdProd) = 800
       .ColWidth(conCOL_SonProgEntr_DataEntrega) = 1000
       
       .ColWidth(conCOL_SonProgEntr_Action2Do) = 0
       .ColWidth(conCOL_SonProgEntr_Action2DoDtEntrega) = 0
       
       .ColWidth(conCOL_SonProgEntr_OBSOP) = 0
       .ColWidth(conCOL_SonProgEntr_StatusOP) = 0
       .ColWidth(conCOL_SonProgEntr_FechTpFr) = 0
       .ColWidth(conCOL_SonProgEntr_INDICE) = 0
       .ColWidth(conCOL_SonProgEntr_INDICEBKP) = 0
       .ColWidth(conCOL_SonProgEntr_DataEntregaBKP) = 0
       .ColWidth(conCOL_SonProgEntr_DescStatusOP) = 900
       .ColWidth(conCOL_SonProgEntr_GrpPlanMestre) = 0
       
       '' Esta Coluna Será Suprimida qq Coisa Voltar ela
       '' Valor antido do ColWidth = 300
       .ColWidth(conCOL_SonProgEntr_PegaPlanMestre) = 0
       
       '' Depois Voltar
       .ColWidth(conCOL_SonProgEntr_QTDENOPALHET) = 0
       .ColWidth(conCOL_SonProgEntr_PALHET) = 0
       
       If cTipOper = "I" Then
            .ColWidth(conCOL_SonProgEntr_CodOP) = 0
       Else
            .ColWidth(conCOL_SonProgEntr_CodOP) = 1000
       End If
       
       .ColWidth(conCOL_SonProgEntr_IDINTERNO) = 0
       .ColWidth(conCOL_SonProgEntr_DataPrevLito) = 0
       .ColWidth(conCOL_SonProgEntr_CODIDPROG) = 0
       .ColWidth(conCOL_SonProgEntr_CODSTATAPONT) = 0
       
       .ColWidth(conCOL_SonProgEntr_DataPrevProd) = 0
       .ColWidth(conCOL_SonProgEntr_DESCSTATUSAPONT) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGridProg()
    
    
On Error GoTo Err_IncRegGridProg
    
    Dim curQtdProdInf   As Currency
    Dim curPegaSaldo    As Currency
    Dim i               As Integer
    Dim lngSTATUSPROD   As Long
    
    With grdProduto
        If .RowSel <= 0 Then
            MsgBox "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If (.Rows - 1) <= 0 Then
            MsgBox "Selecione um Produto !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If Len(Trim(.Cell(flexcpText, .RowSel, conCOL_SonProd_QtdProd))) = 0 Then
            MsgBox "Primeiro Informe a Qtde do Produto !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If Len(Trim(.Cell(flexcpText, .RowSel, conCOL_SonProd_Repeticao))) = 0 And _
           Len(Trim(.Cell(flexcpText, .RowSel, conCOL_SonProd_AltFilme))) = 0 And _
           Len(Trim(.Cell(flexcpText, .RowSel, conCOL_SonProd_FotNovo))) = 0 Then
           MsgBox "ATENÇÃO" & vbCrLf & _
                  "Primeiro Escolha uma ação para fotolito Repetição/Novo/Alteração !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
        
        
        If Len(Trim(.Cell(flexcpText, .RowSel, conCOL_SonProd_FechTpFr))) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Informe primeiro o tipo de fechamento !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If TravaEntregas(CCur(.Cell(flexcpText, .RowSel, conCOL_SonProd_QtdProd)), CLng(.Cell(flexcpText, .Row, conCOL_SonProd_IdProduto))) Then Exit Sub
        
        '' ======================================
        '' Verificando Campos em Branco
        If objBLBFunc.FcExisteLinhaVaziaFilho(grdProgEntrega, conCOL_SonProgEntr_QtdProd, conCOL_SonProgEntr_IdProduto, conCOL_SonProgEntr_Action2Do, .Cell(flexcpText, .Row, conCOL_SonProd_IdProduto)) = False Then
            MsgBox "ATENÇÃO - Primeiro informe a Quantidade do Produto !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If .Cell(flexcpText, .RowSel, conCOL_SonProd_Repeticao) = 1 And _
           .Cell(flexcpText, .RowSel, conCOL_SonProd_AltFilme) = 0 And _
           .Cell(flexcpText, .RowSel, conCOL_SonProd_FotNovo) = 0 Then
            If objBLBFunc.FcExisteLinhaVaziaFilho(grdProgEntrega, conCOL_SonProgEntr_DataEntrega, conCOL_SonProgEntr_IdProduto, conCOL_SonProgEntr_Action2Do, .Cell(flexcpText, .Row, conCOL_SonProd_IdProduto)) = False Then
                MsgBox "ATENÇÃO - Primeiro informe a Data de Entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Sub
            End If
        End If
        '' ======================================
        
        curPegaSaldo = 0
        If Len(Trim(.Cell(flexcpText, .RowSel, conCOL_SonProd_QtdProd))) > 0 Then
           curPegaSaldo = CCur(.Cell(flexcpText, .RowSel, conCOL_SonProd_QtdProd))
           For i = 1 To (grdProgEntrega.Rows - 1)
               If .Cell(flexcpText, .RowSel, conCOL_SonProd_IdProduto) = grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) And _
                  grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) <> dacEnumUpdateAction_delete Then
                    If Len(Trim(grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd))) > 0 Then
                         curQtdProdInf = curQtdProdInf + CCur(grdProgEntrega.Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd))
                    End If
               End If
           Next i
           curPegaSaldo = (curPegaSaldo - curQtdProdInf)
        End If
        
        If Len(Trim(grdProduto.Cell(flexcpText, grdProduto.RowSel, conCOL_SonProd_GrpPlanMestre))) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Informe ao PCP, não foi cadastrado o grupo de linha !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        grdProgEntrega.AddItem .Cell(flexcpText, .RowSel, conCOL_SonProd_IdProduto) & vbTab & _
                               curPegaSaldo & vbTab & _
                               "" & vbTab & _
                               dacEnumUpdateAction_Insert & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               0 & vbTab & _
                               grdProduto.Cell(flexcpData, grdProduto.RowSel, conCOL_SonProd_FechTpFr) & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               CLng(grdProduto.Cell(flexcpText, grdProduto.RowSel, conCOL_SonProd_GrpPlanMestre)) & vbTab & _
                               "" & vbTab & "" & vbTab & "" & vbTab & "" & dacEnumUpdateAction_Insert & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               ""
        
        lngSTATUSPROD = .FindRow(.Cell(flexcpText, .RowSel, conCOL_SonProd_IdProduto), , conCOL_SonProd_IdProduto)
        If lngSTATUSPROD > -1 Then
            If .Cell(flexcpText, lngSTATUSPROD, conCOL_SonProd_StatusProd) = 2 Then grdProgEntrega.Cell(flexcpText, (grdProgEntrega.Rows - 1), conCOL_SonProgEntr_StatusOP) = 4
        End If
        
        grdProgEntrega.Cell(flexcpText, (grdProgEntrega.Rows - 1), conCOL_SonProgEntr_DescStatusOP) = PegaStatusOP(grdProgEntrega.Cell(flexcpText, (grdProgEntrega.Rows - 1), conCOL_SonProgEntr_StatusOP))
        grdProgEntrega.Cell(flexcpText, (grdProgEntrega.Rows - 1), conCOL_SonProgEntr_QTDENOPALHET) = .Cell(flexcpText, .Row, conCOL_SonProd_QTDELATASPALLETS)
        grdProgEntrega.Cell(flexcpText, (grdProgEntrega.Rows - 1), conCOL_SonProgEntr_FechTpFr) = grdProduto.Cell(flexcpData, grdProduto.Row, conCOL_SonProd_FechTpFr)
        
        '' Depois retornar
        ''Call ConferePalhetsProgEntrg((grdProgEntrega.Rows - 1), CLng(Str(curPegaSaldo)))
        
        ''If .Cell(flexcpText, .Row, conCOL_SonProd_AltFilme) = 1 Or _
        ''   .Cell(flexcpText, .Row, conCOL_SonProd_FotNovo) = 1 Then
           ''grdProgEntrega.Cell(flexcpText, (grdProgEntrega.Rows - 1), conCOL_SonProgEntr_DataEntrega) = objCADPEDVENDA.DtEntregaLDTIME(mskDATAPED.Text)
        ''End If
        
        Call RefazIndice
        
    End With
    
    Exit Sub
    
Err_IncRegGridProg:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : IncRegGridProg()", Me.Name, "IncRegGridProg()", strCAMARQERRO)
    
    
End Sub

Private Function CalcTotEntregas(lngQtdTotal As Long, IdProduto As Long) As Boolean

On Error GoTo Err_CalcTotEntregas
    
    CalcTotEntregas = True
    
    Dim lngQtdGrd   As Long
    Dim i           As Integer
    Dim lngIndice   As Long

    lngQtdGrd = 0
    
    With grdProgEntrega
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) = IdProduto And _
               Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd))) > 0 And _
                (.Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Insert Or _
                .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update Or _
                .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Ignore) Then
                lngQtdGrd = lngQtdGrd + CLng(.Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd))
            End If
        Next i
    End With
    
    If lngQtdGrd > lngQtdTotal Then
        lngIndice = -1
        lngIndice = grdProduto.FindRow(IdProduto, , conCOL_SonProd_IdProduto)
        If lngIndice > -1 Then
            MsgBox "ATENÇÃO" & vbCrLf & "A soma do iten " & grdProduto.Cell(flexcpText, lngIndice, conCOL_SonProd_Codigo) & " - " & grdProduto.Cell(flexcpText, lngIndice, conCOL_SonProd_DescProd) & vbCrLf & _
            "na gride de programação de entrega não pode ser maior que a qtde total do Produto !!!", vbOKOnly + vbExclamation, "Aviso"
        End If
        CalcTotEntregas = False
    End If
    
    Exit Function
    
Err_CalcTotEntregas:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CalcTotEntregas()", Me.Name, "CalcTotEntregas()", strCAMARQERRO)
    
End Function


Private Function TravaEntregas(curQtdTotal As Currency, IdProduto As Long) As Boolean

On Error GoTo Err_TravaEntregas
    
    TravaEntregas = False
    
    Dim curQtdGrd   As Currency
    Dim i           As Integer

    curQtdGrd = 0
    
    With grdProgEntrega
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) = IdProduto And _
               Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd))) > 0 And _
                (.Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Insert Or _
                .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update Or _
                .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Ignore) Then
                curQtdGrd = curQtdGrd + .Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd)
            End If
        Next i
    End With
    
    If curQtdGrd >= curQtdTotal Then TravaEntregas = True
    
    Exit Function
    
Err_TravaEntregas:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : TravaEntregas()", Me.Name, "TravaEntregas()", strCAMARQERRO)
    
    
End Function

Private Sub CarregaPlanoEntrega()

On Error GoTo Err_CarregaPlanoEntrega
    
    Dim i                   As Long
    Dim dtDTLIDTIME         As Date
    Dim strDATA             As String
    Dim intAction2Do        As Integer
    Dim intStatusOP         As Integer
    Dim strCODLIN           As String
    Dim intLINHAPROD        As Integer
    Dim boolTemPData        As Boolean
    Dim strDTPREVMONTGEM    As String
    Dim strCODOP            As String
    
    arrENTREGAS = objCADPEDVENDA.PROGENTREGAS
    If IsArray(arrENTREGAS) Then
        With grdProgEntrega
            For i = 1 To UBound(arrENTREGAS)
                
                .AddItem arrENTREGAS(i, 0) & vbTab & _
                         arrENTREGAS(i, 1) & vbTab & _
                         Format(arrENTREGAS(i, 2), "DD/MM/YYYY") & vbTab & _
                         arrENTREGAS(i, 3) & vbTab & _
                         arrENTREGAS(i, 4) & vbTab & _
                         "" & vbTab & _
                         arrENTREGAS(i, 8) & vbTab & _
                         arrENTREGAS(i, 5) & vbTab & _
                         arrENTREGAS(i, 6) & vbTab & _
                         arrENTREGAS(i, 6) & vbTab & _
                         Format(arrENTREGAS(i, 2), "DD/MM/YYYY") & vbTab & _
                         arrENTREGAS(i, 7) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         dacEnumUpdateAction_Ignore & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         ""

                
                Call PegaOP2(Str(objCADPEDVENDA.CODPEDIDO), Trim(arrENTREGAS(i, 7)), i)
                
                If objCADPEDVENDA.STATUS <> "P" Then
                    .Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_StatusOP) = arrENTREGAS(i, 8)
                Else
                    arrENTREGAS(i, 8) = .Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_StatusOP)
                End If
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_GrpPlanMestre) = PegaGrdPMestre(.Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_IdProduto), grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_NECKIN))
        
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_QTDENOPALHET) = arrENTREGAS(i, 9)
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_PALHET) = arrENTREGAS(i, 10)
            
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_DescStatusOP) = PegaStatusOP(CLng(arrENTREGAS(i, 8)))
            
                strCODOP = Replace(.Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_CodOP), "/", "")
                Call PegaPrevMontagem(Trim(arrENTREGAS(i, 7)), Trim(strCODOP), (.Rows - 1))
                
                
                Call PegaStatusMontagem(Trim(arrENTREGAS(i, 7)), Trim(strCODOP), Trim(.Cell(flexcpText, (.Rows - 1), conCOL_SonProgEntr_CODIDPROG)), (.Rows - 1))
                
            
            Next i
        End With
    End If
    
    ''Call LidTimeProducao
    
    If (grdProduto.Rows - 1) > 0 Then
       grdProduto.Row = 1
       grdProduto.Col = 1
    End If

    Exit Sub
    
Err_CarregaPlanoEntrega:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CarregaPlanoEntrega()", Me.Name, "CarregaPlanoEntrega()", strCAMARQERRO)


End Sub

Private Sub DeletaGrdFilho(lngIDProduto As Long)

On Error GoTo Err_DeletaGrdFilho
    
    Dim i As Integer
    With grdProgEntrega
Retorna:
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) = lngIDProduto Then
                If .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Ignore Or _
                   .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update Then
                   .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_delete
                ElseIf .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_Insert Then
                    If (.Rows - 1) = 1 Then .Rows = 1
                    If (.Rows - 1) > 1 Then
                        .RemoveItem i
                        GoTo Retorna
                    End If
                End If
            End If
        Next i
    End With
    
    Exit Sub
    
Err_DeletaGrdFilho:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : DeletaGrdFilho()", Me.Name, "DeletaGrdFilho()", strCAMARQERRO)
    
    
End Sub

Private Sub ExcLinhaGrd(grdGenerico As VSFlexGrid, lngROL As Long, lngCOLAction2Do As Long)

On Error GoTo Err_ExcLinhaGrd
    
    With grdGenerico
        If .Cell(flexcpText, lngROL, lngCOLAction2Do) = dacEnumUpdateAction_Ignore Or _
           .Cell(flexcpText, lngROL, lngCOLAction2Do) = dacEnumUpdateAction_update Then
           .Cell(flexcpText, lngROL, lngCOLAction2Do) = dacEnumUpdateAction_delete
        ElseIf .Cell(flexcpText, lngROL, lngCOLAction2Do) = dacEnumUpdateAction_Insert Then
            If (.Rows - 1) = 1 Then .Rows = 1
            If (.Rows - 1) > 1 Then .RemoveItem lngROL
        End If
    End With
    
    Exit Sub
    
Err_ExcLinhaGrd:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : ExcLinhaGrd()", Me.Name, "ExcLinhaGrd()", strCAMARQERRO)
    
End Sub

Private Function VerificaProdPai(strSQLPAI As String) As Boolean

On Error GoTo Err_VerificaProdPai
    
    If BREC.State = 1 Then BREC.Close
    
    VerificaProdPai = False
    
    BREC.Open strSQLPAI, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then VerificaProdPai = True
    BREC.Close
    
    Exit Function
    
Err_VerificaProdPai:
    
    If BREC.State = 1 Then BREC.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : VerificaProdPai()", Me.Name, "VerificaProdPai()", strCAMARQERRO)
    
End Function

Private Sub VerifOrdFat(strIDPRODUTO As String, lngROW As Long, strCODOP As String)

On Error GoTo Err_VerifOrdFat
    
    If BREC10.State = 1 Then BREC10.Close
    If BREC7.State = 1 Then BREC7.Close
    
    Dim strModulo As String
    
    strModulo = ""
    If intFILIALPED = 1 Then strModulo = "_STEEL"
    
    Dim curQTDREAL   As Currency
    Dim curQTDFAT    As Currency
    Dim curSALDO     As Currency
    
    Dim curQTDROT    As Currency
    Dim curSALDOPED  As Currency
    Dim strNF        As String
    
    Call InitGridOrdemFat
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH" & strModulo & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED = " & objCADPEDVENDA.CODPEDIDO

    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       
       curSALDOPED = CCur(grdProduto.Cell(flexcpText, grdProduto.RowSel, conCOL_SonProd_QtdProd))
       
       Do While Not BREC10.EOF()
            
            '' ===========================================================
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADORDFATI" & strModulo & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODORD    = " & BREC10!SGI_CODORD & vbCrLf
            sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(strIDPRODUTO) & vbCrLf
            sSql = sSql & "   And SGI_CODORDFAB = " & Trim(strCODOP) & vbCrLf
            sSql = sSql & "   And SGI_QTDFAT Is Not Null" & vbCrLf
            sSql = sSql & "Order bY SGI_CODORD"
            
            
            BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
            Do While Not BREC7.EOF()
               
               If IsNull(BREC7!SGI_QTDJAFAT) Then
                  curQTDREAL = BREC7!SGI_QTDREAL
                  curQTDFAT = BREC7!SGI_QTDFAT
               Else
                  curQTDREAL = (BREC7!SGI_QTDREAL - BREC7!SGI_QTDJAFAT)
                  curQTDFAT = BREC7!SGI_QTDFAT
               End If
               
               curSALDO = (curQTDREAL - curQTDFAT)
               
               curSALDOPED = (curSALDOPED - curQTDFAT)
               
               With grdOrdFat
               
                    strNF = PegaCodNF(Str(BREC10!SGI_CODORD), strModulo)
                    
                    .AddItem BREC7!SGI_IDPRODUTO & vbTab & _
                             Format(BREC7!SGI_VLUNIT, "#,##0.00") & vbTab & _
                             curQTDREAL & vbTab & _
                             curQTDFAT & vbTab & _
                             curSALDO & vbTab & _
                             Format(BREC7!SGI_CODORD, "#/####") & vbTab & _
                             Format(BREC10!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                             dacEnumUpdateAction_Ignore & vbTab & _
                             Format(BREC7!SGI_CODORDFAB, "#/####") & vbTab & _
                             curSALDOPED & vbTab & _
                             strNF
               
               End With
               
               BREC7.MoveNext
            Loop
            BREC7.Close
            '' ===========================================================
            
            BREC10.MoveNext
       Loop
       
    End If
    BREC10.Close
    
    Exit Sub
    
Err_VerifOrdFat:

    If BREC10.State = 1 Then BREC10.Close
    If BREC7.State = 1 Then BREC7.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : VerifOrdFat()", Me.Name, "VerifOrdFat()", strCAMARQERRO)
    
    
End Sub

Private Sub LimpaCamposLabel()
    
On Error GoTo Err_LimpaCamposLabel
    
    lblDescVendedor.Caption = ""
    lblDescTpPed.Caption = ""
    lblDescCliente.Caption = ""
    lblDescCondPgto.Caption = ""
    lblDescTransp.Caption = ""
    
    Exit Sub
    
Err_LimpaCamposLabel:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : LimpaCamposLabel()", Me.Name, "LimpaCamposLabel()", strCAMARQERRO)
    
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label, strFUNCAOPAI As String)

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
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDescTabelas()" & vbCrLf & "Campo Nome : " & lblLabel.Name & vbCrLf & "Função Pai : " & strFUNCAOPAI, Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub

Private Sub AtivaDesativacampos(AtivaDesativa As Boolean)

On Error GoTo Err_AtivaDesativacampos
    
    txtCODVEND.Enabled = AtivaDesativa
    txtTIPPED.Enabled = AtivaDesativa
    Command2.Enabled = AtivaDesativa
    Command3.Enabled = AtivaDesativa
    lblDescVendedor.Enabled = AtivaDesativa
    lblDescTpPed.Enabled = AtivaDesativa

    Exit Sub
    
Err_AtivaDesativacampos:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : AtivaDesativacampos()", Me.Name, "AtivaDesativacampos()", strCAMARQERRO)
    

End Sub

Private Sub PegaDadosLabel()

On Error GoTo Err_PegaDadosLabel

        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVEND.Text, lblDescVendedor, "PegaDadosLabel()")
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADESPORCA", txtTIPPED.Text, lblDescTpPed, "PegaDadosLabel()")
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente, "PegaDadosLabel()")
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADCONDPGTO", txtCodCondPgto.Text, lblDescCondPgto, "PegaDadosLabel()")
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTRANSP", txtCODTRANSP.Text, lblDescTransp, "PegaDadosLabel()")
        If Len(Trim(txtCODMOTLIQ.Text)) > 0 Then Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMOTLIQ", txtCODMOTLIQ.Text, lblDescMotLiq, "PegaDadosLabel()")


        objCADPEDVENDA.NOMEVEND = Trim(lblDescVendedor.Caption)

    Exit Sub

Err_PegaDadosLabel:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDadosLabel()", Me.Name, "PegaDadosLabel()", strCAMARQERRO)

End Sub

Private Sub DesativasCampos()

On Error GoTo Err_DesativasCampos


    If lngCodVendedor > 0 Then
        Command2.Enabled = False
        txtCODVEND.Enabled = False
        Command3.Enabled = False
        txtTIPPED.Enabled = False
    End If

    Exit Sub

Err_DesativasCampos:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : DesativasCampos()", Me.Name, "DesativasCampos()", strCAMARQERRO)

End Sub

Private Sub InitGridOrdemFat()

    With grdOrdFat
    
       .Cols = conColumnsIn_SonOrdemFat
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonOrdemFat_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_IdProduto) = ""
       .ColDataType(conCOL_SonOrdemFat_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_VlUnit) = ""
       .ColDataType(conCOL_SonOrdemFat_VlUnit) = flexDTCurrency
       
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_QtdOP) = ""
       .ColDataType(conCOL_SonOrdemFat_QtdOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_QtdProd) = ""
       .ColDataType(conCOL_SonOrdemFat_QtdProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_Saldo) = ""
       .ColDataType(conCOL_SonOrdemFat_Saldo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_CodOrdem) = ""
       .ColDataType(conCOL_SonOrdemFat_CodOrdem) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_DatOrdem) = ""
       .ColDataType(conCOL_SonOrdemFat_DatOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_Action2Do) = ""
       .ColDataType(conCOL_SonOrdemFat_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_CodOP) = ""
       .ColDataType(conCOL_SonOrdemFat_CodOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_SaldoPed) = ""
       .ColDataType(conCOL_SonOrdemFat_SaldoPed) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_NF) = ""
       .ColDataType(conCOL_SonOrdemFat_NF) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_DataNF) = ""
       .ColDataType(conCOL_SonOrdemFat_DataNF) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonOrdemFat_CodORDFAT) = ""
       .ColDataType(conCOL_SonOrdemFat_CodORDFAT) = flexDTLong
       
       .ColWidth(conCOL_SonOrdemFat_IdProduto) = 0
       .ColWidth(conCOL_SonOrdemFat_VlUnit) = 0
       .ColWidth(conCOL_SonOrdemFat_QtdOP) = 750
       .ColWidth(conCOL_SonOrdemFat_QtdProd) = 750
       .ColWidth(conCOL_SonOrdemFat_Saldo) = 750
       .ColWidth(conCOL_SonOrdemFat_CodOrdem) = 1300
       .ColWidth(conCOL_SonOrdemFat_DatOrdem) = 1300
       .ColWidth(conCOL_SonOrdemFat_Action2Do) = 0
       .ColWidth(conCOL_SonOrdemFat_CodOP) = 0
       .ColWidth(conCOL_SonOrdemFat_SaldoPed) = 0
       .ColWidth(conCOL_SonOrdemFat_NF) = 1000
       .ColWidth(conCOL_SonOrdemFat_DataNF) = 1000
       .ColWidth(conCOL_SonOrdemFat_CodORDFAT) = 1000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub PopGrdOrdFat_Prod(strIDPRODUTO As String, lngROW As Long, strCODOP As String)

On Error GoTo Err_PopGrdOrdFat_Prod
            
            
            If Len(Trim(strCODOP)) = 0 Then Exit Sub
            
            If objCADPEDVENDA.STATUS = "F" Or _
               objCADPEDVENDA.STATUS = "P" Or _
               objCADPEDVENDA.STATUS = "L" Or objCADPEDVENDA.STATUS = "M" Then
               Call VerifOrdFat(Trim(strIDPRODUTO), lngROW, strCODOP)
            End If

    Exit Sub
    
Err_PopGrdOrdFat_Prod:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PopGrdOrdFat_Prod()", Me.Name, "PopGrdOrdFat_Prod()", strCAMARQERRO)

End Sub



Private Sub InitGridConfFat()

    With grdConfFat
       
       .Cols = conColumnsIn_SonConfFat
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonConfFat_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonConfFat_IdProduto) = ""
       .ColDataType(conCOL_SonConfFat_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonConfFat_CodOrdem) = ""
       .ColDataType(conCOL_SonConfFat_CodOrdem) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonConfFat_CodConf) = ""
       .ColDataType(conCOL_SonConfFat_CodConf) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonConfFat_QtdProd) = ""
       .ColDataType(conCOL_SonConfFat_QtdProd) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonConfFat_VlUnit) = ""
       .ColDataType(conCOL_SonConfFat_VlUnit) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonConfFat_NF) = ""
       .ColDataType(conCOL_SonConfFat_NF) = flexDTLong
       
       .ColWidth(conCOL_SonConfFat_IdProduto) = 0
       .ColWidth(conCOL_SonConfFat_CodOrdem) = 0
       .ColWidth(conCOL_SonConfFat_CodConf) = 750
       .ColWidth(conCOL_SonConfFat_QtdProd) = 750
       .ColWidth(conCOL_SonConfFat_VlUnit) = 1000
       .ColWidth(conCOL_SonConfFat_NF) = 1000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub PopGrdConfFat(strCODORD As String, strIDPRODUTO As String)

On Error GoTo Err_PopGrdConfFat
    
    If Len(Trim(strCODORD)) = 0 Then Exit Sub
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDCONFH " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODROD = " & Trim(strCODORD)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
    
    
    End If
    BREC.Close
    
    Exit Sub
    
Err_PopGrdConfFat:
    
    If BREC.State = 1 Then BREC.Close
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PopGrdConfFat()", Me.Name, "PopGrdConfFat()", strCAMARQERRO)
    
    
End Sub

Private Function ValidaFornecedor(strCODFORNEC As String) As Boolean
    
On Error GoTo Err_ValidaFornecedor
     
     If BREC7.State = 1 Then BREC7.Close
     
     ValidaFornecedor = True
     
     If Len(Trim(strCODFORNEC)) = 0 Then Exit Function
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_CODIGO = " & Trim(strCODFORNEC)
     
     BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC7.EOF() Then
        ValidaFornecedor = False
     Else
        MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbExclamation, "Aviso"
     End If
     BREC7.Close
     
     Exit Function
     
Err_ValidaFornecedor:

     If BREC7.State = 1 Then BREC7.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : ValidaFornecedor()", Me.Name, "ValidaFornecedor()", strCAMARQERRO)
     
     
End Function

Private Sub DestroiObjeto()
    Set objBLBFunc = Nothing
    Set objCADPEDVENDA = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub LimpaCamposDadosAdicionais()

On Error GoTo Err_LimpaCamposDadosAdicionais

    lblDescFecham.Caption = ""
    lblDescCorpo.Caption = ""
    lblDescTampa.Caption = ""
    lblDescFundo.Caption = ""
    lblDescArgola.Caption = ""
    lblDescFechTPFURO.Caption = ""
    lblDescAltFilme.Caption = ""
    lblFotNovo.Caption = ""
    lblDescRepet.Caption = "'"

    Exit Sub
    
Err_LimpaCamposDadosAdicionais:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : LimpaCamposDadosAdicionais()", Me.Name, "LimpaCamposDadosAdicionais()", strCAMARQERRO)
    
    

End Sub

Private Sub PegadadosGrid(lngROW As Long)
    
On Error GoTo Err_PegadadosGrid
    
    With grdProduto
        
        lblDescFecham.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_Fechamento)
        lblDescCorpo.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_Corpo)
        lblDescTampa.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_Tampa)
        lblDescFundo.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_Fundo)
        
        lblDescArgola.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_Argola)
        lblDescFechTPFURO.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_FechTpFr)
        lblDescAltFilme.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_AltFilme)
        lblFotNovo.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_FotNovo)
        lblDescRepet.Caption = .Cell(flexcpTextDisplay, lngROW, conCOL_SonProd_Repeticao)
        
    End With
    
    Exit Sub
    
Err_PegadadosGrid:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegadadosGrid()", Me.Name, "PegadadosGrid()", strCAMARQERRO)
    
    
End Sub

Private Sub LimpaCampoSaldoRot()
    lblSaldRot.Caption = ""
    lblSaldoJaFat.Caption = ""
End Sub

Private Sub SaldoRotulo(strCODPED As String, strIDPRODUTO As String, strQtdRot As String)

On Error GoTo Err_SaldoRotulo
    
    If BREC.State = 1 Then BREC.Close
    
    Dim dblTotalFat As Double
    Dim dblSaldo    As Double
    Dim dblQTDROT   As Double
    Dim strModulo   As String
    
    strModulo = ""
    If intFILIALPED = 1 Then strModulo = "_STEEL"
    
    dblTotalFat = 0
    dblSaldo = 0
    dblQTDROT = 0
    
    If Len(Trim(strQtdRot)) > 0 Then dblQTDROT = CDbl(strQtdRot)
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(ORDFATI.SGI_QTDFAT) As SGI_TOTQTDFAT " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH" & strModulo & " ORDFATC " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strModulo & " ORDFATI " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDFATC.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDFATC.SGI_CODPED    = " & strCODPED & vbCrLf
    sSql = sSql & "   And ORDFATI.SGI_FILIAL    = ORDFATC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ORDFATI.SGI_CODORD    = ORDFATC.SGI_CODORD" & vbCrLf
    sSql = sSql & "   And ORDFATI.SGI_IDPRODUTO = " & strIDPRODUTO

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       If Not IsNull(BREC!SGI_TOTQTDFAT) Then dblTotalFat = BREC!SGI_TOTQTDFAT
    End If
    BREC.Close
    
    dblSaldo = (dblQTDROT - dblTotalFat)
    
    lblSaldoJaFat.Caption = Trim(Str(dblTotalFat))
    lblSaldRot.Caption = Trim(Str(dblSaldo))
    
    Exit Sub

Err_SaldoRotulo:
    
    If BREC.State = 1 Then BREC.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : VerifOrdFat()", Me.Name, "VerifOrdFat()", strCAMARQERRO)
    
End Sub

Private Sub PegaOP(strCODPED As String, strIDPRODUTO As String, strDTENTREGA As String, strQtdePed As String, lngLINHA As Long, strINDICE As String)

On Error GoTo Err_PegaOP
    
    If BREC12.State = 1 Then BREC12.Close
    
    Dim strModulo As String
    strModulo = ""
    If intFILIALPED = 1 Then strModulo = "_STEEL"
    
    If Len(Trim(strCODPED)) = 0 Then Exit Sub
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Sub
    If Len(Trim(strDTENTREGA)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strModulo & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED     = " & Trim(strCODPED) & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO  = " & Trim(strIDPRODUTO) & vbCrLf
    sSql = sSql & "   And SGI_DATENTREGA = '" & Trim(strDTENTREGA) & "'" & vbCrLf
    ''sSql = sSql & "   And SGI_QTDEPED    = " & Trim(strQtdePed)
    
    BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC12.EOF() Then
       grdProgEntrega.Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_CodOP) = Format(BREC12!SGI_CODIGO, "#/####")
       grdProgEntrega.Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_StatusOP) = BREC12!SGI_STATUS
       
       If BREC12!SGI_STATUS = 0 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_StatusOP) = &H8080FF '' Vermelho
       If BREC12!SGI_STATUS = 1 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_StatusOP) = &H80FFFF '' Amarelo
       If BREC12!SGI_STATUS = 2 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_StatusOP) = &H80FF80 '' Verde
    
    End If
    BREC12.Close

    Exit Sub
    
Err_PegaOP:

    If BREC12.State = 1 Then BREC12.Close
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaOP()", Me.Name, "PegaOP()", strCAMARQERRO)

End Sub

Private Sub LimpaSaldoPedido()
    
On Error GoTo err_LimpaSaldoPedido

    lblTotGer(0).Caption = ""
    lblTotGer(1).Caption = ""
    lblTotGer(2).Caption = ""
    
    Exit Sub
    
err_LimpaSaldoPedido:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : LimpaSaldoPedido()", Me.Name, "LimpaSaldoPedido()", strCAMARQERRO)

End Sub

Private Sub GeraSaldoPedido(strCODPED As String)

On Error GoTo Err_GeraSaldoPedido
    
    If BREC10.State = 1 Then BREC10.Close
    If BREC11.State = 1 Then BREC11.Close
    
    If Len(Trim(strCODPED)) = 0 Then Exit Sub
    
    
    Dim curTOTPED       As Currency
    Dim curQTDFAT       As Currency
    Dim curSALDOPED     As Currency

    curTOTPED = 0
    curQTDFAT = 0
    curSALDOPED = 0
    
    '' Total do Pedido
    sSql = "Select" & vbCrLf
    sSql = sSql & "      Sum(SGI_QTDE) As SGI_QTDE" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADPEDVENDI" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And SGI_CODIGO = " & Trim(strCODPED)
    
    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then
        If Not IsNull(BREC11!SGI_QTDE) Then curTOTPED = BREC11!SGI_QTDE
    End If
    BREC11.Close
    
    
    '' Total já faturado
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(ORDFATI.SGI_QTDFAT) As SGI_TOTQTDFAT " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH ORDFATC " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI ORDFATI " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDFATC.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDFATC.SGI_CODPED    = " & Trim(strCODPED) & vbCrLf
    sSql = sSql & "   And ORDFATI.SGI_FILIAL    = ORDFATC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ORDFATI.SGI_CODORD    = ORDFATC.SGI_CODORD" & vbCrLf

    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
        If Not IsNull(BREC10!SGI_TOTQTDFAT) Then curQTDFAT = BREC10!SGI_TOTQTDFAT
    End If
    BREC10.Close
    
    curSALDOPED = (curTOTPED - curQTDFAT)
    
    lblTotGer(0).Caption = curTOTPED
    lblTotGer(1).Caption = curQTDFAT
    lblTotGer(2).Caption = curSALDOPED

    Exit Sub

Err_GeraSaldoPedido:

    If BREC10.State = 1 Then BREC10.Close
    If BREC11.State = 1 Then BREC11.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : GeraSaldoPedido()", Me.Name, "GeraSaldoPedido()", strCAMARQERRO)
    

End Sub

Private Function PegaPlanMestre(strCODPROD As String, strMES As String, strANO As String, strDTENTREGA As String, lngQTDENTREGA As Long) As Long

On Error GoTo Err_PegaPlanMestre
    
    PegaPlanMestre = 0
    
    If BREC.State = 1 Then BREC.Close
    
    If Len(Trim(strCODPROD)) = 0 Then Exit Function
    If Len(Trim(strDTENTREGA)) = 0 Then Exit Function
    
    Dim lngPLANMESTRE  As Long
    Dim lngOPEMITIDA   As Long
    Dim lngCodLinha    As Long
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PM.SGI_QTDE" & vbCrLf
    sSql = sSql & "      ,DI.SGI_QTDE As SGI_QTDE_PM" & vbCrLf
    sSql = sSql & "      ,PR.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPLANMESTRE" & strNOMFILIAL & "   PM" & vbCrLf
    sSql = sSql & "      ,SGI_CADDIASPMSEMANA" & strNOMFILIAL & " DI" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO LI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO      PR" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PR.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PR.SGI_IDPRODUTO = " & Trim(strCODPROD) & vbCrLf
    sSql = sSql & "   And LI.SGI_FILIAL    = PR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And LI.SGI_CODLIN    = PR.SGI_CODLINPROD " & vbCrLf
    sSql = sSql & "   And PM.SGI_FILIAL    = LI.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PM.SGI_CODLINHA  = LI.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And PM.SGI_MES       = " & Trim(strMES) & vbCrLf
    sSql = sSql & "   And PM.SGI_ANO       = " & Trim(strANO) & vbCrLf
    sSql = sSql & "   And DI.SGI_FILIAL    = PM.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And DI.SGI_CODIGO    = PM.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And DI.SGI_DTSEMANA  = '" & Trim(strDTENTREGA) & "'"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       lngPLANMESTRE = BREC!SGI_QTDE_PM
       lngCodLinha = BREC!SGI_CODLINPROD
    End If
    BREC.Close
    
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Sum(OP.SGI_QTDE) As SGI_QTDE" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PR" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & "  OP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PR.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And PR.SGI_CODLINPROD = " & lngCodLinha & vbCrLf
    sSql = sSql & "   And OP.SGI_FILIAL     = PR.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And OP.SGI_IDPRODUTO  = PR.SGI_IDPRODUTO" & vbCrLf
    ''sSql = sSql & "   And OP.SGI_OPENVIADA is Null" & vbCrLf
    sSql = sSql & "   And OP.SGI_DATENTREGA = '" & Trim(strDTENTREGA) & "'" & vbCrLf
    sSql = sSql & "   And OP.SGI_STATUS     = 0"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDE) Then lngOPEMITIDA = BREC!SGI_QTDE
    End If
    BREC.Close
    
    PegaPlanMestre = (lngPLANMESTRE - (lngOPEMITIDA + lngQTDENTREGA))

    Exit Function
    
Err_PegaPlanMestre:

    If BREC.State = 1 Then BREC.Close
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaPlanMestre()", Me.Name, "PegaPlanMestre()", strCAMARQERRO)

End Function

Private Function PegaQtdeEntrega(strIDPRODUTO As String, strMES As String, strANO As String, strDia As String) As Long
    
On Error GoTo Err_PegaQtdeEntrega
    
    PegaQtdeEntrega = 0
    
    Dim i       As Integer
    
    With grdProgEntrega
        For i = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega))) > 0 And _
               Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd))) > 0 Then
               
                If Trim(strIDPRODUTO) = Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_GrpPlanMestre)) And _
                   Trim(strMES) = Trim(Str(Month(CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega))))) And _
                   Trim(strANO) = Trim(Str(Year(CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega))))) And _
                   CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega)) = CDate(strDia) Then
                    
                    PegaQtdeEntrega = PegaQtdeEntrega + CLng(.Cell(flexcpText, i, conCOL_SonProgEntr_QtdProd))
                End If
            End If
        Next i
    End With
    
    Exit Function
    
Err_PegaQtdeEntrega:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaQtdeEntrega()", Me.Name, "PegaQtdeEntrega()", strCAMARQERRO)
    
End Function

Private Function ChamaSenhaUsuario() As Boolean
    
On Error GoTo Err_ChamaSenhaUsuario

    ChamaSenhaUsuario = False
    
    frmUSULIB.cCaminho = cCaminho
    frmUSULIB.Linha = Linha
    frmUSULIB.iCodigo = iCodigo
    frmUSULIB.FILIAL = FILIAL
    frmUSULIB.strACESSO = strACESSO
    frmUSULIB.strMODPAI = Me.Name
    frmUSULIB.strUSUARIO = strUSUARIO
    frmUSULIB.lngCodVendedor = lngCodVendedor
    frmUSULIB.Show vbModal
    
    ChamaSenhaUsuario = frmUSULIB.boolLib

    Exit Function
    
Err_ChamaSenhaUsuario:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : ChamaSenhaUsuario()", Me.Name, "ChamaSenhaUsuario()", strCAMARQERRO)

End Function

Private Sub PopGrdProdutos()

On Error GoTo Err_PopGrdProdutos
        
    Dim i As Integer
    
    '' Produtos
    arrPRODCOPIA = objCADPEDVENDA.PRODUTOS
    arrItensPedido = objCADPEDVENDA.PRODUTOS
    
    If IsArray(arrItensPedido) Then
        With grdProduto
            For i = 1 To UBound(arrItensPedido)
                
                .AddItem arrItensPedido(i, 10) & vbTab & arrItensPedido(i, 1) & vbTab & "" & vbTab & _
                         arrItensPedido(i, 9) & vbTab & arrItensPedido(i, 2) & vbTab & _
                         Format(arrItensPedido(i, 3), "#,##0.00") & vbTab & _
                         Format(arrItensPedido(i, 4), "#,##0.00") & vbTab & _
                         Format(arrItensPedido(i, 5), "#,##0.00") & vbTab & _
                         Format(arrItensPedido(i, 6), "#,##0.00") & vbTab & _
                         Format(arrItensPedido(i, 7), "#,##0.00") & vbTab & _
                         Format(arrItensPedido(i, 8), "#,##0.00") & vbTab & _
                         Format((arrItensPedido(i, 2) * arrItensPedido(i, 3)), "#,##0.00") & vbTab & _
                         arrItensPedido(i, 11) & vbTab & arrItensPedido(i, 12) & vbTab & _
                         arrItensPedido(i, 13) & vbTab & arrItensPedido(i, 21) & vbTab & _
                         "" & vbTab & arrItensPedido(i, 14) & vbTab & _
                         arrItensPedido(i, 15) & vbTab & _
                         "" & vbTab & "" & vbTab & _
                         arrItensPedido(i, 17) & vbTab & arrItensPedido(i, 18) & vbTab & _
                         arrItensPedido(i, 19) & vbTab & arrItensPedido(i, 20) & vbTab & Trim(arrItensPedido(i, 22)) & vbTab & _
                         arrItensPedido(i, 10) & vbTab & Format(arrItensPedido(i, 3), "#,##0.00") & vbTab & arrItensPedido(i, 2) & vbTab & _
                         arrItensPedido(i, 16) & vbTab & arrItensPedido(i, 17) & vbTab & arrItensPedido(i, 18) & vbTab & _
                         arrItensPedido(i, 19) & vbTab & dacEnumUpdateAction_Ignore & vbTab & "N" & vbTab & arrItensPedido(i, 23) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & 0 & vbTab & ""
            
                If TemOP(Trim(Str(arrItensPedido(i, 10)))) = True Then .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_TemOP) = "S"
                Call CorRotulo(i)
                
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_NECKIN) = objCADPEDVENDA.PegaNECKIN(.Cell(flexcpText, (.Rows - 1), conCOL_SonProd_IdProduto))
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_HOMOLOGADO) = objCADPEDVENDA.PegaHOMOLOGADO(.Cell(flexcpText, (.Rows - 1), conCOL_SonProd_IdProduto))
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_GrpPlanMestre) = PegaGrdPMestre(.Cell(flexcpText, (.Rows - 1), conCOL_SonProd_IdProduto), CLng(.Cell(flexcpText, (.Rows - 1), conCOL_SonProd_NECKIN)))
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_CodCapacidade) = PegaGrdCodCapac(.Cell(flexcpText, (.Rows - 1), conCOL_SonProd_IdProduto))
                
                .Cell(flexcpData, (.Rows - 1), conCOL_SonProd_FechTpFr) = arrItensPedido(i, 16)
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_FechTpFr) = PegaDescTabelasGrd("SGI_CODIGO", "SGI_DESCRI", "SGI_CADFECHAM", Trim(Str(arrItensPedido(i, 16))))
                
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_QTDELATASPALLETS) = arrItensPedido(i, 24)
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_PALLETS) = arrItensPedido(i, 25)
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_Conferido) = arrItensPedido(i, 26)
                .Cell(flexcpText, (.Rows - 1), conCOL_SonProd_PalhetPadrao) = arrItensPedido(i, 27)
            
            Next i
            .Row = 1
        End With
    End If
    Call CalcTotPedido
    
    Exit Sub
        
Err_PopGrdProdutos:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PopGrdProdutos()", Me.Name, "PopGrdProdutos()", strCAMARQERRO)
        

End Sub

Private Function TemOP(strIDPRODUTO As String) As Boolean

    
On Error GoTo Err_TemOP

    TemOP = False
    
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
    
    If BREC.State = 1 Then BREC.Close
    
    Dim strModulo As String
    strModulo = ""
    If intFILIALPED = 1 Then strModulo = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strModulo & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED    = " & objCADPEDVENDA.CODPEDIDO & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPRODUTO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then TemOP = True
    BREC.Close
    
    Exit Function

Err_TemOP:
    
    If BREC.State = 1 Then BREC.Close
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : TemOP()", Me.Name, "TemOP()", strCAMARQERRO)

End Function

Private Sub LimpaCamposGrid(lngROW As Long)
    
On Error GoTo Err_LimpaCamposGrid
    
    With grdProduto
        .Cell(flexcpText, lngROW, conCOL_SonProd_Codigo) = ""
        .Cell(flexcpText, lngROW, conCOL_SonProd_DescProd) = ""
        .Cell(flexcpText, lngROW, conCOL_SonProd_CodLinProd) = ""
        .Cell(flexcpText, lngROW, conCOL_SonProd_HOMOLOGADO) = ""
    End With
    
    Exit Sub
    
Err_LimpaCamposGrid:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : LimpaCamposGrid()", Me.Name, "LimpaCamposGrid()", strCAMARQERRO)
    
End Sub

Private Sub PosCol(lngPOSCOL As Long, lngPOSROL As Long)
    
On Error GoTo Err_PosCol
    
    grdProduto.SetFocus
    grdProduto.Row = lngPOSROL
    grdProduto.Col = lngPOSCOL
    
    Exit Sub
    
Err_PosCol:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PosCol()", Me.Name, "PosCol()", strCAMARQERRO)
    
End Sub

Private Sub MostraDados()
    
On Error GoTo Err_MostraDados

    
    With grdProduto
        txtOBSROT.Text = ""
        Call LimpaCamposDadosAdicionais
        If (.Rows - 1) > 0 And .RowSel > 0 Then
            Dim lngCODIDPROD    As Long
            Dim strCODOP        As String
            lngCODIDPROD = 0
            strCODOP = ""
            If Len(Trim(.Cell(flexcpText, .RowSel, conCOL_SonProd_IdProduto))) > 0 Then lngCODIDPROD = CLng(.Cell(flexcpText, .RowSel, conCOL_SonProd_IdProduto))
            
            Call objBLBFunc.CarregaDadosGrdFilho(grdProgEntrega, conCOL_SonProgEntr_Action2Do, conCOL_SonProgEntr_IdProduto, lngCODIDPROD)
            Call MostraDadosProgEntr
            
            Call InitGridOrdemFat
            Call SaldoRotulo(Str(objCADPEDVENDA.CODPEDIDO), Str(lngCODIDPROD), grdProduto.Cell(flexcpText, .RowSel, conCOL_SonProd_QtdProd))
            
            txtOBSROT.Text = .Cell(flexcpText, .RowSel, conCOL_SonProd_OBSOP)
            Call PegadadosGrid(.RowSel)
        End If
    End With
    
    Exit Sub
    
Err_MostraDados:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : MostraDados()", Me.Name, "MostraDados()", strCAMARQERRO)
    
End Sub

Private Sub RefazIndice()
    
On Error GoTo Err_RefazIndice
    
    Dim i       As Integer
    Dim lngREGS As Long
    With grdProgEntrega
        lngREGS = 0
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) <> dacEnumUpdateAction_delete Then
                lngREGS = (lngREGS + 1)
                .Cell(flexcpText, i, conCOL_SonProgEntr_INDICE) = lngREGS
            End If
        Next i
    End With
    
    Exit Sub
    
Err_RefazIndice:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : RefazIndice()", Me.Name, "RefazIndice()", strCAMARQERRO)
    
End Sub


Private Sub MudaActio2DoFilho(grdGenerica As VSFlexGrid, lngCOLActio2Do As Long, lngCOLINDICE As Long, strVALORPAI As String)
    
On Error GoTo Err_MudaActio2DoFilho
    
    Dim i As Integer
    With grdGenerica
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, lngCOLActio2Do) <> dacEnumUpdateAction_delete Then
                If Trim(.Cell(flexcpText, i, lngCOLINDICE)) = Trim(strVALORPAI) Then
                    If .Cell(flexcpText, i, lngCOLActio2Do) = dacEnumUpdateAction_Ignore Then .Cell(flexcpText, i, lngCOLActio2Do) = dacEnumUpdateAction_update
                End If
            End If
        Next i
    End With
    
    Exit Sub
    
Err_MudaActio2DoFilho:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : MudaActio2DoFilho()", Me.Name, "MudaActio2DoFilho()", strCAMARQERRO)
    
End Sub

Private Sub CorRotulo(lngROW As Integer)
    
On Error GoTo Err_CorRotulo
    With grdProduto
        If .Cell(flexcpText, lngROW, conCOL_SonProd_StatusProd) <> 2 Then
           .Cell(flexcpBackColor, lngROW, 0, lngROW, (grdProduto.Cols - 1)) = &H80000005
        ElseIf .Cell(flexcpText, lngROW, conCOL_SonProd_StatusProd) = 2 Then
           .Cell(flexcpBackColor, lngROW, 0, lngROW, (grdProduto.Cols - 1)) = &H8080FF
        End If
    End With
    Exit Sub
    
Err_CorRotulo:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CorRotulo()", Me.Name, "CorRotulo()", strCAMARQERRO)
    
End Sub


Private Sub PegaOP2(strCODPED As String, strIDPai As String, lngLINHA As Long)

On Error GoTo Err_PegaOP2
    
    If BREC12.State = 1 Then BREC12.Close
    
    Dim strModulo As String
    strModulo = ""
    If intFILIALPED = 1 Then strModulo = "_STEEL"
    
    If Len(Trim(strCODPED)) = 0 Then Exit Sub
    If Len(Trim(strIDPai)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strModulo & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED     = " & Trim(strCODPED) & vbCrLf
    sSql = sSql & "   And SGI_IDPAI      = " & Trim(strIDPai) & vbCrLf
    
    BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC12.EOF() Then
       grdProgEntrega.Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_CodOP) = Format(BREC12!SGI_CODIGO, "#/####")
       grdProgEntrega.Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_StatusOP) = BREC12!SGI_STATUS
       
       If BREC12!SGI_STATUS = 0 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_DescStatusOP) = &H8080FF '' Vermelho
       If BREC12!SGI_STATUS = 1 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_DescStatusOP) = &H80FFFF '' Amarelo
       If BREC12!SGI_STATUS = 2 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_DescStatusOP) = &H80FF80 '' Verde
       If BREC12!SGI_STATUS = 9 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_DescStatusOP) = &H80FF&  '' Abobora

''       If BREC12!SGI_STATUS = 6 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_DescStatusOP) = &H80FF&  '' P.Cota
''       If BREC12!SGI_STATUS = 7 Then grdProgEntrega.Cell(flexcpBackColor, lngLINHA, conCOL_SonProgEntr_IdProduto, lngLINHA, conCOL_SonProgEntr_DescStatusOP) = &HFFFF00     '' P.Data
    
    End If
    BREC12.Close

    Exit Sub
    
Err_PegaOP2:

    If BREC12.State = 1 Then BREC12.Close
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaOP()", Me.Name, "PegaOP2()", strCAMARQERRO)

End Sub

Private Function PermiteLibPDataPCota() As Boolean

    PermiteLibPDataPCota = False
    
    If lngCodUsuario = 0 Then
       PermiteLibPDataPCota = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBPDATAPCOTA" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO         = " & lngCodUsuario & vbCrLf
    sSql = sSql & "   And SGI_LIBPDATAPCOTA  = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibPDataPCota = True
    BREC.Close

End Function


Private Sub AtivaDesativaBotoes()

On Error GoTo Err_AtivaDesativaBotoes
    
    If cTipOper = "LC" Then
       cmdLIBPDATAPCOTA.Visible = PermiteLibPDataPCota
    Else
       cmdLIBPDATAPCOTA.Visible = False
    End If
    
    Exit Sub
    
Err_AtivaDesativaBotoes:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : AtivaDesativaBotoes()", Me.Name, "AtivaDesativaBotoes()", strCAMARQERRO)
    
End Sub


Private Function PegaStatusOP(lngSTATUS As Long) As String

    PegaStatusOP = ""
    
    If lngSTATUS = 0 Then PegaStatusOP = "Liberado"
    If lngSTATUS = 1 Then PegaStatusOP = "Fat.Parcial"
    If lngSTATUS = 2 Then PegaStatusOP = "Fat.Total"
    If lngSTATUS = 3 Then PegaStatusOP = "Reprovado"
    If lngSTATUS = 4 Then PegaStatusOP = "Bloqueado"
    If lngSTATUS = 6 Then PegaStatusOP = "P.Cota"
    If lngSTATUS = 7 Then PegaStatusOP = "P.Data"
    If lngSTATUS = 9 Then PegaStatusOP = "Liq.Man"

End Function

Private Function PegaPedBloq(strCODLIN As String, strDTENTREGA As String) As Long

        PegaPedBloq = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDVH" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & objCADPEDVENDA.CODPEDIDO
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And (PEDVH.SGI_STATUS     = 'B' Or PEDVH.SGI_STATUS = 'N')" & vbCrLf
        
        BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC5.EOF() Then
           If Not IsNull(BREC5!SGI_QTDE) Then PegaPedBloq = BREC5!SGI_QTDE
        End If
        BREC5.Close

End Function

Private Function PegaOPDia(strCODLIN As String, strDTENTREGA As String) As Long

        PegaOPDia = 0
        
        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function
        
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(OP.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO      PROD" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & " OP" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PROD.SGI_CODLINPROD = " & strCODLIN & vbCrLf
        sSql = sSql & "   And OP.SGI_FILIAL       = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And OP.SGI_IDPRODUTO    = PROD.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And OP.SGI_DATENTREGA   = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And OP.SGI_STATUS       = 0"
        
        BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC3.EOF() Then
            If Not IsNull(BREC3!SGI_QTDE) Then PegaOPDia = BREC3!SGI_QTDE
        End If
        BREC3.Close

End Function

Private Function PegaPedBloqAlt(strCODLIN As String, strDTENTREGA As String) As Long

        PegaPedBloqAlt = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDVH" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & objCADPEDVENDA.CODPEDIDO
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_STATUS     = 'S'" & vbCrLf
        
        BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC5.EOF() Then
           If Not IsNull(BREC5!SGI_QTDE) Then PegaPedBloqAlt = BREC5!SGI_QTDE
        End If
        BREC5.Close

End Function

Private Function PegaPedBloqFot(strCODLIN As String, strDTENTREGA As String) As Long

        PegaPedBloqFot = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & "  PEDVH" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & objCADPEDVENDA.CODPEDIDO
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_STATUS     = 'V'" & vbCrLf
        
        BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC5.EOF() Then
           If Not IsNull(BREC5!SGI_QTDE) Then PegaPedBloqFot = BREC5!SGI_QTDE
        End If
        BREC5.Close

End Function

Private Function PegaPedBloqPcPd(strCODLIN As String, strDTENTREGA As String) As Long

        PegaPedBloqPcPd = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & "  PEDVH" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & strCODLIN & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & objCADPEDVENDA.CODPEDIDO
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_STATUS     = 'C'" & vbCrLf
        
        BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC5.EOF() Then
           If Not IsNull(BREC5!SGI_QTDE) Then PegaPedBloqPcPd = BREC5!SGI_QTDE
        End If
        BREC5.Close

End Function

Private Function PegaGrdPMestre(strIDPRODUTO As String, lngNECKIN As Long) As String

    PegaGrdPMestre = ""
    
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       GRP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO       PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO  LINP" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHAIT" & strNOMFILIAL & "   GRPI" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHA" & strNOMFILIAL & "     GRP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL           = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO        = " & strIDPRODUTO & vbCrLf
    sSql = sSql & "   And LINP.SGI_FILIAL           = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINP.SGI_CODLIN           = PROD.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "   And GRPI.SGI_FILIAL           = LINP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPI.SGI_CODLIN           = LINP.SGI_CODIGO" & vbCrLf
    ''sSql = sSql & "   And GRPI.SGI_OPTCOMNECKINSN   = " & lngNECKIN & vbCrLf
    sSql = sSql & "   And GRP.SGI_FILIAL            = GRPI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRP.SGI_CODIGO            = GRPI.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And GRP.SGI_ATIVO             = 1"
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC8.EOF()
        PegaGrdPMestre = PegaGrdPMestre & Trim(Str(BREC8!SGI_CODIGO))
        BREC8.MoveNext
    Loop
    BREC8.Close
    
End Function

Private Function ConfereCotas(strIDPRODUTO As String, strDTENTREGA As String) As Boolean

    ConfereCotas = True
    
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
    If Len(Trim(strDTENTREGA)) = 0 Then Exit Function
    
    Dim lngSALDO                As Long
    Dim lngSALDOQTDENTR         As Long
    Dim lngQTDE                 As Long
    Dim lngQTDEOP               As Long
    Dim lngQTDEPEDBLOQ          As Long
    Dim lngQTDEPEDBLOQ2         As Long
    Dim lngQTDEPEDBLOQALT       As Long
    Dim lngQTDEPEDBLOQALT2      As Long
    Dim lngQTDEPEDLBOQFOT       As Long
    Dim lngQTDEPEDLBOQFOT2      As Long
    Dim lngQTDEPEDBLOQPCPD      As Long
    Dim lngQTDEPEDBLOQPCPD2     As Long
    Dim lngALOCADOPDIA          As Long
    Dim lngQTDCOTA              As Long
    Dim intRESP                 As Integer
    Dim arrGRPLIN()             As String
    Dim i                       As Integer
    
    Dim strCODLIN2              As String
    Dim strGRPCOD               As String
    Dim strCODLIN               As String
        
    lngSALDO = 0
    lngSALDOQTDENTR = 0
    lngQTDE = 0
    lngQTDEOP = 0
    lngQTDEPEDBLOQ = 0
    lngQTDEPEDBLOQ2 = 0
    lngQTDEPEDBLOQALT = 0
    lngQTDEPEDBLOQALT2 = 0
    lngQTDEPEDLBOQFOT = 0
    lngQTDEPEDLBOQFOT2 = 0
    lngQTDEPEDBLOQPCPD = 0
    lngQTDEPEDBLOQPCPD2 = 0
    
    strCODLIN2 = ""
    strGRPCOD = ""
    strCODLIN = ""
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "  From"
    sSql = sSql & "      SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_IDPRODUTO = " & strIDPRODUTO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then strCODLIN = Trim(BREC!SGI_CODLINPROD)
    BREC.Close
    
    
    '' =========================
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       GRPI.*" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO       LIMP" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHAIT" & strNOMFILIAL & "  GRPI" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       LIMP.SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "   And LIMP.SGI_CODLIN         = " & Trim(strCODLIN) & vbCrLf
    sSql = sSql & "   And GRPI.SGI_FILIAL         = LIMP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPI.SGI_CODLIN         = LIMP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And GRPI.SGI_OPTCOMNECKINSN = 0"
   
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC7.EOF()
        strGRPCOD = strGRPCOD & BREC7!SGI_CODIGO
        BREC7.MoveNext
        If Not BREC7.EOF() Then strGRPCOD = strGRPCOD & ","
    Loop
    BREC7.Close
    
    lngSALDOQTDENTR = PegaQtdeEntrega(Trim(Replace(strGRPCOD, ",", "")), Str(Month(CDate(Replace(strDTENTREGA, "'", "")))), Str(Year(CDate(Replace(strDTENTREGA, "'", "")))), Replace(strDTENTREGA, "'", ""))
    
    '' Pega Cota
    lngQTDCOTA = objCADPEDVENDA.PegaCota(Trim(grdProgEntrega.Cell(flexcpText, grdProgEntrega.Row, conCOL_SonProgEntr_IdProduto)), strDTENTREGA, strNOMFILIAL, 0)
    
    '' =========================
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "      LIMP.SGI_CODLIN" & vbCrLf
    
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaQueryOPDia("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, objCADPEDVENDA.CODPEDIDO, 0) & ") As QTDEOP" & vbCrLf
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaQueryPedBloq("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, objCADPEDVENDA.CODPEDIDO, 0) & ") As QTDEPEDBLOQ" & vbCrLf
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaPedQueryBloqAlt("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, objCADPEDVENDA.CODPEDIDO, 0) & ") As QTDEPEDBLOQALT" & vbCrLf
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaPedQueryBloqFot("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, objCADPEDVENDA.CODPEDIDO, 0) & ") As QTDEPEDLBOQFOT" & vbCrLf
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaPedQueryBloqPcPd("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, objCADPEDVENDA.CODPEDIDO, 0) & ") As QTDEPEDBLOQPCPD" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADGRUPLINHAIT" & strNOMFILIAL & " GRPI" & vbCrLf
    sSql = sSql & "     ,SGI_CADLINHAPRODUTO      LIMP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      GRPI.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And GRPI.SGI_CODIGO IN(" & Trim(strGRPCOD) & ")" & vbCrLf
    sSql = sSql & "  And LIMP.SGI_FILIAL = GRPI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And LIMP.SGI_CODIGO = GRPI.SGI_CODLIN"
    
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC7.EOF() Then
        Do While Not BREC7.EOF()
        
            lngQTDEOP = 0
            lngQTDEPEDBLOQ2 = 0
            lngQTDEPEDBLOQALT2 = 0
            lngQTDEPEDLBOQFOT2 = 0
            lngQTDEPEDBLOQPCPD2 = 0
            
            If Not IsNull(BREC7!QtdeOP) Then lngQTDEOP = BREC7!QtdeOP
            If Not IsNull(BREC7!QTDEPEDBLOQ) Then lngQTDEPEDBLOQ2 = BREC7!QTDEPEDBLOQ
            If Not IsNull(BREC7!QTDEPEDBLOQALT) Then lngQTDEPEDBLOQALT2 = BREC7!QTDEPEDBLOQALT
            If Not IsNull(BREC7!QTDEPEDLBOQFOT) Then lngQTDEPEDLBOQFOT2 = BREC7!QTDEPEDLBOQFOT
            If Not IsNull(BREC7!QTDEPEDBLOQPCPD) Then lngQTDEPEDBLOQPCPD2 = BREC7!QTDEPEDBLOQPCPD
            
            lngQTDE = (lngQTDE + lngQTDEOP)
            lngQTDEPEDBLOQ = (lngQTDEPEDBLOQ + lngQTDEPEDBLOQ2)
            lngQTDEPEDBLOQALT = (lngQTDEPEDBLOQALT + lngQTDEPEDBLOQALT2)
            lngQTDEPEDLBOQFOT = (lngQTDEPEDLBOQFOT + lngQTDEPEDLBOQFOT2)
            lngQTDEPEDBLOQPCPD = (lngQTDEPEDBLOQPCPD + lngQTDEPEDBLOQPCPD2)
            
            BREC7.MoveNext
        Loop
    
        lngALOCADOPDIA = (lngQTDE + lngSALDOQTDENTR + lngQTDEPEDBLOQ + lngQTDEPEDBLOQALT + lngQTDEPEDLBOQFOT + lngQTDEPEDBLOQPCPD)
    End If
    BREC7.Close

    lngSALDO = (lngQTDCOTA - lngALOCADOPDIA)
    If lngSALDO >= 0 Then ConfereCotas = False

    If ConfereCotas Then
        MsgBox "ATENÇÃO" & vbCrLf & "A Cota para para o dia " & Format(CDate(strDTENTREGA), "DD/MM/YYYY") & " já está estourada.", vbOKOnly + vbExclamation, "Aviso"
    End If
    
End Function

Private Sub MudaStatusOP_PDPC()
    Dim i As Integer
    With grdProgEntrega
        For i = 1 To (.Rows - 1)
            .Cell(flexcpText, i, conCOL_SonProgEntr_Action2Do) = dacEnumUpdateAction_update
            .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 0
        Next i
    End With
End Sub

Private Function PegaGrdCodCapac(strIDPRODUTO As String) As String

    PegaGrdCodCapac = ""
    
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       LINP.*" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO       PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO  LINP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = " & strIDPRODUTO & vbCrLf
    sSql = sSql & "   And LINP.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINP.SGI_CODLIN     = PROD.SGI_CODLINPROD" & vbCrLf
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then PegaGrdCodCapac = Trim(Str(BREC8!SGI_CODIGO))
    BREC8.Close
    
End Function


Private Function ConfereCliente(strCODIGO As String, strCODVEND As String) As Boolean

On Error GoTo Err_ConfereCliente

    If BREC10.State = 1 Then BREC10.Close
    
    ConfereCliente = False
    
    Dim boolDadosInv As Boolean
    
    If Len(Trim(strCODVEND)) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Informe o Vendedor !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       CLIE.*" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADCLIEVEND CVEN" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       CVEN.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And CVEN.SGI_CODIGO   = " & txtCODVEND.Text & vbCrLf
    sSql = sSql & "   And CVEN.SGI_CODCLI   = " & strCODIGO & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL   = CVEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO   = CVEN.SGI_CODCLI"
    
    boolDadosInv = True
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC10.EOF() Then
       MsgBox "Este Cliente não pertence a este vendedor !!!", vbOKOnly + vbExclamation, "Aviso"
       boolDadosInv = False
    Else
        If BREC10!SGI_DESBCLIE = 0 Then
            MsgBox "Este Cliente está desabilitado !!!", vbOKOnly + vbExclamation, "Aviso"
            boolDadosInv = False
        End If
    End If
    BREC10.Close
    
    If boolDadosInv = False Then Exit Function
    
    ConfereCliente = True
    
    Exit Function
    
Err_ConfereCliente:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : ConfereCliente()", Me.Name, "ConfereCliente()", strCAMARQERRO)

End Function

Private Sub PreenchCboFechTPFR(strIDPROD As String)
    
    cboFechTPFR.Clear
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(strIDPROD)) = 0 Then Exit Sub
    
    Dim intFECH_PADRAO As Integer
    Dim i              As Integer
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       LFTF.SGI_COD" & vbCrLf
    sSql = sSql & "      ,FECH.SGI_DESCRI" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_FechTampaFuro" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO                PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO           LIMP" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO_FECHTPFR  LFTF" & vbCrLf
    sSql = sSql & "      ,SGI_CADFECHAM                 FECH" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = " & Trim(strIDPROD) & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL     = LIMP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODLINPROD = LIMP.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "   And LIMP.SGI_FILIAL     = LFTF.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LIMP.SGI_CODIGO     = LFTF.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And LFTF.SGI_FILIAL     = FECH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LFTF.SGI_COD        = FECH.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "Order By LFTF.SGI_COD"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
        
        intFECH_PADRAO = BREC10!SGI_FechTampaFuro
        Do While Not BREC10.EOF()
           cboFechTPFR.AddItem Trim(BREC10!SGI_DESCRI)
           cboFechTPFR.ItemData(cboFechTPFR.NewIndex) = BREC10!SGI_COD
           BREC10.MoveNext
        Loop
        
        With cboFechTPFR
            For i = 0 To (.ListCount - 1)
                If .ItemData(i) = intFECH_PADRAO Then
                    .ListIndex = i
                    
                    grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_FechTpFr) = .Text
                    grdProduto.Cell(flexcpData, grdProduto.Row, conCOL_SonProd_FechTpFr) = .ItemData(.ListIndex)
                    
                    Exit For
                End If
            Next i
        End With
        
    End If
    BREC10.Close
    
End Sub


Private Function PegaDescTabelasGrd(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String) As String

On Error GoTo Err_PegaDescTabelasGrd

    PegaDescTabelasGrd = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCODIGO)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then PegaDescTabelasGrd = Trim(BREC10(Trim(StrCampoRetorno)))
    BREC10.Close
    
    Exit Function
    
Err_PegaDescTabelasGrd:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDescTabelasGrd()", Me.Name, "PegaDescTabelasGrd()", strCAMARQERRO)

End Function

Public Function PegaQtdeLT_4_Palhets(strIDPROD As String) As String

    PegaQtdeLT_4_Palhets = ""

    If Len(Trim(strIDPROD)) = 0 Then Exit Function
    
    If BREC11.State = 1 Then BREC11.Close
    
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_QTDEEMPILHAMENTO" & vbCrLf
    sSql = sSql & "      ,SGI_USARPADRPALSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPROD & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
        
        grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_PalhetPadrao) = BREC!SGI_USARPADRPALSN
        
        If BREC!SGI_USARPADRPALSN = 1 Then PegaQtdeLT_4_Palhets = BREC!SGI_QTDEEMPILHAMENTO
    
    End If
    BREC.Close
    
End Function

Private Function ConferePalhets(lngROW As Long, lngQtd_do_Item As Long) As Boolean

        ConferePalhets = False

        Dim lngTot_Latas_4_Palhet   As Long
        Dim lngRestoPAlhets         As Long
        Dim curTot_Latas_Palhet     As Currency
        
        lngTot_Latas_4_Palhet = 0
        lngRestoPAlhets = 0
        curTot_Latas_Palhet = 0
        
        If Len(Trim(grdProduto.Cell(flexcpText, lngROW, conCOL_SonProd_QTDELATASPALLETS))) > 0 Then lngTot_Latas_4_Palhet = CLng(grdProduto.Cell(flexcpText, lngROW, conCOL_SonProd_QTDELATASPALLETS))

        If lngTot_Latas_4_Palhet > 0 Then
           curTot_Latas_Palhet = Round((lngQtd_do_Item / lngTot_Latas_4_Palhet))
           lngRestoPAlhets = (lngQtd_do_Item Mod lngTot_Latas_4_Palhet)
        End If
           
        If lngRestoPAlhets = 0 Then
           If curTot_Latas_Palhet > 0 Then grdProduto.Cell(flexcpText, lngROW, conCOL_SonProd_PALLETS) = curTot_Latas_Palhet
        Else
           MsgBox "ATENÇÃO" & vbCrLf & _
                  "Esta Quantidade de Latas, irá gerar : " & Fix(curTot_Latas_Palhet) & " Palhet(s) Inteiros , " & vbCrLf & _
                  "e ira gerar um resto de " & lngRestoPAlhets & " lata(s), quantidade sugerida : " & (lngQtd_do_Item - lngRestoPAlhets) & ".", vbOKOnly + vbExclamation, "Aviso"
           
           grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_QtdProd) = Empty
           grdProduto.Cell(flexcpText, grdProduto.Row, conCOL_SonProd_PALLETS) = Empty
           Exit Function
        End If


        ConferePalhets = True
End Function


Private Function ConferePalhetsProgEntrg(lngROW As Long, lngQtd_do_Item As Long) As Boolean

        ConferePalhetsProgEntrg = False

        Dim lngTot_Latas_4_Palhet   As Long
        Dim lngRestoPAlhets         As Long
        Dim curTot_Latas_Palhet     As Currency
        
        lngTot_Latas_4_Palhet = 0
        lngRestoPAlhets = 0
        curTot_Latas_Palhet = 0
        
        If Len(Trim(grdProgEntrega.Cell(flexcpText, lngROW, conCOL_SonProgEntr_QTDENOPALHET))) > 0 Then lngTot_Latas_4_Palhet = CLng(grdProgEntrega.Cell(flexcpText, lngROW, conCOL_SonProgEntr_QTDENOPALHET))

        If lngTot_Latas_4_Palhet > 0 Then
           curTot_Latas_Palhet = Round((lngQtd_do_Item / lngTot_Latas_4_Palhet))
           lngRestoPAlhets = (lngQtd_do_Item Mod lngTot_Latas_4_Palhet)
        End If
           
        If lngRestoPAlhets = 0 Then
           If curTot_Latas_Palhet > 0 Then grdProgEntrega.Cell(flexcpText, lngROW, conCOL_SonProgEntr_PALHET) = curTot_Latas_Palhet
        Else
           MsgBox "ATENÇÃO" & vbCrLf & _
                  "Esta Quantidade de Latas, irá gerar : " & Fix(curTot_Latas_Palhet) & " Palhet(s) Inteiros , " & vbCrLf & _
                  "e ira gerar um resto de " & lngRestoPAlhets & " lata(s), quantidade sugerida : " & (lngQtd_do_Item - lngRestoPAlhets) & ".", vbOKOnly + vbExclamation, "Aviso"
           
           Exit Function
        End If


        ConferePalhetsProgEntrg = True
End Function

Private Sub AbilDesConferido(boolAbilDes As Boolean, intAbilDes As Integer)
    chkVerificado.Value = intAbilDes
    chkVerificado.Enabled = boolAbilDes
End Sub

Private Sub AbilConferido()

    chkVerificado.Enabled = True
    chkVerificado.Value = objCADPEDVENDA.CONFERIDO
    
    Frame4.Enabled = True
    
    txtCODVEND.Enabled = False
    txtTIPPED.Enabled = False
    txtCIDCLIE.Enabled = False
    txtCodCondPgto.Enabled = False
    
    Command2.Enabled = False
    Command3.Enabled = False
    Command1.Enabled = False
    cmdCondPgto.Enabled = False

End Sub


Private Sub PreenchCboPallet(strIDPROD As String)
    
    cboQtdePorPalhet.Clear
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(strIDPROD)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       LFTF.SGI_QTDELATAS" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO                   PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO              LIMP" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO_FECHTPFRARM  LFTF" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = " & Trim(strIDPROD) & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL     = LIMP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODLINPROD = LIMP.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "   And LIMP.SGI_FILIAL     = LFTF.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LIMP.SGI_CODIGO     = LFTF.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "Order By LFTF.SGI_QTDELATAS"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC10.EOF()
       cboQtdePorPalhet.AddItem BREC10!SGI_QTDELATAS
       ''cboQtdePorPalhet.ItemData(cboFechTPFR.NewIndex) = BREC10!SGI_COD
       BREC10.MoveNext
    Loop
    BREC10.Close
    
End Sub


Private Sub GeraLog()
    Call objBLBFunc.LOG_ACAO(FILIAL, Linha, Me.Name, objBLBFunc.Crypt(strUSUARIO), strNOMCOMP, strVERSAO, cTipOper, objCADPEDVENDA.CODPEDIDO, "Null", "Null", "Null")
End Sub

Private Sub LimpaCamposProEntr(strIDPRODUTO As String)

    Dim i As Integer
    Dim lngLINHA As Long
    
    With grdProgEntrega
        lngLINHA = grdProduto.FindRow(strIDPRODUTO, , conCOL_SonProd_IdProduto)
        If lngLINHA > -1 Then
            If grdProduto.Cell(flexcpText, lngLINHA, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_delete And _
               grdProduto.Cell(flexcpText, lngLINHA, conCOL_SonProd_Action2Do) <> dacEnumUpdateAction_Ignore And _
               grdProduto.Cell(flexcpText, lngLINHA, conCOL_SonProd_StatusProd) = 2 Then
                For i = 1 To (.Rows - 1)
                    If .Cell(flexcpText, i, conCOL_SonProgEntr_IdProduto) = strIDPRODUTO Then
                        .Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega) = Empty
                        .Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP) = 4
                        .Cell(flexcpText, i, conCOL_SonProgEntr_DescStatusOP) = PegaStatusOP(.Cell(flexcpText, i, conCOL_SonProgEntr_StatusOP))
                    End If
                Next i
            End If
        End If
    
    End With

End Sub

Private Function PegaPedidosAberto(strCODCLIE As String, strIDPRODUTO As String) As Boolean

    PegaPedidosAberto = True
    
    If Len(Trim(strCODCLIE)) = 0 Then Exit Function
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
    
    Dim intQTDEREGS       As Integer
    Dim arrPEDIDOS_PEND() As String
    
    intQTDEREGS = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       OP.SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,SUM(OP.SGI_QTDEPED) As QtdePed" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & " OP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDVH" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       OP.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And OP.SGI_IDPRODUTO = " & strIDPRODUTO & vbCrLf
    sSql = sSql & "   And (OP.SGI_STATUS = 0 Or OP.SGI_STATUS = 1)" & vbCrLf
    sSql = sSql & "   And PEDVH.SGI_FILIAL = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PEDVH.SGI_CODIGO = OP.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And PEDVH.SGI_CODCLI = " & strCODCLIE
    sSql = sSql & "Group By OP.SGI_CODPED"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        PegaPedidosAberto = True
    
        Do While Not BREC.EOF()
            intQTDEREGS = (intQTDEREGS + 1)
            BREC.MoveNext
        Loop
    
        If intQTDEREGS > 0 Then
    
            ReDim arrPEDIDOS_PEND(1 To intQTDEREGS, 1 To 2) As String
            BREC.MoveFirst
            Do While Not BREC.EOF()
    
                BREC.MoveNext
            Loop
        End If
    
    End If
    BREC.Close


End Function


Private Function CotaEstourada(strCODLIN As String, strDTENTREGA As String, strIDPRODUTO As String, intAction2Do As Integer, strIDINTERNO As String, lngSALDOQTDENTR As Long, lngCODPED As Long) As Long
    
On Error GoTo Err_CotaEstourada
    
    Dim boolCotaEstourada       As Boolean
    boolCotaEstourada = True
    
    If Len(Trim(strCODLIN)) = 0 Then Exit Function
    If Len(Trim(strDTENTREGA)) = 0 Then Exit Function
    
    Dim lngSALDO                As Long
    Dim lngQTDE                 As Long
    Dim lngQTDEOP               As Long
    Dim lngQTDEPEDBLOQ          As Long
    Dim lngQTDEPEDBLOQ2         As Long
    Dim lngQTDEPEDBLOQALT       As Long
    Dim lngQTDEPEDBLOQALT2      As Long
    Dim lngQTDEPEDLBOQFOT       As Long
    Dim lngQTDEPEDLBOQFOT2      As Long
    Dim lngQTDEPEDBLOQPCPD      As Long
    Dim lngQTDEPEDBLOQPCPD2     As Long
    Dim lngSALDOQTDENTR3        As Long
    Dim lngSALDOQTDENTR4        As Long
    Dim lngALOCADOPDIA          As Long
    Dim lngQTDCOTA              As Long
    Dim lngTOTALSALDO           As Long
    Dim intRESP                 As Integer
    Dim arrGRPLIN()             As String
    Dim i                       As Integer
    
    Dim strCODLIN2              As String
    Dim strGRPCOD               As String
    Dim lngNECKIN               As Long
    Dim arrDADOS()              As String
    Dim intHOMOLOGADO           As Integer
    Dim lngTOTALATRAZADO        As Long
        
    lngSALDO = 0
    lngQTDE = 0
    lngQTDEOP = 0
    lngQTDEPEDBLOQ = 0
    lngQTDEPEDBLOQ2 = 0
    lngQTDEPEDBLOQALT = 0
    lngQTDEPEDBLOQALT2 = 0
    lngQTDEPEDLBOQFOT = 0
    lngQTDEPEDLBOQFOT2 = 0
    lngQTDEPEDBLOQPCPD = 0
    lngQTDEPEDBLOQPCPD2 = 0
    lngTOTALSALDO = 0
    
    strCODLIN2 = ""
    strGRPCOD = ""
    
    lngNECKIN = objCADPEDVENDA.PegaNECKIN(strIDPRODUTO)
    intHOMOLOGADO = objCADPEDVENDA.PegaHOMOLOGADO(strIDPRODUTO)
    lngTOTALATRAZADO = objCADPEDVENDA.PegaAtrazados(strCODLIN, strNOMFILIAL)
    
    '' =========================
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       GRPI.*" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO       LIMP" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHAIT" & strNOMFILIAL & "  GRPI" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       LIMP.SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "   And LIMP.SGI_CODLIN         = " & Trim(strCODLIN) & vbCrLf
    sSql = sSql & "   And GRPI.SGI_FILIAL         = LIMP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPI.SGI_CODLIN         = LIMP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And GRPI.SGI_OPTCOMNECKINSN = " & lngNECKIN & vbCrLf
    sSql = sSql & "   And GRPI.SGI_HOMOLOGSN      = " & intHOMOLOGADO & vbCrLf
   
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC7.EOF()
        strGRPCOD = strGRPCOD & BREC7!SGI_CODIGO
        BREC7.MoveNext
        If Not BREC7.EOF() Then strGRPCOD = strGRPCOD & ","
    Loop
    BREC7.Close
    
    lngSALDOQTDENTR3 = PegaQtdeEntregaPcota(Trim(Replace(strGRPCOD, ",", "")), Month(CDate(strDTENTREGA)), Year(CDate(strDTENTREGA)), strDTENTREGA, strIDINTERNO)
    lngSALDOQTDENTR3 = lngSALDOQTDENTR3 + lngSALDOQTDENTR
    '' Pega Cota
    lngQTDCOTA = objCADPEDVENDA.PegaCota(Trim(strIDPRODUTO), strDTENTREGA, strNOMFILIAL, intHOMOLOGADO)
    '' =========================
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "      LIMP.SGI_CODLIN" & vbCrLf
    
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaQueryOPDia("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEOP" & vbCrLf
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaQueryPedBloq("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQ" & vbCrLf
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaPedQueryBloqAlt("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQALT" & vbCrLf
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaPedQueryBloqFot("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDLBOQFOT" & vbCrLf
    ''sSql = sSql & "     ,(" & objCADPEDVENDA.PegaPedQueryBloqPcPd("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQPCPD" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADGRUPLINHAIT" & strNOMFILIAL & " GRPI" & vbCrLf
    sSql = sSql & "     ,SGI_CADLINHAPRODUTO      LIMP" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      GRPI.SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "  And GRPI.SGI_CODIGO IN(" & Trim(strGRPCOD) & ")" & vbCrLf
    sSql = sSql & "  And GRPI.SGI_OPTCOMNECKINSN = " & lngNECKIN & vbCrLf
    sSql = sSql & "  And GRPI.SGI_HOMOLOGSN      = " & intHOMOLOGADO & vbCrLf
    sSql = sSql & "  And LIMP.SGI_FILIAL         = GRPI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And LIMP.SGI_CODIGO         = GRPI.SGI_CODLIN"
    
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC7.EOF() Then
        Do While Not BREC7.EOF()
        
            
            lngQTDEOP = 0
            lngQTDEPEDBLOQ2 = 0
            lngQTDEPEDBLOQALT2 = 0
            lngQTDEPEDLBOQFOT2 = 0
            lngQTDEPEDBLOQPCPD2 = 0
            
            lngQTDEOP = PegaQueryOPDia(BREC7!SGI_CODLIN, strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO)
            lngQTDEPEDBLOQ2 = PegaQueryPedBloq(BREC7!SGI_CODLIN, strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO)
            lngQTDEPEDBLOQALT2 = PegaPedQueryBloqAlt(BREC7!SGI_CODLIN, strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO)
            lngQTDEPEDLBOQFOT2 = PegaPedQueryBloqFot(BREC7!SGI_CODLIN, strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO)
            lngQTDEPEDBLOQPCPD2 = PegaPedQueryBloqPcPd(BREC7!SGI_CODLIN, strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO)
            
            ''If Not IsNull(BREC7!QtdeOP) Then lngQTDEOP = BREC7!QtdeOP
            ''If Not IsNull(BREC7!QTDEPEDBLOQ) Then lngQTDEPEDBLOQ2 = BREC7!QTDEPEDBLOQ
            ''If Not IsNull(BREC7!QTDEPEDBLOQALT) Then lngQTDEPEDBLOQALT2 = BREC7!QTDEPEDBLOQALT
            ''If Not IsNull(BREC7!QTDEPEDLBOQFOT) Then lngQTDEPEDLBOQFOT2 = BREC7!QTDEPEDLBOQFOT
            ''If Not IsNull(BREC7!QTDEPEDBLOQPCPD) Then lngQTDEPEDBLOQPCPD2 = BREC7!QTDEPEDBLOQPCPD
            
            lngQTDE = (lngQTDE + lngQTDEOP)
            lngQTDEPEDBLOQ = (lngQTDEPEDBLOQ + lngQTDEPEDBLOQ2)
            lngQTDEPEDBLOQALT = (lngQTDEPEDBLOQALT + lngQTDEPEDBLOQALT2)
            lngQTDEPEDLBOQFOT = (lngQTDEPEDLBOQFOT + lngQTDEPEDLBOQFOT2)
            lngQTDEPEDBLOQPCPD = (lngQTDEPEDBLOQPCPD + lngQTDEPEDBLOQPCPD2)
            
            BREC7.MoveNext
        Loop
        
        lngALOCADOPDIA = (lngQTDE + lngSALDOQTDENTR3 + lngQTDEPEDBLOQ + lngQTDEPEDBLOQALT + lngQTDEPEDLBOQFOT + lngQTDEPEDBLOQPCPD + lngTOTALATRAZADO)
    
    End If
    BREC7.Close

    lngSALDO = (lngQTDCOTA - lngALOCADOPDIA)
    
    If lngSALDO >= 0 Then boolCotaEstourada = False
    
    If boolCotaEstourada Then
        If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
        CotaEstourada = 6
    Else
        If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
        CotaEstourada = 0
    End If
    
    
    Exit Function
    
Err_CotaEstourada:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CotaEstourada()", Me.Name, "CotaEstourada()", strCAMARQERRO, sSql)
    
End Function



Private Function PegaQtdeEntregaPcota(strIDPRODUTO As String, strMES As String, strANO As String, strDia As String, strIDINTERNO As String) As Long
    
On Error GoTo Err_PegaQtdeEntregaPcota
    
    PegaQtdeEntregaPcota = 0
    
    Dim i       As Integer
    
    For i = 1 To UBound(arrDIASCOTAS)
        If Len(Trim(arrDIASCOTAS(i, 1))) > 0 And _
           Len(Trim(arrDIASCOTAS(i, 2))) > 0 Then
           
            If Trim(strIDPRODUTO) = Trim(arrDIASCOTAS(i, 4)) And _
               Trim(strMES) = Trim(Str(Month(CDate(arrDIASCOTAS(i, 1))))) And _
               Trim(strANO) = Trim(Str(Year(CDate(arrDIASCOTAS(i, 1))))) And _
               CDate(arrDIASCOTAS(i, 1)) = CDate(strDia) Then
               PegaQtdeEntregaPcota = PegaQtdeEntregaPcota + CLng(arrDIASCOTAS(i, 2))
            End If
        End If
    Next i
    
    Exit Function
    
Err_PegaQtdeEntregaPcota:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaQtdeEntregaPcota()", Me.Name, "PegaQtdeEntregaPcota()", strCAMARQERRO)
    
End Function


Private Function PegaQueryOPDia(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOMOLOGSN As Integer) As Long

        PegaQueryOPDia = 0
        
        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function
        
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(OP.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO      PROD" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & " OP" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD      TP" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PROD.SGI_CODLINPROD = " & strCODLIN & vbCrLf
        sSql = sSql & "   And PROD.SGI_NECKIN     = " & lngNECKIN & vbCrLf
        sSql = sSql & "   And OP.SGI_FILIAL       = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And OP.SGI_IDPRODUTO    = PROD.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And OP.SGI_DATENTREGA   = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And OP.SGI_STATUS       = 0" & vbCrLf
        sSql = sSql & "   And OP.SGI_CODPED       <> " & lngCODPED & vbCrLf
        sSql = sSql & "   And TP.SGI_FILIAL       = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO       = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN    = " & intHOMOLOGSN
        
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
           If Not IsNull(BREC10!SGI_QTDE) Then PegaQueryOPDia = BREC10!SGI_QTDE
        End If
        BREC10.Close

End Function


Private Function PegaQueryPedBloq(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOMOLOGSN As Integer) As Long

        PegaQueryPedBloq = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TP" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & lngCODPED & vbCrLf
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And (PEDVH.SGI_STATUS    = 'B' Or PEDVH.SGI_STATUS = 'N')" & vbCrLf
        
        sSql = sSql & "   And PROD.SGI_FILIAL      = PROGE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO   = PROGE.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_NECKIN      = " & lngNECKIN & vbCrLf
        
        sSql = sSql & "   And TP.SGI_FILIAL        = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO        = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN     = " & intHOMOLOGSN
        
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
           If Not IsNull(BREC10!SGI_QTDE) Then PegaQueryPedBloq = BREC10!SGI_QTDE
        End If
        BREC10.Close
        
End Function


Private Function PegaPedQueryBloqAlt(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOLOGSN As Integer) As Long

        PegaPedQueryBloqAlt = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TP" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & lngCODPED & vbCrLf
        
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_STATUS     = 'S'" & vbCrLf
        
        sSql = sSql & "   And PROD.SGI_FILIAL      = PROGE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO   = PROGE.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_NECKIN      = " & lngNECKIN & vbCrLf
        
        sSql = sSql & "   And TP.SGI_FILIAL        = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO        = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN     = " & intHOLOGSN
        
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
           If Not IsNull(BREC10!SGI_QTDE) Then PegaPedQueryBloqAlt = BREC10!SGI_QTDE
        End If
        BREC10.Close

End Function


Private Function PegaPedQueryBloqFot(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOMOLOGSN As Integer) As Long

        PegaPedQueryBloqFot = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & "  PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TP" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & lngCODPED & vbCrLf
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_STATUS     = 'V'" & vbCrLf
        
        sSql = sSql & "   And PROD.SGI_FILIAL      = PROGE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO   = PROGE.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_NECKIN      = " & lngNECKIN & vbCrLf
        
        sSql = sSql & "   And TP.SGI_FILIAL        = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO        = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN     = " & intHOMOLOGSN
        
        
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
            If Not IsNull(BREC10!SGI_QTDE) Then PegaPedQueryBloqFot = BREC10!SGI_QTDE
        End If
        BREC10.Close
        
End Function


Public Function PegaPedQueryBloqPcPd(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOMOLOGSN As Integer) As Long

        PegaPedQueryBloqPcPd = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & "  PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TP" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & strCODLIN & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & lngCODPED & vbCrLf
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_STATUS     = 'C'" & vbCrLf
        
        sSql = sSql & "   And PROD.SGI_FILIAL      = PROGE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO   = PROGE.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_NECKIN      = " & lngNECKIN & vbCrLf
        
        sSql = sSql & "   And TP.SGI_FILIAL        = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO        = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN     = " & intHOMOLOGSN
       
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
            If Not IsNull(BREC10!SGI_QTDE) Then PegaPedQueryBloqPcPd = BREC10!SGI_QTDE
        End If
        BREC10.Close

End Function

Private Sub InitGridLogPed()

    With grdLogPed
    
       .Cols = conColumnsIn_SonLogPed
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonLogPed_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Data) = ""
       .ColDataType(conCOL_SonLogPed_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Hora) = ""
       .ColDataType(conCOL_SonLogPed_Hora) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_CodUsuario) = ""
       .ColDataType(conCOL_SonLogPed_CodUsuario) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Usuario) = ""
       .ColDataType(conCOL_SonLogPed_Usuario) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_CodAcao) = ""
       .ColDataType(conCOL_SonLogPed_CodAcao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Acao) = ""
       .ColDataType(conCOL_SonLogPed_Acao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Tipo) = ""
       .ColDataType(conCOL_SonLogPed_Tipo) = flexDTString
       
       .ColWidth(conCOL_SonLogPed_Data) = 1000
       .ColWidth(conCOL_SonLogPed_Hora) = 1000
       .ColWidth(conCOL_SonLogPed_CodUsuario) = 0
       .ColWidth(conCOL_SonLogPed_Usuario) = 2000
       .ColWidth(conCOL_SonLogPed_CodAcao) = 0
       .ColWidth(conCOL_SonLogPed_Acao) = 7000
       .ColWidth(conCOL_SonLogPed_Tipo) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub PopLogPedidos()

    Dim i           As Integer
    Dim strDescAcao As String
    
    If IsArray(objCADPEDVENDA.LOGPEDIDOS) Then
    
        With grdLogPed
        
            For i = 1 To UBound(objCADPEDVENDA.LOGPEDIDOS)
                
                .AddItem objCADPEDVENDA.LOGPEDIDOS(i, 1) & vbTab & _
                         objCADPEDVENDA.LOGPEDIDOS(i, 2) & vbTab & _
                         objCADPEDVENDA.LOGPEDIDOS(i, 3) & vbTab & _
                         PegaUsuario(CLng(objCADPEDVENDA.LOGPEDIDOS(i, 3))) & vbTab & _
                         objCADPEDVENDA.LOGPEDIDOS(i, 4) & vbTab & _
                         "" & vbTab & _
                         ""
                         
                strDescAcao = objCADPEDVENDA.LOGPEDIDOS(i, 5)
                         
                .Cell(flexcpText, (.Rows - 1), conCOL_SonLogPed_Acao) = DescAcao(.Cell(flexcpText, (.Rows - 1), conCOL_SonLogPed_CodAcao), strDescAcao)
                         
            Next i
        
        End With
    
    End If

End Sub


Private Function DescAcao(strACAO As String, strDescAcao As String) As String

    DescAcao = ""
    
    If strACAO = "I" Then DescAcao = "Inclusão"
    If (strACAO = "LN" Or strACAO = "N") Then DescAcao = "Liberado Comercial"
    If (strACAO = "LF" Or strACAO = "L") Then DescAcao = "Liberado Financeiro"
    If strACAO = "LC" Then DescAcao = "Liberado P.Cota/P.Data"
    If strACAO = "D" Then DescAcao = "Pedido Bloqueado"
    If strACAO = "A" Then DescAcao = "Pedido Alterado"
    If strACAO = "R" Then DescAcao = "Pedido Reprovado"
    If strACAO = "LV" Then DescAcao = "Arte Liberada"
    If strACAO = "M" Then DescAcao = "Pedido Liguidado Manualmente"
    If strACAO = "LS" Or strACAO = "S" Then DescAcao = "Pedido Liberado do Bloqueio"
    
    If strACAO = "AD" Then DescAcao = "Campo data de Entrega foi Alterado de " & Trim(strDescAcao)
    If strACAO = "POP" Then DescAcao = "Campo data de Entrega foi Alterado de " & Trim(strDescAcao) & " na Progrmação de OP's"
    If strACAO = "PPR" Then DescAcao = "OP " & Trim(strDescAcao) & " Programada"
    If strACAO = "PPE" Then DescAcao = "OP " & Trim(strDescAcao) & " Removida da Programação"
    If strACAO = "OPB" Then DescAcao = "OP " & Trim(strDescAcao) & " Apontada/Montada"
    If strACAO = "OAE" Then DescAcao = "OP " & Trim(strDescAcao) & " Excluida do Apontamento"
    
    If strACAO = "OF" Then DescAcao = "Ordem de Faturamento Gerada / OP nº : " & Trim(strDescAcao)
    If strACAO = "OFA" Then DescAcao = "Ordem de Faturamento Alterada / OP nº : " & Trim(strDescAcao)
    If strACAO = "CF" Then DescAcao = "OP Faturada / OP nº : " & Trim(strDescAcao)
        
End Function

Private Function PegaUsuario(lngCodUsu As Long) As String

    PegaUsuario = ""
    
    If lngCodUsu = 0 Then
        PegaUsuario = "CWS"
        Exit Function
    End If

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_NOME " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_USUARIO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsu
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaUsuario = objBLBFunc.Crypt(BREC!SGI_NOME)
    BREC.Close

End Function

Private Sub LiquidaPedido()
    Dim i As Integer
    
    For i = 0 To (stCAMPOSVENDA.Tabs - 1)
        stCAMPOSVENDA.TabEnabled(i) = False
    Next i
    stCAMPOSVENDA.TabEnabled(2) = True
    stCAMPOSVENDA.Tab = 2
End Sub


Private Sub Liquida()

On Error GoTo Err_Liquida
    
    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Pedido de Venda - [ LIQUIDA ]"
    
    objBLBFunc.LimpaCampos Me

    Call LiquidaPedido
    
    objBLBFunc.Preenche_Estado cboESTENTR
    objBLBFunc.Preenche_Estado cboESTCOBR
    
    objCADPEDVENDA.CODPEDIDO = iCodigo
    
    Call InitGridReprovacao
    Call InitGridProd
    Call InitGridProg
    Call InitGridOrdemFat
    Call InitGridLogPed
    Call InitGridProducao
    
    
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    txtOutrDesp.Text = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    lblDescMotLiq.Caption = ""
    
    '' --------------------
    '' Desconto
    lblVLDESCONTO.Caption = ""
    '' --------------------
   
    Call LimpaCamposLabel
    Call LimpaCamposDadosAdicionais
    Call LimpaCampoSaldoRot
    Call LimpaSaldoPedido
    
    objCADPEDVENDA.FILIALPED = intFILIALPED
    
    optESPECIAL(0).Value = True
    
    Call AbilDesConferido(False, 0)
    
    Frame28.Enabled = True
    txtOBS_MotLiq.Locked = False
    
    If objCADPEDVENDA.Carrega_Campos = True Then
    
       lblSTATUS.Caption = "LIQUIDADO MANUALMENTE"
       objCADPEDVENDA.STATUS = "M"
       cmdAltera.Enabled = False
       
       lblCODIGO.Caption = objCADPEDVENDA.CODPEDIDO
       mskDATAPED.Text = Format(objCADPEDVENDA.DATAPED, "DD/MM/YYYY")
       
       txtCIDCLIE.Text = objCADPEDVENDA.CODCLIE
       txtCodCondPgto.Text = objCADPEDVENDA.CODCONDPGTO
       txtCODVEND.Text = objCADPEDVENDA.CODVEND
       txtTIPPED.Text = objCADPEDVENDA.TIPPED
       txtOBSERVACAO.Text = objCADPEDVENDA.OBSERVACAO
       
       '' Dados de Entrega
       txtENDENTR.Text = objCADPEDVENDA.ENDENTR
       txtBAIENTR.Text = objCADPEDVENDA.BAIENTR
       txtCIDENTR.Text = objCADPEDVENDA.CIDENTR
       If objCADPEDVENDA.ESTENTREGA > 0 Then cboESTENTR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTENTREGA)
       txtCEPENTR.Text = objCADPEDVENDA.CEPENTREGA
       txtTELENTR.Text = objCADPEDVENDA.TELENTR
       txtFAXENTRE.Text = objCADPEDVENDA.FAXENTR
       
       '' Dados de Cobrança
       txtENDCOBR.Text = objCADPEDVENDA.ENDCOBRA
       txtBAICOBR.Text = objCADPEDVENDA.BAICOBRA
       txtCIDCOBR.Text = objCADPEDVENDA.CIDCOBRA
       If objCADPEDVENDA.ESTCOBRA > 0 Then cboESTCOBR.Text = objBLBFunc.PegaEstados(objCADPEDVENDA.ESTCOBRA)
       txtCEPCOBR.Text = objCADPEDVENDA.CEPCOBRA
       txtTELCOBR.Text = objCADPEDVENDA.TELCOBRA
       txtFAXCOBR.Text = objCADPEDVENDA.FAXCOBRA
       
       txtOBSERVACAO.Text = objCADPEDVENDA.OBSERVACAO
       txtOBS2.Text = objCADPEDVENDA.OBS2
       
       txtCODTRANSP.Text = objCADPEDVENDA.CODTRANSP
       txtORDCOMPCLI.Text = objCADPEDVENDA.ORDCOMPCLI
       txtCONTATO.Text = objCADPEDVENDA.CONTATO
       txtDEPARTAMENTO.Text = objCADPEDVENDA.DEPARTAMENTO
       
       optESPECIAL(objCADPEDVENDA.ESPECIAL).Value = True
       optPARAESTOQUE(objCADPEDVENDA.PARAESTOQUE).Value = True
       
       '' Totais
       If objCADPEDVENDA.VALBASICMS > 0 Then lblBASICMS.Caption = Format(objCADPEDVENDA.VALBASICMS, "#,##0.00")
       If objCADPEDVENDA.ALIQICMS > 0 Then txtALIQICMS.Text = Format(objCADPEDVENDA.ALIQICMS, "#,##0.00")
       If objCADPEDVENDA.VLICMS > 0 Then lblVLICMS.Caption = Format(objCADPEDVENDA.VLICMS, "#,##0.00")
       If objCADPEDVENDA.OUTRDESPESAS > 0 Then txtOutrDesp.Text = Format(objCADPEDVENDA.OUTRDESPESAS, "#,##0.00")
       If objCADPEDVENDA.VLFRETE > 0 Then txtFRETE.Text = Format(objCADPEDVENDA.VLFRETE, "#,##0.00")
       If objCADPEDVENDA.VLIPI > 0 Then lblVLIPI.Caption = Format(objCADPEDVENDA.VLIPI, "#,##0.00")
       If objCADPEDVENDA.VLDESCTO > 0 Then lblVLDESCONTO.Caption = Format(objCADPEDVENDA.VLDESCTO, "#,##0.00")
       If objCADPEDVENDA.TOTORCTO > 0 Then lblVLTOTAL.Caption = Format(objCADPEDVENDA.TOTORCTO, "#,##0.00")
       
        chkVerificado.Value = objCADPEDVENDA.CONFERIDO
        objCADPEDVENDA.PERMITEFECHOP = objCADPEDVENDA.PERMITEFECHOP
       
        Call PopGrdProdutos
        Call CarregaPlanoEntrega
        Call GeraSaldoPedido(Str(objCADPEDVENDA.CODPEDIDO))
        Call PegaDadosLabel
        Call PopLogPedidos
    
        Call VisualizaBotoesPCD(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibAlteracao(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibFotolito(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibComercial(objCADPEDVENDA.STATUS, cTipOper)
        Call VisualizaBotoesLibFinanceira(objCADPEDVENDA.STATUS, cTipOper)
    
    End If
        
    Exit Sub
    
Err_Liquida:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Liquida()", Me.Name, "Liquida()", strCAMARQERRO)
    
        
End Sub



Private Sub PegaOPS(strCODPED As String)

On Error GoTo Err_PegaOPS
    
    If Len(Trim(strCODPED)) = 0 Then Exit Sub
    
    If BREC10.State = 1 Then BREC.Close
    
    objCADPEDVENDA.OPS = Empty
    
    Dim lngQTDREGS  As Long
    Dim strFILIAL   As String
    Dim arrOPS()    As String

    
    strFILIAL = ""
    If intFILIALPED = 1 Then strFILIAL = "_STEEL"
    
    sSql = ""
    
    sSql = "Select * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_ORDEMPROD" & strFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And  SGI_CODPED = " & strCODPED & vbCrLf
    sSql = sSql & "   And  SGI_STATUS In(0,1,6,7)"
    
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
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaOPS()", Me.Name, "PegaOPS()", strCAMARQERRO)
    
End Sub


Public Sub VisualizaBotoesPCD(strSTATUS As String, cTipOper As String)

    Dim boolPermite As Boolean
    
    '' Botão P.Cota/P.data
    cmdLibPcotaPData.Visible = False
    cmdCancLibPcotaPdata.Visible = False
    
    boolPermite = PermiteLibPDataPCota
        
    If cTipOper = "C" Then
        If strSTATUS = "4" Or strSTATUS = "C" Then
           '' P.Cota/P.data
           cmdLibPcotaPData.Visible = boolPermite
           cmdLibPcotaPData.Enabled = True
           
           cmdCancLibPcotaPdata.Visible = boolPermite
           cmdCancLibPcotaPdata.Enabled = False
        End If
    ElseIf cTipOper = "LC" Then
        If strSTATUS = "4" Or strSTATUS = "C" Then
           '' P.Cota/P.data
           cmdLibPcotaPData.Visible = boolPermite
           cmdLibPcotaPData.Enabled = False
        
           cmdCancLibPcotaPdata.Visible = boolPermite
           cmdCancLibPcotaPdata.Enabled = True
        End If
    End If

End Sub

Private Function PegaStatus(cTipOper As String) As String

    PegaStatus = ""
    
    If cTipOper = "I" Then PegaStatus = "incluso"
    If cTipOper = "A" Then PegaStatus = "alterado"
    If cTipOper = "LF" Or cTipOper = "N" Or cTipOper = "LS" Or cTipOper = "LN" Or cTipOper = "LC" Or cTipOper = "LV" Then PegaStatus = "Liberado"
    If cTipOper = "R" Then PegaStatus = "Reprovado"
    If cTipOper = "D" Then PegaStatus = "Bloqueado"

End Function


Private Function Consulta_StatusProd(strIDPRODUTO As String) As Boolean

    Consulta_StatusProd = False
    
    If BREC12.State = 1 Then BREC12.Close
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_STATUS" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(strIDPRODUTO) & vbCrLf
    sSql = sSql & "   And SGI_STATUS    = 2" '' Aguardando Liberação

    BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC12.EOF() Then Consulta_StatusProd = True
    BREC12.Close
    
End Function


Public Sub VisualizaBotoesLibAlteracao(strSTATUS As String, cTipOper As String)

    Dim boolPermite As Boolean
    
    '' Botão Libera Alteração
    cmdLibAlteracao.Visible = False
    cmdCancAlteracao.Visible = False
    
    boolPermite = PermiteLibPedBloq
        
    If cTipOper = "C" Or cTipOper = "A" Then
        If strSTATUS = "S" Then
           If cTipOper = "A" Then
                cmdLibAlteracao.Visible = boolPermite
                cmdLibAlteracao.Enabled = False
           Else
                cmdLibAlteracao.Visible = boolPermite
                cmdLibAlteracao.Enabled = True
           End If
           
           cmdCancAlteracao.Visible = boolPermite
           cmdCancAlteracao.Enabled = False
        ElseIf strSTATUS = "G" Then
           cmdLibAlteracao.Visible = boolPermite
           cmdLibAlteracao.Enabled = True
           
           cmdCancAlteracao.Visible = boolPermite
           cmdCancAlteracao.Enabled = False
        End If
    ElseIf cTipOper = "LS" Then
        If strSTATUS = "S" Then
           cmdLibAlteracao.Visible = boolPermite
           cmdLibAlteracao.Enabled = False
        
           cmdCancAlteracao.Visible = boolPermite
           cmdCancAlteracao.Enabled = True
        End If
    End If

End Sub


Private Function PermiteLibPedBloq() As Boolean

    PermiteLibPedBloq = False
    
    If lngCodUsuario = 0 Then
       PermiteLibPedBloq = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBCOMSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO       = " & lngCodUsuario & vbCrLf
    sSql = sSql & "   And SGI_LIBPEDBLOQSN = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibPedBloq = True
    BREC.Close

End Function


Private Function PermiteLibPedFotolito() As Boolean

    PermiteLibPedFotolito = False
    
    If lngCodUsuario = 0 Then
       PermiteLibPedFotolito = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBCOMSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO       = " & lngCodUsuario & vbCrLf
    sSql = sSql & "   And SGI_LIBPEDFOTSN  = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibPedFotolito = True
    BREC.Close

End Function


Public Sub VisualizaBotoesLibFotolito(strSTATUS As String, cTipOper As String)

    Dim boolPermite As Boolean
    
    '' Botão Libera Fotolito
    cmdLibFot.Visible = False
    cmdCancLibFot.Visible = False
    
    boolPermite = PermiteLibPedFotolito
        
    If cTipOper = "C" Then
        If strSTATUS = "V" Then
           cmdLibFot.Visible = boolPermite
           cmdLibFot.Enabled = True
           
           cmdCancLibFot.Visible = boolPermite
           cmdCancLibFot.Enabled = False
        End If
    ElseIf cTipOper = "LV" Then
        cmdLibFot.Visible = boolPermite
        cmdLibFot.Enabled = False
        
        cmdCancLibFot.Visible = boolPermite
        cmdCancLibFot.Enabled = True
    End If

End Sub


Private Function PegaCodNF(strCODORD As String, strNOMTABE As String) As String

    PegaCodNF = ""
    
    If BREC11.State = 1 Then BREC11.Close
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_CODCONF" & vbCrLf
    sSql = sSql & "      ,SGI_CODFATURA" & vbCrLf
    sSql = sSql & "      ,SGI_DATACONF" & vbCrLf
     
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADORDCONFH" & strNOMTABE & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODORD = " & strCODORD

    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then PegaCodNF = Trim(Str(BREC11!SGI_CODFATURA)) & vbTab & Format(BREC11!SGI_DATACONF, "DD/MM/YYYY") & vbTab & Trim(Str(BREC11!SGI_CODCONF))
    BREC11.Close
    
End Function

Private Sub MostraDadosProgEntr()

    With grdProgEntrega
        If (.Rows - 1) = 0 And .RowSel = 0 Then Exit Sub
        
        If .RowSel > 0 Then
            Dim lngCODIDPROD    As Long
            Dim strCODOP        As String
            lngCODIDPROD = 0
            strCODOP = ""
            If Len(Trim(.Cell(flexcpText, .RowSel, conCOL_SonProgEntr_IdProduto))) > 0 Then
                lngCODIDPROD = CLng(.Cell(flexcpText, .RowSel, conCOL_SonProgEntr_IdProduto))
                strCODOP = Trim(Replace(.Cell(flexcpText, .RowSel, conCOL_SonProgEntr_CodOP), "/", ""))
                Call PopGrdOrdFat_Prod(Str(lngCODIDPROD), .RowSel, strCODOP)
                Call PopGrdProgramacao(objCADPEDVENDA.CODPEDIDO, strCODOP, Str(lngCODIDPROD))
            End If
        End If
    End With

End Sub

Private Sub PegaDescTabelasVend(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label, strFUNCAOPAI As String)

On Error GoTo Err_PegaDescTabelasVend

    lblLabel.Caption = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(Replace(Replace(strCODIGO, ".", ""), ",", ""))) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCODIGO) & vbCrLf
    sSql = sSql & "   And SGI_ATIVO = 1"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
    Exit Sub
    
Err_PegaDescTabelasVend:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDescTabelasVend()" & vbCrLf & "Função Pai :" & strFUNCAOPAI, Me.Name, "PegaDescTabelasVend()", strCAMARQERRO)

End Sub

Public Sub VisualizaBotoesLibComercial(strSTATUS As String, cTipOper As String)

    Dim boolPermite As Boolean
    
    '' Botão Libera Comercial
    cmdLiberaCom.Visible = False
    cmdCancLibCom.Visible = False
    
    boolPermite = objCADPEDVENDA.PermiteLibComercial(lngCodUsuario)
        
    If cTipOper = "C" Then
        If strSTATUS = "B" Then
           cmdLiberaCom.Visible = boolPermite
           cmdLiberaCom.Enabled = True
        
           cmdCancLibCom.Visible = boolPermite
           cmdCancLibCom.Enabled = False
        End If
    ElseIf cTipOper = "LN" Then
        cmdLiberaCom.Visible = boolPermite
        cmdLiberaCom.Enabled = False
    
        cmdCancLibCom.Visible = boolPermite
        cmdCancLibCom.Enabled = True
    End If

End Sub

Public Sub VisualizaBotoesLibFinanceira(strSTATUS As String, cTipOper As String)

    Dim boolPermite As Boolean
    
    '' Botão Libera Financeiro
    cmdLiberaFinanceiro.Visible = False
    cmdCancLibFin.Visible = False
    
    boolPermite = objCADPEDVENDA.PermiteLibFinanceiro(lngCodUsuario)
        
    If cTipOper = "C" Then
        If strSTATUS = "N" Then
           cmdLiberaFinanceiro.Visible = boolPermite
           cmdLiberaFinanceiro.Enabled = True
        
           cmdCancLibFin.Visible = boolPermite
           cmdCancLibFin.Enabled = False
        End If
    ElseIf cTipOper = "LF" Then
        cmdLiberaFinanceiro.Visible = boolPermite
        cmdLiberaFinanceiro.Enabled = False
    
        cmdCancLibFin.Visible = boolPermite
        cmdCancLibFin.Enabled = True
    End If

End Sub

Private Sub InitGridReprovacao()

    With grdTIPREPROV
    
       .Cols = conColumnsIn_SonRep
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonRep_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonRep_Codigo) = ""
       .ColDataType(conCOL_SonRep_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonRep_Pesq) = ""
       .ColDataType(conCOL_SonRep_Pesq) = flexDTString
       .ColComboList(conCOL_SonRep_Pesq) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonRep_Desc) = ""
       .ColDataType(conCOL_SonRep_Desc) = flexDTString
       
       .ColWidth(conCOL_SonRep_Codigo) = 1200
       .ColWidth(conCOL_SonRep_Pesq) = 300
       .ColWidth(conCOL_SonRep_Desc) = 5000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGridRep()
   
On Error GoTo Err_IncRegGridRep
    
    Dim strCampos01 As String
    
    
    If objBLBFunc.FcExisteLinhaVazia(grdTIPREPROV, conCOL_SonRep_Codigo) = False Then Exit Sub
    
    strCampos01 = "" & vbTab & _
                  "" & vbTab & _
                  ""
                       
    grdTIPREPROV.AddItem strCampos01
    
    Exit Sub

Err_IncRegGridRep:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : IncRegGridRep()", Me.Name, "IncRegGridRep()", strCAMARQERRO)
                            
End Sub

Private Function PegaDescrReprovacao(lngCodRep As Long) As String
    
    PegaDescrReprovacao = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPREP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodRep
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrReprovacao = BREC2!SGI_DESCRI
    BREC2.Close
    
End Function

Private Sub PopGrdRep()

    Dim i As Integer

    '' Tipo de Reprovação
    arrTIPREPROVA = objCADPEDVENDA.TIPREPROVA
    If IsArray(arrTIPREPROVA) = True Then
        With grdTIPREPROV
            For i = 1 To UBound(arrTIPREPROVA)
                .AddItem arrTIPREPROVA(i) & vbTab & _
                         "" & vbTab & _
                         PegaDescrReprovacao(CLng(arrTIPREPROVA(i)))
            Next i
        End With
    End If

End Sub

Private Sub MostraDadosReprovacao()

       '' Se reprovado
       If objCADPEDVENDA.STATUS = "R" Or objCADPEDVENDA.STATUS = "N" Or objCADPEDVENDA.STATUS = "L" Then
            
          stCAMPOSVENDA.TabVisible(2) = True
          If cTipOper = "LN" Then stCAMPOSVENDA.TabVisible(2) = False
          txtOBS.Locked = True
          txtOBS.Text = objCADPEDVENDA.OBSCOMERCIAL
          Call PopGrdRep
       
       End If

End Sub

Private Sub Volta_Cres_Grid(lngRowSel As Long)
    grdProduto.Cell(flexcpBackColor, grdProduto.RowSel, conCOL_SonProd_IdProduto, grdProduto.RowSel, conCOL_SonProd_OS_Artes) = &H8000000E
    grdProduto.Cell(flexcpForeColor, grdProduto.RowSel, conCOL_SonProd_IdProduto, grdProduto.RowSel, conCOL_SonProd_OS_Artes) = &H80000008
End Sub


Private Sub LidTimeProducao()
    
    Dim i           As Long
    Dim dtDataEntr  As Date
    Dim dtDataLito  As Date
    Dim lngQTDIAS   As Long
    
    With grdProgEntrega
        For i = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega))) > 0 Then
                dtDataEntr = CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataEntrega))
                lngQTDIAS = (dtDataEntr - Date)
                
                '' Previsão de Produção ( Montagem )
                .Cell(flexcpText, i, conCOL_SonProgEntr_DataPrevProd) = Format((dtDataEntr - 2), "DD/MM/YYYY")
                
                '' Previsão de Litografia
                dtDataLito = CDate(.Cell(flexcpText, i, conCOL_SonProgEntr_DataPrevProd))
                .Cell(flexcpText, i, conCOL_SonProgEntr_DataPrevLito) = Format((dtDataLito - 4), "DD/MM/YYYY")
                
                
                ''If lngQTDIAS > 3 Then
                ''    .Cell(flexcpBackColor, i, lngDTENTREGA, i, lngDtPrevEntr) = &H80FF80
                ''ElseIf lngQTDIAS <= 3 And lngQTDIAS >= 1 Then
                ''    .Cell(flexcpBackColor, i, lngDTENTREGA, i, lngDtPrevEntr) = &H80FFFF
                ''ElseIf lngQTDIAS <= 0 Then
                ''    .Cell(flexcpBackColor, i, lngDTENTREGA, i, lngDtPrevEntr) = &H8080FF
                ''End If
                
                ''If lngQTDIAS > 1 Then .Cell(flexcpText, i, lngDiasVenc) = Format(lngQTDIAS, "##00") & " Dias"
                ''If lngQTDIAS = 1 Then .Cell(flexcpText, i, lngDiasVenc) = Format(lngQTDIAS, "##00") & " Dia"
                ''If lngQTDIAS = 0 Then .Cell(flexcpText, i, lngDiasVenc) = "VENCIDO"
                ''If lngQTDIAS < 0 Then
                ''    If (lngQTDIAS * -1) = 1 Then .Cell(flexcpText, i, lngDiasVenc) = Format((lngQTDIAS * -1), "##00") & " DIA VENCIDO"
                ''    If (lngQTDIAS * -1) > 1 Then .Cell(flexcpText, i, lngDiasVenc) = Format((lngQTDIAS * -1), "##00") & " DIAS VENCIDO"
                ''End If
            End If
        Next i
    End With

End Sub


Private Sub PegaPrevMontagem(strCODIDOP As String, strCODOP As String, lngLINHA As Long)

    
    If Len(Trim(strCODOP)) = 0 Then Exit Sub
    If Len(Trim(strCODIDOP)) = 0 Then Exit Sub
    
    Dim strModulo As String
    
    strModulo = ""
    If intFILIALPED = 1 Then strModulo = "_STEEL"
    
    With grdProgEntrega
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       *" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADMOVPCP" & strModulo & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODOP     = " & strCODOP & vbCrLf
        sSql = sSql & "   And SGI_IDINTERNO = " & strCODIDOP
        
        BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC11.EOF() Then
            .Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_DataPrevProd) = Format(BREC11!SGI_DATAPROG, "DD/MM/YYYY")
            .Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_CODIDPROG) = BREC11!SGI_CODINTENO
        Else
            .Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_DataPrevProd) = Empty
            .Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_CODIDPROG) = Empty
        End If
        BREC11.Close
    End With

    
End Sub


Private Sub PegaStatusMontagem(strCODIDOP As String, strCODOP As String, strIDINTPROG As String, lngLINHA As Long)

    
    If Len(Trim(strCODOP)) = 0 Then Exit Sub
    If Len(Trim(strCODIDOP)) = 0 Then Exit Sub
    If Len(Trim(strIDINTPROG)) = 0 Then Exit Sub
    
    Dim strModulo As String
    
    strModulo = ""
    If intFILIALPED = 1 Then strModulo = "_STEEL"
    
    With grdProgEntrega
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       *" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADAPONTPROG" & strModulo & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODOP       = " & strCODOP & vbCrLf
        sSql = sSql & "   And SGI_IDINTOP     = " & strCODIDOP & vbCrLf
        sSql = sSql & "   And SGI_IDINTPROG   = " & strIDINTPROG
        
        BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC11.EOF() Then
            .Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_CODSTATAPONT) = BREC11!SGI_STATUSAPONT
            .Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_DESCSTATUSAPONT) = PegaDescrStatusApontamento(Str(BREC11!SGI_STATUSAPONT))
        Else
            .Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_CODSTATAPONT) = Empty
            .Cell(flexcpText, lngLINHA, conCOL_SonProgEntr_DESCSTATUSAPONT) = Empty
        End If
        BREC11.Close
        
    End With
    
End Sub

Private Function PegaDescrStatusApontamento(strCODSTATUS As String)

    PegaDescrStatusApontamento = ""
    
    If Len(Trim(strCODSTATUS)) = 0 Then Exit Function
    If CLng(strCODSTATUS) = 0 Then Exit Function

    If CLng(strCODSTATUS) = 1 Then PegaDescrStatusApontamento = "Concluido"
    If CLng(strCODSTATUS) = 2 Then PegaDescrStatusApontamento = "Parcial"
    If CLng(strCODSTATUS) = 3 Then PegaDescrStatusApontamento = "Em Produção"
    
End Function



Private Sub InitGridProducao()

    With grdPRODUCAO
    
       .Cols = conColumnsIn_SonProducao
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProducao_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProducao_IDPROG) = ""
       .ColDataType(conCOL_SonProducao_IDPROG) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProducao_IDITERNO) = ""
       .ColDataType(conCOL_SonProducao_IDITERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProducao_CODOP) = ""
       .ColDataType(conCOL_SonProducao_CODOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProducao_IDOP) = ""
       .ColDataType(conCOL_SonProducao_IDOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProducao_IDPROD) = ""
       .ColDataType(conCOL_SonProducao_IDPROD) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProducao_DTPROG) = ""
       .ColDataType(conCOL_SonProducao_DTPROG) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonProducao_QTDPROG) = ""
       .ColDataType(conCOL_SonProducao_QTDPROG) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProducao_CODSTATUS) = ""
       .ColDataType(conCOL_SonProducao_CODSTATUS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProducao_DESCSTATUS) = ""
       .ColDataType(conCOL_SonProducao_DESCSTATUS) = flexDTString
       
       .ColWidth(conCOL_SonProducao_IDPROG) = 0
       .ColWidth(conCOL_SonProducao_IDITERNO) = 0
       .ColWidth(conCOL_SonProducao_CODOP) = 0
       .ColWidth(conCOL_SonProducao_IDOP) = 0
       .ColWidth(conCOL_SonProducao_IDPROD) = 0
       .ColWidth(conCOL_SonProducao_DTPROG) = 1000
       .ColWidth(conCOL_SonProducao_QTDPROG) = 900
       .ColWidth(conCOL_SonProducao_CODSTATUS) = 0
       .ColWidth(conCOL_SonProducao_DESCSTATUS) = 1000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub PopGrdProgramacao(strCODPED As String, strCODOP As String, strIDPROD As String)

    If Len(Trim(strCODPED)) = 0 Then Exit Sub
    If Len(Trim(strCODOP)) = 0 Then Exit Sub
    If Len(Trim(strIDPROD)) = 0 Then Exit Sub
    
    Call InitGridProducao
    
    With grdPRODUCAO
    
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "      MOV.* " & vbCrLf
        sSql = sSql & " From" & vbCrLf
        sSql = sSql & "      SGI_CADMOVPCP" & strNOMFILIAL & " MOV" & vbCrLf
        sSql = sSql & "Where" & vbCrLf
        sSql = sSql & "      MOV.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "  And MOV.SGI_CODPED     = " & Trim(strCODPED) & vbCrLf
        sSql = sSql & "  And MOV.SGI_CODOP      = " & Trim(strCODOP) & vbCrLf
        sSql = sSql & "  And MOV.SGI_IDPRODUTO  = " & Trim(strIDPROD)
        
        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic, adLockReadOnly
        Do While Not BREC10.EOF()
        
            .AddItem BREC10!SGI_CODIGO & vbTab & _
                     BREC10!SGI_CODINTENO & vbTab & _
                     BREC10!SGI_CODOP & vbTab & _
                     BREC10!SGI_IDINTERNO & vbTab & _
                     BREC10!SGI_IDPRODUTO & vbTab & _
                     BREC10!SGI_DATAPROG & vbTab & _
                     BREC10!SGI_QTDEPROD & vbTab & _
                     IIf(IsNull(BREC10!SGI_STATUSAPONT) = True, 0, BREC10!SGI_STATUSAPONT) & vbTab & _
                     ""
        
            
            If Not IsNull(BREC10!SGI_STATUSAPONT) Then .Cell(flexcpText, (.Rows - 1), conCOL_SonProducao_DESCSTATUS) = PegaDescrStatusApontamento(Str(BREC10!SGI_STATUSAPONT))
        
            BREC10.MoveNext
        Loop
        BREC10.Close
    
    End With
End Sub


Private Function PermiteAltPedidoFatParc() As Boolean

    PermiteAltPedidoFatParc = False
    
    If lngCodUsuario = 0 Then
       PermiteAltPedidoFatParc = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_PERMALTPEDFAT" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO        = " & lngCodUsuario & vbCrLf
    sSql = sSql & "   And SGI_PERMALTPEDFAT = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteAltPedidoFatParc = True
    BREC.Close

End Function


Private Function PermiteFechamOP(strCODCLIE As String) As Integer

    PermiteFechamOP = 0
    
    Dim intRESP As Integer
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_PERMFECHOP" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCODCLIE)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        PermiteFechamOP = BREC!SGI_PERMFECHOP
    
        If PermiteFechamOP = 0 Then
            intRESP = MsgBox("Permite Fechar a OP faltando 10% ?", vbYesNo + vbQuestion + vbDefaultButton2)
            If intRESP = vbYes Then PermiteFechamOP = 1
        End If
    
    End If
    BREC.Close

End Function
