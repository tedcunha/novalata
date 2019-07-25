VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCADVENDEDOR 
   Caption         =   "Cadastro de vendedores"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7635
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3855
      Left            =   0
      TabIndex        =   24
      Top             =   3480
      Width           =   7575
      Begin TabDlg.SSTab STniveis 
         Height          =   3495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         _Version        =   393216
         Style           =   1
         Tabs            =   8
         TabsPerRow      =   5
         TabHeight       =   520
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Niveis de desconto"
         TabPicture(0)   =   "frmCADVENDEDOR.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Metas"
         TabPicture(1)   =   "frmCADVENDEDOR.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame7"
         Tab(1).Control(1)=   "Frame6"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Clientes"
         TabPicture(2)   =   "frmCADVENDEDOR.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame9"
         Tab(2).Control(1)=   "Frame8"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Assinatura"
         TabPicture(3)   =   "frmCADVENDEDOR.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame10"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Tipos de Orcamentos"
         TabPicture(4)   =   "frmCADVENDEDOR.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "flxGrdTipoOrca"
         Tab(4).Control(1)=   "Command26"
         Tab(4).Control(2)=   "cmdIncTipoOca"
         Tab(4).ControlCount=   3
         TabCaption(5)   =   "Diversos"
         TabPicture(5)   =   "frmCADVENDEDOR.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label10"
         Tab(5).Control(1)=   "cmdCarrClie"
         Tab(5).Control(2)=   "txtNomPla"
         Tab(5).ControlCount=   3
         TabCaption(6)   =   "Gerenciador"
         TabPicture(6)   =   "frmCADVENDEDOR.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "lnlNomComp(1)"
         Tab(6).Control(1)=   "txtNomComp"
         Tab(6).ControlCount=   2
         TabCaption(7)   =   "Vendedores"
         TabPicture(7)   =   "frmCADVENDEDOR.frx":00C4
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Command5"
         Tab(7).Control(1)=   "Command4"
         Tab(7).Control(2)=   "grdVendedores"
         Tab(7).ControlCount=   3
         Begin VB.CommandButton Command5 
            Height          =   300
            Left            =   -68160
            Picture         =   "frmCADVENDEDOR.frx":00E0
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   1080
            Width           =   300
         End
         Begin VB.CommandButton Command4 
            Height          =   300
            Left            =   -68160
            Picture         =   "frmCADVENDEDOR.frx":022A
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   720
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid grdVendedores 
            Height          =   2655
            Left            =   -74880
            TabIndex        =   61
            Top             =   720
            Width           =   6615
            _cx             =   11668
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
         Begin VB.TextBox txtNomComp 
            Height          =   285
            Left            =   -72960
            MaxLength       =   50
            TabIndex        =   56
            Text            =   "txtNomComp"
            Top             =   720
            Width           =   4935
         End
         Begin VB.TextBox txtNomPla 
            Height          =   285
            Left            =   -73200
            TabIndex        =   54
            Text            =   "txtNomPla"
            Top             =   840
            Width           =   3015
         End
         Begin VB.CommandButton cmdCarrClie 
            Caption         =   "&Carrega Clientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74880
            TabIndex        =   52
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton cmdIncTipoOca 
            Height          =   300
            Left            =   -68280
            Picture         =   "frmCADVENDEDOR.frx":0374
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   720
            Width           =   300
         End
         Begin VB.CommandButton Command26 
            Height          =   300
            Left            =   -68280
            Picture         =   "frmCADVENDEDOR.frx":04BE
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1080
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid flxGrdTipoOrca 
            Height          =   2415
            Left            =   -74880
            TabIndex        =   48
            Top             =   720
            Width           =   6495
            _cx             =   11456
            _cy             =   4260
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
         Begin VB.Frame Frame10 
            Height          =   2535
            Left            =   -74880
            TabIndex        =   44
            Top             =   600
            Width           =   6975
            Begin VB.CommandButton cmdAbreArq 
               Caption         =   "Abre Arquivo"
               Height          =   375
               Left            =   2280
               TabIndex        =   45
               Top             =   1080
               Width           =   1335
            End
            Begin MSComDlg.CommonDialog cmoAbreArq 
               Left            =   2400
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Image Assinatura 
               Height          =   1215
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   2085
            End
         End
         Begin VB.Frame Frame9 
            Height          =   2175
            Left            =   -74880
            TabIndex        =   40
            Top             =   1200
            Width           =   6975
            Begin MSFlexGridLib.MSFlexGrid flxCLIENTES 
               Height          =   1815
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   3201
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
            TabIndex        =   36
            Top             =   600
            Width           =   6975
            Begin VB.CommandButton Command3 
               Height          =   315
               Left            =   1440
               Picture         =   "frmCADVENDEDOR.frx":0608
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton Command2 
               Height          =   315
               Left            =   6600
               Picture         =   "frmCADVENDEDOR.frx":070A
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtCODCLI 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               MaxLength       =   10
               TabIndex        =   38
               Text            =   "txtCODCLI"
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblNomClie 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblNomClie"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   1800
               TabIndex        =   51
               Top             =   240
               Width           =   4815
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Código :"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   585
            End
         End
         Begin VB.Frame Frame7 
            Height          =   2175
            Left            =   -74880
            TabIndex        =   31
            Top             =   1200
            Width           =   6975
            Begin MSFlexGridLib.MSFlexGrid flxMETASVENDAS 
               Height          =   1815
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   3201
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame6 
            Height          =   615
            Left            =   -74880
            TabIndex        =   30
            Top             =   600
            Width           =   6975
            Begin VB.TextBox txtVLVENDAS1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   13
               Text            =   "txtVLVENDA"
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtVLVENDAS2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   14
               Text            =   "txtVLVENDA"
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtCOMISS2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5880
               MaxLength       =   10
               TabIndex        =   15
               Text            =   "txtCOMISS2"
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton Command1 
               Height          =   315
               Left            =   6480
               Picture         =   "frmCADVENDEDOR.frx":080C
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Vendas entre:"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   270
               Width           =   990
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "á"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   2400
               TabIndex        =   33
               Top             =   270
               Width           =   90
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Comissão de :"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   4800
               TabIndex        =   32
               Top             =   270
               Width           =   990
            End
         End
         Begin VB.Frame Frame5 
            Height          =   2175
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   6495
            Begin MSFlexGridLib.MSFlexGrid flxNIVEISDESCONTO 
               Height          =   1815
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   3201
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   6495
            Begin VB.CommandButton cmbGravEsp 
               Height          =   315
               Left            =   6000
               Picture         =   "frmCADVENDEDOR.frx":090E
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtPORCCOM 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5400
               MaxLength       =   10
               TabIndex        =   11
               Text            =   "txtPORCCOM"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtPORC2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2400
               MaxLength       =   10
               TabIndex        =   10
               Text            =   "txtPORC2"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtPORC1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1440
               MaxLength       =   10
               TabIndex        =   9
               Text            =   "txtPORC1"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Comissão de :"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   4320
               TabIndex        =   28
               Top             =   270
               Width           =   990
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "e"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   2160
               TabIndex        =   27
               Top             =   270
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Desconto entre :"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   270
               Width           =   1185
            End
         End
         Begin VB.Label lnlNomComp 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Computador"
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
            Left            =   -74880
            TabIndex        =   55
            Top             =   720
            Width           =   1830
         End
         Begin VB.Label Label10 
            Caption         =   "Nome da Planilha"
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
            Left            =   -74880
            TabIndex        =   53
            Top             =   840
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7575
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
         Picture         =   "frmCADVENDEDOR.frx":0A10
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "frmCADVENDEDOR.frx":0B12
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmCADVENDEDOR.frx":0C14
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   7575
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3000
         TabIndex        =   58
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optAtivo 
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
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optAtivo 
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
            Left            =   840
            TabIndex        =   59
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   1200
         TabIndex        =   47
         Text            =   "cboUsuario"
         Top             =   2040
         Width           =   5175
      End
      Begin VB.TextBox txtMSN 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "txtMSN"
         Top             =   1680
         Width           =   5175
      End
      Begin VB.TextBox txtSKYPE 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "txtSKYPE"
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox txtEMAIL 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtEMAIL"
         Top             =   960
         Width           =   5175
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ativo"
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
         Index           =   5
         Left            =   2280
         TabIndex        =   57
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Left            =   360
         TabIndex        =   46
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "MSN:"
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
         Left            =   600
         TabIndex        =   5
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Skype:"
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
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
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
         TabIndex        =   43
         Top             =   960
         Width           =   570
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
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmCADVENDEDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Integer
Public FILIAL      As Integer
Public strAcesso   As String
Public strMODPAI   As String
Public strUSUARIO  As String
Dim objBLBFunc     As Object
Dim objCADVENDEDOR As Object
Dim objPESQPADRAO  As Object
Dim arrNIVEISCOMIS As Variant
Dim arrMETAS       As Variant
Dim arrCLIENTES    As Variant
Dim arrTIPOORCA    As Variant
Dim arrVENDEDORES  As Variant
Dim strNOMARG      As String
Dim strCAMINHO     As String

Const conCOL_SonTipoOrca_Cod                     As Integer = 0
Const conCOL_SonTipoOrca_Pesq                    As Integer = 1
Const conCOL_SonTipoOrca_Desc                    As Integer = 2
Const conCOL_SonTipoOrca_FormatString            As String = "=Código|...|Descrição"
Const conColumnsIn_SonTipoOrca                   As Integer = 3

Const conCOL_SonVend_Cod                        As Integer = 0
Const conCOL_SonVend_Pesq                       As Integer = 1
Const conCOL_SonVend_Desc                       As Integer = 2
Const conCOL_SonVend_FormatString               As String = "=Código|...|Descrição"
Const conColumnsIn_SonVend                      As Integer = 3


Private Sub cmbGravEsp_Click()
    If cTipOper = "I" Or cTipOper = "A" Then InserNiveisComiss
End Sub

Private Sub cmdAbreArq_Click()

    Dim arrArquivo As Variant
    Dim I As Integer
    
    cmoAbreArq.FileName = ""
    Call LoadPicture("")
    
    cmoAbreArq.ShowOpen
    
    If Len(Trim(cmoAbreArq.FileName)) = "" Then Exit Sub
    
    Assinatura.Picture = LoadPicture(cmoAbreArq.FileName)
    
    arrArquivo = Split(cmoAbreArq.FileName, "\")
    
    If IsArray(arrArquivo) Then
        strCAMINHO = ""
        For I = 0 To (UBound(arrArquivo) - 1)
            strCAMINHO = strCAMINHO & Trim(arrArquivo(I)) + "\"
        Next I
        If UBound(arrArquivo) > 0 Then strNOMARG = Trim(arrArquivo(UBound(arrArquivo)))
    End If
End Sub

Private Sub cmdAltera_Click()
    
    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame4.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame10.Enabled = True
   
    Me.Caption = "Cadastro de vendedores - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
    txtDescricao.SetFocus
    
    STniveis.Tab = 0

End Sub

Private Sub cmdCarrClie_Click()

    If Len(Trim(txtCodigo.Text)) = 0 Then
        MsgBox "ATENÇÃO - Informe o Vendedor !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If Len(Trim(txtNomPla.Text)) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Informe o nome da planilha !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim intQUEST As Integer
    intQUEST = MsgBox("Tem Certeza que esta carregando a planilha correta ?", vbYesNo + vbQuestion + vbDefaultButton2, "ATENÇÂO")
    If intQUEST = 7 Then Exit Sub
    
    Dim oConn           As ADODB.Connection
    Dim oCmd            As ADODB.Command
    Dim oRS             As ADODB.Recordset
    
    Dim strCODCLIE      As String
    
    ''"Data Source=C:\ricardo\PROGRAMAS\ENTRADAS.xls;"
    ''"Data Source=\\SRVLATA\PROGRAMAS\ENTRADAS.xls;" -- Antigo Novalata
    
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "C:\Ricardo\SGI\NOVALATA\DOCS-NOVALATA\ClientesxVendedores\" & Trim(txtNomPla.Text) & ".xls;" & _
                         "Extended Properties=""Excel 8.0;HDR=Yes;"";"
    
    ' cria o objecto command e define a conexao ativa
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oConn
    
    ' abre a planilha
    oCmd.CommandText = "SELECT * from [Plan1$]"
    
    ' cria o recordset com os dados
    Set oRS = New ADODB.Recordset
    oRS.Open oCmd, , adOpenKeyset, adLockOptimistic

        Do While Not oRS.EOF()
        
            strCODCLIE = ""
            If Not IsNull(oRS(0).Value) Then strCODCLIE = Trim(Replace(oRS(0).Value, " ", ""))
            
            oRS.MoveNext
        Loop
    oRS.Close
    
    Set oCmd = Nothing
    
    Exit Sub
    
Err_Prog:

    MsgBox "Erro       : " & Err.Number & vbCrLf & _
           "Erro Desc. : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"

End Sub

Private Sub cmdIncTipoOca_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridTipoOrca
End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADVENDEDOR.VENDCODIGO = objCADVENDEDOR.Gera_Codigo(Me.Name)
    objCADVENDEDOR.VENDDESCRI = txtDescricao.Text
    objCADVENDEDOR.EMAIL = txtEMAIL.Text
    objCADVENDEDOR.SKYPE = txtSKYPE.Text
    objCADVENDEDOR.MSN = txtMSN.Text
    
    objCADVENDEDOR.NOMCOMP = "Null"
    If Len(Trim(Trim(txtNomComp.Text))) > 0 Then objCADVENDEDOR.NOMCOMP = "'" & Trim(txtNomComp.Text) & "'"
    
    objCADVENDEDOR.CodUsuario = 0
    If cboUsuario.ListIndex > -1 Then objCADVENDEDOR.CodUsuario = cboUsuario.ItemData(cboUsuario.ListIndex)
    
    If optAtivo(0).Value = True Then objCADVENDEDOR.ATIVO = 0
    If optAtivo(1).Value = True Then objCADVENDEDOR.ATIVO = 1
    
    If (flxNIVEISDESCONTO.Rows - 1) > 0 Then
       ReDim arrNIVEISCOMIS(1 To (flxNIVEISDESCONTO.Rows - 1), 1 To 3) As String
       For I = 1 To (flxNIVEISDESCONTO.Rows - 1)
           arrNIVEISCOMIS(I, 1) = flxNIVEISDESCONTO.TextMatrix(I, 1)
           arrNIVEISCOMIS(I, 2) = flxNIVEISDESCONTO.TextMatrix(I, 2)
           arrNIVEISCOMIS(I, 3) = flxNIVEISDESCONTO.TextMatrix(I, 3)
       Next I
       objCADVENDEDOR.NIVELCOMIS = arrNIVEISCOMIS
    End If
    
    If (flxMETASVENDAS.Rows - 1) > 0 Then
       ReDim arrMETAS(1 To (flxMETASVENDAS.Rows - 1), 1 To 3) As String
       For I = 1 To (flxMETASVENDAS.Rows - 1)
           arrMETAS(I, 1) = flxMETASVENDAS.TextMatrix(I, 1)
           arrMETAS(I, 2) = flxMETASVENDAS.TextMatrix(I, 2)
           arrMETAS(I, 3) = flxMETASVENDAS.TextMatrix(I, 3)
       Next I
       objCADVENDEDOR.METAS = arrMETAS
    End If
    
    
    arrCLIENTES = Empty
    If (flxCLIENTES.Rows - 1) > 0 Then
       ReDim arrCLIENTES(1 To (flxCLIENTES.Rows - 1)) As String
       For I = 1 To (flxCLIENTES.Rows - 1)
           arrCLIENTES(I) = flxCLIENTES.TextMatrix(I, 1)
       Next I
    End If
    objCADVENDEDOR.CLIENTES = arrCLIENTES
    
    '' Tipos de Orçamento
    arrTIPOORCA = Empty
    If (flxGrdTipoOrca.Rows - 1) Then
        ReDim arrTIPOORCA(1 To (flxGrdTipoOrca.Rows - 1)) As String
        For I = 1 To (flxGrdTipoOrca.Rows - 1)
            arrTIPOORCA(I) = flxGrdTipoOrca.Cell(flexcpText, I, conCOL_SonTipoOrca_Cod)
        Next I
    End If
    objCADVENDEDOR.TIPOORCA = arrTIPOORCA
    
    '' Vendedores
    arrVENDEDORES = Empty
    With grdVendedores
        If (.Rows - 1) > 0 Then
            ReDim arrVENDEDORES(1 To (.Rows - 1)) As String
            For I = 1 To (.Rows - 1)
                arrVENDEDORES(I) = .Cell(flexcpText, I, conCOL_SonVend_Cod)
            Next I
        End If
    End With
    objCADVENDEDOR.VENDEDORES = arrVENDEDORES
    
    
    objCADVENDEDOR.NOMEARQ = strNOMARG
    objCADVENDEDOR.CAMINHO = strCAMINHO
    
    '' ------------------------------
    '' Gravando o Arquivo Imagem
    objCADVENDEDOR.CONTARQ = Empty
    
    sSql = "Select SGI_IMAGEM From SGI_CADVENDEDOR Where SGI_FILIAL = " & FILIAL & " And SGI_CODIGO = " & objCADVENDEDOR.VENDCODIGO
    BREC2.CursorType = adOpenDynamic
    BREC2.LockType = adLockOptimistic
    BREC2.Open sSql, adoBanco_Dados
    If Not BREC2.EOF Then
       objCADVENDEDOR.CONTARQ = objBLBFunc.GravaBlobParaBanco(BREC2, "SGI_IMAGEM", strCAMINHO + strNOMARG)
       BREC2.Update
    End If
    BREC2.Close
    '' ------------------------------
    
    
    If objCADVENDEDOR.GRAVA(cTipOper) = True Then
          
       MsgBox "O vendedor foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
       If cTipOper = "I" Then
          Set objBLBFunc = Nothing
          Set objCADVENDEDOR = Nothing
          Unload Me
       End If
          
    End If
    

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADVENDEDOR = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()
    If cTipOper = "I" Then InserMetas
    If cTipOper = "A" Then InserMetas
End Sub

Private Sub Command2_Click()
    If cTipOper = "I" Or cTipOper = "A" Then InseriClientes
End Sub

Private Sub Command26_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(flxGrdTipoOrca, flxGrdTipoOrca.Row)
End Sub

Private Sub Command3_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão social"
    arrCAMPOS(3, 4) = "3000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODCLI.Text = varRETORNO
       Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblNomClie)
       If Len(Trim(lblNomClie.Caption)) = 0 Then txtCODCLI.Text = ""
       txtCODCLI.SetFocus
    End If
    

End Sub


Private Sub Command4_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridVendedor
End Sub

Private Sub Command5_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdVendedores, grdVendedores.Row)
End Sub

Private Sub flxCLIENTES_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxCLIENTES.Rows = 2 Then flxCLIENTES.Rows = 1
       If flxCLIENTES.Rows > 2 Then flxCLIENTES.RemoveItem (flxCLIENTES.RowSel)
    End If
End Sub

Private Sub flxGrdTipoOrca_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonTipoOrca_Desc
         Cancel = True
    Case conCOL_SonTipoOrca_Cod, _
         conCOL_SonTipoOrca_Pesq
         If cTipOper = "C" Then Cancel = True
    Case Else
        flxGrdTipoOrca.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub flxGrdTipoOrca_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (flxGrdTipoOrca.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonTipoOrca_Pesq
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADESPORCA " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "4000"
            arrCAMPOS(2, 5) = "SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Tipos de Orçamento")
            
            If Len(Trim(varRETORNO)) > 0 Then
               flxGrdTipoOrca.Cell(flexcpText, Row, conCOL_SonTipoOrca_Cod) = varRETORNO
               flxGrdTipoOrca.Cell(flexcpText, Row, conCOL_SonTipoOrca_Desc) = PegaDescrTipoOrcamento(CLng(varRETORNO))
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(flxGrdTipoOrca, Row, conCOL_SonTipoOrca_Cod, varRETORNO) = False Then
               MsgBox "Este Tipo de Orçamento ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               flxGrdTipoOrca.Cell(flexcpText, Row, conCOL_SonTipoOrca_Cod) = ""
               flxGrdTipoOrca.Cell(flexcpText, Row, conCOL_SonTipoOrca_Desc) = ""
               Exit Sub
            End If

    End Select

End Sub

Private Sub flxGrdTipoOrca_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With flxGrdTipoOrca
          Select Case Col
                    Case conCOL_SonTipoOrca_Cod
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                         ''KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
          End Select
     End With
End Sub

Private Sub flxGrdTipoOrca_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With flxGrdTipoOrca
          Select Case Col
                 Case conCOL_SonTipoOrca_Cod
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(flxGrdTipoOrca, Row, conCOL_SonTipoOrca_Cod, .EditText) = False Then
                           MsgBox "Este Tipo de Orçamento ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonTipoOrca_Cod) = ""
                           .Cell(flexcpText, Row, conCOL_SonTipoOrca_Desc) = ""
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrTipoOrcamento(CLng(.EditText)))) = 0 Then
                           MsgBox "Este Tipo de Orçamento não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonTipoOrca_Desc) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonTipoOrca_Cod) = .EditText
                        .Cell(flexcpText, Row, conCOL_SonTipoOrca_Desc) = PegaDescrTipoOrcamento(CLng(.EditText))
          End Select
     End With
End Sub

Private Sub flxMETASVENDAS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxMETASVENDAS.Rows = 2 Then flxMETASVENDAS.Rows = 1
       If flxMETASVENDAS.Rows > 2 Then flxMETASVENDAS.RemoveItem (flxMETASVENDAS.RowSel)
    End If
End Sub

Private Sub flxNIVEISDESCONTO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxNIVEISDESCONTO.Rows = 2 Then flxNIVEISDESCONTO.Rows = 1
       If flxNIVEISDESCONTO.Rows > 2 Then flxNIVEISDESCONTO.RemoveItem (flxNIVEISDESCONTO.RowSel)
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
   Set objCADVENDEDOR = CreateObject("CADVENDEDOR.clsCADVENDEDOR")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
      
   objCADVENDEDOR.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame4.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame10.Enabled = True
   
    Me.Caption = "Cadastro de vendedores - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos Me
    
    txtCodigo.Text = ""
    strNOMARG = ""
    strCAMINHO = ""
    
    Call ConfGridNiveisDesc
    Call ConfGridMEtas
    Call ConfGridCliente
    Call ConfGrdTipoOrca
    Call ConfGrdVendedores
    
    STniveis.Tab = 0
    
    objCADVENDEDOR.PreenchComboUsuario cboUsuario

    cboUsuario.ListIndex = -1
    optAtivo(1).Value = True

    Call LimpaLabel

End Sub


Private Sub ConfGridNiveisDesc()

    flxNIVEISDESCONTO.Rows = 1
    flxNIVEISDESCONTO.Cols = 4
    
    flxNIVEISDESCONTO.TextMatrix(0, 0) = ""
    flxNIVEISDESCONTO.TextMatrix(0, 1) = "% Desconto"
    flxNIVEISDESCONTO.TextMatrix(0, 2) = "% Desconto"
    flxNIVEISDESCONTO.TextMatrix(0, 3) = "% Comissão"
    
    flxNIVEISDESCONTO.ColWidth(0) = 0
    flxNIVEISDESCONTO.ColWidth(1) = 1500
    flxNIVEISDESCONTO.ColWidth(2) = 1500
    flxNIVEISDESCONTO.ColWidth(3) = 1500
    
End Sub

Private Sub grdVendedores_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonVend_Desc
         Cancel = True
    Case conCOL_SonVend_Cod, _
         conCOL_SonVend_Pesq
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdVendedores.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdVendedores_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdVendedores.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonVend_Pesq
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
            
            If cTipOper <> "I" Then
                sSql = sSql & "   And SGI_CODIGO     <> " & objCADVENDEDOR.VENDCODIGO
            End If
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "4500"
            arrCAMPOS(2, 5) = "SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Pesquisa vendedores")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdVendedores.Cell(flexcpText, Row, conCOL_SonVend_Cod) = varRETORNO
               grdVendedores.Cell(flexcpText, Row, conCOL_SonVend_Desc) = PegaDescrVendedor(CLng(varRETORNO))
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(grdVendedores, Row, conCOL_SonVend_Cod, varRETORNO) = False Then
               MsgBox "Este vendedor já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdVendedores.Cell(flexcpText, Row, conCOL_SonVend_Cod) = ""
               grdVendedores.Cell(flexcpText, Row, conCOL_SonVend_Desc) = ""
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdVendedores_KeyPress(KeyAscii As Integer)
     With grdVendedores
          Select Case Col
                    Case conCOL_SonVend_Cod
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                         ''KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
          End Select
     End With
End Sub

Private Sub grdVendedores_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdVendedores
          Select Case Col
                 Case conCOL_SonVend_Cod
                        If .EditText = Empty Then Exit Sub
                        If cTipOper <> "I" Then
                            If CLng(.EditText) = objCADVENDEDOR.VENDCODIGO Then
                                MsgBox "Este vendedor não pode ser incluso !!!", vbOKOnly + vbExclamation, "Aviso"
                                Cancel = True
                                Exit Sub
                            End If
                        End If
                        If objBLBFunc.FcVerifItensRepetidos(grdVendedores, Row, conCOL_SonVend_Cod, .EditText) = False Then
                           MsgBox "Este vendedor já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonVend_Cod) = ""
                           .Cell(flexcpText, Row, conCOL_SonVend_Desc) = ""
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrVendedor(CLng(.EditText)))) = 0 Then
                           MsgBox "Este vendedor não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonVend_Desc) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonVend_Cod) = .EditText
                        .Cell(flexcpText, Row, conCOL_SonVend_Desc) = PegaDescrVendedor(CLng(.EditText))
          End Select
     End With
End Sub

Private Sub txtCODCLI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLI.Name, frmCADVENDEDOR
End Sub

Private Sub txtCODCLI_Validate(Cancel As Boolean)
    
    Dim I       As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtCODCLI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODCLI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCODCLI.Text, lblNomClie)
    If Len(Trim(lblNomClie.Caption)) = 0 Then
       txtCODCLI.Text = ""
       lblNomClie.Caption = ""
       Cancel = True
    End If
    
End Sub

Private Sub txtCOMISS2_GotFocus()
    objBLBFunc.SelecionaCampos txtCOMISS2.Name, frmCADVENDEDOR
End Sub

Private Sub txtCOMISS2_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCOMISS2.Text
End Sub

Private Sub txtCOMISS2_Validate(Cancel As Boolean)

    If Len(Trim(txtCOMISS2.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCOMISS2.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCOMISS2.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCOMISS2.Text = Format(txtCOMISS2.Text, "#,##0.00")

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADVENDEDOR
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtEMAIL_GotFocus()
    objBLBFunc.SelecionaCampos txtEMAIL.Name, frmCADVENDEDOR
End Sub

Private Sub txtMSN_GotFocus()
    objBLBFunc.SelecionaCampos txtMSN.Name, frmCADVENDEDOR
End Sub

Private Sub txtNomComp_GotFocus()
    objBLBFunc.SelecionaCampos txtNomComp.Name, Me
End Sub

Private Sub txtNomComp_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtPORC1_GotFocus()
    objBLBFunc.SelecionaCampos txtPORC1.Name, frmCADVENDEDOR
End Sub

Private Sub txtPORC1_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPORC1.Text
End Sub

Private Sub txtPORC1_Validate(Cancel As Boolean)
    
    If Len(Trim(txtPORC1.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPORC1.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtPORC1.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPORC1.Text = Format(txtPORC1.Text, "#,##0.00")
    
End Sub

Private Sub txtPORC2_GotFocus()
    objBLBFunc.SelecionaCampos txtPORC2.Name, frmCADVENDEDOR
End Sub

Private Sub txtPORC2_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPORC2.Text
End Sub

Private Sub txtPORC2_Validate(Cancel As Boolean)

    If Len(Trim(txtPORC2.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPORC2.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtPORC2.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPORC2.Text = Format(txtPORC2.Text, "#,##0.00")

End Sub

Private Sub txtPORCCOM_GotFocus()
    objBLBFunc.SelecionaCampos txtPORCCOM.Name, frmCADVENDEDOR
End Sub

Private Sub txtPORCCOM_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPORCCOM.Text
End Sub

Private Sub txtPORCCOM_Validate(Cancel As Boolean)

    If Len(Trim(txtPORCCOM.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPORCCOM.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtPORCCOM.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPORCCOM.Text = Format(txtPORCCOM.Text, "#,##0.00")

End Sub

Private Sub InserNiveisComiss()

    Dim dblPorcDesc1 As Double
    Dim dblPorcDesc2 As Double
    Dim dblPorcComis As Double
    
    If Len(Trim(txtPORC1.Text)) = 0 Then
       MsgBox "Informe o primeiro desconto !!!", vbOKOnly + vbCritical, "aviso"
       txtPORC1.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtPORC2.Text)) = 0 Then
       MsgBox "Informe o segundo desconto !!!", vbOKOnly + vbCritical, "aviso"
       txtPORC2.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtPORCCOM.Text)) = 0 Then
       MsgBox "Informe a comissão !!!", vbOKOnly + vbCritical, "aviso"
       txtPORCCOM.SetFocus
       Exit Sub
    End If
    If CDbl(txtPORC1.Text) < 0 Then
       MsgBox "Somente é permitido valores maior que 0 !!!", vbOKOnly + vbCritical, "aviso"
       txtPORC1.Text = 0
       txtPORC1.SetFocus
       Exit Sub
    End If
    If CDbl(txtPORC2.Text) < 0 Then
       MsgBox "Somente é permitido valores maior que 0 !!!", vbOKOnly + vbCritical, "aviso"
       txtPORC2.Text = 0
       txtPORC2.SetFocus
       Exit Sub
    End If
    If CDbl(txtPORCCOM.Text) < 0 Then
       MsgBox "Somente é permitido valores maior que 0 !!!", vbOKOnly + vbCritical, "aviso"
       txtPORCCOM.Text = 0
       txtPORCCOM.SetFocus
       Exit Sub
    End If
    
    dblPorcDesc1 = CDbl(txtPORC1.Text)
    dblPorcDesc2 = CDbl(txtPORC2.Text)
    dblPorcComis = CDbl(txtPORCCOM.Text)
    
    If dblPorcDesc1 > dblPorcDesc2 Then
       MsgBox "A primeira porcentagen não pode ser maior que a segunda !!!", vbOKOnly + vbCritical, "avisa"
       txtPORC1.Text = ""
       txtPORC1.SetFocus
       Exit Sub
    End If
    
    flxNIVEISDESCONTO.AddItem "" & vbTab & txtPORC1.Text & vbTab & txtPORC2.Text & vbTab & txtPORCCOM.Text
    txtPORC1.Text = ""
    txtPORC2.Text = ""
    txtPORCCOM.Text = ""
    txtPORC1.SetFocus
    
End Sub

Private Sub txtSKYPE_GotFocus()
    objBLBFunc.SelecionaCampos txtSKYPE.Name, frmCADVENDEDOR
End Sub

Private Sub txtVLVENDAS1_GotFocus()
    objBLBFunc.SelecionaCampos txtVLVENDAS1.Name, frmCADVENDEDOR
End Sub

Private Sub txtVLVENDAS1_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtVLVENDAS1.Text
End Sub

Private Sub txtVLVENDAS1_Validate(Cancel As Boolean)

    If Len(Trim(txtVLVENDAS1.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtVLVENDAS1.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtVLVENDAS1.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtVLVENDAS1.Text = Format(txtVLVENDAS1.Text, "#,##0.00")

End Sub

Private Sub txtVLVENDAS2_GotFocus()
    objBLBFunc.SelecionaCampos txtVLVENDAS2.Name, frmCADVENDEDOR
End Sub

Private Sub txtVLVENDAS2_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtVLVENDAS2.Text
End Sub

Private Sub txtVLVENDAS2_Validate(Cancel As Boolean)

    If Len(Trim(txtVLVENDAS2.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtVLVENDAS2.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtVLVENDAS2.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtVLVENDAS2.Text = Format(txtVLVENDAS2.Text, "#,##0.00")

End Sub

Private Sub ConfGridMEtas()

    flxMETASVENDAS.Rows = 1
    flxMETASVENDAS.Cols = 4
    
    flxMETASVENDAS.TextMatrix(0, 0) = 0
    flxMETASVENDAS.TextMatrix(0, 1) = "Valor de Venda"
    flxMETASVENDAS.TextMatrix(0, 2) = "Valor de Venda"
    flxMETASVENDAS.TextMatrix(0, 3) = "% Comissão"
    
    flxMETASVENDAS.ColWidth(0) = 0
    flxMETASVENDAS.ColWidth(1) = 1500
    flxMETASVENDAS.ColWidth(2) = 1500
    flxMETASVENDAS.ColWidth(3) = 1500
    
End Sub

Private Sub InserMetas()

    Dim dblMetas1    As Double
    Dim dblMetas2    As Double
    Dim dblPorcComis As Double
    
    If Len(Trim(txtVLVENDAS1.Text)) = 0 Then
       MsgBox "Informe o primeiro valor !!!", vbOKOnly + vbCritical, "aviso"
       txtVLVENDAS1.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtVLVENDAS2.Text)) = 0 Then
       MsgBox "Informe o segundo valor !!!", vbOKOnly + vbCritical, "aviso"
       txtVLVENDAS2.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtCOMISS2.Text)) = 0 Then
       MsgBox "Informe a comissão !!!", vbOKOnly + vbCritical, "aviso"
       txtCOMISS2.SetFocus
       Exit Sub
    End If
    If CDbl(txtVLVENDAS1.Text) < 0 Then
       MsgBox "Somente é permitido valores maior que 0 !!!", vbOKOnly + vbCritical, "aviso"
       txtVLVENDAS1.Text = 0
       txtVLVENDAS1.SetFocus
       Exit Sub
    End If
    If CDbl(txtVLVENDAS2.Text) < 0 Then
       MsgBox "Somente é permitido valores maior que 0 !!!", vbOKOnly + vbCritical, "aviso"
       txtVLVENDAS2.Text = 0
       txtVLVENDAS2.SetFocus
       Exit Sub
    End If
    If CDbl(txtCOMISS2.Text) < 0 Then
       MsgBox "Somente é permitido valores maior que 0 !!!", vbOKOnly + vbCritical, "aviso"
       txtCOMISS2.Text = 0
       txtCOMISS2.SetFocus
       Exit Sub
    End If
    
    dblMetas1 = CDbl(txtVLVENDAS1.Text)
    dblMetas2 = CDbl(txtVLVENDAS2.Text)
    dblPorcComis = CDbl(txtCOMISS2.Text)
    
    If dblMetas1 > dblMetas2 Then
       MsgBox "O primeiro valor não pode ser maior que o segundo !!!", vbOKOnly + vbCritical, "avisa"
       txtVLVENDAS1.Text = ""
       txtVLVENDAS1.SetFocus
       Exit Sub
    End If
    
    flxMETASVENDAS.AddItem "" & vbTab & txtVLVENDAS1.Text & vbTab & txtVLVENDAS2.Text & vbTab & txtCOMISS2.Text
    txtVLVENDAS1.Text = ""
    txtVLVENDAS2.Text = ""
    txtCOMISS2.Text = ""
    txtVLVENDAS1.SetFocus
    
End Sub

Private Sub ConfGridCliente()

    flxCLIENTES.Rows = 1
    flxCLIENTES.Cols = 3
    
    flxCLIENTES.TextMatrix(0, 0) = ""
    flxCLIENTES.TextMatrix(0, 1) = "Código"
    flxCLIENTES.TextMatrix(0, 2) = "Razão Social"
    
    flxCLIENTES.ColWidth(0) = 0
    flxCLIENTES.ColWidth(1) = 700
    flxCLIENTES.ColWidth(2) = 4000
    
End Sub

Private Sub InseriClientes()

    Dim I As Integer
    If Len(Trim(txtCODCLI.Text)) = 0 Then
       MsgBox "Informe o código do cliente !!!", vbOKOnly + vbCritical, "aviso"
       txtCODCLI.SetFocus
       Exit Sub
    End If
    For I = 1 To (flxCLIENTES.Rows - 1)
        If flxCLIENTES.TextMatrix(I, 1) = txtCODCLI.Text Then
           MsgBox "Este cliente já foi relacionado !!!", vbOKOnly + vbCritical, "aviso"
           txtCODCLI.SetFocus
           Exit Sub
        End If
    Next I
    
    ''If VerificaCliente(txtCODCLI.Text) = True Then
    ''    MsgBox "ATENÇÃO" & vbCrLf & _
    ''           "Este cliente já esta relacionado em outro vendedor !!!", vbOKOnly + vbExclamation, "Aviso"
    ''    txtCODCLI.Text = ""
    ''    lblNomClie.Caption = ""
    ''    Exit Sub
    ''End If
    
    flxCLIENTES.AddItem "" & vbTab & txtCODCLI.Text & vbTab & lblNomClie.Caption
    txtCODCLI.Text = ""
    lblNomClie.Caption = ""
    txtCODCLI.SetFocus
    
    
End Sub

Private Function ValidaCampos() As Boolean
   
   ValidaCampos = False
   
   If Len(Trim(txtDescricao.Text)) = 0 Then
      MsgBox "Informe o nome do vendedor !!!", vbOKOnly + vbCritical, "aviso"
      txtDescricao.SetFocus
      Exit Function
   End If
   
   If cTipOper = "I" Then
   
      sSql = "Select " & vbCrLf
      sSql = sSql & "       * " & vbCrLf
      sSql = sSql & "  From " & vbCrLf
      sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
      sSql = sSql & " Where " & vbCrLf
      sSql = sSql & "       SGI_FILIAL    =  " & FILIAL & vbCrLf
      sSql = sSql & "   And SGI_DESCRICAO = '" & txtDescricao.Text & "'"
      
      BREC.Open sSql, adoBanco_Dados, adOpenDynamic
      
      If Not BREC.EOF Then
         MsgBox "Este fornecedor já existe !!!", vbOKOnly + vbCritical, "aviso"
         BREC.Close
         txtDescricao.Text = ""
         txtDescricao.SetFocus
         Exit Function
      End If
      
      BREC.Close
      
   End If
   
   If cTipOper = "A" Then
      If objCADVENDEDOR.VENDDESCRI <> txtDescricao.Text Then
      
         sSql = "Select " & vbCrLf
         sSql = sSql & "       * " & vbCrLf
         sSql = sSql & "  From " & vbCrLf
         sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_FILIAL    =  " & FILIAL & vbCrLf
         sSql = sSql & "   And SGI_DESCRICAO = '" & txtDescricao.Text & "'"
      
         BREC.Open sSql, adoBanco_Dados, adOpenDynamic
      
         If Not BREC.EOF Then
            MsgBox "Este fornecedor já existe !!!", vbOKOnly + vbCritical, "aviso"
            BREC.Close
            txtDescricao.Text = objCADVENDEDOR.VENDDESCRI
            txtDescricao.SetFocus
            Exit Function
         End If
      
         BREC.Close
      
      End If
   End If
   
   ValidaCampos = True

End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame4.Enabled = False
    Frame6.Enabled = False
    Frame8.Enabled = False
    Frame10.Enabled = False
   
    Me.Caption = "Cadastro de vendedores - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADVENDEDOR
    
    txtCodigo.Text = ""
    
    strNOMARG = ""
    strCAMINHO = ""
    
    Call ConfGridNiveisDesc
    Call ConfGridMEtas
    Call ConfGridCliente
    Call ConfGrdTipoOrca
    Call ConfGrdVendedores

    
    STniveis.Tab = 0
    
    objCADVENDEDOR.PreenchComboUsuario cboUsuario
    cboUsuario.ListIndex = -1
    optAtivo(objCADVENDEDOR.ATIVO).Value = True
    
    objCADVENDEDOR.VENDCODIGO = iCodigo
    
    Call LimpaLabel
    
    
    If objCADVENDEDOR.Pesq_Vendedor = True Then
       
       txtCodigo.Text = objCADVENDEDOR.VENDCODIGO
       txtDescricao.Text = objCADVENDEDOR.VENDDESCRI
       txtEMAIL.Text = objCADVENDEDOR.EMAIL
       txtSKYPE.Text = objCADVENDEDOR.SKYPE
       txtMSN.Text = objCADVENDEDOR.MSN
       txtNomComp.Text = objCADVENDEDOR.NOMCOMP
       optAtivo(objCADVENDEDOR.ATIVO).Value = True
       
       arrNIVEISCOMIS = objCADVENDEDOR.NIVELCOMIS
       arrMETAS = objCADVENDEDOR.METAS
       arrCLIENTES = objCADVENDEDOR.CLIENTES
       
       
       strNOMARG = objCADVENDEDOR.NOMEARQ
       strCAMINHO = objCADVENDEDOR.CAMINHO
       
       If IsArray(arrNIVEISCOMIS) = True Then
          For I = 1 To UBound(arrNIVEISCOMIS)
              flxNIVEISDESCONTO.AddItem "" & vbTab & arrNIVEISCOMIS(I, 1) & vbTab & arrNIVEISCOMIS(I, 2) & vbTab & arrNIVEISCOMIS(I, 3)
          Next I
       End If
       If IsArray(arrMETAS) = True Then
          For I = 1 To UBound(arrMETAS)
              flxMETASVENDAS.AddItem "" & vbTab & arrMETAS(I, 1) & vbTab & arrMETAS(I, 2) & vbTab & arrMETAS(I, 3)
          Next I
       End If
       If IsArray(arrCLIENTES) = True Then
          For I = 1 To UBound(arrCLIENTES)
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrCLIENTES(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxCLIENTES.AddItem arrCLIENTES(I) & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_RAZAOSOC
              BREC.Close
          Next I
       End If
       
             
       If Len(Trim(strNOMARG)) > 0 And Len(Trim(Dir(strCAMINHO + strNOMARG))) > 0 Then
          sSql = "Select SGI_IMAGEM from SGI_CADVENDEDOR Where SGI_FILIAL = " & FILIAL & " And SGI_CODIGO = " & objCADVENDEDOR.VENDCODIGO
          BREC2.Open sSql, adoBanco_Dados, adOpenDynamic, adLockOptimistic
          If Not BREC2.EOF Then
             Call objBLBFunc.LeCampoBlobDoDB(BREC2, "SGI_IMAGEM", strNOMARG)
             Assinatura.Picture = LoadPicture(strCAMINHO + strNOMARG)
          End If
          BREC2.Close
       End If
       
       
       If (cboUsuario.ListCount - 1) > 0 Then
            For I = 0 To (cboUsuario.ListCount - 1)
                If cboUsuario.ItemData(I) = objCADVENDEDOR.CodUsuario Then cboUsuario.ListIndex = I
            Next I
       End If
       
       Call PopGrdTipOrca
       Call PopGrdVendedores
       
    End If
    
End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame4.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame10.Enabled = True
   
    Me.Caption = "Cadastro de vendedores - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADVENDEDOR
    
    txtCodigo.Text = ""
    strNOMARG = ""
    strCAMINHO = ""
    
    Call ConfGridNiveisDesc
    Call ConfGridMEtas
    Call ConfGridCliente
    Call ConfGrdTipoOrca
    Call ConfGrdVendedores
    
    STniveis.Tab = 0
    
    objCADVENDEDOR.PreenchComboUsuario cboUsuario
    cboUsuario.ListIndex = -1
    optAtivo(objCADVENDEDOR.ATIVO).Value = True

    objCADVENDEDOR.VENDCODIGO = iCodigo
        
    Call LimpaLabel
    
    If objCADVENDEDOR.Pesq_Vendedor = True Then
       txtCodigo.Text = objCADVENDEDOR.VENDCODIGO
       txtDescricao.Text = objCADVENDEDOR.VENDDESCRI
       txtEMAIL.Text = objCADVENDEDOR.EMAIL
       txtSKYPE.Text = objCADVENDEDOR.SKYPE
       txtMSN.Text = objCADVENDEDOR.MSN
       txtNomComp.Text = objCADVENDEDOR.NOMCOMP
       
       arrNIVEISCOMIS = objCADVENDEDOR.NIVELCOMIS
       arrMETAS = objCADVENDEDOR.METAS
       arrCLIENTES = objCADVENDEDOR.CLIENTES
       
       strNOMARG = objCADVENDEDOR.NOMEARQ
       strCAMINHO = objCADVENDEDOR.CAMINHO
       
       optAtivo(objCADVENDEDOR.ATIVO).Value = True
       
       If IsArray(arrNIVEISCOMIS) = True Then
          For I = 1 To UBound(arrNIVEISCOMIS)
              flxNIVEISDESCONTO.AddItem "" & vbTab & arrNIVEISCOMIS(I, 1) & vbTab & arrNIVEISCOMIS(I, 2) & vbTab & arrNIVEISCOMIS(I, 3)
          Next I
       End If
       If IsArray(arrMETAS) = True Then
          For I = 1 To UBound(arrMETAS)
              flxMETASVENDAS.AddItem "" & vbTab & arrMETAS(I, 1) & vbTab & arrMETAS(I, 2) & vbTab & arrMETAS(I, 3)
          Next I
       End If
       If IsArray(arrCLIENTES) = True Then
          For I = 1 To UBound(arrCLIENTES)
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrCLIENTES(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxCLIENTES.AddItem arrCLIENTES(I) & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_RAZAOSOC
              BREC.Close
          Next I
       End If
       
       If Len(Trim(strNOMARG)) > 0 Then
          sSql = "Select SGI_IMAGEM from SGI_CADVENDEDOR Where SGI_FILIAL = " & FILIAL & " And SGI_CODIGO = " & objCADVENDEDOR.VENDCODIGO
          BREC2.Open sSql, adoBanco_Dados, adOpenDynamic, adLockOptimistic
          If Not BREC2.EOF Then
             Call objBLBFunc.LeCampoBlobDoDB(BREC2, "SGI_IMAGEM", strNOMARG)
             ''if Dir(strCAMINHO & strNOMARG)
             Assinatura.Picture = LoadPicture(strCAMINHO & strNOMARG)
          End If
          BREC2.Close
       End If
    
    End If
    
    If (cboUsuario.ListCount - 1) > 0 Then
         For I = 0 To (cboUsuario.ListCount - 1)
             If cboUsuario.ItemData(I) = objCADVENDEDOR.CodUsuario Then cboUsuario.ListIndex = I
         Next I
    End If
    
    STniveis.Tab = 0
    
    Call PopGrdTipOrca
    Call PopGrdVendedores

End Sub


Private Sub ConfGrdTipoOrca()

    With flxGrdTipoOrca
    
       .Cols = conColumnsIn_SonTipoOrca
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonTipoOrca_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonTipoOrca_Cod) = ""
       .ColDataType(conCOL_SonTipoOrca_Cod) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonTipoOrca_Pesq) = ""
       .ColDataType(conCOL_SonTipoOrca_Pesq) = flexDTString
       .ColComboList(conCOL_SonTipoOrca_Pesq) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonTipoOrca_Desc) = ""
       .ColDataType(conCOL_SonTipoOrca_Desc) = flexDTString
       
       .ColWidth(conCOL_SonTipoOrca_Cod) = 1000
       .ColWidth(conCOL_SonTipoOrca_Pesq) = 300
       .ColWidth(conCOL_SonTipoOrca_Desc) = 4000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub

Private Sub IncRegGridTipoOrca()
   
    If objBLBFunc.FcExisteLinhaVazia(flxGrdTipoOrca, conCOL_SonTipoOrca_Cod) = False Then Exit Sub
    
    flxGrdTipoOrca.AddItem "" & vbTab & _
                           "" & vbTab & _
                           ""
                          
                            
End Sub


Private Function PegaDescrTipoOrcamento(lngCodClie As Long) As String
    PegaDescrTipoOrcamento = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADESPORCA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodClie
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrTipoOrcamento = BREC2!SGI_DESCRICAO
    BREC2.Close
    
End Function


Private Sub PopGrdTipOrca()

    Dim I As Integer
    
    arrTIPOORCA = objCADVENDEDOR.TIPOORCA
    
    If IsArray(arrTIPOORCA) Then
        For I = 1 To UBound(arrTIPOORCA)
            flxGrdTipoOrca.AddItem arrTIPOORCA(I) & vbTab & _
                                   "" & vbTab & _
                                   PegaDescrTipoOrcamento(CLng(arrTIPOORCA(I)))
        Next I
    End If
    
End Sub

Private Function VerificaCliente(strCODCLIE As String) As Boolean

    VerificaCliente = False

    If Len(Trim(strCODCLIE)) = 0 Then Exit Function

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADCLIEVEND" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    If cTipOper = "A" Then
        sSql = sSql & "   And SGI_CODIGO <> " & objCADVENDEDOR.VENDCODIGO & vbCrLf
    End If
    sSql = sSql & "   And SGI_CODCLI  = " & strCODCLIE

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then VerificaCliente = True
    BREC.Close

End Function

Private Sub LimpaLabel()
    lblNomClie.Caption = ""
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label)

On Error GoTo Err_PegaDescTabelas

    lblLabel.Caption = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Sub
    
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


Private Sub ConfGrdVendedores()

    With grdVendedores
    
       .Cols = conColumnsIn_SonVend
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonVend_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonVend_Cod) = ""
       .ColDataType(conCOL_SonVend_Cod) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonVend_Pesq) = ""
       .ColDataType(conCOL_SonVend_Pesq) = flexDTString
       .ColComboList(conCOL_SonVend_Pesq) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonVend_Desc) = ""
       .ColDataType(conCOL_SonVend_Desc) = flexDTString
       
       .ColWidth(conCOL_SonVend_Cod) = 1000
       .ColWidth(conCOL_SonVend_Pesq) = 300
       .ColWidth(conCOL_SonVend_Desc) = 4500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub


Private Function PegaDescrVendedor(lngCodVend As Long) As String
    PegaDescrVendedor = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodVend
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrVendedor = BREC2!SGI_DESCRICAO
    BREC2.Close
    
End Function


Private Sub IncRegGridVendedor()
   
    If objBLBFunc.FcExisteLinhaVazia(grdVendedores, conCOL_SonVend_Cod) = False Then Exit Sub
    
    grdVendedores.AddItem "" & vbTab & _
                          "" & vbTab & _
                          ""
                          
                            
End Sub

Private Sub PopGrdVendedores()

    Dim I As Integer
    
    arrVENDEDORES = objCADVENDEDOR.VENDEDORES
    
    If IsArray(arrVENDEDORES) Then
        With grdVendedores
            For I = 1 To UBound(arrVENDEDORES)
                .AddItem arrVENDEDORES(I) & vbTab & _
                         "" & vbTab & _
                         PegaDescrVendedor(CLng(arrVENDEDORES(I)))
            Next I
        End With
    End If
    
End Sub


