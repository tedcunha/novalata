VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADSENHA 
   Caption         =   "Cadastro de Senhas e Usuários"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab2 
      Height          =   7695
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "frmCADSENHA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Menus"
      TabPicture(1)   =   "frmCADSENHA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton Command5 
         Caption         =   "Carrega Menu"
         Height          =   495
         Left            =   -69240
         TabIndex        =   76
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Frame Frame12 
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
         Height          =   3735
         Left            =   -75000
         TabIndex        =   59
         Top             =   3840
         Width           =   10335
         Begin VSFlex8LCtl.VSFlexGrid grdMENUNETO 
            Height          =   3375
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   10095
            _cx             =   17806
            _cy             =   5953
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
         Caption         =   "[ Principal ]"
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
         Height          =   3495
         Left            =   -75000
         TabIndex        =   56
         Top             =   360
         Width           =   5655
         Begin VB.CommandButton Command4 
            Height          =   300
            Left            =   5280
            Picture         =   "frmCADSENHA.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   240
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid grdMENUPAI 
            Height          =   3135
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   5055
            _cx             =   8916
            _cy             =   5530
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
      Begin VB.Frame Frame10 
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
         Height          =   1575
         Left            =   -69240
         TabIndex        =   54
         Top             =   360
         Width           =   4575
         Begin VSFlex8LCtl.VSFlexGrid grdMENUFILHO 
            Height          =   1215
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   4335
            _cx             =   7646
            _cy             =   2143
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   3855
         Left            =   120
         TabIndex        =   29
         Top             =   3720
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6800
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
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
         TabCaption(0)   =   "E-Mail"
         TabPicture(0)   =   "frmCADSENHA.frx":0182
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "grdEMAIL"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdIncGrid"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdDelGrid"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Acessos - Pedidos"
         TabPicture(1)   =   "frmCADSENHA.frx":019E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label10"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame3"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame4"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame5"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame6"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Frame7"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Frame8"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Frame9"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Frame17(0)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Frame19"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Frame20"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Frame23"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Frame17(1)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).ControlCount=   13
         TabCaption(2)   =   "Acessos - Produtos"
         TabPicture(2)   =   "frmCADSENHA.frx":01BA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame14"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame15"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Acessos - Ordem de Faturamento"
         TabPicture(3)   =   "frmCADSENHA.frx":01D6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame16"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Clientes"
         TabPicture(4)   =   "frmCADSENHA.frx":01F2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label9(0)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "Label9(1)"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "Frame18"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "Frame22"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).ControlCount=   4
         TabCaption(5)   =   "Diversos"
         TabPicture(5)   =   "frmCADSENHA.frx":020E
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame24"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).ControlCount=   1
         Begin VB.Frame Frame17 
            Caption         =   "[ Permite Excluir Pedidos ]"
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
            Left            =   3360
            TabIndex        =   101
            Top             =   2880
            Width           =   3135
            Begin VB.OptionButton optPermExcPedSN 
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
               Left            =   120
               TabIndex        =   103
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optPermExcPedSN 
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
               Left            =   1200
               TabIndex        =   102
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame24 
            Caption         =   "[ Este usuário Mascara OP's ]"
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
            Left            =   -75000
            TabIndex        =   98
            Top             =   360
            Width           =   2895
            Begin VB.OptionButton optMOPSN 
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
               Left            =   360
               TabIndex        =   100
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optMOPSN 
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
               Left            =   1560
               TabIndex        =   99
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame23 
            Caption         =   "[ Permite Alterar Pedidos já faturados Parcialmente ]"
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
            Left            =   3360
            TabIndex        =   95
            Top             =   2280
            Width           =   4815
            Begin VB.OptionButton optPERMALTPEDFAT 
               Caption         =   "NÃO"
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
               Left            =   1440
               TabIndex        =   97
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optPERMALTPEDFAT 
               Caption         =   "SIM"
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
               Left            =   360
               TabIndex        =   96
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame22 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   -68760
            TabIndex        =   92
            Top             =   960
            Width           =   3135
            Begin VB.OptionButton optLimCred 
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
               TabIndex        =   94
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton optLimCred 
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
               Left            =   1080
               TabIndex        =   93
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "[ É Vendedor ]"
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
            Left            =   6600
            TabIndex        =   85
            Top             =   420
            Width           =   3495
            Begin VB.OptionButton optEVENDEDOR 
               Caption         =   "SIM"
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
               Left            =   120
               TabIndex        =   87
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optEVENDEDOR 
               Caption         =   "NÃO"
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
               Left            =   1080
               TabIndex        =   86
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame19 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   7920
            TabIndex        =   82
            Top             =   3540
            Width           =   1815
            Begin VB.OptionButton optPVCLIE 
               Caption         =   "SIM"
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
               Left            =   0
               TabIndex        =   84
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton optPVCLIE 
               Caption         =   "NÃO"
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
               Left            =   840
               TabIndex        =   83
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame Frame18 
            BorderStyle     =   0  'None
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
            Left            =   -68640
            TabIndex        =   77
            Top             =   540
            Width           =   3015
            Begin VB.OptionButton optPermFatRotDifSN 
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
               Left            =   0
               TabIndex        =   79
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton optPermFatRotDifSN 
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
               TabIndex        =   78
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "[ Libera P.Data / P.Cota ]"
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
            Left            =   3360
            TabIndex        =   73
            Top             =   1620
            Width           =   3135
            Begin VB.OptionButton optLIBPDATAPCOTA 
               Caption         =   "SIM"
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
               Left            =   480
               TabIndex        =   75
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optLIBPDATAPCOTA 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   74
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "[ Permite Faturar com mais de 10% ]"
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
            Left            =   -74880
            TabIndex        =   70
            Top             =   420
            Width           =   3375
            Begin VB.OptionButton optPERMFAT10POR 
               Caption         =   "SIM"
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
               Left            =   600
               TabIndex        =   72
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optPERMFAT10POR 
               Caption         =   "NÃO"
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
               Left            =   1800
               TabIndex        =   71
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "[ Permite Liberar Fotolito ]"
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
            Left            =   -74880
            TabIndex        =   67
            Top             =   900
            Width           =   3855
            Begin VB.OptionButton optPERMLIBFOT 
               Caption         =   "SIM"
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
               Left            =   480
               TabIndex        =   69
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optPERMLIBFOT 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   68
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "[ Desabilita Produto ]"
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
            Left            =   -74880
            TabIndex        =   64
            Top             =   300
            Width           =   3855
            Begin VB.OptionButton optDESABPROD 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   66
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optDESABPROD 
               Caption         =   "SIM"
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
               Left            =   480
               TabIndex        =   65
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "[ Libera Fotolito ]"
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
            Left            =   3360
            TabIndex        =   51
            Top             =   1020
            Width           =   3135
            Begin VB.OptionButton optLIBFOTSN 
               Caption         =   "SIM"
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
               Left            =   480
               TabIndex        =   53
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optLIBFOTSN 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   52
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "[ Libera Pedidos Bloqueados ]"
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
            Left            =   3360
            TabIndex        =   48
            Top             =   420
            Width           =   3135
            Begin VB.OptionButton optLIBPEDBLOQ 
               Caption         =   "SIM"
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
               Left            =   480
               TabIndex        =   50
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optLIBPEDBLOQ 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   49
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "[ Liquida Pedido ]"
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
            Left            =   0
            TabIndex        =   45
            Top             =   2820
            Width           =   3255
            Begin VB.OptionButton optLIQPEDSN 
               Caption         =   "SIM"
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
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optLIQPEDSN 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   46
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Reprova Pedido ]"
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
            Left            =   0
            TabIndex        =   42
            Top             =   2220
            Width           =   3255
            Begin VB.OptionButton optREPEDSN 
               Caption         =   "SIM"
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
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton optREPEDSN 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   43
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Libera Comercial ]"
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
            Left            =   0
            TabIndex        =   39
            Top             =   1620
            Width           =   3255
            Begin VB.OptionButton optLIBCOMSN 
               Caption         =   "SIM"
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
               Left            =   240
               TabIndex        =   41
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optLIBCOMSN 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   40
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Libera Financeiro ]"
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
            Left            =   0
            TabIndex        =   36
            Top             =   1020
            Width           =   3255
            Begin VB.OptionButton optLIBFINSN 
               Caption         =   "SIM"
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
               Left            =   240
               TabIndex        =   38
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optLIBFINSN 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   37
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Bloqueia Pedido ]"
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
            Left            =   0
            TabIndex        =   33
            Top             =   420
            Width           =   3255
            Begin VB.OptionButton optPERMBPEDSN 
               Caption         =   "SIM"
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
               Left            =   240
               TabIndex        =   35
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton optPERMBPEDSN 
               Caption         =   "NÃO"
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
               Left            =   1560
               TabIndex        =   34
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdDelGrid 
            Height          =   300
            Left            =   -65160
            Picture         =   "frmCADSENHA.frx":022A
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   780
            Width           =   300
         End
         Begin VB.CommandButton cmdIncGrid 
            Height          =   300
            Left            =   -65160
            Picture         =   "frmCADSENHA.frx":0374
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   420
            Width           =   300
         End
         Begin VSFlex8LCtl.VSFlexGrid grdEMAIL 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   32
            Top             =   420
            Width           =   9615
            _cx             =   16960
            _cy             =   5953
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Permite este Usuário Mudar o limite de crédito"
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
            Left            =   -74760
            TabIndex        =   91
            Top             =   960
            Width           =   3915
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Permite este Usuário de pesquisar outro vendedor na tela de pesquisa inicial de pedidos"
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
            TabIndex        =   81
            Top             =   3540
            Width           =   7530
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Permite este Usuário Liberar Faturamento de Rótulos Separadamente"
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
            Left            =   -74760
            TabIndex        =   80
            Top             =   540
            Width           =   5865
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   10215
         Begin VB.Frame Frame21 
            Caption         =   "[ Ativo ]"
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
            Left            =   6600
            TabIndex        =   88
            Top             =   1200
            Width           =   2415
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
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   90
               Top             =   240
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
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   89
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame13 
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
            Height          =   615
            Left            =   6600
            TabIndex        =   61
            Top             =   600
            Width           =   2415
            Begin VB.OptionButton optTIPO 
               Caption         =   "ANTIGO"
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
               Left            =   1200
               TabIndex        =   63
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optTIPO 
               Caption         =   "NOVO"
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
               TabIndex        =   62
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   20
            Text            =   "txtCodigo"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtNome 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   19
            Text            =   "txtNome"
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox txtSenha1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "txtSenha1"
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtSenha2 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   17
            Text            =   "txtSenha1"
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton cmdPesq 
            Height          =   315
            Left            =   1680
            Picture         =   "frmCADSENHA.frx":04BE
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1080
            Width           =   375
         End
         Begin VB.ComboBox cboAcesso 
            Height          =   315
            Left            =   2040
            TabIndex        =   15
            Text            =   "cboAcesso"
            Top             =   1080
            Width           =   4455
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCADSENHA.frx":05C0
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2160
            Width           =   375
         End
         Begin VB.ComboBox cboFUNCAO 
            Height          =   315
            Left            =   2880
            TabIndex        =   13
            Text            =   "cboFUNCAO"
            Top             =   2160
            Width           =   3615
         End
         Begin VB.TextBox txtCODFUNCAO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Text            =   "txtCODFUNCAO"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtCODSETOR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   11
            Text            =   "txtCODSETOR"
            Top             =   2520
            Width           =   855
         End
         Begin VB.ComboBox cboSETOR 
            Height          =   315
            Left            =   2880
            TabIndex        =   10
            Text            =   "cboSETOR"
            Top             =   2520
            Width           =   3615
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCADSENHA.frx":06C2
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2520
            Width           =   375
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCADSENHA.frx":07C4
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2880
            Width           =   375
         End
         Begin VB.ComboBox cboSECAO 
            Height          =   315
            Left            =   2880
            TabIndex        =   7
            Text            =   "cboSECAO"
            Top             =   2880
            Width           =   3615
         End
         Begin VB.TextBox txtCODSECAO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Text            =   "txtCODSECAO"
            Top             =   2880
            Width           =   855
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
            Height          =   195
            Left            =   105
            TabIndex        =   28
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Usuário:"
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
            Left            =   105
            TabIndex        =   27
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Digite a Senha:"
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
            Left            =   105
            TabIndex        =   26
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Confirmar Senha:"
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
            Left            =   105
            TabIndex        =   25
            Top             =   1800
            Width           =   1470
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Acesso:"
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
            Left            =   105
            TabIndex        =   24
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Função"
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
            Left            =   105
            TabIndex        =   23
            Top             =   2160
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Setor"
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
            TabIndex        =   22
            Top             =   2520
            Width           =   465
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Seção"
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
            TabIndex        =   21
            Top             =   2880
            Width           =   555
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
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
         Picture         =   "frmCADSENHA.frx":08C6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   735
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
         Left            =   1560
         Picture         =   "frmCADSENHA.frx":09C8
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   840
         Picture         =   "frmCADSENHA.frx":0ACA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADSENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho   As String
Public Linha      As Variant
Public cTipOper   As String
Public iCodigo    As Integer
Public FILIAL     As Integer
Public strAcesso  As String
Dim objBLBFunc    As Object
Dim objSENHA      As Object
Dim objACESSO     As Object
Dim objPESQPADRAO As Object
Dim arrEMAIL      As Variant

Dim arrMENUPAI    As Variant
Dim arrMENUFILHO  As Variant
Dim arrMENUNETO   As Variant

Const conCOL_SonSenha_Email                     As Integer = 0
Const conCOL_SonSenha_Ativo                     As Integer = 1
Const conCOL_SonSenha_FormatString              As String = "=E-Mail|Ativo"
Const conColumnsIn_SonSenha                     As Integer = 2

Const conCOL_SonMenuPai_Filial                  As Integer = 0
Const conCOL_SonMenuPai_Codigo                  As Integer = 1
Const conCOL_SonMenuPai_Texto                   As Integer = 2
Const conCOL_SonMenuPai_Tipo                    As Integer = 3
Const conCOL_SonMenuPai_CIGLA                   As Integer = 4
Const conCOL_SonMenuPai_Ativo                   As Integer = 5
Const conCOL_SonMenuPai_FormatString            As String = "=Filial|Codigo|Descrição|Tipo|Cigla|Ativo"
Const conColumnsIn_SonMenuPai                   As Integer = 6

Const conCOL_SonMenuFilho_Filial                As Integer = 0
Const conCOL_SonMenuFilho_Codigo                As Integer = 1
Const conCOL_SonMenuFilho_Texto                 As Integer = 2
Const conCOL_SonMenuFilho_Tipo                  As Integer = 3
Const conCOL_SonMenuFilho_CIGLA                 As Integer = 4
Const conCOL_SonMenuFilho_CIGLA2                As Integer = 5
Const conCOL_SonMenuFilho_Ativo                 As Integer = 6
Const conCOL_SonMenuFilho_FormatString          As String = "=Filial|Codigo|Descrição|Tipo|Cigla|Cigla2|Ativo"
Const conColumnsIn_SonMenuFilho                 As Integer = 7

Const conCOL_SonMenuNeto_Filial                As Integer = 0
Const conCOL_SonMenuNeto_Codigo                As Integer = 1
Const conCOL_SonMenuNeto_Texto                 As Integer = 2
Const conCOL_SonMenuNeto_Tipo                  As Integer = 3
Const conCOL_SonMenuNeto_CIGLA                 As Integer = 4
Const conCOL_SonMenuNeto_CIGLA2                As Integer = 5
Const conCOL_SonMenuNeto_MODULO                As Integer = 6
Const conCOL_SonMenuNeto_Ativo                 As Integer = 7
Const conCOL_SonMenuNeto_INCLUIR               As Integer = 8
Const conCOL_SonMenuNeto_ALTERAR               As Integer = 9
Const conCOL_SonMenuNeto_EXCLUIR               As Integer = 10
Const conCOL_SonMenuNeto_CONSULTAR             As Integer = 11
Const conCOL_SonMenuNeto_IMPRIMIR              As Integer = 12
Const conCOL_SonMenuNeto_FormatString          As String = "=Filial|Codigo|Descrição|Tipo|Cigla|Cigla2|Modulo|Ativo|Incluir|Alterar|Excluir|Consultar|Imprimir"
Const conColumnsIn_SonMenuNeto                 As Integer = 13

Private Sub cboAcesso_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboAcesso, KeyAscii
End Sub

Private Sub cboFUNCAO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFUNCAO, KeyAscii
End Sub

Private Sub cboFUNCAO_Validate(Cancel As Boolean)
    If cboFUNCAO.ListIndex > -1 Then txtCODFUNCAO.Text = cboFUNCAO.ItemData(cboFUNCAO.ListIndex)
End Sub

Private Sub cboSECAO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSECAO, KeyAscii
End Sub

Private Sub cboSECAO_Validate(Cancel As Boolean)
    If cboSECAO.ListIndex > -1 Then txtCODSECAO.Text = cboSECAO.ItemData(cboSECAO.ListIndex)
End Sub

Private Sub cboSETOR_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSETOR, KeyAscii
End Sub

Private Sub cboSETOR_Validate(Cancel As Boolean)
    If cboSETOR.ListIndex > -1 Then
       txtCODSETOR.Text = cboSETOR.ItemData(cboSETOR.ListIndex)
       Call objSENHA.PreencheComboSecao(cboSECAO, CLng(txtCODSETOR.Text))
       cboSECAO.ListIndex = -1
       txtCODSECAO.Text = ""
    End If
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Senhas e Usuários - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtNome.SetFocus

End Sub

Private Sub cmdDelGrid_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdEMAIL, grdEMAIL.Row)
End Sub

Private Sub cmdIncGrid_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGrid
End Sub

Private Sub CmdSalva_Click()

    Dim I           As Integer
    Dim strAcesso   As String
    
    If ValidaCampos = True Then
       
       If cTipOper = "I" Then objSENHA.SENCODIGO = objSENHA.Gera_Codigo(Me.Name)
       
       objSENHA.SENNOME = objBLBFunc.Crypt(txtNome.Text)
       objSENHA.SENDEPTO = 0
       objSENHA.SENSENHA = objBLBFunc.Crypt(txtSenha1.Text)
       objSENHA.SENACESSO = cboAcesso.ItemData(cboAcesso.ListIndex)
       
       objSENHA.FUNCAO = cboFUNCAO.ItemData(cboFUNCAO.ListIndex)
       objSENHA.SETOR = cboSETOR.ItemData(cboSETOR.ListIndex)
       objSENHA.SECAO = cboSECAO.ItemData(cboSECAO.ListIndex)
       
       If optPERMBPEDSN(0).Value = True Then objSENHA.PERMBLOQPED = 0
       If optPERMBPEDSN(1).Value = True Then objSENHA.PERMBLOQPED = 1
       
       If optLIBFINSN(0).Value = True Then objSENHA.LIBFINANCEIRO = 0
       If optLIBFINSN(1).Value = True Then objSENHA.LIBFINANCEIRO = 1
       
       If optLIBCOMSN(0).Value = True Then objSENHA.LIBCOMERCIAL = 0
       If optLIBCOMSN(1).Value = True Then objSENHA.LIBCOMERCIAL = 1
       
       If optREPEDSN(0).Value = True Then objSENHA.REPEDSN = 0
       If optREPEDSN(1).Value = True Then objSENHA.REPEDSN = 1
       
       If optLIQPEDSN(0).Value = True Then objSENHA.LIQPEDSN = 0
       If optLIQPEDSN(1).Value = True Then objSENHA.LIQPEDSN = 1
       
       If optLIBPEDBLOQ(0).Value = True Then objSENHA.LIBPEDBLOQSN = 0
       If optLIBPEDBLOQ(1).Value = True Then objSENHA.LIBPEDBLOQSN = 1
        
       If optLIBFOTSN(0).Value = True Then objSENHA.LIBPEDFOTSN = 0
       If optLIBFOTSN(1).Value = True Then objSENHA.LIBPEDFOTSN = 1
        
       If optTIPO(0).Value = True Then objSENHA.NOVO = 0
       If optTIPO(1).Value = True Then objSENHA.NOVO = 1
        
       If optDESABPROD(0).Value = True Then objSENHA.DESABPROD = 0
       If optDESABPROD(1).Value = True Then objSENHA.DESABPROD = 1
        
       If optPERMLIBFOT(0).Value = True Then objSENHA.PERMLIBFOT = 0
       If optPERMLIBFOT(1).Value = True Then objSENHA.PERMLIBFOT = 1
        
       If optPERMFAT10POR(0).Value = True Then objSENHA.PERMFAT10POR = 0
       If optPERMFAT10POR(1).Value = True Then objSENHA.PERMFAT10POR = 1
        
       If optLIBPDATAPCOTA(0).Value = True Then objSENHA.LIBPDATAPCOTA = 0
       If optLIBPDATAPCOTA(1).Value = True Then objSENHA.LIBPDATAPCOTA = 1
        
       If optPermFatRotDifSN(0).Value = True Then objSENHA.PERMFATROTDIFSN = 0
       If optPermFatRotDifSN(1).Value = True Then objSENHA.PERMFATROTDIFSN = 1
            
        If optPVCLIE(0).Value = True Then objSENHA.PVCLIE = 0
        If optPVCLIE(1).Value = True Then objSENHA.PVCLIE = 1
        
        If optEVENDEDOR(0).Value = True Then objSENHA.EVENDEDOR = 0
        If optEVENDEDOR(1).Value = True Then objSENHA.EVENDEDOR = 1
        
        If optAtivo(0).Value = True Then objSENHA.ATIVO = 0
        If optAtivo(1).Value = True Then objSENHA.ATIVO = 1
        
        If optLimCred(0).Value = True Then objSENHA.BLOQCRED = 0
        If optLimCred(1).Value = True Then objSENHA.BLOQCRED = 1
        
        If optPERMALTPEDFAT(0).Value = True Then objSENHA.PERMALTPEDFAT = 0
        If optPERMALTPEDFAT(1).Value = True Then objSENHA.PERMALTPEDFAT = 1
        
        If optMOPSN(0).Value = True Then objSENHA.MOP = 0
        If optMOPSN(1).Value = True Then objSENHA.MOP = 1
        
        If optPermExcPedSN(0).Value = True Then objSENHA.PermExcPedSN = 0
        If optPermExcPedSN(1).Value = True Then objSENHA.PermExcPedSN = 1
        
        arrEMAIL = Empty
        If (grdEMAIL.Rows - 1) > 0 Then
            ReDim arrEMAIL(1 To (grdEMAIL.Rows - 1), 1 To 2) As Variant
            For I = 1 To (grdEMAIL.Rows - 1)
                arrEMAIL(I, 1) = grdEMAIL.Cell(flexcpText, I, conCOL_SonSenha_Email)
                arrEMAIL(I, 2) = IIf(grdEMAIL.Cell(flexcpTextDisplay, I, conCOL_SonSenha_Ativo) = "Não", 0, 1)
            Next I
        End If
        objSENHA.EMAIL = arrEMAIL
       
        '' ========================
        '' Menu Pai
        arrMENUPAI = Empty
        With grdMENUPAI
            If (.Rows - 1) > 0 Then
                ReDim arrMENUPAI(1 To (.Rows - 1), 1 To .Cols) As String
                For I = 1 To (.Rows - 1)
                    arrMENUPAI(I, (conCOL_SonMenuPai_Filial + 1)) = .Cell(flexcpText, I, conCOL_SonMenuPai_Filial)
                    arrMENUPAI(I, (conCOL_SonMenuPai_Codigo + 1)) = .Cell(flexcpText, I, conCOL_SonMenuPai_Codigo)
                    arrMENUPAI(I, (conCOL_SonMenuPai_Texto + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuPai_Texto) & "'"
                    arrMENUPAI(I, (conCOL_SonMenuPai_Tipo + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuPai_Tipo) & "'"
                    arrMENUPAI(I, (conCOL_SonMenuPai_CIGLA + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuPai_CIGLA) & "'"
                    arrMENUPAI(I, (conCOL_SonMenuPai_Ativo + 1)) = .Cell(flexcpText, I, conCOL_SonMenuPai_Ativo)
                Next I
            End If
        End With
        objSENHA.MENUPAI = arrMENUPAI
       
        '' ========================
        '' Menu Filho
        arrMENUFILHO = Empty
        With grdMENUFILHO
            If (.Rows - 1) > 0 Then
                ReDim arrMENUFILHO(1 To (.Rows - 1), 1 To .Cols) As String
                For I = 1 To (.Rows - 1)
                    arrMENUFILHO(I, (conCOL_SonMenuFilho_Filial + 1)) = .Cell(flexcpText, I, conCOL_SonMenuFilho_Filial)
                    arrMENUFILHO(I, (conCOL_SonMenuFilho_Codigo + 1)) = .Cell(flexcpText, I, conCOL_SonMenuFilho_Codigo)
                    arrMENUFILHO(I, (conCOL_SonMenuFilho_Texto + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuFilho_Texto) & "'"
                    arrMENUFILHO(I, (conCOL_SonMenuFilho_Tipo + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuFilho_Tipo) & "'"
                    arrMENUFILHO(I, (conCOL_SonMenuFilho_CIGLA + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuFilho_CIGLA) & "'"
                    arrMENUFILHO(I, (conCOL_SonMenuFilho_CIGLA2 + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuFilho_CIGLA2) & "'"
                    arrMENUFILHO(I, (conCOL_SonMenuFilho_Ativo + 1)) = .Cell(flexcpText, I, conCOL_SonMenuFilho_Ativo)
                Next I
            End If
        End With
        objSENHA.MENUFILHO = arrMENUFILHO
        
        '' ========================
        '' Menu Neto
        arrMENUNETO = Empty
        strAcesso = ""
        With grdMENUNETO
            If (.Rows - 1) > 0 Then
                ReDim arrMENUNETO(1 To (.Rows - 1), 1 To 9) As String
                For I = 1 To (.Rows - 1)
                    arrMENUNETO(I, (conCOL_SonMenuNeto_Filial + 1)) = .Cell(flexcpText, I, conCOL_SonMenuNeto_Filial)
                    arrMENUNETO(I, (conCOL_SonMenuNeto_Codigo + 1)) = .Cell(flexcpText, I, conCOL_SonMenuNeto_Codigo)
                    arrMENUNETO(I, (conCOL_SonMenuNeto_Texto + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuNeto_Texto) & "'"
                    arrMENUNETO(I, (conCOL_SonMenuNeto_Tipo + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuNeto_Tipo) & "'"
                    arrMENUNETO(I, (conCOL_SonMenuNeto_CIGLA + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuNeto_CIGLA) & "'"
                    arrMENUNETO(I, (conCOL_SonMenuNeto_CIGLA2 + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuNeto_CIGLA2) & "'"
                    arrMENUNETO(I, (conCOL_SonMenuNeto_MODULO + 1)) = "'" & .Cell(flexcpText, I, conCOL_SonMenuNeto_MODULO) & "'"
                    arrMENUNETO(I, (conCOL_SonMenuNeto_Ativo + 1)) = .Cell(flexcpText, I, conCOL_SonMenuNeto_Ativo)
                    
                    strAcesso = ""
                    If .Cell(flexcpText, I, conCOL_SonMenuNeto_INCLUIR) = 1 Then strAcesso = "I"
                    If .Cell(flexcpText, I, conCOL_SonMenuNeto_ALTERAR) = 1 Then strAcesso = strAcesso & "A"
                    If .Cell(flexcpText, I, conCOL_SonMenuNeto_EXCLUIR) = 1 Then strAcesso = strAcesso & "E"
                    If .Cell(flexcpText, I, conCOL_SonMenuNeto_CONSULTAR) = 1 Then strAcesso = strAcesso & "C"
                    If .Cell(flexcpText, I, conCOL_SonMenuNeto_IMPRIMIR) = 1 Then strAcesso = strAcesso & "R"
                    
                    arrMENUNETO(I, 9) = "'" & strAcesso & "'"
                
                Next I
            End If
        End With
        objSENHA.MENUNETO = arrMENUNETO
        
        If objSENHA.GRAVA(cTipOper) = True Then
           
           MsgBox "O Usuário foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
           
           If cTipOper = "I" Then
              Set objBLBFunc = Nothing
              Set objSENHA = Nothing
              Unload Me
           End If
           
        End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objSENHA = Nothing
    Set objACESSO = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFUNCAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Função")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFUNCAO.Text = varRETORNO
        
    cboFUNCAO.ListIndex = -1
    txtCODFUNCAO.SetFocus

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADSETOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Setor")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODSETOR.Text = varRETORNO
        
    cboSETOR.ListIndex = -1
    txtCODSETOR.SetFocus

End Sub

Private Sub Command3_Click()

    If cboSETOR.ListIndex = -1 Then Exit Sub

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "      SE.*  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "      SGI_CADITESET IT " & vbCrLf
    sSql = sSql & "     ,SGI_CADSECAO  SE " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      IT.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "  And IT.SGI_CODIGO   = " & txtCODSETOR.Text & vbCrLf
    sSql = sSql & "  And SE.SGI_FILIAL   = IT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SE.SGI_CODIGO   = IT.SGI_CODSECAO "
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SE.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SE.SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Seção")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODSECAO.Text = varRETORNO
        
    cboSECAO.ListIndex = -1
    txtCODSECAO.SetFocus

End Sub

Private Sub Command4_Click()
    Call PopGrdMenuPai
    Call PopGrdMenuFilho
    Call PopGrdMenuNeto
    grdMENUPAI.Row = 1
    grdMENUFILHO.Row = 1
End Sub

Private Sub Command5_Click()

    Call InitGridMenuPai
    Call InitGridMenuFilho
    Call InitGridMenuNeto

    Call PopGrdMenuPai
    Call PopGrdMenuFilho
    Call PopGrdMenuNeto

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objSENHA = CreateObject("CADSENHA.clsCADSENHA")
   ''Set objACESSO = CreateObject("CADMENU.clsCADMENU")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objSENHA.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta
   
   objBLBFunc.ChecaAcesso frmCADSENHA, strAcesso
   
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Senhas e Usuários - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSENHA
    objSENHA.PreencheComboAcesso cboAcesso
    objSENHA.PreencheComboSetor cboSETOR
    objSENHA.PreencheComboFuncao cboFUNCAO
    
    txtCodigo.Text = ""
    optPERMBPEDSN(0).Value = True
    optLIBFINSN(0).Value = True
    optLIBCOMSN(0).Value = True
    optREPEDSN(0).Value = True
    optLIQPEDSN(0).Value = True
    optLIBPEDBLOQ(0).Value = True
    optLIBFOTSN(0).Value = True
    optTIPO(0).Value = True
    optDESABPROD(0).Value = True
    optPERMLIBFOT(0).Value = True
    optPERMFAT10POR(0).Value = True
    optPVCLIE(0).Value = True
    optEVENDEDOR(0).Value = True
    optAtivo(1).Value = True
    optLIBPDATAPCOTA(0).Value = True
    optLimCred(0).Value = True
    optPERMALTPEDFAT(0).Value = True
    optMOPSN(0).Value = True
    optPermExcPedSN(1).Value = True
    
    Call InitGridEMail
    Call InitGridMenuPai
    Call InitGridMenuFilho
    Call InitGridMenuNeto
   
    Call PopGrdMenuPai
    Call PopGrdMenuFilho
    Call PopGrdMenuNeto
    grdMENUPAI.Row = 1
    grdMENUFILHO.Row = 1

    optPermFatRotDifSN(0).Value = True

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Senhas e Usuários - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSENHA
    objSENHA.PreencheComboAcesso cboAcesso
    objSENHA.PreencheComboSetor cboSETOR
    objSENHA.PreencheComboFuncao cboFUNCAO
    
    objSENHA.SENCODIGO = iCodigo
    optPERMBPEDSN(0).Value = True
    optLIBFINSN(0).Value = True
    optLIBCOMSN(0).Value = True
    optREPEDSN(0).Value = True
    optLIQPEDSN(0).Value = True
    optLIBPEDBLOQ(0).Value = True
    optLIBFOTSN(0).Value = True
    optTIPO(0).Value = True
    optDESABPROD(0).Value = True
    optPERMLIBFOT(0).Value = True
    optPERMFAT10POR(0).Value = True
    optLIBPDATAPCOTA(0).Value = True
    optPermFatRotDifSN(0).Value = True
    optPVCLIE(0).Value = True
    optEVENDEDOR(0).Value = True
    optAtivo(1).Value = True
    optLimCred(0).Value = True
    optPERMALTPEDFAT(0).Value = True
    optMOPSN(0).Value = True
    optPermExcPedSN(1).Value = True
    
    Call InitGridEMail
    Call InitGridMenuPai
    Call InitGridMenuFilho
    Call InitGridMenuNeto
    
    If objSENHA.Carrega_campos = True Then
    
        arrMENUPAI = objSENHA.MENUPAI
        arrMENUFILHO = objSENHA.MENUFILHO
        arrMENUNETO = objSENHA.MENUNETO
        
        txtCodigo.Text = Str(objSENHA.SENCODIGO)
        txtNome.Text = objBLBFunc.Crypt(objSENHA.SENNOME)
        txtSenha1.Text = objBLBFunc.Crypt(objSENHA.SENSENHA)
        txtSenha2.Text = objBLBFunc.Crypt(objSENHA.SENSENHA)
       
        For I = 0 To (cboAcesso.ListCount - 1)
            If cboAcesso.ItemData(I) = objSENHA.SENACESSO Then cboAcesso.ListIndex = I
        Next I
       
        txtCODFUNCAO.Text = objSENHA.FUNCAO
        For I = 0 To (cboFUNCAO.ListCount - 1)
            If cboFUNCAO.ItemData(I) = objSENHA.FUNCAO Then cboFUNCAO.ListIndex = I
        Next I
       
        txtCODSETOR.Text = objSENHA.SETOR
        For I = 0 To (cboSETOR.ListCount - 1)
            If cboSETOR.ItemData(I) = objSENHA.SETOR Then cboSETOR.ListIndex = I
        Next I
       
        Call objSENHA.PreencheComboSecao(cboSECAO, objSENHA.SETOR)
        txtCODSECAO.Text = objSENHA.SECAO
        For I = 0 To (cboSECAO.ListCount - 1)
            If cboSECAO.ItemData(I) = objSENHA.SECAO Then cboSECAO.ListIndex = I
        Next I
    
        Call PopGrdEmail
        Call PopGrdMenuPai
        Call PopGrdMenuFilho
        Call PopGrdMenuNeto
        grdMENUPAI.Row = 1
        grdMENUFILHO.Row = 1
       
        optPERMBPEDSN(objSENHA.PERMBLOQPED).Value = True
        optLIBFINSN(objSENHA.LIBFINANCEIRO).Value = True
        optLIBCOMSN(objSENHA.LIBCOMERCIAL).Value = True
        optREPEDSN(objSENHA.REPEDSN).Value = True
        optLIQPEDSN(objSENHA.LIQPEDSN).Value = True
        optLIBPEDBLOQ(objSENHA.LIBPEDBLOQSN).Value = True
        optLIBFOTSN(objSENHA.LIBPEDFOTSN).Value = True
        optTIPO(objSENHA.NOVO).Value = True
        optPERMLIBFOT(objSENHA.PERMLIBFOT).Value = True
        optPERMFAT10POR(objSENHA.PERMFAT10POR).Value = True
        optLIBPDATAPCOTA(objSENHA.LIBPDATAPCOTA).Value = True
        optPermFatRotDifSN(objSENHA.PERMFATROTDIFSN).Value = True
        optPVCLIE(objSENHA.PVCLIE).Value = True
        optEVENDEDOR(objSENHA.EVENDEDOR).Value = True
        optAtivo(objSENHA.ATIVO).Value = True
        optLimCred(objSENHA.BLOQCRED).Value = True
        optPERMALTPEDFAT(objSENHA.PERMALTPEDFAT).Value = True
        optMOPSN(objSENHA.MOP).Value = True
        optPermExcPedSN(objSENHA.PermExcPedSN).Value = True

        Call CarregaGrdMenuPai
        Call CarregaGrdMenuFilho
        Call CarregaGrdMenuNeto
    
    End If
    
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de Senhas e Usuários - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADSENHA
    objSENHA.PreencheComboAcesso cboAcesso
    objSENHA.PreencheComboSetor cboSETOR
    objSENHA.PreencheComboFuncao cboFUNCAO
    
    objSENHA.SENCODIGO = iCodigo
    optPERMBPEDSN(0).Value = True
    optLIBFINSN(0).Value = True
    optLIBCOMSN(0).Value = True
    optREPEDSN(0).Value = True
    optLIQPEDSN(0).Value = True
    optLIBPEDBLOQ(0).Value = True
    optLIBFOTSN(0).Value = True
    optTIPO(0).Value = True
    optDESABPROD(0).Value = True
    optPERMLIBFOT(0).Value = True
    optPERMFAT10POR(0).Value = True
    optLIBPDATAPCOTA(0).Value = True
    optPermFatRotDifSN(0).Value = True
    optPVCLIE(0).Value = True
    optEVENDEDOR(0).Value = True
    optAtivo(1).Value = True
    optLimCred(0).Value = True
    optPERMALTPEDFAT(0).Value = True
    optMOPSN(0).Value = True
    optPermExcPedSN(1).Value = True
    
    Call InitGridEMail
    Call InitGridMenuPai
    Call InitGridMenuFilho
    Call InitGridMenuNeto
    
    If objSENHA.Carrega_campos = True Then
    
        arrMENUPAI = objSENHA.MENUPAI
        arrMENUFILHO = objSENHA.MENUFILHO
        arrMENUNETO = objSENHA.MENUNETO
        
        txtCodigo.Text = Str(objSENHA.SENCODIGO)
        txtNome.Text = objBLBFunc.Crypt(objSENHA.SENNOME)
        txtSenha1.Text = objBLBFunc.Crypt(objSENHA.SENSENHA)
        txtSenha2.Text = objBLBFunc.Crypt(objSENHA.SENSENHA)
        
        For I = 0 To (cboAcesso.ListCount - 1)
            If cboAcesso.ItemData(I) = objSENHA.SENACESSO Then cboAcesso.ListIndex = I
        Next I
        
        txtCODFUNCAO.Text = objSENHA.FUNCAO
        For I = 0 To (cboFUNCAO.ListCount - 1)
            If cboFUNCAO.ItemData(I) = objSENHA.FUNCAO Then cboFUNCAO.ListIndex = I
        Next I
        
        txtCODSETOR.Text = objSENHA.SETOR
        For I = 0 To (cboSETOR.ListCount - 1)
            If cboSETOR.ItemData(I) = objSENHA.SETOR Then cboSETOR.ListIndex = I
        Next I
        
        Call objSENHA.PreencheComboSecao(cboSECAO, objSENHA.SETOR)
        txtCODSECAO.Text = objSENHA.SECAO
        For I = 0 To (cboSECAO.ListCount - 1)
            If cboSECAO.ItemData(I) = objSENHA.SECAO Then cboSECAO.ListIndex = I
        Next I
       
        Call PopGrdEmail
        Call PopGrdMenuPai
        Call PopGrdMenuFilho
        Call PopGrdMenuNeto
        grdMENUPAI.Row = 1
        grdMENUFILHO.Row = 1
       
        optPERMBPEDSN(objSENHA.PERMBLOQPED).Value = True
        optLIBFINSN(objSENHA.LIBFINANCEIRO).Value = True
        optLIBCOMSN(objSENHA.LIBCOMERCIAL).Value = True
        optREPEDSN(objSENHA.REPEDSN).Value = True
        optLIQPEDSN(objSENHA.LIQPEDSN).Value = True
        optLIBPEDBLOQ(objSENHA.LIBPEDBLOQSN).Value = True
        optLIBFOTSN(objSENHA.LIBPEDFOTSN).Value = True
        optTIPO(objSENHA.NOVO).Value = True
        optDESABPROD(objSENHA.DESABPROD).Value = True
        optPERMLIBFOT(objSENHA.PERMLIBFOT).Value = True
        optPERMFAT10POR(objSENHA.PERMFAT10POR).Value = True
        optLIBPDATAPCOTA(objSENHA.LIBPDATAPCOTA).Value = True
        optPermFatRotDifSN(objSENHA.PERMFATROTDIFSN).Value = True
        optPVCLIE(objSENHA.PVCLIE).Value = True
        optEVENDEDOR(objSENHA.EVENDEDOR).Value = True
        optAtivo(objSENHA.ATIVO).Value = True
        optLimCred(objSENHA.BLOQCRED).Value = True
        optPERMALTPEDFAT(objSENHA.PERMALTPEDFAT).Value = True
        optMOPSN(objSENHA.MOP).Value = True
        optPermExcPedSN(objSENHA.PermExcPedSN).Value = True

        Call CarregaGrdMenuPai
        Call CarregaGrdMenuFilho
        Call CarregaGrdMenuNeto
    
    End If

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtNome.Text)) = 0 Then
        MsgBox "Nome do usuário inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtNome.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtSenha1.Text)) = 0 Then
        MsgBox "Primeira senha inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtSenha1.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtSenha2.Text)) = 0 Then
        MsgBox "Segunda senha inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtSenha2.SetFocus
        Exit Function
     End If
     
     If Trim(txtSenha2.Text) <> Trim(txtSenha1.Text) Then
        MsgBox "Confirmação da senha inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtSenha2.SetFocus
        Exit Function
     End If
     
     If cboAcesso.ListIndex = -1 Then
        MsgBox "Nivel de acesso inválido !!!", vbOKOnly + vbCritical, "Aviso"
        cboAcesso.ListIndex = 0
        cboAcesso.SetFocus
        Exit Function
     End If
     
     If cboFUNCAO.ListIndex = -1 Then
        MsgBox "A função deve ser informada !!!", vbOKOnly + vbCritical, "Aviso"
        cboFUNCAO.ListIndex = -1
        cboFUNCAO.SetFocus
        Exit Function
     End If
     
     If cboSETOR.ListIndex = -1 Then
        MsgBox "O setor deve ser informado !!!", vbOKOnly + vbCritical, "Aviso"
        cboSETOR.ListIndex = -1
        cboSETOR.SetFocus
        Exit Function
     End If
     
     If cboSECAO.ListIndex = -1 Then
        MsgBox "A seção deve ser informada !!!", vbOKOnly + vbCritical, "Aviso"
        cboSECAO.ListIndex = -1
        cboSECAO.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select * from SGI_USUARIO Where SGI_NOME ='" & txtNome.Text & "'"
        sSql = sSql & " And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Usuário já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtNome.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objSENHA.SENNOME <> txtNome.Text Then
        
           sSql = "Select * from SGI_USUARIO Where SGI_NOME ='" & txtNome.Text & "'"
           sSql = sSql & " And SGI_FILIAL = " & FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Usuário já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtNome.Text = objSENHA.SENNOME
              txtNome.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub grdEMAIL_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonSenha_Email, _
         conCOL_SonSenha_Ativo
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdEMAIL.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdEMAIL_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdEMAIL
          Select Case Col
                 Case conCOL_SonSenha_Email
                        If .EditText = Empty Then Exit Sub
                        If VerifItensRepetidos(Row, conCOL_SonSenha_Email, .EditText) = False Then
                           MsgBox "Este E-Mial ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdEMAIL.Cell(flexcpText, Row, conCOL_SonSenha_Email) = ""
                           Cancel = True
                           Exit Sub
                        End If
          End Select
     End With
End Sub

Private Sub grdMENUFILHO_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Select Case Col
    Case conCOL_SonMenuFilho_Ativo
         Call AtivoSNFilho
    End Select
    
    Exit Sub

End Sub

Private Sub grdMENUFILHO_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Select Case Col
    Case conCOL_SonMenuFilho_Texto
         Cancel = True
    Case conCOL_SonMenuFilho_Ativo
         If cTipOper = "C" Then
            Cancel = True
            Exit Sub
         End If
         If grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_Ativo) = 0 Then
            Cancel = True
            Exit Sub
         End If
    Case Else
        grdMENUFILHO.ComboList = ""
    End Select
    
    Exit Sub

End Sub

Private Sub grdMENUFILHO_Click()
    Call MostrDadosNeto
    
End Sub

Private Sub grdMENUFILHO_RowColChange()
    Call MostrDadosNeto
End Sub

Private Sub grdMENUNETO_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Select Case Col
    Case conCOL_SonMenuNeto_Texto
         Cancel = True
    Case conCOL_SonMenuNeto_Ativo
         If cTipOper = "C" Then
            Cancel = True
            Exit Sub
         End If
         If grdMENUFILHO.Cell(flexcpText, grdMENUFILHO.Row, conCOL_SonMenuFilho_Ativo) = 0 Then
            Cancel = True
            Exit Sub
         End If
    Case Else
        grdMENUNETO.ComboList = ""
    End Select
    
    Exit Sub

End Sub

Private Sub grdMENUPAI_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Select Case Col
    Case conCOL_SonMenuPai_Ativo
            Call AtivoSN
            If grdMENUPAI.Cell(flexcpText, Row, conCOL_SonMenuPai_CIGLA) = "G" Then Call AtivoSNFilho2
    End Select
    
    Exit Sub

End Sub

Private Sub grdMENUPAI_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Select Case Col
    Case conCOL_SonMenuPai_Texto
         Cancel = True
    Case conCOL_SonMenuPai_Ativo
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdMENUPAI.ComboList = ""
    End Select
    
    Exit Sub

End Sub

Private Sub grdMENUPAI_Click()
    Call MostrDados
    Call MostrDadosNeto
End Sub


Private Sub grdMENUPAI_RowColChange()
    Call MostrDados
    Call MostrDadosNeto
End Sub


Private Sub txtCODFUNCAO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFUNCAO.Name, frmCADSENHA
End Sub

Private Sub txtCODFUNCAO_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtCODFUNCAO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFUNCAO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODFUNCAO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    For I = 0 To (cboFUNCAO.ListCount - 1)
        If CInt(txtCODFUNCAO.Text) = cboFUNCAO.ItemData(I) Then cboFUNCAO.ListIndex = I
    Next I
    
    If cboFUNCAO.ListIndex = -1 Then
       MsgBox "Está função não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCODFUNCAO.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODSECAO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODSECAO.Name, frmCADSENHA
End Sub

Private Sub txtCODSECAO_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtCODSECAO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODSECAO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSECAO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    For I = 0 To (cboSECAO.ListCount - 1)
        If CInt(txtCODSECAO.Text) = cboSECAO.ItemData(I) Then cboSECAO.ListIndex = I
    Next I
    
    If cboSECAO.ListIndex = -1 Then
       MsgBox "Esta seção não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSECAO.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODSETOR_GotFocus()
    objBLBFunc.SelecionaCampos txtCODSETOR.Name, frmCADSENHA
End Sub

Private Sub txtCODSETOR_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtCODSETOR.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODSETOR.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSETOR.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    For I = 0 To (cboSETOR.ListCount - 1)
        If CInt(txtCODSETOR.Text) = cboSETOR.ItemData(I) Then cboSETOR.ListIndex = I
    Next I
    
    If cboSETOR.ListIndex = -1 Then
       MsgBox "Esta setor não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSETOR.Text = ""
       Cancel = True
       Exit Sub
    End If

    Call objSENHA.PreencheComboSecao(cboSECAO, CLng(txtCODSETOR.Text))
    cboSECAO.ListIndex = -1
    txtCODSECAO.Text = ""
    
End Sub

Private Sub txtNome_GotFocus()
    objBLBFunc.SelecionaCampos txtNome.Name, frmCADSENHA
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtSenha1_GotFocus()
    objBLBFunc.SelecionaCampos txtSenha1.Name, frmCADSENHA
End Sub

Private Sub txtSenha1_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtSenha2_GotFocus()
    objBLBFunc.SelecionaCampos txtSenha2.Name, frmCADSENHA
End Sub

Private Sub txtSenha2_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub InitGridEMail()

    With grdEMAIL
    
       .Cols = conColumnsIn_SonSenha
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonSenha_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonSenha_Email) = ""
       .ColDataType(conCOL_SonSenha_Email) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonSenha_Ativo) = ""
       .ColDataType(conCOL_SonSenha_Ativo) = flexDTBoolean
       .ColFormat(conCOL_SonSenha_Ativo) = "Sim;Não"
       
       .ColWidth(conCOL_SonSenha_Email) = 4500
       .ColWidth(conCOL_SonSenha_Ativo) = 1200
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncRegGrid()
   
    If ExisteLinhaVazia = False Then Exit Sub
    
    grdEMAIL.AddItem "" & vbTab & _
                     ""
                            
End Sub


Private Function ExisteLinhaVazia() As Boolean
    ExisteLinhaVazia = False
    
    Dim I As Integer
    
    For I = 1 To (grdEMAIL.Rows - 1)
        If grdEMAIL.Cell(flexcpText, I, conCOL_SonSenha_Email) = Empty Then Exit Function
    Next I
    
    ExisteLinhaVazia = True
End Function


Private Function VerifItensRepetidos(intRow As Long, intCol As Long, varCampo As Variant) As Boolean
    VerifItensRepetidos = False
    Dim I As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    For I = 1 To (grdEMAIL.Rows - 1)
        If I <> intRow And UCase(grdEMAIL.Cell(flexcpText, I, intCol)) = varCampo Then Exit Function
    Next I
    VerifItensRepetidos = True
End Function


Private Sub PopGrdEmail()
    Dim I As Integer
    arrEMAIL = objSENHA.EMAIL
    If IsArray(arrEMAIL) Then
       For I = 1 To UBound(arrEMAIL)
           grdEMAIL.AddItem arrEMAIL(I, 1) & vbTab & arrEMAIL(I, 2)
       Next I
    End If
End Sub

Private Sub InitGridMenuPai()

    With grdMENUPAI
    
       .Cols = conColumnsIn_SonMenuPai
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMenuPai_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMenuPai_Filial) = ""
       .ColDataType(conCOL_SonMenuPai_Filial) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMenuPai_Codigo) = ""
       .ColDataType(conCOL_SonMenuPai_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMenuPai_Texto) = ""
       .ColDataType(conCOL_SonMenuPai_Texto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuPai_Tipo) = ""
       .ColDataType(conCOL_SonMenuPai_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuPai_CIGLA) = ""
       .ColDataType(conCOL_SonMenuPai_CIGLA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuPai_Ativo) = ""
       .ColDataType(conCOL_SonMenuPai_Ativo) = flexDTString
       .ColComboList(conCOL_SonMenuPai_Ativo) = objSENHA.PreenchComboAtivo
       
       .ColWidth(conCOL_SonMenuPai_Filial) = 0
       .ColWidth(conCOL_SonMenuPai_Codigo) = 0
       .ColWidth(conCOL_SonMenuPai_Texto) = 3000
       .ColWidth(conCOL_SonMenuPai_Tipo) = 0
       .ColWidth(conCOL_SonMenuPai_CIGLA) = 0
       .ColWidth(conCOL_SonMenuPai_Ativo) = 700
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PopGrdMenuPai()
    
    Call InitGridMenuPai
    
    Dim intATIVO    As Integer
    
    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_MENUP" & vbCrLf
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_FILIAL     = 0" & vbCrLf
    sSql = sSql & "  And  SGI_TIPO       = 'P'" & vbCrLf
    sSql = sSql & "Order By SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        With grdMENUPAI
            Do While Not BREC.EOF()
                
                .AddItem FILIAL & vbTab & _
                         BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_TEXTO & vbTab & _
                         BREC!SGI_TIPO & vbTab & _
                         BREC!SGI_CIGLA & vbTab & _
                         BREC!SGI_ATIVO
                
                BREC.MoveNext
            Loop
        End With
    End If
    BREC.Close
    
End Sub
 

Private Sub InitGridMenuFilho()

    With grdMENUFILHO
    
       .Cols = conColumnsIn_SonMenuFilho
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMenuFilho_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMenuFilho_Filial) = ""
       .ColDataType(conCOL_SonMenuFilho_Filial) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMenuFilho_Codigo) = ""
       .ColDataType(conCOL_SonMenuFilho_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMenuFilho_Texto) = ""
       .ColDataType(conCOL_SonMenuFilho_Texto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuFilho_Tipo) = ""
       .ColDataType(conCOL_SonMenuFilho_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuFilho_CIGLA) = ""
       .ColDataType(conCOL_SonMenuFilho_CIGLA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuFilho_CIGLA2) = ""
       .ColDataType(conCOL_SonMenuFilho_CIGLA2) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuFilho_Ativo) = ""
       .ColDataType(conCOL_SonMenuFilho_Ativo) = flexDTString
       .ColComboList(conCOL_SonMenuFilho_Ativo) = objSENHA.PreenchComboAtivo
       
       .ColWidth(conCOL_SonMenuFilho_Filial) = 0
       .ColWidth(conCOL_SonMenuFilho_Codigo) = 0
       .ColWidth(conCOL_SonMenuFilho_Texto) = 2000
       .ColWidth(conCOL_SonMenuFilho_Tipo) = 0
       .ColWidth(conCOL_SonMenuFilho_CIGLA) = 0
       .ColWidth(conCOL_SonMenuFilho_CIGLA2) = 0
       .ColWidth(conCOL_SonMenuFilho_Ativo) = 700
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub



Private Sub PopGrdMenuFilho()
    
    Call InitGridMenuFilho
    
    Dim intATIVO    As Integer
    
    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_MENUP" & vbCrLf
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_FILIAL     = 0" & vbCrLf
    sSql = sSql & "  And  SGI_TIPO       = 'S'" & vbCrLf
    sSql = sSql & "Order By SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        With grdMENUFILHO
            Do While Not BREC.EOF()
                
                .AddItem FILIAL & vbTab & _
                         BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_TEXTO & vbTab & _
                         BREC!SGI_TIPO & vbTab & _
                         BREC!SGI_CIGLA & vbTab & _
                         BREC!SGI_CIGLA2 & vbTab & _
                         BREC!SGI_ATIVO
                
                BREC.MoveNext
            Loop
        End With
    End If
    BREC.Close
    
End Sub

Private Sub MostrDados()
    
    If (grdMENUPAI.Rows - 1) = 0 Then Exit Sub
    If grdMENUPAI.Row = 0 Then Exit Sub
    
    Frame10.Caption = "Menu [ " & grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_Texto) & " ]"
    Call objBLBFunc.CarregaDadosGrdFilho(grdMENUFILHO, dacEnumUpdateAction_Ignore, conCOL_SonMenuFilho_CIGLA, grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_CIGLA))

    Dim lngLINHA As Long
    lngLINHA = -1
    lngLINHA = grdMENUFILHO.FindRow(grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_CIGLA), , conCOL_SonMenuFilho_CIGLA)
    If lngLINHA > -1 Then grdMENUFILHO.Row = lngLINHA

End Sub


Private Sub InitGridMenuNeto()

    With grdMENUNETO
    
       .Cols = conColumnsIn_SonMenuNeto
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMenuNeto_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_Filial) = ""
       .ColDataType(conCOL_SonMenuNeto_Filial) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_Codigo) = ""
       .ColDataType(conCOL_SonMenuNeto_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_Texto) = ""
       .ColDataType(conCOL_SonMenuNeto_Texto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_Tipo) = ""
       .ColDataType(conCOL_SonMenuNeto_Tipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_CIGLA) = ""
       .ColDataType(conCOL_SonMenuNeto_CIGLA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_CIGLA2) = ""
       .ColDataType(conCOL_SonMenuNeto_CIGLA2) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_MODULO) = ""
       .ColDataType(conCOL_SonMenuNeto_MODULO) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_Ativo) = ""
       .ColDataType(conCOL_SonMenuNeto_Ativo) = flexDTString
       .ColComboList(conCOL_SonMenuNeto_Ativo) = objSENHA.PreenchComboAtivo
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_INCLUIR) = ""
       .ColDataType(conCOL_SonMenuNeto_INCLUIR) = flexDTString
       .ColComboList(conCOL_SonMenuNeto_INCLUIR) = objSENHA.PreenchComboAtivo
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_ALTERAR) = ""
       .ColDataType(conCOL_SonMenuNeto_ALTERAR) = flexDTString
       .ColComboList(conCOL_SonMenuNeto_ALTERAR) = objSENHA.PreenchComboAtivo
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_EXCLUIR) = ""
       .ColDataType(conCOL_SonMenuNeto_EXCLUIR) = flexDTString
       .ColComboList(conCOL_SonMenuNeto_EXCLUIR) = objSENHA.PreenchComboAtivo
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_CONSULTAR) = ""
       .ColDataType(conCOL_SonMenuNeto_CONSULTAR) = flexDTString
       .ColComboList(conCOL_SonMenuNeto_CONSULTAR) = objSENHA.PreenchComboAtivo
       
       .Cell(flexcpData, 0, conCOL_SonMenuNeto_IMPRIMIR) = ""
       .ColDataType(conCOL_SonMenuNeto_IMPRIMIR) = flexDTString
       .ColComboList(conCOL_SonMenuNeto_IMPRIMIR) = objSENHA.PreenchComboAtivo
       
       .ColWidth(conCOL_SonMenuNeto_Filial) = 0
       .ColWidth(conCOL_SonMenuNeto_Codigo) = 0
       .ColWidth(conCOL_SonMenuNeto_Texto) = 4000
       .ColWidth(conCOL_SonMenuNeto_Tipo) = 0
       .ColWidth(conCOL_SonMenuNeto_CIGLA) = 0
       .ColWidth(conCOL_SonMenuNeto_CIGLA2) = 0
       .ColWidth(conCOL_SonMenuNeto_MODULO) = 0
       .ColWidth(conCOL_SonMenuNeto_Ativo) = 600
       .ColWidth(conCOL_SonMenuNeto_INCLUIR) = 600
       .ColWidth(conCOL_SonMenuNeto_ALTERAR) = 600
       .ColWidth(conCOL_SonMenuNeto_EXCLUIR) = 600
       .ColWidth(conCOL_SonMenuNeto_CONSULTAR) = 800
       .ColWidth(conCOL_SonMenuNeto_IMPRIMIR) = 800
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PopGrdMenuNeto()
    
    Call InitGridMenuNeto
    
    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_MENUP" & vbCrLf
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_FILIAL     = 0" & vbCrLf
    sSql = sSql & "  And  SGI_TIPO       = 'M'" & vbCrLf
    sSql = sSql & "Order By SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        With grdMENUNETO
            Do While Not BREC.EOF()
                
                .AddItem FILIAL & vbTab & _
                         BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_TEXTO & vbTab & _
                         BREC!SGI_TIPO & vbTab & _
                         BREC!SGI_CIGLA & vbTab & _
                         BREC!SGI_CIGLA2 & vbTab & _
                         BREC!SGI_MODULO & vbTab & _
                         IIf(IsNull(BREC!SGI_ATIVO) = True, 1, BREC!SGI_ATIVO) & vbTab & _
                         1 & vbTab & _
                         1 & vbTab & _
                         1 & vbTab & _
                         1 & vbTab & _
                         1
                
                BREC.MoveNext
            Loop
        End With
    End If
    BREC.Close
    
End Sub

Private Sub MostrDadosNeto()
    
    If (grdMENUFILHO.Rows - 1) = 0 Then Exit Sub
    If grdMENUFILHO.Row = 0 Then Exit Sub
    
    Frame12.Caption = "Menu [ " & grdMENUFILHO.Cell(flexcpText, grdMENUFILHO.Row, conCOL_SonMenuFilho_Texto) & " ]"
    Call objBLBFunc.CarregaDadosGrdFilho(grdMENUNETO, dacEnumUpdateAction_Ignore, conCOL_SonMenuNeto_CIGLA, grdMENUFILHO.Cell(flexcpText, grdMENUFILHO.Row, conCOL_SonMenuFilho_CIGLA2))
    
    If grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_CIGLA) = "G" Then
          Frame12.Caption = ""
          Call objBLBFunc.CarregaDadosGrdFilho(grdMENUNETO, dacEnumUpdateAction_Ignore, conCOL_SonMenuNeto_CIGLA, grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_CIGLA))
    End If

End Sub

Private Sub AtivoSN()
    
    Dim I   As Long
    Dim j   As Long
    
    For I = 1 To (grdMENUFILHO.Rows - 1)
        
        If grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_CIGLA) = grdMENUFILHO.Cell(flexcpText, I, conCOL_SonMenuFilho_CIGLA) Then
            grdMENUFILHO.Cell(flexcpText, I, conCOL_SonMenuFilho_Ativo) = grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_Ativo)
            
            For j = 1 To (grdMENUNETO.Rows - 1)
                If grdMENUFILHO.Cell(flexcpText, I, conCOL_SonMenuFilho_CIGLA2) = grdMENUNETO.Cell(flexcpText, j, conCOL_SonMenuNeto_CIGLA) Then
                    grdMENUNETO.Cell(flexcpText, j, conCOL_SonMenuNeto_Ativo) = grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_Ativo)
                End If
            Next j
        End If
    
    Next I

End Sub


Private Sub AtivoSNFilho()
    
    Dim I   As Long
    
    For I = 1 To (grdMENUNETO.Rows - 1)
        If grdMENUFILHO.Cell(flexcpText, grdMENUFILHO.Row, conCOL_SonMenuFilho_CIGLA2) = grdMENUNETO.Cell(flexcpText, I, conCOL_SonMenuNeto_CIGLA) Then
            grdMENUNETO.Cell(flexcpText, I, conCOL_SonMenuNeto_Ativo) = grdMENUFILHO.Cell(flexcpText, grdMENUFILHO.Row, conCOL_SonMenuFilho_Ativo)
        End If
    Next I

End Sub

Private Sub CarregaGrdMenuPai()

    Dim I           As Integer
    Dim lngLINHA    As Long
    If IsArray(arrMENUPAI) Then
        With grdMENUPAI
            For I = 1 To UBound(arrMENUPAI)
                lngLINHA = .FindRow(arrMENUPAI(I, 2), , conCOL_SonMenuPai_Codigo)
                If lngLINHA > -1 Then .Cell(flexcpText, lngLINHA, conCOL_SonMenuPai_Ativo) = arrMENUPAI(I, 6)
            Next I
        End With
    End If

End Sub

Private Sub CarregaGrdMenuFilho()

    Dim I           As Integer
    Dim lngLINHA    As Long
    
    If IsArray(arrMENUFILHO) Then
        With grdMENUFILHO
            For I = 1 To UBound(arrMENUFILHO)
                lngLINHA = .FindRow(arrMENUFILHO(I, 2), , conCOL_SonMenuFilho_Codigo)
                If lngLINHA > -1 Then .Cell(flexcpText, lngLINHA, conCOL_SonMenuFilho_Ativo) = arrMENUFILHO(I, 7)
            Next I
        End With
    End If

End Sub


Private Sub CarregaGrdMenuNeto()

    Dim I               As Integer
    Dim j               As Integer
    Dim lngLINHA        As Long
    Dim lngTAMACESSO    As Long
    
    If IsArray(arrMENUNETO) Then
        With grdMENUNETO
            For I = 1 To UBound(arrMENUNETO)
                lngLINHA = .FindRow(arrMENUNETO(I, 2), , conCOL_SonMenuNeto_Codigo)
                If lngLINHA > -1 Then
                    .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_Ativo) = arrMENUNETO(I, 9)
                   
                    .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_INCLUIR) = 0
                    .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_ALTERAR) = 0
                    .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_EXCLUIR) = 0
                    .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_CONSULTAR) = 0
                    .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_IMPRIMIR) = 0
                   
                    lngTAMACESSO = Len(Trim(arrMENUNETO(I, 8)))
                    For j = 1 To lngTAMACESSO
                        If Mid(arrMENUNETO(I, 8), j, 1) = "I" Then .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_INCLUIR) = 1
                        If Mid(arrMENUNETO(I, 8), j, 1) = "A" Then .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_ALTERAR) = 1
                        If Mid(arrMENUNETO(I, 8), j, 1) = "E" Then .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_EXCLUIR) = 1
                        If Mid(arrMENUNETO(I, 8), j, 1) = "C" Then .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_CONSULTAR) = 1
                        If Mid(arrMENUNETO(I, 8), j, 1) = "R" Then .Cell(flexcpText, lngLINHA, conCOL_SonMenuNeto_IMPRIMIR) = 1
                    Next j
                End If
            Next I
        End With
    End If

End Sub


Private Sub AtivoSNFilho2()
    
    Dim I   As Long
    
    For I = 1 To (grdMENUNETO.Rows - 1)
        If grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_CIGLA) = grdMENUNETO.Cell(flexcpText, I, conCOL_SonMenuNeto_CIGLA) Then
            grdMENUNETO.Cell(flexcpText, I, conCOL_SonMenuNeto_Ativo) = grdMENUPAI.Cell(flexcpText, grdMENUPAI.Row, conCOL_SonMenuPai_Ativo)
        End If
    Next I

End Sub


