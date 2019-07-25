VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADFORNEC 
   Caption         =   "Cadastro de Fornecedores"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stFornec 
      Height          =   5175
      Left            =   0
      TabIndex        =   27
      Top             =   1080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Dados Cadastrais"
      TabPicture(0)   =   "frmCADFORNEC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dados de Cobrança"
      TabPicture(1)   =   "frmCADFORNEC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "End. Retirada"
      TabPicture(2)   =   "frmCADFORNEC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Produtos"
      TabPicture(3)   =   "frmCADFORNEC.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Cotação/Pedidos"
      TabPicture(4)   =   "frmCADFORNEC.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Fiscal"
      TabPicture(5)   =   "frmCADFORNEC.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame11"
      Tab(5).Control(1)=   "Frame12"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Qualidade"
      TabPicture(6)   =   "frmCADFORNEC.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "SSTab1"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Transportadoras"
      TabPicture(7)   =   "frmCADFORNEC.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame14"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame15"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Dados Adicionais"
      TabPicture(8)   =   "frmCADFORNEC.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame17"
      Tab(8).ControlCount=   1
      Begin TabDlg.SSTab SSTab1 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   96
         Top             =   720
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7646
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
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
         TabCaption(0)   =   "Não Conf. / Niveis de Risco"
         TabPicture(0)   =   "frmCADFORNEC.frx":00FC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame18"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame13"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Parâmetros"
         TabPicture(1)   =   "frmCADFORNEC.frx":0118
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame21"
         Tab(1).Control(1)=   "Frame20"
         Tab(1).Control(2)=   "Frame19"
         Tab(1).ControlCount=   3
         Begin VB.Frame Frame21 
            Caption         =   "[ Requer Sempre Inspeção de Qualidade ]"
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
            Height          =   615
            Left            =   -74880
            TabIndex        =   121
            Top             =   3600
            Width           =   8775
            Begin VB.OptionButton optREQQUALID 
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
               ForeColor       =   &H8000000D&
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   123
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton optREQQUALID 
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
               ForeColor       =   &H8000000D&
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   122
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame20 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   111
            Top             =   1320
            Width           =   8775
            Begin MSFlexGridLib.MSFlexGrid flxProdNaoRecebInsp 
               Height          =   1935
               Left            =   120
               TabIndex        =   112
               Top             =   240
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   3413
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "[ Produtos que Não Precisan receber Inspeção ]"
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
            Left            =   -74880
            TabIndex        =   105
            Top             =   480
            Width           =   8775
            Begin VB.TextBox txtPRODNAO 
               Height          =   285
               Left            =   960
               MaxLength       =   10
               TabIndex        =   109
               Text            =   "txtPRODNAO"
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cboPRODNAO 
               Height          =   315
               Left            =   2640
               TabIndex        =   108
               Text            =   "cboPRODNAO"
               Top             =   360
               Width           =   5655
            End
            Begin VB.CommandButton Command5 
               Height          =   315
               Left            =   2280
               Picture         =   "frmCADFORNEC.frx":0134
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   360
               Width           =   375
            End
            Begin VB.CommandButton Command4 
               Height          =   315
               Left            =   8280
               Picture         =   "frmCADFORNEC.frx":0236
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label25 
               Caption         =   "Produto:"
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
               Left            =   120
               TabIndex        =   110
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Frame Frame13 
            Height          =   3255
            Left            =   120
            TabIndex        =   103
            Top             =   960
            Width           =   8775
            Begin MSFlexGridLib.MSFlexGrid flxQualidade 
               Height          =   2895
               Left            =   120
               TabIndex        =   104
               Top             =   240
               Width           =   8535
               _ExtentX        =   15055
               _ExtentY        =   5106
               _Version        =   393216
               FixedCols       =   0
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame18 
            Height          =   615
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   8775
            Begin VB.CommandButton Command3 
               Height          =   315
               Left            =   2640
               Picture         =   "frmCADFORNEC.frx":0338
               Style           =   1  'Graphical
               TabIndex        =   100
               Top             =   240
               Width           =   375
            End
            Begin VB.ComboBox cboNivelRisco 
               Height          =   315
               Left            =   3000
               TabIndex        =   99
               Text            =   "cboNivelRisco"
               Top             =   240
               Width           =   4935
            End
            Begin VB.TextBox txtNivelRisco 
               Height          =   285
               Left            =   1680
               MaxLength       =   10
               TabIndex        =   98
               Text            =   "txtNivelRi"
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nivel de Risco"
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
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   1260
            End
            Begin VB.Label lblCorRisc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   8040
               TabIndex        =   101
               Top             =   240
               Width           =   615
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   -74880
         TabIndex        =   78
         Top             =   720
         Width           =   9015
         Begin VB.TextBox txtPORC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7920
            TabIndex        =   86
            Text            =   "txtPORC"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtCodProdFornec 
            Height          =   285
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   84
            Text            =   "txtCodProdFornec"
            Top             =   600
            Width           =   3375
         End
         Begin VB.CommandButton cmdGravProd 
            Height          =   315
            Left            =   8520
            Picture         =   "frmCADFORNEC.frx":043A
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2280
            Picture         =   "frmCADFORNEC.frx":053C
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboProduto 
            Height          =   315
            Left            =   2640
            TabIndex        =   80
            Text            =   "cboProduto"
            Top             =   240
            Width           =   5895
         End
         Begin VB.TextBox txtProduto 
            Height          =   285
            Left            =   960
            MaxLength       =   10
            TabIndex        =   79
            Text            =   "txtProduto"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Qtde. Lote/p Inspeção %"
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
            Left            =   5640
            TabIndex        =   85
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label23 
            Caption         =   "Código do Fornecedor :"
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
            Left            =   120
            TabIndex        =   83
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label21 
            Caption         =   "Produto:"
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
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   76
         Top             =   1680
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxPrcProd 
            Height          =   855
            Left            =   120
            TabIndex        =   88
            Top             =   2400
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   1508
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxProduto 
            Height          =   2175
            Left            =   120
            TabIndex        =   77
            Top             =   120
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame17 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   69
         Top             =   720
         Width           =   8895
         Begin VB.TextBox txtBLOQFORNEC 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6840
            TabIndex        =   75
            Text            =   "txtBLOQFORNEC"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtSOMLIBIGF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6840
            TabIndex        =   74
            Text            =   "txtSOMLIBIGF"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtAVISIGF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6840
            TabIndex        =   73
            Text            =   "txtAVISIGF"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Bloquear fornecedor Quando IQF estiver ABAIXO de"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   4455
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Liberar fornecedor somente com altorização quando IQF estiver a baixo de"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   600
            Width           =   6360
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Avisar quando IQF estiver ABAIXO de"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   3225
         End
      End
      Begin VB.Frame Frame15 
         Height          =   735
         Left            =   -74880
         TabIndex        =   67
         Top             =   720
         Width           =   9015
         Begin VB.TextBox txtCODTRANSP 
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   92
            Text            =   "txtCODTRAN"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboTRANSP 
            Height          =   315
            Left            =   3120
            TabIndex        =   91
            Text            =   "cboTRANSP"
            Top             =   240
            Width           =   5415
         End
         Begin VB.CommandButton cmdTRANSP 
            Height          =   315
            Left            =   2760
            Picture         =   "frmCADFORNEC.frx":063E
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdIncTransp 
            Height          =   315
            Left            =   8520
            Picture         =   "frmCADFORNEC.frx":0740
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   240
            Width           =   375
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
            TabIndex        =   93
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame14 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   65
         Top             =   1440
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxTRANSP 
            Height          =   3255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   5741
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame12 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   59
         Top             =   1440
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxFISCAL 
            Height          =   3255
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   5741
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   -74880
         TabIndex        =   55
         Top             =   720
         Width           =   9015
         Begin VB.TextBox txtCODNATOPER 
            Height          =   285
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   58
            Text            =   "txtCODNATO"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboNATOPERCAO 
            Height          =   315
            Left            =   3600
            TabIndex        =   60
            Text            =   "cboNATOPERCAO"
            Top             =   240
            Width           =   4935
         End
         Begin VB.CommandButton cmdNATOPERACAO 
            Height          =   315
            Left            =   3240
            Picture         =   "frmCADFORNEC.frx":0842
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   8520
            Picture         =   "frmCADFORNEC.frx":0944
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Natureza de Operação:"
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
            Left            =   240
            TabIndex        =   56
            Top             =   240
            Width           =   1995
         End
      End
      Begin VB.Frame Frame8 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   53
         Top             =   720
         Width           =   9015
         Begin TabDlg.SSTab SSTab2 
            Height          =   3975
            Left            =   240
            TabIndex        =   113
            Top             =   240
            Width           =   8655
            _ExtentX        =   15266
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
            TabCaption(0)   =   "Cotações"
            TabPicture(0)   =   "frmCADFORNEC.frx":0A46
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame9"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Pedidos"
            TabPicture(1)   =   "frmCADFORNEC.frx":0A62
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame10"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Notas Fiscais"
            TabPicture(2)   =   "frmCADFORNEC.frx":0A7E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame16"
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame16 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3495
               Left            =   -74880
               TabIndex        =   118
               Top             =   360
               Width           =   8415
               Begin MSFlexGridLib.MSFlexGrid flxNF 
                  Height          =   1695
                  Left            =   120
                  TabIndex        =   119
                  Top             =   1680
                  Width           =   8175
                  _ExtentX        =   14420
                  _ExtentY        =   2990
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
                  Appearance      =   0
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
               Height          =   3495
               Left            =   -74880
               TabIndex        =   116
               Top             =   360
               Width           =   8415
               Begin MSFlexGridLib.MSFlexGrid flxPedido 
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   117
                  Top             =   1800
                  Width           =   8175
                  _ExtentX        =   14420
                  _ExtentY        =   2778
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
                  Appearance      =   0
               End
            End
            Begin VB.Frame Frame9 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3495
               Left            =   120
               TabIndex        =   114
               Top             =   360
               Width           =   8415
               Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   120
                  Top             =   120
                  Width           =   8175
                  _ExtentX        =   14420
                  _ExtentY        =   2778
                  _Version        =   393216
                  FixedCols       =   0
                  Appearance      =   0
               End
               Begin MSFlexGridLib.MSFlexGrid FlxCotacoes 
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   115
                  Top             =   1800
                  Width           =   8175
                  _ExtentX        =   14420
                  _ExtentY        =   2778
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
                  Appearance      =   0
               End
            End
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   47
         Top             =   720
         Width           =   9015
         Begin VB.ComboBox cboEstReti 
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Text            =   "cboEstReti"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtCepReti 
            Height          =   285
            Left            =   1440
            MaxLength       =   9
            TabIndex        =   21
            Text            =   "txtCepRet"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtCidRetirada 
            Height          =   285
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   19
            Text            =   "txtCidRetirada"
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtEndReti 
            Height          =   285
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   20
            Text            =   "txtEndReti"
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox txtEndRetirada 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   18
            Text            =   "txtEndRetirada"
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label20 
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
            Height          =   195
            Left            =   600
            TabIndex        =   52
            Top             =   1730
            Width           =   660
         End
         Begin VB.Label Label19 
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
            Height          =   195
            Left            =   840
            TabIndex        =   51
            Top             =   1320
            Width           =   405
         End
         Begin VB.Label Label18 
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
            Height          =   195
            Left            =   600
            TabIndex        =   50
            Top             =   600
            Width           =   660
         End
         Begin VB.Label Label17 
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
            Height          =   195
            Left            =   660
            TabIndex        =   49
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label16 
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
            Height          =   195
            Left            =   360
            TabIndex        =   48
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   45
         Top             =   1320
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid flxDadCobr 
            Height          =   3375
            Left            =   120
            TabIndex        =   46
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
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   -74880
         TabIndex        =   43
         Top             =   720
         Width           =   9015
         Begin VB.CommandButton cmbGravPagto 
            Height          =   315
            Left            =   8520
            Picture         =   "frmCADFORNEC.frx":0A9A
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdPesq 
            Height          =   315
            Left            =   1800
            Picture         =   "frmCADFORNEC.frx":0B9C
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   2160
            TabIndex        =   16
            Text            =   "cboBanco"
            Top             =   240
            Width           =   6375
         End
         Begin VB.TextBox txtBanco 
            Height          =   285
            Left            =   840
            MaxLength       =   10
            TabIndex        =   15
            Text            =   "txtBanco"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label15 
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
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4335
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   9015
         Begin VB.ComboBox cboStatus 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   95
            Text            =   "cboStatus"
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cboGRPDESP 
            Height          =   315
            ItemData        =   "frmCADFORNEC.frx":0C9E
            Left            =   1920
            List            =   "frmCADFORNEC.frx":0CA0
            TabIndex        =   14
            Text            =   "cboSite"
            Top             =   3840
            Visible         =   0   'False
            Width           =   4470
         End
         Begin VB.TextBox txtIQF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6360
            MaxLength       =   10
            TabIndex        =   64
            Text            =   "txtIQF"
            Top             =   3120
            Width           =   1335
         End
         Begin MSMask.MaskEdBox mskDtCadastro 
            Height          =   285
            Left            =   6360
            TabIndex        =   13
            Top             =   3480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboSite 
            Height          =   315
            ItemData        =   "frmCADFORNEC.frx":0CA2
            Left            =   1920
            List            =   "frmCADFORNEC.frx":0CA4
            TabIndex        =   12
            Text            =   "cboSite"
            Top             =   3480
            Width           =   2670
         End
         Begin VB.ComboBox cboEmail 
            Height          =   315
            Left            =   1920
            TabIndex        =   11
            Text            =   "cboEmail"
            Top             =   3120
            Width           =   2655
         End
         Begin VB.ComboBox cboContato 
            Height          =   315
            Left            =   6360
            TabIndex        =   10
            Text            =   "cboContato"
            Top             =   2760
            Width           =   2055
         End
         Begin VB.ComboBox cboTel 
            Height          =   315
            Left            =   1920
            TabIndex        =   9
            Text            =   "cboTel"
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox txtCEP 
            Height          =   285
            Left            =   6360
            MaxLength       =   9
            TabIndex        =   8
            Text            =   "txtCEP"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Text            =   "cboEstado"
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox txtBairro 
            Height          =   285
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   6
            Text            =   "txtBairro"
            Top             =   2040
            Width           =   2535
         End
         Begin VB.TextBox txtCidade 
            Height          =   285
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   5
            Text            =   "txtCidade"
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox txtEndereco 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "txtEndereco"
            Top             =   1320
            Width           =   4575
         End
         Begin VB.TextBox txtNomFantasia 
            Height          =   285
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "txtNomFantasia"
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox txtCNPJCPF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4800
            MaxLength       =   15
            TabIndex        =   1
            Text            =   "txtCNPJCPF"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   0
            Text            =   "txtCodigo"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtRazaoSoc 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "txtDescricao"
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Status"
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
            Left            =   6600
            TabIndex        =   94
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Grupo de Despesas:"
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
            Left            =   75
            TabIndex        =   89
            Top             =   3840
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "IQF:"
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
            Left            =   5880
            TabIndex        =   63
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label Label14 
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
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   4920
            TabIndex        =   42
            Top             =   3495
            Width           =   1290
         End
         Begin VB.Label Label13 
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
            Index           =   0
            Left            =   1395
            TabIndex        =   41
            Top             =   3495
            Width           =   405
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail:"
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
            Left            =   1200
            TabIndex        =   40
            Top             =   3150
            Width           =   600
         End
         Begin VB.Label Label11 
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
            Left            =   5520
            TabIndex        =   39
            Top             =   2805
            Width           =   735
         End
         Begin VB.Label Label10 
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
            Left            =   975
            TabIndex        =   38
            Top             =   2805
            Width           =   825
         End
         Begin VB.Label Label9 
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
            Left            =   5880
            TabIndex        =   37
            Top             =   2445
            Width           =   405
         End
         Begin VB.Label Label8 
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
            Left            =   1140
            TabIndex        =   36
            Top             =   2445
            Width           =   660
         End
         Begin VB.Label Label7 
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
            Left            =   1230
            TabIndex        =   35
            Top             =   2040
            Width           =   570
         End
         Begin VB.Label Label6 
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
            Left            =   1155
            TabIndex        =   34
            Top             =   1680
            Width           =   660
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
            Left            =   930
            TabIndex        =   33
            Top             =   1320
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantasia:"
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
            Left            =   480
            TabIndex        =   32
            Top             =   960
            Width           =   1335
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
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3720
            TabIndex        =   31
            Top             =   285
            Width           =   975
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
            Left            =   1140
            TabIndex        =   30
            Top             =   255
            Width           =   660
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
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   600
            TabIndex        =   29
            Top             =   600
            Width           =   1200
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9375
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
         Picture         =   "frmCADFORNEC.frx":0CA6
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Picture         =   "frmCADFORNEC.frx":11D8
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Picture         =   "frmCADFORNEC.frx":12DA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADFORNEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Integer
Public FILIAL      As Integer
Public strAcesso   As String
Public strMODPAI   As String
Dim objBLBFunc     As Object
Dim objCADFORNEC   As Object
Dim objPESQPADRAO  As Object
Dim arrBANCOS      As Variant
Dim arrTELEFONE    As Variant
Dim arrCONTATO     As Variant
Dim arrEMAIL       As Variant
Dim arrSITE        As Variant
Dim arrPRODUTO     As Variant
Dim arrNATOPER     As Variant
Dim arrTRANSP      As Variant
Dim arrPRODNAOINSP As Variant

Private Sub cboBanco_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboBanco, KeyAscii
End Sub
Private Sub cboBanco_Validate(Cancel As Boolean)
    If cboBanco.ListIndex > -1 Then txtBanco.Text = cboBanco.ItemData(cboBanco.ListIndex)
End Sub

Private Sub cboContato_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
       If cboContato.ListIndex = -1 Then Exit Sub
       cboContato.RemoveItem cboContato.ListIndex
       cboContato.Text = ""
    End If

End Sub

Private Sub cboContato_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub cboContato_Validate(Cancel As Boolean)

   If Len(Trim(cboContato.Text)) > 20 Then
      MsgBox "Somente é permitido 20 Digitos !!!", vbOKOnly + vbCritical, "aviso"
      cboContato.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If Len(Trim(cboContato.Text)) = 0 Then Exit Sub
   
   cboContato.AddItem cboContato.Text
   cboContato.Text = ""

End Sub

Private Sub cboEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       
       If cboEmail.ListIndex = -1 Then Exit Sub
       
       cboEmail.RemoveItem cboEmail.ListIndex
       cboEmail.Text = ""
       
    End If
End Sub

Private Sub cboEmail_Validate(Cancel As Boolean)

   If Len(Trim(cboEmail.Text)) > 60 Then
      MsgBox "Somente é permitido 60 Digitos !!!", vbOKOnly + vbCritical, "aviso"
      cboEmail.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If Len(Trim(cboEmail.Text)) = 0 Then Exit Sub
   
   cboEmail.AddItem cboEmail.Text
   cboEmail.Text = ""

End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboEstado, KeyAscii
End Sub

Private Sub cboEstReti_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboEstReti, KeyAscii
End Sub

Private Sub cboNATOPERCAO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboNATOPERCAO, KeyAscii
End Sub

Private Sub cboNATOPERCAO_Validate(Cancel As Boolean)
    If Len(Trim(cboNATOPERCAO.Text)) > 0 Then txtCODNATOPER.Text = Mid(cboNATOPERCAO.Text, 1, 5)
End Sub

Private Sub cboNivelRisco_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboNivelRisco, KeyAscii
End Sub

Private Sub cboNivelRisco_Validate(Cancel As Boolean)
    If cboNivelRisco.ListIndex > -1 Then txtNivelRisco.Text = cboNivelRisco.ItemData(cboNivelRisco.ListIndex)
End Sub


Private Sub cboPRODNAO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboPRODNAO, KeyAscii
End Sub

Private Sub cboPRODNAO_Validate(Cancel As Boolean)
      If Len(Trim(cboPRODNAO.Text)) > 0 Then txtPRODNAO.Text = Mid(cboPRODNAO.Text, 1, 10)
End Sub



Private Sub cboSite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       
       If cboSite.ListIndex = -1 Then Exit Sub
       
       cboSite.RemoveItem cboSite.ListIndex
       cboSite.Text = ""
       
    End If
End Sub

Private Sub cboSite_Validate(Cancel As Boolean)

   If Len(Trim(cboSite.Text)) > 60 Then
      MsgBox "Somente é permitido 60 Digitos !!!", vbOKOnly + vbCritical, "aviso"
      cboSite.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If Len(Trim(cboSite.Text)) = 0 Then Exit Sub
   
   cboSite.AddItem cboSite.Text
   cboSite.Text = ""

End Sub

Private Sub cboTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       
       If cboTel.ListIndex = -1 Then Exit Sub
       
       cboTel.RemoveItem cboTel.ListIndex
       cboTel.Text = ""
       
    End If
End Sub

Private Sub cboTel_Validate(Cancel As Boolean)

   If Len(Trim(cboTel.Text)) > 13 Then
      MsgBox "Somente é permitido 13 Digitos !!!", vbOKOnly + vbCritical, "aviso"
      cboTel.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If Len(Trim(cboTel.Text)) = 0 Then Exit Sub
   
   cboTel.AddItem cboTel.Text
   cboTel.Text = ""

End Sub


Private Sub cboTRANSP_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTRANSP, KeyAscii
End Sub

Private Sub cboTRANSP_Validate(Cancel As Boolean)
    If cboTRANSP.ListIndex > -1 Then txtCODTRANSP.Text = cboTRANSP.ItemData(cboTRANSP.ListIndex)
End Sub

Private Sub cmbGravPagto_Click()
    If cTipOper = "I" Or cTipOper = "A" Then InclBancos
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    mskDtCadastro.Enabled = False
    
    txtCNPJCPF.Enabled = True
    txtRazaoSoc.Enabled = True
    txtNomFantasia.Enabled = True
    txtEndereco.Enabled = True
    txtCidade.Enabled = True
    txtBairro.Enabled = True
    txtCEP.Enabled = True
    cboEstado.Enabled = True
    
    cboTel.Locked = False
    cboContato.Locked = False
    cboEmail.Locked = False
    cboSite.Locked = False
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame18.Enabled = True
    Frame19.Enabled = True
    Frame21.Enabled = True
    
    Me.Caption = "Cadastro de fornecedores - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
    stFornec.Tab = 0
    txtCNPJCPF.SetFocus

End Sub

Private Sub cmdGravProd_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGrid
End Sub

Private Sub cmdIncTransp_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridTransp
End Sub

Private Sub cmdNATOPERACAO_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADNATOPERACAO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And SGI_ENTSAI = 0"
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Natureza de operação")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODNATOPER.Text = varRETORNO
    
    cboNATOPERCAO.ListIndex = -1
    txtCODNATOPER.SetFocus

End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 4, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADBANCOS" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
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
    
    If Len(Trim(varRETORNO)) > 0 Then txtBanco.Text = varRETORNO
    
    cboBanco.ListIndex = -1
    txtBanco.SetFocus

End Sub

Private Sub CmdSalva_Click()
    
    Dim I As Integer
    
    If ValidaCampos = True Then
       
       objCADFORNEC.MODULO = Me.Name
       
       If cTipOper = "I" Then objCADFORNEC.CodigoFOR = objCADFORNEC.Gera_Codigo(Me.Name)
       
       
       objCADFORNEC.CNPJCPFFOR = txtCNPJCPF.Text
       objCADFORNEC.RAZSOCFOR = txtRazaoSoc.Text
       objCADFORNEC.NOMFANTFOR = txtNomFantasia.Text
       objCADFORNEC.ENDFOR = txtEndereco.Text
       objCADFORNEC.CIDFOR = txtCidade.Text
       objCADFORNEC.BAIFOR = txtBairro.Text
       If cboEstado.ListIndex > -1 Then objCADFORNEC.ESTFOR = cboEstado.ItemData(cboEstado.ListIndex)
       objCADFORNEC.CEPFOR = txtCEP.Text
       objCADFORNEC.DTCADFOR = CDate(mskDtCadastro.Text)
       objCADFORNEC.ENDRETI = txtEndRetirada.Text
       objCADFORNEC.CIDRETI = txtCidRetirada.Text
       objCADFORNEC.BAIRETI = txtEndReti.Text
       objCADFORNEC.STATUS = cboStatus.ItemData(cboStatus.ListIndex)
       
       If optREQQUALID(0).Value = True Then objCADFORNEC.REQINSPQUA = 0
       If optREQQUALID(1).Value = True Then objCADFORNEC.REQINSPQUA = 1
       
       If cboEstReti.ListIndex > -1 Then objCADFORNEC.ESTRETI = cboEstReti.ItemData(cboEstReti.ListIndex)
       
       If Len(Trim(txtNivelRisco.Text)) > 0 Then objCADFORNEC.CODNIVRISCO = txtNivelRisco.Text
       
       If cboGRPDESP.ListIndex > -1 Then objCADFORNEC.GRPDESP = cboGRPDESP.ItemData(cboGRPDESP.ListIndex)
                 
       objCADFORNEC.CEPRETI = txtCepReti.Text
       ''If Len(Trim(txtIQF.Text)) > 0 Then objCADFORNEC.IQF = CLng(txtIQF.Text)
             
       ' Telefone
       If cboTel.ListCount > 0 Then
          ReDim arrTELEFONE(0 To (cboTel.ListCount - 1)) As String
          For I = 0 To (cboTel.ListCount - 1)
              arrTELEFONE(I) = cboTel.List(I) & "^" & cboTel.Name
          Next I
          objCADFORNEC.TELFOR = arrTELEFONE
       Else
          objCADFORNEC.TELFOR = Empty
       End If
       
       ' Contato
       If cboContato.ListCount > 0 Then
          ReDim arrCONTATO(0 To (cboContato.ListCount - 1)) As String
          For I = 0 To (cboContato.ListCount - 1)
              arrCONTATO(I) = cboContato.List(I) & "^" & cboContato.Name
          Next I
          objCADFORNEC.CONTATOFOR = arrCONTATO
       Else
          objCADFORNEC.CONTATOFOR = Empty
       End If
       ' E-Mail
       If cboEmail.ListCount > 0 Then
          ReDim arrEMAIL(0 To (cboEmail.ListCount - 1)) As String
          For I = 0 To (cboEmail.ListCount - 1)
              arrEMAIL(I) = cboEmail.List(I) & "^" & cboEmail.Name
          Next I
          objCADFORNEC.EMAILFOR = arrEMAIL
       Else
          objCADFORNEC.EMAILFOR = Empty
       End If
       'Site
       If cboSite.ListCount > 0 Then
          ReDim arrSITE(0 To (cboSite.ListCount - 1)) As String
          For I = 0 To (cboSite.ListCount - 1)
              arrSITE(I) = cboSite.List(I) & "^" & cboSite.Name
          Next I
          objCADFORNEC.SITEFOR = arrSITE
       Else
          objCADFORNEC.SITEFOR = Empty
       End If
       'Bancos
       If flxDadCobr.Rows > 1 Then
          ReDim arrBANCOS(1 To (flxDadCobr.Rows - 1)) As String
          For I = 1 To (flxDadCobr.Rows - 1)
              arrBANCOS(I) = flxDadCobr.TextMatrix(I, 1)
          Next I
          objCADFORNEC.BANCOS = arrBANCOS
       Else
          ReDim arrBANCOS(0) As String
          objCADFORNEC.BANCOS = arrBANCOS
       End If
       'Produtos
       If (flxProduto.Rows - 1) > 0 Then
          ReDim arrPRODUTO(1 To (flxProduto.Rows - 1), 1 To 3) As String
          For I = 1 To (flxProduto.Rows - 1)
              arrPRODUTO(I, 1) = flxProduto.TextMatrix(I, 0)
              arrPRODUTO(I, 2) = flxProduto.TextMatrix(I, 2)
              arrPRODUTO(I, 3) = flxProduto.TextMatrix(I, 3)
          Next I
          objCADFORNEC.PRODUTO = arrPRODUTO
       Else
          ReDim arrPRODUTO(0) As String
          objCADFORNEC.PRODUTO = arrPRODUTO
       End If
       '' Natureza de Operação
       If flxFISCAL.Rows > 1 Then
          ReDim arrNATOPER(1 To (flxFISCAL.Rows - 1)) As String
          For I = 1 To (flxFISCAL.Rows - 1)
              arrNATOPER(I) = flxFISCAL.TextMatrix(I, 1)
          Next I
          objCADFORNEC.NATOPER = arrNATOPER
       Else
          ReDim arrNATOPER(0) As String
          objCADFORNEC.NATOPER = arrNATOPER
       End If
       '' Transportadoras
       If flxTRANSP.Rows > 1 Then
          ReDim arrTRANSP(1 To (flxTRANSP.Rows - 1)) As Long
          For I = 1 To (flxTRANSP.Rows - 1)
              arrTRANSP(I) = flxTRANSP.TextMatrix(I, 1)
          Next I
          objCADFORNEC.TRANSP = arrTRANSP
       Else
          ReDim arrTRANSP(0) As String
          objCADFORNEC.TRANSP = arrTRANSP
       End If
       
       If Len(Trim(txtAVISIGF.Text)) > 0 Then objCADFORNEC.AVISAIGF = CInt(txtAVISIGF.Text)
       If Len(Trim(txtSOMLIBIGF.Text)) > 0 Then objCADFORNEC.LIBIGF = CInt(txtSOMLIBIGF.Text)
       If Len(Trim(txtBLOQFORNEC)) > 0 Then objCADFORNEC.TRAVAIGF = CInt(txtBLOQFORNEC.Text)
    
       '' Produtos que não precisam de Inspeção
       arrPRODNAOINSP = Empty
       If (flxProdNaoRecebInsp.Rows - 1) > 0 Then
          ReDim arrPRODNAOINSP(1 To (flxProdNaoRecebInsp.Rows - 1)) As String
          For I = 1 To (flxProdNaoRecebInsp.Rows - 1)
              arrPRODNAOINSP(I) = Trim(flxProdNaoRecebInsp.TextMatrix(I, 0))
          Next I
       End If
       objCADFORNEC.PRODNAOINSP = arrPRODNAOINSP
       
       '' Grava as informações
       If objCADFORNEC.GRAVA(cTipOper) = True Then
          
          MsgBox "O Fornecedor foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If objCADFORNEC.Atualiza(cTipOper, objCADFORNEC.CodigoFOR, FILIAL, Me.Name) = False Then Exit Sub
          
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objCADFORNEC = Nothing
             Unload Me
          End If
          
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
    Set objBLBFunc = Nothing
    Set objCADFORNEC = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PROD.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PROD" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And PROD.SGI_PRODUTOTIPO = 0"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtProduto.Text = varRETORNO
    
    cboProduto.ListIndex = -1
    txtProduto.SetFocus

End Sub


Private Sub Command2_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridNatOper
End Sub



Private Sub Command3_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADRISCO" & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Risco do Fornecedor")
    
    If Len(Trim(varRETORNO)) > 0 Then txtNivelRisco.Text = varRETORNO
    
    cboNivelRisco.ListIndex = -1
    txtNivelRisco.SetFocus

End Sub

Private Sub Command4_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridProdNao
End Sub

Private Sub Command5_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PROD.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PROD" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtPRODNAO.Text = varRETORNO
    
    cboPRODNAO.ListIndex = -1
    txtPRODNAO.SetFocus

End Sub

Private Sub FlxCotacoes_Click()
    ConfGridPedido
    ConfGridNF
    If ((FlxCotacoes.Rows - 1) > 0) And (cTipOper = "C" Or cTipOper = "A") And (Len(Trim(FlxCotacoes.TextMatrix(FlxCotacoes.Row, 0))) > 0) Then CarregaGridPedidos
End Sub

Private Sub FlxCotacoes_RowColChange()
    ConfGridPedido
    ConfGridNF
    If ((FlxCotacoes.Rows - 1) > 0) And (cTipOper = "C" Or cTipOper = "A") And (Len(Trim(FlxCotacoes.TextMatrix(FlxCotacoes.Row, 0))) > 0) Then CarregaGridPedidos
End Sub

Private Sub flxDadCobr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxDadCobr.Rows = 2 Then flxDadCobr.Rows = 1
       If flxDadCobr.Rows > 2 Then flxDadCobr.RemoveItem flxDadCobr.RowSel
    End If
End Sub

Private Sub flxFISCAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxFISCAL.Rows = 2 Then flxFISCAL.Rows = 1
       If flxFISCAL.Rows > 2 Then flxFISCAL.RemoveItem flxFISCAL.RowSel
    End If
End Sub

Private Sub flxPedido_Click()
    ConfGridNF
    If ((flxPedido.Rows - 1) > 0) And (cTipOper = "C" Or cTipOper = "A") And (Len(Trim(flxPedido.TextMatrix(flxPedido.Row, 0))) > 0) Then CarregaGridNf
End Sub

Private Sub flxPedido_RowColChange()
    ConfGridNF
    If ((flxPedido.Rows - 1) > 0) And (cTipOper = "C" Or cTipOper = "A") And (Len(Trim(flxPedido.TextMatrix(flxPedido.Row, 0))) > 0) Then CarregaGridNf
End Sub

Private Sub flxProdNaoRecebInsp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If (cTipOper = "I") Or (cTipOper = "A") Then
          If flxProdNaoRecebInsp.Rows = 2 Then flxProdNaoRecebInsp.Rows = 1
          If flxProdNaoRecebInsp.Rows > 2 Then flxProdNaoRecebInsp.RemoveItem flxProdNaoRecebInsp.RowSel
          txtPRODNAO.SetFocus
       End If
    End If
End Sub

Private Sub flxProduto_Click()
    If (flxProduto.Rows - 1 > 0) And (cTipOper = "C" Or cTipOper = "A") Then CarregaGridCitacao
End Sub

Private Sub flxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If (cTipOper = "I") Or (cTipOper = "A") Then
          If flxProduto.Rows = 2 Then flxProduto.Rows = 1
          If flxProduto.Rows > 2 Then flxProduto.RemoveItem flxProduto.RowSel
          txtProduto.SetFocus
       End If
    End If
End Sub

Private Sub flxProduto_RowColChange()
    If (flxProduto.Rows - 1 > 0) And (cTipOper = "C" Or cTipOper = "A") Then
       CarregaGridCitacao
       Call CarregaGridPrecos(flxProduto.TextMatrix(flxProduto.Row, 0))
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

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADFORNEC = CreateObject("CADFORNEC.clsCADFORNEC")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADFORNEC.FILIAL = FILIAL
   
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
    Frame3.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame11.Enabled = True
    Frame15.Enabled = True
    Frame17.Enabled = True
    Frame18.Enabled = True
    Frame19.Enabled = True
    Frame21.Enabled = True
    
    Me.Caption = "Cadastro de fornecedores - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFORNEC
    optREQQUALID(0).Value = True
    
    objBLBFunc.Preenche_Estado cboEstado
    objBLBFunc.Preenche_Estado cboEstReti
    
    objCADFORNEC.PreenchComboBancos cboBanco
    objCADFORNEC.PreencheComboProd cboProduto
    objCADFORNEC.PreencheComboNatOper cboNATOPERCAO
    objCADFORNEC.PreenchComboTransportadoras cboTRANSP
    objCADFORNEC.PreenchComboGrpDespesas cboGRPDESP
    objCADFORNEC.PreenchComboNivelRiscos cboNivelRisco
    objBLBFunc.PreenchComboStatus cboStatus
    objCADFORNEC.PreencheComboProd cboPRODNAO
    
    cboStatus.ListIndex = 1
    
    ConfGridProd
    ConfGridBancos
    ConfGridCota
    ConfGridPedido
    ConfGridNatOper
    ConfGridTransp
    ConfGridNF
    ConfGridPrcProd
    ConfGridProdNao
    
    txtCodigo.Text = ""
    mskDtCadastro.Text = Format(Date, "DD/MM/YYYY")
    
    cboBanco.ListIndex = -1
    cboProduto.ListIndex = -1
    cboEstado.ListIndex = -1
    cboNivelRisco.ListIndex = -1
    lblCorRisc.Caption = ""
    
    stFornec.Tab = 0
    
    Call Desabilita_Tables
   
End Sub

Private Sub stFornec_Click(PreviousTab As Integer)

   If stFornec.Tab = 1 Then
      If Frame3.Enabled = True Then txtBanco.SetFocus
   ElseIf stFornec.Tab = 2 Then
      If Frame5.Enabled = True Then txtEndRetirada.SetFocus
   ElseIf stFornec.Tab = 3 Then
      If Frame6.Enabled = True Then txtProduto.SetFocus
   ElseIf stFornec.Tab = 4 Then
   End If
   
End Sub

Private Sub txtAVISIGF_Validate(Cancel As Boolean)

    If Len(Trim(txtAVISIGF.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtAVISIGF.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtAVISIGF.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If CInt(txtAVISIGF.Text) < 0 Then
       MsgBox "Somente é permitido dados positivos !!!", vbOKOnly + vbExclamation, "Aviso"
       txtAVISIGF.Text = ""
       Cancel = True
       Exit Sub
    End If
    If CInt(txtAVISIGF.Text) > 100 Then
       MsgBox "Somente é permitido IGF até 100 !!!", vbOKOnly + vbExclamation, "Aviso"
       txtAVISIGF.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If Len(Trim(txtSOMLIBIGF.Text)) > 0 Then
       If (CInt(txtAVISIGF.Text) <= CInt(txtSOMLIBIGF.Text)) Then
          MsgBox "aviso de IGF não pode ser menor que os demais valores !!!", vbOKOnly + vbExclamation, "Aviso"
          Cancel = True
          Exit Sub
       End If
    End If
    If Len(Trim(txtBLOQFORNEC.Text)) > 0 Then
       If (CInt(txtAVISIGF.Text) <= CInt(txtBLOQFORNEC.Text)) Then
          MsgBox "aviso de IGF não pode ser menor que os demais valores !!!", vbOKOnly + vbExclamation, "Aviso"
          Cancel = True
          Exit Sub
       End If
    End If

End Sub

Private Sub txtBairro_GotFocus()
    objBLBFunc.SelecionaCampos txtBairro.Name, frmCADFORNEC
End Sub

Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtBanco_GotFocus()
    objBLBFunc.SelecionaCampos txtBanco.Name, frmCADFORNEC
End Sub

Private Sub txtBanco_Validate(Cancel As Boolean)
    
    Dim I       As Integer
    Dim blACHOU As Boolean
    
    If Len(Trim(txtBanco.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtBanco.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtBanco.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    blACHOU = False
    For I = 0 To (cboBanco.ListCount - 1)
        If CInt(txtBanco.Text) = cboBanco.ItemData(I) Then
           blACHOU = True
           cboBanco.ListIndex = I
        End If
    Next I
    
    If blACHOU = False Then
       MsgBox "Este banco não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtBanco.Text = ""
       cboBanco.ListIndex = -1
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtBLOQFORNEC_Validate(Cancel As Boolean)

    If Len(Trim(txtBLOQFORNEC.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtBLOQFORNEC.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtBLOQFORNEC.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If CInt(txtBLOQFORNEC.Text) < 0 Then
       MsgBox "Somente é permitido dados positivos !!!", vbOKOnly + vbExclamation, "Aviso"
       txtBLOQFORNEC.Text = ""
       Cancel = True
       Exit Sub
    End If
    If CInt(txtBLOQFORNEC.Text) > 100 Then
       MsgBox "Somente é permitido IGF até 100 !!!", vbOKOnly + vbExclamation, "Aviso"
       txtBLOQFORNEC.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If Len(Trim(txtAVISIGF.Text)) > 0 Then
       If (CInt(txtBLOQFORNEC.Text) >= CInt(txtAVISIGF.Text)) Then
          MsgBox "Aviso de IGF não pode ser meior que valor acima !!!", vbOKOnly + vbExclamation, "Aviso"
          Cancel = True
          Exit Sub
       End If
    End If
    If Len(Trim(txtSOMLIBIGF.Text)) > 0 Then
       If (CInt(txtBLOQFORNEC.Text) >= CInt(txtSOMLIBIGF.Text)) Then
          MsgBox "Aviso de IGF não pode ser maior que valor acima !!!", vbOKOnly + vbExclamation, "Aviso"
          Cancel = True
          Exit Sub
       End If
    End If

End Sub

Private Sub txtCEP_GotFocus()
    objBLBFunc.SelecionaCampos txtCEP.Name, frmCADFORNEC
End Sub

Private Sub txtCEP_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCepReti_GotFocus()
    objBLBFunc.SelecionaCampos txtCepReti.Name, frmCADFORNEC
End Sub

Private Sub txtCepReti_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCidade_GotFocus()
    objBLBFunc.SelecionaCampos txtCidade.Name, frmCADFORNEC
End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCidRetirada_GotFocus()
    objBLBFunc.SelecionaCampos txtCidRetirada.Name, frmCADFORNEC
End Sub

Private Sub txtCidRetirada_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCNPJCPF_GotFocus()
    objBLBFunc.SelecionaCampos txtCNPJCPF.Name, frmCADFORNEC
End Sub

Private Sub txtCNPJCPF_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCODNATOPER_GotFocus()
    objBLBFunc.SelecionaCampos txtCODNATOPER.Name, frmCADFORNEC
End Sub

Private Sub txtCODNATOPER_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCODNATOPER_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCODNATOPER.Text)) = 0 Then Exit Sub
    
   cboNATOPERCAO.ListIndex = -1
   For I = 0 To (cboNATOPERCAO.ListCount - 1)
       If Trim(Mid(cboNATOPERCAO.List(I), 1, 5)) = txtCODNATOPER.Text Then cboNATOPERCAO.ListIndex = I
   Next I
    
   If cboNATOPERCAO.ListIndex = -1 Then
      MsgBox "Esta natureza não existe !!!", vbOKOnly + vbExclamation, "Aviso"
      txtCODNATOPER.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub


Private Sub txtCODTRANSP_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTRANSP.Name, frmCADFORNEC
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

Private Sub txtEndereco_GotFocus()
    objBLBFunc.SelecionaCampos txtEndereco.Name, frmCADFORNEC
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtEndReti_GotFocus()
    objBLBFunc.SelecionaCampos txtEndReti.Name, frmCADFORNEC
End Sub

Private Sub txtEndReti_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtEndRetirada_GotFocus()
    objBLBFunc.SelecionaCampos txtEndRetirada.Name, frmCADFORNEC
End Sub

Private Sub txtEndRetirada_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtIQF_Validate(Cancel As Boolean)

    If Len(Trim(txtIQF.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtIQF.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtIQF.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtNivelRisco_GotFocus()
    objBLBFunc.SelecionaCampos txtNivelRisco.Name, frmCADFORNEC
End Sub

Private Sub txtNivelRisco_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blACHOU As Boolean
    
    If Len(Trim(txtNivelRisco.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtNivelRisco.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtNivelRisco.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    blACHOU = False
    For I = 0 To (cboNivelRisco.ListCount - 1)
        If CInt(txtNivelRisco.Text) = cboNivelRisco.ItemData(I) Then
           blACHOU = True
           cboNivelRisco.ListIndex = I
        End If
    Next I
    
    If blACHOU = False Then
       MsgBox "Este Nivel de Risco não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtNivelRisco.Text = ""
       cboNivelRisco.ListIndex = -1
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtNomFantasia_GotFocus()
    objBLBFunc.SelecionaCampos txtNomFantasia.Name, frmCADFORNEC
End Sub

Private Sub txtNomFantasia_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub



Private Sub txtPORC_GotFocus()
    objBLBFunc.SelecionaCampos txtPORC.Name, frmCADFORNEC
End Sub

Private Sub txtPORC_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPORC.Text
End Sub

Private Sub txtPORC_Validate(Cancel As Boolean)
    If IsNumeric(txtPORC.Text) Then
       If CLng(txtPORC.Text) > 100 Then
          MsgBox "Somente é permito 100% !!!"
          txtPORC.Text = ""
          Cancel = True
       End If
    End If
End Sub

Private Sub txtPRODNAO_GotFocus()
    objBLBFunc.SelecionaCampos txtPRODNAO.Name, frmCADFORNEC
End Sub

Private Sub txtPRODNAO_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtPRODNAO_Validate(Cancel As Boolean)
   
   Dim I As Integer

   If Len(Trim(txtPRODNAO.Text)) = 0 Then Exit Sub
    
   cboPRODNAO.ListIndex = -1
   For I = 0 To (cboPRODNAO.ListCount - 1)
       If Trim(Mid(cboPRODNAO.List(I), 1, 10)) = txtPRODNAO.Text Then cboPRODNAO.ListIndex = I
   Next I
    
   If cboPRODNAO.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtPRODNAO.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtProduto_GotFocus()
    objBLBFunc.SelecionaCampos txtProduto.Name, frmCADFORNEC
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtProduto_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtProduto.Text)) = 0 Then Exit Sub
    
   cboProduto.ListIndex = -1
   For I = 0 To (cboProduto.ListCount - 1)
       If Trim(Mid(cboProduto.List(I), 1, 10)) = txtProduto.Text Then cboProduto.ListIndex = I
   Next I
    
   If cboProduto.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtProduto.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtRazaoSoc_GotFocus()
    objBLBFunc.SelecionaCampos txtRazaoSoc.Name, frmCADFORNEC
End Sub

Private Sub txtRazaoSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub IncGrid()

   Dim I As Integer
   
   If (Len(Trim(txtProduto.Text)) = 0) Or (cboProduto.ListIndex = -1) Then
      MsgBox "Informe o código do produto !!!", vbOKOnly + vbCritical, "aviso"
      txtProduto.SetFocus
      Exit Sub
   End If
   
   For I = 1 To (flxProduto.Rows - 1)
       If Trim(flxProduto.TextMatrix(I, 0)) = cboProduto.ItemData(cboProduto.ListIndex) Then
          MsgBox "Este produto já esta relacionado !!!", vbOKOnly + vbCritical, "aviso"
          txtProduto.Text = ""
          cboProduto.ListIndex = -1
          txtProduto.SetFocus
          Exit Sub
       End If
   Next I
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "      PRO.* " & vbCrLf
   sSql = sSql & " from " & vbCrLf
   sSql = sSql & "      SGI_CADPRODUTO PRO" & vbCrLf
   sSql = sSql & "Where " & vbCrLf
   sSql = sSql & "      PRO.SGI_FILIAL = " & FILIAL & vbCrLf
   sSql = sSql & "  And PRO.SGI_CODIGO = '" & Trim(txtProduto.Text) & "'" & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      flxProduto.AddItem BREC!SGI_IDPRODUTO & vbTab & _
                         cboProduto.Text & vbTab & _
                         Trim(txtCodProdFornec.Text) & vbTab & _
                         IIf(Len(Trim(txtPORC.Text)) = 0, "", txtPORC.Text)
   End If
   
   flxProduto.ColAlignment(1) = 0
   
   BREC.Close
   
   txtProduto.Text = ""
   cboProduto.ListIndex = -1
   txtCodProdFornec.Text = ""
   txtPORC.Text = ""
   
   txtProduto.SetFocus

End Sub

Private Sub ConfGridProd()

    flxProduto.Rows = 1
    flxProduto.Cols = 4
    
    flxProduto.TextMatrix(0, 0) = ""
    flxProduto.TextMatrix(0, 1) = "Código/Descrição"
    flxProduto.TextMatrix(0, 2) = "Código do Fornecedor"
    flxProduto.TextMatrix(0, 3) = "Qtde. insp no Lote %"
    
    
    flxProduto.ColWidth(0) = 0
    flxProduto.ColWidth(1) = 5000
    flxProduto.ColWidth(2) = 1700
    flxProduto.ColWidth(3) = 1700
    
End Sub


Private Function ValidaCampos() As Boolean
    
    ValidaCampos = False
    
    Dim intRESP As Integer
    Dim I       As Integer
    
    If Len(Trim(txtCNPJCPF.Text)) > 0 Then
       
       If Not IsNumeric(txtCNPJCPF.Text) Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "aviso"
          txtCNPJCPF.Text = ""
          txtCNPJCPF.SetFocus
          Exit Function
       End If
       
       If Len(Trim(txtCNPJCPF.Text)) < 14 And Len(Trim(txtCNPJCPF.Text)) > 11 Then
          MsgBox "CPF/CNPJ Inválido !!!", vbOKOnly + vbCritical, "Aviso"
          txtCNPJCPF.Text = ""
          txtCNPJCPF.SetFocus
          Exit Function
       End If
       
       If Len(Trim(txtCNPJCPF.Text)) < 11 Then
          MsgBox "CPF/CNPJ Inválido !!!", vbOKOnly + vbCritical, "Aviso"
          txtCNPJCPF.Text = ""
          txtCNPJCPF.SetFocus
          Exit Function
       End If
       
       If Len(Trim(txtCNPJCPF.Text)) = 11 Then
          If objBLBFunc.ViewCPF(txtCNPJCPF.Text) = False Then
             MsgBox "CPF Inválido !!!", vbOKOnly + vbCritical, "Aviso"
             txtCNPJCPF.Text = ""
             txtCNPJCPF.SetFocus
             Exit Function
          End If
       End If
       
       If Len(Trim(txtCNPJCPF.Text)) = 14 Then
          If objBLBFunc.ViewCGC(txtCNPJCPF.Text) = False Then
             MsgBox "CNPJ Inválido !!!", vbOKOnly + vbCritical, "aviso"
             txtCNPJCPF.Text = ""
             txtCNPJCPF.SetFocus
             Exit Function
          End If
       End If
       
    End If
    
    If Len(Trim(txtRazaoSoc.Text)) = 0 Then
       MsgBox "Razão Social inválida !!!", vbOKOnly + vbCritical, "aviso"
       txtRazaoSoc.Text = ""
       txtRazaoSoc.SetFocus
       Exit Function
    End If
    
    If cTipOper = "I" Then
       
       '' Pesquisa descrição
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL   =  " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_RAZAOSOC = '" & txtRazaoSoc.Text & "'"
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       
       If Not BREC.EOF Then
          MsgBox "Está ração social já existe !!!", vbOKOnly + vbCritical, "aviso"
          txtRazaoSoc.Text = ""
          txtRazaoSoc.SetFocus
          BREC.Close
          Exit Function
       End If
       
       BREC.Close
       
       If Len(Trim(txtCNPJCPF.Text)) > 0 Then
       
          '' Pesquisa CNPJCPF
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL   =  " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CPFCNPJ  = '" & txtCNPJCPF.Text & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       
          If Not BREC.EOF Then
             MsgBox "CNPJ/CPF já existe !!!", vbOKOnly + vbCritical, "aviso"
             txtCNPJCPF.Text = ""
             txtCNPJCPF.SetFocus
             BREC.Close
             Exit Function
          End If
       
          BREC.Close
       
       End If
       
    End If
    
    If cTipOper = "A" Then
    
      If Trim(txtCNPJCPF.Text) <> Trim(objCADFORNEC.CNPJCPFFOR) Then
      
         '' Pesquisa CNPJCPF
         sSql = "Select " & vbCrLf
         sSql = sSql & "       * " & vbCrLf
         sSql = sSql & "  From " & vbCrLf
         sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_FILIAL   =  " & FILIAL & vbCrLf
         sSql = sSql & "   And SGI_CPFCNPJ  = '" & txtCNPJCPF.Text & "'"
        
         BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
         If Not BREC.EOF Then
            MsgBox "CNPJ/CPF já existe !!!", vbOKOnly + vbCritical, "aviso"
            txtCNPJCPF.Text = objCADFORNEC.CNPJCPFFOR
            txtCNPJCPF.SetFocus
            BREC.Close
            Exit Function
         End If
       
         BREC.Close
      
      End If
      
      If Trim(txtRazaoSoc.Text) <> Trim(objCADFORNEC.RAZSOCFOR) Then
      
          '' Pesquisa descrição
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL   =  " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_RAZAOSOC = '" & txtRazaoSoc.Text & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       
          If Not BREC.EOF Then
             MsgBox "Está ração social já existe !!!", vbOKOnly + vbCritical, "aviso"
             txtRazaoSoc.Text = objCADFORNEC.RAZSOCFOR
             txtRazaoSoc.SetFocus
             BREC.Close
             Exit Function
          End If
       
          BREC.Close
      
      End If
      
      If cboGRPDESP.ListIndex > -1 Then
         If cboGRPDESP.ItemData(cboGRPDESP.ListIndex) <> objCADFORNEC.GRPDESPBKP Then
            intRESP = MsgBox("O Grupo de Despesa está sendo alterado , " & vbCrLf & _
                             "o grupo de despesas do contas a pagar será alterado !!!, Continua ?", vbYesNo + vbQuestion, "Aviso")
                
            If intRESP = vbNo Then
               For I = 0 To (cboGRPDESP.ListCount - 1)
                   If objCADFORNEC.GRPDESPBKP = cboGRPDESP.ItemData(I) Then cboGRPDESP.ListIndex = I
               Next I
               Exit Function
            End If
         End If
      End If
      
    End If
    
    ValidaCampos = True
    
End Function


Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    If Len(Trim(strMODPAI)) > 0 Then cmdAltera.Enabled = True
    If Len(Trim(strMODPAI)) = 0 Then cmdAltera.Enabled = False
    
    Frame2.Enabled = True
    
    txtCNPJCPF.Enabled = False
    txtRazaoSoc.Enabled = False
    txtNomFantasia.Enabled = False
    txtEndereco.Enabled = False
    txtCidade.Enabled = False
    txtBairro.Enabled = False
    txtCEP.Enabled = False
    cboEstado.Enabled = False
    mskDtCadastro.Enabled = False
    txtIQF.Enabled = False
    
    cboTel.Locked = True
    cboContato.Locked = True
    cboEmail.Locked = True
    cboSite.Locked = True
    
    Frame3.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame8.Enabled = True
    Frame11.Enabled = False
    Frame15.Enabled = False
    Frame17.Enabled = False
    Frame18.Enabled = False
    Frame19.Enabled = False
    Frame21.Enabled = False
    
    Me.Caption = "Cadastro de fornecedores - [ CONSULTA ]"
    
    objCADFORNEC.PreenchComboBancos cboBanco
    objBLBFunc.LimpaCampos frmCADFORNEC
    
    objBLBFunc.Preenche_Estado cboEstado
    objBLBFunc.Preenche_Estado cboEstReti
    objCADFORNEC.MODULO = Me.Name
    
    optREQQUALID(0).Value = True
    
    objCADFORNEC.PreencheComboProd cboProduto
    objCADFORNEC.PreencheComboNatOper cboNATOPERCAO
    objCADFORNEC.PreenchComboTransportadoras cboTRANSP
    objCADFORNEC.PreenchComboGrpDespesas cboGRPDESP
    objCADFORNEC.PreenchComboNivelRiscos cboNivelRisco
    objCADFORNEC.PreencheComboProd cboPRODNAO
    objBLBFunc.PreenchComboStatus cboStatus
    
    ConfGridProd
    ConfGridBancos
    ConfGridCota
    ConfGridPedido
    ConfGridNatOper
    ConfGridTransp
    ConfGridNaoConf
    ConfGridNF
    ConfGridPrcProd
    ConfGridProdNao
    
    txtCodigo.Text = ""
    lblCorRisc.Caption = ""
    
    cboProduto.ListIndex = -1
    cboEstado.ListIndex = -1
    cboBanco.ListIndex = -1
    cboNivelRisco.ListIndex = -1
    
    objCADFORNEC.CodigoFOR = iCodigo
    
    If objCADFORNEC.Carrega_campos Then
       
       txtCodigo.Text = Str(objCADFORNEC.CodigoFOR)
       txtCNPJCPF.Text = objCADFORNEC.CNPJCPFFOR
       txtRazaoSoc.Text = objCADFORNEC.RAZSOCFOR
       txtNomFantasia.Text = objCADFORNEC.NOMFANTFOR
       txtEndereco.Text = objCADFORNEC.ENDFOR
       txtCidade.Text = objCADFORNEC.CIDFOR
       txtBairro.Text = objCADFORNEC.BAIFOR
       
       optREQQUALID(objCADFORNEC.REQINSPQUA).Value = True
       
       cboStatus.ListIndex = objCADFORNEC.STATUS
       
       For I = 0 To (cboEstado.ListCount - 1)
           If cboEstado.ItemData(I) = objCADFORNEC.ESTFOR Then cboEstado.ListIndex = I
       Next I
       
       txtCEP.Text = objCADFORNEC.CEPFOR
       mskDtCadastro.Text = Format(objCADFORNEC.DTCADFOR, "DD/MM/YYYY")
       
       txtEndRetirada.Text = objCADFORNEC.ENDRETI
       txtCidRetirada.Text = objCADFORNEC.CIDRETI
       txtEndReti.Text = objCADFORNEC.BAIRETI
       txtCepReti.Text = objCADFORNEC.CEPRETI
       txtIQF.Text = objCADFORNEC.IQF
       
       '' Nivel de Risco
       If objCADFORNEC.CODNIVRISCO > 0 Then
          txtNivelRisco.Text = objCADFORNEC.CODNIVRISCO
          lblCorRisc.BackColor = PegaCorRisco(objCADFORNEC.CODNIVRISCO)
          For I = 0 To (cboNivelRisco.ListCount - 1)
              If cboNivelRisco.ItemData(I) = objCADFORNEC.CODNIVRISCO Then cboNivelRisco.ListIndex = I
          Next I
       End If
       
       If objCADFORNEC.AVISAIGF > 0 Then txtAVISIGF.Text = Str(objCADFORNEC.AVISAIGF)
       If objCADFORNEC.LIBIGF > 0 Then txtSOMLIBIGF.Text = Str(objCADFORNEC.LIBIGF)
       If objCADFORNEC.TRAVAIGF > 0 Then txtBLOQFORNEC = Str(objCADFORNEC.TRAVAIGF)
       
       For I = 0 To (cboEstReti.ListCount - 1)
           If cboEstReti.ItemData(I) = objCADFORNEC.ESTRETI Then cboEstReti.ListIndex = I
       Next I
       
       If objCADFORNEC.GRPDESP > 0 Then
          For I = 0 To (cboGRPDESP.ListCount - 1)
              If cboGRPDESP.ItemData(I) = objCADFORNEC.GRPDESP Then cboGRPDESP.ListIndex = I
          Next I
       End If
       
       'Telefone
       arrTELEFONE = objCADFORNEC.TELFOR
       If IsArray(arrTELEFONE) = True Then
          For I = 1 To UBound(arrTELEFONE)
              cboTel.AddItem arrTELEFONE(I)
          Next I
       End If
       'Contato
       arrCONTATO = objCADFORNEC.CONTATOFOR
       If IsArray(arrCONTATO) = True Then
          For I = 1 To UBound(arrCONTATO)
              cboContato.AddItem arrCONTATO(I)
          Next I
       End If
       'E-MAIL
       arrEMAIL = objCADFORNEC.EMAILFOR
       If IsArray(arrEMAIL) = True Then
          For I = 1 To UBound(arrEMAIL)
              cboEmail.AddItem arrEMAIL(I)
          Next I
       End If
       'Site
       arrSITE = objCADFORNEC.SITEFOR
       If IsArray(arrSITE) = True Then
          For I = 1 To UBound(arrSITE)
              cboSite.AddItem arrSITE(I)
          Next I
       End If
       'Produto
       arrPRODUTO = objCADFORNEC.PRODUTO
       If IsArray(arrPRODUTO) = True Then
          For I = 1 To UBound(arrPRODUTO)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
              sSql = sSql & "  Where " & vbCrLf
              sSql = sSql & "        SGI_FILIAL    =  " & FILIAL & vbCrLf
              sSql = sSql & "    AND SGI_IDPRODUTO = " & Trim(arrPRODUTO(I, 1))
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxProduto.AddItem Trim(arrPRODUTO(I, 1)) & vbTab & _
                                                      Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO) & vbTab & _
                                                      Trim(arrPRODUTO(I, 2)) & vbTab & _
                                                      Trim(arrPRODUTO(I, 3))

              BREC.Close
              
          Next I
       End If
       '' Bancos
       arrBANCOS = objCADFORNEC.BANCOS
       If IsArray(arrBANCOS) = True Then
          For I = 1 To UBound(arrBANCOS)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrBANCOS(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxDadCobr.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_AGENCIA & vbTab & BREC!SGI_CC
              BREC.Close
              
          Next I
       End If
       '' Natureza de Operação
       arrNATOPER = objCADFORNEC.NATOPER
       If IsArray(arrNATOPER) = True Then
          For I = 1 To UBound(arrNATOPER)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = '" & arrNATOPER(I) & "'"
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxFISCAL.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
              BREC.Close
              
          Next I
       End If
       '' Transportadoras
       arrTRANSP = objCADFORNEC.TRANSP
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
       '' Qualidade
       PopGridQualidade
       
       'Produto Não Inspecionado
       arrPRODNAOINSP = objCADFORNEC.PRODNAOINSP
       PopProdNaoInsp
       
    End If
    
    stFornec.Tab = 0
    Call Desabilita_Tables
   
End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    mskDtCadastro.Enabled = False
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    Frame8.Enabled = True
    Frame11.Enabled = True
    Frame15.Enabled = True
    Frame17.Enabled = True
    Frame18.Enabled = True
    Frame19.Enabled = True
    Frame21.Enabled = True
    
    Me.Caption = "Cadastro de fornecedores - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFORNEC
    
    objBLBFunc.Preenche_Estado cboEstado
    objBLBFunc.Preenche_Estado cboEstReti
    objCADFORNEC.MODULO = Me.Name
    
    objCADFORNEC.PreenchComboBancos cboBanco
    objCADFORNEC.PreencheComboProd cboProduto
    objCADFORNEC.PreencheComboNatOper cboNATOPERCAO
    objCADFORNEC.PreenchComboTransportadoras cboTRANSP
    objCADFORNEC.PreenchComboGrpDespesas cboGRPDESP
    objCADFORNEC.PreenchComboNivelRiscos cboNivelRisco
    objCADFORNEC.PreencheComboProd cboPRODNAO
    objBLBFunc.PreenchComboStatus cboStatus
    
    cboStatus.ListIndex = objCADFORNEC.STATUS
    
    optREQQUALID(0).Value = True
    
    ConfGridProd
    ConfGridBancos
    ConfGridCota
    ConfGridPedido
    ConfGridNatOper
    ConfGridTransp
    ConfGridNaoConf
    ConfGridNF
    ConfGridPrcProd
    ConfGridProdNao
    
    txtCodigo.Text = ""
    
    cboProduto.ListIndex = -1
    cboEstado.ListIndex = -1
    cboBanco.ListIndex = -1
    cboNivelRisco.ListIndex = -1
    lblCorRisc.Caption = ""
    
    objCADFORNEC.CodigoFOR = iCodigo
    
    If objCADFORNEC.Carrega_campos Then
       
       txtCodigo.Text = Str(objCADFORNEC.CodigoFOR)
       txtCNPJCPF.Text = objCADFORNEC.CNPJCPFFOR
       txtRazaoSoc.Text = objCADFORNEC.RAZSOCFOR
       txtNomFantasia.Text = objCADFORNEC.NOMFANTFOR
       txtEndereco.Text = objCADFORNEC.ENDFOR
       txtCidade.Text = objCADFORNEC.CIDFOR
       txtBairro.Text = objCADFORNEC.BAIFOR
       txtIQF.Text = objCADFORNEC.IQF
       
       optREQQUALID(objCADFORNEC.REQINSPQUA).Value = True
       
       If objCADFORNEC.AVISAIGF > 0 Then txtAVISIGF.Text = Str(objCADFORNEC.AVISAIGF)
       If objCADFORNEC.LIBIGF > 0 Then txtSOMLIBIGF.Text = Str(objCADFORNEC.LIBIGF)
       If objCADFORNEC.TRAVAIGF > 0 Then txtBLOQFORNEC = Str(objCADFORNEC.TRAVAIGF)
       
       For I = 0 To (cboEstado.ListCount - 1)
           If cboEstado.ItemData(I) = objCADFORNEC.ESTFOR Then cboEstado.ListIndex = I
       Next I
       
       txtCEP.Text = objCADFORNEC.CEPFOR
       mskDtCadastro.Text = Format(objCADFORNEC.DTCADFOR, "DD/MM/YYYY")
       
       txtEndRetirada.Text = objCADFORNEC.ENDRETI
       txtCidRetirada.Text = objCADFORNEC.CIDRETI
       txtEndReti.Text = objCADFORNEC.BAIRETI
       txtCepReti.Text = objCADFORNEC.CEPRETI
       
       '' Nivel de Risco
       If objCADFORNEC.CODNIVRISCO > 0 Then
          txtNivelRisco.Text = objCADFORNEC.CODNIVRISCO
          lblCorRisc.BackColor = PegaCorRisco(objCADFORNEC.CODNIVRISCO)
          For I = 0 To (cboNivelRisco.ListCount - 1)
              If cboNivelRisco.ItemData(I) = objCADFORNEC.CODNIVRISCO Then cboNivelRisco.ListIndex = I
          Next I
       End If
       
       
       If objCADFORNEC.GRPDESP > 0 Then
          For I = 0 To (cboGRPDESP.ListCount - 1)
              If cboGRPDESP.ItemData(I) = objCADFORNEC.GRPDESP Then cboGRPDESP.ListIndex = I
          Next I
       End If
       
       For I = 0 To (cboEstReti.ListCount - 1)
           If cboEstReti.ItemData(I) = objCADFORNEC.ESTRETI Then cboEstReti.ListIndex = I
       Next I
       
       'Telefone
       arrTELEFONE = objCADFORNEC.TELFOR
       If IsArray(arrTELEFONE) = True Then
          For I = 1 To UBound(arrTELEFONE)
              cboTel.AddItem arrTELEFONE(I)
          Next I
       End If
       'Contato
       arrCONTATO = objCADFORNEC.CONTATOFOR
       If IsArray(arrCONTATO) = True Then
          For I = 1 To UBound(arrCONTATO)
              cboContato.AddItem arrCONTATO(I)
          Next I
       End If
       'E-MAIL
       arrEMAIL = objCADFORNEC.EMAILFOR
       If IsArray(arrEMAIL) = True Then
          For I = 1 To UBound(arrEMAIL)
              cboEmail.AddItem arrEMAIL(I)
          Next I
       End If
       'Site
       arrSITE = objCADFORNEC.SITEFOR
       If IsArray(arrSITE) = True Then
          For I = 1 To UBound(arrSITE)
              cboSite.AddItem arrSITE(I)
          Next I
       End If
       'Produto
       arrPRODUTO = objCADFORNEC.PRODUTO
       If IsArray(arrPRODUTO) = True Then
          For I = 1 To UBound(arrPRODUTO)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
              sSql = sSql & "  Where " & vbCrLf
              sSql = sSql & "        SGI_FILIAL    =  " & FILIAL & vbCrLf
              sSql = sSql & "    AND SGI_IDPRODUTO =  " & Trim(arrPRODUTO(I, 1))
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxProduto.AddItem Trim(arrPRODUTO(I, 1)) & vbTab & _
                                                      Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO) & vbTab & _
                                                      Trim(arrPRODUTO(I, 2)) & vbTab & _
                                                      Trim(arrPRODUTO(I, 3))
              BREC.Close
              
          Next I
       End If
       arrBANCOS = objCADFORNEC.BANCOS
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
              If Not BREC.EOF Then flxDadCobr.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_AGENCIA & vbTab & BREC!SGI_CC
              BREC.Close
              
          Next I
       End If
       '' Natureza de Operação
       arrNATOPER = objCADFORNEC.NATOPER
       If IsArray(arrNATOPER) = True Then
          For I = 1 To UBound(arrNATOPER)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = '" & arrNATOPER(I) & "'"
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxFISCAL.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
              BREC.Close
              
          Next I
       End If
       '' Transportadoras
       arrTRANSP = objCADFORNEC.TRANSP
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
       '' Qualidade
       PopGridQualidade
   
       'Produto Não Inspecionado
       arrPRODNAOINSP = objCADFORNEC.PRODNAOINSP
       PopProdNaoInsp
   
    End If
    
    stFornec.Tab = 0
    Call Desabilita_Tables
   
End Sub

Private Sub ConfGridBancos()

    flxDadCobr.Rows = 1
    flxDadCobr.Cols = 5
    
    flxDadCobr.TextMatrix(0, 0) = ""
    flxDadCobr.TextMatrix(0, 1) = "Código"
    flxDadCobr.TextMatrix(0, 2) = "Banco"
    flxDadCobr.TextMatrix(0, 3) = "Agência"
    flxDadCobr.TextMatrix(0, 4) = "C/C"
    
    flxDadCobr.ColWidth(0) = 0
    flxDadCobr.ColWidth(1) = 700
    flxDadCobr.ColWidth(2) = 3000
    flxDadCobr.ColWidth(3) = 1500
    flxDadCobr.ColWidth(4) = 1500
    
End Sub

Private Sub InclBancos()

    Dim I As Integer
    
    If Len(Trim(txtBanco.Text)) = 0 Or cboBanco.ListIndex = -1 Then
       MsgBox "Informe o banco !!!", vbOKOnly + vbCritical, "aviso"
       txtBanco.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxDadCobr.Rows - 1)
        If txtBanco.Text = flxDadCobr.TextMatrix(I, 1) Then
           MsgBox "Este banco já foi incluso !!!", vbOKOnly + vbCritical, "aviso"
           txtBanco.Text = ""
           cboBanco.ListIndex = -1
           txtBanco.SetFocus
           Exit Sub
        End If
    Next I
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & txtBanco.Text
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then flxDadCobr.AddItem "" & vbTab & txtBanco.Text & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_AGENCIA & vbTab & BREC!SGI_CC
    BREC.Close
    
    cboBanco.ListIndex = -1
    txtBanco.Text = ""
    txtBanco.SetFocus
    
End Sub


Private Sub ConfGridCota()

    FlxCotacoes.Rows = 1
    FlxCotacoes.Cols = 4
    
    FlxCotacoes.TextMatrix(0, 0) = ""
    FlxCotacoes.TextMatrix(0, 1) = "Nº Cotação"
    FlxCotacoes.TextMatrix(0, 2) = "Data"
    FlxCotacoes.TextMatrix(0, 3) = "Status"
    
    FlxCotacoes.ColWidth(0) = 0
    FlxCotacoes.ColWidth(1) = 1000
    FlxCotacoes.ColWidth(2) = 1000
    FlxCotacoes.ColWidth(3) = 1000
    
End Sub

Private Sub CarregaGridCitacao()

    ConfGridCota
    ConfGridPedido
    ConfGridNF
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ITENS.SGI_CODIGO  " & vbCrLf
    sSql = sSql & "      ,CABEC.SGI_DATA    " & vbCrLf
    sSql = sSql & "      ,CABEC.SGI_STATUS  " & vbCrLf
    sSql = sSql & "      ,ITENS.SGI_CODPED  " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_COTAITENS ITENS  " & vbCrLf
    sSql = sSql & "      ,SGI_COTAHEADER CABEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ITENS.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And ITENS.SGI_CODFOR  = " & txtCodigo.Text & vbCrLf
    sSql = sSql & "   And ITENS.SGI_PRODUTO = '" & flxProduto.TextMatrix(flxProduto.Row, 0) & "'" & vbCrLf
    sSql = sSql & "   And CABEC.SGI_CODIGO  = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & " Group By " & vbCrLf
    sSql = sSql & "          ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "         ,CABEC.SGI_DATA   " & vbCrLf
    sSql = sSql & "         ,CABEC.SGI_STATUS " & vbCrLf
    sSql = sSql & "         ,ITENS.SGI_CODPED  " & vbCrLf
    sSql = sSql & " Order by CABEC.SGI_DATA,ITENS.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       FlxCotacoes.AddItem IIf(Not IsNull(BREC!SGI_CODPED), BREC!SGI_CODPED, "") & vbTab & BREC!SGI_CODIGO & vbTab & Format(BREC!SGI_DATA, "DD/MM/YYYY") & vbTab & BREC!SGI_STATUS
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Private Sub ConfGridPedido()

    flxPedido.Clear
    flxPedido.Rows = 1
    flxPedido.Cols = 4
    
    flxPedido.TextMatrix(0, 0) = ""
    flxPedido.TextMatrix(0, 1) = "Nº Pedido"
    flxPedido.TextMatrix(0, 2) = "Data"
    flxPedido.TextMatrix(0, 3) = "Status"
    
    flxPedido.ColWidth(0) = 0
    flxPedido.ColWidth(1) = 1000
    flxPedido.ColWidth(2) = 1000
    flxPedido.ColWidth(3) = 1000
    
End Sub


Private Sub CarregaGridPedidos()
    
    ConfGridPedido
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_PEDIDOHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & FlxCotacoes.TextMatrix(FlxCotacoes.Row, 0)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxPedido.AddItem BREC!SGI_CODIGO & vbTab & BREC!SGI_CODIGO & vbTab & Format(BREC!SGI_DATAPEDIDO, "DD/MM/YYYY") & vbTab & IIf(BREC!SGI_STATUS = "A", "ABERTO", "")
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Private Sub ConfGridNatOper()

    flxFISCAL.Rows = 1
    flxFISCAL.Cols = 3
    
    flxFISCAL.TextMatrix(0, 0) = ""
    flxFISCAL.TextMatrix(0, 1) = "Código"
    flxFISCAL.TextMatrix(0, 2) = "Descrição"
    
    flxFISCAL.ColWidth(0) = 0
    flxFISCAL.ColWidth(1) = 1000
    flxFISCAL.ColWidth(2) = 5000

End Sub


Private Sub IncGridNatOper()

   Dim I As Integer
   
   If (Len(Trim(txtCODNATOPER.Text)) = 0) Or (cboNATOPERCAO.ListIndex = -1) Then
      MsgBox "Informe o código do produto !!!", vbOKOnly + vbExclamation, "aviso"
      txtCODNATOPER.SetFocus
      Exit Sub
   End If
      
   For I = 1 To (flxFISCAL.Rows - 1)
       If flxFISCAL.TextMatrix(I, 0) = txtCODNATOPER.Text Then
          MsgBox "Esta natureza já esta relacionada !!!", vbOKOnly + vbExclamation, "aviso"
          txtCODNATOPER.Text = ""
          cboNATOPERCAO.ListIndex = -1
          txtCODNATOPER.SetFocus
          Exit Sub
       End If
   Next I
   
   flxFISCAL.AddItem "" & vbTab & _
                     txtCODNATOPER.Text & vbTab & _
                     Trim(Mid(cboNATOPERCAO.Text, 8, 50))
   
   txtCODNATOPER.Text = ""
   cboNATOPERCAO.ListIndex = -1
   
   txtCODNATOPER.SetFocus

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

Private Sub PopGridQualidade()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       CADNA.* " & vbCrLf
    sSql = sSql & "      ,CABEC.SGI_DATLCTO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_NFENTRADACABEC CABEC" & vbCrLf
    sSql = sSql & "      ,SGI_NFNAOCONF      NAOCO" & vbCrLf
    sSql = sSql & "      ,SGI_CADNAOCONF     CADNA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CABEC.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And CABEC.SGI_CODFORNEC = " & objCADFORNEC.CodigoFOR & vbCrLf
    sSql = sSql & "   And NAOCO.SGI_FILIAL    = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And NAOCO.SGI_CODIGO    = CABEC.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And CADNA.SGI_FILIAL    = NAOCO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CADNA.SGI_CODIGO    = NAOCO.SGI_CODNAOCONF"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       flxQualidade.AddItem "" & vbTab & _
                            BREC!SGI_CODIGO & vbTab & _
                            BREC!SGI_DESCRICAO & vbTab & _
                            Format(BREC!SGI_DATLCTO, "DD/MM/YYYY")
    
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Private Sub ConfGridNaoConf()

    flxQualidade.Rows = 1
    flxQualidade.Cols = 4
    
    flxQualidade.TextMatrix(0, 0) = ""
    flxQualidade.TextMatrix(0, 1) = "Código"
    flxQualidade.TextMatrix(0, 2) = "Descrição"
    flxQualidade.TextMatrix(0, 3) = "Data"
    
    flxQualidade.ColWidth(0) = 0
    flxQualidade.ColWidth(1) = 1000
    flxQualidade.ColWidth(2) = 3000
    flxQualidade.ColWidth(3) = 1000
    
End Sub

Private Sub ConfGridNF()

    flxNF.Rows = 1
    flxNF.Cols = 3
    flxNF.Cols = 4
    
    flxNF.TextMatrix(0, 0) = ""
    flxNF.TextMatrix(0, 1) = "Código NF"
    flxNF.TextMatrix(0, 2) = "Série"
    flxNF.TextMatrix(0, 3) = "Data"
    
    flxNF.ColWidth(0) = 0
    flxNF.ColWidth(1) = 1000
    flxNF.ColWidth(2) = 1000
    flxNF.ColWidth(3) = 1000
    
End Sub

Private Sub txtSOMLIBIGF_Validate(Cancel As Boolean)

    If Len(Trim(txtSOMLIBIGF.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtSOMLIBIGF.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtSOMLIBIGF.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If CInt(txtSOMLIBIGF.Text) < 0 Then
       MsgBox "Somente é permitido dados positivos !!!", vbOKOnly + vbExclamation, "Aviso"
       txtSOMLIBIGF.Text = ""
       Cancel = True
       Exit Sub
    End If
    If CInt(txtSOMLIBIGF.Text) > 100 Then
       MsgBox "Somente é permitido IGF até 100 !!!", vbOKOnly + vbExclamation, "Aviso"
       txtSOMLIBIGF.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If Len(Trim(txtAVISIGF.Text)) > 0 Then
       If (CInt(txtSOMLIBIGF.Text) >= CInt(txtAVISIGF.Text)) Then
          MsgBox "aviso de IGF não pode ser meior que valor acima !!!", vbOKOnly + vbExclamation, "Aviso"
          Cancel = True
          Exit Sub
       End If
    End If
    If Len(Trim(txtBLOQFORNEC.Text)) > 0 Then
       If (CInt(txtSOMLIBIGF.Text) <= CInt(txtBLOQFORNEC.Text)) Then
          MsgBox "aviso de IGF não pode ser menor que valor abaixo !!!", vbOKOnly + vbExclamation, "Aviso"
          Cancel = True
          Exit Sub
       End If
    End If
    

End Sub


Private Sub CarregaGridNf()
    
    ConfGridNF
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_NFENTRADACABEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED = " & flxPedido.TextMatrix(flxPedido.Row, 0)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxNF.AddItem "" & vbTab & BREC!SGI_CODNF & vbTab & BREC!SGI_SERIE & vbTab & Format(BREC!SGI_DTEMISS, "DD/MM/YYYY")
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Private Sub ConfGridPrcProd()

    
    flxPrcProd.Rows = 1
    flxPrcProd.Cols = 5
    
    flxPrcProd.TextMatrix(0, 0) = ""
    flxPrcProd.TextMatrix(0, 1) = "Data Nf"
    flxPrcProd.TextMatrix(0, 2) = "Cod. Nf"
    flxPrcProd.TextMatrix(0, 3) = "Preço"
    flxPrcProd.TextMatrix(0, 4) = "Qtde."
    
    
    flxPrcProd.ColWidth(0) = 0
    flxPrcProd.ColWidth(1) = 1000
    flxPrcProd.ColWidth(2) = 1000
    flxPrcProd.ColWidth(3) = 1000
    flxPrcProd.ColWidth(3) = 1000
    
End Sub

Private Sub CarregaGridPrecos(strCODPRODUTO As String)
   
   ConfGridPrcProd
   
   '' Histórico de Preços
   sSql = "Select " & vbCrLf
   sSql = sSql & "       HIST.SGI_DTPRC " & vbCrLf
   sSql = sSql & "      ,HIST.SGI_CODNF " & vbCrLf
   sSql = sSql & "      ,HIST.SGI_PRCCOMPRA " & vbCrLf
   sSql = sSql & "      ,ITEN.SGI_QTD " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_HISTPRC        HIST " & vbCrLf
   sSql = sSql & "      ,SGI_NFENTRADACABEC NF " & vbCrLf
   sSql = sSql & "      ,SGI_NFENTRADAITENS ITEN " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       HIST.SGI_FILIAL  = " & FILIAL & vbCrLf
   sSql = sSql & "   And HIST.SGI_CODFORN = " & objCADFORNEC.CodigoFOR & vbCrLf
   sSql = sSql & "   And HIST.SGI_CODPROD = '" & Trim(strCODPRODUTO) & "'" & vbCrLf
   sSql = sSql & "   And NF.SGI_FILIAL    = HIST.SGI_FILIAL  " & vbCrLf
   sSql = sSql & "   And NF.SGI_CODFORNEC = HIST.SGI_CODFORN " & vbCrLf
   sSql = sSql & "   And NF.SGI_CODNF     = HIST.SGI_CODNF " & vbCrLf
   sSql = sSql & "   And ITEN.SGI_FILIAL  = NF.SGI_FILIAL " & vbCrLf
   sSql = sSql & "   And ITEN.SGI_CODIGO  = NF.SGI_CODIGO " & vbCrLf
   sSql = sSql & "   And ITEN.SGI_PRODUTO = HIST.SGI_CODPROD " & vbCrLf
   sSql = sSql & " Group By HIST.SGI_DTPRC,HIST.SGI_CODNF,HIST.SGI_PRCCOMPRA,ITEN.SGI_QTD " & vbCrLf
   sSql = sSql & " Order By HIST.SGI_DTPRC DESC,HIST.SGI_PRCCOMPRA DESC "
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   Do While Not BREC.EOF
      
      flxPrcProd.AddItem "" & vbTab & _
                         Format(BREC!SGI_DTPRC, "DD/MM/YYYY") & vbTab & _
                         Trim(BREC!SGI_CODNF) & vbTab & _
                         Format(BREC!SGI_PRCCOMPRA, "#,##0.00") & vbTab & _
                         Format(BREC!SGI_QTD, "#,###0.000")
      
      BREC.MoveNext
   Loop
   BREC.Close

End Sub

Private Function PegaCorRisco(lngCodNivel As Long) As String

    PegaCorRisco = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODCOR " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADRISCO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodNivel
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaCorRisco = BREC!SGI_CODCOR
    BREC.Close
    
End Function

Private Sub ConfGridProdNao()

    flxProdNaoRecebInsp.Rows = 1
    flxProdNaoRecebInsp.Cols = 2
    
    flxProdNaoRecebInsp.TextMatrix(0, 0) = ""
    flxProdNaoRecebInsp.TextMatrix(0, 1) = "Código/Descrição"
    
    
    flxProdNaoRecebInsp.ColWidth(0) = 0
    flxProdNaoRecebInsp.ColWidth(1) = 6000
    
End Sub

Private Sub IncGridProdNao()

   Dim I As Integer
   
   If (Len(Trim(txtPRODNAO.Text)) = 0) Or (cboPRODNAO.ListIndex = -1) Then
      MsgBox "Informe o código do produto !!!", vbOKOnly + vbCritical, "aviso"
      txtPRODNAO.SetFocus
      Exit Sub
   End If
   
   For I = 1 To (flxProdNaoRecebInsp.Rows - 1)
       If flxProdNaoRecebInsp.TextMatrix(I, 0) = txtPRODNAO.Text Then
          MsgBox "Este produto já esta relacionado !!!", vbOKOnly + vbCritical, "aviso"
          txtPRODNAO.Text = ""
          cboPRODNAO.ListIndex = -1
          txtPRODNAO.SetFocus
          Exit Sub
       End If
   Next I
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "      PRO.* " & vbCrLf
   sSql = sSql & " from " & vbCrLf
   sSql = sSql & "      SGI_CADPRODUTO PRO" & vbCrLf
   sSql = sSql & "Where " & vbCrLf
   sSql = sSql & "      PRO.SGI_FILIAL = " & FILIAL & vbCrLf
   sSql = sSql & "  And PRO.SGI_CODIGO = '" & Trim(txtPRODNAO.Text) & "'" & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      flxProdNaoRecebInsp.AddItem txtPRODNAO.Text & vbTab & _
                         cboPRODNAO.Text
   End If
   
   flxProdNaoRecebInsp.ColAlignment(1) = 0
   
   BREC.Close
   
   txtPRODNAO.Text = ""
   cboPRODNAO.ListIndex = -1
   txtPRODNAO.SetFocus

End Sub

Private Sub PopProdNaoInsp()
       
       Dim I As Integer
       
       If IsArray(arrPRODNAOINSP) = True Then
          For I = 1 To UBound(arrPRODNAOINSP)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
              sSql = sSql & "  Where " & vbCrLf
              sSql = sSql & "        SGI_FILIAL =  " & FILIAL & vbCrLf
              sSql = sSql & "    AND SGI_CODIGO = '" & Trim(arrPRODNAOINSP(I)) & "'"
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxProdNaoRecebInsp.AddItem Trim(arrPRODNAOINSP(I)) & vbTab & _
                                                               Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)

              BREC.Close
              
          Next I
          
          flxProdNaoRecebInsp.ColAlignment(1) = 0
          
       End If

End Sub

Private Sub Desabilita_Tables()
    stFornec.TabVisible(1) = False
    stFornec.TabVisible(2) = False
    stFornec.TabVisible(3) = False
    stFornec.TabVisible(4) = False
    stFornec.TabVisible(5) = False
    stFornec.TabVisible(6) = False
    stFornec.TabVisible(7) = False
    stFornec.TabVisible(8) = False
End Sub
