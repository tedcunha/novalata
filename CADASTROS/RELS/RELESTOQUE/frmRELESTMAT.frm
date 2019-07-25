VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRELESTMAT 
   Caption         =   "Relatório de Estoque de Produtos"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8625
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stRels 
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      Tab             =   7
      TabsPerRow      =   7
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
      TabCaption(0)   =   "Básico"
      TabPicture(0)   =   "frmRELESTMAT.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(3)=   "Frame4"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Lista de Material"
      TabPicture(1)   =   "frmRELESTMAT.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame29"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Grupo"
      TabPicture(2)   =   "frmRELESTMAT.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(2)=   "Frame8"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Sub-Grupo"
      TabPicture(3)   =   "frmRELESTMAT.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame13"
      Tab(3).Control(1)=   "Frame12"
      Tab(3).Control(2)=   "Frame11"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Espécie"
      TabPicture(4)   =   "frmRELESTMAT.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame16"
      Tab(4).Control(1)=   "Frame15"
      Tab(4).Control(2)=   "Frame14"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Tipo"
      TabPicture(5)   =   "frmRELESTMAT.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame19"
      Tab(5).Control(1)=   "Frame18"
      Tab(5).Control(2)=   "Frame17"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Esp. Técnica"
      TabPicture(6)   =   "frmRELESTMAT.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame22"
      Tab(6).Control(1)=   "Frame21"
      Tab(6).Control(2)=   "Frame20"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Fornecedores"
      TabPicture(7)   =   "frmRELESTMAT.frx":00C4
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "Frame23"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame24"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Frame25"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "Processos"
      TabPicture(8)   =   "frmRELESTMAT.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame28"
      Tab(8).Control(1)=   "Frame27"
      Tab(8).Control(2)=   "Frame26"
      Tab(8).ControlCount=   3
      TabCaption(9)   =   "Padrões"
      TabPicture(9)   =   "frmRELESTMAT.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      Begin VB.Frame Frame29 
         Caption         =   "[ Com Processos ]"
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
         TabIndex        =   104
         Top             =   2040
         Width           =   2775
         Begin VB.OptionButton optComProc 
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
            Index           =   1
            Left            =   1200
            TabIndex        =   106
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optComProc 
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
            Index           =   0
            Left            =   240
            TabIndex        =   105
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame28 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   101
         Top             =   2400
         Width           =   8055
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   14
            Left            =   1680
            TabIndex        =   102
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "[ Processo Final ]"
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
         Left            =   -74880
         TabIndex        =   94
         Top             =   1560
         Width           =   8055
         Begin VB.TextBox txtCADPROCESSOFIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   100
            Text            =   "txtCADPROC"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboProcessoFIN 
            Height          =   315
            Left            =   1680
            TabIndex        =   99
            Text            =   "cboProcesso"
            Top             =   240
            Width           =   5775
         End
         Begin VB.CommandButton Command15 
            Height          =   315
            Left            =   1320
            Picture         =   "frmRELESTMAT.frx":0118
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "[ Processo Inicial ]"
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
         Left            =   -74880
         TabIndex        =   93
         Top             =   720
         Width           =   8055
         Begin VB.TextBox txtCADPROCESSOINI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   97
            Text            =   "txtCADPROC"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboProcessoINI 
            Height          =   315
            Left            =   1680
            TabIndex        =   96
            Text            =   "cboProcesso"
            Top             =   240
            Width           =   5775
         End
         Begin VB.CommandButton Command14 
            Height          =   315
            Left            =   1320
            Picture         =   "frmRELESTMAT.frx":021A
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame25 
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
         Height          =   615
         Left            =   120
         TabIndex        =   84
         Top             =   2400
         Width           =   7935
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   13
            Left            =   1680
            TabIndex        =   86
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "[ Fornecedor Final ]"
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
         TabIndex        =   83
         Top             =   1560
         Width           =   7935
         Begin VB.TextBox txtCODFORNECFIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   92
            Text            =   "txtCODFORNEC"
            Top             =   255
            Width           =   735
         End
         Begin VB.ComboBox cboFornecFIN 
            Height          =   315
            Left            =   1200
            TabIndex        =   91
            Text            =   "cboFornec"
            Top             =   255
            Width           =   5535
         End
         Begin VB.CommandButton Command13 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":031C
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "[ Fornecedor Inicial ]"
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
         TabIndex        =   82
         Top             =   720
         Width           =   7935
         Begin VB.TextBox txtCODFORNECINI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   89
            Text            =   "txtCODFORNEC"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboFornecINI 
            Height          =   315
            Left            =   1200
            TabIndex        =   88
            Text            =   "cboFornec"
            Top             =   255
            Width           =   5535
         End
         Begin VB.CommandButton cmdPesqFor 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":041E
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame22 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   73
         Top             =   2400
         Width           =   8055
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   10
            Left            =   1680
            TabIndex        =   74
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "[ Especificação Final ]"
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
         Left            =   -74880
         TabIndex        =   72
         Top             =   1560
         Width           =   8055
         Begin VB.TextBox txtCODESPTECFIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   81
            Text            =   "txtCODESPT"
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboESPTECFIN 
            Height          =   315
            Left            =   1560
            TabIndex        =   80
            Text            =   "cboESPTEC"
            Top             =   240
            Width           =   6375
         End
         Begin VB.CommandButton Command12 
            Height          =   315
            Left            =   1200
            Picture         =   "frmRELESTMAT.frx":0520
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "[ Especificação Inicial ]"
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
         Left            =   -74880
         TabIndex        =   71
         Top             =   720
         Width           =   8055
         Begin VB.TextBox txtCODESPTECINI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   78
            Text            =   "txtCODESPT"
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboESPTECINI 
            Height          =   315
            Left            =   1560
            TabIndex        =   77
            Text            =   "cboESPTEC"
            Top             =   240
            Width           =   6375
         End
         Begin VB.CommandButton Command11 
            Height          =   315
            Left            =   1200
            Picture         =   "frmRELESTMAT.frx":0622
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame19 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   62
         Top             =   2400
         Width           =   7935
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   9
            Left            =   1680
            TabIndex        =   64
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "[ Tipo Final ]"
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
         Left            =   -74880
         TabIndex        =   61
         Top             =   1560
         Width           =   7935
         Begin VB.TextBox txtTipofin 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   70
            Text            =   "txtTipo"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboTipofin 
            Height          =   315
            Left            =   1200
            TabIndex        =   69
            Text            =   "cboTipo"
            Top             =   240
            Width           =   4455
         End
         Begin VB.CommandButton Command10 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":0724
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "[ Tipo Inicial ]"
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
         Left            =   -74880
         TabIndex        =   60
         Top             =   720
         Width           =   7935
         Begin VB.TextBox txtTipoini 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   67
            Text            =   "txtTipo"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboTipoini 
            Height          =   315
            Left            =   1200
            TabIndex        =   66
            Text            =   "cboTipo"
            Top             =   240
            Width           =   4455
         End
         Begin VB.CommandButton Command9 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":0826
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame16 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   51
         Top             =   2400
         Width           =   7935
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   6
            Left            =   1680
            TabIndex        =   52
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "[ Espécie Final ]"
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
         Left            =   -74880
         TabIndex        =   50
         Top             =   1560
         Width           =   7935
         Begin VB.ComboBox cboEspeciefin 
            Height          =   315
            Left            =   1200
            TabIndex        =   59
            Text            =   "cboEspecie"
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox txtEspeciefin 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   58
            Text            =   "txtEspecie"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":0928
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "[ Espécie Inicial ]"
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
         Left            =   -74880
         TabIndex        =   49
         Top             =   720
         Width           =   7935
         Begin VB.ComboBox cboEspecieini 
            Height          =   315
            Left            =   1200
            TabIndex        =   56
            Text            =   "cboEspecie"
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox txtEspecieini 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   55
            Text            =   "txtEspecie"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command7 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":0A2A
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame13 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   46
         Top             =   2400
         Width           =   7935
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   48
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "[ Sub-Grupo Final ]"
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
         Left            =   -74880
         TabIndex        =   42
         Top             =   1560
         Width           =   7935
         Begin VB.ComboBox cboSUBGRUPOFIN 
            Height          =   315
            Left            =   1200
            TabIndex        =   45
            Text            =   "cboSUBGRUPO"
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox txtCODSUBGRUPFIN 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   44
            Text            =   "txtCODSUBG"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":0B2C
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "[ Sub-Grupo Inicial ]"
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   720
         Width           =   7935
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":0C2E
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtCODSUBGRUPINI 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   40
            Text            =   "txtCODSUBG"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboSUBGRUPOINI 
            Height          =   315
            Left            =   1200
            TabIndex        =   39
            Text            =   "cboSUBGRUPO"
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame10 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   35
         Top             =   2280
         Width           =   8055
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   36
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "[ Grupo Final ]"
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
         Left            =   -74880
         TabIndex        =   31
         Top             =   1560
         Width           =   8055
         Begin VB.ComboBox cboCODGRUPFIN 
            Height          =   315
            Left            =   1200
            TabIndex        =   34
            Text            =   "cboCODGRUP"
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox txtCODGRUPFIN 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   33
            Text            =   "txtCODGRUP"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":0D30
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "[ Grupo Inicial ]"
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   720
         Width           =   8055
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   840
            Picture         =   "frmRELESTMAT.frx":0E32
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtCODGRUPINI 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   29
            Text            =   "txtCODGRUP"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboCODGRUPINI 
            Height          =   315
            Left            =   1200
            TabIndex        =   28
            Text            =   "cboCODGRUP"
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame7 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   24
         Top             =   2760
         Width           =   7935
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   26
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdem 
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "[ Produto - Final ]"
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
         Left            =   -74880
         TabIndex        =   20
         Top             =   2040
         Width           =   7935
         Begin VB.ComboBox cboProdutoFim 
            Height          =   315
            Left            =   1920
            TabIndex        =   23
            Text            =   "cboProduto"
            Top             =   240
            Width           =   5895
         End
         Begin VB.TextBox txtProdutoFim 
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   22
            Text            =   "txtProduto"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   1560
            Picture         =   "frmRELESTMAT.frx":0F34
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "[ Produto - Inicial ]"
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
         Left            =   -74880
         TabIndex        =   16
         Top             =   1320
         Width           =   7935
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   1560
            Picture         =   "frmRELESTMAT.frx":1036
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtProdutoIni 
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "txtProduto"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cboProdutoIni 
            Height          =   315
            Left            =   1920
            TabIndex        =   17
            Text            =   "cboProduto"
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   12
         Top             =   720
         Width           =   7935
         Begin VB.OptionButton optEstProduto 
            Caption         =   "Peças"
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
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optEstProduto 
            Caption         =   "Normalizados"
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
            Height          =   255
            Index           =   4
            Left            =   1680
            TabIndex        =   14
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optEstProduto 
            Caption         =   "Aparelhos"
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
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   8
         Top             =   720
         Width           =   8055
         Begin VB.OptionButton optEstProduto 
            Caption         =   "Aparelhos"
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
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optEstProduto 
            Caption         =   "Normalizados"
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
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optEstProduto 
            Caption         =   "Peças"
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
            Index           =   2
            Left            =   3600
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "[ Produto ]"
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
         Left            =   -74880
         TabIndex        =   4
         Top             =   1320
         Width           =   8055
         Begin VB.ComboBox cboProduto 
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Text            =   "cboProduto"
            Top             =   240
            Width           =   6015
         End
         Begin VB.TextBox txtProduto 
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "txtProduto"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdPesq 
            Height          =   315
            Left            =   1560
            Picture         =   "frmRELESTMAT.frx":1138
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
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
         Picture         =   "frmRELESTMAT.frx":123A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
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
         Left            =   960
         Picture         =   "frmRELESTMAT.frx":133C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELESTMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Dim objBLBFunc      As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object
Dim objRELESTMAT    As Object
''Dim cCamRel         As String

Dim arrPRODPAI      As Variant '' Dados do Pai
Dim arrNIVEL01_PROD As Variant '' Filho 1
Dim arrNIVEL02_PROD As Variant '' Filho 2
Dim arrNIVEL03_PROD As Variant '' Filho 3
Dim arrNIVEL04_PROD As Variant '' Filho 4
Dim arrNIVEL05_PROD As Variant '' Filho 5
Dim arrNIVEL06_PROD As Variant '' Filho 6
Dim arrNIVEL07_PROD As Variant '' Filho 7
Dim arrNIVEL08_PROD As Variant '' Filho 8
Dim arrNIVEL09_PROD As Variant '' Filho 9

Dim arrPROCPAI      As Variant '' Processos do Pai

Dim strRESULTADO    As String

Dim strCABEC2       As String
Dim strCABEC3       As String

Dim I                As Integer
Dim intQTDREG        As Integer
Dim intQTDREG2       As Integer
Dim intQTDREG3       As Integer
Dim intQTDREG4       As Integer
Dim intQTDREG5       As Integer
Dim intQTDREG6       As Integer
Dim intQTDREG7       As Integer
Dim intQTDREG8       As Integer
Dim intQTDREG9       As Integer
Dim intQTDREG10      As Integer

Dim intQTDREGPROC1   As Integer
Dim intQTDREGPROC2   As Integer
Dim intQTDREGPROC3   As Integer
Dim intQTDREGPROC4   As Integer
Dim intQTDREGPROC5   As Integer
Dim intQTDREGPROC6   As Integer
Dim intQTDREGPROC7   As Integer
Dim intQTDREGPROC8   As Integer
Dim intQTDREGPROC9   As Integer

Dim arrNIVEL0PROC1_PROD As Variant '' Processo Filho 1
Dim arrNIVEL0PROC2_PROD As Variant '' Processo Filho 2
Dim arrNIVEL0PROC3_PROD As Variant '' Processo Filho 3
Dim arrNIVEL0PROC4_PROD As Variant '' Processo Filho 4
Dim arrNIVEL0PROC5_PROD As Variant '' Processo Filho 5
Dim arrNIVEL0PROC6_PROD As Variant '' Processo Filho 6
Dim arrNIVEL0PROC7_PROD As Variant '' Processo Filho 7
Dim arrNIVEL0PROC8_PROD As Variant '' Processo Filho 8
Dim arrNIVEL0PROC9_PROD As Variant '' Processo Filho 9


Private Sub cboCODGRUPFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCODGRUPFIN, KeyAscii
End Sub

Private Sub cboCODGRUPFIN_Validate(Cancel As Boolean)
    If cboCODGRUPFIN.ListIndex > -1 Then txtCODGRUPFIN.Text = cboCODGRUPFIN.ItemData(cboCODGRUPFIN.ListIndex)
End Sub

Private Sub cboCODGRUPINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCODGRUPINI, KeyAscii
End Sub

Private Sub cboCODGRUPINI_Validate(Cancel As Boolean)
    If cboCODGRUPINI.ListIndex > -1 Then txtCODGRUPINI.Text = cboCODGRUPINI.ItemData(cboCODGRUPINI.ListIndex)
End Sub

Private Sub cboEspeciefin_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboEspecieini, KeyAscii
End Sub

Private Sub cboEspeciefin_Validate(Cancel As Boolean)
    If cboEspeciefin.ListIndex > -1 Then txtEspeciefin.Text = cboEspeciefin.ItemData(cboEspeciefin.ListIndex)
End Sub

Private Sub cboEspecieini_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboEspeciefin, KeyAscii
End Sub

Private Sub cboEspecieini_Validate(Cancel As Boolean)
    If cboEspecieini.ListIndex > -1 Then txtEspecieini.Text = cboEspecieini.ItemData(cboEspecieini.ListIndex)
End Sub

Private Sub cboESPTECFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboESPTECFIN, KeyAscii
End Sub

Private Sub cboESPTECFIN_Validate(Cancel As Boolean)
    If cboESPTECFIN.ListIndex > -1 Then txtCODESPTECFIN.Text = Str(cboESPTECFIN.ItemData(cboESPTECFIN.ListIndex))
End Sub

Private Sub cboESPTECINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboESPTECINI, KeyAscii
End Sub

Private Sub cboESPTECINI_Validate(Cancel As Boolean)
    If cboESPTECINI.ListIndex > -1 Then txtCODESPTECINI.Text = Str(cboESPTECINI.ItemData(cboESPTECINI.ListIndex))
End Sub

Private Sub cboFornecFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornecFIN, KeyAscii
End Sub

Private Sub cboFornecFIN_Validate(Cancel As Boolean)
    If cboFornecFIN.ListIndex > -1 Then txtCODFORNECFIN.Text = cboFornecFIN.ItemData(cboFornecFIN.ListIndex)
End Sub

Private Sub cboFornecINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornecINI, KeyAscii
End Sub

Private Sub cboFornecINI_Validate(Cancel As Boolean)
    If cboFornecINI.ListIndex > -1 Then txtCODFORNECINI.Text = cboFornecINI.ItemData(cboFornecINI.ListIndex)
End Sub

Private Sub cboProcessoFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProcessoFIN, KeyAscii
End Sub

Private Sub cboProcessoFIN_Validate(Cancel As Boolean)
    If cboProcessoFIN.ListIndex > -1 Then txtCADPROCESSOFIN.Text = Str(cboProcessoFIN.ItemData(cboProcessoFIN.ListIndex))
End Sub

Private Sub cboProcessoINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProcessoINI, KeyAscii
End Sub

Private Sub cboProcessoINI_Validate(Cancel As Boolean)
    If cboProcessoINI.ListIndex > -1 Then txtCADPROCESSOINI.Text = Str(cboProcessoINI.ItemData(cboProcessoINI.ListIndex))
End Sub

Private Sub cboProduto_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProduto, KeyAscii
End Sub

Private Sub cboProduto_Validate(Cancel As Boolean)
    If cboProduto.ListIndex > -1 Then txtProduto.Text = Trim(Mid(cboProduto.Text, 1, 10))
End Sub

Private Sub cboProdutoFim_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProdutoFim, KeyAscii
End Sub

Private Sub cboProdutoFim_Validate(Cancel As Boolean)
    If cboProdutoFim.ListIndex > -1 Then txtProdutoFim.Text = Trim(Mid(cboProdutoFim.Text, 1, 10))
End Sub

Private Sub cboProdutoIni_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProdutoIni, KeyAscii
End Sub

Private Sub cboProdutoIni_Validate(Cancel As Boolean)
    If cboProdutoIni.ListIndex > -1 Then txtProdutoIni.Text = Trim(Mid(cboProdutoIni.Text, 1, 10))
End Sub

Private Sub cboSUBGRUPOFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSUBGRUPOFIN, KeyAscii
End Sub

Private Sub cboSUBGRUPOFIN_Validate(Cancel As Boolean)
    If cboSUBGRUPOFIN.ListIndex > -1 Then txtCODSUBGRUPFIN.Text = cboSUBGRUPOFIN.ItemData(cboSUBGRUPOFIN.ListIndex)
End Sub

Private Sub cboSUBGRUPOINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSUBGRUPOINI, KeyAscii
End Sub

Private Sub cboSUBGRUPOINI_Validate(Cancel As Boolean)
    If cboSUBGRUPOINI.ListIndex > -1 Then txtCODSUBGRUPINI.Text = cboSUBGRUPOINI.ItemData(cboSUBGRUPOINI.ListIndex)
End Sub
Private Sub cboTipofin_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTipofin, KeyAscii
End Sub

Private Sub cboTipofin_Validate(Cancel As Boolean)
    If cboTipofin.ListIndex > -1 Then txtTipofin.Text = cboTipofin.ItemData(cboTipofin.ListIndex)
End Sub

Private Sub cboTipoini_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTipoini, KeyAscii
End Sub

Private Sub cboTipoini_Validate(Cancel As Boolean)
    If cboTipoini.ListIndex > -1 Then txtTipoini.Text = cboTipoini.ItemData(cboTipoini.ListIndex)
End Sub

Private Sub cmdImpressao_Click()
    If stRels.Tab = 0 Then ImpProdBasico
    If stRels.Tab = 1 And optComProc(1).Value = True Then ImpListaMat
    If stRels.Tab = 1 And optComProc(0).Value = True Then ImpListaMatProc
    If stRels.Tab = 2 Then ImpRelGrupProd
    If stRels.Tab = 3 Then ImpRelSubGrupProd
    If stRels.Tab = 4 Then ImpRelEspProd
    If stRels.Tab = 5 Then ImpRelTipoProd
    If stRels.Tab = 6 Then ImpRelEspTecProd
    If stRels.Tab = 7 Then ImpRelFornProd
    If stRels.Tab = 8 Then ImpRelProcProd
End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    
    If optEstProduto(0).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 0"
    If optEstProduto(1).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 1"
    If optEstProduto(2).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 2"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtProduto.Text = varRETORNO
    
    cboProduto.ListIndex = -1
    txtProduto.SetFocus

End Sub

Private Sub cmdPesqFor_Click()

    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    
    arrCAMPOS(4, 1) = "SGI_NOMFANTA"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Fantasia"
    arrCAMPOS(4, 4) = "2000"
    arrCAMPOS(4, 5) = "SGI_NOMFANTA"
    
    arrCAMPOS(5, 1) = "SGI_CIDADE"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Cidade"
    arrCAMPOS(5, 4) = "1500"
    arrCAMPOS(5, 5) = "SGI_CIDADE"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNECINI.Text = varRETORNO
    
    cboFornecINI.ListIndex = -1
    txtCODFORNECINI.SetFocus


End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboEspeciefin, KeyAscii
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    
    If optEstProduto(3).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 0"
    If optEstProduto(4).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 1"
    If optEstProduto(5).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 2"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtProdutoIni.Text = varRETORNO
    
    cboProdutoIni.ListIndex = -1
    txtProdutoIni.SetFocus

End Sub

Private Sub Command10_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPPROD " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "tipos de produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtTipofin.Text = varRETORNO
    
    cboTipofin.ListIndex = -1
    txtTipofin.SetFocus

End Sub

Private Sub Command11_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADESPTEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCESPTEC"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "6000"
    arrCAMPOS(2, 5) = "SGI_DESCESPTEC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Especificação técnica")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODESPTECINI.Text = varRETORNO
    
    cboESPTECINI.ListIndex = -1
    txtCODESPTECINI.SetFocus

End Sub

Private Sub Command12_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADESPTEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCESPTEC"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "6000"
    arrCAMPOS(2, 5) = "SGI_DESCESPTEC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Especificação técnica")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODESPTECFIN.Text = varRETORNO
    
    cboESPTECFIN.ListIndex = -1
    txtCODESPTECFIN.SetFocus

End Sub

Private Sub Command13_Click()

    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    
    arrCAMPOS(4, 1) = "SGI_NOMFANTA"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Fantasia"
    arrCAMPOS(4, 4) = "2000"
    arrCAMPOS(4, 5) = "SGI_NOMFANTA"
    
    arrCAMPOS(5, 1) = "SGI_CIDADE"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Cidade"
    arrCAMPOS(5, 4) = "1500"
    arrCAMPOS(5, 5) = "SGI_CIDADE"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNECFIN.Text = varRETORNO
    
    cboFornecFIN.ListIndex = -1
    txtCODFORNECFIN.SetFocus

End Sub

Private Sub Command14_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "6000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Processos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCADPROCESSOINI.Text = varRETORNO
    
    cboProcessoINI.ListIndex = -1
    txtCADPROCESSOINI.SetFocus

End Sub

Private Sub Command15_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "6000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Processos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCADPROCESSOFIN.Text = varRETORNO
    
    cboProcessoFIN.ListIndex = -1
    txtCADPROCESSOFIN.SetFocus

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    
    If optEstProduto(3).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 0"
    If optEstProduto(4).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 1"
    If optEstProduto(5).Value = True Then sSql = sSql & "   And SGI_PRODUTOESTILO = 2"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtProdutoFim.Text = varRETORNO
    
    cboProdutoFim.ListIndex = -1
    txtProdutoFim.SetFocus

End Sub

Private Sub Command3_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPROD " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODGRUPFIN.Text = varRETORNO
    
    cboCODGRUPFIN.ListIndex = -1
    txtCODGRUPFIN.SetFocus

End Sub

Private Sub Command4_Click()

        ReDim arrCAMPOS(1 To 2, 1 To 5) As String
        ReDim arrTABELA(1 To 1) As String
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_CODIGO     " & vbCrLf
        sSql = sSql & "      ,SGI_DESCRICAO  " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "      SGI_CADSUBGRPROD" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    
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
    
        varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Sub Grupo de produtos")
    
        If Len(Trim(varRETORNO)) > 0 Then txtCODSUBGRUPFIN.Text = varRETORNO
    
        cboSUBGRUPOFIN.ListIndex = -1
        txtCODSUBGRUPFIN.SetFocus

End Sub

Private Sub Command5_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPROD " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODGRUPINI.Text = varRETORNO
    
    cboCODGRUPINI.ListIndex = -1
    txtCODGRUPINI.SetFocus

End Sub

Private Sub Command6_Click()

        ReDim arrCAMPOS(1 To 2, 1 To 5) As String
        ReDim arrTABELA(1 To 1) As String
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_CODIGO     " & vbCrLf
        sSql = sSql & "      ,SGI_DESCRICAO  " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "      SGI_CADSUBGRPROD" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    
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
    
        varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Sub Grupo de produtos")
    
        If Len(Trim(varRETORNO)) > 0 Then txtCODSUBGRUPINI.Text = varRETORNO
    
        cboSUBGRUPOINI.ListIndex = -1
        txtCODSUBGRUPINI.SetFocus

End Sub

Private Sub Command7_Click()

       ReDim arrCAMPOS(1 To 2, 1 To 5) As String
       ReDim arrTABELA(1 To 1) As String
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       SGI_CODIGO    " & vbCrLf
       sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADESPPROD " & vbCrLf
       sSql = sSql & "Where " & vbCrLf
       sSql = sSql & "       SGI_CODIGO = " & FILIAL
    
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
    
       varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Especie de produtos")
    
       If Len(Trim(varRETORNO)) > 0 Then txtEspecieini.Text = varRETORNO
    
       cboEspecieini.ListIndex = -1
       txtEspecieini.SetFocus

End Sub

Private Sub Command8_Click()

       ReDim arrCAMPOS(1 To 2, 1 To 5) As String
       ReDim arrTABELA(1 To 1) As String
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       SGI_CODIGO    " & vbCrLf
       sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADESPPROD " & vbCrLf
       sSql = sSql & "Where " & vbCrLf
       sSql = sSql & "       SGI_CODIGO = " & FILIAL
    
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
    
       varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Especie de produtos")
    
       If Len(Trim(varRETORNO)) > 0 Then txtEspeciefin.Text = varRETORNO
    
       cboEspeciefin.ListIndex = -1
       txtEspeciefin.SetFocus

End Sub

Private Sub Command9_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPPROD " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "tipos de produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtTipoini.Text = varRETORNO
    
    cboTipoini.ListIndex = -1
    txtTipoini.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    Set objRELESTMAT = CreateObject("RELESTOQUE.clsRELESTMAT")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    stRels.Tab = 0
    
    objRELESTMAT.FILIAL = FILIAL
    
    objBLBFunc.LimpaCampos frmRELESTMAT
    
    optOrdem(0).Value = True
    optEstProduto(3).Value = True
        
    objRELESTMAT.PreenchComboGrupProduto cboCODGRUPINI
    objRELESTMAT.PreenchComboGrupProduto cboCODGRUPFIN
    
    objRELESTMAT.PreenchComboSubGrupProduto cboSUBGRUPOINI
    objRELESTMAT.PreenchComboSubGrupProduto cboSUBGRUPOFIN
    
    objRELESTMAT.PreenchComboEspecie cboEspecieini
    objRELESTMAT.PreenchComboEspecie cboEspeciefin
    
    objRELESTMAT.PreenchComboTipo cboTipoini
    objRELESTMAT.PreenchComboTipo cboTipofin
    
    objRELESTMAT.PreenchComboEspTecnica cboESPTECINI
    objRELESTMAT.PreenchComboEspTecnica cboESPTECFIN
    
    objRELESTMAT.PreencheComboFornec cboFornecINI
    objRELESTMAT.PreencheComboFornec cboFornecFIN
    
    objRELESTMAT.PreenchComboProcesso cboProcessoINI
    objRELESTMAT.PreenchComboProcesso cboProcessoFIN
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    '' --------------------------------------
    ''cCamRel = "C:\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\ESTOQUE\"
    ''cCamRel = "\\pc6\HD\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\ESTOQUE\"
End Sub


Private Sub ImpListaMat()

    If CamposOKLstMat = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADPRODUTO.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADPRODUTO.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODIGO = '" & Trim(txtProduto.Text) & "' "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Lista de Materiais"
    strCABEC3 = "Estrutura de produtos"
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELLISTAMAT.rpt", Linha, 1, strCABEC2, strCABEC3, True

End Sub

Private Function CamposOKLstMat() As Boolean
    
    CamposOKLstMat = False
    
    If Len(Trim(txtProduto.Text)) = 0 Or cboProduto.ListIndex = -1 Then
       MsgBox "Não foi selecionado nenhum produto para impressão !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    
    CamposOKLstMat = True

End Function

Private Sub optEstProduto_Click(Index As Integer)
    cboProdutoIni.Clear
    cboProdutoFim.Clear
    cboProduto.Clear
    txtProduto.Text = ""
    txtProdutoIni.Text = ""
    txtProdutoFim.Text = ""
    cboProduto.ListIndex = -1
    cboProdutoIni.ListIndex = -1
    cboProdutoFim.ListIndex = -1
    If Index = 0 Then objRELESTMAT.PreencheComboProd cboProduto, Index
    If Index = 1 Then objRELESTMAT.PreencheComboProd cboProduto, Index
    If Index = 2 Then objRELESTMAT.PreencheComboProd cboProduto, Index
    If Index = 3 Then
       objRELESTMAT.PreencheComboProd cboProdutoIni, 0
       objRELESTMAT.PreencheComboProd cboProdutoFim, 0
    End If
    If Index = 4 Then
       objRELESTMAT.PreencheComboProd cboProdutoIni, 1
       objRELESTMAT.PreencheComboProd cboProdutoFim, 1
    End If
    If Index = 5 Then
       objRELESTMAT.PreencheComboProd cboProdutoIni, 2
       objRELESTMAT.PreencheComboProd cboProdutoFim, 2
    End If
End Sub


Private Sub stRels_Click(PreviousTab As Integer)
    If stRels.Tab = 0 Then
       optEstProduto(3).Value = True
       optOrdem(0).Value = True
    End If
    If stRels.Tab = 1 Then
       optEstProduto(0).Value = True
       optComProc(1).Value = True
    End If
    If stRels.Tab = 2 Then optOrdem(3).Value = True
    If stRels.Tab = 3 Then optOrdem(4).Value = True
    If stRels.Tab = 4 Then optOrdem(7).Value = True
    If stRels.Tab = 5 Then optOrdem(8).Value = True
    If stRels.Tab = 6 Then optOrdem(11).Value = True
    If stRels.Tab = 7 Then optOrdem(12).Value = True
    If stRels.Tab = 8 Then optOrdem(15).Value = True
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtCADPROCESSOFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCADPROCESSOFIN.Name, frmRELESTMAT
End Sub

Private Sub txtCADPROCESSOFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCADPROCESSOFIN.Text
End Sub

Private Sub txtCADPROCESSOFIN_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCADPROCESSOFIN.Text)) = 0 Then Exit Sub
    
   cboProcessoFIN.ListIndex = -1
   For I = 0 To (cboProcessoFIN.ListCount - 1)
       If cboProcessoFIN.ItemData(I) = CInt(txtCADPROCESSOFIN.Text) Then cboProcessoFIN.ListIndex = I
   Next I
    
   If cboProcessoFIN.ListIndex = -1 Then
      MsgBox "Este Processo não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtCADPROCESSOFIN.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtCADPROCESSOINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCADPROCESSOINI.Name, frmRELESTMAT
End Sub

Private Sub txtCADPROCESSOINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCADPROCESSOINI.Text
End Sub

Private Sub txtCADPROCESSOINI_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCADPROCESSOINI.Text)) = 0 Then Exit Sub
    
   cboProcessoINI.ListIndex = -1
   For I = 0 To (cboProcessoINI.ListCount - 1)
       If cboProcessoINI.ItemData(I) = CInt(txtCADPROCESSOINI.Text) Then cboProcessoINI.ListIndex = I
   Next I
    
   If cboProcessoINI.ListIndex = -1 Then
      MsgBox "Este Processo não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtCADPROCESSOINI.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtCODESPTECFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODESPTECFIN.Name, frmRELESTMAT
End Sub

Private Sub txtCODESPTECFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODESPTECFIN.Text
End Sub

Private Sub txtCODESPTECFIN_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCODESPTECFIN.Text)) = 0 Then Exit Sub
    
   cboESPTECFIN.ListIndex = -1
   For I = 0 To (cboESPTECFIN.ListCount - 1)
       If cboESPTECFIN.ItemData(I) = CInt(txtCODESPTECFIN.Text) Then cboESPTECFIN.ListIndex = I
   Next I
    
   If cboESPTECFIN.ListIndex = -1 Then
      MsgBox "Esta especificação técnica não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtCODESPTECFIN.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtCODESPTECINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODESPTECINI.Name, frmRELESTMAT
End Sub

Private Sub txtCODESPTECINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODESPTECINI.Text
End Sub

Private Sub txtCODESPTECINI_Validate(Cancel As Boolean)
   
   Dim I As Integer

   If Len(Trim(txtCODESPTECINI.Text)) = 0 Then Exit Sub
    
   cboESPTECINI.ListIndex = -1
   For I = 0 To (cboESPTECINI.ListCount - 1)
       If cboESPTECINI.ItemData(I) = CInt(txtCODESPTECINI.Text) Then cboESPTECINI.ListIndex = I
   Next I
    
   If cboESPTECINI.ListIndex = -1 Then
      MsgBox "Esta especificação técnica não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtCODESPTECINI.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtCODFORNECFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNECFIN.Name, frmRELESTMAT
End Sub

Private Sub txtCODFORNECFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFORNECFIN.Text
End Sub

Private Sub txtCODFORNECFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNECFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNECFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFornecFIN.ListIndex = -1
    For I = 0 To (cboFornecFIN.ListCount - 1)
        If cboFornecFIN.ItemData(I) = Str(Val(txtCODFORNECFIN.Text)) Then cboFornecFIN.ListIndex = I
    Next I
    
    If cboFornecFIN.ListIndex = -1 Then
       MsgBox "Esta fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODFORNECINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNECINI.Name, frmRELESTMAT
End Sub

Private Sub txtCODFORNECINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFORNECINI.Text
End Sub

Private Sub txtCODFORNECINI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNECINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNECINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFornecINI.ListIndex = -1
    For I = 0 To (cboFornecINI.ListCount - 1)
        If cboFornecINI.ItemData(I) = Str(Val(txtCODFORNECINI.Text)) Then cboFornecINI.ListIndex = I
    Next I
    
    If cboFornecINI.ListIndex = -1 Then
       MsgBox "Esta fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODGRUPFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODGRUPFIN.Name, frmRELESTMAT
End Sub

Private Sub txtCODGRUPFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODGRUPFIN.Text
End Sub

Private Sub txtCODGRUPFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODGRUPFIN.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCODGRUPFIN.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUPFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCODGRUPFIN.ListIndex = -1
    For I = 0 To (cboCODGRUPFIN.ListCount - 1)
        If cboCODGRUPFIN.ItemData(I) = Str(Val(txtCODGRUPFIN.Text)) Then cboCODGRUPFIN.ListIndex = I
    Next I
    
    If cboCODGRUPFIN.ListIndex = -1 Then
       MsgBox "Este grupo de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUPFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODGRUPINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODGRUPINI.Name, frmRELESTMAT
End Sub

Private Sub txtCODGRUPINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODGRUPINI.Text
End Sub

Private Sub txtCODGRUPINI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODGRUPINI.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCODGRUPINI.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUPINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCODGRUPINI.ListIndex = -1
    For I = 0 To (cboCODGRUPINI.ListCount - 1)
        If cboCODGRUPINI.ItemData(I) = Str(Val(txtCODGRUPINI.Text)) Then cboCODGRUPINI.ListIndex = I
    Next I
    
    If cboCODGRUPINI.ListIndex = -1 Then
       MsgBox "Este grupo de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUPINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub
Private Sub txtCODSUBGRUPFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODSUBGRUPFIN.Name, frmRELESTMAT
End Sub

Private Sub txtCODSUBGRUPFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODSUBGRUPFIN.Text
End Sub

Private Sub txtCODSUBGRUPFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODSUBGRUPFIN.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCODSUBGRUPFIN.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODSUBGRUPFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboSUBGRUPOFIN.ListIndex = -1
    For I = 0 To (cboSUBGRUPOFIN.ListCount - 1)
        If cboSUBGRUPOFIN.ItemData(I) = Str(Val(txtCODSUBGRUPFIN.Text)) Then cboSUBGRUPOFIN.ListIndex = I
    Next I
    
    If cboSUBGRUPOFIN.ListIndex = -1 Then
       MsgBox "Este sub grupo de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODSUBGRUPFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODSUBGRUPINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODSUBGRUPINI.Name, frmRELESTMAT
End Sub

Private Sub txtCODSUBGRUPINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODSUBGRUPINI.Text
End Sub

Private Sub txtCODSUBGRUPINI_Validate(Cancel As Boolean)
    
    Dim I As Integer
    
    If Len(Trim(txtCODSUBGRUPINI.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCODSUBGRUPINI.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODSUBGRUPINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboSUBGRUPOINI.ListIndex = -1
    For I = 0 To (cboSUBGRUPOINI.ListCount - 1)
        If cboSUBGRUPOINI.ItemData(I) = Str(Val(txtCODSUBGRUPINI.Text)) Then cboSUBGRUPOINI.ListIndex = I
    Next I
    
    If cboSUBGRUPOINI.ListIndex = -1 Then
       MsgBox "Este sub grupo de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODSUBGRUPINI.Text = ""
       Cancel = True
       Exit Sub
    End If
End Sub

Private Sub txtEspeciefin_GotFocus()
    objBLBFunc.SelecionaCampos txtEspeciefin.Name, frmRELESTMAT
End Sub

Private Sub txtEspeciefin_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtEspeciefin.Text
End Sub

Private Sub txtEspeciefin_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtEspeciefin.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtEspeciefin.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtEspeciefin.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboEspeciefin.ListIndex = -1
    For I = 0 To (cboEspeciefin.ListCount - 1)
        If cboEspeciefin.ItemData(I) = Str(Val(txtEspeciefin.Text)) Then cboEspeciefin.ListIndex = I
    Next I
    
    If cboEspeciefin.ListIndex = -1 Then
       MsgBox "Esta espécie de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtEspeciefin.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtEspecieini_GotFocus()
    objBLBFunc.SelecionaCampos txtEspecieini.Name, frmRELESTMAT
End Sub

Private Sub txtEspecieini_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtEspecieini.Text
End Sub

Private Sub txtEspecieini_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtEspecieini.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtEspecieini.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtEspecieini.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboEspecieini.ListIndex = -1
    For I = 0 To (cboEspecieini.ListCount - 1)
        If cboEspecieini.ItemData(I) = Str(Val(txtEspecieini.Text)) Then cboEspecieini.ListIndex = I
    Next I
    
    If cboEspecieini.ListIndex = -1 Then
       MsgBox "Esta espécie de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtEspecieini.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtProduto_GotFocus()
    objBLBFunc.SelecionaCampos txtProduto.Name, frmRELESTMAT
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtProduto_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtProduto.Text)) = 0 Then Exit Sub
    
   cboProduto.ListIndex = -1
   For I = 0 To (cboProduto.ListCount - 1)
       If Trim(Mid(cboProduto.List(I), 1, 10)) = Trim(txtProduto.Text) Then cboProduto.ListIndex = I
   Next I
    
   If cboProduto.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtProduto.Text = ""
      Cancel = True
      Exit Sub
   End If
   
End Sub
Private Sub txtProdutoFim_GotFocus()
    objBLBFunc.SelecionaCampos txtProdutoFim.Name, frmRELESTMAT
End Sub

Private Sub txtProdutoFim_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtProdutoFim_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtProdutoFim.Text)) = 0 Then Exit Sub
    
   cboProdutoFim.ListIndex = -1
   For I = 0 To (cboProdutoFim.ListCount - 1)
       If Trim(Mid(cboProdutoFim.List(I), 1, 10)) = Trim(txtProdutoFim.Text) Then cboProdutoFim.ListIndex = I
   Next I
    
   If cboProdutoFim.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtProdutoFim.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtProdutoIni_GotFocus()
   objBLBFunc.SelecionaCampos txtProdutoIni.Name, frmRELESTMAT
End Sub

Private Sub txtProdutoIni_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtProdutoIni_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtProdutoIni.Text)) = 0 Then Exit Sub
    
   cboProdutoIni.ListIndex = -1
   For I = 0 To (cboProdutoIni.ListCount - 1)
       If Trim(Mid(cboProdutoIni.List(I), 1, 10)) = Trim(txtProdutoIni.Text) Then cboProdutoIni.ListIndex = I
   Next I
    
   If cboProdutoIni.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtProdutoIni.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub ImpProdBasico()

    If CamposOKProdBasico = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADPRODUTO.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODGPROD "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODSUBGPROD "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODESPECIE "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_UNIDMEDIDA "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_PRODUTOTIPO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_PRODUTOESTILO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_PESOUNIT "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODEAN "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_SALDO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_ESTMINIMO "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADPRODUTO.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODIGO = '" & Trim(txtProdutoIni.Text) & "' "
    ElseIf Len(Trim(txtProdutoFim.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADPRODUTO.SGI_CODIGO >= '" & Trim(txtProdutoIni.Text) & "' "
       sSql = sSql & "        And SGI_CADPRODUTO.SGI_CODIGO <= '" & Trim(txtProdutoFim.Text) & "') "
    End If
    
    If optOrdem(0).Value = True Then sSql = sSql & "   Order by SGI_CADPRODUTO.SGI_CODIGO "
    If optOrdem(1).Value = True Then sSql = sSql & "   Order by SGI_CADPRODUTO.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Relatório de produtos"
    
    If optOrdem(0).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(0).Caption)
    If optOrdem(1).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(1).Caption)
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELESTBAS.rpt", Linha, 1, strCABEC2, strCABEC3, False


End Sub

Private Function CamposOKProdBasico() As Boolean
    
    CamposOKProdBasico = False
    
    If Len(Trim(txtProdutoFim.Text)) > 0 Or cboProdutoFim.ListIndex <> -1 Then
       If Len(Trim(txtProdutoIni.Text)) = 0 Or cboProdutoIni.ListIndex = -1 Then
          MsgBox "Não foi selecionado o produto inicial !!!", vbOKOnly + vbExclamation, "Aviso"
          txtProdutoFim.Text = ""
          cboProdutoFim.ListIndex = -1
          txtProdutoIni.SetFocus
          Exit Function
       End If
    End If
    
    CamposOKProdBasico = True

End Function


Private Sub ImpRelGrupProd()

    If CamposOKRelGrpProd = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADGRUPROD.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADGRUPROD.SGI_DESCRICAO "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADGRUPROD SGI_CADGRUPROD "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADGRUPROD.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODGRUPINI.Text)) > 0 And Len(Trim(txtCODGRUPFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADGRUPROD.SGI_CODIGO = " & Trim(txtCODGRUPINI.Text)
    ElseIf Len(Trim(txtCODGRUPINI.Text)) > 0 And Len(Trim(txtCODGRUPFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADGRUPROD.SGI_CODIGO >= " & Trim(txtCODGRUPINI.Text) & " And SGI_CADGRUPROD.SGI_CODIGO <= " & Trim(txtCODGRUPFIN.Text) & ") "
    End If
    
    If optOrdem(3).Value = True Then sSql = sSql & "   Order by SGI_CADGRUPROD.SGI_CODIGO "
    If optOrdem(2).Value = True Then sSql = sSql & "   Order by SGI_CADGRUPROD.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Este Grupo não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Relatório de produtos por Grupo"
    
    If optOrdem(3).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(0).Caption)
    If optOrdem(2).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(1).Caption)
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELPRODGRUP.rpt", Linha, 1, strCABEC2, strCABEC3, False

End Sub

Private Function CamposOKRelGrpProd() As Boolean
    
    CamposOKRelGrpProd = False
    
    If Len(Trim(txtCODGRUPINI.Text)) = 0 And Len(Trim(txtCODGRUPFIN.Text)) > 0 Then
       MsgBox "Escolha o grupo inicial !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCODGRUPINI.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODGRUPINI.Text)) > 0 And Len(Trim(txtCODGRUPFIN.Text)) > 0 Then
       If CLng(txtCODGRUPINI.Text) > CLng(txtCODGRUPFIN.Text) Then
          MsgBox "O Código do grupo inicial não pode ser maior que o do grupo final !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODGRUPINI.Text = ""
          txtCODGRUPINI.SetFocus
          cboCODGRUPINI.ListIndex = -1
          Exit Function
       End If
    End If
    
    CamposOKRelGrpProd = True

End Function


Private Sub ImpRelSubGrupProd()

    If CamposOKRelSubGrpProd = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADSUBGRPROD.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADSUBGRPROD.SGI_DESCRICAO "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADSUBGRPROD SGI_CADSUBGRPROD "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADSUBGRPROD.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODSUBGRUPINI.Text)) > 0 And Len(Trim(txtCODSUBGRUPFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADSUBGRPROD.SGI_CODIGO = " & Trim(txtCODSUBGRUPINI.Text)
    ElseIf Len(Trim(txtCODSUBGRUPINI.Text)) > 0 And Len(Trim(txtCODSUBGRUPFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADSUBGRPROD.SGI_CODIGO >= " & Trim(txtCODSUBGRUPINI.Text) & " And SGI_CADSUBGRPROD.SGI_CODIGO <= " & Trim(txtCODSUBGRUPFIN.Text) & ") "
    End If
    
    If optOrdem(4).Value = True Then sSql = sSql & "   Order by SGI_CADSUBGRPROD.SGI_CODIGO "
    If optOrdem(5).Value = True Then sSql = sSql & "   Order by SGI_CADSUBGRPROD.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Este Sub-Grupo não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Relatório de produtos por Sub-Grupo"
    
    If optOrdem(4).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(0).Caption)
    If optOrdem(5).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(1).Caption)
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELPRODSUBGRP.rpt", Linha, 1, strCABEC2, strCABEC3, False

End Sub

Private Function CamposOKRelEspProd() As Boolean
    
    CamposOKRelEspProd = False
    
    If Len(Trim(txtEspecieini.Text)) = 0 And Len(Trim(txtEspeciefin.Text)) > 0 Then
       MsgBox "Escolha a espécie inicial !!!", vbOKOnly + vbExclamation, "Aviso"
       txtEspecieini.SetFocus
       Exit Function
    End If
    If Len(Trim(txtEspecieini.Text)) > 0 And Len(Trim(txtEspeciefin.Text)) > 0 Then
       If CLng(txtEspecieini.Text) > CLng(txtEspeciefin.Text) Then
          MsgBox "O Código da espécie inicial não pode ser maior que o da espécie final !!!", vbOKOnly + vbExclamation, "Aviso"
          txtEspecieini.Text = ""
          txtEspecieini.SetFocus
          cboEspecieini.ListIndex = -1
          Exit Function
       End If
    End If
    
    CamposOKRelEspProd = True

End Function

Private Sub ImpRelEspProd()

    If CamposOKRelEspProd = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADESPPROD.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADESPPROD.SGI_DESCRICAO "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADESPPROD SGI_CADESPPROD "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADESPPROD.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtEspecieini.Text)) > 0 And Len(Trim(txtEspeciefin.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADESPPROD.SGI_CODIGO = " & Trim(txtEspecieini.Text)
    ElseIf Len(Trim(txtEspecieini.Text)) > 0 And Len(Trim(txtEspeciefin.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADESPPROD.SGI_CODIGO >= " & Trim(txtEspecieini.Text) & " And SGI_CADESPPROD.SGI_CODIGO <= " & Trim(txtEspeciefin.Text) & ") "
    End If
    
    If optOrdem(7).Value = True Then sSql = sSql & "   Order by SGI_CADESPPROD.SGI_CODIGO "
    If optOrdem(6).Value = True Then sSql = sSql & "   Order by SGI_CADESPPROD.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Esta Espécie não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Relatório de produtos por Espécie"
    
    If optOrdem(7).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(0).Caption)
    If optOrdem(6).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(1).Caption)
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELPRODESPECIE.rpt", Linha, 1, strCABEC2, strCABEC3, False

End Sub

Private Function CamposOKRelSubGrpProd() As Boolean
    
    CamposOKRelSubGrpProd = False
    
    If Len(Trim(txtCODSUBGRUPINI.Text)) = 0 And Len(Trim(txtCODSUBGRUPFIN.Text)) > 0 Then
       MsgBox "Escolha o Sub-grupo inicial !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCODSUBGRUPINI.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODSUBGRUPINI.Text)) > 0 And Len(Trim(txtCODSUBGRUPFIN.Text)) > 0 Then
       If CLng(txtCODSUBGRUPINI.Text) > CLng(txtCODSUBGRUPFIN.Text) Then
          MsgBox "O Código do Sub-Grupo inicial não pode ser maior que o Sub-Grupo final !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODSUBGRUPINI.Text = ""
          txtCODSUBGRUPINI.SetFocus
          cboSUBGRUPOINI.ListIndex = -1
          Exit Function
       End If
    End If
    
    CamposOKRelSubGrpProd = True

End Function

Private Sub txtTipofin_GotFocus()
    objBLBFunc.SelecionaCampos txtTipofin.Name, frmRELESTMAT
End Sub

Private Sub txtTipofin_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtTipofin.Text
End Sub

Private Sub txtTipofin_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtTipofin.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtTipofin.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtTipofin.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTipofin.ListIndex = -1
    For I = 0 To (cboTipoini.ListCount - 1)
        If cboTipofin.ItemData(I) = CInt(txtTipofin.Text) Then cboTipofin.ListIndex = I
    Next I
    
    If cboTipofin.ListIndex = -1 Then
       MsgBox "Esta tipo de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtTipofin.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtTipoini_GotFocus()
    objBLBFunc.SelecionaCampos txtTipoini.Name, frmRELESTMAT
End Sub

Private Sub txtTipoini_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtTipoini.Text
End Sub


Private Sub ImpRelTipoProd()

    If CamposOKRelTipoProd = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADTIPPROD.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADTIPPROD.SGI_DESCRICAO "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADTIPPROD SGI_CADTIPPROD "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADTIPPROD.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtTipoini.Text)) > 0 And Len(Trim(txtTipofin.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADTIPPROD.SGI_CODIGO = " & Trim(txtTipoini.Text)
    ElseIf Len(Trim(txtTipoini.Text)) > 0 And Len(Trim(txtTipofin.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADTIPPROD.SGI_CODIGO >= " & Trim(txtTipoini.Text) & " And SGI_CADTIPPROD.SGI_CODIGO <= " & Trim(txtTipofin.Text) & ") "
    End If
    
    If optOrdem(8).Value = True Then sSql = sSql & "   Order by SGI_CADTIPPROD.SGI_CODIGO "
    If optOrdem(9).Value = True Then sSql = sSql & "   Order by SGI_CADTIPPROD.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Este tipo de produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Relatório de produtos por Tipo de Produto"
    
    If optOrdem(8).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(8).Caption)
    If optOrdem(9).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(9).Caption)
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELPRODTIPO.rpt", Linha, 1, strCABEC2, strCABEC3, False

End Sub

Private Function CamposOKRelTipoProd() As Boolean
    
    CamposOKRelTipoProd = False
    
    If Len(Trim(txtEspecieini.Text)) = 0 And Len(Trim(txtEspeciefin.Text)) > 0 Then
       MsgBox "Escolha a espécie inicial !!!", vbOKOnly + vbExclamation, "Aviso"
       txtEspecieini.SetFocus
       Exit Function
    End If
    If Len(Trim(txtEspecieini.Text)) > 0 And Len(Trim(txtEspeciefin.Text)) > 0 Then
       If CLng(txtEspecieini.Text) > CLng(txtEspeciefin.Text) Then
          MsgBox "O Código da espécie inicial não pode ser maior que o da espécie final !!!", vbOKOnly + vbExclamation, "Aviso"
          txtEspecieini.Text = ""
          txtEspecieini.SetFocus
          cboEspecieini.ListIndex = -1
          Exit Function
       End If
    End If
    
    CamposOKRelTipoProd = True

End Function

Private Sub txtTipoini_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtTipoini.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtTipoini.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtTipoini.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTipoini.ListIndex = -1
    For I = 0 To (cboTipoini.ListCount - 1)
        If cboTipoini.ItemData(I) = CInt(txtTipoini.Text) Then cboTipoini.ListIndex = I
    Next I
    
    If cboTipoini.ListIndex = -1 Then
       MsgBox "Esta tipo de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtTipoini.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub ImpRelEspTecProd()

    If CamposOKRelTipoProd = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADESPTEC.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADESPTEC.SGI_DESCESPTEC "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADESPTEC SGI_CADESPTEC "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADESPTEC.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODESPTECINI.Text)) > 0 And Len(Trim(txtCODESPTECFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADESPTEC.SGI_CODIGO = " & Trim(txtCODESPTECINI.Text)
    ElseIf Len(Trim(txtCODESPTECINI.Text)) > 0 And Len(Trim(txtCODESPTECFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADESPTEC.SGI_CODIGO >= " & Trim(txtCODESPTECINI.Text) & " And SGI_CADESPTEC.SGI_CODIGO <= " & Trim(txtCODESPTECFIN.Text) & ") "
    End If
    
    If optOrdem(11).Value = True Then sSql = sSql & "   Order by SGI_CADESPTEC.SGI_CODIGO "
    If optOrdem(10).Value = True Then sSql = sSql & "   Order by SGI_CADESPTEC.SGI_DESCESPTEC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Esta Especificação Técnica de produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Relatório de produtos por Especificação técnica"
    
    If optOrdem(11).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(11).Caption)
    If optOrdem(10).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(10).Caption)
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELPRODESPTEC.rpt", Linha, 1, strCABEC2, strCABEC3, False

End Sub


Private Sub ImpRelFornProd()

    If CamposOKRelFornec = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADFORNEC.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADFORNEC.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADFORNEC.SGI_CODIGO = " & Trim(txtCODFORNECINI.Text)
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADFORNEC.SGI_CODIGO >= " & Trim(txtCODFORNECINI.Text) & " And SGI_CADFORNEC.SGI_CODIGO <= " & Trim(txtCODFORNECFIN.Text) & ") "
    End If
    
    If optOrdem(12).Value = True Then sSql = sSql & "   Order by SGI_CADFORNEC.SGI_CODIGO "
    If optOrdem(13).Value = True Then sSql = sSql & "   Order by SGI_CADFORNEC.SGI_RAZAOSOC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Este Fornecedor de produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Relatório de produtos por Fornecedor"
    
    If optOrdem(12).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(12).Caption)
    If optOrdem(13).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(13).Caption)
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELPRODFOR.rpt", Linha, 1, strCABEC2, strCABEC3, False

End Sub

Private Function CamposOKRelFornec() As Boolean
    
    CamposOKRelFornec = False
    
    If Len(Trim(txtCODFORNECINI.Text)) = 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       MsgBox "Escolha o fornecedor inicial !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCODFORNECINI.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODFORNECFIN.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       If CLng(txtCODFORNECINI.Text) > CLng(txtCODFORNECFIN.Text) Then
          MsgBox "O Código do fornecedor inicial não pode ser maior que a do fornecedor final !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODFORNECINI.Text = ""
          txtCODFORNECINI.SetFocus
          cboFornecINI.ListIndex = -1
          Exit Function
       End If
    End If
    
    CamposOKRelFornec = True

End Function


Private Sub ImpRelProcProd()

    If CamposOKRelProcesso = False Then Exit Sub
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADPROCESSO.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADPROCESSO.SGI_DESCRI "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADPROCESSO SGI_CADPROCESSO "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADPROCESSO.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCADPROCESSOINI.Text)) > 0 And Len(Trim(txtCADPROCESSOFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADPROCESSO.SGI_CODIGO = " & Trim(txtCADPROCESSOINI.Text)
    ElseIf Len(Trim(txtCADPROCESSOINI.Text)) > 0 And Len(Trim(txtCADPROCESSOFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADPROCESSO.SGI_CODIGO >= " & Trim(txtCADPROCESSOINI.Text) & " And SGI_CADPROCESSO.SGI_CODIGO <= " & Trim(txtCADPROCESSOFIN.Text) & ") "
    End If
    
    If optOrdem(15).Value = True Then sSql = sSql & "   Order by SGI_CADPROCESSO.SGI_CODIGO "
    If optOrdem(14).Value = True Then sSql = sSql & "   Order by SGI_CADPROCESSO.SGI_DESCRI "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Este Processo de produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Relatório de produtos por Processos"
    
    If optOrdem(15).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(15).Caption)
    If optOrdem(14).Value = True Then strCABEC3 = "Básico por ordem de " & Trim(optOrdem(14).Caption)
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELPRODPROCESSO.rpt", Linha, 1, strCABEC2, strCABEC3, False

End Sub

Private Function CamposOKRelProcesso() As Boolean
    
    CamposOKRelProcesso = False
    
    If Len(Trim(txtCADPROCESSOINI.Text)) = 0 And Len(Trim(txtCADPROCESSOFIN.Text)) > 0 Then
       MsgBox "Escolha o processo inicial !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCADPROCESSOINI.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCADPROCESSOINI.Text)) > 0 And Len(Trim(txtCADPROCESSOFIN.Text)) > 0 Then
       If CLng(txtCADPROCESSOINI.Text) > CLng(txtCADPROCESSOFIN.Text) Then
          MsgBox "O Código do processo inicial não pode ser maior que a do processo final !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCADPROCESSOINI.Text = ""
          txtCADPROCESSOINI.SetFocus
          cboProcessoINI.ListIndex = -1
          Exit Function
       End If
    End If
    
    CamposOKRelProcesso = True

End Function


Private Sub ImpListaMatProc()

    On Error GoTo err_grava
    
    If CamposOKLstMat = False Then Exit Sub
    
    Dim strPRODUTO As String
    Dim I02 As Integer
    Dim I03 As Integer
    Dim I04 As Integer
    Dim I05 As Integer
    Dim I06 As Integer
    Dim I07 As Integer
    Dim I08 As Integer
    Dim I09 As Integer
    Dim I10 As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.* " & vbCrLf
    sSql = sSql & "      ,UNI.SGI_UNIDADE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO " & vbCrLf
    sSql = sSql & "      ,SGI_CADUNIMED  UNI " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_CODIGO = '" & Trim(txtProduto.Text) & "' " & vbCrLf
    sSql = sSql & "   And UNI.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And UNI.SGI_CODIGO = PRO.SGI_UNIDMEDIDA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
        ReDim arrPRODPAI(1 To 1, 1 To 8) As Variant
        arrPRODPAI(1, 1) = BREC!SGI_CODIGO
        arrPRODPAI(1, 2) = BREC!SGI_DESCRICAO
        arrPRODPAI(1, 3) = 1
        arrPRODPAI(1, 4) = BREC!SGI_PRCCUSTO
        arrPRODPAI(1, 5) = BREC!SGI_PRCCUSTO
        arrPRODPAI(1, 6) = BREC!SGI_UNIDADE
        arrPRODPAI(1, 7) = BREC!SGI_CODTABCONV
        arrPRODPAI(1, 8) = "1"
        objRELESTMAT.PRODPAI = arrPRODPAI
        
        '' Pegando Processo do Produto Pai
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_CADPROPROC.* " & vbCrLf
        sSql = sSql & "      ,SGI_CADPROCESSO.SGI_DESCRI " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPROPROC  SGI_CADPROPROC" & vbCrLf
        sSql = sSql & "      ,SGI_CADPROCESSO SGI_CADPROCESSO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_CADPROPROC.SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CADPROPROC.SGI_CODPROD = '" & Trim(BREC!SGI_CODIGO) & "'" & vbCrLf
        sSql = sSql & "   And SGI_CADPROCESSO.SGI_FILIAL = SGI_CADPROPROC.SGI_FILIAL  " & vbCrLf
        sSql = sSql & "   And SGI_CADPROCESSO.SGI_CODIGO = SGI_CADPROPROC.SGI_CODPROC " & vbCrLf
       
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        intQTDREG = 0
        If Not BREC2.EOF Then
           Do While Not BREC2.EOF
              intQTDREG = intQTDREG + 1
              BREC2.MoveNext
           Loop
           
           ReDim arrPROCPAI(1 To intQTDREG, 1 To 6)
           BREC2.MoveFirst
           intQTDREG = 1
           Do While Not BREC2.EOF
              arrPROCPAI(intQTDREG, 1) = 1
              arrPROCPAI(intQTDREG, 2) = BREC2!SGI_CODPROC
              arrPROCPAI(intQTDREG, 3) = BREC2!SGI_TEMPO
              arrPROCPAI(intQTDREG, 4) = BREC2!SGI_VALORHORA
              arrPROCPAI(intQTDREG, 5) = BREC2!SGI_TOTVALOR
              arrPROCPAI(intQTDREG, 6) = BREC2!SGI_DESCRI
              
              BREC2.MoveNext
              intQTDREG = (intQTDREG + 1)
          Loop
          objRELESTMAT.PROCPAI = arrPROCPAI
        End If
        BREC2.Close
        
    End If
    BREC.Close
        
    If objRELESTMAT.Grava_ListMatTemp_ProdPai = False Then Exit Sub  '' Gravando o Pai
    arrNIVEL01_PROD = Empty
    
    '' Montando Array dos Filhos
    ''Filho 01
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.*              " & vbCrLf
    sSql = sSql & "      ,LST.SGI_PRODLST    " & vbCrLf
    sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
    sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
    sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
    sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
    sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrPRODPAI(1, 1)) & "'" & vbCrLf
    sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
   
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then
       intQTDREG = 0
       Do While Not BREC2.EOF
          intQTDREG = intQTDREG + 1
          BREC2.MoveNext
       Loop
       ReDim arrNIVEL01_PROD(1 To intQTDREG, 1 To 8)
       BREC2.MoveFirst
       intQTDREG = 1
       Do While Not BREC2.EOF
          arrNIVEL01_PROD(intQTDREG, 1) = BREC2!SGI_CODIGO
          arrNIVEL01_PROD(intQTDREG, 2) = BREC2!SGI_DESCRICAO
          arrNIVEL01_PROD(intQTDREG, 3) = BREC2!SGI_QTDE
          arrNIVEL01_PROD(intQTDREG, 4) = BREC2!SGI_CUSTOUNIT
          arrNIVEL01_PROD(intQTDREG, 5) = BREC2!SGI_CUSTOTOTAL
          arrNIVEL01_PROD(intQTDREG, 6) = BREC2!SGI_UNIDCONS
          arrNIVEL01_PROD(intQTDREG, 7) = BREC2!SGI_CODTABCONV
          arrNIVEL01_PROD(intQTDREG, 8) = "1" & Trim(Str(intQTDREG))
          
          '' ==========================================================================
          '' Pegando Processo do Nivel 01
          sSql = "Select " & vbCrLf
          sSql = sSql & "       SGI_CADPROPROC.* " & vbCrLf
          sSql = sSql & "      ,SGI_CADPROCESSO.SGI_DESCRI " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADPROPROC  SGI_CADPROPROC  " & vbCrLf
          sSql = sSql & "      ,SGI_CADPROCESSO SGI_CADPROCESSO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_CADPROPROC.SGI_FILIAL  = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CADPROPROC.SGI_CODPROD = '" & Trim(BREC2!SGI_CODIGO) & "'" & vbCrLf
          sSql = sSql & "   And SGI_CADPROCESSO.SGI_FILIAL = SGI_CADPROPROC.SGI_FILIAL  " & vbCrLf
          sSql = sSql & "   And SGI_CADPROCESSO.SGI_CODIGO = SGI_CADPROPROC.SGI_CODPROC " & vbCrLf
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          intQTDREGPROC1 = 0
          If Not BREC.EOF Then
             Do While Not BREC.EOF
                intQTDREGPROC1 = intQTDREGPROC1 + 1
                BREC.MoveNext
             Loop
           
             ReDim arrNIVEL0PROC1_PROD(1 To intQTDREGPROC1, 1 To 6)
             BREC.MoveFirst
             intQTDREGPROC1 = 1
             Do While Not BREC.EOF
                arrNIVEL0PROC1_PROD(intQTDREGPROC1, 1) = "1" & Trim(Str(intQTDREG))
                arrNIVEL0PROC1_PROD(intQTDREGPROC1, 2) = BREC!SGI_CODPROC
                arrNIVEL0PROC1_PROD(intQTDREGPROC1, 3) = BREC!SGI_TEMPO
                arrNIVEL0PROC1_PROD(intQTDREGPROC1, 4) = BREC!SGI_VALORHORA
                arrNIVEL0PROC1_PROD(intQTDREGPROC1, 5) = BREC!SGI_TOTVALOR
                arrNIVEL0PROC1_PROD(intQTDREGPROC1, 6) = BREC!SGI_DESCRI
                BREC.MoveNext
                intQTDREGPROC1 = (intQTDREGPROC1 + 1)
            Loop
            objRELESTMAT.PROCPROD01 = arrNIVEL0PROC1_PROD
            
          End If
          BREC.Close
          '' ==========================================================================
          
          intQTDREG = intQTDREG + 1
          BREC2.MoveNext
       Loop
       objRELESTMAT.NIVEL01 = arrNIVEL01_PROD
    End If
    BREC2.Close
          
    If Not IsArray(arrNIVEL01_PROD) Then Exit Sub
    For I02 = 1 To UBound(arrNIVEL01_PROD)
        arrNIVEL02_PROD = Empty
        If objRELESTMAT.Grava_ListMatTemp_Filho1(I02) = False Then Exit Sub '' Gravando Filho 1
        objRELESTMAT.PROCPROD01 = Empty
        '' ==========================================================================
        '' Montando Array dos Filhos
        ''Filho 02
        sSql = "Select " & vbCrLf
        sSql = sSql & "       PRO.*              " & vbCrLf
        sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
        sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
        sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
        sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
        sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrNIVEL01_PROD(I02, 1)) & "'" & vbCrLf
        sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
       
        BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC3.EOF Then
           intQTDREG2 = 0
           Do While Not BREC3.EOF
              intQTDREG2 = intQTDREG2 + 1
              BREC3.MoveNext
           Loop
           ReDim arrNIVEL02_PROD(1 To intQTDREG2, 1 To 8)
           BREC3.MoveFirst
           intQTDREG2 = 1
           Do While Not BREC3.EOF
              arrNIVEL02_PROD(intQTDREG2, 1) = BREC3!SGI_CODIGO
              arrNIVEL02_PROD(intQTDREG2, 2) = BREC3!SGI_DESCRICAO
              arrNIVEL02_PROD(intQTDREG2, 3) = BREC3!SGI_QTDE
              arrNIVEL02_PROD(intQTDREG2, 4) = BREC3!SGI_CUSTOUNIT
              arrNIVEL02_PROD(intQTDREG2, 5) = BREC3!SGI_CUSTOTOTAL
              arrNIVEL02_PROD(intQTDREG2, 6) = BREC3!SGI_UNIDCONS
              arrNIVEL02_PROD(intQTDREG2, 7) = BREC3!SGI_CODTABCONV
              arrNIVEL02_PROD(intQTDREG2, 8) = "1" & Trim(Str(I02)) & Trim(Str(intQTDREG2))
              intQTDREG2 = intQTDREG2 + 1
              BREC3.MoveNext
           Loop
           objRELESTMAT.NIVEL02 = arrNIVEL02_PROD
        End If
        BREC3.Close
        '' ==========================================================================
        
        If IsArray(arrNIVEL02_PROD) Then
            For I03 = 1 To UBound(arrNIVEL02_PROD)
                arrNIVEL03_PROD = Empty
                If objRELESTMAT.Grava_ListMatTemp_Filho2(I03) = False Then Exit Sub '' Gravando Filho 2
            
                '' ==========================================================================
                '' Montando Array dos Filhos
                ''Filho 03
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PRO.*              " & vbCrLf
                sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
                sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
                sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
                sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
                sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
                sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrNIVEL02_PROD(I03, 1)) & "'" & vbCrLf
                sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
       
                BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC4.EOF Then
                   intQTDREG3 = 0
                   Do While Not BREC4.EOF
                      intQTDREG3 = intQTDREG3 + 1
                      BREC4.MoveNext
                   Loop
                   ReDim arrNIVEL03_PROD(1 To intQTDREG3, 1 To 8)
                   BREC4.MoveFirst
                   intQTDREG3 = 1
                   Do While Not BREC4.EOF
                      arrNIVEL03_PROD(intQTDREG3, 1) = BREC4!SGI_CODIGO
                      arrNIVEL03_PROD(intQTDREG3, 2) = BREC4!SGI_DESCRICAO
                      arrNIVEL03_PROD(intQTDREG3, 3) = BREC4!SGI_QTDE
                      arrNIVEL03_PROD(intQTDREG3, 4) = BREC4!SGI_CUSTOUNIT
                      arrNIVEL03_PROD(intQTDREG3, 5) = BREC4!SGI_CUSTOTOTAL
                      arrNIVEL03_PROD(intQTDREG3, 6) = BREC4!SGI_UNIDCONS
                      arrNIVEL03_PROD(intQTDREG3, 7) = BREC4!SGI_CODTABCONV
                      arrNIVEL03_PROD(intQTDREG3, 8) = "1" & Trim(Str(I02)) & Trim(Str(I03)) & Trim(Str(intQTDREG3))
                      intQTDREG3 = intQTDREG3 + 1
                      BREC4.MoveNext
                   Loop
                   objRELESTMAT.NIVEL03 = arrNIVEL03_PROD
                End If
                BREC4.Close
                '' ==========================================================================
                
                If IsArray(arrNIVEL03_PROD) Then
                    For I04 = 1 To UBound(arrNIVEL03_PROD)
                        arrNIVEL04_PROD = Empty
                        If objRELESTMAT.Grava_ListMatTemp_Filho3(I04) = False Then Exit Sub '' Gravando Filho 3
                        
                        '' ==========================================================================
                        '' Montando Array dos Filhos
                        ''Filho 04
                        sSql = "Select " & vbCrLf
                        sSql = sSql & "       PRO.*              " & vbCrLf
                        sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
                        sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
                        sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
                        sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
                        sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
                        sSql = sSql & "  From " & vbCrLf
                        sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
                        sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
                        sSql = sSql & " Where " & vbCrLf
                        sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
                        sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrNIVEL03_PROD(I04, 1)) & "'" & vbCrLf
                        sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
                        sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
            
                        BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
                        If Not BREC5.EOF Then
                           intQTDREG4 = 0
                           Do While Not BREC5.EOF
                              intQTDREG4 = intQTDREG4 + 1
                              BREC5.MoveNext
                           Loop
                           ReDim arrNIVEL04_PROD(1 To intQTDREG4, 1 To 8)
                           BREC5.MoveFirst
                           intQTDREG4 = 1
                           Do While Not BREC5.EOF
                              arrNIVEL04_PROD(intQTDREG4, 1) = BREC5!SGI_CODIGO
                              arrNIVEL04_PROD(intQTDREG4, 2) = BREC5!SGI_DESCRICAO
                              arrNIVEL04_PROD(intQTDREG4, 3) = BREC5!SGI_QTDE
                              arrNIVEL04_PROD(intQTDREG4, 4) = BREC5!SGI_CUSTOUNIT
                              arrNIVEL04_PROD(intQTDREG4, 5) = BREC5!SGI_CUSTOTOTAL
                              arrNIVEL04_PROD(intQTDREG4, 6) = BREC5!SGI_UNIDCONS
                              arrNIVEL04_PROD(intQTDREG4, 7) = BREC5!SGI_CODTABCONV
                              arrNIVEL04_PROD(intQTDREG4, 8) = "1" & Trim(Str(I02)) & Trim(Str(I03)) & Trim(Str(I04)) & Trim(Str(intQTDREG4))
                              intQTDREG4 = intQTDREG4 + 1
                              BREC5.MoveNext
                           Loop
                           objRELESTMAT.NIVEL04 = arrNIVEL04_PROD
                        End If
                        BREC5.Close
                        '' ==========================================================================
                
                        If IsArray(arrNIVEL04_PROD) Then
                           For I05 = 1 To UBound(arrNIVEL04_PROD)
                               arrNIVEL05_PROD = Empty
                               If objRELESTMAT.Grava_ListMatTemp_Filho4(I05) = False Then Exit Sub '' Gravando Filho 4
                               
                               '' ==========================================================================
                               '' Montando Array dos Filhos
                               '' Filho 05
                                sSql = "Select " & vbCrLf
                                sSql = sSql & "       PRO.*              " & vbCrLf
                                sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
                                sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
                                sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
                                sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
                                sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
                                sSql = sSql & "  From " & vbCrLf
                                sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
                                sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
                                sSql = sSql & " Where " & vbCrLf
                                sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
                                sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrNIVEL04_PROD(I05, 1)) & "'" & vbCrLf
                                sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
                                sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
                
                                BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
                                If Not BREC6.EOF Then
                                   intQTDREG5 = 0
                                   Do While Not BREC6.EOF
                                      intQTDREG5 = intQTDREG5 + 1
                                      BREC6.MoveNext
                                   Loop
                                   ReDim arrNIVEL05_PROD(1 To intQTDREG5, 1 To 8)
                                   BREC6.MoveFirst
                                   intQTDREG5 = 1
                                   Do While Not BREC6.EOF
                                      arrNIVEL05_PROD(intQTDREG5, 1) = BREC6!SGI_CODIGO
                                      arrNIVEL05_PROD(intQTDREG5, 2) = BREC6!SGI_DESCRICAO
                                      arrNIVEL05_PROD(intQTDREG5, 3) = BREC6!SGI_QTDE
                                      arrNIVEL05_PROD(intQTDREG5, 4) = BREC6!SGI_CUSTOUNIT
                                      arrNIVEL05_PROD(intQTDREG5, 5) = BREC6!SGI_CUSTOTOTAL
                                      arrNIVEL05_PROD(intQTDREG5, 6) = BREC6!SGI_UNIDCONS
                                      arrNIVEL05_PROD(intQTDREG5, 7) = BREC6!SGI_CODTABCONV
                                      arrNIVEL05_PROD(intQTDREG5, 8) = "1" & Trim(Str(I02)) & Trim(Str(I03)) & Trim(Str(I04)) & Trim(Str(I05)) & Trim(Str(intQTDREG5))
                                      intQTDREG5 = intQTDREG5 + 1
                                      BREC6.MoveNext
                                   Loop
                                   objRELESTMAT.NIVEL05 = arrNIVEL05_PROD
                                End If
                                BREC6.Close
                                '' ==========================================================================
                                
                                If IsArray(arrNIVEL05_PROD) Then
                                   For I06 = 1 To UBound(arrNIVEL05_PROD)
                                       arrNIVEL06_PROD = Empty
                                       If objRELESTMAT.Grava_ListMatTemp_Filho5(I06) = False Then Exit Sub '' Gravando Filho 5
                                       
                                       '' ==========================================================================
                                       '' Montando Array dos Filhos
                                       '' Filho 06
                                        sSql = "Select " & vbCrLf
                                        sSql = sSql & "       PRO.*              " & vbCrLf
                                        sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
                                        sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
                                        sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
                                        sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
                                        sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
                                        sSql = sSql & "  From " & vbCrLf
                                        sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
                                        sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
                                        sSql = sSql & " Where " & vbCrLf
                                        sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
                                        sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrNIVEL05_PROD(I06, 1)) & "'" & vbCrLf
                                        sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
                                        sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
                        
                                        BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
                                        If Not BREC7.EOF Then
                                           intQTDREG6 = 0
                                           Do While Not BREC7.EOF
                                              intQTDREG6 = intQTDREG6 + 1
                                              BREC7.MoveNext
                                           Loop
                                           ReDim arrNIVEL06_PROD(1 To intQTDREG6, 1 To 8)
                                           BREC7.MoveFirst
                                           intQTDREG6 = 1
                                           Do While Not BREC7.EOF
                                              arrNIVEL06_PROD(intQTDREG6, 1) = BREC7!SGI_CODIGO
                                              arrNIVEL06_PROD(intQTDREG6, 2) = BREC7!SGI_DESCRICAO
                                              arrNIVEL06_PROD(intQTDREG6, 3) = BREC7!SGI_QTDE
                                              arrNIVEL06_PROD(intQTDREG6, 4) = BREC7!SGI_CUSTOUNIT
                                              arrNIVEL06_PROD(intQTDREG6, 5) = BREC7!SGI_CUSTOTOTAL
                                              arrNIVEL06_PROD(intQTDREG6, 6) = BREC7!SGI_UNIDCONS
                                              arrNIVEL06_PROD(intQTDREG6, 7) = BREC7!SGI_CODTABCONV
                                              arrNIVEL06_PROD(intQTDREG6, 8) = "1" & Trim(Str(I02)) & Trim(Str(I03)) & Trim(Str(I04)) & Trim(Str(I05)) & Trim(Str(I06)) & Trim(Str(intQTDREG6))
                                              intQTDREG6 = intQTDREG6 + 1
                                              BREC7.MoveNext
                                           Loop
                                           objRELESTMAT.NIVEL06 = arrNIVEL06_PROD
                                        End If
                                        BREC7.Close
                                        '' ==========================================================================
                                        
                                        If IsArray(arrNIVEL06_PROD) Then
                                           For I07 = 1 To UBound(arrNIVEL06_PROD)
                                               arrNIVEL07_PROD = Empty
                                               If objRELESTMAT.Grava_ListMatTemp_Filho6(I07) = False Then Exit Sub '' Gravando Filho 6
                                               
                                               '' ==========================================================================
                                               '' Montando Array dos Filhos
                                               '' Filho 07
                                                sSql = "Select " & vbCrLf
                                                sSql = sSql & "       PRO.*              " & vbCrLf
                                                sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
                                                sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
                                                sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
                                                sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
                                                sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
                                                sSql = sSql & "  From " & vbCrLf
                                                sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
                                                sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
                                                sSql = sSql & " Where " & vbCrLf
                                                sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
                                                sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrNIVEL06_PROD(I07, 1)) & "'" & vbCrLf
                                                sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
                                                sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
                                
                                                BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
                                                If Not BREC8.EOF Then
                                                   intQTDREG7 = 0
                                                   Do While Not BREC7.EOF
                                                      intQTDREG7 = intQTDREG7 + 1
                                                      BREC8.MoveNext
                                                   Loop
                                                   ReDim arrNIVEL07_PROD(1 To intQTDREG7, 1 To 8)
                                                   BREC8.MoveFirst
                                                   intQTDREG7 = 1
                                                   Do While Not BREC8.EOF
                                                      arrNIVEL07_PROD(intQTDREG7, 1) = BREC8!SGI_CODIGO
                                                      arrNIVEL07_PROD(intQTDREG7, 2) = BREC8!SGI_DESCRICAO
                                                      arrNIVEL07_PROD(intQTDREG7, 3) = BREC8!SGI_QTDE
                                                      arrNIVEL07_PROD(intQTDREG7, 4) = BREC8!SGI_CUSTOUNIT
                                                      arrNIVEL07_PROD(intQTDREG7, 5) = BREC8!SGI_CUSTOTOTAL
                                                      arrNIVEL07_PROD(intQTDREG7, 6) = BREC8!SGI_UNIDCONS
                                                      arrNIVEL07_PROD(intQTDREG7, 7) = BREC8!SGI_CODTABCONV
                                                      arrNIVEL07_PROD(intQTDREG7, 8) = "1" & Trim(Str(I02)) & Trim(Str(I03)) & Trim(Str(I04)) & Trim(Str(I05)) & Trim(Str(I06)) & Trim(Str(I07)) & Trim(Str(intQTDREG7))
                                                      intQTDREG7 = intQTDREG7 + 1
                                                      BREC8.MoveNext
                                                   Loop
                                                   objRELESTMAT.NIVEL07 = arrNIVEL07_PROD
                                                End If
                                                BREC8.Close
                                                '' ==========================================================================
                                               
                                                If IsArray(arrNIVEL07_PROD) Then
                                                   For I08 = 1 To UBound(arrNIVEL07_PROD)
                                                       arrNIVEL08_PROD = Empty
                                                       If objRELESTMAT.Grava_ListMatTemp_Filho7(I08) = False Then Exit Sub '' Gravando Filho 7
                                                       
                                                       '' ==========================================================================
                                                       '' Montando Array dos Filhos
                                                       '' Filho 08
                                                        sSql = "Select " & vbCrLf
                                                        sSql = sSql & "       PRO.*              " & vbCrLf
                                                        sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
                                                        sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
                                                        sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
                                                        sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
                                                        sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
                                                        sSql = sSql & "  From " & vbCrLf
                                                        sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
                                                        sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
                                                        sSql = sSql & " Where " & vbCrLf
                                                        sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
                                                        sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrNIVEL07_PROD(I08, 1)) & "'" & vbCrLf
                                                        sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
                                                        sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
                                        
                                                        BREC9.Open sSql, adoBanco_Dados, adOpenDynamic
                                                        If Not BREC9.EOF Then
                                                           intQTDREG8 = 0
                                                           Do While Not BREC9.EOF
                                                              intQTDREG8 = intQTDREG8 + 1
                                                              BREC9.MoveNext
                                                           Loop
                                                           ReDim arrNIVEL08_PROD(1 To intQTDREG8, 1 To 8)
                                                           BREC9.MoveFirst
                                                           intQTDREG8 = 1
                                                           Do While Not BREC9.EOF
                                                              arrNIVEL08_PROD(intQTDREG8, 1) = BREC9!SGI_CODIGO
                                                              arrNIVEL08_PROD(intQTDREG8, 2) = BREC9!SGI_DESCRICAO
                                                              arrNIVEL08_PROD(intQTDREG8, 3) = BREC9!SGI_QTDE
                                                              arrNIVEL08_PROD(intQTDREG8, 4) = BREC9!SGI_CUSTOUNIT
                                                              arrNIVEL08_PROD(intQTDREG8, 5) = BREC9!SGI_CUSTOTOTAL
                                                              arrNIVEL08_PROD(intQTDREG8, 6) = BREC9!SGI_UNIDCONS
                                                              arrNIVEL08_PROD(intQTDREG8, 7) = BREC9!SGI_CODTABCONV
                                                              arrNIVEL08_PROD(intQTDREG8, 8) = "1" & Trim(Str(I02)) & Trim(Str(I03)) & Trim(Str(I04)) & Trim(Str(I05)) & Trim(Str(I06)) & Trim(Str(I07)) & Trim(Str(I08)) & Trim(Str(intQTDREG8))
                                                              intQTDREG8 = intQTDREG8 + 1
                                                              BREC9.MoveNext
                                                           Loop
                                                           objRELESTMAT.NIVEL08 = arrNIVEL08_PROD
                                                        End If
                                                        BREC9.Close
                                                        '' ==========================================================================
                                                        
                                                        If IsArray(arrNIVEL08_PROD) Then
                                                           For I09 = 1 To UBound(arrNIVEL08_PROD)
                                                               arrNIVEL09_PROD = Empty
                                                               If objRELESTMAT.Grava_ListMatTemp_Filho8(I09) = False Then Exit Sub '' Gravando Filho 8
                                                               
                                                               '' ==========================================================================
                                                               '' Montando Array dos Filhos
                                                               '' Filho 09
                                                                sSql = "Select " & vbCrLf
                                                                sSql = sSql & "       PRO.*              " & vbCrLf
                                                                sSql = sSql & "      ,LST.SGI_QTDE       " & vbCrLf
                                                                sSql = sSql & "      ,LST.SGI_CUSTOUNIT  " & vbCrLf
                                                                sSql = sSql & "      ,LST.SGI_CUSTOTOTAL " & vbCrLf
                                                                sSql = sSql & "      ,LST.SGI_UNIDCONS   " & vbCrLf
                                                                sSql = sSql & "      ,LST.SGI_CODTABCONV " & vbCrLf
                                                                sSql = sSql & "  From " & vbCrLf
                                                                sSql = sSql & "       SGI_LISTAMAT   LST" & vbCrLf
                                                                sSql = sSql & "      ,SGI_CADPRODUTO PRO" & vbCrLf
                                                                sSql = sSql & " Where " & vbCrLf
                                                                sSql = sSql & "       LST.SGI_FILIAL  = " & FILIAL & vbCrLf
                                                                sSql = sSql & "   AND LST.SGI_PRODUTO = '" & Trim(arrNIVEL08_PROD(I09, 1)) & "'" & vbCrLf
                                                                sSql = sSql & "   And PRO.SGI_FILIAL  = LST.SGI_FILIAL" & vbCrLf
                                                                sSql = sSql & "   And PRO.SGI_CODIGO  = LST.SGI_PRODLST"
                                        
                                                                BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
                                                                If Not BREC10.EOF Then
                                                                   intQTDREG9 = 0
                                                                   Do While Not BREC10.EOF
                                                                      intQTDREG9 = intQTDREG9 + 1
                                                                      BREC10.MoveNext
                                                                   Loop
                                                                   ReDim arrNIVEL09_PROD(1 To intQTDREG9, 1 To 8)
                                                                   BREC10.MoveFirst
                                                                   intQTDREG9 = 1
                                                                   Do While Not BREC10.EOF
                                                                      arrNIVEL09_PROD(intQTDREG9, 1) = BREC10!SGI_CODIGO
                                                                      arrNIVEL09_PROD(intQTDREG9, 2) = BREC10!SGI_DESCRICAO
                                                                      arrNIVEL09_PROD(intQTDREG9, 3) = BREC10!SGI_QTDE
                                                                      arrNIVEL09_PROD(intQTDREG9, 4) = BREC10!SGI_CUSTOUNIT
                                                                      arrNIVEL09_PROD(intQTDREG9, 5) = BREC10!SGI_CUSTOTOTAL
                                                                      arrNIVEL09_PROD(intQTDREG9, 6) = BREC10!SGI_UNIDCONS
                                                                      arrNIVEL09_PROD(intQTDREG9, 7) = BREC10!SGI_CODTABCONV
                                                                      arrNIVEL09_PROD(intQTDREG9, 8) = "1" & Trim(Str(I02)) & Trim(Str(I03)) & Trim(Str(I04)) & Trim(Str(I05)) & Trim(Str(I06)) & Trim(Str(I07)) & Trim(Str(I08)) & Trim(Str(I09)) & Trim(Str(intQTDREG9))
                                                                      intQTDREG9 = intQTDREG9 + 1
                                                                      BREC9.MoveNext
                                                                   Loop
                                                                   objRELESTMAT.NIVEL09 = arrNIVEL09_PROD
                                                                End If
                                                                BREC9.Close
                                                                '' ==========================================================================
                                                                
                                                                If IsArray(arrNIVEL09_PROD) Then
                                                                   For I10 = 1 To UBound(arrNIVEL09_PROD)
                                                                       If objRELESTMAT.Grava_ListMatTemp_Filho9(I10) = False Then Exit Sub '' Gravando Filho 9
                                                                   Next I10
                                                                End If
                                                           Next I09
                                                        End If
                                                   Next I08
                                                End If
                                           Next I07
                                        End If
                                   Next I06
                                End If
                           Next I05
                        End If
                    Next I04
                End If
            Next I03
        End If
    Next I02
       
    sSql = "Select * From SGI_LISTAMAT_TEMP Where SGI_FILIAL = " & FILIAL
       
    strCABEC2 = "Lista de Materiais"
    strCABEC3 = "Estrutura de produtos"
       
    Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELLISTPROC02.rpt", Linha, 1, strCABEC2, strCABEC3, False)
    
    Exit Sub
    
err_grava:

    MsgBox "Erro Nº : " & Err.Number & vbCrLf & " Descrição : " & Err.Description, vbExclamation + vbOKOnly, "Aviso"
    
    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    If BREC3.State = 1 Then BREC3.Close
    If BREC4.State = 1 Then BREC4.Close
    If BREC5.State = 1 Then BREC5.Close
    If BREC6.State = 1 Then BREC6.Close
    If BREC7.State = 1 Then BREC7.Close
    If BREC8.State = 1 Then BREC8.Close
    If BREC9.State = 1 Then BREC9.Close

End Sub

