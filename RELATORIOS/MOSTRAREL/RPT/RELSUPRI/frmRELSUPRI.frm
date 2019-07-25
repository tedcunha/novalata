VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELSUPRI 
   Caption         =   "Relatório de Suprimentos"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10395
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stCotaPedidos 
      Height          =   4815
      Left            =   0
      TabIndex        =   21
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
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
      TabCaption(0)   =   "Cotaçôes"
      TabPicture(0)   =   "frmRELSUPRI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "frmRELSUPRI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "stTabPedidos"
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   4335
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7646
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
         TabCaption(0)   =   "Emitidas"
         TabPicture(0)   =   "frmRELSUPRI.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame5(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame5(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame5(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         Begin VB.Frame Frame5 
            Caption         =   "[ Fornecedores ]"
            ForeColor       =   &H8000000D&
            Height          =   975
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   2520
            Width           =   9855
            Begin VB.CommandButton Command13 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":0054
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   600
               Width           =   375
            End
            Begin VB.ComboBox cboFornecFIN 
               Height          =   315
               Left            =   3600
               TabIndex        =   15
               Text            =   "cboFornec"
               Top             =   615
               Width           =   6135
            End
            Begin VB.TextBox txtCODFORNECFIN 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   14
               Text            =   "txtCODFORNEC"
               Top             =   615
               Width           =   1455
            End
            Begin VB.CommandButton cmdPesqFor 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":0156
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   240
               Width           =   375
            End
            Begin VB.ComboBox cboFornecINI 
               Height          =   315
               Left            =   3600
               TabIndex        =   13
               Text            =   "cboFornec"
               Top             =   255
               Width           =   6135
            End
            Begin VB.TextBox txtCODFORNECINI 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   12
               Text            =   "txtCODFORNEC"
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Fornecedor Inicial"
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
               Index           =   5
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "Fornecedor Final"
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
               Index           =   4
               Left            =   120
               TabIndex        =   34
               Top             =   600
               Width           =   1575
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Produtos ]"
            ForeColor       =   &H8000000D&
            Height          =   975
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   1560
            Width           =   9855
            Begin VB.CommandButton Command2 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":0258
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox txtProdutoFim 
               Height          =   315
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   10
               Text            =   "txtProduto"
               Top             =   600
               Width           =   1455
            End
            Begin VB.ComboBox cboProdutoFim 
               Height          =   315
               Left            =   3600
               TabIndex        =   11
               Text            =   "cboProduto"
               Top             =   600
               Width           =   6135
            End
            Begin VB.ComboBox cboProdutoIni 
               Height          =   315
               Left            =   3600
               TabIndex        =   9
               Text            =   "cboProduto"
               Top             =   240
               Width           =   6135
            End
            Begin VB.TextBox txtProdutoIni 
               Height          =   315
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   8
               Text            =   "txtProduto"
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton Command1 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":035A
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "Produto Final"
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
               Index           =   3
               Left            =   120
               TabIndex        =   32
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Produto Inicial"
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
               Index           =   2
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Data de Emissão ]"
            ForeColor       =   &H8000000D&
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   9855
            Begin MSMask.MaskEdBox mskDtFinal 
               Height          =   285
               Left            =   3720
               TabIndex        =   7
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtInicial 
               Height          =   285
               Left            =   1320
               TabIndex        =   6
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Data Final"
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
               Left            =   2640
               TabIndex        =   30
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Data Inicial"
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
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Cotações com Pedidos Emitidos ]"
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   120
            TabIndex        =   26
            Top             =   3600
            Width           =   3735
            Begin VB.OptionButton optCotPedSimNao 
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
               Height          =   255
               Index           =   4
               Left            =   2400
               TabIndex        =   75
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optCotPedSimNao 
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
               Index           =   1
               Left            =   1440
               TabIndex        =   17
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optCotPedSimNao 
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
               Index           =   0
               Left            =   360
               TabIndex        =   16
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Status da Cotação ]"
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   5400
            TabIndex        =   25
            Top             =   360
            Width           =   4575
            Begin VB.OptionButton optStatus 
               Caption         =   "Liberado"
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
               Index           =   3
               Left            =   3360
               TabIndex        =   40
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "Cotado"
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
               Index           =   2
               Left            =   2280
               TabIndex        =   5
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "Digitado"
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
               Left            =   1080
               TabIndex        =   4
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optStatus 
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
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   3
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "[ Ordem ]"
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   5175
            Begin VB.OptionButton optOrdem 
               Caption         =   "Fornecedor"
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
               Index           =   2
               Left            =   3480
               TabIndex        =   2
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton optOrdem 
               Caption         =   "Produto"
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
               TabIndex        =   1
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optOrdem 
               Caption         =   "Dt. Emissão"
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
               Left            =   120
               TabIndex        =   0
               Top             =   240
               Width           =   1335
            End
         End
      End
      Begin TabDlg.SSTab stTabPedidos 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
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
         TabCaption(0)   =   "Emitidos"
         TabPicture(0)   =   "frmRELSUPRI.frx":045C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame7"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame5(3)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame5(4)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame5(5)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame8"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Com Risco"
         TabPicture(1)   =   "frmRELSUPRI.frx":0478
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5(6)"
         Tab(1).Control(1)=   "Frame5(7)"
         Tab(1).Control(2)=   "Frame9"
         Tab(1).ControlCount=   3
         Begin VB.Frame Frame9 
            Caption         =   "[ Status do Pedido ]"
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
            TabIndex        =   91
            Top             =   2160
            Width           =   3975
            Begin VB.OptionButton optRisPEdSt 
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
               Height          =   255
               Index           =   2
               Left            =   2880
               TabIndex        =   94
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optRisPEdSt 
               Caption         =   "Baixados"
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
               Left            =   1560
               TabIndex        =   93
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton optRisPEdSt 
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
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   92
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Riscos ]"
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
            Height          =   975
            Index           =   7
            Left            =   -74880
            TabIndex        =   82
            Top             =   1080
            Width           =   9855
            Begin VB.CommandButton Command8 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":0494
               Style           =   1  'Graphical
               TabIndex        =   88
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox txtCODROSCOFIN 
               Height          =   315
               Left            =   1800
               TabIndex        =   87
               Text            =   "txtCODROSCOFIN"
               Top             =   600
               Width           =   1455
            End
            Begin VB.ComboBox cboRISCOFIN 
               Height          =   315
               Left            =   3600
               TabIndex        =   86
               Text            =   "cboRISCOFIN"
               Top             =   600
               Width           =   6135
            End
            Begin VB.ComboBox cboRISCOINI 
               Height          =   315
               Left            =   3600
               TabIndex        =   85
               Text            =   "cboRISCOINI"
               Top             =   240
               Width           =   6135
            End
            Begin VB.TextBox txtCODROSCOINI 
               Height          =   315
               Left            =   1800
               TabIndex        =   84
               Text            =   "txtCODROSCOINI"
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton Command7 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":0596
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "Risco Final"
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
               Index           =   15
               Left            =   120
               TabIndex        =   90
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Risco Inicial"
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
               Index           =   14
               Left            =   120
               TabIndex        =   89
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Data de Entrega ]"
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
            Index           =   6
            Left            =   -74880
            TabIndex        =   77
            Top             =   360
            Width           =   9855
            Begin MSMask.MaskEdBox mskDTPEDFIN 
               Height          =   285
               Left            =   3720
               TabIndex        =   78
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDTPEDINI 
               Height          =   285
               Left            =   1320
               TabIndex        =   79
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Data Final"
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
               Index           =   13
               Left            =   2640
               TabIndex        =   81
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Data Inicial"
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
               Index           =   12
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "[ Somente Pedidos Atrazados ]"
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   120
            TabIndex        =   72
            Top             =   3600
            Width           =   3375
            Begin VB.OptionButton optCotPedSimNao 
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
               Height          =   255
               Index           =   5
               Left            =   2280
               TabIndex        =   76
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optCotPedSimNao 
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
               Index           =   3
               Left            =   360
               TabIndex        =   74
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optCotPedSimNao 
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
               Index           =   2
               Left            =   1320
               TabIndex        =   73
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Fornecedores ]"
            ForeColor       =   &H8000000D&
            Height          =   975
            Index           =   5
            Left            =   120
            TabIndex        =   63
            Top             =   2520
            Width           =   9855
            Begin VB.TextBox txtCODFORNECINIPed 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   69
               Text            =   "txtCODFORNEC"
               Top             =   240
               Width           =   1455
            End
            Begin VB.ComboBox cboFornecINIPed 
               Height          =   315
               Left            =   3600
               TabIndex        =   68
               Text            =   "cboFornec"
               Top             =   255
               Width           =   6135
            End
            Begin VB.CommandButton Command6 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":0698
               Style           =   1  'Graphical
               TabIndex        =   67
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtCODFORNECFINPed 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   66
               Text            =   "txtCODFORNEC"
               Top             =   615
               Width           =   1455
            End
            Begin VB.ComboBox cboFornecFINPed 
               Height          =   315
               Left            =   3600
               TabIndex        =   65
               Text            =   "cboFornec"
               Top             =   615
               Width           =   6135
            End
            Begin VB.CommandButton Command5 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":079A
               Style           =   1  'Graphical
               TabIndex        =   64
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "Fornecedor Final"
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
               Index           =   11
               Left            =   120
               TabIndex        =   71
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Fornecedor Inicial"
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
               Index           =   10
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Produtos ]"
            ForeColor       =   &H8000000D&
            Height          =   975
            Index           =   4
            Left            =   120
            TabIndex        =   54
            Top             =   1560
            Width           =   9855
            Begin VB.CommandButton Command4 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":089C
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtProdutoIniPed 
               Height          =   315
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   59
               Text            =   "txtProduto"
               Top             =   240
               Width           =   1455
            End
            Begin VB.ComboBox cboProdutoIniPed 
               Height          =   315
               Left            =   3600
               TabIndex        =   58
               Text            =   "cboProduto"
               Top             =   240
               Width           =   6135
            End
            Begin VB.ComboBox cboProdutoFimPed 
               Height          =   315
               Left            =   3600
               TabIndex        =   57
               Text            =   "cboProduto"
               Top             =   600
               Width           =   6135
            End
            Begin VB.TextBox txtProdutoFimPed 
               Height          =   315
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   56
               Text            =   "txtProduto"
               Top             =   600
               Width           =   1455
            End
            Begin VB.CommandButton Command3 
               Height          =   315
               Left            =   3240
               Picture         =   "frmRELSUPRI.frx":099E
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "Produto Inicial"
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
               Index           =   9
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Produto Final"
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
               Index           =   8
               Left            =   120
               TabIndex        =   61
               Top             =   600
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Data de Entrega ]"
            ForeColor       =   &H8000000D&
            Height          =   615
            Index           =   3
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   9855
            Begin MSMask.MaskEdBox mskDataFinPed 
               Height          =   285
               Left            =   3720
               TabIndex        =   50
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDataIniPed 
               Height          =   285
               Left            =   1320
               TabIndex        =   51
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Data Inicial"
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
               Index           =   7
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Data Final"
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
               Index           =   6
               Left            =   2640
               TabIndex        =   52
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "[ Status do Pedido ]"
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   5400
            TabIndex        =   45
            Top             =   360
            Width           =   4575
            Begin VB.OptionButton optStatus 
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
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "Aberto"
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
               Index           =   6
               Left            =   1800
               TabIndex        =   47
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "Baixado"
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
               Index           =   5
               Left            =   3240
               TabIndex        =   46
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Ordem ]"
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   5175
            Begin VB.OptionButton optOrdem 
               Caption         =   "Dt. Entrega"
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
               Index           =   5
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton optOrdem 
               Caption         =   "Produto"
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
               Index           =   4
               Left            =   1920
               TabIndex        =   43
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optOrdem 
               Caption         =   "Fornecedor"
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
               Index           =   3
               Left            =   3480
               TabIndex        =   42
               Top             =   240
               Width           =   1335
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10335
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
         Picture         =   "frmRELSUPRI.frx":0AA0
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "frmRELSUPRI.frx":0BA2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELSUPRI"
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
Dim objRELSUPRI     As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Private Sub cboFornecFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornecFIN, KeyAscii
End Sub

Private Sub cboFornecFIN_Validate(Cancel As Boolean)
    If cboFornecFIN.ListIndex > -1 Then txtCODFORNECFIN.Text = cboFornecFIN.ItemData(cboFornecFIN.ListIndex)
End Sub


Private Sub cboFornecFINPed_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornecFINPed, KeyAscii
End Sub

Private Sub cboFornecINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornecINI, KeyAscii
End Sub

Private Sub cboFornecINI_Validate(Cancel As Boolean)
    If cboFornecINI.ListIndex > -1 Then txtCODFORNECINI.Text = cboFornecINI.ItemData(cboFornecINI.ListIndex)
End Sub

Private Sub cboFornecINIPed_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornecINIPed, KeyAscii
End Sub

Private Sub cboProdutoFim_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProdutoFim, KeyAscii
End Sub

Private Sub cboProdutoFim_Validate(Cancel As Boolean)
    If cboProdutoFim.ListIndex > -1 Then txtProdutoFim.Text = Trim(Mid(cboProdutoFim.Text, 1, 10))
End Sub


Private Sub cboProdutoFimPed_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProdutoFimPed, KeyAscii
End Sub

Private Sub cboProdutoIni_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProdutoIni, KeyAscii
End Sub

Private Sub cboProdutoIni_Validate(Cancel As Boolean)
    If cboProdutoIni.ListIndex > -1 Then txtProdutoIni.Text = Trim(Mid(cboProdutoIni.Text, 1, 10))
End Sub

Private Sub cboProdutoIniPed_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProdutoIniPed, KeyAscii
End Sub

Private Sub cboRiscoFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboRiscoFIN, KeyAscii
End Sub

Private Sub cboRiscoFIN_Validate(Cancel As Boolean)
    If cboRiscoFIN.ListIndex > -1 Then txtCODROSCOFIN.Text = cboRiscoFIN.ItemData(cboRiscoFIN.ListIndex)
End Sub

Private Sub cboRiscoINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboRiscoINI, KeyAscii
End Sub

Private Sub cboRiscoINI_Validate(Cancel As Boolean)
    If cboRiscoINI.ListIndex > -1 Then txtCODROSCOINI.Text = cboRiscoINI.ItemData(cboRiscoINI.ListIndex)
End Sub

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    
    If stCotaPedidos.Tab = 0 Then
       If optOrdem(0).Value Then
          If Imprimir = False Then Exit Sub
       ElseIf optOrdem(1).Value Then
          If ImprimirProdCota = False Then Exit Sub
       ElseIf optOrdem(2).Value Then
          If ImprimirFornCota = False Then Exit Sub
       End If
    ElseIf stCotaPedidos.Tab = 1 Then
       If stTabPedidos.Tab = 0 Then
          If optOrdem(5).Value Then
             If ImprimirPedDtNao = False Then Exit Sub
          ElseIf optOrdem(4).Value Then
             If ImprimirPedPdNao = False Then Exit Sub
          ElseIf optOrdem(3).Value Then
             If ImprimirPedForNao = False Then Exit Sub
          End If
       ElseIf stTabPedidos.Tab = 1 Then
          Call ImprimirPedDtRisco
       End If
    End If
    
End Sub

Private Sub cmdPesqFor_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADFORNEC"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "3000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNECINI.Text = varRETORNO
    
    cboFornecINI.ListIndex = -1
    txtCODFORNECINI.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objRELSUPRI = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
    Unload Me
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

Private Sub Command13_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADFORNEC"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "3000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNECFIN.Text = varRETORNO
    
    cboFornecFIN.ListIndex = -1
    txtCODFORNECFIN.SetFocus

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
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    
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
    
    If Len(Trim(varRETORNO)) > 0 Then txtProdutoFimPed.Text = varRETORNO
    
    cboProdutoFimPed.ListIndex = -1
    txtProdutoFimPed.SetFocus

End Sub

Private Sub Command4_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    
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
    
    If Len(Trim(varRETORNO)) > 0 Then txtProdutoIniPed.Text = varRETORNO
    
    cboProdutoIniPed.ListIndex = -1
    txtProdutoIniPed.SetFocus

End Sub

Private Sub Command5_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADFORNEC"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "3000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNECFINPed.Text = varRETORNO
    
    cboFornecFINPed.ListIndex = -1
    txtCODFORNECFINPed.SetFocus

End Sub

Private Sub Command6_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADFORNEC"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "3000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNECINIPed.Text = varRETORNO
    
    cboFornecINIPed.ListIndex = -1
    txtCODFORNECINIPed.SetFocus

End Sub

Private Sub Command7_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADRISCO"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Riscos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODROSCOINI.Text = varRETORNO
    
    cboRiscoINI.ListIndex = -1
    txtCODROSCOINI.SetFocus

End Sub

Private Sub Command8_Click()
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADRISCO"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Riscos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODROSCOFIN.Text = varRETORNO
    
    cboRiscoFIN.ListIndex = -1
    txtCODROSCOFIN.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELSUPRI = CreateObject("RELSUPRI.clsRELSUPRI")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELSUPRI
    
    objRELSUPRI.FILIAL = FILIAL
    
    objRELSUPRI.PreencheComboFornec cboFornecINI
    objRELSUPRI.PreencheComboFornec cboFornecFIN
    
    objRELSUPRI.PreencheComboFornec cboFornecINIPed
    objRELSUPRI.PreencheComboFornec cboFornecFINPed
    
    objRELSUPRI.PreencheComboProd cboProdutoIni
    objRELSUPRI.PreencheComboProd cboProdutoFim
    
    objRELSUPRI.PreencheComboProd cboProdutoIniPed
    objRELSUPRI.PreencheComboProd cboProdutoFimPed
    
    objRELSUPRI.PreencheComboCadRisco cboRiscoINI
    objRELSUPRI.PreencheComboCadRisco cboRiscoFIN
    
    optOrdem(0).Value = True
    optStatus(0).Value = True
    optCotPedSimNao(0).Value = True
    
    optOrdem(5).Value = True
    optStatus(7).Value = True
    optCotPedSimNao(2).Value = True
    
    mskDtInicial.Text = Format(Now, "DD/MM/YYYY")
    mskDtFinal.Text = Format(Now + 30, "DD/MM/YYYY")
    
    mskDataIniPed.Text = Format(Now, "DD/MM/YYYY")
    mskDataFinPed.Text = Format(Now + 30, "DD/MM/YYYY")
    
    mskDTPEDINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTPEDFIN.Text = Format(Now + 30, "DD/MM/YYYY")
    
    stCotaPedidos.Tab = 0
    optRisPEdSt(0).Value = True
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    '' --------------------------------------
End Sub



Private Sub mskDataFinPed_GotFocus()
    objBLBFunc.SelecionaCampos mskDataFinPed.Name, frmRELSUPRI
End Sub

Private Sub mskDataIniPed_GotFocus()
    objBLBFunc.SelecionaCampos mskDataIniPed.Name, frmRELSUPRI
End Sub

Private Sub mskDtFinal_GotFocus()
    objBLBFunc.SelecionaCampos mskDtFinal.Name, frmRELSUPRI
End Sub

Private Sub mskDtInicial_GotFocus()
    objBLBFunc.SelecionaCampos mskDtInicial.Name, frmRELSUPRI
End Sub

Private Sub stCotaPedidos_DblClick()
    If stCotaPedidos.Tab = 0 Then '' Cotações
    ElseIf stCotaPedidos.Tab = 1 Then '' Pedidos
    End If
End Sub

Private Sub txtCODFORNECFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNECFIN.Name, frmRELSUPRI
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
       MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODFORNECFINPed_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNECFINPed.Name, frmRELSUPRI
End Sub

Private Sub txtCODFORNECFINPed_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNECFINPed.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNECFINPed.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECFINPed.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFornecFINPed.ListIndex = -1
    For I = 0 To (cboFornecINIPed.ListCount - 1)
        If cboFornecFINPed.ItemData(I) = Str(Val(txtCODFORNECFINPed.Text)) Then cboFornecFINPed.ListIndex = I
    Next I
    
    If cboFornecFINPed.ListIndex = -1 Then
       MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECFINPed.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODFORNECINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNECINI.Name, frmRELSUPRI
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
       MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub


Private Sub txtCODFORNECINIPed_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNECINIPed.Name, frmRELSUPRI
End Sub

Private Sub txtCODFORNECINIPed_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNECINIPed.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNECINIPed.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECINIPed.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFornecINIPed.ListIndex = -1
    For I = 0 To (cboFornecINIPed.ListCount - 1)
        If cboFornecINIPed.ItemData(I) = Str(Val(txtCODFORNECINIPed.Text)) Then cboFornecINIPed.ListIndex = I
    Next I
    
    If cboFornecINIPed.ListIndex = -1 Then
       MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNECINIPed.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODROSCOFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODROSCOFIN.Name, frmRELSUPRI
End Sub

Private Sub txtCODROSCOFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODROSCOFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODROSCOFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODROSCOFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboRiscoFIN.ListIndex = -1
    For I = 0 To (cboRiscoFIN.ListCount - 1)
        If cboRiscoFIN.ItemData(I) = Str(Val(txtCODROSCOFIN.Text)) Then cboRiscoFIN.ListIndex = I
    Next I
    
    If cboRiscoFIN.ListIndex = -1 Then
       MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODROSCOFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODROSCOINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODROSCOINI.Name, frmRELSUPRI
End Sub

Private Sub txtCODROSCOINI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODROSCOINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODROSCOINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODROSCOINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboRiscoINI.ListIndex = -1
    For I = 0 To (cboRiscoINI.ListCount - 1)
        If cboRiscoINI.ItemData(I) = Str(Val(txtCODROSCOINI.Text)) Then cboRiscoINI.ListIndex = I
    Next I
    
    If cboRiscoINI.ListIndex = -1 Then
       MsgBox "Este fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODROSCOINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtProdutoFim_GotFocus()
    objBLBFunc.SelecionaCampos txtProdutoFim.Name, frmRELSUPRI
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

Private Sub txtProdutoFimPed_GotFocus()
    objBLBFunc.SelecionaCampos txtProdutoFimPed.Name, frmRELSUPRI
End Sub

Private Sub txtProdutoFimPed_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtProdutoFimPed.Text)) = 0 Then Exit Sub
    
   cboProdutoFimPed.ListIndex = -1
   For I = 0 To (cboProdutoFimPed.ListCount - 1)
       If Trim(Mid(cboProdutoFimPed.List(I), 1, 10)) = Trim(txtProdutoFimPed.Text) Then cboProdutoFimPed.ListIndex = I
   Next I
    
   If cboProdutoFimPed.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtProdutoFimPed.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtProdutoIni_GotFocus()
   objBLBFunc.SelecionaCampos txtProdutoIni.Name, frmRELSUPRI
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

Private Function ConfereCampos() As Boolean

    
    ConfereCampos = False
    
    If stCotaPedidos.Tab = 0 Then
        '' Emissão
        If Not IsDate(mskDtInicial.Text) Then
           MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
           mskDtInicial.SetFocus
           Exit Function
        End If
        If Not IsDate(mskDtFinal.Text) Then
           MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
           mskDtFinal.SetFocus
           Exit Function
        End If
        If IsNull(mskDtInicial.Text) And Not IsNull(mskDtFinal.Text) Then
           MsgBox "Data Inicial deve ser Preenchida !!!", vbOKOnly + vbExclamation, "Aviso"
           mskDtInicial.SetFocus
           Exit Function
        End If
        If CDate(mskDtInicial.Text) > CDate(mskDtFinal.Text) Then
           MsgBox "Data Inicial não pode ser maior que data Final !!!", vbOKOnly + vbExclamation, "Aviso"
           mskDtInicial.SetFocus
           Exit Function
        End If
        
        '' Produto Inicial/Final
        If Len(Trim(txtProdutoIni.Text)) = 0 And Len(Trim(txtProdutoFim.Text)) > 0 Then
           MsgBox "Produto Inicial Deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
           txtProdutoIni.SetFocus
           Exit Function
        End If
        
        '' Fornecedor
        
        If Len(Trim(txtCODFORNECINI.Text)) = 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
           MsgBox "Fornecedor Inicial Deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
           txtCODFORNECINI.SetFocus
           Exit Function
        End If
        If Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
           If Not IsNumeric(txtCODFORNECINI.Text) Then
              MsgBox "Somente é permitido numero no campo fornecedor inicial !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODFORNECINI.SetFocus
              Exit Function
           End If
           If Not IsNumeric(txtCODFORNECFIN.Text) Then
              MsgBox "Somente é permitido numero no campo fornecedor final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODFORNECFIN.SetFocus
              Exit Function
           End If
           If CLng(txtCODFORNECINI.Text) > CLng(txtCODFORNECFIN.Text) Then
              MsgBox "Fornecedor não póde ser maior que fornecedor final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODFORNECINI.SetFocus
              Exit Function
           End If
        End If
    ElseIf stCotaPedidos.Tab = 1 Then
        If stTabPedidos.Tab = 0 Then
            '' Emissão
            If Not IsDate(mskDataIniPed.Text) Then
               MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
               mskDataIniPed.SetFocus
               Exit Function
            End If
            If Not IsDate(mskDataFinPed.Text) Then
               MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
               mskDataFinPed.SetFocus
               Exit Function
            End If
            If IsNull(mskDataIniPed.Text) And Not IsNull(mskDataFinPed.Text) Then
               MsgBox "Data Inicial deve ser Preenchida !!!", vbOKOnly + vbExclamation, "Aviso"
               mskDataIniPed.SetFocus
               Exit Function
            End If
            If CDate(mskDataIniPed.Text) > CDate(mskDataFinPed.Text) Then
               MsgBox "Data Inicial não pode ser maior que data Final !!!", vbOKOnly + vbExclamation, "Aviso"
               mskDataIniPed.SetFocus
               Exit Function
            End If
            
            '' Produto Inicial/Final
            If Len(Trim(txtProdutoIniPed.Text)) = 0 And Len(Trim(txtProdutoFimPed.Text)) > 0 Then
               MsgBox "Produto Inicial Deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
               txtProdutoIniPed.SetFocus
               Exit Function
            End If
            
            '' Fornecedor
            
            If Len(Trim(txtCODFORNECINIPed.Text)) = 0 And Len(Trim(txtCODFORNECFINPed.Text)) > 0 Then
               MsgBox "Fornecedor Inicial Deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
               txtCODFORNECINIPed.SetFocus
               Exit Function
            End If
            If Len(Trim(txtCODFORNECINIPed.Text)) > 0 And Len(Trim(txtCODFORNECFINPed.Text)) > 0 Then
               If Not IsNumeric(txtCODFORNECINI.Text) Then
                  MsgBox "Somente é permitido numero no campo fornecedor inicial !!!", vbOKOnly + vbExclamation, "Aviso"
                  txtCODFORNECINIPed.SetFocus
                  Exit Function
               End If
               If Not IsNumeric(txtCODFORNECFINPed.Text) Then
                  MsgBox "Somente é permitido numero no campo fornecedor final !!!", vbOKOnly + vbExclamation, "Aviso"
                  txtCODFORNECFINPed.SetFocus
                  Exit Function
               End If
               If CLng(txtCODFORNECINIPed.Text) > CLng(txtCODFORNECFINPed.Text) Then
                  MsgBox "Fornecedor não póde ser maior que fornecedor final !!!", vbOKOnly + vbExclamation, "Aviso"
                  txtCODFORNECINIPed.SetFocus
                  Exit Function
               End If
            End If
        ElseIf stTabPedidos.Tab = 1 Then
            
        End If
    End If
    ConfereCampos = True

End Function

Private Function Imprimir() As Boolean
    Imprimir = False
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    Dim strARQ    As String
    
    sSql = "Select "
    
    sSql = sSql & "       SGI_COTAHEADER.SGI_DATA   "
    
    sSql = sSql & "      ,SGI_COTAITENS.SGI_PRODUTO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO "
    sSql = sSql & "      ,SGI_CADUNIMED.SGI_UNIDADE "
    
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODIGO  "
    sSql = sSql & "      ,SGI_COTAHEADER.SGI_STATUS "
    
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODFOR "
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_QTD "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_VLUNIT "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_VLTOT "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_PRZENTR "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODCONDPGT "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODPED "
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO "
    
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_UNIDMEDIDA "
    sSql = sSql & "      ,SGI_CADUNIMED.SGI_CODIGO "
    sSql = sSql & "      ,SGI_COTAHEADER.SGI_CODIGO   "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADCONDPGTO SGI_CADCONDPGTO "
    sSql = sSql & "      ,SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & "      ,SGI_CADUNIMED SGI_CADUNIMED "
    sSql = sSql & "      ,SGI_COTAHEADER SGI_COTAHEADER "
    sSql = sSql & "      ,SGI_COTAITENS SGI_COTAITENS "
    sSql = sSql & " Where "
    
    sSql = sSql & "       SGI_COTAITENS.SGI_FILIAL  = SGI_COTAHEADER.SGI_FILIAL "
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODIGO  = SGI_COTAHEADER.SGI_CODIGO "
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL  = SGI_CADPRODUTO.SGI_FILIAL "
    sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO = SGI_CADPRODUTO.SGI_CODIGO "
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL  = SGI_CADFORNEC.SGI_FILIAL "
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR  = SGI_CADFORNEC.SGI_CODIGO "
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL  = SGI_CADCONDPGTO.SGI_FILIAL "
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODCONDPGT = SGI_CADCONDPGTO.SGI_CODIGO "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL = SGI_CADUNIMED.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_UNIDMEDIDA = SGI_CADUNIMED.SGI_CODIGO "
    
    If Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR = " & Trim(txtCODFORNECINI.Text)
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR >= " & Trim(txtCODFORNECINI.Text) & " And SGI_COTAITENS.SGI_CODFOR <= " & Trim(txtCODFORNECFIN.Text)
    End If
    
    If Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) = 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO = '" & Trim(txtProdutoIni.Text) & "'"
    ElseIf Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) > 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO >= '" & Trim(txtProdutoIni.Text) & "' And SGI_COTAITENS.SGI_PRODUTO <= '" & Trim(txtProdutoFim.Text) & "'"
    End If
    
      
    If optCotPedSimNao(0).Value = True Then
       sSql = sSql & "    And SGI_COTAITENS.SGI_CODPED IS NOT NULL"
    ElseIf optCotPedSimNao(1).Value = True Then
       sSql = sSql & "    And SGI_COTAITENS.SGI_CODPED IS NULL"
    End If
    
    If optStatus(1).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'DIGITADO'"
    ElseIf optStatus(2).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'COTADO'"
    ElseIf optStatus(3).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'LIBERADO'"
    End If
    
    sSql = sSql & "   And SGI_COTAHEADER.SGI_FILIAL = " & FILIAL
    If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_DATA = '" & Trim(Format(CDate(mskDtInicial.Text), "MM/DD/YYYY")) & "'"
    Else
       sSql = sSql & "   And SGI_COTAHEADER.SGI_DATA >= '" & Trim(Format(CDate(mskDtInicial.Text), "MM/DD/YYYY")) & "' And SGI_COTAHEADER.SGI_DATA <= '" & Trim(Format(CDate(mskDtFinal.Text), "MM/DD/YYYY")) & "'"
    End If
    
    sSql = sSql & " Order By "

       strCABEC1 = "Relatório de Cotação de Compras no periodo"
       If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
          strCABEC2 = "Na data de " & mskDtInicial.Text
       Else
          strCABEC2 = "De " & mskDtInicial.Text & " a " & mskDtFinal.Text
       End If
       sSql = sSql & "          SGI_COTAHEADER.SGI_DATA "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    
    '' Chamada do Relatório
    If optOrdem(0).Value Then strARQ = "RELSUPRIDATA.rpt"
    If optOrdem(1).Value Then strARQ = "RELSUPRIPROD.rpt"
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelSupri & strARQ, Linha, 1, strCABEC1, strCABEC2, True
   
    Imprimir = True
End Function

Private Function ImprimirProdCota() As Boolean
    ImprimirProdCota = False
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    
    sSql = sSql & "       SGI_COTAITENS.SGI_PRODUTO "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO "
    sSql = sSql & "      ,SGI_CADUNIMED.SGI_UNIDADE "
    
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODIGO "
    sSql = sSql & "      ,SGI_COTAHEADER.SGI_DATA "
    sSql = sSql & "      ,SGI_COTAHEADER.SGI_STATUS "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODFOR "
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_QTD "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_VLUNIT "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_VLTOT "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_PRZENTR "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODCONDPGT "
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO "
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODPED "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADCONDPGTO SGI_CADCONDPGTO "
    sSql = sSql & "      ,SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & "      ,SGI_CADUNIMED SGI_CADUNIMED "
    sSql = sSql & "      ,SGI_COTAHEADER SGI_COTAHEADER "
    sSql = sSql & "      ,SGI_COTAITENS SGI_COTAITENS "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_COTAITENS.SGI_FILIAL      = " & FILIAL
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL      = SGI_CADPRODUTO.SGI_FILIAL "
    sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO     = SGI_CADPRODUTO.SGI_CODIGO "
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL      = SGI_COTAHEADER.SGI_FILIAL "
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODIGO      = SGI_COTAHEADER.SGI_CODIGO "
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL      = SGI_CADFORNEC.SGI_FILIAL "
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR      = SGI_CADFORNEC.SGI_CODIGO "
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL      = SGI_CADCONDPGTO.SGI_FILIAL "
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODCONDPGT  = SGI_CADCONDPGTO.SGI_CODIGO "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADUNIMED.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_UNIDMEDIDA = SGI_CADUNIMED.SGI_CODIGO "
    
    If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_DATA = '" & Trim(Format(CDate(mskDtInicial.Text), "MM/DD/YYYY")) & "'"
    Else
       sSql = sSql & "   And SGI_COTAHEADER.SGI_DATA >= '" & Trim(Format(CDate(mskDtInicial.Text), "MM/DD/YYYY")) & "' And SGI_COTAHEADER.SGI_DATA <= '" & Trim(Format(CDate(mskDtFinal.Text), "MM/DD/YYYY")) & "'"
    End If
    
    If Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR = " & Trim(txtCODFORNECINI.Text)
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR >= " & Trim(txtCODFORNECINI.Text) & " And SGI_COTAITENS.SGI_CODFOR <= " & Trim(txtCODFORNECFIN.Text)
    End If
    
    If optStatus(1).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'DIGITADO'"
    ElseIf optStatus(2).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'COTADO'"
    ElseIf optStatus(3).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'LIBERADO'"
    End If
    
    If Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) = 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO = '" & Trim(txtProdutoIni.Text) & "'"
    ElseIf Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) > 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO >= '" & Trim(txtProdutoIni.Text) & "' And SGI_COTAITENS.SGI_PRODUTO <= '" & Trim(txtProdutoFim.Text) & "'"
    End If
    
    If optCotPedSimNao(0).Value = True Then
       sSql = sSql & "    And SGI_COTAITENS.SGI_CODPED IS NOT NULL"
    ElseIf optCotPedSimNao(1).Value = True Then
       sSql = sSql & "    And SGI_COTAITENS.SGI_CODPED IS NULL"
    End If
    
    sSql = sSql & " Order by SGI_COTAITENS.SGI_PRODUTO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    
    strCABEC1 = "Relatório de Cotação de Compras Por Produto"
    If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
       strCABEC2 = "Na data de " & mskDtInicial.Text
    Else
       strCABEC2 = "De " & mskDtInicial.Text & " a " & mskDtFinal.Text
    End If
    
    '' Chamada do Relatório
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELSUPRIPROD.rpt", Linha, 1, strCABEC1, strCABEC2, True
   
    ImprimirProdCota = True
End Function


Private Function ImprimirFornCota() As Boolean
    ImprimirFornCota = False
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_COTAITENS.SGI_CODFOR " & vbCrLf
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_COTAHEADER.SGI_DATA " & vbCrLf
    sSql = sSql & "      ,SGI_COTAHEADER.SGI_STATUS " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_PRODUTO " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,SGI_CADUNIMED.SGI_UNIDADE " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_QTD " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_VLUNIT " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_VLTOT " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_PRZENTR " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODCONDPGT "
    sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_CODPED "
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADCONDPGTO SGI_CADCONDPGTO " & vbCrLf
    sSql = sSql & "       ,SGI_CADFORNEC SGI_CADFORNEC " & vbCrLf
    sSql = sSql & "       ,SGI_CADPRODUTO SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & "       ,SGI_CADUNIMED SGI_CADUNIMED " & vbCrLf
    sSql = sSql & "       ,SGI_COTAHEADER SGI_COTAHEADER " & vbCrLf
    sSql = sSql & "       ,SGI_COTAITENS SGI_COTAITENS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_COTAITENS.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL = SGI_CADFORNEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR = SGI_CADFORNEC.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL = SGI_COTAHEADER.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODIGO = SGI_COTAHEADER.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO = SGI_CADPRODUTO.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_FILIAL = SGI_CADCONDPGTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODCONDPGT = SGI_CADCONDPGTO.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL = SGI_CADUNIMED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_UNIDMEDIDA = SGI_CADUNIMED.SGI_CODIGO " & vbCrLf
    
    If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_DATA = '" & Trim(Format(CDate(mskDtInicial.Text), "MM/DD/YYYY")) & "'"
    Else
       sSql = sSql & "   And SGI_COTAHEADER.SGI_DATA >= '" & Trim(Format(CDate(mskDtInicial.Text), "MM/DD/YYYY")) & "' And SGI_COTAHEADER.SGI_DATA <= '" & Trim(Format(CDate(mskDtFinal.Text), "MM/DD/YYYY")) & "'"
    End If
    
    If Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR = " & Trim(txtCODFORNECINI.Text)
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR >= " & Trim(txtCODFORNECINI.Text) & " And SGI_COTAITENS.SGI_CODFOR <= " & Trim(txtCODFORNECFIN.Text)
    End If
    
    If optStatus(1).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'DIGITADO'"
    ElseIf optStatus(2).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'COTADO'"
    ElseIf optStatus(3).Value = True Then
       sSql = sSql & "   And SGI_COTAHEADER.SGI_STATUS = 'LIBERADO'"
    End If
    
    If Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) = 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO = '" & Trim(txtProdutoIni.Text) & "'"
    ElseIf Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) > 0 Then
       sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO >= '" & Trim(txtProdutoIni.Text) & "' And SGI_COTAITENS.SGI_PRODUTO <= '" & Trim(txtProdutoFim.Text) & "'"
    End If
    
    If optCotPedSimNao(0).Value = True Then
       sSql = sSql & "    And SGI_COTAITENS.SGI_CODPED IS NOT NULL"
    ElseIf optCotPedSimNao(1).Value = True Then
       sSql = sSql & "    And SGI_COTAITENS.SGI_CODPED IS NULL"
    End If
    
    sSql = sSql & " Order by SGI_COTAITENS.SGI_CODFOR "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    
    strCABEC1 = "Relatório de Cotação de Compras Por Fornecedor"
    If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
       strCABEC2 = "Na data de " & mskDtInicial.Text
    Else
       strCABEC2 = "De " & mskDtInicial.Text & " a " & mskDtFinal.Text
    End If
    
    '' Chamada do Relatório
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELSUPRIFOR.rpt", Linha, 1, strCABEC1, strCABEC2, True
   
    ImprimirFornCota = True
End Function

Private Sub txtProdutoIniPed_GotFocus()
   objBLBFunc.SelecionaCampos txtProdutoIniPed.Name, frmRELSUPRI
End Sub

Private Sub txtProdutoIniPed_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtProdutoIniPed.Text)) = 0 Then Exit Sub
    
   cboProdutoIniPed.ListIndex = -1
   For I = 0 To (cboProdutoIniPed.ListCount - 1)
       If Trim(Mid(cboProdutoIniPed.List(I), 1, 10)) = Trim(txtProdutoIniPed.Text) Then cboProdutoIniPed.ListIndex = I
   Next I
    
   If cboProdutoIniPed.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtProdutoIniPed.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Function ImprimirPedDtNao() As Boolean
    ImprimirPedDtNao = False
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_DTAENTREGA "
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_PRODUTO "
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_CODIGO "
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_CODFOR "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO "
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC "
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_CODIGO "
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_STATUS "
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_QTD "
    sSql = sSql & "      ,SGI_CADUNIMED.SGI_UNIDADE "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & "      ,SGI_CADUNIMED SGI_CADUNIMED "
    sSql = sSql & "      ,SGI_PEDIDOHEADER SGI_PEDIDOHEADER "
    sSql = sSql & "      ,SGI_PEDIDOITENS SGI_PEDIDOITENS "
    
    sSql = sSql & " Where "
    
    
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_FILIAL  = SGI_PEDIDOHEADER.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_CODIGO  = SGI_PEDIDOHEADER.SGI_CODIGO "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_FILIAL  = SGI_CADPRODUTO.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO = SGI_CADPRODUTO.SGI_CODIGO "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADUNIMED.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_UNIDMEDIDA = SGI_CADUNIMED.SGI_CODIGO "
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_FILIAL = SGI_CADFORNEC.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR = SGI_CADFORNEC.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODFORNECINIPed.Text)) > 0 And Len(Trim(txtCODFORNECFINPed.Text)) = 0 Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR = " & Trim(txtCODFORNECINIPed.Text)
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_PEDIDOHEADER.SGI_CODFOR >= " & Trim(txtCODFORNECINIPed.Text) & " And SGI_PEDIDOHEADER.SGI_CODFOR <= " & Trim(txtCODFORNECFINPed.Text) & ")"
    End If
    
    If Len(Trim(txtProdutoIniPed.Text)) > 0 And Len(Trim(txtProdutoFimPed.Text)) = 0 Then
       sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO = '" & Trim(txtProdutoIniPed.Text) & "'"
    ElseIf Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) > 0 Then
       sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO >= '" & Trim(txtProdutoIniPed.Text) & "' And SGI_PEDIDOITENS.SGI_PRODUTO <= '" & Trim(txtProdutoFimPed.Text) & "'"
    End If
    
    If optStatus(6).Value = True Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_STATUS = 'A'"
    ElseIf optStatus(5).Value = True Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_STATUS = 'B'"
    End If
    
    If CDate(mskDataIniPed.Text) = CDate(mskDataFinPed.Text) Then
       sSql = sSql & "   And (SGI_PEDIDOITENS.SGI_DTAENTREGA = '" & Trim(Format(CDate(mskDataIniPed.Text), "MM/DD/YYYY")) & "')"
    Else
       sSql = sSql & "   And (SGI_PEDIDOITENS.SGI_DTAENTREGA >= '" & Trim(Format(CDate(mskDataIniPed.Text), "MM/DD/YYYY")) & "' And SGI_PEDIDOITENS.SGI_DTAENTREGA <= '" & Trim(Format(CDate(mskDataFinPed.Text), "MM/DD/YYYY")) & "')"
    End If
    
    If optCotPedSimNao(3).Value = True Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_PEDIDOITENS.SGI_DTAENTREGA, GETDATE()) * -1) < 0 "
    ElseIf optCotPedSimNao(2).Value = True Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_PEDIDOITENS.SGI_DTAENTREGA, GETDATE()) * -1) >= 0 "
    End If
    
    sSql = sSql & " Order By "

    strCABEC1 = "Relatório de Pedidos de Compras no periodo"
    If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
       strCABEC2 = "Na data de " & mskDataIniPed.Text
    Else
       strCABEC2 = "De " & mskDataIniPed.Text & " a " & mskDataFinPed.Text
    End If
    sSql = sSql & "          SGI_PEDIDOITENS.SGI_DTAENTREGA "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    
    '' Chamada do Relatório
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELPEDCODATANAO.rpt", Linha, 1, strCABEC1, strCABEC2, True
   
    ImprimirPedDtNao = True
End Function

Private Function ImprimirPedPdNao() As Boolean
    ImprimirPedPdNao = False
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_DTAENTREGA "
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_PRODUTO "
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_CODIGO "
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_CODFOR "
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO "
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC "
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_CODIGO "
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_STATUS "
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_QTD "
    sSql = sSql & "      ,SGI_CADUNIMED.SGI_UNIDADE "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & "      ,SGI_CADUNIMED SGI_CADUNIMED "
    sSql = sSql & "      ,SGI_PEDIDOHEADER SGI_PEDIDOHEADER "
    sSql = sSql & "      ,SGI_PEDIDOITENS SGI_PEDIDOITENS "
    
    sSql = sSql & " Where "
    
    
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_FILIAL  = SGI_PEDIDOHEADER.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_CODIGO  = SGI_PEDIDOHEADER.SGI_CODIGO "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_FILIAL  = SGI_CADPRODUTO.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO = SGI_CADPRODUTO.SGI_CODIGO "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADUNIMED.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_UNIDMEDIDA = SGI_CADUNIMED.SGI_CODIGO "
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_FILIAL = SGI_CADFORNEC.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR = SGI_CADFORNEC.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODFORNECINIPed.Text)) > 0 And Len(Trim(txtCODFORNECFINPed.Text)) = 0 Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR = " & Trim(txtCODFORNECINIPed.Text)
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_PEDIDOHEADER.SGI_CODFOR >= " & Trim(txtCODFORNECINIPed.Text) & " And SGI_PEDIDOHEADER.SGI_CODFOR <= " & Trim(txtCODFORNECFINPed.Text) & ")"
    End If
    
    If Len(Trim(txtProdutoIniPed.Text)) > 0 And Len(Trim(txtProdutoFimPed.Text)) = 0 Then
       sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO = '" & Trim(txtProdutoIniPed.Text) & "'"
    ElseIf Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) > 0 Then
       sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO >= '" & Trim(txtProdutoIniPed.Text) & "' And SGI_PEDIDOITENS.SGI_PRODUTO <= '" & Trim(txtProdutoFimPed.Text) & "'"
    End If
    
    If optStatus(6).Value = True Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_STATUS = 'A'"
    ElseIf optStatus(5).Value = True Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_STATUS = 'B'"
    End If
    
    If CDate(mskDataIniPed.Text) = CDate(mskDataFinPed.Text) Then
       sSql = sSql & "   And (SGI_PEDIDOITENS.SGI_DTAENTREGA = '" & Trim(Format(CDate(mskDataIniPed.Text), "MM/DD/YYYY")) & "')"
    Else
       sSql = sSql & "   And (SGI_PEDIDOITENS.SGI_DTAENTREGA >= '" & Trim(Format(CDate(mskDataIniPed.Text), "MM/DD/YYYY")) & "' And SGI_PEDIDOITENS.SGI_DTAENTREGA <= '" & Trim(Format(CDate(mskDataFinPed.Text), "MM/DD/YYYY")) & "')"
    End If
    
    If optCotPedSimNao(3).Value = True Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_PEDIDOITENS.SGI_DTAENTREGA, GETDATE()) * -1) < 0 "
    ElseIf optCotPedSimNao(2).Value = True Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_PEDIDOITENS.SGI_DTAENTREGA, GETDATE()) * -1) >= 0 "
    End If
    
    sSql = sSql & " Order By "

    strCABEC1 = "Relatório de Pedidos de Compras no periodo"
    If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
       strCABEC2 = "Na data de " & mskDataIniPed.Text
    Else
       strCABEC2 = "De " & mskDataIniPed.Text & " a " & mskDataFinPed.Text
    End If
    sSql = sSql & "          SGI_PEDIDOITENS.SGI_DTAENTREGA "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    
    '' Chamada do Relatório
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELPEDCOPRONAO.rpt", Linha, 1, strCABEC1, strCABEC2, True
   
    ImprimirPedPdNao = True
End Function


Private Function ImprimirPedForNao() As Boolean
    ImprimirPedForNao = False
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_DTAENTREGA "
    sSql = sSql & "     , SGI_PEDIDOHEADER.SGI_FILIAL "
    sSql = sSql & "     , SGI_PEDIDOHEADER.SGI_CODIGO "
    sSql = sSql & "     , SGI_PEDIDOHEADER.SGI_CODFOR "
    sSql = sSql & "     , SGI_PEDIDOITENS.SGI_PRODUTO "
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_DESCRICAO "
    
    sSql = sSql & "  From "
    
    sSql = sSql & "       SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO "
    sSql = sSql & "      ,SGI_CADUNIMED SGI_CADUNIMED "
    sSql = sSql & "      ,SGI_PEDIDOHEADER SGI_PEDIDOHEADER "
    sSql = sSql & "      ,SGI_PEDIDOITENS SGI_PEDIDOITENS "
    
    sSql = sSql & " Where "
    
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_FILIAL  = SGI_PEDIDOHEADER.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_CODIGO  = SGI_PEDIDOHEADER.SGI_CODIGO "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_FILIAL  = SGI_CADPRODUTO.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO = SGI_CADPRODUTO.SGI_CODIGO "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADUNIMED.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_UNIDMEDIDA = SGI_CADUNIMED.SGI_CODIGO "
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_FILIAL = SGI_CADFORNEC.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR = SGI_CADFORNEC.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODFORNECINIPed.Text)) > 0 And Len(Trim(txtCODFORNECFINPed.Text)) = 0 Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR = " & Trim(txtCODFORNECINIPed.Text)
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_PEDIDOHEADER.SGI_CODFOR >= " & Trim(txtCODFORNECINIPed.Text) & " And SGI_PEDIDOHEADER.SGI_CODFOR <= " & Trim(txtCODFORNECFINPed.Text) & ")"
    End If
    
    If Len(Trim(txtProdutoIniPed.Text)) > 0 And Len(Trim(txtProdutoFimPed.Text)) = 0 Then
       sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO = '" & Trim(txtProdutoIniPed.Text) & "'"
    ElseIf Len(Trim(txtProdutoIni.Text)) > 0 And Len(Trim(txtProdutoFim.Text)) > 0 Then
       sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO >= '" & Trim(txtProdutoIniPed.Text) & "' And SGI_PEDIDOITENS.SGI_PRODUTO <= '" & Trim(txtProdutoFimPed.Text) & "'"
    End If
    
    If optStatus(6).Value = True Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_STATUS = 'A'"
    ElseIf optStatus(5).Value = True Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_STATUS = 'B'"
    End If
    
    If CDate(mskDataIniPed.Text) = CDate(mskDataFinPed.Text) Then
       sSql = sSql & "   And (SGI_PEDIDOITENS.SGI_DTAENTREGA = '" & Trim(Format(CDate(mskDataIniPed.Text), "MM/DD/YYYY")) & "')"
    Else
       sSql = sSql & "   And (SGI_PEDIDOITENS.SGI_DTAENTREGA Between '" & Trim(Format(CDate(mskDataIniPed.Text), "MM/DD/YYYY")) & "' And '" & Trim(Format(CDate(mskDataFinPed.Text), "MM/DD/YYYY")) & "')"
    End If
    
    If optCotPedSimNao(3).Value = True Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_PEDIDOITENS.SGI_DTAENTREGA, GETDATE()) * -1) < 0 "
    ElseIf optCotPedSimNao(2).Value = True Then
       sSql = sSql & "  And (DATEDIFF(day, SGI_PEDIDOITENS.SGI_DTAENTREGA, GETDATE()) * -1) >= 0 "
    End If
    
    sSql = sSql & " Order By "

    strCABEC1 = "Relatório de Pedidos de Compras no periodo"
    If CDate(mskDtInicial.Text) = CDate(mskDtFinal.Text) Then
       strCABEC2 = "Na data de " & mskDataIniPed.Text
    Else
       strCABEC2 = "De " & mskDataIniPed.Text & " a " & mskDataFinPed.Text
    End If
    sSql = sSql & "          SGI_PEDIDOHEADER.SGI_CODFOR "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    
    '' Chamada do Relatório
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELPEDCOFORNAO.rpt", Linha, 1, strCABEC1, strCABEC2, True
   
    ImprimirPedForNao = True
End Function

Private Sub ImprimirPedDtRisco()
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_DTAENTREGA "
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_CODIGO "
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_CODFOR "
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADFORNEC SGI_CADFORNEC "
    sSql = sSql & "      ,SGI_CADRISCO SGI_CADRISCO "
    sSql = sSql & "      ,SGI_CADRISCOFORNEC SGI_CADRISCOFORNEC "
    sSql = sSql & "      ,SGI_PEDIDOHEADER SGI_PEDIDOHEADER "
    sSql = sSql & "      ,SGI_PEDIDOITENS SGI_PEDIDOITENS "
    
    sSql = sSql & " Where "
    
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_FILIAL = SGI_PEDIDOHEADER.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_CODIGO = SGI_PEDIDOHEADER.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_FILIAL = SGI_CADFORNEC.SGI_FILIAL "
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR = SGI_CADFORNEC.SGI_CODIGO "
    
    sSql = sSql & "   And SGI_CADFORNEC.SGI_FILIAL = SGI_CADRISCOFORNEC.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADFORNEC.SGI_CODIGO = SGI_CADRISCOFORNEC.SGI_CODFORNEC "
    
    sSql = sSql & "   And SGI_CADRISCOFORNEC.SGI_FILIAL = SGI_CADRISCO.SGI_FILIAL "
    sSql = sSql & "   And SGI_CADRISCOFORNEC.SGI_CODIGO = SGI_CADRISCO.SGI_CODIGO "
    
    If CDate(mskDTPEDINI.Text) = CDate(mskDTPEDFIN.Text) Then
       sSql = sSql & "   And (SGI_PEDIDOITENS.SGI_DTAENTREGA = '" & Trim(Format(CDate(mskDTPEDINI.Text), "MM/DD/YYYY")) & "')"
    Else
       sSql = sSql & "   And (SGI_PEDIDOITENS.SGI_DTAENTREGA >= '" & Trim(Format(CDate(mskDTPEDINI.Text), "MM/DD/YYYY")) & "' And SGI_PEDIDOITENS.SGI_DTAENTREGA <= '" & Trim(Format(CDate(mskDTPEDFIN.Text), "MM/DD/YYYY")) & "')"
    End If
    
    If Len(Trim(txtCODROSCOINI.Text)) > 0 And Len(Trim(txtCODROSCOFIN.Text)) = 0 Then
       sSql = sSql & "   And SGI_CADRISCO.SGI_CODIGO = " & Trim(txtCODROSCOINI.Text)
    ElseIf Len(Trim(txtCODFORNECINI.Text)) > 0 And Len(Trim(txtCODFORNECFIN.Text)) > 0 Then
       sSql = sSql & "   And (SGI_CADRISCO.SGI_CODIGO >= " & Trim(txtCODROSCOINI.Text) & " And SGI_CADRISCO.SGI_CODIGO <= " & Trim(txtCODROSCOFIN.Text) & ")"
    End If
    
    strCABEC1 = "Relatório de Pedidos de Compras com Riscos no periodo"
    If CDate(mskDTPEDINI.Text) = CDate(mskDTPEDFIN.Text) Then
       strCABEC2 = "Na data de " & mskDTPEDINI.Text
    Else
       strCABEC2 = "De " & mskDTPEDINI.Text & " a " & mskDTPEDFIN.Text
    End If
    
    If optRisPEdSt(0).Value = True Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_STATUS = 'A'"
       strCABEC2 = strCABEC2 & " (Abertos)"
    ElseIf optRisPEdSt(1).Value = True Then
       sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_STATUS = 'B'"
       strCABEC2 = strCABEC2 & " (Baixados)"
    Else
       strCABEC2 = strCABEC2 & " (Todos)"
    End If
    
    sSql = sSql & " Order By "
    sSql = sSql & "          SGI_PEDIDOITENS.SGI_DTAENTREGA "

    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    '' Chamada do Relatório
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelSupri & "RELRISCOPEDIDOS.rpt", Linha, 1, strCABEC1, strCABEC2, True
   
End Sub

