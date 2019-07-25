VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADCONTASAPG 
   Caption         =   "Cadastro de Contas a Pagar"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab StContAPG 
      Height          =   5415
      Left            =   0
      TabIndex        =   24
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9551
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
      TabCaption(0)   =   "Dados do Titulo"
      TabPicture(0)   =   "frmCADCONTASAPG.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Dados da Baixa"
      TabPicture(1)   =   "frmCADCONTASAPG.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame7 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   72
         Top             =   3480
         Width           =   8055
         Begin VB.TextBox txtCODTIPPGTO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2480
            MaxLength       =   10
            TabIndex        =   22
            Text            =   "txtCODFORNEC"
            Top             =   615
            Width           =   735
         End
         Begin VB.ComboBox cboTIPOPGTO 
            Height          =   315
            Left            =   3600
            TabIndex        =   23
            Text            =   "cboTIPOPGTO"
            Top             =   615
            Width           =   4335
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   3220
            Picture         =   "frmCADCONTASAPG.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   620
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskDTPGTO 
            Height          =   285
            Left            =   6690
            TabIndex        =   21
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtVlPagto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2490
            MaxLength       =   15
            TabIndex        =   20
            Text            =   "txtVlPagto"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Pagto:"
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
            Index           =   30
            Left            =   1020
            TabIndex        =   83
            Top             =   675
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data do Pgto:"
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
            Index           =   29
            Left            =   5400
            TabIndex        =   82
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   16
            Left            =   2490
            TabIndex        =   81
            Top             =   1440
            Width           =   1440
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   15
            Left            =   6690
            TabIndex        =   80
            Top             =   1440
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Acrescimo:"
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
            Index           =   28
            Left            =   5415
            TabIndex        =   79
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Acrescimo:"
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
            Index           =   27
            Left            =   1365
            TabIndex        =   78
            Top             =   1440
            Width           =   945
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   14
            Left            =   2490
            TabIndex        =   77
            Top             =   1080
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor pago:"
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
            Index           =   26
            Left            =   1320
            TabIndex        =   76
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   13
            Left            =   6690
            TabIndex        =   75
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Desconto:"
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
            Index           =   25
            Left            =   5490
            TabIndex        =   74
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desconto:"
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
            Index           =   24
            Left            =   1425
            TabIndex        =   73
            Top             =   1080
            Width           =   885
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   61
         Top             =   2160
         Width           =   8055
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   12
            Left            =   2520
            TabIndex        =   71
            Top             =   960
            Width           =   5400
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   11
            Left            =   6720
            TabIndex        =   70
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   10
            Left            =   2520
            TabIndex        =   69
            Top             =   600
            Width           =   1560
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   8
            Left            =   6720
            TabIndex        =   68
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   7
            Left            =   2520
            TabIndex        =   67
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Numero do documento:"
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
            Index           =   23
            Left            =   360
            TabIndex        =   66
            Top             =   250
            Width           =   1980
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data vencimento:"
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
            Index           =   22
            Left            =   5040
            TabIndex        =   65
            Top             =   250
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Index           =   21
            Left            =   1845
            TabIndex        =   64
            Top             =   610
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de documento:"
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
            Index           =   20
            Left            =   645
            TabIndex        =   63
            Top             =   960
            Width           =   1710
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Parcela:"
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
            Index           =   19
            Left            =   5835
            TabIndex        =   62
            Top             =   610
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   38
         Top             =   360
         Width           =   8055
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   6
            Left            =   6720
            TabIndex        =   60
            Top             =   1450
            Width           =   1200
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   5
            Left            =   2520
            TabIndex        =   59
            Top             =   1450
            Width           =   1560
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   58
            Top             =   1150
            Width           =   5400
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   57
            Top             =   840
            Width           =   5400
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   56
            Top             =   540
            Width           =   5400
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
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
            Index           =   18
            Left            =   1335
            TabIndex        =   55
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Condição de pagamento:"
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
            Index           =   17
            Left            =   240
            TabIndex        =   54
            Top             =   840
            Width           =   2130
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Grupo de despesa:"
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
            Index           =   16
            Left            =   720
            TabIndex        =   53
            Top             =   1150
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total Doc:"
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
            Index           =   15
            Left            =   930
            TabIndex        =   52
            Top             =   1450
            Width           =   1410
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Documento Pai:"
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
            Index           =   14
            Left            =   5205
            TabIndex        =   51
            Top             =   1450
            Width           =   1365
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   1
            Left            =   6720
            TabIndex        =   50
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   49
            Top             =   240
            Width           =   1080
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
            Index           =   13
            Left            =   1680
            TabIndex        =   48
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data do Lançamento:"
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
            Index           =   12
            Left            =   4770
            TabIndex        =   47
            Top             =   240
            Width           =   1845
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1815
         Left            =   120
         TabIndex        =   37
         Top             =   3480
         Width           =   8055
         Begin MSFlexGridLib.MSFlexGrid flxCONTAPGT 
            Height          =   1575
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   120
         TabIndex        =   32
         Top             =   2280
         Width           =   8055
         Begin VB.CommandButton cmdTIPDOC 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTASAPG.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmbGravPGT 
            Height          =   315
            Left            =   7560
            Picture         =   "frmCADCONTASAPG.frx":023C
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   840
            Width           =   375
         End
         Begin VB.ComboBox cboTIPCOD 
            Height          =   315
            Left            =   3480
            TabIndex        =   14
            Text            =   "cboTIPCOD"
            Top             =   840
            Width           =   4095
         End
         Begin VB.TextBox txtCODTIPDOC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   13
            Text            =   "txtCODTIPDOC"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   12
            Text            =   "txtValor"
            Top             =   550
            Width           =   1695
         End
         Begin VB.TextBox txtNUMDOC 
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "txtNUMDOC"
            Top             =   240
            Width           =   1695
         End
         Begin MSMask.MaskEdBox mskDTVencto 
            Height          =   285
            Left            =   6720
            TabIndex        =   11
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPARCELAS 
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
            Height          =   255
            Index           =   0
            Left            =   6720
            TabIndex        =   44
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Parcela:"
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
            Index           =   9
            Left            =   5835
            TabIndex        =   43
            Top             =   570
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de documento:"
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
            Index           =   8
            Left            =   525
            TabIndex        =   36
            Top             =   850
            Width           =   1710
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Left            =   1725
            TabIndex        =   35
            Top             =   570
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data vencimento:"
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
            Left            =   5040
            TabIndex        =   34
            Top             =   285
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Numero do documento:"
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
            Left            =   240
            TabIndex        =   33
            Top             =   280
            Width           =   1980
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   8055
         Begin VB.TextBox txtDOCPAI 
            Height          =   285
            Left            =   6720
            MaxLength       =   10
            TabIndex        =   9
            Text            =   "txtDOCPAI"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtVlTotDoc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   8
            Text            =   "txtVlTotDoc"
            Top             =   1520
            Width           =   1455
         End
         Begin VB.CommandButton cmdPesqGrupDesp 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTASAPG.frx":033E
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton cmdPesqCondPgt 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTASAPG.frx":0440
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdPesqFor 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTASAPG.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   510
            Width           =   375
         End
         Begin VB.ComboBox cboGrupDesp 
            Height          =   315
            Left            =   3480
            TabIndex        =   7
            Text            =   "cboGrupDesp"
            Top             =   1200
            Width           =   4455
         End
         Begin VB.TextBox txtCODGRUPDESP 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "txtCODGRUPDESP"
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cboCondPgto 
            Height          =   315
            Left            =   3480
            TabIndex        =   5
            Text            =   "cboCondPgto"
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox txtCODCONDPGT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "txtCODCONDPGT"
            Top             =   840
            Width           =   735
         End
         Begin VB.ComboBox cboFornec 
            Height          =   315
            Left            =   3480
            TabIndex        =   3
            Text            =   "cboFornec"
            Top             =   530
            Width           =   4455
         End
         Begin VB.TextBox txtCODFORNEC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txtCODFORNEC"
            Top             =   530
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskDTLANCTO 
            Height          =   285
            Left            =   6720
            TabIndex        =   1
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Documento Pai:"
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
            Index           =   11
            Left            =   5205
            TabIndex        =   46
            Top             =   1575
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total Doc:"
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
            Index           =   10
            Left            =   810
            TabIndex        =   45
            Top             =   1560
            Width           =   1410
         End
         Begin VB.Label lblCODPGTO 
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
            Height          =   255
            Index           =   9
            Left            =   2400
            TabIndex        =   0
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data do Lançamento:"
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
            Left            =   4680
            TabIndex        =   31
            Top             =   240
            Width           =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Grupo de despesa:"
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
            Left            =   600
            TabIndex        =   30
            Top             =   1200
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Condição de pagamento:"
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
            TabIndex        =   29
            Top             =   840
            Width           =   2130
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
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
            Left            =   1215
            TabIndex        =   28
            Top             =   550
            Width           =   1035
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
            Left            =   1590
            TabIndex        =   27
            Top             =   240
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8295
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
         Picture         =   "frmCADCONTASAPG.frx":0644
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
         Picture         =   "frmCADCONTASAPG.frx":0746
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Picture         =   "frmCADCONTASAPG.frx":0848
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADCONTASAPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public cTipOper      As String
Public iCodigo       As Integer
Public iParcela      As Integer
Public FILIAL        As Integer
Public strAcesso     As String
Public strMODPAI     As String
Public strUSUARIO    As String
Dim objBLBFunc       As Object
Dim objCADCONTASAPG  As Object
Dim objPESQPADRAO    As Object
Dim arrGRIDPGTOS     As Variant
Dim intQtdParc       As Integer
Dim arrDiasParc      As Variant
Dim strSTATGRID      As String
Dim intLinhaIndice   As Integer


Private Sub cboCondPgto_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCondPgto, KeyAscii
End Sub

Private Sub cboCondPgto_Validate(Cancel As Boolean)
    
    If cboCondPgto.ListIndex > -1 Then
       
       If Len(Trim(txtCODCONDPGT.Text)) > 0 Then
          If txtCODCONDPGT.Text <> cboCondPgto.ItemData(cboCondPgto.ListIndex) Then
             ConfGridCondPGTO
             LimpaCamposPGTO
          End If
       End If
       
       txtCODCONDPGT.Text = cboCondPgto.ItemData(cboCondPgto.ListIndex)
       Call PegaParcelas
       strSTATGRID = "I"
       
    End If
    
End Sub

Private Sub cboFornec_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornec, KeyAscii
End Sub

Private Sub cboFornec_Validate(Cancel As Boolean)
       Dim I As Integer
       If cboFornec.ListIndex > -1 Then
          
          txtCODFORNEC.Text = cboFornec.ItemData(cboFornec.ListIndex)
          txtCODGRUPDESP.Text = PegaGrpDesp(cboFornec.ItemData(cboFornec.ListIndex))
          If Len(Trim(txtCODGRUPDESP.Text)) > 0 Then
             If CInt(txtCODGRUPDESP.Text) = 0 Then txtCODGRUPDESP.Text = ""
          End If
          
          cboGrupDesp.Enabled = True
          txtCODGRUPDESP.Enabled = True
          cmdPesqGrupDesp.Enabled = True
          
          cboGrupDesp.ListIndex = -1
          For I = 0 To (cboGrupDesp.ListCount - 1)
              If cboGrupDesp.ItemData(I) = Str(Val(txtCODGRUPDESP.Text)) Then cboGrupDesp.ListIndex = I
          Next I
          
          If cboGrupDesp.ListIndex > -1 Then
             cboGrupDesp.Enabled = False
             txtCODGRUPDESP.Enabled = False
             cmdPesqGrupDesp.Enabled = False
          End If
          
       End If
End Sub

Private Sub cboGrupDesp_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboGrupDesp, KeyAscii
End Sub

Private Sub cboGrupDesp_Validate(Cancel As Boolean)
    If cboGrupDesp.ListIndex > -1 Then txtCODGRUPDESP.Text = cboGrupDesp.ItemData(cboGrupDesp.ListIndex)
End Sub

Private Sub cboTIPCOD_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTIPCOD, KeyAscii
End Sub

Private Sub cboTIPCOD_Validate(Cancel As Boolean)
    If cboTIPCOD.ListIndex > -1 Then txtCODTIPDOC.Text = cboTIPCOD.ItemData(cboTIPCOD.ListIndex)
End Sub

Private Sub cboTIPOPGTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTIPOPGTO, KeyAscii
End Sub

Private Sub cboTIPOPGTO_Validate(Cancel As Boolean)
    If cboTIPOPGTO.ListIndex > -1 Then txtCODTIPPGTO.Text = cboTIPOPGTO.ItemData(cboTIPOPGTO.ListIndex)
End Sub

Private Sub cmbGravPGT_Click()
    If cTipOper = "I" Or cTipOper = "A" Then AddGridDoc strSTATGRID
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    If VerifNF = True Then Exit Sub
    
    If cTipOper <> "CB" Then
    
        '' Verifica se há baixados
        sSql = "Select" & vbCrLf
        sSql = sSql & "      * " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIAPG " & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_CODIGO = " & objCADCONTASAPG.CODPGTO & vbCrLf
        sSql = sSql & "  And SGI_VLPAGO IS NOT NULL "
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF Then
           MsgBox "Há titulos já baixado !!!", vbOKOnly + vbExclamation, "Aviso"
           BREC.Close
           Exit Sub
        End If
        BREC.Close
        '' ------------------------------

        CmdSalva.Enabled = True
        cmdAltera.Enabled = False
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
   
        Me.Caption = "Cadastro de contas a pagar - [ ALTERAÇÃO ]"
    
        StContAPG.TabEnabled(0) = True
        StContAPG.TabEnabled(1) = False
    
        cTipOper = "A"
        strSTATGRID = "I"
        intLinhaIndice = 0
        
    End If
    
    If cTipOper = "CB" Then
    
        CmdSalva.Enabled = True
        cmdAltera.Enabled = False
        Frame7.Enabled = True
   
        Me.Caption = "Cadastro de contas a pagar - [ ALTERA BAIXA ]"
    
        cTipOper = "AB"
        txtVlPagto.SetFocus
    
    End If

End Sub

Private Sub cmdPesqCondPgt_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCONDPGTO " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Condição de Pagamento", "CADCONDPAGTO.clsCADCONDPAGTO")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCONDPGT.Text = varRETORNO
    
    cboCondPgto.ListIndex = -1
    txtCODCONDPGT.SetFocus

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
    sSql = sSql & "   And SGI_STATUS in(1,2)"
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores", "CADFORNEC.clsCADFORNEC")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNEC.Text = varRETORNO
    
    cboFornec.ListIndex = -1
    txtCODFORNEC.SetFocus

End Sub

Private Sub cmdPesqGrupDesp_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPDESP " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de Despesa", "CADGRUPDESP.clsCADGRUPDESP")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODGRUPDESP.Text = varRETORNO
    
    cboCondPgto.ListIndex = -1
    txtCODGRUPDESP.SetFocus

End Sub

Private Sub CmdSalva_Click()
    
    Dim I       As Integer
    Dim intResp As Integer
    
    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Or cTipOper = "A" Then
       
       If cTipOper = "I" Then objCADCONTASAPG.CODPGTO = objCADCONTASAPG.Gera_Codigo(Me.Name)
       objCADCONTASAPG.DATALCTO = CDate(mskDTLANCTO.Text)
       objCADCONTASAPG.CODFORN = txtCODFORNEC.Text
       objCADCONTASAPG.CODCONDPGTO = txtCODCONDPGT.Text
       If Len(Trim(txtCODGRUPDESP.Text)) > 0 Then objCADCONTASAPG.CODGRPDESP = CLng(txtCODGRUPDESP.Text)
       objCADCONTASAPG.VLTOTLCTO = CCur(txtVlTotDoc.Text)
       objCADCONTASAPG.DOCPAI = txtDOCPAI.Text
    
       If (flxCONTAPGT.Rows - 1) > 0 Then
          ReDim arrGRIDPGTOS(1 To (flxCONTAPGT.Rows - 1), 1 To 5) As String
          For I = 1 To (flxCONTAPGT.Rows - 1)
              arrGRIDPGTOS(I, 1) = flxCONTAPGT.TextMatrix(I, 1)
              arrGRIDPGTOS(I, 2) = flxCONTAPGT.TextMatrix(I, 2)
              arrGRIDPGTOS(I, 3) = flxCONTAPGT.TextMatrix(I, 3)
              arrGRIDPGTOS(I, 4) = flxCONTAPGT.TextMatrix(I, 4)
              arrGRIDPGTOS(I, 5) = flxCONTAPGT.TextMatrix(I, 6)
          Next I
       End If
    
       objCADCONTASAPG.DOCPGTO = arrGRIDPGTOS
       
    End If
    
    If cTipOper = "B" Or cTipOper = "AB" Then
    
       objCADCONTASAPG.VLPGTO = CCur(txtVlPagto.Text)
       objCADCONTASAPG.DTPAGTO = CDate(mskDTPGTO.Text)
       objCADCONTASAPG.NUMDOC = Trim(lblCODPGTO(7).Caption)
       
       objCADCONTASAPG.TIPOPGTO = 0
       If Len(Trim(txtCODTIPPGTO.Text)) > 0 Then objCADCONTASAPG.TIPOPGTO = CLng(txtCODTIPPGTO.Text)
       
       objCADCONTASAPG.DESCONTOPGTO = 0
       If Len(Trim(lblCODPGTO(14).Caption)) > 0 Then objCADCONTASAPG.DESCONTOPGTO = CCur(lblCODPGTO(14).Caption)
       
       objCADCONTASAPG.PORCDESC = 0
       If Len(Trim(lblCODPGTO(13).Caption)) > 0 Then objCADCONTASAPG.PORCDESC = CCur(lblCODPGTO(13).Caption)
       
       objCADCONTASAPG.ACRESCPGTO = 0
       If Len(Trim(lblCODPGTO(16).Caption)) > 0 Then objCADCONTASAPG.ACRESCPGTO = CCur(lblCODPGTO(16).Caption)
       
       objCADCONTASAPG.PORCACRES = 0
       If Len(Trim(lblCODPGTO(15).Caption)) > 0 Then objCADCONTASAPG.PORCACRES = CCur(lblCODPGTO(15).Caption)
       
       objCADCONTASAPG.PARCPGTO = iParcela
    
    End If
    
    If objCADCONTASAPG.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O titulo foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", IIf(cTipOper = "B", "baixado", IIf(cTipOper = "AB", "alterado", "")))) & " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    If objCADCONTASAPG.Atualiza(cTipOper, Str(objCADCONTASAPG.CODPGTO), FILIAL, Me.Name) = False Then Exit Sub
          
    If cTipOper = "I" Then
       intResp = MsgBox("Deseja incluir novo titulo ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
       If intResp = 7 Then
          Set objBLBFunc = Nothing
          Set objCADCONTASAPG = Nothing
          Set objPESQPADRAO = Nothing
          Unload Me
       Else
          Inclui
          txtCODFORNEC.SetFocus
       End If
    ElseIf cTipOper = "B" Or cTipOper = "AB" Then
       Set objBLBFunc = Nothing
       Set objCADCONTASAPG = Nothing
       Set objPESQPADRAO = Nothing
       Unload Me
    End If
    
End Sub

Private Sub cmdTIPDOC_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADTIPODOC " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Tipo de Documento", "CADTIPODOC.clsCADTIPODOC")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODTIPDOC.Text = varRETORNO
    
    cboTIPCOD.ListIndex = -1
    txtCODTIPDOC.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADCONTASAPG = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPOPGTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_OPERACAO = 1"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    arrCAMPOS(3, 1) = "SGI_SINAL"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Sinal"
    arrCAMPOS(3, 4) = "500"
    arrCAMPOS(3, 5) = "SGI_SINAL"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Tipo de Pagamento", "CADTIPOPGTO.clsCADTIPOPGTO")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODTIPPGTO.Text = varRETORNO
    
    cboTIPOPGTO.ListIndex = -1
    txtCODTIPPGTO.SetFocus

End Sub

Private Sub flxCONTAPGT_DblClick()
    If (flxCONTAPGT.Rows - 1) > 0 And (cTipOper = "I" Or cTipOper = "A") Then CarregaCampos
End Sub


Private Sub flxCONTAPGT_RowColChange()
    If (flxCONTAPGT.Rows - 1) > 0 And (cTipOper = "I" Or cTipOper = "A") Then CarregaCampos
End Sub

Private Sub Form_Activate()
    If cTipOper = "B" Or cTipOper = "AB" Then txtVlPagto.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADCONTASAPG = CreateObject("CADCONTASAPG.clsCADCONTASAPG")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADCONTASAPG.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta
   If cTipOper = "B" Then Baixa
   If cTipOper = "CB" Then ConsBaixa
   If cTipOper = "AB" Then AlteraBaixa

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de contas a pagar - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCONTASAPG
    
    lblCODPGTO(9).Caption = ""
    
    mskDTLANCTO.Text = Format(Date, "DD/MM/YYYY")
    
    objCADCONTASAPG.PreencheComboFornec cboFornec
    objCADCONTASAPG.PreencheComboCondPgto cboCondPgto
    
    objCADCONTASAPG.PreencheCombo cboGrupDesp, "SGI_CADGRUPDESP"
    objCADCONTASAPG.PreencheCombo cboTIPCOD, "SGI_CADTIPODOC"
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGridCondPGTO
    
    StContAPG.Tab = 0
    StContAPG.TabEnabled(1) = False
    
    strSTATGRID = "I"
    intLinhaIndice = 0
    
    cboGrupDesp.Enabled = True
    txtCODGRUPDESP.Enabled = True
    cmdPesqGrupDesp.Enabled = True
   
End Sub

Private Sub mskDTLANCTO_GotFocus()
    objBLBFunc.SelecionaCampos mskDTLANCTO.Name, frmCADCONTASAPG
End Sub

Private Sub mskDTLANCTO_Validate(Cancel As Boolean)

    If Not IsDate(mskDTLANCTO.Text) Then
       MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If
    
    If VerifCaixa(CDate(mskDTLANCTO.Text)) = True Then Cancel = True
    
End Sub

Private Sub mskDTPGTO_GotFocus()
    objBLBFunc.SelecionaCampos mskDTPGTO.Name, frmCADCONTASAPG
End Sub

Private Sub mskDTPGTO_Validate(Cancel As Boolean)

    If IsNull(mskDTPGTO.Text) Then
       MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If
    
    If VerifCaixa(CDate(mskDTPGTO.Text)) = True Then Cancel = True

End Sub

Private Sub txtCODCONDPGT_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCONDPGT.Name, frmCADCONTASAPG
End Sub

Private Sub txtCODCONDPGT_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCONDPGT.Text
End Sub

Private Sub txtCODCONDPGT_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCONDPGT.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCONDPGT.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCONDPGT.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If cboCondPgto.ListCount > 0 And cboCondPgto.ListIndex > -1 Then
       If txtCODCONDPGT.Text <> cboCondPgto.ItemData(cboCondPgto.ListIndex) Then
          ConfGridCondPGTO
          LimpaCamposPGTO
       End If
    End If
    
    cboCondPgto.ListIndex = -1
    For I = 0 To (cboCondPgto.ListCount - 1)
        If cboCondPgto.ItemData(I) = Str(Val(txtCODCONDPGT.Text)) Then cboCondPgto.ListIndex = I
    Next I
    
    If cboCondPgto.ListIndex = -1 Then
       MsgBox "Esta Condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCONDPGT.Text = ""
       Cancel = True
       Exit Sub
    End If
   
    PegaParcelas
    strSTATGRID = "I"

End Sub

Private Sub txtCODFORNEC_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNEC.Name, frmCADCONTASAPG
End Sub

Private Sub txtCODFORNEC_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFORNEC.Text
End Sub

Private Sub txtCODFORNEC_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNEC.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNEC.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNEC.Text = ""
       txtCODGRUPDESP.Text = ""
       cboGrupDesp.ListIndex = -1
       Cancel = True
       Exit Sub
    End If
    
    cboFornec.ListIndex = -1
    For I = 0 To (cboFornec.ListCount - 1)
        If cboFornec.ItemData(I) = Str(Val(txtCODFORNEC.Text)) Then cboFornec.ListIndex = I
    Next I
    
    If cboFornec.ListIndex = -1 Then
       MsgBox "Esta fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNEC.Text = ""
       txtCODGRUPDESP.Text = ""
       cboGrupDesp.ListIndex = -1
       Cancel = True
       Exit Sub
    End If
    
    cboGrupDesp.Enabled = True
    txtCODGRUPDESP.Enabled = True
    cmdPesqGrupDesp.Enabled = True
    
    txtCODGRUPDESP.Text = PegaGrpDesp(cboFornec.ItemData(cboFornec.ListIndex))
    If Len(Trim(txtCODGRUPDESP.Text)) > 0 Then
       If CInt(txtCODGRUPDESP.Text) = 0 Then txtCODGRUPDESP.Text = ""
    End If
    
    cboGrupDesp.ListIndex = -1
    For I = 0 To (cboGrupDesp.ListCount - 1)
        If cboGrupDesp.ItemData(I) = Str(Val(txtCODGRUPDESP.Text)) Then cboGrupDesp.ListIndex = I
    Next I
    
    If cboGrupDesp.ListIndex > -1 Then
        cboGrupDesp.Enabled = False
        txtCODGRUPDESP.Enabled = False
        cmdPesqGrupDesp.Enabled = False
    End If

End Sub

Private Sub txtCODGRUPDESP_GotFocus()
    objBLBFunc.SelecionaCampos txtCODGRUPDESP.Name, frmCADCONTASAPG
End Sub

Private Sub txtCODGRUPDESP_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODGRUPDESP.Text
End Sub

Private Sub txtCODGRUPDESP_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODGRUPDESP.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODGRUPDESP.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUPDESP.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboGrupDesp.ListIndex = -1
    For I = 0 To (cboGrupDesp.ListCount - 1)
        If cboGrupDesp.ItemData(I) = Str(Val(txtCODGRUPDESP.Text)) Then cboGrupDesp.ListIndex = I
    Next I
    
    If cboGrupDesp.ListIndex = -1 Then
       MsgBox "Esta Condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUPDESP.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODTIPDOC_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPDOC.Name, frmCADCONTASAPG
End Sub

Private Sub txtCODTIPDOC_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODTIPDOC.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTIPDOC.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPDOC.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTIPCOD.ListIndex = -1
    For I = 0 To (cboTIPCOD.ListCount - 1)
        If cboTIPCOD.ItemData(I) = Str(Val(txtCODTIPDOC.Text)) Then cboTIPCOD.ListIndex = I
    Next I
    
    If cboTIPCOD.ListIndex = -1 Then
       MsgBox "Esta Condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPDOC.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODTIPPGTO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPPGTO.Name, frmCADCONTASAPG
End Sub

Private Sub txtCODTIPPGTO_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODTIPPGTO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTIPPGTO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPPGTO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTIPOPGTO.ListIndex = -1
    For I = 0 To (cboTIPOPGTO.ListCount - 1)
        If cboTIPOPGTO.ItemData(I) = Trim(Str(Val(txtCODTIPPGTO.Text))) Then cboTIPOPGTO.ListIndex = I
    Next I
    
    If cboTIPOPGTO.ListIndex = -1 Then
       MsgBox "Esta Condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPPGTO.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtDOCPAI_GotFocus()
    objBLBFunc.SelecionaCampos txtDOCPAI.Name, frmCADCONTASAPG
End Sub

Private Sub txtNUMDOC_GotFocus()
    objBLBFunc.SelecionaCampos txtNUMDOC.Name, frmCADCONTASAPG
End Sub

Private Sub txtValor_GotFocus()
    objBLBFunc.SelecionaCampos txtValor.Name, frmCADCONTASAPG
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtValor.Text
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)

    If Len(Trim(txtValor.Text)) = 0 Then Exit Sub

    If Not IsNumeric(txtValor.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If Val(txtValor.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtValor.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtValor.Text = Format(txtValor.Text, "#,##0.00")

End Sub

Private Sub txtVlPagto_GotFocus()
    objBLBFunc.SelecionaCampos txtVlPagto.Name, frmCADCONTASAPG
    lblCODPGTO(14).Caption = ""
    lblCODPGTO(13).Caption = ""
    lblCODPGTO(16).Caption = ""
    lblCODPGTO(15).Caption = ""
End Sub

Private Sub txtVlPagto_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtVlPagto.Text
End Sub

Private Sub txtVlPagto_Validate(Cancel As Boolean)

    If Len(Trim(txtVlPagto.Text)) = 0 Then Exit Sub

    If Not IsNumeric(txtVlPagto.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If Val(txtVlPagto.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtVlPagto.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtVlPagto.Text = Format(txtVlPagto.Text, "#,##0.00")
    
    If CalDesconto > 0 Then lblCODPGTO(14).Caption = Format(CalDesconto, "#,##0.00")
    If CalAcrescimo > 0 Then lblCODPGTO(16).Caption = Format(CalAcrescimo, "#,##0.00")

End Sub

Private Sub txtVlTotDoc_GotFocus()
    objBLBFunc.SelecionaCampos txtVlTotDoc.Name, frmCADCONTASAPG
    ConfGridCondPGTO
    LimpaCamposPGTO
    strSTATGRID = "I"
End Sub

Private Sub txtVlTotDoc_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtVlTotDoc.Text
End Sub

Private Sub PegaParcelas()

    intQtdParc = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCONDPGTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & cboCondPgto.ItemData(cboCondPgto.ListIndex)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
       intQtdParc = BREC!SGI_PARCELAS
       ReDim arrDiasParc(1 To intQtdParc) As Integer
    End If
    
    BREC.Close
    
    '' Pegando Parcelas Para Incrementar a Matrix
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CONDPGTOPARC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & cboCondPgto.ItemData(cboCondPgto.ListIndex)
    sSql = sSql & "Order by SGI_PARC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       arrDiasParc(BREC!SGI_PARC) = BREC!SGI_DIAS
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub

Private Sub txtVlTotDoc_Validate(Cancel As Boolean)

    If Len(Trim(txtVlTotDoc.Text)) = 0 Then Exit Sub

    If Not IsNumeric(txtVlTotDoc.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If Val(txtVlTotDoc.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtVlTotDoc.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtVlTotDoc.Text = Format(txtVlTotDoc.Text, "#,##0.00")
    
    ConfGridCondPGTO
    PopGridDocmentos

End Sub

Private Sub ConfGridCondPGTO()

    flxCONTAPGT.Rows = 1
    flxCONTAPGT.Cols = 7
    
    flxCONTAPGT.TextMatrix(0, 0) = ""
    flxCONTAPGT.TextMatrix(0, 1) = "Numero Doc."
    flxCONTAPGT.TextMatrix(0, 2) = "Data Venc."
    flxCONTAPGT.TextMatrix(0, 3) = "Valor"
    flxCONTAPGT.TextMatrix(0, 4) = "Parcela"
    flxCONTAPGT.TextMatrix(0, 5) = "Tipo. Doc"
    flxCONTAPGT.TextMatrix(0, 6) = "Cod.Tip.Doc"
    
    flxCONTAPGT.ColWidth(0) = 0
    flxCONTAPGT.ColWidth(1) = 1000
    flxCONTAPGT.ColWidth(2) = 1000
    flxCONTAPGT.ColWidth(3) = 1000
    flxCONTAPGT.ColWidth(4) = 700
    flxCONTAPGT.ColWidth(5) = 5000
    flxCONTAPGT.ColWidth(6) = 0
    
    flxCONTAPGT.ColAlignment(5) = 1
    
End Sub

Private Sub PopGridDocmentos()

    Dim curValor     As Double
    Dim I            As Integer
    Dim dtVcto       As Date
    
    If Not IsArray(arrDiasParc) Then Exit Sub
    
    lblPARCELAS(0).Caption = "01/" & Format(UBound(arrDiasParc), "##00")
    
    curValor = (CDbl(txtVlTotDoc.Text) / UBound(arrDiasParc))
    txtValor.Text = Format(curValor, "#,##0.00")
     
    dtVcto = CDate(mskDTLANCTO.Text) + arrDiasParc(flxCONTAPGT.Rows) '' Pega a Data do Vcto
    dtVcto = PegaDiaUtil(dtVcto)
    
    mskDTVencto.Text = Format(dtVcto, "DD/MM/YYYY")
    
End Sub

Private Sub AddGridDoc(strStat As String)

    Dim iLinha As Integer
    
    If (flxCONTAPGT.Rows - 1) >= UBound(arrDiasParc) And strStat = "I" Then
       MsgBox "Parcelas já concluidas !!!", vbOKOnly + vbExclamation, "Aviso"
       LimpaCamposPGTO
       CmdSalva.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtVlTotDoc.Text)) = 0 Then
       MsgBox "Valor do doc. não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
       LimpaCamposPGTO
       txtVlTotDoc.SetFocus
       Exit Sub
    End If
    If CCur(txtVlTotDoc.Text) <= 0 Then
       MsgBox "Valor do doc. não pode ser <= 0 !!!", vbOKOnly + vbExclamation, "Aviso"
       LimpaCamposPGTO
       txtVlTotDoc.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtNUMDOC.Text)) = 0 Then
       MsgBox "Numero de doc. não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
       txtNUMDOC.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtValor.Text)) = 0 Then
       MsgBox "Valor não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
       txtValor.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(txtValor.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtValor.Text = ""
       txtValor.SetFocus
       Exit Sub
    End If
    If CCur(txtValor.Text) <= 0 Then
       MsgBox "Não é permitodo valores <= 0 !!!", vbOKOnly + vbExclamation, "Aviso"
       txtValor.Text = ""
       txtValor.SetFocus
       Exit Sub
    End If
    If Not IsDate(mskDTLANCTO.Text) Then
       MsgBox "Data do Lançamento Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTLANCTO.SetFocus
       Exit Sub
    End If
    If Not IsDate(mskDTVencto.Text) Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTVencto.SetFocus
       Exit Sub
    Else
       If Not IsDate(mskDTLANCTO.Text) Then
          MsgBox "Data do Lançamento Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTLANCTO.SetFocus
          Exit Sub
       End If
       If CDate(mskDTLANCTO.Text) > CDate(mskDTVencto.Text) Then
          MsgBox "Data do Lançamento não pode ser maior que data do Vencimento !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTLANCTO.SetFocus
          Exit Sub
       End If
    End If
    If CCur(txtValor.Text) > CCur(txtVlTotDoc.Text) Then
        MsgBox "O valor da parcela não pode ser maior que valor total do titulo !!!", vbOKOnly + vbExclamation, "Aviso"
        txtValor.SetFocus
        Exit Sub
    End If
    If (cboTIPCOD.ListIndex) = -1 Or Len(Trim(txtCODTIPDOC.Text)) = 0 Then
        MsgBox "Infor o tipo de documento a ser pago !!!", vbOKOnly + vbExclamation, "Aviso"
        cboTIPCOD.SetFocus
        Exit Sub
    End If
    
    If strStat = "I" Then
       flxCONTAPGT.AddItem "" & vbTab & _
                           txtNUMDOC.Text & vbTab & _
                           mskDTVencto.Text & vbTab & _
                           txtValor.Text & vbTab & _
                           lblPARCELAS(0).Caption & vbTab & _
                           IIf(Len(Trim(txtCODTIPDOC.Text)) > 0, cboTIPCOD.Text, "") & vbTab & _
                           txtCODTIPDOC.Text
    
        If flxCONTAPGT.Rows = UBound(arrDiasParc) Then txtValor.Text = Format(CCur(txtValor.Text) - CalcDif, "#,##0.00")
        If flxCONTAPGT.Rows > UBound(arrDiasParc) Then iLinha = (flxCONTAPGT.Rows - 1)
        If flxCONTAPGT.Rows <= UBound(arrDiasParc) Then iLinha = flxCONTAPGT.Rows
    
        lblPARCELAS(0).Caption = Format(iLinha, "##00") & "/" & Format(UBound(arrDiasParc), "##00")
        mskDTVencto.Text = Format(CDate(mskDTLANCTO.Text) + arrDiasParc(iLinha), "DD/MM/YYYY")
    
        If (flxCONTAPGT.Rows - 1) >= UBound(arrDiasParc) Then
           LimpaCamposPGTO
           CmdSalva.SetFocus
        Else
           txtNUMDOC.SetFocus
        End If
    
    End If
    
    If strStat = "A" Then
        
       If intLinhaIndice > 0 Then flxCONTAPGT.TextMatrix(intLinhaIndice, 1) = txtNUMDOC.Text
       If intLinhaIndice > 0 Then flxCONTAPGT.TextMatrix(intLinhaIndice, 2) = mskDTVencto.Text
       If intLinhaIndice > 0 Then flxCONTAPGT.TextMatrix(intLinhaIndice, 5) = cboTIPCOD.Text
       If intLinhaIndice > 0 Then flxCONTAPGT.TextMatrix(intLinhaIndice, 6) = txtCODTIPDOC.Text
       
       intLinhaIndice = 0
       LimpaCamposPGTO
       CmdSalva.SetFocus
    End If
    
    strSTATGRID = "I"
    
End Sub

Private Sub LimpaCamposPGTO()
       txtNUMDOC.Text = ""
       mskDTVencto.Text = "__/__/____"
       txtValor.Text = ""
       lblPARCELAS(0).Caption = ""
       txtCODTIPDOC.Text = ""
       cboTIPCOD.ListIndex = -1
End Sub

Private Function CalcDif() As Currency
    
    Dim curTotDoc As Currency
    Dim curTotParc As Currency
    Dim curTotLanc As Currency
    
    CalcDif = 0
    
    curTotDoc = CCur(txtVlTotDoc.Text)
    curTotLanc = CCur(Format(curTotDoc / UBound(arrDiasParc), "#,##0.00")) * UBound(arrDiasParc)
    
    If curTotLanc > curTotDoc Then CalcDif = (curTotLanc - curTotDoc)
    
End Function

Public Sub CarregaCampos()
    
    txtNUMDOC.Text = flxCONTAPGT.TextMatrix(flxCONTAPGT.Row, 1)
    mskDTVencto.Text = flxCONTAPGT.TextMatrix(flxCONTAPGT.Row, 2)
    txtValor.Text = flxCONTAPGT.TextMatrix(flxCONTAPGT.Row, 3)
    lblPARCELAS(0).Caption = flxCONTAPGT.TextMatrix(flxCONTAPGT.Row, 4)
    txtCODTIPDOC.Text = flxCONTAPGT.TextMatrix(flxCONTAPGT.Row, 6)
    
    If Len(Trim(txtCODTIPDOC.Text)) > 0 Then txtCODTIPDOC_Validate False
    
    ''txtNUMDOC.SetFocus
    strSTATGRID = "A"
    intLinhaIndice = flxCONTAPGT.Row
    
    
    
End Sub

Private Function Valida_Campos() As Boolean

    Valida_Campos = False
    
    If cTipOper = "I" Or cTipOper = "A" Then
    
       If Len(Trim(txtCODFORNEC.Text)) = 0 Then
          MsgBox "Código do fornecedor não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODFORNEC.SetFocus
          Exit Function
       End If
       If Len(Trim(txtCODCONDPGT.Text)) = 0 Then
          MsgBox "Código da condição de pagamento não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODCONDPGT.SetFocus
          Exit Function
       End If
       If Len(Trim(txtVlTotDoc.Text)) = 0 Then
          MsgBox "Valor total do doc. não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
          txtVlTotDoc.SetFocus
          Exit Function
       End If
       If (flxCONTAPGT.Rows - 1) = 0 Then
          MsgBox "Não foram informado parcelas !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODFORNEC.SetFocus
          Exit Function
       End If
       If (flxCONTAPGT.Rows - 1) < UBound(arrDiasParc) Then
          MsgBox "Falta incluir parcelas !!!", vbOKOnly + vbExclamation, "Aviso"
          Exit Function
       End If
    End If
    
    If cTipOper = "B" Then
       If Len(Trim(txtVlPagto.Text)) = 0 Then
          MsgBox "Valor de pagamento não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
          txtVlPagto.SetFocus
          Exit Function
       End If
       If Not IsDate(CDate(mskDTPGTO.Text)) Then
          MsgBox "Data de pagamento inválida !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTPGTO.Text = Format(Now, "DD/MM/YYYY")
          mskDTPGTO.SetFocus
          Exit Function
       End If
       If CDate(mskDTPGTO.Text) < CDate(lblCODPGTO(1).Caption) Then
          MsgBox "Data de pagamento não pode ser menor que a data de lançamento !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTPGTO.Text = Format(Now, "DD/MM/YYYY")
          mskDTPGTO.SetFocus
          Exit Function
       End If
    End If
    
    Valida_Campos = True

End Function

Private Sub Consulta()

    Dim I           As Integer
    Dim j           As Integer
    Dim strTipoPgto As String
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de contas a pagar - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADCONTASAPG
    
    lblCODPGTO(9).Caption = ""
    
    objCADCONTASAPG.PreencheComboFornec cboFornec
    objCADCONTASAPG.PreencheComboCondPgto cboCondPgto
    
    objCADCONTASAPG.PreencheCombo cboGrupDesp, "SGI_CADGRUPDESP"
    objCADCONTASAPG.PreencheCombo cboTIPCOD, "SGI_CADTIPODOC"
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGridCondPGTO
    
    StContAPG.Tab = 0
    StContAPG.TabEnabled(0) = True
    StContAPG.TabEnabled(1) = False
    
    objCADCONTASAPG.CODPGTO = iCodigo
    
    If objCADCONTASAPG.Carrega_campos = True Then
      
       lblCODPGTO(9).Caption = objCADCONTASAPG.CODPGTO
       mskDTLANCTO.Text = Format(objCADCONTASAPG.DATALCTO, "DD/MM/YYYY")
       txtCODFORNEC.Text = objCADCONTASAPG.CODFORN
       txtCODCONDPGT.Text = objCADCONTASAPG.CODCONDPGTO
       txtCODGRUPDESP.Text = IIf(objCADCONTASAPG.CODGRPDESP > 0, objCADCONTASAPG.CODGRPDESP, "")
       txtVlTotDoc.Text = Format(objCADCONTASAPG.VLTOTLCTO, "#,##0.00")
       txtDOCPAI.Text = objCADCONTASAPG.DOCPAI
       
       arrGRIDPGTOS = objCADCONTASAPG.DOCPGTO
       
       '' Fornecedor
       For I = 0 To (cboFornec.ListCount - 1)
           If objCADCONTASAPG.CODFORN = cboFornec.ItemData(I) Then cboFornec.ListIndex = I
       Next I
       
       '' Condição de Pagamento
       For I = 0 To (cboCondPgto.ListCount - 1)
           If objCADCONTASAPG.CODCONDPGTO = cboCondPgto.ItemData(I) Then cboCondPgto.ListIndex = I
       Next I
       
       '' Grupo de Despesas
       For I = 0 To (cboGrupDesp.ListCount - 1)
           If objCADCONTASAPG.CODGRPDESP = cboGrupDesp.ItemData(I) Then cboGrupDesp.ListIndex = I
       Next I
       
       PegaParcelas
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
       
           strTipoPgto = ""
           For j = 0 To (cboTIPCOD.ListCount - 1)
               If Not IsNull(arrGRIDPGTOS(I, 5)) Then
                  If cboTIPCOD.ItemData(j) = Str(Val(arrGRIDPGTOS(I, 5))) Then
                     cboTIPCOD.ListIndex = j
                     strTipoPgto = cboTIPCOD.Text
                     cboTIPCOD.ListIndex = -1
                     Exit For
                  End If
               End If
           Next j
       
           flxCONTAPGT.AddItem "" & vbTab & _
                               arrGRIDPGTOS(I, 1) & vbTab & _
                               Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY") & vbTab & _
                               Format(arrGRIDPGTOS(I, 4), "#,##0.00") & vbTab & _
                               Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00") & vbTab & _
                               strTipoPgto & vbTab & _
                               arrGRIDPGTOS(I, 5)
                               
       Next I
       
    
    End If
    
End Sub


Private Sub Altera()

    Dim I           As Integer
    Dim j           As Integer
    Dim strTipoPgto As String
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de contas a pagar - [ ALTERACAO ]"
    
    objBLBFunc.LimpaCampos frmCADCONTASAPG
    
    lblCODPGTO(9).Caption = ""
    
    objCADCONTASAPG.PreencheComboFornec cboFornec
    objCADCONTASAPG.PreencheComboCondPgto cboCondPgto
    
    objCADCONTASAPG.PreencheCombo cboGrupDesp, "SGI_CADGRUPDESP"
    objCADCONTASAPG.PreencheCombo cboTIPCOD, "SGI_CADTIPODOC"
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGridCondPGTO
    
    StContAPG.Tab = 0
    StContAPG.TabEnabled(0) = True
    StContAPG.TabEnabled(1) = False
    
    objCADCONTASAPG.CODPGTO = iCodigo
    
    If objCADCONTASAPG.Carrega_campos = True Then
      
       lblCODPGTO(9).Caption = objCADCONTASAPG.CODPGTO
       mskDTLANCTO.Text = Format(objCADCONTASAPG.DATALCTO, "DD/MM/YYYY")
       txtCODFORNEC.Text = objCADCONTASAPG.CODFORN
       txtCODCONDPGT.Text = objCADCONTASAPG.CODCONDPGTO
       txtCODGRUPDESP.Text = IIf(objCADCONTASAPG.CODGRPDESP > 0, objCADCONTASAPG.CODGRPDESP, "")
       txtVlTotDoc.Text = Format(objCADCONTASAPG.VLTOTLCTO, "#,##0.00")
       txtDOCPAI.Text = objCADCONTASAPG.DOCPAI
       
       arrGRIDPGTOS = objCADCONTASAPG.DOCPGTO
       
       '' Fornecedor
       For I = 0 To (cboFornec.ListCount - 1)
           If objCADCONTASAPG.CODFORN = cboFornec.ItemData(I) Then cboFornec.ListIndex = I
       Next I
       
       '' Condição de Pagamento
       For I = 0 To (cboCondPgto.ListCount - 1)
           If objCADCONTASAPG.CODCONDPGTO = cboCondPgto.ItemData(I) Then cboCondPgto.ListIndex = I
       Next I
       
       '' Grupo de Despesas
       For I = 0 To (cboGrupDesp.ListCount - 1)
           If objCADCONTASAPG.CODGRPDESP = cboGrupDesp.ItemData(I) Then cboGrupDesp.ListIndex = I
       Next I
       
       If PegaGrpDesp(objCADCONTASAPG.CODFORN) > 0 Then
          txtCODGRUPDESP.Enabled = False
          cmdPesqGrupDesp.Enabled = False
          cboGrupDesp.Enabled = False
       End If
      
       PegaParcelas
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
           
           strTipoPgto = ""
           For j = 0 To (cboTIPCOD.ListCount - 1)
               If Not IsNull(arrGRIDPGTOS(I, 5)) Then
                  If cboTIPCOD.ItemData(j) = Str(Val(arrGRIDPGTOS(I, 5))) Then
                     cboTIPCOD.ListIndex = j
                     strTipoPgto = cboTIPCOD.Text
                     cboTIPCOD.ListIndex = -1
                     Exit For
                  End If
               End If
           Next j
           
           
           flxCONTAPGT.AddItem "" & vbTab & _
                               arrGRIDPGTOS(I, 1) & vbTab & _
                               Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY") & vbTab & _
                               Format(arrGRIDPGTOS(I, 4), "#,##0.00") & vbTab & _
                               Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00") & vbTab & _
                               strTipoPgto & vbTab & _
                               arrGRIDPGTOS(I, 5)
                               
       Next I
       
    
    End If
    
End Sub

Private Function Baixa()

    Dim I           As Integer
    Dim j           As Integer
    Dim strTipoPgto As String
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame7.Enabled = True
   
    Me.Caption = "Cadastro de contas a pagar - [ BAIXA MANUAL ]"
    
    objBLBFunc.LimpaCampos frmCADCONTASAPG
    
    lblCODPGTO(9).Caption = ""
    For I = 0 To (lblCODPGTO.Count - 1)
        lblCODPGTO(I).Caption = ""
    Next I
    
    mskDTPGTO.Text = Format(Date, "DD/MM/YYYY")
    
    objCADCONTASAPG.PreencheComboFornec cboFornec
    objCADCONTASAPG.PreencheComboCondPgto cboCondPgto
    
    objCADCONTASAPG.PreencheCombo cboGrupDesp, "SGI_CADGRUPDESP"
    objCADCONTASAPG.PreencheCombo cboTIPCOD, "SGI_CADTIPODOC"
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGridCondPGTO
    
    StContAPG.Tab = 1
    StContAPG.TabEnabled(0) = False
    StContAPG.TabEnabled(1) = True
    
    objCADCONTASAPG.CODPGTO = iCodigo
    
    If objCADCONTASAPG.Carrega_campos = True Then
      
       lblCODPGTO(0).Caption = objCADCONTASAPG.CODPGTO
       lblCODPGTO(1).Caption = Format(objCADCONTASAPG.DATALCTO, "DD/MM/YYYY")
       
       txtCODFORNEC.Text = objCADCONTASAPG.CODFORN
       txtCODCONDPGT.Text = objCADCONTASAPG.CODCONDPGTO
       txtCODGRUPDESP.Text = IIf(objCADCONTASAPG.CODGRPDESP > 0, objCADCONTASAPG.CODGRPDESP, "")
       
       lblCODPGTO(5).Caption = Format(objCADCONTASAPG.VLTOTLCTO, "#,##0.00")
       lblCODPGTO(6).Caption = objCADCONTASAPG.DOCPAI
       
       arrGRIDPGTOS = objCADCONTASAPG.DOCPGTO
       
       '' Fornecedor
       For I = 0 To (cboFornec.ListCount - 1)
           If objCADCONTASAPG.CODFORN = cboFornec.ItemData(I) Then cboFornec.ListIndex = I
       Next I
       lblCODPGTO(2).Caption = cboFornec.Text
       
       '' Condição de Pagamento
       For I = 0 To (cboCondPgto.ListCount - 1)
           If objCADCONTASAPG.CODCONDPGTO = cboCondPgto.ItemData(I) Then cboCondPgto.ListIndex = I
       Next I
       lblCODPGTO(3).Caption = cboCondPgto.Text
       
       '' Grupo de Despesas
       For I = 0 To (cboGrupDesp.ListCount - 1)
           If objCADCONTASAPG.CODGRPDESP = cboGrupDesp.ItemData(I) Then cboGrupDesp.ListIndex = I
       Next I
       lblCODPGTO(4).Caption = cboGrupDesp.Text
       
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
       
           strTipoPgto = ""
           For j = 0 To (cboTIPCOD.ListCount - 1)
               If Not IsNull(arrGRIDPGTOS(I, 5)) Then
                  If cboTIPCOD.ItemData(j) = Str(Val(arrGRIDPGTOS(I, 5))) Then
                     cboTIPCOD.ListIndex = j
                     strTipoPgto = cboTIPCOD.Text
                     cboTIPCOD.ListIndex = -1
                     Exit For
                  End If
               End If
           Next j
       
           If arrGRIDPGTOS(I, 3) = iParcela Then
           
              lblCODPGTO(7).Caption = arrGRIDPGTOS(I, 1)
              lblCODPGTO(8).Caption = Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY")
              lblCODPGTO(10).Caption = Format(arrGRIDPGTOS(I, 4), "#,##0.00")
              lblCODPGTO(11).Caption = Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00")
              lblCODPGTO(12).Caption = strTipoPgto
              
              txtVlPagto.Text = Format(arrGRIDPGTOS(I, 4), "#,##0.00")
           End If
       Next I
        
    End If


End Function

Private Function CalDesconto() As Currency
    
    CalDesconto = 0
    
    Dim curVLPARC As Currency
    Dim curVLPAGO As Currency
    
    curVLPARC = CCur(lblCODPGTO(10).Caption)
    curVLPAGO = CCur(txtVlPagto.Text)
    
    If curVLPAGO >= curVLPARC Then Exit Function
    
    CalDesconto = (curVLPARC - curVLPAGO)
    lblCODPGTO(13).Caption = Format(((CalDesconto / curVLPARC) * 100), "#,##0.00")
    
End Function

Private Function CalAcrescimo() As Currency

    CalAcrescimo = 0
    
    Dim curVLPARC As Currency
    Dim curVLPAGO As Currency
    
    curVLPARC = CCur(lblCODPGTO(10).Caption)
    curVLPAGO = CCur(txtVlPagto.Text)
    
    If curVLPAGO <= curVLPARC Then Exit Function
    
    CalAcrescimo = (curVLPAGO - curVLPARC)
    lblCODPGTO(15).Caption = Format(((CalAcrescimo / curVLPARC) * 100), "#,##0.00")

End Function

Private Sub ConsBaixa()

    Dim I           As Integer
    Dim j           As Integer
    Dim strTipoPgto As String
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame7.Enabled = False
   
    Me.Caption = "Cadastro de contas a pagar - [ CONSULTA BAIXA ]"
    
    objBLBFunc.LimpaCampos frmCADCONTASAPG
    
    lblCODPGTO(9).Caption = ""
    For I = 0 To (lblCODPGTO.Count - 1)
        lblCODPGTO(I).Caption = ""
    Next I
    
    objCADCONTASAPG.PreencheComboFornec cboFornec
    objCADCONTASAPG.PreencheComboCondPgto cboCondPgto
    
    objCADCONTASAPG.PreencheCombo cboGrupDesp, "SGI_CADGRUPDESP"
    objCADCONTASAPG.PreencheCombo cboTIPCOD, "SGI_CADTIPODOC"
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGridCondPGTO
    
    StContAPG.Tab = 1
    StContAPG.TabEnabled(0) = False
    StContAPG.TabEnabled(1) = True
    
    objCADCONTASAPG.CODPGTO = iCodigo
    objCADCONTASAPG.PARCPGTO = iParcela
    
    If objCADCONTASAPG.Carrega_campos = True Then
      
       lblCODPGTO(0).Caption = objCADCONTASAPG.CODPGTO
       lblCODPGTO(1).Caption = Format(objCADCONTASAPG.DATALCTO, "DD/MM/YYYY")
       
       txtCODFORNEC.Text = objCADCONTASAPG.CODFORN
       txtCODCONDPGT.Text = objCADCONTASAPG.CODCONDPGTO
       txtCODGRUPDESP.Text = IIf(objCADCONTASAPG.CODGRPDESP > 0, objCADCONTASAPG.CODGRPDESP, "")
       
       lblCODPGTO(5).Caption = Format(objCADCONTASAPG.VLTOTLCTO, "#,##0.00")
       lblCODPGTO(6).Caption = objCADCONTASAPG.DOCPAI
       
       arrGRIDPGTOS = objCADCONTASAPG.DOCPGTO
       
       '' Fornecedor
       For I = 0 To (cboFornec.ListCount - 1)
           If objCADCONTASAPG.CODFORN = cboFornec.ItemData(I) Then cboFornec.ListIndex = I
       Next I
       lblCODPGTO(2).Caption = cboFornec.Text
       
       '' Condição de Pagamento
       For I = 0 To (cboCondPgto.ListCount - 1)
           If objCADCONTASAPG.CODCONDPGTO = cboCondPgto.ItemData(I) Then cboCondPgto.ListIndex = I
       Next I
       lblCODPGTO(3).Caption = cboCondPgto.Text
       
       '' Grupo de Despesas
       For I = 0 To (cboGrupDesp.ListCount - 1)
           If objCADCONTASAPG.CODGRPDESP = cboGrupDesp.ItemData(I) Then cboGrupDesp.ListIndex = I
       Next I
       lblCODPGTO(4).Caption = cboGrupDesp.Text
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
           If arrGRIDPGTOS(I, 3) = iParcela Then
       
              strTipoPgto = ""
              For j = 0 To (cboTIPCOD.ListCount - 1)
                 If Not IsNull(arrGRIDPGTOS(I, 5)) Then
                    If cboTIPCOD.ItemData(j) = Str(Val(arrGRIDPGTOS(I, 5))) Then
                       cboTIPCOD.ListIndex = j
                       strTipoPgto = cboTIPCOD.Text
                       cboTIPCOD.ListIndex = -1
                       Exit For
                    End If
                 End If
              Next j
              
              lblCODPGTO(7).Caption = arrGRIDPGTOS(I, 1)
              lblCODPGTO(8).Caption = Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY")
              lblCODPGTO(10).Caption = Format(arrGRIDPGTOS(I, 4), "#,##0.00")
              lblCODPGTO(11).Caption = Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00")
              lblCODPGTO(12).Caption = strTipoPgto
       
           End If
       Next I
       
       '' Iquala Campos já Pagos
       If cTipOper = "CB" Then
          
          mskDTPGTO.Text = Format(objCADCONTASAPG.DTPAGTO, "DD/MM/YYYY")
          txtVlPagto.Text = Format(objCADCONTASAPG.VLPGTO, "#,##0.00")
          
          If objCADCONTASAPG.DESCONTOPGTO > 0 Then lblCODPGTO(14).Caption = Format(objCADCONTASAPG.DESCONTOPGTO, "#,##0.00")
          If objCADCONTASAPG.PORCDESC > 0 Then lblCODPGTO(13).Caption = Format(objCADCONTASAPG.PORCDESC, "#,##0.00")
          If objCADCONTASAPG.ACRESCPGTO > 0 Then lblCODPGTO(16).Caption = Format(objCADCONTASAPG.ACRESCPGTO, "#,##0.00")
          If objCADCONTASAPG.PORCACRES > 0 Then lblCODPGTO(15).Caption = Format(objCADCONTASAPG.PORCACRES, "#,##0.00")
          
          If objCADCONTASAPG.TIPOPGTO > 0 Then
             txtCODTIPPGTO.Text = Trim(Str(objCADCONTASAPG.TIPOPGTO))
             txtCODTIPPGTO_Validate False
          End If
          
       End If
       
        
    End If

End Sub

Private Sub AlteraBaixa()

    Dim I           As Integer
    Dim j           As Integer
    Dim strTipoPgto As String
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame7.Enabled = True
   
    Me.Caption = "Cadastro de contas a pagar - [ ALTERA BAIXA ]"
    
    objBLBFunc.LimpaCampos frmCADCONTASAPG
    
    lblCODPGTO(9).Caption = ""
    For I = 0 To (lblCODPGTO.Count - 1)
        lblCODPGTO(I).Caption = ""
    Next I
    
    objCADCONTASAPG.PreencheComboFornec cboFornec
    objCADCONTASAPG.PreencheComboCondPgto cboCondPgto
    
    objCADCONTASAPG.PreencheCombo cboGrupDesp, "SGI_CADGRUPDESP"
    objCADCONTASAPG.PreencheCombo cboTIPCOD, "SGI_CADTIPODOC"
    objCADCONTASAPG.PreencheCombo cboTIPOPGTO, "SGI_CADTIPOPGTO"
    
    ConfGridCondPGTO
    
    StContAPG.Tab = 1
    StContAPG.TabEnabled(0) = False
    StContAPG.TabEnabled(1) = True
    
    objCADCONTASAPG.CODPGTO = iCodigo
    objCADCONTASAPG.PARCPGTO = iParcela
    
    If objCADCONTASAPG.Carrega_campos = True Then
      
       lblCODPGTO(0).Caption = objCADCONTASAPG.CODPGTO
       lblCODPGTO(1).Caption = Format(objCADCONTASAPG.DATALCTO, "DD/MM/YYYY")
       
       txtCODFORNEC.Text = objCADCONTASAPG.CODFORN
       txtCODCONDPGT.Text = objCADCONTASAPG.CODCONDPGTO
       txtCODGRUPDESP.Text = IIf(objCADCONTASAPG.CODGRPDESP > 0, objCADCONTASAPG.CODGRPDESP, "")
       
       lblCODPGTO(5).Caption = Format(objCADCONTASAPG.VLTOTLCTO, "#,##0.00")
       lblCODPGTO(6).Caption = objCADCONTASAPG.DOCPAI
       
       arrGRIDPGTOS = objCADCONTASAPG.DOCPGTO
       
       '' Fornecedor
       For I = 0 To (cboFornec.ListCount - 1)
           If objCADCONTASAPG.CODFORN = cboFornec.ItemData(I) Then cboFornec.ListIndex = I
       Next I
       lblCODPGTO(2).Caption = cboFornec.Text
       
       '' Condição de Pagamento
       For I = 0 To (cboCondPgto.ListCount - 1)
           If objCADCONTASAPG.CODCONDPGTO = cboCondPgto.ItemData(I) Then cboCondPgto.ListIndex = I
       Next I
       lblCODPGTO(3).Caption = cboCondPgto.Text
       
       '' Grupo de Despesas
       For I = 0 To (cboGrupDesp.ListCount - 1)
           If objCADCONTASAPG.CODGRPDESP = cboGrupDesp.ItemData(I) Then cboGrupDesp.ListIndex = I
       Next I
       lblCODPGTO(4).Caption = cboGrupDesp.Text
       
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
           If arrGRIDPGTOS(I, 3) = iParcela Then
       
              strTipoPgto = ""
              For j = 0 To (cboTIPCOD.ListCount - 1)
                 If Not IsNull(arrGRIDPGTOS(I, 5)) Then
                    If cboTIPCOD.ItemData(j) = Str(Val(arrGRIDPGTOS(I, 5))) Then
                       cboTIPCOD.ListIndex = j
                       strTipoPgto = cboTIPCOD.Text
                       cboTIPCOD.ListIndex = -1
                       Exit For
                    End If
                 End If
              Next j
              
              lblCODPGTO(7).Caption = arrGRIDPGTOS(I, 1)
              lblCODPGTO(8).Caption = Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY")
              lblCODPGTO(10).Caption = Format(arrGRIDPGTOS(I, 4), "#,##0.00")
              lblCODPGTO(11).Caption = Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00")
              lblCODPGTO(12).Caption = strTipoPgto
       
           End If
       Next I
       
       '' Iquala Campos já Pagos
       If cTipOper = "AB" Then
          
          mskDTPGTO.Text = Format(objCADCONTASAPG.DTPAGTO, "DD/MM/YYYY")
          txtVlPagto.Text = Format(objCADCONTASAPG.VLPGTO, "#,##0.00")
          
          If objCADCONTASAPG.DESCONTOPGTO > 0 Then lblCODPGTO(14).Caption = Format(objCADCONTASAPG.DESCONTOPGTO, "#,##0.00")
          If objCADCONTASAPG.PORCDESC > 0 Then lblCODPGTO(13).Caption = Format(objCADCONTASAPG.PORCDESC, "#,##0.00")
          If objCADCONTASAPG.ACRESCPGTO > 0 Then lblCODPGTO(16).Caption = Format(objCADCONTASAPG.ACRESCPGTO, "#,##0.00")
          If objCADCONTASAPG.PORCACRES > 0 Then lblCODPGTO(15).Caption = Format(objCADCONTASAPG.PORCACRES, "#,##0.00")
          
          If objCADCONTASAPG.TIPOPGTO > 0 Then
             txtCODTIPPGTO.Text = Str(objCADCONTASAPG.TIPOPGTO)
             txtCODTIPPGTO_Validate False
          End If
          
       End If
       
        
    End If


End Sub

Private Function VerifCaixa(dtData As Date) As Boolean

    VerifCaixa = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFLXCXHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_DATA   = '" & Format(dtData, "MM/DD/YYYY") & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Existe fluxo de caixa criado !!!", vbOKOnly + vbExclamation, "Aviso"
       If cTipOper = "I" Then mskDTLANCTO.Text = Format(Now, "DD/MM/YYYY")
       If cTipOper = "B" Then mskDTPGTO.Text = Format(Now, "DD/MM/YYYY")
       VerifCaixa = True
    End If
    BREC.Close
    
    If VerifCaixa = True Then Exit Function

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFLXCXHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "Order by SGI_DATA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       If dtData < CDate(Format(BREC!SGI_DATA, "DD/MM/YYYY")) Then
          MsgBox "Data de lançamento menor que data do fluxo de caixa !!!", vbOKOnly + vbExclamation, "Aviso"
          If cTipOper = "I" Then mskDTLANCTO.Text = Format(Now, "DD/MM/YYYY")
          If cTipOper = "B" Then mskDTPGTO.Text = Format(Now, "DD/MM/YYYY")
          VerifCaixa = True
       End If
    End If
    BREC.Close

End Function

Private Function VerifNF() As Boolean
    
    VerifNF = False
    
    '' ----------------------------------------------------------------
    '' Verifica se existe NF
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_NFENTRADACABEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODCONTAPG = " & objCADCONTASAPG.CODPGTO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Este titulo está ligado a uma NF !!!", vbOKOnly + vbExclamation, "Aviso"
       VerifNF = True
    End If
    BREC.Close
    '' ----------------------------------------------------------------

End Function

Private Function PegaGrpDesp(intCODFORNEC As Integer) As Integer

    PegaGrpDesp = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_GRPDESPFORN " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODFORN = " & intCODFORNEC
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaGrpDesp = BREC!SGI_CODGRPDSP
    BREC.Close

End Function

Private Function PegaDiaUtil(DtDiaUtil As Date) As Date
   
    PegaDiaUtil = DtDiaUtil
    Dim intDiaSemana As Integer
    
loopVolta:

    intDiaSemana = Weekday(DtDiaUtil)
    If intDiaSemana = 1 Then DtDiaUtil = (DtDiaUtil + 1) '' Se for Domingo Joga para a segunda
    If intDiaSemana = 7 Then DtDiaUtil = (DtDiaUtil + 2) '' Se for Sabado Joga para a segunda
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDIASUTEIS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_DATENVENTO = '" & Format(DtDiaUtil, "MM/DD/YYYY") & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       intDiaSemana = Weekday(BREC!SGI_DATENVENTO)
       If intDiaSemana = 7 Then
          DtDiaUtil = (BREC!SGI_DATENVENTO + 2)
          BREC.Close
          GoTo loopVolta
       ElseIf intDiaSemana >= 1 And intDiaSemana < 7 Then
          DtDiaUtil = (BREC!SGI_DATENVENTO + 1)
          BREC.Close
          GoTo loopVolta
       End If
    End If
    BREC.Close
    
    PegaDiaUtil = DtDiaUtil
End Function
