VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADCONTRECEB 
   Caption         =   "Cadastro de Contas a Receber Manual"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab StContAPG 
      Height          =   5655
      Left            =   0
      TabIndex        =   18
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9975
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
      TabPicture(0)   =   "frmCADCONTRECEB.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Dados da Baixa"
      TabPicture(1)   =   "frmCADCONTRECEB.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame7 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   62
         Top             =   3720
         Width           =   8055
         Begin VB.TextBox txtCODTIPPGTO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2355
            MaxLength       =   10
            TabIndex        =   65
            Text            =   "txtCODTIPP"
            Top             =   615
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTRECEB.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   620
            Width           =   375
         End
         Begin VB.TextBox txtVlPagto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2355
            MaxLength       =   15
            TabIndex        =   63
            Text            =   "txtVlPagto"
            Top             =   240
            Width           =   1935
         End
         Begin MSMask.MaskEdBox mskDTPGTO 
            Height          =   285
            Left            =   6690
            TabIndex        =   66
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPORACRESC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblPORACRESC"
            Height          =   285
            Left            =   5880
            TabIndex        =   79
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lblVALACRESC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVALACRESC"
            Height          =   285
            Left            =   2355
            TabIndex        =   78
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lblPORCDESCTO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblPORCDESCTO"
            Height          =   285
            Left            =   5880
            TabIndex        =   77
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblVALDESCTO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVALDESCTO"
            Height          =   285
            Left            =   2355
            TabIndex        =   76
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label33 
            Caption         =   "% Acrescimo"
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
            Left            =   4680
            TabIndex        =   74
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label32 
            Caption         =   "% Desconto"
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
            Left            =   4680
            TabIndex        =   73
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label31 
            Caption         =   "Data do Pgto"
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
            Left            =   4680
            TabIndex        =   72
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Acrescimo"
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
            TabIndex        =   71
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Desconto"
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
            TabIndex        =   70
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label28 
            Caption         =   "Tipo de Pagto"
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
            TabIndex        =   69
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblDESCTPPGTO 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCTPPGTO"
            Height          =   285
            Left            =   3480
            TabIndex        =   68
            Top             =   600
            Width           =   4455
         End
         Begin VB.Label Label27 
            Caption         =   "Valor pago"
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
            TabIndex        =   67
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   8055
         Begin VB.Label lblDESCGRPRECEBBAIXA 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCGRPRECEBBAIXA"
            Height          =   285
            Left            =   2400
            TabIndex        =   88
            Top             =   1530
            Width           =   5535
         End
         Begin VB.Label lblDESCTIPODOCBAIXA 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCTIPODOCBAIXA"
            Height          =   285
            Left            =   2400
            TabIndex        =   87
            Top             =   1200
            Width           =   5535
         End
         Begin VB.Label lblDESCCONDPGTOBAIXA 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCCONDPGTOBAIXA"
            Height          =   285
            Left            =   2400
            TabIndex        =   86
            Top             =   870
            Width           =   5535
         End
         Begin VB.Label lblDOCPAIBAIXA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDOCPAIBAIXA"
            Height          =   285
            Left            =   6360
            TabIndex        =   82
            Top             =   1860
            Width           =   1575
         End
         Begin VB.Label lblVLTOTDOC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLTOTDOC"
            Height          =   285
            Left            =   2400
            TabIndex        =   81
            Top             =   1860
            Width           =   1815
         End
         Begin VB.Label lblDTLANCTO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDTLANCTO"
            Height          =   285
            Left            =   6360
            TabIndex        =   80
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblCODBAIXA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCODBAIXA"
            Height          =   285
            Left            =   2400
            TabIndex        =   61
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblDESCCLIENTEBAIXA 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCCLIENTEBAIXA"
            Height          =   285
            Left            =   2400
            TabIndex        =   59
            Top             =   540
            Width           =   5535
         End
         Begin VB.Label Label22 
            Caption         =   "Documento Pai"
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
            Left            =   4680
            TabIndex        =   54
            Top             =   1860
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Data do Lançamento"
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
            Left            =   4440
            TabIndex        =   53
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Valor Total Doc"
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
            TabIndex        =   52
            Top             =   1860
            Width           =   2175
         End
         Begin VB.Label Label19 
            Caption         =   "Grupo de Recebimento"
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
            TabIndex        =   51
            Top             =   1530
            Width           =   2175
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de documento"
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
            TabIndex        =   50
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label17 
            Caption         =   "Condição de pagamento"
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
            TabIndex        =   49
            Top             =   870
            Width           =   2295
         End
         Begin VB.Label Label16 
            Caption         =   "Cliente"
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
            TabIndex        =   48
            Top             =   540
            Width           =   615
         End
         Begin VB.Label Label15 
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   -74880
         TabIndex        =   25
         Top             =   2640
         Width           =   8055
         Begin VB.Label lblVALDOC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVALDOC"
            Height          =   285
            Left            =   2400
            TabIndex        =   85
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblDTVENCTO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDTVENCTO"
            Height          =   285
            Left            =   6360
            TabIndex        =   84
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblParcela 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblParcela"
            Height          =   285
            Left            =   6360
            TabIndex        =   83
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblNUMDOCBAIXA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblNUMDOCBAIXA"
            Height          =   285
            Left            =   2400
            TabIndex        =   75
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label26 
            Caption         =   "Parcela"
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
            Left            =   4680
            TabIndex        =   58
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label25 
            Caption         =   "Data vencimento"
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
            Left            =   4680
            TabIndex        =   57
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label24 
            Caption         =   "Valor"
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
            TabIndex        =   56
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label23 
            Caption         =   "Numero do documento"
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
            TabIndex        =   55
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   120
         TabIndex        =   23
         Top             =   4080
         Width           =   8055
         Begin MSFlexGridLib.MSFlexGrid flxCONTAPGT 
            Height          =   1215
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2143
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   8055
         Begin VB.TextBox txtNUMDOC 
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "txtNUMDOC"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   9
            Text            =   "txtValor"
            Top             =   550
            Width           =   1695
         End
         Begin VB.CommandButton cmbGravPGT 
            Height          =   315
            Left            =   7560
            Picture         =   "frmCADCONTRECEB.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   480
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskDTVencto 
            Height          =   285
            Left            =   6360
            TabIndex        =   10
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            Caption         =   "Parcela"
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
            Left            =   4320
            TabIndex        =   46
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Data vencimento"
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
            Left            =   4320
            TabIndex        =   45
            Top             =   285
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Valor"
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
            TabIndex        =   44
            Top             =   570
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Numero do documento"
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
            TabIndex        =   43
            Top             =   285
            Width           =   1935
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
            Left            =   6360
            TabIndex        =   11
            Top             =   555
            Width           =   1200
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   8055
         Begin VB.TextBox txtCODGRPRECEB 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   5
            Text            =   "txtCODGRPRECEB"
            Top             =   1920
            Width           =   735
         End
         Begin VB.CommandButton cmdGRPRECEB 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTRECEB.frx":023C
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1920
            Width           =   375
         End
         Begin VB.CommandButton cmdBANCO 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTRECEB.frx":033E
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtCODBANCO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   4
            Text            =   "txtCODBANCO"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtCODTIPDOC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   3
            Text            =   "txtCODTIPDOC"
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmdTIPDOC 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTRECEB.frx":0440
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1200
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskDTLANCTO 
            Height          =   285
            Left            =   6600
            TabIndex        =   0
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtCODCLI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   1
            Text            =   "txtCODCLI"
            Top             =   530
            Width           =   735
         End
         Begin VB.TextBox txtCODCONDPGT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txtCODCONDPGT"
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdPesqFor 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTRECEB.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   510
            Width           =   375
         End
         Begin VB.CommandButton cmdPesqCondPgt 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCADCONTRECEB.frx":0644
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtVlTotDoc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Text            =   "txtVlTotDoc"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txtDOCPAI 
            Height          =   285
            Left            =   6600
            MaxLength       =   10
            TabIndex        =   7
            Text            =   "txtDOCPAI"
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblCODTIT 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCODTIT"
            Height          =   285
            Left            =   2400
            TabIndex        =   60
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Data do Lançamento"
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
            Left            =   4680
            TabIndex        =   42
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label9 
            Caption         =   "Nota Fiscal Numero"
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
            Left            =   4800
            TabIndex        =   41
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "Valor Total Documento"
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
            TabIndex        =   40
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Grupo de Recebimento"
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
            TabIndex        =   39
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "Banco"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Condição de pagamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   555
            Width           =   615
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblDESCGRPRECEB 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCGRPRECEB"
            Height          =   285
            Left            =   3480
            TabIndex        =   33
            Top             =   1920
            Width           =   4455
         End
         Begin VB.Label lblDESCBANCO 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCBANCO"
            Height          =   285
            Left            =   3480
            TabIndex        =   32
            Top             =   1560
            Width           =   4455
         End
         Begin VB.Label lblDESCTIPDOC 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCTIPDOC"
            Height          =   285
            Left            =   3480
            TabIndex        =   31
            Top             =   1200
            Width           =   4455
         End
         Begin VB.Label lblDESCCONDPGTO 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCCONDPGTO"
            Height          =   285
            Left            =   3480
            TabIndex        =   30
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label lblDESCCLIENTE 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESCCLIENTE"
            Height          =   285
            Left            =   3480
            TabIndex        =   29
            Top             =   530
            Width           =   4455
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8295
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
         Picture         =   "frmCADCONTRECEB.frx":0746
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmCADCONTRECEB.frx":0C78
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmCADCONTRECEB.frx":0D7A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADCONTRECEB"
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
Public lngCodUsuario As Long
Dim objBLBFunc       As Object
Dim objCADCONTRECEB  As Object
Dim objPESQPADRAO    As Object
Dim arrGRIDPGTOS     As Variant
Dim intQtdParc       As Integer
Dim arrDiasParc      As Variant
Dim strSTATGRID      As String
Dim intLinhaIndice   As Integer

Private Sub cmbGravPGT_Click()
    If cTipOper = "I" Or cTipOper = "A" Then AddGridDoc strSTATGRID
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    If cTipOper <> "CB" Then
    
        '' Verifica se há baixados
        sSql = "Select" & vbCrLf
        sSql = sSql & "      * " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIARC " & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_CODIGO = " & objCADCONTRECEB.CODPGTO & vbCrLf
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
   
        Me.Caption = "Cadastro de Contas a Receber Manual - [ ALTERAÇÃO ]"
    
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
   
        Me.Caption = "Cadastro de Contas a Receber Manual - [ ALTERA BAIXA ]"
    
        cTipOper = "AB"
        txtVlPagto.SetFocus
    
    End If

End Sub

Private Sub cmdBANCO_Click()

    ReDim arrCAMPOS(1 To 4, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Bancos", "CADBANCOS.clsCADBANCOS")
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODBANCO.Text = varRETORNO
       Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADBANCOS", txtCODBANCO, lblDESCBANCO)
    End If
    
    txtCODBANCO.SetFocus

End Sub

Private Sub cmdGRPRECEB_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPREC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de Recebimento", "CADGRUPREC.clsCADGRUPREC")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODGRPRECEB.Text = varRETORNO
       Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADGRUPREC", txtCODGRPRECEB, lblDESCGRPRECEB)
    End If
    txtCODGRPRECEB.SetFocus

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
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Condição de Pagamento", "CADCONDPAGTO.clsCADCONDPAGTO")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODCONDPGT.Text = varRETORNO
       Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADCONDPGTO", txtCODCONDPGT, lblDESCCONDPGTO)
    End If
    txtCODCONDPGT.SetFocus

End Sub

Private Sub cmdPesqFor_Click()

    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
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
    arrCAMPOS(3, 4) = "3000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    arrCAMPOS(4, 1) = "SGI_NOMFANTA"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Nome Fantasia"
    arrCAMPOS(4, 4) = "2000"
    arrCAMPOS(4, 5) = "SGI_NOMFANTA"
    
    arrCAMPOS(5, 1) = "SGI_CIDNORM"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Cidade"
    arrCAMPOS(5, 4) = "1500"
    arrCAMPOS(5, 5) = "SGI_CIDNORM"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODCLI.Text = varRETORNO
       Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCODCLI.Text, lblDESCCLIENTE)
    End If
    txtCODCLI.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I           As Integer
    Dim intResp     As Integer
    Dim lngCodLog   As Long
    
    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Or cTipOper = "A" Then
       
       If cTipOper = "I" Then objCADCONTRECEB.CODPGTO = objBLBFunc.Gera_Codigo(Me.Name, FILIAL, Linha)
       
       objCADCONTRECEB.DATALCTO = CDate(mskDTLANCTO.Text)
       objCADCONTRECEB.CODFORN = txtCODCLI.Text
       objCADCONTRECEB.CODCONDPGTO = txtCODCONDPGT.Text
       objCADCONTRECEB.CODTIPDOC = txtCODTIPDOC.Text
       If Len(Trim(txtCODBANCO.Text)) > 0 Then objCADCONTRECEB.CODBANCO = txtCODBANCO.Text
       objCADCONTRECEB.VLTOTLCTO = CCur(txtVlTotDoc.Text)
       objCADCONTRECEB.DOCPAI = txtDOCPAI.Text
       objCADCONTRECEB.GRPRECEB = CInt(txtCODGRPRECEB.Text)
    
       If (flxCONTAPGT.Rows - 1) > 0 Then
          ReDim arrGRIDPGTOS(1 To (flxCONTAPGT.Rows - 1), 1 To 4) As String
          For I = 1 To (flxCONTAPGT.Rows - 1)
              arrGRIDPGTOS(I, 1) = flxCONTAPGT.TextMatrix(I, 1)
              arrGRIDPGTOS(I, 2) = flxCONTAPGT.TextMatrix(I, 2)
              arrGRIDPGTOS(I, 3) = flxCONTAPGT.TextMatrix(I, 3)
              arrGRIDPGTOS(I, 4) = flxCONTAPGT.TextMatrix(I, 4)
          Next I
       End If
    
       objCADCONTRECEB.DOCPGTO = arrGRIDPGTOS
       
    End If
    
    If cTipOper = "B" Or cTipOper = "AB" Then
    
       objCADCONTRECEB.VLPGTO = CCur(txtVlPagto.Text)
       objCADCONTRECEB.DTPAGTO = CDate(mskDTPGTO.Text)
       objCADCONTRECEB.NUMDOC = Trim(lblNUMDOCBAIXA.Caption)
    
       objCADCONTRECEB.TIPOPGTO = 0
       If Len(Trim(txtCODTIPPGTO.Text)) > 0 Then objCADCONTRECEB.TIPOPGTO = CLng(txtCODTIPPGTO.Text)
    
       objCADCONTRECEB.DESCONTOPGTO = 0
       If Len(Trim(lblVALDESCTO.Caption)) > 0 Then objCADCONTRECEB.DESCONTOPGTO = CCur(lblVALDESCTO.Caption)
    
       objCADCONTRECEB.PORCDESC = 0
       If Len(Trim(lblPORCDESCTO.Caption)) > 0 Then objCADCONTRECEB.PORCDESC = CCur(lblPORCDESCTO.Caption)
    
       objCADCONTRECEB.ACRESCPGTO = 0
       If Len(Trim(lblVALACRESC.Caption)) > 0 Then objCADCONTRECEB.ACRESCPGTO = CCur(lblVALACRESC.Caption)
    
       objCADCONTRECEB.PORCACRES = 0
       If Len(Trim(lblPORACRESC.Caption)) > 0 Then objCADCONTRECEB.PORCACRES = CCur(lblPORACRESC.Caption)
    
       objCADCONTRECEB.PARCPGTO = iParcela
    
    End If
    
    If objCADCONTRECEB.GRAVA(cTipOper) = False Then Exit Sub
          
    '' Atualizando os Dados
    If objBLBFunc.Atualiza(cTipOper, Str(objCADCONTRECEB.CODPGTO), FILIAL, Me.Name, Linha) = False Then Exit Sub

    '' Gerando Log de Sistema
    lngCodLog = objBLBFunc.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objBLBFunc.GravaLogModulo(FILIAL, lngCodLog, Me.Name, cTipOper, lngCodUsuario, Str(objCADCONTRECEB.CODPGTO), Linha)
    
    MsgBox "O titulo foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", IIf(cTipOper = "B", "baixado", IIf(cTipOper = "AB", "alterado", "")))) & " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       intResp = MsgBox("Deseja incluir novo titulo ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
       If intResp = 7 Then
          Set objBLBFunc = Nothing
          Set objCADCONTRECEB = Nothing
          Set objPESQPADRAO = Nothing
          Unload Me
       Else
          Inclui
          txtCODCLI.SetFocus
       End If
    ElseIf cTipOper = "B" Or cTipOper = "AB" Then
       Set objBLBFunc = Nothing
       Set objCADCONTRECEB = Nothing
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
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Tipo de documento", "CADTIPODOC.clsCADTIPODOC")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODTIPDOC.Text = varRETORNO
       Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPODOC", txtCODTIPDOC, lblDESCTIPDOC)
    End If
    txtCODTIPDOC.SetFocus

End Sub

Private Sub cmdVoltar_Click()
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
    sSql = sSql & "   And SGI_OPERACAO = 2"
    
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
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODTIPPGTO.Text = varRETORNO
       Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPOPGTO", txtCODTIPPGTO.Text, lblDESCTPPGTO)
    End If
    txtCODTIPPGTO.SetFocus

End Sub

Private Sub flxCONTAPGT_DblClick()
    If (flxCONTAPGT.Rows - 1) > 0 And (cTipOper = "I" Or cTipOper = "A") Then CarregaCampos
End Sub

Private Sub flxCONTAPGT_RowColChange()
    If (flxCONTAPGT.Rows - 1) > 0 And (cTipOper = "I" Or cTipOper = "A") Then CarregaCampos
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADCONTRECEB = CreateObject("CADCONTRECEB.clsCADCONTRECEB")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADCONTRECEB.FILIAL = FILIAL
   
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
   
    Me.Caption = "Cadastro de Contas a Receber Manual - [ INCLUSÃO ]"
    
    Call objBLBFunc.LimpaCampos(frmCADCONTRECEB)
    
    lblCODBAIXA.Caption = ""
    
    mskDTLANCTO.Text = Format(Date, "DD/MM/YYYY")
    
    Call ConfGridCondPGTO
    Call LimpaLabels
    
    StContAPG.Tab = 0
    StContAPG.TabEnabled(1) = False
    
    strSTATGRID = "I"
    intLinhaIndice = 0
   
End Sub

Private Sub ConfGridCondPGTO()

    flxCONTAPGT.Rows = 1
    flxCONTAPGT.Cols = 5
    
    flxCONTAPGT.TextMatrix(0, 0) = ""
    flxCONTAPGT.TextMatrix(0, 1) = "Numero Doc."
    flxCONTAPGT.TextMatrix(0, 2) = "Data Venc."
    flxCONTAPGT.TextMatrix(0, 3) = "Valor"
    flxCONTAPGT.TextMatrix(0, 4) = "Parcela"
    
    flxCONTAPGT.ColWidth(0) = 0
    flxCONTAPGT.ColWidth(1) = 1000
    flxCONTAPGT.ColWidth(2) = 1000
    flxCONTAPGT.ColWidth(3) = 1000
    flxCONTAPGT.ColWidth(4) = 700
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyObjetos
End Sub


Private Sub mskDTLANCTO_GotFocus()
    objBLBFunc.SelecionaCampos mskDTLANCTO.Name, frmCADCONTRECEB
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
    objBLBFunc.SelecionaCampos mskDTPGTO.Name, frmCADCONTRECEB
End Sub

Private Sub mskDTPGTO_Validate(Cancel As Boolean)
    
    If IsNull(mskDTPGTO.Text) Then
       MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If
    
    If VerifCaixa(CDate(mskDTPGTO.Text)) = True Then Cancel = True

End Sub

Private Sub txtCODBANCO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODBANCO.Name, frmCADCONTRECEB
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
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADBANCOS", txtCODBANCO.Text, lblDESCBANCO)
    If Len(Trim(lblDESCBANCO.Caption)) = 0 Then
       txtCODBANCO.Text = ""
       Cancel = True
    End If
    
End Sub

Private Sub txtCODCLI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLI.Name, frmCADCONTRECEB
End Sub

Private Sub txtCODCLI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLI.Text
End Sub

Private Sub txtCODCLI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCLI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLI.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCODCLI.Text, lblDESCCLIENTE)
    If Len(Trim(lblDESCCLIENTE.Caption)) = 0 Then
       txtCODCLI.Text = ""
       Cancel = True
    End If
        
End Sub

Private Sub txtCODCONDPGT_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCONDPGT.Name, frmCADCONTRECEB
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
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADCONDPGTO", txtCODCONDPGT.Text, lblDESCCONDPGTO)
    If Len(Trim(lblDESCCONDPGTO.Caption)) = 0 Then
       txtCODCONDPGT.Text = ""
       Cancel = True
    End If
    Call ConfGridCondPGTO
    Call LimpaCamposPGTO
    Call PegaParcelas
    strSTATGRID = "I"

End Sub

Private Sub LimpaCamposPGTO()
       txtNUMDOC.Text = ""
       mskDTVencto.Text = "__/__/____"
       txtValor.Text = ""
       lblPARCELAS(0).Caption = ""
End Sub

Private Sub PegaParcelas()

    intQtdParc = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCONDPGTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(txtCODCONDPGT.Text)
    
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
    sSql = sSql & "   And SGI_CODIGO = " & Trim(txtCODCONDPGT.Text)
    sSql = sSql & "Order by SGI_PARC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       arrDiasParc(BREC!SGI_PARC) = BREC!SGI_DIAS
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub

Private Sub txtCODGRPRECEB_GotFocus()
    objBLBFunc.SelecionaCampos txtCODGRPRECEB.Name, frmCADCONTRECEB
End Sub

Private Sub txtCODGRPRECEB_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODGRPRECEB.Text
End Sub

Private Sub txtCODGRPRECEB_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODGRPRECEB.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODGRPRECEB.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRPRECEB.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADGRUPREC", txtCODGRPRECEB.Text, lblDESCGRPRECEB)
    If Len(Trim(lblDESCGRPRECEB.Caption)) = 0 Then
       txtCODGRPRECEB.Text = ""
       Cancel = True
    End If
        
End Sub

Private Sub txtCODTIPDOC_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPDOC.Name, frmCADCONTRECEB
End Sub

Private Sub txtCODTIPDOC_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODTIPDOC.Text
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
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPODOC", txtCODTIPDOC.Text, lblDESCTIPDOC)
    If Len(Trim(lblDESCTIPDOC.Caption)) = 0 Then
       txtCODTIPDOC.Text = ""
       Cancel = True
    End If
    
    
End Sub

Private Sub txtCODTIPPGTO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPPGTO.Name, frmCADCONTRECEB
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
    
       Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPOPGTO", txtCODTIPPGTO.Text, lblDESCTPPGTO)
    If Len(Trim(lblDESCTPPGTO.Caption)) = 0 Then
       txtCODTIPPGTO.Text = ""
       Cancel = True
    End If
    
End Sub

Private Sub txtDOCPAI_GotFocus()
    objBLBFunc.SelecionaCampos txtDOCPAI.Name, frmCADCONTRECEB
End Sub

Private Sub txtNUMDOC_GotFocus()
    objBLBFunc.SelecionaCampos txtNUMDOC.Name, frmCADCONTRECEB
End Sub

Private Sub txtVlPagto_GotFocus()
    objBLBFunc.SelecionaCampos txtVlPagto.Name, frmCADCONTRECEB
    lblVALDESCTO.Caption = ""
    lblVALACRESC.Caption = ""
    lblPORCDESCTO.Caption = ""
    lblPORACRESC.Caption = ""
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
    
    If CalDesconto > 0 Then lblVALDESCTO.Caption = Format(CalDesconto, "#,##0.00")
    If CalAcrescimo > 0 Then lblVALACRESC.Caption = Format(CalAcrescimo, "#,##0.00")

End Sub

Private Sub txtVlTotDoc_GotFocus()
    objBLBFunc.SelecionaCampos txtVlTotDoc.Name, frmCADCONTRECEB
    ConfGridCondPGTO
    LimpaCamposPGTO
    strSTATGRID = "I"
End Sub

Private Sub txtVlTotDoc_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtVlTotDoc.Text
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
    If Not IsDate(mskDTVencto.Text) Then
       MsgBox "Data de Vencimento inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTVencto.SetFocus
       Exit Sub
    Else
       If Not IsDate(mskDTLANCTO.Text) Then
          MsgBox "Data de Lançamento inválida !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTLANCTO.SetFocus
          Exit Sub
       End If
       If CDate(mskDTLANCTO.Text) > CDate(mskDTVencto.Text) Then
          MsgBox "Data de Vencimento não pode ser menor que date de Lançamento !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTVencto.SetFocus
          Exit Sub
       End If
    End If
    
    If strStat = "I" Then
       flxCONTAPGT.AddItem "" & vbTab & _
                           txtNUMDOC.Text & vbTab & _
                           mskDTVencto.Text & vbTab & _
                           txtValor.Text & vbTab & _
                           lblPARCELAS(0).Caption

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
       
       intLinhaIndice = 0
       LimpaCamposPGTO
       CmdSalva.SetFocus
    End If
    
    strSTATGRID = "I"
    
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
    
    strSTATGRID = "A"
    intLinhaIndice = flxCONTAPGT.Row
    
End Sub

Private Function Valida_Campos() As Boolean

    Valida_Campos = False
    
    If cTipOper = "I" Or cTipOper = "A" Then
       If Len(Trim(txtCODCLI.Text)) = 0 Then
          MsgBox "Código do cliente não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODCLI.SetFocus
          Exit Function
       End If
       If Len(Trim(txtCODCONDPGT.Text)) = 0 Then
          MsgBox "Código da condição de pagamento não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODCONDPGT.SetFocus
          Exit Function
       End If
       If Len(Trim(txtCODTIPDOC.Text)) = 0 Then
          MsgBox "Código de tipo de documento não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODTIPDOC.SetFocus
          Exit Function
       End If
       If Len(Trim(txtVlTotDoc.Text)) = 0 Then
          MsgBox "Valor total do doc. não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
          txtVlTotDoc.SetFocus
          Exit Function
       End If
       If (flxCONTAPGT.Rows - 1) = 0 Then
          MsgBox "Não foram informado parcelas !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODCLI.SetFocus
          Exit Function
       End If
       If Len(Trim(txtCODGRPRECEB.Text)) = 0 Then
          MsgBox "Informe o Grupo de Recebimento !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCODGRPRECEB.SetFocus
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
        If CDate(mskDTPGTO.Text) < CDate(lblDTLANCTO.Caption) Then
           MsgBox "Data de pagamento não pode ser menor que data de lançamento !!!", vbOKOnly + vbExclamation, "Aviso"
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
   
    Me.Caption = "Cadastro de Contas a Receber Manual - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADCONTRECEB
    
    lblCODBAIXA.Caption = ""
    
    Call ConfGridCondPGTO
    Call LimpaLabels
    
    StContAPG.Tab = 0
    StContAPG.TabEnabled(0) = True
    StContAPG.TabEnabled(1) = False
    
    objCADCONTRECEB.CODPGTO = iCodigo
    
    If objCADCONTRECEB.Carrega_campos = True Then
      
       lblCODTIT.Caption = objCADCONTRECEB.CODPGTO
       mskDTLANCTO.Text = Format(objCADCONTRECEB.DATALCTO, "DD/MM/YYYY")
       txtCODCLI.Text = objCADCONTRECEB.CODFORN
       txtCODCONDPGT.Text = objCADCONTRECEB.CODCONDPGTO
       txtCODTIPDOC.Text = objCADCONTRECEB.CODTIPDOC
       If objCADCONTRECEB.CODBANCO > 0 Then txtCODBANCO.Text = objCADCONTRECEB.CODBANCO
       txtVlTotDoc.Text = Format(objCADCONTRECEB.VLTOTLCTO, "#,##0.00")
       txtDOCPAI.Text = objCADCONTRECEB.DOCPAI
       If objCADCONTRECEB.GRPRECEB > 0 Then txtCODGRPRECEB.Text = objCADCONTRECEB.GRPRECEB
              
       arrGRIDPGTOS = objCADCONTRECEB.DOCPGTO
       
       Call PopulaLabels
       Call PegaParcelas
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
           flxCONTAPGT.AddItem "" & vbTab & _
                               arrGRIDPGTOS(I, 1) & vbTab & _
                               Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY") & vbTab & _
                               Format(arrGRIDPGTOS(I, 4), "#,##0.00") & vbTab & _
                               Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00")
                               
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
   
    Me.Caption = "Cadastro de Contas a Receber Manual - [ ALTERACAO ]"
    
    objBLBFunc.LimpaCampos frmCADCONTRECEB
    
    lblCODTIT.Caption = ""
    
    ConfGridCondPGTO
    Call LimpaLabels
    
    StContAPG.Tab = 0
    StContAPG.TabEnabled(0) = True
    StContAPG.TabEnabled(1) = False
    
    objCADCONTRECEB.CODPGTO = iCodigo
    
    If objCADCONTRECEB.Carrega_campos = True Then
      
       lblCODTIT.Caption = objCADCONTRECEB.CODPGTO
       mskDTLANCTO.Text = Format(objCADCONTRECEB.DATALCTO, "DD/MM/YYYY")
       txtCODCLI.Text = objCADCONTRECEB.CODFORN
       txtCODCONDPGT.Text = objCADCONTRECEB.CODCONDPGTO
       txtCODTIPDOC.Text = objCADCONTRECEB.CODTIPDOC
       If objCADCONTRECEB.CODBANCO > 0 Then txtCODBANCO.Text = objCADCONTRECEB.CODBANCO
       txtVlTotDoc.Text = Format(objCADCONTRECEB.VLTOTLCTO, "#,##0.00")
       txtDOCPAI.Text = objCADCONTRECEB.DOCPAI
       If objCADCONTRECEB.GRPRECEB > 0 Then txtCODGRPRECEB.Text = objCADCONTRECEB.GRPRECEB
       
       arrGRIDPGTOS = objCADCONTRECEB.DOCPGTO
       
       Call PegaParcelas
       Call PopulaLabels
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
           
           flxCONTAPGT.AddItem "" & vbTab & _
                               arrGRIDPGTOS(I, 1) & vbTab & _
                               Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY") & vbTab & _
                               Format(arrGRIDPGTOS(I, 4), "#,##0.00") & vbTab & _
                               Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00")
                               
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
   
    Me.Caption = "Cadastro de Contas a Receber Manual - [ BAIXA MANUAL ]"
    
    objBLBFunc.LimpaCampos frmCADCONTRECEB
    
    Call LimpaLabels
    Call LimpaLabelsAcrDesc
    
    
    mskDTPGTO.Text = Format(Date, "DD/MM/YYYY")
    
    Call ConfGridCondPGTO
    
    StContAPG.Tab = 1
    StContAPG.TabEnabled(0) = False
    StContAPG.TabEnabled(1) = True
    
    objCADCONTRECEB.CODPGTO = iCodigo
    
    If objCADCONTRECEB.Carrega_campos = True Then
      
       lblCODBAIXA.Caption = objCADCONTRECEB.CODPGTO
       lblDTLANCTO.Caption = Format(objCADCONTRECEB.DATALCTO, "DD/MM/YYYY")
       
       txtCODCLI.Text = objCADCONTRECEB.CODFORN
       txtCODCONDPGT.Text = objCADCONTRECEB.CODCONDPGTO
       txtCODTIPDOC.Text = objCADCONTRECEB.CODTIPDOC
       If objCADCONTRECEB.GRPRECEB > 0 Then txtCODGRPRECEB.Text = objCADCONTRECEB.GRPRECEB
       
       lblVLTOTDOC.Caption = Format(objCADCONTRECEB.VLTOTLCTO, "#,##0.00")
       lblDOCPAIBAIXA.Caption = objCADCONTRECEB.DOCPAI
       
       arrGRIDPGTOS = objCADCONTRECEB.DOCPGTO
       
       Call PopLabBaixa
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
       
           If arrGRIDPGTOS(I, 3) = iParcela Then
           
              lblNUMDOCBAIXA.Caption = arrGRIDPGTOS(I, 1)
              lblDTVENCTO.Caption = Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY")
              lblVALDOC.Caption = Format(arrGRIDPGTOS(I, 4), "#,##0.00")
              lblParcela.Caption = Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00")
              txtVlPagto.Text = Format(arrGRIDPGTOS(I, 4), "#,##0.00")
              
           End If
       Next I
        
    End If


End Function

Private Function CalDesconto() As Currency
    
    CalDesconto = 0
    
    Dim curVLPARC As Currency
    Dim curVLPAGO As Currency
    
    curVLPARC = CCur(lblVALDOC.Caption)
    curVLPAGO = CCur(txtVlPagto.Text)
    
    If curVLPAGO >= curVLPARC Then Exit Function
    
    CalDesconto = (curVLPARC - curVLPAGO)
    lblPORCDESCTO.Caption = Format(((CalDesconto / curVLPARC) * 100), "#,##0.00")
    
End Function

Private Function CalAcrescimo() As Currency

    CalAcrescimo = 0
    
    Dim curVLPARC As Currency
    Dim curVLPAGO As Currency
    
    curVLPARC = CCur(lblVALDOC.Caption)
    curVLPAGO = CCur(txtVlPagto.Text)
    
    If curVLPAGO <= curVLPARC Then Exit Function
    
    CalAcrescimo = (curVLPAGO - curVLPARC)
    lblPORACRESC.Caption = Format(((CalAcrescimo / curVLPARC) * 100), "#,##0.00")

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
   
    Me.Caption = "Cadastro de Contas a Receber Manual - [ CONSULTA BAIXA ]"
    
    objBLBFunc.LimpaCampos frmCADCONTRECEB
    Call LimpaLabels
    Call LimpaLabelsAcrDesc
    
    Call ConfGridCondPGTO
    
    StContAPG.Tab = 1
    StContAPG.TabEnabled(0) = False
    StContAPG.TabEnabled(1) = True
    
    objCADCONTRECEB.CODPGTO = iCodigo
    objCADCONTRECEB.PARCPGTO = iParcela
    
    If objCADCONTRECEB.Carrega_campos = True Then
      
       lblCODBAIXA.Caption = objCADCONTRECEB.CODPGTO
       lblDTLANCTO.Caption = Format(objCADCONTRECEB.DATALCTO, "DD/MM/YYYY")
       
       txtCODCLI.Text = objCADCONTRECEB.CODFORN
       txtCODCONDPGT.Text = objCADCONTRECEB.CODCONDPGTO
       txtCODTIPDOC.Text = objCADCONTRECEB.CODTIPDOC
       If objCADCONTRECEB.GRPRECEB > 0 Then txtCODGRPRECEB.Text = objCADCONTRECEB.GRPRECEB
       
       
       lblVLTOTDOC.Caption = Format(objCADCONTRECEB.VLTOTLCTO, "#,##0.00")
       lblDOCPAIBAIXA.Caption = objCADCONTRECEB.DOCPAI
       
       arrGRIDPGTOS = objCADCONTRECEB.DOCPGTO
       
       Call PopLabBaixa
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
           If arrGRIDPGTOS(I, 3) = iParcela Then
              
              lblNUMDOCBAIXA.Caption = arrGRIDPGTOS(I, 1)
              lblDTVENCTO.Caption = Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY")
              lblVALDOC.Caption = Format(arrGRIDPGTOS(I, 4), "#,##0.00")
              lblParcela.Caption = Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00")
       
           End If
       Next I
       
       '' Iquala Campos já Pagos
       If cTipOper = "CB" Then
          
          mskDTPGTO.Text = Format(objCADCONTRECEB.DTPAGTO, "DD/MM/YYYY")
          txtVlPagto.Text = Format(objCADCONTRECEB.VLPGTO, "#,##0.00")
          
          If objCADCONTRECEB.DESCONTOPGTO > 0 Then lblVALDESCTO.Caption = Format(objCADCONTRECEB.DESCONTOPGTO, "#,##0.00")
          If objCADCONTRECEB.PORCDESC > 0 Then lblPORCDESCTO.Caption = Format(objCADCONTRECEB.PORCDESC, "#,##0.00")
          If objCADCONTRECEB.ACRESCPGTO > 0 Then lblVALACRESC.Caption = Format(objCADCONTRECEB.ACRESCPGTO, "#,##0.00")
          If objCADCONTRECEB.PORCACRES > 0 Then lblPORACRESC.Caption = Format(objCADCONTRECEB.PORCACRES, "#,##0.00")
          
          If objCADCONTRECEB.TIPOPGTO > 0 Then
             txtCODTIPPGTO.Text = Str(objCADCONTRECEB.TIPOPGTO)
             Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPOPGTO", txtCODTIPPGTO.Text, lblDESCTPPGTO)
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
   
    Me.Caption = "Cadastro de Contas a Receber Manual - [ ALTERA BAIXA ]"
    
    objBLBFunc.LimpaCampos frmCADCONTRECEB
    Call LimpaLabels
    Call LimpaLabelsAcrDesc
    
    Call ConfGridCondPGTO
    
    StContAPG.Tab = 1
    StContAPG.TabEnabled(0) = False
    StContAPG.TabEnabled(1) = True
    
    objCADCONTRECEB.CODPGTO = iCodigo
    objCADCONTRECEB.PARCPGTO = iParcela
    
    If objCADCONTRECEB.Carrega_campos = True Then
       
       If objCADCONTRECEB.GRPRECEB > 0 Then txtCODGRPRECEB.Text = objCADCONTRECEB.GRPRECEB
       
       lblCODBAIXA.Caption = objCADCONTRECEB.CODPGTO
       lblDTLANCTO.Caption = Format(objCADCONTRECEB.DATALCTO, "DD/MM/YYYY")
       
       txtCODCLI.Text = objCADCONTRECEB.CODFORN
       txtCODCONDPGT.Text = objCADCONTRECEB.CODCONDPGTO
       txtCODTIPDOC.Text = objCADCONTRECEB.CODTIPDOC
       If objCADCONTRECEB.GRPRECEB > 0 Then txtCODGRPRECEB.Text = objCADCONTRECEB.GRPRECEB
       
       lblVLTOTDOC.Caption = Format(objCADCONTRECEB.VLTOTLCTO, "#,##0.00")
       lblDOCPAIBAIXA.Caption = objCADCONTRECEB.DOCPAI
       
       arrGRIDPGTOS = objCADCONTRECEB.DOCPGTO
       
       '' Preenchendo grid de titulos
       For I = 1 To UBound(arrGRIDPGTOS)
           If arrGRIDPGTOS(I, 3) = iParcela Then
              lblNUMDOCBAIXA.Caption = arrGRIDPGTOS(I, 1)
              lblDTVENCTO.Caption = Format(arrGRIDPGTOS(I, 2), "DD/MM/YYYY")
              lblVALDOC.Caption = Format(arrGRIDPGTOS(I, 4), "#,##0.00")
              lblParcela.Caption = Format(arrGRIDPGTOS(I, 3), "##00") & "/" & Format(UBound(arrGRIDPGTOS), "##00")
           End If
       Next I
       
       '' Iquala Campos já Pagos
       If cTipOper = "AB" Then
          
          mskDTPGTO.Text = Format(objCADCONTRECEB.DTPAGTO, "DD/MM/YYYY")
          txtVlPagto.Text = Format(objCADCONTRECEB.VLPGTO, "#,##0.00")
          
          If objCADCONTRECEB.DESCONTOPGTO > 0 Then lblVALDESCTO.Caption = Format(objCADCONTRECEB.DESCONTOPGTO, "#,##0.00")
          If objCADCONTRECEB.PORCDESC > 0 Then lblPORCDESCTO.Caption = Format(objCADCONTRECEB.PORCDESC, "#,##0.00")
          If objCADCONTRECEB.ACRESCPGTO > 0 Then lblVALACRESC.Caption = Format(objCADCONTRECEB.ACRESCPGTO, "#,##0.00")
          If objCADCONTRECEB.PORCACRES > 0 Then lblPORACRESC.Caption = Format(objCADCONTRECEB.PORCACRES, "#,##0.00")
          
          If objCADCONTRECEB.TIPOPGTO > 0 Then
             txtCODTIPPGTO.Text = Str(objCADCONTRECEB.TIPOPGTO)
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

Private Sub LimpaLabels()
    lblDESCCLIENTE.Caption = ""
    lblDESCCONDPGTO.Caption = ""
    lblDESCTIPDOC.Caption = ""
    lblDESCBANCO.Caption = ""
    lblDESCGRPRECEB.Caption = ""
    lblDESCTPPGTO.Caption = ""
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCodigo As String, lblLabel As Variant)

    lblLabel.Caption = ""
    
    If Len(Trim(strCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCodigo)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub PopulaLabels()
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCODCLI.Text, lblDESCCLIENTE)
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADCONDPGTO", txtCODCONDPGT.Text, lblDESCCONDPGTO)
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPODOC", txtCODTIPDOC.Text, lblDESCTIPDOC)
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADBANCOS", txtCODBANCO.Text, lblDESCBANCO)
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADGRUPREC", txtCODGRPRECEB.Text, lblDESCGRPRECEB)
End Sub

Private Sub DestroyObjetos()
    Set objBLBFunc = Nothing
    Set objCADCONTRECEB = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub PopLabBaixa()
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCODCLI.Text, lblDESCCLIENTEBAIXA)
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADCONDPGTO", txtCODCONDPGT.Text, lblDESCCONDPGTOBAIXA)
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPODOC", txtCODTIPDOC.Text, lblDESCTIPODOCBAIXA)
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADGRUPREC", txtCODGRPRECEB.Text, lblDESCGRPRECEBBAIXA)
End Sub

Private Sub LimpaLabelsAcrDesc()
    lblVALDESCTO.Caption = ""
    lblPORCDESCTO.Caption = ""
    lblVALACRESC.Caption = ""
    lblPORACRESC.Caption = ""
End Sub
