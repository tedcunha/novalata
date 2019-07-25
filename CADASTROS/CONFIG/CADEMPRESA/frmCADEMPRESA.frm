VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCADEMPRESA 
   Caption         =   "Cadastro de Empresas"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   0
      TabIndex        =   35
      Top             =   1080
      Width           =   9735
      Begin VB.TextBox txtDtExpira 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   46
         Text            =   "txtDtExpir"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtInscEst 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "txtInscEst"
         Top             =   1320
         Width           =   2295
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2775
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
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
         TabCaption(0)   =   "Endereço de Cobrança"
         TabPicture(0)   =   "frmCADEMPRESA.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label3(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3(3)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label3(4)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label3(5)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label3(6)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label3(7)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtEndCobr"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtBairroCobr"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtCidCobr"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtCepCobr"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtFonCobr"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtFaxCobr"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "cboEstCob"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         TabCaption(1)   =   "Endereço de Entrega"
         TabPicture(1)   =   "frmCADEMPRESA.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cboEstEnt"
         Tab(1).Control(1)=   "txtFaxEnt"
         Tab(1).Control(2)=   "txtFonEnt"
         Tab(1).Control(3)=   "txtCepEnt"
         Tab(1).Control(4)=   "txtCidEnt"
         Tab(1).Control(5)=   "txtBaiEnt"
         Tab(1).Control(6)=   "txtEndEnt"
         Tab(1).Control(7)=   "Label3(14)"
         Tab(1).Control(8)=   "Label3(13)"
         Tab(1).Control(9)=   "Label3(12)"
         Tab(1).Control(10)=   "Label3(11)"
         Tab(1).Control(11)=   "Label3(10)"
         Tab(1).Control(12)=   "Label3(9)"
         Tab(1).Control(13)=   "Label3(8)"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "Logo Relatórios"
         TabPicture(2)   =   "frmCADEMPRESA.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmoAbreArq"
         Tab(2).Control(1)=   "cmdAbreArq"
         Tab(2).Control(2)=   "Assinatura"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "E-Mail"
         TabPicture(3)   =   "frmCADEMPRESA.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtEMAILCC"
         Tab(3).Control(1)=   "txtPortaSMTP"
         Tab(3).Control(2)=   "txtSMTP"
         Tab(3).Control(3)=   "txtSenhaEmail"
         Tab(3).Control(4)=   "txtEmail"
         Tab(3).Control(5)=   "Label4(4)"
         Tab(3).Control(6)=   "Label4(3)"
         Tab(3).Control(7)=   "Label4(2)"
         Tab(3).Control(8)=   "Label4(1)"
         Tab(3).Control(9)=   "Label4(0)"
         Tab(3).ControlCount=   10
         Begin VB.TextBox txtEMAILCC 
            Height          =   285
            Left            =   -73320
            MaxLength       =   100
            TabIndex        =   56
            Text            =   "txtEMAILCC"
            Top             =   1920
            Width           =   5655
         End
         Begin VB.TextBox txtPortaSMTP 
            Height          =   285
            Left            =   -73320
            MaxLength       =   2
            TabIndex        =   55
            Top             =   1560
            Width           =   5655
         End
         Begin VB.TextBox txtSMTP 
            Height          =   285
            Left            =   -73320
            MaxLength       =   100
            TabIndex        =   54
            Text            =   "txtSMTP"
            Top             =   1200
            Width           =   5655
         End
         Begin VB.TextBox txtSenhaEmail 
            Height          =   285
            Left            =   -73320
            MaxLength       =   50
            TabIndex        =   53
            Text            =   "txtSenhaEmail"
            Top             =   840
            Width           =   5655
         End
         Begin VB.TextBox txtEmail 
            Height          =   285
            Left            =   -73320
            MaxLength       =   100
            TabIndex        =   52
            Text            =   "txtEmail"
            Top             =   480
            Width           =   5655
         End
         Begin MSComDlg.CommonDialog cmoAbreArq 
            Left            =   -70320
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.ComboBox cboEstEnt 
            Height          =   315
            Left            =   -70200
            TabIndex        =   25
            Text            =   "cboEstEnt"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.ComboBox cboEstCob 
            Height          =   315
            Left            =   4800
            TabIndex        =   11
            Text            =   "cboEstCob"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtFaxEnt 
            Height          =   285
            Left            =   -73920
            MaxLength       =   30
            TabIndex        =   32
            Text            =   "txtFaxEnt"
            Top             =   2280
            Width           =   4815
         End
         Begin VB.TextBox txtFonEnt 
            Height          =   285
            Left            =   -73920
            MaxLength       =   30
            TabIndex        =   30
            Text            =   "txtFonEnt"
            Top             =   1920
            Width           =   4815
         End
         Begin VB.TextBox txtCepEnt 
            Height          =   285
            Left            =   -73920
            MaxLength       =   9
            TabIndex        =   27
            Text            =   "txtCepEnt"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txtCidEnt 
            Height          =   285
            Left            =   -73920
            MaxLength       =   30
            TabIndex        =   23
            Text            =   "txtCidEnt"
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox txtBaiEnt 
            Height          =   285
            Left            =   -73920
            MaxLength       =   50
            TabIndex        =   21
            Text            =   "txtBaiEnt"
            Top             =   840
            Width           =   4815
         End
         Begin VB.TextBox txtEndEnt 
            Height          =   285
            Left            =   -73920
            MaxLength       =   50
            TabIndex        =   19
            Text            =   "txtEndEnt"
            Top             =   480
            Width           =   4815
         End
         Begin VB.TextBox txtFaxCobr 
            Height          =   285
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   17
            Text            =   "txtFaxCobr"
            Top             =   2280
            Width           =   4815
         End
         Begin VB.TextBox txtFonCobr 
            Height          =   285
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   15
            Text            =   "txtFonCobr"
            Top             =   1920
            Width           =   4815
         End
         Begin VB.TextBox txtCepCobr 
            Height          =   285
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   13
            Text            =   "txtCepCobr"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txtCidCobr 
            Height          =   285
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "txtCidCobr"
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox txtBairroCobr 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "txtBairroCobr"
            Top             =   840
            Width           =   4815
         End
         Begin VB.TextBox txtEndCobr 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "txtEndCobr"
            Top             =   480
            Width           =   4815
         End
         Begin VB.CommandButton cmdAbreArq 
            Caption         =   "Abre Arquivo"
            Height          =   375
            Left            =   -70440
            TabIndex        =   42
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Email CC"
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
            Index           =   4
            Left            =   -74880
            TabIndex        =   51
            Top             =   1920
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Porta SMTP"
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
            Index           =   3
            Left            =   -74880
            TabIndex        =   50
            Top             =   1560
            Width           =   1035
         End
         Begin VB.Label Label4 
            Caption         =   "SMTP"
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
            Left            =   -74880
            TabIndex        =   49
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Senha"
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
            Left            =   -74880
            TabIndex        =   48
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Email"
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
            Left            =   -74880
            TabIndex        =   47
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
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
            Index           =   14
            Left            =   -74880
            TabIndex        =   31
            Top             =   2280
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fone"
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
            Index           =   13
            Left            =   -74880
            TabIndex        =   43
            Top             =   1920
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
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
            Left            =   -74880
            TabIndex        =   26
            Top             =   1560
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
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
            Index           =   11
            Left            =   -70920
            TabIndex        =   24
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
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
            Index           =   10
            Left            =   -74880
            TabIndex        =   22
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
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
            Index           =   9
            Left            =   -74880
            TabIndex        =   20
            Top             =   840
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
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
            Index           =   8
            Left            =   -74880
            TabIndex        =   18
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
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
            Left            =   120
            TabIndex        =   16
            Top             =   2280
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fone"
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
            Index           =   6
            Left            =   120
            TabIndex        =   14
            Top             =   1920
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
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
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
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
            TabIndex        =   10
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
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
            TabIndex        =   8
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
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
            TabIndex        =   6
            Top             =   840
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
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
            TabIndex        =   4
            Top             =   480
            Width           =   825
         End
         Begin VB.Image Assinatura 
            Height          =   2295
            Left            =   -74880
            Stretch         =   -1  'True
            Top             =   360
            Width           =   4365
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   960
         Width           =   2295
         Begin VB.CheckBox chkPadrao 
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
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.TextBox txtCNPJ 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "txtCNPJ"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
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
         Index           =   16
         Left            =   240
         TabIndex        =   45
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Insc. Est:"
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
         Index           =   15
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
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
         Left            =   480
         TabIndex        =   38
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   600
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   9735
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
         Picture         =   "frmCADEMPRESA.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   39
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
         Picture         =   "frmCADEMPRESA.frx":05A2
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   735
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
         Picture         =   "frmCADEMPRESA.frx":06A4
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADEMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho  As String
Public Linha     As Variant
Public cTipOper  As String
Public iCodigo   As Integer
Public strAcesso As String
Dim objBLBFunc   As Object
Dim objEMPRESA   As Object
Dim strNOMARG      As String
Dim strCAMINHO     As String
Private Sub cboEstCob_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboEstCob, KeyAscii
End Sub

Private Sub cboEstEnt_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboEstEnt, KeyAscii
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
    
    
    strCAMINHO = ""
    For I = 0 To (UBound(arrArquivo) - 1)
        strCAMINHO = strCAMINHO & Trim(arrArquivo(I)) + "\"
    Next I
    If IsArray(arrArquivo) Then strNOMARG = Trim(arrArquivo(UBound(arrArquivo)))

End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Empresas - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = True Then
       
       If cTipOper = "I" Then objEMPRESA.EMPCOD = objEMPRESA.Gera_Codigo(Me.Name)
       
       objEMPRESA.DATEXP = "''"
       If Len(Trim(txtDtExpira.Text)) > 0 Then objEMPRESA.DATEXP = "'" & DateSerial(Year(CDate(txtDtExpira.Text)), Month(CDate(txtDtExpira.Text)), Day(CDate(txtDtExpira.Text))) & "'"
       
       objEMPRESA.EMPDESC = txtDescricao.Text
       objEMPRESA.EMPCNPJ = txtCNPJ.Text
       objEMPRESA.Pradrao = chkPadrao.Value
       
       objEMPRESA.EMPENDCOB = txtEndCobr.Text
       objEMPRESA.EMPBAICOB = txtBairroCobr.Text
       objEMPRESA.EMPESTCOB = cboEstCob.Text
       objEMPRESA.EMPCIDCOB = txtCidCobr.Text
       objEMPRESA.EMPCEPCOB = txtCepCobr.Text
       objEMPRESA.EMPFONCOB = txtFonCobr.Text
       objEMPRESA.EMPFAXCOB = txtFaxCobr.Text
       
       objEMPRESA.EMPENDENT = txtEndEnt.Text
       objEMPRESA.EMPBAIENT = txtBaiEnt.Text
       objEMPRESA.EMPESTENT = cboEstEnt.Text
       objEMPRESA.EMPCIDENT = txtCidEnt.Text
       objEMPRESA.EMPCEPENT = txtCepEnt.Text
       objEMPRESA.EMPFONENT = txtFonEnt.Text
       objEMPRESA.EMPFAXENT = txtFaxEnt.Text
       objEMPRESA.EMPINSCEST = txtInscEst.Text
       
       objEMPRESA.ARQUIVO = strNOMARG
       objEMPRESA.CAMINHO = strCAMINHO
       
       If Len(Trim(strNOMARG)) > 0 Then
          '' Gravando Arquivo
          sSql = "Select SGI_IMAGEM From SGI_FILIAL Where SGI_FILIAL = " & objEMPRESA.EMPCOD
          BREC2.CursorType = adOpenDynamic
          BREC2.LockType = adLockOptimistic
          BREC2.Open sSql, adoBanco_Dados
          If Not BREC2.EOF Then
             Call objBLBFunc.GravaBlobParaBanco(BREC2, "SGI_IMAGEM", strCAMINHO + strNOMARG)
             BREC2.Update
          End If
          BREC2.Close
          '' ------------------------------
       End If
       
        objEMPRESA.EMAIL = "Null"
        objEMPRESA.SENHAEMAIL = "Null"
        objEMPRESA.SMTP = "Null"
        objEMPRESA.PORTASMTP = "Null"
        objEMPRESA.EMAILCC = "Null"
       
        ''If Len(Trim(txtEmail.Text)) > 0 Then objEMPRESA.EMAIL = "'" & txtEmail.Text & "'"
        ''If Len(Trim(txtSenhaEmail.Text)) > 0 Then objEMPRESA.SENHAEMAIL = "'" & txtSenhaEmail.Text & "'"
        ''If Len(Trim(txtSMTP.Text)) > 0 Then objEMPRESA.SMTP = "'" & txtSMTP.Text & "'"
        ''If Len(Trim(txtPortaSMTP.Text)) > 0 Then objEMPRESA.PORTASMTP = "'" & txtPortaSMTP.Text & "'"
        ''If Len(Trim(txtEMAILCC.Text)) > 0 Then objEMPRESA.EMAILCC = "'" & txtEMAILCC.Text & "'"
       
       If objEMPRESA.GRAVA(cTipOper) = True Then
          
          MsgBox "A Empresa foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objEMPRESA = Nothing
             Unload Me
          End If
          
       End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objEMPRESA = Nothing
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objEMPRESA = CreateObject("CADEMPRESA.clsCADEMPRESA")
        
    If cTipOper = "I" Then Inclui
    If cTipOper = "A" Then Altera
    If cTipOper = "C" Then Consulta
    
    objBLBFunc.ChecaAcesso frmCADEMPRESA, strAcesso
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objBLBFunc = Nothing
    Set objEMPRESA = Nothing
    Unload Me
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Empresas - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADEMPRESA
    
    txtCodigo.Text = ""
    chkPadrao.Value = 0
    
    strNOMARG = ""
    strCAMINHO = ""
    
    objBLBFunc.Preenche_Estado cboEstCob
    objBLBFunc.Preenche_Estado cboEstEnt
   
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição da Impresa Inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     'If Len(Trim(txtCNPJ.Text)) > 0 Then
        'If objBLBFunc.ViewCGC(txtCNPJ.Text) = False Then
        '   MsgBox "CNPJ Inválido !!!", vbOKOnly + vbCritical, "Aviso"
        '   txtCNPJ.SetFocus
        '  Exit Function
        'End If
     'End If
     
     If Len(Trim(txtDtExpira.Text)) > 0 Then
        If Not IsDate(txtDtExpira.Text) Then
            MsgBox "Senha Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
     End If
     
     
     If cTipOper = "I" Then
        
        sSql = "Select * from SGI_FILIAL Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "A descrição da filial já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
        
        If Len(Trim(txtCNPJ.Text)) > 0 Then
        
           sSql = "Select * from SGI_FILIAL Where SGI_CNPJ ='" & txtCNPJ.Text & "'"
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "O CNPJ da filial já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtCNPJ.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
        
        If chkPadrao.Value = 1 Then
           
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_FILIAL " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_PADRAO = 1"
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              MsgBox "Já existe empresa como padrão !!!", vbOKOnly + vbCritical, "Aviso"
              chkPadrao.Value = 0
              chkPadrao.SetFocus
              BREC.Close
              Exit Function
           End If
           BREC.Close
           
        End If
     
     End If
     
     If cTipOper = "A" Then
        
        If objEMPRESA.EMPDESC <> txtDescricao.Text Then
        
           sSql = "Select * from SGI_FILIAL Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "A descrição da filial já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objEMPRESA.EMPDESC
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
        
        If Len(Trim(txtCNPJ.Text)) > 0 Then
        
           If objEMPRESA.EMPCNPJ <> txtCNPJ.Text Then
           
              sSql = "Select * from SGI_FILIAL Where SGI_CNPJ ='" & txtCNPJ.Text & "'"
              BREC.Open sSql, adoBanco_Dados
        
              If Not BREC.EOF Then
                 MsgBox "O CNPJ da filial já existe !!!", vbOKOnly + vbCritical, "Aviso"
                 txtCNPJ.Text = objEMPRESA.EMPCNPJ
                 txtCNPJ.SetFocus
                 BREC.Close
                 Exit Function
              End If
        
              BREC.Close
           
           End If
           
        End If
        
        If (objEMPRESA.Pradrao <> chkPadrao.Value) And (chkPadrao.Value = 1) Then
           
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_FILIAL " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_PADRAO = 1"
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              MsgBox "Já existe empresa como padrão !!!", vbOKOnly + vbCritical, "Aviso"
              chkPadrao.Value = 0
              chkPadrao.SetFocus
              BREC.Close
              Exit Function
           End If
           BREC.Close
           
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub txtBaiEnt_GotFocus()
    objBLBFunc.SelecionaCampos txtBaiEnt.Name, frmCADEMPRESA
End Sub

Private Sub txtBaiEnt_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtBairroCobr_GotFocus()
    objBLBFunc.SelecionaCampos txtBairroCobr.Name, frmCADEMPRESA
End Sub

Private Sub txtBairroCobr_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCepCobr_GotFocus()
    objBLBFunc.SelecionaCampos txtCepCobr.Name, frmCADEMPRESA
End Sub

Private Sub txtCepCobr_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCepEnt_GotFocus()
    objBLBFunc.SelecionaCampos txtCepEnt.Name, frmCADEMPRESA
End Sub

Private Sub txtCepEnt_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCidCobr_GotFocus()
    objBLBFunc.SelecionaCampos txtCidCobr.Name, frmCADEMPRESA
End Sub

Private Sub txtCidCobr_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCidEnt_GotFocus()
    objBLBFunc.SelecionaCampos txtCidEnt.Name, frmCADEMPRESA
End Sub

Private Sub txtCidEnt_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCNPJ_GotFocus()
    objBLBFunc.SelecionaCampos txtCNPJ.Name, frmCADEMPRESA
End Sub

Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADEMPRESA
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Public Sub Altera()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Empresas - [ ALTERAÇÃO ]"
    
    strNOMARG = ""
    strCAMINHO = ""
    
    objBLBFunc.LimpaCampos frmCADEMPRESA
    
    objBLBFunc.Preenche_Estado cboEstCob
    objBLBFunc.Preenche_Estado cboEstEnt
    
    objEMPRESA.EMPCOD = iCodigo
    
    If objEMPRESA.Carrega_campos = True Then
       txtCodigo.Text = Str(objEMPRESA.EMPCOD)
       txtDescricao.Text = objEMPRESA.EMPDESC
       txtCNPJ.Text = objEMPRESA.EMPCNPJ
       chkPadrao.Value = objEMPRESA.Pradrao
    
       txtEndCobr.Text = objEMPRESA.EMPENDCOB
       txtBairroCobr.Text = objEMPRESA.EMPBAICOB
       cboEstCob.Text = objEMPRESA.EMPESTCOB
       txtCidCobr.Text = objEMPRESA.EMPCIDCOB
       txtCepCobr.Text = objEMPRESA.EMPCEPCOB
       txtFonCobr.Text = objEMPRESA.EMPFONCOB
       txtFaxCobr.Text = objEMPRESA.EMPFAXCOB
       
       txtEndEnt.Text = objEMPRESA.EMPENDENT
       txtBaiEnt.Text = objEMPRESA.EMPBAIENT
       cboEstEnt.Text = objEMPRESA.EMPESTENT
       txtCidEnt.Text = objEMPRESA.EMPCIDENT
       txtCepEnt.Text = objEMPRESA.EMPCEPENT
       txtFonEnt.Text = objEMPRESA.EMPFONENT
       txtFaxEnt.Text = objEMPRESA.EMPFAXENT
       txtInscEst.Text = objEMPRESA.EMPINSCEST
       
       strNOMARG = objEMPRESA.ARQUIVO
       strCAMINHO = objEMPRESA.CAMINHO
    
       txtDtExpira.Text = objEMPRESA.DATEXP
       
       If Len(Trim(strNOMARG)) > 0 Then
          sSql = "Select SGI_IMAGEM from SGI_FILIAL Where SGI_FILIAL = " & objEMPRESA.EMPCOD
          BREC2.Open sSql, adoBanco_Dados, adOpenDynamic, adLockOptimistic
          If Not BREC2.EOF Then
             Call objBLBFunc.LeCampoBlobDoDB(BREC2, "SGI_IMAGEM", strNOMARG)
             Assinatura.Picture = LoadPicture(strCAMINHO + strNOMARG)
          End If
          BREC2.Close
       End If
    
        txtEmail.Text = objEMPRESA.EMAIL
        txtSenhaEmail.Text = objEMPRESA.SENHAEMAIL
        txtSMTP.Text = objEMPRESA.SMTP
        txtPortaSMTP.Text = objEMPRESA.PORTASMTP
        txtEMAILCC.Text = objEMPRESA.EMAILCC
    
    End If
    
End Sub

Private Sub Consulta()

    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    strNOMARG = ""
    strCAMINHO = ""
    
    Me.Caption = "Cadastro de Empresas - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADEMPRESA
    
    objEMPRESA.EMPCOD = iCodigo
    
    objBLBFunc.Preenche_Estado cboEstCob
    objBLBFunc.Preenche_Estado cboEstEnt
    
    If objEMPRESA.Carrega_campos = True Then
       txtCodigo.Text = Str(objEMPRESA.EMPCOD)
       txtDescricao.Text = objEMPRESA.EMPDESC
       txtCNPJ.Text = objEMPRESA.EMPCNPJ
       chkPadrao.Value = objEMPRESA.Pradrao
       
       txtEndCobr.Text = objEMPRESA.EMPENDCOB
       txtBairroCobr.Text = objEMPRESA.EMPBAICOB
       cboEstCob.Text = objEMPRESA.EMPESTCOB
       txtCidCobr.Text = objEMPRESA.EMPCIDCOB
       txtCepCobr.Text = objEMPRESA.EMPCEPCOB
       txtFonCobr.Text = objEMPRESA.EMPFONCOB
       txtFaxCobr.Text = objEMPRESA.EMPFAXCOB
       
       txtEndEnt.Text = objEMPRESA.EMPENDENT
       txtBaiEnt.Text = objEMPRESA.EMPBAIENT
       cboEstEnt.Text = objEMPRESA.EMPESTENT
       txtCidEnt.Text = objEMPRESA.EMPCIDENT
       txtCepEnt.Text = objEMPRESA.EMPCEPENT
       txtFonEnt.Text = objEMPRESA.EMPFONENT
       txtFaxEnt.Text = objEMPRESA.EMPFAXENT
       txtInscEst.Text = objEMPRESA.EMPINSCEST
       
       strNOMARG = objEMPRESA.ARQUIVO
       strCAMINHO = objEMPRESA.CAMINHO
       
       txtDtExpira.Text = objEMPRESA.DATEXP
       
       If Len(Trim(strNOMARG)) > 0 Then
          sSql = "Select SGI_IMAGEM from SGI_FILIAL Where SGI_FILIAL = " & objEMPRESA.EMPCOD
          BREC2.Open sSql, adoBanco_Dados, adOpenDynamic, adLockOptimistic
          If Not BREC2.EOF Then
             Call objBLBFunc.LeCampoBlobDoDB(BREC2, "SGI_IMAGEM", strNOMARG)
             Assinatura.Picture = LoadPicture(strCAMINHO + strNOMARG)
          End If
          BREC2.Close
       End If
    
        txtEmail.Text = objEMPRESA.EMAIL
        txtSenhaEmail.Text = objEMPRESA.SENHAEMAIL
        txtSMTP.Text = objEMPRESA.SMTP
        txtPortaSMTP.Text = objEMPRESA.PORTASMTP
        txtEMAILCC.Text = objEMPRESA.EMAILCC
    
    End If

End Sub

Private Sub txtEndCobr_GotFocus()
    objBLBFunc.SelecionaCampos txtEndCobr.Name, frmCADEMPRESA
End Sub

Private Sub txtEndCobr_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtEndEnt_GotFocus()
    objBLBFunc.SelecionaCampos txtEndEnt.Name, frmCADEMPRESA
End Sub

Private Sub txtEndEnt_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFaxCobr_GotFocus()
    objBLBFunc.SelecionaCampos txtFaxCobr.Name, frmCADEMPRESA
End Sub

Private Sub txtFaxCobr_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFaxEnt_GotFocus()
    objBLBFunc.SelecionaCampos txtFaxEnt.Name, frmCADEMPRESA
End Sub

Private Sub txtFaxEnt_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFonCobr_GotFocus()
    objBLBFunc.SelecionaCampos txtFonCobr.Name, frmCADEMPRESA
End Sub

Private Sub txtFonCobr_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFonEnt_GotFocus()
    objBLBFunc.SelecionaCampos txtFonEnt.Name, frmCADEMPRESA
End Sub

Private Sub txtFonEnt_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub
