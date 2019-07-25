VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDEPARAGERAL 
   Caption         =   "De Para Geral"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   11400
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5953
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
      TabCaption(0)   =   "Pedidos de Vendedores"
      TabPicture(0)   =   "frmDEPARAGERAL.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDescVendedor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDescVendedorF"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDescClienteF"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCODVEND"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCODVENDF"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraPeriodo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fraProgVend"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command7"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCIDCLIEF"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Carteira de Clientes"
      TabPicture(1)   =   "frmDEPARAGERAL.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Command8"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "Command5"
      Tab(1).Control(5)=   "Command4"
      Tab(1).Control(6)=   "txtCODVENDC2"
      Tab(1).Control(7)=   "txtCODVENDC1"
      Tab(1).Control(8)=   "lblDescVendedorC2"
      Tab(1).Control(9)=   "lblDescVendedorC1"
      Tab(1).Control(10)=   "Label1(3)"
      Tab(1).Control(11)=   "Label1(2)"
      Tab(1).ControlCount=   12
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   -74880
         TabIndex        =   45
         Top             =   2280
         Width           =   10935
         Begin VB.PictureBox pgbClie 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   10665
            TabIndex        =   46
            Top             =   240
            Width           =   10695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Label4"
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
            TabIndex        =   47
            Top             =   600
            Width           =   585
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Iniciar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66120
         TabIndex        =   44
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtCIDCLIEF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   42
         Text            =   "txtCIDCLIEF"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   3840
         Picture         =   "frmDEPARAGERAL.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1920
         Width           =   375
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   375
         Left            =   -72120
         TabIndex        =   37
         Top             =   1200
         Width           =   8175
         Begin VB.TextBox txtCIDCLIE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   0
            MaxLength       =   10
            TabIndex        =   39
            Text            =   "txtCIDCLIE"
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   1200
            Picture         =   "frmDEPARAGERAL.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblDescCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescCliente"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1560
            TabIndex        =   40
            Top             =   0
            Width           =   6615
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -75000
         TabIndex        =   33
         Top             =   1200
         Width           =   2655
         Begin VB.OptionButton optTodosClie 
            Caption         =   "Unico Cliente"
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
            Left            =   1080
            TabIndex        =   35
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton optTodosClie 
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   -70920
         Picture         =   "frmDEPARAGERAL.frx":023C
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   -70920
         Picture         =   "frmDEPARAGERAL.frx":033E
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCODVENDC2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   28
         Text            =   "txtCODVEND"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCODVENDC1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "txtCODVEND"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Iniciar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   24
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Frame fraProgVend 
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   10935
         Begin MSComctlLib.ProgressBar pgbVendedor 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Min             =   1e-4
            Max             =   100
            Scrolling       =   1
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
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
            TabIndex        =   23
            Top             =   600
            Width           =   585
         End
      End
      Begin VB.Frame fraPeriodo 
         Caption         =   "[ Periodo ]"
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
         Left            =   2640
         TabIndex        =   20
         Top             =   1200
         Width           =   3135
         Begin MSMask.MaskEdBox mskDTINI 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDTFIN 
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   1440
            TabIndex        =   21
            Top             =   240
            Width           =   120
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   1440
         Width           =   2775
         Begin VB.OptionButton optPeriodo 
            Caption         =   "Periodo"
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
            Left            =   1320
            TabIndex        =   19
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optPeriodo 
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
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   3840
         Picture         =   "frmDEPARAGERAL.frx":0440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtCODVENDF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Text            =   "txtCODVENDF"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCODVEND 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCODVEND"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3840
         Picture         =   "frmDEPARAGERAL.frx":0542
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblDescClienteF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescClienteF"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4200
         TabIndex        =   43
         Top             =   1920
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unico Cliente"
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
         TabIndex        =   36
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label lblDescVendedorC2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendedorC2"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -70560
         TabIndex        =   32
         Top             =   840
         Width           =   6615
      End
      Begin VB.Label lblDescVendedorC1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendedorC1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -70560
         TabIndex        =   31
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mudar Clientes do vendedor"
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
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Para o vendedor"
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
         Left            =   -74880
         TabIndex        =   25
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label lblDescVendedorF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendedorF"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4200
         TabIndex        =   16
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Para o vendedor"
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
         TabIndex        =   14
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label lblDescVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendedor"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4200
         TabIndex        =   13
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mudar Pedidos do vendedor"
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
         TabIndex        =   11
         Top             =   480
         Width           =   2400
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11055
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
         Picture         =   "frmDEPARAGERAL.frx":0644
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   5775
         Begin VB.OptionButton optEmpresa 
            Caption         =   "NOVALATA/STEEL ROL"
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
            Left            =   3240
            TabIndex        =   9
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton optEmpresa 
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
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optEmpresa 
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
            Index           =   1
            Left            =   1560
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmDEPARAGERAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Dim objBLBFunc      As Object
Dim objPESQPADRAO   As Object

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

On Error GoTo Err_Command1_Click


    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL     = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVENDF.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedorF)
    If Len(Trim(lblDescVendedorF.Caption)) = 0 Then txtCODVENDF.Text = ""
    
    If txtCODVENDF.Enabled = True Then txtCODVENDF.SetFocus

    Exit Sub
    
Err_Command1_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command1_Click()", Me.Name, "Command1_Click()", strCAMARQERRO)

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
    sSql = sSql & "        SGI_FILIAL     = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVEND.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedor)
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then txtCODVEND.Text = ""
    
    If txtCODVEND.Enabled = True Then txtCODVEND.SetFocus

    Exit Sub
    
Err_Command2_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command2_Click()", Me.Name, "Command2_Click()", strCAMARQERRO)

End Sub

Private Sub Command3_Click()
    If ConfereCampos = False Then Exit Sub

End Sub

Private Sub Command4_Click()

On Error GoTo Err_Command4_Click


    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL     = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVENDC1.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedorC1)
    If Len(Trim(lblDescVendedorC1.Caption)) = 0 Then txtCODVENDC1.Text = ""
    
    If txtCODVENDC1.Enabled = True Then txtCODVENDC1.SetFocus

    Exit Sub
    
Err_Command4_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command4_Click()", Me.Name, "Command4_Click()", strCAMARQERRO)

End Sub

Private Sub Command5_Click()

On Error GoTo Err_Command5_Click


    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL     = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVENDC2.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedorC2)
    If Len(Trim(lblDescVendedorC2.Caption)) = 0 Then txtCODVENDC2.Text = ""
    
    If txtCODVENDC2.Enabled = True Then txtCODVENDC2.SetFocus

    Exit Sub
    
Err_Command5_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command5_Click()", Me.Name, "Command5_Click()", strCAMARQERRO)

End Sub

Private Sub Command6_Click()

On Error GoTo Err_Command6_Click

    If Len(Trim(txtCODVENDC1.Text)) = 0 Then
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
    sSql = sSql & "       CVEN.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CVEN.SGI_CODIGO = " & txtCODVENDC1.Text & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = CVEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = CVEN.SGI_CODCLI"
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCIDCLIE.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescCliente)
    If Len(Trim(lblDescCliente.Caption)) = 0 Then txtCIDCLIE.Text = ""
    
    If txtCIDCLIE.Enabled = True Then txtCIDCLIE.SetFocus

    Exit Sub
    
Err_Command6_Click:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command6_Click()", Me.Name, "Command6_Click()", strCAMARQERRO)

End Sub

Private Sub Command7_Click()

On Error GoTo Err_Command7_Click

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
    sSql = sSql & "       CVEN.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CVEN.SGI_CODIGO = " & txtCODVEND.Text & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = CVEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = CVEN.SGI_CODCLI"
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCIDCLIEF.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescClienteF)
    If Len(Trim(lblDescClienteF.Caption)) = 0 Then txtCIDCLIEF.Text = ""
    
    If txtCIDCLIEF.Enabled = True Then txtCIDCLIEF.SetFocus

    Exit Sub
    
Err_Command7_Click:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command7_Click()", Me.Name, "Command7_Click()", strCAMARQERRO)

End Sub

Private Sub Command8_Click()
    If ConfereCampos = False Then Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   strCAMARQERRO = Right(Linha(9), Len(Trim(Linha(9))) - 8)
   
   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
    
   Call objBLBFunc.LimpaCampos(Me)
   Call LimpaCamposLabel
   
   optEmpresa(0).Value = True
   optPeriodo(0).Value = True
   optTodosClie(0).Value = True

   fraProgVend.Visible = False
   Frame6.Visible = False
   
   SSTab1.Tab = 0

End Sub

Private Sub Destroy_Object()
    Set objBLBFunc = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Object
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


Private Sub mskDTFIN_GotFocus()

On Error GoTo Err_mskDTFIN_GotFocus
    
    Call objBLBFunc.SelecionaCampos(mskDTFIN.Name, Me)

    Exit Sub
    
Err_mskDTFIN_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : mskDTFIN_GotFocus()", Me.Name, "mskDTFIN_GotFocus()", strCAMARQERRO)

End Sub

Private Sub mskDTINI_GotFocus()

On Error GoTo Err_mskDTINI_GotFocus
    
    Call objBLBFunc.SelecionaCampos(mskDTINI.Name, Me)

    Exit Sub
    
Err_mskDTINI_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : mskDTINI_GotFocus()", Me.Name, "mskDTINI_GotFocus()", strCAMARQERRO)

End Sub

Private Sub optPeriodo_Click(Index As Integer)
    If Index = 0 Then fraPeriodo.Enabled = False
    If Index = 1 Then
        fraPeriodo.Enabled = True
        mskDTINI.SetFocus
    End If
End Sub


Private Sub optTodosClie_Click(Index As Integer)
    If Index = 0 Then Frame5.Enabled = False
    If Index = 1 Then Frame5.Enabled = True
End Sub

Private Sub txtCIDCLIE_GotFocus()

On Error GoTo Err_txtCIDCLIE_GotFocus

    Call objBLBFunc.SelecionaCampos(txtCIDCLIE.Name, Me)

    Exit Sub
    
Err_txtCIDCLIE_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIE_GotFocus()", Me.Name, "txtCIDCLIE_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCIDCLIE_KeyPress
    
    Call objBLBFunc.SoNumeroPonto(KeyAscii, txtCIDCLIE.Text)

    Exit Sub
    
Err_txtCIDCLIE_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIE_KeyPress()", Me.Name, "txtCIDCLIE_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

On Error GoTo Err_txtCIDCLIE_Validate

    Dim I As Integer
    
    If Len(Trim(txtCIDCLIE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCIDCLIE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If ConfereCliente(txtCIDCLIE.Text, txtCODVENDC2.Text) = False Then
       txtCIDCLIE.Text = ""
       lblDescCliente.Caption = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente)
    If Len(Trim(lblDescCliente.Caption)) = 0 Then
       txtCIDCLIE.Text = ""
       lblDescCliente.Caption = ""
       Cancel = True
       Exit Sub
    End If

    Exit Sub
    
Err_txtCIDCLIE_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIE_Validate()", Me.Name, "txtCIDCLIE_Validate()", strCAMARQERRO)

End Sub

Private Sub txtCIDCLIEF_GotFocus()

On Error GoTo Err_txtCIDCLIEF_GotFocus

    Call objBLBFunc.SelecionaCampos(txtCIDCLIEF.Name, Me)

    Exit Sub
    
Err_txtCIDCLIEF_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIEF_GotFocus()", Me.Name, "txtCIDCLIEF_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCIDCLIEF_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCIDCLIEF_KeyPress
    
    Call objBLBFunc.SoNumeroPonto(KeyAscii, txtCIDCLIEF.Text)

    Exit Sub
    
Err_txtCIDCLIEF_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIEF_KeyPress()", Me.Name, "txtCIDCLIEF_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCIDCLIEF_Validate(Cancel As Boolean)

On Error GoTo Err_txtCIDCLIEF_Validate

    Dim I As Integer
    
    If Len(Trim(txtCIDCLIEF.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCIDCLIEF.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCIDCLIEF.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If ConfereCliente(txtCIDCLIEF.Text, txtCODVEND.Text) = False Then
       txtCIDCLIEF.Text = ""
       lblDescClienteF.Caption = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIEF.Text, lblDescClienteF)
    If Len(Trim(lblDescClienteF.Caption)) = 0 Then
       txtCIDCLIEF.Text = ""
       lblDescClienteF.Caption = ""
       Cancel = True
       Exit Sub
    End If

    Exit Sub
    
Err_txtCIDCLIEF_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCIDCLIEF_Validate()", Me.Name, "txtCIDCLIEF_Validate()", strCAMARQERRO)

End Sub

Private Sub txtCODVEND_GotFocus()

On Error GoTo Err_txtCODVEND_GotFocus
    
    objBLBFunc.SelecionaCampos txtCODVEND.Name, Me

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

    Dim I As Integer
    
    If Len(Trim(txtCODVEND.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVEND.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVEND.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVEND.Text, lblDescVendedor)
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then
       txtCODVEND.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCODVEND_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVEND_Validate()", Me.Name, "txtCODVEND_Validate()", strCAMARQERRO)

End Sub

Private Sub txtCODVENDC1_GotFocus()

On Error GoTo Err_txtCODVENDC1_GotFocus
    
    Call objBLBFunc.SelecionaCampos(txtCODVENDC1.Name, Me)

    Exit Sub
    
Err_txtCODVENDC1_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDC1_GotFocus()", Me.Name, "txtCODVENDC1_GotFocus()", strCAMARQERRO)

End Sub


Private Sub txtCODVENDC1_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCODVENDC1_KeyPress

    Call objBLBFunc.SoNumeroPonto(KeyAscii, txtCODVENDC1.Text)

    Exit Sub
    
Err_txtCODVENDC1_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDC1_KeyPress()", Me.Name, "txtCODVENDC1_KeyPress()", strCAMARQERRO)

End Sub


Private Sub txtCODVENDC1_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODVENDC1_Validate

    Dim I As Integer
    
    If Len(Trim(txtCODVENDC1.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVENDC1.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVENDC1.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVENDC1.Text, lblDescVendedorC1)
    If Len(Trim(lblDescVendedorC1.Caption)) = 0 Then
       txtCODVENDC1.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCODVENDC1_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDC1_Validate()", Me.Name, "txtCODVENDC1_Validate()", strCAMARQERRO)

End Sub

Private Sub txtCODVENDC2_GotFocus()

On Error GoTo Err_txtCODVENDC2_GotFocus
    
    Call objBLBFunc.SelecionaCampos(txtCODVENDC2.Name, Me)

    Exit Sub
    
Err_txtCODVENDC2_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDC2_GotFocus()", Me.Name, "txtCODVENDC2_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCODVENDC2_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCODVENDC2_KeyPress

    Call objBLBFunc.SoNumeroPonto(KeyAscii, txtCODVENDC2.Text)

    Exit Sub
    
Err_txtCODVENDC2_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDC2_KeyPress()", Me.Name, "txtCODVENDC2_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCODVENDC2_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODVENDC2_Validate

    Dim I As Integer
    
    If Len(Trim(txtCODVENDC2.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVENDC2.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVENDC2.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVENDC2.Text, lblDescVendedorC2)
    If Len(Trim(lblDescVendedorC2.Caption)) = 0 Then
       txtCODVENDC2.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCODVENDC2_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDC2_Validate()", Me.Name, "txtCODVENDC2_Validate()", strCAMARQERRO)

End Sub

Private Sub txtCODVENDF_GotFocus()

On Error GoTo Err_txtCODVENDF_GotFocus
    
    objBLBFunc.SelecionaCampos txtCODVENDF.Name, Me

    Exit Sub
    
Err_txtCODVENDF_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDF_GotFocus()", Me.Name, "txtCODVENDF_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCODVENDF_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCODVENDF_KeyPress

    objBLBFunc.SoNumeroPonto KeyAscii, txtCODVENDF.Text

    Exit Sub
    
Err_txtCODVENDF_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDF_KeyPress()", Me.Name, "txtCODVENDF_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCODVENDF_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODVENDF_Validate

    Dim I As Integer
    
    If Len(Trim(txtCODVENDF.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVENDF.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVENDF.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVENDF.Text, lblDescVendedorF)
    If Len(Trim(lblDescVendedorF.Caption)) = 0 Then
       txtCODVENDF.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCODVENDF_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : txtCODVENDF_Validate()", Me.Name, "txtCODVENDF_Validate()", strCAMARQERRO)

End Sub

Private Sub LimpaCamposLabel()
    
On Error GoTo Err_LimpaCamposLabel
    
    lblDescVendedor.Caption = ""
    lblDescVendedorF.Caption = ""
    Label2.Caption = ""
    lblDescVendedorC1.Caption = ""
    lblDescVendedorC2.Caption = ""
    lblDescCliente.Caption = ""
    lblDescClienteF.Caption = ""
    Label4.Caption = ""
    
    Exit Sub
    
Err_LimpaCamposLabel:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : LimpaCamposLabel()", Me.Name, "LimpaCamposLabel()", strCAMARQERRO)
    
End Sub


Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
    If SSTab1.Tab = 0 Then
        If Len(Trim(txtCODVEND.Text)) > 0 And Len(Trim(txtCODVENDF.Text)) > 0 Then
           If CLng(txtCODVEND.Text) > CLng(txtCODVENDF.Text) Then
              MsgBox "Vendedor Inicial não pode ser maior que vendedor Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODVEND.SetFocus
              Exit Function
           End If
        End If
    ElseIf SSTab1.Tab = 0 Then
        If Len(Trim(txtCODVENDC1.Text)) > 0 And Len(Trim(txtCODVENDC2.Text)) > 0 Then
           If CLng(txtCODVENDC1.Text) > CLng(txtCODVENDC2.Text) Then
              MsgBox "Vendedor Inicial não pode ser maior que vendedor Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODVENDC1.SetFocus
              Exit Function
           End If
        End If
    End If
    
    If SSTab1.Tab = 0 And optPeriodo(1).Value = True Then
        If Not IsDate(mskDTINI.Text) Then
            MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
        If Not IsDate(mskDTFIN.Text) Then
            MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTFIN.SetFocus
            Exit Function
        End If
        
        If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
            MsgBox "Data Inicial não pode ser maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
    End If
    
    ConfereCampos = True

End Function



Private Function ConfereCliente(strCODIGO As String, strCODVEND As String) As Boolean

On Error GoTo Err_PegaDescTabelas

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
    sSql = sSql & "       CVEN.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CVEN.SGI_CODIGO = " & txtCODVEND.Text & vbCrLf
    sSql = sSql & "   And CVEN.SGI_CODCLI = " & strCODIGO & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL = CVEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO = CVEN.SGI_CODCLI"
    
    boolDadosInv = True
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC10.EOF() Then
       MsgBox "Este Cliente não pertence a este vendedor !!!", vbOKOnly + vbExclamation, "Aviso"
       boolDadosInv = False
    End If
    BREC10.Close
    
    If boolDadosInv = False Then Exit Function
    
    ConfereCliente = True
    
    Exit Function
    
Err_PegaDescTabelas:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Function


