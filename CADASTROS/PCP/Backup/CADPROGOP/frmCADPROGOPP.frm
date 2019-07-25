VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADPROGOPP 
   Caption         =   "Programação de Produção"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   17730
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDESCPROD 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      TabIndex        =   26
      Text            =   "txtDESCPROD"
      Top             =   1200
      Width           =   6735
   End
   Begin VB.Frame Frame3 
      Height          =   5175
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   17655
      Begin VSFlex8LCtl.VSFlexGrid grdMOVDIARIO 
         Height          =   4935
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   17415
         _cx             =   30718
         _cy             =   8705
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
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   17655
      Begin VB.Frame fraPeriodo 
         Caption         =   "[ Mês/Ano ]"
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
         Left            =   3960
         TabIndex        =   39
         Top             =   1920
         Width           =   6135
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   4320
            TabIndex        =   41
            Text            =   "cboAno"
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmCADPROGOPP.frx":0000
            Left            =   1920
            List            =   "frmCADPROGOPP.frx":0002
            TabIndex        =   40
            Text            =   "cboMes"
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label1 
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
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "[ Tipo de Relatório ]"
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
         TabIndex        =   36
         Top             =   1920
         Width           =   3735
         Begin VB.OptionButton optTIPREL 
            Caption         =   "Por Linha"
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
            TabIndex        =   38
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optTIPREL 
            Caption         =   "Por Data"
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
            Left            =   1920
            TabIndex        =   37
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtCODPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12600
         TabIndex        =   35
         Text            =   "txtCODPED"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCODOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9120
         TabIndex        =   34
         Text            =   "txtCODOP"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   15960
         Picture         =   "frmCADPROGOPP.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtDESCLIN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9120
         TabIndex        =   30
         Text            =   "txtDESCLIN"
         Top             =   1560
         Width           =   6735
      End
      Begin VB.TextBox txtCODLIN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   28
         Text            =   "txtCODLIN"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   15960
         Picture         =   "frmCADPROGOPP.frx":0106
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtCODPROD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   23
         Text            =   "txtCODPROD"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtNOMECLIE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9120
         TabIndex        =   21
         Text            =   "txtNOMECLIE"
         Top             =   840
         Width           =   6735
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   15960
         Picture         =   "frmCADPROGOPP.frx":0208
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtCODCLIE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   18
         Text            =   "txtCODCLIE"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data da Programação"
         DragMode        =   1  'Automatic
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
         Left            =   3960
         TabIndex        =   13
         Tag             =   "MOVP.SGI_DATAPROG"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   120
         Width           =   3135
         Begin MSMask.MaskEdBox mskDataI 
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataF 
            Height          =   255
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblCampo 
            Alignment       =   2  'Center
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
            Index           =   5
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Width           =   120
         End
      End
      Begin VB.Frame fraOrdem 
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
         ForeColor       =   &H00800000&
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3735
         Begin VB.ListBox lstOrdem 
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
            ForeColor       =   &H00800000&
            Height          =   1395
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código da Pedido"
         DragMode        =   1  'Automatic
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
         Left            =   10800
         TabIndex        =   33
         Tag             =   "MOVP.SGI_CODPED"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código da OP"
         DragMode        =   1  'Automatic
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
         Left            =   7200
         TabIndex        =   32
         Tag             =   "MOVP.SGI_CODOP"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição da Linha"
         DragMode        =   1  'Automatic
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
         Left            =   7200
         TabIndex        =   29
         Tag             =   "LINH.SGI_DESCRI"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   1560
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código da Linha"
         DragMode        =   1  'Automatic
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
         Left            =   3960
         TabIndex        =   27
         Tag             =   "MOVP.SGI_CODLIN"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   1560
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição do Produto"
         DragMode        =   1  'Automatic
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
         Left            =   7200
         TabIndex        =   24
         Tag             =   "PROD.SGI_DESCRICAO"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   1200
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código do Produto"
         DragMode        =   1  'Automatic
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
         Left            =   3960
         TabIndex        =   22
         Tag             =   "PROD.SGI_CODIGO"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Cliente"
         DragMode        =   1  'Automatic
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
         Left            =   7200
         TabIndex        =   20
         Tag             =   "CLIE.SGI_RAZAOSOC"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código do Cliente"
         DragMode        =   1  'Automatic
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
         Left            =   3960
         TabIndex        =   17
         Tag             =   "PROD.SGI_CODCLIE"
         ToolTipText     =   "Click e Arraste para o Campo Ordem"
         Top             =   840
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   17655
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         Picture         =   "frmCADPROGOPP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar"
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
         Picture         =   "frmCADPROGOPP.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdInclui 
         Caption         =   "&Inclui"
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
         Picture         =   "frmCADPROGOPP.frx":093E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Inclui um novo registro"
         Top             =   120
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
         Left            =   1800
         Picture         =   "frmCADPROGOPP.frx":0E70
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Altera Registro"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdExclui 
         Caption         =   "&Exclui"
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
         Left            =   2640
         Picture         =   "frmCADPROGOPP.frx":0F72
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Registro"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCanFiltro 
         Caption         =   "Desfas"
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
         Left            =   16080
         Picture         =   "frmCADPROGOPP.frx":1074
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Desfas Ultima Pesqusa"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdOrden 
         Caption         =   "Ordem"
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
         Left            =   16800
         Picture         =   "frmCADPROGOPP.frx":15A6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ordena os Registros"
         Top             =   120
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   13440
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmCADPROGOPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Public strUsuario   As String
Public lngCodUsuaro As Long
Public intFILIALPED As Integer

Dim lngCodVendedor  As Long
Dim objFuncoes      As Object
Dim objCADMOVPCPP   As Object
Dim objRel          As Object
Dim objPESQPADRAO   As Object
Dim iCodigo         As Long
Dim strModulo       As String
Dim strNomTab       As String

Const conCOL_Mov_Codigo                          As Integer = 0
Const conCOL_Mov_DatMov                          As Integer = 1
Const conCOL_Mov_CodOP                           As Integer = 2
Const conCOL_Mov_CodPED                          As Integer = 3
Const conCOL_Mov_CodLin                          As Integer = 4
Const conCOL_Mov_DescLin                         As Integer = 5
Const conCOL_Mov_CodProd                         As Integer = 6
Const conCOL_Mov_DescProd                        As Integer = 7
Const conCOL_Mov_QtdePed                         As Integer = 8
Const conCOL_Mov_CodClie                         As Integer = 9
Const conCOL_Mov_DescClie                        As Integer = 10
Const conCOL_Mov_FormatString                    As String = "=Cód.Programação|Data Programação|Cod.OP|Cod.Ped|Linha|Descrição da Linha|Cód. Produto|Descrição do Produto|Qtde.Ped|Cód.Cliente|Descrição do Cliente"
Const conColumnsIn_Mov                           As Integer = 11


Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
    Call ConfGrid
End Sub

Private Sub cmdExclui_Click()

    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
      
    Dim iResp As Integer
      
    iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
      
    If iResp <> 6 Then Exit Sub
      
    If objCADMOVPCPP.Carrega_Campos(strNomTab) = False Then Exit Sub
    If objCADMOVPCPP.GRAVA("E", strNomTab) = False Then Exit Sub
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    Call AbilitaCampos
    Call ConfGrid

End Sub

Private Sub cmdImpressao_Click()
    If ConsisteCampos = False Then Exit Sub
    Call Imprime
End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    If ConsisteCampos = False Then Exit Sub
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PROD.* " & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO  PROD" & vbCrLf
    If Len(Trim(txtNOMECLIE.Text)) > 0 Then
        sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
    End If
    If Len(Trim(txtDESCLIN.Text)) > 0 Then
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO  LINH" & vbCrLf
    End If
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL  = " & FILIAL & vbCrLf
    If Len(Trim(txtCODCLIE.Text)) > 0 Then
        sSql = sSql & "   And PROD.SGI_CODCLIE = " & Trim(txtCODCLIE.Text) & vbCrLf
    End If
    If Len(Trim(txtCODLIN.Text)) > 0 Then
        sSql = sSql & "   And PROD.SGI_CODLINPROD Like '%" & Trim(txtCODLIN.Text) & "%'" & vbCrLf
    End If
    If Len(Trim(txtCODPROD.Text)) > 0 Then
        sSql = sSql & "   And PROD.SGI_CODIGO Like '%" & Trim(txtCODPROD.Text) & "%'" & vbCrLf
    End If
    
    If Len(Trim(txtNOMECLIE.Text)) > 0 Then
        sSql = sSql & "   And CLIE.SGI_FILIAL  = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_RAZAOSOC Like '%" & txtNOMECLIE.Text & "%'" & vbCrLf
        sSql = sSql & "   And PROD.SGI_CODCLIE = CLIE.SGI_CODIGO" & vbCrLf
    End If
    If Len(Trim(txtDESCLIN.Text)) > 0 Then
        sSql = sSql & "   And LINH.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And LINH.SGI_DESCRI Like '%" & txtDESCLIN.Text & "%'" & vbCrLf
        sSql = sSql & "   And PROD.SGI_CODLINPROD = LINH.SGI_CODLIN" & vbCrLf
    End If
    
    arrTABELA(1) = sSql
    
    
    arrCAMPOS(1, 1) = "SGI_IDPRODUTO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código Interno"
    arrCAMPOS(1, 4) = "1600"
    arrCAMPOS(1, 5) = "PROD.SGI_IDPRODUTO"
    
    arrCAMPOS(2, 1) = "SGI_CODIGO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Código do Produto"
    arrCAMPOS(2, 4) = "1600"
    arrCAMPOS(2, 5) = "PROD.SGI_CODIGO"
    
    arrCAMPOS(3, 1) = "SGI_DESCRICAO"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Descrição do Produto"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "PROD.SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) = 0 Then Exit Sub
    
    txtCODPROD.Text = PegaCodigoProduto(CLng(varRETORNO))
    
    Call PegaDescTabelas("SGI_IDPRODUTO", "SGI_DESCRICAO", "SGI_CADPRODUTO", varRETORNO, txtDESCPROD, "Command1_Click()", False)
    If Len(Trim(txtDESCPROD.Text)) = 0 Then txtCODPROD.Text = ""
    
    txtCODPROD.SetFocus

    Exit Sub

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
    If Len(Trim(txtDESCLIN.Text)) > 0 Then
        sSql = sSql & "   And SGI_DESCRI Like '%" & Trim(txtDESCLIN.Text) & "%'"
    End If
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODLIN"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código da Linha"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODLIN"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição da Linha"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Linha de Produto", "CADLINHAPROD.clsCADLINHAPROD")
    
    If Len(Trim(varRETORNO)) = 0 Then Exit Sub
    
    txtCODLIN.Text = varRETORNO
       
    Call PegaDescTabelas("SGI_CODLIN", "SGI_DESCRI", "SGI_CADLINHAPRODUTO", "'" & varRETORNO & "'", txtDESCLIN, "Command2_Click()", False)
    If Len(Trim(txtDESCLIN.Text)) = 0 Then txtDESCLIN.Text = ""
    
    txtCODLIN.SetFocus
       

End Sub

Private Sub Command9_Click()


    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CLIE.* " & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE  CLIE" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CLIE.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If Len(Trim(txtNOMECLIE.Text)) > 0 Then
        sSql = sSql & "   And CLIE.SGI_RAZAOSOC Like '%" & Trim(txtNOMECLIE.Text) & "%'"
    End If
    
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
    arrCAMPOS(3, 4) = "5000"
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLIE.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, txtNOMECLIE, "Command8_Click()", False)
    If Len(Trim(txtNOMECLIE.Text)) = 0 Then txtCODCLIE.Text = ""
    
    txtCODCLIE.SetFocus

    Exit Sub
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
   Set objCADMOVPCPP = CreateObject("CADPROGOP.clsCADPROGOP")
   Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    
   objCADMOVPCPP.FILIAL = FILIAL
   objFuncoes.LimpaCampos Me
    
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
    
   strNomTab = ""
   If intFILIALPED = 0 Then strModulo = "NOVALATA"
   If intFILIALPED = 1 Then
        strModulo = "STEEL ROL"
        strNomTab = "_STEEL"
   End If
   
   Call AbilitaCampos
   Call ConfGrid
    
   strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

   Me.Caption = Me.Caption & " / " & strModulo
   
   optTIPREL(0).Value = True

    Call objFuncoes.Preenche_Mes(cboMes)
    cboMes.ListIndex = (Month(Date) - 1)
    
    Call objFuncoes.Preenche_Ano(cboAno)
    cboAno.ListIndex = 0


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub grdMOVDIARIO_Click()
    With grdMOVDIARIO
        If .Row = 0 Then Exit Sub
        If (.Rows - 1) > 0 Then objCADMOVPCPP.CODIGO = CInt(.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
    End With
End Sub

Private Sub grdMOVDIARIO_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    With grdMOVDIARIO
        If .Row = 0 Then Exit Sub
        If (.Rows - 1) > 0 Then Call Operacao("C")
    End With
End Sub

Private Sub grdMOVDIARIO_RowColChange()
    With grdMOVDIARIO
        If .Row = 0 Then Exit Sub
        If (.Rows - 1) > 0 Then objCADMOVPCPP.CODIGO = CInt(.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
    End With
End Sub



Private Sub lstOrdem_DragDrop(Source As Control, x As Single, Y As Single)
    Dim i As Integer
    For i = 0 To (lstOrdem.ListCount - 1)
        If lstOrdem.ItemData(i) = Source.TabIndex Then Exit Sub
    Next i
    
    lstOrdem.AddItem Source
    lstOrdem.ItemData(lstOrdem.NewIndex) = Source.TabIndex
End Sub

Private Sub lstOrdem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If lstOrdem.ListCount = 0 Then Exit Sub
        If lstOrdem.ListIndex = -1 Then Exit Sub
        lstOrdem.RemoveItem lstOrdem.ListIndex
    End If
End Sub

Private Sub mskDataF_GotFocus()
    objFuncoes.SelecionaCampos mskDataF.Name, Me
End Sub

Private Sub mskDataI_GotFocus()
    objFuncoes.SelecionaCampos mskDataI.Name, Me
End Sub


Private Sub Operacao(strOperacao As String)
  
    With grdMOVDIARIO
        If strOperacao <> "I" Then
            If (.Rows - 1) = 0 Or .Row = 0 Then
                MsgBox "ATENÇÂO" & vbCrLf & _
                       "Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
                       Exit Sub
            End If
            If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
        End If
    End With
    
    frmCADPROGOP.cCaminho = cCaminho
    frmCADPROGOP.Linha = Linha
    frmCADPROGOP.iCodigo = iCodigo
    frmCADPROGOP.cTipOper = strOperacao
    frmCADPROGOP.FILIAL = FILIAL
    frmCADPROGOP.strAcesso = strAcesso
    frmCADPROGOP.strMODPAI = Me.Name
    frmCADPROGOP.strUsuario = strUsuario
    frmCADPROGOP.lngCodVendedor = lngCodVendedor
    frmCADPROGOP.lngCodUsuario = lngCodUsuaro
    frmCADPROGOP.intFILIALPED = intFILIALPED
    frmCADPROGOP.Show vbModal
    
    ''Call Atualiza_Grid
    Call AbilitaCampos
    Call ConfGrid

End Sub


Private Sub ConfGrid()

    With grdMOVDIARIO
    
       .Cols = conColumnsIn_Mov
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Mov_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_Mov_Codigo) = ""
       .ColDataType(conCOL_Mov_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_DatMov) = ""
       .ColDataType(conCOL_Mov_DatMov) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_Mov_CodOP) = ""
       .ColDataType(conCOL_Mov_CodOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_CodPED) = ""
       .ColDataType(conCOL_Mov_CodPED) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_CodLin) = ""
       .ColDataType(conCOL_Mov_CodLin) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Mov_DescLin) = ""
       .ColDataType(conCOL_Mov_DescLin) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Mov_CodProd) = ""
       .ColDataType(conCOL_Mov_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Mov_DescProd) = ""
       .ColDataType(conCOL_Mov_DescProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Mov_QtdePed) = ""
       .ColDataType(conCOL_Mov_QtdePed) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_CodClie) = ""
       .ColDataType(conCOL_Mov_CodClie) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_DescClie) = ""
       .ColDataType(conCOL_Mov_DescClie) = flexDTString
       
       .ColWidth(conCOL_Mov_Codigo) = 0
       .ColWidth(conCOL_Mov_DatMov) = 1500
       .ColWidth(conCOL_Mov_CodOP) = 1200
       .ColWidth(conCOL_Mov_CodPED) = 1200
       .ColWidth(conCOL_Mov_CodLin) = 1200
       .ColWidth(conCOL_Mov_DescLin) = 2500
       .ColWidth(conCOL_Mov_CodProd) = 1200
       .ColWidth(conCOL_Mov_DescProd) = 5000
       .ColWidth(conCOL_Mov_QtdePed) = 1200
       .ColWidth(conCOL_Mov_CodClie) = 1200
       .ColWidth(conCOL_Mov_DescClie) = 5000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADMOVPCPP.AtivoDesativo(strNomTab)
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdImpressao.Enabled = boolAtivoDesativo
    cmdCanFiltro.Enabled = boolAtivoDesativo
    cmdOrden.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo
    Frame3.Enabled = boolAtivoDesativo

End Sub

Private Sub Ordem()

    Call ConfGrid
  
    
    ''If cboMes.ListIndex = -1 Then Exit Sub
    ''If cboAno.ListIndex = -1 Then Exit Sub
    
    ''Dim strDTINICIAL        As String
    ''Dim strDTFINAL          As String

    ''strDTINICIAL = "'" & Format(CDate("01/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "MM/DD/YYYY") & "'"
    ''If cboMes.ItemData(cboMes.ListIndex) = 12 Then
    ''    strDTFINAL = "'" & Format(CDate("31/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "MM/DD/YYYY") & "'"
    ''Else
    ''    strDTFINAL = "'" & Format((CDate("01/" & (cboMes.ItemData(cboMes.ListIndex) + 1) & "/" & cboAno.ItemData(cboAno.ListIndex)) - 1), "MM/DD/YYYY") & "'"
    ''End If
    
    sSql = ""
    
    sSql = " Select " & vbCrLf
    sSql = sSql & "        MOVP.SGI_CODIGO " & vbCrLf
    sSql = sSql & "       ,MOVP.SGI_DATAPROG" & vbCrLf
    sSql = sSql & "       ,MOVP.SGI_CODOP" & vbCrLf
    sSql = sSql & "       ,MOVP.SGI_CODPED" & vbCrLf
    sSql = sSql & "       ,MOVP.SGI_CODLIN" & vbCrLf
    sSql = sSql & "       ,MOVP.SGI_QTDEPROD" & vbCrLf
    
    sSql = sSql & "       ,LINH.SGI_DESCRI As SGI_DESCLIN" & vbCrLf
    
    sSql = sSql & "       ,PROD.SGI_CODIGO    As SGI_CODPROD" & vbCrLf
    sSql = sSql & "       ,PROD.SGI_DESCRICAO As SGI_DESCPROD" & vbCrLf
    sSql = sSql & "       ,PROD.SGI_CODCLIE" & vbCrLf
    
    sSql = sSql & "       ,CLIE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "   from " & vbCrLf
    sSql = sSql & "        SGI_CADMOVPCP" & strNomTab & " MOVP" & vbCrLf
    sSql = sSql & "       ,SGI_CADLINHAPRODUTO  LINH" & vbCrLf
    sSql = sSql & "       ,SGI_CADPRODUTO       PROD" & vbCrLf
    sSql = sSql & "       ,SGI_CADCLIENTE       CLIE" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MOVP.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And LINH.SGI_FILIAL    = MOVP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINH.SGI_CODLIN    = MOVP.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL    = MOVP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = MOVP.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And CLIE.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO    = PROD.SGI_CODCLIE" & vbCrLf
    
    If Len(Trim(Replace(Replace(mskDataI.Text, "/", ""), "_", ""))) > 0 And _
       Len(Trim(Replace(Replace(mskDataF.Text, "/", ""), "_", ""))) > 0 And _
       Len(Trim(Replace(Replace(mskDataI.Text, "/", ""), "_", ""))) = 8 And _
       Len(Trim(Replace(Replace(mskDataF.Text, "/", ""), "_", ""))) = 8 Then
        sSql = sSql & "   And MOVP.SGI_DATAPROG Between '" & Format(CDate(mskDataI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDataF.Text), "MM/DD/YYYY") & "'" & vbCrLf
    End If
    If cboMes.ListIndex > -1 And _
       cboAno.ListIndex > -1 Then
       sSql = sSql & "   And Month(MOVP.SGI_DATAPROG) = " & cboMes.ItemData(cboMes.ListIndex) & vbCrLf
       sSql = sSql & "   And Year(MOVP.SGI_DATAPROG)  = " & cboAno.ItemData(cboAno.ListIndex) & vbCrLf
    End If
    If Len(Trim(txtCODOP.Text)) > 0 Then
        sSql = sSql & "   And MOVP.SGI_CODOP   = " & Trim(txtCODOP.Text) & vbCrLf
    End If
    If Len(Trim(txtCODPED.Text)) > 0 Then
        sSql = sSql & "   And MOVP.SGI_CODPED  = " & Trim(txtCODPED.Text) & vbCrLf
    End If
    If Len(Trim(txtCODLIN.Text)) > 0 Then
        sSql = sSql & "   And MOVP.SGI_CODLIN  = " & Trim(txtCODLIN.Text) & vbCrLf
    End If
    If Len(Trim(txtCODCLIE.Text)) > 0 Then
        sSql = sSql & "   And PROD.SGI_CODCLIE = " & Trim(txtCODCLIE.Text) & vbCrLf
    End If
    If Len(Trim(txtCODPROD.Text)) > 0 Then
        sSql = sSql & "   And PROD.SGI_CODIGO Like '%" & Trim(txtCODPROD.Text) & "%'" & vbCrLf
    End If
    If Len(Trim(txtNOMECLIE.Text)) > 0 Then
        sSql = sSql & "   And CLIE.SGI_RAZAOSOC Like '%" & Trim(txtNOMECLIE.Text) & "%'" & vbCrLf
    End If
    If Len(Trim(txtDESCPROD.Text)) > 0 Then
        sSql = sSql & "   And PROD.SGI_DESCRICAO Like '%" & Trim(txtDESCPROD.Text) & "%'" & vbCrLf
    End If
    If Len(Trim(txtDESCLIN.Text)) > 0 Then
        sSql = sSql & "   And LINH.SGI_DESCRI Like '%" & Trim(txtDESCLIN.Text) & "%'" & vbCrLf
    End If
    
    If lstOrdem.ListCount > 0 Then sSql = sSql & ConfCamposOrdem
    
  
    BREC.Open sSql, adoBanco_Dados
    If Not BREC.EOF() Then
        With grdMOVDIARIO
            Do While Not BREC.EOF
               .AddItem BREC!SGI_CODIGO & vbTab & _
                        Format(BREC!SGI_DATAPROG, "DD/MM/YYYY") & vbTab & _
                        BREC!SGI_CODOP & vbTab & _
                        BREC!SGI_CODPED & vbTab & _
                        Trim(BREC!SGI_CODLIN) & vbTab & _
                        Trim(BREC!SGI_DESCLIN) & vbTab & _
                        Trim(BREC!SGI_CODPROD) & vbTab & _
                        Trim(BREC!SGI_DESCPROD) & vbTab & _
                        BREC!SGI_QTDEPROD & vbTab & _
                        BREC!SGI_CODCLIE & vbTab & _
                        Trim(BREC!SGI_RAZAOSOC)

               BREC.MoveNext
            Loop
        End With
    Else
        MsgBox "Não há dados para consultar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

End Sub
Private Sub Destroy_Objeto()
    Set objFuncoes = Nothing
    Set objCADMOVPCPP = Nothing
    Set objRel = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub txtCODCLIE_GotFocus()
    objFuncoes.SelecionaCampos txtCODCLIE.Name, Me
End Sub

Private Sub txtCODCLIE_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODCLIE.Text
End Sub

Private Sub txtCODCLIE_Validate(Cancel As Boolean)


    Dim i As Integer
    
    If Len(Trim(txtCODCLIE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLIE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCODCLIE.Text = Trim(Replace(Replace(txtCODCLIE.Text, ",", ""), ".", ""))
    

End Sub

Private Sub txtCODLIN_GotFocus()
    objFuncoes.SelecionaCampos txtCODLIN.Name, Me
End Sub

Private Sub txtCODLIN_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCODLIN_Validate(Cancel As Boolean)

    If Len(Trim(txtCODLIN.Text)) = 0 Then Exit Sub
    
    txtCODLIN.Text = Trim(Replace(Replace(txtCODLIN.Text, ",", ""), ".", ""))
    
End Sub

Private Sub txtCODOP_GotFocus()
    objFuncoes.SelecionaCampos txtCODOP.Name, Me
End Sub

Private Sub txtCODOP_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODOP.Text
End Sub

Private Sub txtCODPED_GotFocus()
    objFuncoes.SelecionaCampos txtCODPED.Name, Me
End Sub

Private Sub txtCODPED_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODPED.Text
End Sub

Private Sub txtCODPROD_GotFocus()
    objFuncoes.SelecionaCampos txtCODPROD.Name, Me
End Sub

Private Sub txtCODPROD_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODPROD.Text
End Sub

Private Sub txtCODPROD_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCODPROD.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODPROD.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODPROD.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCODPROD.Text = Trim(Replace(txtCODPROD.Text, ",", ""))

End Sub

Private Sub txtDESCLIN_GotFocus()
    objFuncoes.SelecionaCampos txtDESCLIN.Name, Me
End Sub

Private Sub txtDESCLIN_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtDESCPROD_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtNOMECLIE_GotFocus()
    objFuncoes.SelecionaCampos txtNOMECLIE.Name, Me
End Sub

Private Sub txtDESCPROD_GotFocus()
    objFuncoes.SelecionaCampos txtDESCPROD.Name, Me
End Sub

Private Sub txtNOMECLIE_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, txtGeral As TextBox, strFUNCAOPAI As String, boolLIKE As Boolean)

    txtGeral.Text = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    
    If boolLIKE = False Then
        sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(Replace(strCODIGO, ",", ""))
    ElseIf boolLIKE = True Then
        sSql = sSql & "       " & Trim(StrCampoPesq) & " Like " & Trim(Replace(strCODIGO, ",", ""))
    End If
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       txtGeral.Text = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
    Exit Sub
    

End Sub


Private Function PegaCodigoProduto(lngIDPRODUTO As Long) As String

    PegaCodigoProduto = ""
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PROD.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PROD" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO   = " & lngIDPRODUTO
    
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC7.EOF() Then PegaCodigoProduto = BREC7!SGI_CODIGO
    BREC7.Close
    
End Function


Private Function ConsisteCampos() As Boolean

    Dim strDATAI As String
    Dim strDATAF As String
    
    ConsisteCampos = True
    
    strDATAI = Replace(Replace(mskDataI.Text, "/", ""), "_", "")
    strDATAF = Replace(Replace(mskDataF.Text, "/", ""), "_", "")
    
    If (Len(Trim(strDATAI)) = 0 Or Len(Trim(strDATAF)) = 0) Then Exit Function
    If (Len(Trim(strDATAI)) < 8 Or Len(Trim(strDATAF)) < 8) Then Exit Function
    
    If (Len(Trim(strDATAI)) = 0 And Len(Trim(strDATAF)) > 0) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data inicial não pode ser vázia !!!"
        ConsisteCampos = False
        Exit Function
    End If
    If (Len(Trim(strDATAI)) > 0 And Len(Trim(strDATAF)) = 0) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data final não pode ser vázia !!!"
        ConsisteCampos = False
        Exit Function
    End If
    
    If Not IsDate(mskDataI.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data inicial inválida !!!"
        mskDataI.SetFocus
        ConsisteCampos = False
        Exit Function
    End If
    If Not IsDate(mskDataF.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data final inválida !!!"
        mskDataF.SetFocus
        ConsisteCampos = False
        Exit Function
    End If
    
    
    If CDate(mskDataI.Text) > CDate(mskDataF.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "A data inicial não pode ser maior que data final !!!"
        mskDataI.SetFocus
        ConsisteCampos = False
        Exit Function
    End If
    
    
    
    
End Function


Private Function ConfCamposOrdem() As String

    ConfCamposOrdem = ""
    
    If lstOrdem.ListCount = 0 Then Exit Function
    
    Dim strCAMPOOrd         As String
    Dim i                   As Integer
    Dim intTABINDEX         As Integer
    Dim vControl
    
    ConfCamposOrdem = "Order By" & vbCrLf
    
    For i = 0 To (lstOrdem.ListCount - 1)
        For Each vControl In Me.Controls
            If TypeOf vControl Is TextBox Then
            ElseIf TypeOf vControl Is OptionButton Then
            ElseIf TypeOf vControl Is ComboBox Then
            ElseIf TypeOf vControl Is ListBox Then
            ElseIf TypeOf vControl Is Label Then
                If vControl.TabIndex = lstOrdem.ItemData(i) Then
                    ConfCamposOrdem = ConfCamposOrdem & Trim(vControl.Tag)
                    If i < (lstOrdem.ListCount - 1) Then ConfCamposOrdem = ConfCamposOrdem & ","
                End If
            ElseIf TypeOf vControl Is Frame Then
                If vControl.TabIndex = lstOrdem.ItemData(i) Then
                    ConfCamposOrdem = ConfCamposOrdem & Trim(vControl.Tag)
                    If i < (lstOrdem.ListCount - 1) Then ConfCamposOrdem = ConfCamposOrdem & ","
                End If
            End If
            
        Next
    Next i

End Function


Private Sub Imprime()
    
    Dim strNomRel       As String
    Dim boolTEMDADOS    As Boolean
    Dim strCABEC1       As String
    Dim strCABEC2       As String
    Dim strNOMTABELA    As String
    Dim strNOMEMPRESA   As String
    
    strNOMTABELA = strNomTab
    strNOMEMPRESA = strModulo
    
    boolTEMDADOS = False
    
    sSql = ""
    
    sSql = sSql & "Select" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAPROG" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAENTR" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODOP" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_QTDEPROD" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_QTDEAPONTADA" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_STATUSAPONT" & vbCrLf
    
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_CODIGO" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_VernTampa" & vbCrLf
    
    sSql = sSql & "     , SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "     , SGI_ORDEMPROD" & strNOMTABELA & ".SGI_FECHTPFU" & vbCrLf
    
    sSql = sSql & "     , SGI_CADFECHAM.SGI_DESCRI"
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADFECHAM SGI_CADFECHAM" & vbCrLf
    sSql = sSql & "     , SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & " SGI_CADMOVPCP" & strNOMTABELA & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_ORDEMPROD" & strNOMTABELA & " SGI_ORDEMPROD" & strNOMTABELA & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & ".SGI_FILIAL    = " & FILIAL & vbCrLf
    
    If objCADMOVPCPP.CODIGO > 0 Then
        sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODIGO    = " & objCADMOVPCPP.CODIGO & vbCrLf
    End If
    
    If Len(Trim(Replace(Replace(mskDataI.Text, "/", ""), "_", ""))) > 0 And _
       Len(Trim(Replace(Replace(mskDataF.Text, "/", ""), "_", ""))) > 0 And _
       Len(Trim(Replace(Replace(mskDataI.Text, "/", ""), "_", ""))) = 8 And _
       Len(Trim(Replace(Replace(mskDataF.Text, "/", ""), "_", ""))) = 8 Then
    
       sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAPROG  Between '" & Format(CDate(mskDataI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDataF.Text), "MM/DD/YYYY") & "'" & vbCrLf
    End If
    
    If Len(Trim(txtCODLIN.Text)) > 0 Then
        sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODLIN    = " & txtCODLIN.Text & vbCrLf
    End If
    
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_FILIAL    = SGI_ORDEMPROD" & strNOMTABELA & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODOP     = SGI_ORDEMPROD" & strNOMTABELA & ".SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL                       = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD                   = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD" & strNOMTABELA & ".SGI_FILIAL    = SGI_CADFECHAM.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strNOMTABELA & ".SGI_FECHTPFU  = SGI_CADFECHAM.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then boolTEMDADOS = True
    BREC.Close
    
    If boolTEMDADOS = False Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Não há dados para Imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    strCABEC1 = "Programação de Montagem " & strNOMEMPRESA
    
    strNomRel = ""
    strNomRel = "RELPCPPROGMONT" & strNOMTABELA
    
    If optTIPREL(0).Value = True Then strNomRel = strNomRel & "_LIN.rpt"
    If optTIPREL(1).Value = True Then strNomRel = strNomRel & "_DT.rpt"

    If Len(Trim(strNomRel)) > 0 Then
        Call objRel.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If


End Sub


