VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCADPROGLITO 
   Caption         =   "Cadastro de Programação de Litografia"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   18075
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar pbDADOS 
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   8640
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Op's já Programadas ]"
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
      Height          =   2415
      Left            =   0
      TabIndex        =   29
      Top             =   6120
      Width           =   18015
      Begin VSFlex8LCtl.VSFlexGrid grdOPSPROGRAMADAS 
         Height          =   2055
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   17775
         _cx             =   31353
         _cy             =   3625
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
   Begin VB.Frame Frame3 
      Caption         =   "[ OP's há Programar ]"
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
      Height          =   3015
      Left            =   0
      TabIndex        =   28
      Top             =   3120
      Width           =   18015
      Begin VSFlex8LCtl.VSFlexGrid grdOPS 
         Height          =   2655
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   17775
         _cx             =   31353
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
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   18015
      Begin VB.TextBox txtSTATUS 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   8040
         TabIndex        =   39
         Text            =   "txtSTATUS"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdCarrega 
         Caption         =   "&Carrega Dados"
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
         Left            =   10320
         Picture         =   "frmCADPROGLITO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3480
         Picture         =   "frmCADPROGLITO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtCODCAPAC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "txtCODCAPA"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   3480
         Picture         =   "frmCADPROGLITO.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtCODCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtCODCLIE"
         Top             =   1320
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtUSUARIOALT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   15360
         TabIndex        =   24
         Text            =   "txtUSUARIOALT"
         Top             =   1200
         Width           =   2535
      End
      Begin MSMask.MaskEdBox mskDTCRIACAOUSU 
         Height          =   285
         Left            =   15360
         TabIndex        =   21
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtUSUARIOCRIA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   15360
         TabIndex        =   19
         Text            =   "txtUSUARIOCRIA"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtCADMAQ 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCADMAQ"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdFornec 
         Height          =   315
         Left            =   3480
         Picture         =   "frmCADPROGLITO.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDTMOV 
         Height          =   285
         Left            =   5880
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCODIGO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "txtCODIGO"
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskDTALTERACAO 
         Height          =   285
         Left            =   15360
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   12
         Left            =   7320
         TabIndex        =   38
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblDESCAPAC 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDESCAPAC"
         Height          =   285
         Left            =   3840
         TabIndex        =   37
         Top             =   1680
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Capacidade"
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
         Index           =   11
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   -1920
         TabIndex        =   34
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblDESCLIE 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDESCLIE"
         Height          =   285
         Left            =   3840
         TabIndex        =   33
         Top             =   1320
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "á"
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
         Index           =   8
         Left            =   3600
         TabIndex        =   27
         Top             =   960
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data de Alteração"
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
         Index           =   6
         Left            =   13560
         TabIndex        =   23
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alterado Por"
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
         Index           =   5
         Left            =   13560
         TabIndex        =   22
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data de Criação"
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
         Index           =   4
         Left            =   13560
         TabIndex        =   20
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário de Criação"
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
         Index           =   3
         Left            =   13560
         TabIndex        =   18
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label lblDESCMAQ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDESCMAQ"
         Height          =   285
         Left            =   3840
         TabIndex        =   17
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Máquina"
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
         TabIndex        =   15
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data da Programação"
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
         Left            =   3840
         TabIndex        =   14
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código da Programação"
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
         TabIndex        =   13
         Top             =   240
         Width           =   2025
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   18015
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Volta <ESC>"
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
         Picture         =   "frmCADPROGLITO.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva <F2>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         Picture         =   "frmCADPROGLITO.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1695
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
         Left            =   1560
         Picture         =   "frmCADPROGLITO.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADPROGLITO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public iParcela         As Integer
Public FILIAL           As Integer
Public strACESSO        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCODUSUARIO    As Long
Public intFILIALPED     As Integer
Public strFILIAL        As String

Dim lngCodLog           As Long
Dim strValor            As String
Dim strCAPTION          As String
Dim strNOMFILIAL        As String

Dim objBLBFunc          As New clsFuncoes
Dim objCADPROGLITO      As New clsCADPROGLITO
Dim objPESQPADRAO       As Object

Const conCOL_Prod_OP                              As Integer = 0
Const conCOL_Prod_Pedido                          As Integer = 1
Const conCOL_Prod_DTENTREGA                       As Integer = 2
Const conCOL_Prod_QTDEOP                          As Integer = 3
Const conCOL_Prod_QTDFAT                          As Integer = 4
Const conCOL_Prod_SALDO                           As Integer = 5
Const conCOL_Prod_ID                              As Integer = 6
Const conCOL_Prod_ROTULO                          As Integer = 7
Const conCOL_Prod_DESCROT                         As Integer = 8
Const conCOL_Prod_CODCAPAC                        As Integer = 9
Const conCOL_Prod_CAPACIDADE                      As Integer = 10
Const conCOL_Prod_CODCLIE                         As Integer = 11
Const conCOL_Prod_DESCCLIE                        As Integer = 12
Const conCOL_Prod_FotNovo                         As Integer = 13
Const conCOL_Prod_NeckIN                          As Integer = 14
Const conCOL_Prod_Fechamento                      As Integer = 15
Const conCOL_Prod_CodVerniz                       As Integer = 16
Const conCOL_Prod_Verniz                          As Integer = 17
Const conCOL_Prod_CodEsmalte                      As Integer = 18
Const conCOL_Prod_Esmalte                         As Integer = 19
Const conCOL_Prod_QtdeFls                         As Integer = 20
Const conCOL_Prod_QtdeLatas                       As Integer = 21
Const conCOL_Prod_DtPEdido                        As Integer = 22
Const conCOL_Prod_FormatString                    As String = "=OP|Pedido|Dt.Entrega|Qtde.OP|Qtde.Fat|Saldo.OP|IDPROD|Rótulo|Desc.Rótulo|CODCPAC|Capacidade|CODCLIE|Cliente|Fot.Novo S/N|Neck-IN S/N|Fechamento|CodVerniz|Verniz|CodEsmalte|Esmalte|Qte.Fls|Qte.Latas|Dt.Pedido"
Const conColumnsIn_Prod                           As Integer = 23

Const conCOL_ProdProg_OP                          As Integer = 0
Const conCOL_ProdProg_DTENTREGA                   As Integer = 1
Const conCOL_ProdProg_QTDEOP                      As Integer = 2
Const conCOL_ProdProg_QTDEFAT                     As Integer = 3
Const conCOL_ProdProg_SALDO                       As Integer = 4
Const conCOL_ProdProg_ID                          As Integer = 5
Const conCOL_ProdProg_ROTULO                      As Integer = 6
Const conCOL_ProdProg_DESCROT                     As Integer = 7
Const conCOL_ProdProg_CODCAPAC                    As Integer = 8
Const conCOL_ProdProg_CAPACIDADE                  As Integer = 9
Const conCOL_ProdProg_CODCLIE                     As Integer = 10
Const conCOL_ProdProg_DESCCLIE                    As Integer = 11
Const conCOL_ProdProg_FormatString                As String = "=OP|Dt.Entrega|Qtde.OP|Qtde.Fat|Saldo.OP|IDPROD|Rótulo|Desc.Rótulo|CODCPAC|Capacidade|CODCLIE|Cliente"
Const conColumnsIn_ProdProg                       As Integer = 12

Private Sub cmdCarrega_Click()
    If ConfereCampos = False Then Exit Sub
    Call CarregaDadosOP("'" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "'", "'" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'")
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADPROGLITO = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Load()

''    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
''    Set objCADPROGLITO = CreateObject("CADPROGLITO.clsCADPROGLITO")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
    objCADPROGLITO.FILIAL = FILIAL
   
    If intFILIALPED = 0 Then strFILIAL = "NOVALATA "
    If intFILIALPED = 1 Then strFILIAL = "STEEL "
    
    strNOMFILIAL = ""
    If intFILIALPED = 1 Then strNOMFILIAL = "_STEEL"
    
    strCAPTION = Me.Caption & " / " & strFILIAL
    
    Call IniciaForm

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub IniciaForm()

    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call LimpaCamposlabel
    Call DesabilitaCampos
    
    objCADPROGLITO.Codigo = iCodigo
    
    If cTipOper = "I" Then
        txtUSUARIOCRIA.Text = strUsuario
        mskDTCRIACAOUSU.Text = Format(Now, "DD/MM/YYYY")
        mskDTMOV.Text = Format(Now, "DD/MM/YYYY")
    End If
    
    
    Call ConfGrid
    Call ConfGridOPProgramadas
     
    pbDADOS.Visible = False
    pbDADOS.Min = 0
    
    ''Call CarregaCampos
    
End Sub

Private Sub DesabilitaCampos()
    If cTipOper = "I" Then Frame2.Enabled = True
    If cTipOper = "C" Or cTipOper = "A" Then Frame2.Enabled = False
End Sub


Private Sub LimpaCamposlabel()
    lblDESCMAQ.Caption = ""
    lblDESCLIE.Caption = ""
    lblDESCAPAC.Caption = ""
End Sub

Private Sub ConfGrid()

    With grdOPS
    
       .Cols = conColumnsIn_Prod
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Prod_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_Prod_OP) = ""
       .ColDataType(conCOL_Prod_OP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_Pedido) = ""
       .ColDataType(conCOL_Prod_Pedido) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_DTENTREGA) = ""
       .ColDataType(conCOL_Prod_DTENTREGA) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_Prod_QTDEOP) = ""
       .ColDataType(conCOL_Prod_QTDEOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_QTDFAT) = ""
       .ColDataType(conCOL_Prod_QTDFAT) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_SALDO) = ""
       .ColDataType(conCOL_Prod_SALDO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_ID) = ""
       .ColDataType(conCOL_Prod_ID) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_ROTULO) = ""
       .ColDataType(conCOL_Prod_ROTULO) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Prod_DESCROT) = ""
       .ColDataType(conCOL_Prod_DESCROT) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Prod_CODCAPAC) = ""
       .ColDataType(conCOL_Prod_CODCAPAC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_CAPACIDADE) = ""
       .ColDataType(conCOL_Prod_CAPACIDADE) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Prod_CODCLIE) = ""
       .ColDataType(conCOL_Prod_CODCLIE) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_DESCCLIE) = ""
       .ColDataType(conCOL_Prod_DESCCLIE) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Prod_FotNovo) = ""
       .ColDataType(conCOL_Prod_FotNovo) = flexDTString
       .ColComboList(conCOL_Prod_FotNovo) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_Prod_NeckIN) = ""
       .ColDataType(conCOL_Prod_NeckIN) = flexDTString
       .ColComboList(conCOL_Prod_NeckIN) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_Prod_Fechamento) = ""
       .ColDataType(conCOL_Prod_Fechamento) = flexDTString
       .ColComboList(conCOL_Prod_Fechamento) = objCADPROGLITO.PreenchComboFechamentoGrdSA
       
       .Cell(flexcpData, 0, conCOL_Prod_QtdeFls) = ""
       .ColDataType(conCOL_Prod_QtdeFls) = flexDTLong
       
       
       .Cell(flexcpData, 0, conCOL_Prod_QtdeLatas) = ""
       .ColDataType(conCOL_Prod_QtdeLatas) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_DtPEdido) = ""
       .ColDataType(conCOL_Prod_DtPEdido) = flexDTDate
       
       
       .ColWidth(conCOL_Prod_OP) = 1000
       .ColWidth(conCOL_Prod_Pedido) = 1000
       .ColWidth(conCOL_Prod_DTENTREGA) = 1000
       .ColWidth(conCOL_Prod_QTDEOP) = 1000
       .ColWidth(conCOL_Prod_QTDFAT) = 1000
       .ColWidth(conCOL_Prod_SALDO) = 1000
       .ColWidth(conCOL_Prod_ID) = 0
       .ColWidth(conCOL_Prod_ROTULO) = 1200
       .ColWidth(conCOL_Prod_DESCROT) = 5000
       .ColWidth(conCOL_Prod_CODCAPAC) = 0
       .ColWidth(conCOL_Prod_CAPACIDADE) = 2000
       .ColWidth(conCOL_Prod_CODCLIE) = 0
       .ColWidth(conCOL_Prod_DESCCLIE) = 5000
       .ColWidth(conCOL_Prod_FotNovo) = 1200
       .ColWidth(conCOL_Prod_NeckIN) = 1200
       .ColWidth(conCOL_Prod_Fechamento) = 1200
       .ColWidth(conCOL_Prod_CodVerniz) = 0
       .ColWidth(conCOL_Prod_Verniz) = 2500
       .ColWidth(conCOL_Prod_CodEsmalte) = 0
       .ColWidth(conCOL_Prod_Esmalte) = 2500
       .ColWidth(conCOL_Prod_QtdeFls) = 1200
       .ColWidth(conCOL_Prod_QtdeLatas) = 1200
       .ColWidth(conCOL_Prod_DtPEdido) = 1200
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub

Private Sub ConfGridOPProgramadas()

    With grdOPSPROGRAMADAS
    
       .Cols = conColumnsIn_ProdProg
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ProdProg_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_ProdProg_OP) = ""
       .ColDataType(conCOL_ProdProg_OP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdProg_DTENTREGA) = ""
       .ColDataType(conCOL_ProdProg_DTENTREGA) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_ProdProg_QTDEOP) = ""
       .ColDataType(conCOL_ProdProg_QTDEOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdProg_QTDEFAT) = ""
       .ColDataType(conCOL_ProdProg_QTDEFAT) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdProg_SALDO) = ""
       .ColDataType(conCOL_ProdProg_SALDO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdProg_ID) = ""
       .ColDataType(conCOL_ProdProg_ID) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdProg_ROTULO) = ""
       .ColDataType(conCOL_ProdProg_ROTULO) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdProg_DESCROT) = ""
       .ColDataType(conCOL_ProdProg_DESCROT) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdProg_CODCAPAC) = ""
       .ColDataType(conCOL_ProdProg_CODCAPAC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdProg_CAPACIDADE) = ""
       .ColDataType(conCOL_ProdProg_CAPACIDADE) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ProdProg_CODCLIE) = ""
       .ColDataType(conCOL_ProdProg_CODCLIE) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ProdProg_DESCCLIE) = ""
       .ColDataType(conCOL_ProdProg_DESCCLIE) = flexDTString
       
       .ColWidth(conCOL_ProdProg_OP) = 1000
       .ColWidth(conCOL_ProdProg_DTENTREGA) = 1000
       .ColWidth(conCOL_ProdProg_QTDEOP) = 1000
       .ColWidth(conCOL_ProdProg_QTDEFAT) = 1000
       .ColWidth(conCOL_ProdProg_SALDO) = 1000
       .ColWidth(conCOL_ProdProg_ID) = 0
       .ColWidth(conCOL_ProdProg_ROTULO) = 1500
       .ColWidth(conCOL_ProdProg_DESCROT) = 4000
       .ColWidth(conCOL_ProdProg_CODCAPAC) = 0
       .ColWidth(conCOL_ProdProg_CAPACIDADE) = 3000
       .ColWidth(conCOL_ProdProg_CODCLIE) = 0
       .ColWidth(conCOL_ProdProg_DESCCLIE) = 4000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub


Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
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
    
    If Year(CDate(mskDTINI.Text)) <> Year(CDate(mskDTFIN.Text)) Then
        MsgBox "O Ano não pode ser diferente !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTINI.SetFocus
        Exit Function
    End If
    
    ConfereCampos = True

End Function

Private Sub CarregaDadosOP(strDTINI As String, strDTFIN As String)

    Call ConfGrid

    Dim lngQTDFOLHAS  As Long
    Dim dblPERDAPROC  As Long
    Dim lngQTDEFOLHAS As Long
    
    

    pbDADOS.Visible = True
    pbDADOS.Min = 0

    Dim lngQTDREGS As Long

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ORDP.*" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO   As SGI_DESCPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_NECKIN" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_FechSoldaAgrafado" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_QTDCORPSPADRAOSN" & vbCrLf       '' Pega Produto S/N - 0 = Não / 1 = Sim
    sSql = sSql & "      ,PROD.SGI_QTDEPORFOLHA" & vbCrLf
    sSql = sSql & "      ,LINH.SGI_DESCRI      As SGI_DESCLINHA" & vbCrLf
    sSql = sSql & "      ,LINH.SGI_PERDPROC" & vbCrLf
    sSql = sSql & "      ,LINH.SGI_QTDECORPOS" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & " ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINH" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf

    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_DATENTREGA Between " & strDTINI & " And " & strDTFIN & vbCrLf
    sSql = sSql & "   And (ORDP.SGI_STATUS = 0 or ORDP.SGI_STATUS = 1)" & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ORDP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And LINH.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINH.SGI_CODLIN     = PROD.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO     = PROD.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "Order By ORDP.SGI_DATENTREGA"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
    
        With grdOPS
            
            lngQTDREGS = 0
            Do While Not BREC.EOF()
                lngQTDREGS = (lngQTDREGS + 1)
                BREC.MoveNext
            Loop
        
        
            BREC.MoveFirst
            pbDADOS.Min = 0
            pbDADOS.Max = lngQTDREGS
            lngQTDREGS = 0
            
            Me.MousePointer = 11
            
            Do While Not BREC.EOF()
    
                lngQTDREGS = (lngQTDREGS + 1)
                pbDADOS.Value = lngQTDREGS
                
                
                lngQTDFOLHAS = 0
                dblPERDAPROC = 0
                lngQTDFOLHAS = 0
                If BREC!SGI_QTDCORPSPADRAOSN = 0 Then '' Folhas Padrão 0 = Não
                   If Not IsNull(BREC!SGI_QTDEPORFOLHA) Then lngQTDFOLHAS = BREC!SGI_QTDEPORFOLHA
                   dblPERDAPROC = 1.5
                ElseIf BREC!SGI_QTDCORPSPADRAOSN = 1 Then '' Folhas Padrão 1 = Sim
                   If Not IsNull(BREC!SGI_QTDECORPOS) Then lngQTDFOLHAS = BREC!SGI_QTDECORPOS
                   If Not IsNull(BREC!SGI_PERDPROC) Then dblPERDAPROC = BREC!SGI_PERDPROC
                End If
                If lngQTDFOLHAS > 0 Then lngQTDEFOLHAS = ((BREC!SGI_QTDEPED * dblPERDAPROC) / lngQTDFOLHAS)
    
                .AddItem BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_CODPED & vbTab & _
                         BREC!SGI_DATENTREGA & vbTab & _
                         BREC!SGI_QTDEPED & vbTab & _
                         BREC!SGI_QTDFAT & vbTab & _
                         BREC!SGI_SALDO & vbTab & _
                         BREC!SGI_IDPRODUTO & vbTab & _
                         BREC!SGI_CODPROD & vbTab & _
                         BREC!SGI_DESCPROD & vbTab & _
                         BREC!SGI_CODLINPROD & vbTab & _
                         BREC!SGI_DESCLINHA & vbTab & _
                         BREC!SGI_CODCLIE & vbTab & _
                         BREC!SGI_RAZAOSOC & vbTab & _
                         BREC!SGI_FOTNOVO & vbTab & _
                         BREC!SGI_NECKIN & vbTab & _
                         BREC!SGI_FechSoldaAgrafado & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         lngQTDEFOLHAS & vbTab & _
                         "" & vbTab & _
                         Format(BREC!SGI_DATAPED, "DD/MM/YYYY")
                         
                '' -------------------
                '' Verniz
                sSql = ""
                
                sSql = "Select" & vbCrLf
                sSql = sSql & "       VERNIZ.SGI_PRODUTO" & vbCrLf
                sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_VERNIZPROD VERNIZ" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       VERNIZ.SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & "   And VERNIZ.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "   And PROD.SGI_FILIAL      = VERNIZ.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And PROD.SGI_IDPRODUTO   = VERNIZ.SGI_PRODUTO" & vbCrLf
                
                BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC12.EOF() Then
                   .Cell(flexcpText, (.Rows - 1), conCOL_Prod_CodVerniz) = BREC12!SGI_PRODUTO
                   .Cell(flexcpText, (.Rows - 1), conCOL_Prod_Verniz) = BREC12!SGI_DESCRICAO
                End If
                BREC12.Close
                
                '' -------------------
                '' Esmalte
                sSql = ""
                
                sSql = "Select" & vbCrLf
                sSql = sSql & "       ESMALTE.SGI_PRODUTO" & vbCrLf
                sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_ESMALTEPROD ESMALTE" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO  PROD" & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       ESMALTE.SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & "   And ESMALTE.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "   And PROD.SGI_FILIAL       = ESMALTE.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And PROD.SGI_IDPRODUTO    = ESMALTE.SGI_PRODUTO" & vbCrLf
                
                BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC12.EOF() Then
                   .Cell(flexcpText, (.Rows - 1), conCOL_Prod_CodEsmalte) = BREC12!SGI_PRODUTO
                   .Cell(flexcpText, (.Rows - 1), conCOL_Prod_Esmalte) = BREC12!SGI_DESCRICAO
                End If
                BREC12.Close
                
                BREC.MoveNext
            Loop
        End With
    
    Else
        MsgBox "ATENÇÃO - Não há dados de OP's para carregar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

    pbDADOS.Visible = False
    Me.MousePointer = 0


End Sub
