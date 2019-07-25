VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADNATOPERACAO 
   Caption         =   "Cadastro de natureza de operação"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12690
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame12 
      Caption         =   "[ Protocolo no Simples ]"
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
      Left            =   6360
      TabIndex        =   53
      Top             =   7560
      Width           =   6255
      Begin VB.TextBox txtProtocoloSimp 
         Height          =   615
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Text            =   "frmCADNATOPERACAO.frx":0000
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "[ Protocolo ]"
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
      Left            =   0
      TabIndex        =   28
      Top             =   7560
      Width           =   6255
      Begin VB.TextBox txtProtocolo 
         Height          =   615
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "frmCADNATOPERACAO.frx":0013
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   300
      Left            =   12360
      Picture         =   "frmCADNATOPERACAO.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exclui a linha da Gride Selecionada"
      Top             =   4320
      Width           =   300
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   12360
      Picture         =   "frmCADNATOPERACAO.frx":016A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Inclui uma nova linha na Gride"
      Top             =   3960
      Width           =   300
   End
   Begin VSFlex8LCtl.VSFlexGrid grdNATOPER 
      Height          =   3495
      Left            =   0
      TabIndex        =   25
      Top             =   3960
      Width           =   12255
      _cx             =   21616
      _cy             =   6165
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
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   12615
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2160
         TabIndex        =   50
         Top             =   2520
         Width           =   1935
         Begin VB.OptionButton optRegimeSTEsp 
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
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optRegimeSTEsp 
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
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   51
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.TextBox txtALIQCOFINS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         TabIndex        =   47
         Text            =   "txtALIQCOFINS"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtALIQPIS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         TabIndex        =   46
         Text            =   "txtALIQPIS"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2160
         TabIndex        =   41
         Top             =   2280
         Width           =   1575
         Begin VB.Frame Frame10 
            Caption         =   "Frame10"
            Height          =   15
            Left            =   240
            TabIndex        =   49
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optEspecial02 
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
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optEspecial02 
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
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   42
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame Frame34 
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
         Left            =   2160
         TabIndex        =   36
         Top             =   2040
         Width           =   1575
         Begin VB.OptionButton optExpecial 
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
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   38
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optExpecial 
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
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtSitTrib 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   34
         Text            =   "txtSitTrib"
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2160
         TabIndex        =   30
         Top             =   1800
         Width           =   4215
         Begin VB.OptionButton optPessoaFJ 
            Caption         =   "Pessoa Fisica"
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
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton optPessoaFJ 
            Caption         =   "Pessoa Juridica"
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
            TabIndex        =   31
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton optDefault 
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
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optDefault 
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
            Left            =   720
            TabIndex        =   22
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtCODIGO 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "txtCODIGO"
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
         Begin VB.OptionButton optImpExp 
            Caption         =   "Exportação"
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
            TabIndex        =   18
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optImpExp 
            Caption         =   "Importação"
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
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   1560
         Width           =   3855
         Begin VB.OptionButton optDentroForaEst 
            Caption         =   "Fora do Estado"
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
            TabIndex        =   15
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton optDentroForaEst 
            Caption         =   "Dentro do Estado"
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
            TabIndex        =   14
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1335
         Width           =   4335
         Begin VB.OptionButton optEntSai 
            Caption         =   "Transferência"
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
            Left            =   2760
            TabIndex        =   35
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton optEntSai 
            Caption         =   "Saida"
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
            Left            =   1800
            TabIndex        =   12
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optEntSai 
            Caption         =   "Entrada"
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
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNomeCla 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtNomecla"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Regime ST Especial"
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
         Left            =   240
         TabIndex        =   48
         Top             =   2520
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aliq COFINS"
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
         Left            =   6240
         TabIndex        =   45
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aliq PIS"
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
         Left            =   6600
         TabIndex        =   44
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Especial - 02"
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
         Left            =   240
         TabIndex        =   40
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Especial - 01"
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
         Left            =   240
         TabIndex        =   39
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Sit.Tributária"
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
         Left            =   2880
         TabIndex        =   33
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Default"
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
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Códgo"
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
         Left            =   210
         TabIndex        =   19
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Nomeclatura"
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
         Left            =   210
         TabIndex        =   9
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Operação"
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
         Left            =   210
         TabIndex        =   7
         Top             =   1320
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12615
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
         Picture         =   "frmCADNATOPERACAO.frx":02B4
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADNATOPERACAO.frx":07E6
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCADNATOPERACAO.frx":08E8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADNATOPERACAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public lngCODIGO        As Long
Public FILIAL           As Integer
Public strAcesso        As String
Public lngCodUsuario    As Long
Dim objBLBFunc          As Object
Dim objCADNATOPERACAO   As Object
Dim arrALIQICMS         As Variant

Const conCOL_SonNatOper_EstOrigem               As Integer = 0
Const conCOL_SonNatOper_EstDestino              As Integer = 1
Const conCOL_SonNatOper_ALiqICMS                As Integer = 2

Const conCOL_SonNatOper_SubstTribSN             As Integer = 3
Const conCOL_SonNatOper_AliqIVASTORIG           As Integer = 4
Const conCOL_SonNatOper_AliqICMSINT             As Integer = 5
Const conCOL_SonNatOper_AliqST                  As Integer = 6

Const conCOL_SonNatOper_ESTPESQ                 As Integer = 7
Const conCOL_SonNatOper_Protocolo               As Integer = 8

Const conCOL_SonNatOper_EnqSimpSN               As Integer = 9
Const conCOL_SonNatOper_ProtOptSimp             As Integer = 10
Const conCOL_SonNatOper_ALiqICMSSIMPL           As Integer = 11
Const conCOL_SonNatOper_AliqIVASTORIGSIMPL      As Integer = 12
Const conCOL_SonNatOper_AliqICMSINTSIMPL        As Integer = 13
Const conCOL_SonNatOper_AliqSTSIMPL             As Integer = 14

Const conCOL_SonNatOper_FormatString            As String = "=Est.Orig|Est.Dest|%ICMS|St.S/N|%IVA.Orig|%ICMS.Interno|%IVA.Ajust|EstPSQ|Protocolo|Enq.Simples S/N|ProtEnqSimp|%ICMS|%IVA.Orig|%ICMS.Interno|%IVA.Ajust"
Const conColumnsIn_SonNatOper                   As Integer = 15

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de natureza de operação - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtNomeCla.Enabled = True
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim lngCodLog   As Long
    Dim sValor      As String
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADNATOPERACAO.CODIGO = objBLBFunc.Gera_Codigo(Me.Name, FILIAL, Linha)
    
    objCADNATOPERACAO.NOMECLA = txtNomeCla.Text
    objCADNATOPERACAO.DESCRICAO = txtDescricao.Text
    objCADNATOPERACAO.SITTRIB = txtSitTrib.Text
    
    If optEntSai(0).Value = True Then objCADNATOPERACAO.ENTSAI = 0
    If optEntSai(1).Value = True Then objCADNATOPERACAO.ENTSAI = 1
    If optEntSai(2).Value = True Then objCADNATOPERACAO.ENTSAI = 2
    
    If optDentroForaEst(0).Value = True Then objCADNATOPERACAO.DENFOREST = 0
    If optDentroForaEst(1).Value = True Then objCADNATOPERACAO.DENFOREST = 1
    
    If optImpExp(0).Value = True Then objCADNATOPERACAO.IMPEXT = 0
    If optImpExp(1).Value = True Then objCADNATOPERACAO.IMPEXT = 1
    
    If optDefault(0).Value = True Then objCADNATOPERACAO.DEFAULT = 0
    If optDefault(1).Value = True Then objCADNATOPERACAO.DEFAULT = 1
        
    If optPessoaFJ(1).Value = True Then objCADNATOPERACAO.PessoaFJ = 1
    If optPessoaFJ(0).Value = True Then objCADNATOPERACAO.PessoaFJ = 0
    
    If optExpecial(0).Value = True Then objCADNATOPERACAO.ESPECIAL = 0
    If optExpecial(1).Value = True Then objCADNATOPERACAO.ESPECIAL = 1
    
    If optEspecial02(0).Value = True Then objCADNATOPERACAO.ESPECIAL02 = 0
    If optEspecial02(1).Value = True Then objCADNATOPERACAO.ESPECIAL02 = 1
    
    objCADNATOPERACAO.ALIQPIS = "Null"
    objCADNATOPERACAO.ALIQCOFINS = "Null"
    If Len(Trim(txtALIQPIS.Text)) > 0 Then
        strValor = Replace(txtALIQPIS.Text, ".", "")
        strValor = Replace(strValor, ",", ".")
        objCADNATOPERACAO.ALIQPIS = strValor
    End If
    If Len(Trim(txtALIQCOFINS.Text)) > 0 Then
        strValor = Replace(txtALIQCOFINS.Text, ".", "")
        strValor = Replace(strValor, ",", ".")
        objCADNATOPERACAO.ALIQCOFINS = strValor
    End If
    
    arrALIQICMS = Empty
    With grdNATOPER
        If (.Rows - 1) > 0 Then
           ReDim arrALIQICMS(1 To (.Rows - 1), 1 To 14) As String
           For I = 1 To (.Rows - 1)
                arrALIQICMS(I, 1) = .Cell(flexcpText, I, conCOL_SonNatOper_EstOrigem)
                arrALIQICMS(I, 2) = .Cell(flexcpText, I, conCOL_SonNatOper_EstDestino)
               
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_ALiqICMS))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_SonNatOper_ALiqICMS), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrALIQICMS(I, 3) = sValor
           
                arrALIQICMS(I, 4) = .Cell(flexcpText, I, conCOL_SonNatOper_SubstTribSN)
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_AliqST))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_SonNatOper_AliqST), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrALIQICMS(I, 5) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_AliqICMSINT))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_SonNatOper_AliqICMSINT), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrALIQICMS(I, 6) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_AliqIVASTORIG))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_SonNatOper_AliqIVASTORIG), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrALIQICMS(I, 7) = sValor
                arrALIQICMS(I, 8) = Trim(Replace(.Cell(flexcpText, I, conCOL_SonNatOper_Protocolo), "'", ""))
                
                arrALIQICMS(I, 9) = .Cell(flexcpText, I, conCOL_SonNatOper_EnqSimpSN)
                
                arrALIQICMS(I, 10) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_ProtOptSimp))) > 0 Then
                    arrALIQICMS(I, 10) = "'" & Trim(Replace(.Cell(flexcpText, I, conCOL_SonNatOper_ProtOptSimp), "'", "")) & "'"
                End If
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_ALiqICMSSIMPL))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_SonNatOper_ALiqICMSSIMPL), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrALIQICMS(I, 11) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_AliqIVASTORIGSIMPL))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_SonNatOper_AliqIVASTORIGSIMPL), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrALIQICMS(I, 12) = sValor
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_AliqICMSINTSIMPL))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_SonNatOper_AliqICMSINTSIMPL), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrALIQICMS(I, 13) = sValor
           
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonNatOper_AliqSTSIMPL))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_SonNatOper_AliqSTSIMPL), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrALIQICMS(I, 14) = sValor
           
           Next I
        End If
    End With
    objCADNATOPERACAO.ALIQICMS = arrALIQICMS
    If optExpecial(0).Value = True Then objCADNATOPERACAO.ESPECIAL = 0
    If optExpecial(1).Value = True Then objCADNATOPERACAO.ESPECIAL = 1
        
    If optRegimeSTEsp(0).Value = True Then objCADNATOPERACAO.REGIMESTESP = 0
    If optRegimeSTEsp(1).Value = True Then objCADNATOPERACAO.REGIMESTESP = 1
    
    If objCADNATOPERACAO.GRAVA(cTipOper) = False Then Exit Sub
          
    '' Atualizando os Dados
    If objBLBFunc.Atualiza(cTipOper, Str(objCADNATOPERACAO.CODIGO), FILIAL, Me.Name, Linha) = False Then Exit Sub
    
    '' Gerando Log de Sistema
    lngCodLog = objBLBFunc.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objBLBFunc.GravaLogModulo(FILIAL, lngCodLog, Me.Name, cTipOper, lngCodUsuario, Str(objCADNATOPERACAO.CODIGO), Linha)
    
    MsgBox "Natureza de operação foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADNATOPERACAO = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADNATOPERACAO = Nothing
    Unload Me
End Sub

Private Sub Command4_Click()
    If cTipOper = "C" Then Exit Sub
    Call objBLBFunc.ExclLinhaGrid(grdNATOPER, grdNATOPER.Row)
End Sub

Private Sub Command5_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGridNatOper
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADNATOPERACAO = CreateObject("CADNATOPERACAO.clsCADNATOPERACAO")
   
   objCADNATOPERACAO.FILIAL = FILIAL
   
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de natureza de operação - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADNATOPERACAO
    
    optEntSai(0).Value = False
    optEntSai(1).Value = False
    
    optDentroForaEst(0).Value = False
    optDentroForaEst(1).Value = False
    
    optImpExp(0).Value = False
    optImpExp(1).Value = False
    
    Call InitGridNatOper
    
    optDefault(0).Value = True
    optImpExp(0).Value = True
   
    optExpecial(0).Value = True
    optEspecial02(0).Value = True
    
    optRegimeSTEsp(0).Value = True
    
End Sub

Private Sub grdNATOPER_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     Dim curIVAAJUSTADO         As Double
     Dim curIVAAJUSTADOSIMPL    As Double
     With grdNATOPER
          Select Case Col
                 Case conCOL_SonNatOper_ALiqICMS, _
                      conCOL_SonNatOper_ALiqICMSSIMPL, _
                      conCOL_SonNatOper_AliqST, _
                      conCOL_SonNatOper_AliqSTSIMPL, _
                      conCOL_SonNatOper_AliqICMSINT, _
                      conCOL_SonNatOper_AliqIVASTORIG, _
                      conCOL_SonNatOper_AliqIVASTORIGSIMPL, _
                      conCOL_SonNatOper_AliqICMSINTSIMPL
                        .Cell(flexcpText, Row, Col) = Format(.Cell(flexcpText, Row, Col), "#,##0.00")
                        curIVAAJUSTADO = CalcIVA_AJUSTADO(Row)
                        curIVAAJUSTADOSIMPL = CalcIVA_AJUSTADO_SIMPLES(Row)
                        If curIVAAJUSTADO > 0 Then .Cell(flexcpText, Row, conCOL_SonNatOper_AliqST) = Format(curIVAAJUSTADO, "#,##0.00")
                        If curIVAAJUSTADOSIMPL > 0 Then .Cell(flexcpText, Row, conCOL_SonNatOper_AliqSTSIMPL) = Format(curIVAAJUSTADOSIMPL, "#,##0.00")
                 Case conCOL_SonNatOper_SubstTribSN
                        If .Cell(flexcpText, Row, conCOL_SonNatOper_SubstTribSN) = -1 Then Call PosCol(grdNATOPER, conCOL_SonNatOper_AliqIVASTORIG, Row)
          End Select
     End With
End Sub

Private Sub grdNATOPER_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdNATOPER
        Select Case Col
        Case conCOL_SonNatOper_EstOrigem, _
             conCOL_SonNatOper_EstDestino, _
             conCOL_SonNatOper_ALiqICMS
             If cTipOper = "C" Then Cancel = True
        Case conCOL_SonNatOper_AliqST
             If cTipOper = "C" Then Cancel = True
             If .Cell(flexcpText, Row, conCOL_SonNatOper_SubstTribSN) = 0 Then Cancel = True
        Case conCOL_SonNatOper_AliqIVASTORIG, _
             conCOL_SonNatOper_AliqICMSINT, _
             conCOL_SonNatOper_AliqST, _
             conCOL_SonNatOper_EnqSimpSN
             If cTipOper = "C" Then Cancel = True
             If .Cell(flexcpText, Row, conCOL_SonNatOper_SubstTribSN) = 0 Then Cancel = True
        Case conCOL_SonNatOper_ALiqICMSSIMPL, _
             conCOL_SonNatOper_AliqIVASTORIGSIMPL, _
             conCOL_SonNatOper_AliqICMSINTSIMPL, _
             conCOL_SonNatOper_AliqSTSIMPL
             If cTipOper = "C" Then Cancel = True
             If .Cell(flexcpText, Row, conCOL_SonNatOper_EnqSimpSN) = 0 Then Cancel = True
        Case Else
            .ComboList = ""
        End Select
    End With
    Exit Sub
End Sub

Private Sub grdNATOPER_Click()
    Call PopCampoProtocolo(grdNATOPER.Row)
End Sub

Private Sub grdNATOPER_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdNATOPER
          Select Case Col
                    Case conCOL_SonNatOper_ALiqICMS, _
                         conCOL_SonNatOper_ALiqICMSSIMPL, _
                         conCOL_SonNatOper_AliqST, _
                         conCOL_SonNatOper_AliqICMSINT, _
                         conCOL_SonNatOper_AliqIVASTORIG, _
                         conCOL_SonNatOper_AliqIVASTORIGSIMPL, _
                         conCOL_SonNatOper_AliqICMSINTSIMPL, _
                         conCOL_SonNatOper_AliqSTSIMPL
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsCurrency)
          End Select
     End With
End Sub

Private Sub grdNATOPER_RowColChange()
    Call PopCampoProtocolo(grdNATOPER.Row)
End Sub

Private Sub grdNATOPER_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     Dim strIDPESQ  As String
     Dim lngRow     As Long
     
     With grdNATOPER
          Select Case Col
                 Case conCOL_SonNatOper_EstDestino
                        If .EditText = Empty Then Exit Sub
                        strIDPESQ = Trim(.Cell(flexcpText, Row, conCOL_SonNatOper_EstOrigem)) & Trim(.EditText)
                        lngRow = .FindRow(strIDPESQ, , conCOL_SonNatOper_ESTPESQ)
                        
                        If lngRow <> -1 Then
                           If lngRow = Row Then Exit Sub
                           MsgBox "ATENÇÃO - Estado Origem e Destino já existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_SonNatOper_ESTPESQ) = Trim(strIDPESQ)
                        Call PosCol(grdNATOPER, conCOL_SonNatOper_ALiqICMS, Row)
                Case conCOL_SonNatOper_AliqIVASTORIG, _
                     conCOL_SonNatOper_AliqIVASTORIGSIMPL
                        If .EditText = Empty Then Exit Sub
          End Select
     End With

End Sub

Private Sub Text1_GotFocus()
    objBLBFunc.SelecionaCampos txtALIQPIS.Name, frmCADCLASSFISC
End Sub

Private Sub txtALIQCOFINS_GotFocus()
    objBLBFunc.SelecionaCampos txtALIQCOFINS.Name, frmCADNATOPERACAO
End Sub

Private Sub txtALIQCOFINS_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.SoNumeroPonto(KeyAscii, txtALIQCOFINS.Text)
End Sub

Private Sub txtALIQCOFINS_Validate(Cancel As Boolean)
    If Len(Trim(txtALIQCOFINS.Text)) = 0 Then Exit Sub
    txtALIQCOFINS.Text = Format(txtALIQCOFINS.Text, "#,##0.00")
End Sub

Private Sub txtALIQPIS_GotFocus()
    objBLBFunc.SelecionaCampos txtALIQPIS.Name, frmCADNATOPERACAO
End Sub

Private Sub txtALIQPIS_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.SoNumeroPonto(KeyAscii, txtALIQPIS.Text)
End Sub

Private Sub txtALIQPIS_Validate(Cancel As Boolean)
    If Len(Trim(txtALIQPIS.Text)) = 0 Then Exit Sub
    txtALIQPIS.Text = Format(txtALIQPIS.Text, "#,##0.00")
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADNATOPERACAO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     Dim intDefault As Integer
     
     If Len(Trim(txtDescricao.Text)) = 0 Then
        MsgBox "Especificação técnica inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If optEntSai(0).Value = False And optEntSai(1).Value = False And optEntSai(2).Value = False Then
        MsgBox "Entrada/saida/transferência não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     If optDentroForaEst(0).Value = False And optDentroForaEst(1).Value = False Then
        MsgBox "Dentro e fora do estado não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     If optImpExp(0).Value = False And optImpExp(1).Value = False Then
        MsgBox "Importação e exportação não pode ser nulo !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     If optPessoaFJ(0).Value = False And optPessoaFJ(1).Value = False Then
        MsgBox "Informe se é pessoa fisica ou juridica !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     If cTipOper = "I" Then
        If optDefault(1).Value = True Then
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
            If optEntSai(0).Value = True Then sSql = sSql & "   And SGI_ENTSAI = 0" & vbCrLf
            If optEntSai(1).Value = True Then sSql = sSql & "   And SGI_ENTSAI = 1" & vbCrLf
            If optDefault(1).Value = True Then sSql = sSql & "   And SGI_DEFAULT = 1"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF Then
                MsgBox "Já existe natureza de operação padrão !!!", vbOKOnly + vbExclamation, "Aviso"
                optDefault(0).Value = True
                BREC2.Close
                Exit Function
            End If
            BREC2.Close
        
        End If
        
        '' -- Código
        ''sSql = "Select " & vbCrLf
        ''sSql = sSql & "       * " & vbCrLf
        ''sSql = sSql & "  from " & vbCrLf
        ''sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
        ''sSql = sSql & " Where " & vbCrLf
        ''sSql = sSql & "       SGI_NOMECLCOD    = '" & txtNomeCla.Text & "'" & vbCrLf
        ''sSql = sSql & "   And SGI_FILIAL       = " & FILIAL & vbCrLf
        ''If optPessoaFJ(0).Value = True Then sSql = sSql & "   And SGI_PESSOAFJ     = 0" & vbCrLf
        ''If optPessoaFJ(1).Value = True Then sSql = sSql & "   And SGI_PESSOAFJ     = 1"
        
        ''BREC.Open sSql, adoBanco_Dados
        ''If Not BREC.EOF Then
        ''   MsgBox "Código já existe !!!", vbOKOnly + vbExclamation, "Aviso"
        ''   txtNomeCla.SetFocus
        ''   BREC.Close
        ''   Exit Function
        ''End If
        ''BREC.Close
                
     End If
     
     If cTipOper = "A" Then
                                
        If optDefault(0).Value = True Then intDefault = 0
        If optDefault(1).Value = True Then intDefault = 1
        
        If optDefault(1).Value = True Then
            If objCADNATOPERACAO.DEFAULT <> intDefault Then
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                If optEntSai(0).Value = True Then sSql = sSql & "   And SGI_ENTSAI = 0" & vbCrLf
                If optEntSai(1).Value = True Then sSql = sSql & "   And SGI_ENTSAI = 1" & vbCrLf
                If optDefault(1).Value = True Then sSql = sSql & "   And SGI_DEFAULT = 1"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF Then
                    MsgBox "Já existe natureza de operação padrão !!!", vbOKOnly + vbExclamation, "Aviso"
                    optDefault(0).Value = True
                    BREC2.Close
                    Exit Function
                End If
                BREC2.Close
            End If
        End If
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de natureza de operação - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADNATOPERACAO
    
    objCADNATOPERACAO.CODIGO = lngCODIGO
    
    Call InitGridNatOper
    optImpExp(0).Value = True

    optExpecial(0).Value = True
    optEspecial02(0).Value = True
    
    optRegimeSTEsp(0).Value = True
    
    If objCADNATOPERACAO.Carrega_campos = True Then

       txtCODIGO.Text = objCADNATOPERACAO.CODIGO
       txtNomeCla.Text = objCADNATOPERACAO.NOMECLA
       txtDescricao.Text = objCADNATOPERACAO.DESCRICAO
       txtSitTrib.Text = objCADNATOPERACAO.SITTRIB
       
       optEntSai(objCADNATOPERACAO.ENTSAI).Value = True
       optDentroForaEst(objCADNATOPERACAO.DENFOREST).Value = True
       optImpExp(objCADNATOPERACAO.IMPEXT).Value = True
       optDefault(objCADNATOPERACAO.DEFAULT).Value = True
       optPessoaFJ(objCADNATOPERACAO.PessoaFJ).Value = True
       
       If Len(Trim(objCADNATOPERACAO.ALIQPIS)) > 0 Then txtALIQPIS.Text = objCADNATOPERACAO.ALIQPIS
       If Len(Trim(objCADNATOPERACAO.ALIQCOFINS)) > 0 Then txtALIQCOFINS.Text = objCADNATOPERACAO.ALIQCOFINS
       
       optRegimeSTEsp(objCADNATOPERACAO.REGIMESTESP).Value = True
       
       arrALIQICMS = objCADNATOPERACAO.ALIQICMS
       If IsArray(arrALIQICMS) Then
          With grdNATOPER
                For I = 1 To UBound(arrALIQICMS)
                    .AddItem arrALIQICMS(I, 1) & vbTab & _
                             arrALIQICMS(I, 2) & vbTab & _
                             Format(arrALIQICMS(I, 3), "#,##0.00") & vbTab & _
                             arrALIQICMS(I, 4) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 7))) = 0, "", Format(arrALIQICMS(I, 7), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 6))) = 0, "", Format(arrALIQICMS(I, 6), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 5))) = 0, "", Format(arrALIQICMS(I, 5), "#,##0.00")) & vbTab & _
                             "" & vbTab & _
                             arrALIQICMS(I, 8) & vbTab & _
                             arrALIQICMS(I, 9) & vbTab & _
                             arrALIQICMS(I, 10) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 11))) = 0, "", Format(arrALIQICMS(I, 11), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 12))) = 0, "", Format(arrALIQICMS(I, 12), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 13))) = 0, "", Format(arrALIQICMS(I, 13), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 14))) = 0, "", Format(arrALIQICMS(I, 14), "#,##0.00"))
                             
                     .Cell(flexcpText, (.Rows - 1), conCOL_SonNatOper_ESTPESQ) = Trim(arrALIQICMS(I, 1)) & Trim(.Cell(flexcpTextDisplay, .Rows - 1, conCOL_SonNatOper_EstDestino))
                Next I
          End With
       End If
       
       optExpecial(objCADNATOPERACAO.ESPECIAL).Value = True
       optEspecial02(objCADNATOPERACAO.ESPECIAL02).Value = True
       
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de natureza de operação - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADNATOPERACAO
    
    objCADNATOPERACAO.CODIGO = lngCODIGO
    
    Call InitGridNatOper
    optImpExp(0).Value = True
    
    optExpecial(0).Value = True
    optEspecial02(0).Value = True
    
    optRegimeSTEsp(0).Value = True
    
    If objCADNATOPERACAO.Carrega_campos = True Then

       txtCODIGO.Text = objCADNATOPERACAO.CODIGO
       txtNomeCla.Text = objCADNATOPERACAO.NOMECLA
       txtDescricao.Text = objCADNATOPERACAO.DESCRICAO
       txtSitTrib.Text = objCADNATOPERACAO.SITTRIB
       
       optEntSai(objCADNATOPERACAO.ENTSAI).Value = True
       optDentroForaEst(objCADNATOPERACAO.DENFOREST).Value = True
       optImpExp(objCADNATOPERACAO.IMPEXT).Value = True
       optDefault(objCADNATOPERACAO.DEFAULT).Value = True
       optPessoaFJ(objCADNATOPERACAO.PessoaFJ).Value = True
       
       If Len(Trim(objCADNATOPERACAO.ALIQPIS)) > 0 Then txtALIQPIS.Text = objCADNATOPERACAO.ALIQPIS
       If Len(Trim(objCADNATOPERACAO.ALIQCOFINS)) > 0 Then txtALIQCOFINS.Text = objCADNATOPERACAO.ALIQCOFINS
       
       optRegimeSTEsp(objCADNATOPERACAO.REGIMESTESP).Value = True
       
       arrALIQICMS = objCADNATOPERACAO.ALIQICMS
       If IsArray(arrALIQICMS) Then
          With grdNATOPER
                For I = 1 To UBound(arrALIQICMS)
                    .AddItem arrALIQICMS(I, 1) & vbTab & _
                             arrALIQICMS(I, 2) & vbTab & _
                             Format(arrALIQICMS(I, 3), "#,##0.00") & vbTab & _
                             arrALIQICMS(I, 4) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 7))) = 0, "", Format(arrALIQICMS(I, 7), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 6))) = 0, "", Format(arrALIQICMS(I, 6), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 5))) = 0, "", Format(arrALIQICMS(I, 5), "#,##0.00")) & vbTab & _
                             "" & vbTab & _
                             arrALIQICMS(I, 8) & vbTab & _
                             arrALIQICMS(I, 9) & vbTab & _
                             arrALIQICMS(I, 10) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 11))) = 0, "", Format(arrALIQICMS(I, 11), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 12))) = 0, "", Format(arrALIQICMS(I, 12), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 13))) = 0, "", Format(arrALIQICMS(I, 13), "#,##0.00")) & vbTab & _
                             IIf(Len(Trim(arrALIQICMS(I, 14))) = 0, "", Format(arrALIQICMS(I, 14), "#,##0.00"))
                             
                     .Cell(flexcpText, (.Rows - 1), conCOL_SonNatOper_ESTPESQ) = Trim(arrALIQICMS(I, 1)) & Trim(.Cell(flexcpTextDisplay, .Rows - 1, conCOL_SonNatOper_EstDestino))
                Next I
          End With
       End If
    
       optExpecial(objCADNATOPERACAO.ESPECIAL).Value = True
       optEspecial02(objCADNATOPERACAO.ESPECIAL02).Value = True
       
    End If

End Sub


Private Sub txtNomeCla_GotFocus()
    objBLBFunc.SelecionaCampos txtNomeCla.Name, frmCADNATOPERACAO
End Sub


Private Sub InitGridNatOper()

    With grdNATOPER
    
        .Cols = conColumnsIn_SonNatOper
        .Rows = 1
        .FixedCols = 0
        .FormatString = conCOL_SonNatOper_FormatString
        
        .AutoSizeMouse = False
        
        .AllowUserResizing = flexResizeBoth
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_EstOrigem) = ""
        .ColDataType(conCOL_SonNatOper_EstOrigem) = flexDTString
        .ColComboList(conCOL_SonNatOper_EstOrigem) = objBLBFunc.Preenche_EstadoGrid
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_EstDestino) = ""
        .ColDataType(conCOL_SonNatOper_EstDestino) = flexDTString
        .ColComboList(conCOL_SonNatOper_EstDestino) = objBLBFunc.Preenche_EstadoGrid
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_ALiqICMS) = ""
        .ColDataType(conCOL_SonNatOper_ALiqICMS) = flexDTCurrency
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_SubstTribSN) = ""
        .ColDataType(conCOL_SonNatOper_SubstTribSN) = flexDTBoolean
        .ColFormat(conCOL_SonNatOper_SubstTribSN) = "S;N"
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_AliqIVASTORIG) = ""
        .ColDataType(conCOL_SonNatOper_AliqIVASTORIG) = flexDTCurrency
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_AliqICMSINT) = ""
        .ColDataType(conCOL_SonNatOper_AliqICMSINT) = flexDTCurrency
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_AliqST) = ""
        .ColDataType(conCOL_SonNatOper_AliqST) = flexDTCurrency
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_ESTPESQ) = ""
        .ColDataType(conCOL_SonNatOper_ESTPESQ) = flexDTString
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_Protocolo) = ""
        .ColDataType(conCOL_SonNatOper_Protocolo) = flexDTString
       
        .Cell(flexcpData, 0, conCOL_SonNatOper_EnqSimpSN) = ""
        .ColDataType(conCOL_SonNatOper_EnqSimpSN) = flexDTBoolean
        .ColFormat(conCOL_SonNatOper_EnqSimpSN) = "Sim;Não"
        .ColAlignment(conCOL_SonNatOper_EnqSimpSN) = flexAlignCenterCenter
        
        .Cell(flexcpData, 0, conCOL_SonNatOper_ProtOptSimp) = ""
        .ColDataType(conCOL_SonNatOper_ProtOptSimp) = flexDTString
        
        .Cell(flexcpData, 0, conCOL_SonNatOper_ALiqICMSSIMPL) = ""
        .ColDataType(conCOL_SonNatOper_ALiqICMSSIMPL) = flexDTCurrency
        
        .Cell(flexcpData, 0, conCOL_SonNatOper_AliqIVASTORIGSIMPL) = ""
        .ColDataType(conCOL_SonNatOper_AliqIVASTORIGSIMPL) = flexDTCurrency
        
        .Cell(flexcpData, 0, conCOL_SonNatOper_AliqICMSINTSIMPL) = ""
        .ColDataType(conCOL_SonNatOper_AliqICMSINTSIMPL) = flexDTCurrency
        
        .Cell(flexcpData, 0, conCOL_SonNatOper_AliqSTSIMPL) = ""
        .ColDataType(conCOL_SonNatOper_AliqSTSIMPL) = flexDTCurrency
        
        .ColWidth(conCOL_SonNatOper_EstOrigem) = 800
        .ColWidth(conCOL_SonNatOper_EstDestino) = 800
        .ColWidth(conCOL_SonNatOper_ALiqICMS) = 700
        .ColWidth(conCOL_SonNatOper_SubstTribSN) = 600
        .ColWidth(conCOL_SonNatOper_EnqSimpSN) = 1300
        
        .ColWidth(conCOL_SonNatOper_AliqIVASTORIG) = 900
        .ColWidth(conCOL_SonNatOper_AliqICMSINT) = 1200
        .ColWidth(conCOL_SonNatOper_AliqST) = 900
        .ColWidth(conCOL_SonNatOper_ESTPESQ) = 0
        .ColWidth(conCOL_SonNatOper_Protocolo) = 0
        .ColWidth(conCOL_SonNatOper_ProtOptSimp) = 0
       
        .ColWidth(conCOL_SonNatOper_ALiqICMSSIMPL) = 700
        .ColWidth(conCOL_SonNatOper_AliqIVASTORIGSIMPL) = 900
        .ColWidth(conCOL_SonNatOper_AliqICMSINTSIMPL) = 1200
        .ColWidth(conCOL_SonNatOper_AliqSTSIMPL) = 900
       
        .Editable = flexEDKbdMouse
        .AllowSelection = False
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .BackColor = &H80000018
        .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub IncRegGridNatOper()
    
    '' ======================================
    '' Verificando Campos em Branco
    If objBLBFunc.TemLinhaVazia(grdNATOPER, conCOL_SonNatOper_EstOrigem) = True Then Exit Sub
    If objBLBFunc.TemLinhaVazia(grdNATOPER, conCOL_SonNatOper_EstDestino) = True Then Exit Sub
    '' ======================================
    
    Dim lngRow As Long
    
    With grdNATOPER
        
        .AddItem "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 0 & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 0 & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 ""
        
        Call PosRegistro(grdNATOPER, conCOL_SonNatOper_EstOrigem)
                          
    End With
End Sub

Private Sub PosRegistro(grdGenerica As VSFlexGrid, lngCODIGO As Long)
    Dim I As Long
    
    With grdGenerica
    
        I = -1
        For I = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, I, lngCODIGO))) = 0 Then
                .Row = I
                .Col = lngCODIGO
                .EditCell
                Exit For
            End If
        Next I
        
    End With
End Sub

Private Sub PosCol(grdGenerico As VSFlexGrid, lngCol As Long, lngRow As Long)
    With grdGenerico
        .Col = lngCol
        .Row = lngRow
        .EditCell
    End With

End Sub


Private Function CalcIVA_AJUSTADO(lngRow As Long) As Double
    
    CalcIVA_AJUSTADO = 0
    
    Dim curIVAORIG       As Double
    Dim curALIQINT       As Double
    Dim curALIQICMS      As Double
    
    Dim curIVAAJUST1     As Double
    Dim curIVAAJUST2     As Double
    Dim curIVAAJUST3     As Double
    Dim curIVAAJUST4     As Double
    
    
    With grdNATOPER
    
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_SonNatOper_AliqIVASTORIG))) > 0 Then curIVAORIG = .Cell(flexcpText, lngRow, conCOL_SonNatOper_AliqIVASTORIG)
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_SonNatOper_AliqICMSINT))) > 0 Then curALIQINT = .Cell(flexcpText, lngRow, conCOL_SonNatOper_AliqICMSINT)
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_SonNatOper_ALiqICMS))) > 0 Then curALIQICMS = .Cell(flexcpText, lngRow, conCOL_SonNatOper_ALiqICMS)
        
        curIVAORIG = (1 + (curIVAORIG / 100))
        curALIQINT = (1 - (curALIQINT / 100))
        curALIQICMS = (1 - (curALIQICMS / 100))
        
        If curIVAORIG > 1 Or curALIQINT > 1 Or curALIQICMS > 1 Then
            curIVAAJUST1 = (curALIQICMS / curALIQINT)
            curIVAAJUST2 = (curIVAORIG * curIVAAJUST1)
            curIVAAJUST3 = (curIVAAJUST2 - 1)
            curIVAAJUST4 = (curIVAAJUST3 * 100)
        End If
        
    End With
    
    CalcIVA_AJUSTADO = curIVAAJUST4
    
End Function

Private Sub PopCampoProtocolo(lnrRow As Long)
    
    If lnrRow = 0 Then Exit Sub
    
    With grdNATOPER
        txtProtocolo.Text = Trim(Replace(.Cell(flexcpText, lnrRow, conCOL_SonNatOper_Protocolo), "'", ""))
        txtProtocoloSimp.Text = Trim(Replace(.Cell(flexcpText, lnrRow, conCOL_SonNatOper_ProtOptSimp), "'", ""))
    End With
    
End Sub

Private Sub txtProtocolo_Validate(Cancel As Boolean)
    Call PopGrdNatOperProtocolo(grdNATOPER.Row)
End Sub

Private Sub PopGrdNatOperProtocolo(lngRow As Long)

    If (grdNATOPER.Rows - 1) = 0 Then Exit Sub
    
    If lngRow = 0 Then
        MsgBox "Selecione uma linha da gride !!!", vbOKOnly + vbExclamation, "Aviso"
        txtProtocolo.Text = ""
        Exit Sub
    End If
    
    With grdNATOPER
        .Cell(flexcpText, lngRow, conCOL_SonNatOper_Protocolo) = Trim(Replace(txtProtocolo.Text, "'", ""))
        .Cell(flexcpText, lngRow, conCOL_SonNatOper_ProtOptSimp) = Trim(Replace(txtProtocoloSimp.Text, "'", ""))
    End With
    
End Sub


Private Sub txtProtocoloSimp_Validate(Cancel As Boolean)
    Call PopGrdNatOperProtocolo(grdNATOPER.Row)
End Sub


Private Function CalcIVA_AJUSTADO_SIMPLES(lngRow As Long) As Double
    
    CalcIVA_AJUSTADO_SIMPLES = 0
    
    Dim curIVAORIG       As Double
    Dim curALIQINT       As Double
    Dim curALIQICMS      As Double
    
    Dim curIVAAJUST1     As Double
    Dim curIVAAJUST2     As Double
    Dim curIVAAJUST3     As Double
    Dim curIVAAJUST4     As Double
    
    
    With grdNATOPER
    
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_SonNatOper_AliqIVASTORIGSIMPL))) > 0 Then curIVAORIG = .Cell(flexcpText, lngRow, conCOL_SonNatOper_AliqIVASTORIGSIMPL)
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_SonNatOper_AliqICMSINTSIMPL))) > 0 Then curALIQINT = .Cell(flexcpText, lngRow, conCOL_SonNatOper_AliqICMSINTSIMPL)
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_SonNatOper_ALiqICMSSIMPL))) > 0 Then curALIQICMS = .Cell(flexcpText, lngRow, conCOL_SonNatOper_ALiqICMSSIMPL)
        
        curIVAORIG = (1 + (curIVAORIG / 100))
        curALIQINT = (1 - (curALIQINT / 100))
        curALIQICMS = (1 - (curALIQICMS / 100))
        
        If curIVAORIG > 1 Or curALIQINT > 1 Or curALIQICMS > 1 Then
            curIVAAJUST1 = (curALIQICMS / curALIQINT)
            curIVAAJUST2 = (curIVAORIG * curIVAAJUST1)
            curIVAAJUST3 = (curIVAAJUST2 - 1)
            curIVAAJUST4 = (curIVAAJUST3 * 100)
        End If
        
    End With
    
    CalcIVA_AJUSTADO_SIMPLES = curIVAAJUST4
    
End Function

