VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADMOVPCP 
   Caption         =   "Programação de Produção"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   19545
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "[ Estoque ]"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   37
      Top             =   7440
      Width           =   14775
      Begin VSFlex8LCtl.VSFlexGrid grdEstoque 
         Height          =   1695
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   14535
         _cx             =   25638
         _cy             =   2990
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
      Caption         =   "[ Filtro ]"
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
      Height          =   1095
      Left            =   8880
      TabIndex        =   34
      Top             =   960
      Width           =   10575
      Begin VB.ListBox lstStatus 
         Appearance      =   0  'Flat
         Height          =   705
         ItemData        =   "frmCADMOVPCP.frx":0000
         Left            =   840
         List            =   "frmCADMOVPCP.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   35
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Linha"
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
         Index           =   11
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ OP ]"
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
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   18975
      Begin MSMask.MaskEdBox mskDTENTREGA 
         Height          =   285
         Left            =   10800
         TabIndex        =   4
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtQTDOPPROG 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   15960
         TabIndex        =   5
         Text            =   "txtQTDOPPROG"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtCODOP 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "txtCODOP"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblSTATUS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSTATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   17880
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblQTDOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblQTDOP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   14160
         TabIndex        =   32
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblCOMP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCOMP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   13560
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblFECH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFECH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   12960
         TabIndex        =   30
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblNECK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNECK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   12360
         TabIndex        =   29
         Top             =   480
         Width           =   495
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   17880
         TabIndex        =   28
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.OP Programada"
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
         Left            =   15960
         TabIndex        =   27
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Qtde.OP"
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
         Left            =   14160
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "COMP"
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
         Left            =   13560
         TabIndex        =   25
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FECH"
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
         Left            =   12960
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NECK"
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
         Left            =   12360
         TabIndex        =   23
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data de Entrega"
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
         Left            =   10800
         TabIndex        =   22
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição do Rótulo"
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
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rótulo"
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
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblDESCROT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDESCROT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Top             =   480
         Width           =   7455
      End
      Begin VB.Label lblCODROT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODROT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código OP"
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
         TabIndex        =   17
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   19200
      Picture         =   "frmCADMOVPCP.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   300
   End
   Begin VB.CommandButton Command9 
      Height          =   300
      Left            =   19200
      Picture         =   "frmCADMOVPCP.frx":014E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Inclui uma nova linha na Gride"
      Top             =   2160
      Width           =   300
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "[ Dados para Filtragem ]"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   8655
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   4320
         TabIndex        =   1
         Text            =   "cboAno"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmCADMOVPCP.frx":0298
         Left            =   1920
         List            =   "frmCADMOVPCP.frx":029A
         TabIndex        =   0
         Text            =   "cboMes"
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdCARREGA 
         Caption         =   "&Carrega Gride - <F5>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   2
         Top             =   240
         Width           =   2295
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
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   19335
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
         Left            =   18480
         Picture         =   "frmCADMOVPCP.frx":029C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprime Registro"
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
         Picture         =   "frmCADMOVPCP.frx":039E
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmCADMOVPCP.frx":04A0
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmCADMOVPCP.frx":05A2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid grdPROGOP 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   19335
      _cx             =   34105
      _cy             =   7646
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
Attribute VB_Name = "frmCADMOVPCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCodVendedor   As Long
Public lngCodUsuario    As Long
Public intFILIALPED     As Integer


Dim objBLBFunc          As Object
Dim objCADMOVPCP        As Object
Dim objPESQPADRAO       As Object
Dim objRel              As Object

Dim strNOMEMPRESA       As String
Dim strModulo           As String
Dim strNOMTABELA        As String

Dim arrPROGRAMADO()     As GRPLinha
Dim arrPROGROP          As Variant
Dim arrPROGROPDEL       As Variant
Dim arrMUDASTATUSOP     As Variant
Dim arrGERAFOLHAS       As Variant
Dim arrGRPLINAS()       As GRPLinha
Dim arrGRPLINHAESP()    As GRPLinha
Dim arrGRPFRACIONADA()  As GRPLinha
Dim arrGRPLINHADEL()    As GRPLinha

Dim arrLINHAS()         As LINHAS
Dim arrDIASLINHA()      As DIAS_LINHAS
Dim arrOPSINCUSAS()     As OPS_INCLUSAS
Dim arrOPSPROG()        As OPS_INCLUSAS
Dim arrFOLHASUSADAS()   As OPS_INCLUSAS_FOLHAS_USADAS

Dim strCAPTION          As String
Dim boolTEMDADOS        As Boolean
Dim lngLINHASEL         As Long
Dim lngULTLINSEL        As Long
Dim intFILTROSEL        As Long
Dim lngLINHASELEST      As Long

'' -----------------------------------------------------------------------------------
Const conCOL_PRODOP_DESCLINHA                       As Integer = 0
Const conCOL_PRODOP_Data                            As Integer = 1
Const conCOL_PRODOP_Capacidade                      As Integer = 2
Const conCOL_PRODOP_TotalProgramado                 As Integer = 3
Const conCOL_PRODOP_Disponivel                      As Integer = 4
Const conCOL_PRODOP_TotalOPS                        As Integer = 5
Const conCOL_PRODOP_Programado                      As Integer = 6
Const conCOL_PRODOP_CodOP                           As Integer = 7
Const conCOL_PRODOP_DataEntrega                     As Integer = 8
Const conCOL_PRODOP_CodRotulo                       As Integer = 9
Const conCOL_PRODOP_DescRotulo                      As Integer = 10
Const conCOL_PRODOP_NECK                            As Integer = 11
Const conCOL_PRODOP_FECH                            As Integer = 12
Const conCOL_PRODOP_COMP                            As Integer = 13
Const conCOL_PRODOP_QtdeOP                          As Integer = 14
Const conCOL_PRODOP_QtdeOPProgramada                As Integer = 15
Const conCOL_PRODOP_QtdeReal                        As Integer = 16
Const conCOL_PRODOP_StatusLito                      As Integer = 17
Const conCOL_PRODOP_Status                          As Integer = 18
Const conCOL_PRODOP_CodGRPLINHA                     As Integer = 19
Const conCOL_PRODOP_CodOPBKP                        As Integer = 20
Const conCOL_PRODOP_BLOCOdeOPS                      As Integer = 21
Const conCOL_PRODOP_CODLINHA                        As Integer = 22
Const conCOL_PRODOP_IDOP                            As Integer = 23
Const conCOL_PRODOP_CODPED                          As Integer = 24
Const conCOL_PRODOP_IDPRODUTO                       As Integer = 25
Const conCOL_PRODOP_CODSTATUS                       As Integer = 26
Const conCOL_PRODOP_DTENTREGAORIG                   As Integer = 27
Const conCOL_PRODOP_CODINTERNO                      As Integer = 28
Const conCOL_PRODOP_QTDREALORIG                     As Integer = 29
Const conCOL_PRODOP_Action2Do                       As Integer = 30
Const conCOL_PRODOP_CODSTATUSORIG                   As Integer = 31
Const conCOL_PRODOP_CODSTATUSAPONT                  As Integer = 32
Const conCOL_PRODOP_STATUSAPONT                     As Integer = 33
Const conCOL_PRODOP_TIPO                            As Integer = 34
Const conCOL_PRODOP_INDCEARRAYLINHA                 As Integer = 35
Const conCOL_PRODOP_INDCEARRAYDIA                   As Integer = 36
Const conCOL_PRODOP_INDCEARRAYOP                    As Integer = 37
Const conCOL_PRODOP_FRACIONADA                      As Integer = 38
Const conCOL_PRODOP_IDLINHA                         As Integer = 39
Const conCOL_PRODOP_LINHAEXP                        As Integer = 40
Const conCOL_PRODOP_INDICE                          As Integer = 41
Const conCOL_PRODOP_FormatString                    As String = "=Linha|Data|Capacidade|Total Progr.|Disponivel|Total OP's|Programado|Cód.OP|Dat.Entrega|Rótulo|Descrição do Produto|NECK|FECH|COMP|Qtde.OP|Qtde.OP Progr.|Qtde.Real|Status Lito|Status OP|CodGRPLINHA|CODOPBKP|BLOCOdeOPS|CODLINHA|IDOP|CODPED|IDPRODUTO|CODSTATUS|DTENTREGAORIG|CODINTERNO|QTDREALORIG|Action2Do|CODSTATUSORIG|CODSTATUSAPONT|Status Apont.|TIPO|IDARRAYLINHA|IDARRAYDIA|IDARRAYOP|FRACIONADA|IDLINHA|LINEXP|INDICE"
Const conColumnsIn_PRODOP                           As Integer = 42

Const conCOL_PRODEST_FOLHAUSADA                     As Integer = 0
Const conCOL_PRODEST_IDPROD                         As Integer = 1
Const conCOL_PRODEST_CODPROD                        As Integer = 2
Const conCOL_PRODEST_CODCAPAC                       As Integer = 3
Const conCOL_PRODEST_CAPAC                          As Integer = 4
Const conCOL_PRODEST_CODFOLHAUSADA                  As Integer = 5
Const conCOL_PRODEST_DESCFOLHAUSADA                 As Integer = 6
Const conCOL_PRODEST_ESPESS                         As Integer = 7
Const conCOL_PRODEST_LARG                           As Integer = 8
Const conCOL_PRODEST_COMP                           As Integer = 9
Const conCOL_PRODEST_QTDECORP                       As Integer = 10
Const conCOL_PRODEST_PERDPRODC                      As Integer = 11
Const conCOL_PRODEST_QTDEFOLHAS                     As Integer = 12
Const conCOL_PRODEST_PESO                           As Integer = 13
Const conCOL_PRODEST_QTDELATAS                      As Integer = 14
Const conCOL_PRODEST_CODOP                          As Integer = 15
Const conCOL_PRODEST_INDICE                         As Integer = 16
Const conCOL_PRODEST_NEFOLHAS                       As Integer = 17
Const conCOL_PRODEST_LINHA                          As Integer = 18
Const conCOL_PRODEST_FormatString                   As String = "=   |IDPROD|CODPROD|CODCAPAC|CAPAC|CODFOLHAUSADA|Descrição da Folha|EXPESSURA|LARGURA|COMPRIMENTO|QTDECORPOS|PERDAPROCESSO|Qtde. de Folhas|PESO|Qtde. de Latas|CODOP|INDICE|Necessidade em Folhas|LINHA"
Const conColumnsIn_PRODEST                          As Integer = 19

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub cmdCARREGA_Click()
    '' Carregando das Linhas com as OP's
    Call CarregaLinha
End Sub

Private Sub cmdImpressao_Click()
    Call Imprime
End Sub

Private Sub CmdSalva_Click()

    If Valida_Campos = False Then Exit Sub
    
    Dim I               As Integer
    Dim J               As Long
    Dim K               As Long
    Dim L               As Long
    Dim lngQTDRES       As Long
    Dim lngQTDREGSFOL   As Long
    
    
    If cTipOper = "I" Then objCADMOVPCP.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMTABELA, FILIAL, Linha)
    
    '' ===============================
    '' Carregando a Programação
    lngQTDRES = 0
    lngQTDREGSFOL = 0
    For I = 1 To UBound(arrLINHAS)
    
        If arrLINHAS(I).lngQTDLINHAS > 0 Then
            For J = 1 To arrLINHAS(I).lngQTDLINHAS
                If arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS > 0 Then
                    For K = 1 To arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS
                        If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngAction2Do <> dacEnumUpdateAction_delete Then
                        
                            lngQTDRES = (lngQTDRES + 1)
                            ReDim Preserve arrPROGRAMADO(1 To lngQTDRES) As GRPLinha
                        
                            arrPROGRAMADO(lngQTDRES).dtDATA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).dtDATAPROG
                            arrPROGRAMADO(lngQTDRES).lngCODOP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP
                            arrPROGRAMADO(lngQTDRES).strCODLINHA = Str(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODLINA)
                            arrPROGRAMADO(lngQTDRES).lngCodGRPLinha = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCIDGRPLIN
                            
                            arrPROGRAMADO(lngQTDRES).dtDATAENTREGA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).dtDATAENTREGA
                            arrPROGRAMADO(lngQTDRES).dtDATAENTREGAORIG = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).dtDATAENTREGAORIGINAL
                            arrPROGRAMADO(lngQTDRES).lngQTDREALPROOGOP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOPPROGRAMADA
                            
                            arrPROGRAMADO(lngQTDRES).lngSTATUS = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATUS
                            arrPROGRAMADO(lngQTDRES).lngSTATUSORIG = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATUSORIGINAL
                            arrPROGRAMADO(lngQTDRES).lngIDPRODUTO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDPRODUTO
                            arrPROGRAMADO(lngQTDRES).lngIDPAI = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDOP
                            arrPROGRAMADO(lngQTDRES).lngCODPED = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODPED
                            arrPROGRAMADO(lngQTDRES).lngAction2Do = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngAction2Do
                            arrPROGRAMADO(lngQTDRES).intFRACIONADO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).intFRACIONADA
                            arrPROGRAMADO(lngQTDRES).lngID_LINHA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngID_LINHA
                            
                            arrPROGRAMADO(lngQTDRES).lngCODINTERNO = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMTABELA & "_ID", FILIAL, Linha)
                            
                            '' =====================================
                            '' Pegando as Folhas selecionadas
                            For L = 1 To arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDFOLHASUSADAS
                                If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngFOLHAUSADA = 1 Then
                                    lngQTDREGSFOL = (lngQTDREGSFOL + 1)
                                    ReDim Preserve arrFOLHASUSADAS(1 To lngQTDREGSFOL) As OPS_INCLUSAS_FOLHAS_USADAS
                                    
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngFOLHAUSADA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngFOLHAUSADA
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngCODOP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngCODOP
                                    arrFOLHASUSADAS(lngQTDREGSFOL).strCODPROD = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).strCODPROD
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngIDLIN = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngIDLIN
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngCODLIN = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngCODLIN
                                    arrFOLHASUSADAS(lngQTDREGSFOL).strDESCFOLHAUSADA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).strDESCFOLHAUSADA
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngCODFOLHAUSADA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngCODFOLHAUSADA
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngIDPROD = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngIDPROD
                                    arrFOLHASUSADAS(lngQTDREGSFOL).strINDICE = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).strINDICE
                                    arrFOLHASUSADAS(lngQTDREGSFOL).dblESPESS = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblESPESS
                                    arrFOLHASUSADAS(lngQTDREGSFOL).dblLARG = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblLARG
                                    arrFOLHASUSADAS(lngQTDREGSFOL).dblCOMP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblCOMP
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngQTDECORP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngQTDECORP
                                    arrFOLHASUSADAS(lngQTDREGSFOL).dblPERDPRODC = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblPERDPRODC
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngNECEFOLHAS = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngNECEFOLHAS
                                    arrFOLHASUSADAS(lngQTDREGSFOL).dblPESO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblPESO
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngQTDEFOLHAS = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngQTDEFOLHAS
                                    arrFOLHASUSADAS(lngQTDREGSFOL).lngQTDELATAS = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngQTDELATAS
                                End If
                            Next L
                            '' =====================================
                        
                        End If
                    Next K
                End If
            Next J
        End If
        
    Next I
        
    arrPROGROP = Empty
    If lngQTDRES > 0 Then
        ReDim arrPROGROP(1 To UBound(arrPROGRAMADO), 1 To 16) As String
        For I = 1 To UBound(arrPROGRAMADO)
            arrPROGROP(I, 1) = "'" & Format(arrPROGRAMADO(I).dtDATA, "MM/DD/YYYY") & "'"
            arrPROGROP(I, 2) = arrPROGRAMADO(I).lngCODOP
            arrPROGROP(I, 3) = "'" & Format(arrPROGRAMADO(I).dtDATAENTREGA, "MM/DD/YYYY") & "'"
            arrPROGROP(I, 4) = arrPROGRAMADO(I).lngSTATUS
            arrPROGROP(I, 5) = arrPROGRAMADO(I).lngIDPRODUTO
            arrPROGROP(I, 6) = arrPROGRAMADO(I).lngIDPAI
            arrPROGROP(I, 7) = arrPROGRAMADO(I).lngCODPED
            arrPROGROP(I, 8) = arrPROGRAMADO(I).lngAction2Do
            arrPROGROP(I, 9) = arrPROGRAMADO(I).lngQTDREALPROOGOP
            arrPROGROP(I, 10) = arrPROGRAMADO(I).lngCODINTERNO
            arrPROGROP(I, 11) = "'" & Format(arrPROGRAMADO(I).dtDATAENTREGAORIG, "MM/DD/YYYY") & "'"
            arrPROGROP(I, 12) = arrPROGRAMADO(I).lngSTATUSORIG
            arrPROGROP(I, 13) = arrPROGRAMADO(I).intFRACIONADO
            arrPROGROP(I, 14) = arrPROGRAMADO(I).lngID_LINHA
            arrPROGROP(I, 15) = arrPROGRAMADO(I).strCODLINHA
            arrPROGROP(I, 16) = arrPROGRAMADO(I).lngCodGRPLinha
        Next I
    End If
    objCADMOVPCP.PROGRAMADO = arrPROGROP
    '' ===============================
    
    '' ===============================
    '' Carregando as Folhas Usadas Na Programação
    arrGERAFOLHAS = Empty
    If lngQTDREGSFOL > 0 Then
        ReDim arrGERAFOLHAS(1 To UBound(arrFOLHASUSADAS), 1 To 4) As String
        For I = 1 To UBound(arrFOLHASUSADAS)
            arrGERAFOLHAS(I, 1) = arrFOLHASUSADAS(I).lngIDPROD
            arrGERAFOLHAS(I, 2) = arrFOLHASUSADAS(I).lngCODFOLHAUSADA
            arrGERAFOLHAS(I, 3) = arrFOLHASUSADAS(I).lngCODOP
            arrGERAFOLHAS(I, 4) = "'" & Trim(arrFOLHASUSADAS(I).strINDICE) & "'"
        Next I
    End If
    objCADMOVPCP.GERAFOLHAS = arrGERAFOLHAS
    '' ===============================
    
    If objCADMOVPCP.GRAVA(cTipOper, strModulo) = False Then Exit Sub

    MsgBox "A Programação [ " & objCADMOVPCP.CODIGO & " ] foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    boolTEMDADOS = False
    
    If cTipOper = "I" Then cTipOper = "C"
    
    Call IniciaForm
    
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    
    If cTipOper = "C" Then Exit Sub
    
    With grdPROGOP
        
        If (.Rows - 1) = 0 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "A programação não foi carregado !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        ElseIf .Row <= 0 Or _
               lngLINHASEL <= 0 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Selecione um Registro !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        Dim lngLINHA        As Long
        Dim lngTOTLINHAS    As Long
        Dim I               As Long
        Dim lngLINHAATU     As Long
        
        If cTipOper <> "C" Then
            If Len(Trim(.Cell(flexcpText, .Row, conCOL_PRODOP_CODSTATUSAPONT))) > 0 Then
                If .Cell(flexcpText, .Row, conCOL_PRODOP_CODSTATUSAPONT) = 1 Then
                    MsgBox "ATENÇÂO" & vbCrLf & "Não é possivel excluir esta OP da Programação ela já foi concluida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Sub
                ElseIf .Cell(flexcpText, .Row, conCOL_PRODOP_CODSTATUSAPONT) = 2 Then
                    MsgBox "ATENÇÂO" & vbCrLf & "Não é possivel excluir esta OP da Programação ela já foi concluida Parcialmente !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Sub
                ElseIf .Cell(flexcpText, .Row, conCOL_PRODOP_CODSTATUSAPONT) = 3 Then
                    MsgBox "ATENÇÂO" & vbCrLf & "Não é possivel excluir esta OP da Programação , OP em Processo !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Sub
                End If
            End If
        End If
        
        If Len(Trim(.Cell(flexcpText, .Row, conCOL_PRODOP_Action2Do))) = 0 Then Exit Sub
        
        If (cTipOper = "A" And .Cell(flexcpText, .Row, conCOL_PRODOP_Action2Do) = dacEnumUpdateAction_Ignore) Then
            Call DeletaLinhaDaArray(.Cell(flexcpText, .Row, conCOL_PRODOP_CodOP))
        ElseIf (cTipOper = "A" And .Cell(flexcpText, .Row, conCOL_PRODOP_Action2Do) = dacEnumUpdateAction_Insert) Then
            Call DeletaLinhaDaArray(.Cell(flexcpText, .Row, conCOL_PRODOP_CodOP))
        ElseIf (cTipOper = "I" And .Cell(flexcpText, .Row, conCOL_PRODOP_Action2Do) = dacEnumUpdateAction_Insert) Then
            Call DeletaLinhaDaArray(.Cell(flexcpText, .Row, conCOL_PRODOP_CodOP))
        End If
        
        Call ConfGridProgOP
        Call CarregaLinhasDoArray
        Call lstStatus_Click

        Call LimpaCamposLabel
        txtCODOP.Text = Empty
        mskDTENTREGA.Text = "__/__/____"
        txtQTDOPPROG.Text = Empty
        
        
    End With
End Sub

Private Sub Command9_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADMOVPCP = Nothing
    Set objPESQPADRAO = Nothing
    Set objRel = Nothing
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADMOVPCP = CreateObject("CADMOVPCP.clsCADMOVPCP")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
   
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strNOMTABELA = ""
    If intFILIALPED = 1 Then strNOMTABELA = "_STEEL"
    
    strModulo = ""
    If intFILIALPED = 1 Then strModulo = "_STEEL"
    
    strNOMEMPRESA = "NOVALATA"
    If intFILIALPED = 1 Then strNOMEMPRESA = "STEEL ROL"
    
    strCAPTION = "Programação de Produção /" & strNOMEMPRESA & " - "
    
    objCADMOVPCP.FILIAL = FILIAL
    objCADMOVPCP.CODUSUARIO = lngCodUsuario
    
    Call IniciaForm
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Destroy_Objeto
End Sub

Private Sub grdEstoque_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    With grdEstoque
        
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        
        Select Case Col
               Case conCOL_PRODEST_FOLHAUSADA
               
                    If (.Rows - 1) <= 0 Then Exit Sub
                    If (Row) = 0 Then Exit Sub
                    
                    If lngLINHASEL <= 0 Then
                        MsgBox "ATENÇÂO" & vbCrLf & _
                               "Selecione uma OP !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Sub
                    End If
                    
                    If lngLINHASELEST <= 0 Then
                        MsgBox "ATENÇÂO" & vbCrLf & _
                               "Selecione uma registro !!!", vbOKOnly + vbExclamation, "Aviso"
                        Exit Sub
                    End If
                        
                    Dim lngLINHA_LINHA  As Long
                    Dim lngLINHA_DIA    As Long
                    Dim lngLINHA_OP     As Long
                    Dim lngLINHA_FOLHA  As Long
                    
                    lngLINHA_LINHA = grdPROGOP.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_INDCEARRAYLINHA)
                    lngLINHA_DIA = grdPROGOP.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_INDCEARRAYDIA)
                    lngLINHA_OP = grdPROGOP.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_INDCEARRAYOP)
                    lngLINHA_FOLHA = .Cell(flexcpText, lngLINHASELEST, conCOL_PRODEST_LINHA)
                    
                    If .Cell(flexcpChecked, lngLINHASELEST, conCOL_PRODEST_FOLHAUSADA) = 2 Then
                         arrLINHAS(lngLINHA_LINHA).arrDIAS_LINHA(lngLINHA_DIA).arrOPS_INCLUSAS(lngLINHA_OP).arrFOLHAS_USADAS(lngLINHA_FOLHA).lngFOLHAUSADA = 2
                    ElseIf .Cell(flexcpChecked, lngLINHASELEST, conCOL_PRODEST_FOLHAUSADA) = 1 Then
                         arrLINHAS(lngLINHA_LINHA).arrDIAS_LINHA(lngLINHA_DIA).arrOPS_INCLUSAS(lngLINHA_OP).arrFOLHAS_USADAS(lngLINHA_FOLHA).lngFOLHAUSADA = 1
                    End If
                        
               
        End Select
    
    End With

End Sub

Private Sub grdEstoque_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With grdEstoque
        Select Case Col
               Case conCOL_PRODEST_INDICE, _
                    conCOL_PRODEST_NEFOLHAS, _
                    conCOL_PRODEST_LINHA
                    Cancel = True
               Case conCOL_PRODEST_FOLHAUSADA
                    If cTipOper = "C" Then
                        Cancel = True
                    Else
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_PRODEST_QTDELATAS))) = 0 Then Cancel = True
                    End If
               Case Else
                    Cancel = True
                   .ComboList = ""
               End Select
    End With
    Exit Sub

End Sub

Private Sub grdEstoque_Click()
    Call PegaLinhaEstoque
End Sub

Private Sub grdEstoque_RowColChange()
    Call PegaLinhaEstoque
End Sub

Private Sub grdPROGOP_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    With grdPROGOP
        
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        
        Select Case Col
               Case conCOL_PRODOP_DESCLINHA
               Case conCOL_PRODOP_Programado
               Case conCOL_PRODOP_QtdeOPProgramada
                    Call CalcTotDiaProgramado(.Cell(flexcpText, Row, conCOL_PRODOP_BLOCOdeOPS))
        End Select
    
    End With

End Sub

Private Sub grdPROGOP_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    
    Dim lngPESQ As Long
    
    With grdPROGOP
        Select Case Col
               Case conCOL_PRODOP_DESCLINHA, conCOL_PRODOP_Data, _
                    conCOL_PRODOP_Capacidade, conCOL_PRODOP_TotalProgramado, _
                    conCOL_PRODOP_Disponivel, conCOL_PRODOP_TotalOPS, _
                    conCOL_PRODOP_CodOP, conCOL_PRODOP_CodRotulo, _
                    conCOL_PRODOP_DescRotulo, conCOL_PRODOP_NECK, _
                    conCOL_PRODOP_FECH, conCOL_PRODOP_COMP, _
                    conCOL_PRODOP_QtdeOP, conCOL_PRODOP_Status, _
                    conCOL_PRODOP_CodGRPLINHA, conCOL_PRODOP_CodOPBKP, _
                    conCOL_PRODOP_BLOCOdeOPS, _
                    conCOL_PRODOP_CODLINHA, conCOL_PRODOP_IDOP, _
                    conCOL_PRODOP_CODPED, conCOL_PRODOP_IDPRODUTO, _
                    conCOL_PRODOP_CODSTATUS, conCOL_PRODOP_DTENTREGAORIG, _
                    conCOL_PRODOP_CODINTERNO, conCOL_PRODOP_Action2Do, _
                    conCOL_PRODOP_QTDREALORIG, conCOL_PRODOP_Programado, _
                    conCOL_PRODOP_CODSTATUSORIG, conCOL_PRODOP_QtdeReal, _
                    conCOL_PRODOP_CODSTATUSAPONT, conCOL_PRODOP_STATUSAPONT, _
                    conCOL_PRODOP_INDICE
                    Cancel = True
               Case conCOL_PRODOP_QtdeOPProgramada, _
                    conCOL_PRODOP_StatusLito, _
                    conCOL_PRODOP_DataEntrega
                    If .Row <= 3 Then
                        Cancel = True
                    Else
                        If .Cell(flexcpText, Row, conCOL_PRODOP_CodOP) = 0 Or _
                           Len(Trim(.Cell(flexcpText, Row, conCOL_PRODOP_CodOP))) = 0 Then
                            Cancel = True
                        Else
                            If cTipOper = "C" Then
                                Cancel = True
                            Else
                                If Col <> conCOL_PRODOP_Programado Then
                                    If .Cell(flexcpChecked, Row, conCOL_PRODOP_Programado) = 2 Then Cancel = True
                                ElseIf Col = conCOL_PRODOP_Programado Then
                                    lngPESQ = .FindRow(.Cell(flexcpText, Row, conCOL_PRODOP_BLOCOdeOPS), , conCOL_PRODOP_BLOCOdeOPS)
                                    If lngPESQ > -1 Then
                                        If CLng(.Cell(flexcpText, lngPESQ, conCOL_PRODOP_Capacidade)) = 0 Then Cancel = True
                                    End If
                                End If
                            End If
                        End If
                    End If
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub

End Sub

Private Sub grdPROGOP_Click()
    Call PegaLinhaAtual
End Sub


Private Sub grdPROGOP_DblClick()
    With grdPROGOP
        
        If (.Rows - 1) = 0 Then Exit Sub
        If (.Row) <= 0 Then Exit Sub
        
        Dim boolEXP     As Boolean
        
        If .GetNode(.Row).Expanded = True Then
            .GetNode(.Row).Expanded = False
        ElseIf .GetNode(.Row).Expanded = False Then
            .GetNode(.Row).Expanded = True
        End If
        
        boolEXP = .GetNode(.Row).Expanded
        If boolEXP = False Then .Cell(flexcpText, .Row, conCOL_PRODOP_LINHAEXP) = 0
        If boolEXP = True Then .Cell(flexcpText, .Row, conCOL_PRODOP_LINHAEXP) = 1
        
        If Len(Trim(.Cell(flexcpText, .Row, conCOL_PRODOP_INDCEARRAYLINHA))) > 0 Then
            arrLINHAS(.Cell(flexcpText, .Row, conCOL_PRODOP_INDCEARRAYLINHA)).lngEXPLIN = .Cell(flexcpText, .Row, conCOL_PRODOP_LINHAEXP)
            arrLINHAS(.Cell(flexcpText, .Row, conCOL_PRODOP_INDCEARRAYLINHA)).arrDIAS_LINHA(.Cell(flexcpText, .Row, conCOL_PRODOP_INDCEARRAYDIA)).lngEXPLIN = .Cell(flexcpText, .Row, conCOL_PRODOP_LINHAEXP)
            arrLINHAS(.Cell(flexcpText, .Row, conCOL_PRODOP_INDCEARRAYLINHA)).arrDIAS_LINHA(.Cell(flexcpText, .Row, conCOL_PRODOP_INDCEARRAYDIA)).arrOPS_INCLUSAS(.Cell(flexcpText, .Row, conCOL_PRODOP_INDCEARRAYOP)).lngEXPLIN = .Cell(flexcpText, .Row, conCOL_PRODOP_LINHAEXP)
        End If
        
    End With
End Sub

Private Sub grdPROGOP_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdPROGOP
          Select Case Col
                    Case conCOL_PRODOP_QtdeOP, _
                         conCOL_PRODOP_QtdeReal
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_PRODOP_StatusLito
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_PRODOP_DataEntrega
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          End Select
     End With
End Sub


Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, StrCodigo As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(StrCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(StrCodigo)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub





Private Function Valida_Campos() As Boolean

     Valida_Campos = False
     
     With grdPROGOP
        If (.Rows - 1) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "A programação não foi carregada !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
        End If
     End With
     
     Valida_Campos = True

End Function


Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    If cTipOper = "I" Then
        iCodigo = 0
    Else
        If iCodigo = 0 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Não foi selecionado nenhum registro para Consulta ou Alterar !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    objCADMOVPCP.CODIGO = iCodigo
    
    Call ConfGridProgOP
    Call ConfGrd
    
    Call objBLBFunc.Preenche_Mes(cboMes)
    cboMes.ListIndex = (Month(Date) - 1)
    
    Call objBLBFunc.Preenche_Ano(cboAno)
    cboAno.ListIndex = 0
    
    Call LimpaCamposLabel
    Call CarregaCampos
    
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        fraPeriodo.Enabled = True
        Frame2.Enabled = True
        cmdCARREGA.Enabled = True
        If strTipOper = "A" Then
            cboMes.Enabled = False
            cboAno.Enabled = False
            cmdCARREGA.Enabled = False
        End If
    ElseIf strTipOper = "C" Then
        fraPeriodo.Enabled = False
        Frame2.Enabled = False
        cboMes.Enabled = False
        cboAno.Enabled = False
        cmdCARREGA.Enabled = False
    End If
End Sub

Private Sub CarregaCampos()

    If objCADMOVPCP.Carrega_Campos(strNOMTABELA) = False Then Exit Sub
    
    Dim I As Integer
    
    If cTipOper = "I" Then
        For I = 0 To (cboMes.ListCount - 1)
            If cboMes.ItemData(I) = Month(Now) Then
                cboMes.ListIndex = I
                Exit For
            End If
        Next I
        
        For I = 0 To (cboAno.ListCount - 1)
            If cboAno.ItemData(I) = Year(Now) Then
                cboAno.ListIndex = I
                Exit For
            End If
        Next I
    ElseIf cTipOper = "C" Or cTipOper = "A" Then
        For I = 0 To (cboMes.ListCount - 1)
            If cboMes.ItemData(I) = Month(CDate(objCADMOVPCP.DTPROGRAMA)) Then
                cboMes.ListIndex = I
                Exit For
            End If
        Next I
        
        For I = 0 To (cboAno.ListCount - 1)
            If cboAno.ItemData(I) = Year(CDate(objCADMOVPCP.DTPROGRAMA)) Then
                cboAno.ListIndex = I
                Exit For
            End If
        Next I
    End If
    
    Call CarregaLinha

End Sub


Private Sub CarregaLinha()

    Call ConfGridProgOP
    Call ConfLstLinhas
    
    If cboMes.ListIndex = -1 Then Exit Sub
    If cboAno.ListIndex = -1 Then Exit Sub
    
    Dim strDTINICIAL        As String
    Dim strDTFINAL          As String

    strDTINICIAL = "'" & Format(CDate("01/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "MM/DD/YYYY") & "'"
    If cboMes.ItemData(cboMes.ListIndex) = 12 Then
        strDTFINAL = "'" & Format(CDate("31/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "MM/DD/YYYY") & "'"
    Else
        strDTFINAL = "'" & Format((CDate("01/" & (cboMes.ItemData(cboMes.ListIndex) + 1) & "/" & cboAno.ItemData(cboAno.ListIndex)) - 1), "MM/DD/YYYY") & "'"
    End If
    
    '' Verificando se Já Existe o Mês Criado
    If cTipOper = "I" Then
        If VerificaSeExisteMes(cboMes.ItemData(cboMes.ListIndex), cboAno.ItemData(cboAno.ListIndex)) = True Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Vocês esta tentando Incluir o Mês " & cboMes.Text & "/" & cboAno.Text & " !!!" & vbCrLf & _
                   "Este Mês já Está Criado Entre com a Opção Alterar !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    '' ------------------------------------
    '' Carregando a Array
    Dim lngLINREG           As Long
    Dim lngLINREG_TOT       As Long
    Dim lngLINREG_DATAS     As Long
    Dim lngQTDOPS           As Long
    Dim lngQTDFOLHAS        As Long
    Dim arrCAMPOS()         As String
    Dim strCAMPOS           As String
    Dim lngPESQCODCORT      As Long
    Dim lngQTDNECFOLHAS     As Long
    
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       LINHA.SGI_CODLIN" & vbCrLf
    sSql = sSql & "     , LINHA.SGI_DESCRI As SGI_DESCRLINHA" & vbCrLf
    sSql = sSql & "     , GRPLINMESANO.SGI_Mes" & vbCrLf
    sSql = sSql & "     , GRPLINMESANO.SGI_ANO" & vbCrLf
    sSql = sSql & "     , GRPLINITE.SGI_CODIGO As SGI_CODGRUPLIN" & vbCrLf
    sSql = sSql & "     , GRPLINITE.SGI_IDINTERNO" & vbCrLf
    sSql = sSql & "     , GRPLINMESANO.SGI_QTDECAPACIDADE" & vbCrLf
    sSql = sSql & "     , Count(*) as SGI_QtdeRegs" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADGRUPLINHAIT" & strNOMTABELA & " GRPLINITE" & vbCrLf
    sSql = sSql & "     , SGI_MAQULIN_MESANO" & strNOMTABELA & " GRPLINMESANO" & vbCrLf
    sSql = sSql & "     , SGI_MAQULIN_CAPAC" & strNOMTABELA & "  GRPCAPACDIA" & vbCrLf
    sSql = sSql & "     , SGI_CADLINHAPRODUTO      LINHA" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       GRPLINITE.SGI_FILIAL        = " & FILIAL & vbCrLf
    
    sSql = sSql & "   And GRPLINMESANO.SGI_FILIAL     = GRPLINITE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPLINMESANO.SGI_CODIGO     = GRPLINITE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And GRPLINMESANO.SGI_IDPAI      = GRPLINITE.SGI_IDINTERNO" & vbCrLf
    sSql = sSql & "   And GRPLINMESANO.SGI_MES        = " & cboMes.ItemData(cboMes.ListIndex) & vbCrLf
    sSql = sSql & "   And GRPLINMESANO.SGI_ANO        = " & cboAno.ItemData(cboAno.ListIndex) & vbCrLf
    
    sSql = sSql & "   And GRPCAPACDIA.SGI_FILIAL      = GRPLINMESANO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPCAPACDIA.SGI_CODIGO      = GRPLINMESANO.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And GRPCAPACDIA.SGI_IDPAI       = GRPLINMESANO.SGI_IDPAI" & vbCrLf
    sSql = sSql & "   And Month(GRPCAPACDIA.SGI_DATA) = GRPLINMESANO.SGI_MES" & vbCrLf
    sSql = sSql & "   And  Year(GRPCAPACDIA.SGI_DATA) = GRPLINMESANO.SGI_ANO" & vbCrLf
    sSql = sSql & "   And GRPCAPACDIA.SGI_ATIVO       = 1" & vbCrLf
    
    sSql = sSql & "   And LINHA.SGI_FILIAL            = GRPCAPACDIA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINHA.SGI_CODLIN            = GRPCAPACDIA.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "Group By LINHA.SGI_CODLIN,LINHA.SGI_DESCRI,GRPLINMESANO.SGI_Mes,GRPLINMESANO.SGI_ANO,GRPLINITE.SGI_CODIGO,GRPLINITE.SGI_IDINTERNO,GRPLINMESANO.SGI_QTDECAPACIDADE" & vbCrLf
    sSql = sSql & "Order By LINHA.SGI_CODLIN,LINHA.SGI_DESCRI,GRPLINMESANO.SGI_Mes,GRPLINMESANO.SGI_ANO,GRPLINITE.SGI_CODIGO,GRPLINITE.SGI_IDINTERNO,GRPLINMESANO.SGI_QTDECAPACIDADE"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
       lngLINREG = 0
       lngLINREG_TOT = 0
       Do While Not BREC.EOF()
       
            lngLINREG = (lngLINREG + 1)
            ReDim Preserve arrLINHAS(1 To lngLINREG) As LINHAS
            
            arrLINHAS(lngLINREG).lngCodLinha = BREC!SGI_CODLIN
            arrLINHAS(lngLINREG).strDESCGRPLINHA = Trim(BREC!SGI_DESCRLINHA)
            arrLINHAS(lngLINREG).lngMES = BREC!SGI_Mes
            arrLINHAS(lngLINREG).lngANO = BREC!SGI_ANO
            arrLINHAS(lngLINREG).lngCODGRPLIN = BREC!SGI_CODGRUPLIN
            arrLINHAS(lngLINREG).lngID_INTERNO = BREC!SGI_IDINTERNO
            arrLINHAS(lngLINREG).lngQTDLINHAS = BREC!SGI_QtdeRegs
            arrLINHAS(lngLINREG).lngQTDECAPACIDADE = BREC!SGI_QTDECAPACIDADE
            arrLINHAS(lngLINREG).lngEXPLIN = 0
                        
            '' ------------------------------------
            '' Carregando as Datas para Programação
            sSql = ""
            sSql = sSql & "Select" & vbCrLf
            sSql = sSql & "       *" & vbCrLf
            sSql = sSql & "  From" & vbCrLf
            sSql = sSql & "        SGI_MAQULIN_CAPAC" & strNOMTABELA & vbCrLf
            sSql = sSql & " Where" & vbCrLf
            sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "    And SGI_IDPAI  = " & BREC!SGI_IDINTERNO & vbCrLf
            sSql = sSql & "    And SGI_ATIVO  = 1" & vbCrLf
            sSql = sSql & "Order BY SGI_DATA"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then
                lngLINREG_DATAS = 0
                Do While Not BREC2.EOF()
                    lngLINREG_DATAS = (lngLINREG_DATAS + 1)
                    lngLINREG_TOT = (lngLINREG_TOT + 1)
                    ReDim Preserve arrLINHAS(lngLINREG).arrDIAS_LINHA(1 To lngLINREG_DATAS) As DIAS_LINHAS
                    
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).lngCodLinha = BREC!SGI_CODLIN
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).lngCODGRPLIN = BREC!SGI_CODGRUPLIN
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).lngID_PAI = BREC!SGI_IDINTERNO
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).dtDATAPROG = BREC2!SGI_DATA
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).lngQTDOPS = 0
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).lngTOTALPECAS = BREC2!SGI_TOTALPECAS
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).lngTOTPROG = 0
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).lngTOTDISP = 0
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).strBLOCOOP = "B" & lngLINREG_TOT
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).strTIPO = "H"
                    arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).lngEXPLIN = 0
                    
                    '' --------------------------------------------
                    '' Pegando a OP
                    sSql = ""
                    sSql = "Select" & vbCrLf
                    sSql = sSql & "       MOVPCP.*" & vbCrLf
                    sSql = sSql & "     , PROD.SGI_CODIGO       As SGI_CODROTULO" & vbCrLf
                    sSql = sSql & "     , PROD.SGI_DESCRICAO    As SGI_DESCOTULO" & vbCrLf
                    sSql = sSql & "     , PROD.SGI_NECKIN" & vbCrLf
                    sSql = sSql & "     , ORDP.SGI_QTDE" & vbCrLf
                    sSql = sSql & "  From" & vbCrLf
                    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & " MOVPCP" & vbCrLf
                    sSql = sSql & "     , SGI_ORDEMPROD" & strNOMTABELA & " ORDP" & vbCrLf
                    sSql = sSql & "     , SGI_CADPRODUTO      PROD" & vbCrLf
                    sSql = sSql & " Where" & vbCrLf
                    sSql = sSql & "       MOVPCP.SGI_FILIAL     = " & FILIAL & vbCrLf
                    sSql = sSql & "   And MOVPCP.SGI_DATAPROG   = '" & Format(BREC2!SGI_DATA, "MM/DD/YYYY") & "'" & vbCrLf
                    sSql = sSql & "   And MOVPCP.SGI_IDLINHA    = " & BREC!SGI_IDINTERNO
                    sSql = sSql & "   And MOVPCP.SGI_CODGRPLIN  = " & BREC!SGI_CODGRUPLIN
                    sSql = sSql & "   And MOVPCP.SGI_CODLIN     = " & BREC!SGI_CODLIN
                    
                    sSql = sSql & "   And ORDP.SGI_FILIAL     = MOVPCP.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And ORDP.SGI_IDPAI      = MOVPCP.SGI_IDINTERNO" & vbCrLf
                    
                    sSql = sSql & "   And PROD.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ORDP.SGI_IDPRODUTO" & vbCrLf
                    sSql = sSql & "   And PROD.SGI_CODLINPROD = " & BREC!SGI_CODLIN
                    
                    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                    If Not BREC3.EOF() Then
                        Do While Not BREC3.EOF()
                        
                            '' Adicionando elemento na Array
                            Call Add_ElArray(lngLINREG _
                                            , lngLINREG_DATAS _
                                            , Trim(Str(BREC3!SGI_CODOP)) _
                                            , BREC!SGI_IDINTERNO _
                                            , BREC3!SGI_DATAPROG _
                                            , BREC!SGI_CODGRUPLIN _
                                            , BREC!SGI_CODLIN _
                                            , arrLINHAS(lngLINREG).arrDIAS_LINHA(lngLINREG_DATAS).strBLOCOOP _
                                            , BREC3!SGI_DATAENTR _
                                            , BREC3!SGI_DATAENTRANT _
                                            , BREC3!SGI_QTDE _
                                            , BREC3!SGI_QTDEPROD _
                                            , True _
                                            , BREC3!SGI_CODINTENO _
                                            , dacEnumUpdateAction_Ignore _
                                            , 0)
                                            
                            
                            ''arrOPSINCUSAS(lngQTDOPS).lngCODSTATAPONT = PegaStatusApontamento(Str(BREC3!SGI_CODOP), Str(BREC3!SGI_IDINTERNO), Str(BREC3!SGI_CODINTENO))
                            
                            '' -------------------
                            '' Pegando a Qtde Real Produzida
                            ''arrOPSINCUSAS(lngQTDOPS).lngQTDOAPONTADAORIGINAL = PegaQtdeRealProduzida(Str(BREC3!SGI_CODOP), Str(BREC3!SGI_IDINTERNO), Str(BREC3!SGI_CODINTENO))
                            '' -------------------
                            
                            BREC3.MoveNext
                        Loop
                    End If
                    BREC3.Close
                    
                    BREC2.MoveNext
                Loop
            End If
            BREC2.Close
            '' ------------------------------------
            
            BREC.MoveNext
       Loop
       
       Call CarregaLinhasDoArray
       If (grdPROGOP.Rows - 1) > 0 Then
            grdPROGOP.Row = 1
            Call PegaLinhaAtual
       End If
    
    Else
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Não existe dados para carregar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close
    
End Sub

Private Function PegaStatus(lngCODSTATUS As Long) As String

    PegaStatus = ""
    
    If lngCODSTATUS = 0 Then PegaStatus = "Liberado"
    If lngCODSTATUS = 1 Then PegaStatus = "Fat.Parcial"
    If lngCODSTATUS = 6 Then PegaStatus = "P.Cota"
    If lngCODSTATUS = 7 Then PegaStatus = "P.Data"

End Function


Private Sub ConfGridProgOP()

    ' reset the control
    Call SetDefaults(grdPROGOP)
    
    With grdPROGOP
        
        .Redraw = True
        
        ' set the properties we want
        .Rows = 1
        .Cols = conColumnsIn_PRODOP
        .FixedRows = 1
        .FormatString = conCOL_PRODOP_FormatString
        
        .AllowUserResizing = flexResizeBoth
        .OutlineBar = flexOutlineBarComplete
        .OutlineCol = 0
        .SubtotalPosition = flexSTAbove
        
        .GridLines = flexGridInsetVert
        .FontName = "Arial"
        .FontSize = 8
        
        
        .Cell(flexcpData, 0, conCOL_PRODOP_Programado) = ""
        .ColDataType(conCOL_PRODOP_Programado) = flexDTBoolean
    
        .Cell(flexcpData, 0, conCOL_PRODOP_QtdeOPProgramada) = ""
        .ColDataType(conCOL_PRODOP_QtdeOPProgramada) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_Action2Do) = ""
        .ColDataType(conCOL_PRODOP_Action2Do) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_CodOPBKP) = ""
        .ColDataType(conCOL_PRODOP_CodOPBKP) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_BLOCOdeOPS) = ""
        .ColDataType(conCOL_PRODOP_BLOCOdeOPS) = flexDTString
    
        .Cell(flexcpData, 0, conCOL_PRODOP_IDOP) = ""
        .ColDataType(conCOL_PRODOP_IDOP) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_CODPED) = ""
        .ColDataType(conCOL_PRODOP_CODPED) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_IDPRODUTO) = ""
        .ColDataType(conCOL_PRODOP_IDPRODUTO) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_CODSTATUS) = ""
        .ColDataType(conCOL_PRODOP_CODSTATUS) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_DTENTREGAORIG) = ""
        .ColDataType(conCOL_PRODOP_DTENTREGAORIG) = flexDTDate
    
        .Cell(flexcpData, 0, conCOL_PRODOP_DataEntrega) = ""
        .ColDataType(conCOL_PRODOP_DataEntrega) = flexDTDate
    
        .Cell(flexcpData, 0, conCOL_PRODOP_CODINTERNO) = ""
        .ColDataType(conCOL_PRODOP_CODINTERNO) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_QTDREALORIG) = ""
        .ColDataType(conCOL_PRODOP_QTDREALORIG) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_CODSTATUSORIG) = ""
        .ColDataType(conCOL_PRODOP_CODSTATUSORIG) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_CODLINHA) = ""
        .ColDataType(conCOL_PRODOP_CODLINHA) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_CODSTATUSAPONT) = ""
        .ColDataType(conCOL_PRODOP_CODSTATUSAPONT) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_STATUSAPONT) = ""
        .ColDataType(conCOL_PRODOP_STATUSAPONT) = flexDTString
    
        .Cell(flexcpData, 0, conCOL_PRODOP_TIPO) = ""
        .ColDataType(conCOL_PRODOP_TIPO) = flexDTString
    
        .Cell(flexcpData, 0, conCOL_PRODOP_INDCEARRAYLINHA) = ""
        .ColDataType(conCOL_PRODOP_INDCEARRAYLINHA) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_INDCEARRAYDIA) = ""
        .ColDataType(conCOL_PRODOP_INDCEARRAYDIA) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_INDCEARRAYOP) = ""
        .ColDataType(conCOL_PRODOP_INDCEARRAYOP) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_FRACIONADA) = ""
        .ColDataType(conCOL_PRODOP_FRACIONADA) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_IDLINHA) = ""
        .ColDataType(conCOL_PRODOP_IDLINHA) = flexDTLong
    
        .Cell(flexcpData, 0, conCOL_PRODOP_LINHAEXP) = ""
        .ColDataType(conCOL_PRODOP_LINHAEXP) = flexDTLong
    
        .ColWidth(conCOL_PRODOP_DESCLINHA) = 2300
        .ColWidth(conCOL_PRODOP_Data) = 1400
        .ColWidth(conCOL_PRODOP_Capacidade) = 1000
        .ColWidth(conCOL_PRODOP_TotalProgramado) = 1000
        .ColWidth(conCOL_PRODOP_Disponivel) = 1000
        .ColWidth(conCOL_PRODOP_TotalOPS) = 800
        .ColWidth(conCOL_PRODOP_Programado) = 0
        .ColWidth(conCOL_PRODOP_CodOP) = 900
        .ColWidth(conCOL_PRODOP_DataEntrega) = 1000
        .ColWidth(conCOL_PRODOP_CodRotulo) = 1200
        .ColWidth(conCOL_PRODOP_DescRotulo) = 4500
        .ColWidth(conCOL_PRODOP_NECK) = 500
        .ColWidth(conCOL_PRODOP_FECH) = 500
        .ColWidth(conCOL_PRODOP_COMP) = 600
        .ColWidth(conCOL_PRODOP_QtdeOP) = 900
        .ColWidth(conCOL_PRODOP_QtdeOPProgramada) = 1300
        .ColWidth(conCOL_PRODOP_QtdeReal) = 900
        .ColWidth(conCOL_PRODOP_StatusLito) = 1500
        .ColWidth(conCOL_PRODOP_Status) = 900
        .ColWidth(conCOL_PRODOP_CodGRPLINHA) = 0
        .ColWidth(conCOL_PRODOP_CodOPBKP) = 0
        .ColWidth(conCOL_PRODOP_BLOCOdeOPS) = 0
        .ColWidth(conCOL_PRODOP_CODLINHA) = 0
        .ColWidth(conCOL_PRODOP_IDOP) = 0
        
        .ColWidth(conCOL_PRODOP_CODPED) = 0
        .ColWidth(conCOL_PRODOP_IDPRODUTO) = 0
        .ColWidth(conCOL_PRODOP_CODSTATUS) = 0
        .ColWidth(conCOL_PRODOP_DTENTREGAORIG) = 0
        .ColWidth(conCOL_PRODOP_CODINTERNO) = 0
        .ColWidth(conCOL_PRODOP_QTDREALORIG) = 0
        .ColWidth(conCOL_PRODOP_Action2Do) = 0
        .ColWidth(conCOL_PRODOP_CODSTATUSORIG) = 0
        .ColWidth(conCOL_PRODOP_CODSTATUSAPONT) = 0
        .ColWidth(conCOL_PRODOP_STATUSAPONT) = 1100
        .ColWidth(conCOL_PRODOP_TIPO) = 0
        .ColWidth(conCOL_PRODOP_INDCEARRAYLINHA) = 0
        .ColWidth(conCOL_PRODOP_INDCEARRAYDIA) = 0
        .ColWidth(conCOL_PRODOP_INDCEARRAYOP) = 0
        .ColWidth(conCOL_PRODOP_FRACIONADA) = 0
        .ColWidth(conCOL_PRODOP_IDLINHA) = 0
        .ColWidth(conCOL_PRODOP_LINHAEXP) = 0
        
    End With
End Sub



Sub SetDefaults(fa As VSFlexGrid)
    
    With fa
        .BindToArray Null
        .Rows = 0
        .Cols = 0
        .ScrollTrack = False
        .ExplorerBar = flexExNone
        .AutoSearch = flexSearchNone
        
        .Editable = flexEDKbdMouse
        
        .AllowUserResizing = flexResizeNone
        .SelectionMode = flexSelectionFree
        .OutlineBar = flexOutlineBarNone
        .OLEDragMode = flexOLEDragManual
        .OLEDropMode = flexOLEDropNone
        .ScrollTips = False
        .ToolTipText = ""
    
    End With
    
End Sub

Private Sub CalcTotDiaProgramado(strDTDIA As String)

    Dim lngLINHADIAINI  As Long
    Dim lngQTDPROG      As Long
    Dim lngTOTPROG      As Long
    Dim lngTOTDISP      As Long
    Dim lngQTDOPS       As Long
    Dim I               As Long
    
    lngQTDPROG = 0
    lngTOTPROG = 0
    lngTOTDISP = 0
    lngQTDOPS = 0
    
    With grdPROGOP
        lngLINHADIAINI = 1
    
        lngLINHADIAINI = .FindRow(strDTDIA, , conCOL_PRODOP_BLOCOdeOPS)
        If lngLINHADIAINI > -1 Then
            '' Total Disponivel
            lngTOTDISP = .Cell(flexcpText, lngLINHADIAINI, conCOL_PRODOP_Capacidade)
            
            For I = 1 To (.Rows - 1)
                If .Cell(flexcpText, I, conCOL_PRODOP_BLOCOdeOPS) = strDTDIA And IsDate(.Cell(flexcpText, I, conCOL_PRODOP_Data)) Then
                    
                    If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOP_CodOP))) > 0 And _
                       .Cell(flexcpText, I, conCOL_PRODOP_Action2Do) <> dacEnumUpdateAction_delete Then lngQTDOPS = (lngQTDOPS + 1)
                    
                    
                    If .Cell(flexcpChecked, I, conCOL_PRODOP_Programado) = 1 And _
                       .Cell(flexcpText, I, conCOL_PRODOP_Action2Do) <> dacEnumUpdateAction_delete Then
                       If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOP_QtdeOPProgramada))) > 0 Then
                            lngQTDPROG = lngQTDPROG + CLng(.Cell(flexcpText, I, conCOL_PRODOP_QtdeOPProgramada))
                       End If
                    End If
                       
                End If
            Next I
        
            lngTOTPROG = (lngTOTPROG + lngQTDPROG)
            lngTOTDISP = (lngTOTDISP - lngTOTPROG)
            
            .Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_TotalProgramado) = lngTOTPROG
            .Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_Disponivel) = lngTOTDISP
            .Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_TotalOPS) = lngQTDOPS
            
            If lngTOTDISP < 0 Then
                .Cell(flexcpBackColor, (lngLINHADIAINI - 1), conCOL_PRODOP_Disponivel) = vbRed
                .Cell(flexcpForeColor, (lngLINHADIAINI - 1), conCOL_PRODOP_Disponivel) = vbBlack
            ElseIf lngTOTDISP >= 0 Then
               .Cell(flexcpBackColor, (lngLINHADIAINI - 1), conCOL_PRODOP_Disponivel) = &HC0C0C0
               .Cell(flexcpForeColor, (lngLINHADIAINI - 1), conCOL_PRODOP_Disponivel) = vbBlack
            End If
            
        End If
    
    End With
    
    

End Sub

Private Sub IncRegGrid()

    If (grdPROGOP.Rows - 1) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "A Programação não foi carregada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If grdPROGOP.Row <= 0 Or lngLINHASEL <= 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Selecione um Dia !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If Not IsNumeric(txtCODOP.Text) Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Informe a OP !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If Not IsDate(mskDTENTREGA.Text) Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Data de entraga da OP inválido !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If Not IsNumeric(txtQTDOPPROG.Text) Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Qtde Programada da OP inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim lngPESQ     As Long
    
    With grdPROGOP
        lngPESQ = .FindRow(txtCODOP.Text, , conCOL_PRODOP_CodOP)
        If lngPESQ > -1 Then
            If .Cell(flexcpText, lngPESQ, conCOL_PRODOP_Action2Do) <> dacEnumUpdateAction_delete Then
                 MsgBox "ATENÇÂO" & vbCrLf & _
                        "Esta OP, já esta relacionada neste Gride !!!", vbOKOnly + vbExclamation, "Aviso"
                Call LimpaCamposLabel
                txtCODOP.Text = Empty
                mskDTENTREGA.Text = "__/__/____"
                txtQTDOPPROG.Text = Empty
                Exit Sub
            End If
        End If
    End With
    
    Dim I               As Long
    Dim J               As Long
    Dim K               As Long
    Dim L               As Long
    
    Dim dtDATASEL       As Date
    Dim lngGRPLINHASEL  As Long
    Dim lngCODLINHASEL  As Long
    Dim lngQTDOPS       As Long
    Dim strBLOCOSEL     As String
    Dim boolSAIU        As Boolean
    Dim intRESP         As Integer
    Dim boolFRACIONA    As Boolean
    Dim lngQTDPROG      As Long
    Dim lngQTDDIFOP     As Long
    Dim lngCODOP        As Long
    Dim lngQTDOP        As Long
    Dim lngQTDFOLHAS    As Long
    Dim lngQTDNECFOLHAS As Long
    Dim strCAMPOS       As String
    Dim arrCAMPOS()     As String
    Dim dtDATAENTREGA   As Date
    
    With grdPROGOP
        If .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_TIPO) = "H" Or _
           .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_TIPO) = "I" Then
           
            '' Não deixa incluir com Data Retroativa
            If CDate(.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_Data)) < Date Then
                 MsgBox "ATENÇÂO" & vbCrLf & _
                        "Não é possivel inserir neste dia, data atual do sistema maior que data selecionada, Usar outro dia !", vbOKOnly + vbExclamation, "Aviso"
                 Exit Sub
            End If
            If VerifDispLinha(.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_BLOCOdeOPS)) = False Then
                MsgBox "ATENÇÂO" & vbCrLf & "Não há quatidade disponivel para inserir OP's, Use outro dia !", vbOKOnly + vbExclamation, "Aviso"
                Exit Sub
            End If
           
            '' Verificando a Disponibilidade da Linha
            boolFRACIONA = False
            If VerificaDispLinha(.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_BLOCOdeOPS), txtQTDOPPROG.Text) = False Then
                intRESP = MsgBox("ATENÇÂO" & vbCrLf & _
                                 "Estourou a Capacidade da Linha Fracionar a OP ?", vbYesNo + vbQuestion + vbDefaultButton2, "Pergunta")
                
                If intRESP = vbYes Then
                
                    '' =================================
                    '' Fraciona
                    boolFRACIONA = True
                    lngCODOP = CLng(txtCODOP.Text)
                    
                    '' Quantidade real da OP
                    lngQTDOP = CLng(txtQTDOPPROG.Text)
                    
                    lngQTDDIFOP = DiferencaOP(.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_BLOCOdeOPS), txtQTDOPPROG.Text)
                    txtQTDOPPROG.Text = lngQTDDIFOP
                                
                    dtDATAENTREGA = CDate(mskDTENTREGA.Text)
                    
                Else
                    
                    Call LimpaCamposLabel
                    txtCODOP.Text = Empty
                    mskDTENTREGA.Text = "__/__/____"
                    txtQTDOPPROG.Text = Empty
                
                    Exit Sub
                End If
            End If
            
            dtDATASEL = CDate(.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_Data))
            lngGRPLINHASEL = .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_CodGRPLINHA)
            lngCODLINHASEL = .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_CODLINHA)
            strBLOCOSEL = .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_BLOCOdeOPS)
            
        
        End If
    End With
    
    For I = 1 To UBound(arrLINHAS)
        If arrLINHAS(I).lngQTDLINHAS > 0 Then
            For J = 1 To arrLINHAS(I).lngQTDLINHAS
                
                If arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG = dtDATASEL And _
                   arrLINHAS(I).arrDIAS_LINHA(J).lngCODGRPLIN = lngGRPLINHASEL And _
                   arrLINHAS(I).arrDIAS_LINHA(J).lngCodLinha = lngCODLINHASEL And _
                   arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP = strBLOCOSEL Then
                   
                   '' Adicionando elemento na Array
                   Call Add_ElArray(I _
                                   , J _
                                   , Trim(txtCODOP.Text) _
                                   , arrLINHAS(I).arrDIAS_LINHA(J).lngID_PAI _
                                   , dtDATASEL _
                                   , lngGRPLINHASEL _
                                   , lngCODLINHASEL _
                                   , strBLOCOSEL _
                                   , dtDATAENTREGA _
                                   , dtDATAENTREGA _
                                   , CLng(lblQTDOP.Caption) _
                                   , CLng(txtQTDOPPROG.Text) _
                                   , boolFRACIONA _
                                   , -1 _
                                   , dacEnumUpdateAction_Insert _
                                   , 1)
                    
                    '' Ultima Linha selecionada
                    lngULTLINSEL = lngLINHASEL
                    
                    Call LimpaCamposLabel
                    txtCODOP.Text = Empty
                    mskDTENTREGA.Text = "__/__/____"
                    txtQTDOPPROG.Text = Empty
                    
                    Exit For
                    
                End If
            Next J
        End If
    Next I
    
    Call ConfGridProgOP
    Call ConfGrd
    
    Call CarregaLinhasDoArray
    Call lstStatus_Click
    
    If lngLINHASEL <= 0 Then
        lngLINHASEL = lngULTLINSEL
        grdPROGOP.Row = lngLINHASEL
        txtCODOP.SetFocus
    End If
    
    
    If boolFRACIONA = True Then
        
        '' Se for Fracionado Realocar o Saldo em Outro Lugar
        Dim dtDATAFLUT          As Date
        Dim lngDISPLINHA        As Long
        Dim lngQTDJAPROG        As Long
        Dim lngSALDOLINHA       As Long
        Dim lngSALDOOP          As Long
        Dim lngQTDOPPROGORIG    As Long
        Dim lngQTDOPSNOVAS      As Long
        Dim boolOPFRAC          As Boolean
        
        lngQTDOPPROGORIG = 0
        dtDATAFLUT = PegaDataValida(dtDATASEL, lngGRPLINHASEL, lngCODLINHASEL)
        lngSALDOOP = CalcSaldoOP(lngCODOP, lngQTDOP)
        
        boolOPFRAC = True
        
        lngQTDOPSNOVAS = 0
        For I = 1 To UBound(arrLINHAS)
            If arrLINHAS(I).lngQTDLINHAS > 0 Then
            
                If arrLINHAS(I).lngCODGRPLIN = lngGRPLINHASEL And _
                   arrLINHAS(I).lngCodLinha = lngCODLINHASEL Then
            
                    For J = 1 To arrLINHAS(I).lngQTDLINHAS
                    
                        If arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG = dtDATAFLUT Then
                                
                            '' Começando o Dia Novo
                            '' Pegando a Disponibilidade da Linha
                            
                            '' Verificando se a OP Existe (Fracionada)
                            '' Esta Incluindo a OP no Proximo dia
                            If boolOPFRAC = True Then
                                lngDISPLINHA = PegaDispLinha(arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP)
                                If lngSALDOOP >= lngDISPLINHA Then
                                    lngSALDOOP = lngDISPLINHA
                                End If
                                
                                '' Adicionando elemento na Array
                                Call Add_ElArray(I _
                                                , J _
                                                , Trim(Str(lngCODOP)) _
                                                , arrLINHAS(I).arrDIAS_LINHA(J).lngID_PAI _
                                                , arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG _
                                                , arrLINHAS(I).arrDIAS_LINHA(J).lngCODGRPLIN _
                                                , arrLINHAS(I).arrDIAS_LINHA(J).lngCodLinha _
                                                , arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP _
                                                , dtDATAENTREGA _
                                                , dtDATAENTREGA _
                                                , lngSALDOOP _
                                                , lngSALDOOP _
                                                , True _
                                                , -1 _
                                                , dacEnumUpdateAction_Insert _
                                                , 1)
                            
                                boolOPFRAC = False
                                lngSALDOOP = (lngQTDOP - PegaTotalJaProgOP(lngCODOP))
                                If lngSALDOOP > 0 Then
                                    boolOPFRAC = True
                                    dtDATAFLUT = PegaDataValida(arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG, lngGRPLINHASEL, lngCODLINHASEL)
                                End If
                                
                            End If
                        
                        'ElseIf arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG > dtDATAFLUT Then
                        '    If arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS > 0 Then
                           
                        '        For K = 1 To arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS
                        '            If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP <> lngCODOP Then
                        '
                        '                    lngQTDOPSNOVAS = (lngQTDOPSNOVAS + 1)
                        '                    ReDim Preserve arrOPSINCUSAS(1 To lngQTDOPSNOVAS) As OPS_INCLUSAS
                        '
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngID_LINHA = arrLINHAS(I).arrDIAS_LINHA(J).lngID_PAI
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).dtDATAPROG = arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngCIDGRPLIN = arrLINHAS(I).arrDIAS_LINHA(J).lngCODGRPLIN
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngCODLINA = arrLINHAS(I).arrDIAS_LINHA(J).lngCodLinha
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).strBLOCOOP = arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngAction2Do = dacEnumUpdateAction_delete
                                            
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngCODOP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).strCODROTULO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).strCODROTULO
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).strDESCROTULO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).strDESCROTULO
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).dtDATAENTREGA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).dtDATAENTREGA
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).dtDATAENTREGAORIGINAL = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).dtDATAENTREGAORIGINAL
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngQTDOPORIGINAL = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOPORIGINAL
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngQTDOPPROGRAMADA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOPPROGRAMADA
                                            
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).intNECK = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).intNECK
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngCODOPBKP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOPBKP
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngIDOP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDOP
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngCODPED = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODPED
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngIDPRODUTO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDPRODUTO
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngCODSTATUS = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATUS
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngCODSTATUSORIGINAL = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATUSORIGINAL
                                            
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).strTIPO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).strTIPO
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).intSELECIONADO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).intSELECIONADO
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngIDINTERNO = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDINTERNO
                                            
                                            '' Grando os Indeces da Linha do Array
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngIDARRAYLINHA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDARRAYLINHA
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngIDARRAYDIA = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDARRAYDIA
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngIDARRAYOP = lngQTDOPSNOVAS
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).strINDICE = Trim(Str(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP)) & Trim(Str(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDPRODUTO)) & Trim(arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP)
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).intFRACIONADA = 0
                                            
                        '                    arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngAction2Do = dacEnumUpdateAction_delete
                                            
                        '                    lngCODOP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP
                                            
                                            '' =============================================================================
                        '                    '' Incluindo as Folhas que serão usadas
                        '                    lngQTDFOLHAS = 0
                        '                    sSql = ""
                                            
                        '                    sSql = "Select" & vbCrLf
                        '                    sSql = sSql & "       PROD.SGI_IDPRODUTO    " & vbCrLf
                        '                    sSql = sSql & "      ,PROD.SGI_CODIGO       " & vbCrLf
                        '                    sSql = sSql & "      ,PROD.SGI_DESCRICAO    As SGI_DESCPROD" & vbCrLf
                        '                    sSql = sSql & "      ,PROD.SGI_CODLINPROD   " & vbCrLf
                        '                    sSql = sSql & "      ,LINH.SGI_CODIGO       As SGI_IDLINHA" & vbCrLf
                        '                    sSql = sSql & "      ,MEDC.SGI_CODMEDCORT   As SGI_CODFOLHA" & vbCrLf
                        '                    sSql = sSql & "      ,MEDC.SGI_EXPESS       " & vbCrLf
                        '                    sSql = sSql & "      ,MEDC.SGI_LARGUR       " & vbCrLf
                        '                    sSql = sSql & "      ,MEDC.SGI_COMPRI       " & vbCrLf
                        '                    sSql = sSql & "      ,MEDC.SGI_QTDECORPOS   " & vbCrLf
                        '                    sSql = sSql & "      ,MEDC.SGI_PERDPROC     " & vbCrLf
                        '                    sSql = sSql & "      ,DIMC.SGI_DESCORTE     " & vbCrLf
                                            
                        '                    sSql = sSql & "  From" & vbCrLf
                        '                    sSql = sSql & "       SGI_CADPRODUTO        PROD" & vbCrLf
                        '                    sSql = sSql & "      ,SGI_CADLINHAPRODUTO   LINH" & vbCrLf
                        '                    sSql = sSql & "      ,SGI_MEDCORTELINHA     MEDC" & vbCrLf
                        '                    sSql = sSql & "      ,SGI_CADDIMCORTE       DIMC" & vbCrLf
                                            
                        '                    sSql = sSql & " Where" & vbCrLf
                        '                    sSql = sSql & "       PROD.SGI_FILIAL       = " & FILIAL & vbCrLf
                        '                    sSql = sSql & "   And PROD.SGI_IDPRODUTO    = " & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDPRODUTO & vbCrLf
                                            
                        '                    sSql = sSql & "   And LINH.SGI_FILIAL       = PROD.SGI_FILIAL" & vbCrLf
                        '                    sSql = sSql & "   And LINH.SGI_CODLIN       = PROD.SGI_CODLINPROD" & vbCrLf
                                            
                        '                    sSql = sSql & "   And MEDC.SGI_FILIAL       = LINH.SGI_FILIAL" & vbCrLf
                        '                    sSql = sSql & "   And MEDC.SGI_CODIGO       = LINH.SGI_CODIGO" & vbCrLf
                                            
                        '                    sSql = sSql & "   And DIMC.SGI_FILIAL       = MEDC.SGI_FILIAL" & vbCrLf
                        '                    sSql = sSql & "   And DIMC.SGI_CODIGO       = MEDC.SGI_CODMEDCORT"
                    
                        '                    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
                        '                    If Not BREC11.EOF() Then
                                                
                        '                        Do While Not BREC11.EOF()
                                                    
                        '                            lngQTDFOLHAS = (lngQTDFOLHAS + 1)
                        '                            ReDim Preserve arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(1 To lngQTDFOLHAS) As OPS_INCLUSAS_FOLHAS_USADAS
                    
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngFOLHAUSADA = 2
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngCODOP = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).strCODPROD = arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).strCODROTULO
                                                   
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngIDLIN = BREC11!SGI_IDLINHA
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngCODLIN = BREC11!SGI_CODLINPROD
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).strDESCFOLHAUSADA = BREC11!SGI_DESCORTE
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngCODFOLHAUSADA = BREC11!SGI_CODFOLHA
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngIDPROD = BREC11!SGI_IDPRODUTO
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).strINDICE = Trim(Str(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP)) & Trim(Str(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDPRODUTO)) & Trim(arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP)
                        '                            arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngLINHA = lngQTDFOLHAS
                                                    
                        '                            If Not IsNull(BREC11!SGI_EXPESS) Then arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblESPESS = BREC11!SGI_EXPESS
                        '                            If Not IsNull(BREC11!SGI_LARGUR) Then arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblLARG = BREC11!SGI_LARGUR
                        '                            If Not IsNull(BREC11!SGI_COMPRI) Then arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblCOMP = BREC11!SGI_COMPRI
                        '                            If Not IsNull(BREC11!SGI_QTDECORPOS) Then arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDECORP = BREC11!SGI_QTDECORPOS
                        '                            If Not IsNull(BREC11!SGI_PERDPROC) Then arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblPERDPRODC = BREC11!SGI_PERDPROC
                        '
                        '                            If Not IsNull(BREC11!SGI_QTDECORPOS) Then
                                                    
                        '                                lngQTDNECFOLHAS = 0
                        '                                If BREC11!SGI_QTDECORPOS > 0 Then lngQTDNECFOLHAS = (arrOPSINCUSAS(lngQTDOPSNOVAS).lngQTDOPPROGRAMADA / BREC11!SGI_QTDECORPOS)
                        '                                arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngNECEFOLHAS = lngQTDNECFOLHAS
                                                    
                        '                                arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblPESO = 0
                        '                                arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDEFOLHAS = 0
                        '                                arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDELATAS = 0
                                                    
                        '                                strCAMPOS = SomaSaldos(Str(BREC11!SGI_IDPRODUTO), BREC11!SGI_CODFOLHA, BREC11!SGI_QTDECORPOS)
                        '                                If Len(Trim(strCAMPOS)) > 0 Then
                        '                                    arrCAMPOS = Split(strCAMPOS, "|")
                        '                                    arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblPESO = CDbl(arrCAMPOS(0))
                        '                                    arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDEFOLHAS = CLng(arrCAMPOS(1))
                        '                                    arrOPSINCUSAS(lngQTDOPSNOVAS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDELATAS = CLng(arrCAMPOS(2))
                        '                                End If
                        '                            End If
                                                         
                        '                            BREC11.MoveNext
                        '                        Loop
                        '                    End If
                        '                    BREC11.Close
                        '                    arrOPSINCUSAS(lngQTDOPSNOVAS).lngQTDFOLHASUSADAS = lngQTDFOLHAS
                        '                    '' =============================================================================
                                        
                        '            End If
                        '        Next K
                            
                        '    End If
                        End If
                    Next J
                End If
            End If
        Next I
        
        
        '' ===================================
        '' Recalculando a Array
        '' Começando o Dia Novo
        '' Pegando a Disponibilidade da Linha
        'lngQTDJAPROG = 0
        'For I = 1 To UBound(arrOPSINCUSAS)
        '    lngDISPLINHA = PegaDispLinha(arrOPSINCUSAS(I).strBLOCOOP)
        '    If lngSALDOLINHA = 0 Then lngSALDOLINHA = lngDISPLINHA
        '
        '    lngSALDOOP = (arrOPSINCUSAS(I).lngQTDOPORIGINAL - PegaTotalJaProgOP(arrOPSINCUSAS(I).lngCODOP))
        '
        '    Do While lngSALDOOP > 0
        '        lngSALDOOP = (arrOPSINCUSAS(I).lngQTDOPORIGINAL - PegaTotalJaProgOP(arrOPSINCUSAS(I).lngCODOP))
        '
        '        If lngSALDOOP >= lngSALDOLINHA Then
        '            lngSALDOOP = (lngSALDOOP - lngSALDOLINHA)
        '            lngQTDOP = lngSALDOLINHA
        '            arrOPSINCUSAS(I).lngQTDOPPROGRAMADA = lngQTDOP
        '            boolOPFRAC = True
        '        Else
        '            lngQTDOP = (arrOPSINCUSAS(I).lngQTDOPORIGINAL - PegaTotalJaProgOP(arrOPSINCUSAS(I).lngCODOP))
        '            lngSALDOOP = (arrOPSINCUSAS(I).lngQTDOPORIGINAL - (lngQTDOP + PegaTotalJaProgOP(arrOPSINCUSAS(I).lngCODOP)))
        '            arrOPSINCUSAS(I).lngQTDOPPROGRAMADA = lngQTDOP
        '            boolOPFRAC = True
        '        End If
        '
        '        lngQTDJAPROG = (lngQTDJAPROG + lngQTDOP)
        '        lngSALDOLINHA = (lngDISPLINHA - lngQTDJAPROG)
        '
        '        If boolOPFRAC = True Then
        '            For J = 1 To UBound(arrLINHAS)
        '                If arrLINHAS(J).lngCODGRPLIN = lngGRPLINHASEL And _
        '                   arrLINHAS(J).lngCodLinha = lngCODLINHASEL Then
        '                   For K = 1 To arrLINHAS(J).lngQTDLINHAS
        '                        If arrLINHAS(J).arrDIAS_LINHA(K).dtDATAPROG = dtDATAFLUT Then
        '
        '                            '' Adicionando elemento na Array
        '                            Call Add_ElArray(J _
        '                                            , K _
        '                                            , Trim(Str(arrOPSINCUSAS(I).lngCODOP)) _
        '                                            , arrLINHAS(J).arrDIAS_LINHA(K).lngID_PAI _
        '                                            , arrLINHAS(J).arrDIAS_LINHA(K).dtDATAPROG _
        '                                            , arrLINHAS(J).arrDIAS_LINHA(K).lngCODGRPLIN _
        '                                            , arrLINHAS(J).arrDIAS_LINHA(K).lngCodLinha _
        '                                            , arrLINHAS(J).arrDIAS_LINHA(K).strBLOCOOP _
        '                                            , arrOPSINCUSAS(I).dtDATAENTREGA _
        '                                            , arrOPSINCUSAS(I).lngQTDOPORIGINAL _
        '                                            , arrOPSINCUSAS(I).lngQTDOPPROGRAMADA _
        '                                            , True _
        '                                            , arrOPSINCUSAS(I).lngIDINTERNO)
        '                        End If
        '                   Next K
        '                End If
        '            Next J
        '
        '        End If
        '
        '        If lngSALDOLINHA = 0 Then
        '            dtDATAFLUT = PegaDataValida(dtDATAFLUT, lngGRPLINHASEL, lngCODLINHASEL)
        '            lngDISPLINHA = PegaDispLinha(PegaBlocoOP(dtDATAFLUT, lngGRPLINHASEL, lngCODLINHASEL))
        '            lngSALDOLINHA = lngDISPLINHA
        '            lngQTDJAPROG = 0
        '        End If
        '
        '    Loop
        'Next I
        '' ===================================
        
        Call ConfGridProgOP
        Call ConfGrd
        
        Call CarregaLinhasDoArray
        Call lstStatus_Click
        
        If lngLINHASEL <= 0 Then
            lngLINHASEL = lngULTLINSEL
            grdPROGOP.Row = lngLINHASEL
            txtCODOP.SetFocus
        End If

    End If
    
End Sub

Private Sub grdPROGOP_RowColChange()
    Call PegaLinhaAtual
End Sub

Private Sub grdPROGOP_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim strINDICE   As String

     With grdPROGOP
          Select Case Col
                 Case conCOL_PRODOP_QtdeOPProgramada
                        If .EditText = Empty Then
                            MsgBox "Não é Permitido inserir valores nulos !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Valore Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        If Len(Trim(.EditText)) = 0 Then
                            MsgBox "Não é Permitido inserir valores nulos !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        If .EditText = 0 Then
                            MsgBox "Não é Permitido inserir valores nulos !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        '' Se caso já estiver Lotado não deixa mais incluir OP's
                        If VerificaDispLinha(.Cell(flexcpText, .Row, conCOL_PRODOP_BLOCOdeOPS), CLng(.EditText)) = False Then
                            Cancel = True
                            Exit Sub
                        End If
                        
                        
                 Case conCOL_PRODOP_DataEntrega
                        If .EditText = Empty Then
                            MsgBox "Não é Permitido inserir valores nulos !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        If Len(Trim(.EditText)) < 10 Then
                            MsgBox "Data de Entrega Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                       
                        If Not IsDate(.EditText) Then
                            MsgBox "Data de Entrega Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
          
          End Select
     End With

End Sub

Private Sub LimpaCamposLabel()
    lblCODROT.Caption = ""
    lblDESCROT.Caption = ""
    lblNECK.Caption = ""
    lblFECH.Caption = ""
    lblCOMP.Caption = ""
    lblQTDOP.Caption = ""
    lblSTATUS.Caption = ""
End Sub


Private Sub lstStatus_Click()
    Call Filtro
End Sub

Private Sub lstStatus_ItemCheck(Item As Integer)
    intFILTROSEL = Item
End Sub

Private Sub mskDTENTREGA_GotFocus()
    objBLBFunc.SelecionaCampos mskDTENTREGA.Name, Me
End Sub

Private Sub mskDTENTREGA_Validate(Cancel As Boolean)

    If Len(Trim(Replace(Replace(mskDTENTREGA.Text, "/", ""), "_", ""))) = 0 Then Exit Sub
    
    If Not IsDate(mskDTENTREGA.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Data Inválida !!!"
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtCODOP_GotFocus()
    objBLBFunc.SelecionaCampos txtCODOP.Name, Me
End Sub

Private Sub txtCODOP_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODOP.Text
End Sub

Private Sub txtCODOP_Validate(Cancel As Boolean)

    If Len(Trim(txtCODOP.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODOP.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Somente é permitido numeros !!!"
        txtCODOP.Text = ""
        Cancel = True
        Exit Sub
    End If
    If (grdPROGOP.Rows - 1) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "A Programação não foi carregada !!!", vbOKOnly + vbExclamation, "Aviso"
        Call LimpaCamposLabel
        txtCODOP.Text = ""
        Cancel = True
        Exit Sub
    End If
    
    Dim lngPESQ As Long
    
    With grdPROGOP
        lngPESQ = .FindRow(txtCODOP.Text, , conCOL_PRODOP_CodOP)
        If lngPESQ > -1 Then
            If .Cell(flexcpText, lngPESQ, conCOL_PRODOP_Action2Do) <> dacEnumUpdateAction_delete Then
                MsgBox "ATENÇÂO" & vbCrLf & "Esta OP já esta relacionada na Gride !!!", vbOKOnly + vbExclamation, "Aviso"
                Call LimpaCamposLabel
                txtCODOP.Text = ""
                Cancel = True
                Exit Sub
            End If
        End If
    End With
    
    
    txtCODOP.Text = Trim(Replace(Replace(txtCODOP.Text, ",", ""), ".", ""))
    If PegaOP(txtCODOP.Text) = False Then Cancel = True

End Sub

Private Function PegaOP(strCODOP As String) As Boolean

    PegaOP = False
    
    If Len(Trim(strCODOP)) = 0 Then Exit Function
    
    Dim lngPESQOP   As Long
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PROD.*" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_QTDE" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_STATUS" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_NECKIN" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_SALDO As SGI_SALDO_OP"
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMTABELA & " ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO    = " & strCODOP & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL    = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = ORDP.SGI_IDPRODUTO" & vbCrLf
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF() Then
        If BREC4!SGI_STATUS <> 0 And _
           BREC4!SGI_STATUS <> 1 And _
           BREC4!SGI_STATUS <> 6 And _
           BREC4!SGI_STATUS <> 7 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "OP não esta com o Status de acordo para ser Inclusa !!!", vbOKOnly + vbExclamation, "Aviso"
            Call LimpaCamposLabel
            txtCODOP.Text = ""
        Else
            PegaOP = True
            lblCODROT.Caption = BREC4!SGI_CODIGO
            lblDESCROT.Caption = BREC4!SGI_DESCRICAO
            mskDTENTREGA.Text = Format(BREC4!SGI_DATENTREGA, "DD/MM/YYYY")
            lblNECK.Caption = IIf(BREC4!SGI_NECKIN = 1, "Sim", "Não")
            lblFECH.Caption = PegaFechamentoTampaFuro(Str(BREC4!SGI_IDPRODUTO))
            lblCOMP.Caption = ""
            lblQTDOP.Caption = BREC4!SGI_QTDE
            txtQTDOPPROG.Text = BREC4!SGI_QTDE
            lblSTATUS.Caption = PegaStatus(BREC4!SGI_STATUS)
            
            '' Verificando o Estoque do Produto
            '' Call PegaEstoqueProdutos(Str(BREC4!SGI_IDPRODUTO), strCODOP)
            
            '' Verificando se já esta na Gride
            lngPESQOP = grdPROGOP.FindRow(strCODOP, , conCOL_PRODOP_CodOP)
            If lngPESQOP > 0 Then
                Call PegaSaldo(BREC4!SGI_QTDE, strCODOP)
            Else
                If BREC4!SGI_STATUS = 1 Then '' Saldo da OP
                    lblQTDOP.Caption = BREC4!SGI_SALDO_OP
                    txtQTDOPPROG.Text = BREC4!SGI_SALDO_OP
                End If
            End If
            
        End If
    Else
        MsgBox "ATENÇÃO" & vbCrLf & _
               "OP não existe !!!", vbOKOnly + vbExclamation, "Aviso"
        Call LimpaCamposLabel
    End If
    BREC4.Close
    
End Function

Private Sub txtQTDOPPROG_GotFocus()
    objBLBFunc.SelecionaCampos txtQTDOPPROG.Name, Me
End Sub

Private Sub txtQTDOPPROG_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQTDOPPROG.Text
End Sub

Private Sub txtQTDOPPROG_Validate(Cancel As Boolean)

    If Len(Trim(txtQTDOPPROG.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtQTDOPPROG.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Somente é permitido numeros !!!"
        txtQTDOPPROG.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub

Private Function VerifOP(strCODOP As String, strTOTALOP As String, strTOTALPROG As String) As Boolean

    VerifOP = False

    If Len(Trim(strCODOP)) = 0 Then Exit Function
    If Len(Trim(strTOTALOP)) = 0 Then Exit Function
    
    Dim I               As Long
    Dim lngTOTALOP      As Long
    Dim lngTOTALOPPROG  As Long
    Dim lngSOMAQTDPROG  As Long
    
    lngTOTALOP = CLng(strTOTALOP)
    lngTOTALOPPROG = CLng(strTOTALPROG)
    
    lngSOMAQTDPROG = 0
    With grdPROGOP
        For I = 1 To (.Rows - 1)
            If Trim(strCODOP) = .Cell(flexcpText, I, conCOL_PRODOP_CodOP) And _
               .Cell(flexcpText, I, conCOL_PRODOP_Action2Do) <> dacEnumUpdateAction_delete Then
               lngSOMAQTDPROG = lngSOMAQTDPROG + CLng(.Cell(flexcpText, I, conCOL_PRODOP_QtdeOPProgramada))
            End If
        Next I
    End With
    lngSOMAQTDPROG = (lngSOMAQTDPROG + lngTOTALOPPROG)

    If lngSOMAQTDPROG > lngTOTALOP Then Exit Function

    VerifOP = True
    
End Function

Private Function VerificaSeExisteMes(lngMES As Long, lngANO As Long) As Boolean

    VerificaSeExisteMes = False
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And Month(SGI_DATAPROG) = " & lngMES & vbCrLf
    sSql = sSql & "   And Year(SGI_DATAPROG)  = " & lngANO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then VerificaSeExisteMes = True
    BREC.Close
    
End Function

Private Function VerificaDispLinha(strDTDIA As String, lngQTDDPROG As Long) As Boolean

    VerificaDispLinha = False
    
    Dim lngTOTISP       As Long
    Dim lngTOTPROG      As Long
    Dim lngTOTJAPROG    As Long
    Dim lngLINHADIAINI  As Long
    
    With grdPROGOP
        lngLINHADIAINI = .FindRow(strDTDIA, , conCOL_PRODOP_BLOCOdeOPS)
        If lngLINHADIAINI > -1 Then
                
                lngTOTISP = CLng(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_Capacidade))
                
                lngTOTPROG = 0
                If Len(Trim(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_TotalProgramado))) > 0 Then lngTOTPROG = CLng(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_TotalProgramado))
                
                lngTOTJAPROG = (lngTOTISP - (lngTOTPROG + lngQTDDPROG))
                
                If lngTOTJAPROG < 0 Then
''                    MsgBox "ATENÇÃO" & vbCrLf & _
''                           "A capacidade já esta estourada não e permitido acrescentar mais OP's !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                ElseIf lngTOTJAPROG < 0 Then
                    MsgBox "ATENÇÂO" & vbCrLf & _
                           "Capacidade da Linha esta zerada !!!, Impossivel incluir OP's", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
            
        End If
    End With
    
    VerificaDispLinha = True
    
End Function

Private Sub PegaSaldo(lngQTDOPREAL As Long, strCODOP As String)

    Dim lngLINHA As Long
    Dim lngSALDO As Long
    Dim lngQTDJAINF As Long
    Dim I As Long

    
    With grdPROGOP
    
        lngLINHA = .FindRow(strCODOP, , conCOL_PRODOP_CodOP)

        If lngLINHA > -1 Then
        
            For I = 1 To (.Rows - 1)
                If .Cell(flexcpText, I, conCOL_PRODOP_CodOP) = strCODOP Then
                    lngQTDJAINF = lngQTDJAINF + CLng(.Cell(flexcpText, I, conCOL_PRODOP_QtdeOPProgramada))
                End If
            Next I
            
            lngSALDO = (lngQTDOPREAL - lngQTDJAINF)
            If lngSALDO > 0 Then txtQTDOPPROG.Text = lngSALDO
            
        End If
    End With
End Sub

Private Function PegaQtdeRealProduzida(strCODOP As String, strIDOP As String, strIDINTERNO As String) As Long

    PegaQtdeRealProduzida = 0
    
    If Len(Trim(strCODOP)) = 0 Then Exit Function
    If Len(Trim(strIDOP)) = 0 Then Exit Function
    If Len(Trim(strIDINTERNO)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODOP      = " & Trim(strCODOP) & vbCrLf
    sSql = sSql & "   And SGI_IDINTOP    = " & Trim(strIDOP) & vbCrLf
    sSql = sSql & "   And SGI_IDINTPROG  = " & Trim(strIDINTERNO)
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then PegaQtdeRealProduzida = BREC8!SGI_QTDEPROD
    BREC8.Close
    
End Function

Private Function PegaStatusApontamento(strCODOP As String, strIDOP As String, strIDINTERNO As String) As Long

    PegaStatusApontamento = 0
    
    If Len(Trim(strCODOP)) = 0 Then Exit Function
    If Len(Trim(strIDOP)) = 0 Then Exit Function
    If Len(Trim(strIDINTERNO)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODOP      = " & Trim(strCODOP) & vbCrLf
    sSql = sSql & "   And SGI_IDINTOP    = " & Trim(strIDOP) & vbCrLf
    sSql = sSql & "   And SGI_IDINTPROG  = " & Trim(strIDINTERNO)
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then PegaStatusApontamento = BREC8!SGI_STATUSAPONT
    BREC8.Close
    
End Function


Private Function PegaDescrStatusApontamento(strCODSTATUS As String)

    PegaDescrStatusApontamento = ""
    
    If Len(Trim(strCODSTATUS)) = 0 Then Exit Function
    If CLng(strCODSTATUS) = 0 Then Exit Function

    If CLng(strCODSTATUS) = 1 Then PegaDescrStatusApontamento = "Concluido"
    If CLng(strCODSTATUS) = 2 Then PegaDescrStatusApontamento = "Parcial"
    If CLng(strCODSTATUS) = 3 Then PegaDescrStatusApontamento = "Em Produção"
    
End Function

Private Sub MudaCorStatus(lngCODSTATUS As Long, lngROW As Long)
    With grdPROGOP
        If lngCODSTATUS = 1 Then        '' Concluido
            .Cell(flexcpBackColor, lngROW, conCOL_PRODOP_CodOP, lngROW, (.Cols - 1)) = &H4000&
            .Cell(flexcpForeColor, lngROW, conCOL_PRODOP_CodOP, lngROW, (.Cols - 1)) = vbWhite
        ElseIf lngCODSTATUS = 2 Then    '' Parcial
            .Cell(flexcpBackColor, lngROW, conCOL_PRODOP_CodOP, lngROW, (.Cols - 1)) = &HC0C0&
            .Cell(flexcpForeColor, lngROW, conCOL_PRODOP_CodOP, lngROW, (.Cols - 1)) = vbBlack
        Else
            .Cell(flexcpBackColor, lngROW, conCOL_PRODOP_CodOP, lngROW, (.Cols - 1)) = vbWhite
            .Cell(flexcpForeColor, lngROW, conCOL_PRODOP_CodOP, lngROW, (.Cols - 1)) = vbBlack
        End If
    End With
End Sub

Private Sub GeraSubtotais()
    With grdPROGOP
            '' GroupOn -> se refere qual coluna ira pegar para agrupar (Somar os Valores)
            '' TotalOn -> Se refere em qual coluna será mostrado o Agurpamento do Total (Soma dos Valores)
            .Subtotal flexSTSum, -1, 2, "#,", 2, vbWhite, True
            .Subtotal flexSTSum, 0, 2, "#,", &H404040, vbWhite, True
            .Subtotal flexSTSum, 1, 2, "#,", &HC0C0C0, vbBlack, True
    
            ' merge
            .MergeCells = flexMergeRestrictAll
            .MergeCol(conCOL_PRODOP_DESCLINHA) = True
            .MergeCol(conCOL_PRODOP_Data) = True
    End With
End Sub

Private Sub ConfLstLinhas()
    With lstStatus
            .Clear

            sSql = ""
            
            sSql = "Select Distinct" & vbCrLf
            sSql = sSql & "       LINHA.SGI_CODIGO                 As SGI_IDLINHA" & vbCrLf
            sSql = sSql & "      ,LINHA.SGI_CODLIN" & vbCrLf
            sSql = sSql & "      ,LINHA.SGI_DESCRI                 As SGI_DESCLINHA" & vbCrLf
            
            sSql = sSql & "  From" & vbCrLf
            sSql = sSql & "       SGI_CADGRUPLINHAIT" & strNOMTABELA & " GRPLINITE" & vbCrLf
            sSql = sSql & "      ,SGI_MAQULIN_MESANO" & strNOMTABELA & " GRPLINMESANO" & vbCrLf
            sSql = sSql & "      ,SGI_MAQULIN_CAPAC" & strNOMTABELA & "  GRPCAPACDIA" & vbCrLf
            sSql = sSql & "      ,SGI_CADLINHAPRODUTO      LINHA" & vbCrLf
            
            sSql = sSql & " Where" & vbCrLf
            sSql = sSql & "       GRPLINITE.SGI_FILIAL        = " & FILIAL & vbCrLf
            sSql = sSql & "   And GRPLINMESANO.SGI_FILIAL     = GRPLINITE.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And GRPLINMESANO.SGI_CODIGO     = GRPLINITE.SGI_CODIGO" & vbCrLf
            sSql = sSql & "   And GRPLINMESANO.SGI_IDPAI      = GRPLINITE.SGI_IDINTERNO" & vbCrLf
            sSql = sSql & "   And GRPLINMESANO.SGI_MES        = " & cboMes.ItemData(cboMes.ListIndex) & vbCrLf
            sSql = sSql & "   And GRPLINMESANO.SGI_ANO        = " & cboAno.ItemData(cboAno.ListIndex) & vbCrLf
            
            sSql = sSql & "   And GRPCAPACDIA.SGI_FILIAL      = GRPLINMESANO.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And GRPCAPACDIA.SGI_CODIGO      = GRPLINMESANO.SGI_CODIGO" & vbCrLf
            sSql = sSql & "   And GRPCAPACDIA.SGI_IDPAI       = GRPLINMESANO.SGI_IDPAI" & vbCrLf
            sSql = sSql & "   And Month(GRPCAPACDIA.SGI_DATA) = GRPLINMESANO.SGI_MES" & vbCrLf
            sSql = sSql & "   And Year(GRPCAPACDIA.SGI_DATA)  = GRPLINMESANO.SGI_ANO" & vbCrLf
            sSql = sSql & "   And GRPCAPACDIA.SGI_ATIVO       = 1" & vbCrLf
            
            sSql = sSql & "   And LINHA.SGI_FILIAL            = GRPCAPACDIA.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And LINHA.SGI_CODLIN            = GRPCAPACDIA.SGI_CODLIN" & vbCrLf
            
            sSql = sSql & "Order By LINHA.SGI_CODLIN,LINHA.SGI_DESCRI" & vbCrLf
            
            BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
            Do While Not BREC10.EOF()
                .AddItem Trim(BREC10!SGI_DESCLINHA)
                BREC10.MoveNext
            Loop
            BREC10.Close
            
    End With
End Sub


Private Sub Filtro()

    Dim I           As Long
    Dim J           As Long
    Dim strINDICE   As String
    Dim boolTEMSELECIONADO  As Boolean
    
    '' Selecionou
    boolTEMSELECIONADO = False
    With lstStatus
        For I = 0 To (.ListCount - 1)
            strINDICE = Trim(.List(I))
            With grdPROGOP
                For J = 1 To (.Rows - 1)
                    If Trim(Replace(.Cell(flexcpText, J, conCOL_PRODOP_DESCLINHA), "Total", "")) = strINDICE Then
                       If lstStatus.Selected(I) = True Then
                            .RowHidden(J) = False
                            boolTEMSELECIONADO = lstStatus.Selected(I)
                       ElseIf lstStatus.Selected(I) = False Then
                            .RowHidden(J) = True
                       End If
                    End If
                Next J
            End With
        Next I
    End With
    
    '' Tirou Todas a Seleções
    If boolTEMSELECIONADO = False Then
        With grdPROGOP
            For J = 1 To (.Rows - 1)
                .RowHidden(J) = False
            Next J
        End With
    End If
    
    Call EncolheLinhasGrid

End Sub

Private Sub EncolheLinhasGrid()
    Dim I   As Long
    With grdPROGOP
        For I = 1 To (.Rows - 1)
            If I > 1 Then
                If .Cell(flexcpText, I, conCOL_PRODOP_LINHAEXP) <> "" Then
                    If .Cell(flexcpText, I, conCOL_PRODOP_LINHAEXP) = 0 Then
                       .GetNode(I).Expanded = False
                    ElseIf .Cell(flexcpText, I, conCOL_PRODOP_LINHAEXP) = 1 Then
                      .GetNode(I).Expanded = True
                    Else
                      .GetNode(I).Expanded = False
                    End If
                Else
                    .GetNode(I).Expanded = False
                End If
            End If
        Next I
    End With
End Sub

Private Function DiferencaOP(strDTDIA As String, lngQTDDPROG As Long) As Long
    
    DiferencaOP = 0
    
    Dim lngTOTISP       As Long
    Dim lngTOTPROG      As Long
    Dim lngTOTJAPROG    As Long
    Dim lngLINHADIAINI  As Long
    
    With grdPROGOP
        lngLINHADIAINI = .FindRow(strDTDIA, , conCOL_PRODOP_BLOCOdeOPS)
        If lngLINHADIAINI > -1 Then
                
                lngTOTISP = CLng(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_Capacidade))
                
                lngTOTPROG = 0
                If Len(Trim(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_TotalProgramado))) > 0 Then lngTOTPROG = CLng(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_TotalProgramado))
                
                lngTOTJAPROG = (lngTOTISP - lngTOTPROG)
                
        End If
    End With

    DiferencaOP = lngTOTJAPROG
    
End Function


Private Function CalcSaldoOP(lngCODOP As Long, lngQTDOP As Long) As Long

    CalcSaldoOP = 0
    
    Dim I           As Long
    Dim J           As Long
    Dim K           As Long
    Dim lngSALDO    As Long
    Dim lngQTDPROG  As Long
    
    
    lngQTDPROG = 0
    For I = 1 To UBound(arrLINHAS)
        If arrLINHAS(I).lngQTDLINHAS > 0 Then
            For J = 1 To arrLINHAS(I).lngQTDLINHAS
                If arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS > 0 Then
                    For K = 1 To arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS
                        If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP = lngCODOP And _
                           arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngAction2Do <> dacEnumUpdateAction_delete Then
                           lngQTDPROG = (lngQTDPROG + arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOPPROGRAMADA)
                        End If
                    Next K
                End If
            Next J
        End If
    Next I
    
    lngSALDO = (lngQTDOP - lngQTDPROG)
    CalcSaldoOP = lngSALDO

End Function

Private Sub CarregaLinhasDoArray()

    Dim I                   As Long
    Dim J                   As Long
    Dim K                   As Long
    Dim L                   As Long
    Dim lngINDCODFOLHA      As Long
    Dim lngQTDNECFOLHAS     As Long

    With grdPROGOP
        For I = 1 To UBound(arrLINHAS)
            For J = 1 To arrLINHAS(I).lngQTDLINHAS
                .AddItem Trim(arrLINHAS(I).strDESCGRPLINHA) & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG & vbTab & _
                         arrLINHAS(I).arrDIAS_LINHA(J).lngTOTALPECAS & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrLINHAS(I).lngCODGRPLIN & vbTab & _
                         "" & vbTab & _
                         arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP & vbTab & _
                         arrLINHAS(I).lngCodLinha & vbTab & _
                         "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                         "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                         "" & vbTab & "" & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).strTIPO & vbTab & _
                         "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & arrLINHAS(I).lngID_INTERNO & vbTab & arrLINHAS(I).lngEXPLIN
                         
                
                '' ==================================================
                '' Incluindo as OPS ao Dia
                If arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS > 0 Then
                    For K = 1 To (arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS)
                        If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngAction2Do <> dacEnumUpdateAction_delete Then
                            .AddItem Trim(arrLINHAS(I).strDESCGRPLINHA) & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG & vbTab & _
                                     "" & vbTab & "" & vbTab & _
                                     "" & vbTab & "" & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).intSELECIONADO & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).dtDATAENTREGA & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).strCODROTULO & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).strDESCROTULO & vbTab & _
                                     IIf(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).intNECK = 0, "Não", "Sim") & vbTab & _
                                     "" & vbTab & "" & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOPORIGINAL & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOPPROGRAMADA & vbTab & _
                                     IIf(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOAPONTADAORIGINAL = 0, "", arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOAPONTADAORIGINAL) & vbTab & _
                                     "" & vbTab & _
                                     PegaStatus(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATUS) & vbTab & _
                                     arrLINHAS(I).lngCODGRPLIN & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOPBKP & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP & vbTab & _
                                     arrLINHAS(I).lngCodLinha & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDOP & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODPED & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDPRODUTO & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATUS & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).dtDATAENTREGAORIGINAL & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDINTERNO & vbTab & "" & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngAction2Do & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATUSORIGINAL & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATAPONT & vbTab & PegaDescrStatusApontamento(Str(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATAPONT)) & vbTab & "" & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).strTIPO & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDARRAYLINHA & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDARRAYDIA & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngIDARRAYOP & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).intFRACIONADA & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngID_LINHA & vbTab & _
                                     arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngEXPLIN & vbTab & arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).strINDICE
                        
                                     '' ==============================================
                                     '' Icluindo Estoque na Gride de Estoque
                                     If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDFOLHASUSADAS > 0 Then
                                        With grdEstoque
                                           For L = 1 To arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDFOLHASUSADAS
                                                .AddItem "" & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngIDPROD & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).strCODPROD & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngCODLIN & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngIDLIN & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngCODFOLHAUSADA & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).strDESCFOLHAUSADA & vbTab & _
                                                         Format(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblESPESS, "#,###0.000") & vbTab & _
                                                         Format(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblLARG, "#,##0.00") & vbTab & _
                                                         Format(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblCOMP, "#,##0.00") & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngQTDECORP & vbTab & _
                                                         Format(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblPERDPRODC, "#,##0.00") & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngQTDEFOLHAS & vbTab & _
                                                         Format(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).dblPESO, "#,####0.0000") & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngQTDELATAS & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngCODOP & vbTab & _
                                                         Trim(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).strINDICE) & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngNECEFOLHAS & vbTab & _
                                                         arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngLINHA
                                                         
                                                         
                                                '' Pesquisa se Já esta Selecionada
                                                ''sSql = ""
                                                ''sSql = "Select" & vbCrLf
                                                ''sSql = "       *" & vbCrLf
                                                ''sSql = "  From " & vbCrLf
                                                ''sSql = "       SGI_CADMOVPCP_FOLHAS_" & vbCrLf
                                                ''sSql = " Where " & vbCrLf
                                                ''sSql = "       SGI_FILIAL    = " & FILIAL & vbCrLf
                                                ''sSql = "   And SGI_CODIGO    = " & objCADMOVPCP.CODIGO & vbCrLf
                                                ''sSql = "   And SGI_IDPRODUTO = " & arrLINHAS(i).arrDIAS_LINHA(j).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngIDPROD & vbCrLf
                                                ''sSql = "   And SGI_CODFOLHA  = " & arrLINHAS(i).arrDIAS_LINHA(j).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngCODFOLHAUSADA & vbCrLf
                                                ''sSql = "   And SGI_CODOP     = " & arrLINHAS(i).arrDIAS_LINHA(j).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).lngCODOP & vbCrLf
                                                ''sSql = "   And SGI_INDICE    = '" & Trim(arrLINHAS(i).arrDIAS_LINHA(j).arrOPS_INCLUSAS(K).arrFOLHAS_USADAS(L).strINDICE) & "'"
                                                
                                                
                                           Next L
                                        End With
                                     End If
                                     '' ==============================================
                        
                            Call MudaCorStatus(arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODSTATAPONT, (.Rows - 1))
                        End If
                    Next K
                End If
                '' ==================================================
            Next J
        Next I
    End With
    
    Call GeraSubtotais
    
    For I = 1 To UBound(arrLINHAS)
        For J = 1 To arrLINHAS(I).lngQTDLINHAS
            Call CalcTotDiaProgramado(arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP)
        Next J
    Next I
    
    Call EncolheLinhasGrid

End Sub

Private Sub PegaLinhaAtual()
    
    Dim lngQTDNECFOLHAS As Long
    Dim lngINDLINHA     As Long
    Dim lngINDDIAS      As Long
    Dim lngINDOP        As Long
    
    With grdPROGOP
        lngLINHASEL = 0
        If (.Rows - 1) <= 0 Then Exit Sub
        If (.Row) = 0 Then Exit Sub
        If .Cell(flexcpText, .Row, conCOL_PRODOP_TIPO) = "H" Or _
           .Cell(flexcpText, .Row, conCOL_PRODOP_TIPO) = "I" Then lngLINHASEL = .Row
           
        If .Cell(flexcpText, .Row, conCOL_PRODOP_TIPO) = "I" And _
           Len(Trim(.Cell(flexcpText, .Row, conCOL_PRODOP_CodOP))) > 0 Then
                
                Call PopCampos(.Row)
                Call objBLBFunc.CarregaDadosGrdFilhoSemAction2Do(grdEstoque, conCOL_PRODEST_INDICE, .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_INDICE))
                
                ''lngINDLINHA = .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_INDCEARRAYLINHA)
                ''lngINDDIAS = .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_INDCEARRAYDIA)
                ''lngINDOP = .Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_INDCEARRAYOP)
                ''Call CalcNecessFolhas(arrLINHAS(lngINDLINHA).arrDIAS_LINHA(lngINDDIAS).arrOPS_INCLUSAS(lngINDOP).lngCODOP, arrLINHAS(lngINDLINHA).arrDIAS_LINHA(lngINDDIAS).arrOPS_INCLUSAS(lngINDOP).lngQTDOPPROGRAMADA)
                
        Else
            Call objBLBFunc.CarregaDadosGrdFilhoSemAction2Do(grdEstoque, conCOL_PRODEST_INDICE, -1)
            
            Call LimpaCamposLabel
            txtCODOP.Text = Empty
            mskDTENTREGA.Text = "__/__/____"
            txtQTDOPPROG.Text = Empty
        End If
    End With
End Sub

Private Sub EscondeLinhaExcluida()
    Dim I   As Long
    
    With grdPROGOP
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_PRODOP_Action2Do) = dacEnumUpdateAction_delete Then .RowHidden(I) = True
            If .Cell(flexcpText, I, conCOL_PRODOP_Action2Do) <> dacEnumUpdateAction_delete Then .RowHidden(I) = False
        Next I
    End With
End Sub


Private Function VerifDispLinha(strBLOCOP As String) As Boolean
    
    VerifDispLinha = True
    
    Dim lngLINHA    As Long
    
    With grdPROGOP
        lngLINHA = .FindRow(Trim(strBLOCOP), , conCOL_PRODOP_BLOCOdeOPS)
        If lngLINHA > 0 Then
            If CLng(.Cell(flexcpText, (lngLINHA - 1), conCOL_PRODOP_Disponivel)) = 0 Then VerifDispLinha = False
        End If
    End With
    
End Function

Private Sub Imprime()
    
    If (grdPROGOP.Rows - 1) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "A Programação não foi carregada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If grdPROGOP.Row <= 0 Or lngLINHASEL <= 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Selecione um Dia !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    
    Dim strNomRel       As String
    Dim boolTEMDADOS    As Boolean
    Dim strCABEC1       As String
    Dim strCABEC2       As String
    
    boolTEMDADOS = False
    
    sSql = ""
    
    sSql = sSql & "Select" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAPROG" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAENTR" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODOP" & vbCrLf
    
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_CODIGO" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
    
    sSql = sSql & "     , SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & " SGI_CADMOVPCP" & strNOMTABELA & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & ".SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODIGO    = " & objCADMOVPCP.CODIGO & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAPROG  = '" & Format(CDate(grdPROGOP.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_Data)), "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODLIN    = " & grdPROGOP.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_CODLINHA) & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODGRPLIN = " & grdPROGOP.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_CodGRPLINHA) & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_IDLINHA   = " & grdPROGOP.Cell(flexcpText, lngLINHASEL, conCOL_PRODOP_IDLINHA) & vbCrLf
    
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN"
    
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
    strNomRel = strNomRel & "_LIN.rpt"

    If Len(Trim(strNomRel)) > 0 Then
        Call objRel.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If


End Sub



Public Function PegaFechamentoTampaFuro(strIDPROD As String) As String

    If BREC10.State = 1 Then BREC10.Close
    
    PegaFechamentoTampaFuro = ""
    
    If Len(Trim(strIDPROD)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       LFTF.SGI_COD" & vbCrLf
    sSql = sSql & "      ,FECH.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO                PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO           LIMP" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO_FECHTPFR  LFTF" & vbCrLf
    sSql = sSql & "      ,SGI_CADFECHAM                 FECH" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = " & strIDPROD & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL     = LIMP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODLINPROD = LIMP.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "   And LIMP.SGI_FILIAL     = LFTF.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LIMP.SGI_CODIGO     = LFTF.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And LFTF.SGI_FILIAL     = FECH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LFTF.SGI_COD        = FECH.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "Order By LFTF.SGI_COD"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then PegaFechamentoTampaFuro = Trim(BREC10!SGI_DESCRI)
    BREC10.Close
    
End Function

Private Sub ConfGrd()

    With grdEstoque

       .Cols = conColumnsIn_PRODEST
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_PRODEST_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_PRODEST_FOLHAUSADA) = ""
       .ColDataType(conCOL_PRODEST_FOLHAUSADA) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_PRODEST_IDPROD) = ""
       .ColDataType(conCOL_PRODEST_IDPROD) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODEST_CODPROD) = ""
       .ColDataType(conCOL_PRODEST_CODPROD) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODEST_CODCAPAC) = ""
       .ColDataType(conCOL_PRODEST_CODCAPAC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODEST_CAPAC) = ""
       .ColDataType(conCOL_PRODEST_CAPAC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODEST_CODFOLHAUSADA) = ""
       .ColDataType(conCOL_PRODEST_CODFOLHAUSADA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODEST_DESCFOLHAUSADA) = ""
       .ColDataType(conCOL_PRODEST_DESCFOLHAUSADA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODEST_ESPESS) = ""
       .ColDataType(conCOL_PRODEST_ESPESS) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_PRODEST_LARG) = ""
       .ColDataType(conCOL_PRODEST_LARG) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_PRODEST_COMP) = ""
       .ColDataType(conCOL_PRODEST_COMP) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_PRODEST_QTDECORP) = ""
       .ColDataType(conCOL_PRODEST_QTDECORP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODEST_PERDPRODC) = ""
       .ColDataType(conCOL_PRODEST_PERDPRODC) = flexDTCurrency

       .Cell(flexcpData, 0, conCOL_PRODEST_QTDEFOLHAS) = ""
       .ColDataType(conCOL_PRODEST_QTDEFOLHAS) = flexDTLong

       .Cell(flexcpData, 0, conCOL_PRODEST_PESO) = ""
       .ColDataType(conCOL_PRODEST_PESO) = flexDTCurrency

       .Cell(flexcpData, 0, conCOL_PRODEST_QTDELATAS) = ""
       .ColDataType(conCOL_PRODEST_QTDELATAS) = flexDTLong

       .Cell(flexcpData, 0, conCOL_PRODEST_INDICE) = ""
       .ColDataType(conCOL_PRODEST_INDICE) = flexDTString

       .Cell(flexcpData, 0, conCOL_PRODEST_NEFOLHAS) = ""
       .ColDataType(conCOL_PRODEST_NEFOLHAS) = flexDTLong

       .Cell(flexcpData, 0, conCOL_PRODEST_LINHA) = ""
       .ColDataType(conCOL_PRODEST_LINHA) = flexDTLong

       .ColWidth(conCOL_PRODEST_FOLHAUSADA) = 300
       .ColWidth(conCOL_PRODEST_IDPROD) = 0
       .ColWidth(conCOL_PRODEST_CODPROD) = 0
       .ColWidth(conCOL_PRODEST_CODCAPAC) = 0
       .ColWidth(conCOL_PRODEST_CAPAC) = 0
       .ColWidth(conCOL_PRODEST_CODFOLHAUSADA) = 0
       .ColWidth(conCOL_PRODEST_DESCFOLHAUSADA) = 3000
       .ColWidth(conCOL_PRODEST_ESPESS) = 0
       .ColWidth(conCOL_PRODEST_LARG) = 0
       .ColWidth(conCOL_PRODEST_COMP) = 0
       .ColWidth(conCOL_PRODEST_QTDECORP) = 0
       .ColWidth(conCOL_PRODEST_PERDPRODC) = 0
       .ColWidth(conCOL_PRODEST_QTDEFOLHAS) = 1300
       .ColWidth(conCOL_PRODEST_PESO) = 0
       .ColWidth(conCOL_PRODEST_QTDELATAS) = 1300
       .ColWidth(conCOL_PRODEST_CODOP) = 1300
       .ColWidth(conCOL_PRODEST_INDICE) = 2000
       .ColWidth(conCOL_PRODEST_NEFOLHAS) = 1800
       .ColWidth(conCOL_PRODEST_LINHA) = 1800
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
       
    End With
    
End Sub


Private Sub PegaEstoqueProdutos(strIDPROD As String, strCODOP As String)
        
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PROD.SGI_IDPRODUTO    " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODIGO       " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO    As SGI_DESCPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODLINPROD   " & vbCrLf
    sSql = sSql & "      ,LINH.SGI_CODIGO       As SGI_IDLINHA" & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_CODMEDCORT   As SGI_CODFOLHA" & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_EXPESS       " & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_LARGUR       " & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_COMPRI       " & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_QTDECORPOS   " & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_PERDPROC     " & vbCrLf
    sSql = sSql & "      ,DIMC.SGI_DESCORTE     " & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO        PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO   LINH" & vbCrLf
    sSql = sSql & "      ,SGI_MEDCORTELINHA     MEDC" & vbCrLf
    sSql = sSql & "      ,SGI_CADDIMCORTE       DIMC" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO    = " & strIDPROD & vbCrLf
    
    sSql = sSql & "   And LINH.SGI_FILIAL       = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINH.SGI_CODLIN       = PROD.SGI_CODLINPROD" & vbCrLf
    
    sSql = sSql & "   And MEDC.SGI_FILIAL       = LINH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And MEDC.SGI_CODIGO       = LINH.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And DIMC.SGI_FILIAL       = MEDC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And DIMC.SGI_CODIGO       = MEDC.SGI_CODMEDCORT"
    

    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then
        With grdEstoque
            Do While Not BREC11.EOF()
            
                .AddItem "" & vbTab & _
                         BREC11!SGI_IDPRODUTO & vbTab & _
                         BREC11!SGI_CODIGO & vbTab & _
                         BREC11!SGI_CODLINPROD & vbTab & _
                         BREC11!SGI_IDLINHA & vbTab & _
                         BREC11!SGI_CODFOLHA & vbTab & _
                         BREC11!SGI_DESCORTE & vbTab & _
                         Format(BREC11!SGI_EXPESS, "#,####0.0000") & vbTab & _
                         Format(BREC11!SGI_LARGUR, "#,##0.00") & vbTab & _
                         Format(BREC11!SGI_COMPRI, "#,##0.00") & vbTab & _
                         BREC11!SGI_QTDECORPOS & vbTab & _
                         Format(BREC11!SGI_PERDPROC, "#,##0.00") & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         strCODOP
                                         
                If Not IsNull(BREC11!SGI_QTDECORPOS) Then
                    Call SomaPesoSaldo(Str(BREC11!SGI_IDPRODUTO), BREC11!SGI_CODFOLHA, (.Rows - 1), BREC11!SGI_QTDECORPOS)
                End If
                                         
                         
                BREC11.MoveNext
            Loop
        End With
    End If
    BREC11.Close
End Sub

Private Sub SomaPesoSaldo(strID, strCODCORTE As String, lngLINEST As Long, lngQTDCORPOS As Long)

        Dim dblENTRADA      As Double
        
        Dim lngQTDFOLHAS    As Long
        Dim lngQTDLATAS     As Long
        
        '' Entrada
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "      CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "     ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_CONFPESO)         As SGI_SALDOPESO" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_QTDEFOLHASREC)    As SGI_SALDOFOLHAS" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADRECROTLIT_IT ITEN" & vbCrLf
        sSql = sSql & "      ,SGI_CADRECROTLIT    CABEC" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And ITEN.SGI_IDPRODUTO = " & strID & vbCrLf
        sSql = sSql & "   And ITEN.SGI_STATUS    = 'REC'" & vbCrLf       '' Entrada no Destino
        sSql = sSql & "   And ITEN.SGI_CODCODTE  = " & strCODCORTE & vbCrLf
        
        sSql = sSql & "   And CABEC.SGI_FILIAL   = ITEN.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CABEC.SGI_CODIGO   = ITEN.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO    = CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "Group By CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "        ,CLIE.SGI_RAZAOSOC"

        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
            Do While Not BREC10.EOF()
                With grdEstoque
                
                    dblENTRADA = (dblENTRADA + BREC10!SGI_SALDOPESO)
                    
                    lngQTDFOLHAS = (lngQTDFOLHAS + BREC10!SGI_SALDOFOLHAS)
                    lngQTDLATAS = (lngQTDCORPOS * lngQTDFOLHAS)
                    
                    .Cell(flexcpText, lngLINEST, conCOL_PRODEST_PESO) = Format(dblENTRADA, "#,##0.00")
                    .Cell(flexcpText, lngLINEST, conCOL_PRODEST_QTDEFOLHAS) = lngQTDFOLHAS
                    .Cell(flexcpText, lngLINEST, conCOL_PRODEST_QTDELATAS) = lngQTDLATAS
                    
                End With
                BREC10.MoveNext
            Loop
        End If
        BREC10.Close
End Sub

Private Function SomaSaldos(strID, strCODCORTE As String, lngQTDCORPOS As Long) As String

        SomaSaldos = ""
        
        Dim dblENTRADA      As Double
        
        Dim lngQTDFOLHAS    As Long
        Dim lngQTDLATAS     As Long
        
        '' Entrada
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "      CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "     ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_CONFPESO)         As SGI_SALDOPESO" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_QTDEFOLHASREC)    As SGI_SALDOFOLHAS" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADRECROTLIT_IT ITEN" & vbCrLf
        sSql = sSql & "      ,SGI_CADRECROTLIT    CABEC" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And ITEN.SGI_IDPRODUTO = " & strID & vbCrLf
        sSql = sSql & "   And ITEN.SGI_STATUS    = 'REC'" & vbCrLf       '' Entrada no Destino
        sSql = sSql & "   And ITEN.SGI_CODCODTE  = " & strCODCORTE & vbCrLf
        
        sSql = sSql & "   And CABEC.SGI_FILIAL   = ITEN.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CABEC.SGI_CODIGO   = ITEN.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO    = CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "Group By CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "        ,CLIE.SGI_RAZAOSOC"

        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
            Do While Not BREC10.EOF()
                With grdEstoque
                
                    dblENTRADA = (dblENTRADA + BREC10!SGI_SALDOPESO)
                    
                    lngQTDFOLHAS = (lngQTDFOLHAS + BREC10!SGI_SALDOFOLHAS)
                    lngQTDLATAS = (lngQTDCORPOS * lngQTDFOLHAS)

                End With
                BREC10.MoveNext
            Loop
            SomaSaldos = Format(dblENTRADA, "#,##0.00") & "|" & Str(lngQTDFOLHAS) & "|" & lngQTDLATAS
        End If
        BREC10.Close
        
End Function



Private Function CalcSaldoLinha(strDTDIA As String, lngQTDDPROG As Long) As Boolean

    CalcSaldoLinha = False
    
    Dim lngTOTISP       As Long
    Dim lngTOTPROG      As Long
    Dim lngTOTJAPROG    As Long
    Dim lngLINHADIAINI  As Long
    
    With grdPROGOP
        lngLINHADIAINI = .FindRow(strDTDIA, , conCOL_PRODOP_BLOCOdeOPS)
        If lngLINHADIAINI > -1 Then
                
                lngTOTISP = CLng(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_Capacidade))
                
                lngTOTPROG = 0
                ''If Len(Trim(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_TotalProgramado))) > 0 Then lngTOTPROG = CLng(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_TotalProgramado))
                
                lngTOTJAPROG = (lngTOTISP - (lngTOTPROG + lngQTDDPROG))
                
''                If lngTOTJAPROG < 0 Then
''                    MsgBox "ATENÇÃO" & vbCrLf & _
''                           "A capacidade já esta estourada não e permitido acrescentar mais OP's !!!", vbOKOnly + vbExclamation, "Aviso"
''                    Exit Function
''                ElseIf lngTOTJAPROG < 0 Then
''                    MsgBox "ATENÇÂO" & vbCrLf & _
''                           "Capacidade da Linha esta zerada !!!, Impossivel incluir OP's", vbOKOnly + vbExclamation, "Aviso"
''                    Exit Function
''                End If
''
        End If
    End With
    
    CalcSaldoLinha = True
    
End Function

Private Function PegaDispLinha(strDTDIA As String) As Long

    '' Pegando a Dipnibidade da Linha
    
    PegaDispLinha = 0
    
    Dim lngTOTISP       As Long
    Dim lngTOTPROG      As Long
    Dim lngTOTJAPROG    As Long
    Dim lngLINHADIAINI  As Long
    
    With grdPROGOP
        lngLINHADIAINI = .FindRow(strDTDIA, , conCOL_PRODOP_BLOCOdeOPS)
        If lngLINHADIAINI > -1 Then PegaDispLinha = CLng(.Cell(flexcpText, (lngLINHADIAINI - 1), conCOL_PRODOP_Capacidade))
    End With
    
End Function



Private Function PegaDataValida(dtDATAFLUT As Date, lngGRPLINHASEL As Long, lngCODLINHASEL As Long) As Date

        PegaDataValida = Empty
        
        Dim I   As Long
        Dim J   As Long
        
        '' Pegar Proxima Data Válida
        For I = 1 To UBound(arrLINHAS)
            If arrLINHAS(I).lngQTDLINHAS > 0 Then
                If arrLINHAS(I).lngCODGRPLIN = lngGRPLINHASEL And _
                   arrLINHAS(I).lngCodLinha = lngCODLINHASEL Then
                   For J = 1 To arrLINHAS(I).lngQTDLINHAS
                        If arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG > dtDATAFLUT Then
                            PegaDataValida = arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG
                            Exit Function
                        End If
                   Next J
                End If
            End If
        Next I

End Function

Private Function PegaTotalJaProgOP(lngCODOP As Long) As Long

    Dim I               As Long
    Dim J               As Long
    Dim K               As Long
    Dim lngOPPROG       As Long
    Dim lngOPJAPROG     As Long
    
    lngOPPROG = 0
    For I = 1 To UBound(arrLINHAS)
        For J = 1 To arrLINHAS(I).lngQTDLINHAS
            For K = 1 To arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS
                If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngAction2Do <> dacEnumUpdateAction_delete Then
                    If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP = lngCODOP Then lngOPPROG = lngOPPROG + (arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngQTDOPPROGRAMADA)
                End If
            Next K
        Next J
    Next I
    
    lngOPJAPROG = (lngOPPROG)
    PegaTotalJaProgOP = lngOPJAPROG

End Function

Private Function PegaBlocoOP(dtDTBLOCO As Date, lngCODGRP As Long, lngCODLIN As Long) As String
    
    PegaBlocoOP = ""
    
    Dim I       As Long
    Dim J       As Long
    
    For I = 1 To UBound(arrLINHAS)
        If arrLINHAS(I).lngCODGRPLIN = lngCODGRP And _
           arrLINHAS(I).lngCodLinha = lngCODLIN Then
            For J = 1 To arrLINHAS(I).lngQTDLINHAS
                If arrLINHAS(I).arrDIAS_LINHA(J).dtDATAPROG = dtDTBLOCO Then
                    PegaBlocoOP = Trim(arrLINHAS(I).arrDIAS_LINHA(J).strBLOCOOP)
                    Exit Function
                End If
            Next J
        End If
    Next I
    
End Function

Private Sub DeletaLinhaDaArray(lngCODOP As Long)
    If lngCODOP = 0 Then Exit Sub
    
    Dim I   As Long
    Dim J   As Long
    Dim K   As Long
    
    If UBound(arrLINHAS) > 0 Then
        For I = 1 To UBound(arrLINHAS)
            For J = 1 To arrLINHAS(I).lngQTDLINHAS
                For K = 1 To arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS
                    If arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngCODOP = lngCODOP Then arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(K).lngAction2Do = dacEnumUpdateAction_delete
                Next K
            Next J
        Next I
    End If
    
End Sub

Private Sub PopCampos(lngLINHA As Long)
    With grdPROGOP
        txtCODOP.Text = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_CodOP)
        lblCODROT.Caption = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_CodRotulo)
        lblDESCROT.Caption = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_DescRotulo)
        mskDTENTREGA.Text = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_DataEntrega)
        lblNECK.Caption = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_NECK)
        lblFECH.Caption = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_FECH)
        lblCOMP.Caption = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_COMP)
        lblQTDOP.Caption = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_QtdeOP)
        txtQTDOPPROG.Text = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_QtdeOPProgramada)
        lblSTATUS.Caption = .Cell(flexcpText, lngLINHA, conCOL_PRODOP_Status)
    End With
End Sub

Private Sub CalcNecessFolhas(lngCODOP As Long, lngQTDEOP As Long)
    
    If lngCODOP = 0 Then Exit Sub
    If lngQTDEOP = 0 Then Exit Sub
    
    Dim I           As Long
    Dim lngQTDCORP  As Long
    Dim lngQTDNECF  As Long
    
    With grdEstoque
        For I = 1 To (.Rows - 1)
            If CLng(.Cell(flexcpText, I, conCOL_PRODEST_CODOP)) = lngCODOP Then
            
                lngQTDCORP = 0
                lngQTDNECF = 0
                If Len(Trim(.Cell(flexcpText, I, conCOL_PRODEST_QTDECORP))) > 0 Then lngQTDCORP = .Cell(flexcpText, I, conCOL_PRODEST_QTDECORP)
                If lngQTDCORP > 0 Then lngQTDNECF = (lngQTDEOP / lngQTDCORP)
                If lngQTDNECF > 0 Then .Cell(flexcpText, I, conCOL_PRODEST_NEFOLHAS) = lngQTDNECF
                
            End If
        Next I
    End With
    
End Sub

Private Sub PegaLinhaEstoque()
    With grdEstoque
        
        lngLINHASELEST = 0
        If .Row = 0 Then Exit Sub
        If (.Rows - 1) = 0 Then Exit Sub
        lngLINHASELEST = .Row
        
    End With
End Sub

Private Sub Add_ElArray(lngROWI As Long _
                      , lngROWJ As Long _
                      , strCODOP As String _
                      , lngID_LINHA As Long _
                      , dtDATASEL As Date _
                      , lngGRPLINHASEL As Long _
                      , lngCODLINHASEL As Long _
                      , strBLOCOSEL As String _
                      , dtDTENTREGA As Date _
                      , dtDTENTREGAANT As Date _
                      , lblQTDOP As Long _
                      , lngQTDOPPROG As Long _
                      , boolFRACIONA As Boolean _
                      , lngIDINTERNO As Long _
                      , lngAction2Do _
                      , lngEXPLIN As Long)

    Dim I               As Long
    Dim J               As Long
    Dim lngQTDOPS       As Long
    Dim lngQTDFOLHAS    As Long
    Dim lngQTDNECFOLHAS As Long
    Dim strCAMPOS       As String
    Dim arrCAMPOS()     As String
    
    I = lngROWI
    J = lngROWJ
    
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       ORDP.SGI_CODIGO      As SGI_CODOP" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_STATUS      " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_CODPED      " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_IDPRODUTO   " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_IDPAI       " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_PROGRAMADO  " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_CODPROD     " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_QTDE        As SGI_QTDEOP" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_SALDO       " & vbCrLf
    sSql = sSql & "      ,LINH.SGI_CODIGO      As SGI_CODLINHA" & vbCrLf
    sSql = sSql & "      ,LINH.SGI_CODLIN      " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_NECKIN      " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO   As SGI_DESCPROD" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMTABELA & "      ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO           PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO      LINH" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL              = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO              = " & Trim(strCODOP) & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL              = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO           = ORDP.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And LINH.SGI_FILIAL              = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINH.SGI_CODLIN              = PROD.SGI_CODLINPROD" & vbCrLf

    BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC5.EOF() Then
    
        '' Pegando as OP's ja inseridas
        lngQTDOPS = (arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS + 1)
        
        '' Incluindo Mais Uma Linha no Array()
        ReDim Preserve arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(1 To lngQTDOPS) As OPS_INCLUSAS
    
        arrLINHAS(I).arrDIAS_LINHA(J).lngQTDOPS = lngQTDOPS
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngID_LINHA = lngID_LINHA
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).dtDATAPROG = dtDATASEL
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngCIDGRPLIN = lngGRPLINHASEL
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngCODLINA = lngCODLINHASEL
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).strBLOCOOP = strBLOCOSEL
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngAction2Do = lngAction2Do
    
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngCODOP = CLng(strCODOP)
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).strCODROTULO = BREC5!SGI_CODPROD
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).strDESCROTULO = BREC5!SGI_DESCPROD
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).dtDATAENTREGA = dtDTENTREGA
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).dtDATAENTREGAORIGINAL = dtDTENTREGAANT
        
        If BREC5!SGI_STATUS = 1 Then
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngQTDOPORIGINAL = BREC5!SGI_SALDO
        Else
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngQTDOPORIGINAL = lblQTDOP
        End If
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngQTDOPPROGRAMADA = lngQTDOPPROG
    
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).intNECK = BREC5!SGI_NECKIN
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngCODOPBKP = CLng(strCODOP)
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngIDOP = BREC5!SGI_IDPAI
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngCODPED = BREC5!SGI_CODPED
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngIDPRODUTO = BREC5!SGI_IDPRODUTO
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngCODSTATUS = BREC5!SGI_STATUS
        
        
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngCODSTATUSORIGINAL = BREC5!SGI_STATUS
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).strINDICE = Trim(strCODOP) & Trim(Str(BREC5!SGI_IDPRODUTO)) & Trim(strBLOCOSEL)
    
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).strTIPO = "I"
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).intSELECIONADO = 1
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngIDINTERNO = lngIDINTERNO
        
        '' Grando os Indeces da Linha do Array
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngIDARRAYLINHA = I
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngIDARRAYDIA = J
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngIDARRAYOP = lngQTDOPS
        arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngEXPLIN = lngEXPLIN
    
        If boolFRACIONA = True Then arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).intFRACIONADA = 1
        If boolFRACIONA = False Then arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).intFRACIONADA = 0
            
        '' =============================================================================
        '' Incluindo as Folhas que serão usadas
        Call AddFolhArray(I _
                        , J _
                        , lngQTDOPS _
                        , BREC5!SGI_IDPRODUTO _
                        , CLng(strCODOP) _
                        , strBLOCOSEL _
                        , lngQTDOPPROG)
    
    End If
    BREC5.Close
    
End Sub

Private Sub AddFolhArray(lngROWI As Long _
                       , lngROWJ As Long _
                       , lngQTDOPS As Long _
                       , lngIDPROD As Long _
                       , lngCODOP As Long _
                       , strBLOCOOP As String _
                       , lngQTDOPPROGRAMADA As Long)

    Dim arrCAMPOS()         As String
    Dim strCAMPOS           As String
    Dim lngQTDFOLHAS        As Long
    Dim lngQTDNECFOLHAS     As Long
    Dim I                   As Long
    Dim J                   As Long
    
    I = lngROWI
    J = lngROWJ
        
    '' =============================================================================
    '' Incluindo as Folhas que serão usadas
    lngQTDFOLHAS = 0
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PROD.SGI_IDPRODUTO    " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODIGO       As SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO    As SGI_DESCPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODLINPROD   " & vbCrLf
    sSql = sSql & "      ,LINH.SGI_CODIGO       As SGI_IDLINHA" & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_CODMEDCORT   As SGI_CODFOLHA" & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_EXPESS       " & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_LARGUR       " & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_COMPRI       " & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_QTDECORPOS   " & vbCrLf
    sSql = sSql & "      ,MEDC.SGI_PERDPROC     " & vbCrLf
    sSql = sSql & "      ,DIMC.SGI_DESCORTE     " & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO        PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO   LINH" & vbCrLf
    sSql = sSql & "      ,SGI_MEDCORTELINHA     MEDC" & vbCrLf
    sSql = sSql & "      ,SGI_CADDIMCORTE       DIMC" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO    = " & lngIDPROD & vbCrLf
    
    sSql = sSql & "   And LINH.SGI_FILIAL       = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINH.SGI_CODLIN       = PROD.SGI_CODLINPROD" & vbCrLf
    
    sSql = sSql & "   And MEDC.SGI_FILIAL       = LINH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And MEDC.SGI_CODIGO       = LINH.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And DIMC.SGI_FILIAL       = MEDC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And DIMC.SGI_CODIGO       = MEDC.SGI_CODMEDCORT"

    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then
        
        Do While Not BREC11.EOF()
            
            lngQTDFOLHAS = (lngQTDFOLHAS + 1)
            ReDim Preserve arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(1 To lngQTDFOLHAS) As OPS_INCLUSAS_FOLHAS_USADAS

            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngFOLHAUSADA = 2
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngCODOP = lngCODOP
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).strCODPROD = BREC11!SGI_CODPROD
            
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngIDLIN = BREC11!SGI_IDLINHA
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngCODLIN = BREC11!SGI_CODLINPROD
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).strDESCFOLHAUSADA = BREC11!SGI_DESCORTE
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngCODFOLHAUSADA = BREC11!SGI_CODFOLHA
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngIDPROD = BREC11!SGI_IDPRODUTO
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).strINDICE = Trim(Str(lngCODOP)) & Trim(Str(lngIDPROD)) & Trim(strBLOCOOP)
            arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngLINHA = lngQTDFOLHAS
            
            If Not IsNull(BREC11!SGI_EXPESS) Then arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblESPESS = BREC11!SGI_EXPESS
            If Not IsNull(BREC11!SGI_LARGUR) Then arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblLARG = BREC11!SGI_LARGUR
            If Not IsNull(BREC11!SGI_COMPRI) Then arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblCOMP = BREC11!SGI_COMPRI
            If Not IsNull(BREC11!SGI_QTDECORPOS) Then arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDECORP = BREC11!SGI_QTDECORPOS
            If Not IsNull(BREC11!SGI_PERDPROC) Then arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblPERDPRODC = BREC11!SGI_PERDPROC
                 
            If Not IsNull(BREC11!SGI_QTDECORPOS) Then
            
                lngQTDNECFOLHAS = 0
                If BREC11!SGI_QTDECORPOS > 0 Then lngQTDNECFOLHAS = (lngQTDOPPROGRAMADA / BREC11!SGI_QTDECORPOS)
                arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngNECEFOLHAS = lngQTDNECFOLHAS
            
                arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblPESO = 0
                arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDEFOLHAS = 0
                arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDELATAS = 0
            
                strCAMPOS = SomaSaldos(Str(BREC11!SGI_IDPRODUTO), BREC11!SGI_CODFOLHA, BREC11!SGI_QTDECORPOS)
                If Len(Trim(strCAMPOS)) > 0 Then
                    arrCAMPOS = Split(strCAMPOS, "|")
                    arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).dblPESO = CDbl(arrCAMPOS(0))
                    arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDEFOLHAS = CLng(arrCAMPOS(1))
                    arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).arrFOLHAS_USADAS(lngQTDFOLHAS).lngQTDELATAS = CLng(arrCAMPOS(2))
                End If
            End If
                 
            BREC11.MoveNext
        Loop
    End If
    BREC11.Close
    arrLINHAS(I).arrDIAS_LINHA(J).arrOPS_INCLUSAS(lngQTDOPS).lngQTDFOLHASUSADAS = lngQTDFOLHAS
    '' =============================================================================

End Sub
