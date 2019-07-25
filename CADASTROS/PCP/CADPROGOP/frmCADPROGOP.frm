VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADPROGOP 
   Caption         =   "Programação de Produção"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   18555
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   18120
      Top             =   8400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Indice"
      Height          =   255
      Left            =   18120
      TabIndex        =   31
      Top             =   8880
      Width           =   375
   End
   Begin VB.Frame Frame8 
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
      Height          =   735
      Left            =   15720
      TabIndex        =   26
      Top             =   960
      Width           =   2775
      Begin MSMask.MaskEdBox mskDataI 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataF 
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   1320
         TabIndex        =   29
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "[ Periodo Completo ]"
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
      Height          =   735
      Left            =   11400
      TabIndex        =   22
      Top             =   960
      Width           =   4215
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Escolhe  Dias "
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
         Index           =   3
         Left            =   2520
         TabIndex        =   30
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   " Dia"
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
         Left            =   1800
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optPeriodo 
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
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optPeriodo 
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
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   360
         Width           =   735
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
      Height          =   735
      Left            =   8640
      TabIndex        =   19
      Top             =   960
      Width           =   2655
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
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
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
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid grdAUX 
      Height          =   255
      Left            =   18120
      TabIndex        =   18
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
      _cx             =   450
      _cy             =   450
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
   Begin VB.Frame Frame5 
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
      Height          =   1455
      Left            =   5280
      TabIndex        =   17
      Top             =   7680
      Width           =   12735
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ OP's Programadas ]"
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
      Height          =   6015
      Left            =   5280
      TabIndex        =   14
      Top             =   1680
      Width           =   13215
      Begin VB.CommandButton Command6 
         Height          =   300
         Left            =   12840
         Picture         =   "frmCADPROGOP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Realoca a OP para o Proximo dia"
         Top             =   1320
         Width           =   300
      End
      Begin VB.CommandButton Command5 
         Height          =   300
         Left            =   12840
         Picture         =   "frmCADPROGOP.frx":2EFA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Carrega OP's não apontadas, e realoca no dia vigente"
         Top             =   960
         Width           =   300
      End
      Begin VB.CommandButton Command4 
         Height          =   300
         Left            =   12840
         Picture         =   "frmCADPROGOP.frx":4BF4
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   2400
         Width           =   300
      End
      Begin VB.CommandButton Command9 
         Height          =   300
         Left            =   12840
         Picture         =   "frmCADPROGOP.frx":7AEE
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   12840
         Picture         =   "frmCADPROGOP.frx":7C38
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   2040
         Width           =   300
      End
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   12840
         Picture         =   "frmCADPROGOP.frx":AB32
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdPROGRAMACAO 
         Height          =   5655
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   12615
         _cx             =   22251
         _cy             =   9975
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
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
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
      Caption         =   "[ Dias de Programação ]"
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
      Height          =   5415
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   5175
      Begin VSFlex8LCtl.VSFlexGrid grdDIASPROG 
         Height          =   5055
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4935
         _cx             =   8705
         _cy             =   8916
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
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
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
      Caption         =   "[ Linha ]"
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
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   5295
      Begin VSFlex8LCtl.VSFlexGrid grdLinha 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5055
         _cx             =   8916
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
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
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
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   8535
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
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmCADPROGOP.frx":AC7C
         Left            =   1920
         List            =   "frmCADPROGOP.frx":AC7E
         TabIndex        =   7
         Text            =   "cboMes"
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   4320
         TabIndex        =   6
         Text            =   "cboAno"
         Top             =   240
         Width           =   1695
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
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18495
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
         Picture         =   "frmCADPROGOP.frx":AC80
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCADPROGOP.frx":B1B2
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADPROGOP.frx":B2B4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
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
         Left            =   17520
         Picture         =   "frmCADPROGOP.frx":B3B6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Registro"
         Top             =   240
         Width           =   855
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   255
      Left            =   27840
      TabIndex        =   35
      Top             =   12720
      Visible         =   0   'False
      Width           =   255
      _cx             =   450
      _cy             =   450
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
   Begin VSFlex8LCtl.VSFlexGrid grdEXC 
      Height          =   255
      Left            =   18120
      TabIndex        =   36
      Top             =   7800
      Visible         =   0   'False
      Width           =   255
      _cx             =   450
      _cy             =   450
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
Attribute VB_Name = "frmCADPROGOP"
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
Dim strCAPTION          As String
Dim strDTINICIAL        As String
Dim strDTFINAL          As String

Dim arrPROGROP          As Variant
Dim arrPROGROPEXC       As Variant

Dim lngSCROOLMANT       As Long
Dim lngSCROOLDANT       As Long
Dim lngSCROOLPANT       As Long

Dim boolINS_LINHA       As Boolean
Dim boolFRACIONA        As Boolean
Dim boolREMOPSEL        As Boolean

Const conCOL_LINHA_CODLIN                           As Integer = 0
Const conCOL_LINHA_CODGRPLIN                        As Integer = 1
Const conCOL_LINHA_DESCLIN                          As Integer = 2
Const conCOL_LINHA_MES                              As Integer = 3
Const conCOL_LINHA_ANO                              As Integer = 4
Const conCOL_LINHA_IDINTERNO                        As Integer = 5
Const conCOL_LINHA_CAPACLINHA                       As Integer = 6
Const conCOL_LINHA_TOTALPROG                        As Integer = 7
Const conCOL_LINHA_TOTALDISP                        As Integer = 8
Const conCOL_LINHA_FormatString                     As String = "=CodLin|CodGrpLin|Descrição da Linha|MES|ANO|IDINTERNO|Capacidade|Programado|Disponivel"
Const conColumnsIn_LINHA                            As Integer = 9

Const conCOL_DTP_IDPAI                              As Integer = 0
Const conCOL_DTP_DTPROG                             As Integer = 1
Const conCOL_DTP_TOTDIA                             As Integer = 2
Const conCOL_DTP_PROGAM                             As Integer = 3
Const conCOL_DTP_QTDISP                             As Integer = 4
Const conCOL_DTP_IDINTERNO                          As Integer = 5
Const conCOL_DTP_INDICEPAI                          As Integer = 6
Const conCOL_DTP_CODLIN                             As Integer = 7
Const conCOL_DTP_CODGRPLIN                          As Integer = 8
Const conCOL_DTP_SELDIA                             As Integer = 9
Const conCOL_DTP_FormatString                       As String = "=IDPAI|Data Valida|Qtde. do Dia|Qtde. Progr.|Qtde. Disp.|IDINTERNO|INDICEPAI|CODLIN|CODGRPLIN|  "
Const conColumnsIn_DTP                              As Integer = 10

Const conCOL_PROG_IDDIA                             As Integer = 0
Const conCOL_PROG_CODOP                             As Integer = 1
Const conCOL_PROG_CODPROD                           As Integer = 2
Const conCOL_PROG_DESCPROD                          As Integer = 3
Const conCOL_PROG_DTENTREGA                         As Integer = 4
Const conCOL_PROG_QtdeOP                            As Integer = 5
Const conCOL_PROG_QtdeOPProgramada                  As Integer = 6
Const conCOL_PROG_QtdeReal                          As Integer = 7
Const conCOL_PROG_CODPED                            As Integer = 8
Const conCOL_PROG_IDPRODUTO                         As Integer = 9
Const conCOL_PROG_IDINTERNOOP                       As Integer = 10
Const conCOL_PROG_CODINTENO                         As Integer = 11
Const conCOL_PROG_CODLIN                            As Integer = 12
Const conCOL_PROG_CODGRPLIN                         As Integer = 13
Const conCOL_PROG_IDLINHA                           As Integer = 14
Const conCOL_PROG_DTENTREGABKP                      As Integer = 15
Const conCOL_PROG_CODOPBKP                          As Integer = 16
Const conCOL_PROG_CODPEDBKP                         As Integer = 17
Const conCOL_PROG_IDPRODBKP                         As Integer = 18
Const conCOL_PROG_IDINTERNOOPBKP                    As Integer = 19
Const conCOL_PROG_INDICEPROG                        As Integer = 20
Const conCOL_PROG_CODSTATUS                         As Integer = 21
Const conCOL_PROG_CODSTATUSBKP                      As Integer = 22
Const conCOL_PROG_FRACIONADO                        As Integer = 23
Const conCOL_PROG_DTPROG                            As Integer = 24
Const conCOL_PROG_NECKIN                            As Integer = 25
Const conCOL_PROG_FECH                              As Integer = 26
Const conCOL_PROG_COMP                              As Integer = 27
Const conCOL_PROG_Action2Do                         As Integer = 28
Const conCOL_PROG_Marca                             As Integer = 29
Const conCOL_PROG_CODSTATUSAPONT                    As Integer = 30
Const conCOL_PROG_DESCSTATSAPONT                    As Integer = 31
Const conCOL_PROG_ORDEM                             As Integer = 32
Const conCOL_PROG_TIMESTAMP                         As Integer = 33
Const conCOL_PROG_FormatString                      As String = "=INDICE|Código OP|Cód. Produto|Descrição do Produto|Dt.Entrega|Qtde.OP|Qtde.Progr.|Qtde.Real|CODPED|IDPRODUTO|IDINTERNOOP|CODINTENO|CODLIN|CODGRPLIN|IDLINHA|DTENTREGABKP|CODOPBKP|CODPEDBKP|IDPRODBKP|IDINTERNOOPBKP|INDICEPROG|CODSTATUS|CODSTATUSBKP|FRACIONADO|DTPROG|NECK|FECH|COMP|Action2Do|  |CODSTATUSAPONT|Status|ORDEM|TIMESTAMP"
Const conColumnsIn_PROG                             As Integer = 34

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    cTipOper = "A"
    
    Dim lngROLLINANT    As Long
    Dim lngROLDIAANT    As Long
    Dim lngROLOPANT     As Long
    
    lngROLLINANT = grdLinha.Row
    lngROLDIAANT = grdDIASPROG.Row
    lngROLOPANT = grdPROGRAMACAO.Row
    
    Call IniciaForm

    If grdLinha.Row > 0 Then grdLinha.Row = lngROLLINANT
    If grdLinha.RowSel > 0 Then grdLinha.RowSel = lngROLLINANT
    grdLinha.TopRow = lngSCROOLMANT

    If lngROLDIAANT > 0 Then
       grdDIASPROG.Row = lngROLDIAANT
       Call MarcaLinha(grdDIASPROG, lngROLDIAANT)
    End If
    If lngROLDIAANT > 0 Then grdDIASPROG.RowSel = lngROLDIAANT
    grdDIASPROG.TopRow = lngSCROOLDANT
    
    If grdPROGRAMACAO.Row > 0 Then grdPROGRAMACAO.Row = lngROLOPANT
    If grdPROGRAMACAO.RowSel > 0 Then grdPROGRAMACAO.RowSel = lngROLOPANT
    grdPROGRAMACAO.TopRow = lngSCROOLPANT

End Sub

Private Sub cmdCARREGA_Click()
    '' Carregando das Linhas com as OP's
    Call CarregaLinha
End Sub

Private Sub cmdImpressao_Click()
    If cTipOper = "I" Then
        MsgBox "Somente pode ser impresso no modo de Consulta ou Alteração !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If optPeriodo(0).Value = True Then
        If ConsisteCampos = False Then Exit Sub
    End If
    Call Imprime
End Sub

Private Sub CmdSalva_Click()
    
    Call objBLBFunc.RemoveLinhaVazia(grdPROGRAMACAO, conCOL_PROG_CODOP)
    '' Refazendo os Indices
    Call Command1_Click
    
    If Valida_Campos = False Then Exit Sub
    
    Dim i               As Long
    Dim lngROLLINANT    As Long
    Dim lngROLDIAANT    As Long
    Dim lngROLOPANT     As Long
    
    lngROLLINANT = grdLinha.Row
    lngROLDIAANT = grdDIASPROG.Row
    lngROLOPANT = grdPROGRAMACAO.Row
    
    If cTipOper = "I" Then objCADMOVPCP.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMTABELA, FILIAL, Linha)
    
    arrPROGROPEXC = Empty
    With grdEXC
        If (.Rows - 1) > 0 Then
            ReDim arrPROGROPEXC(1 To (.Rows - 1), 1 To 21) As String
            For i = 1 To UBound(arrPROGROPEXC)
                arrPROGROPEXC(i, 1) = .Cell(flexcpText, i, conCOL_PROG_DTPROG)
                arrPROGROPEXC(i, 2) = .Cell(flexcpText, i, conCOL_PROG_CODOP)
                arrPROGROPEXC(i, 3) = .Cell(flexcpText, i, conCOL_PROG_DTENTREGA)
                arrPROGROPEXC(i, 4) = .Cell(flexcpText, i, conCOL_PROG_CODSTATUS)
                arrPROGROPEXC(i, 5) = .Cell(flexcpText, i, conCOL_PROG_IDPRODUTO)
                arrPROGROPEXC(i, 6) = .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOP)
                arrPROGROPEXC(i, 7) = .Cell(flexcpText, i, conCOL_PROG_CODPED)
                arrPROGROPEXC(i, 8) = .Cell(flexcpText, i, conCOL_PROG_Action2Do)
                arrPROGROPEXC(i, 9) = .Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada)
                
                If .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Ignore Or _
                   .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update Or _
                   .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_delete Then
                    arrPROGROPEXC(i, 10) = .Cell(flexcpText, i, conCOL_PROG_CODINTENO)
                ElseIf .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Insert Then
                    arrPROGROPEXC(i, 10) = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMTABELA & "_IDPROG", FILIAL, Linha)
                End If
                
                arrPROGROPEXC(i, 13) = .Cell(flexcpText, i, conCOL_PROG_FRACIONADO)
                arrPROGROPEXC(i, 14) = .Cell(flexcpText, i, conCOL_PROG_IDLINHA)
                arrPROGROPEXC(i, 15) = .Cell(flexcpText, i, conCOL_PROG_CODLIN)
                arrPROGROPEXC(i, 16) = .Cell(flexcpText, i, conCOL_PROG_CODGRPLIN)
                
                If .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Insert Then
                    arrPROGROPEXC(i, 11) = .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP)
                    arrPROGROPEXC(i, 12) = .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP)
                    arrPROGROPEXC(i, 17) = .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP)
                    arrPROGROPEXC(i, 18) = .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP)
                    arrPROGROPEXC(i, 19) = .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP)
                    arrPROGROPEXC(i, 20) = .Cell(flexcpText, i, conCOL_PROG_CODOP)
                
                ElseIf .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update Or _
                       .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_delete Then
                    arrPROGROPEXC(i, 11) = .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP)
                    arrPROGROPEXC(i, 12) = .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP)
                    arrPROGROPEXC(i, 17) = .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP)
                    arrPROGROPEXC(i, 18) = .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP)
                    arrPROGROPEXC(i, 19) = .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP)
                    arrPROGROPEXC(i, 20) = .Cell(flexcpText, i, conCOL_PROG_CODOPBKP)
                    
                End If
                
                arrPROGROPEXC(i, 21) = .Cell(flexcpText, i, conCOL_PROG_ORDEM)
            Next i
        End If
    End With
    objCADMOVPCP.PROGRAMADODEL = arrPROGROPEXC
    '' ===============================
    
    
    
    arrPROGROP = Empty
    With grdPROGRAMACAO
        If (.Rows - 1) > 0 Then
            ReDim arrPROGROP(1 To (.Rows - 1), 1 To 21) As String
            For i = 1 To UBound(arrPROGROP)
''                If CLng(.Cell(flexcpText, i, conCOL_PROG_CODOP)) = 336992016 Then
''                    MsgBox "Parar"
''                End If
                arrPROGROP(i, 1) = .Cell(flexcpText, i, conCOL_PROG_DTPROG)
                arrPROGROP(i, 2) = .Cell(flexcpText, i, conCOL_PROG_CODOP)
                arrPROGROP(i, 3) = .Cell(flexcpText, i, conCOL_PROG_DTENTREGA)
                arrPROGROP(i, 4) = .Cell(flexcpText, i, conCOL_PROG_CODSTATUS)
                arrPROGROP(i, 5) = .Cell(flexcpText, i, conCOL_PROG_IDPRODUTO)
                arrPROGROP(i, 6) = .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOP)
                arrPROGROP(i, 7) = .Cell(flexcpText, i, conCOL_PROG_CODPED)
                arrPROGROP(i, 8) = .Cell(flexcpText, i, conCOL_PROG_Action2Do)
                arrPROGROP(i, 9) = .Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada)
                
                If .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Ignore Or _
                   .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update Or _
                   .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_delete Then
                    arrPROGROP(i, 10) = .Cell(flexcpText, i, conCOL_PROG_CODINTENO)
                ElseIf .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Insert Then
                    arrPROGROP(i, 10) = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMTABELA & "_IDPROG", FILIAL, Linha)
                End If
                
                arrPROGROP(i, 13) = .Cell(flexcpText, i, conCOL_PROG_FRACIONADO)
                arrPROGROP(i, 14) = .Cell(flexcpText, i, conCOL_PROG_IDLINHA)
                arrPROGROP(i, 15) = .Cell(flexcpText, i, conCOL_PROG_CODLIN)
                arrPROGROP(i, 16) = .Cell(flexcpText, i, conCOL_PROG_CODGRPLIN)
                
                If .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Insert Then
                    arrPROGROP(i, 11) = .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP)
                    arrPROGROP(i, 12) = .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP)
                    arrPROGROP(i, 17) = .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP)
                    arrPROGROP(i, 18) = .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP)
                    arrPROGROP(i, 19) = .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP)
                    arrPROGROP(i, 20) = .Cell(flexcpText, i, conCOL_PROG_CODOP)
                
                ElseIf .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update Or _
                       .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_delete Then
                    arrPROGROP(i, 11) = .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP)
                    arrPROGROP(i, 12) = .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP)
                    arrPROGROP(i, 17) = .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP)
                    arrPROGROP(i, 18) = .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP)
                    arrPROGROP(i, 19) = .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP)
                    arrPROGROP(i, 20) = .Cell(flexcpText, i, conCOL_PROG_CODOPBKP)
                    
                End If
                
                arrPROGROP(i, 21) = .Cell(flexcpText, i, conCOL_PROG_ORDEM)
            Next i
        End If
    End With
    objCADMOVPCP.PROGRAMADO = arrPROGROP
    '' ===============================
    
    If objCADMOVPCP.GRAVA(cTipOper, strModulo) = False Then Exit Sub

    MsgBox "A Programação [ " & objCADMOVPCP.CODIGO & " ] foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    If cTipOper = "I" Then cTipOper = "C"
    
    If objCADMOVPCP.AtivoDesativo(strModulo) = False Then
        Unload Me
    Else
        iCodigo = objCADMOVPCP.CODIGO
        Call IniciaForm
    End If
    
    If grdLinha.Row > 0 Then grdLinha.Row = lngROLLINANT
    If grdLinha.RowSel > 0 Then grdLinha.RowSel = lngROLLINANT
    grdLinha.TopRow = lngSCROOLMANT

    If lngROLDIAANT > 0 Then
       If (grdDIASPROG.Rows - 1) > 0 Then
        grdDIASPROG.Row = lngROLDIAANT
        Call MarcaLinha(grdDIASPROG, lngROLDIAANT)
       End If
    End If
    If (grdDIASPROG.Rows - 1) > 0 Then
        If lngROLDIAANT > 0 Then grdDIASPROG.RowSel = lngROLDIAANT
        grdDIASPROG.TopRow = lngSCROOLDANT
    End If
    
    If grdPROGRAMACAO.Row > 0 Then grdPROGRAMACAO.Row = lngROLOPANT
    If grdPROGRAMACAO.RowSel > 0 Then grdPROGRAMACAO.RowSel = lngROLOPANT
    grdPROGRAMACAO.TopRow = lngSCROOLPANT

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim i           As Long
    Dim lngORDEM    As Long
    With grdPROGRAMACAO
        lngORDEM = 0
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
                If .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Ignore Then .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update
                lngORDEM = (lngORDEM + 1)
                .Cell(flexcpText, i, conCOL_PROG_ORDEM) = lngORDEM
            End If
        Next i
    End With
    
    Call OrderByGridProgramacao
End Sub

Private Sub Command2_Click()
    With grdPROGRAMACAO
        Call MoveIten(.Row, "C")
    End With
End Sub

Private Sub Command3_Click()

    If cTipOper = "C" Then Exit Sub
    
    
    If boolREMOPSEL = False Then
    
        Dim intRESP As Integer
        
        intRESP = MsgBox("ATENÇÂO" & vbCrLf & _
                         "As OP's Selecionada(s) Serão Excluida(s) !" & vbCrLf & _
                         "Tem Certeza ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                         
        If intRESP = vbNo Then Exit Sub
                         
        
        If (grdLinha.Rows - 1) = 0 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "A Linha não Foi Carregada !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If grdLinha.Row = 0 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Selecione uma Linha !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If (grdDIASPROG.Rows - 1) = 0 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Os dias para Programar Não Foram Carregados !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If grdDIASPROG.Row = 0 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Selecione Uma Data de Programação !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        With grdPROGRAMACAO
            If .Row = 0 Then
                MsgBox "ATENÇÂO" & vbCrLf & _
                       "Selecione um Registro !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Sub
            End If
            
            If Len(Trim(.Cell(flexcpText, .RowSel, conCOL_PROG_CODOP))) > 0 Then
            
                Dim lngCODOP    As Long
                Dim i           As Long
                        
Volta:
                For i = 1 To (.Rows - 1)
                    If .Cell(flexcpChecked, i, conCOL_PROG_Marca) = 1 Then
                        If Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 Then lngCODOP = .Cell(flexcpText, i, conCOL_PROG_CODOP)
                        
                        If .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) = 1 Or _
                           .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) = 2 Or _
                           .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) = 3 Then
                            MsgBox "ATENÇÂO" & vbCrLf & _
                                   "A OP " & lngCODOP & " já foi apontada esta OP não pode ser excluida !!!", vbOKOnly + vbExclamation, "Atenção"
                            .Cell(flexcpText, i, conCOL_PROG_Marca) = 0
                        Else
                            If .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Insert Then
                                If (.Rows - 1) = 1 Then
                                    .Rows = 1
                                Else
                                    .RemoveItem i
                                End If
                                Call ApagaDemaisOPs(lngCODOP)
                                GoTo Volta
                                
                            ElseIf .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Ignore Or _
                                   .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update Then
                                .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_delete
                            End If
                            
                            Call ApagaDemaisOPs(lngCODOP)
                        End If
                    End If
                Next i
            Else
                If .Cell(flexcpText, .RowSel, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Insert Then
                    If (.Rows - 1) = 1 Then
                        .Rows = 1
                    Else
                        .RemoveItem .RowSel
                    End If
                ElseIf .Cell(flexcpText, .RowSel, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Ignore Or _
                       .Cell(flexcpText, .RowSel, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update Then
                    .Cell(flexcpText, .RowSel, conCOL_PROG_Action2Do) = dacEnumUpdateAction_delete
                End If
            End If
            
            '' Calculando a Qtde já Programada
            Call MostraOPProg
            Call CalcQTdJaProg
            Call CalcProgLinha
            
            '' Refazendo os Indices
            Call Command1_Click
            
            Call MarcaLinha(grdDIASPROG, grdDIASPROG.RowSel)
        
            boolREMOPSEL = False
            boolFRACIONA = False
            boolINS_LINHA = False
        
        End With
        
    ElseIf boolREMOPSEL = True Then
    
        Dim dtPROGRAMACAO As Date
    
        dtPROGRAMACAO = grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_DTPROG)
        
        With grdPROGRAMACAO
            Call RetiraSelDiasHaFrente(dtPROGRAMACAO, .Cell(flexcpText, .Row, conCOL_PROG_IDLINHA), .Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN))
            boolREMOPSEL = False
            boolFRACIONA = False
            boolINS_LINHA = False
            
            MsgBox "As Marcações para Mover as OP's foram removidas !!!", vbOKOnly + vbInformation, "Aviso"
        End With
    
    End If
End Sub

Private Sub Command4_Click()
    With grdPROGRAMACAO
        Call MoveIten(.Row, "B")
    End With
End Sub

Private Sub Command5_Click()
    
    If cTipOper = "C" Then
        MsgBox "Você está no modo de Consulta escolha 'Alterar' !!!", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
    
    Dim lngROLLINANT    As Long
    Dim lngROLDIAANT    As Long
    Dim lngROLOPANT     As Long
    
    lngROLLINANT = grdLinha.Row
    lngROLDIAANT = grdDIASPROG.Row
    lngROLOPANT = grdPROGRAMACAO.Row
    
    Call PopGrdProgramadoMesDiaAnterior

    '' Calculando a Qtde já Programada
    Call CalcQTdJaProg
    Call PintaColunasEditaveis
    Call CalcProgLinha
    Call PintaOPAPontada
    
    '' Dando um Order By na Gride de Programação
    Call OrderByGridProgramacao
    
    If grdLinha.Row > 0 Then grdLinha.Row = lngROLLINANT
    If grdLinha.RowSel > 0 Then grdLinha.RowSel = lngROLLINANT
    grdLinha.TopRow = lngSCROOLMANT

    If lngROLDIAANT > 0 Then
       grdDIASPROG.Row = lngROLDIAANT
       Call MarcaLinha(grdDIASPROG, lngROLDIAANT)
    End If
    If lngROLDIAANT > 0 Then grdDIASPROG.RowSel = lngROLDIAANT
    grdDIASPROG.TopRow = lngSCROOLDANT
    
    If grdPROGRAMACAO.Row > 0 Then grdPROGRAMACAO.Row = lngROLOPANT
    If grdPROGRAMACAO.RowSel > 0 Then grdPROGRAMACAO.RowSel = lngROLOPANT
    grdPROGRAMACAO.TopRow = lngSCROOLPANT

    Call MostraOPProg

End Sub

Private Sub Command6_Click()
    
    If cTipOper = "C" Then
        MsgBox "Você está no modo de Consulta escolha 'Alterar' !!!", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
    
    With grdPROGRAMACAO
        
        Dim dtDATAPROG      As Date
        Dim dtPROGRAMACAO   As Date
        Dim lngMESATU       As Long
        Dim lngPESQ         As Long
        Dim lngSALDODISP    As Long
        Dim i               As Long
        Dim lngQTDSEL       As Long
        Dim strINDICE       As String
        
        If (.Rows - 1) = 0 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        lngQTDSEL = 0
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpChecked, i, conCOL_PROG_Marca) = 1 Then
               lngQTDSEL = (lngQTDSEL + 1)
               boolREMOPSEL = True
            End If
        Next i
        If lngQTDSEL > 1 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Não pode ser selecionado mais de Uma OP !", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If .Cell(flexcpChecked, .Row, conCOL_PROG_Marca) = 1 Then
            If Len(Trim(.Cell(flexcpText, .Row, conCOL_PROG_CODOP))) > 0 Then
            
                dtDATAPROG = CDate(.Cell(flexcpText, .Row, conCOL_PROG_DTPROG))
                lngMESATU = Month(dtDATAPROG)
                
                dtDATAPROG = (dtDATAPROG + 1)
                
                ''Trim(.Cell(flexcpText, .Row, conCOL_PROG_IDLINHA))
                strINDICE = Trim(.Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN)) & Trim(Str(Day(dtDATAPROG)) & Trim(Str(Month(dtDATAPROG)))) & Trim(Str(Year(dtDATAPROG)))
            
                Do While lngMESATU = Month(dtDATAPROG)
                    lngPESQ = grdDIASPROG.FindRow(strINDICE, , conCOL_DTP_INDICEPAI)
                    If lngPESQ <> -1 Then
                       lngSALDODISP = SaldoDisponivel(strINDICE)
                       If lngSALDODISP > 0 Then Exit Do
                    End If
                    dtDATAPROG = (dtDATAPROG + 1)
                    ''Trim(.Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN))
                    strINDICE = Trim(.Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN)) & Trim(Str(Day(dtDATAPROG)) & Trim(Str(Month(dtDATAPROG)))) & Trim(Str(Year(dtDATAPROG)))
                Loop
                
                .Cell(flexcpText, .Row, conCOL_PROG_IDDIA) = strINDICE
                .Cell(flexcpText, .Row, conCOL_PROG_DTPROG) = dtDATAPROG
                .Cell(flexcpChecked, .Row, conCOL_PROG_Marca) = 2
                If .Cell(flexcpText, .Row, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Ignore Then .Cell(flexcpText, .Row, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update
                
                boolFRACIONA = True
                dtPROGRAMACAO = DisparaFracionamento(.Cell(flexcpText, .Row, conCOL_PROG_CODOP), .Cell(flexcpText, .Row, conCOL_PROG_QtdeOPProgramada), .Row)
                
                If Len(Trim(.Cell(flexcpText, .Row, conCOL_PROG_CODOP))) > 0 Then
                    If boolFRACIONA = True Then
                        Call RemanejaOPs(.Cell(flexcpText, .Row, conCOL_PROG_CODOP), _
                                         Trim(.Cell(flexcpText, .Row, conCOL_PROG_IDDIA)), _
                                         dtPROGRAMACAO, _
                                         .Cell(flexcpText, .Row, conCOL_PROG_IDLINHA), _
                                         .Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN), _
                                         .Cell(flexcpText, .Row, conCOL_PROG_ORDEM), _
                                         .Row)
                    End If
                End If
                boolFRACIONA = False
                boolREMOPSEL = False
                
                
                grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_PROGAM) = TotalProgramado(Trim(.Cell(flexcpText, .Row, conCOL_PROG_IDDIA)))
                grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_QTDISP) = SaldoDisponivel(Trim(.Cell(flexcpText, .Row, conCOL_PROG_IDDIA)))
                
                Call PintaCelula(grdDIASPROG.RowSel)
                Call CalcQTdJaProg
                Call CalcProgLinha
                
                Call MarcaLinha(grdDIASPROG, grdDIASPROG.RowSel)
                
                Call MostraOPProg
                
            End If
        End If
        
   End With
End Sub

Private Sub Command9_Click()
    If cTipOper = "I" Or cTipOper = "A" Then
        Call IncRegGrid
        ''If boolREMOPSEL = False Then Call IncRegGrid
        ''If boolREMOPSEL = True Then Call MoveOP
    ElseIf cTipOper = "C" Then
        MsgBox "Você está no modo de Consulta escolha 'Alterar' !!!", vbOKOnly + vbInformation, "Aviso"
    End If
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
   Set objCADMOVPCP = CreateObject("CADPROGOP.clsCADPROGOP")
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
    
    Command1.Visible = False
    If lngCodUsuario = 0 Then Command1.Visible = True
    
    Call IniciaForm

End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

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
    
    Call ConfGrdLinha
    Call ConfGrdDtProg
    Call ConfGrdOPProgramada
    Call ConfGrdAux
    Call ConfGrdExc
    
    Call objBLBFunc.Preenche_Mes(cboMes)
    cboMes.ListIndex = (Month(Date) - 1)
    
    Call objBLBFunc.Preenche_Ano(cboAno)
    cboAno.ListIndex = 0
    
    optTIPREL(0).Value = True
    optPeriodo(1).Value = True
    boolINS_LINHA = False
    boolFRACIONA = False
    boolREMOPSEL = False
    
    
    strDTINICIAL = "01/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)
    mskDataI.Text = CDate(strDTINICIAL)
    
    If cboMes.ItemData(cboMes.ListIndex) = 12 Then
        strDTFINAL = "31/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)
        mskDataF.Text = (CDate(strDTFINAL) - 1)
    Else
        strDTFINAL = "01/" & (cboMes.ItemData(cboMes.ListIndex) + 1) & "/" & cboAno.ItemData(cboAno.ListIndex)
        mskDataF.Text = (CDate(strDTFINAL) - 1)
    End If
    
    Call CarregaCampos
    
End Sub


Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        fraPeriodo.Enabled = True
        'Frame2.Enabled = True
        cmdCARREGA.Enabled = True
        If strTipOper = "A" Then
            cboMes.Enabled = False
            cboAno.Enabled = False
            cmdCARREGA.Enabled = False
        End If
    ElseIf strTipOper = "C" Then
        fraPeriodo.Enabled = False
        'Frame2.Enabled = False
        cboMes.Enabled = False
        cboAno.Enabled = False
        cmdCARREGA.Enabled = False
    End If
End Sub


Private Sub ConfGrdLinha()

    With grdLinha

       .Cols = conColumnsIn_LINHA
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_LINHA_FormatString
       .AutoSizeMouse = False
       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_LINHA_CODLIN) = ""
       .ColDataType(conCOL_LINHA_CODLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_LINHA_CODGRPLIN) = ""
       .ColDataType(conCOL_LINHA_CODGRPLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_LINHA_DESCLIN) = ""
       .ColDataType(conCOL_LINHA_DESCLIN) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_LINHA_MES) = ""
       .ColDataType(conCOL_LINHA_MES) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_LINHA_ANO) = ""
       .ColDataType(conCOL_LINHA_ANO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_LINHA_IDINTERNO) = ""
       .ColDataType(conCOL_LINHA_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_LINHA_CAPACLINHA) = ""
       .ColDataType(conCOL_LINHA_CAPACLINHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_LINHA_TOTALPROG) = ""
       .ColDataType(conCOL_LINHA_TOTALPROG) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_LINHA_TOTALDISP) = ""
       .ColDataType(conCOL_LINHA_TOTALDISP) = flexDTLong
       
       .ColWidth(conCOL_LINHA_CODLIN) = 0
       .ColWidth(conCOL_LINHA_CODGRPLIN) = 0
       .ColWidth(conCOL_LINHA_DESCLIN) = 3000
       .ColWidth(conCOL_LINHA_MES) = 0
       .ColWidth(conCOL_LINHA_ANO) = 0
       .ColWidth(conCOL_LINHA_IDINTERNO) = 0
       .ColWidth(conCOL_LINHA_CAPACLINHA) = 900
       .ColWidth(conCOL_LINHA_TOTALPROG) = 0
       .ColWidth(conCOL_LINHA_TOTALDISP) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       .FontName = "Arial"
       .FontBold = True
       .FontSize = 7
       
    End With
    
End Sub


Private Sub CarregaLinha()

    Call ConfGrdLinha
    Call ConfGrdDtProg
    Call ConfGrdOPProgramada
    Call ConfGrdAux
    Call ConfGrdExc
    
    If cboMes.ListIndex = -1 Then Exit Sub
    If cboAno.ListIndex = -1 Then Exit Sub
    
    Dim strDTINICIAL        As String
    Dim strDTFINAL          As String
    
    Dim strDTPESQINI        As String
    Dim strDTPESQFIN        As String

    strDTINICIAL = "'" & Format(CDate("01/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "MM/DD/YYYY") & "'"
    strDTPESQINI = Format(CDate("01/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "DD/MM/YYYY")
    mskDataI.Text = strDTPESQINI
    
    If cboMes.ItemData(cboMes.ListIndex) = 12 Then
        strDTFINAL = "'" & Format(CDate("31/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "MM/DD/YYYY") & "'"
        strDTPESQFIN = Format(CDate("31/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "DD/MM/YYYY")
    Else
        strDTFINAL = "'" & Format((CDate("01/" & (cboMes.ItemData(cboMes.ListIndex) + 1) & "/" & cboAno.ItemData(cboAno.ListIndex)) - 1), "MM/DD/YYYY") & "'"
        strDTPESQFIN = Format((CDate("01/" & (cboMes.ItemData(cboMes.ListIndex) + 1) & "/" & cboAno.ItemData(cboAno.ListIndex)) - 1), "DD/MM/YYYY")
    End If
    mskDataF.Text = strDTPESQFIN
    
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
    sSql = sSql & "        GRPL.SGI_CODIGO        As SGI_CODGRUPLIN" & vbCrLf
    sSql = sSql & "       ,GRPL.SGI_DESCRI        As SGI_DESCRLINHA" & vbCrLf
    sSql = sSql & "       ,CAPC.SGI_IDPAI" & vbCrLf
    sSql = sSql & "       ,Month(CAPC.SGI_DATA)   As SGI_MES" & vbCrLf
    sSql = sSql & "       ,Year(CAPC.SGI_DATA)    As SGI_ANO" & vbCrLf
    sSql = sSql & "       ,Sum(SGI_TOTALPECAS)    As SGI_TOTALPECAS" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "        SGI_MAQULIN_CAPAC" & strNOMTABELA & " CAPC" & vbCrLf
    sSql = sSql & "       ,SGI_CADGRUPLINHA" & strNOMTABELA & "  GRPL" & vbCrLf
    
    sSql = sSql & "  Where" & vbCrLf
    sSql = sSql & "        CAPC.SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "    And CAPC.SGI_ATIVO       = 1" & vbCrLf
    sSql = sSql & "    And Month(CAPC.SGI_DATA) = " & cboMes.ItemData(cboMes.ListIndex) & vbCrLf
    sSql = sSql & "    And Year(CAPC.SGI_DATA)  = " & cboAno.ItemData(cboAno.ListIndex) & vbCrLf
    sSql = sSql & "    And GRPL.SGI_FILIAL      = CAPC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "    And GRPL.SGI_CODIGO      = CAPC.SGI_CODIGO" & vbCrLf
    sSql = sSql & "Group By" & vbCrLf
    sSql = sSql & "         GRPL.SGI_CODIGO" & vbCrLf
    sSql = sSql & "        ,GRPL.SGI_DESCRI" & vbCrLf
    sSql = sSql & "        ,CAPC.SGI_IDPAI" & vbCrLf
    sSql = sSql & "        ,Month(CAPC.SGI_DATA)" & vbCrLf
    sSql = sSql & "        ,Year(CAPC.SGI_DATA)" & vbCrLf
    sSql = sSql & "Order By GRPL.SGI_DESCRI,GRPL.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
       With grdLinha
            Do While Not BREC.EOF()
                .AddItem BREC!SGI_CODGRUPLIN & vbTab & _
                         BREC!SGI_CODGRUPLIN & vbTab & _
                         Trim(BREC!SGI_DESCRLINHA) & vbTab & _
                         BREC!SGI_Mes & vbTab & _
                         BREC!SGI_ANO & vbTab & _
                         BREC!SGI_IDPAI & vbTab & _
                         BREC!SGI_TOTALPECAS & vbTab & _
                         "" & vbTab & _
                         ""
                         
                 '' Populando a Gride de Dias de Programação Válidas
                 Call PopGrdDiasProd(BREC!SGI_CODGRUPLIN)
                 BREC.MoveNext
            Loop
       End With
    Else
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Não existe dados para carregar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close
    
    '' Pegando as OP's Que ainda não foram apontadas
    Call PopGrdProgramado(strDTINICIAL, strDTFINAL)
    boolINS_LINHA = False
    boolFRACIONA = False
    boolREMOPSEL = False
    
    '' Calculando a Qtde já Programada
    Call CalcQTdJaProg
    Call PintaColunasEditaveis
    Call CalcProgLinha
    Call PintaOPAPontada
    
    If (grdLinha.Rows - 1) > 0 Then
        grdLinha.Row = 1
    End If
    
    '' Dando um Order By na Gride de Programação
    Call OrderByGridProgramacao
    
    
    
End Sub


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

Private Sub CarregaCampos()

    If objCADMOVPCP.Carrega_Campos(strNOMTABELA) = False Then Exit Sub
    
    Dim i As Integer
    
    If cTipOper = "I" Then
        For i = 0 To (cboMes.ListCount - 1)
            If cboMes.ItemData(i) = Month(Now) Then
                cboMes.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To (cboAno.ListCount - 1)
            If cboAno.ItemData(i) = Year(Now) Then
                cboAno.ListIndex = i
                Exit For
            End If
        Next i
    ElseIf cTipOper = "C" Or cTipOper = "A" Then
        For i = 0 To (cboMes.ListCount - 1)
            If cboMes.ItemData(i) = Month(CDate(objCADMOVPCP.DTPROGRAMA)) Then
                cboMes.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To (cboAno.ListCount - 1)
            If cboAno.ItemData(i) = Year(CDate(objCADMOVPCP.DTPROGRAMA)) Then
                cboAno.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    Call CarregaLinha

End Sub

Private Sub grdDIASPROG_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    lngSCROOLDANT = NewTopRow
End Sub

Private Sub grdDIASPROG_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Call MarcaLinha(grdDIASPROG, NewRowSel)
End Sub

Private Sub grdDIASPROG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdDIASPROG
        Select Case Col
               Case conCOL_DTP_IDPAI, _
                    conCOL_DTP_DTPROG, _
                    conCOL_DTP_TOTDIA, _
                    conCOL_DTP_PROGAM, _
                    conCOL_DTP_QTDISP, _
                    conCOL_DTP_IDINTERNO, _
                    conCOL_DTP_INDICEPAI, _
                    conCOL_DTP_CODLIN, _
                    conCOL_DTP_CODGRPLIN
                    Cancel = True
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub
End Sub

Private Sub grdDIASPROG_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    Call Desmarca(grdDIASPROG, OldRowSel)
End Sub

Private Sub grdDIASPROG_Click()
    Call MostraOPProg
End Sub

Private Sub grdDIASPROG_RowColChange()
    Call MostraOPProg
End Sub

Private Sub grdLinha_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    lngSCROOLMANT = NewTopRow
End Sub

Private Sub grdLinha_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Call MarcaLinha(grdLinha, NewRowSel)
End Sub

Private Sub grdLinha_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdLinha
        Select Case Col
               Case conCOL_LINHA_CODLIN, _
                    conCOL_LINHA_CODGRPLIN, _
                    conCOL_LINHA_DESCLIN, _
                    conCOL_LINHA_MES, _
                    conCOL_LINHA_ANO, _
                    conCOL_LINHA_IDINTERNO, _
                    conCOL_LINHA_CAPACLINHA, _
                    conCOL_LINHA_TOTALPROG, _
                    conCOL_LINHA_TOTALDISP
                    Cancel = True
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub
End Sub


Private Sub PopGrdDiasProd(lngIDINTERNO As Long)

    If lngIDINTERNO = 0 Then Exit Sub
    
    Dim strINDICE   As String

    '' ------------------------------------
    '' Carregando as Datas para Programação
    
    sSql = ""
    
    sSql = sSql & "Select" & vbCrLf
    sSql = sSql & "       SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_DATA" & vbCrLf
    sSql = sSql & "      ,SGI_IDPAI" & vbCrLf
    sSql = sSql & "      ,Max(SGI_TOTALPECAS) As SGI_TOTALPECAS" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "        SGI_MAQULIN_CAPAC" & strNOMTABELA & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "        SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "    And SGI_CODIGO      = " & lngIDINTERNO & vbCrLf
    sSql = sSql & "    And SGI_ATIVO       = 1" & vbCrLf
    sSql = sSql & "    And Month(SGI_DATA) = " & cboMes.ItemData(cboMes.ListIndex) & vbCrLf
    sSql = sSql & "    And Year(SGI_DATA)  = " & cboAno.ItemData(cboAno.ListIndex) & vbCrLf
    sSql = sSql & "Group By SGI_CODIGO" & vbCrLf
    sSql = sSql & "     ,SGI_DATA" & vbCrLf
    sSql = sSql & "     ,SGI_IDPAI" & vbCrLf
    sSql = sSql & "Order BY SGI_DATA"
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then
        With grdDIASPROG
            Do While Not BREC2.EOF()
                
                strINDICE = Trim(Str(lngIDINTERNO)) & Trim(Str(Day(BREC2!SGI_DATA)) & Trim(Str(Month(BREC2!SGI_DATA)))) & Trim(Str(Year(BREC2!SGI_DATA)))
                
                .AddItem lngIDINTERNO & vbTab & _
                         Format(BREC2!SGI_DATA, "DD/MM/YYYY") & vbTab & _
                         BREC2!SGI_TOTALPECAS & vbTab & _
                         0 & vbTab & _
                         0 & vbTab & _
                         BREC2!SGI_IDPAI & vbTab & _
                         strINDICE & vbTab & _
                         "" & vbTab & _
                         BREC2!SGI_CODIGO & vbTab & _
                         0
               
               BREC2.MoveNext
            Loop
        End With
    End If
    BREC2.Close
End Sub

Private Sub ConfGrdDtProg()

    With grdDIASPROG

       .Cols = conColumnsIn_DTP
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_DTP_FormatString
       .AutoSizeMouse = False
       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_DTP_IDPAI) = ""
       .ColDataType(conCOL_DTP_IDPAI) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_DTP_DTPROG) = ""
       .ColDataType(conCOL_DTP_DTPROG) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_DTP_TOTDIA) = ""
       .ColDataType(conCOL_DTP_TOTDIA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_DTP_PROGAM) = ""
       .ColDataType(conCOL_DTP_PROGAM) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_DTP_QTDISP) = ""
       .ColDataType(conCOL_DTP_QTDISP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_DTP_IDINTERNO) = ""
       .ColDataType(conCOL_DTP_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_DTP_INDICEPAI) = ""
       .ColDataType(conCOL_DTP_INDICEPAI) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_DTP_CODLIN) = ""
       .ColDataType(conCOL_DTP_CODLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_DTP_CODGRPLIN) = ""
       .ColDataType(conCOL_DTP_CODGRPLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_DTP_SELDIA) = ""
       .ColDataType(conCOL_DTP_SELDIA) = flexDTBoolean
       
       .ColWidth(conCOL_DTP_IDPAI) = 0
       .ColWidth(conCOL_DTP_DTPROG) = 900
       .ColWidth(conCOL_DTP_TOTDIA) = 1000
       .ColWidth(conCOL_DTP_PROGAM) = 1000
       .ColWidth(conCOL_DTP_QTDISP) = 1000
       .ColWidth(conCOL_DTP_IDINTERNO) = 0
       .ColWidth(conCOL_DTP_INDICEPAI) = 0
       .ColWidth(conCOL_DTP_CODLIN) = 0
       .ColWidth(conCOL_DTP_CODGRPLIN) = 0
       .ColWidth(conCOL_DTP_SELDIA) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       .FontName = "Arial"
       .FontSize = 7
       .FontBold = True
       
    End With
    
End Sub

Private Sub grdLinha_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
   ''Label2.Caption = lngSCROOLMANT
End Sub

Private Sub grdLinha_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
        Call Desmarca(grdLinha, OldRowSel)
End Sub

Private Sub grdLinha_Click()
    Call MostraDiasProg
End Sub

Private Sub MostraDiasProg()
    With grdLinha
        If (.Rows - 1) = 0 Then Exit Sub
        Call objBLBFunc.CarregaDadosGrdFilhoSemAction2Do(grdDIASPROG, conCOL_DTP_IDPAI, -1)
        
        Call objBLBFunc.CarregaDadosGrdFilho(grdPROGRAMACAO, conCOL_PROG_Action2Do, conCOL_PROG_IDDIA, -1)
        
        If .Row = 0 Then Exit Sub
        Call objBLBFunc.CarregaDadosGrdFilhoSemAction2Do(grdDIASPROG, conCOL_DTP_IDPAI, .Cell(flexcpText, .Row, conCOL_LINHA_CODGRPLIN))
        
    End With
End Sub

Private Sub grdLinha_GotFocus()
    With grdLinha
        If (.Rows - 1) > 0 And .RowSel > 0 Then
            .Cell(flexcpBackColor, .RowSel, 0, .RowSel, (.Cols - 1)) = &H8000000D
            .Cell(flexcpForeColor, .RowSel, 0, .RowSel, (.Cols - 1)) = &H8000000E
        End If
    End With
End Sub

Private Sub grdLinha_RowColChange()
    Call MostraDiasProg
End Sub


Private Sub ConfGrdOPProgramada()

    With grdPROGRAMACAO

       .Cols = conColumnsIn_PROG
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_PROG_FormatString
       .AutoSizeMouse = False
       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_PROG_IDDIA) = ""
       .ColDataType(conCOL_PROG_IDDIA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_CODOP) = ""
       .ColDataType(conCOL_PROG_CODOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPROD) = ""
       .ColDataType(conCOL_PROG_CODPROD) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_DESCPROD) = ""
       .ColDataType(conCOL_PROG_DESCPROD) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_DTENTREGA) = ""
       .ColDataType(conCOL_PROG_DTENTREGA) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeOP) = ""
       .ColDataType(conCOL_PROG_QtdeOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeOPProgramada) = ""
       .ColDataType(conCOL_PROG_QtdeOPProgramada) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeReal) = ""
       .ColDataType(conCOL_PROG_QtdeReal) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPED) = ""
       .ColDataType(conCOL_PROG_CODPED) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDPRODUTO) = ""
       .ColDataType(conCOL_PROG_IDPRODUTO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDINTERNOOP) = ""
       .ColDataType(conCOL_PROG_IDINTERNOOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODINTENO) = ""
       .ColDataType(conCOL_PROG_CODINTENO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODLIN) = ""
       .ColDataType(conCOL_PROG_CODLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODGRPLIN) = ""
       .ColDataType(conCOL_PROG_CODGRPLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDLINHA) = ""
       .ColDataType(conCOL_PROG_IDLINHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_DTENTREGABKP) = ""
       .ColDataType(conCOL_PROG_DTENTREGABKP) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PROG_CODOPBKP) = ""
       .ColDataType(conCOL_PROG_CODOPBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPEDBKP) = ""
       .ColDataType(conCOL_PROG_CODPEDBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDPRODBKP) = ""
       .ColDataType(conCOL_PROG_IDPRODBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDINTERNOOPBKP) = ""
       .ColDataType(conCOL_PROG_IDINTERNOOPBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_INDICEPROG) = ""
       .ColDataType(conCOL_PROG_INDICEPROG) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_CODSTATUS) = ""
       .ColDataType(conCOL_PROG_CODSTATUS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODSTATUSBKP) = ""
       .ColDataType(conCOL_PROG_CODSTATUSBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_NECKIN) = ""
       .ColDataType(conCOL_PROG_NECKIN) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_FECH) = ""
       .ColDataType(conCOL_PROG_FECH) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_COMP) = ""
       .ColDataType(conCOL_PROG_COMP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_Action2Do) = ""
       .ColDataType(conCOL_PROG_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_Marca) = ""
       .ColDataType(conCOL_PROG_Marca) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_PROG_CODSTATUSAPONT) = ""
       .ColDataType(conCOL_PROG_CODSTATUSAPONT) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_DESCSTATSAPONT) = ""
       .ColDataType(conCOL_PROG_DESCSTATSAPONT) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_ORDEM) = ""
       .ColDataType(conCOL_PROG_ORDEM) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_TIMESTAMP) = ""
       .ColDataType(conCOL_PROG_TIMESTAMP) = flexDTLong
       
       .ColWidth(conCOL_PROG_IDDIA) = 0
       .ColWidth(conCOL_PROG_CODOP) = 1000
       .ColWidth(conCOL_PROG_CODPROD) = 1000
       .ColWidth(conCOL_PROG_DESCPROD) = 4000
       .ColWidth(conCOL_PROG_DTENTREGA) = 900
       .ColWidth(conCOL_PROG_QtdeOP) = 900
       .ColWidth(conCOL_PROG_QtdeOPProgramada) = 900
       .ColWidth(conCOL_PROG_QtdeReal) = 900
       .ColWidth(conCOL_PROG_CODPED) = 0
       .ColWidth(conCOL_PROG_IDPRODUTO) = 0
       .ColWidth(conCOL_PROG_IDINTERNOOP) = 0
       .ColWidth(conCOL_PROG_CODINTENO) = 0
       .ColWidth(conCOL_PROG_CODLIN) = 0
       .ColWidth(conCOL_PROG_CODGRPLIN) = 0
       .ColWidth(conCOL_PROG_IDLINHA) = 0
       .ColWidth(conCOL_PROG_DTENTREGABKP) = 0
       .ColWidth(conCOL_PROG_CODOPBKP) = 0
       .ColWidth(conCOL_PROG_CODPEDBKP) = 0
       .ColWidth(conCOL_PROG_IDPRODBKP) = 0
       .ColWidth(conCOL_PROG_IDINTERNOOPBKP) = 0
       .ColWidth(conCOL_PROG_INDICEPROG) = 0
       .ColWidth(conCOL_PROG_CODSTATUS) = 0
       .ColWidth(conCOL_PROG_CODSTATUSBKP) = 0
       .ColWidth(conCOL_PROG_FRACIONADO) = 0
       .ColWidth(conCOL_PROG_DTPROG) = 1200                 '' Data de Programação
       .ColWidth(conCOL_PROG_NECKIN) = 500
       .ColWidth(conCOL_PROG_FECH) = 500
       .ColWidth(conCOL_PROG_COMP) = 500
       .ColWidth(conCOL_PROG_Action2Do) = 1200              '' Action To Do
       .ColWidth(conCOL_PROG_Marca) = 200
       .ColWidth(conCOL_PROG_CODSTATUSAPONT) = 0
       .ColWidth(conCOL_PROG_DESCSTATSAPONT) = 1100
       .ColWidth(conCOL_PROG_ORDEM) = 0
       .ColWidth(conCOL_PROG_TIMESTAMP) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       .FontName = "Arial"
       .FontSize = 7
       .FontBold = True
       
    End With
    
End Sub


Private Sub PopGrdProgramado(strDTINICIAL As String, strDTFINAL As String)

    '' dtDATAPROG As Date, lngCODLIN As Long, strINDICE As String

    Dim dtDATAPROG      As Date
    Dim strINDICE       As String
    Dim lngMESATU       As Long
    Dim lngPESQ         As Long
    Dim lngActioToDo    As Long

    dtDATAPROG = Date
    lngMESATU = Month(dtDATAPROG)
    
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       SGI_IDLINHA" & vbCrLf
    sSql = sSql & "     , SGI_CODGRPLIN" & vbCrLf
    sSql = sSql & "     , SGI_DATAPROG" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAl = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_DATAPROG between " & strDTINICIAL & " And " & strDTFINAL & vbCrLf
    ''sSql = sSql & "   And (SGI_STATUSAPONT Is Null or SGI_STATUSAPONT = 2)" & vbCrLf
    sSql = sSql & "Order By SGI_IDLINHA,SGI_CODGRPLIN,SGI_DATAPROG"
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC4.EOF()
        
        
        '' --------------------------------------------
        '' Pegando a OP
        sSql = ""
        sSql = "Select" & vbCrLf
        sSql = sSql & "       MOVPCP.*" & vbCrLf
        sSql = sSql & "     , PROD.SGI_CODIGO       As SGI_CODROTULO" & vbCrLf
        sSql = sSql & "     , PROD.SGI_DESCRICAO    As SGI_DESCOTULO" & vbCrLf
        sSql = sSql & "     , PROD.SGI_NECKIN" & vbCrLf
        sSql = sSql & "     , PROD.SGI_VernTampa" & vbCrLf
        sSql = sSql & "     , ORDP.SGI_QTDE" & vbCrLf
        sSql = sSql & "     , ORDP.SGI_FECHTPFU" & vbCrLf
        sSql = sSql & "     , ORDP.SGI_DATENTREGA" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & " MOVPCP" & vbCrLf
        sSql = sSql & "     , SGI_ORDEMPROD" & strNOMTABELA & " ORDP" & vbCrLf
        sSql = sSql & "     , SGI_CADPRODUTO      PROD" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       MOVPCP.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And MOVPCP.SGI_DATAPROG   = '" & Format(BREC4!SGI_DATAPROG, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And MOVPCP.SGI_IDLINHA    = " & BREC4!SGI_IDLINHA & vbCrLf
        sSql = sSql & "   And MOVPCP.SGI_CODGRPLIN  = " & BREC4!SGI_CODGRPLIN & vbCrLf
        ''sSql = sSql & "   And (MOVPCP.SGI_STATUSAPONT Is Null or MOVPCP.SGI_STATUSAPONT = 2)" & vbCrLf
        
        sSql = sSql & "   And ORDP.SGI_FILIAL     = MOVPCP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And ORDP.SGI_IDPAI      = MOVPCP.SGI_IDINTERNO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ORDP.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "Order By MOVPCP.SGI_DATAPROG,MOVPCP.SGI_ORDEM"
        
        BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC3.EOF() Then
                            
            ''If BREC4!SGI_DATAPROG <= dtDATAPROG Then
            ''    strINDICE = Trim(Str(BREC3!SGI_IDLINHA)) & Trim(Str(Day(dtDATAPROG)) & Trim(Str(Month(dtDATAPROG)))) & Trim(Str(Year(dtDATAPROG))) & Trim(Str(BREC3!SGI_CODGRPLIN))
            ''    lngActioToDo = dacEnumUpdateAction_Insert
            ''ElseIf BREC4!SGI_DATAPROG > dtDATAPROG Then
                ''Trim(Str(BREC3!SGI_IDLINHA))
                strINDICE = Trim(Str(BREC3!SGI_CODGRPLIN)) & Trim(Str(Day(BREC4!SGI_DATAPROG)) & Trim(Str(Month(BREC4!SGI_DATAPROG)))) & Trim(Str(Year(BREC4!SGI_DATAPROG)))
                lngActioToDo = dacEnumUpdateAction_Ignore
            '' End If
            
            ''Do While lngMESATU = Month(dtDATAPROG)
            ''    lngPESQ = grdDIASPROG.FindRow(strINDICE, , conCOL_DTP_INDICEPAI)
            ''    If lngPESQ <> -1 Then Exit Do
            ''    dtDATAPROG = (dtDATAPROG + 1)
            ''Loop
            
            Do While Not BREC3.EOF()
            
            ''    If BREC4!SGI_DATAPROG <= dtDATAPROG Then
            ''        With grdEXC
            ''
            ''            .AddItem strINDICE & vbTab & BREC3!SGI_CODOP & vbTab & _
            ''                     BREC3!SGI_CODROTULO & vbTab & BREC3!SGI_DESCOTULO & vbTab & _
            ''                     Format(BREC3!SGI_DATAENTR, "DD/MM/YYYY") & vbTab & _
            ''                     BREC3!SGI_QTDE & vbTab & BREC3!SGI_QTDEPROD & vbTab & _
            ''                     "" & vbTab & BREC3!SGI_CODPED & vbTab & _
            ''                     BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
            ''                     BREC3!SGI_CODINTENO & vbTab & BREC3!SGI_CODLIN & vbTab & _
            ''                     BREC3!SGI_CODGRPLIN & vbTab & BREC3!SGI_IDLINHA & vbTab & _
            ''                     Format(BREC3!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
            ''                     BREC3!SGI_CODOP & vbTab & BREC3!SGI_CODPED & vbTab & _
            ''                     BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
            ''                     Trim(strINDICE) & Trim(Str(BREC3!SGI_CODOP)) & vbTab & _
            ''                     BREC3!SGI_CODSTATUS & vbTab & _
            ''                     BREC3!SGI_STATUSORIG & vbTab & _
            ''                     BREC3!SGI_FRACIONADA & vbTab & _
            ''                     dtDATAPROG & vbTab & _
            ''                     IIf(BREC3!SGI_NECKIN = 1, "Sim", "Não") & vbTab & _
            ''                     PegaFechamentoTampaFuro(Str(BREC3!SGI_FECHTPFU)) & vbTab & _
            ''                     IIf(IsNull(BREC3!SGI_VernTampa) = False, PegaComp(BREC3!SGI_VernTampa), "") & vbTab & _
            ''                     dacEnumUpdateAction_delete & vbTab & 0 & vbTab & _
            ''                     "" & vbTab & "" & vbTab & _
            ''                     BREC3!SGI_ORDEM & vbTab & _
            ''                     IIf(IsNull(BREC3!SGI_TimeStamp), -1, BREC3!SGI_TimeStamp)
            ''
            ''        End With
            ''    End If
                
                With grdPROGRAMACAO
                    
                    .AddItem strINDICE & vbTab & BREC3!SGI_CODOP & vbTab & _
                             BREC3!SGI_CODROTULO & vbTab & BREC3!SGI_DESCOTULO & vbTab & _
                             Format(BREC3!SGI_DATAENTR, "DD/MM/YYYY") & vbTab & _
                             BREC3!SGI_QTDE & vbTab & BREC3!SGI_QTDEPROD & vbTab & _
                             "" & vbTab & BREC3!SGI_CODPED & vbTab & _
                             BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                             BREC3!SGI_CODINTENO & vbTab & BREC3!SGI_CODLIN & vbTab & _
                             BREC3!SGI_CODGRPLIN & vbTab & BREC3!SGI_IDLINHA & vbTab & _
                             Format(BREC3!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                             BREC3!SGI_CODOP & vbTab & BREC3!SGI_CODPED & vbTab & _
                             BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                             Trim(strINDICE) & Trim(Str(BREC3!SGI_CODOP)) & vbTab & _
                             BREC3!SGI_CODSTATUS & vbTab & _
                             BREC3!SGI_STATUSORIG & vbTab & _
                             BREC3!SGI_FRACIONADA & vbTab & _
                             BREC3!SGI_DATAPROG & vbTab & _
                             IIf(BREC3!SGI_NECKIN = 1, "Sim", "Não") & vbTab & _
                             PegaFechamentoTampaFuro(Str(BREC3!SGI_FECHTPFU)) & vbTab & _
                             IIf(IsNull(BREC3!SGI_VernTampa) = False, PegaComp(BREC3!SGI_VernTampa), "") & vbTab & _
                             lngActioToDo & vbTab & 0 & vbTab & _
                             "" & vbTab & "" & vbTab & _
                             BREC3!SGI_ORDEM & vbTab & _
                             IIf(IsNull(BREC3!SGI_TimeStamp), -1, BREC3!SGI_TimeStamp)
                                             
                    If Not IsNull(BREC3!SGI_STATUSAPONT) Then
                        .Cell(flexcpText, (.Rows - 1), conCOL_PROG_CODSTATUSAPONT) = BREC3!SGI_STATUSAPONT
                        .Cell(flexcpText, (.Rows - 1), conCOL_PROG_QtdeReal) = BREC3!SGI_QTDEAPONTADA
                        .Cell(flexcpText, (.Rows - 1), conCOL_PROG_DESCSTATSAPONT) = PegaDescrStatusApontamento(Str(BREC3!SGI_STATUSAPONT))
                    End If
                End With
                
                
                BREC3.MoveNext
            Loop
        End If
        BREC3.Close
        
        ''dtDATAPROG = (dtDATAPROG + 1)
        BREC4.MoveNext
    Loop
    BREC4.Close
    
End Sub

Private Sub grdPROGRAMACAO_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    With grdPROGRAMACAO
        
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        
        Dim dtPROGRAMACAO As Date
        
        Select Case Col
               Case conCOL_PROG_CODOP, _
                    conCOL_PROG_QtdeOPProgramada
                    
mov_manual:
                    If Len(Trim(.Cell(flexcpText, .Row, conCOL_PROG_CODOP))) > 0 Then
                        
                        boolFRACIONA = False
                        dtPROGRAMACAO = DisparaFracionamento(.Cell(flexcpText, .Row, conCOL_PROG_CODOP), .Cell(flexcpText, .Row, conCOL_PROG_QtdeOPProgramada), .Row)
                        
                        If Len(Trim(.Cell(flexcpText, .Row, conCOL_PROG_CODOP))) > 0 Then
                            If boolFRACIONA = True Then
                                Call RemanejaOPs(.Cell(flexcpText, .Row, conCOL_PROG_CODOP), _
                                                 Trim(.Cell(flexcpText, .Row, conCOL_PROG_IDDIA)), _
                                                 dtPROGRAMACAO, _
                                                 .Cell(flexcpText, .Row, conCOL_PROG_IDLINHA), _
                                                 .Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN), _
                                                 .Cell(flexcpText, .Row, conCOL_PROG_ORDEM), _
                                                 .Row)
                                boolREMOPSEL = False
                            End If
                        End If
                        boolINS_LINHA = False
                        boolFRACIONA = False
                        
                                     
                    End If
               Case conCOL_PROG_DTENTREGA
               Case conCOL_PROG_Marca
                    
                    If boolREMOPSEL = True Then
                        
                        If .Cell(flexcpChecked, Row, conCOL_PROG_Marca) = 1 Then
                            Dim lngSALDODISP As Long
                            
                            lngSALDODISP = SaldoDisponivel(.Cell(flexcpText, Row, conCOL_PROG_IDDIA))
                            If lngSALDODISP = 0 Then
                                .Cell(flexcpChecked, Row, conCOL_PROG_Marca) = 2
                                MsgBox "ATENÇÂO" & vbCrLf & _
                                       "Não há mais saldo para alocar !!!", vbOKOnly + vbExclamation, "Aviso"
                                Exit Sub
                            End If
                            
                            .Cell(flexcpChecked, Row, conCOL_PROG_Marca) = 2
                            GoTo mov_manual
                            
                        
                        ElseIf .Cell(flexcpChecked, Row, conCOL_PROG_Marca) = 2 Then
                            boolREMOPSEL = False
                            .Cell(flexcpText, Row, conCOL_PROG_QtdeOPProgramada) = .Cell(flexcpText, Row, conCOL_PROG_QtdeOP)
                        End If
                        
                    End If
        End Select
    
        grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_PROGAM) = TotalProgramado(Trim(.Cell(flexcpText, Row, conCOL_PROG_IDDIA)))
        grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_QTDISP) = SaldoDisponivel(Trim(.Cell(flexcpText, Row, conCOL_PROG_IDDIA)))
        
        Call PintaCelula(grdDIASPROG.RowSel)
        Call CalcQTdJaProg
        Call CalcProgLinha
        
        Call MarcaLinha(grdDIASPROG, grdDIASPROG.RowSel)
    
    End With

End Sub

Private Sub grdPROGRAMACAO_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
        lngSCROOLPANT = NewTopRow
End Sub

Private Sub grdPROGRAMACAO_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Call PintaColunasEditaveis
    Call MarcaLinha(grdPROGRAMACAO, NewRowSel)
End Sub

Private Sub grdPROGRAMACAO_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdPROGRAMACAO
        Select Case Col
               Case conCOL_PROG_IDDIA, conCOL_PROG_CODPROD, _
                    conCOL_PROG_DESCPROD, conCOL_PROG_QtdeOP, _
                    conCOL_PROG_QtdeReal, conCOL_PROG_CODPED, _
                    conCOL_PROG_IDPRODUTO, conCOL_PROG_IDINTERNOOP, _
                    conCOL_PROG_CODINTENO, conCOL_PROG_CODLIN, _
                    conCOL_PROG_CODGRPLIN, conCOL_PROG_IDLINHA, _
                    conCOL_PROG_DTENTREGABKP, conCOL_PROG_CODOPBKP, _
                    conCOL_PROG_CODPEDBKP, conCOL_PROG_IDPRODBKP, _
                    conCOL_PROG_IDINTERNOOPBKP, conCOL_PROG_INDICEPROG, _
                    conCOL_PROG_CODSTATUS, conCOL_PROG_CODSTATUSBKP, _
                    conCOL_PROG_FRACIONADO, conCOL_PROG_DTPROG, _
                    conCOL_PROG_NECKIN, conCOL_PROG_FECH, _
                    conCOL_PROG_COMP, conCOL_PROG_Action2Do, _
                    conCOL_PROG_CODSTATUSAPONT, conCOL_PROG_DESCSTATSAPONT, _
                    conCOL_PROG_ORDEM, conCOL_PROG_TIMESTAMP
                    Cancel = True
               Case conCOL_PROG_CODOP, _
                    conCOL_PROG_DTENTREGA, _
                    conCOL_PROG_QtdeOPProgramada, _
                    conCOL_PROG_Marca
                    If cTipOper = "C" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If .Cell(flexcpText, Row, conCOL_PROG_CODSTATUSAPONT) = 1 Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Col = conCOL_PROG_CODOP Then
                        If cTipOper = "A" Then
                            If (Len(Trim(.Cell(flexcpText, Row, Col))) > 0 And _
                               (.Cell(flexcpText, Row, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Ignore Or _
                               .Cell(flexcpText, Row, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update)) Then Cancel = True
                        End If
                    End If
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub

End Sub

Private Sub MostraOPProg()
    With grdDIASPROG
    
        If (.Rows - 1) = 0 Then Exit Sub
        Call objBLBFunc.CarregaDadosGrdFilho(grdPROGRAMACAO, conCOL_PROG_Action2Do, conCOL_PROG_IDDIA, -1)
        
        If .Row = 0 Then Exit Sub
        Call objBLBFunc.CarregaDadosGrdFilho(grdPROGRAMACAO, conCOL_PROG_Action2Do, conCOL_PROG_IDDIA, .Cell(flexcpText, .Row, conCOL_DTP_INDICEPAI))
        
    End With
End Sub

Private Sub grdPROGRAMACAO_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
        Call Desmarca(grdPROGRAMACAO, OldRowSel)
        Call PintaColunasEditaveis
End Sub

Private Sub grdPROGRAMACAO_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdPROGRAMACAO
          Select Case Col
                    Case conCOL_PROG_CODOP, _
                         conCOL_PROG_QtdeOPProgramada
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_PROG_DTENTREGA
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          End Select
     End With
End Sub

Private Sub CalcQTdJaProg()
    Dim i               As Long
    Dim j               As Long
    Dim lngQTDJAPROG    As Long
    Dim lngTOTDISP      As Long
    
    For i = 1 To (grdDIASPROG.Rows - 1)
        
        lngQTDJAPROG = 0
        With grdPROGRAMACAO
            For j = 1 To (.Rows - 1)
                If .Cell(flexcpText, j, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete And _
                   .Cell(flexcpText, j, conCOL_PROG_CODSTATUSAPONT) <> 1 Then
                   If grdDIASPROG.Cell(flexcpText, i, conCOL_DTP_INDICEPAI) = .Cell(flexcpText, j, conCOL_PROG_IDDIA) Then
                        If Len(Trim(.Cell(flexcpText, j, conCOL_PROG_QtdeOPProgramada))) > 0 Then lngQTDJAPROG = lngQTDJAPROG + (.Cell(flexcpText, j, conCOL_PROG_QtdeOPProgramada))
                   End If
                End If
            Next j
        End With
        lngTOTDISP = (CLng(grdDIASPROG.Cell(flexcpText, i, conCOL_DTP_TOTDIA)) - lngQTDJAPROG)
        
        grdDIASPROG.Cell(flexcpText, i, conCOL_DTP_PROGAM) = lngQTDJAPROG
        grdDIASPROG.Cell(flexcpText, i, conCOL_DTP_QTDISP) = lngTOTDISP
        
        Call PintaCelula(i)
        
    Next i
End Sub

Private Sub IncRegGrid()

    Dim intRESP         As Long

    If grdLinha.Row = 0 Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Selecione uma Linha para programar !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If (grdLinha.Rows - 1) = 0 Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Não foram carregadas as Linhas para Programação !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If grdDIASPROG.Row = 0 Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Selecione uma dia para programar !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If (grdDIASPROG.Rows - 1) = 0 Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Não há dias para programar !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    
    If objBLBFunc.FcExisteLinhaVazia(grdPROGRAMACAO, conCOL_PROG_CODOP) = False Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Existe linha vázia não e permitido Incluir !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If grdLinha.Cell(flexcpText, grdLinha.Row, conCOL_LINHA_CODGRPLIN) <> grdDIASPROG.Cell(flexcpText, grdDIASPROG.Row, conCOL_DTP_IDPAI) Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Selecione um dia da para a Linha !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    
    '' Não deixa incluir com Data Retroativa
    If CDate(grdDIASPROG.Cell(flexcpText, grdDIASPROG.Row, conCOL_DTP_DTPROG)) < Date Then
         MsgBox "ATENÇÂO" & vbCrLf & _
                "Não é possivel inserir neste dia, data atual do sistema maior que data selecionada, Usar outro dia !", vbOKOnly + vbExclamation, "Aviso"
         Exit Sub
    End If
    
    '' Se não tiver saldo Disponivel não deixa Programar
    boolINS_LINHA = False
    If grdDIASPROG.Cell(flexcpText, grdDIASPROG.Row, conCOL_DTP_QTDISP) <= 0 Then
         intRESP = MsgBox("ATENÇÂO" & vbCrLf & _
                "Esta tentando inserir neste dia, Não Há Saldo Disponivel, Deseja realmente Inserir Programação para este Dia ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
         If intRESP = vbNo Then Exit Sub
         If intRESP = vbYes Then boolINS_LINHA = True
         
         If grdPROGRAMACAO.Row = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Selecione um registro para poder inserir !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
         End If
    End If
    
    With grdPROGRAMACAO
        
        '' Pega o Indice desta Linha
        Dim lngORDEM        As Long
        Dim lngORDEM_ORIG   As Long
        Dim lngLINSEL       As Long
        Dim lngLINVAL       As Long
        Dim i               As Long
        Dim lngCODOP        As Long
        Dim strIDPAI        As String
        
        
        If boolINS_LINHA = True Then
            '' ====================================
            lngLINSEL = .Row
            lngORDEM = CLng(.Cell(flexcpText, lngLINSEL, conCOL_PROG_ORDEM))
            lngORDEM_ORIG = lngORDEM
            lngCODOP = .Cell(flexcpText, lngLINSEL, conCOL_PROG_CODOP)
            '' ====================================
        End If
        
        .AddItem grdDIASPROG.Cell(flexcpText, grdDIASPROG.Row, conCOL_DTP_INDICEPAI) & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 -1 & vbTab & _
                 grdDIASPROG.Cell(flexcpText, grdDIASPROG.Row, conCOL_DTP_CODLIN) & vbTab & _
                 grdDIASPROG.Cell(flexcpText, grdDIASPROG.Row, conCOL_DTP_CODGRPLIN) & vbTab & _
                 grdDIASPROG.Cell(flexcpText, grdDIASPROG.Row, conCOL_DTP_IDPAI) & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & _
                 grdDIASPROG.Cell(flexcpText, grdDIASPROG.Row, conCOL_DTP_DTPROG) & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & dacEnumUpdateAction_Insert & vbTab & _
                 0 & vbTab & "" & vbTab & "" & vbTab & _
                 IIf(boolINS_LINHA = True, lngORDEM, "") & vbTab & _
                 -1
         
         
        If boolINS_LINHA = False Then
            '' ====================================
            '' Refaz o Indice pra Frente
            '' A partir da Linha Selecionada
            lngORDEM = 0
            For i = 1 To (.Rows - 1)
               If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
                  lngORDEM = (lngORDEM + 1)
                  .Cell(flexcpText, i, conCOL_PROG_ORDEM) = lngORDEM
               End If
            Next i
             '' ====================================
        Else
            For i = lngLINSEL To (.Rows - 1)
               If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete And _
                  Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 Then
                  lngORDEM = (lngORDEM + 1)
                  .Cell(flexcpText, i, conCOL_PROG_ORDEM) = lngORDEM
               End If
            Next i
        End If
        
        '' Dando um Order By na Gride de Programação
        Call OrderByGridProgramacao
        Call PintaColunasEditaveis
        
        .SetFocus
        .Col = conCOL_PROG_CODOP
        If boolINS_LINHA = False Then
            .Row = (.Rows - 1)
        ElseIf boolINS_LINHA = True Then
            ''lngLINSEL = (lngLINSEL + 1)
            ''.Row = lngLINSEL
            ''.Col = conCOL_PROG_QtdeOPProgramada
            ''.Select .Row, .Col, .Row, (.Cols - 1)
            For i = 1 To (.Rows - 1)
                If Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODPROD))) = 0 Then
                    .Row = i
                    Exit For
                End If
            Next i
        End If
        .Select .Row, .Col, .Row, (.Cols - 1)
        .EditCell
    
    End With
    
    
    
End Sub


Private Sub Desmarca(grdGENERICA As Variant, lngRowSel As Long)
    With grdGENERICA
        If lngRowSel > 0 And (.Rows - 1) > 0 Then
            If lngRowSel <= (.Rows - 1) Then
                .Cell(flexcpBackColor, lngRowSel, 0, lngRowSel, (.Cols - 1)) = &H8000000E
                .Cell(flexcpForeColor, lngRowSel, 0, lngRowSel, (.Cols - 1)) = &H80000008
                Call PintaCelula(lngRowSel)
            End If
        End If
    End With
End Sub

Private Sub MarcaLinha(grdGENERICA As Variant, lngRowSel As Long)
    With grdGENERICA
        If lngRowSel > 0 And (.Rows - 1) > 0 Then
            .Cell(flexcpBackColor, .RowSel, 0, .RowSel, (.Cols - 1)) = &H8000000D
            .Cell(flexcpForeColor, .RowSel, 0, .RowSel, (.Cols - 1)) = &H8000000E
        End If
    End With
End Sub

Private Function Valida_Campos() As Boolean

     Valida_Campos = False
     
     Dim i                  As Long
     Dim boolTEMDADOS       As Boolean
     Dim lngPESQROW         As Long
     
     With grdLinha
        If (.Rows - 1) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "A programação não foi carregada !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
        End If
     End With
     
     With grdPROGRAMACAO
        
        If (.Rows - 1) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Não foi informado OP's para Programar !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
        End If
        
        'boolTEMDADOS = False
        'For I = 1 To (.Rows - 1)
        '    If .Cell(flexcpText, I, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
        '        boolTEMDADOS = True
        '        Exit For
        '    End If
        'Next I
        'If boolTEMDADOS = False Then
        '    MsgBox "ATENÇÃO" & vbCrLf & _
        '           "Não foi informado OP's para Programar !!!", vbOKOnly + vbExclamation, "Aviso"
        '           Exit Function
        'End If
        
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
                If .Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada) = Empty Then
                    
                    lngPESQROW = grdLinha.FindRow(.Cell(flexcpText, i, conCOL_PROG_IDLINHA), , conCOL_LINHA_IDINTERNO)
                    
                    MsgBox "ATENÇÂO" & vbCrLf & _
                           "A OP : " & .Cell(flexcpText, i, conCOL_PROG_CODOP) & " Programada no Dia : " & .Cell(flexcpText, i, conCOL_PROG_DTPROG) & vbCrLf & _
                           "Na Linha : " & IIf(lngPESQROW > -1, grdLinha.Cell(flexcpText, lngPESQROW, conCOL_LINHA_DESCLIN), "") & vbCrLf & _
                           "Não pode ter valor nulo na quantidade Programada !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
            End If
        Next i
        
        
        
     End With
     
     Valida_Campos = True

End Function



Private Sub grdPROGRAMACAO_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim strINDICE       As String
    Dim lngSALDODISP    As Long
    Dim dtPROGRAMACAO   As Date
    Dim lngPEGADIA      As Long
    Dim lngSALDOP       As Long
    
     With grdPROGRAMACAO
          Select Case Col
                 Case conCOL_PROG_CODOP
                        
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(.EditText)) = 0 Then Exit Sub
                        If Not IsNumeric(.EditText) Then Exit Sub
                        
                        
                        '' Verificando se a OP já esta programada
                        If JaEstaProgramada(.EditText) = True Then
                           Cancel = True
                           Exit Sub
                        End If
                        
                        lngSALDOP = PegaSaldoOP(.EditText)
                        
                        If objBLBFunc.FcVerifItensRepetidosAct2Do(grdPROGRAMACAO, Row, conCOL_PROG_CODOP, Trim(.EditText), conCOL_PROG_Action2Do) = False Then
                           If lngSALDOP <= 0 Then
                                lngPEGADIA = .FindRow(.EditText, , conCOL_PROG_CODOP)
                                If lngPEGADIA > -1 Then
                                     MsgBox "ATENÇÃO" & vbCrLf & "A OP  " & .EditText & " já esta Programada no dia " & .Cell(flexcpText, lngPEGADIA, conCOL_PROG_DTPROG) & " !!!", vbOKOnly + vbExclamation, "Aviso"
                                Else
                                     MsgBox "ATENÇÃO" & vbCrLf & "A " & .EditText & " OP já esta Programada !!!", vbOKOnly + vbExclamation, "Aviso"
                                End If
                                Cancel = True
                                Exit Sub
                           End If
                        End If
                        
                        Cancel = PegaOP(.EditText)
                        If Cancel = True Then
                           Call LimpaColsGrid(Row)
                           Exit Sub
                        End If
                        
                        strINDICE = Trim(.Cell(flexcpText, Row, conCOL_PROG_IDDIA)) & Trim(.EditText)
                        .Cell(flexcpText, Row, conCOL_PROG_INDICEPROG) = Trim(strINDICE)
                        Call PosColCapac(conCOL_PROG_DTENTREGA, Row)
                        
                Case conCOL_PROG_DTENTREGA
                        
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(Replace(.EditText, "/", ""))) = 0 Then Exit Sub
                        If Not IsDate(.EditText) Then Exit Sub
                        
                        Call objBLBFunc.TrocaAction2Do(grdPROGRAMACAO, Row, conCOL_PROG_Action2Do, .Cell(flexcpText, Row, conCOL_PROG_DTENTREGA), .EditText)
                        Call PosColCapac(conCOL_PROG_QtdeOPProgramada, Row)
                        
                Case conCOL_PROG_QtdeOPProgramada
          
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(.EditText)) = 0 Then Exit Sub
                        If Not IsNumeric(.EditText) Then Exit Sub
                        
                        If Calcula20Porc(CLng(.Cell(flexcpText, Row, conCOL_PROG_QtdeOP)), CLng(.EditText)) = False Then
                            Cancel = True
                            Exit Sub
                        End If
                        
                        
                        Call objBLBFunc.TrocaAction2Do(grdPROGRAMACAO, Row, conCOL_PROG_Action2Do, .Cell(flexcpText, Row, conCOL_PROG_QtdeOPProgramada), .EditText)
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_PROG_QtdeOPProgramada))) = 0 Then Call Command9_Click
                        
                        
          End Select
     End With

End Sub

Private Function PegaOP(strCODOP As String) As Boolean

    PegaOP = True
    
    If Len(Trim(strCODOP)) = 0 Then Exit Function
    
    
    Dim lngSALDOP   As Long
    
    strCODOP = Trim(Replace(Replace(strCODOP, ",", ""), ".", ""))
    
    Dim lngPESQOP   As Long
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PROD.*" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_QTDE" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_STATUS" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_NECKIN" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_SALDO As SGI_SALDO_OP" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_IDPAI" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_FECHTPFU" & vbCrLf
    sSql = sSql & "      ,GRPL.SGI_CODIGO As SGI_CODGRPLIN" & vbCrLf
    sSql = sSql & "      ,LINH.SGI_CODIGO As SGI_CODLINHA" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMTABELA & " ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINH" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHAIT" & strNOMTABELA & " GRPL" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO    = " & strCODOP & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL    = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = ORDP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And LINH.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINH.SGI_CODLIN    = PROD.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "   And GRPL.SGI_FILIAL    = LINH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPL.SGI_CODLIN    = LINH.SGI_CODIGO" & vbCrLf
    
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF() Then
        If BREC4!SGI_STATUS <> 0 And _
           BREC4!SGI_STATUS <> 1 And _
           BREC4!SGI_STATUS <> 6 And _
           BREC4!SGI_STATUS <> 7 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "OP não esta com o Status de acordo para ser Inclusa !!!", vbOKOnly + vbExclamation, "Aviso"
        Else
            
            With grdPROGRAMACAO
                 If BREC4!SGI_CODGRPLIN = CLng(.Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN)) Then
                    PegaOP = False
                    lngSALDOP = PegaSaldoOP(.EditText)
                    
                    .Cell(flexcpText, .Row, conCOL_PROG_CODPROD) = Trim(BREC4!SGI_CODIGO)
                    .Cell(flexcpText, .Row, conCOL_PROG_DESCPROD) = Trim(BREC4!SGI_DESCRICAO)
                    .Cell(flexcpText, .Row, conCOL_PROG_DTENTREGA) = Format(BREC4!SGI_DATENTREGA, "DD/MM/YYYY")
                    .Cell(flexcpText, .Row, conCOL_PROG_DTENTREGABKP) = Format(BREC4!SGI_DATENTREGA, "DD/MM/YYYY")
                    .Cell(flexcpText, .Row, conCOL_PROG_QtdeOP) = BREC4!SGI_QTDE
                    
                    If lngSALDOP = 0 Then
                        .Cell(flexcpText, .Row, conCOL_PROG_QtdeOPProgramada) = BREC4!SGI_QTDE
                    Else
                        If lngSALDOP < 0 Then
                            lngSALDOP = (BREC4!SGI_QTDE + lngSALDOP)
                        End If
                        .Cell(flexcpText, .Row, conCOL_PROG_QtdeOPProgramada) = lngSALDOP
                    End If
                    
                    If BREC4!SGI_STATUS = 1 Then
                        .Cell(flexcpText, .Row, conCOL_PROG_QtdeOP) = BREC4!SGI_SALDO_OP
                        .Cell(flexcpText, .Row, conCOL_PROG_QtdeOPProgramada) = BREC4!SGI_SALDO_OP
                    End If
                    .Cell(flexcpText, .Row, conCOL_PROG_CODPED) = BREC4!SGI_CODPED
                    .Cell(flexcpText, .Row, conCOL_PROG_CODPEDBKP) = BREC4!SGI_CODPED
                    .Cell(flexcpText, .Row, conCOL_PROG_IDPRODUTO) = BREC4!SGI_IDPRODUTO
                    .Cell(flexcpText, .Row, conCOL_PROG_IDPRODBKP) = BREC4!SGI_IDPRODUTO
                    .Cell(flexcpText, .Row, conCOL_PROG_IDINTERNOOP) = BREC4!SGI_IDPAI
                    .Cell(flexcpText, .Row, conCOL_PROG_IDINTERNOOPBKP) = BREC4!SGI_IDPAI
                    .Cell(flexcpText, .Row, conCOL_PROG_CODSTATUS) = BREC4!SGI_STATUS
                    .Cell(flexcpText, .Row, conCOL_PROG_CODSTATUSBKP) = BREC4!SGI_STATUS
                    .Cell(flexcpText, .Row, conCOL_PROG_FRACIONADO) = 0
                    .Cell(flexcpText, .Row, conCOL_PROG_NECKIN) = IIf(BREC4!SGI_NECKIN = 1, "Sim", "Não")
                    .Cell(flexcpText, .Row, conCOL_PROG_FECH) = PegaFechamentoTampaFuro(BREC4!SGI_FECHTPFU)
                    .Cell(flexcpText, .Row, conCOL_PROG_COMP) = IIf(IsNull(BREC4!SGI_VernTampa) = False, PegaComp(BREC4!SGI_VernTampa), "")
                    .Cell(flexcpText, .Row, conCOL_PROG_CODLIN) = BREC4!SGI_CODLINPROD
                 Else
                    MsgBox "ATENÇÂO" & vbCrLf & _
                           "Este Produto não pertence a esta Linha !!!", vbOKOnly + vbExclamation, "Aviso"
                 End If
            End With
        End If
    Else
        MsgBox "ATENÇÃO" & vbCrLf & _
               "OP não existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC4.Close
    
End Function

Private Sub LimpaColsGrid(lngROW As Long)
    With grdPROGRAMACAO
        .Cell(flexcpText, lngROW, conCOL_PROG_CODOP) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_CODPROD) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_DESCPROD) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_DTENTREGA) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_QtdeOP) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_QtdeOPProgramada) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_CODPED) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_IDPRODUTO) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_IDINTERNOOP) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_CODSTATUS) = Empty
        .Cell(flexcpText, lngROW, conCOL_PROG_INDICEPROG) = Empty
    End With
End Sub

Private Function SaldoDisponivel(strINDICEPAI As String) As Long

    SaldoDisponivel = 0
    
    Dim lngLINPESQ      As Long
    Dim lngTOTCAPAC     As Long
    
    lngTOTCAPAC = 0
    With grdDIASPROG
        lngLINPESQ = .FindRow(strINDICEPAI, , conCOL_DTP_INDICEPAI)
        If lngLINPESQ <> -1 Then lngTOTCAPAC = .Cell(flexcpText, lngLINPESQ, conCOL_DTP_TOTDIA)
    End With
    
    SaldoDisponivel = (lngTOTCAPAC - TotalProgramado(Trim(strINDICEPAI)))
    
End Function


Private Function TotalProgramado(strINDICEPAI As String) As Long
    
    TotalProgramado = 0

    Dim i               As Long
    
    With grdPROGRAMACAO
        For i = 1 To (.Rows - 1)
            If Trim(.Cell(flexcpText, i, conCOL_PROG_IDDIA)) = Trim(strINDICEPAI) Then
                If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete And _
                   .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) <> 1 Then
                    If Len(Trim(.Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada))) > 0 Then TotalProgramado = TotalProgramado + CLng(.Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada))
                End If
            End If
        Next i
    End With

End Function

Private Sub PosColCapac(lngPOSCOL As Long, lngPOSROL As Long)
    
On Error GoTo Err_PosCol
    
    With grdPROGRAMACAO
        .SetFocus
        .Row = lngPOSROL
        .Col = lngPOSCOL
    End With
    
    Exit Sub
    
Err_PosCol:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PosCol()", Me.Name, "PosCol()", strCAMARQERRO)
    
End Sub


Private Sub PintaCelula(lngROLSEL As Long)
        If grdDIASPROG.Cell(flexcpText, lngROLSEL, conCOL_DTP_QTDISP) < 0 Then
           grdDIASPROG.Cell(flexcpBackColor, lngROLSEL, conCOL_DTP_QTDISP, lngROLSEL, conCOL_DTP_QTDISP) = vbRed
           grdDIASPROG.Cell(flexcpForeColor, lngROLSEL, conCOL_DTP_QTDISP, lngROLSEL, conCOL_DTP_QTDISP) = vbWhite
        ElseIf grdDIASPROG.Cell(flexcpText, lngROLSEL, conCOL_DTP_QTDISP) >= 0 Then
           grdDIASPROG.Cell(flexcpBackColor, lngROLSEL, conCOL_DTP_QTDISP, lngROLSEL, conCOL_DTP_QTDISP) = vbWhite
           grdDIASPROG.Cell(flexcpForeColor, lngROLSEL, conCOL_DTP_QTDISP, lngROLSEL, conCOL_DTP_QTDISP) = vbBlack
        End If
End Sub

Private Sub PintaColunasEditaveis()
    Dim i   As Long
    With grdPROGRAMACAO
        For i = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT))) = 0 Then
                .Cell(flexcpBackColor, i, conCOL_PROG_CODOP, i, conCOL_PROG_CODOP) = &H80FF&
                .Cell(flexcpBackColor, i, conCOL_PROG_DTENTREGA, i, conCOL_PROG_DTENTREGA) = &H80FF&
                .Cell(flexcpBackColor, i, conCOL_PROG_QtdeOPProgramada, i, conCOL_PROG_QtdeOPProgramada) = &H80FF&
                
                .Cell(flexcpForeColor, i, conCOL_PROG_CODOP, i, conCOL_PROG_CODOP) = vbWhite
                .Cell(flexcpForeColor, i, conCOL_PROG_DTENTREGA, i, conCOL_PROG_DTENTREGA) = vbWhite
                .Cell(flexcpForeColor, i, conCOL_PROG_QtdeOPProgramada, i, conCOL_PROG_QtdeOPProgramada) = vbWhite
            End If
        Next i
    End With
End Sub

Private Sub Imprime()
    
    If (grdLinha.Rows - 1) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "A Linha de Programação não foi carregada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If (grdDIASPROG.Rows - 1) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Os Dias e Capacidades não foi carregado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If grdLinha.RowSel <= 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Selecione uma Linha !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If optPeriodo(2).Value = True Then
        If grdDIASPROG.RowSel <= 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Selecione uma Data de Programação !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    Dim strNomRel           As String
    Dim boolTEMDADOS        As Boolean
    Dim strCABEC1           As String
    Dim strCABEC2           As String
    Dim strDATAS            As String
    Dim i                   As Long

    strDTINICIAL = "'" & Format(CDate("01/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "MM/DD/YYYY") & "'"
    If cboMes.ItemData(cboMes.ListIndex) = 12 Then
        strDTFINAL = "'" & Format(CDate("31/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)), "MM/DD/YYYY") & "'"
    Else
        strDTFINAL = "'" & Format((CDate("01/" & (cboMes.ItemData(cboMes.ListIndex) + 1) & "/" & cboAno.ItemData(cboAno.ListIndex)) - 1), "MM/DD/YYYY") & "'"
    End If
    
    
    If optPeriodo(3).Value = True Then
        strDATAS = ""
        With grdDIASPROG
            For i = 1 To (.Rows - 1)
                If .Cell(flexcpChecked, i, conCOL_DTP_SELDIA) = 1 Then
                    strDATAS = strDATAS & "'" & Format(CDate(.Cell(flexcpText, i, conCOL_DTP_DTPROG)), "MM/DD/YYYY") & "'" & ","
                End If
            Next i
            If Len(Trim(strDATAS)) = 0 Then
                MsgBox "ATENÇÂO" & vbCrLf & _
                       "Selecione Um ou Mais Dias para Impressão !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Sub
            End If
            strDATAS = Mid(strDATAS, 1, (Len(strDATAS) - 1))
        End With
    End If
    
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
    
    sSql = sSql & "     , SGI_CADFECHAM.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADFECHAM SGI_CADFECHAM" & vbCrLf
    sSql = sSql & "     , SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strNOMTABELA & " SGI_CADMOVPCP" & strNOMTABELA & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_ORDEMPROD" & strNOMTABELA & " SGI_ORDEMPROD" & strNOMTABELA & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & ".SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODIGO    = " & objCADMOVPCP.CODIGO & vbCrLf
    
    If optPeriodo(0).Value = True Then
        sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAPROG  Between '" & Format(CDate(mskDataI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDataF.Text), "MM/DD/YYYY") & "'" & vbCrLf
    ElseIf optPeriodo(1).Value = True Then
        sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAPROG  Between " & strDTINICIAL & " And " & strDTFINAL & vbCrLf
    ElseIf optPeriodo(2).Value = True Then
        sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAPROG  = '" & Format(CDate(grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_DTPROG)), "MM/DD/YYYY") & "'" & vbCrLf
    ElseIf optPeriodo(3).Value = True Then
        sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_DATAPROG  In(" & Trim(strDATAS) & ")" & vbCrLf
    
    End If
    
    ''sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODLIN    = " & grdLinha.Cell(flexcpText, grdLinha.RowSel, conCOL_LINHA_CODLIN) & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODGRPLIN = " & grdLinha.Cell(flexcpText, grdLinha.RowSel, conCOL_LINHA_CODGRPLIN) & vbCrLf
    ''sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_IDLINHA   = " & grdLinha.Cell(flexcpText, grdLinha.RowSel, conCOL_LINHA_IDINTERNO) & vbCrLf
    
    ''sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODLIN    = " & grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_CODLIN) & vbCrLf
    ''sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_CODGRPLIN = " & grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_CODGRPLIN) & vbCrLf
    ''sSql = sSql & "   And SGI_CADMOVPCP" & strNOMTABELA & ".SGI_IDLINHA   = " & grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_IDPAI) & vbCrLf
    
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

Private Function DisparaFracionamento(lngCODOP As Long, lngQTDPROGRAM As Long, lngROLATU As Long) As Date

    Dim intRESP         As Integer
    Dim lngSALDODISP    As Long
    Dim lngPROGRMADO    As Long
    Dim lngTOTCAPLINHA  As Long
    Dim lngSALDOVALIDO  As Long
    Dim lngQTDPROG      As Long
    Dim lngSALDO_OP     As Long
    Dim dtDATAPROG      As Date
    Dim lngMESVALIDO    As Long
    Dim lngIDINTERNO    As Long
    Dim lngGRPLINHA     As Long
    Dim strINDICE       As String
    
    With grdPROGRAMACAO
        boolFRACIONA = False
        
        If boolINS_LINHA = False Then
            lngSALDODISP = SaldoDisponivel(.Cell(flexcpText, lngROLATU, conCOL_PROG_IDDIA))
        Else
            lngSALDODISP = SaldoDisponivel2(.Cell(flexcpText, lngROLATU, conCOL_PROG_IDDIA), .Cell(flexcpText, lngROLATU, conCOL_PROG_ORDEM), lngROLATU)
        End If
        
        '' Se não houver saldo
        If lngSALDODISP < 0 Then
            
            If boolREMOPSEL = True Then
                intRESP = vbYes
            Else
                intRESP = ChamaResp(lngSALDODISP)
            End If
            
            If intRESP = vbYes Then
                
                boolFRACIONA = True
                Call objBLBFunc.RemoveLinhaVazia(grdPROGRAMACAO, conCOL_PROG_CODOP)
                Call objBLBFunc.RemoveLinhaVazia(grdPROGRAMACAO, conCOL_PROG_QtdeOPProgramada)
                
                lngIDINTERNO = .Cell(flexcpText, lngROLATU, conCOL_PROG_IDLINHA)
                lngGRPLINHA = .Cell(flexcpText, lngROLATU, conCOL_PROG_CODGRPLIN)
                    
                lngTOTCAPLINHA = PegaCapacLinhaDia(.Cell(flexcpText, lngROLATU, conCOL_PROG_IDDIA))
                lngSALDO_OP = PegaSaldoOP(lngCODOP)
                lngSALDOVALIDO = (lngTOTCAPLINHA - lngQTDPROGRAM)
                
                lngQTDPROG = (lngQTDPROGRAM + lngSALDODISP)
                .Cell(flexcpText, lngROLATU, conCOL_PROG_QtdeOPProgramada) = lngQTDPROG
                .Cell(flexcpText, lngROLATU, conCOL_PROG_FRACIONADO) = 1
                
                dtDATAPROG = .Cell(flexcpText, lngROLATU, conCOL_PROG_DTPROG)
                lngMESVALIDO = Month(.Cell(flexcpText, lngROLATU, conCOL_PROG_DTPROG))
                
                Do While Month(dtDATAPROG) = lngMESVALIDO
                    dtDATAPROG = (dtDATAPROG + 1)
                    ''Trim(Str(lngIDINTERNO))
                    strINDICE = Trim(Str(lngGRPLINHA)) & Trim(Str(Day(dtDATAPROG)) & Trim(Str(Month(dtDATAPROG)))) & Trim(Str(Year(dtDATAPROG)))
                    
                    lngTOTCAPLINHA = PegaCapacLinhaDia(strINDICE)
                    If lngTOTCAPLINHA > 0 Then
                        lngSALDO_OP = PegaSaldoOP(lngCODOP)
                        If lngSALDO_OP > 0 Then
                            lngSALDOVALIDO = (lngTOTCAPLINHA - lngSALDO_OP)
                            If lngSALDOVALIDO > 0 Then
                                lngQTDPROG = lngSALDO_OP
                            Else
                                lngQTDPROG = (lngSALDO_OP + lngSALDOVALIDO)
                            End If
                        
''================================
'INDICE                         1
'Código OP|Cód. Produto         2
'Descrição do Produto           3
'Dt.Entrega                     4
'Qtde.OP                        5
'Qtde.Programada                6
'Qtde.Real                      7
'CODPED                         8
'IDPRODUTO                      9
'IDINTERNOOP                    10
'CODINTENO                      11
'CODLIN                         12
'CODGRPLIN                      13
'IDLINHA                        14
'DTENTREGABKP                   15
'CODOPBKP                       16
'CODPEDBKP                      17
'IDPRODBKP                      18
'IDINTERNOOPBKP                 19
'INDICEPROG                     20
'CODSTATUS                      21
'CODSTATUSBKP                   22
'FRACIONADO                     23
'DTPROG|Action2Do               24
'NECK                           25
'FECH                           26
'COMP                           27
''================================

                            .AddItem strINDICE & vbTab & _
                                     lngCODOP & vbTab & .Cell(flexcpText, lngROLATU, conCOL_PROG_CODPROD) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_DESCPROD) & vbTab & .Cell(flexcpText, lngROLATU, conCOL_PROG_DTENTREGA) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_QtdeOP) & vbTab & lngQTDPROG & vbTab & "" & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_CODPED) & vbTab & .Cell(flexcpText, lngROLATU, conCOL_PROG_IDPRODUTO) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_IDINTERNOOP) & vbTab & _
                                     -1 & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_CODLIN) & vbTab & .Cell(flexcpText, lngROLATU, conCOL_PROG_CODGRPLIN) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_IDLINHA) & vbTab & .Cell(flexcpText, lngROLATU, conCOL_PROG_DTENTREGABKP) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_CODOPBKP) & vbTab & .Cell(flexcpText, lngROLATU, conCOL_PROG_CODPEDBKP) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_IDPRODBKP) & vbTab & .Cell(flexcpText, lngROLATU, conCOL_PROG_IDINTERNOOPBKP) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_INDICEPROG) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_CODSTATUS) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_CODSTATUSBKP) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_FRACIONADO) & vbTab & _
                                     dtDATAPROG & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_NECKIN) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_FECH) & vbTab & _
                                     .Cell(flexcpText, lngROLATU, conCOL_PROG_COMP) & vbTab & _
                                     dacEnumUpdateAction_Insert & vbTab & 0 & vbTab & _
                                     "" & vbTab & "" & vbTab & _
                                     "" & vbTab & -1
                                     
                            Call PintaColunasEditaveis
                            Call MostraOPProg
                        End If
                        If PegaSaldoOP(lngCODOP) <= 0 Then Exit Do
                    End If
                    
                Loop
                
                '' Refazendo os Indices
                Call Command1_Click
                
            ElseIf intRESP = 1 Then
                .Col = conCOL_PROG_CODOP
            ElseIf intRESP = 2 Then
                .Col = conCOL_PROG_CODOP
                boolREMOPSEL = True
            ElseIf intRESP = vbNo Then
                Call LimpaColsGrid(lngROLATU)
                .Col = conCOL_PROG_CODOP
            End If
            
        ElseIf lngSALDODISP >= 0 Then
            If boolINS_LINHA = True Then boolFRACIONA = True
            dtDATAPROG = .Cell(flexcpText, lngROLATU, conCOL_PROG_DTPROG)
        End If
    End With
    
    DisparaFracionamento = dtDATAPROG
    
End Function

Private Function SomaJaProgramado_OPs_Diferente(lngCODOP As Long, strINDICEPAI As String) As Long

    SomaJaProgramado_OPs_Diferente = 0
    
    Dim i   As Long
    
    With grdPROGRAMACAO
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
                If Trim(strINDICEPAI) = Trim(.Cell(flexcpText, i, conCOL_PROG_IDDIA)) And Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 Then
                    If lngCODOP <> CLng(.Cell(flexcpText, i, conCOL_PROG_CODOP)) Then SomaJaProgramado_OPs_Diferente = SomaJaProgramado_OPs_Diferente + CLng(.Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada))
                End If
            End If
        Next i
    End With
    
End Function

Private Function PegaCapacLinhaDia(strINDICEPAI As String) As Long
    
    PegaCapacLinhaDia = 0
    
    Dim lngLINPESQ  As Long
    
    With grdDIASPROG
        lngLINPESQ = .FindRow(strINDICEPAI, , conCOL_DTP_INDICEPAI)
        If lngLINPESQ <> -1 Then PegaCapacLinhaDia = .Cell(flexcpText, lngLINPESQ, conCOL_DTP_TOTDIA)
    End With
    
End Function

Private Function PegaJaProgramadoOPVigente(lngCODOP As Long) As Long

    PegaJaProgramadoOPVigente = 0
    
    Dim i   As Long
    With grdPROGRAMACAO
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
                If .Cell(flexcpText, i, conCOL_PROG_CODOP) = lngCODOP Then PegaJaProgramadoOPVigente = PegaJaProgramadoOPVigente + CLng(.Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada))
            End If
        Next i
    End With
    
End Function

Private Function PegaSaldoOP(lngCODOP As Long) As Long
    PegaSaldoOP = 0
    
    Dim i                   As Long
    Dim lngLINPESQ          As Long
    Dim lngQTDOP            As Long
    Dim lngPROGRAMADO       As Long
    Dim lngPROGEMOUTROMES   As Long
    
    With grdPROGRAMACAO
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
                If lngCODOP = .Cell(flexcpText, i, conCOL_PROG_CODOP) Then
                    lngQTDOP = CLng(.Cell(flexcpText, i, conCOL_PROG_QtdeOP))
                    Exit For
                End If
            End If
        Next i
    End With
    
    lngPROGRAMADO = PegaJaProgramadoOPVigente(lngCODOP)
    lngPROGEMOUTROMES = PegaSaldoOPMesesAnteriores(lngCODOP)
    
    PegaSaldoOP = (lngQTDOP - (lngPROGRAMADO + lngPROGEMOUTROMES))
    
End Function

Private Sub ApagaDemaisOPs(lngCODOP As Long)
    Dim i   As Long
    With grdPROGRAMACAO

Volta:
        
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete And _
               .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) <> 1 Then
                If lngCODOP = .Cell(flexcpText, i, conCOL_PROG_CODOP) Then
                    If .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Insert Then
                        If (.Rows - 1) = 1 Then
                            .Row = 1
                        ElseIf (.Rows - 1) > 1 Then
                            .RemoveItem i
                        End If
                        GoTo Volta
                    ElseIf .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_Ignore Or _
                       .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_update Then
                       .Cell(flexcpText, i, conCOL_PROG_Action2Do) = dacEnumUpdateAction_delete
                    End If
                End If
            End If
        Next i
    End With
End Sub

Private Sub CalcProgLinha()

    Dim i               As Long
    Dim j               As Long
    Dim lngPROGRAMADO   As Long
    Dim lngQTDCAPAC     As Long
    Dim lngTOTDISP      As Long
    
    With grdLinha
        For i = 1 To (.Rows - 1)
        
            lngPROGRAMADO = 0
            lngQTDCAPAC = 0
            lngQTDCAPAC = CLng(.Cell(flexcpText, i, conCOL_LINHA_CAPACLINHA))
            
            For j = 1 To (grdDIASPROG.Rows - 1)
                If .Cell(flexcpText, i, conCOL_LINHA_CODGRPLIN) = grdDIASPROG.Cell(flexcpText, j, conCOL_DTP_IDPAI) Then
                    lngPROGRAMADO = lngPROGRAMADO + CLng(grdDIASPROG.Cell(flexcpText, j, conCOL_DTP_PROGAM))
                End If
            Next j
            
            .Cell(flexcpText, i, conCOL_LINHA_TOTALPROG) = lngPROGRAMADO
            
            lngTOTDISP = (lngQTDCAPAC - lngPROGRAMADO)
            .Cell(flexcpText, i, conCOL_LINHA_TOTALDISP) = lngTOTDISP
            
        Next i
    End With

End Sub

Public Function PegaFechamentoTampaFuro(strCODFECH As String) As String

    If BREC10.State = 1 Then BREC10.Close
    
    PegaFechamentoTampaFuro = ""
    
    If Len(Trim(strCODFECH)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       FECH.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,FECH.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADFECHAM                 FECH" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    
    sSql = sSql & "       FECH.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And FECH.SGI_CODIGO = " & Trim(strCODFECH)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then PegaFechamentoTampaFuro = Trim(BREC10!SGI_DESCRI)
    BREC10.Close
    
End Function


Private Sub RemanejaOPs(lngCODOP As Long, strINDICE As String, dtDATAPROG As Date, lngIDINTERNO As Long, lngGRPLINHA As Long, lngORDEM As Long, lngROLATU As Long)

    Dim i                   As Long
    Dim j                   As Long
    Dim lngSALDODISP        As Long
    Dim lngQTDPROG          As Long
    Dim lngSALDO_OP         As Long
    Dim lngSALDOVALIDO      As Long
    Dim lngMESVALIDO        As Long
    Dim lngTOTCAPLINHA      As Long
    Dim lngPESQLIN          As Long
    Dim dtPROGAMACAO        As Date
    
    dtPROGAMACAO = dtDATAPROG
    lngMESVALIDO = Month(dtPROGAMACAO)
    Do While Month(dtPROGAMACAO) = lngMESVALIDO
        '' Trim(Str(lngIDINTERNO))
        strINDICE = Trim(Str(lngGRPLINHA)) & Trim(Str(Day(dtPROGAMACAO)) & Trim(Str(Month(dtPROGAMACAO)))) & Trim(Str(Year(dtPROGAMACAO)))
        
        If boolINS_LINHA = False Then
            lngSALDODISP = SaldoDisponivel(strINDICE)
        Else
            lngSALDODISP = SaldoDisponivel2(strINDICE, lngORDEM, lngROLATU)
            If lngSALDODISP = 0 Then Exit Do
        End If
        
        If lngSALDODISP <> 0 Then
            Exit Do
        End If
        dtPROGAMACAO = (dtPROGAMACAO + 1)
    Loop
    
    Call ConfGrdAux
    
    With grdPROGRAMACAO
        
        lngMESVALIDO = Month(dtPROGAMACAO)
        
        Do While Month(dtPROGAMACAO) = lngMESVALIDO
            
            lngTOTCAPLINHA = PegaCapacLinhaDia(strINDICE)
            If lngTOTCAPLINHA > 0 Then
                For i = 1 To (.Rows - 1)
                    If Trim(.Cell(flexcpText, i, conCOL_PROG_IDDIA)) = Trim(strINDICE) And _
                       .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete And _
                       .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) <> 1 Then
                       
                       If boolINS_LINHA = False Then
                            
                            If .Cell(flexcpText, i, conCOL_PROG_CODOP) <> lngCODOP And _
                               Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 Then
                                 lngPESQLIN = grdAUX.FindRow(.Cell(flexcpText, i, conCOL_PROG_CODOP), , conCOL_PROG_CODOP)
                                 If lngPESQLIN = -1 Then
                                        grdAUX.AddItem strINDICE & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPROD) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_DESCPROD) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGA) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_QtdeOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_QtdeOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_QtdeReal) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODPED) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDPRODUTO) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOP) & vbTab & _
                                                      -1 & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODLIN) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODGRPLIN) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_IDLINHA) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODOPBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_INDICEPROG) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODSTATUS) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_FRACIONADO) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_DTPROG) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_NECKIN) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_FECH) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_COMP) & vbTab & _
                                                      dacEnumUpdateAction_Insert & vbTab & 0 & vbTab & _
                                                      "" & vbTab & "" & _
                                                      "" & vbTab & -1
                                 End If
                            End If
                       ElseIf boolINS_LINHA = True Then
                            If .Cell(flexcpText, i, conCOL_PROG_ORDEM) > lngORDEM And _
                               Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 Then
                                 lngPESQLIN = grdAUX.FindRow(.Cell(flexcpText, i, conCOL_PROG_CODOP), , conCOL_PROG_CODOP)
                                 If lngPESQLIN = -1 Then
                                        grdAUX.AddItem strINDICE & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPROD) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_DESCPROD) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGA) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_QtdeOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_QtdeOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_QtdeReal) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODPED) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDPRODUTO) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOP) & vbTab & _
                                                      -1 & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODLIN) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODGRPLIN) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_IDLINHA) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODOPBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_INDICEPROG) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODSTATUS) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_FRACIONADO) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_DTPROG) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_NECKIN) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_FECH) & vbTab & _
                                                      .Cell(flexcpText, i, conCOL_PROG_COMP) & vbTab & _
                                                      dacEnumUpdateAction_Insert & vbTab & 0 & vbTab & _
                                                      "" & vbTab & "" & _
                                                      "" & vbTab & -1
                                 End If
                            End If
                       
                       End If
                       
                    End If
                Next i
            End If
            dtPROGAMACAO = (dtPROGAMACAO + 1)
            '' Trim(Str(lngIDINTERNO))
            strINDICE = Trim(Str(lngGRPLINHA)) & Trim(Str(Day(dtPROGAMACAO)) & Trim(Str(Month(dtPROGAMACAO)))) & Trim(Str(Year(dtPROGAMACAO)))
        Loop
        
    End With

    
    With grdAUX
        
        '' Removendo as OP'Selecionadas
        For i = 1 To (.Rows - 1)
            '' Apagando as OPs Relacionadas
            If Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 Then Call ApagaDemaisOPs(.Cell(flexcpText, i, conCOL_PROG_CODOP))
        Next i
        
        For i = 1 To (.Rows - 1)
        
            '' Realocando as OP's
            dtPROGAMACAO = .Cell(flexcpText, i, conCOL_PROG_DTPROG)
            ''If boolREMOPSEL = True Then dtPROGAMACAO = (dtPROGAMACAO + 1)
            
            lngMESVALIDO = Month(dtPROGAMACAO)
            
            Do While Month(dtPROGAMACAO) = lngMESVALIDO
                '' Trim(Str(lngIDINTERNO))
                strINDICE = Trim(Str(lngGRPLINHA)) & Trim(Str(Day(dtPROGAMACAO)) & Trim(Str(Month(dtPROGAMACAO)))) & Trim(Str(Year(dtPROGAMACAO)))
                lngSALDODISP = SaldoDisponivel(strINDICE)
                If lngSALDODISP <> 0 Then
                    .Cell(flexcpText, i, conCOL_PROG_DTPROG) = dtPROGAMACAO
                    Exit Do
                End If
                dtPROGAMACAO = (dtPROGAMACAO + 1)
            Loop
        
            '' Se não houver saldo
            If lngSALDODISP <> 0 Then
                    
                lngTOTCAPLINHA = PegaCapacLinhaDia(strINDICE)
                dtPROGAMACAO = .Cell(flexcpText, i, conCOL_PROG_DTPROG)
                
                If lngSALDODISP < 0 Then lngQTDPROG = (lngTOTCAPLINHA + lngSALDODISP)
                If lngSALDODISP > 0 Then
                    If .Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada) < lngSALDODISP Then
                        lngQTDPROG = .Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada)
                    Else
                        lngQTDPROG = lngSALDODISP
                    End If
                End If
                
                .Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada) = lngQTDPROG
                .Cell(flexcpText, i, conCOL_PROG_FRACIONADO) = 1
                
                grdPROGRAMACAO.AddItem strINDICE & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_CODOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPROD) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_DESCPROD) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGA) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_QtdeOP) & vbTab & lngQTDPROG & vbTab & .Cell(flexcpText, i, conCOL_PROG_QtdeReal) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_CODPED) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDPRODUTO) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOP) & vbTab & _
                                        -1 & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_CODLIN) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODGRPLIN) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_IDLINHA) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_CODOPBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_INDICEPROG) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODSTATUS) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_FRACIONADO) & vbTab & _
                                        dtPROGAMACAO & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_NECKIN) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_FECH) & vbTab & _
                                       .Cell(flexcpText, i, conCOL_PROG_COMP) & vbTab & _
                                       dacEnumUpdateAction_Insert & vbTab & 0 & vbTab & _
                                       "" & vbTab & "" & vbTab & _
                                       ""
                
                lngMESVALIDO = Month(dtPROGAMACAO)
                
                
                Do While Month(dtPROGAMACAO) = lngMESVALIDO
                    dtPROGAMACAO = (dtPROGAMACAO + 1)
                    '' Trim(Str(lngIDINTERNO))
                    strINDICE = Trim(Str(lngGRPLINHA)) & Trim(Str(Day(dtPROGAMACAO)) & Trim(Str(Month(dtPROGAMACAO)))) & Trim(Str(Year(dtPROGAMACAO)))
                    
                    lngTOTCAPLINHA = PegaCapacLinhaDia(strINDICE)
                    If lngTOTCAPLINHA > 0 Then
                        If Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 Then lngSALDO_OP = PegaSaldoOP(.Cell(flexcpText, i, conCOL_PROG_CODOP))
                        If lngSALDO_OP > 0 Then
                            lngSALDOVALIDO = (lngTOTCAPLINHA - lngSALDO_OP)
                            If lngSALDOVALIDO > 0 Then
                                lngQTDPROG = lngSALDO_OP
                            Else
                                lngQTDPROG = (lngSALDO_OP + lngSALDOVALIDO)
                            End If
                            
                            grdPROGRAMACAO.AddItem strINDICE & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_CODOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPROD) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_DESCPROD) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGA) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_QtdeOP) & vbTab & lngQTDPROG & vbTab & .Cell(flexcpText, i, conCOL_PROG_QtdeReal) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_CODPED) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDPRODUTO) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOP) & vbTab & _
                                        -1 & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_CODLIN) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODGRPLIN) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_IDLINHA) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_CODOPBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_INDICEPROG) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_CODSTATUS) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_FRACIONADO) & vbTab & _
                                        dtPROGAMACAO & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_NECKIN) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_FECH) & vbTab & _
                                        .Cell(flexcpText, i, conCOL_PROG_COMP) & vbTab & _
                                        dacEnumUpdateAction_Insert & vbTab & 0 & vbTab & _
                                        "" & vbTab & "" & vbTab & _
                                        ""
                        Else
                            Exit Do
                        End If
                    End If
                Loop
            End If
        Next i
    End With

    '' Refazendo os Indices
    If boolINS_LINHA = True Or boolREMOPSEL = True Then Call Command1_Click

    Call PintaColunasEditaveis
    Call MostraOPProg
    Call ConfGrdAux

End Sub

Private Sub ConfGrdAux()

    With grdAUX

       .Cols = conColumnsIn_PROG
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_PROG_FormatString
       .AutoSizeMouse = False
       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_PROG_IDDIA) = ""
       .ColDataType(conCOL_PROG_IDDIA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_CODOP) = ""
       .ColDataType(conCOL_PROG_CODOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPROD) = ""
       .ColDataType(conCOL_PROG_CODPROD) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_DESCPROD) = ""
       .ColDataType(conCOL_PROG_DESCPROD) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_DTENTREGA) = ""
       .ColDataType(conCOL_PROG_DTENTREGA) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeOP) = ""
       .ColDataType(conCOL_PROG_QtdeOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeOPProgramada) = ""
       .ColDataType(conCOL_PROG_QtdeOPProgramada) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeReal) = ""
       .ColDataType(conCOL_PROG_QtdeReal) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPED) = ""
       .ColDataType(conCOL_PROG_CODPED) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDPRODUTO) = ""
       .ColDataType(conCOL_PROG_IDPRODUTO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDINTERNOOP) = ""
       .ColDataType(conCOL_PROG_IDINTERNOOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODINTENO) = ""
       .ColDataType(conCOL_PROG_CODINTENO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODLIN) = ""
       .ColDataType(conCOL_PROG_CODLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODGRPLIN) = ""
       .ColDataType(conCOL_PROG_CODGRPLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDLINHA) = ""
       .ColDataType(conCOL_PROG_IDLINHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_DTENTREGABKP) = ""
       .ColDataType(conCOL_PROG_DTENTREGABKP) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PROG_CODOPBKP) = ""
       .ColDataType(conCOL_PROG_CODOPBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPEDBKP) = ""
       .ColDataType(conCOL_PROG_CODPEDBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDPRODBKP) = ""
       .ColDataType(conCOL_PROG_IDPRODBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDINTERNOOPBKP) = ""
       .ColDataType(conCOL_PROG_IDINTERNOOPBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_INDICEPROG) = ""
       .ColDataType(conCOL_PROG_INDICEPROG) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_CODSTATUS) = ""
       .ColDataType(conCOL_PROG_CODSTATUS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODSTATUSBKP) = ""
       .ColDataType(conCOL_PROG_CODSTATUSBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_NECKIN) = ""
       .ColDataType(conCOL_PROG_NECKIN) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_FECH) = ""
       .ColDataType(conCOL_PROG_FECH) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_COMP) = ""
       .ColDataType(conCOL_PROG_COMP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_Action2Do) = ""
       .ColDataType(conCOL_PROG_Action2Do) = flexDTLong
       
       .ColWidth(conCOL_PROG_IDDIA) = 0
       .ColWidth(conCOL_PROG_CODOP) = 1000
       .ColWidth(conCOL_PROG_CODPROD) = 1000
       .ColWidth(conCOL_PROG_DESCPROD) = 5000
       .ColWidth(conCOL_PROG_DTENTREGA) = 900
       .ColWidth(conCOL_PROG_QtdeOP) = 1100
       .ColWidth(conCOL_PROG_QtdeOPProgramada) = 1300
       .ColWidth(conCOL_PROG_QtdeReal) = 1100
       .ColWidth(conCOL_PROG_CODPED) = 0
       .ColWidth(conCOL_PROG_IDPRODUTO) = 0
       .ColWidth(conCOL_PROG_IDINTERNOOP) = 0
       .ColWidth(conCOL_PROG_CODINTENO) = 0
       .ColWidth(conCOL_PROG_CODLIN) = 0
       .ColWidth(conCOL_PROG_CODGRPLIN) = 0
       .ColWidth(conCOL_PROG_IDLINHA) = 0
       .ColWidth(conCOL_PROG_DTENTREGABKP) = 0
       .ColWidth(conCOL_PROG_CODOPBKP) = 0
       .ColWidth(conCOL_PROG_CODPEDBKP) = 0
       .ColWidth(conCOL_PROG_IDPRODBKP) = 0
       .ColWidth(conCOL_PROG_IDINTERNOOPBKP) = 0
       .ColWidth(conCOL_PROG_INDICEPROG) = 0
       .ColWidth(conCOL_PROG_CODSTATUS) = 0
       .ColWidth(conCOL_PROG_CODSTATUSBKP) = 0
       .ColWidth(conCOL_PROG_FRACIONADO) = 0
       .ColWidth(conCOL_PROG_DTPROG) = 0
       .ColWidth(conCOL_PROG_NECKIN) = 500
       .ColWidth(conCOL_PROG_FECH) = 500
       .ColWidth(conCOL_PROG_COMP) = 500
       .ColWidth(conCOL_PROG_Action2Do) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       .FontName = "Arial"
       .FontSize = 7
       .FontBold = True
       
    End With
    
End Sub

Private Function PegaComp(lngCOMP As Long) As String
    PegaComp = ""
    If lngCOMP = 1 Then PegaComp = "VEX"
    If lngCOMP = 2 Then PegaComp = "VZ"
    If lngCOMP = 3 Then PegaComp = "NAT"
    If lngCOMP = 4 Then PegaComp = "VI"
End Function

Private Function Calcula20Porc(lngQTDOP As Long, lgnQTDPROG As Long) As Boolean
    
    Calcula20Porc = False
    
    Dim lngRESULTADO    As Long
    Dim lngQTDOPSUP     As Long
    
    lngQTDOPSUP = (lngQTDOP + (lngQTDOP * 0.2))
    
    If lgnQTDPROG > lngQTDOPSUP Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "A Quantidade programada ultrapassou a 20% da quantidade da OP !", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
    Calcula20Porc = True
    
End Function
Private Function PegaDescrStatusApontamento(strCODSTATUS As String)

    PegaDescrStatusApontamento = ""
    
    If Len(Trim(strCODSTATUS)) = 0 Then Exit Function
    If CLng(strCODSTATUS) = 0 Then Exit Function

    If CLng(strCODSTATUS) = 1 Then PegaDescrStatusApontamento = "Concluido"
    If CLng(strCODSTATUS) = 2 Then PegaDescrStatusApontamento = "Parcial"
    If CLng(strCODSTATUS) = 3 Then PegaDescrStatusApontamento = "Em Produção"
    
End Function

Private Sub PintaOPAPontada()

    Dim i    As Long
    With grdPROGRAMACAO
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
                If .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) = 1 Then        '' Concluido
                   .Cell(flexcpBackColor, i, conCOL_PROG_CODOP, i, (.Cols - 1)) = &H4000& '' Verde
                   .Cell(flexcpForeColor, i, conCOL_PROG_CODOP, i, (.Cols - 1)) = vbWhite  '' Branco
                ElseIf .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) = 2 Then    '' Parcial
                    .Cell(flexcpBackColor, i, conCOL_PROG_CODOP, i, (.Cols - 1)) = &HFFFF& '' Amarelo
                    .Cell(flexcpForeColor, i, conCOL_PROG_CODOP, i, (.Cols - 1)) = vbBlack '' Preto
                ElseIf .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) = 3 Then  '' Em Produção / Azul
                    .Cell(flexcpBackColor, i, conCOL_PROG_CODOP, i, (.Cols - 1)) = &HFF8080
                End If
            End If
        Next i
    End With
    
End Sub

Private Sub OrderByGridProgramacao()

    Dim i As Long
    
    With grdPROGRAMACAO
          '' Dando um Order By na Gride de Programação
          If (.Rows - 1) > 0 Then .Cell(flexcpSort, 1, conCOL_PROG_ORDEM, (.Rows - 1), conCOL_PROG_ORDEM) = flexSortNumericAscending
    End With
End Sub

Private Function SaldoDisponivel2(strINDICEPAI As String, lngORDEM As Long, lngLINHA As Long) As Long

    SaldoDisponivel2 = 0
    
    Dim lngLINPESQ      As Long
    Dim lngTOTCAPAC     As Long
    Dim lngLINHA2       As Long
    lngLINHA2 = lngLINHA
    
    lngTOTCAPAC = 0
    With grdDIASPROG
        lngLINPESQ = .FindRow(strINDICEPAI, , conCOL_DTP_INDICEPAI)
        If lngLINPESQ <> -1 Then lngTOTCAPAC = .Cell(flexcpText, lngLINPESQ, conCOL_DTP_TOTDIA)
    End With
    
    SaldoDisponivel2 = (lngTOTCAPAC - TotalProgramado3(Trim(strINDICEPAI), lngORDEM))
    
End Function

Private Function TotalProgramado2(strINDICEPAI As String, lngLINHA As Long) As Long
    
    TotalProgramado2 = 0

    Dim i               As Long
    lngLINHA = (lngLINHA + 1)
    
    With grdPROGRAMACAO
        For i = lngLINHA To (.Rows - 1)
            If Trim(.Cell(flexcpText, i, conCOL_PROG_IDDIA)) = Trim(strINDICEPAI) Then
                If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete Then
                    If Len(Trim(.Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada))) > 0 Then TotalProgramado2 = TotalProgramado2 + CLng(.Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada))
                End If
            End If
        Next i
    End With

End Function

Private Function TotalProgramado3(strINDICEPAI As String, lngORDEM As Long) As Long
    
    TotalProgramado3 = 0

    Dim i               As Long
    
    With grdPROGRAMACAO
        For i = 1 To (.Rows - 1)
            If Trim(.Cell(flexcpText, i, conCOL_PROG_IDDIA)) = Trim(strINDICEPAI) Then
                If .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete And _
                   .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) <> 1 Then
                    If Len(Trim(.Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada))) > 0 Then TotalProgramado3 = TotalProgramado3 + CLng(.Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada))
                    If lngORDEM = CLng(.Cell(flexcpText, i, conCOL_PROG_ORDEM)) Then Exit For
                End If
            End If
        Next i
    End With

End Function


Private Sub mskDataF_GotFocus()
    objBLBFunc.SelecionaCampos mskDataF.Name, Me
End Sub

Private Sub mskDataI_GotFocus()
    objBLBFunc.SelecionaCampos mskDataI.Name, Me
End Sub

Private Sub optPeriodo_Click(Index As Integer)
    Frame8.Visible = False
    Call MostraCel(0)
    If Index = 0 Then Frame8.Visible = True
    If Index = 3 Then Call MostraCel(200)
End Sub

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
    
    If optPeriodo(0).Value = True Then
        If Month(CDate(mskDataI.Text)) <> cboMes.ItemData(cboMes.ListIndex) Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "O Mês da Data inicial não pode ser diferente no Mês de Programação !!!"
            mskDataI.Text = CDate("01/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex))
            mskDataI.SetFocus
            ConsisteCampos = False
            Exit Function
            
        End If
        If Month(CDate(mskDataF.Text)) <> cboMes.ItemData(cboMes.ListIndex) Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "O Mês da Data final não pode ser diferente no Mês de Programação !!!"
            
            If cboMes.ItemData(cboMes.ListIndex) = 12 Then
                strDTFINAL = "31/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)
                mskDataF.Text = (CDate(strDTFINAL) - 1)
            Else
                strDTFINAL = "01/" & (cboMes.ItemData(cboMes.ListIndex) + 1) & "/" & cboAno.ItemData(cboAno.ListIndex)
                mskDataF.Text = (CDate(strDTFINAL) - 1)
            End If
            
            mskDataF.SetFocus
            ConsisteCampos = False
            Exit Function
        End If
    End If

End Function

Private Sub MostraCel(lngTamCol As Long)
    Dim i   As Long
    With grdDIASPROG
        For i = 1 To (.Rows - 1)
            .ColWidth(conCOL_DTP_SELDIA) = lngTamCol
        Next i
    End With
End Sub

Private Sub Timer1_Timer()
    
        
    If cTipOper = "C" Then
        Dim lngROLLINANT    As Long
        Dim lngROLDIAANT    As Long
        Dim lngROLOPANT     As Long
        
        lngROLLINANT = grdLinha.Row
        lngROLDIAANT = grdDIASPROG.Row
        lngROLOPANT = grdPROGRAMACAO.Row
        
        Call IniciaForm
    
        grdLinha.Row = lngROLLINANT
        grdLinha.RowSel = lngROLLINANT
        grdLinha.TopRow = lngSCROOLMANT
        
        grdDIASPROG.Row = lngROLDIAANT
        grdDIASPROG.RowSel = lngROLDIAANT
        grdDIASPROG.TopRow = lngSCROOLDANT
        
        If lngROLOPANT > 0 Then grdPROGRAMACAO.Row = lngROLOPANT
        If lngROLOPANT > 0 Then grdPROGRAMACAO.RowSel = lngROLOPANT
        grdPROGRAMACAO.TopRow = lngSCROOLPANT
    End If
    
End Sub


Private Function ChamaResp(lngSALDODISP As Long) As Integer
    
    frmRESP.intRETORNO = vbNo
    frmRESP.lngSALDODISP = lngSALDODISP
    frmRESP.Show vbModal

    ChamaResp = frmRESP.intRETORNO
    
End Function

Private Sub MoveOP()
            
    If grdDIASPROG.Row = 0 Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Selecione um dia de Programação", vbOKOnly + vbExclamation, "Aviso"
               Exit Sub
    End If
    
    Dim intRESP             As Integer
    Dim dtPROGRAMACAO       As Date
    
    intRESP = MsgBox("ATENÇÂO" & vbCrLf & _
                     "As OP's Selecionada(s) Serão Realocadas(s) !" & vbCrLf & _
                     "Tem Certeza ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                     
    If intRESP = vbNo Then
        boolREMOPSEL = False
        boolFRACIONA = False
        Exit Sub
    End If

    With grdPROGRAMACAO
        
        
        .Col = conCOL_PROG_CODOP
        boolFRACIONA = True
        boolREMOPSEL = True
        dtPROGRAMACAO = .Cell(flexcpText, .Row, conCOL_PROG_DTPROG)
    
        If Len(Trim(.Cell(flexcpText, .Row, conCOL_PROG_CODOP))) > 0 Then
            
            
            If VerifDiasSelMov(dtPROGRAMACAO, .Cell(flexcpText, .Row, conCOL_PROG_IDLINHA), .Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN)) = False Then
                MsgBox "ATENÇÂO" & vbCrLf & _
                       "Selecione uma ou mais OP's para remanejar !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Sub
            End If
            
            Call SelDiasHaFrente(dtPROGRAMACAO, .Cell(flexcpText, .Row, conCOL_PROG_IDLINHA), .Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN))
            
''            dtPROGRAMACAO = DisparaFracionamento(.Cell(flexcpText, .Row, conCOL_PROG_CODOP), .Cell(flexcpText, .Row, conCOL_PROG_QtdeOPProgramada), .Row)
            
            
            If boolFRACIONA = True Then
                Call RemanejaOPs(.Cell(flexcpText, .Row, conCOL_PROG_CODOP), _
                                 Trim(.Cell(flexcpText, .Row, conCOL_PROG_IDDIA)), _
                                 dtPROGRAMACAO, _
                                 .Cell(flexcpText, .Row, conCOL_PROG_IDLINHA), _
                                 .Cell(flexcpText, .Row, conCOL_PROG_CODGRPLIN), _
                                 .Cell(flexcpText, .Row, conCOL_PROG_ORDEM), _
                                 .Row)
            End If
        End If
        boolINS_LINHA = False
        boolFRACIONA = False
        boolREMOPSEL = False
    
        grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_PROGAM) = TotalProgramado(Trim(.Cell(flexcpText, .Row, conCOL_PROG_IDDIA)))
        grdDIASPROG.Cell(flexcpText, grdDIASPROG.RowSel, conCOL_DTP_QTDISP) = SaldoDisponivel(Trim(.Cell(flexcpText, .Row, conCOL_PROG_IDDIA)))
        
        Call PintaCelula(grdDIASPROG.RowSel)
        Call CalcQTdJaProg
        Call CalcProgLinha
    
        Call MarcaLinha(grdDIASPROG, grdDIASPROG.RowSel)
    
    
    End With
    
End Sub

Private Sub SelDiasHaFrente(dtPROG As Date, lngIDINTERNO As Long, lngGRPLINHA As Long)

    Dim i               As Long
    Dim lngMESVALIDO    As Long
    Dim lngPESQ         As Long
    Dim dtPROGRAMACAO   As Date
    Dim strINDICE       As String
    
    
    dtPROGRAMACAO = (dtPROG + 1)
    
    With grdPROGRAMACAO
        lngMESVALIDO = Month(dtPROGRAMACAO)
        Do While Month(dtPROGRAMACAO) = lngMESVALIDO
            strINDICE = Trim(Str(lngIDINTERNO)) & Trim(Str(Day(dtPROGRAMACAO)) & Trim(Str(Month(dtPROGRAMACAO)))) & Trim(Str(Year(dtPROGRAMACAO))) & Trim(Str(lngGRPLINHA))
            lngPESQ = .FindRow(strINDICE, , conCOL_PROG_IDDIA)
            If lngPESQ <> -1 Then
                For i = 1 To (.Rows - 1)
                    If Trim(.Cell(flexcpText, i, conCOL_PROG_IDDIA)) = Trim(strINDICE) And _
                       .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete And Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 Then
                       .Cell(flexcpChecked, i, conCOL_PROG_Marca) = 1
                    End If
                Next i
            End If
            dtPROGRAMACAO = (dtPROGRAMACAO + 1)
        Loop
    End With
End Sub


Private Function VerifDiasSelMov(dtPROG As Date, lngIDINTERNO As Long, lngGRPLINHA As Long) As Boolean

    Dim i               As Long
    Dim lngPESQ         As Long
    Dim dtPROGRAMACAO   As Date
    Dim strINDICE       As String
    Dim boolSelecionado As Boolean
    
    dtPROGRAMACAO = dtPROG
    VerifDiasSelMov = False
    
    With grdPROGRAMACAO
        strINDICE = Trim(Str(lngIDINTERNO)) & Trim(Str(Day(dtPROGRAMACAO)) & Trim(Str(Month(dtPROGRAMACAO)))) & Trim(Str(Year(dtPROGRAMACAO))) & Trim(Str(lngGRPLINHA))
        lngPESQ = .FindRow(strINDICE, , conCOL_PROG_IDDIA)
        If lngPESQ <> -1 Then
            For i = 1 To (.Rows - 1)
                If Trim(.Cell(flexcpText, i, conCOL_PROG_IDDIA)) = Trim(strINDICE) And _
                   .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete And _
                   Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 And _
                   .Cell(flexcpChecked, i, conCOL_PROG_Marca) = 1 Then
                   VerifDiasSelMov = True
                End If
            Next i
        End If
    End With
    
End Function



Private Sub RetiraSelDiasHaFrente(dtPROG As Date, lngIDINTERNO As Long, lngGRPLINHA As Long)

    Dim i               As Long
    Dim lngMESVALIDO    As Long
    Dim lngPESQ         As Long
    Dim dtPROGRAMACAO   As Date
    Dim strINDICE       As String
    
    dtPROGRAMACAO = dtPROG
    
    With grdPROGRAMACAO
        lngMESVALIDO = Month(dtPROGRAMACAO)
        Do While Month(dtPROGRAMACAO) = lngMESVALIDO
            strINDICE = Trim(Str(lngIDINTERNO)) & Trim(Str(Day(dtPROGRAMACAO)) & Trim(Str(Month(dtPROGRAMACAO)))) & Trim(Str(Year(dtPROGRAMACAO))) & Trim(Str(lngGRPLINHA))
            lngPESQ = .FindRow(strINDICE, , conCOL_PROG_IDDIA)
            If lngPESQ <> -1 Then
                For i = 1 To (.Rows - 1)
                    If Trim(.Cell(flexcpText, i, conCOL_PROG_IDDIA)) = Trim(strINDICE) And _
                       .Cell(flexcpText, i, conCOL_PROG_Action2Do) <> dacEnumUpdateAction_delete _
                       And Len(Trim(.Cell(flexcpText, i, conCOL_PROG_CODOP))) > 0 And _
                       .Cell(flexcpChecked, i, conCOL_PROG_Marca) = 1 Then
                       .Cell(flexcpChecked, i, conCOL_PROG_Marca) = 2
                    End If
                Next i
            End If
            dtPROGRAMACAO = (dtPROGRAMACAO + 1)
        Loop
    End With
End Sub



Private Sub MoveIten(lngLINHA_ATU As Long, strTIPO As String)

    If cTipOper = "C" Then
        MsgBox "Você está no modo de Consulta escolha 'Alterar' !!!", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
    
    If boolINS_LINHA = True Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Não é permitido esta no modo de Inserir Linha !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If boolREMOPSEL = True Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Não é permitido esta no modo de excluir OP !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    
    If lngLINHA_ATU = 0 Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Selecione uma Linha !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If

    Dim i           As Long
    Dim lngPROX_LIN As Long
    Dim lngLINAUX   As Long
    
    i = lngLINHA_ATU
    
    If strTIPO = "B" Then lngPROX_LIN = (lngLINHA_ATU + 1)
    If strTIPO = "C" Then lngPROX_LIN = (lngLINHA_ATU - 1)
    
    Call ConfGrdAux
                                        
    With grdPROGRAMACAO
        
        If lngPROX_LIN > (.Rows - 1) Or lngPROX_LIN <= 0 Then
           MsgBox "ATENÇÂO" & vbCrLf & "Não é permitido avançar !!!", vbOKOnly + vbExclamation, "aviso"
           Exit Sub
        End If
        
        '' Verificando se esta no Msm Dia
        If .Cell(flexcpText, i, conCOL_PROG_IDDIA) <> .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDDIA) Then
           MsgBox "ATENÇÂO" & vbCrLf & "Não é permitido avançar !!!", vbOKOnly + vbExclamation, "aviso"
           Exit Sub
        End If
    
        '' Pegando o Item Atual
        grdAUX.AddItem .Cell(flexcpText, i, conCOL_PROG_IDDIA) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_CODOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPROD) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_DESCPROD) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGA) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_QtdeOP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada) & vbTab & .Cell(flexcpText, i, conCOL_PROG_QtdeReal) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_CODPED) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDPRODUTO) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOP) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_CODINTENO) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_CODLIN) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODGRPLIN) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_IDLINHA) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_CODOPBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP) & vbTab & .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_INDICEPROG) & vbTab & .Cell(flexcpText, i, conCOL_PROG_CODSTATUS) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_FRACIONADO) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_DTPROG) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_NECKIN) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_FECH) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_COMP) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_Action2Do) & vbTab & .Cell(flexcpText, i, conCOL_PROG_Marca) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) & vbTab & .Cell(flexcpText, i, conCOL_PROG_DESCSTATSAPONT) & vbTab & _
                       .Cell(flexcpText, i, conCOL_PROG_ORDEM) & vbTab & .Cell(flexcpText, i, conCOL_PROG_TIMESTAMP)
        
        lngLINAUX = (grdAUX.Rows - 1)
        
        .Cell(flexcpText, i, conCOL_PROG_IDDIA) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDDIA)
        .Cell(flexcpText, i, conCOL_PROG_CODOP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODOP)
        .Cell(flexcpText, i, conCOL_PROG_CODPROD) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODPROD)
        .Cell(flexcpText, i, conCOL_PROG_DESCPROD) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DESCPROD)
        .Cell(flexcpText, i, conCOL_PROG_DTENTREGA) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DTENTREGA)
        .Cell(flexcpText, i, conCOL_PROG_QtdeOP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_QtdeOP)
        .Cell(flexcpText, i, conCOL_PROG_QtdeOPProgramada) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_QtdeOPProgramada)
        .Cell(flexcpText, i, conCOL_PROG_QtdeReal) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_QtdeReal)
        .Cell(flexcpText, i, conCOL_PROG_CODPED) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODPED)
        .Cell(flexcpText, i, conCOL_PROG_IDPRODUTO) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDPRODUTO)
        .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDINTERNOOP)
        .Cell(flexcpText, i, conCOL_PROG_CODINTENO) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODINTENO)
        .Cell(flexcpText, i, conCOL_PROG_CODLIN) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODLIN)
        .Cell(flexcpText, i, conCOL_PROG_CODGRPLIN) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODGRPLIN)
        .Cell(flexcpText, i, conCOL_PROG_IDLINHA) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDLINHA)
        .Cell(flexcpText, i, conCOL_PROG_DTENTREGABKP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DTENTREGABKP)
        .Cell(flexcpText, i, conCOL_PROG_CODOPBKP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODOPBKP)
        .Cell(flexcpText, i, conCOL_PROG_CODPEDBKP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODPEDBKP)
        .Cell(flexcpText, i, conCOL_PROG_IDPRODBKP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDPRODBKP)
        .Cell(flexcpText, i, conCOL_PROG_IDINTERNOOPBKP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDINTERNOOPBKP)
        .Cell(flexcpText, i, conCOL_PROG_INDICEPROG) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_INDICEPROG)
        .Cell(flexcpText, i, conCOL_PROG_CODSTATUS) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODSTATUS)
        .Cell(flexcpText, i, conCOL_PROG_CODSTATUSBKP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODSTATUSBKP)
        .Cell(flexcpText, i, conCOL_PROG_FRACIONADO) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_FRACIONADO)
        .Cell(flexcpText, i, conCOL_PROG_DTPROG) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DTPROG)
        .Cell(flexcpText, i, conCOL_PROG_NECKIN) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_NECKIN)
        .Cell(flexcpText, i, conCOL_PROG_FECH) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_FECH)
        .Cell(flexcpText, i, conCOL_PROG_COMP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_COMP)
        .Cell(flexcpText, i, conCOL_PROG_Action2Do) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_Action2Do)
        .Cell(flexcpText, i, conCOL_PROG_Marca) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_Marca)
        .Cell(flexcpText, i, conCOL_PROG_CODSTATUSAPONT) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODSTATUSAPONT)
        .Cell(flexcpText, i, conCOL_PROG_DESCSTATSAPONT) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DESCSTATSAPONT)
        .Cell(flexcpText, i, conCOL_PROG_ORDEM) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_ORDEM)
        .Cell(flexcpText, i, conCOL_PROG_TIMESTAMP) = .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_TIMESTAMP)
        
        
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDDIA) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_IDDIA)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODOP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODOP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODPROD) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODPROD)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DESCPROD) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_DESCPROD)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DTENTREGA) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_DTENTREGA)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_QtdeOP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_QtdeOP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_QtdeOPProgramada) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_QtdeOPProgramada)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_QtdeReal) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_QtdeReal)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODPED) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODPED)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDPRODUTO) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_IDPRODUTO)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDINTERNOOP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_IDINTERNOOP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODINTENO) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODINTENO)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODLIN) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODLIN)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODGRPLIN) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODGRPLIN)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDLINHA) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_IDLINHA)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DTENTREGABKP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_DTENTREGABKP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODOPBKP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODOPBKP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODPEDBKP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODPEDBKP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDPRODBKP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_IDPRODBKP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_IDINTERNOOPBKP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_IDINTERNOOPBKP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_INDICEPROG) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_INDICEPROG)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODSTATUS) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODSTATUS)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODSTATUSBKP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODSTATUSBKP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_FRACIONADO) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_FRACIONADO)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DTPROG) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_DTPROG)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_NECKIN) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_NECKIN)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_FECH) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_FECH)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_COMP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_COMP)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_Action2Do) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_Action2Do)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_Marca) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_Marca)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_CODSTATUSAPONT) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_CODSTATUSAPONT)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_DESCSTATSAPONT) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_DESCSTATSAPONT)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_ORDEM) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_ORDEM)
        .Cell(flexcpText, lngPROX_LIN, conCOL_PROG_TIMESTAMP) = grdAUX.Cell(flexcpText, lngLINAUX, conCOL_PROG_TIMESTAMP)
        
        Call Command1_Click
        
        .Row = lngPROX_LIN
        .RowSel = lngPROX_LIN
    
         Call PintaColunasEditaveis
         Call MarcaLinha(grdPROGRAMACAO, .RowSel)
    
    End With
End Sub

Private Sub CarregaOP_NaoApontadas()

    Dim i               As Long
    Dim lngIDINTERNO    As Long
    Dim lngGRPLINHA     As Long
    Dim lngPESQOP       As Long
    Dim lngROLATU       As Long
    Dim lngMESVALIDO    As Long
    Dim lngSALDODISP    As Long
    Dim strINDICE       As String
    Dim dtPROGRAMACAO   As Date
    
    Call ConfGrdExc
    
    With grdDIASPROG
    
        For i = 1 To (.Rows - 1)
        
            lngIDINTERNO = .Cell(flexcpText, i, conCOL_DTP_IDPAI)
            lngGRPLINHA = .Cell(flexcpText, i, conCOL_DTP_CODGRPLIN)
            strINDICE = .Cell(flexcpText, i, conCOL_DTP_INDICEPAI)
            
            If CDate(.Cell(flexcpText, i, conCOL_DTP_DTPROG)) = Date Then
            
                dtPROGRAMACAO = .Cell(flexcpText, i, conCOL_DTP_DTPROG)
                
                sSql = ""
                
                sSql = "Select" & vbCrLf
                sSql = sSql & "       MOVPCP.*" & vbCrLf
                sSql = sSql & "     , PROD.SGI_CODIGO       As SGI_CODROTULO" & vbCrLf
                sSql = sSql & "     , PROD.SGI_DESCRICAO    As SGI_DESCOTULO" & vbCrLf
                sSql = sSql & "     , PROD.SGI_NECKIN" & vbCrLf
                sSql = sSql & "     , PROD.SGI_VernTampa" & vbCrLf
                sSql = sSql & "     , ORDP.SGI_QTDE" & vbCrLf
                sSql = sSql & "     , ORDP.SGI_FECHTPFU" & vbCrLf
                sSql = sSql & "     , ORDP.SGI_DATENTREGA" & vbCrLf
                
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & " MOVPCP" & vbCrLf
                sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMTABELA & " ORDP" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
                
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       MOVPCP.SGI_FILIAL   = " & FILIAL & vbCrLf
                sSql = sSql & "   And MOVPCP.SGI_IDLINHA  = " & lngIDINTERNO & vbCrLf
                sSql = sSql & "   And (MOVPCP.SGI_STATUSAPONT Is Null or MOVPCP.SGI_STATUSAPONT = 2)" & vbCrLf
                
                sSql = sSql & "   And ORDP.SGI_FILIAL     = MOVPCP.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And ORDP.SGI_IDPAI      = MOVPCP.SGI_IDINTERNO" & vbCrLf
                sSql = sSql & "   And PROD.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ORDP.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "Order By MOVPCP.SGI_DATAPROG,MOVPCP.SGI_ORDEM"
                
                With grdPROGRAMACAO
                    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                    Do While Not BREC3.EOF()
                        
                        '' =========================================================================
                        '' Colocando na Grid Para Exclusão
                        
                        grdEXC.AddItem strINDICE & vbTab & BREC3!SGI_CODOP & vbTab & _
                                       BREC3!SGI_CODROTULO & vbTab & BREC3!SGI_DESCOTULO & vbTab & _
                                       Format(BREC3!SGI_DATAENTR, "DD/MM/YYYY") & vbTab & _
                                       BREC3!SGI_QTDE & vbTab & BREC3!SGI_QTDEPROD & vbTab & _
                                       "" & vbTab & BREC3!SGI_CODPED & vbTab & _
                                       BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                                       BREC3!SGI_CODINTENO & vbTab & BREC3!SGI_CODLIN & vbTab & _
                                       BREC3!SGI_CODGRPLIN & vbTab & BREC3!SGI_IDLINHA & vbTab & _
                                       Format(BREC3!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                                       BREC3!SGI_CODOP & vbTab & BREC3!SGI_CODPED & vbTab & _
                                       BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                                       Trim(strINDICE) & Trim(Str(BREC3!SGI_CODOP)) & vbTab & _
                                       BREC3!SGI_CODSTATUS & vbTab & _
                                       BREC3!SGI_STATUSORIG & vbTab & _
                                       BREC3!SGI_FRACIONADA & vbTab & _
                                       dtPROGRAMACAO & vbTab & _
                                       IIf(BREC3!SGI_NECKIN = 1, "Sim", "Não") & vbTab & _
                                       PegaFechamentoTampaFuro(Str(BREC3!SGI_FECHTPFU)) & vbTab & _
                                       IIf(IsNull(BREC3!SGI_VernTampa) = False, PegaComp(BREC3!SGI_VernTampa), "") & vbTab & _
                                       dacEnumUpdateAction_delete & vbTab & 0 & vbTab & _
                                       "" & vbTab & "" & vbTab & _
                                       BREC3!SGI_ORDEM & vbTab & _
                                       IIf(IsNull(BREC3!SGI_TimeStamp), -1, BREC3!SGI_TimeStamp)
                                       
                        lngROLATU = (.Rows - 1)
                                                     
                        If Not IsNull(BREC3!SGI_STATUSAPONT) Then
                            .Cell(flexcpText, lngROLATU, conCOL_PROG_CODSTATUSAPONT) = BREC3!SGI_STATUSAPONT
                            .Cell(flexcpText, lngROLATU, conCOL_PROG_QtdeReal) = BREC3!SGI_QTDEAPONTADA
                            .Cell(flexcpText, lngROLATU, conCOL_PROG_DESCSTATSAPONT) = PegaDescrStatusApontamento(Str(BREC3!SGI_STATUSAPONT))
                        End If
                                       
                                       
                        '' =========================================================================
                        
                        ''lngPESQOP = .FindRow(BREC3!SGI_CODOP, , conCOL_PROG_CODOP)
                        ''If lngPESQOP = -1 Then
                        ''    .AddItem strINDICE & vbTab & BREC3!SGI_CODOP & vbTab & _
                        ''             BREC3!SGI_CODROTULO & vbTab & BREC3!SGI_DESCOTULO & vbTab & _
                        ''             Format(BREC3!SGI_DATAENTR, "DD/MM/YYYY") & vbTab & _
                        ''             BREC3!SGI_QTDE & vbTab & BREC3!SGI_QTDEPROD & vbTab & _
                        ''             "" & vbTab & BREC3!SGI_CODPED & vbTab & _
                        ''             BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                        ''             BREC3!SGI_CODINTENO & vbTab & BREC3!SGI_CODLIN & vbTab & _
                        ''             BREC3!SGI_CODGRPLIN & vbTab & BREC3!SGI_IDLINHA & vbTab & _
                        ''             Format(BREC3!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                        ''             BREC3!SGI_CODOP & vbTab & BREC3!SGI_CODPED & vbTab & _
                        ''             BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                        ''             Trim(strINDICE) & Trim(Str(BREC3!SGI_CODOP)) & vbTab & _
                        ''             BREC3!SGI_CODSTATUS & vbTab & _
                        ''             BREC3!SGI_STATUSORIG & vbTab & _
                        ''             BREC3!SGI_FRACIONADA & vbTab & _
                        ''             dtPROGRAMACAO & vbTab & _
                        ''             IIf(BREC3!SGI_NECKIN = 1, "Sim", "Não") & vbTab & _
                        ''             PegaFechamentoTampaFuro(Str(BREC3!SGI_FECHTPFU)) & vbTab & _
                        ''             IIf(IsNull(BREC3!SGI_VernTampa) = False, PegaComp(BREC3!SGI_VernTampa), "") & vbTab & _
                        ''             dacEnumUpdateAction_Insert & vbTab & 0 & vbTab & _
                        ''             "" & vbTab & "" & vbTab & _
                        ''             BREC3!SGI_ORDEM & vbTab & _
                        ''             IIf(IsNull(BREC3!SGI_TimeStamp), -1, BREC3!SGI_TimeStamp)
                                                     
                            
                        ''    boolFRACIONA = False
                        ''    boolINS_LINHA = False
                        ''    boolREMOPSEL = True
                        ''    dtPROGRAMACAO = DisparaFracionamento(BREC3!SGI_CODOP, BREC3!SGI_QTDE, lngROLATU)

                            '' ======================================================
                            '' Vendo Saldo Disponivel e posiciionando no Dia
                        ''    lngMESVALIDO = Month(dtPROGRAMACAO)
                        ''    Do While Month(dtPROGRAMACAO) = lngMESVALIDO
                        ''        strINDICE = Trim(Str(lngIDINTERNO)) & Trim(Str(Day(dtPROGRAMACAO)) & Trim(Str(Month(dtPROGRAMACAO)))) & Trim(Str(Year(dtPROGRAMACAO))) & Trim(Str(lngGRPLINHA))
                        ''
                        ''        If boolINS_LINHA = False Then
                        ''            lngSALDODISP = SaldoDisponivel(strINDICE)
                        ''        Else
                                ''    lngSALDODISP = SaldoDisponivel2(strINDICE, lngORDEM, lngROLATU)
                                ''    If lngSALDODISP = 0 Then Exit Do
                        ''        End If
                                
                        ''        If lngSALDODISP <> 0 Then
                        ''            Exit Do
                        ''        End If
                        ''        dtPROGRAMACAO = (dtPROGRAMACAO + 1)
                        ''    Loop
                            '' ======================================================
                        ''End If
                        BREC3.MoveNext
                        
                    Loop
                    BREC3.Close
                End With
                
            End If
        Next i
    End With
    
End Sub

Private Sub ConfGrdExc()

    With grdEXC

       .Cols = conColumnsIn_PROG
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_PROG_FormatString
       .AutoSizeMouse = False
       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_PROG_IDDIA) = ""
       .ColDataType(conCOL_PROG_IDDIA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_CODOP) = ""
       .ColDataType(conCOL_PROG_CODOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPROD) = ""
       .ColDataType(conCOL_PROG_CODPROD) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_DESCPROD) = ""
       .ColDataType(conCOL_PROG_DESCPROD) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_DTENTREGA) = ""
       .ColDataType(conCOL_PROG_DTENTREGA) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeOP) = ""
       .ColDataType(conCOL_PROG_QtdeOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeOPProgramada) = ""
       .ColDataType(conCOL_PROG_QtdeOPProgramada) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_QtdeReal) = ""
       .ColDataType(conCOL_PROG_QtdeReal) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPED) = ""
       .ColDataType(conCOL_PROG_CODPED) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDPRODUTO) = ""
       .ColDataType(conCOL_PROG_IDPRODUTO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDINTERNOOP) = ""
       .ColDataType(conCOL_PROG_IDINTERNOOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODINTENO) = ""
       .ColDataType(conCOL_PROG_CODINTENO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODLIN) = ""
       .ColDataType(conCOL_PROG_CODLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODGRPLIN) = ""
       .ColDataType(conCOL_PROG_CODGRPLIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDLINHA) = ""
       .ColDataType(conCOL_PROG_IDLINHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_DTENTREGABKP) = ""
       .ColDataType(conCOL_PROG_DTENTREGABKP) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PROG_CODOPBKP) = ""
       .ColDataType(conCOL_PROG_CODOPBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODPEDBKP) = ""
       .ColDataType(conCOL_PROG_CODPEDBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDPRODBKP) = ""
       .ColDataType(conCOL_PROG_IDPRODBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_IDINTERNOOPBKP) = ""
       .ColDataType(conCOL_PROG_IDINTERNOOPBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_INDICEPROG) = ""
       .ColDataType(conCOL_PROG_INDICEPROG) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_CODSTATUS) = ""
       .ColDataType(conCOL_PROG_CODSTATUS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_CODSTATUSBKP) = ""
       .ColDataType(conCOL_PROG_CODSTATUSBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PROG_NECKIN) = ""
       .ColDataType(conCOL_PROG_NECKIN) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_FECH) = ""
       .ColDataType(conCOL_PROG_FECH) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_COMP) = ""
       .ColDataType(conCOL_PROG_COMP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PROG_Action2Do) = ""
       .ColDataType(conCOL_PROG_Action2Do) = flexDTLong
       
       .ColWidth(conCOL_PROG_IDDIA) = 0
       .ColWidth(conCOL_PROG_CODOP) = 1000
       .ColWidth(conCOL_PROG_CODPROD) = 1000
       .ColWidth(conCOL_PROG_DESCPROD) = 5000
       .ColWidth(conCOL_PROG_DTENTREGA) = 900
       .ColWidth(conCOL_PROG_QtdeOP) = 1100
       .ColWidth(conCOL_PROG_QtdeOPProgramada) = 1300
       .ColWidth(conCOL_PROG_QtdeReal) = 1100
       .ColWidth(conCOL_PROG_CODPED) = 0
       .ColWidth(conCOL_PROG_IDPRODUTO) = 0
       .ColWidth(conCOL_PROG_IDINTERNOOP) = 0
       .ColWidth(conCOL_PROG_CODINTENO) = 0
       .ColWidth(conCOL_PROG_CODLIN) = 0
       .ColWidth(conCOL_PROG_CODGRPLIN) = 0
       .ColWidth(conCOL_PROG_IDLINHA) = 0
       .ColWidth(conCOL_PROG_DTENTREGABKP) = 0
       .ColWidth(conCOL_PROG_CODOPBKP) = 0
       .ColWidth(conCOL_PROG_CODPEDBKP) = 0
       .ColWidth(conCOL_PROG_IDPRODBKP) = 0
       .ColWidth(conCOL_PROG_IDINTERNOOPBKP) = 0
       .ColWidth(conCOL_PROG_INDICEPROG) = 0
       .ColWidth(conCOL_PROG_CODSTATUS) = 0
       .ColWidth(conCOL_PROG_CODSTATUSBKP) = 0
       .ColWidth(conCOL_PROG_FRACIONADO) = 0
       .ColWidth(conCOL_PROG_DTPROG) = 0
       .ColWidth(conCOL_PROG_NECKIN) = 500
       .ColWidth(conCOL_PROG_FECH) = 500
       .ColWidth(conCOL_PROG_COMP) = 500
       .ColWidth(conCOL_PROG_Action2Do) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       .FontName = "Arial"
       .FontSize = 7
       .FontBold = True
       
    End With
    
End Sub


Private Function PegaQtdeOPJaApontada(lngCODOP As Long) As Long

End Function


Private Function PegaSaldoOPMesesAnteriores(lngCODOP As Long) As Long
        
    PegaSaldoOPMesesAnteriores = 0
    
    '' VERIFICANDO SE EXISTE A op EM MESES ANTERIORES
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Sum(MOV.SGI_QTDEPROD) As SGI_QTDEPROD" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & " MOV" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       MOV.SGI_FILIAL          = " & FILIAL & vbCrLf
    sSql = sSql & "   And MOV.SGI_CODOP           = " & lngCODOP & vbCrLf
    sSql = sSql & "   And Month(MOV.SGI_DATAPROG) < " & cboMes.ItemData(cboMes.ListIndex) & vbCrLf

    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then
        If Not IsNull(BREC11!SGI_QTDEPROD) Then PegaSaldoOPMesesAnteriores = BREC11!SGI_QTDEPROD
    End If
    BREC11.Close
    

End Function


Private Sub PopGrdProgramadoMesDiaAnterior()

    '' dtDATAPROG As Date, lngCODLIN As Long, strINDICE As String

    Dim dtDATAPROG      As Date
    Dim strINDICE       As String
    Dim lngMESATU       As Long
    Dim lngPESQ         As Long
    Dim lngActioToDo    As Long
    Dim i               As Long
    Dim j               As Long

    dtDATAPROG = Date
    lngMESATU = Month(dtDATAPROG)
    
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       SGI_IDLINHA" & vbCrLf
    sSql = sSql & "     , SGI_CODGRPLIN" & vbCrLf
    sSql = sSql & "     , SGI_DATAPROG" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAl = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_DATAPROG < '" & Format(dtDATAPROG, "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And (SGI_STATUSAPONT Is Null or SGI_STATUSAPONT = 2)" & vbCrLf
    sSql = sSql & "Order By SGI_IDLINHA,SGI_CODGRPLIN,SGI_DATAPROG"
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC4.EOF()
        
        
        '' --------------------------------------------
        '' Pegando a OP
        sSql = ""
        sSql = "Select" & vbCrLf
        sSql = sSql & "       MOVPCP.*" & vbCrLf
        sSql = sSql & "     , PROD.SGI_CODIGO       As SGI_CODROTULO" & vbCrLf
        sSql = sSql & "     , PROD.SGI_DESCRICAO    As SGI_DESCOTULO" & vbCrLf
        sSql = sSql & "     , PROD.SGI_NECKIN" & vbCrLf
        sSql = sSql & "     , PROD.SGI_VernTampa" & vbCrLf
        sSql = sSql & "     , ORDP.SGI_QTDE" & vbCrLf
        sSql = sSql & "     , ORDP.SGI_FECHTPFU" & vbCrLf
        sSql = sSql & "     , ORDP.SGI_DATENTREGA" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & " MOVPCP" & vbCrLf
        sSql = sSql & "     , SGI_ORDEMPROD" & strNOMTABELA & " ORDP" & vbCrLf
        sSql = sSql & "     , SGI_CADPRODUTO      PROD" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       MOVPCP.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And MOVPCP.SGI_DATAPROG   = '" & Format(BREC4!SGI_DATAPROG, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And MOVPCP.SGI_IDLINHA    = " & BREC4!SGI_IDLINHA & vbCrLf
        sSql = sSql & "   And MOVPCP.SGI_CODGRPLIN  = " & BREC4!SGI_CODGRPLIN & vbCrLf
        sSql = sSql & "   And (MOVPCP.SGI_STATUSAPONT Is Null or MOVPCP.SGI_STATUSAPONT = 2)" & vbCrLf
        
        sSql = sSql & "   And ORDP.SGI_FILIAL     = MOVPCP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And ORDP.SGI_IDPAI      = MOVPCP.SGI_IDINTERNO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ORDP.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "Order By MOVPCP.SGI_DATAPROG,MOVPCP.SGI_ORDEM"
        
        BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC3.EOF() Then
            
            Do While Not BREC3.EOF()
                
                If BREC4!SGI_DATAPROG < dtDATAPROG Then
                    ''Trim(Str(BREC3!SGI_IDLINHA))
                    strINDICE = Trim(Str(BREC3!SGI_CODGRPLIN)) & Trim(Str(Day(dtDATAPROG)) & Trim(Str(Month(dtDATAPROG)))) & Trim(Str(Year(dtDATAPROG)))
                    lngActioToDo = dacEnumUpdateAction_Insert
                ''ElseIf BREC4!SGI_DATAPROG > dtDATAPROG Then
                ''   strINDICE = Trim(Str(BREC3!SGI_IDLINHA)) & Trim(Str(Day(BREC4!SGI_DATAPROG)) & Trim(Str(Month(BREC4!SGI_DATAPROG)))) & Trim(Str(Year(BREC4!SGI_DATAPROG))) & Trim(Str(BREC3!SGI_CODGRPLIN))
                ''    lngActioToDo = dacEnumUpdateAction_Ignore
                End If
                
                Do While lngMESATU = Month(dtDATAPROG)
                    lngPESQ = grdDIASPROG.FindRow(strINDICE, , conCOL_DTP_INDICEPAI)
                    If lngPESQ <> -1 Then Exit Do
                    dtDATAPROG = (dtDATAPROG + 1)
                    ''Trim(Str(BREC3!SGI_IDLINHA))
                    strINDICE = Trim(Str(BREC3!SGI_CODGRPLIN)) & Trim(Str(Day(dtDATAPROG)) & Trim(Str(Month(dtDATAPROG)))) & Trim(Str(Year(dtDATAPROG)))
                Loop
                
                If BREC4!SGI_DATAPROG < dtDATAPROG Then
                    With grdEXC
                        
                        ''Trim(Str(BREC3!SGI_IDLINHA))
                        strINDICE = Trim(Str(BREC3!SGI_CODGRPLIN)) & Trim(Str(Day(BREC4!SGI_DATAPROG)) & Trim(Str(Month(BREC4!SGI_DATAPROG)))) & Trim(Str(Year(BREC4!SGI_DATAPROG)))
                        .AddItem strINDICE & vbTab & BREC3!SGI_CODOP & vbTab & _
                                 BREC3!SGI_CODROTULO & vbTab & BREC3!SGI_DESCOTULO & vbTab & _
                                 Format(BREC3!SGI_DATAENTR, "DD/MM/YYYY") & vbTab & _
                                 BREC3!SGI_QTDE & vbTab & BREC3!SGI_QTDEPROD & vbTab & _
                                 "" & vbTab & BREC3!SGI_CODPED & vbTab & _
                                 BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                                 BREC3!SGI_CODINTENO & vbTab & BREC3!SGI_CODLIN & vbTab & _
                                 BREC3!SGI_CODGRPLIN & vbTab & BREC3!SGI_IDLINHA & vbTab & _
                                 Format(BREC3!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                                 BREC3!SGI_CODOP & vbTab & BREC3!SGI_CODPED & vbTab & _
                                 BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                                 Trim(strINDICE) & Trim(Str(BREC3!SGI_CODOP)) & vbTab & _
                                 BREC3!SGI_CODSTATUS & vbTab & _
                                 BREC3!SGI_STATUSORIG & vbTab & _
                                 BREC3!SGI_FRACIONADA & vbTab & _
                                 BREC4!SGI_DATAPROG & vbTab & _
                                 IIf(BREC3!SGI_NECKIN = 1, "Sim", "Não") & vbTab & _
                                 PegaFechamentoTampaFuro(Str(BREC3!SGI_FECHTPFU)) & vbTab & _
                                 IIf(IsNull(BREC3!SGI_VernTampa) = False, PegaComp(BREC3!SGI_VernTampa), "") & vbTab & _
                                 dacEnumUpdateAction_delete & vbTab & 0 & vbTab & _
                                 "" & vbTab & "" & vbTab & _
                                 BREC3!SGI_ORDEM & vbTab & _
                                 IIf(IsNull(BREC3!SGI_TimeStamp), -1, BREC3!SGI_TimeStamp)
            
                    End With
                End If
                
                With grdPROGRAMACAO
                    
                    ''Trim(Str(BREC3!SGI_IDLINHA))
                    strINDICE = Trim(Str(BREC3!SGI_CODGRPLIN)) & Trim(Str(Day(dtDATAPROG)) & Trim(Str(Month(dtDATAPROG)))) & Trim(Str(Year(dtDATAPROG)))
                    .AddItem strINDICE & vbTab & BREC3!SGI_CODOP & vbTab & _
                             BREC3!SGI_CODROTULO & vbTab & BREC3!SGI_DESCOTULO & vbTab & _
                             Format(BREC3!SGI_DATAENTR, "DD/MM/YYYY") & vbTab & _
                             BREC3!SGI_QTDE & vbTab & BREC3!SGI_QTDEPROD & vbTab & _
                             "" & vbTab & BREC3!SGI_CODPED & vbTab & _
                             BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                             BREC3!SGI_CODINTENO & vbTab & BREC3!SGI_CODLIN & vbTab & _
                             BREC3!SGI_CODGRPLIN & vbTab & BREC3!SGI_IDLINHA & vbTab & _
                             Format(BREC3!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                             BREC3!SGI_CODOP & vbTab & BREC3!SGI_CODPED & vbTab & _
                             BREC3!SGI_IDPRODUTO & vbTab & BREC3!SGI_IDINTERNO & vbTab & _
                             Trim(strINDICE) & Trim(Str(BREC3!SGI_CODOP)) & vbTab & _
                             BREC3!SGI_CODSTATUS & vbTab & _
                             BREC3!SGI_STATUSORIG & vbTab & _
                             BREC3!SGI_FRACIONADA & vbTab & _
                             dtDATAPROG & vbTab & _
                             IIf(BREC3!SGI_NECKIN = 1, "Sim", "Não") & vbTab & _
                             PegaFechamentoTampaFuro(Str(BREC3!SGI_FECHTPFU)) & vbTab & _
                             IIf(IsNull(BREC3!SGI_VernTampa) = False, PegaComp(BREC3!SGI_VernTampa), "") & vbTab & _
                             lngActioToDo & vbTab & 0 & vbTab & _
                             "" & vbTab & "" & vbTab & _
                             BREC3!SGI_ORDEM & vbTab & _
                             IIf(IsNull(BREC3!SGI_TimeStamp), -1, BREC3!SGI_TimeStamp)
                                             
                    If Not IsNull(BREC3!SGI_STATUSAPONT) Then
                        .Cell(flexcpText, (.Rows - 1), conCOL_PROG_CODSTATUSAPONT) = BREC3!SGI_STATUSAPONT
                        .Cell(flexcpText, (.Rows - 1), conCOL_PROG_QtdeReal) = BREC3!SGI_QTDEAPONTADA
                        .Cell(flexcpText, (.Rows - 1), conCOL_PROG_DESCSTATSAPONT) = PegaDescrStatusApontamento(Str(BREC3!SGI_STATUSAPONT))
                    End If
                End With
                
                
                BREC3.MoveNext
            Loop
        End If
        BREC3.Close
        
        ''dtDATAPROG = (dtDATAPROG + 1)
        BREC4.MoveNext
    Loop
    BREC4.Close
    
    
    '' Verificando Para Excluir da Gride
    With grdEXC
        For i = 1 To (.Rows - 1)
            For j = 1 To (grdPROGRAMACAO.Rows - 1)
                If .Cell(flexcpText, i, conCOL_PROG_CODOP) = grdPROGRAMACAO.Cell(flexcpText, j, conCOL_PROG_CODOP) And _
                   .Cell(flexcpText, i, conCOL_PROG_IDDIA) = grdPROGRAMACAO.Cell(flexcpText, j, conCOL_PROG_IDDIA) Then
                   grdPROGRAMACAO.Cell(flexcpText, j, conCOL_PROG_Action2Do) = dacEnumUpdateAction_delete
                End If
            Next j
        Next i
    End With
    
End Sub



Private Function JaEstaProgramada(lngCODOP As Long) As Boolean

    JaEstaProgramada = False
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      (Sum(MOV.SGI_QTDEPROD) - OP.SGI_QTDE) As SGI_QTDEPROD" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & " MOV" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMTABELA & " OP" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       MOV.SGI_FILIAL          = " & FILIAL & vbCrLf
    sSql = sSql & "   And MOV.SGI_CODOP           = " & lngCODOP & vbCrLf
    sSql = sSql & "   And Month(MOV.SGI_DATAPROG) < " & cboMes.ItemData(cboMes.ListIndex) & vbCrLf
    sSql = sSql & "   And OP.SGI_FILIAL           = MOV.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And OP.SGI_CODIGO           = MOV.SGI_CODOP" & vbCrLf
    sSql = sSql & "Group By MOV.SGI_QTDEPROD,OP.SGI_QTDE"
    
    BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC12.EOF() Then
        If Not IsNull(BREC12!SGI_QTDEPROD) Then
            If BREC12!SGI_QTDEPROD = 0 Then JaEstaProgramada = True
        End If
    End If
    BREC12.Close
    
    If JaEstaProgramada = True Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Esta OP: " & lngCODOP & " já esta programada em outro Mês !!!"
    End If

End Function



