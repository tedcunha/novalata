VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADQTDTURNOS 
   Caption         =   "Cadastro de quantidade de turnos"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Turnos"
      TabPicture(0)   =   "frmCADQTDTURNOS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraParadas"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Operadores"
      TabPicture(1)   =   "frmCADQTDTURNOS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdOperadores"
      Tab(1).Control(1)=   "cmdIncIten"
      Tab(1).Control(2)=   "cmdExcIten"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Máquinas"
      TabPicture(2)   =   "frmCADQTDTURNOS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdMAQUINAS"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraParadas 
         Caption         =   "[ Paradas ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   8535
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   8040
            Picture         =   "frmCADQTDTURNOS.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   8040
            Picture         =   "frmCADQTDTURNOS.frx":0192
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   600
            Width           =   375
         End
         Begin VSFlex8LCtl.VSFlexGrid grdParadas 
            Height          =   2655
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   7815
            _cx             =   13785
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
      Begin VSFlex8LCtl.VSFlexGrid grdMAQUINAS 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   8535
         _cx             =   15055
         _cy             =   10610
         Appearance      =   1
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
      Begin VB.CommandButton cmdExcIten 
         Height          =   315
         Left            =   -66720
         Picture         =   "frmCADQTDTURNOS.frx":071C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdIncIten 
         Height          =   315
         Left            =   -66720
         Picture         =   "frmCADQTDTURNOS.frx":0CA6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VSFlex8LCtl.VSFlexGrid grdOperadores 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   8055
         _cx             =   14208
         _cy             =   10821
         Appearance      =   1
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
      Begin VB.Frame Frame4 
         Height          =   3015
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   8535
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   8040
            Picture         =   "frmCADQTDTURNOS.frx":0DE4
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   8040
            Picture         =   "frmCADQTDTURNOS.frx":136E
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   120
            Width           =   375
         End
         Begin VSFlex8LCtl.VSFlexGrid grdPeriodo 
            Height          =   2775
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   7815
            _cx             =   13785
            _cy             =   4895
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
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8775
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
         Picture         =   "frmCADQTDTURNOS.frx":14AC
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
         MaskColor       =   &H8000000F&
         Picture         =   "frmCADQTDTURNOS.frx":19DE
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADQTDTURNOS.frx":1AE0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   8775
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   960
         Width           =   2295
         Begin VB.OptionButton optAtivoSN 
            Caption         =   "Não"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   16
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optAtivoSN 
            Caption         =   "Sim"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1450
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   7215
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ativo:"
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
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   510
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
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   930
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
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmCADQTDTURNOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Public strUSUARIO   As String
Public iCodigo      As Integer
Public cTipOper     As String
Dim objBLBFunc      As Object
Dim objCADQTDTURNOS As Object
Dim objPESQPADRAO   As Object
Dim arrDIASSEMANA   As Variant
Dim arrOPERADORES   As Variant
Dim arrPARADAS      As Variant

Const conCOL_SonPe_Periodo                          As Integer = 0
Const conCOL_SonPe_HorEnt                           As Integer = 1
Const conCOL_SonPe_HorSai                           As Integer = 2
Const conCOL_SonPe_HorPar                           As Integer = 3
Const conCOL_SonPe_QtdPar                           As Integer = 4
Const conCOL_SonPe_TotalLiq                         As Integer = 5
Const conCOL_SonPe_FormatString                     As String = "=Periodo|Hor. Entrada|Hor. Saida|Paradas|Qtd. Paradas|Hor. Liquido"
Const conColumnsIn_SonPe                            As Integer = 6

Const conCOL_SonPa_Parada                           As Integer = 0
Const conCOL_SonPa_HorIni                           As Integer = 1
Const conCOL_SonPa_HorFin                           As Integer = 2
Const conCOL_SonPa_Total                            As Integer = 3
Const conCOL_SonPa_ComParada                        As Integer = 4
Const conCOL_SonPa_Pai                              As Integer = 5
Const conCOL_SonPa_FormatString                     As String = "=Parada|Hor. Inicial|Hor. Final|Total|Com parada|Pai"
Const conColumnsIn_SonPa                            As Integer = 6

Const conCOL_SonOper_CodOper                        As Integer = 0
Const conCOL_SonOper_PesqOper                       As Integer = 1
Const conCOL_SonOper_Desc_Oper                      As Integer = 2
Const conCOL_SonOper_FormatString                   As String = "=Cód. Operador|...|Descrição Operador"
Const conColumnsIn_SonOper                          As Integer = 3

Const conCOL_SonMaq_CodMaq                          As Integer = 0
Const conCOL_SonMaq_Desc_Maq                        As Integer = 1
Const conCOL_SonMaq_FormatString                    As String = "=Cód. Máquina|Descrição da Máquina"
Const conColumnsIn_SonMaq                           As Integer = 2


Private Sub cmdAltera_Click()
    
    Dim I As Integer
    Dim j As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de quantidade de turnos - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
End Sub

Private Sub cmdExcIten_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdOperadores, grdOperadores.Row)
End Sub

Private Sub cmdIncIten_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGrid
End Sub

Private Sub CmdSalva_Click()
    
    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADQTDTURNOS.CODIGO = objCADQTDTURNOS.Gera_Codigo(Me.Name)
    
    objCADQTDTURNOS.DESCRI = txtDescricao.Text
    If optAtivoSN(1).Value = True Then objCADQTDTURNOS.ATIVO = 1
    If optAtivoSN(0).Value = True Then objCADQTDTURNOS.ATIVO = 0
    
    '' ======================================================================
    arrDIASSEMANA = Empty
    With grdPeriodo
        If (.Rows - 1) > 0 Then
           ReDim arrDIASSEMANA(1 To (.Rows - 1), 1 To 6) As String
           For I = 1 To (.Rows - 1)
               arrDIASSEMANA(I, 1) = .Cell(flexcpText, I, conCOL_SonPe_Periodo)
               arrDIASSEMANA(I, 2) = .Cell(flexcpText, I, conCOL_SonPe_HorEnt)
               arrDIASSEMANA(I, 3) = .Cell(flexcpText, I, conCOL_SonPe_HorSai)
               If Len(Trim(.Cell(flexcpText, I, conCOL_SonPe_HorPar))) > 0 Then
                    arrDIASSEMANA(I, 4) = .Cell(flexcpText, I, conCOL_SonPe_HorPar)
               Else
                    arrDIASSEMANA(I, 4) = "Null"
               End If
               If Len(Trim(.Cell(flexcpText, I, conCOL_SonPe_QtdPar))) > 0 Then
                    arrDIASSEMANA(I, 5) = .Cell(flexcpText, I, conCOL_SonPe_QtdPar)
               Else
                    arrDIASSEMANA(I, 5) = "Null"
               End If
               arrDIASSEMANA(I, 6) = .Cell(flexcpText, I, conCOL_SonPe_TotalLiq)
           Next I
        End If
    End With
    objCADQTDTURNOS.DIASSEMANA = arrDIASSEMANA
    
    '' ======================================================================
    arrOPERADORES = Empty
    If (grdOperadores.Rows - 1) > 0 Then
       ReDim arrOPERADORES(1 To (grdOperadores.Rows - 1)) As String
       For I = 1 To (grdOperadores.Rows - 1)
            arrOPERADORES(I) = grdOperadores.Cell(flexcpText, I, conCOL_SonOper_CodOper)
       Next I
    End If
    objCADQTDTURNOS.OPERADORES = arrOPERADORES
    
    '' ======================================================================
    arrPARADAS = Empty
    With grdParadas
        If (.Rows - 1) > 0 Then
           ReDim arrPARADAS(1 To (.Rows - 1), 1 To 6) As String
           For I = 1 To (.Rows - 1)
                arrPARADAS(I, 1) = .Cell(flexcpText, I, conCOL_SonPa_Parada)
                arrPARADAS(I, 2) = .Cell(flexcpText, I, conCOL_SonPa_HorIni)
                arrPARADAS(I, 3) = .Cell(flexcpText, I, conCOL_SonPa_HorFin)
                arrPARADAS(I, 4) = .Cell(flexcpText, I, conCOL_SonPa_Total)
                arrPARADAS(I, 5) = IIf(Trim(.Cell(flexcpTextDisplay, I, conCOL_SonPa_ComParada)) = "Sim", 1, 0)
                arrPARADAS(I, 6) = .Cell(flexcpText, I, conCOL_SonPa_Pai)
           Next I
        End If
    End With
    objCADQTDTURNOS.PARADAS = arrPARADAS
    
    
    If objCADQTDTURNOS.GRAVA(cTipOper) = False Then Exit Sub
    If objCADQTDTURNOS.Atualiza(cTipOper, Str(objCADQTDTURNOS.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
          
    MsgBox "A Qtde de turnos " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADQTDTURNOS = Nothing
       Unload Me
    End If
    
End Sub

Private Sub cmdVoltar_Click()

On Error GoTo teste

    Set objBLBFunc = Nothing
    Set objCADQTDTURNOS = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
    
    Exit Sub

teste:
    MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
           "Descrição   : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"
           
End Sub

Private Sub Command1_Click()
    
    Dim strTOTALPERIODO As String
    Dim dtTotalLiquido  As Date
    
    If cTipOper = "I" Or cTipOper = "A" Then
       Call objBLBFunc.ExclLinhaGrid(grdParadas, grdParadas.Row)
       grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_QtdPar) = QtdParadas(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo))
       grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorPar) = TotalHoraParada(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo))
       
       strTOTALPERIODO = objBLBFunc.CalcTempo(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorEnt), grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorSai))
       If Len(Trim(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorPar))) > 0 Or _
          grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorPar) <> "00:00:00" Then
          dtTotalLiquido = CDate(strTOTALPERIODO & ":00") - CDate(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorPar) & ":00")
       Else
          dtTotalLiquido = CDate(strTOTALPERIODO & ":00")
       End If
       grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_TotalLiq) = Format(dtTotalLiquido, "HH:MM")
       
    End If
End Sub

Private Sub Command2_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridParadas
End Sub

Private Sub Command3_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridPeriodo
End Sub

Private Sub Command4_Click()
    Dim strPERIODO As String
    Dim I          As Integer
    If cTipOper = "I" Or cTipOper = "A" Then
VOLTA:
       strPERIODO = grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo)
       For I = 1 To (grdParadas.Rows - 1)
           If Trim(grdParadas.Cell(flexcpText, I, conCOL_SonPa_Pai)) = Trim(strPERIODO) Then
              grdParadas.RemoveItem I
              GoTo VOLTA
           End If
       Next I
       Call objBLBFunc.ExclLinhaGrid(grdPeriodo, grdPeriodo.Row)
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADQTDTURNOS = CreateObject("CADQTDTURNOS.clsCADQTDTURNOS")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
      
   objCADQTDTURNOS.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de quantidade de turnos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADQTDTURNOS
    
    txtCodigo.Text = ""
    
    Call InitGridMaquinas
    
    Call InitGridOperadores
    Call InitGridParadas
    Call InitGridPeriodo
    
    optAtivoSN(1).Value = True
   
End Sub

Private Sub grdOperadores_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case conCOL_SonOper_CodOper
         If (grdOperadores.Cols - 1) <> grdOperadores.Col Then grdOperadores.Col = Col + 2
    End Select
    Exit Sub
End Sub

Private Sub grdOperadores_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonOper_Desc_Oper
         Cancel = True
    Case conCOL_SonOper_CodOper, _
         conCOL_SonOper_PesqOper
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdOperadores.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdOperadores_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdOperadores.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonOper_PesqOper
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADOPERADOR " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRI"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Nome"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "SGI_DESCRI"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Operadores")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdOperadores.Cell(flexcpText, Row, conCOL_SonOper_CodOper) = varRETORNO
               grdOperadores.Cell(flexcpText, Row, conCOL_SonOper_Desc_Oper) = PegaDescrOperador(CLng(grdOperadores.Cell(flexcpText, Row, conCOL_SonOper_CodOper)))
            End If
            
            If VerifItensRepetidos(Row, conCOL_SonOper_CodOper, varRETORNO) = False Then
               MsgBox "Este operador já está relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdOperadores.Cell(flexcpText, Row, conCOL_SonOper_CodOper) = Empty
               grdOperadores.Cell(flexcpText, Row, conCOL_SonOper_Desc_Oper) = Empty
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdOperadores_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdOperadores
          Select Case Col
                    Case conCOL_SonOper_CodOper
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdOperadores_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdOperadores
          Select Case Col
                 Case conCOL_SonOper_CodOper
                        If .EditText = Empty Then Exit Sub
                        If VerifItensRepetidos(Row, conCOL_SonOper_CodOper, .EditText) = False Then
                           MsgBox "Este Operador ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdOperadores.Cell(flexcpText, Row, conCOL_SonOper_CodOper) = Empty
                           grdOperadores.Cell(flexcpText, Row, conCOL_SonOper_Desc_Oper) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrOperador(CLng(.EditText)))) = 0 Then
                           MsgBox "Este Operador não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonOper_Desc_Oper) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonOper_Desc_Oper) = PegaDescrOperador(CLng(.EditText))
          End Select
     End With

End Sub

Private Sub grdParadas_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Dim dtTotalLiquido  As Date
    
    Dim lngHoraInicial      As Long
    Dim lngHoraFinal        As Long
    Dim lngTotHoraParada    As Long
    
    Dim strTOTALPERIODO     As String

    With grdParadas
        Select Case Col
            Case conCOL_SonPa_HorIni, _
                 conCOL_SonPa_HorFin, _
                 conCOL_SonPa_ComParada
                 
                 If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonPa_HorIni), ":", ""))) > 0 And _
                    Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonPa_HorFin), ":", ""))) > 0 Then
                    
                    If CDate(.Cell(flexcpText, Row, conCOL_SonPa_HorIni)) > CDate(.Cell(flexcpText, Row, conCOL_SonPa_HorFin)) Then
                       MsgBox "Data Inicial não pode ser maior que data final !!!", vbOKOnly + vbExclamation, "Aviso"
                       .Cell(flexcpText, Row, conCOL_SonPa_HorIni) = ""
                       .Cell(flexcpText, Row, conCOL_SonPa_HorFin) = ""
                       .Cell(flexcpText, Row, conCOL_SonPa_Total) = ""
                       Exit Sub
                    End If
                    
                    .Cell(flexcpText, Row, conCOL_SonPa_Total) = Format(CDate(objBLBFunc.CalcTempo(Trim(.Cell(flexcpText, Row, conCOL_SonPa_HorIni)), Trim(.Cell(flexcpText, Row, conCOL_SonPa_HorFin)))), "HH:MM")
                    grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_QtdPar) = QtdParadas(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo))
                    grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorPar) = Format(CDate(TotalHoraParada(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo))), "HH:MM")
                    
                    strTOTALPERIODO = objBLBFunc.CalcTempo(Trim(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorEnt)), Trim(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorSai)))
                    
                    lngHoraInicial = objBLBFunc.CONVHRMIN(strTOTALPERIODO)
                    lngHoraFinal = objBLBFunc.CONVHRMIN(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_HorPar))
                    lngTotHoraParada = (lngHoraInicial - lngHoraFinal)
                    
                    strTOTALPERIODO = objBLBFunc.CONVMINHR(lngTotHoraParada)
                    dtTotalLiquido = CDate(strTOTALPERIODO)
                    grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_TotalLiq) = Format(dtTotalLiquido, "HH:MM")
                    
                 End If
        End Select
    End With
End Sub

Private Sub grdParadas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonPa_Total
         Cancel = True
    Case conCOL_SonPa_Parada, conCOL_SonPa_HorIni, conCOL_SonPa_HorFin, conCOL_SonPa_ComParada
         If cTipOper = "C" Then Cancel = True
    End Select
    Exit Sub
End Sub

Private Sub grdParadas_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdMAQUINAS
          Select Case Col
                    Case conCOL_SonPa_HorIni, conCOL_SonPa_HorFin
                         ''KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          End Select
     End With
End Sub

Private Sub grdParadas_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim intHoras As Integer
    Dim intMinutos As Integer
    
    With grdParadas
        Select Case Col
            Case conCOL_SonPa_HorIni, _
                 conCOL_SonPa_HorFin
                 If .EditText = "  :  " Then
                    Cancel = False
                    Exit Sub
                 End If
                 
                 '' ==================================
                 '' Validando Campo Horas
                 If Len(Trim(Mid(.EditText, 1, 2))) = 0 And Len(Trim(Mid(.EditText, 4, 2))) > 0 Then
                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 If Len(Trim(Mid(.EditText, 1, 2))) > 0 And Len(Trim(Mid(.EditText, 4, 2))) = 0 Then
                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 
                 
                 '' ==================================
                 '' Validando Horas
                 intHoras = CInt(Mid(.EditText, 1, 2))
                 intMinutos = CInt(Mid(.EditText, 4, 2))
                 If intHoras >= 24 Or intHoras < 0 Then
                    MsgBox "Hora Inválida o Dia vai somente até 24:00 !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 If intMinutos >= 60 Then
                    MsgBox "Minutos Inválido os minutos somente devem ser informados de 00 a 59 !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
        End Select
    End With

End Sub

Private Sub grdPeriodo_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Dim strTOTALPERIODO As String
    Dim dtTotalLiquido  As Date
    Dim lngMinutos      As Long
    
    With grdPeriodo
        Select Case Col
            Case conCOL_SonPe_HorEnt, _
                 conCOL_SonPe_HorSai
                 
                 If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonPe_HorEnt), ":", ""))) > 0 And _
                    Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonPe_HorSai), ":", ""))) > 0 Then
                    
                    strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, .Row, conCOL_SonPe_HorEnt), .Cell(flexcpText, .Row, conCOL_SonPe_HorSai))
                    
                    dtTotalLiquido = CDate(strTOTALPERIODO)
                    .Cell(flexcpText, .Row, conCOL_SonPe_TotalLiq) = Format(dtTotalLiquido, "HH:MM")
                    
                 End If
        End Select
    End With

End Sub

Private Sub grdPeriodo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonPe_TotalLiq, conCOL_SonPe_QtdPar, conCOL_SonPe_HorPar
         Cancel = True
    Case conCOL_SonPe_Periodo, conCOL_SonPe_HorEnt, conCOL_SonPe_HorSai
         If cTipOper = "C" Then Cancel = True
    End Select
    Exit Sub
End Sub

Private Sub grdPeriodo_Click()
    If (grdPeriodo.Rows - 1) > 0 And grdPeriodo.Row > 0 Then Call PosRegGrdParadas(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo))
End Sub

Private Sub grdPeriodo_RowColChange()
    If (grdPeriodo.Rows - 1) > 0 And grdPeriodo.Row > 0 Then Call PosRegGrdParadas(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo))
End Sub

Private Sub grdPeriodo_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim intHoras As Integer
    Dim intMinutos As Integer
    
    With grdPeriodo
        Select Case Col
            Case conCOL_SonPe_HorEnt, _
                 conCOL_SonPe_HorSai
                 If .EditText = "  :  " Then
                    Cancel = False
                    Exit Sub
                 End If
                 
                 '' ==================================
                 '' Validando Campo Horas
                 If Len(Trim(Mid(.EditText, 1, 2))) = 0 And Len(Trim(Mid(.EditText, 4, 2))) > 0 Then
                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 If Len(Trim(Mid(.EditText, 1, 2))) > 0 And Len(Trim(Mid(.EditText, 4, 2))) = 0 Then
                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 
                 
                 '' ==================================
                 '' Validando Horas
                 intHoras = CInt(Mid(.EditText, 1, 2))
                 intMinutos = CInt(Mid(.EditText, 4, 2))
                 If intHoras >= 24 Or intHoras < 0 Then
                    MsgBox "Hora Inválida o Dia vai somente até 24:00 !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 If intMinutos >= 60 Then
                    MsgBox "Minutos Inválido os minutos somente devem ser informados de 00 a 59 !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
        End Select
    End With

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADQTDTURNOS
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     
     Dim I As Long
     
     If (grdPeriodo.Rows - 1) = 0 Then
        MsgBox "Informe pelo menos 1 dia da semana !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     If Len(Trim(txtDescricao.Text)) = 0 Then
        MsgBox "Informe a descrição !!!", vbOKOnly + vbExclamation, "aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     With grdPeriodo
        For I = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, I, conCOL_SonPe_QtdPar))) = 0 Then
                MsgBox " Informe pelo menos 1 Parada !!!", vbOKOnly + vbExclamation, " Aviso"
                Exit Function
            ElseIf Len(Trim(.Cell(flexcpText, I, conCOL_SonPe_QtdPar))) > 0 Then
                If CLng(.Cell(flexcpText, I, conCOL_SonPe_QtdPar)) = 0 Then
                    MsgBox " Informe pelo menos 1 Parada !!!", vbOKOnly + vbExclamation, " Aviso"
                    Exit Function
                End If
            End If
        Next I
     End With
     
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    Dim j As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    
    Me.Caption = "Cadastro de quantidade de turnos - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADQTDTURNOS
    objCADQTDTURNOS.CODIGO = iCodigo
    
    Call InitGridOperadores
    Call InitGridMaquinas
    Call InitGridParadas
    Call InitGridPeriodo
    
    optAtivoSN(1).Value = True
    
    If objCADQTDTURNOS.Carrega_campos = True Then
       
       txtCodigo.Text = Str(objCADQTDTURNOS.CODIGO)
       txtDescricao.Text = objCADQTDTURNOS.DESCRI
       arrDIASSEMANA = objCADQTDTURNOS.DIASSEMANA
       arrPARADAS = objCADQTDTURNOS.PARADAS
       optAtivoSN(objCADQTDTURNOS.ATIVO).Value = True
       
       Call CarregaGrid
       Call PopGrdMaquinas
       Call PopGrdPeriodo
       Call PopGrdParadas
       
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    Dim j As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de quantidade de turnos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADQTDTURNOS
    
    objCADQTDTURNOS.CODIGO = iCodigo
    
    Call InitGridOperadores
    Call InitGridMaquinas
    Call InitGridParadas
    Call InitGridPeriodo
    
    optAtivoSN(1).Value = True
    
    If objCADQTDTURNOS.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADQTDTURNOS.CODIGO)
       txtDescricao.Text = objCADQTDTURNOS.DESCRI
       arrDIASSEMANA = objCADQTDTURNOS.DIASSEMANA
       arrPARADAS = objCADQTDTURNOS.PARADAS
       optAtivoSN(objCADQTDTURNOS.ATIVO).Value = True
       
       Call CarregaGrid
       Call PopGrdMaquinas
       Call PopGrdPeriodo
       Call PopGrdParadas
       
    End If
    
End Sub

Private Sub InitGridOperadores()

    With grdOperadores
    
       .Cols = conColumnsIn_SonOper
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonOper_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonOper_CodOper) = ""
       .ColDataType(conCOL_SonOper_CodOper) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOper_PesqOper) = ""
       .ColDataType(conCOL_SonOper_PesqOper) = flexDTString
       .ColComboList(conCOL_SonOper_PesqOper) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonOper_Desc_Oper) = ""
       .ColDataType(conCOL_SonOper_Desc_Oper) = flexDTString
       
       .ColWidth(conCOL_SonOper_CodOper) = 1500
       .ColWidth(conCOL_SonOper_PesqOper) = 300
       .ColWidth(conCOL_SonOper_Desc_Oper) = 4000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncRegGrid()
   
    If ExisteLinhaVazia = False Then Exit Sub
    
    grdOperadores.AddItem "" & vbTab & _
                          "" & vbTab & _
                          "" & vbTab & _
                          "" & vbTab & _
                          "" & vbTab & _
                          "" & vbTab & _
                          ""
                            
End Sub


Private Function ExisteLinhaVazia() As Boolean
    ExisteLinhaVazia = False
    
    Dim I As Integer
    
    For I = 1 To (grdOperadores.Rows - 1)
        If grdOperadores.Cell(flexcpText, I, conCOL_SonOper_CodOper) = Empty Then Exit Function
    Next I
    
    ExisteLinhaVazia = True
End Function

Private Function PegaDescrOperador(lngCodUsuario As Long) As String
    PegaDescrOperador = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADOPERADOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuario
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescrOperador = BREC!SGI_DESCRI
    BREC.Close
    
End Function

Private Function VerifItensRepetidos(intRow As Long, intCol As Long, varCampo As Variant) As Boolean
    VerifItensRepetidos = False
    Dim I As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    
    For I = 1 To (grdOperadores.Rows - 1)
        If I <> intRow And grdOperadores.Cell(flexcpText, I, intCol) = varCampo Then Exit Function
    Next I
    VerifItensRepetidos = True
End Function

Private Sub CarregaGrid()

    Dim I As Integer
    arrOPERADORES = objCADQTDTURNOS.OPERADORES
    
    If IsArray(arrOPERADORES) Then
       For I = 1 To UBound(arrOPERADORES)
           grdOperadores.AddItem arrOPERADORES(I) & vbTab & _
                                 "" & vbTab & _
                                 PegaDescrOperador(CLng(arrOPERADORES(I)))
       Next I
    End If
    
End Sub


Private Sub InitGridMaquinas()

    With grdMAQUINAS
    
       .Cols = conColumnsIn_SonMaq
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMaq_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMaq_CodMaq) = ""
       .ColDataType(conCOL_SonMaq_CodMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMaq_Desc_Maq) = ""
       .ColDataType(conCOL_SonMaq_Desc_Maq) = flexDTString
       
       .ColWidth(conCOL_SonMaq_CodMaq) = 1500
       .ColWidth(conCOL_SonMaq_Desc_Maq) = 4000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PopGrdMaquinas()

    sSql = "Select " & vbTab
    sSql = sSql & "       HEAD.* " & vbTab
    sSql = sSql & "  From " & vbTab
    sSql = sSql & "       SGI_CADMAQTURN ITEN" & vbTab
    sSql = sSql & "      ,SGI_CADMAQUINA HEAD" & vbTab
    sSql = sSql & " Where " & vbTab
    sSql = sSql & "       ITEN.SGI_FILIAL  = " & FILIAL & vbTab
    sSql = sSql & "   And ITEN.SGI_CODTURN = " & objCADQTDTURNOS.CODIGO & vbTab
    sSql = sSql & "   And HEAD.SGI_FILIAL = ITEN.SGI_FILIAL " & vbTab
    sSql = sSql & "   And HEAD.SGI_CODIGO = ITEN.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       grdMAQUINAS.AddItem BREC!SGI_CODIGO & vbTab & _
                           BREC!SGI_DESCRI
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Private Sub InitGridParadas()

    With grdParadas
    
       .Cols = conColumnsIn_SonPa
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonPa_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonPa_Parada) = ""
       .ColDataType(conCOL_SonPa_Parada) = flexDTString
       .ColComboList(conCOL_SonPa_Parada) = "|#1;Parada para almoço|#2;Parada para jantar|#3;Parada para café|#4;Para para descanso|#5;Parada para treinamento"
       
       .Cell(flexcpData, 0, conCOL_SonPa_HorIni) = ""
       .ColDataType(conCOL_SonPa_HorIni) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPa_HorFin) = ""
       .ColDataType(conCOL_SonPa_HorFin) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPa_Total) = ""
       .ColDataType(conCOL_SonPa_Total) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPa_ComParada) = ""
       .ColDataType(conCOL_SonPa_ComParada) = flexDTBoolean
       .ColFormat(conCOL_SonPa_ComParada) = "Sim;Não"
       
       .Cell(flexcpData, 0, conCOL_SonPa_Pai) = ""
       .ColDataType(conCOL_SonPa_Pai) = flexDTString
       
       .ColWidth(conCOL_SonPa_Parada) = 2000
       .ColWidth(conCOL_SonPa_HorIni) = 1000
       .ColWidth(conCOL_SonPa_HorFin) = 1000
       .ColWidth(conCOL_SonPa_Total) = 1000
       .ColHidden(conCOL_SonPa_Pai) = True
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncRegGridParadas()
   
    If grdPeriodo.Row = 0 Then
       MsgBox "Primeiro Selecione o Periodo !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    If ExisteLinhaVaziaParada = False Then Exit Sub
    
    With grdParadas
         .AddItem "" & vbTab & _
                  "" & vbTab & _
                  "" & vbTab & _
                  "" & vbTab & _
                  "" & vbTab & _
                  "" & Trim(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo))
                       
        '' Formatado para Campo Hora
        .ColEditMask(conCOL_SonPa_HorIni) = "##:##"
        .ColEditMask(conCOL_SonPa_HorFin) = "##:##"
        .ColEditMask(conCOL_SonPa_Total) = "##:##"
                      
        '' Alinhamento dos Campos
        .ColAlignment(conCOL_SonPa_HorIni) = flexAlignRightCenter
        .ColAlignment(conCOL_SonPa_HorFin) = flexAlignRightCenter
        .ColAlignment(conCOL_SonPa_Total) = flexAlignRightCenter
    End With
                       
End Sub


Private Function ExisteLinhaVaziaParada() As Boolean
    ExisteLinhaVaziaParada = False
    
    Dim I As Integer
    
    For I = 1 To (grdParadas.Rows - 1)
        If grdParadas.Cell(flexcpText, I, conCOL_SonPa_Parada) = Empty Then Exit Function
    Next I
    
    ExisteLinhaVaziaParada = True
End Function

''Private Function CalcTempo(strHORINI As String, strHORFIN As String) As String
''
''    CalcTempo = ""
'
'    Dim dtTotalHora   As Date
'
'    dtTotalHora = (CDate(Trim(strHORINI) & ":00") - CDate(Trim(strHORFIN) & ":00"))
'
'    CalcTempo = Format(dtTotalHora, "HH:MM")
'
'End Function

Private Sub InitGridPeriodo()

    With grdPeriodo
    
       .Cols = conColumnsIn_SonPe
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonPe_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonPe_Periodo) = ""
       .ColDataType(conCOL_SonPe_Periodo) = flexDTString
       .ColComboList(conCOL_SonPe_Periodo) = ComboDiasSemanaGrd
       
       .Cell(flexcpData, 0, conCOL_SonPe_HorEnt) = ""
       .ColDataType(conCOL_SonPe_HorEnt) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPe_HorSai) = ""
       .ColDataType(conCOL_SonPe_HorSai) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPe_HorPar) = ""
       .ColDataType(conCOL_SonPe_HorPar) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPe_QtdPar) = ""
       .ColDataType(conCOL_SonPe_QtdPar) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonPe_TotalLiq) = ""
       .ColDataType(conCOL_SonPe_TotalLiq) = flexDTString
       
       .ColWidth(conCOL_SonPe_Periodo) = 1800
       .ColWidth(conCOL_SonPe_HorEnt) = 1000
       .ColWidth(conCOL_SonPe_HorSai) = 1000
       .ColWidth(conCOL_SonPe_HorPar) = 1000
       .ColWidth(conCOL_SonPe_QtdPar) = 1000
       .ColWidth(conCOL_SonPe_TotalLiq) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Function ComboDiasSemanaGrd() As String
    ComboDiasSemanaGrd = "|#1;Domingo|#2;Segunda|#3;Terça|#4;Quarta|#5;Quinta|#6;Sexta|#7;Sabado"
End Function


Private Sub IncRegGridPeriodo()
   
    If ExisteLinhaVaziaPeriodo = False Then Exit Sub
    
    With grdPeriodo
         .AddItem "" & vbTab & _
                  "" & vbTab & _
                  "" & vbTab & _
                  "" & vbTab & _
                  "" & vbTab & _
                  ""
                       
        '' Formatado para Campo Hora
        .ColEditMask(conCOL_SonPe_HorEnt) = "##:##"
        .ColEditMask(conCOL_SonPe_HorSai) = "##:##"
        .ColEditMask(conCOL_SonPe_HorPar) = "##:##"
        .ColEditMask(conCOL_SonPe_TotalLiq) = "##:##"
                      
        '' Alinhamento dos Campos
        .ColAlignment(conCOL_SonPe_HorEnt) = flexAlignRightCenter
        .ColAlignment(conCOL_SonPe_HorSai) = flexAlignRightCenter
        .ColAlignment(conCOL_SonPe_HorPar) = flexAlignRightCenter
        .ColAlignment(conCOL_SonPe_TotalLiq) = flexAlignRightCenter
    End With
                       
End Sub


Private Function ExisteLinhaVaziaPeriodo() As Boolean
    ExisteLinhaVaziaPeriodo = False
    
    Dim I As Integer
    
    For I = 1 To (grdPeriodo.Rows - 1)
        If grdPeriodo.Cell(flexcpText, I, conCOL_SonPe_Periodo) = Empty Then Exit Function
    Next I
    
    ExisteLinhaVaziaPeriodo = True
End Function


Private Function QtdParadas(strPAI As String) As String

    QtdParadas = ""
    
    Dim I As Integer
    Dim lngTOTPARADAS As Long
    
    lngTOTPARADAS = 0
    For I = 1 To (grdParadas.Rows - 1)
        If Trim(grdParadas.Cell(flexcpText, I, conCOL_SonPa_Pai)) = Trim(strPAI) Then lngTOTPARADAS = (lngTOTPARADAS + 1)
    Next I
    
    QtdParadas = Format(lngTOTPARADAS, "##00")
End Function

Private Function TotalHoraParada(strPAI As String) As String
    
    TotalHoraParada = ""
    
    Dim dtTotHora   As Date
    Dim I           As Integer
    Dim lngMinutos  As Long
    
    With grdParadas
        lngMinutos = 0
        For I = 1 To (.Rows - 1)
            If Trim(.Cell(flexcpText, I, conCOL_SonPa_Pai)) = Trim(strPAI) Then
               If .Cell(flexcpTextDisplay, I, conCOL_SonPa_ComParada) = "Sim" Then
                  lngMinutos = lngMinutos + objBLBFunc.CONVHRMIN(.Cell(flexcpText, I, conCOL_SonPa_Total))
               End If
            End If
        Next I
    End With
    
    TotalHoraParada = objBLBFunc.CONVMINHR(lngMinutos)
    
    
    If Len(Trim(TotalHoraParada)) = 0 Then TotalHoraParada = "00:00:00"
    
End Function

Private Sub PosRegGrdParadas(strCODPROD As String)
    Dim I As Integer
    With grdParadas
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonPa_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With
End Sub


Private Sub PopGrdPeriodo()
    
    Dim I          As Integer
    Dim strPERIODO As String
    Dim strPARADAS As String
    
    Dim lngHORAINI As Long
    Dim lngHORAFIN As Long
    
    Dim dtTotalLiq As Date
    
    If IsArray(arrDIASSEMANA) = True Then
       With grdPeriodo
            For I = 1 To UBound(arrDIASSEMANA)
                .AddItem Trim(arrDIASSEMANA(I, 1)) & vbTab & _
                              arrDIASSEMANA(I, 2) & vbTab & _
                              arrDIASSEMANA(I, 3) & vbTab & _
                              arrDIASSEMANA(I, 4) & vbTab & _
                              arrDIASSEMANA(I, 5) & vbTab & _
                              "00:00" & vbTab & _
                               ""
                                       
                '' Formatado para Campo Hora
                .ColEditMask(conCOL_SonPe_HorEnt) = "##:##"
                .ColEditMask(conCOL_SonPe_HorSai) = "##:##"
                .ColEditMask(conCOL_SonPe_HorPar) = "##:##"
                .ColEditMask(conCOL_SonPe_TotalLiq) = "##:##"
                           
                '' Alinhamento dos Campos
                .ColAlignment(conCOL_SonPe_HorEnt) = flexAlignRightCenter
                .ColAlignment(conCOL_SonPe_HorSai) = flexAlignRightCenter
                .ColAlignment(conCOL_SonPe_HorPar) = flexAlignRightCenter
                .ColAlignment(conCOL_SonPe_TotalLiq) = flexAlignRightCenter
                
                strPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, (.Rows - 1), conCOL_SonPe_HorEnt), .Cell(flexcpText, (.Rows - 1), conCOL_SonPe_HorSai))
                strPARADAS = .Cell(flexcpText, (.Rows - 1), conCOL_SonPe_HorPar)
                
                lngHORAINI = objBLBFunc.CONVHRMIN(strPERIODO)
                lngHORAFIN = objBLBFunc.CONVHRMIN(strPARADAS)
                
                dtTotalLiq = CDate(objBLBFunc.CONVMINHR(lngHORAINI - lngHORAFIN))
                .Cell(flexcpText, .Rows - 1, conCOL_SonPe_TotalLiq) = Format(dtTotalLiq, "HH:MM")
            Next I
       End With
    End If
    
End Sub

Private Sub PopGrdParadas()

    Dim I As Integer
    
    If IsArray(arrPARADAS) Then
       With grdParadas
            For I = 1 To UBound(arrPARADAS)
                
                .AddItem Trim(arrPARADAS(I, 1)) & vbTab & _
                                   arrPARADAS(I, 2) & vbTab & _
                                   arrPARADAS(I, 3) & vbTab & _
                                   arrPARADAS(I, 4) & vbTab & _
                                   arrPARADAS(I, 5) & vbTab & _
                                   arrPARADAS(I, 6)
                                   
                '' Formatado para Campo Hora
                .ColEditMask(conCOL_SonPa_HorIni) = "##:##"
                .ColEditMask(conCOL_SonPa_HorFin) = "##:##"
                .ColEditMask(conCOL_SonPa_Total) = "##:##"
                               
                '' Alinhamento dos Campos
                .ColAlignment(conCOL_SonPa_HorIni) = flexAlignRightCenter
                .ColAlignment(conCOL_SonPa_HorFin) = flexAlignRightCenter
                .ColAlignment(conCOL_SonPa_Total) = flexAlignRightCenter
                                   
            Next I
            
            If (grdPeriodo.Rows - 1) > 0 Then
               grdPeriodo.Row = 1
               Call PosRegGrdParadas(grdPeriodo.Cell(flexcpText, grdPeriodo.Row, conCOL_SonPe_Periodo))
            End If
            
       End With
    End If
End Sub
