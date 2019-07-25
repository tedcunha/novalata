VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADGRUPLINHA 
   Caption         =   "Cadastro de Grupos de Linha"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   14535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame12 
      Caption         =   "[ Capacidade Produtiva ]"
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
      Height          =   3735
      Left            =   0
      TabIndex        =   23
      Top             =   4800
      Width           =   14535
      Begin VSFlex8LCtl.VSFlexGrid grdCAPACPROD 
         Height          =   3375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   13935
         _cx             =   24580
         _cy             =   5953
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
   Begin VB.Frame Frame11 
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
      Height          =   2535
      Left            =   9480
      TabIndex        =   19
      Top             =   2280
      Width           =   5055
      Begin VB.CommandButton Command9 
         Height          =   300
         Left            =   4680
         Picture         =   "frmCADGRUPLINHA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton Command10 
         Height          =   300
         Left            =   4680
         Picture         =   "frmCADGRUPLINHA.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exclui a linha da Gride Selecionada"
         Top             =   600
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdMesAno 
         Height          =   2175
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4455
         _cx             =   7858
         _cy             =   3836
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
      Caption         =   "[ Capacidade ]"
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
      Height          =   2535
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   9375
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   9000
         Picture         =   "frmCADGRUPLINHA.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   9000
         Picture         =   "frmCADGRUPLINHA.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdCAPACIDADE 
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8775
         _cx             =   15478
         _cy             =   3836
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
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   14535
      Begin VB.TextBox txtQtdePorHora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Text            =   "txtQtdePorHora"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtQTDECAPACGERAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "txtQTDECAPACGERAL"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8040
         TabIndex        =   12
         Top             =   600
         Width           =   1695
         Begin VB.OptionButton optATIVOSN 
            Caption         =   "SIM"
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
            TabIndex        =   14
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optATIVOSN 
            Caption         =   "NÃO"
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
            TabIndex        =   13
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qtde/Hora"
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
         TabIndex        =   25
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capacidade Total"
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
         Left            =   9960
         TabIndex        =   17
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ativo"
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
         Left            =   7440
         TabIndex        =   11
         Top             =   600
         Width           =   450
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
         Left            =   120
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
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
         Picture         =   "frmCADGRUPLINHA.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADGRUPLINHA.frx":062A
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADGRUPLINHA.frx":072C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADGRUPLINHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public strMODPAI        As String
Public FILIAL           As Integer
Public strAcesso        As String
Public strUsuario       As String
Public lngCodUsuario    As Long
Public strNOMETABELA    As String
Public strNOMEFILIAL    As String

Dim objBLBFunc          As Object
Dim objCADCADGRUPLINHA  As Object
Dim objPESQPADRAO       As Object
Dim strCAPTION          As String
Dim arrCAPACIDADEIT     As Variant
Dim arrMESANO           As Variant
Dim arrCAPACPROD        As Variant

Const conCOL_GRPLINHA_CodCor                        As Integer = 0
Const conCOL_GRPLINHA_PesqCor                       As Integer = 1
Const conCOL_GRPLINHA_DescCor                       As Integer = 2
Const conCOL_GRPLINHA_IDCAPC                        As Integer = 3
Const conCOL_GRPLINHA_IDINTERNO                     As Integer = 4
Const conCOL_GRPLINHA_NECKINSN                      As Integer = 5
Const conCOL_GRPLINHA_INDICE                        As Integer = 6
Const conCOL_GRPLINHA_HOMOLOGSN                     As Integer = 7
Const conCOL_GRPLINHA_FormatString                  As String = "=Cód.Capc|...|Descrição da Capacidade|IDCAPC|IDInterno|Neck-IN(S/N)|INDICE|Homologada(S/N)"
Const conColumnsIn_GRPLINHA                         As Integer = 8

Const conCOL_SonMesAno_Mes                          As Integer = 0
Const conCOL_SonMesAno_Ano                          As Integer = 1
Const conCOL_SonMesAno_Indice                       As Integer = 2
Const conCOL_SonMesAno_QtdeCapac                    As Integer = 3
Const conCOL_SonMesAno_Carrega                      As Integer = 4
Const conCOL_SonMesAno_IndiceBKP                    As Integer = 5
Const conCOL_SonMesAno_MesBKP                       As Integer = 6
Const conCOL_SonMesAno_AnoBKP                       As Integer = 7
Const conCOL_SonMesAno_IDPAI                        As Integer = 8
Const conCOL_SonMesAno_FormatString                 As String = "=Mês|Ano|Indice|Qtde Capacidade|...|IndiceBKP|MesBKP|AnoBKP|IDPAI"
Const conColumnsIn_SonMesAno                        As Integer = 9

Const conCOL_SonCapacProd_Data                  As Integer = 0
Const conCOL_SonCapacProd_DiaDesc               As Integer = 1
Const conCOL_SonCapacProd_HorInicial            As Integer = 2
Const conCOL_SonCapacProd_Parada                As Integer = 3
Const conCOL_SonCapacProd_HoraFinal             As Integer = 4
Const conCOL_SonCapacProd_TotalHoras            As Integer = 5
Const conCOL_SonCapacProd_TotalPecas            As Integer = 6
Const conCOL_SonCapacProd_CodIndice             As Integer = 7
Const conCOL_SonCapacProd_DiaSem                As Integer = 8
Const conCOL_SonCapacProd_AtivoSN               As Integer = 9
Const conCOL_SonCapacProd_INDICE                As Integer = 10
Const conCOL_SonCapacProd_IDPAI                 As Integer = 11
Const conCOL_SonCapacProd_FormatString          As String = "=Data|Dia da Semana|Hora Inicial|Parada|Hora Final|Total de Horas|Total de Peças|CodIndice|DiaSem|Ativo|Indice|IDPAI"
Const conColumnsIn_SonCapacProd                 As Integer = 12

Private Sub cmdAltera_Click()
    
    cTipOper = "A"
    If objBLBFunc.ChecaAcesso2(cTipOper, strAcesso) = False Then Exit Sub
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub CmdSalva_Click()

    Dim I   As Long
    Dim j   As Long
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADCADGRUPLINHA.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMETABELA, FILIAL, Linha)
    
    objCADCADGRUPLINHA.DESCRI = "'" & Trim(Replace(Replace(txtDescricao.Text, ",", ""), "'", "")) & "'"
    
    If optATIVOSN(0).Value = True Then objCADCADGRUPLINHA.ATIVO = 0
    If optATIVOSN(1).Value = True Then objCADCADGRUPLINHA.ATIVO = 1
    
    objCADCADGRUPLINHA.CAPACTOTLINHA = "Null"
    If Len(Trim(txtQTDECAPACGERAL.Text)) > 0 Then objCADCADGRUPLINHA.CAPACTOTLINHA = txtQTDECAPACGERAL.Text
    
    objCADCADGRUPLINHA.QTDPORHORA = txtQtdePorHora.Text
    
    
    arrCAPACIDADEIT = Empty
    With grdCAPACIDADE
        If (.Rows - 1) > 0 Then
            ReDim arrCAPACIDADEIT(1 To (.Rows - 1), 1 To 5) As String
            For I = 1 To (.Rows - 1)
                arrCAPACIDADEIT(I, 1) = .Cell(flexcpText, I, conCOL_GRPLINHA_IDCAPC)
                arrCAPACIDADEIT(I, 2) = .Cell(flexcpText, I, conCOL_GRPLINHA_NECKINSN)
                arrCAPACIDADEIT(I, 3) = .Cell(flexcpText, I, conCOL_GRPLINHA_HOMOLOGSN)
                arrCAPACIDADEIT(I, 4) = .Cell(flexcpText, I, conCOL_GRPLINHA_INDICE)
                
                '' Gerando o ID Interno
                If cTipOper = "I" Then
                    arrCAPACIDADEIT(I, 5) = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMETABELA & "_CAPAC", FILIAL, Linha)
                ElseIf cTipOper = "A" Then
                    If Len(Trim(.Cell(flexcpText, I, conCOL_GRPLINHA_IDINTERNO))) = 0 Then
                        arrCAPACIDADEIT(I, 5) = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMETABELA & "_CAPAC", FILIAL, Linha)
                    Else
                        arrCAPACIDADEIT(I, 5) = .Cell(flexcpText, I, conCOL_GRPLINHA_IDINTERNO)
                    End If
                End If
                
            Next I
        End If
    End With
    objCADCADGRUPLINHA.CAPACIDADELIN = arrCAPACIDADEIT
    
    ''---------------------------------------
    '' Mes/Ano
    Call objBLBFunc.RemoveLinhaVazia(grdMesAno, conCOL_SonMesAno_Indice)
    arrMESANO = Empty
    With grdMesAno
        If (.Rows - 1) > 0 Then
            ReDim arrMESANO(1 To (.Rows - 1), 1 To 5) As String
            For I = 1 To (.Rows - 1)
                arrMESANO(I, 1) = .Cell(flexcpText, I, conCOL_SonMesAno_Mes)
                arrMESANO(I, 2) = .Cell(flexcpText, I, conCOL_SonMesAno_Ano)
                
                arrMESANO(I, 4) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonMesAno_QtdeCapac))) > 0 Then
                    arrMESANO(I, 4) = .Cell(flexcpText, I, conCOL_SonMesAno_QtdeCapac)
                End If
                
                If cTipOper = "I" Then
                    arrMESANO(I, 3) = .Cell(flexcpText, I, conCOL_SonMesAno_Indice) & arrMESANO(I, 5) & objCADCADGRUPLINHA.CODIGO
                    arrMESANO(I, 5) = objCADCADGRUPLINHA.CODIGO
                Else
                    arrMESANO(I, 3) = .Cell(flexcpText, I, conCOL_SonMesAno_Indice)
                    arrMESANO(I, 5) = .Cell(flexcpText, I, conCOL_SonMesAno_IDPAI)
                End If
                '' ============================
                
            Next I
        End If
    End With
    objCADCADGRUPLINHA.MESANO = arrMESANO
    ''---------------------------------------
    
    ''---------------------------------------
    '' Capacidade Produtiva
    arrCAPACPROD = Empty
    With grdCAPACPROD
        If (.Rows - 1) > 0 Then
            ReDim arrCAPACPROD(1 To (.Rows - 1), 1 To 11) As String
            For I = 1 To (.Rows - 1)
                arrCAPACPROD(I, 1) = "'" & Format(CDate(.Cell(flexcpText, I, conCOL_SonCapacProd_Data)), "MM/DD/YYYY") & "'"
                
                arrCAPACPROD(I, 2) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_HorInicial))) > 0 Then arrCAPACPROD(I, 2) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_HorInicial)) & "'"
                
                arrCAPACPROD(I, 3) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_Parada))) > 0 Then arrCAPACPROD(I, 3) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_Parada)) & "'"
                
                arrCAPACPROD(I, 4) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_HoraFinal))) > 0 Then arrCAPACPROD(I, 4) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_HoraFinal)) & "'"
            
                arrCAPACPROD(I, 5) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_TotalHoras))) > 0 Then arrCAPACPROD(I, 5) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_TotalHoras)) & "'"
            
                arrCAPACPROD(I, 6) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_TotalPecas))) > 0 Then arrCAPACPROD(I, 6) = .Cell(flexcpText, I, conCOL_SonCapacProd_TotalPecas)
            
                If cTipOper = "I" Then
                    arrCAPACPROD(I, 7) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_CodIndice)) & objCADCADGRUPLINHA.CODIGO & "'"
                Else
                    arrCAPACPROD(I, 7) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_CodIndice)) & "'"
                End If
                
                arrCAPACPROD(I, 8) = .Cell(flexcpText, I, conCOL_SonCapacProd_DiaSem)
                arrCAPACPROD(I, 9) = .Cell(flexcpText, I, conCOL_SonCapacProd_AtivoSN)
                
                If cTipOper = "I" Then
                    arrCAPACPROD(I, 10) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_INDICE)) & objCADCADGRUPLINHA.CODIGO & "'"
                    arrCAPACPROD(I, 11) = objCADCADGRUPLINHA.CODIGO
                Else
                    arrCAPACPROD(I, 10) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_INDICE)) & "'"
                    arrCAPACPROD(I, 11) = .Cell(flexcpText, I, conCOL_SonCapacProd_IDPAI)
                End If
            
            Next I
        End If
    End With
    objCADCADGRUPLINHA.CAPACPROD = arrCAPACPROD
    ''---------------------------------------
    
    If objCADCADGRUPLINHA.GRAVA(cTipOper, strNOMETABELA) = False Then Exit Sub
          
    MsgBox "O Grupo de Linha foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command10_Click()
    If cTipOper = "I" Or cTipOper = "A" Then
        With grdMesAno
            If (.Rows - 1) = 0 Or .Row = 0 Then Exit Sub
            Call objBLBFunc.ExcLinhaGrdFilho(grdCAPACPROD, conCOL_SonCapacProd_CodIndice, .Cell(flexcpText, .Row, conCOL_SonMesAno_Indice))
            Call objBLBFunc.ExclLinhaGrid(grdMesAno, .Row)
        End With
    End If
End Sub

Private Sub Command2_Click()
    Call IncRegGrid
End Sub

Private Sub Command3_Click()
    If cTipOper = "C" Then Exit Sub
    With grdCAPACIDADE
        If (.Rows - 1) = 0 Or .Row = 0 Then Exit Sub
        Call objBLBFunc.ExclLinhaGrid(grdCAPACIDADE, .Row)
    End With
End Sub

Private Sub Command9_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGridMesAno
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADCADGRUPLINHA = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADCADGRUPLINHA = CreateObject("CADGRUPLINHA.clsCADGRUPLINHA")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strCAPTION = "Cadastro de Grupos de Linha - " & strNOMEFILIAL & " - "
   
    objCADCADGRUPLINHA.FILIAL = FILIAL
   
    Call IniciaForm

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub grdCAPACIDADE_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdCAPACIDADE
        
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        
        Select Case Col
               Case conCOL_GRPLINHA_CodCor
        End Select
    
    End With
End Sub

Private Sub grdCAPACIDADE_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdCAPACIDADE
        Select Case Col
               Case conCOL_GRPLINHA_DescCor, _
                    conCOL_GRPLINHA_IDCAPC, _
                    conCOL_GRPLINHA_IDINTERNO, _
                    conCOL_GRPLINHA_INDICE
                    Cancel = True
               Case conCOL_GRPLINHA_CodCor, _
                    conCOL_GRPLINHA_PesqCor, _
                    conCOL_GRPLINHA_NECKINSN, _
                    conCOL_GRPLINHA_HOMOLOGSN
                    If cTipOper = "C" Then Cancel = True
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub
End Sub

Private Sub grdCAPACIDADE_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    With grdCAPACIDADE
        If (.Rows - 1) = 0 Then Exit Sub
        
        Select Case Col
            Case conCOL_GRPLINHA_PesqCor
                
                Call PesqCapc(Row)
                
                Exit Sub
                
        End Select
    End With

End Sub


Private Sub grdCAPACIDADE_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdCAPACIDADE
          Select Case Col
                    Case conCOL_GRPLINHA_CodCor
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdCAPACIDADE_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
     Dim strINDICE As String

     With grdCAPACIDADE
          Select Case Col
                 Case conCOL_GRPLINHA_CodCor
                        If .EditText = Empty Then Exit Sub
                        
                        .Cell(flexcpText, Row, conCOL_GRPLINHA_IDCAPC) = PegaIDLinha(.EditText)
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_GRPLINHA_IDCAPC))) = 0 Then
                           MsgBox "ATENÇÂO" & vbCrLf & "Capacidade não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        ''strINDICE = .Cell(flexcpText, Row, conCOL_GRPLINHA_IDCAPC) & .Cell(flexcpText, Row, conCOL_GRPLINHA_NECKINSN) & .Cell(flexcpText, Row, conCOL_GRPLINHA_HOMOLOGSN)
                        
                        strINDICE = .Cell(flexcpText, Row, conCOL_GRPLINHA_IDCAPC)
                        If objBLBFunc.FcVerifItensRepetidos(grdCAPACIDADE, Row, conCOL_GRPLINHA_INDICE, Trim(strINDICE)) = False Then
                           MsgBox "ATENÇÃO" & vbCrLf & "A Capacidade ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_CodCor) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_DescCor) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_IDINTERNO) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_INDICE) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_NECKINSN) = 0
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_HOMOLOGSN) = 0
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_GRPLINHA_CodCor) = .EditText
                        .Cell(flexcpText, Row, conCOL_GRPLINHA_DescCor) = PegaDescrLinha(.Cell(flexcpText, Row, conCOL_GRPLINHA_IDCAPC))
                        .Cell(flexcpText, Row, conCOL_GRPLINHA_INDICE) = strINDICE
                        
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_GRPLINHA_DescCor))) = 0 Then
                           MsgBox "ATENÇÃO" & vbCrLf & "Capacidade não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_GRPLINHA_DescCor) = Empty
                           .Cell(flexcpText, Row, conCOL_GRPLINHA_IDCAPC) = Empty
                           .Cell(flexcpText, Row, conCOL_GRPLINHA_INDICE) = Empty
                           .Cell(flexcpText, Row, conCOL_GRPLINHA_NECKINSN) = 0
                           .Cell(flexcpText, Row, conCOL_GRPLINHA_HOMOLOGSN) = 0
                           Cancel = True
                           Exit Sub
                        End If
                Case conCOL_GRPLINHA_NECKINSN
                        If .EditText = Empty Then Exit Sub
                
                        strINDICE = .Cell(flexcpText, Row, conCOL_GRPLINHA_IDCAPC) & IIf(.EditText = "Sim", 1, 0) & .Cell(flexcpText, Row, conCOL_GRPLINHA_HOMOLOGSN)
                        If objBLBFunc.FcVerifItensRepetidos(grdCAPACIDADE, Row, conCOL_GRPLINHA_INDICE, Trim(strINDICE)) = False Then
                           MsgBox "ATENÇÃO" & vbCrLf & "A Capacidade ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_CodCor) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_DescCor) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_IDINTERNO) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_INDICE) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_NECKINSN) = 0
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_HOMOLOGSN) = 0
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_GRPLINHA_INDICE) = Trim(strINDICE)
                
                Case conCOL_GRPLINHA_HOMOLOGSN
                        If .EditText = Empty Then Exit Sub
                
                        strINDICE = .Cell(flexcpText, Row, conCOL_GRPLINHA_IDCAPC) & .Cell(flexcpText, Row, conCOL_GRPLINHA_NECKINSN) & IIf(.EditText = "Sim", 1, 0)
                        If objBLBFunc.FcVerifItensRepetidos(grdCAPACIDADE, Row, conCOL_GRPLINHA_INDICE, Trim(strINDICE)) = False Then
                           MsgBox "ATENÇÃO" & vbCrLf & "A Capacidade ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_CodCor) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_DescCor) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_IDINTERNO) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_INDICE) = Empty
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_NECKINSN) = 0
                           grdCAPACIDADE.Cell(flexcpText, Row, conCOL_GRPLINHA_HOMOLOGSN) = 0
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_GRPLINHA_INDICE) = Trim(strINDICE)
                
          End Select
     End With

End Sub

Private Sub grdCAPACPROD_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Dim strTOTALPERIODO     As String
    Dim dtTotalLiquido      As Date
    Dim lngMinutos          As Long
    Dim lngHoraInicial      As Long
    Dim lngHoraFinal        As Long
    Dim lngTotHoraParada    As Long
    Dim arrHORAMIN          As Variant
    Dim lngTotPecasHora     As Long
    
    With grdCAPACPROD
        Select Case Col
            Case conCOL_SonCapacProd_AtivoSN
                If .Cell(flexcpText, .Row, conCOL_SonCapacProd_AtivoSN) = 0 Then
                    .Cell(flexcpText, Row, conCOL_SonCapacProd_HorInicial) = Empty
                    .Cell(flexcpText, Row, conCOL_SonCapacProd_Parada) = Empty
                    .Cell(flexcpText, Row, conCOL_SonCapacProd_HoraFinal) = Empty
                    .Cell(flexcpText, Row, conCOL_SonCapacProd_TotalHoras) = Empty
                    .Cell(flexcpText, Row, conCOL_SonCapacProd_TotalPecas) = Empty
                End If
            Case conCOL_SonCapacProd_HorInicial, _
                 conCOL_SonCapacProd_Parada, _
                 conCOL_SonCapacProd_HoraFinal
                 
                 If (grdMesAno.Rows - 1) = 0 Or grdMesAno.Row = 0 Then
                    MsgBox "ATENÇÂO - Selecione um Mes/Ano !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Sub
                 End If
                 
                 If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonCapacProd_HorInicial), ":", ""))) > 0 And _
                    Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonCapacProd_Parada), ":", ""))) > 0 And _
                    Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonCapacProd_HoraFinal), ":", ""))) > 0 Then
                    
                    '' Ativo (S/N)
                    .Cell(flexcpText, .Row, conCOL_SonCapacProd_AtivoSN) = 1
                    
                    strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, .Row, conCOL_SonCapacProd_HorInicial), .Cell(flexcpText, .Row, conCOL_SonCapacProd_HoraFinal))
                    
                    lngHoraInicial = objBLBFunc.CONVHRMIN(strTOTALPERIODO)
                    lngHoraFinal = objBLBFunc.CONVHRMIN(.Cell(flexcpText, .Row, conCOL_SonCapacProd_Parada))
                    lngTotHoraParada = (lngHoraInicial - lngHoraFinal)
                    
                    strTOTALPERIODO = objBLBFunc.CONVMINHR(lngTotHoraParada)
                    If Len(Trim(strTOTALPERIODO)) > 0 Then
                         dtTotalLiquido = CDate(strTOTALPERIODO)
                         .Cell(flexcpText, .Row, conCOL_SonCapacProd_TotalHoras) = Format(dtTotalLiquido, "HH:MM")
                         
                         
                        If Len(Trim(txtQtdePorHora.Text)) > 0 Then
                             '' -------------------------
                             '' Qtde de Pecas
                             .Cell(flexcpText, Row, conCOL_SonCapacProd_TotalPecas) = CalcTotPecasHora(.Cell(flexcpText, Row, conCOL_SonCapacProd_TotalHoras), CLng(txtQtdePorHora.Text))
                            '' -------------------------
                        End If
                        
                        '' Total de Peças no Mês
                        grdMesAno.Cell(flexcpText, grdMesAno.Row, conCOL_SonMesAno_QtdeCapac) = CalcTotCapacidade(.Cell(flexcpText, .Row, conCOL_SonCapacProd_CodIndice))
                        
                        txtQTDECAPACGERAL.Text = CalcTotalGerLinha
                        
                        If Row < (.Rows - 1) Then
                            .Row = (Row + 1)
                        End If
                        
                    End If
                 Else
                    If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonCapacProd_HorInicial), ":", ""))) = 0 And _
                       Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonCapacProd_Parada), ":", ""))) = 0 And _
                       Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonCapacProd_HoraFinal), ":", ""))) = 0 Then
                        '' Ativo (S/N)
                        .Cell(flexcpText, .Row, conCOL_SonCapacProd_AtivoSN) = 0
                        .Cell(flexcpText, Row, conCOL_SonCapacProd_HorInicial) = Empty
                        .Cell(flexcpText, Row, conCOL_SonCapacProd_Parada) = Empty
                        .Cell(flexcpText, Row, conCOL_SonCapacProd_HoraFinal) = Empty
                        .Cell(flexcpText, Row, conCOL_SonCapacProd_TotalHoras) = Empty
                        .Cell(flexcpText, Row, conCOL_SonCapacProd_TotalPecas) = Empty
                    End If
                 End If
        End Select
    End With

End Sub

Private Sub grdCAPACPROD_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


    With grdCAPACPROD
        Select Case Col
        Case conCOL_SonCapacProd_Data, _
             conCOL_SonCapacProd_DiaDesc, _
             conCOL_SonCapacProd_CodIndice, _
             conCOL_SonCapacProd_DiaSem, _
             conCOL_SonCapacProd_TotalHoras, _
             conCOL_SonCapacProd_TotalPecas, _
             conCOL_SonCapacProd_INDICE, _
             conCOL_SonCapacProd_IDPAI
             Cancel = True
        Case conCOL_SonCapacProd_HorInicial, _
             conCOL_SonCapacProd_Parada, _
             conCOL_SonCapacProd_HoraFinal
             If cTipOper = "C" Then
                Cancel = True
             Else
                ''If .Cell(flexcpText, Row, conCOL_SonCapacProd_AtivoSN) = 0 Then Cancel = True
             End If
        Case conCOL_SonCapacProd_AtivoSN
             If cTipOper = "C" Then Cancel = True
        Case Else
            .ComboList = ""
        End Select
    End With
    Exit Sub

End Sub


Private Sub grdCAPACPROD_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim intHoras    As Integer
    Dim intMinutos  As Integer
    Dim strHORA     As String
    
    With grdCAPACPROD
        Select Case Col
            Case conCOL_SonCapacProd_HorInicial, _
                 conCOL_SonCapacProd_Parada, _
                 conCOL_SonCapacProd_HoraFinal
                 If Len(Trim(.EditText)) = 0 Then
                    Cancel = False
                    Exit Sub
                 End If
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
'                 If .EditText = "00:00" Then
'                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
'                    Cancel = True
'                    Exit Sub
'                 End If
                 
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
                 
                 If Col = conCOL_SonCapacProd_HorInicial Then
                    Call PosColCapac(conCOL_SonCapacProd_Parada, Row)
                 ElseIf Col = conCOL_SonCapacProd_Parada Then
                    Call PosColCapac(conCOL_SonCapacProd_HoraFinal, Row)
                 ElseIf Col = conCOL_SonCapacProd_HoraFinal Then
                    If Row < (.Rows - 1) Then
                        Call PosColCapac(conCOL_SonCapacProd_HorInicial, Row)
                    End If
                 End If
                 
        End Select
    End With

End Sub

Private Sub grdMesAno_AfterEdit(ByVal Row As Long, ByVal Col As Long)

     Dim lngLINHA   As Long
     Dim strINDICE  As String
     Dim intRESP    As Integer

     With grdMesAno
          Select Case Col
                 Case conCOL_SonMesAno_Mes, _
                      conCOL_SonMesAno_Ano
                        
                        If Len(Trim(txtQtdePorHora.Text)) = 0 Then
                            MsgBox "ATENÇÂO" & vbCrLf & _
                                   "Informe a Qtde/Hora !!!", vbOKOnly + vbExclamation, "Aviso"
                            Exit Sub
                        End If
                        
                        strINDICE = ""
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_SonMesAno_Mes))) > 0 And Len(Trim(.Cell(flexcpText, Row, conCOL_SonMesAno_Ano))) > 0 Then
                            If Col = conCOL_SonMesAno_Mes Then
                                If cTipOper = "I" Then
                                    strINDICE = .Cell(flexcpText, Row, conCOL_SonMesAno_Mes) & .Cell(flexcpText, Row, conCOL_SonMesAno_Ano)
                                Else
                                    strINDICE = .Cell(flexcpText, Row, conCOL_SonMesAno_Mes) & .Cell(flexcpText, Row, conCOL_SonMesAno_Ano) & txtCodigo.Text
                                End If
                            ElseIf Col = conCOL_SonMesAno_Ano Then
                                If cTipOper = "I" Then
                                    strINDICE = .Cell(flexcpText, Row, conCOL_SonMesAno_Mes) & .Cell(flexcpText, Row, conCOL_SonMesAno_Ano)
                                Else
                                    strINDICE = .Cell(flexcpText, Row, conCOL_SonMesAno_Mes) & .Cell(flexcpText, Row, conCOL_SonMesAno_Ano) & txtCodigo.Text
                                End If
                            End If
                        End If
                        
                        If Len(Trim(strINDICE)) > 0 Then
                            If Trim(strINDICE) <> Trim(.Cell(flexcpText, Row, conCOL_SonMesAno_IndiceBKP)) Then
                                lngLINHA = grdCAPACPROD.FindRow(Trim(.Cell(flexcpText, Row, conCOL_SonMesAno_IndiceBKP)), , conCOL_SonCapacProd_CodIndice)
                                If lngLINHA > -1 Then
                                    intRESP = MsgBox("ATENÇÂO - Os dados serão apagados , Deseja continuar (Sim ou Não)", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
                                    If intRESP = vbYes Then
                                        Call objBLBFunc.ExcLinhaGrdFilho(grdCAPACPROD, conCOL_SonCapacProd_CodIndice, .Cell(flexcpText, .Row, conCOL_SonMesAno_IndiceBKP))
                                    Else
                                        .Cell(flexcpText, Row, conCOL_SonMesAno_Mes) = .Cell(flexcpText, Row, conCOL_SonMesAno_MesBKP)
                                        .Cell(flexcpText, Row, conCOL_SonMesAno_Ano) = .Cell(flexcpText, Row, conCOL_SonMesAno_AnoBKP)
                                        Exit Sub
                                    End If
                                End If
                            End If
                            
                            .Cell(flexcpText, Row, conCOL_SonMesAno_Indice) = strINDICE
                            .Cell(flexcpText, Row, conCOL_SonMesAno_IndiceBKP) = strINDICE
                            .Cell(flexcpText, Row, conCOL_SonMesAno_MesBKP) = .Cell(flexcpText, Row, conCOL_SonMesAno_Mes)
                            .Cell(flexcpText, Row, conCOL_SonMesAno_AnoBKP) = .Cell(flexcpText, Row, conCOL_SonMesAno_Ano)
                            
                            If objBLBFunc.FcVerifItensRepetidos(grdMesAno, Row, conCOL_SonMesAno_Indice, strINDICE) = False Then
                               MsgBox "Este Mês/Ano ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                               Call LimpaCamposGridMesAno(Row)
                               Exit Sub
                            End If
                            Call IncGrCapacProd(CLng(.Cell(flexcpText, Row, conCOL_SonMesAno_Mes)), CLng(.Cell(flexcpText, Row, conCOL_SonMesAno_Ano)), strINDICE)
                        End If
                        
          End Select
     End With

End Sub

Private Sub grdMesAno_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonMesAno_Indice, _
         conCOL_SonMesAno_QtdeCapac, _
         conCOL_SonMesAno_IndiceBKP, _
         conCOL_SonMesAno_IDPAI
         Cancel = True
    Case conCOL_SonMesAno_Mes, _
         conCOL_SonMesAno_Ano, _
         conCOL_SonMesAno_Carrega
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdMesAno.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdMesAno_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If (grdCAPACIDADE.Rows - 1) = 0 Or (grdCAPACIDADE.Row) = 0 Then
        MsgBox "ATENÇÃO - Selecione uma Linha de Produto !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim strINDICE As String
    
    Select Case Col
        Case conCOL_SonMesAno_Carrega
    
            If cTipOper = "C" Then Exit Sub
            
            With grdMesAno
                strINDICE = .Cell(flexcpText, Row, conCOL_SonMesAno_Indice)
                Call IncGrCapacProd(CLng(.Cell(flexcpText, Row, conCOL_SonMesAno_Mes)), CLng(.Cell(flexcpText, Row, conCOL_SonMesAno_Ano)), strINDICE)
            End With
    End Select

End Sub

Private Sub grdMesAno_Click()
    Call MostraFilhoMesAno
End Sub

Private Sub grdMesAno_RowColChange()
    Call MostraFilhoMesAno
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, Me
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

    ValidaCampos = False
    
    Dim I As Long
    
    
    If Len(Trim(txtDescricao.Text)) = 0 Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Campo Descrição não pode ser vázio !!!", vbOKOnly + vbExclamation, "Acviso"
               Exit Function
    End If
    
    If Len(Trim(txtQtdePorHora.Text)) = 0 Then
        MsgBox "ATENÇÂO" & vbCrLf & "Informe a Qtde Por Hora !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    If Not IsNumeric(txtQtdePorHora.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & "Informe a Qtde Por Hora !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
    
    ''Mês Ano
    For I = 1 To (grdMesAno.Rows - 1)
         If Len(Trim(grdMesAno.Cell(flexcpText, I, conCOL_SonMesAno_Mes))) = 0 Then
             MsgBox "ATENÇÃO - O Campo Mês na grid de Mês/Ano não pode ser vázio !!!", vbOKOnly + vbExclamation, "Aviso"
             Exit Function
         End If
         If Len(Trim(grdMesAno.Cell(flexcpText, I, conCOL_SonMesAno_Ano))) = 0 Then
             MsgBox "ATENÇÃO - O Campo Ano na grid de Mês/Ano não pode ser vázio !!!", vbOKOnly + vbExclamation, "Aviso"
             Exit Function
         End If
    Next I
            
    ''Capacidade Produtiva
    For I = 1 To (grdCAPACPROD.Rows - 1)
         If grdCAPACPROD.Cell(flexcpText, I, conCOL_SonCapacProd_AtivoSN) = 1 Then
            If Len(Trim(Replace(grdCAPACPROD.Cell(flexcpText, I, conCOL_SonCapacProd_HorInicial), ":", ""))) = 0 Then
                MsgBox "ATENÇÃO - O Campo Hora Inicial na grid de Capacidade Produtiva não pode ser vázio !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
            If Len(Trim(Replace(grdCAPACPROD.Cell(flexcpText, I, conCOL_SonCapacProd_Parada), ":", ""))) = 0 Then
                MsgBox "ATENÇÃO - O Campo Parada na grid de Capacidade Produtiva não pode ser vázio !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
            If Len(Trim(Replace(grdCAPACPROD.Cell(flexcpText, I, conCOL_SonCapacProd_HoraFinal), ":", ""))) = 0 Then
                MsgBox "ATENÇÃO - O Campo Hora Final na grid de Capacidade Produtiva não pode ser vázio !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
         End If
    Next I
        
        
    ValidaCampos = True
     
End Function


Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    Call ConfGrd
    Call InitGridMesAno
    Call InitGridCapacProd
    
    If cTipOper = "I" Then iCodigo = 0
    objCADCADGRUPLINHA.CODIGO = iCodigo
    optATIVOSN(0).Value = True
    
    Call CarregaCampos
    
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
    End If
End Sub

Private Sub CarregaCampos()

On Error GoTo Err_CarregaCampos
    
    If objCADCADGRUPLINHA.Carrega_campos(strNOMETABELA) = True Then
    
        txtCodigo.Text = objCADCADGRUPLINHA.CODIGO
        txtDescricao.Text = objCADCADGRUPLINHA.DESCRI
        optATIVOSN(objCADCADGRUPLINHA.ATIVO).Value = True
        txtQTDECAPACGERAL.Text = objCADCADGRUPLINHA.CAPACTOTLINHA
        txtQtdePorHora.Text = objCADCADGRUPLINHA.QTDPORHORA
        
        Call PopGrdLanctos
        Call PopGrdMesAno
        Call PopGrdCapacProd
    
        With grdMesAno
            If (.Rows - 1) > 0 Then
                .Row = 1
                Call grdMesAno_Click
            End If
        End With
    
    End If
    
    Exit Sub

Err_CarregaCampos:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CarregaCampos", Me.Name, "CarregaCampos")
    
End Sub

Private Sub ConfGrd()

    With grdCAPACIDADE

       .Cols = conColumnsIn_GRPLINHA
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_GRPLINHA_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_CodCor) = ""
       .ColDataType(conCOL_GRPLINHA_CodCor) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_PesqCor) = ""
       .ColDataType(conCOL_GRPLINHA_PesqCor) = flexDTString
       .ColComboList(conCOL_GRPLINHA_PesqCor) = "..."
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_DescCor) = ""
       .ColDataType(conCOL_GRPLINHA_DescCor) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_IDCAPC) = ""
       .ColDataType(conCOL_GRPLINHA_IDCAPC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_IDINTERNO) = ""
       .ColDataType(conCOL_GRPLINHA_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_IDINTERNO) = ""
       .ColDataType(conCOL_GRPLINHA_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_NECKINSN) = ""
       .ColDataType(conCOL_GRPLINHA_NECKINSN) = flexDTString
       .ColComboList(conCOL_GRPLINHA_NECKINSN) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_INDICE) = ""
       .ColDataType(conCOL_GRPLINHA_INDICE) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_GRPLINHA_HOMOLOGSN) = ""
       .ColDataType(conCOL_GRPLINHA_HOMOLOGSN) = flexDTString
       .ColComboList(conCOL_GRPLINHA_HOMOLOGSN) = "|#1;Sim|#0;Não"
       
       
       .ColWidth(conCOL_GRPLINHA_CodCor) = 1000
       .ColWidth(conCOL_GRPLINHA_PesqCor) = 300
       .ColWidth(conCOL_GRPLINHA_DescCor) = 3000
       .ColWidth(conCOL_GRPLINHA_IDCAPC) = 0
       .ColWidth(conCOL_GRPLINHA_IDINTERNO) = 0
       .ColWidth(conCOL_GRPLINHA_NECKINSN) = 0
       .ColWidth(conCOL_GRPLINHA_INDICE) = 0
       .ColWidth(conCOL_GRPLINHA_HOMOLOGSN) = 1400
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack

    End With
    
End Sub

Private Sub IncRegGrid()
   
    If cTipOper = "C" Then Exit Sub
    
    If objBLBFunc.FcExisteLinhaVazia(grdCAPACIDADE, conCOL_GRPLINHA_CodCor) = False Then Exit Sub
    
    With grdCAPACIDADE
        .AddItem "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 0 & vbTab & _
                 "" & vbTab & _
                 0 & vbTab & _
                 "" & vbTab & _
                 ""
    End With
   
End Sub

Private Sub PesqCapc(lngROW As Long)
    With grdCAPACIDADE
        If (.Rows - 1) = 0 Then Exit Sub
            
        If cTipOper = "C" Then Exit Sub
        
        Dim lngLINHA                    As Long
        Dim strINDICE                   As String
        ReDim arrCAMPOS(1 To 3, 1 To 5) As String
        ReDim arrTABELA(1 To 1) As String
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLINHAPRODUTO" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
        
        arrTABELA(1) = sSql
        
        arrCAMPOS(1, 1) = "SGI_CODIGO"
        arrCAMPOS(1, 2) = "N"
        arrCAMPOS(1, 3) = "Código"
        arrCAMPOS(1, 4) = "1000"
        arrCAMPOS(1, 5) = "SGI_CODIGO"
        
        arrCAMPOS(2, 1) = "SGI_CODLIN"
        arrCAMPOS(2, 2) = "N"
        arrCAMPOS(2, 3) = "Co.Linha"
        arrCAMPOS(2, 4) = "1000"
        arrCAMPOS(2, 5) = "SGI_CODLIN"
        
        arrCAMPOS(3, 1) = "SGI_DESCRI"
        arrCAMPOS(3, 2) = "S"
        arrCAMPOS(3, 3) = "Descrição"
        arrCAMPOS(3, 4) = "5000"
        arrCAMPOS(3, 5) = "SGI_DESCRI"
        
        varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Capacidade")
        
        If Len(Trim(varRETORNO)) = 0 Then Exit Sub
        
        ''strINDICE = Trim(varRETORNO) & .Cell(flexcpText, lngROW, conCOL_GRPLINHA_NECKINSN) & .Cell(flexcpText, lngROW, conCOL_GRPLINHA_HOMOLOGSN)
        strINDICE = Trim(varRETORNO)
        If objBLBFunc.FcVerifItensRepetidos(grdCAPACIDADE, lngROW, conCOL_GRPLINHA_INDICE, Trim(strINDICE)) = False Then
           MsgBox "ATENÇÃO" & vbCrLf & " A Capacidade ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
           grdCAPACIDADE.Cell(flexcpText, lngROW, conCOL_GRPLINHA_CodCor) = Empty
           grdCAPACIDADE.Cell(flexcpText, lngROW, conCOL_GRPLINHA_DescCor) = Empty
           grdCAPACIDADE.Cell(flexcpText, lngROW, conCOL_GRPLINHA_IDCAPC) = Empty
           grdCAPACIDADE.Cell(flexcpText, lngROW, conCOL_GRPLINHA_INDICE) = Empty
           grdCAPACIDADE.Cell(flexcpText, lngROW, conCOL_GRPLINHA_NECKINSN) = 0
           grdCAPACIDADE.Cell(flexcpText, lngROW, conCOL_GRPLINHA_HOMOLOGSN) = 0
           Exit Sub
        End If
        
        .Cell(flexcpText, lngROW, conCOL_GRPLINHA_IDCAPC) = varRETORNO
        .Cell(flexcpText, lngROW, conCOL_GRPLINHA_CodCor) = PegaCodLinha(varRETORNO)
        .Cell(flexcpText, lngROW, conCOL_GRPLINHA_DescCor) = PegaDescrLinha(varRETORNO)
        .Cell(flexcpText, lngROW, conCOL_GRPLINHA_INDICE) = strINDICE
    
    End With
End Sub


Private Function PegaCodLinha(strCODLINHA As String) As String

    PegaCodLinha = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = " & Trim(strCODLINHA) & vbCrLf
    sSql = sSql & "   And SGI_FILIAL = " & FILIAL

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaCodLinha = BREC!SGI_CODLIN
    BREC.Close
    
End Function

Private Function PegaDescrLinha(strIDlinha As String) As String
    
    PegaDescrLinha = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strIDlinha
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrLinha = BREC2!SGI_DESCRI
    BREC2.Close
    
End Function

Private Function PegaIDLinha(strCODLINHA As String) As String

    PegaIDLinha = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODLIN = " & Trim(strCODLINHA) & vbCrLf
    sSql = sSql & "   And SGI_FILIAL = " & FILIAL

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaIDLinha = BREC!SGI_CODIGO
    BREC.Close
    
End Function

Private Sub PopGrdLanctos()

    Dim I As Integer
    
    arrCAPACIDADEIT = objCADCADGRUPLINHA.CAPACIDADELIN
    If IsArray(arrCAPACIDADEIT) Then
        With grdCAPACIDADE
            For I = 1 To UBound(arrCAPACIDADEIT)
                .AddItem "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrCAPACIDADEIT(I, 1) & vbTab & _
                         "" & vbTab & _
                         arrCAPACIDADEIT(I, 2) & vbTab & _
                         "" & vbTab & _
                         arrCAPACIDADEIT(I, 3)
                         
        
                .Cell(flexcpText, (.Rows - 1), conCOL_GRPLINHA_CodCor) = PegaCodLinha(Str(arrCAPACIDADEIT(I, 1)))
                .Cell(flexcpText, (.Rows - 1), conCOL_GRPLINHA_DescCor) = PegaDescrLinha(Str(arrCAPACIDADEIT(I, 1)))
                
                If Len(Trim(arrCAPACIDADEIT(I, 5))) = 0 Then
                    .Cell(flexcpText, (.Rows - 1), conCOL_GRPLINHA_INDICE) = Trim(Str(arrCAPACIDADEIT(I, 1) & arrCAPACIDADEIT(I, 2)) & arrCAPACIDADEIT(I, 3))
                Else
                    .Cell(flexcpText, (.Rows - 1), conCOL_GRPLINHA_INDICE) = arrCAPACIDADEIT(I, 4)
                End If
                
                .Cell(flexcpText, (.Rows - 1), conCOL_GRPLINHA_IDINTERNO) = arrCAPACIDADEIT(I, 5)
            
            
            Next I
        End With
    End If

End Sub

Private Sub InitGridMesAno()

    With grdMesAno
    
       .Cols = conColumnsIn_SonMesAno
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMesAno_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_Mes) = ""
       .ColDataType(conCOL_SonMesAno_Mes) = flexDTString
       .ColComboList(conCOL_SonMesAno_Mes) = objCADCADGRUPLINHA.PreenchComboMes
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_Ano) = ""
       .ColDataType(conCOL_SonMesAno_Ano) = flexDTString
       .ColComboList(conCOL_SonMesAno_Ano) = objCADCADGRUPLINHA.PreenchComboAno
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_Indice) = ""
       .ColDataType(conCOL_SonMesAno_Indice) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_QtdeCapac) = ""
       .ColDataType(conCOL_SonMesAno_QtdeCapac) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_Carrega) = ""
       .ColDataType(conCOL_SonMesAno_Carrega) = flexDTString
       .ColComboList(conCOL_SonMesAno_Carrega) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_IndiceBKP) = ""
       .ColDataType(conCOL_SonMesAno_IndiceBKP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_MesBKP) = ""
       .ColDataType(conCOL_SonMesAno_MesBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_AnoBKP) = ""
       .ColDataType(conCOL_SonMesAno_AnoBKP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMesAno_IDPAI) = ""
       .ColDataType(conCOL_SonMesAno_IDPAI) = flexDTLong
       
       .ColWidth(conCOL_SonMesAno_Mes) = 1200
       .ColWidth(conCOL_SonMesAno_Ano) = 1000
       .ColWidth(conCOL_SonMesAno_Indice) = 0
       .ColWidth(conCOL_SonMesAno_QtdeCapac) = 1400
       .ColWidth(conCOL_SonMesAno_Carrega) = 300
       .ColWidth(conCOL_SonMesAno_IndiceBKP) = 0
       .ColWidth(conCOL_SonMesAno_AnoBKP) = 0
       .ColWidth(conCOL_SonMesAno_MesBKP) = 0
       .ColWidth(conCOL_SonMesAno_IDPAI) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGridMesAno()
   
    If objBLBFunc.FcExisteLinhaVazia(grdMesAno, conCOL_SonMesAno_Mes) = False Then Exit Sub
    If objBLBFunc.FcExisteLinhaVazia(grdMesAno, conCOL_SonMesAno_Ano) = False Then Exit Sub
    If (grdCAPACIDADE.Rows - 1) = 0 Then
        MsgBox "ATENÇÂO - Não existe dados na Gride de Capacidade !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    grdMesAno.AddItem "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      ""
                      
    If cTipOper <> "I" Then grdMesAno.Cell(flexcpText, (grdMesAno.Rows - 1), conCOL_SonMesAno_IDPAI) = Trim(txtCodigo.Text)
                       
End Sub


Private Sub LimpaCamposGridMesAno(lngROWS As Long)
    With grdMesAno
            .Cell(flexcpText, lngROWS, conCOL_SonMesAno_Mes) = Empty
            .Cell(flexcpText, lngROWS, conCOL_SonMesAno_Ano) = Empty
            .Cell(flexcpText, lngROWS, conCOL_SonMesAno_Indice) = Empty
    End With
End Sub

Private Sub IncGrCapacProd(lngMES As Long, lngANO As Long, strCODINDICE As String)
    
    Dim I               As Integer
    Dim lngQtdDias      As Long
    Dim dtInicial       As Date
    Dim dtfinal         As Date
    Dim dtDia           As Date
    Dim arrDIASEMANA()  As String
    Dim strINDICE       As String
    Dim lngPESQUISA     As Long
    
    dtInicial = CDate(1 & "/" & lngMES & "/" & lngANO)
    
    If lngMES < 12 Then dtfinal = (CDate(1 & "/" & (lngMES + 1) & "/" & lngANO) - 1)
    If lngMES = 12 Then dtfinal = CDate(31 & "/" & lngMES & "/" & lngANO)
    
    lngQtdDias = Day(dtfinal)
    
    With grdCAPACPROD
        For I = 1 To lngQtdDias
            dtDia = CDate(I & "/" & lngMES & "/" & lngANO)
            
            If cTipOper = "I" Then
                strINDICE = Trim(Replace(Format(dtDia, "DD/MM/YYYY"), "/", ""))
            Else
                strINDICE = Trim(Replace(Format(dtDia, "DD/MM/YYYY"), "/", "")) & txtCodigo.Text
            End If
            
            lngPESQUISA = .FindRow(strINDICE, , conCOL_SonCapacProd_INDICE)
            If lngPESQUISA = -1 Then
                .AddItem Format(dtDia, "DD/MM/YYYY") & vbTab & _
                         PegaDescDiaSemana(Weekday(dtDia)) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         strCODINDICE & vbTab & _
                         Weekday(dtDia) & vbTab & _
                         0 & vbTab & _
                         strINDICE & vbTab & _
                         ""
                
                If cTipOper <> "I" Then .Cell(flexcpText, (.Rows - 1), conCOL_SonCapacProd_IDPAI) = txtCodigo.Text
                         
            End If
        Next I
    End With
    
End Sub


Private Sub InitGridCapacProd()

    With grdCAPACPROD
    
       .Cols = conColumnsIn_SonCapacProd
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonCapacProd_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_Data) = ""
       .ColDataType(conCOL_SonMesAno_Mes) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_DiaDesc) = ""
       .ColDataType(conCOL_SonCapacProd_DiaDesc) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_HorInicial) = ""
       .ColDataType(conCOL_SonCapacProd_HorInicial) = flexDTString
       .ColEditMask(conCOL_SonCapacProd_HorInicial) = "##:##"
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_Parada) = ""
       .ColDataType(conCOL_SonCapacProd_Parada) = flexDTString
       .ColEditMask(conCOL_SonCapacProd_Parada) = "##:##"
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_HoraFinal) = ""
       .ColDataType(conCOL_SonCapacProd_HoraFinal) = flexDTString
       .ColEditMask(conCOL_SonCapacProd_HoraFinal) = "##:##"
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_TotalHoras) = ""
       .ColDataType(conCOL_SonCapacProd_TotalHoras) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_TotalPecas) = ""
       .ColDataType(conCOL_SonCapacProd_TotalPecas) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_CodIndice) = ""
       .ColDataType(conCOL_SonCapacProd_CodIndice) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_DiaSem) = ""
       .ColDataType(conCOL_SonCapacProd_DiaSem) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_AtivoSN) = ""
       .ColDataType(conCOL_SonCapacProd_AtivoSN) = flexDTString
       .ColComboList(conCOL_SonCapacProd_AtivoSN) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_INDICE) = ""
       .ColDataType(conCOL_SonCapacProd_INDICE) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCapacProd_IDPAI) = ""
       .ColDataType(conCOL_SonCapacProd_IDPAI) = flexDTLong
       
       .ColWidth(conCOL_SonCapacProd_Data) = 1000
       .ColWidth(conCOL_SonCapacProd_DiaDesc) = 1200
       .ColWidth(conCOL_SonCapacProd_HorInicial) = 1000
       .ColWidth(conCOL_SonCapacProd_Parada) = 1000
       .ColWidth(conCOL_SonCapacProd_HoraFinal) = 1000
       .ColWidth(conCOL_SonCapacProd_TotalPecas) = 1300
       .ColWidth(conCOL_SonCapacProd_CodIndice) = 0
       .ColWidth(conCOL_SonCapacProd_DiaSem) = 0
       .ColWidth(conCOL_SonCapacProd_AtivoSN) = 800
       .ColWidth(conCOL_SonCapacProd_INDICE) = 0
       .ColWidth(conCOL_SonCapacProd_IDPAI) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub MostraFilhoMesAno()
    With grdMesAno
        If (.Rows - 1) = 0 Then Exit Sub
        If .Rows = 0 Then Exit Sub
        Call MostraLinhasFilho(.Cell(flexcpText, .Row, conCOL_SonMesAno_Indice), grdCAPACPROD, conCOL_SonCapacProd_CodIndice)
    End With
End Sub

Private Sub MostraLinhasFilho(strINDICEPAI As String, grdFILHO As Variant, lngCOLFILHO As Long)
    Dim I As Integer
    With grdFILHO
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, lngCOLFILHO) = strINDICEPAI Then
               .RowHidden(I) = False
            Else
               .RowHidden(I) = True
            End If
        Next I
    End With
End Sub

Private Function CalcTotPecasHora(strTOTHORAS As String, lngQTPCDEHORAS As Long) As Long

    CalcTotPecasHora = 0
    
    Dim arrHORAMIN      As Variant
    Dim lngTotPecasHora As Long

    
    '' -------------------------
    '' Qtde de Pecas
    arrHORAMIN = Split(strTOTHORAS, ":")
    lngTotPecasHora = (arrHORAMIN(0) * lngQTPCDEHORAS)
    
    CalcTotPecasHora = lngTotPecasHora
    '' -------------------------

End Function

Private Function CalcTotCapacidade(strINDICE As String) As Long

    CalcTotCapacidade = 0
    
    Dim I                   As Long
    Dim lngSomaQtdeCapac    As Long
    
    lngSomaQtdeCapac = 0
    
    With grdCAPACPROD
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_SonCapacProd_AtivoSN) = 1 Then
                If Trim(grdMesAno.Cell(flexcpText, grdMesAno.Row, conCOL_SonMesAno_Indice)) = Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_CodIndice)) Then
                    If Len(Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_TotalPecas))) > 0 Then lngSomaQtdeCapac = lngSomaQtdeCapac + .Cell(flexcpText, I, conCOL_SonCapacProd_TotalPecas)
                End If
            End If
        Next I
    End With
    
    CalcTotCapacidade = lngSomaQtdeCapac

End Function

Private Sub PopGrdMesAno()
    Dim I As Long
    arrMESANO = objCADCADGRUPLINHA.MESANO
    If IsArray(arrMESANO) Then
        With grdMesAno
            For I = 1 To UBound(arrMESANO)
                .AddItem arrMESANO(I, 1) & vbTab & _
                         arrMESANO(I, 2) & vbTab & _
                         arrMESANO(I, 3) & vbTab & _
                         IIf(Len(Trim(arrMESANO(I, 4))) > 0, arrMESANO(I, 4), "") & vbTab & _
                         "" & vbTab & _
                         arrMESANO(I, 3) & vbTab & _
                         arrMESANO(I, 1) & vbTab & _
                         arrMESANO(I, 2) & vbTab & _
                         arrMESANO(I, 5)

                         
            Next I
        End With
    End If
End Sub


Private Sub PopGrdCapacProd()
    
    Dim I               As Long
    Dim arrDIASEMANA()  As String
    
    arrCAPACPROD = objCADCADGRUPLINHA.CAPACPROD
    If IsArray(arrCAPACPROD) Then
        With grdCAPACPROD
            For I = 1 To UBound(arrCAPACPROD)
                .AddItem arrCAPACPROD(I, 1) & vbTab & _
                         PegaDescDiaSemana(Weekday(CDate(arrCAPACPROD(I, 1)))) & vbTab & _
                         arrCAPACPROD(I, 2) & vbTab & _
                         arrCAPACPROD(I, 3) & vbTab & _
                         arrCAPACPROD(I, 4) & vbTab & _
                         arrCAPACPROD(I, 5) & vbTab & _
                         arrCAPACPROD(I, 6) & vbTab & _
                         arrCAPACPROD(I, 7) & vbTab & _
                         arrCAPACPROD(I, 8) & vbTab & _
                         arrCAPACPROD(I, 9) & vbTab & _
                         arrCAPACPROD(I, 10) & vbTab & _
                         arrCAPACPROD(I, 11)

            Next I
        End With
    End If
End Sub

Private Function CalcTotalGerLinha() As Long

    CalcTotalGerLinha = 0
    
    Dim I As Long

    With grdMesAno
        For I = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, I, conCOL_SonMesAno_QtdeCapac))) > 0 Then
                CalcTotalGerLinha = CalcTotalGerLinha + CLng(.Cell(flexcpText, I, conCOL_SonMesAno_QtdeCapac))
            End If
        Next I
    End With

End Function

Private Sub PosColCapac(lngPOSCOL As Long, lngPOSROL As Long)
    
On Error GoTo Err_PosCol
    
    With grdCAPACPROD
        .SetFocus
        .Row = lngPOSROL
        .Col = lngPOSCOL
    End With
    
    Exit Sub
    
Err_PosCol:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PosCol()", Me.Name, "PosCol()", strCAMARQERRO)
    
End Sub


Private Function PegaDescDiaSemana(lngDIASEMANA As Long) As String
    
    PegaDescDiaSemana = ""
    
    ReDim arrDIASEMANA(1 To 7) As String
    
    arrDIASEMANA(1) = "Domingo"
    arrDIASEMANA(2) = "Segunda"
    arrDIASEMANA(3) = "Terça"
    arrDIASEMANA(4) = "Quarta"
    arrDIASEMANA(5) = "Quinta"
    arrDIASEMANA(6) = "Sexta"
    arrDIASEMANA(7) = "Sabado"

    PegaDescDiaSemana = Trim(arrDIASEMANA(lngDIASEMANA))

End Function

Private Sub txtQtdePorHora_GotFocus()
    objBLBFunc.SelecionaCampos txtQtdePorHora.Name, Me
End Sub

Private Sub txtQtdePorHora_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQtdePorHora.Text
End Sub

Private Sub txtQtdePorHora_Validate(Cancel As Boolean)

    If Len(Trim(txtQtdePorHora.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtQtdePorHora.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtQtdePorHora.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtQtdePorHora.Text = Trim(Replace(Replace(txtQtdePorHora.Text, ",", ""), ".", ""))

    If (grdMesAno.Rows - 1) = 0 Or grdMesAno.Row = 0 Then Exit Sub

    Dim I               As Integer
    Dim lngQTDPORHORA   As Long
    Dim strINDICE       As String
    
    lngQTDPORHORA = CLng(txtQtdePorHora.Text)
    
    '' Recalcular Qtde Total Produzida
    strINDICE = grdMesAno.Cell(flexcpText, grdMesAno.Row, conCOL_SonMesAno_Indice)
    With grdCAPACPROD
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_SonCapacProd_CodIndice) = strINDICE Then
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonCapacProd_TotalHoras))) > 0 Then
                    .Cell(flexcpText, I, conCOL_SonCapacProd_TotalPecas) = CalcTotPecasHora(.Cell(flexcpText, I, conCOL_SonCapacProd_TotalHoras), lngQTDPORHORA)
                End If
            End If
        Next I
    End With
    
    '' Total de Peças no Mês
    grdMesAno.Cell(flexcpText, grdMesAno.Row, conCOL_SonMesAno_QtdeCapac) = CalcTotCapacidade(grdMesAno.Cell(flexcpText, grdMesAno.Row, conCOL_SonMesAno_Indice))
    '' Total Geral da Capacidade
    txtQTDECAPACGERAL.Text = CalcTotalGerLinha

End Sub
