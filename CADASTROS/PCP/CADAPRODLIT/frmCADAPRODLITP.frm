VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADAPRODLITP 
   Caption         =   "Boletim de Produção de Litografia"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15405
   LinkTopic       =   "Form2"
   ScaleHeight     =   8565
   ScaleWidth      =   15405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   15375
      Begin VSFlex8LCtl.VSFlexGrid grdAPONT 
         Height          =   6615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   15135
         _cx             =   26696
         _cy             =   11668
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
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15375
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   200
         Width           =   1935
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "txtCampos"
         Top             =   200
         Width           =   11535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtro:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo:"
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
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame cmdFECHA 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   15375
      Begin VB.CommandButton cmdOrden 
         Caption         =   "&Ordem"
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
         Left            =   14400
         Picture         =   "frmCADAPRODLITP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCanFiltro 
         Caption         =   "&Desfas"
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
         Left            =   13560
         Picture         =   "frmCADAPRODLITP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdExclui 
         Caption         =   "&Excluir"
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
         Left            =   3840
         Picture         =   "frmCADAPRODLITP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdAltera 
         Caption         =   "&Alterar <F6>"
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
         Picture         =   "frmCADAPRODLITP.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Altera Empresa "
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdInclui 
         Caption         =   "&Incluir <F5>"
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
         Left            =   1440
         Picture         =   "frmCADAPRODLITP.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar <ESC>"
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
         Picture         =   "frmCADAPRODLITP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Interval        =   50000
         Left            =   5760
         Top             =   240
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&primir"
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
         Left            =   4680
         Picture         =   "frmCADAPRODLITP.frx":129C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime o Vale"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADAPRODLITP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strACESSO        As String
Public strUsuario       As String
Public lngCodUsuaro     As Long
Public intFILIALPED     As Integer
Public lngCodVendedor   As Long

Dim objFuncoes          As New clsFuncoes
Dim objCADAPRODLITP     As New clsCADAPRODLIT

Dim iCodigo             As Long
Dim lngCodLog           As Long
Dim strFILIAL           As String
Dim strNOMTABELA        As String

Const conCOL_SonApont_Codigo                   As Integer = 0
Const conCOL_SonApont_DataMov                  As Integer = 1
Const conCOL_SonApont_CodMaq                   As Integer = 2
Const conCOL_SonApont_DescMaq                  As Integer = 3
Const conCOL_SonApont_CodTurno                 As Integer = 4
Const conCOL_SonApont_DescTurno                As Integer = 5
Const conCOL_SonApont_FormatString             As String = "=Cód.Lcto|Data Lcto|Cod.Maq|Desc Máquina|Cod.Turno|Desc.Turno"
Const conColumnsIn_SonApont                    As Integer = 6

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    ''Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    ''Set objCADAPONTPRODP = CreateObject("CADAPONTPROD.clsCADAPONTPROD")
    ''Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")

    objCADAPRODLITP.FILIAL = FILIAL
    objFuncoes.LimpaCampos Me
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strNOMTABELA = ""
    If intFILIALPED = 0 Then strFILIAL = "NOVALATA"
    If intFILIALPED = 1 Then
        strFILIAL = "STEEL"
        strNOMTABELA = "_STEEL"
    End If
   
    Call ConfTooTipText
    Call AbilitaCampos
    Call ConfGrid
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    Call ConfFiltro


    Me.Caption = Me.Caption

End Sub

Private Sub DestroyObjeto()
    Set objFuncoes = Nothing
    Set objCADAPRODLITP = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyObjeto
End Sub

Private Sub ConfTooTipText()
    cmdVoltar.ToolTipText = "Volta ao menu Principal"
    cmdInclui.ToolTipText = "Inclui um novo movimento"
    cmdAltera.ToolTipText = "Altera o movimento gerado"
    cmdExclui.ToolTipText = "Exclui o movimento gerado"
    cmdCanFiltro.ToolTipText = "Desfaz o filtro"
    cmdOrden.ToolTipText = "Ordena os dados conforme o filtro"
End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADAPRODLITP.AtivoDesativo(strNOMTABELA)
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo

End Sub

Private Sub ConfGrid()

    With grdAPONT
    
       .Cols = conColumnsIn_SonApont
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonApont_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonApont_Codigo) = ""
       .ColDataType(conCOL_SonApont_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonApont_DataMov) = ""
       .ColDataType(conCOL_SonApont_DataMov) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonApont_CodMaq) = ""
       .ColDataType(conCOL_SonApont_CodMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonApont_DescMaq) = ""
       .ColDataType(conCOL_SonApont_DescMaq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonApont_CodTurno) = ""
       .ColDataType(conCOL_SonApont_CodTurno) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonApont_DescTurno) = ""
       .ColDataType(conCOL_SonApont_DescTurno) = flexDTString
       
       .ColWidth(conCOL_SonApont_Codigo) = 1000
       .ColWidth(conCOL_SonApont_DataMov) = 1200
       .ColWidth(conCOL_SonApont_CodMaq) = 1000
       .ColWidth(conCOL_SonApont_DescMaq) = 5000
       .ColWidth(conCOL_SonApont_CodTurno) = 1000
       .ColWidth(conCOL_SonApont_DescTurno) = 5000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       .GridLineWidth = 6
       .GridLines = flexGridExplorer
       
    End With
    
End Sub

Private Sub ConfFiltro()
    cboFiltro.Clear
    cboFiltro.AddItem "Cód.Lcto"
    cboFiltro.AddItem "Data.Lcto"
    cboFiltro.AddItem "Cód.Maquina"
    cboFiltro.AddItem "Descrição da Máquina"
    cboFiltro.AddItem "Cód.Turno"
    cboFiltro.AddItem "Descrição do Turno"
    cboFiltro.ListIndex = 0
End Sub

