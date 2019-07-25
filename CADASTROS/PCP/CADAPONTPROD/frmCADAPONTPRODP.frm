VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADAPONTPRODP 
   Caption         =   "Apontamento de Produção"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   15390
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   15375
      Begin VSFlex8LCtl.VSFlexGrid grdAPONT 
         Height          =   6735
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   15135
         _cx             =   26696
         _cy             =   11880
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
   Begin VB.Frame cmdFECHA 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   15375
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
         Picture         =   "frmCADAPONTPRODP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprime o Vale"
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   50000
         Left            =   5760
         Top             =   240
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
         Picture         =   "frmCADAPONTPRODP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   1335
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
         Picture         =   "frmCADAPONTPRODP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   120
         Width           =   1215
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
         Picture         =   "frmCADAPONTPRODP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Altera Empresa "
         Top             =   120
         Width           =   1215
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
         Picture         =   "frmCADAPONTPRODP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmCADAPONTPRODP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
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
         Picture         =   "frmCADAPONTPRODP.frx":129C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   11535
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   200
         Width           =   1935
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
         TabIndex        =   4
         Top             =   240
         Width           =   645
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
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCADAPONTPRODP"
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
Dim objCADAPONTPRODP    As New clsCADAPONTPROD
Dim objREL              As Object

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

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("I", strACESSO) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()

On Error GoTo Err_cmdCanFiltro_Click
   
   Call AbilitaCampos
   Call ConfGrid
   
   Exit Sub
   
Err_cmdCanFiltro_Click:
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "V", "Função : cmdCanFiltro_Click()", Me.Name, "cmdCanFiltro_Click()", strCAMARQERRO)

End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strACESSO) = False Then Exit Sub
  
  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADAPONTPRODP.GRAVA("E", strNOMTABELA) = False Then Exit Sub
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call AbilitaCampos
  Call ConfGrid

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strACESSO) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
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

Private Sub Destroy_ObjetoP()
    Set objFuncoes = Nothing
    Set objCADAPONTPRODP = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Load()

    ''Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    ''Set objCADAPONTPRODP = CreateObject("CADAPONTPROD.clsCADAPONTPROD")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")

    objCADAPONTPRODP.FILIAL = FILIAL
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

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_ObjetoP
End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADAPONTPRODP.AtivoDesativo(strNOMTABELA)
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo

End Sub

Private Sub ConfTooTipText()
    cmdVoltar.ToolTipText = "Volta ao menu Principal"
    cmdInclui.ToolTipText = "Inclui um novo movimento"
    cmdAltera.ToolTipText = "Altera o movimento gerado"
    cmdExclui.ToolTipText = "Exclui o movimento gerado"
    cmdCanFiltro.ToolTipText = "Desfaz o filtro"
    cmdOrden.ToolTipText = "Ordena os dados conforme o filtro"
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

Private Sub Operacao(strOperacao As String)
 
    iCodigo = 0
 
    With grdAPONT
        If strOperacao <> "I" Then
            If (.Rows - 1) = 0 Or .Row = 0 Then
                MsgBox "Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Sub
            End If
        End If
        If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_SonApont_Codigo))
    End With
    
    frmCADAPONTPROD.cCaminho = cCaminho
    frmCADAPONTPROD.Linha = Linha
    frmCADAPONTPROD.iCodigo = iCodigo
    frmCADAPONTPROD.cTipOper = strOperacao
    frmCADAPONTPROD.FILIAL = FILIAL
    frmCADAPONTPROD.strACESSO = strACESSO
    frmCADAPONTPROD.strMODPAI = Me.Name
    frmCADAPONTPROD.strUsuario = strUsuario
    frmCADAPONTPROD.lngCODUSUARIO = lngCodUsuaro
    frmCADAPONTPROD.intFILIALPED = intFILIALPED
    frmCADAPONTPROD.Show vbModal
  
    Call ConfGrid
    Call AbilitaCampos

End Sub


Private Sub Ordem()

On Error GoTo Err_Ordem
    
    If BREC.State = 1 Then BREC.Close
    
    Call ConfGrid
     
    Dim strCAMPO        As String
    
    txtCampos.Text = ""
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CABEC.* " & vbCrLf
    sSql = sSql & "      ,CADMAQ.SGI_DESCRI As SGI_DESCRIMAQ" & vbCrLf
    sSql = sSql & "      ,TURNO.SGI_DESCRI  As SGI_DESCRITURN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPPRODUCAO" & strNOMTABELA & " CABEC" & vbCrLf
    sSql = sSql & "      ,SGI_CADMAQUINA    CADMAQ" & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN   TURNO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CABEC.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And CADMAQ.SGI_FILIAL = CABEC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CADMAQ.SGI_CODIGO = CABEC.SGI_CODMAQ" & vbCrLf
    sSql = sSql & "   And TURNO.SGI_FILIAL  = CABEC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And TURNO.SGI_CODIGO  = CABEC.SGI_CODTUN" & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order By CABEC.SGI_CODIGO"
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order By CABEC.SGI_DTLANC"
    If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order By CABEC.SGI_CODMAQ"
    If cboFiltro.ListIndex = 3 Then sSql = sSql & "Order By CADMAQ.SGI_DESCRIMAQ"
    If cboFiltro.ListIndex = 4 Then sSql = sSql & "Order By CABEC.SGI_CODTUN"
    If cboFiltro.ListIndex = 5 Then sSql = sSql & "Order By TURNO.SGI_DESCRITURN"
    
    BREC.Open sSql, adoBanco_Dados
      
    Do While Not BREC.EOF
        
        strCAMPO = BREC!SGI_CODIGO & vbTab & _
                   Format(BREC!SGI_DTLANC, "DD/MM/YYYY") & vbTab & _
                   BREC!SGI_CODMAQ & vbTab & _
                   BREC!SGI_DESCRIMAQ & vbTab & _
                   BREC!SGI_CODTUN & vbTab & _
                   BREC!SGI_DESCRITURN
        
        grdAPONT.AddItem strCAMPO
        
        BREC.MoveNext
    Loop
    BREC.Close

    Exit Sub

Err_Ordem:
    
    If BREC.State = 1 Then BREC.Close
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : Ordem()", Me.Name, "Ordem()", strCAMARQERRO)


End Sub


Private Sub grdAPONT_Click()
    With grdAPONT
        If (.Rows - 1) > 0 And .Row > 0 Then objCADAPONTPRODP.Codigo = CLng(.Cell(flexcpText, .Row, conCOL_SonApont_Codigo))
    End With
End Sub

Private Sub grdAPONT_DblClick()
   If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
   With grdAPONT
        If (.Rows - 1) > 0 And .Row > 0 Then Call Operacao("C")
   End With
End Sub

Private Sub grdAPONT_RowColChange()
    With grdAPONT
        If (.Rows - 1) > 0 And .Row > 0 Then objCADAPONTPRODP.Codigo = CLng(.Cell(flexcpText, .Row, conCOL_SonApont_Codigo))
    End With
End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, Me
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim lngCodSep   As Long
    Dim strCampos   As String
    Dim strDEBCRED  As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    
    
    If cboFiltro.ListIndex = 0 Or _
       cboFiltro.ListIndex = 2 Or _
       cboFiltro.ListIndex = 4 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
    ElseIf cboFiltro.ListIndex = 1 Then
        If IsDate(txtCampos.Text) = False Then
           MsgBox "Somente é permitido Datas !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
    End If
        
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CABEC.*" & vbCrLf
    sSql = sSql & "      ,CADMAQ.SGI_DESCRI As SGI_DESCRIMAQ" & vbCrLf
    sSql = sSql & "      ,TURNO.SGI_DESCRI As SGI_DESCRITURN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPPRODUCAO" & strNOMTABELA & " CABEC" & vbCrLf
    sSql = sSql & "      ,SGI_CADMAQUINA  CADMAQ" & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN TURNO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CABEC.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CADMAQ.SGI_FILIAL = CABEC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CADMAQ.SGI_CODIGO = CABEC.SGI_CODMAQ" & vbCrLf
    sSql = sSql & "   And TURNO.SGI_FILIAL = CABEC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And TURNO.SGI_CODIGO = CABEC.SGI_CODTUN" & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
       sSql = sSql & "   And CABEC.SGI_CODIGO = " & txtCampos.Text & vbCrLf
       sSql = sSql & "Order By CABEC.SGI_CODIGO"
    ElseIf cboFiltro.ListIndex = 1 Then
       sSql = sSql & "   And CABEC.SGI_DTLANC = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
       sSql = sSql & "Order By CABEC.SGI_DTLANC"
    ElseIf cboFiltro.ListIndex = 2 Then
       sSql = sSql & "   And CABEC.SGI_CODMAQ = " & txtCampos.Text & vbCrLf
       sSql = sSql & "Order By CABEC.SGI_CODMAQ"
    ElseIf cboFiltro.ListIndex = 3 Then
       sSql = sSql & "   And CADMAQ.SGI_DESCRI Like '%" & txtCampos.Text & "%'" & vbCrLf
       sSql = sSql & "Order By CADMAQ.SGI_DESCRIMAQ"
    ElseIf cboFiltro.ListIndex = 4 Then
       sSql = sSql & "   And CABEC.SGI_CODTUN = " & txtCampos.Text & vbCrLf
       sSql = sSql & "Order By CABEC.SGI_CODTUN"
    ElseIf cboFiltro.ListIndex = 5 Then
       sSql = sSql & "   And TURNO.SGI_DESCRI Like '%" & txtCampos.Text & "%'" & vbCrLf
       sSql = sSql & "Order By TURNO.SGI_DESCRITURN"
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
    
        Call ConfGrid

        Do While Not BREC.EOF()
            
            strCampos = BREC!SGI_CODIGO & vbTab & _
                        Format(BREC!SGI_DTLANC, "DD/MM/YYYY") & vbTab & _
                        BREC!SGI_CODMAQ & vbTab & _
                        BREC!SGI_DESCRIMAQ & vbTab & _
                        BREC!SGI_CODTUN & vbTab & _
                        BREC!SGI_DESCRITURN
           
            grdAPONT.AddItem strCampos
           
            BREC.MoveNext
        Loop
        
    Else
        MsgBox "Este Registro não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

End Sub
