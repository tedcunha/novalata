VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADTIPOSLIQP 
   Caption         =   "Cadastro de Motivos de Liquida��o de Pedidos"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10260
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   5175
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   10215
      Begin VSFlex8LCtl.VSFlexGrid grdMotLiq 
         Height          =   4815
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9975
         _cx             =   17595
         _cy             =   8493
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
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   10215
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
         Left            =   9240
         Picture         =   "frmCADTIPOSLIQ.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   8400
         Picture         =   "frmCADTIPOSLIQ.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmCADTIPOSLIQ.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmCADTIPOSLIQ.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Altera Empresa "
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
         Picture         =   "frmCADTIPOSLIQ.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   120
         Width           =   855
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
         Picture         =   "frmCADTIPOSLIQ.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3720
         Top             =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   5655
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   200
         Width           =   2535
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
         Left            =   3720
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
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCADTIPOSLIQP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public lngIDUSUARIO     As Long
Public strUsuario       As String

Dim objFUNCOES              As Object
Dim objCADTIPOSLIQP         As Object
Dim iCodigo                 As Integer

Const conCOL_Mov_Codigo                          As Integer = 0
Const conCOL_Mov_Descri                          As Integer = 1
Const conCOL_Mov_FormatString                    As String = "=C�digo|Descri��o"
Const conColumnsIn_Mov                           As Integer = 2


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub cmdAltera_Click()
    If objFUNCOES.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    Call ConfGrid
End Sub

Private Sub cmdExclui_Click()

  If objFUNCOES.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclus�o do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADTIPOSLIQP.GRAVA("E") = False Then Exit Sub
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call AbilitaCampos
  Call ConfGrid

End Sub

Private Sub cmdInclui_Click()
    If objFUNCOES.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set objFUNCOES = CreateObject("BLBCWS.clsFuncoes")
    Set objCADTIPOSLIQP = CreateObject("CADTIPOSLIQ.clsCADTIPOSLIQ")
    
    objFUNCOES.LimpaCampos Me
    
    objCADTIPOSLIQP.FILIAL = FILIAL
    
    Set adoBanco_Dados = objFUNCOES.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "N�o foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call AbilitaCampos
    Call ConfGrid
    
    cboFiltro.AddItem "C�digo"
    cboFiltro.AddItem "Descri��o"
    cboFiltro.ListIndex = 0

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)


End Sub

Private Sub Destroy_Objeto()
    Set objFUNCOES = Nothing
    Set objCADTIPOSLIQP = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub


Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADTIPOSLIQP.AtivoDesativo
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdCanFiltro.Enabled = boolAtivoDesativo
    cmdOrden.Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo
    Frame3.Enabled = boolAtivoDesativo

End Sub

Private Sub ConfGrid()

    With grdMotLiq
    
       .Cols = conColumnsIn_Mov
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Mov_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_Mov_Codigo) = ""
       .ColDataType(conCOL_Mov_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Mov_Descri) = ""
       .ColDataType(conCOL_Mov_Descri) = flexDTString
       
       .ColWidth(conCOL_Mov_Codigo) = 1000
       .ColWidth(conCOL_Mov_Descri) = 5000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub grdMotLiq_Click()
    With grdMotLiq
        If .Row = 0 Then Exit Sub
        If (.Rows - 1) > 0 Then objCADTIPOSLIQP.CODIGO = CInt(.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
    End With
End Sub

Private Sub grdMotLiq_DblClick()
    If objFUNCOES.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    With grdMotLiq
        If .Row = 0 Then Exit Sub
        If (.Rows - 1) > 0 Then Call Operacao("C")
    End With
End Sub

Private Sub grdMotLiq_RowColChange()
    With grdMotLiq
        If .Row = 0 Then Exit Sub
        If (.Rows - 1) > 0 Then objCADTIPOSLIQP.CODIGO = CInt(.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
    End With
End Sub

Private Sub Ordem()

    Call ConfGrid
  
    txtCampos.Text = ""
  
    sSql = ""
    sSql = " Select " & vbCrLf
    sSql = sSql & "        * " & vbCrLf
    sSql = sSql & "   from " & vbCrLf
    sSql = sSql & "        SGI_CADMOTLIQ " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
  
    If cboFiltro.ListIndex = 0 Then
       sSql = sSql & " Order by SGI_CODIGO "
    ElseIf cboFiltro.ListIndex = 1 Then
       sSql = sSql & " Order by SGI_DESCRI "
    End If
  
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       grdMotLiq.AddItem BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_DESCRI
       BREC.MoveNext
    Loop
  
    BREC.Close

End Sub

Private Sub txtCampos_Change()
    objFUNCOES.SelecionaCampos txtCampos.Name, Me
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFUNCOES.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim lngCodSep   As Long
    Dim strCampos   As String
    Dim strDEBCRED  As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    If cboFiltro.ListIndex = 0 Then
        If IsNumeric(txtCampos.Text) = False Then
           MsgBox "Somente � permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
           txtCampos.Text = ""
           txtCampos.SetFocus
           Exit Sub
        End If
    End If
        
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOTLIQ" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "   And SGI_CODIGO   = " & Trim(txtCampos.Text)
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "   And SGI_DESCRI LIKE '%" & Trim(txtCampos.Text) & "%'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
    
        Call ConfGrid

        Do While Not BREC.EOF()
            
            strCampos = BREC!SGI_CODIGO & vbTab & _
                        BREC!SGI_DESCRI
           
            grdMotLiq.AddItem strCampos
           
            BREC.MoveNext
        Loop
        
    Else
        MsgBox "Este Registro n�o Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

End Sub


Private Sub Operacao(strOperacao As String)
  
    With grdMotLiq
        If strOperacao <> "I" Then
            If (.Row = 0) Then Exit Sub
            If (.Rows - 1) > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_Mov_Codigo))
        End If
    End With
    
    frmCADTIPOSLIQ.cCaminho = cCaminho
    frmCADTIPOSLIQ.Linha = Linha
    frmCADTIPOSLIQ.iCodigo = iCodigo
    frmCADTIPOSLIQ.cTipOper = strOperacao
    frmCADTIPOSLIQ.FILIAL = FILIAL
    frmCADTIPOSLIQ.strAcesso = strAcesso
    frmCADTIPOSLIQ.lngCODUSUARIO = lngIDUSUARIO
    frmCADTIPOSLIQ.strUsuario = strUsuario
    frmCADTIPOSLIQ.Show vbModal
    
    Call AbilitaCampos

End Sub

