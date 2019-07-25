VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADFLUXOPRODP 
   Caption         =   "Cadastro de Fluxo de Produção"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11550
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   5535
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   11535
      Begin VSFlex8LCtl.VSFlexGrid grdFluxoProd 
         Height          =   5175
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   11295
         _cx             =   19923
         _cy             =   9128
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
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11535
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   200
         Width           =   1695
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "txtCampos"
         Top             =   200
         Width           =   8055
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
         TabIndex        =   11
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
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11535
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3720
         Top             =   360
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
         Picture         =   "frmCADFLUXOPRODP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   240
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
         Picture         =   "frmCADFLUXOPRODP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Inclui uma nova empresa"
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
         Left            =   1800
         Picture         =   "frmCADFLUXOPRODP.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Altera Empresa "
         Top             =   240
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
         Picture         =   "frmCADFLUXOPRODP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
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
         Left            =   9840
         Picture         =   "frmCADFLUXOPRODP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
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
         Left            =   10680
         Picture         =   "frmCADFLUXOPRODP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADFLUXOPRODP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho             As String
Public Linha                As Variant
Public FILIAL               As Integer
Public strAcesso            As String
Public strUSUARIO           As String
Dim objFuncoes              As Object
Dim objCADFLUXOPROD         As Object
Dim iCodigo                 As Integer

Const conCOL_SonCadFlx_Codigo                     As Integer = 0
Const conCOL_SonCadFlx_Produto                    As Integer = 1
Const conCOL_SonCadFlx_DescrProd                  As Integer = 2
Const conCOL_SonCadFlx_FormatString               As String = "=Cód.Fluxo|linha|Descrição"
Const conColumnsIn_SonFlx                         As Integer = 3

Private Sub cboFiltro_Change()
    txtCampos.SetFocus
    Call InitGridFlxProd
    Call PreencheGrid
End Sub

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    Call InitGridFlxProd
    Call PreencheGrid
End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub

  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  objCADFLUXOPROD.CODPROD = grdFluxoProd.Cell(flexcpText, grdFluxoProd.Row, conCOL_SonCadFlx_Produto)
  
  If objCADFLUXOPROD.GRAVA("E") = False Then Exit Sub
  If objCADFLUXOPROD.Atualiza("E", Str(objCADFLUXOPROD.CODIGO), FILIAL, "frmCADFLUXOPROD") = False Then Exit Sub
  
  MsgBox "Registro Excluso com Sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call Atualiza_Grid
  Call AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If grdFluxoProd.Rows > 1 Then Ordem
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADFLUXOPROD = Nothing
    Unload Me
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADFLUXOPROD = CreateObject("CADFLUXOPROD.clsCADFLUXOPROD")
    
    objCADFLUXOPROD.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADFLUXOPRODP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call AbilitaCampos
    Call InitGridFlxProd
    Call PreencheGrid
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Produto"
    cboFiltro.AddItem "Descrição"
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()
    If objCADFLUXOPROD.Pesq_CadFlxProd = False Then
       cmdAltera.Enabled = False
       cmdExclui.Enabled = False
       Frame1.Enabled = False
       Frame3.Enabled = False
    Else
       cmdAltera.Enabled = True
       cmdExclui.Enabled = True
       Frame1.Enabled = True
       Frame3.Enabled = True
    End If
End Sub

Private Sub InitGridFlxProd()

    With grdFluxoProd
    
       .Cols = conColumnsIn_SonFlx
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonCadFlx_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCadFlx_Codigo) = ""
       .ColDataType(conCOL_SonCadFlx_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCadFlx_Produto) = ""
       .ColDataType(conCOL_SonCadFlx_Produto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCadFlx_DescrProd) = ""
       .ColDataType(conCOL_SonCadFlx_DescrProd) = flexDTString
       
       .ColWidth(conCOL_SonCadFlx_Codigo) = 1000
       .ColWidth(conCOL_SonCadFlx_Produto) = 1000
       .ColWidth(conCOL_SonCadFlx_DescrProd) = 4000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       HEAD.*          " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRI " & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFLUXPROD     HEAD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO PROD" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       HEAD.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL = HEAD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODIGO = HEAD.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & " Order by HEAD.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       grdFluxoProd.AddItem BREC!SGI_CODIGO & vbTab & _
                            Trim(BREC!SGI_CODPROD) & vbTab & _
                            Trim(BREC!SGI_DESCRI)
       BREC.MoveNext
    Loop
    
    BREC.Close
    
       
End Sub

Private Sub grdFluxoProd_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonCadFlx_Codigo, conCOL_SonCadFlx_Produto, conCOL_SonCadFlx_DescrProd
    Cancel = True
    Case Else
        grdFluxoProd.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdFluxoProd_Click()
    If (grdFluxoProd.Rows - 1) > 0 And (grdFluxoProd.Row) > 0 Then objCADFLUXOPROD.CODIGO = CLng(grdFluxoProd.Cell(flexcpText, grdFluxoProd.RowSel, conCOL_SonCadFlx_Codigo))
End Sub

Private Sub grdFluxoProd_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If (grdFluxoProd.Rows - 1) > 0 Then Call Operacao("C")
End Sub

Private Sub grdFluxoProd_RowColChange()
    If (grdFluxoProd.Rows - 1) > 0 And (grdFluxoProd.Row) > 0 Then objCADFLUXOPROD.CODIGO = CLng(grdFluxoProd.Cell(flexcpText, grdFluxoProd.RowSel, conCOL_SonCadFlx_Codigo))
End Sub

Private Sub Timer1_Timer()
  Call Atualiza_Grid
  Call AbilitaCampos
End Sub

Private Sub Atualiza_Grid()
    
     Dim I        As Integer
     Dim bolAchou As Boolean
      
     bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADFLUXOPROD'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (grdFluxoProd.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If grdFluxoProd.Cell(flexcpText, I, conCOL_SonCadFlx_Codigo) = Trim(BREC!SGI_CODIGO) Then
                  If grdFluxoProd.Rows = 2 Then grdFluxoProd.Rows = 1
                  If grdFluxoProd.Rows > 2 Then grdFluxoProd.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(grdFluxoProd.Cell(flexcpText, I, conCOL_SonCadFlx_Codigo)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        sSql = "Select " & vbCrLf
        sSql = sSql & "       HEAD.* " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_DESCRI " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADFLUXPROD HEAD" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO  PROD" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       HEAD.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And HEAD.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL = HEAD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_CODIGO = HEAD.SGI_IDPRODUTO " & vbCrLf
        
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              grdFluxoProd.AddItem BREC2!SGI_CODIGO & vbTab & _
                                   Trim(BREC2!SGI_CODPROD) & vbTab & _
                                   Trim(BREC2!SGI_DESCRI)
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              grdFluxoProd.Cell(flexcpText, I, conCOL_SonCadFlx_Codigo) = BREC2!SGI_CODIGO
              grdFluxoProd.Cell(flexcpText, I, conCOL_SonCadFlx_Produto) = Trim(BREC2!SGI_CODPROD)
              grdFluxoProd.Cell(flexcpText, I, conCOL_SonCadFlx_DescrProd) = Trim(BREC2!SGI_DESCRI)
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
      
End Sub

Private Sub Operacao(Operacao As String)
 
  Dim Pesquisa As String
  
  If (grdFluxoProd.Row) > 0 Then iCodigo = CLng(grdFluxoProd.Cell(flexcpText, grdFluxoProd.Row, conCOL_SonCadFlx_Codigo))
  
  frmCADFLUXOPROD.cCaminho = cCaminho
  frmCADFLUXOPROD.Linha = Linha
  frmCADFLUXOPROD.iCodigo = iCodigo
  frmCADFLUXOPROD.cTipOper = Operacao
  frmCADFLUXOPROD.FILIAL = FILIAL
  frmCADFLUXOPROD.strAcesso = strAcesso
  frmCADFLUXOPROD.Show vbModal
  
  Call Atualiza_Grid
  Call AbilitaCampos

End Sub

Private Sub Ordem()

  InitGridFlxProd
  
  txtCampos.Text = ""
  
  sSql = ""
  
  sSql = "Select " & vbCrLf
  sSql = sSql & "       HEAD.* " & vbCrLf
  sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
  sSql = sSql & "  From " & vbCrLf
  sSql = sSql & "       SGI_CADFLUXPROD HEAD" & vbCrLf
  sSql = sSql & "      ,SGI_CADPRODUTO  PROD" & vbCrLf
  sSql = sSql & " Where " & vbCrLf
  sSql = sSql & "       HEAD.SGI_FILIAL = " & FILIAL & vbCrLf
  sSql = sSql & "   And PROD.SGI_FILIAL = HEAD.SGI_FILIAL " & vbCrLf
  sSql = sSql & "   And PROD.SGI_CODIGO = HEAD.SGI_CODPROD " & vbCrLf
  
  If cboFiltro.ListIndex = 0 Then
     sSql = sSql & " Order by HEAD.SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = sSql & " Order by HEAD.SGI_CODPROD "
  ElseIf cboFiltro.ListIndex = 2 Then
     sSql = sSql & " Order by PROD.SGI_DESCRICAO "
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     grdFluxoProd.AddItem BREC!SGI_CODIGO & vbTab & _
                          BREC!SGI_CODPROD & vbTab & _
                          BREC!SGI_DESCRICAO
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub
Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADFLUXOPRODP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       HEAD.* " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFLUXPROD HEAD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  PROD" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       HEAD.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL = HEAD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODIGO = HEAD.SGI_CODPROD " & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
       sSql = sSql & "   And HEAD.SGI_CODIGO = " & txtCampos.Text
         
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = sSql & "   And HEAD.SGI_CODPROD  = '" & Trim(txtCampos.Text) & "'"
    
    ElseIf cboFiltro.ListIndex = 3 Then
       sSql = sSql & "   And PROD.SGI_DESCRICAO Like '" & txtCampos.Text & "%'"
    End If

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
         
       InitGridFlxProd
           
       Do While Not BREC.EOF
          grdFluxoProd.AddItem BREC!SGI_CODIGO & vbTab & _
                               Trim(BREC!SGI_CODPROD) & vbTab & _
                               Trim(BREC!SGI_DESCRICAO)
          BREC.MoveNext
       Loop
           
       BREC.Close
       grdFluxoProd.SetFocus
       Exit Sub
       
    End If
    BREC.Close
    
    InitGridFlxProd
    PreencheGrid

End Sub
