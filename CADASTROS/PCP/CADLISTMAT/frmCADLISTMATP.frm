VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADLISTMATP 
   Caption         =   "Cadastro de Estrutura de Produto"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   9855
      Begin VSFlex8LCtl.VSFlexGrid grdESTRPROD 
         Height          =   5895
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9615
         _cx             =   16960
         _cy             =   10398
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
      TabIndex        =   6
      Top             =   0
      Width           =   9855
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   200
         Width           =   1695
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "txtCampos"
         Top             =   200
         Width           =   6375
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9855
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
         Picture         =   "frmCADLISTMATP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmCADLISTMATP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Volta ao Menu Principal"
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
         Picture         =   "frmCADLISTMATP.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Altera Empresa "
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
         Picture         =   "frmCADLISTMATP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Empresa"
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
         Left            =   8040
         Picture         =   "frmCADLISTMATP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   855
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
         Left            =   8880
         Picture         =   "frmCADLISTMATP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   4920
         Top             =   120
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
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
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   7800
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
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
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7800
      Width           =   2895
   End
End
Attribute VB_Name = "frmCADLISTMATP"
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
Dim objCADLISTAMAT          As Object
Dim iCodigo                 As String

Const conCOL_SonCadLst_IDProduto                  As Integer = 0
Const conCOL_SonCadLst_Produto                    As Integer = 1
Const conCOL_SonCadLst_DescrProd                  As Integer = 2
Const conCOL_SonCadLst_TemArvore                  As Integer = 3
Const conCOL_SonCadLst_FormatString               As String = "=IDProduto|Produto|Descrição|Tem Arvore"
Const conColumnsIn_SonLst                         As Integer = 4

Private Sub cboFiltro_Change()
    txtCampos.SetFocus
    Call InitGridLstProd
    Call PreencheGrid
End Sub

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    Call InitGridLstProd
    Call PreencheGrid
End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub

  Dim iResp As Integer
  
  If VerifArvore(PegaPaiEstrutura(Str(objCADLISTAMAT.CODIGO))) = False Then
     MsgBox "A estrutura não foi criada !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  End If
  
  iResp = MsgBox("Confirma a exclusão da estrutura ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If (grdESTRPROD.Row) > 0 Then iCodigo = grdESTRPROD.Cell(flexcpText, grdESTRPROD.Row, conCOL_SonCadLst_IDProduto)
  objCADLISTAMAT.CODIGO = grdESTRPROD.Cell(flexcpText, grdESTRPROD.Row, conCOL_SonCadLst_IDProduto)
  
  Call frmCADARVPROD.CarregaLista(PegaPaiEstrutura(Str(iCodigo)), FILIAL)
  If objCADLISTAMAT.GRAVA("E") = False Then Exit Sub
  
  MsgBox "Estrutura exclusa com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call Atualiza_Grid
  Call AbilitaCampos

End Sub


Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADLISTAMAT = Nothing
    Unload Me
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADLISTAMAT = CreateObject("CADLISTMAT.clsCADLISTMAT")
    
    objCADLISTAMAT.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADLISTMATP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    InitGridLstProd
    ''PreencheGrid
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Descrição"
    
    cboFiltro.ListIndex = 0
    
''    Call AtualizaFlagTemArv

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    Label3.Caption = "Arvores já Prontas : " & QtdeArvoresProntas
    Label4.Caption = "Arvores Pendentes : " & QtdeArvoresPendentes

End Sub

Private Sub AbilitaCampos()
    If objCADLISTAMAT.Pesq_CadLstProd = False Then
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

Private Sub InitGridLstProd()

    With grdESTRPROD
    
       .Cols = conColumnsIn_SonLst
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonCadLst_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCadLst_IDProduto) = ""
       .ColDataType(conCOL_SonCadLst_IDProduto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCadLst_Produto) = ""
       .ColDataType(conCOL_SonCadLst_Produto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCadLst_DescrProd) = ""
       .ColDataType(conCOL_SonCadLst_DescrProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCadLst_TemArvore) = ""
       .ColDataType(conCOL_SonCadLst_TemArvore) = flexDTString
       
       .ColWidth(conCOL_SonCadLst_IDProduto) = 0
       .ColWidth(conCOL_SonCadLst_Produto) = 1200
       .ColWidth(conCOL_SonCadLst_DescrProd) = 5000
       .ColWidth(conCOL_SonCadLst_TemArvore) = 1200
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub PreencheGrid()

    Dim strTemArvore As String
    
    sSql = ""
    
    sSql = "Select SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & ",Case PRO.SGI_PRODUTOTIPO" & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' +" & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' +" & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' +" & vbCrLf
    sSql = sSql & "                        (Case" & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End)" & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGOACAB" & vbCrLf
    sSql = sSql & "      ,PRO.*" & vbCrLf
    sSql = sSql & "      ,TIP.SGI_DESCRICAO as DESC_TIPO" & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO as DESC_ESP" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD ESP" & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPPROD TIP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PRO.SGI_CODESPECIE" & vbCrLf
    sSql = sSql & "   And TIP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And TIP.SGI_CODIGO = PRO.SGI_CODTIPO" & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
    
        grdESTRPROD.AddItem Trim(Str(BREC!SGI_IDPRODUTO)) & vbTab & _
                            Trim(BREC!SGI_CODIGOACAB) & vbTab & _
                            Trim(BREC!SGI_DESCRICAO) & vbTab & _
                            IIf(Trim(BREC!SGI_TEMARVORE) = "S", "Sim", "Não")
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub

Private Sub grdFluxoProd_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonCadLst_Produto, conCOL_SonCadLst_DescrProd
    Cancel = True
    Case Else
        grdESTRPROD.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdESTRPROD_Click()
    If (grdESTRPROD.Rows - 1) > 0 And (grdESTRPROD.Row) > 0 Then objCADLISTAMAT.CODIGO = Trim(grdESTRPROD.Cell(flexcpText, grdESTRPROD.RowSel, conCOL_SonCadLst_IDProduto))
End Sub

Private Sub grdESTRPROD_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If (grdESTRPROD.Rows - 1) > 0 Then Call Operacao("C")
End Sub

Private Sub grdESTRPROD_RowColChange()
    If (grdESTRPROD.Rows - 1) > 0 And (grdESTRPROD.Row) > 0 Then objCADLISTAMAT.CODIGO = Trim(grdESTRPROD.Cell(flexcpText, grdESTRPROD.RowSel, conCOL_SonCadLst_IDProduto))
End Sub

Private Sub Timer1_Timer()
  Call Atualiza_Grid
  Call AbilitaCampos
End Sub

Private Sub Atualiza_Grid()
    
     Dim I              As Integer
     Dim bolAchou       As Boolean
     Dim strTemArvore   As String
      
     bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADARVPROD'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (grdESTRPROD.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If grdESTRPROD.Cell(flexcpText, I, conCOL_SonCadLst_IDProduto) = Trim(BREC!SGI_CODIGO) Then
                  If grdESTRPROD.Rows = 2 Then grdESTRPROD.Rows = 1
                  If grdESTRPROD.Rows > 2 Then grdESTRPROD.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(grdESTRPROD.Cell(flexcpText, I, conCOL_SonCadLst_IDProduto)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        sSql = "Select PRO.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & ",Case PRO.SGI_PRODUTOTIPO" & vbCrLf
        sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' +" & vbCrLf
        sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' +" & vbCrLf
        sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' +" & vbCrLf
        sSql = sSql & "                        (Case" & vbCrLf
        sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
        sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRO.SGI_DIGVERIF))) End)" & vbCrLf
        sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO" & vbCrLf
        sSql = sSql & ",PRO.SGI_DESCRICAO" & vbCrLf
        sSql = sSql & ",Case PRO.SGI_TEMARVORE When 'N' then 'Não'" & vbCrLf
        sSql = sSql & "                        When 'S' then 'Sim' End As SGI_TEMARVORE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "         SGI_CADPRODUTO PRO" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PRO.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And PRO.SGI_IDPRODUTO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
           
           If Not BREC2.EOF() Then
              
              grdESTRPROD.AddItem Trim(Str(BREC2!SGI_IDPRODUTO)) & vbTab & _
                                  Trim(BREC2!SGI_CODIGO) & vbTab & _
                                  Trim(BREC2!SGI_DESCRICAO) & vbTab & _
                                  Trim(BREC2!SGI_TEMARVORE)
           End If
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
           
           If Not BREC2.EOF() Then
              grdESTRPROD.Cell(flexcpText, I, conCOL_SonCadLst_IDProduto) = Trim(Str(BREC2!SGI_IDPRODUTO))
              grdESTRPROD.Cell(flexcpText, I, conCOL_SonCadLst_Produto) = Trim(BREC2!SGI_CODIGO)
              grdESTRPROD.Cell(flexcpText, I, conCOL_SonCadLst_DescrProd) = Trim(BREC2!SGI_DESCRICAO)
              grdESTRPROD.Cell(flexcpText, I, conCOL_SonCadLst_TemArvore) = Trim(BREC2!SGI_TEMARVORE)
           End If
        
        End If
        BREC2.Close
        
     End If
     BREC.Close
      
End Sub

Private Sub Operacao(Operacao As String)
 
  Dim Pesquisa As String
  
  If (grdESTRPROD.Row) = 0 Then
        MsgBox "Selecione um produto !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
  End If
  
  If (grdESTRPROD.Row) > 0 Then iCodigo = Trim(grdESTRPROD.Cell(flexcpText, grdESTRPROD.Row, conCOL_SonCadLst_IDProduto))
  
  If Operacao = "A" Or Operacao = "C" Then
     If VerifArvore(PegaPaiEstrutura(iCodigo)) = False Then
        MsgBox "A estrutura não foi criada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
  ElseIf Operacao = "AL" Then
     If VerifArvore(PegaPaiEstrutura(iCodigo)) = True Then
        MsgBox "A estrutura já foi alterada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
  ElseIf Operacao = "I" Then
     If VerifArvore(PegaPaiEstrutura(iCodigo)) = True Then
        MsgBox "A estrutura já foi criada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
  End If
  
  frmCADARVPROD.cCaminho = cCaminho
  frmCADARVPROD.Linha = Linha
  frmCADARVPROD.iCodigo = PegaPaiEstrutura(iCodigo)
  frmCADARVPROD.strCODPROD = iCodigo
  frmCADARVPROD.cTipOper = Operacao
  frmCADARVPROD.FILIAL = FILIAL
  frmCADARVPROD.strAcesso = strAcesso
  frmCADARVPROD.Show vbModal
  
  Call Atualiza_Grid
  Call AbilitaCampos

    Label3.Caption = "Arvores já Prontas : " & QtdeArvoresProntas

End Sub

Private Sub Ordem()

    Dim strTemArvore As String
    
    InitGridLstProd
  
    txtCampos.Text = ""
  
    sSql = ""
  
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,PRO.*" & vbCrLf
    sSql = sSql & "      ,PRO.SGI_CODIGO As SGI_CODIGOACAB" & vbCrLf
    sSql = sSql & "      ,TIP.SGI_DESCRICAO as DESC_TIPO" & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO as DESC_ESP" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD ESP" & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPPROD TIP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PRO.SGI_CODESPECIE" & vbCrLf
    sSql = sSql & "   And TIP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And TIP.SGI_CODIGO = PRO.SGI_CODTIPO" & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order By PRO.SGI_CODIGO "
    If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order By PRO.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       
       grdESTRPROD.AddItem Trim(Str(BREC!SGI_IDPRODUTO)) & vbTab & _
                           Trim(BREC!SGI_CODIGOACAB) & vbTab & _
                           Trim(BREC!SGI_DESCRICAO) & vbTab & _
                           IIf(Trim(BREC!SGI_TEMARVORE) = "S", "Sim", "Não")
       BREC.MoveNext
    Loop
  
    BREC.Close

End Sub
Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADLISTMATP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim strTemArvore As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & ",Case PRO.SGI_PRODUTOTIPO" & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' +" & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' +" & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' +" & vbCrLf
    sSql = sSql & "                        (Case" & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End)" & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGOACAB" & vbCrLf
    sSql = sSql & "      ,PRO.*" & vbCrLf
    sSql = sSql & "      ,TIP.SGI_DESCRICAO as DESC_TIPO" & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO as DESC_ESP" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD ESP" & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPPROD TIP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
    sSql = sSql & "   And ESP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PRO.SGI_CODESPECIE" & vbCrLf
    sSql = sSql & "   And TIP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And TIP.SGI_CODIGO = PRO.SGI_CODTIPO" & vbCrLf
    sSql = sSql & "   And  Case PRO.SGI_PRODUTOTIPO" & vbCrLf
    sSql = sSql & "        When 1 Then" & vbCrLf
    sSql = sSql & "                replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' +" & vbCrLf
    sSql = sSql & "                replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' +" & vbCrLf
    sSql = sSql & "                replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' +" & vbCrLf
    sSql = sSql & "        Case" & vbCrLf
    sSql = sSql & "              When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
    sSql = sSql & "              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End" & vbCrLf
    sSql = sSql & "              When 0 Then PRO.SGI_CODIGO End Like '" & Trim(txtCampos.Text) & "%'"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
         
       Call InitGridLstProd
           
       Do While Not BREC.EOF
            grdESTRPROD.AddItem Trim(BREC!SGI_IDPRODUTO) & vbTab & _
                                Trim(BREC!SGI_CODIGOACAB) & vbTab & _
                                Trim(BREC!SGI_DESCRICAO) & vbTab & _
                                IIf(Trim(BREC!SGI_TEMARVORE) = "S", "Sim", "Não")
            BREC.MoveNext
       Loop
           
    End If
    BREC.Close
    
End Sub

Private Function VerifArvore(lngPRODUTO As Long) As Boolean

    VerifArvore = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMATPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO   = " & lngPRODUTO & vbCrLf
    sSql = sSql & "   And SGI_FILIAL   = " & FILIAL
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF Then VerifArvore = True
    BREC4.Close

End Function


Private Function PegaPaiEstrutura(strPRODUTO As String) As Long

    PegaPaiEstrutura = -1
    
    If Len(Trim(strPRODUTO)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMATPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_IDPRODUTO  = " & Trim(strPRODUTO) & vbCrLf
    sSql = sSql & "   And SGI_IDPRODLST  = -1" & vbCrLf
    sSql = sSql & "   And SGI_FILIAL   = " & FILIAL
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF Then PegaPaiEstrutura = BREC4!SGI_CODIGO
    BREC4.Close

End Function


Private Sub AtualizaFlagTemArv()

    Dim ArrProdutos() As String
    Dim intQtdProdutos As Integer
    Dim I As Integer

    sSql = "Select" & vbCrLf
    sSql = sSql & "      Count(SGI_IDPRODUTO) as QtdProd" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       ReDim ArrProdutos(1 To BREC!QtdProd, 1 To 2) As String
    End If
    BREC.Close
    
    '' Preenchendo a Gride
    sSql = "Select" & vbCrLf
    sSql = sSql & "      SGI_TEMARVORE" & vbCrLf
    sSql = sSql & "     ,SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    intQtdProdutos = 1
    Do While Not BREC.EOF()
    
       ArrProdutos(intQtdProdutos, 1) = BREC!SGI_IDPRODUTO
       ArrProdutos(intQtdProdutos, 2) = BREC!SGI_TEMARVORE
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       *" & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_LISTAMATPROD" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
       sSql = sSql & "   And SGI_IDPRODLST  = -1" & vbCrLf
       sSql = sSql & "   And SGI_FILIAL     = " & FILIAL
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF() Then ArrProdutos(intQtdProdutos, 2) = "S"
       BREC2.Close
       
       intQtdProdutos = intQtdProdutos + 1
       BREC.MoveNext
    Loop
    BREC.Close
    
    
    '' --------------------------------------------
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    If IsArray(ArrProdutos) Then
    
        For I = 1 To UBound(ArrProdutos)
        
            '' Alterando o Produto
            sSql = "Update SGI_CADPRODUTO Set SGI_TEMARVORE = '" & Trim(ArrProdutos(I, 2)) & "'" & vbCrLf
            sSql = sSql & " Where SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_IDPRODUTO = " & ArrProdutos(I, 1)
        
            BGRV.CommandText = sSql
            BGRV.Execute
    
        Next I
        
    End If
    
    adoBanco_Dados.CommitTrans
    '' --------------------------------------------
    
End Sub

Private Function QtdeArvoresProntas() As Long

    QtdeArvoresProntas = 0
    
    sSql = "Select Count(SGI_TEMARVORE) As SGI_QTDE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_TEMARVORE   = 'S'"
    sSql = sSql & "   And SGI_PRODUTOTIPO = 1" & vbCrLf

    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC7.EOF() Then QtdeArvoresProntas = BREC7!SGI_QTDE
    BREC7.Close

End Function

Private Function QtdeArvoresPendentes() As Long

    QtdeArvoresPendentes = 0
    
    sSql = "Select Count(SGI_TEMARVORE) As SGI_QTDE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_TEMARVORE   = 'N'"
    sSql = sSql & "   And SGI_PRODUTOTIPO = 1" & vbCrLf

    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC7.EOF() Then QtdeArvoresPendentes = BREC7!SGI_QTDE
    BREC7.Close

End Function

