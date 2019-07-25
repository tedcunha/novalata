VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADTIPOOPERACAOP 
   Caption         =   "Cadastro de Tipos de Operação"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   10215
      Begin VSFlex8LCtl.VSFlexGrid grdTIPOOPERACAO 
         Height          =   4695
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9975
         _cx             =   17595
         _cy             =   8281
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
      Width           =   10215
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   200
         Width           =   2535
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "txtCampos"
         Top             =   200
         Width           =   5655
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
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10215
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3720
         Top             =   240
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
         Picture         =   "frmCADTIPOOPERACAOP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
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
         Picture         =   "frmCADTIPOOPERACAOP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Inclui uma nova empresa"
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
         Picture         =   "frmCADTIPOOPERACAOP.frx":0A64
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
         Picture         =   "frmCADTIPOOPERACAOP.frx":0B66
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
         Left            =   8400
         Picture         =   "frmCADTIPOOPERACAOP.frx":0C68
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
         Left            =   9240
         Picture         =   "frmCADTIPOOPERACAOP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADTIPOOPERACAOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public Filial           As Integer
Public strAcesso        As String
Public strUSUARIO       As String
Dim objFuncoes          As Object
Dim objCADTIPOPER       As Object
Dim iCodigo             As Integer

Const conCOL_SonTipOper_Codigo                    As Integer = 0
Const conCOL_SonTipOper_Descricao                 As Integer = 1
Const conCOL_SonTipOper_FormatString              As String = "=Código|Descrição"
Const conColumnsIn_SonTipOper                     As Integer = 2


Private Sub cboFiltro_Change()
    txtCampos.Text = ""
    txtCampos.SetFocus
    Call ConfGrdTipOper
    Call PreencheGrid
End Sub

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    Call ConfGrdTipOper
    Call PreencheGrid
End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub

  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
    
  If objCADTIPOPER.GRAVA("E") = False Then Exit Sub
  If objCADTIPOPER.Atualiza("E", Str(objCADTIPOPER.CODIGO), Filial, "frmCADTIPOOPERACAO") = False Then Exit Sub
  
  MsgBox "Registro Excluso com Sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call Atualiza_Grid
  Call AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    If (grdTIPOOPERACAO.Rows - 1) > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objCADTIPOPER = Nothing
    Set objFuncoes = Nothing
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    
    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADTIPOPER = CreateObject("CADTIPOOPERACAO.clsCADTIPOOPERACAO")
        
    objCADTIPOPER.Filial = Filial
    
    objFuncoes.LimpaCampos frmCADTIPOOPERACAOP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call AbilitaCampos
    Call ConfGrdTipOper
    Call PreencheGrid
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Descrição"
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub ConfGrdTipOper()

    With grdTIPOOPERACAO
    
       .Cols = conColumnsIn_SonTipOper
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonTipOper_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonTipOper_Codigo) = ""
       .ColDataType(conCOL_SonTipOper_Codigo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonTipOper_Descricao) = ""
       .ColDataType(conCOL_SonTipOper_Descricao) = flexDTString
       
       .ColWidth(conCOL_SonTipOper_Codigo) = 1000
       .ColWidth(conCOL_SonTipOper_Descricao) = 5000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub AbilitaCampos()
    
    If objCADTIPOPER.Pesq_CadTipOper = False Then
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


Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_TIPOPERACAO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       grdTIPOOPERACAO.AddItem BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRI
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub

Private Sub Operacao(Operacao As String)
 
  Dim Pesquisa As String
  
  If (grdTIPOOPERACAO.Row) > 0 Then iCodigo = CLng(grdTIPOOPERACAO.Cell(flexcpText, grdTIPOOPERACAO.Row, conCOL_SonTipOper_Codigo))
  
  frmCADTIPOOPERACAO.cCaminho = cCaminho
  frmCADTIPOOPERACAO.Linha = Linha
  frmCADTIPOOPERACAO.iCodigo = iCodigo
  frmCADTIPOOPERACAO.cTipOper = Operacao
  frmCADTIPOOPERACAO.Filial = Filial
  frmCADTIPOOPERACAO.strAcesso = strAcesso
  frmCADTIPOOPERACAO.Show vbModal
  
  Call Atualiza_Grid
  Call AbilitaCampos

End Sub

Private Sub grdTIPOOPERACAO_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonTipOper_Codigo, conCOL_SonTipOper_Descricao
         Cancel = True
    Case Else
        grdTIPOOPERACAO.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdTIPOOPERACAO_Click()
    If (grdTIPOOPERACAO.Rows - 1) > 0 Then objCADTIPOPER.CODIGO = CLng(grdTIPOOPERACAO.Cell(flexcpText, grdTIPOOPERACAO.RowSel, conCOL_SonTipOper_Codigo))
End Sub

Private Sub grdTIPOOPERACAO_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If (grdTIPOOPERACAO.Rows - 1) > 0 Then Call Operacao("C")
End Sub

Private Sub grdTIPOOPERACAO_RowColChange()
    If (grdTIPOOPERACAO.Rows - 1) > 0 Then objCADTIPOPER.CODIGO = CLng(grdTIPOOPERACAO.Cell(flexcpText, grdTIPOOPERACAO.RowSel, conCOL_SonTipOper_Codigo))
End Sub

Private Sub Timer1_Timer()
  Call Atualiza_Grid
  Call AbilitaCampos
End Sub

Private Sub Ordem()

  Call ConfGrdTipOper
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_TIPOPERACAO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_TIPOPERACAO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
     sSql = sSql & " Order by SGI_DESCRI "
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     grdTIPOOPERACAO.AddItem BREC!SGI_CODIGO & vbTab & _
                             BREC!SGI_DESCRI
     BREC.MoveNext
  Loop
  BREC.Close

End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADTIPOOPERACAO
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    If cboFiltro.ListIndex = 0 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
             
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_TIPOPERACAO " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & Filial & vbCrLf
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrdTipOper
              
          Do While Not BREC.EOF
             grdTIPOOPERACAO.AddItem BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRI
             BREC.MoveNext
          Loop
              
          BREC.Close
          grdTIPOOPERACAO.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_TIPOPERACAO " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL  =  " & Filial & vbCrLf
       sSql = sSql & "  And SGI_DESCRI LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          Call ConfGrdTipOper
              
          Do While Not BREC.EOF
             grdTIPOOPERACAO.AddItem BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRI
             BREC.MoveNext
          Loop
              
          BREC.Close
          grdTIPOOPERACAO.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    
    Call ConfGrdTipOper
    Call PreencheGrid

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
     sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADTIPOOPERACAO'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (grdTIPOOPERACAO.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If grdTIPOOPERACAO.Cell(flexcpText, I, conCOL_SonTipOper_Codigo) = Trim(BREC!SGI_CODIGO) Then
                  If grdTIPOOPERACAO.Rows = 2 Then grdTIPOOPERACAO.Rows = 1
                  If grdTIPOOPERACAO.Rows > 2 Then grdTIPOOPERACAO.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(grdTIPOOPERACAO.Cell(flexcpText, I, conCOL_SonTipOper_Codigo)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              grdTIPOOPERACAO.AddItem BREC2!SGI_CODIGO & vbTab & _
                                      BREC2!SGI_DESCRI
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              grdTIPOOPERACAO.Cell(flexcpText, I, conCOL_SonTipOper_Codigo) = BREC2!SGI_CODIGO
              grdTIPOOPERACAO.Cell(flexcpText, I, conCOL_SonTipOper_Descricao) = BREC2!SGI_DESCRI
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
      
End Sub


