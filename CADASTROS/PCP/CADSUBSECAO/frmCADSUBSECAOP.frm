VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADSUBSECAOP 
   Caption         =   "Cadastro se Sub-Seção"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3855
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxSUBSECAO 
         Height          =   3615
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7935
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
         Width           =   4455
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
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7935
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
         Picture         =   "frmCADSUBSECAOP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   735
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
         Left            =   840
         Picture         =   "frmCADSUBSECAOP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   120
         Width           =   735
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
         Left            =   1560
         Picture         =   "frmCADSUBSECAOP.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Altera Empresa "
         Top             =   120
         Width           =   735
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
         Left            =   2280
         Picture         =   "frmCADSUBSECAOP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
         Width           =   735
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
         Left            =   6360
         Picture         =   "frmCADSUBSECAOP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
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
         Left            =   7080
         Picture         =   "frmCADSUBSECAOP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3600
         Top             =   240
      End
   End
End
Attribute VB_Name = "frmCADSUBSECAOP"
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
Dim objFuncoes      As Object
Dim objCADSUBSECAO  As Object
Dim iCodigo         As Integer

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Operacao "A"
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    ConfGrid
    PreencheGrid
End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  Dim iResp As Integer
  
  If Consistencia = False Then
     MsgBox "No cadastro de seção existe esta sub-seção cadastrada !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  End If
  
  iResp = MsgBox("Confirma a Exclusão do Registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADSUBSECAO.GRAVA("E") = False Then Exit Sub
  If objCADSUBSECAO.Atualiza("E", Str(objCADSUBSECAO.CODIGO), FILIAL, "frmCADSUBSECAO") = False Then Exit Sub
  
  MsgBox "Registro Excluso com Sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Atualiza_Grid
  AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxSUBSECAO.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADSUBSECAO = Nothing
    Unload Me
End Sub

Private Sub flxSUBSECAO_Click()
    If flxSUBSECAO.Rows > 1 Then objCADSUBSECAO.CODIGO = CInt(flxSUBSECAO.TextMatrix(flxSUBSECAO.RowSel, 1))
End Sub

Private Sub flxSUBSECAO_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxSUBSECAO.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxSUBSECAO_RowColChange()
    If flxSUBSECAO.Rows > 1 Then objCADSUBSECAO.CODIGO = CInt(flxSUBSECAO.TextMatrix(flxSUBSECAO.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADSUBSECAO = CreateObject("CADSUBSECAO.clsCADSUBSECAO")
    
    objCADSUBSECAO.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADSUBSECAOP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Descrição"
    cboFiltro.AddItem "Sigla"
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()

    If objCADSUBSECAO.Carrega_CadSubSecao = False Then
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

Private Sub ConfGrid()
            
    flxSUBSECAO.Rows = 1
    flxSUBSECAO.Cols = 4
    
    flxSUBSECAO.TextMatrix(0, 0) = ""
    flxSUBSECAO.TextMatrix(0, 1) = "Código"
    flxSUBSECAO.TextMatrix(0, 2) = "Descrição"
    flxSUBSECAO.TextMatrix(0, 3) = "Sigla"
    
    flxSUBSECAO.ColWidth(0) = 0
    flxSUBSECAO.ColWidth(1) = 700
    flxSUBSECAO.ColWidth(2) = 5000
    flxSUBSECAO.ColWidth(3) = 2000
    
End Sub

Private Sub PreencheGrid()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxSUBSECAO.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           BREC!SGI_DESCRI & vbTab & _
                           BREC!SGI_SIGLA
       BREC.MoveNext
    Loop
    PosGrid
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxSUBSECAO.Rows - 1)
             
            If flxSUBSECAO.TextMatrix(I, 1) = iCodigo Then
               flxSUBSECAO.Row = I
               flxSUBSECAO.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub


Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
  
  If flxSUBSECAO.Rows > 1 Then iCodigo = CInt(flxSUBSECAO.TextMatrix(flxSUBSECAO.RowSel, 1))
  
  frmCADSUBSECAO.cCaminho = cCaminho
  frmCADSUBSECAO.Linha = Linha
  frmCADSUBSECAO.iCodigo = iCodigo
  frmCADSUBSECAO.cTipOper = strOperacao
  frmCADSUBSECAO.FILIAL = FILIAL
  frmCADSUBSECAO.strAcesso = strAcesso
  frmCADSUBSECAO.strMODPAI = Me.Name
  frmCADSUBSECAO.Show vbModal
  
  Atualiza_Grid
  AbilitaCampos

End Sub


Public Function Consistencia() As Boolean
    
    Consistencia = True
    
    '' Sub-Seção dacadastra na Seção
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADITESEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODSUBSECAO = " & objCADSUBSECAO.CODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then Consistencia = False
    BREC.Close
    
End Function

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  sSql = ""
  
  sSql = " Select " & vbCrLf
  sSql = sSql & "        * " & vbCrLf
  sSql = sSql & "   from " & vbCrLf
  sSql = sSql & "        SGI_CADSUBSECAO " & vbCrLf
  sSql = sSql & " Where " & vbCrLf
  sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
  
  If cboFiltro.ListIndex = 0 Then
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = sSql & " Order by SGI_DESCRI "
  ElseIf cboFiltro.ListIndex = 2 Then
     sSql = sSql & " Order by SGI_SIGLA "
  End If
  
  BREC.Open sSql, adoBanco_Dados
  Do While Not BREC.EOF
     flxSUBSECAO.AddItem "" & vbTab & _
                         BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_DESCRI & vbTab & _
                         BREC!SGI_SIGLA
     BREC.MoveNext
  Loop
  BREC.Close

End Sub

Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADSUBSECAOP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    Call ConfGrid
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADSUBSECAO " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Call cmdCanFiltro_Click
          Exit Sub
       End If
             
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
    ElseIf cboFiltro.ListIndex = 1 Then
       sSql = sSql & "  And SGI_DESCRI LIKE '" & txtCampos.Text & "%'"
    ElseIf cboFiltro.ListIndex = 2 Then
       sSql = sSql & "  And SGI_SIGLA LIKE '" & txtCampos.Text & "%'"
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       flxSUBSECAO.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           BREC!SGI_DESCRI & vbTab & _
                           BREC!SGI_SIGLA
       BREC.MoveNext
    Loop
    BREC.Close
    flxSUBSECAO.SetFocus
    
    Exit Sub

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
     sSql = sSql & "   And SGI_MODULO = 'frmCADSUBSECAO'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (flxSUBSECAO.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If flxSUBSECAO.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                  If flxSUBSECAO.Rows = 2 Then flxSUBSECAO.Rows = 1
                  If flxSUBSECAO.Rows > 2 Then flxSUBSECAO.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(flxSUBSECAO.TextMatrix(I, 1)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxSUBSECAO.AddItem "" & vbTab & _
                                   BREC2!SGI_CODIGO & vbTab & _
                                   BREC2!SGI_DESCRI & vbTab & _
                                   BREC2!SGI_SIGLA
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxSUBSECAO.TextMatrix(I, 0) = ""
              flxSUBSECAO.TextMatrix(I, 1) = BREC2!SGI_CODIGO
              flxSUBSECAO.TextMatrix(I, 2) = BREC2!SGI_DESCRI
              flxSUBSECAO.TextMatrix(I, 3) = BREC2!SGI_SIGLA
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
     
End Sub


