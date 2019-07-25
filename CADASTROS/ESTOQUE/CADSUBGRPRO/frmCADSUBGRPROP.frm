VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADSUBGRPROP 
   Caption         =   "Cadastro de sub-familia de produtos"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxSUBGRPROD 
         Height          =   3135
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5530
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   7935
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3600
         Top             =   240
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
         Picture         =   "frmCADSUBGRPROP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmCADSUBGRPROP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmCADSUBGRPROP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmCADSUBGRPROP.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Altera Empresa "
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
         Picture         =   "frmCADSUBGRPROP.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   120
         Width           =   735
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
         Picture         =   "frmCADSUBGRPROP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   4455
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   200
         Width           =   1695
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
Attribute VB_Name = "frmCADSUBGRPROP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Dim objFuncoes      As Object
Dim objCADSUBGRPROD As Object
Dim iCodigo         As Integer

Private Sub cboFiltro_Validate(Cancel As Boolean)

    txtCampos.Text = ""
    txtCampos.SetFocus
    
    ConfGrid
    PreencheGrid

End Sub

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
  
  ''If Verifica_Integracao = False Then Exit Sub
  
  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADSUBGRPROD.GRAVA("E") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxSUBGRPROD.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADSUBGRPROD = Nothing
    Unload Me
End Sub

Private Sub flxSUBGRPROD_Click()
    If flxSUBGRPROD.Rows > 1 Then objCADSUBGRPROD.SUBGRPRODCOD = CInt(flxSUBGRPROD.TextMatrix(flxSUBGRPROD.RowSel, 1))
End Sub

Private Sub flxSUBGRPROD_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxSUBGRPROD.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxSUBGRPROD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
       If flxSUBGRPROD.Rows > 1 Then Operacao "C"
    End If
End Sub

Private Sub flxSUBGRPROD_RowColChange()
    If flxSUBGRPROD.Rows > 1 Then objCADSUBGRPROD.SUBGRPRODCOD = CInt(flxSUBGRPROD.TextMatrix(flxSUBGRPROD.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADSUBGRPROD = CreateObject("CADSUBGRPRO.clsCADSUBGRPRO")
    
    objCADSUBGRPROD.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADSUBGRPROP
    
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
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()

    If objCADSUBGRPROD.Pesq_CadSubGruProd = False Then
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
    
    flxSUBGRPROD.Rows = 1
    flxSUBGRPROD.Cols = 3
    
    flxSUBGRPROD.TextMatrix(0, 1) = "Código"
    flxSUBGRPROD.TextMatrix(0, 2) = "Descrição"
    
    flxSUBGRPROD.ColWidth(0) = 0
    flxSUBGRPROD.ColWidth(1) = 700
    flxSUBGRPROD.ColWidth(2) = 5000
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SUB.SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,SUB.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADSUBGRPROD SUB" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SUB.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SUB.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxSUBGRPROD.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxSUBGRPROD.Rows - 1)
             
            If flxSUBGRPROD.TextMatrix(I, 1) = iCodigo Then
               flxSUBGRPROD.Row = I
               flxSUBGRPROD.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
  
  If flxSUBGRPROD.Rows > 1 Then iCodigo = CInt(flxSUBGRPROD.TextMatrix(flxSUBGRPROD.RowSel, 1))
  
  frmCADSUBGRPRO.cCaminho = cCaminho
  frmCADSUBGRPRO.Linha = Linha
  frmCADSUBGRPRO.iCodigo = iCodigo
  frmCADSUBGRPRO.cTipOper = strOperacao
  frmCADSUBGRPRO.FILIAL = FILIAL
  frmCADSUBGRPRO.strAcesso = strAcesso
  frmCADSUBGRPRO.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     
     sSql = " Select "
     sSql = sSql & "        SGI_CODIGO    " & vbCrLf
     sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADSUBGRPROD " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
     
  ElseIf cboFiltro.ListIndex = 1 Then
  
     sSql = " Select "
     sSql = sSql & "        SGI_CODIGO    " & vbCrLf
     sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADSUBGRPROD " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL         = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_DESCRICAO "
  
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flxSUBGRPROD.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub
Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADSUBGRPROP
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
       sSql = sSql & "      SGI_CADSUBGRPROD " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxSUBGRPROD.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxSUBGRPROD.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADSUBGRPROD " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL
       sSql = sSql & "  And SGI_DESCRICAO LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxSUBGRPROD.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxSUBGRPROD.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    
    ConfGrid
    PreencheGrid

End Sub

Private Function Verifica_Integracao() As Boolean

     Verifica_Integracao = False
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  from " & vbCrLf
     sSql = sSql & "       SGI_SUBGRUPRODITEN " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_CODESPECIE = " & objCADSUBGRPROD.SUBGRPRODCOD
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     If Not BREC.EOF Then
        MsgBox "Há grupos de produtos com sub grupo cadastrado impossivel Excluir !!!", vbInformation + vbOKOnly, "Aviso"
        BREC.Close
        Exit Function
     End If

     Verifica_Integracao = True
     BREC.Close
     
End Function

