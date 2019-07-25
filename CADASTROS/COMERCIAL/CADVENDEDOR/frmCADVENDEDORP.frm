VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADVENDEDORP 
   Caption         =   "Cadastro de vendedores"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3855
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxCADVENDEDORES 
         Height          =   3495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6165
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
         Left            =   3720
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
         Left            =   6960
         Picture         =   "frmCADVENDEDORP.frx":0000
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
         Left            =   6120
         Picture         =   "frmCADVENDEDORP.frx":0102
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
         Picture         =   "frmCADVENDEDORP.frx":0634
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
         Picture         =   "frmCADVENDEDORP.frx":0736
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
         Picture         =   "frmCADVENDEDORP.frx":0838
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
         Picture         =   "frmCADVENDEDORP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
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
Attribute VB_Name = "frmCADVENDEDORP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Public strUSUARIO   As String
Dim objFuncoes      As Object
Dim objCADVENDEDOR  As Object
Dim iCodigo         As Integer

Private Sub cboFiltro_Change()
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

  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclus�o do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADVENDEDOR.GRAVA("E") = False Then Exit Sub
  
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
    If flxCADVENDEDORES.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADVENDEDOR = Nothing
    Unload Me
End Sub

Private Sub flxCADVENDEDORES_Click()
    If flxCADVENDEDORES.Rows > 1 Then objCADVENDEDOR.VENDCODIGO = CInt(flxCADVENDEDORES.TextMatrix(flxCADVENDEDORES.RowSel, 1))
End Sub

Private Sub flxCADVENDEDORES_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCADVENDEDORES.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCADVENDEDORES_RowColChange()
    If flxCADVENDEDORES.Rows > 1 Then objCADVENDEDOR.VENDCODIGO = CInt(flxCADVENDEDORES.TextMatrix(flxCADVENDEDORES.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADVENDEDOR = CreateObject("CADVENDEDOR.clsCADVENDEDOR")
    
    objCADVENDEDOR.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADVENDEDORP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "N�o foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
    cboFiltro.AddItem "C�digo"
    cboFiltro.AddItem "Nome"
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()

    If objCADVENDEDOR.Pesq_CadVendedor = False Then
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
        
    flxCADVENDEDORES.Rows = 1
    flxCADVENDEDORES.Cols = 3
    
    flxCADVENDEDORES.TextMatrix(0, 0) = ""
    flxCADVENDEDORES.TextMatrix(0, 1) = "C�digo"
    flxCADVENDEDORES.TextMatrix(0, 2) = "Nome"
    
    flxCADVENDEDORES.ColWidth(0) = 0
    flxCADVENDEDORES.ColWidth(1) = 700
    flxCADVENDEDORES.ColWidth(2) = 5000
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxCADVENDEDORES.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxCADVENDEDORES.Rows - 1)
             
            If flxCADVENDEDORES.TextMatrix(I, 1) = iCodigo Then
               flxCADVENDEDORES.Row = I
               flxCADVENDEDORES.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_DESCRICAO "
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flxCADESPORCA.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub


Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADVENDEDORP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    If cboFiltro.ListIndex = 0 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente � permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
             
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADVENDEDOR " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxCADVENDEDORES.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxCADVENDEDORES.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADVENDEDOR " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_DESCRICAO = '" & txtCampos.Text & "'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxCADVENDEDORES.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxCADVENDEDORES.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    
    ConfGrid
    PreencheGrid

End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
  
  If flxCADVENDEDORES.Rows > 1 Then iCodigo = CInt(flxCADVENDEDORES.TextMatrix(flxCADVENDEDORES.RowSel, 1))
  
  frmCADVENDEDOR.cCaminho = cCaminho
  frmCADVENDEDOR.Linha = Linha
  frmCADVENDEDOR.iCodigo = iCodigo
  frmCADVENDEDOR.cTipOper = strOperacao
  frmCADVENDEDOR.FILIAL = FILIAL
  frmCADVENDEDOR.strAcesso = strAcesso
  frmCADVENDEDOR.strMODPAI = Me.Name
  frmCADVENDEDOR.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub

