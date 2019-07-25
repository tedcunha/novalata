VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADSENHAP 
   Caption         =   "Cadastro de Usuários e Senha"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7935
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   200
         Width           =   1695
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   10
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   7935
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3240
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
         Picture         =   "frmCADSENHAP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmCADSENHAP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADSENHAP.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADSENHAP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADSENHAP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCADSENHAP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxSenha 
         Height          =   3615
         Left            =   120
         TabIndex        =   14
         Top             =   240
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
   Begin VB.PictureBox stMensagen 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
End
Attribute VB_Name = "frmCADSENHAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho  As String
Public Linha     As Variant
Public FILIAL    As Integer
Public strAcesso As String
Dim objFuncoes   As Object
Dim objCADSENHA  As Object
Dim iCodigo      As Integer

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
  
  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADSENHA.GRAVA("E") = False Then Exit Sub
  
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
    If flxSenha.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADSENHA = Nothing
    Unload Me
End Sub

Private Sub flxSenha_Click()
    If flxSenha.Rows > 1 Then
       objCADSENHA.SENCODIGO = CInt(flxSenha.TextMatrix(flxSenha.RowSel, 1))
    End If
End Sub

Private Sub flxSenha_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxSenha.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
       If flxSenha.Rows > 1 Then Operacao "C"
    End If
End Sub

Private Sub flxSenha_RowColChange()
    If flxSenha.Rows > 1 Then
       objCADSENHA.SENCODIGO = CInt(flxSenha.TextMatrix(flxSenha.RowSel, 1))
    End If
End Sub

Private Sub Form_Activate()
   If flxSenha.Enabled = True Then flxSenha.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADSENHA = CreateObject("CADSENHA.clsCADSENHA")
    
    objCADSENHA.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADSENHAP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Nome"
    
    objFuncoes.ChecaAcesso frmCADSENHAP, strAcesso
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()

    If objCADSENHA.Pesq_CadSenha = False Then
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

    flxSenha.Rows = 1
    flxSenha.Cols = 4
    
    flxSenha.TextMatrix(0, 1) = "Código"
    flxSenha.TextMatrix(0, 2) = "Descrição"
    flxSenha.TextMatrix(0, 3) = "Departamento"
    
    flxSenha.ColWidth(0) = 0
    flxSenha.ColWidth(1) = 700
    flxSenha.ColWidth(2) = 5000
    flxSenha.ColWidth(3) = 1500
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select * from SGI_USUARIO " & vbCrLf
    sSql = sSql & " Where SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxSenha.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & objFuncoes.Crypt(BREC!SGI_NOME) & vbTab & BREC!SGI_DEPARTAMENTO
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub Operacao(Operacao As String)
 
  Dim Pesquisa As String
     
  If flxSenha.Rows > 1 Then
     iCodigo = CInt(flxSenha.TextMatrix(flxSenha.RowSel, 1))
  End If
    
  frmCADSENHA.cCaminho = cCaminho
  frmCADSENHA.Linha = Linha
  frmCADSENHA.iCodigo = iCodigo
  frmCADSENHA.cTipOper = Operacao
  frmCADSENHA.FILIAL = FILIAL
  frmCADSENHA.strAcesso = strAcesso
  frmCADSENHA.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid
  
  objFuncoes.ChecaAcesso frmCADSENHAP, strAcesso

End Sub

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select * from SGI_USUARIO " & vbCrLf
     sSql = sSql & " Where SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select * from SGI_USUARIO " & vbCrLf
     sSql = sSql & " Where SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_NOME "
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flxSenha.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_NOME & vbTab & BREC!SGI_DEPARTAMENTO
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
     
        For I = 1 To (flxSenha.Rows - 1)
             
            If flxSenha.TextMatrix(I, 1) = iCodigo Then
               flxSenha.Row = I
               flxSenha.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub
Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADSENHA
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
       sSql = sSql & "      SGI_USUARIO" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_CODIGO = " & txtCampos.Text & vbCrLf
       sSql = sSql & "  And SGI_FILIAL = " & FILIAL
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxSenha.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_NOME & vbTab & BREC!SGI_DEPARTAMENTO
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxSenha.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_USUARIO" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_NOME LIKE '" & txtCampos.Text & "%'" & vbCrLf
       sSql = sSql & "  And SGI_FILIAL = " & FILIAL
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxSenha.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_NOME & vbTab & BREC!SGI_DEPARTAMENTO
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxSenha.SetFocus
          Exit Sub
          
       End If
    
    
    End If

    BREC.Close
    ConfGrid
    PreencheGrid

End Sub
