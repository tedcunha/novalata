VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADFORNECP 
   Caption         =   "Cadastro de fornecedores"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11325
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   11295
      Begin MSFlexGridLib.MSFlexGrid flxCADFORNEC 
         Height          =   4695
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8281
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
      Width           =   11295
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
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "txtCampos"
         Top             =   200
         Width           =   6735
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
      Width           =   11295
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
         Picture         =   "frmCADFORNECP.frx":0000
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
         Picture         =   "frmCADFORNECP.frx":0532
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
         Picture         =   "frmCADFORNECP.frx":0A64
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
         Picture         =   "frmCADFORNECP.frx":0B66
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
         Left            =   9480
         Picture         =   "frmCADFORNECP.frx":0C68
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
         Left            =   10320
         Picture         =   "frmCADFORNECP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADFORNECP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public FILIAL      As Integer
Public strACESSO   As String
Dim objFuncoes     As Object
Dim objCADFORNEC   As Object
Dim iCodigo        As Integer

Private Sub cboFiltro_Change()
    txtCampos.Text = ""
    txtCampos.SetFocus
    
    ConfGrid
    PreencheGrid
End Sub

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strACESSO) = False Then Exit Sub
    Operacao "A"
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    ConfGrid
    PreencheGrid
End Sub

Private Sub cmdExclui_Click()
  
  If objFuncoes.ChecaAcesso2("E", strACESSO) = False Then Exit Sub
  
  If Verif_reg = True Then Exit Sub
  
'  If ConfereContsAPG = True Then
'     MsgBox "Não Foi Possivel Excluir o Fornecedor Pois Existe Contas A Pagar !!!", vbOKOnly + vbExclamation, "Aviso"
'     Exit Sub
'  End If
  
  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADFORNEC.GRAVA("E") = False Then Exit Sub
  If objCADFORNEC.Atualiza("E", objCADFORNEC.CodigoFOR, FILIAL, "frmCADFORNEC") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  AbilitaCampos
  Atualiza_Grid

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strACESSO) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxCADFORNEC.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADFORNEC = Nothing
    Unload Me
End Sub

Private Sub flxCADFORNEC_Click()
    If flxCADFORNEC.Rows > 1 Then objCADFORNEC.CodigoFOR = CInt(flxCADFORNEC.TextMatrix(flxCADFORNEC.RowSel, 1))
End Sub

Private Sub flxCADFORNEC_DblClick()
    If objFuncoes.ChecaAcesso2("C", strACESSO) = False Then Exit Sub
    If flxCADFORNEC.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCADFORNEC_RowColChange()
    If flxCADFORNEC.Rows > 1 Then objCADFORNEC.CodigoFOR = CInt(flxCADFORNEC.TextMatrix(flxCADFORNEC.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADFORNEC = CreateObject("CADFORNEC.clsCADFORNEC")
    
    objCADFORNEC.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADFORNECP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "CPF/CNPJ"
    cboFiltro.AddItem "Razão social"
    cboFiltro.AddItem "Nome fantasia"
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub Operacao(strOperacao As String)
 
  If strOperacao = "A" Or strOperacao = "C" Then
     If Verif_reg = True Then Exit Sub
  End If
  
  Dim Pesquisa As String
  
  If flxCADFORNEC.Rows > 1 Then iCodigo = CInt(flxCADFORNEC.TextMatrix(flxCADFORNEC.RowSel, 1))
  
  frmCADFORNEC.cCaminho = cCaminho
  frmCADFORNEC.Linha = Linha
  frmCADFORNEC.iCodigo = iCodigo
  frmCADFORNEC.cTipOper = strOperacao
  frmCADFORNEC.FILIAL = FILIAL
  frmCADFORNEC.strACESSO = strACESSO
  frmCADFORNEC.strMODPAI = Me.Name
  frmCADFORNEC.Show vbModal
  
  AbilitaCampos
  Atualiza_Grid
  
End Sub


Private Sub AbilitaCampos()

    If objCADFORNEC.Pesq_CadFornec = False Then
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
    
    flxCADFORNEC.Rows = 1
    flxCADFORNEC.Cols = 5
    
    flxCADFORNEC.TextMatrix(0, 0) = ""
    flxCADFORNEC.TextMatrix(0, 1) = "Código"
    flxCADFORNEC.TextMatrix(0, 2) = "CNPJ/CPF"
    flxCADFORNEC.TextMatrix(0, 3) = "Razão social"
    flxCADFORNEC.TextMatrix(0, 4) = "Nome Fantasia"
    
    flxCADFORNEC.ColWidth(0) = 0
    flxCADFORNEC.ColWidth(1) = 700
    flxCADFORNEC.ColWidth(2) = 1500
    flxCADFORNEC.ColWidth(3) = 5000
    flxCADFORNEC.ColWidth(4) = 3000
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxCADFORNEC.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_CPFCNPJ & vbTab & BREC!SGI_RAZAOSOC & vbTab & BREC!SGI_NOMFANTA
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxCADFORNEC.Rows - 1)
             
            If flxCADFORNEC.TextMatrix(I, 1) = iCodigo Then
               flxCADFORNEC.Row = I
               flxCADFORNEC.Col = 1
               
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
     sSql = sSql & "        SGI_CADFORNEC " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
     
  ElseIf cboFiltro.ListIndex = 1 Then
     
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADFORNEC " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL         = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CPFCNPJ "
  
  ElseIf cboFiltro.ListIndex = 2 Then
     
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADFORNEC " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL         = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_RAZAOSOC "
  
  ElseIf cboFiltro.ListIndex = 3 Then
     
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADFORNEC " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL         = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_NOMFANTA "
  
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flxCADFORNEC.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_CPFCNPJ & vbTab & BREC!SGI_RAZAOSOC & vbTab & BREC!SGI_NOMFANTA
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub


Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADFORNECP
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
       sSql = sSql & "      SGI_CADFORNEC " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxCADFORNEC.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_CPFCNPJ & vbTab & BREC!SGI_RAZAOSOC & vbTab & BREC!SGI_NOMFANTA
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxCADFORNEC.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADFORNEC " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_CPFCNPJ = '" & txtCampos.Text & "'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxCADFORNEC.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_CPFCNPJ & vbTab & BREC!SGI_RAZAOSOC & vbTab & BREC!SGI_NOMFANTA
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxCADFORNEC.SetFocus
          Exit Sub
          
       End If
    
    ElseIf cboFiltro.ListIndex = 2 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADFORNEC " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL      =  " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_RAZAOSOC LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxCADFORNEC.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_CPFCNPJ & vbTab & BREC!SGI_RAZAOSOC & vbTab & BREC!SGI_NOMFANTA
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxCADFORNEC.SetFocus
          Exit Sub
          
       End If
    
    ElseIf cboFiltro.ListIndex = 3 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADFORNEC " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL      =  " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_NOMFANTA LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxCADFORNEC.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_CPFCNPJ & vbTab & BREC!SGI_RAZAOSOC & vbTab & BREC!SGI_NOMFANTA
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxCADFORNEC.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    
    ConfGrid
    PreencheGrid

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
     sSql = sSql & "   And SGI_MODULO = 'frmCADFORNEC'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (flxCADFORNEC.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If flxCADFORNEC.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                  If flxCADFORNEC.Rows = 2 Then flxCADFORNEC.RemoveItem I
                  If flxCADFORNEC.Rows > 2 Then flxCADFORNEC.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = flxCADFORNEC.TextMatrix(I, 1) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxCADFORNEC.AddItem "" & vbTab & _
                                   BREC2!SGI_CODIGO & vbTab & _
                                   BREC2!SGI_CPFCNPJ & vbTab & _
                                   BREC2!SGI_RAZAOSOC & vbTab & _
                                   BREC2!SGI_NOMFANTA
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxCADFORNEC.TextMatrix(I, 0) = ""
              flxCADFORNEC.TextMatrix(I, 1) = BREC2!SGI_CODIGO
              flxCADFORNEC.TextMatrix(I, 2) = BREC2!SGI_CPFCNPJ
              flxCADFORNEC.TextMatrix(I, 3) = BREC2!SGI_RAZAOSOC
              flxCADFORNEC.TextMatrix(I, 4) = BREC2!SGI_NOMFANTA
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
      
End Sub


Private Function Verif_reg() As Boolean

    Verif_reg = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & objCADFORNEC.CodigoFOR
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Este registro foi excluso !!!", vbOKOnly + vbExclamation, "Aviso"
       Verif_reg = True
    End If
    BREC.Close

End Function

Private Function ConfereContsAPG() As Boolean

    ConfereContsAPG = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CONTASHAPG " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODFOR = " & objCADFORNEC.CodigoFOR
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then ConfereContsAPG = True
    BREC.Close

End Function
