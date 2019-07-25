VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADCLIENTEP 
   Caption         =   "Cadastro de clientes"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   12315
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   12255
      Begin MSFlexGridLib.MSFlexGrid flxCLIENTES 
         Height          =   4575
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   8070
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   12255
      Begin VB.Timer Timer1 
         Interval        =   50000
         Left            =   4560
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
         Left            =   11400
         Picture         =   "frmCADCLIENTE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
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
         Left            =   10680
         Picture         =   "frmCADCLIENTE.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
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
         Left            =   2640
         Picture         =   "frmCADCLIENTE.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmCADCLIENTE.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Altera Empresa "
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
         Picture         =   "frmCADCLIENTE.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   240
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
         Picture         =   "frmCADCLIENTE.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cliente"
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
         Left            =   3480
         Picture         =   "frmCADCLIENTE.frx":129C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exclui a condição de pagamento"
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   8295
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
Attribute VB_Name = "frmCADCLIENTEP"
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
Public lngIDUsuario As Long

Dim objFuncoes      As New clsFuncoes
Dim objCADCLIENTES  As New clsCADCLIENTE
Dim iCodigo         As Long


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

  If Verif_reg = True Then Exit Sub
  
  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADCLIENTES.GRAVA("E") = False Then Exit Sub
  If objCADCLIENTES.Atualiza("E", objCADCLIENTES.CLIECODIGO, FILIAL, "frmCADCLIENTE") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  ''Atualiza_Grid
  ConfGrid
  AbilitaCampos
  
  
End Sub

Private Sub cmdInclui_Click()
    
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
    
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADCLIENTES = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()
  
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    
    If Verif_reg = True Then Exit Sub

    Converte
    
    Atualiza_Grid
    AbilitaCampos
End Sub


Private Sub flxCLIENTES_Click()
    If flxCLIENTES.Rows > 1 Then objCADCLIENTES.CLIECODIGO = CLng(flxCLIENTES.TextMatrix(flxCLIENTES.RowSel, 1))
End Sub

Private Sub flxCLIENTES_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCLIENTES.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCLIENTES_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If flxCLIENTES.Rows > 1 Then Operacao "C"
    End If
End Sub

Private Sub flxCLIENTES_RowColChange()
    If flxCLIENTES.Rows > 1 Then objCADCLIENTES.CLIECODIGO = CLng(flxCLIENTES.TextMatrix(flxCLIENTES.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    ''Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    ''Set objCADCLIENTES = CreateObject("CADCLIENTE.clsCADCLIENTE")
    
    objCADCLIENTES.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADCLIENTEP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call AbilitaCampos
    Call ConfGrid
    ''PreencheGrid
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "CPF/CNPJ"
    cboFiltro.AddItem "Razão social"
    cboFiltro.AddItem "Nome fantasia"
    
    cboFiltro.ListIndex = 0
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()

    If objCADCLIENTES.Pesq_CadCliente = False Then
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
    
    flxCLIENTES.Rows = 1
    flxCLIENTES.Cols = 6
    
    flxCLIENTES.TextMatrix(0, 0) = ""
    flxCLIENTES.TextMatrix(0, 1) = "Código"
    flxCLIENTES.TextMatrix(0, 2) = "CNPJ/CPF"
    flxCLIENTES.TextMatrix(0, 3) = "Razão social"
    flxCLIENTES.TextMatrix(0, 4) = "Nome Fantasia"
    flxCLIENTES.TextMatrix(0, 5) = "Habilitado"
    
    flxCLIENTES.ColWidth(0) = 0
    flxCLIENTES.ColWidth(1) = 700
    flxCLIENTES.ColWidth(2) = 1600
    flxCLIENTES.ColWidth(3) = 5000
    flxCLIENTES.ColWidth(4) = 3000
    flxCLIENTES.ColWidth(5) = 1000
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       flxCLIENTES.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           objFuncoes.FormataCnpj(BREC!SGI_CPFCNPJ) & vbTab & _
                           BREC!SGI_RAZAOSOC & vbTab & BREC!SGI_NOMFANTA
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub Operacao(strOperacao As String)
 
    If strOperacao = "A" Or strOperacao = "C" Then
       If Verif_reg = True Then Exit Sub
    End If
  
    Dim Pesquisa As String
  
    If flxCLIENTES.Rows > 1 Then iCodigo = CLng(flxCLIENTES.TextMatrix(flxCLIENTES.RowSel, 1))
  
    frmCADCLIENTE.cCaminho = cCaminho
    frmCADCLIENTE.Linha = Linha
    frmCADCLIENTE.iCodigo = iCodigo
    frmCADCLIENTE.cTipOper = strOperacao
    frmCADCLIENTE.FILIAL = FILIAL
    frmCADCLIENTE.strAcesso = strAcesso
    frmCADCLIENTE.strMODPAI = Me.Name
    frmCADCLIENTE.lngIDUsuario = lngIDUsuario
    frmCADCLIENTE.Show vbModal
  
    AbilitaCampos
    ConfGrid
  
  ''Atualiza_Grid
  ''AbilitaCampos
  
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxCLIENTES.Rows - 1)
             
            If flxCLIENTES.TextMatrix(I, 1) = iCodigo Then
               flxCLIENTES.Row = I
               flxCLIENTES.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Ordem()

    Call ConfGrid
    
    txtCampos.Text = ""
  
    sSql = ""
    sSql = " Select " & vbCrLf
    sSql = sSql & "        * " & vbCrLf
    sSql = sSql & "   from " & vbCrLf
    sSql = sSql & "        SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
  
    If cboFiltro.ListIndex = 0 Then sSql = sSql & " Order by SGI_CODIGO"
    If cboFiltro.ListIndex = 1 Then sSql = sSql & " Order by SGI_CPFCNPJ"
    If cboFiltro.ListIndex = 2 Then sSql = sSql & " Order by SGI_RAZAOSOC"
    If cboFiltro.ListIndex = 3 Then sSql = sSql & " Order by SGI_NOMFANTA"
  
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxCLIENTES.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           objFuncoes.FormataCnpj(BREC!SGI_CPFCNPJ) & vbTab & _
                           BREC!SGI_RAZAOSOC & vbTab & _
                           BREC!SGI_NOMFANTA & vbTab & _
                           IIf(BREC!SGI_DESBCLIE = 1, "Sim", "Não")
       BREC.MoveNext
    Loop
  
    BREC.Close

End Sub

Private Sub Timer1_Timer()
    ''Atualiza_Grid
    ''AbilitaCampos
End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADCLIENTEP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
             
       sSql = sSql & "  And SGI_CODIGO = " & Trim(Replace(Replace(txtCampos.Text, ",", ""), ".", ""))
    ElseIf cboFiltro.ListIndex = 1 Then
       sSql = sSql & "  And SGI_CPFCNPJ = '%" & Trim(txtCampos.Text) & "%'"
    ElseIf cboFiltro.ListIndex = 2 Then
       sSql = sSql & "  And SGI_RAZAOSOC LIKE '%" & Trim(txtCampos.Text) & "%'"
    ElseIf cboFiltro.ListIndex = 3 Then
       sSql = sSql & "  And SGI_NOMFANTA LIKE '%" & Trim(txtCampos.Text) & "%'"
    End If

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
        ConfGrid
        Do While Not BREC.EOF()
           flxCLIENTES.AddItem "" & vbTab & _
                               BREC!SGI_CODIGO & vbTab & _
                               BREC!SGI_CPFCNPJ & vbTab & _
                               BREC!SGI_RAZAOSOC & vbTab & _
                               BREC!SGI_NOMFANTA & vbTab & _
                               IIf(BREC!SGI_DESBCLIE = 1, "Sim", "Não")
           BREC.MoveNext
        Loop
    End If
    BREC.Close

End Sub

Private Sub Converte()

On Error GoTo err_grava
     
     Dim lngCODATUAl As Long
     Dim iResp       As Integer
     
     lngCODATUAl = objCADCLIENTES.CLIECODIGO

     If lngCODATUAl = 0 Then Exit Sub
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_CODIGO = " & lngCODATUAl
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     If BREC!SGI_ECLIENTE = "S" Then
        MsgBox "Este cliente não pode ser convertido !!!", vbOKOnly + vbCritical, "aviso"
        BREC.Close
        Exit Sub
     End If
     
     BREC.Close
     
     iResp = MsgBox("Deseja realmente converter este cliente <S/N> ?", vbYesNo + vbQuestion + vbDefaultButton2, "aviso")
    
     If iResp = vbNo Then Exit Sub
     
     '' ---------------------------------------------------------
     '' Trocando o Código
     
     lngCODATUAl = objCADCLIENTES.Gera_Codigo("frmCADCLIENTE", True)
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     '' ---------------------------------------------------------
     sSql = "Update SGI_CADCLIENTE Set " & vbCrLf
     sSql = sSql & "                          SGI_CODIGO   = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                         ,SGI_ECLIENTE = 'S'" & vbCrLf
     sSql = sSql & "         Where " & vbCrLf
     sSql = sSql & "               SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "           And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_DADOSGERAIS Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_MODULO = 'frmCADCLIENTE'" & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_REFBANCARIA Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_REFCOMERCIAL Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_REFPESSOAL Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_RESTRICOES Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_CLIFORNEC Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_CLIECLIENTES Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_CLISISCERT Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     sSql = "Update SGI_CLIATENDIDO Set SGI_CODIGO = " & lngCODATUAl & vbCrLf
     sSql = sSql & "                     Where " & vbCrLf
     sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "                       And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
     
     BGRV.CommandText = sSql
     BGRV.Execute
     '' ---------------------------------------------------------
     
     adoBanco_Dados.CommitTrans
     
     Exit Sub

err_grava:

    MsgBox "Ocorreu um erro !!!", vbOKOnly + vbCritical, "aviso"
    adoBanco_Dados.RollbackTrans
    
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
     sSql = sSql & "   And SGI_MODULO = 'frmCADCLIENTE'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (flxCLIENTES.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If flxCLIENTES.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                  If flxCLIENTES.Rows = 2 Then flxCLIENTES.Rows = 1
                  If flxCLIENTES.Rows > 2 Then flxCLIENTES.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(flxCLIENTES.TextMatrix(I, 1)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              
              flxCLIENTES.AddItem "" & vbTab & _
                                  BREC2!SGI_CODIGO & vbTab & _
                                  objFuncoes.FormataCnpj(BREC2!SGI_CPFCNPJ) & vbTab & _
                                  BREC2!SGI_RAZAOSOC & vbTab & _
                                  BREC2!SGI_NOMFANTA
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxCLIENTES.TextMatrix(I, 0) = ""
              flxCLIENTES.TextMatrix(I, 1) = BREC2!SGI_CODIGO
              flxCLIENTES.TextMatrix(I, 2) = objFuncoes.FormataCnpj(BREC2!SGI_CPFCNPJ)
              flxCLIENTES.TextMatrix(I, 3) = BREC2!SGI_RAZAOSOC
              flxCLIENTES.TextMatrix(I, 4) = BREC2!SGI_NOMFANTA
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
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & objCADCLIENTES.CLIECODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Este registro foi excluso !!!", vbOKOnly + vbExclamation, "Aviso"
       Verif_reg = True
    End If
    BREC.Close

End Function


Private Sub PopTab_ProdClie()

    On Error GoTo Err_PopTab_ProdClie

    Dim lngQTDREGS  As Long
    Dim arrPRODUTOS As Variant
    
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       CABE.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,ITEN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & " From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI ITEN" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH CABE" & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "       ITEN.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CABE.SGI_FILIAL = ITEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CABE.SGI_CODIGO = ITEN.SGI_CODIGO" & vbCrLf
    sSql = sSql & "Union" & vbCrLf
    sSql = sSql & "Select Distinct" & vbCrLf
    sSql = sSql & "       CABE.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,ITEN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI_STEEL ITEN" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH_STEEL CABE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       ITEN.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CABE.SGI_FILIAL = ITEN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CABE.SGI_CODIGO = ITEN.SGI_CODIGO" & vbCrLf
    sSql = sSql & "Order By CABE.SGI_CODCLI" & vbCrLf
    sSql = sSql & "        ,ITEN.SGI_IDPRODUTO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        lngQTDREGS = 0
        Do While Not BREC.EOF()
            lngQTDREGS = (lngQTDREGS + 1)
            BREC.MoveNext
        Loop
        
        ReDim arrPRODUTOS(1 To lngQTDREGS, 1 To 2) As Variant
        BREC.MoveFirst
        
        lngQTDREGS = 1
        Do While Not BREC.EOF()
            arrPRODUTOS(lngQTDREGS, 1) = BREC!SGI_CODCLI
            arrPRODUTOS(lngQTDREGS, 2) = BREC!SGI_IDPRODUTO
            lngQTDREGS = (lngQTDREGS + 1)
            BREC.MoveNext
        Loop
        objCADCLIENTES.PRODUTOS = arrPRODUTOS
    
    End If
    BREC.Close

    If objCADCLIENTES.GRAVA("IP") = True Then MsgBox "Produtos Inclusos com Exito !!!", vbOKOnly + vbExclamation, "Aviso"
        

    Exit Sub

Err_PopTab_ProdClie:
End Sub
 
 
