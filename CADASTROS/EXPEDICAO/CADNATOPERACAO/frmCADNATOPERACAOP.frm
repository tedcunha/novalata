VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADNATOPERACAOP 
   Caption         =   "Cadastro de natureza de operação"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   4335
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   8775
      Begin MSFlexGridLib.MSFlexGrid flxNATOPERACAO 
         Height          =   3975
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7011
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
      Width           =   8775
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3600
         Top             =   120
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
         Picture         =   "frmCADNATOPERACAOP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmCADNATOPERACAOP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmCADNATOPERACAOP.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmCADNATOPERACAOP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   6960
         Picture         =   "frmCADNATOPERACAOP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   7800
         Picture         =   "frmCADNATOPERACAOP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   200
         Width           =   1695
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtCampos"
         Top             =   200
         Width           =   5295
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmCADNATOPERACAOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public strUsuario       As String
Public lngCodUsuaro     As Long
Dim objFuncoes          As Object
Dim objCADNATOPERACAO   As Object
Dim lngCODIGO           As Long

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
    Dim lngCodLog As Long
  
    iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
    If iResp <> 6 Then Exit Sub
  
    If objCADNATOPERACAO.GRAVA("E") = False Then Exit Sub
  
    If objFuncoes.Atualiza("E", objCADNATOPERACAO.CODIGO, FILIAL, "frmCADNATOPERACAO", Linha) = False Then Exit Sub
    
    lngCodLog = objFuncoes.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objFuncoes.GravaLogModulo(FILIAL, lngCodLog, "frmCADNATOPERACAO", "E", lngCodUsuaro, Str(objCADNATOPERACAO.CODIGO), Linha)
  
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
    If flxNATOPERACAO.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub flxNATOPERACAO_Click()
    If (flxNATOPERACAO.Rows - 1) > 0 And flxNATOPERACAO.Row > 0 Then objCADNATOPERACAO.CODIGO = CLng(flxNATOPERACAO.TextMatrix(flxNATOPERACAO.RowSel, 0))
End Sub

Private Sub flxNATOPERACAO_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If (flxNATOPERACAO.Rows - 1) > 0 And flxNATOPERACAO.Row > 0 Then Operacao "C"
End Sub

Private Sub flxNATOPERACAO_RowColChange()
    If (flxNATOPERACAO.Rows - 1) > 0 And flxNATOPERACAO.Row > 0 Then objCADNATOPERACAO.CODIGO = CLng(flxNATOPERACAO.TextMatrix(flxNATOPERACAO.RowSel, 0))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADNATOPERACAO = CreateObject("CADNATOPERACAO.clsCADNATOPERACAO")
    
    objCADNATOPERACAO.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADNATOPERACAOP
    
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
    cboFiltro.AddItem "Tipo"
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()
    
    If objCADNATOPERACAO.Pesq_CadNatOperacao = False Then
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
    flxNATOPERACAO.Rows = 1
    flxNATOPERACAO.Cols = 5
    
    flxNATOPERACAO.TextMatrix(0, 1) = "Código"
    flxNATOPERACAO.TextMatrix(0, 2) = "Descrição"
    flxNATOPERACAO.TextMatrix(0, 3) = "Tipo"
    flxNATOPERACAO.TextMatrix(0, 4) = "Default"
    
    flxNATOPERACAO.ColWidth(0) = 1000
    flxNATOPERACAO.ColWidth(1) = 1000
    flxNATOPERACAO.ColWidth(2) = 4000
    flxNATOPERACAO.ColWidth(3) = 1500
    flxNATOPERACAO.ColWidth(4) = 1000
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxNATOPERACAO.AddItem BREC!SGI_CODIGO & vbTab & _
                              BREC!SGI_CODIGO & vbTab & _
                              BREC!SGI_DESCRICAO & vbTab & _
                              IIf(BREC!SGI_ENTSAI = 1, "SAIDA", IIf(BREC!SGI_ENTSAI = 0, "ENTRADA", IIf(BREC!SGI_ENTSAI = 2, "TRANSFERENCIA", ""))) & vbTab & _
                              IIf(BREC!SGI_DEFAULT = 1, "Sim", "Não")

       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub


Private Sub Operacao(Operacao As String)
 
  Dim Pesquisa As String
  
  If (flxNATOPERACAO.Rows - 1) > 0 And flxNATOPERACAO.Row > 0 Then lngCODIGO = CLng(flxNATOPERACAO.TextMatrix(flxNATOPERACAO.RowSel, 0))
  
  frmCADNATOPERACAO.cCaminho = cCaminho
  frmCADNATOPERACAO.Linha = Linha
  frmCADNATOPERACAO.lngCODIGO = lngCODIGO
  frmCADNATOPERACAO.cTipOper = Operacao
  frmCADNATOPERACAO.FILIAL = FILIAL
  frmCADNATOPERACAO.strAcesso = strAcesso
  frmCADNATOPERACAO.lngCodUsuario = lngCodUsuaro
  frmCADNATOPERACAO.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub


Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADNATOPERACAOP
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
    sSql = sSql & "      SGI_CADNATOPERACAO " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL
    
    If cboFiltro.ListIndex = 0 Then
             
       sSql = sSql & "  And SGI_NOMECLCOD Like '" & Trim(txtCampos.Text) & "%'"
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = sSql & "  And SGI_DESCRICAO Like '" & Trim(txtCampos.Text) & "%'"
    
    ElseIf cboFiltro.ListIndex = 2 Then
    
       If UCase(Trim(Mid(txtCampos.Text, 1, 1))) = "S" Then
          sSql = sSql & "  And SGI_ENTSAI = 1"
       ElseIf UCase(Trim(Mid(txtCampos.Text, 1, 1))) = "E" Then
          sSql = sSql & "  And SGI_ENTSAI = 0"
       End If
       
    End If

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
      
    If Not BREC.EOF Then
         
       Call ConfGrid
       Do While Not BREC.EOF
          flxNATOPERACAO.AddItem BREC!SGI_CODIGO & vbTab & _
                                 Trim(BREC!SGI_NOMECLCOD) & vbTab & _
                                 Trim(BREC!SGI_DESCRICAO) & vbTab & _
                                 IIf(BREC!SGI_ENTSAI = 1, "SAIDA", "ENTRADA") & vbTab & _
                                 IIf(BREC!SGI_DEFAULT = 1, "Sim", "Não")
          BREC.MoveNext
       Loop
       
    Else
       MsgBox "Este Registro Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    
    BREC.Close
    
End Sub

Private Sub Ordem()

    ConfGrid
  
    txtCampos.Text = ""
  
    sSql = ""
    sSql = " Select " & vbCrLf
    sSql = sSql & "        * " & vbCrLf
    sSql = sSql & "   from " & vbCrLf
    sSql = sSql & "        SGI_CADNATOPERACAO " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
  
    If cboFiltro.ListIndex = 0 Then
       sSql = sSql & " Order by SGI_NOMECLCOD "
    ElseIf cboFiltro.ListIndex = 1 Then
       sSql = sSql & " Order by SGI_DESCRICAO "
    ElseIf cboFiltro.ListIndex = 2 Then
       sSql = sSql & " Order by SGI_ENTSAI "
    End If
    
    BREC.Open sSql, adoBanco_Dados
    Do While Not BREC.EOF
       flxNATOPERACAO.AddItem BREC!SGI_CODIGO & vbTab & _
                              BREC!SGI_NOMECLCOD & vbTab & _
                              BREC!SGI_DESCRICAO & vbTab & _
                              IIf(BREC!SGI_ENTSAI = 1, "SAIDA", "ENTRADA") & vbTab & _
                              IIf(BREC!SGI_DEFAULT = 1, "Sim", "Não")
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

