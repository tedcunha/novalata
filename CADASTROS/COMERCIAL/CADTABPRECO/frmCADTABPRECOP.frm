VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADTABPRECOP 
   Caption         =   "Cadastro de Tabela de Preços"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxGRIDTABPRECO 
         Height          =   3375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5953
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
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "txtCampos"
         Top             =   200
         Width           =   4455
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   8
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3840
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
         Picture         =   "frmCADTABPRECOP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   6120
         Picture         =   "frmCADTABPRECOP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADTABPRECOP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCADTABPRECOP.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADTABPRECOP.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADTABPRECOP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADTABPRECOP"
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
Dim objCADTABPRECO  As Object
Dim iCodigo         As Integer


Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Operacao "A"
End Sub

Private Sub cmdExclui_Click()
    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADTABPRECO = Nothing
    Unload Me
End Sub

Private Sub flxGRIDTABPRECO_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxGRIDTABPRECO.Rows > 1 Then Operacao "C"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADTABPRECO = CreateObject("CADTABPRECO.clsCADTABPRECO")
    
    objCADTABPRECO.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADTABPRECOP
    
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

    Exit Sub
    
    If objCADTABPRECO.Pesq_CadTipoAlim = False Then
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
    
    flxGRIDTABPRECO.Rows = 1
    flxGRIDTABPRECO.Cols = 5
    
    flxGRIDTABPRECO.TextMatrix(0, 0) = ""
    flxGRIDTABPRECO.TextMatrix(0, 1) = "Código"
    flxGRIDTABPRECO.TextMatrix(0, 2) = "Data"
    flxGRIDTABPRECO.TextMatrix(0, 3) = "Produto"
    flxGRIDTABPRECO.TextMatrix(0, 4) = "Descrição"
    
    flxGRIDTABPRECO.ColWidth(0) = 0
    flxGRIDTABPRECO.ColWidth(1) = 700
    flxGRIDTABPRECO.ColWidth(2) = 1000
    flxGRIDTABPRECO.ColWidth(3) = 1500
    flxGRIDTABPRECO.ColWidth(4) = 4000
    
End Sub

Private Sub PreencheGrid()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       TAB.SGI_CODTAB  " & vbCrLf
    sSql = sSql & "      ,TAB.SGI_DATATAB " & vbCrLf
    sSql = sSql & "      ,TAB.SGI_CODPROD " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_TABPRECO TAB " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       TAB.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL = TAB.SGI_FILIAL  " & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODIGO = TAB.SGI_CODPROD " & vbCrLf
    sSql = sSql & " Group By " & vbCrLf
    sSql = sSql & "          TAB.SGI_CODTAB  " & vbCrLf
    sSql = sSql & "         ,TAB.SGI_DATATAB " & vbCrLf
    sSql = sSql & "         ,TAB.SGI_CODPROD " & vbCrLf
    sSql = sSql & "         ,PROD.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       flxGRIDTABPRECO.AddItem "" & vbTab & _
                               BREC!SGI_CODTAB & vbTab & _
                               Format(BREC!SGI_DATATAB, "DD/MM/YYYY") & vbTab & _
                               BREC!SGI_CODPROD & vbTab & _
                               BREC!SGI_DESCRICAO
                               
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()
 
    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxGRIDTABPRECO.Rows - 1)
             
            If flxGRIDTABPRECO.TextMatrix(I, 1) = iCodigo Then
               flxGRIDTABPRECO.Row = I
               flxGRIDTABPRECO.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
    
  If flxGRIDTABPRECO.Rows > 1 Then iCodigo = CInt(flxGRIDTABPRECO.TextMatrix(flxGRIDTABPRECO.RowSel, 1))
    
  frmCADTABPRECO.cCaminho = cCaminho
  frmCADTABPRECO.Linha = Linha
  frmCADTABPRECO.iCodigo = iCodigo
  frmCADTABPRECO.cTipOper = strOperacao
  frmCADTABPRECO.FILIAL = FILIAL
  frmCADTABPRECO.strAcesso = strAcesso
  frmCADTABPRECO.strMODPAI = Me.Name
  frmCADTABPRECO.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub

