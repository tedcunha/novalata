VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADFLUXCAIXAP 
   Caption         =   "Fluxo de Caixa"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   9615
      Begin MSFlexGridLib.MSFlexGrid flxFluxCaixa 
         Height          =   3375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   9615
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   5400
         Top             =   120
      End
      Begin VB.CommandButton cmdReabre 
         Caption         =   "&Reabre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         Picture         =   "frmCADFLUXCAIXAP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdFecha 
         Caption         =   "&Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         Picture         =   "frmCADFLUXCAIXAP.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Exclui Empresa"
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
         Height          =   735
         Left            =   120
         Picture         =   "frmCADFLUXCAIXAP.frx":0884
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
         Height          =   735
         Left            =   960
         Picture         =   "frmCADFLUXCAIXAP.frx":0DB6
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
         Height          =   735
         Left            =   1800
         Picture         =   "frmCADFLUXCAIXAP.frx":12E8
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
         Height          =   735
         Left            =   2640
         Picture         =   "frmCADFLUXCAIXAP.frx":13EA
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
         Height          =   735
         Left            =   7800
         Picture         =   "frmCADFLUXCAIXAP.frx":14EC
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
         Height          =   735
         Left            =   8640
         Picture         =   "frmCADFLUXCAIXAP.frx":1A1E
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
      Width           =   9615
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   6135
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
Attribute VB_Name = "frmCADFLUXCAIXAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Long
Public strAcesso    As String
Public strUsuario   As String
Dim objFuncoes      As Object
Dim objCADFLUXCAIXA As Object
Dim iCodigo         As Long

Private Sub cmdAltera_Click()

    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    If flxFluxCaixa.TextMatrix(flxFluxCaixa.Row, 3) = "FECHADO" Then
       MsgBox "Caixa fechado !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    Operacao "A"
    
End Sub

Private Sub cmdExclui_Click()

    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
    
    Dim intResp As Integer
    
    If flxFluxCaixa.TextMatrix(flxFluxCaixa.Row, 3) = "FECHADO" Then
       MsgBox "Caixa fechado reabra o caixa !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    intResp = MsgBox("Deseja realmente excluir o caixa ?", vbYesNo + vbQuestion + vbDefaultButton2, "Pergunta")
    
    If intResp = 7 Then Exit Sub
    
    objCADFLUXCAIXA.DATAFLXCAIXA = CDate(flxFluxCaixa.TextMatrix(flxFluxCaixa.Row, 2))
    If objCADFLUXCAIXA.GRAVA("E") = False Then Exit Sub
    
    MsgBox "O Caixa foi excluso com exito !!!", vbOKOnly + vbInformation, "Aviso"
    
    AbilitaCampos
    ConfGrid
    PreencheGrid

End Sub

Private Sub cmdFecha_Click()

    If ChefaCaixa = False Then Exit Sub
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdReabre_Click()

    If VerifLctos = True Then Exit Sub
    If ReabreCaixa = False Then Exit Sub
    
    AbilitaCampos
    ConfGrid
    PreencheGrid

End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADFLUXCAIXA = Nothing
    Unload Me
End Sub


Private Sub flxFluxCaixa_Click()
    If flxFluxCaixa.Rows > 1 Then objCADFLUXCAIXA.CODFLXCAIXA = CLng(flxFluxCaixa.TextMatrix(flxFluxCaixa.RowSel, 1))
End Sub

Private Sub flxFluxCaixa_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxFluxCaixa.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxFluxCaixa_RowColChange()
    If flxFluxCaixa.Rows > 1 Then objCADFLUXCAIXA.CODFLXCAIXA = CLng(flxFluxCaixa.TextMatrix(flxFluxCaixa.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADFLUXCAIXA = CreateObject("CADFLUXCAIXA.clsCADFLUXCAIXA")
    
    objCADFLUXCAIXA.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADFLUXCAIXAP
    
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
    
    If objCADFLUXCAIXA.Pesq_CadFluxCaixa = False Then
       cmdAltera.Enabled = False
       cmdExclui.Enabled = False
       cmdFecha.Enabled = False
       cmdReabre.Enabled = False
       Frame1.Enabled = False
       Frame3.Enabled = False
    Else
       cmdAltera.Enabled = True
       cmdExclui.Enabled = True
       cmdFecha.Enabled = True
       cmdReabre.Enabled = True
       Frame1.Enabled = True
       Frame3.Enabled = True
    End If

End Sub

Private Sub ConfGrid()
    
    flxFluxCaixa.Rows = 1
    flxFluxCaixa.Cols = 4
    
    flxFluxCaixa.TextMatrix(0, 0) = ""
    flxFluxCaixa.TextMatrix(0, 1) = "Código"
    flxFluxCaixa.TextMatrix(0, 2) = "Data"
    flxFluxCaixa.TextMatrix(0, 3) = "Status"
    
    flxFluxCaixa.ColWidth(0) = 0
    flxFluxCaixa.ColWidth(1) = 1000
    flxFluxCaixa.ColWidth(2) = 1000
    flxFluxCaixa.ColWidth(3) = 1000
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFLXCXHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_DATA DESC"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxFluxCaixa.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & Format(BREC!SGI_DATA, "DD/MM/YYYY") & vbTab & BREC!SGI_STATUS
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxFluxCaixa.Rows - 1)
             
            If flxFluxCaixa.TextMatrix(I, 1) = iCodigo Then
               flxFluxCaixa.Row = I
               flxFluxCaixa.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
  
  If flxFluxCaixa.Rows > 1 Then iCodigo = CInt(flxFluxCaixa.TextMatrix(flxFluxCaixa.RowSel, 1))
  
  frmCADFLUXCAIXA.cCaminho = cCaminho
  frmCADFLUXCAIXA.Linha = Linha
  frmCADFLUXCAIXA.iCodigo = iCodigo
  frmCADFLUXCAIXA.cTipOper = strOperacao
  frmCADFLUXCAIXA.FILIAL = FILIAL
  frmCADFLUXCAIXA.strAcesso = strAcesso
  frmCADFLUXCAIXA.strMODPAI = Me.Name
  frmCADFLUXCAIXA.strUsuario = strUsuario
  
  frmCADFLUXCAIXA.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid
  
End Sub

Private Function ChefaCaixa() As Boolean

    ChefaCaixa = False
    
    Dim intResp As Integer
    
    If flxFluxCaixa.TextMatrix(flxFluxCaixa.Row, 3) = "FECHADO" Then
       MsgBox "Caixa fechado !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    
    intResp = MsgBox("Deseja realmente fechar o caixa ?", vbYesNo + vbQuestion, "Aviso")
    
    If intResp = 7 Then Exit Function
    
    sSql = "Update SGI_CADFLXCXHEADER Set SGI_STATUS = 'FECHADO'" & vbCrLf
    sSql = sSql & "                    Where " & vbCrLf
    sSql = sSql & "                          SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "                      And SGI_CODIGO = " & objCADFLUXCAIXA.CODFLXCAIXA
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    MsgBox "Caixa fechado !!!", vbOKOnly + vbInformation, "Aviso"
    
    ChefaCaixa = True
    
End Function

Private Function ReabreCaixa() As Boolean

    ReabreCaixa = False
    
    Dim intResp As Integer
    
    If flxFluxCaixa.TextMatrix(flxFluxCaixa.Row, 3) = "ABERTO" Then
       MsgBox "Caixa aberto !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    
    intResp = MsgBox("Deseja realmente reabrir o caixa ?", vbYesNo + vbQuestion, "Aviso")
    
    If intResp = 7 Then Exit Function
    
    sSql = "Update SGI_CADFLXCXHEADER Set SGI_STATUS = 'ABERTO'" & vbCrLf
    sSql = sSql & "                    Where " & vbCrLf
    sSql = sSql & "                          SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "                      And SGI_CODIGO = " & objCADFLUXCAIXA.CODFLXCAIXA
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    MsgBox "Caixa reaberto !!!", vbOKOnly + vbInformation, "Aviso"
    
    ReabreCaixa = True
    
End Function


Private Function VerifLctos() As Boolean

    VerifLctos = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFLXCXHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_DATA   > '" & Format(CDate(flxFluxCaixa.TextMatrix(flxFluxCaixa.Row, 2)), "MM/DD/YYYY") & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       VerifLctos = True
       MsgBox "Existe lançamentos posteriores a este caixa, reabra e exclua os caixas posteriores !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

End Function

