VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADENTRMATP 
   Caption         =   "Entrada Materiais Manual"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   13230
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   6975
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   13215
      Begin MSFlexGridLib.MSFlexGrid flxEntrMat 
         Height          =   6735
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   11880
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
      Width           =   13095
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
         Width           =   9615
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
      Width           =   13095
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
         Left            =   12120
         Picture         =   "frmCADENTRMATP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   11280
         Picture         =   "frmCADENTRMATP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADENTRMATP.frx":0634
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
         Picture         =   "frmCADENTRMATP.frx":0736
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
         Picture         =   "frmCADENTRMATP.frx":0838
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
         Picture         =   "frmCADENTRMATP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADENTRMATP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public strUSUARIO       As String
Public lngCodUsuario    As Long

Dim objFuncoes     As Object
Dim objCADENTRMAT  As Object
Dim iCodigo        As Long

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    If VerifNf = False Then Exit Sub
    Operacao "A"
End Sub

Private Sub cmdExclui_Click()
    
  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  Dim iResp As Integer
  
  If VerifNf = False Then Exit Sub
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADENTRMAT.Carrega_campos = False Then Exit Sub
    
  objCADENTRMAT.CODLCTO = objCADENTRMAT.Gera_Codigo("CARDEX")
  
  If objCADENTRMAT.GRAVA("E") = False Then Exit Sub
  
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
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADENTRMAT = Nothing
    Unload Me
End Sub

Private Sub flxEntrMat_Click()
    If flxEntrMat.Rows > 1 Then objCADENTRMAT.CADREQENTCOD = CLng(flxEntrMat.TextMatrix(flxEntrMat.RowSel, 1))
End Sub

Private Sub flxEntrMat_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxEntrMat.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxEntrMat_RowColChange()
    If flxEntrMat.Rows > 1 Then objCADENTRMAT.CADREQENTCOD = CLng(flxEntrMat.TextMatrix(flxEntrMat.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADENTRMAT = CreateObject("CADENTRMAT.clsCADENTRMAT")
    
    objCADENTRMAT.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADENTRMATP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
    cboFiltro.AddItem "Nº Req."
    cboFiltro.AddItem "Data."
    
    cboFiltro.ListIndex = 0

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()
    
    If objCADENTRMAT.Pesq_CadEntrReqMat = False Then
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
        
    flxEntrMat.Rows = 1
    flxEntrMat.Cols = 3
    
    flxEntrMat.TextMatrix(0, 1) = ""
    flxEntrMat.TextMatrix(0, 1) = "Nº Req."
    flxEntrMat.TextMatrix(0, 2) = "Data"
    
    flxEntrMat.ColWidth(0) = 0
    flxEntrMat.ColWidth(1) = 1500
    flxEntrMat.ColWidth(2) = 1500
    
End Sub


Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAT.* " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADREQENTRMAT MAT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MAT.SGI_FILIAL   = " & FILIAL & vbCrLf

    sSql = sSql & " Order by MAT.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxEntrMat.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & Format(BREC!SGI_DATREQ, "DD/MM/YYYY")
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxEntrMat.Rows - 1)
             
            If flxEntrMat.TextMatrix(I, 1) = iCodigo Then
               flxEntrMat.Row = I
               flxEntrMat.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub


Private Sub Operacao(Operacao As String)
 
  Dim Pesquisa As String
    
  If flxEntrMat.Rows > 1 Then iCodigo = CInt(flxEntrMat.TextMatrix(flxEntrMat.RowSel, 1))
  
  frmCADENTRMAT.cCaminho = cCaminho
  frmCADENTRMAT.Linha = Linha
  frmCADENTRMAT.iCodigo = iCodigo
  frmCADENTRMAT.cTipOper = Operacao
  frmCADENTRMAT.FILIAL = FILIAL
  frmCADENTRMAT.strAcesso = strAcesso
  frmCADENTRMAT.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub

Private Function VerifNf() As Boolean

    VerifNf = True
    
    '' ----------------------------------------------------------------
    '' Verifica se existe NF
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_NFENTRADACABEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODREQENTR = " & objCADENTRMAT.CADREQENTCOD
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Este requisição está ligada a uma NF !!!", vbOKOnly + vbExclamation, "Aviso"
       VerifNf = False
    End If
    BREC.Close
    '' ----------------------------------------------------------------

End Function


Private Sub Ordem()

    Call ConfGrid


    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAT.* " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADREQENTRMAT MAT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MAT.SGI_FILIAL   = " & FILIAL & vbCrLf
    

    If cboFiltro.ListIndex = 0 Then sSql = sSql & " Order by MAT.SGI_CODIGO"
    If cboFiltro.ListIndex = 1 Then sSql = sSql & " Order by MAT.SGI_DATREQ"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxEntrMat.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & Format(BREC!SGI_DATREQ, "DD/MM/YYYY")
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

