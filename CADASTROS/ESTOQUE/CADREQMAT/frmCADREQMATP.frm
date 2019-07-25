VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADREQMATP 
   Caption         =   "Cadastro de Requisição de Materiais"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   12975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   6495
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   12975
      Begin MSFlexGridLib.MSFlexGrid flxReqMat 
         Height          =   6255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   11033
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
      Width           =   12975
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
         Width           =   9495
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
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   12975
      Begin VB.CommandButton cmdImpLotes 
         Caption         =   "Im&prime"
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
         Picture         =   "frmCADREQMATP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Baixa titulos ou lote de pagamento"
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   5640
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
         Height          =   735
         Left            =   12000
         Picture         =   "frmCADREQMATP.frx":030A
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
         Height          =   735
         Left            =   11160
         Picture         =   "frmCADREQMATP.frx":040C
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
         Height          =   735
         Left            =   2640
         Picture         =   "frmCADREQMATP.frx":093E
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
         Height          =   735
         Left            =   1800
         Picture         =   "frmCADREQMATP.frx":0A40
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
         Height          =   735
         Left            =   960
         Picture         =   "frmCADREQMATP.frx":0B42
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
         Height          =   735
         Left            =   120
         Picture         =   "frmCADREQMATP.frx":1074
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADREQMATP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public FILIAL      As Integer
Public strAcesso   As String
Public strUSUARIO  As String
Dim objFuncoes     As Object
Dim objCADREQMAT   As Object
Dim objREL         As Object
Dim iCodigo        As Long
Dim cCamRel        As String

Private Sub cboFiltro_Change()
    txtCampos.Text = ""
    txtCampos.SetFocus
    ConfGrid
    PreencheGrid
End Sub

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    If VerifReqSai = True Then
       MsgBox "Existe requisição de saidas já emitida !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    Operacao "A"
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    Call ConfGrid
End Sub

Private Sub cmdExclui_Click()
  
  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  Dim iResp As Integer
  
  If VerifReqSai = True Then
     MsgBox "Existe requisição de saidas já emitida !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  End If
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADREQMAT.GRAVA("E") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  AbilitaCampos
  ConfGrid

End Sub

Private Sub cmdImpLotes_Click()
    Call ImpReqMat(objCADREQMAT.CADREQCOD)
End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADREQMAT = Nothing
    Unload Me
End Sub
Private Sub flxReqMat_Click()
    If (flxReqMat.Rows - 1) > 0 And flxReqMat.Row > 0 Then objCADREQMAT.CADREQCOD = CLng(flxReqMat.TextMatrix(flxReqMat.RowSel, 1))
End Sub

Private Sub flxReqMat_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If (flxReqMat.Rows - 1) > 0 And flxReqMat.Row > 0 Then Call Operacao("C")
End Sub

Private Sub flxReqMat_RowColChange()
    If (flxReqMat.Rows - 1) > 0 And flxReqMat.Row > 0 Then objCADREQMAT.CADREQCOD = CLng(flxReqMat.TextMatrix(flxReqMat.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADREQMAT = CreateObject("CADREQMAT.clsCADREQMAT")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    objCADREQMAT.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADREQMATP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call AbilitaCampos
    Call ConfGrid
    
    cboFiltro.AddItem "Nº Req."
    cboFiltro.AddItem "Departamento"
    cboFiltro.AddItem "Data. Req."
    
    cboFiltro.ListIndex = 0
    
    '' --------------------------------------
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    ''cCamRel = "\\pc6\HD\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\ESTOQUE\"
    

End Sub

Private Sub AbilitaCampos()
    
    If objCADREQMAT.Pesq_CadReqMat = False Then
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
    
    flxReqMat.Rows = 1
    flxReqMat.Cols = 4
    
    flxReqMat.TextMatrix(0, 1) = "Nº Req."
    flxReqMat.TextMatrix(0, 2) = "Departamento"
    flxReqMat.TextMatrix(0, 3) = "Data Req."
    
    flxReqMat.ColWidth(0) = 0
    flxReqMat.ColWidth(1) = 1500
    flxReqMat.ColWidth(2) = 4000
    flxReqMat.ColWidth(3) = 1500
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAT.* " & vbCrLf
    sSql = sSql & "      ,DPT.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADREQMAT MAT" & vbCrLf
    sSql = sSql & "      ,SGI_CADDEPTO  DPT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MAT.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And DPT.SGI_FILIAL   = MAT.SGI_FILIAL   " & vbCrLf
    sSql = sSql & "   And DPT.SGI_CODDEPTO = MAT.SGI_CODDEPTO "

    sSql = sSql & " Order by MAT.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxReqMat.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & Format(BREC!SGI_DATREQ, "DD/MM/YYYY")
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxReqMat.Rows - 1)
             
            If flxReqMat.TextMatrix(I, 1) = iCodigo Then
               flxReqMat.Row = I
               flxReqMat.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(Operacao As String)
 
  Dim Pesquisa As String
  
  If (flxReqMat.Rows - 1) > 0 Then iCodigo = CInt(flxReqMat.TextMatrix(flxReqMat.RowSel, 1))
  
  frmCADREQMAT.cCaminho = cCaminho
  frmCADREQMAT.Linha = Linha
  frmCADREQMAT.iCodigo = iCodigo
  frmCADREQMAT.cTipOper = Operacao
  frmCADREQMAT.FILIAL = FILIAL
  frmCADREQMAT.strAcesso = strAcesso
  frmCADREQMAT.Show vbModal
  
  Call AbilitaCampos
  Call ConfGrid

End Sub

Private Sub Ordem()

  Call ConfGrid
  
  If (flxReqMat.Rows - 1) = 0 Then Exit Sub
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAT.* " & vbCrLf
    sSql = sSql & "      ,DPT.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADREQMAT MAT" & vbCrLf
    sSql = sSql & "      ,SGI_CADDEPTO  DPT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MAT.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And DPT.SGI_FILIAL   = MAT.SGI_FILIAL   " & vbCrLf
    sSql = sSql & "   And DPT.SGI_CODDEPTO = MAT.SGI_CODDEPTO "
    sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAT.* " & vbCrLf
    sSql = sSql & "      ,DPT.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADREQMAT MAT" & vbCrLf
    sSql = sSql & "      ,SGI_CADDEPTO  DPT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MAT.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And DPT.SGI_FILIAL   = MAT.SGI_FILIAL   " & vbCrLf
    sSql = sSql & "   And DPT.SGI_CODDEPTO = MAT.SGI_CODDEPTO "
    sSql = sSql & " Order by SGI_DESCRICAO "
  ElseIf cboFiltro.ListIndex = 2 Then
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAT.* " & vbCrLf
    sSql = sSql & "      ,DPT.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADREQMAT MAT" & vbCrLf
    sSql = sSql & "      ,SGI_CADDEPTO  DPT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MAT.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And DPT.SGI_FILIAL   = MAT.SGI_FILIAL   " & vbCrLf
    sSql = sSql & "   And DPT.SGI_CODDEPTO = MAT.SGI_CODDEPTO "
    sSql = sSql & " Order by SGI_DATREQ "
  End If
  
  BREC.Open sSql, adoBanco_Dados
  Do While Not BREC.EOF
     flxReqMat.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & Format(BREC!SGI_DATREQ, "DD/MM/YYYY")
     BREC.MoveNext
  Loop
  BREC.Close

End Sub


Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADREQMATP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    Call ConfGrid
    
    sSql = ""
    
    If cboFiltro.ListIndex = 0 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
             
       sSql = "Select " & vbCrLf
       sSql = sSql & "       MAT.* " & vbCrLf
       sSql = sSql & "      ,DPT.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADREQMAT MAT" & vbCrLf
       sSql = sSql & "      ,SGI_CADDEPTO  DPT" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       MAT.SGI_FILIAL   = " & FILIAL & vbCrLf
       sSql = sSql & "   And MAT.SGI_CODIGO   = " & txtCampos.Text
       sSql = sSql & "   And DPT.SGI_FILIAL   = MAT.SGI_FILIAL   " & vbCrLf
       sSql = sSql & "   And DPT.SGI_CODDEPTO = MAT.SGI_CODDEPTO "
       sSql = sSql & " Order by SGI_CODIGO "
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxReqMat.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & Format(BREC!SGI_DATREQ, "DD/MM/YYYY")
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxReqMat.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       MAT.* " & vbCrLf
       sSql = sSql & "      ,DPT.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADREQMAT MAT" & vbCrLf
       sSql = sSql & "      ,SGI_CADDEPTO  DPT" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       MAT.SGI_FILIAL    = " & FILIAL & vbCrLf
       sSql = sSql & "   And DPT.SGI_DESCRICAO Like '" & txtCampos.Text & "%'"
       sSql = sSql & "   And DPT.SGI_FILIAL    = MAT.SGI_FILIAL   " & vbCrLf
       sSql = sSql & "   And DPT.SGI_CODDEPTO  = MAT.SGI_CODDEPTO "
       sSql = sSql & " Order by SGI_CODIGO "
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxReqMat.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & Format(BREC!SGI_DATREQ, "DD/MM/YYYY")
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxReqMat.SetFocus
          Exit Sub
          
       End If
       
    ElseIf cboFiltro.ListIndex = 2 Then
    
       If IsDate(txtCampos.Text) = False Then
          MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       MAT.* " & vbCrLf
       sSql = sSql & "      ,DPT.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_CADREQMAT MAT" & vbCrLf
       sSql = sSql & "      ,SGI_CADDEPTO  DPT" & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       MAT.SGI_FILIAL    = " & FILIAL & vbCrLf
       sSql = sSql & "   And MAT.SGI_DATREQ    = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'"
       sSql = sSql & "   And DPT.SGI_FILIAL    = MAT.SGI_FILIAL   " & vbCrLf
       sSql = sSql & "   And DPT.SGI_CODDEPTO  = MAT.SGI_CODDEPTO "
       sSql = sSql & " Order by SGI_CODIGO "
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxReqMat.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO & vbTab & Format(BREC!SGI_DATREQ, "DD/MM/YYYY")
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxReqMat.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    
    ConfGrid
    PreencheGrid

End Sub

Private Function VerifReqSai() As Boolean

    VerifReqSai = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADITREQSAIMAT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODREQ = " & objCADREQMAT.CADREQCOD

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then VerifReqSai = True
    BREC.Close
    
End Function

Private Sub ImpReqMat(lngCODREQ As Long)
    
    If lngCODREQ = 0 Then
       MsgBox "Informe o Código da requisição !!!", vbExclamation + vbYesNo, "Aviso"
       Exit Sub
    End If
    
    Dim strCABEC2 As String
    Dim strCABEC3 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADREQMAT.SGI_CODIGO "
    sSql = sSql & "      ,SGI_CADREQMAT.SGI_DATREQ "
    sSql = sSql & "      ,SGI_CADREQMAT.SGI_CODDEPTO "
    sSql = sSql & "      ,SGI_CADREQMAT.SGI_CODUSUAR "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADREQMAT SGI_CADREQMAT"
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADREQMAT.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And SGI_CADREQMAT.SGI_CODIGO = " & lngCODREQ
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Requisição de material não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC2 = "Requisição de Materials"
    strCABEC3 = ""
    
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelRegMat & "RELREQMAT.rpt", Linha, 1, strCABEC2, strCABEC3, False
    
End Sub
