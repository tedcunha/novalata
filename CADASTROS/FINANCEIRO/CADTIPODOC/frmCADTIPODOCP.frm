VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADTIPODOCP 
   Caption         =   "cadastro de Tipo de Documentos"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxTipDoc 
         Height          =   3495
         Left            =   120
         TabIndex        =   13
         Top             =   120
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
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7935
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
         Left            =   6960
         Picture         =   "frmCADTIPODOCP.frx":0000
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
         Left            =   6120
         Picture         =   "frmCADTIPODOCP.frx":0102
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
         Picture         =   "frmCADTIPODOCP.frx":0634
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
         Picture         =   "frmCADTIPODOCP.frx":0736
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
         Picture         =   "frmCADTIPODOCP.frx":0838
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
         Picture         =   "frmCADTIPODOCP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADTIPODOCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public Filial      As Long
Public strAcesso   As String
Dim objFuncoes     As Object
Dim objCADTIPODOC  As Object
Dim iCodigo        As Long

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
  
  If objCADTIPODOC.GRAVA("E") = False Then Exit Sub
  
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
    If flxTipDoc.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADTIPODOC = Nothing
    Unload Me
End Sub

Private Sub flxTipDoc_Click()
    If flxTipDoc.Rows > 1 Then objCADTIPODOC.TIPDESPCOD = CLng(flxTipDoc.TextMatrix(flxTipDoc.RowSel, 1))
End Sub

Private Sub flxTipDoc_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxTipDoc.Rows > 1 Then Operacao "C"
End Sub
Private Sub flxTipDoc_RowColChange()
    If flxTipDoc.Rows > 1 Then objCADTIPODOC.TIPDESPCOD = CLng(flxTipDoc.TextMatrix(flxTipDoc.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADTIPODOC = CreateObject("CADTIPODOC.clsCADTIPODOC")
    
    objCADTIPODOC.Filial = Filial
    
    objFuncoes.LimpaCampos frmCADTIPODOCP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "N�o foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
    cboFiltro.AddItem "C�digo"
    cboFiltro.AddItem "Descri��o"
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub
Private Sub AbilitaCampos()
    
    If objCADTIPODOC.Pesq_CADTIPODOC = False Then
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
    
    flxTipDoc.Rows = 1
    flxTipDoc.Cols = 3
    
    flxTipDoc.TextMatrix(0, 1) = "C�digo"
    flxTipDoc.TextMatrix(0, 2) = "Descri��o"
    
    flxTipDoc.ColWidth(0) = 0
    flxTipDoc.ColWidth(1) = 1000
    flxTipDoc.ColWidth(2) = 5000
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADTIPODOC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxTipDoc.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxTipDoc.Rows - 1)
             
            If flxTipDoc.TextMatrix(I, 1) = iCodigo Then
               flxTipDoc.Row = I
               flxTipDoc.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub



Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADTIPODOCP
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
       sSql = sSql & "      SGI_CADTIPODOC " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & Filial
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxTipDoc.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxTipDoc.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADTIPODOC " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & Filial
       sSql = sSql & "  And SGI_DESCRICAO LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxTipDoc.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxTipDoc.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    
    ConfGrid
    PreencheGrid

End Sub

Private Sub Operacao(Operacao As String)
 
  Dim Pesquisa As String
  
  If flxTipDoc.Rows > 1 Then iCodigo = CInt(flxTipDoc.TextMatrix(flxTipDoc.RowSel, 1))
  
  frmCADTIPODOC.cCaminho = cCaminho
  frmCADTIPODOC.Linha = Linha
  frmCADTIPODOC.iCodigo = iCodigo
  frmCADTIPODOC.cTipOper = Operacao
  frmCADTIPODOC.Filial = Filial
  frmCADTIPODOC.strAcesso = strAcesso
  frmCADTIPODOC.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADTIPODOC " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL = " & Filial & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADTIPODOC " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL = " & Filial & vbCrLf
     sSql = sSql & " Order by SGI_DESCRICAO "
  End If
  
  BREC.Open sSql, adoBanco_Dados
  Do While Not BREC.EOF
     flxTipDoc.AddItem "" & vbTab & BREC!SGI_CODIGO & vbTab & BREC!SGI_DESCRICAO
     BREC.MoveNext
  Loop
  BREC.Close

End Sub

