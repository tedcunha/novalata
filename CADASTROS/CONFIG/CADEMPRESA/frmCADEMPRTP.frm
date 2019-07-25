VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADEMPRTP 
   Caption         =   "Cadastro de Empresas"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   8535
      Begin MSFlexGridLib.MSFlexGrid flxEmpresa 
         Height          =   3375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
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
      TabIndex        =   1
      Top             =   600
      Width           =   8535
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3240
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
         Left            =   7680
         Picture         =   "frmCADEMPRTP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   6840
         Picture         =   "frmCADEMPRTP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   2280
         Picture         =   "frmCADEMPRTP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
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
         Picture         =   "frmCADEMPRTP.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Altera Empresa "
         Top             =   240
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
         Picture         =   "frmCADEMPRTP.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   240
         Width           =   735
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
         Picture         =   "frmCADEMPRTP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "txtCampos"
         Top             =   200
         Width           =   4455
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
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
         Left            =   3240
         TabIndex        =   5
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
Attribute VB_Name = "frmCADEMPRTP"
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
Dim objCADEMPR   As Object
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
  
  iResp = MsgBox("Confirma a exclisão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADEMPR.GRAVA("E") = False Then Exit Sub
  
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
    If flxEmpresa.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADEMPR = Nothing
    Unload Me
End Sub

Private Sub flxEmpresa_Click()
    If flxEmpresa.Rows > 1 Then
       objCADEMPR.EMPCOD = CInt(flxEmpresa.TextMatrix(flxEmpresa.RowSel, 1))
    End If
End Sub

Private Sub flxEmpresa_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxEmpresa.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
       If flxEmpresa.Rows > 1 Then Operacao "C"
    End If
End Sub

Private Sub flxEmpresa_RowColChange()
    If flxEmpresa.Rows > 1 Then
       objCADEMPR.EMPCOD = CInt(flxEmpresa.TextMatrix(flxEmpresa.RowSel, 1))
    End If
End Sub

Private Sub Form_Activate()
  If flxEmpresa.Enabled = True Then flxEmpresa.SetFocus
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
    Set objCADEMPR = CreateObject("CADEMPRESA.clsCADEMPRESA")
    
    objCADEMPR.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADEMPRTP
    
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
    cboFiltro.AddItem "CNPJ"
    
    objFuncoes.ChecaAcesso frmCADEMPRTP, strAcesso
    
    cboFiltro.ListIndex = 0

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()

    If objCADEMPR.Pesq_Empresa = False Then
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

    flxEmpresa.Rows = 1
    flxEmpresa.Cols = 5
    
    flxEmpresa.TextMatrix(0, 0) = ""
    flxEmpresa.TextMatrix(0, 1) = "Código"
    flxEmpresa.TextMatrix(0, 2) = "Descrição"
    flxEmpresa.TextMatrix(0, 3) = "CNPJ"
    flxEmpresa.TextMatrix(0, 4) = "Padrão"
    
    
    flxEmpresa.ColWidth(0) = 0
    flxEmpresa.ColWidth(1) = 700
    flxEmpresa.ColWidth(2) = 5000
    flxEmpresa.ColWidth(3) = 1500
    flxEmpresa.ColWidth(4) = 600
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select * from SGI_FILIAL " & vbCrLf
    sSql = sSql & "    Where " & vbCrLf
    sSql = sSql & "          SGI_FILIAL <> " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_FILIAL" & vbCrLf
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       
       flxEmpresa.AddItem "" & vbTab & _
                          BREC!SGI_FILIAL & vbTab & _
                          BREC!SGI_DESCRICAO & vbTab & _
                          BREC!SGI_CNPJ & vbTab & _
                          IIf(IsNull(BREC!SGI_PADRAO) = False, IIf(BREC!SGI_PADRAO = 1, "SIM", "NÃO"), "NÃO")
       
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub Operacao(Operacao As String)

  Dim Pesquisa As String
  
  If flxEmpresa.Rows > 1 Then
     iCodigo = CInt(flxEmpresa.TextMatrix(flxEmpresa.RowSel, 1))
  End If
  
  frmCADEMPRESA.iCodigo = iCodigo
  frmCADEMPRESA.cTipOper = Operacao
  frmCADEMPRESA.strAcesso = strAcesso
  frmCADEMPRESA.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid
  
  objFuncoes.ChecaAcesso frmCADEMPRTP, strAcesso

End Sub

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select * from SGI_FILIAL " & vbCrLf
     sSql = sSql & "         Where " & vbCrLf
     sSql = sSql & "               SGI_FILIAL <> " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_FILIAL"
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select * from SGI_FILIAL " & vbCrLf
     sSql = sSql & "         Where " & vbCrLf
     sSql = sSql & "               SGI_FILIAL <> " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_DESCRICAO"
  ElseIf cboFiltro.ListIndex = 2 Then
     sSql = " Select * from SGI_FILIAL " & vbCrLf
     sSql = sSql & "         Where " & vbCrLf
     sSql = sSql & "               SGI_FILIAL <> " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CNPJ"
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flxEmpresa.AddItem "" & vbTab & BREC!SGI_FILIAL & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_CNPJ
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
     
        For I = 1 To (flxEmpresa.Rows - 1)
             
            If flxEmpresa.TextMatrix(I, 1) = iCodigo Then
               flxEmpresa.Row = I
               flxEmpresa.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub


Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADEMPRTP
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
       sSql = sSql & "      SGI_FILIAL" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxEmpresa.AddItem "" & vbTab & BREC!SGI_FILIAL & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_CNPJ
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxEmpresa.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_FILIAL" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_DESCRICAO LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxEmpresa.AddItem "" & vbTab & BREC!SGI_FILIAL & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_CNPJ
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxEmpresa.SetFocus
          Exit Sub
          
       End If
    
    ElseIf cboFiltro.ListIndex = 2 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
             
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_FILIAL" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_CNPJ = '" & txtCampos.Text & "'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxEmpresa.AddItem "" & vbTab & BREC!SGI_FILIAL & vbTab & BREC!SGI_DESCRICAO & vbTab & BREC!SGI_CNPJ
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxEmpresa.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    ConfGrid
    PreencheGrid

End Sub
