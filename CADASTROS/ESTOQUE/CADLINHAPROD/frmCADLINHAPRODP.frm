VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADLINHAPRODP 
   Caption         =   "Cadastro de Linha de Produto"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7950
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxCADLINHAPROD 
         Height          =   3615
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
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
         Picture         =   "frmCADLINHAPRODP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADLINHAPRODP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADLINHAPRODP.frx":0A64
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
         Left            =   2400
         Picture         =   "frmCADLINHAPRODP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADLINHAPRODP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADLINHAPRODP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3600
         Top             =   240
      End
   End
End
Attribute VB_Name = "frmCADLINHAPRODP"
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
Dim objFuncoes          As Object
Dim objCADLINHAPROD     As Object
Dim iCodigo             As Integer

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
  
  ''If Consistencia = False Then
  ''   MsgBox "Existe no cadastro de setor esta seção !!!", vbOKOnly + vbExclamation, "Aviso"
  ''   Exit Sub
  ''End If
  
  iResp = MsgBox("Confirma a Exclusão do Registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADLINHAPROD.GRAVA("E") = False Then Exit Sub
  If objCADLINHAPROD.Atualiza("E", Str(objCADLINHAPROD.CODIGO), FILIAL, "frmCADLINHAPROD") = False Then Exit Sub
  
  MsgBox "Registro Excluso com Sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Atualiza_Grid
  AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxCADLINHAPROD.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADLINHAPROD = Nothing
    Unload Me
End Sub

Private Sub flxCADLINHAPROD_Click()
    If flxCADLINHAPROD.Rows > 1 Then objCADLINHAPROD.CODIGO = CInt(flxCADLINHAPROD.TextMatrix(flxCADLINHAPROD.RowSel, 1))
End Sub

Private Sub flxCADLINHAPROD_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCADLINHAPROD.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCADLINHAPROD_RowColChange()
    If flxCADLINHAPROD.Rows > 1 Then objCADLINHAPROD.CODIGO = CInt(flxCADLINHAPROD.TextMatrix(flxCADLINHAPROD.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADLINHAPROD = CreateObject("CADLINHAPROD.clsCADLINHAPROD")
    
    objCADLINHAPROD.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADLINHAPRODP
    
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

    If objCADLINHAPROD.Carrega_CADLINHAPRODUTO = False Then
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
            
    flxCADLINHAPROD.Rows = 1
    flxCADLINHAPROD.Cols = 4
    
    flxCADLINHAPROD.TextMatrix(0, 0) = ""
    flxCADLINHAPROD.TextMatrix(0, 1) = ""
    flxCADLINHAPROD.TextMatrix(0, 2) = "Código"
    flxCADLINHAPROD.TextMatrix(0, 3) = "Descrição"
    
    flxCADLINHAPROD.ColWidth(0) = 0
    flxCADLINHAPROD.ColWidth(1) = 0
    flxCADLINHAPROD.ColWidth(2) = 1500
    flxCADLINHAPROD.ColWidth(3) = 5000
    
    flxCADLINHAPROD.ColAlignment(3) = 0
    
End Sub

Private Sub PreencheGrid()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       flxCADLINHAPROD.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           Format(BREC!SGI_CODLIN, "###000") & vbTab & _
                           BREC!SGI_DESCRI
       BREC.MoveNext
    Loop
    BREC.Close
    
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
     sSql = sSql & "   And SGI_MODULO = 'frmCADLINHAPROD'" & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (flxCADLINHAPROD.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If flxCADLINHAPROD.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                  If flxCADLINHAPROD.Rows = 2 Then flxCADLINHAPROD.Rows = 1
                  If flxCADLINHAPROD.Rows > 2 Then flxCADLINHAPROD.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(flxCADLINHAPROD.TextMatrix(I, 1)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxCADLINHAPROD.AddItem "" & vbTab & _
                                   BREC2!SGI_CODIGO & vbTab & _
                                   Format(BREC2!SGI_CODLIN, "###000") & vbTab & _
                                   BREC2!SGI_DESCRI
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxCADLINHAPROD.TextMatrix(I, 0) = ""
              flxCADLINHAPROD.TextMatrix(I, 1) = BREC2!SGI_CODIGO
              flxCADLINHAPROD.TextMatrix(I, 2) = Format(BREC2!SGI_CODLIN, "###000")
              flxCADLINHAPROD.TextMatrix(I, 3) = BREC2!SGI_DESCRI
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
End Sub


Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  sSql = ""
  
    sSql = " Select " & vbCrLf
    sSql = sSql & "        * " & vbCrLf
    sSql = sSql & "   from " & vbCrLf
    sSql = sSql & "        SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
  
  If cboFiltro.ListIndex = 0 Then
     sSql = sSql & " Order by SGI_CODLIN "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = sSql & " Order by SGI_DESCRI "
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flxCADLINHAPROD.AddItem "" & vbTab & _
                         BREC!SGI_CODIGO & vbTab & _
                         Format(BREC!SGI_CODLIN, "###000") & vbTab & _
                         BREC!SGI_DESCRI
     BREC.MoveNext
  Loop
  BREC.Close

End Sub


Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADLINHAPRODP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    Call ConfGrid
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL  =  " & FILIAL & vbCrLf
    
    If cboFiltro.ListIndex = 0 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Call cmdCanFiltro_Click
          Exit Sub
       End If
             
       sSql = sSql & "  And SGI_CODLIN = " & txtCampos.Text
    ElseIf cboFiltro.ListIndex = 1 Then
       sSql = sSql & "  And SGI_DESCRI LIKE '" & txtCampos.Text & "%'"
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       flxCADLINHAPROD.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           Format(BREC!SGI_CODLIN, "###000") & vbTab & _
                           BREC!SGI_DESCRI
       BREC.MoveNext
    Loop
    BREC.Close
    flxCADLINHAPROD.SetFocus
    
    Exit Sub

End Sub


Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
  
  If flxCADLINHAPROD.Rows > 1 Then iCodigo = CInt(flxCADLINHAPROD.TextMatrix(flxCADLINHAPROD.RowSel, 1))
  
  frmCADLINHAPROD.cCaminho = cCaminho
  frmCADLINHAPROD.Linha = Linha
  frmCADLINHAPROD.iCodigo = iCodigo
  frmCADLINHAPROD.cTipOper = strOperacao
  frmCADLINHAPROD.FILIAL = FILIAL
  frmCADLINHAPROD.strAcesso = strAcesso
  frmCADLINHAPROD.strMODPAI = Me.Name
  frmCADLINHAPROD.Show vbModal
  
  Atualiza_Grid
  AbilitaCampos

End Sub

