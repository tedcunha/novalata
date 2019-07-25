VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADFERRAP 
   Caption         =   "Cadastro de Ferramentas"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flcCADFERRA 
         Height          =   4215
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7435
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
         Picture         =   "frmCADFERRAP.frx":0000
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
         Picture         =   "frmCADFERRAP.frx":0532
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
         Picture         =   "frmCADFERRAP.frx":0A64
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
         Picture         =   "frmCADFERRAP.frx":0B66
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
         Left            =   6120
         Picture         =   "frmCADFERRAP.frx":0C68
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
         Left            =   6960
         Picture         =   "frmCADFERRAP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3600
         Top             =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   4455
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
Attribute VB_Name = "frmCADFERRAP"
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
Dim objCADFERRA    As Object
Dim iCodigo        As Long

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
  
  If objCADFERRA.GRAVA("E") = False Then Exit Sub
  If objCADFERRA.Atualiza("E", Str(objCADFERRA.CODIGO), Filial, "frmCADFERRA") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Atualiza_Grid
  AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flcCADFERRA.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADFERRA = Nothing
    Unload Me
End Sub

Private Sub flcCADFERRA_Click()
    If flcCADFERRA.Rows > 1 Then objCADFERRA.CODIGO = CInt(flcCADFERRA.TextMatrix(flcCADFERRA.RowSel, 1))
End Sub

Private Sub flcCADFERRA_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flcCADFERRA.Rows > 1 Then Operacao "C"
End Sub

Private Sub flcCADFERRA_RowColChange()
    If flcCADFERRA.Rows > 1 Then objCADFERRA.CODIGO = CInt(flcCADFERRA.TextMatrix(flcCADFERRA.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADFERRA = CreateObject("CADFERRA.clsCADFERRA")
    
    objCADFERRA.Filial = Filial
    
    objFuncoes.LimpaCampos frmCADFERRAP
    
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

Private Sub ConfGrid()
    
    flcCADFERRA.Rows = 1
    flcCADFERRA.Cols = 3
    
    flcCADFERRA.TextMatrix(0, 1) = "Código"
    flcCADFERRA.TextMatrix(0, 2) = "Descrição"
    
    flcCADFERRA.ColWidth(0) = 0
    flcCADFERRA.ColWidth(1) = 1000
    flcCADFERRA.ColWidth(2) = 6000
    
End Sub

Private Sub AbilitaCampos()
    
    If objCADFERRA.Pesq_CadFerramenta = False Then
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

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFERRA " & vbCrLf
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flcCADFERRA.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           BREC!SGI_DESCRI
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub

Private Sub Operacao(strOperacao As String)
 
  If strOperacao = "A" Or strOperacao = "C" Then
     If Verif_reg = True Then Exit Sub
  End If
  
  Dim Pesquisa As String
  
  If flcCADFERRA.Rows > 1 Then iCodigo = CInt(flcCADFERRA.TextMatrix(flcCADFERRA.RowSel, 1))
  
  frmCADFERRA.cCaminho = cCaminho
  frmCADFERRA.Linha = Linha
  frmCADFERRA.iCodigo = iCodigo
  frmCADFERRA.cTipOper = strOperacao
  frmCADFERRA.Filial = Filial
  frmCADFERRA.strAcesso = strAcesso
  frmCADFERRA.Show vbModal
  
  AbilitaCampos
  Atualiza_Grid

End Sub

Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub

Private Function Verif_reg() As Boolean

    Verif_reg = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFERRA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & objCADFERRA.CODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Este registro foi excluso !!!", vbOKOnly + vbExclamation, "Aviso"
       Verif_reg = True
    End If
    BREC.Close

End Function

Private Sub Atualiza_Grid()
    
     Dim I        As Integer
     Dim bolAchou As Boolean
      
     bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADFERRA'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (flcCADFERRA.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If flcCADFERRA.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                  If flcCADFERRA.Rows = 2 Then flcCADFERRA.Rows = 1
                  If flcCADFERRA.Rows > 2 Then flcCADFERRA.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(flcCADFERRA.TextMatrix(I, 1)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADFERRA " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flcCADFERRA.AddItem "" & vbTab & _
                                  BREC2!SGI_CODIGO & vbTab & _
                                  BREC2!SGI_DESCRI
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADFERRA " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flcCADFERRA.TextMatrix(I, 0) = ""
              flcCADFERRA.TextMatrix(I, 1) = BREC2!SGI_CODIGO
              flcCADFERRA.TextMatrix(I, 2) = BREC2!SGI_DESCRI
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
      
End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADFERRAP
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
       sSql = sSql & "      SGI_CADFERRA " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & Filial & vbCrLf
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flcCADFERRA.AddItem "" & vbTab & _
                                 BREC!SGI_CODIGO & vbTab & _
                                 BREC!SGI_DESCRI
             BREC.MoveNext
          Loop
              
          BREC.Close
          flcCADFERRA.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADFERRA " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL =  " & Filial & vbCrLf
       sSql = sSql & "  And SGI_DESCRI LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flcCADFERRA.AddItem "" & vbTab & _
                                 BREC!SGI_CODIGO & vbTab & _
                                 BREC!SGI_DESCRI
             BREC.MoveNext
          Loop
              
          BREC.Close
          flcCADFERRA.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    
End Sub

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADFERRA " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADFERRA " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
     sSql = sSql & " Order by SGI_DESCRI "
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flcCADFERRA.AddItem "" & vbTab & _
                         BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_DESCRI
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub

