VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADOPERADORP 
   Caption         =   "Cadastro de Operadores"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   4095
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxCADOPERADOR 
         Height          =   3855
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6800
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
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   4440
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
         Picture         =   "frmCADOPERADORP.frx":0000
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
         Picture         =   "frmCADOPERADORP.frx":0102
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
         Picture         =   "frmCADOPERADORP.frx":0634
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
         Picture         =   "frmCADOPERADORP.frx":0736
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
         Picture         =   "frmCADOPERADORP.frx":0838
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
         Picture         =   "frmCADOPERADORP.frx":0D6A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADOPERADORP"
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
Dim objCADOPERADOR  As Object
Dim iCodigo         As Integer

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
  
  If Consistencia = False Then
     MsgBox "Existe no cadastro de maquinas este operador cadastrado !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  End If
  
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADOPERADOR.GRAVA("E") = False Then Exit Sub
  If objCADOPERADOR.Atualiza("E", Str(objCADOPERADOR.CODIGO), FILIAL, "frmCADOPERADOR") = False Then Exit Sub
  
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call Atualiza_Grid
  Call AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxCADOPERADOR.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADOPERADOR = Nothing
    Unload Me
End Sub

Private Sub flxCADOPERADOR_Click()
    If flxCADOPERADOR.Rows > 1 Then objCADOPERADOR.CODIGO = CInt(flxCADOPERADOR.TextMatrix(flxCADOPERADOR.RowSel, 1))
End Sub

Private Sub flxCADOPERADOR_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCADOPERADOR.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCADOPERADOR_RowColChange()
    If flxCADOPERADOR.Rows > 1 Then objCADOPERADOR.CODIGO = CInt(flxCADOPERADOR.TextMatrix(flxCADOPERADOR.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
   
    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADOPERADOR = CreateObject("CADOPERADOR.clsCADOPERADOR")
    
    objCADOPERADOR.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADOPERADORP
    
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
    
    If objCADOPERADOR.Pesq_CadOperador = False Then
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
    
    flxCADOPERADOR.Rows = 1
    flxCADOPERADOR.Cols = 4
    
    flxCADOPERADOR.TextMatrix(0, 0) = ""
    flxCADOPERADOR.TextMatrix(0, 1) = "Código"
    flxCADOPERADOR.TextMatrix(0, 2) = "Descrição"
    flxCADOPERADOR.TextMatrix(0, 3) = "Ativo"
    
    flxCADOPERADOR.ColWidth(0) = 0
    flxCADOPERADOR.ColWidth(1) = 1000
    flxCADOPERADOR.ColWidth(2) = 4000
    flxCADOPERADOR.ColWidth(3) = 600
    
End Sub

Private Sub PreencheGrid()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADOPERADOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       
       flxCADOPERADOR.AddItem "" & vbTab & _
                             BREC!SGI_CODIGO & vbTab & _
                             BREC!SGI_DESCRI & vbTab & _
                             IIf(BREC!SGI_ATIVO = 0, "Não", "Sim")
                               
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub


Private Sub PosGrid()
 
    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxCADOPERADOR.Rows - 1)
             
            If flxCADOPERADOR.TextMatrix(I, 1) = iCodigo Then
               flxCADOPERADOR.Row = I
               flxCADOPERADOR.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
    
  If flxCADOPERADOR.Rows > 1 Then iCodigo = CInt(flxCADOPERADOR.TextMatrix(flxCADOPERADOR.RowSel, 1))
    
  frmCADOPERADOR.cCaminho = cCaminho
  frmCADOPERADOR.Linha = Linha
  frmCADOPERADOR.iCodigo = iCodigo
  frmCADOPERADOR.cTipOper = strOperacao
  frmCADOPERADOR.FILIAL = FILIAL
  frmCADOPERADOR.strAcesso = strAcesso
  frmCADOPERADOR.Show vbModal
  
  Call Atualiza_Grid
  Call AbilitaCampos

End Sub


Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADOPERADOR " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADOPERADOR " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_DESCRI "
  End If
  
  BREC.Open sSql, adoBanco_Dados
  Do While Not BREC.EOF
     flxCADOPERADOR.AddItem "" & vbTab & _
                            BREC!SGI_CODIGO & vbTab & _
                            BREC!SGI_DESCRI & vbTab & _
                            IIf(BREC!SGI_ATIVO = 0, "Não", "Sim")
     BREC.MoveNext
  Loop
  BREC.Close

End Sub


Public Function Consistencia() As Boolean
    
    Consistencia = True
    
    '' Operador no cadastro de Maquina
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQOPER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODOPER  = " & objCADOPERADOR.CODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then Consistencia = False
    BREC.Close
    
End Function

Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub

Private Sub Atualiza_Grid()
    
     Dim I           As Integer
     Dim bolAchou    As Boolean
     Dim strRAZAOSOC As String
     Dim strDESCPROD As String
      
     bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADOPERADOR'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (flxCADOPERADOR.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If flxCADOPERADOR.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                  If flxCADOPERADOR.Rows = 2 Then flxCADOPERADOR.Rows = 1
                  If flxCADOPERADOR.Rows > 2 Then flxCADOPERADOR.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(flxCADOPERADOR.TextMatrix(I, 1)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "      *" & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "     SGI_CADOPERADOR " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "  And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
                
                flxCADOPERADOR.AddItem "" & vbTab & _
                                     BREC2!SGI_CODIGO & vbTab & _
                                     BREC2!SGI_DESCRI & vbTab & _
                                     IIf(BREC2!SGI_ATIVO = 0, "Não", "Sim")
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "      * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "      SGI_CADOPERADOR " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "  And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              
              flxCADOPERADOR.TextMatrix(I, 1) = BREC2!SGI_CODIGO
              flxCADOPERADOR.TextMatrix(I, 2) = BREC2!SGI_DESCRI
              flxCADOPERADOR.TextMatrix(I, 3) = IIf(BREC2!SGI_ATIVO = 0, "Não", "Sim")
           
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
      
End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADOPERADORP
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
       sSql = sSql & "      SGI_CADOPERADOR" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxCADOPERADOR.AddItem "" & vbTab & _
                                  BREC!SGI_CODIGO & vbTab & _
                                  BREC!SGI_DESCRI & vbTab & _
                                  IIf(BREC!SGI_ATIVO = 0, "Não", "Sim")
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxCADOPERADOR.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADOPERADOR" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL
       sSql = sSql & "  And SGI_DESCRI LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxCADOPERADOR.AddItem "" & vbTab & _
                                  BREC!SGI_CODIGO & vbTab & _
                                  BREC!SGI_DESCRI
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxCADOPERADOR.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    ConfGrid
    PreencheGrid

End Sub
