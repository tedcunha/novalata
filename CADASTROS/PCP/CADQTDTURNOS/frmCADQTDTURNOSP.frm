VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADQTDTURNOSP 
   Caption         =   "Cadastro de quantidade de turnos"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9180
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid flxQTDTURNOS 
         Height          =   4215
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7435
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
      Width           =   9135
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
         Width           =   5655
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
      Width           =   9135
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
         Picture         =   "frmCADQTDTURNOSP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADQTDTURNOSP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADQTDTURNOSP.frx":0A64
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
         Left            =   2640
         Picture         =   "frmCADQTDTURNOSP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   7320
         Picture         =   "frmCADQTDTURNOSP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   8160
         Picture         =   "frmCADQTDTURNOSP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3840
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmCADQTDTURNOSP"
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
Dim iCodigo         As Integer
Dim objFuncoes      As Object
Dim objCADQTDTURNOS As Object
Private Sub cboFiltro_Validate(Cancel As Boolean)
    txtCampos.Text = ""
    txtCampos.SetFocus
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
  
  If Consistencia = False Then
     MsgBox "Existe no cadastro de maquinas este turno cadastrado !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  End If
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADQTDTURNOS.GRAVA("E") = False Then Exit Sub
  If objCADQTDTURNOS.Atualiza("E", Str(objCADQTDTURNOS.CODIGO), FILIAL, "frmCADQTDTURNOS") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Call Atualiza_Grid
  Call AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxQTDTURNOS.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADQTDTURNOS = Nothing
    Unload Me
End Sub

Private Sub flxQTDTURNOS_Click()
    If flxQTDTURNOS.Rows > 1 Then objCADQTDTURNOS.CODIGO = CInt(flxQTDTURNOS.TextMatrix(flxQTDTURNOS.RowSel, 1))
End Sub

Private Sub flxQTDTURNOS_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxQTDTURNOS.Rows > 1 Then Operacao "C"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADQTDTURNOS = CreateObject("CADQTDTURNOS.clsCADQTDTURNOS")
    
    objCADQTDTURNOS.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADQTDTURNOSP
    
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
    
    If objCADQTDTURNOS.Pesq_CadQtdTurnos = False Then
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
    
    flxQTDTURNOS.Rows = 1
    flxQTDTURNOS.Cols = 4
    
    flxQTDTURNOS.TextMatrix(0, 0) = ""
    flxQTDTURNOS.TextMatrix(0, 1) = "Código"
    flxQTDTURNOS.TextMatrix(0, 2) = "Descrição"
    flxQTDTURNOS.TextMatrix(0, 3) = "Ativo"
    
    flxQTDTURNOS.ColWidth(0) = 0
    flxQTDTURNOS.ColWidth(1) = 700
    flxQTDTURNOS.ColWidth(2) = 5000
    flxQTDTURNOS.ColWidth(3) = 1000
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxQTDTURNOS.AddItem "" & vbTab & _
                            BREC!SGI_CODIGO & vbTab & _
                            BREC!SGI_DESCRI & vbTab & _
                            IIf(BREC!SGI_ATIVO = 0, "Nào", "Sim")
                            
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxQTDTURNOS.Rows - 1)
             
            If flxQTDTURNOS.TextMatrix(I, 1) = iCodigo Then
               flxQTDTURNOS.Row = I
               flxQTDTURNOS.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
On Error GoTo Err_Teste
  
  
  Dim Pesquisa As String
    
  If flxQTDTURNOS.Rows > 1 Then iCodigo = CInt(flxQTDTURNOS.TextMatrix(flxQTDTURNOS.RowSel, 1))
    
  frmCADQTDTURNOS.cCaminho = cCaminho
  frmCADQTDTURNOS.Linha = Linha
  frmCADQTDTURNOS.iCodigo = iCodigo
  frmCADQTDTURNOS.cTipOper = strOperacao
  frmCADQTDTURNOS.FILIAL = FILIAL
  frmCADQTDTURNOS.strAcesso = strAcesso
  frmCADQTDTURNOS.Show vbModal
  
  Call Atualiza_Grid
  Call AbilitaCampos
  
  Exit Sub
  
Err_Teste:

End Sub

Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select * from SGI_CADQTDETURN " & vbCrLf
     sSql = sSql & "         Where " & vbCrLf
     sSql = sSql & "               SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO"
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select * from SGI_CADQTDETURN " & vbCrLf
     sSql = sSql & "         Where " & vbCrLf
     sSql = sSql & "               SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_DESCRI"
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flxQTDTURNOS.AddItem "" & vbTab & _
                          BREC!SGI_CODIGO & vbTab & _
                          BREC!SGI_DESCRI & vbTab & _
                          IIf(BREC!SGI_ATIVO = 0, "Não", "Sim")
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub

Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub

Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADQTDTURNOSP
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
       sSql = sSql & "      SGI_CADQTDETURN" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxQTDTURNOS.AddItem "" & vbTab & _
                                  BREC!SGI_CODIGO & vbTab & _
                                  BREC!SGI_DESCRI & vbTab & _
                                  IIf(BREC!SGI_ATIVO = 0, "Não", "Sim")
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxQTDTURNOS.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADQTDETURN" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL
       sSql = sSql & "  And SGI_DESCRI LIKE '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxQTDTURNOS.AddItem "" & vbTab & _
                                  BREC!SGI_CODIGO & vbTab & _
                                  BREC!SGI_DESCRI
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxQTDTURNOS.SetFocus
          Exit Sub
          
       End If
    
    End If

    BREC.Close
    ConfGrid
    PreencheGrid

End Sub

Public Function Consistencia() As Boolean
    
    Consistencia = True
    
    '' Qtde de Turnos no cadastro de Maquina
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQTURN " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODTURN  = " & objCADQTDTURNOS.CODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then Consistencia = False
    BREC.Close
    
End Function

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
     sSql = sSql & "   And SGI_MODULO = 'frmCADQTDTURNOS'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (flxQTDTURNOS.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If flxQTDTURNOS.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                  If flxQTDTURNOS.Rows = 2 Then flxQTDTURNOS.Rows = 1
                  If flxQTDTURNOS.Rows > 2 Then flxQTDTURNOS.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(flxQTDTURNOS.TextMatrix(I, 1)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "      *" & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "     SGI_CADQTDETURN " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "  And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
                
                flxQTDTURNOS.AddItem "" & vbTab & _
                                     BREC2!SGI_CODIGO & vbTab & _
                                     BREC2!SGI_DESCRI & vbTab & _
                                     IIf(BREC2!SGI_ATIVO = 0, "Não", "Sim")
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "      * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "      SGI_CADQTDETURN " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "  And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              
              flxQTDTURNOS.TextMatrix(I, 1) = BREC2!SGI_CODIGO
              flxQTDTURNOS.TextMatrix(I, 2) = BREC2!SGI_DESCRI
              flxQTDTURNOS.TextMatrix(I, 3) = IIf(BREC2!SGI_ATIVO = 0, "Não", "Sim")
           
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
      
End Sub


