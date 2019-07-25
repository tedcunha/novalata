VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADDIASUTEISP 
   Caption         =   "Cadastro de Dias Úteis"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxDIASUTEIS 
         Height          =   3135
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5530
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
         Picture         =   "frmCADDIASUTEISP.frx":0000
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
         Picture         =   "frmCADDIASUTEISP.frx":0532
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
         Picture         =   "frmCADDIASUTEISP.frx":0A64
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
         Picture         =   "frmCADDIASUTEISP.frx":0B66
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
         Left            =   6120
         Picture         =   "frmCADDIASUTEISP.frx":0C68
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
         Left            =   6960
         Picture         =   "frmCADDIASUTEISP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3600
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmCADDIASUTEISP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public FILIAL        As Integer
Public strAcesso     As String
Dim objFUNCOES       As Object
Dim objCADDIASUTEIS  As Object
Dim iCodigo          As Integer
Private Sub cboFiltro_Validate(Cancel As Boolean)
    txtCampos.Text = ""
    txtCampos.SetFocus
    ConfGrid
    PreencheGrid
End Sub

Private Sub cmdAltera_Click()
    If objFUNCOES.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Operacao "A"
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    ConfGrid
    PreencheGrid
End Sub

Private Sub cmdExclui_Click()
  
  If objFUNCOES.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  If Verif_reg = True Then Exit Sub
  
  Dim iResp As Integer
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADDIASUTEIS.GRAVA("E") = False Then Exit Sub
  If objCADDIASUTEIS.Atualiza("E", Str(objCADDIASUTEIS.CODIGO), FILIAL, "frmCADDIASUTEIS") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Atualiza_Grid
  AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFUNCOES.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxDIASUTEIS.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFUNCOES = Nothing
    Set objCADDIASUTEIS = Nothing
    Unload Me
End Sub

Private Sub flxDIASUTEIS_Click()
    If flxDIASUTEIS.Rows > 1 Then objCADDIASUTEIS.CODIGO = CInt(flxDIASUTEIS.TextMatrix(flxDIASUTEIS.RowSel, 1))
End Sub

Private Sub flxDIASUTEIS_DblClick()
    If objFUNCOES.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxDIASUTEIS.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxDIASUTEIS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If objFUNCOES.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
       If flxDIASUTEIS.Rows > 1 Then Operacao "C"
    End If
End Sub

Private Sub flxDIASUTEIS_RowColChange()
    If flxDIASUTEIS.Rows > 1 Then objCADDIASUTEIS.CODIGO = CInt(flxDIASUTEIS.TextMatrix(flxDIASUTEIS.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFUNCOES = CreateObject("BLBCWS.clsFuncoes")
    Set objCADDIASUTEIS = CreateObject("CADDIASUTEIS.clsCADDIASUTEIS")
    
    objFUNCOES.LimpaCampos frmCADDIASUTEISP
    
    objCADDIASUTEIS.FILIAL = FILIAL
    
    Set adoBanco_Dados = objFUNCOES.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Descrição"
    cboFiltro.AddItem "Data"

    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub ConfGrid()

    flxDIASUTEIS.Rows = 1
    flxDIASUTEIS.Cols = 4
    
    flxDIASUTEIS.TextMatrix(0, 0) = ""
    flxDIASUTEIS.TextMatrix(0, 1) = "Código"
    flxDIASUTEIS.TextMatrix(0, 2) = "Descrição"
    flxDIASUTEIS.TextMatrix(0, 3) = "Data"
    
    flxDIASUTEIS.ColWidth(0) = 0
    flxDIASUTEIS.ColWidth(1) = 700
    flxDIASUTEIS.ColWidth(2) = 5000
    flxDIASUTEIS.ColWidth(3) = 1000
    
End Sub

Private Sub AbilitaCampos()

    If objCADDIASUTEIS.Pesq_Banco = False Then
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
    
    sSql = "Select * from SGI_CADDIASUTEIS " & vbCrLf
    sSql = sSql & " Where SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados
    
    Do While Not BREC.EOF
       flxDIASUTEIS.AddItem "" & vbTab & _
                            BREC!SGI_CODIGO & vbTab & _
                            BREC!SGI_DESCRI & vbTab & _
                            Format(BREC!SGI_DATENVENTO, "DD/MM/YYYY")
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()
    If iCodigo > 0 Then
        Dim I As Integer
        For I = 1 To (flxDIASUTEIS.Rows - 1)
            If flxDIASUTEIS.TextMatrix(I, 1) = iCodigo Then
               flxDIASUTEIS.Row = I
               Exit For
            End If
        Next I
    End If
End Sub


Private Sub Operacao(Operacao As String)

  Dim Pesquisa As String
  
  If flxDIASUTEIS.Rows > 1 Then iCodigo = CInt(flxDIASUTEIS.TextMatrix(flxDIASUTEIS.RowSel, 1))
  
  frmCADDIASUTEIS.iCodigo = iCodigo
  frmCADDIASUTEIS.cTipOper = Operacao
  frmCADDIASUTEIS.FILIAL = FILIAL
  frmCADDIASUTEIS.strAcesso = strAcesso
  frmCADDIASUTEIS.Show vbModal
  
  Atualiza_Grid
  AbilitaCampos

End Sub

Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
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
     sSql = sSql & "   And SGI_MODULO = 'frmCADDIASUTEIS'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        For I = 1 To (flxDIASUTEIS.Rows - 1)
            If Trim(BREC!SGI_ACAO) = "E" Then
               If flxDIASUTEIS.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                  If flxDIASUTEIS.Rows = 2 Then flxDIASUTEIS.Rows = 1
                  If flxDIASUTEIS.Rows > 2 Then flxDIASUTEIS.RemoveItem I
                  Exit For
               End If
            ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
               If Trim(BREC!SGI_CODIGO) = Trim(flxDIASUTEIS.TextMatrix(I, 1)) Then
                  bolAchou = True
                  Exit For
               End If
            End If
        Next I
            
        If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
            
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADDIASUTEIS " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxDIASUTEIS.AddItem "" & vbTab & _
                                   BREC2!SGI_CODIGO & vbTab & _
                                   BREC2!SGI_DESCRI & vbTab & _
                                   Format(BREC2!SGI_DATENVENTO, "DD/MM/YYYY")
           End If
           BREC2.Close
        
        ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADDIASUTEIS " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
              flxDIASUTEIS.TextMatrix(I, 0) = ""
              flxDIASUTEIS.TextMatrix(I, 1) = BREC2!SGI_CODIGO
              flxDIASUTEIS.TextMatrix(I, 2) = BREC2!SGI_DESCRI
              flxDIASUTEIS.TextMatrix(I, 3) = Format(BREC2!SGI_DATENVENTO, "DD/MM/YYYY")
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
      
End Sub


Private Function Verif_reg() As Boolean

    Verif_reg = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDIASUTEIS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & objCADDIASUTEIS.CODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Este registro foi excluso !!!", vbOKOnly + vbExclamation, "Aviso"
       Verif_reg = True
    End If
    BREC.Close

End Function


Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADDIASUTEIS " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADDIASUTEIS " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_DESCRI "
  ElseIf cboFiltro.ListIndex = 2 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADDIASUTEIS " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_DATENVENTO "
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     flxDIASUTEIS.AddItem "" & vbTab & _
                          BREC!SGI_CODIGO & vbTab & _
                          BREC!SGI_DESCRI & vbTab & _
                          Format(BREC!SGI_DATENVENTO, "DD/MM/YYYY")
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub

Private Sub txtCampos_GotFocus()
    objFUNCOES.SelecionaCampos txtCampos.Name, frmCADDIASUTEISP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFUNCOES.Maiuscula(KeyAscii)
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
       sSql = sSql & "      SGI_CADDIASUTEIS " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_CODIGO = " & txtCampos.Text
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxDIASUTEIS.AddItem "" & vbTab & _
                                  BREC!SGI_CODIGO & vbTab & _
                                  BREC!SGI_DESCRI & vbTab & _
                                  Format(BREC!SGI_DATENVENTO, "DD/MM/YYYY")
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxDIASUTEIS.SetFocus
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADDIASUTEIS " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL  =  " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_DESCRI  Like '" & txtCampos.Text & "%'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxDIASUTEIS.AddItem "" & vbTab & _
                                  BREC!SGI_CODIGO & vbTab & _
                                  BREC!SGI_DESCRI & vbTab & _
                                  Format(BREC!SGI_DATENVENTO, "DD/MM/YYYY")
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxDIASUTEIS.SetFocus
          Exit Sub
          
       End If
    
    ElseIf cboFiltro.ListIndex = 2 Then
       
       If Not IsDate(txtCampos.Text) Then
          MsgBox "Somente é permitido Data !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
       
       sSql = "Select" & vbCrLf
       sSql = sSql & "      *" & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CADDIASUTEIS " & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      SGI_FILIAL  =  " & FILIAL & vbCrLf
       sSql = sSql & "  And SGI_DATENVENTO = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'"
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          ConfGrid
              
          Do While Not BREC.EOF
             flxDIASUTEIS.AddItem "" & vbTab & _
                                  BREC!SGI_CODIGO & vbTab & _
                                  BREC!SGI_DESCRI & vbTab & _
                                  Format(BREC!SGI_DATENVENTO, "DD/MM/YYYY")
             BREC.MoveNext
          Loop
              
          BREC.Close
          flxDIASUTEIS.SetFocus
          Exit Sub
          
       End If
    End If

    BREC.Close
    
    ConfGrid
    PreencheGrid

End Sub
