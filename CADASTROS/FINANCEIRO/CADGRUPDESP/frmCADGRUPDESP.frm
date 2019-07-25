VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmCADGRUPDESP 
   Caption         =   "Cadastro de grupo de despesas"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   3135
      Left            =   0
      TabIndex        =   15
      Top             =   2760
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid flxFornec 
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fornecedores"
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   2040
      Width           =   6255
      Begin VB.CommandButton cmbGravPGT 
         Height          =   315
         Left            =   5760
         Picture         =   "frmCADGRUPDESP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCODFORNEC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCODFORNEC"
         Top             =   255
         Width           =   735
      End
      Begin VB.ComboBox cboFornec 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Text            =   "cboFornec"
         Top             =   255
         Width           =   4575
      End
      Begin VB.CommandButton cmdPesqFor 
         Height          =   315
         Left            =   840
         Picture         =   "frmCADGRUPDESP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   6255
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Volta"
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
         Picture         =   "frmCADGRUPDESP.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
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
         Left            =   1920
         Picture         =   "frmCADGRUPDESP.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
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
         Left            =   1080
         Picture         =   "frmCADGRUPDESP.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADGRUPDESP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Long
Public Filial      As Integer
Public strAcesso   As String
Public strUSUARIO  As String
Dim objBLBFunc     As Object
Dim objCADGRUPDESP As Object
Dim objPESQPADRAO  As Object
Dim arrGRPDESP     As Variant
Dim arrGRPDESPBKP  As Variant


Private Sub cboFornec_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFornec, KeyAscii
End Sub

Private Sub cboFornec_Validate(Cancel As Boolean)
       If cboFornec.ListIndex > -1 Then txtCODFORNEC.Text = cboFornec.ItemData(cboFornec.ListIndex)
End Sub

Private Sub cmbGravPGT_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridFornec
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de grupo de despesas - [ ALTERAÇÃO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus
    

End Sub

Private Sub cmdPesqFor_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
    sSql = sSql & "   And SGI_STATUS = 2" & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ/CPF"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, Filial, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Fornecedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFORNEC.Text = varRETORNO
    
    cboFornec.ListIndex = -1
    txtCODFORNEC.SetFocus


End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = True Then
       
       If cTipOper = "I" Then objCADGRUPDESP.GRPDESPCCOD = objCADGRUPDESP.Gera_Codigo(Me.Name)
       
       objCADGRUPDESP.GRPDESPDESC = txtDescricao.Text
       
       If (flxFornec.Rows - 1) > 0 Then
          ReDim arrGRPDESP(1 To (flxFornec.Rows - 1)) As String
          For I = 1 To (flxFornec.Rows - 1)
              arrGRPDESP(I) = flxFornec.TextMatrix(I, 1)
          Next I
          objCADGRUPDESP.FORNEC = arrGRPDESP
       End If
       
       If objCADGRUPDESP.GRAVA(cTipOper) = True Then
          MsgBox "A grupo de despesas foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objCADGRUPDESP = Nothing
             Unload Me
          End If
       End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADGRUPDESP = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub flxFornec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxFornec.Rows = 2 Then flxFornec.Rows = 1
       If flxFornec.Rows > 2 Then flxFornec.RemoveItem flxFornec.RowSel
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADGRUPDESP = CreateObject("CADGRUPDESP.clsCADGRUPDESP")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADGRUPDESP.Filial = Filial
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de grupo de despesas - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADGRUPDESP
    
    txtCodigo.Text = ""
     
    objCADGRUPDESP.PreencheComboFornec cboFornec
    
    ConfGridFornec
   
End Sub

Private Sub Consulta()

    Dim I           As Integer
    Dim strRAZAOSOC As String
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
   
    Me.Caption = "Cadastro de grupo de despesas - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADGRUPDESP
    
    objCADGRUPDESP.GRPDESPCCOD = iCodigo
    
    objCADGRUPDESP.PreencheComboFornec cboFornec
    
    ConfGridFornec
    
    If objCADGRUPDESP.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADGRUPDESP.GRPDESPCCOD)
       txtDescricao.Text = objCADGRUPDESP.GRPDESPDESC
       
       arrGRPDESP = objCADGRUPDESP.FORNEC
       arrGRPDESPBKP = objCADGRUPDESP.FORNECBKP
       
       If IsArray(arrGRPDESP) Then
          strRAZAOSOC = ""
          For I = 1 To UBound(arrGRPDESP)
              
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & Trim(arrGRPDESP(I))
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then strRAZAOSOC = Trim(BREC!SGI_RAZAOSOC)
              BREC.Close
          
              flxFornec.AddItem "" & vbTab & _
                                arrGRPDESP(I) & vbTab & _
                                Trim(strRAZAOSOC)
          Next I
       End If
       
    End If

End Sub

Public Sub Altera()

    Dim I           As Integer
    Dim strRAZAOSOC As String
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Me.Caption = "Cadastro de grupo de despesas - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADGRUPDESP
    
    objCADGRUPDESP.GRPDESPCCOD = iCodigo
    
    objCADGRUPDESP.PreencheComboFornec cboFornec
    
    ConfGridFornec
    
    If objCADGRUPDESP.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADGRUPDESP.GRPDESPCCOD)
       txtDescricao.Text = objCADGRUPDESP.GRPDESPDESC
       
       arrGRPDESP = objCADGRUPDESP.FORNEC
       arrGRPDESPBKP = objCADGRUPDESP.FORNECBKP
       
       If IsArray(arrGRPDESP) Then
          strRAZAOSOC = ""
          For I = 1 To UBound(arrGRPDESP)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & Filial & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrGRPDESP(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then strRAZAOSOC = Trim(BREC!SGI_RAZAOSOC)
              BREC.Close
          
              flxFornec.AddItem "" & vbTab & _
                                arrGRPDESP(I) & vbTab & _
                                Trim(strRAZAOSOC)
          Next I
       End If
    
    End If
    
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Grupo de despesa inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADGRUPDESP " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRICAO = '" & txtDescricao.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL    = " & Filial
        
        BREC.Open sSql, adoBanco_Dados
        If Not BREC.EOF Then
           MsgBox "Grupo de despesa já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
                
        If objCADGRUPDESP.GRPDESPDESC <> txtDescricao.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADGRUPDESP " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRICAO = '" & txtDescricao.Text & "'" & vbCrLf
           sSql = sSql & "    And SGI_FILIAL   =  " & Filial
           
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Grupo de despesa já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADGRUPDESP.GRPDESPDESC
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub txtCODFORNEC_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFORNEC.Name, frmCADGRUPDESP
End Sub

Private Sub txtCODFORNEC_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFORNEC.Text
End Sub

Private Sub txtCODFORNEC_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFORNEC.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFORNEC.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNEC.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFornec.ListIndex = -1
    For I = 0 To (cboFornec.ListCount - 1)
        If cboFornec.ItemData(I) = Str(Val(txtCODFORNEC.Text)) Then cboFornec.ListIndex = I
    Next I
    
    If cboFornec.ListIndex = -1 Then
       MsgBox "Esta fornecedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNEC.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADGRUPDESP
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub ConfGridFornec()

    flxFornec.Rows = 1
    flxFornec.Cols = 3
    
    flxFornec.TextMatrix(0, 0) = ""
    flxFornec.TextMatrix(0, 1) = "Código"
    flxFornec.TextMatrix(0, 2) = "Razão Social"
    
    flxFornec.ColWidth(0) = 0
    flxFornec.ColWidth(1) = 500
    flxFornec.ColWidth(2) = 5000
    
End Sub

Private Sub IncGridFornec()

    Dim I As Integer
    
    If (cboFornec.ListIndex = -1) Then
       MsgBox "Informe o fornecedor !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCODFORNEC.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxFornec.Rows - 1)
        If Trim(flxFornec.TextMatrix(I, 1)) = Trim(txtCODFORNEC.Text) Then
           MsgBox "Este fornecedor já está relacionada na grid !!!", vbOKOnly + vbExclamation, "Aviso"
           txtCODFORNEC.Text = ""
           cboFornec.ListIndex = -1
           txtCODFORNEC.SetFocus
           Exit Sub
        End If
    Next I
    
    '' -----------------------------------
    If ConfereGrupo(CInt(txtCODFORNEC.Text)) = True Then
       MsgBox "Este fornecedor já está relacionado em outro grupo de despesas !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFORNEC.Text = ""
       cboFornec.ListIndex = -1
       txtCODFORNEC.SetFocus
       Exit Sub
    End If
    
    flxFornec.AddItem "" & vbTab & _
                      txtCODFORNEC.Text & vbTab & _
                      cboFornec.Text

    txtCODFORNEC.Text = ""
    cboFornec.ListIndex = -1
    txtCODFORNEC.SetFocus
    
End Sub

Private Function ConfereGrupo(intCODFORN As Integer) As Boolean
    
    ConfereGrupo = False
    
    Dim I         As Integer
    Dim boolAchou As Boolean
    
    If IsArray(arrGRPDESPBKP) And cTipOper = "A" Then
       
       boolAchou = False
       For I = 1 To UBound(arrGRPDESPBKP)
           If arrGRPDESPBKP(I) = intCODFORN Then boolAchou = True
       Next I
    
       If boolAchou = False Then
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_GRPDESPFORN " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL    = " & Filial & vbCrLf
          sSql = sSql & "   And SGI_CODFORN   = " & intCODFORN
    
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then ConfereGrupo = True
          BREC.Close
       End If
    Else
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_GRPDESPFORN " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL    = " & Filial & vbCrLf
       sSql = sSql & "   And SGI_CODFORN   = " & intCODFORN

       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then ConfereGrupo = True
       BREC.Close
    End If
    
End Function
