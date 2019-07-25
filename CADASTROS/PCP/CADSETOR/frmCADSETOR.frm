VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADSETOR 
   Caption         =   "Cadastro de Setores"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab StSetores 
      Height          =   3135
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "frmCADSETOR.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Seção"
      TabPicture(1)   =   "frmCADSETOR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   6135
         Begin VB.ComboBox cboSecao 
            Height          =   315
            Left            =   1850
            TabIndex        =   3
            Text            =   "cboSecao"
            Top             =   120
            Width           =   3855
         End
         Begin VB.TextBox txtCODSECAO 
            Height          =   285
            Left            =   670
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txtCODSECA"
            Top             =   120
            Width           =   750
         End
         Begin VB.CommandButton cmdPesq 
            Height          =   315
            Left            =   1450
            Picture         =   "frmCADSETOR.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmbGravEsp 
            Height          =   315
            Left            =   5700
            Picture         =   "frmCADSETOR.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   10
            TabIndex        =   17
            Top             =   165
            Width           =   660
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   9
         Top             =   960
         Width           =   6135
         Begin MSFlexGridLib.MSFlexGrid flxSECAO 
            Height          =   1815
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   3201
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6135
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "txtDescricao"
            Top             =   600
            Width           =   4935
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   15
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6375
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
         Left            =   960
         Picture         =   "frmCADSETOR.frx":023C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   855
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
         Left            =   1800
         MaskColor       =   &H8000000F&
         Picture         =   "frmCADSETOR.frx":076E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
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
         Picture         =   "frmCADSETOR.frx":0870
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADSETOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Integer
Public FILIAL      As Integer
Public strAcesso   As String
Public strMODPAI   As String
Dim objBLBFunc     As Object
Dim objCADSETOR    As Object
Dim objPESQPADRAO  As Object
Dim arrSECAO       As Variant

Private Sub cboSecao_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSecao, KeyAscii
End Sub

Private Sub cboSecao_Validate(Cancel As Boolean)
    If cboSecao.ListIndex > -1 Then txtCODSECAO.Text = cboSecao.ItemData(cboSecao.ListIndex)
End Sub

Private Sub cmbGravEsp_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncSecao
End Sub

Private Sub cmdAltera_Click()

    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    StSetores.Tab = 0
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    txtDescricao.SetFocus
    
    cTipOper = "A"
    
    Me.Caption = "Cadastro de Setor - [ ALTERAÇÃO ]"

End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADSECAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Seção")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODSECAO.Text = varRETORNO
        
    cboSecao.ListIndex = -1
    txtCODSECAO.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If Verifica_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADSETOR.CODIGO = objCADSETOR.Gera_Codigo(Me.Name)
    
    objCADSETOR.DESCRI = txtDescricao.Text
    
    If (flxSECAO.Rows - 1) > 0 Then
       ReDim arrSECAO(1 To (flxSECAO.Rows - 1)) As String
       For I = 1 To (flxSECAO.Rows - 1)
           arrSECAO(I) = flxSECAO.TextMatrix(I, 1)
       Next I
       objCADSETOR.SECAO = arrSECAO
    Else
       ReDim arrSECAO(0) As String
       objCADSETOR.SECAO = arrSECAO
    End If
    
    '' Grava as informações
    If objCADSETOR.GRAVA(cTipOper) = False Then Exit Sub
    
    MsgBox "O Setor foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
    If objCADSETOR.Atualiza(cTipOper, Str(objCADSETOR.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
    
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADSETOR = Nothing
       Set objPESQPADRAO = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
   Set objBLBFunc = Nothing
   Set objCADSETOR = Nothing
   Set objPESQPADRAO = Nothing
   Unload Me
End Sub

Private Sub flxSECAO_KeyDown(KeyCode As Integer, Shift As Integer)
    If (flxSECAO.Rows - 1) = 0 Then Exit Sub
    If cTipOper = "C" Then Exit Sub
    If KeyCode = vbKeyDelete Then
       If flxSECAO.Rows = 2 Then flxSECAO.Rows = 1
       If flxSECAO.Rows > 2 Then flxSECAO.RemoveItem (flxSECAO.RowSel)
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
   Set objCADSETOR = CreateObject("CADSETOR.clsCADSETOR")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADSETOR.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub txtCODSECAO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODSECAO.Name, frmCADSETOR
End Sub

Private Sub txtCODSECAO_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtCODSECAO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODSECAO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSECAO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    For I = 0 To (cboSecao.ListCount - 1)
        If CInt(txtCODSECAO.Text) = cboSecao.ItemData(I) Then cboSecao.ListIndex = I
    Next I
    
    If cboSecao.ListIndex = -1 Then
       MsgBox "Esta Seção não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSECAO.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub IncSecao()

    Dim I As Integer
    
    If Len(Trim(txtCODSECAO.Text)) = 0 Or cboSecao.ListIndex = -1 Then
       MsgBox "Informe a Seção !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSECAO.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxSECAO.Rows - 1)
        If txtCODSECAO.Text = flxSECAO.TextMatrix(I, 1) Then
           MsgBox "Esta Seção já foi inclusa !!!", vbOKOnly + vbCritical, "aviso"
           txtCODSECAO.Text = ""
           cboSecao.ListIndex = -1
           txtCODSECAO.SetFocus
           Exit Sub
        End If
    Next I
    
    flxSECAO.AddItem "" & vbTab & _
                     txtCODSECAO.Text & vbTab & _
                     cboSecao.Text
                        
    txtCODSECAO.Text = ""
    cboSecao.ListIndex = -1
    txtCODSECAO.SetFocus
    
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADSETOR
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de Setores - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSETOR
    
    StSetores.Tab = 0
    
    ConfGridSecao
    
    objCADSETOR.PreencheComboSecao cboSecao
   
End Sub

Private Sub ConfGridSecao()

    flxSECAO.Rows = 1
    flxSECAO.Cols = 3
    
    flxSECAO.TextMatrix(0, 0) = ""
    flxSECAO.TextMatrix(0, 1) = "Código"
    flxSECAO.TextMatrix(0, 2) = "Descrição"
    
    flxSECAO.ColWidth(0) = 0
    flxSECAO.ColWidth(1) = 1000
    flxSECAO.ColWidth(2) = 3000
    
End Sub

Private Function Verifica_Campos() As Boolean

    Verifica_Campos = False
    
    Dim I As Integer
    Dim j As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtDescricao.Text)) = 0 Then
       MsgBox "Informe a descrição do Setor !!!", vbOKOnly + vbExclamation, "Aviso"
       StSetores.Tab = 0
       txtDescricao.SetFocus
       Exit Function
    End If
    
    'If (flxSECAO.Rows - 1) = 0 Then
    '   MsgBox "Nenhuma Seção foi informada !!!", vbOKOnly + vbExclamation, "Aviso"
    '   StSetores.Tab = 1
    '   txtCODSECAO.SetFocus
    '   Exit Function
    'End If
    
    If cTipOper = "I" Then
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADSETOR " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Está descrição do Setor já existe !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          StSetores.Tab = 0
          txtDescricao.SetFocus
          Exit Function
       End If
       BREC.Close
       
       For I = 1 To (flxSECAO.Rows - 1)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADITESET " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODSECAO = " & flxSECAO.TextMatrix(I, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              MsgBox "Esta Seção já este cadastrado para setor !!!", vbOKOnly + vbExclamation, "Aviso"
              BREC.Close
              Exit Function
           End If
           BREC.Close
           
       Next I
       
    End If
    If cTipOper = "A" Then
    
       If objCADSETOR.DESCRI <> txtDescricao.Text Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADSETOR " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Está descrição do Setor já existe !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtDescricao.Text = objCADSETOR.DESCRI
             StSetores.Tab = 0
             txtDescricao.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
       
       
       If IsArray(arrSECAO) Then
            For I = 1 To UBound(arrSECAO)
                blAchou = False
                For j = 1 To (flxSECAO.Rows - 1)
                    If arrSECAO(I) = flxSECAO.TextMatrix(I, 1) Then blAchou = True
                    If blAchou = False Then
                       
                       sSql = "Select " & vbCrLf
                       sSql = sSql & "       * " & vbCrLf
                       sSql = sSql & "  From " & vbCrLf
                       sSql = sSql & "       SGI_CADITESET " & vbCrLf
                       sSql = sSql & " Where " & vbCrLf
                       sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
                       sSql = sSql & "   And SGI_CODSECAO = " & arrSECAO(I)
                       
                       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
                       If Not BREC.EOF Then
                          MsgBox "Esta Seção já este cadastrado para setor !!!", vbOKOnly + vbExclamation, "Aviso"
                          BREC.Close
                          Exit Function
                       End If
                       BREC.Close
                       
                    End If
                Next j
            Next I
        End If
    End If
    
    Verifica_Campos = True

End Function

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Me.Caption = "Cadastro de Setor - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSETOR
    
    objCADSETOR.CODIGO = iCodigo
    
    StSetores.Tab = 0
    
    ConfGridSecao
    
    objBLBFunc.LimpaCampos frmCADSETOR
    
    objCADSETOR.PreencheComboSecao cboSecao
    
    If objCADSETOR.Carrega_campos = True Then
       
       txtCodigo.Text = objCADSETOR.CODIGO
       txtDescricao.Text = objCADSETOR.DESCRI
       
       arrSECAO = objCADSETOR.SECAO
       
       If IsArray(arrSECAO) = True Then
          For I = 1 To UBound(arrSECAO)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSECAO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrSECAO(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then
                 flxSECAO.AddItem "" & vbTab & _
                                  arrSECAO(I) & vbTab & _
                                  BREC!SGI_DESCRI
              End If
              BREC.Close
              
          Next I
       End If
    
    End If

End Sub


Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    Me.Caption = "Cadastro de Setor - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADSETORP
    
    objCADSETOR.CODIGO = iCodigo
    
    StSetores.Tab = 0
    
    ConfGridSecao
    
    objBLBFunc.LimpaCampos frmCADSETOR
    
    objCADSETOR.PreencheComboSecao cboSecao
    
    If objCADSETOR.Carrega_campos = True Then
       
       txtCodigo.Text = objCADSETOR.CODIGO
       txtDescricao.Text = objCADSETOR.DESCRI
       
       arrSECAO = objCADSETOR.SECAO
       
       If IsArray(arrSECAO) = True Then
          For I = 1 To UBound(arrSECAO)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSECAO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrSECAO(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then
                 flxSECAO.AddItem "" & vbTab & _
                                  arrSECAO(I) & vbTab & _
                                  BREC!SGI_DESCRI
              End If
              BREC.Close
              
          Next I
       End If
    
    End If

End Sub


