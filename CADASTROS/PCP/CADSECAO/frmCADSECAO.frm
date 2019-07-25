VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADSECAO 
   Caption         =   "Cadastro de Seção"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6510
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stSECAO 
      Height          =   3375
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5953
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
      TabPicture(0)   =   "frmCADSECAO.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sub-Seção"
      TabPicture(1)   =   "frmCADSECAO.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   10
         Top             =   960
         Width           =   6135
         Begin MSFlexGridLib.MSFlexGrid flxSUBSECAO 
            Height          =   1935
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   3413
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   6135
         Begin VB.CommandButton cmbGravEsp 
            Height          =   315
            Left            =   5700
            Picture         =   "frmCADSECAO.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdPesq 
            Height          =   315
            Left            =   1450
            Picture         =   "frmCADSECAO.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtCODSUBSECAO 
            Height          =   285
            Left            =   670
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txtCODSUBS"
            Top             =   120
            Width           =   750
         End
         Begin VB.ComboBox cboSubSecao 
            Height          =   315
            Left            =   1850
            TabIndex        =   3
            Text            =   "cboSubSecao"
            Top             =   120
            Width           =   3855
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
            TabIndex        =   16
            Top             =   165
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   6135
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
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "txtDescricao"
            Top             =   600
            Width           =   4935
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
            TabIndex        =   13
            Top             =   600
            Width           =   930
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6375
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
         Picture         =   "frmCADSECAO.frx":023C
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADSECAO.frx":033E
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   960
         Picture         =   "frmCADSECAO.frx":0440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADSECAO"
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
Dim objCADSECAO    As Object
Dim objPESQPADRAO  As Object
Dim arrSUBSECAO    As Variant

Private Sub cboSubSecao_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSubSecao, KeyAscii
End Sub

Private Sub cboSubSecao_Validate(Cancel As Boolean)
    If cboSubSecao.ListIndex > -1 Then txtCODSUBSECAO.Text = cboSubSecao.ItemData(cboSubSecao.ListIndex)
End Sub

Private Sub cmbGravEsp_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncSubSecao
End Sub

Private Sub cmdAltera_Click()

    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    stSECAO.Tab = 0
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    txtDescricao.SetFocus
    
    cTipOper = "A"
    
    Me.Caption = "Cadastro de Seção - [ ALTERAÇÃO ]"

End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Sub-Seção")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODSUBSECAO.Text = varRETORNO
        
    cboSubSecao.ListIndex = -1
    txtCODSUBSECAO.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If Verifica_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADSECAO.CODIGO = objCADSECAO.Gera_Codigo(Me.Name)
    
    objCADSECAO.DESCRI = txtDescricao.Text
    
    If (flxSUBSECAO.Rows - 1) > 0 Then
       ReDim arrSUBSECAO(1 To (flxSUBSECAO.Rows - 1)) As String
       For I = 1 To (flxSUBSECAO.Rows - 1)
           arrSUBSECAO(I) = flxSUBSECAO.TextMatrix(I, 1)
       Next I
       objCADSECAO.SUBSECAO = arrSUBSECAO
    Else
       ReDim arrSUBSECAO(0) As String
       objCADSECAO.SUBSECAO = arrSUBSECAO
    End If
    
    '' Grava as informações
    If objCADSECAO.GRAVA(cTipOper) = False Then Exit Sub
    
    MsgBox "A Seção foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
    If objCADSECAO.Atualiza(cTipOper, Str(objCADSECAO.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
    
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADSECAO = Nothing
       Set objPESQPADRAO = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADSECAO = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub flxSUBSECAO_KeyDown(KeyCode As Integer, Shift As Integer)
    If (flxSUBSECAO.Rows - 1) = 0 Then Exit Sub
    If cTipOper = "C" Then Exit Sub
    If KeyCode = vbKeyDelete Then
       If flxSUBSECAO.Rows = 2 Then flxSUBSECAO.Rows = 1
       If flxSUBSECAO.Rows > 2 Then flxSUBSECAO.RemoveItem (flxSUBSECAO.RowSel)
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
   Set objCADSECAO = CreateObject("CADSECAO.clsCADSECAO")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADSECAO.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de Seção - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSECAO
    
    stSECAO.Tab = 0
    
    ConfGridSecao
    
    objCADSECAO.PreencheComboSubSecao cboSubSecao
   
End Sub

Private Sub ConfGridSecao()

    flxSUBSECAO.Rows = 1
    flxSUBSECAO.Cols = 3
    
    flxSUBSECAO.TextMatrix(0, 0) = ""
    flxSUBSECAO.TextMatrix(0, 1) = "Código"
    flxSUBSECAO.TextMatrix(0, 2) = "Descrição"
    
    flxSUBSECAO.ColWidth(0) = 0
    flxSUBSECAO.ColWidth(1) = 1000
    flxSUBSECAO.ColWidth(2) = 3000
    
End Sub

Private Sub txtCODSUBSECAO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODSUBSECAO.Name, frmCADSECAO
End Sub

Private Sub txtCODSUBSECAO_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtCODSUBSECAO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODSUBSECAO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSUBSECAO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    For I = 0 To (cboSubSecao.ListCount - 1)
        If CInt(txtCODSUBSECAO.Text) = cboSubSecao.ItemData(I) Then cboSubSecao.ListIndex = I
    Next I
    
    If cboSubSecao.ListIndex = -1 Then
       MsgBox "Esta Sub-Seção não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSUBSECAO.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub IncSubSecao()

    Dim I As Integer
    
    If Len(Trim(txtCODSUBSECAO.Text)) = 0 Or cboSubSecao.ListIndex = -1 Then
       MsgBox "Informe a Sub-Seção !!!", vbOKOnly + vbCritical, "aviso"
       txtCODSUBSECAO.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxSUBSECAO.Rows - 1)
        If txtCODSUBSECAO.Text = flxSUBSECAO.TextMatrix(I, 1) Then
           MsgBox "Esta Sub-Seção já foi inclusa !!!", vbOKOnly + vbCritical, "aviso"
           txtCODSUBSECAO.Text = ""
           cboSubSecao.ListIndex = -1
           txtCODSUBSECAO.SetFocus
           Exit Sub
        End If
    Next I
    
    flxSUBSECAO.AddItem "" & vbTab & _
                        txtCODSUBSECAO.Text & vbTab & _
                        cboSubSecao.Text
                        
    txtCODSUBSECAO.Text = ""
    cboSubSecao.ListIndex = -1
    txtCODSUBSECAO.SetFocus
    
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADSECAO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function Verifica_Campos() As Boolean

    Verifica_Campos = False
    
    Dim J As Integer
    Dim I As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtDescricao.Text)) = 0 Then
       MsgBox "Informe a descrição da Seção !!!", vbOKOnly + vbExclamation, "Aviso"
       stSECAO.Tab = 0
       txtDescricao.SetFocus
       Exit Function
    End If
    
    If cTipOper = "I" Then
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADSECAO " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Está descrição da Seção já existe !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          stSECAO.Tab = 0
          txtDescricao.SetFocus
          Exit Function
       End If
       BREC.Close
       
       For I = 1 To (flxSUBSECAO.Rows - 1)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADITESEC " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODSUBSECAO = " & flxSUBSECAO.TextMatrix(I, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              MsgBox "Está Sub-Seção já existe para outra seção !!!", vbOKOnly + vbExclamation, "Aviso"
              BREC.Close
              Exit Function
           End If
           BREC.Close
           
       Next I
       
    End If
    If cTipOper = "A" Then
    
       If objCADSECAO.DESCRI <> txtDescricao.Text Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADSECAO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Está descrição da Seção já existe !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtDescricao.Text = objCADSECAO.DESCRI
             stSECAO.Tab = 0
             txtDescricao.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
       
       '' Verifica se existe alguma nova Sub-Seção
       For I = 1 To (flxSUBSECAO.Rows - 1)
           If IsArray(arrSUBSECAO) = True Then
              For J = 1 To UBound(arrSUBSECAO)
                  blAchou = False
                  If flxSUBSECAO.TextMatrix(I, 1) = arrSUBSECAO(J) Then blAchou = True
                
                  If blAchou = False Then
                     sSql = "Select " & vbCrLf
                     sSql = sSql & "       * " & vbCrLf
                     sSql = sSql & "  From " & vbCrLf
                     sSql = sSql & "       SGI_CADITESEC " & vbCrLf
                     sSql = sSql & " Where " & vbCrLf
                     sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
                     sSql = sSql & "   And SGI_CODSUBSECAO = " & flxSUBSECAO.TextMatrix(I, 1)
           
                     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
                     If Not BREC.EOF Then
                        MsgBox "Está Sub-Seção já existe para outra seção !!!", vbOKOnly + vbExclamation, "Aviso"
                        BREC.Close
                        Exit Function
                     End If
                     BREC.Close
                  End If
           
              Next J
           End If
       Next I
       
    End If
    
    Verifica_Campos = True

End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    Me.Caption = "Cadastro de Seção - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADSECAO
    
    objCADSECAO.CODIGO = iCodigo
    
    stSECAO.Tab = 0
    
    ConfGridSecao
    
    objCADSECAO.PreencheComboSubSecao cboSubSecao
    
    If objCADSECAO.Carrega_campos = True Then
       
       txtCodigo.Text = objCADSECAO.CODIGO
       txtDescricao.Text = objCADSECAO.DESCRI
       
       arrSUBSECAO = objCADSECAO.SUBSECAO
       
       If IsArray(arrSUBSECAO) = True Then
          For I = 1 To UBound(arrSUBSECAO)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrSUBSECAO(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then
                 flxSUBSECAO.AddItem "" & vbTab & _
                                     arrSUBSECAO(I) & vbTab & _
                                     BREC!SGI_DESCRI
              End If
              BREC.Close
              
          Next I
       End If
    
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Me.Caption = "Cadastro de Seção - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSECAO
    
    objCADSECAO.CODIGO = iCodigo
    
    stSECAO.Tab = 0
    
    ConfGridSecao
    
    objCADSECAO.PreencheComboSubSecao cboSubSecao
    
    If objCADSECAO.Carrega_campos = True Then
       
       txtCodigo.Text = objCADSECAO.CODIGO
       txtDescricao.Text = objCADSECAO.DESCRI
       
       arrSUBSECAO = objCADSECAO.SUBSECAO
       
       If IsArray(arrSUBSECAO) = True Then
          For I = 1 To UBound(arrSUBSECAO)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrSUBSECAO(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then
                 flxSUBSECAO.AddItem "" & vbTab & _
                                     arrSUBSECAO(I) & vbTab & _
                                     BREC!SGI_DESCRI
              End If
              BREC.Close
              
          Next I
       End If
    
    End If

End Sub


