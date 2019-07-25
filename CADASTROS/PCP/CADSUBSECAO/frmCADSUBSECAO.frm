VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADSUBSECAO 
   Caption         =   "Cadastro de Sub-Seção"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stSubSecao 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4683
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
      TabPicture(0)   =   "frmCADSUBSECAO.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Maquinas"
      TabPicture(1)   =   "frmCADSUBSECAO.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   18
         Top             =   840
         Width           =   6255
         Begin MSFlexGridLib.MSFlexGrid flxMaquinas 
            Height          =   1455
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   2566
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
         End
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   6255
         Begin VB.ComboBox cboMaquina 
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Text            =   "cboMaquina"
            Top             =   120
            Width           =   3855
         End
         Begin VB.TextBox txtCODMAQ 
            Height          =   285
            Left            =   780
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "txtCODMAQ"
            Top             =   120
            Width           =   750
         End
         Begin VB.CommandButton cmdPesq 
            Height          =   315
            Left            =   1560
            Picture         =   "frmCADSUBSECAO.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmbGravEsp 
            Height          =   315
            Left            =   5760
            Picture         =   "frmCADSUBSECAO.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   6
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
            Left            =   120
            TabIndex        =   17
            Top             =   160
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   6255
         Begin VB.TextBox txtSigla 
            Height          =   285
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   1
            Text            =   "txtCodigo"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "txtDescricao"
            Top             =   960
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sigla:"
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
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   495
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
            Left            =   240
            TabIndex        =   15
            Top             =   960
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
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6615
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
         Picture         =   "frmCADSUBSECAO.frx":023C
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmCADSUBSECAO.frx":076E
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmCADSUBSECAO.frx":0870
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADSUBSECAO"
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
Dim objCADSUBSECAO As Object
Dim objPESQPADRAO  As Object
Dim arrMAQUINAS    As Variant

Private Sub cboMaquina_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboMaquina, KeyAscii
End Sub

Private Sub cboMaquina_Validate(Cancel As Boolean)
    If cboMaquina.ListIndex > -1 Then txtCODMAQ.Text = cboMaquina.ItemData(cboMaquina.ListIndex)
End Sub

Private Sub cmbGravEsp_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncMaquinas
End Sub

Private Sub cmdAltera_Click()
    
    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    stSubSecao.Tab = 0
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    txtDescricao.SetFocus
    
    cTipOper = "A"
    
    Me.Caption = "Cadastro Sub-Seção - [ ALTERAÇÃO ]"
    
End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Maquinas")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODMAQ.Text = varRETORNO
        
    cboMaquina.ListIndex = -1
    txtCODMAQ.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If Verifica_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADSUBSECAO.CODIGO = objCADSUBSECAO.Gera_Codigo(Me.Name)
    
    objCADSUBSECAO.DESCRI = txtDescricao.Text
    objCADSUBSECAO.SIGLA = txtSigla.Text
    
    If (flxMaquinas.Rows - 1) > 0 Then
       ReDim arrMAQUINAS(1 To (flxMaquinas.Rows - 1)) As String
       For I = 1 To (flxMaquinas.Rows - 1)
           arrMAQUINAS(I) = flxMaquinas.TextMatrix(I, 1)
       Next I
       objCADSUBSECAO.Maquinas = arrMAQUINAS
    Else
       ReDim arrMAQUINAS(0) As String
       objCADSUBSECAO.Maquinas = arrMAQUINAS
    End If
    
    '' Grava as informações
    If objCADSUBSECAO.GRAVA(cTipOper) = False Then Exit Sub
    
    MsgBox "A Sub-Seção foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    If objCADSUBSECAO.Atualiza(cTipOper, Str(objCADSUBSECAO.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
    
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADSUBSECAO = Nothing
       Set objPESQPADRAO = Nothing
       Unload Me
    End If
    
End Sub

Private Sub cmdVoltar_Click()
   Set objBLBFunc = Nothing
   Set objCADSUBSECAO = Nothing
   Set objPESQPADRAO = Nothing
   Unload Me
End Sub

Private Sub flxMaquinas_KeyDown(KeyCode As Integer, Shift As Integer)
    If (flxMaquinas.Rows - 1) = 0 Then Exit Sub
    If cTipOper = "C" Then Exit Sub
    If KeyCode = vbKeyDelete Then
       If flxMaquinas.Rows = 2 Then flxMaquinas.Rows = 1
       If flxMaquinas.Rows > 2 Then flxMaquinas.RemoveItem (flxMaquinas.RowSel)
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
   Set objCADSUBSECAO = CreateObject("CADSUBSECAO.clsCADSUBSECAO")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADSUBSECAO.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub


Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de Sub-Seção - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSUBSECAO
    
    stSubSecao.Tab = 0
    
    ConfGridMaq
    
    objCADSUBSECAO.PreencheComboMaquinas cboMaquina
   
End Sub

Private Sub ConfGridMaq()

    flxMaquinas.Rows = 1
    flxMaquinas.Cols = 3
    
    flxMaquinas.TextMatrix(0, 0) = ""
    flxMaquinas.TextMatrix(0, 1) = "Código"
    flxMaquinas.TextMatrix(0, 2) = "Descrição"
    
    flxMaquinas.ColWidth(0) = 0
    flxMaquinas.ColWidth(1) = 1000
    flxMaquinas.ColWidth(2) = 3000
    
End Sub

Private Sub txtCODMAQ_GotFocus()
    objBLBFunc.SelecionaCampos txtCODMAQ.Name, frmCADSUBSECAO
End Sub

Private Sub txtCODMAQ_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blACHOU As Boolean
    
    If Len(Trim(txtCODMAQ.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODMAQ.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCODMAQ.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    For I = 0 To (cboMaquina.ListCount - 1)
        If CInt(txtCODMAQ.Text) = cboMaquina.ItemData(I) Then cboMaquina.ListIndex = I
    Next I
    
    If cboMaquina.ListIndex = -1 Then
       MsgBox "Esta maquina não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCODMAQ.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub IncMaquinas()

    Dim I As Integer
    
    If Len(Trim(txtCODMAQ.Text)) = 0 Or cboMaquina.ListIndex = -1 Then
       MsgBox "Informe a maquina !!!", vbOKOnly + vbCritical, "aviso"
       txtCODMAQ.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxMaquinas.Rows - 1)
        If txtCODMAQ.Text = flxMaquinas.TextMatrix(I, 1) Then
           MsgBox "Esta maquina já foi inclusa !!!", vbOKOnly + vbCritical, "aviso"
           txtCODMAQ.Text = ""
           cboMaquina.ListIndex = -1
           txtCODMAQ.SetFocus
           Exit Sub
        End If
    Next I
    
    flxMaquinas.AddItem "" & vbTab & _
                        txtCODMAQ.Text & vbTab & _
                        cboMaquina.Text
                        
    txtCODMAQ.Text = ""
    cboMaquina.ListIndex = -1
    txtCODMAQ.SetFocus
    
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADSUBSECAO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function Verifica_Campos() As Boolean

    Verifica_Campos = False
    
    Dim I As Integer
    
    If Len(Trim(txtDescricao.Text)) = 0 Then
       MsgBox "Informe a descrição da Sub-Seção !!!", vbOKOnly + vbExclamation, "Aviso"
       stSubSecao.Tab = 0
       txtDescricao.SetFocus
       Exit Function
    End If
    
    'If (flxMaquinas.Rows - 1) = 0 Then
    '   MsgBox "Nenhuma máquina foi informada !!!", vbOKOnly + vbExclamation, "Aviso"
    '   stSubSecao.Tab = 1
    '   txtCODMAQ.SetFocus
    '   Exit Function
    'End If
    
    If cTipOper = "I" Then
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Está descrição da Sub-Seção já existe !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          stSubSecao.Tab = 0
          txtDescricao.SetFocus
          Exit Function
       End If
       BREC.Close
       
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_SIGLA  = '" & Trim(txtSigla.Text) & "'"
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Está Sigla da Sub-Seção já existe !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          stSubSecao.Tab = 0
          txtSigla.SetFocus
          Exit Function
       End If
       BREC.Close
       
       For I = 1 To (flxMaquinas.Rows - 1)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADSUBSECMAQ " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_CODMAQ = " & flxMaquinas.TextMatrix(I, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              MsgBox "Está maquina já esta cadastrada para outra seção !!!", vbOKOnly + vbExclamation, "Aviso"
              BREC.Close
              Exit Function
           End If
           BREC.Close
           
       Next I
       
    End If
    If cTipOper = "A" Then
    
       If objCADSUBSECAO.DESCRI <> txtDescricao.Text Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Está descrição da Sub-Seção já existe !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtDescricao.Text = objCADSUBSECAO.DESCRI
             stSubSecao.Tab = 0
             txtDescricao.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
       
       If objCADSUBSECAO.SIGLA <> txtSigla.Text Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_SIGLA  = '" & Trim(txtSigla.Text) & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Está Sigla da Sub-Seção já existe !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtSigla.Text = objCADSUBSECAO.SIGLA
             stSubSecao.Tab = 0
             txtDescricao.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
       
    End If
    
    Verifica_Campos = True

End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    Me.Caption = "Cadastro Sub-Seção - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADSUBSECAO
    
    objCADSUBSECAO.CODIGO = iCodigo
    
    stSubSecao.Tab = 0
    
    ConfGridMaq
    
    objCADSUBSECAO.PreencheComboMaquinas cboMaquina
    
    If objCADSUBSECAO.Carrega_campos = True Then
       
       txtCodigo.Text = objCADSUBSECAO.CODIGO
       txtDescricao.Text = objCADSUBSECAO.DESCRI
       txtSigla.Text = objCADSUBSECAO.SIGLA
       
       arrMAQUINAS = objCADSUBSECAO.Maquinas
       
       If IsArray(arrMAQUINAS) = True Then
          For I = 1 To UBound(arrMAQUINAS)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrMAQUINAS(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then
                 flxMaquinas.AddItem "" & vbTab & _
                                     arrMAQUINAS(I) & vbTab & _
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
    
    Me.Caption = "Cadastro Sub-Seção - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSUBSECAO
    
    objCADSUBSECAO.CODIGO = iCodigo
    
    stSubSecao.Tab = 0
    
    ConfGridMaq
    
    objCADSUBSECAO.PreencheComboMaquinas cboMaquina
    
    If objCADSUBSECAO.Carrega_campos = True Then
       
       txtCodigo.Text = objCADSUBSECAO.CODIGO
       txtDescricao.Text = objCADSUBSECAO.DESCRI
       txtSigla.Text = objCADSUBSECAO.SIGLA
       
       arrMAQUINAS = objCADSUBSECAO.Maquinas
       
       If IsArray(arrMAQUINAS) = True Then
          For I = 1 To UBound(arrMAQUINAS)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrMAQUINAS(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then
                 flxMaquinas.AddItem "" & vbTab & _
                                     arrMAQUINAS(I) & vbTab & _
                                     BREC!SGI_DESCRI
              End If
              BREC.Close
              
          Next I
       End If
    
    End If

End Sub

Private Sub txtSigla_GotFocus()
    objBLBFunc.SelecionaCampos txtSigla.Name, frmCADSUBSECAO
End Sub
