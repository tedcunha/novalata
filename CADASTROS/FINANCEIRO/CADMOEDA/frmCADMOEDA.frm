VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADMOEDA 
   Caption         =   "Cadastro de Moedas"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4515
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   1920
      Width           =   4455
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   3960
         Picture         =   "frmCADMOEDA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtVlIndice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "txtVlIndice"
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtDtIndice 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indice:"
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
         Index           =   5
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Indice ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   18
      Top             =   2520
      Width           =   4455
      Begin MSFlexGridLib.MSFlexGrid flxINDICES 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4260
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   4455
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtCAPMIN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Text            =   "txtCAPMIN"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox cboUnidade 
         Height          =   315
         Left            =   3600
         TabIndex        =   11
         Text            =   "cboUnidade"
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Index           =   3
         Left            =   2280
         TabIndex        =   17
         Top             =   960
         Width           =   75
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   120
         TabIndex        =   16
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
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCodigo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Capacidade Por Minuto"
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
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
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
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4455
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
         Picture         =   "frmCADMOEDA.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   735
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
         Left            =   1560
         Picture         =   "frmCADMOEDA.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   840
         Picture         =   "frmCADMOEDA.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADMOEDA"
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
Dim objBLBFunc     As Object
Dim objCADMOEDA    As Object
Dim arrVLINDICE    As Variant

Private Sub cmdAltera_Click()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = True
    
    Me.Caption = "Cadastro de Moedas - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADMOEDA.CODIGO = objCADMOEDA.Gera_Codigo(Me.Name)
    
    objCADMOEDA.DESCRI = txtDescricao.Text
    
    '' Indices de Moeda
    If (flxINDICES.Rows - 1) > 0 Then
       ReDim arrVLINDICE(1 To (flxINDICES.Rows - 1), 1 To 2) As String
       For I = 1 To (flxINDICES.Rows - 1)
           arrVLINDICE(I, 1) = flxINDICES.TextMatrix(I, 1)
           arrVLINDICE(I, 2) = flxINDICES.TextMatrix(I, 2)
       Next I
    Else
       ReDim arrVLINDICE(0) As String
    End If
    objCADMOEDA.INDICE = arrVLINDICE
    
    If objCADMOEDA.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O indice da moeda foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADMOEDA = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADMOEDA = Nothing
    Unload Me
End Sub

Private Sub Command5_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridIndice
End Sub

Private Sub flxINDICES_KeyDown(KeyCode As Integer, Shift As Integer)
    If cTipOper = "I" Or cTipOper = "A" Then
       If KeyCode <> vbKeyDelete Then Exit Sub
       If flxINDICES.Rows = 2 Then flxINDICES.Rows = 1
       If flxINDICES.Rows > 2 Then flxINDICES.RemoveItem flxINDICES.RowSel
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
   Set objCADMOEDA = CreateObject("CADMOEDA.clsCADMOEDA")
   
   objCADMOEDA.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub


Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de Moedas - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADMOEDA
    
    lblCodigo.Caption = ""
    txtDtIndice.Text = "__/__/____"
    
    Call ConfGridIndices
    
End Sub


Private Sub ConfGridIndices()

    flxINDICES.Rows = 1
    flxINDICES.Cols = 3
    
    flxINDICES.TextMatrix(0, 0) = ""
    flxINDICES.TextMatrix(0, 1) = "Dt. Indice"
    flxINDICES.TextMatrix(0, 2) = "Vl. Indice"
    
    flxINDICES.ColWidth(0) = 0
    flxINDICES.ColWidth(1) = 1000
    flxINDICES.ColWidth(2) = 1000
    
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADMOEDA
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtDtIndice_GotFocus()
    objBLBFunc.SelecionaCampos txtDtIndice.Name, frmCADMOEDA
End Sub

Private Sub txtVlIndice_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtVlIndice.Text
End Sub

Private Sub txtVlIndice_Validate(Cancel As Boolean)

    If Len(Trim(txtVlIndice.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtVlIndice.Text) Then
       MsgBox "Somente é Permitido Numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If
    
    txtVlIndice.Text = Format(CCur(txtVlIndice.Text), "#,##0.00")

End Sub

Private Sub IncGridIndice()

    Dim I As Integer
    
    If IsDate(txtDtIndice.Text) = False Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       txtDtIndice.Text = "__/__/____"
       txtDtIndice.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtVlIndice.Text)) = 0 Then
       MsgBox "Informe o Valor do Indice !!!", vbOKOnly + vbExclamation, "Aviso"
       txtVlIndice.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxINDICES.Rows - 1)
        If CDate(flxINDICES.TextMatrix(I, 1)) = CDate(txtDtIndice.Text) Then
           MsgBox "Esta data já foi lançada na Grid !!!", vbOKOnly + vbExclamation, "Aviso"
           txtDtIndice.Text = "__/__/____"
           txtDtIndice.SetFocus
           Exit Sub
        End If
    Next I
    
    flxINDICES.AddItem "" & vbTab & _
                       txtDtIndice.Text & vbTab & _
                       txtVlIndice.Text
                       
    txtDtIndice.Text = "__/__/____"
    txtVlIndice.Text = ""
    txtDtIndice.SetFocus
    
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "descrição da máquina inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function


Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = True
    Frame4.Enabled = False
   
    Me.Caption = "Cadastro de Moedas - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADMOEDA
    objCADMOEDA.CODIGO = iCodigo
    
    ConfGridIndices
    
    If objCADMOEDA.Carrega_campos = True Then
    
       lblCodigo.Caption = Str(objCADMOEDA.CODIGO)
       txtDescricao.Text = objCADMOEDA.DESCRI
       arrVLINDICE = objCADMOEDA.INDICE
       
       If IsArray(arrVLINDICE) Then
          For I = 1 To UBound(arrVLINDICE)
              flxINDICES.AddItem "" & vbTab & arrVLINDICE(I, 1) & vbTab & arrVLINDICE(I, 2)
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
    Frame4.Enabled = True
    
    Me.Caption = "Cadastro de Moedas - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADMOEDA
    objCADMOEDA.CODIGO = iCodigo
    
    ConfGridIndices
    
    If objCADMOEDA.Carrega_campos = True Then
    
       lblCodigo.Caption = Str(objCADMOEDA.CODIGO)
       txtDescricao.Text = objCADMOEDA.DESCRI
       arrVLINDICE = objCADMOEDA.INDICE
       
       If IsArray(arrVLINDICE) Then
          For I = 1 To UBound(arrVLINDICE)
              flxINDICES.AddItem "" & vbTab & arrVLINDICE(I, 1) & vbTab & arrVLINDICE(I, 2)
          Next I
       End If
    
    End If

End Sub


