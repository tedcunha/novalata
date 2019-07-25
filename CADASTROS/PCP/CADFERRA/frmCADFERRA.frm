VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADFERRA 
   Caption         =   "Cadastro de Ferramentas"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5106
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
      TabCaption(0)   =   "Dados Gerais"
      TabPicture(0)   =   "frmCADFERRA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dados de Armazenagen"
      TabPicture(1)   =   "frmCADFERRA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   6855
         Begin VB.TextBox txtBoxCaixa 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "txtBoxCaixa"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtPrateleira 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   7
            Text            =   "txtPrateleira"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtArmario 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "txtArmario"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Box/Caixa"
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
            Index           =   3
            Left            =   300
            TabIndex        =   24
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Prateleira"
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
            Index           =   2
            Left            =   300
            TabIndex        =   23
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Armário"
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
            Index           =   1
            Left            =   300
            TabIndex        =   22
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   6855
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   0
            Text            =   "txtCodigo"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "txtDescricao"
            Top             =   600
            Width           =   5055
         End
         Begin VB.TextBox txtQtdEst 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   2
            Text            =   "txtQtdEst"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtEstMin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Text            =   "txtEstMin"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtCapac 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Text            =   "txtCapac"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.ComboBox cboUnid 
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Text            =   "cboUnid"
            Top             =   1320
            Width           =   1095
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
            Index           =   0
            Left            =   300
            TabIndex        =   21
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
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   20
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Qtde.Estoque"
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
            Index           =   1
            Left            =   300
            TabIndex        =   19
            Top             =   960
            Width           =   1170
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   2
            Left            =   300
            TabIndex        =   18
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Est.Minimo"
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
            Index           =   3
            Left            =   300
            TabIndex        =   17
            Top             =   1680
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Capacidade"
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
            Index           =   4
            Left            =   300
            TabIndex        =   16
            Top             =   2040
            Width           =   1020
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7095
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
         Picture         =   "frmCADFERRA.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmCADFERRA.frx":013A
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmCADFERRA.frx":023C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADFERRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho   As String
Public Linha      As Variant
Public cTipOper   As String
Public iCodigo    As Long
Public Filial     As Long
Public strAcesso  As String
Dim objBLBFunc    As Object
Dim objCADFERRA   As Object
Private Sub cboUnid_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboUnid, KeyAscii
End Sub

Private Sub cmdAltera_Click()
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    SSTab1.Tab = 0
    
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()
    
    If ValidaCampos = False Then Exit Sub
       
    If cTipOper = "I" Then objCADFERRA.CODIGO = objCADFERRA.Gera_Codigo(Me.Name)
       
    objCADFERRA.DESCRI = txtDescricao.Text
    objCADFERRA.QTDEST = IIf(Len(Trim(txtQtdEst.Text)) = 0, 0, CCur(txtQtdEst.Text))
    If Len(Trim(txtEstMin.Text)) > 0 Then objCADFERRA.ESTMIN = CCur(txtEstMin.Text)
    objCADFERRA.CAPACI = IIf(Len(Trim(txtCapac.Text)) = 0, 0, CCur(txtCapac.Text))
    objCADFERRA.ARMARIO = txtArmario.Text
    objCADFERRA.PRATELE = txtPrateleira.Text
    objCADFERRA.BOXCAIXA = txtBoxCaixa.Text
    objCADFERRA.UNIDADE = cboUnid.ItemData(cboUnid.ListIndex)
    
    If objCADFERRA.GRAVA(cTipOper) = False Then Exit Sub
    
    MsgBox "A ferramenta foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
      
    If objCADFERRA.Atualiza(cTipOper, objCADFERRA.CODIGO, Filial, Me.Name) = False Then Exit Sub
    
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objBLBFunc = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADFERRA = Nothing
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADFERRA = CreateObject("CADFERRA.clsCADFERRA")
   
   objCADFERRA.Filial = Filial
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta
   
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de ferramentas - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFERRA
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    SSTab1.Tab = 0
    
    objCADFERRA.PreenchComboUnidade cboUnid
   
End Sub
Private Sub txtArmario_GotFocus()
    objBLBFunc.SelecionaCampos txtArmario.Name, frmCADFERRA
End Sub

Private Sub txtArmario_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtBoxCaixa_GotFocus()
    objBLBFunc.SelecionaCampos txtBoxCaixa.Name, frmCADFERRA
End Sub

Private Sub txtBoxCaixa_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCapac_GotFocus()
    objBLBFunc.SelecionaCampos txtCapac.Name, frmCADFERRA
End Sub

Private Sub txtCapac_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCapac.Text
End Sub

Private Sub txtCapac_Validate(Cancel As Boolean)

    If Len(Trim(txtCapac.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCapac.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtCapac.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCapac.Text = Format(txtCapac.Text, "#,###0.000")

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADFERRA
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtEstMin_GotFocus()
    objBLBFunc.SelecionaCampos txtEstMin.Name, frmCADFERRA
End Sub

Private Sub txtEstMin_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtEstMin.Text
End Sub

Private Sub txtEstMin_Validate(Cancel As Boolean)

    If Len(Trim(txtEstMin.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtEstMin.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtEstMin.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtEstMin.Text = Format(txtEstMin.Text, "#,###0.000")

End Sub

Private Sub txtPrateleira_GotFocus()
    objBLBFunc.SelecionaCampos txtPrateleira.Name, frmCADFERRA
End Sub

Private Sub txtPrateleira_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtQtdEst_GotFocus()
    objBLBFunc.SelecionaCampos txtQtdEst.Name, frmCADFERRA
End Sub

Private Sub txtQtdEst_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQtdEst.Text
End Sub

Private Sub txtQtdEst_Validate(Cancel As Boolean)

    If Len(Trim(txtQtdEst.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtQtdEst.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtQtdEst.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtQtdEst.Text = Format(txtQtdEst.Text, "#,###0.000")

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Len(Trim(txtDescricao.Text)) = 0 Then
        MsgBox "Atenção o campos descrição não pode ser vazio !!!", vbOKOnly + vbExclamation, "Aviso"
        SSTab1.Tab = 0
        txtDescricao.SetFocus
        Exit Function
     End If
     If cboUnid.ListIndex = -1 Then
        MsgBox "Atenção o campos unidade não pode ser vazio !!!", vbOKOnly + vbExclamation, "Aviso"
        SSTab1.Tab = 0
        cboUnid.SetFocus
        Exit Function
     End If
     If Len(Trim(txtCapac.Text)) = 0 Then
        MsgBox "Atenção a capacidade unidade não pode ser vazio !!!", vbOKOnly + vbExclamation, "Aviso"
        SSTab1.Tab = 0
        txtCapac.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    Dim I As Integer
   
    Me.Caption = "Cadastro de ferramentas - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADFERRA
    
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    SSTab1.Tab = 0
    
    objCADFERRA.PreenchComboUnidade cboUnid
    
    objCADFERRA.CODIGO = iCodigo
    
    If objCADFERRA.Carrega_campos = True Then
    
       txtCodigo.Text = objCADFERRA.CODIGO
       
       txtDescricao.Text = objCADFERRA.DESCRI
       txtQtdEst.Text = Format(objCADFERRA.QTDEST, "#,###0.000")
       txtEstMin.Text = Format(objCADFERRA.ESTMIN, "#,###0.000")
       txtCapac.Text = Format(objCADFERRA.CAPACI, "#,###0.000")
       txtArmario.Text = objCADFERRA.ARMARIO
       txtPrateleira.Text = objCADFERRA.PRATELE
       txtBoxCaixa.Text = objCADFERRA.BOXCAIXA
       For I = 0 To (cboUnid.ListCount - 1)
           If cboUnid.ItemData(I) = objCADFERRA.UNIDADE Then cboUnid.ListIndex = I
       Next I
    End If
   
End Sub

Private Sub Altera()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Dim I As Integer
   
    Me.Caption = "Cadastro de ferramentas - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFERRA
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    SSTab1.Tab = 0
    
    objCADFERRA.PreenchComboUnidade cboUnid
    
    objCADFERRA.CODIGO = iCodigo
    
    If objCADFERRA.Carrega_campos = True Then
    
       txtCodigo.Text = objCADFERRA.CODIGO
       
       txtDescricao.Text = objCADFERRA.DESCRI
       txtQtdEst.Text = Format(objCADFERRA.QTDEST, "#,###0.000")
       txtEstMin.Text = Format(objCADFERRA.ESTMIN, "#,###0.000")
       txtCapac.Text = Format(objCADFERRA.CAPACI, "#,###0.000")
       txtArmario.Text = objCADFERRA.ARMARIO
       txtPrateleira.Text = objCADFERRA.PRATELE
       txtBoxCaixa.Text = objCADFERRA.BOXCAIXA
       For I = 0 To (cboUnid.ListCount - 1)
           If cboUnid.ItemData(I) = objCADFERRA.UNIDADE Then cboUnid.ListIndex = I
       Next I
       
    End If
   
End Sub

