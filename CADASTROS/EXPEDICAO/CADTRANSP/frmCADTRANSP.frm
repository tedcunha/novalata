VERSION 5.00
Begin VB.Form frmCADTRANSP 
   Caption         =   "Cadastro de Transportadoras"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6420
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
         Picture         =   "frmCADTRANSP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
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
         Picture         =   "frmCADTRANSP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
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
         Picture         =   "frmCADTRANSP.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Width           =   6375
      Begin VB.TextBox txtContato 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "txtContato"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtTelefone 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   6
         Text            =   "txtTelefone"
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtCEP 
         Height          =   285
         Left            =   4800
         MaxLength       =   9
         TabIndex        =   9
         Text            =   "txtCEP"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cboESTADO 
         Height          =   315
         Left            =   4800
         TabIndex        =   8
         Text            =   "cboESTADO"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "txtCidade"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "txtBairro"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtEndereco 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "txtEndereco"
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox txtCGCCNPJ 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "txtCGCCNPJ"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         Index           =   8
         Left            =   600
         TabIndex        =   24
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
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
         Index           =   7
         Left            =   480
         TabIndex        =   23
         Top             =   2400
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
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
         Index           =   6
         Left            =   4270
         TabIndex        =   22
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         Index           =   5
         Left            =   4020
         TabIndex        =   21
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
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
         Left            =   660
         TabIndex        =   20
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
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
         Left            =   750
         TabIndex        =   19
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Left            =   480
         TabIndex        =   18
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CPF/CNPJ:"
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
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   975
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
         Left            =   660
         TabIndex        =   15
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmCADTRANSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho   As String
Public Linha      As Variant
Public cTipOper   As String
Public iCodigo    As Long
Public FILIAL     As Integer
Public strAcesso  As String
Dim objBLBFunc    As Object
Dim objCADTRANSP  As Object

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Transportadoras - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADTRANSP.CODIGO = objCADTRANSP.Gera_Codigo(Me.Name)
    objCADTRANSP.RAZAOSOC = txtDescricao.Text
    objCADTRANSP.CPFCNPJ = txtCGCCNPJ.Text
    objCADTRANSP.ENDERECO = txtEndereco.Text
    objCADTRANSP.BAIRRO = txtBairro.Text
    objCADTRANSP.CIDADE = txtCidade.Text
    If cboESTADO.ListIndex > -1 Then objCADTRANSP.ESTADO = cboESTADO.ItemData(cboESTADO.ListIndex)
    objCADTRANSP.CEP = txtCEP.Text
    objCADTRANSP.TELEFONE = txtTelefone.Text
    objCADTRANSP.CONTATO = txtContato.Text
    
    If objCADTRANSP.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A transportadora foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       Inclui
       txtDescricao.SetFocus
    End If
          
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADTRANSP = Nothing
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
   Set objCADTRANSP = CreateObject("CADTRANSP.clsCADTRANSP")
   
   objCADTRANSP.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Transportadoras - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTRANSP
    objBLBFunc.Preenche_Estado cboESTADO
    
    txtCodigo.Text = ""
   
End Sub

Private Sub txtBairro_GotFocus()
    objBLBFunc.SelecionaCampos txtBairro.Name, frmCADTRANSP
End Sub

Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCEP_GotFocus()
    objBLBFunc.SelecionaCampos txtCEP.Name, frmCADTRANSP
End Sub

Private Sub txtCGCCNPJ_GotFocus()
    objBLBFunc.SelecionaCampos txtCGCCNPJ.Name, frmCADTRANSP
End Sub

Private Sub txtCGCCNPJ_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCidade_GotFocus()
    objBLBFunc.SelecionaCampos txtCidade.Name, frmCADTRANSP
End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADTRANSP
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtEndereco_GotFocus()
    objBLBFunc.SelecionaCampos txtEndereco.Name, frmCADTRANSP
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
    
     If Len(Trim(txtCGCCNPJ.Text)) > 0 Then
       
       If Not IsNumeric(txtCGCCNPJ.Text) Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "aviso"
          txtCGCCNPJ.Text = ""
          txtCGCCNPJ.SetFocus
          Exit Function
       End If
       
       If Len(Trim(txtCGCCNPJ.Text)) < 14 And Len(Trim(txtCGCCNPJ.Text)) > 11 Then
          MsgBox "CPF/CNPJ Inválido !!!", vbOKOnly + vbCritical, "Aviso"
          txtCGCCNPJ.Text = ""
          txtCGCCNPJ.SetFocus
          Exit Function
       End If
       
       If Len(Trim(txtCGCCNPJ.Text)) < 11 Then
          MsgBox "CPF/CNPJ Inválido !!!", vbOKOnly + vbCritical, "Aviso"
          txtCGCCNPJ.Text = ""
          txtCGCCNPJ.SetFocus
          Exit Function
       End If
       
       If Len(Trim(txtCGCCNPJ.Text)) = 11 Then
          If objBLBFunc.ViewCPF(txtCGCCNPJ.Text) = False Then
             MsgBox "CPF Inválido !!!", vbOKOnly + vbCritical, "Aviso"
             txtCGCCNPJ.Text = ""
             txtCGCCNPJ.SetFocus
             Exit Function
          End If
       End If
       
       If Len(Trim(txtCGCCNPJ.Text)) = 14 Then
          If objBLBFunc.ViewCGC(txtCGCCNPJ.Text) = False Then
             MsgBox "CNPJ Inválido !!!", vbOKOnly + vbCritical, "aviso"
             txtCGCCNPJ.Text = ""
             txtCGCCNPJ.SetFocus
             Exit Function
          End If
       End If
       
       If cTipOper = "I" Then
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CPFCNPJ = '" & txtCGCCNPJ.Text & "'" & vbCrLf
          
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "CPFCNPJ já cadastrado !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtCGCCNPJ.Text = ""
             txtCGCCNPJ.SetFocus
             Exit Function
          End If
          BREC.Close
       End If
       
       If cTipOper = "A" Then
          If txtCGCCNPJ.Text <> objCADTRANSP.CPFCNPJ Then
             sSql = "Select " & vbCrLf
             sSql = sSql & "       * " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
             sSql = sSql & "   And SGI_CPFCNPJ = '" & txtCGCCNPJ.Text & "'" & vbCrLf
          
             BREC.Open sSql, adoBanco_Dados, adOpenDynamic
             If Not BREC.EOF Then
                MsgBox "CPFCNPJ já cadastrado !!!", vbOKOnly + vbExclamation, "Aviso"
                BREC.Close
                txtCGCCNPJ.Text = objCADTRANSP.CPFCNPJ
                txtCGCCNPJ.SetFocus
                Exit Function
             End If
             BREC.Close
          End If
       End If
       
       
     End If
     
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Transportadora inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRICAO ='" & txtDescricao.Text & "'"
        sSql = sSql & " And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Razão social da transportadora já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADTRANSP.RAZAOSOC <> txtDescricao.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRICAO ='" & txtDescricao.Text & "'"
           sSql = sSql & " And SGI_FILIAL = " & FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Descrição do tipo de produto existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADTRANSP.RAZAOSOC
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de Transportadoras - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADTRANSP
    objBLBFunc.Preenche_Estado cboESTADO
    
    objCADTRANSP.CODIGO = iCodigo
    
    If objCADTRANSP.Carrega_campos = True Then
    
       txtCodigo.Text = objCADTRANSP.CODIGO
       txtDescricao.Text = objCADTRANSP.RAZAOSOC
       txtCGCCNPJ.Text = objCADTRANSP.CPFCNPJ
       txtEndereco.Text = objCADTRANSP.ENDERECO
       txtBairro.Text = objCADTRANSP.BAIRRO
       txtCidade.Text = objCADTRANSP.CIDADE
       For I = 0 To (cboESTADO.ListCount - 1)
           If cboESTADO.ItemData(I) = objCADTRANSP.ESTADO Then cboESTADO.ListIndex = I
       Next I
       txtTelefone.Text = objCADTRANSP.TELEFONE
       txtContato.Text = objCADTRANSP.CONTATO
       txtCEP.Text = objCADTRANSP.CEP
    
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Transportadoras - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTRANSP
    objBLBFunc.Preenche_Estado cboESTADO
    
    objCADTRANSP.CODIGO = iCodigo
    
    If objCADTRANSP.Carrega_campos = True Then
    
       txtCodigo.Text = objCADTRANSP.CODIGO
       txtDescricao.Text = objCADTRANSP.RAZAOSOC
       txtCGCCNPJ.Text = objCADTRANSP.CPFCNPJ
       txtEndereco.Text = objCADTRANSP.ENDERECO
       txtBairro.Text = objCADTRANSP.BAIRRO
       txtCidade.Text = objCADTRANSP.CIDADE
       For I = 0 To (cboESTADO.ListCount - 1)
           If cboESTADO.ItemData(I) = objCADTRANSP.ESTADO Then cboESTADO.ListIndex = I
       Next I
       txtTelefone.Text = objCADTRANSP.TELEFONE
       txtContato.Text = objCADTRANSP.CONTATO
       txtCEP.Text = objCADTRANSP.CEP
    
    End If

End Sub

