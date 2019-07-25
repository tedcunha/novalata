VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmCADCONDPAGTO 
   Caption         =   "Cadastro de condição de pagamento"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   5895
      Begin VB.TextBox txtPORCACRE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "txtPORCACR"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmbGravEsp 
         Height          =   315
         Left            =   1965
         Picture         =   "frmCADCONDPAGTO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtNDIAS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "txtNDIAS"
         Top             =   480
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid flxParcelas 
         Height          =   2175
         Left            =   2445
         TabIndex        =   16
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3836
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "% Acresc."
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
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Dias:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtParcelas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtParcelas"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4575
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Parcelas:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   810
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
      Width           =   5895
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
         Picture         =   "frmCADCONDPAGTO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmCADCONDPAGTO.frx":0634
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
         Picture         =   "frmCADCONDPAGTO.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADCONDPAGTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public cTipOper      As String
Public iCodigo       As Integer
Public FILIAL        As Integer
Public strAcesso     As String
Public strMODPAI     As String
Public strUSUARIO    As String
Dim objBLBFunc       As Object
Dim objCADCONDPAGTO  As Object
Dim arrPARCPAGTO     As Variant
 
Private Sub cmbGravEsp_Click()
    If cTipOper = "I" Then IncGridParc
    If cTipOper = "A" Then IncGridParc
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro condição de pagamento - [ ALTERAÇÃO ]"
    
    txtNDIAS.Enabled = True
    txtPORCACRE.Enabled = True
    cmbGravEsp.Enabled = True
    
    cTipOper = "A"
    
    txtCodigo.Text = ""
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADCONDPAGTO.CODPGTO = objCADCONDPAGTO.Gera_Codigo(Me.Name)
    
    objCADCONDPAGTO.DESPGTO = txtDescricao.Text
    objCADCONDPAGTO.PARPGTO = Val(txtParcelas.Text)
    
    If flxParcelas.Rows > 1 Then
       
       ReDim arrPARCPAGTO(1 To (flxParcelas.Rows - 1), 1 To 4)
       
       For I = 1 To UBound(arrPARCPAGTO)
           arrPARCPAGTO(I, 1) = flxParcelas.TextMatrix(I, 0)
           arrPARCPAGTO(I, 2) = flxParcelas.TextMatrix(I, 1)
           arrPARCPAGTO(I, 3) = flxParcelas.TextMatrix(I, 2)
           arrPARCPAGTO(I, 4) = flxParcelas.TextMatrix(I, 3)
       Next I
       
       objCADCONDPAGTO.NPCPGTO = arrPARCPAGTO
    End If

    If objCADCONDPAGTO.GRAVA(cTipOper) = True Then
          
       MsgBox "A forma de pagamento foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
       If objCADCONDPAGTO.Atualiza(cTipOper, Str(objCADCONDPAGTO.CODPGTO), FILIAL, Me.Name) = False Then Exit Sub
       
       If cTipOper = "I" Then
          Set objBLBFunc = Nothing
          Set objCADCONDPAGTO = Nothing
          Unload Me
       End If
          
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADCONDPAGTO = Nothing
    Unload Me
End Sub

Private Sub flxParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       If cTipOper = "C" Then Exit Sub
       If flxParcelas.Rows = 2 Then flxParcelas.Rows = 1
       If flxParcelas.Rows > 2 Then flxParcelas.RemoveItem (flxParcelas.RowSel)
       RefazParc
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
   Set objCADCONDPAGTO = CreateObject("CADCONDPAGTO.clsCADCONDPAGTO")
   
   objCADCONDPAGTO.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de condição de pagamento - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCONDPAGTO
    
    ConfGridPgto
    txtCodigo.Text = ""
   
End Sub

Private Sub ConfGridPgto()

    flxParcelas.Rows = 1
    flxParcelas.Cols = 4
    
    flxParcelas.TextMatrix(0, 0) = ""
    flxParcelas.TextMatrix(0, 1) = "Parcelas"
    flxParcelas.TextMatrix(0, 2) = "Dias"
    flxParcelas.TextMatrix(0, 3) = "% Acresc."
    
    flxParcelas.ColWidth(0) = 0
    flxParcelas.ColWidth(1) = 1000
    flxParcelas.ColWidth(2) = 700
    flxParcelas.ColWidth(3) = 1000
    
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADCONDPAGTO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtNDIAS_GotFocus()
    objBLBFunc.SelecionaCampos txtNDIAS.Name, frmCADCONDPAGTO
End Sub

Private Sub txtNDIAS_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtNDIAS.Text
End Sub

Private Sub txtParcelas_GotFocus()
    objBLBFunc.SelecionaCampos txtParcelas.Name, frmCADCONDPAGTO
End Sub

Private Sub txtParcelas_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtParcelas.Text
End Sub


Private Sub IncGridParc()

    Dim I As Integer
    
    If Len(Trim(txtParcelas.Text)) = 0 Then
       MsgBox "Informe o numero de parcelas !!!", vbOKOnly + vbCritical, "Aviso"
       txtPORCACRE.Text = ""
       txtParcelas.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtNDIAS.Text)) = 0 Then
       MsgBox "Informe o numero de dias !!!", vbOKOnly + vbCritical, "Aviso"
       txtPORCACRE.Text = ""
       txtNDIAS.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(txtParcelas.Text) Then
       MsgBox "O campo parcela deve ser numero !!!", vbOKOnly + vbCritical, "aviso"
       txtParcelas.Text = ""
       txtPORCACRE.Text = ""
       txtParcelas.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(txtNDIAS.Text) Then
       MsgBox "O campo numero de dias deve ser numero !!!", vbOKOnly + vbCritical, "aviso"
       txtNDIAS.Text = ""
       txtPORCACRE.Text = ""
       txtNDIAS.SetFocus
       Exit Sub
    End If
    If Val(txtParcelas.Text) = 0 Then
       MsgBox "A quantidade de parcelas deve ser maior que zero !!!", vbOKOnly + vbCritical, "aviso"
       txtParcelas.Text = "1"
       txtPORCACRE.Text = ""
       txtParcelas.SetFocus
       Exit Sub
    End If
    If (flxParcelas.Rows - 1) = Val(txtParcelas.Text) Then
       MsgBox "O numero de parcelas já esta esgotado !!!", vbOKOnly + vbCritical, "aviso"
       txtNDIAS.Text = ""
       txtPORCACRE.Text = ""
       txtParcelas.SetFocus
       Exit Sub
    End If
    For I = 1 To (flxParcelas.Rows - 1)
        If flxParcelas.TextMatrix(I, 2) = txtNDIAS.Text Then
           MsgBox "Esta quantidade de dias já esta relacionada !!!", vbOKOnly + vbCritical, "aviso"
           txtNDIAS.SetFocus
           Exit Sub
        End If
    Next I
    
    flxParcelas.AddItem "" & vbTab & "" & vbTab & txtNDIAS.Text & vbTab & txtPORCACRE.Text
    txtNDIAS.Text = ""
    txtPORCACRE.Text = ""
    RefazParc
    txtNDIAS.SetFocus
    
End Sub

Private Sub txtPORCACRE_GotFocus()
    objBLBFunc.SelecionaCampos txtPORCACRE.Name, frmCADCONDPAGTO
End Sub

Private Sub txtPORCACRE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPORCACRE.Text
End Sub

Private Sub txtPORCACRE_Validate(Cancel As Boolean)

    If Len(Trim(txtPORCACRE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPORCACRE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtPORCACRE.Text = ""
       Cancel = True
       Exit Sub
    End If
    If Val(txtPORCACRE.Text) < 0 Then
       MsgBox "Somente é permitido numeros positivos !!!", vbOKOnly + vbCritical, "Aviso"
       txtPORCACRE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPORCACRE.Text = Format(txtPORCACRE.Text, "#,##0.00")

End Sub

Private Sub RefazParc()
    
    Dim I As Integer

    For I = 1 To (flxParcelas.Rows - 1)
        flxParcelas.TextMatrix(I, 0) = I
        flxParcelas.TextMatrix(I, 1) = Str(I) & "/" & txtParcelas.Text
    Next I

End Sub

Private Function Valida_Campos() As Boolean

       Valida_Campos = False
       
       If Len(Trim(txtDescricao.Text)) = 0 Then
          MsgBox "Informe a descrição do pagamento !!!", vbOKOnly + vbCritical, "aviso"
          txtDescricao.SetFocus
          Exit Function
       End If
       If Len(Trim(txtParcelas.Text)) = 0 Then
          MsgBox "Informe as parcelas !!!", vbOKOnly + vbCritical, "aviso"
          txtParcelas.SetFocus
          Exit Function
       End If
       If (flxParcelas.Rows - 1) = 0 Then
          MsgBox "Não foram informado parcelas !!!", vbOKOnly + vbCritical, "aviso"
          txtNDIAS.SetFocus
          Exit Function
       End If
       If ((flxParcelas.Rows - 1) < Val(txtParcelas.Text)) Or (flxParcelas.Rows - 1) > Val(txtParcelas.Text) Then
          MsgBox "Total de parcelas não esta correto ", vbOKOnly + vbCritical, "aviso"
          txtParcelas.SetFocus
          Exit Function
       End If
       
       Valida_Campos = True

End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = True
    
   
    Me.Caption = "Cadastro condição de pagamento - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADCONDPAGTO
    
    txtNDIAS.Enabled = False
    txtPORCACRE.Enabled = False
    cmbGravEsp.Enabled = False
    
    objCADCONDPAGTO.CODPGTO = iCodigo
    
    ConfGridPgto
    
    If objCADCONDPAGTO.Carrega_campos = True Then
       
       txtCodigo.Text = Str(objCADCONDPAGTO.CODPGTO)
       txtDescricao.Text = objCADCONDPAGTO.DESPGTO
       txtParcelas.Text = Str(objCADCONDPAGTO.PARPGTO)
       arrPARCPAGTO = objCADCONDPAGTO.NPCPGTO
       
       If IsArray(arrPARCPAGTO) Then
            For I = 1 To UBound(arrPARCPAGTO)
                flxParcelas.AddItem arrPARCPAGTO(I, 1) & vbTab & "" & vbTab & Trim(arrPARCPAGTO(I, 2)) & vbTab & Format(arrPARCPAGTO(I, 3), "#,##0.00")
            Next I
            RefazParc
       End If
       
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
   
    Me.Caption = "Cadastro condição de pagamento - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCONDPAGTO
    
    txtNDIAS.Enabled = True
    txtPORCACRE.Enabled = True
    cmbGravEsp.Enabled = True
    
    objCADCONDPAGTO.CODPGTO = iCodigo
    
    ConfGridPgto
    
    If objCADCONDPAGTO.Carrega_campos = True Then
       
       txtCodigo.Text = Str(objCADCONDPAGTO.CODPGTO)
       txtDescricao.Text = objCADCONDPAGTO.DESPGTO
       txtParcelas.Text = Str(objCADCONDPAGTO.PARPGTO)
       arrPARCPAGTO = objCADCONDPAGTO.NPCPGTO
       
        If IsArray(arrPARCPAGTO) Then
            For I = 1 To UBound(arrPARCPAGTO)
                flxParcelas.AddItem arrPARCPAGTO(I, 1) & vbTab & "" & vbTab & Trim(arrPARCPAGTO(I, 2)) & vbTab & Format(arrPARCPAGTO(I, 3), "#,##0.00")
            Next I
            RefazParc
        End If
    End If

End Sub

