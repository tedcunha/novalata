VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELVENDEDORES 
   Caption         =   "Relatorio de Pedidos por vendedores"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "[ Relatório ]"
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
      Left            =   5760
      TabIndex        =   24
      Top             =   1080
      Width           =   4215
      Begin VB.OptionButton optRELCOTAANSIN 
         Caption         =   "Análitico"
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
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optRELCOTAANSIN 
         Caption         =   "Sintético"
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
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[ Periodo ]"
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
      Left            =   5520
      TabIndex        =   20
      Top             =   2640
      Width           =   4455
      Begin VB.OptionButton optDiaMesAno 
         Caption         =   "Dia"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optDiaMesAno 
         Caption         =   "Mês"
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
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optDiaMesAno 
         Caption         =   "Ano"
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
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Pedidos ]"
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
      Left            =   0
      TabIndex        =   15
      Top             =   2640
      Width           =   5415
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Reprovados"
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
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Todos"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Bloqueados"
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
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Liberados"
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
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   9975
      Begin VB.TextBox txtCODCLIFIN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "txtCODCLIFIN"
         Top             =   615
         Width           =   975
      End
      Begin VB.CommandButton cmdPesqCLIFIN 
         Height          =   315
         Left            =   2640
         Picture         =   "frmRELVENDEDORES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCODCLIINI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "txtCODCLIINI"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPesqCLI 
         Height          =   315
         Left            =   2640
         Picture         =   "frmRELVENDEDORES.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3000
         TabIndex        =   28
         Top             =   600
         Width           =   6855
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor Inicial"
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
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor Final"
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
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   5655
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final"
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
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9975
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
         Picture         =   "frmRELVENDEDORES.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         Picture         =   "frmRELVENDEDORES.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELVENDEDORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public lngCodUsuario    As Long
Dim objBLBFunc          As Object
Dim objRELVENDEDORES    As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call ImpRelVendedores
End Sub

Private Sub cmdPesqCLI_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Vendedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLIINI.Text = varRETORNO
    
    Label5.Caption = objRELVENDEDORES.PegaVendedor(txtCODCLIINI.Text)
    If Len(Trim(Label5.Caption)) = 0 Then
        MsgBox "Vendedor Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODCLIINI.Text = ""
    End If
    
    txtCODCLIINI.SetFocus

End Sub

Private Sub cmdPesqCLIFIN_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Vendedores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLIFIN.Text = varRETORNO
    Label6.Caption = objRELVENDEDORES.PegaVendedor(txtCODCLIFIN.Text)
    If Len(Trim(Label6.Caption)) = 0 Then
        MsgBox "Vendedor Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODCLIFIN.Text = ""
    End If
    
    txtCODCLIFIN.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELVENDEDORES = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
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
    Set objRELVENDEDORES = CreateObject("RELCOMERCIAL.clsRELVENDEDORES")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")

    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELVENDEDORES
    
    objRELVENDEDORES.FILIAL = FILIAL
    
    Label5.Caption = ""
    Label6.Caption = ""
    
    mskDTINI.Text = Format(Date, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Date + 30, "DD/MM/YYYY")

    optRELCOTAANSIN(0).Value = True
    optTdSomSemPed(0).Value = True
    optDiaMesAno(0).Value = True
    
    Frame3.Enabled = True
    If lngCodUsuario > 0 Then
        Frame3.Enabled = False
        txtCODCLIINI.Text = Trim(Str(objRELVENDEDORES.PegaIDVendedor(Str(lngCodUsuario))))
        Label5.Caption = objRELVENDEDORES.PegaVendedor(txtCODCLIINI.Text)
    End If

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELVENDEDORES
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELVENDEDORES
End Sub

Private Sub txtCODCLIFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLIFIN.Name, frmRELVENDEDORES
End Sub

Private Sub txtCODCLIFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLIFIN.Text
End Sub

Private Sub txtCODCLIFIN_Validate(Cancel As Boolean)
    
    If Len(Trim(txtCODCLIFIN.Text)) = 0 Then
       Label6.Caption = ""
       Exit Sub
    End If
    
    Label6.Caption = objRELVENDEDORES.PegaVendedor(txtCODCLIFIN.Text)
    If Len(Trim(Label6.Caption)) = 0 Then
        MsgBox "Vendedor Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODCLIFIN.Text = ""
        Label6.Caption = ""
        Cancel = True
    End If
    
End Sub

Private Sub txtCODCLIINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLIINI.Name, frmRELVENDEDORES
End Sub

Private Sub txtCODCLIINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLIINI.Text
End Sub

Private Sub txtCODCLIINI_Validate(Cancel As Boolean)

    If Len(Trim(txtCODCLIINI.Text)) = 0 Then
       Label5.Caption = ""
       Exit Sub
    End If
    
    Label5.Caption = objRELVENDEDORES.PegaVendedor(txtCODCLIINI.Text)

    If Len(Trim(Label5.Caption)) = 0 Then
        MsgBox "Vendedor Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODCLIINI.Text = ""
        Label5.Caption = ""
        Cancel = True
    End If

End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
        If Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) > 0 Then
           If CLng(txtCODCLIINI.Text) > CLng(txtCODCLIFIN.Text) Then
              MsgBox "Cliente Inicial não pode ser maior que Cliente Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODCLIINI.SetFocus
              Exit Function
           End If
        End If
        
        If Not IsDate(mskDTINI.Text) Then
            MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
        If Not IsDate(mskDTFIN.Text) Then
            MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTFIN.SetFocus
            Exit Function
        End If
        
        If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
            MsgBox "Data Inicial não pode ser maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
    
    ConfereCampos = True

End Function



Private Sub ImpRelVendedores()


    Dim strNomRel As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADPEDVENDH.SGI_DATAPED "
    sSql = sSql & "     , SGI_CADPEDVENDH.SGI_CODIGO "
    sSql = sSql & "     , SGI_CADPEDVENDH.SGI_CODCLI "
    sSql = sSql & "     , SGI_CADCLIENTE.SGI_RAZAOSOC "
    sSql = sSql & "     , SGI_CADPEDVENDH.SGI_QTDITENS "
    sSql = sSql & "     , SGI_CADPEDVENDH.SGI_VLTOT "
    sSql = sSql & "     , SGI_CADPEDVENDH.SGI_VLIPI"
    sSql = sSql & "     , SGI_CADPEDVENDH.SGI_CODVEND "
    sSql = sSql & "     , SGI_CADVENDEDOR.SGI_DESCRICAO "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE "
    sSql = sSql & "     , SGI_CADPEDVENDH SGI_CADPEDVENDH "
    sSql = sSql & "     , SGI_CADVENDEDOR SGI_CADVENDEDOR "
    
    sSql = sSql & " Where "
    
    sSql = sSql & "        SGI_CADPEDVENDH.SGI_FILIAL      = SGI_CADCLIENTE.SGI_FILIAL "
    sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_CODCLI      = SGI_CADCLIENTE.SGI_CODIGO "
    sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_FILIAL      = SGI_CADVENDEDOR.SGI_FILIAL "
    sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_CODVEND   = SGI_CADVENDEDOR.SGI_CODIGO "
    
    If Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) > 0 Then
       sSql = sSql & "  And   (SGI_CADPEDVENDH.SGI_CODVEND >= " & Trim(txtCODCLIINI.Text) & " And SGI_CADPEDVENDH.SGI_CODVEND <= " & Trim(txtCODCLIFIN.Text) & ")"
    ElseIf Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) = 0 Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_CODVEND = " & Trim(txtCODCLIINI.Text)
    End If
    
    sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_FILIAL = " & FILIAL
    sSql = sSql & "  And  (SGI_CADPEDVENDH.SGI_DATAPED >= '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And SGI_CADPEDVENDH.SGI_DATAPED <= '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "')"
    
    If optTdSomSemPed(1).Value = True Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_STATUS = 'B'"
    ElseIf optTdSomSemPed(2).Value = True Then
       sSql = sSql & "  And   (SGI_CADPEDVENDH.SGI_STATUS = 'L' or SGI_CADPEDVENDH.SGI_STATUS = 'N')"
    ElseIf optTdSomSemPed(4).Value = True Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_STATUS = 'R'"
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close

    strCABEC1 = "Relatório de Pedidos de Vendedores "
    
    If optTdSomSemPed(0).Value = True Then
       strCABEC1 = strCABEC1 & " [Todos] "
    ElseIf optTdSomSemPed(1).Value = True Then
       strCABEC1 = strCABEC1 & " [Bloqueados] "
    ElseIf optTdSomSemPed(2).Value = True Then
       strCABEC1 = strCABEC1 & " [Liberados] "
    ElseIf optTdSomSemPed(4).Value = True Then
       strCABEC1 = strCABEC1 & " [Reprovados] "
    ''ElseIf optTdSomSemPed(4).Value = True Then
    ''   strCABEC1 = strCABEC1 & " [Faturados] "
    End If

    If optRELCOTAANSIN(0).Value = True Then strCABEC1 = strCABEC1 & " [ Análitico ]"
    If optRELCOTAANSIN(1).Value = True Then strCABEC1 = strCABEC1 & " [ Sintético ]"

    If optDiaMesAno(0).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text
    ElseIf optDiaMesAno(1).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Mês " & Format(Month(CDate(mskDTINI.Text)), "##00") & "/" & Year(CDate(mskDTINI.Text)) & " ao Mês " & Format(Month(CDate(mskDTFIN.Text)), "##00") & "/" & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "No Mês " & Format(Month(CDate(mskDTINI.Text)), "##00") & "/" & Year(CDate(mskDTFIN.Text))
    ElseIf optDiaMesAno(2).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "Do Ano " & Year(CDate(mskDTINI.Text)) & " ao Ano " & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "No Ano " & Year(CDate(mskDTFIN.Text))
    End If


    strNomRel = ""
    If optDiaMesAno(0).Value = True Then
       If optRELCOTAANSIN(0).Value = True Then
          strNomRel = "RELPDVVENDIAAN.rpt"
       ElseIf optRELCOTAANSIN(1).Value = True Then
          strNomRel = "RELPDVVENDIASI.rpt"
       End If
    ElseIf optDiaMesAno(1).Value = True Then
       If optRELCOTAANSIN(0).Value = True Then
          strNomRel = "RELPDVVENMESAN.rpt"
       ElseIf optRELCOTAANSIN(1).Value = True Then
          strNomRel = "RELPDVVENMESSI.rpt"
       End If
    ElseIf optDiaMesAno(2).Value = True Then
       If optRELCOTAANSIN(0).Value = True Then
          strNomRel = "RELPDVVENANOAN.rpt"
       ElseIf optRELCOTAANSIN(1).Value = True Then
          strNomRel = "RELPEDVENDANSI.rpt"
       End If
    End If

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, False)
    End If
    
End Sub
