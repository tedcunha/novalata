VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPEDIDOS 
   Caption         =   "Relatório de Pedidos de Vendas por Cliente"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   10080
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   23
      Top             =   2520
      Width           =   10095
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Liberado Financeiro"
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
         Index           =   5
         Left            =   4560
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
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
         Left            =   8520
         TabIndex        =   10
         Top             =   240
         Width           =   1455
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
         Index           =   3
         Left            =   6960
         TabIndex        =   9
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Faturados"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   1215
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
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
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
      Left            =   0
      TabIndex        =   22
      Top             =   3120
      Width           =   2775
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
         Left            =   120
         TabIndex        =   11
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
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   19
      Top             =   1560
      Width           =   10095
      Begin VB.TextBox txtCODCLIFIN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "txtCODCLIFIN"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdPesqCLIFIN 
         Height          =   315
         Left            =   2640
         Picture         =   "frmPEDIDOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmPEDIDOS.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDesClieFin 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDesClieFin"
         Height          =   285
         Left            =   3000
         TabIndex        =   26
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label lblDesClieIni 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDesClieIni"
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Top             =   240
         Width           =   6975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Inicial"
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
         TabIndex        =   21
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Final"
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
         TabIndex        =   20
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   10095
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         Width           =   1215
         _ExtentX        =   2143
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10095
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
         Picture         =   "frmPEDIDOS.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmPEDIDOS.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPEDIDOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Dim objBLBFunc      As Object
Dim objRELPEDIDOS   As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object
Dim strCABEC1       As String
Dim strCABEC2       As String


Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call InpRelCotAbertasAn
End Sub

Private Sub cmdPesqCLI_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODCLIINI.Text = varRETORNO
       lblDesClieIni.Caption = PegaDescClie(varRETORNO)
    End If
    
    txtCODCLIINI.SetFocus

End Sub

Private Sub cmdPesqCLIFIN_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODCLIFIN.Text = varRETORNO
       lblDesClieFin.Caption = PegaDescClie(varRETORNO)
    End If
    
    txtCODCLIFIN.SetFocus

End Sub

Private Sub cmdVoltar_Click()
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
    Set objRELPEDIDOS = CreateObject("RELCOMERCIAL.clsPEDIDOS")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmPEDIDOS
    Call LimpaCamposLabel
    
    objRELPEDIDOS.FILIAL = FILIAL

    optTdSomSemPed(2).Value = True
    optRELCOTAANSIN(0).Value = True
    
    mskDTINI.Text = Format(Date, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Date + 30, "DD/MM/YYYY")

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmPEDIDOS
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmPEDIDOS
End Sub


Private Sub txtCODCLIFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLIFIN.Name, frmPEDIDOS
End Sub

Private Sub txtCODCLIFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLIFIN.Text
End Sub

Private Sub txtCODCLIFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCLIFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLIFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblDesClieFin.Caption = PegaDescClie(Str(txtCODCLIFIN.Text))
    If Len(Trim(lblDesClieFin.Caption)) = 0 Then
        txtCODCLIFIN.Text = ""
        MsgBox "Cliente não existe !!!", vbOKOnly + vbExclamation, "Aviso"
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub txtCODCLIINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLIINI.Name, frmPEDIDOS
End Sub

Private Sub txtCODCLIINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLIINI.Text
End Sub

Private Sub txtCODCLIINI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCLIINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLIINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblDesClieIni.Caption = PegaDescClie(Str(txtCODCLIINI.Text))
    If Len(Trim(lblDesClieIni.Caption)) = 0 Then
        txtCODCLIINI.Text = ""
        MsgBox "Cliente não existe !!!", vbOKOnly + vbExclamation, "Aviso"
        Cancel = True
        Exit Sub
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

Private Sub InpRelCotAbertasAn()

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Relatório de Pedidos "
    
    If optRELCOTAANSIN(0).Value = True Then strCABEC1 = strCABEC1 & " [ Análitico ]"
    If optRELCOTAANSIN(1).Value = True Then strCABEC1 = strCABEC1 & " [ Sintético ]"
    
    If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text
    
    '' Chamada do Relatório

End Sub

Private Sub DestroiObjeto()
    Set objBLBFunc = Nothing
    Set objRELPEDIDOS = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub LimpaCamposLabel()
    lblDesClieIni.Caption = ""
    lblDesClieFin.Caption = ""
End Sub

Private Function PegaDescClie(strCodclie As String) As String

    PegaDescClie = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCodclie)
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF() Then PegaDescClie = Trim(BREC3!SGI_RAZAOSOC)
    BREC3.Close
    
End Function
