VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCARTPROD 
   Caption         =   "Carteira de Produ��o"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   11340
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "[ Empresa ]"
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
      Height          =   615
      Left            =   5520
      TabIndex        =   17
      Top             =   960
      Width           =   5775
      Begin VB.OptionButton optEmpresa 
         Caption         =   "STEEL ROL"
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
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optEmpresa 
         Caption         =   "NOVALATA"
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
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optEmpresa 
         Caption         =   "NOVALATA e STEEL"
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
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   5415
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
         TabIndex        =   13
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
         Left            =   1440
         TabIndex        =   14
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Frame11"
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
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   2760
      Width           =   11295
      Begin ComctlLib.ProgressBar prgBAR 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Cliente ]"
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
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   8415
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   1080
         Picture         =   "frmCARTPROD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "txtCIDCLIE"
         Top             =   255
         Width           =   975
      End
      Begin VB.Label lblDescCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "[ Vendedor ]"
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
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   8415
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1080
         Picture         =   "frmCARTPROD.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCODVEND 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "txtCODVEND"
         Top             =   255
         Width           =   975
      End
      Begin VB.Label lblDescVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendedor"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
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
         Picture         =   "frmCARTPROD.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
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
         Picture         =   "frmCARTPROD.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCARTPROD"
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
Dim objCARTPROD         As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String
Dim strNomRel           As String
Dim strEMPRESADESC      As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call Imprime
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
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
    arrCAMPOS(1, 3) = "C�digo"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Raz�o Social"
    arrCAMPOS(3, 4) = "3000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    arrCAMPOS(4, 1) = "SGI_NOMFANTA"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Nome Fantasia"
    arrCAMPOS(4, 4) = "2000"
    arrCAMPOS(4, 5) = "SGI_NOMFANTA"
    
    arrCAMPOS(5, 1) = "SGI_CIDNORM"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Cidade"
    arrCAMPOS(5, 4) = "1500"
    arrCAMPOS(5, 5) = "SGI_CIDNORM"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCIDCLIE.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescCliente)
    If Len(Trim(lblDescCliente.Caption)) = 0 Then txtCIDCLIE.Text = ""
    
    txtCIDCLIE.SetFocus

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "C�digo"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descri��o"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVEND.Text = varRETORNO
    
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedor)
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then txtCODVEND.Text = ""
    txtCODVEND.SetFocus

End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCARTPROD = CreateObject("RELCOMERCIAL.clsREMPEDCART")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "N�o foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call objBLBFunc.LimpaCampos(Me)
    
    objCARTPROD.FILIAL = FILIAL
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    Me.Caption = Me.Caption & " / " & Me.Name
    
    optEmpresa(1).value = True

    Call LimpaCamposLabel

    Frame11.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Object
End Sub

Private Sub txtCIDCLIE_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, Me
End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCIDCLIE.Text
End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCIDCLIE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCIDCLIE.Text) Then
       MsgBox "Somente � permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente)
    If Len(Trim(lblDescCliente.Caption)) = 0 Then
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODVEND_GotFocus()
    objBLBFunc.SelecionaCampos txtCODVEND.Name, Me
End Sub

Private Sub txtCODVEND_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODVEND.Text
End Sub

Private Sub txtCODVEND_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODVEND.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVEND.Text) Then
       MsgBox "Somente � permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVEND.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVEND.Text, lblDescVendedor)
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then
       txtCODVEND.Text = ""
       Cancel = True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub Destroy_Object()
    Set objBLBFunc = Nothing
    Set objCARTPROD = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCODIGO)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub LimpaCamposLabel()
    lblDescCliente.Caption = ""
    lblDescVendedor.Caption = ""
End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
    If Not IsDate(mskDTINI.Text) Then
        MsgBox "Data Inicial Inv�lida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTINI.SetFocus
        Exit Function
    End If
    If Not IsDate(mskDTFIN.Text) Then
        MsgBox "Data Final Inv�lida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTFIN.SetFocus
        Exit Function
    End If
    
    If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
        MsgBox "Data Inicial n�o pode ser maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTINI.SetFocus
        Exit Function
    End If
    
    ConfereCampos = True

End Function


Private Sub Imprime()
    
    Dim strEMPRESA              As String
    Dim strNOMCLIE              As String
    Dim strNOMVEND              As String
    Dim boolTemRegs             As Boolean
    Dim strSQLNOVALATA          As String
    Dim strSQLSTEEL             As String
    Dim lngQTDREGSNOVA          As Long
    Dim lngQTDREGSTEEL          As Long
    
    Dim arrDADOSTAB()           As String
    Dim arrDADOSTAB_STEEL()     As String
    
    Dim lngREGS                 As Long
    Dim lngQTDPED               As Long
    Dim lngQTDFAT               As Long
    Dim lngSALDO                As Long
    
    strSQLNOVALATA = ""
    strSQLSTEEL = ""
    
    strEMPRESA = ""
    
    boolTemRegs = False
   
    
    '' Novalata
    sSql = ""
    
    If (optEmpresa(0).value = True Or optEmpresa(2).value = True) Then
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD" & strEMPRESA & ".SGI_DATENTREGA" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODIGO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPROD" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_QTDE" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_QTDEPED" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_QTDFAT" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_NOMEVEND" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIALPED" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_FECHTPFU As SGI_FechTampaFuro" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & ".SGI_DATAPED" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODCLI" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernCorpo" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernTampa" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernFundo" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernArgola" & vbCrLf
    ''    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechTampaFuro" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechSoldaAgrafado" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & ".SGI_VLUNIT" & vbCrLf
        sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO  As SGI_DESCONDPGTO" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADCLIENTE      SGI_CADCLIENTE" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & " SGI_CADPEDVENDH" & strEMPRESA & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO      SGI_CADPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & " SGI_ORDEMPROD" & strEMPRESA & vbCrLf
        sSql = sSql & "      ,SGI_CADCONDPGTO  SGI_CADCONDPGTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & " SGI_CADPEDVENDI" & strEMPRESA & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_DATENTREGA Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
        
        ''If optStatusPED(1).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 0" & vbCrLf
        ''ElseIf optStatusPED(2).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 1" & vbCrLf
        ''ElseIf optStatusPED(3).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 2" & vbCrLf
        ''ElseIf optStatusPED(4).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 6" & vbCrLf  '' P.Cota
        ''ElseIf optStatusPED(5).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 7" & vbCrLf  '' P.Data
        ''ElseIf optStatusPED(6).value = True Then
            '' Pega P.Cota P.Data
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 6" & vbCrLf  '' P.Cota
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 7" & vbCrLf  '' P.Data
        ''End If
        
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIAL = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
        
        strNOMCLIE = ""
        If Len(Trim(txtCIDCLIE.Text)) > 0 Then
            sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODCLI = " & Trim(txtCIDCLIE.Text) & vbCrLf
            strNOMCLIE = " / Cliente : " & lblDescCliente.Caption
        End If
        
        strNOMVEND = ""
        If Len(Trim(txtCODVEND.Text)) > 0 Then
            strNOMVEND = "/ Vendedor : " & lblDescVendedor.Caption
            sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODVEND = " & Trim(txtCODVEND.Text) & vbCrLf
        End If
        
        sSql = sSql & "   And SGI_CADCONDPGTO.SGI_FILIAL     = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADCONDPGTO.SGI_CODIGO     = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODCONDPGT" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_FILIAL     = SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_CODIGO     = SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED" & vbCrLf
        sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_IDPRODUTO  = SGI_ORDEMPROD" & strEMPRESA & ".SGI_IDPRODUTO" & vbCrLf
        
        ''If optAgrup(0).value = True Or optAgrup(1).value = True Then
        ''    sSql = sSql & "Order By SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED" & vbCrLf
        ''    sSql = sSql & "        ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODIGO"
        ''ElseIf optAgrup(2).value = True Or optAgrup(3).value = True Then
            sSql = sSql & "Order By SGI_ORDEMPROD" & strEMPRESA & ".SGI_DATENTREGA" & vbCrLf
            sSql = sSql & "        ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED" & vbCrLf
            sSql = sSql & "        ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODIGO"
        ''End If
        
        strSQLNOVALATA = sSql
        
        If BREC.State = 1 Then BREC.Close
        
        BREC.Open sSql, adoBanco_Dados
        If Not BREC.EOF() Then
            boolTemRegs = True
            
                Frame11.Visible = True
                Frame11.Caption = "[ Aguarde Carregando dados NOVALATA ]"
                Frame11.Refresh
                
                prgBAR.Min = 0
                
                
                lngREGS = 0
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    BREC.MoveNext
                Loop
                lngQTDREGSNOVA = lngREGS
                
                prgBAR.Max = lngREGS
                ReDim arrDADOSTAB(1 To lngREGS, 1 To 23) As String
                BREC.MoveFirst
                lngREGS = 0
            
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    prgBAR.value = lngREGS
                    
                    arrDADOSTAB(lngREGS, 1) = BREC!SGI_CODPED
                    arrDADOSTAB(lngREGS, 2) = BREC!SGI_CODIGO
                    arrDADOSTAB(lngREGS, 3) = Format(BREC!SGI_DATAPED, "DD/MM/YYYY")
                    arrDADOSTAB(lngREGS, 4) = Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY")
                    arrDADOSTAB(lngREGS, 5) = Trim(BREC!SGI_RAZAOSOC)
                    arrDADOSTAB(lngREGS, 6) = Trim(BREC!SGI_CODPROD)
                    arrDADOSTAB(lngREGS, 7) = Trim(BREC!SGI_DESCRICAO)
                    arrDADOSTAB(lngREGS, 8) = Format(BREC!SGI_VLUNIT, "#,##0.00")
                    arrDADOSTAB(lngREGS, 9) = Trim(BREC!SGI_DESCONDPGTO)
                    arrDADOSTAB(lngREGS, 10) = Trim(BREC!SGI_NOMEVEND)
                    arrDADOSTAB(lngREGS, 11) = Trim(BREC!SGI_DESCRI)
                    
                    lngQTDPED = BREC!SGI_QTDEPED
                    arrDADOSTAB(lngREGS, 12) = lngQTDPED
                    
                    lngQTDFAT = 0
                    If Not IsNull(BREC!SGI_QTDFAT) Then lngQTDFAT = BREC!SGI_QTDFAT
                    arrDADOSTAB(lngREGS, 13) = lngQTDFAT
                    
                    lngSALDO = (lngQTDPED - lngQTDFAT)
                    arrDADOSTAB(lngREGS, 14) = lngSALDO
                    
                    If BREC!SGI_STATUS = 0 Then
                        arrDADOSTAB(lngREGS, 15) = "Aberto"
                    ElseIf BREC!SGI_STATUS = 1 Then
                        arrDADOSTAB(lngREGS, 15) = "Parcial"
                    ElseIf BREC!SGI_STATUS = 2 Then
                        arrDADOSTAB(lngREGS, 15) = "Total"
                    ElseIf BREC!SGI_STATUS = 3 Then
                        arrDADOSTAB(lngREGS, 15) = "Bloqueada"
                    ElseIf BREC!SGI_STATUS = 4 Then
                        arrDADOSTAB(lngREGS, 15) = "Reprovado"
                    ElseIf BREC!SGI_STATUS = 6 Then
                        arrDADOSTAB(lngREGS, 15) = "P.Cota"
                    ElseIf BREC!SGI_STATUS = 7 Then
                        arrDADOSTAB(lngREGS, 15) = "P.Data"
                    ElseIf BREC!SGI_STATUS = 9 Then
                        arrDADOSTAB(lngREGS, 15) = "Liq.Man"
                    Else
                        arrDADOSTAB(lngREGS, 15) = "S/I"
                    End If
                    
                    arrDADOSTAB(lngREGS, 16) = "NOVALATA"
                    
                    '' Verniz FF
                    arrDADOSTAB(lngREGS, 17) = ""
                    arrDADOSTAB(lngREGS, 18) = ""
                    arrDADOSTAB(lngREGS, 19) = ""
                    arrDADOSTAB(lngREGS, 20) = ""
                    
                    If Not IsNull(BREC!SGI_VernCorpo) Then arrDADOSTAB(lngREGS, 17) = VernFolhaFrandes(BREC!SGI_VernCorpo)
                    If Not IsNull(BREC!SGI_VernTampa) Then arrDADOSTAB(lngREGS, 18) = VernFolhaFrandes(BREC!SGI_VernTampa)
                    If Not IsNull(BREC!SGI_VernFundo) Then arrDADOSTAB(lngREGS, 19) = VernFolhaFrandes(BREC!SGI_VernFundo)
                    If Not IsNull(BREC!SGI_VernArgola) Then arrDADOSTAB(lngREGS, 20) = VernFolhaFrandes(BREC!SGI_VernArgola)
                    
                    '' Fechamento
                    arrDADOSTAB(lngREGS, 21) = ""
                    arrDADOSTAB(lngREGS, 22) = ""
                    arrDADOSTAB(lngREGS, 23) = ""
                    
                    If Not IsNull(BREC!SGI_FechTampaFuro) Then arrDADOSTAB(lngREGS, 21) = PegaFechTampaFuro(BREC!SGI_FechTampaFuro)
                    If Not IsNull(BREC!SGI_NECKIN) Then arrDADOSTAB(lngREGS, 22) = IIf(BREC!SGI_NECKIN = 0, "N�o", "Sim")
                    If Not IsNull(BREC!SGI_FechSoldaAgrafado) Then arrDADOSTAB(lngREGS, 23) = TipoFecha(BREC!SGI_FechSoldaAgrafado)
                    
                    BREC.MoveNext
                Loop
             Frame11.Visible = False
        End If
        BREC.Close
    
    End If
    
    
    '' Gera��o para NOVALATA e STEEL
    If (optEmpresa(1).value = True Or optEmpresa(2).value = True) Then
        
        strEMPRESA = "_STEEL"
 
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD" & strEMPRESA & ".SGI_DATENTREGA" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODIGO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPROD" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_QTDE" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_QTDEPED" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_QTDFAT" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_NOMEVEND" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIALPED" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_FECHTPFU As SGI_FechTampaFuro" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & ".SGI_DATAPED" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODCLI" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernCorpo" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernTampa" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernFundo" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernArgola" & vbCrLf
''        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechTampaFuro" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechSoldaAgrafado" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADCONDPGTO.SGI_DESCRICAO  As SGI_DESCONDPGTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & ".SGI_VLUNIT" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADCLIENTE      SGI_CADCLIENTE" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & "     SGI_CADPEDVENDH" & strEMPRESA & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO      SGI_CADPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & "       SGI_ORDEMPROD" & strEMPRESA & vbCrLf
        sSql = sSql & "      ,SGI_CADCONDPGTO     SGI_CADCONDPGTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & "     SGI_CADPEDVENDI" & strEMPRESA & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_DATENTREGA Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
        
        ''If optStatusPED(1).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 0" & vbCrLf
        ''ElseIf optStatusPED(2).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 1" & vbCrLf
        ''ElseIf optStatusPED(3).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 2" & vbCrLf
        ''ElseIf optStatusPED(4).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 6" & vbCrLf  '' P.Cota
        ''ElseIf optStatusPED(5).value = True Then
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 7" & vbCrLf  '' P.Data
        ''ElseIf optStatusPED(6).value = True Then
            '' Pega P.Cota P.Data
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 6" & vbCrLf  '' P.Cota
        ''    sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_STATUS = 7" & vbCrLf  '' P.Data
        ''End If
        
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIAL = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strEMPRESA & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
        
        strNOMCLIE = ""
        If Len(Trim(txtCIDCLIE.Text)) > 0 Then
            sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODCLI = " & Trim(txtCIDCLIE.Text) & vbCrLf
            strNOMCLIE = " / Cliente : " & lblDescCliente.Caption
        End If
        
        strNOMVEND = ""
        If Len(Trim(txtCODVEND.Text)) > 0 Then
            strNOMVEND = "/ Vendedor : " & lblDescVendedor.Caption
            sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODVEND = " & Trim(txtCODVEND.Text) & vbCrLf
        End If
        
        sSql = sSql & "   And SGI_CADCONDPGTO.SGI_FILIAL     = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADCONDPGTO.SGI_CODIGO     = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODCONDPGT" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_FILIAL     = SGI_ORDEMPROD" & strEMPRESA & ".SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_CODIGO     = SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED" & vbCrLf
        sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_IDPRODUTO  = SGI_ORDEMPROD" & strEMPRESA & ".SGI_IDPRODUTO" & vbCrLf
    
        ''If optAgrup(0).value = True Or optAgrup(1).value = True Then
        ''    sSql = sSql & "Order By SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED" & vbCrLf
        ''    sSql = sSql & "        ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODIGO"
        ''ElseIf optAgrup(2).value = True Or optAgrup(3).value = True Then
            sSql = sSql & "Order By SGI_ORDEMPROD" & strEMPRESA & ".SGI_DATENTREGA" & vbCrLf
            sSql = sSql & "        ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODPED" & vbCrLf
            sSql = sSql & "        ,SGI_ORDEMPROD" & strEMPRESA & ".SGI_CODIGO"
        ''End If
    
        strSQLSTEEL = sSql
        
        If BREC.State = 1 Then BREC.Close
        
        BREC.Open sSql, adoBanco_Dados
        If Not BREC.EOF() Then
            boolTemRegs = True
            
                '' Steel
            
                Frame11.Visible = True
                Frame11.Caption = "[ Aguarde Carregando dados STEEL ]"
                Frame11.Refresh
                
                lngREGS = 0
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    BREC.MoveNext
                Loop
                lngQTDREGSTEEL = lngREGS
                
                BREC.MoveFirst
                prgBAR.Max = lngREGS
                ReDim arrDADOSTAB_STEEL(1 To lngREGS, 1 To 23) As String
                lngREGS = 0
            
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    prgBAR.value = lngREGS
                    
                    arrDADOSTAB_STEEL(lngREGS, 1) = BREC!SGI_CODPED
                    arrDADOSTAB_STEEL(lngREGS, 2) = BREC!SGI_CODIGO
                    arrDADOSTAB_STEEL(lngREGS, 3) = Format(BREC!SGI_DATAPED, "DD/MM/YYYY")
                    arrDADOSTAB_STEEL(lngREGS, 4) = Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY")
                    arrDADOSTAB_STEEL(lngREGS, 5) = Trim(BREC!SGI_RAZAOSOC)
                    arrDADOSTAB_STEEL(lngREGS, 6) = Trim(BREC!SGI_CODPROD)
                    arrDADOSTAB_STEEL(lngREGS, 7) = Trim(BREC!SGI_DESCRICAO)
                    arrDADOSTAB_STEEL(lngREGS, 8) = Format(BREC!SGI_VLUNIT, "#,##0.00")
                    arrDADOSTAB_STEEL(lngREGS, 9) = Trim(BREC!SGI_DESCONDPGTO)
                    arrDADOSTAB_STEEL(lngREGS, 10) = Trim(BREC!SGI_NOMEVEND)
                    arrDADOSTAB_STEEL(lngREGS, 11) = Trim(BREC!SGI_DESCRI)
                    
                    lngQTDPED = BREC!SGI_QTDEPED
                    arrDADOSTAB_STEEL(lngREGS, 12) = lngQTDPED
                    
                    lngQTDFAT = 0
                    If Not IsNull(BREC!SGI_QTDFAT) Then lngQTDFAT = BREC!SGI_QTDFAT
                    arrDADOSTAB_STEEL(lngREGS, 13) = lngQTDFAT
                    
                    lngSALDO = (lngQTDPED - lngQTDFAT)
                    arrDADOSTAB_STEEL(lngREGS, 14) = lngSALDO
                    
                    If BREC!SGI_STATUS = 0 Then
                        arrDADOSTAB_STEEL(lngREGS, 15) = "Aberto"
                    ElseIf BREC!SGI_STATUS = 1 Then
                        arrDADOSTAB_STEEL(lngREGS, 15) = "Parcial"
                    ElseIf BREC!SGI_STATUS = 2 Then
                        arrDADOSTAB_STEEL(lngREGS, 15) = "Total"
                    ElseIf BREC!SGI_STATUS = 3 Then
                        arrDADOSTAB_STEEL(lngREGS, 15) = "Reprovado"
                    ElseIf BREC!SGI_STATUS = 4 Then
                        arrDADOSTAB_STEEL(lngREGS, 15) = "Bloqueada"
                    ElseIf BREC!SGI_STATUS = 6 Then
                        arrDADOSTAB_STEEL(lngREGS, 15) = "P.Cota"
                    ElseIf BREC!SGI_STATUS = 7 Then
                        arrDADOSTAB_STEEL(lngREGS, 15) = "P.Data"
                    ElseIf BREC!SGI_STATUS = 9 Then
                        arrDADOSTAB_STEEL(lngREGS, 15) = "Liq.Man"
                    Else
                        arrDADOSTAB_STEEL(lngREGS, 15) = "S/I"
                    End If
                    
                    arrDADOSTAB_STEEL(lngREGS, 16) = "STEEL"
                    
                    '' Verniz FF
                    arrDADOSTAB_STEEL(lngREGS, 17) = ""
                    arrDADOSTAB_STEEL(lngREGS, 18) = ""
                    arrDADOSTAB_STEEL(lngREGS, 19) = ""
                    arrDADOSTAB_STEEL(lngREGS, 20) = ""
                    
                    If Not IsNull(BREC!SGI_VernCorpo) Then arrDADOSTAB_STEEL(lngREGS, 17) = VernFolhaFrandes(BREC!SGI_VernCorpo)
                    If Not IsNull(BREC!SGI_VernTampa) Then arrDADOSTAB_STEEL(lngREGS, 18) = VernFolhaFrandes(BREC!SGI_VernTampa)
                    If Not IsNull(BREC!SGI_VernFundo) Then arrDADOSTAB_STEEL(lngREGS, 19) = VernFolhaFrandes(BREC!SGI_VernFundo)
                    If Not IsNull(BREC!SGI_VernArgola) Then arrDADOSTAB_STEEL(lngREGS, 20) = VernFolhaFrandes(BREC!SGI_VernArgola)
                    
                    '' Fechamento
                    arrDADOSTAB_STEEL(lngREGS, 21) = ""
                    arrDADOSTAB_STEEL(lngREGS, 22) = ""
                    arrDADOSTAB_STEEL(lngREGS, 23) = ""
                    
                    If Not IsNull(BREC!SGI_FechTampaFuro) Then arrDADOSTAB_STEEL(lngREGS, 21) = PegaFechTampaFuro(BREC!SGI_FechTampaFuro)
                    If Not IsNull(BREC!SGI_NECKIN) Then arrDADOSTAB_STEEL(lngREGS, 22) = IIf(BREC!SGI_NECKIN = 0, "N�o", "Sim")
                    If Not IsNull(BREC!SGI_FechSoldaAgrafado) Then arrDADOSTAB_STEEL(lngREGS, 23) = TipoFecha(BREC!SGI_FechSoldaAgrafado)
                    
                    BREC.MoveNext
                Loop
                Frame11.Visible = False
            
        End If
        BREC.Close
    
    End If
    
    If optEmpresa(0).value = True Then strEMPRESADESC = "NOVALATA"
    If optEmpresa(1).value = True Then strEMPRESADESC = "STEEL ROL"
   
    If boolTemRegs = False Then
       MsgBox "N�o h� dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If

    strNomRel = ""
    Call ExportaParaExcel(arrDADOSTAB, lngQTDREGSNOVA, arrDADOSTAB_STEEL, lngQTDREGSTEEL)

End Sub

Private Function VernFolhaFrandes(lngCODIGO As Long) As String

    VernFolhaFrandes = ""
    
    If lngCODIGO < 1 And lngCODIGO > 4 Then Exit Function
    
    Dim arrVERFECH()    As String
    ReDim arrVERFECH(1 To 4) As String
    
    arrVERFECH(1) = "VEX"
    arrVERFECH(2) = "VZ"
    arrVERFECH(3) = "NAT"
    arrVERFECH(4) = "VI"
    
    VernFolhaFrandes = arrVERFECH(lngCODIGO)


End Function

Private Function PegaFechTampaFuro(lngCODIGO As Long) As String

    PegaFechTampaFuro = ""
    
    If lngCODIGO = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_DESCRI" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADFECHAM" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO

    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF() Then PegaFechTampaFuro = BREC4!SGI_DESCRI
    BREC4.Close
    
End Function

Private Function TipoFecha(lngCODFECH As Long) As String

    TipoFecha = ""
    
    If lngCODFECH < 0 Then Exit Function

    Dim arrFECHAMENTO(0 To 2) As String
    
    arrFECHAMENTO(0) = "SOLDA"
    arrFECHAMENTO(1) = "AGRAFADO"
    arrFECHAMENTO(2) = "REPUXO"
    
    TipoFecha = arrFECHAMENTO(lngCODFECH)

End Function

Private Sub ExportaParaExcel(arrDADOSTAB() As String, lngQTDREGSNOVA As Long, arrDADOSTAB_STEEL() As String, lngQTDRESSTEEL As Long)

On Error GoTo Handle_Error

    Dim myExcelFile             As New clsExcelFile
    Dim FileName$
    Dim boolTEMDADOS            As Boolean
    
    Dim lngREGS                 As Long
    Dim lngLINHA                As Long
    Dim lngQTDPED               As Long
    Dim lngQTDFAT               As Long
    Dim lngSALDO                As Long

    Frame11.Visible = True
    
    If lngQTDREGSNOVA = 0 And lngQTDRESSTEEL = 0 Then
        MsgBox "Aten��o - N�o h� dados para gerar o arquivo !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If

    With myExcelFile
        'Create the new spreadsheet
        If optEmpresa(0).value = True Then
           FileName$ = strCamRelNovo & "RELPREPARA\CARTPROD_NOVALATA.xls"
        ElseIf optEmpresa(1).value = True Then
           FileName$ = strCamRelNovo & "RELPREPARA\CARTPROD_STEEL.xls"
        ElseIf optEmpresa(2).value = True Then
           FileName$ = strCamRelNovo & "RELPREPARA\CARTPROD_NS.xls"
        End If
        
        .CreateFile FileName$
        
        'set a Password for the file. If set, the rest of the spreadsheet will
        'be encrypted. If a password is used it must immediately follow the
        'CreateFile method.
        'This is different then protecting the spreadsheet (see below).
        'NOTE: For some reason this function does not work. Excel will
        'recognize that the file is password protected, but entering the password
        'will not work. Also, the file is not encrypted. Therefore, do not use
        'this function until I can figure out why it doesn't work. There is not
        'much documentation on this function available.
        '.SetFilePassword "PAUL"
        
        'specify whether to print the gridlines or not
        'this should come before the setting of fonts and margins
        .PrintGridLines = False
        
        'it is a good idea to set margins, fonts and column widths
        'prior to writing any text/numerics to the spreadsheet. These
        'should come before setting the fonts.
        
        .SetMargin xlsTopMargin, 1.5   'set to 1.5 inches
        .SetMargin xlsLeftMargin, 1.5
        .SetMargin xlsRightMargin, 1.5
        .SetMargin xlsBottomMargin, 1.5
        
        'to insert a Horizontal Page Break you need to specify the row just
        'after where you want the page break to occur. You can insert as many
        'page breaks as you wish (in any order).
        .InsertHorizPageBreak 10
        .InsertHorizPageBreak 20
        
        'set a default row height for the entire spreadsheet (1/20th of a point)
        .SetDefaultRowHeight 14
        
        'Up to 4 fonts can be specified for the spreadsheet. This is a
        'limitation of the Excel 2.1 format. For each value written to the
        'spreadsheet you can specify which font to use.
        
        .SetFont "Arial", 10, xlsNoFormat              'font0
        .SetFont "Arial", 10, xlsBold                  'font1
        .SetFont "Arial", 10, xlsBold + xlsUnderline   'font2
        .SetFont "Courier", 16, xlsBold + xlsItalic    'font3
        
        'Column widths are specified in Excel as 1/256th of a character.
        
        .SetColumnWidth 1, 4, 18
        .SetColumnWidth 5, 5, 60
        .SetColumnWidth 6, 6, 25
        .SetColumnWidth 7, 7, 60
        
        If (lngCodUsuario = 0 Or lngCodUsuario = 16) Then
            .SetColumnWidth 8, 8, 18
            .SetColumnWidth 9, 9, 30
            .SetColumnWidth 10, 10, 40
            .SetColumnWidth 11, 11, 30
            .SetColumnWidth 12, 16, 18
            .SetColumnWidth 17, 20, 18
            .SetColumnWidth 21, 23, 18
        Else
            .SetColumnWidth 8, 8, 40
            .SetColumnWidth 9, 9, 30
            .SetColumnWidth 10, 14, 18
            .SetColumnWidth 15, 18, 18
            .SetColumnWidth 19, 21, 18
        End If
        
        
        'Set special row heights for row 1 and 2
        ''.SetRowHeight 1, 30
        ''.SetRowHeight 2, 30
        
        'set any header or footer that you want to print on
        'every page. This text will be centered at the top and/or
        'bottom of each page. The font will always be the font that
        'is specified as font0, therefore you should only set the
        'header/footer after specifying the fonts through SetFont.
        ''.SetHeader "BIFF 2.1 API"
        ''.SetFooter "Paul Squires - Excel BIFF Class"
        
        'write a normal left aligned string using font3 (Courier Italic)
        ''.WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsNormal, 1, 1, "Quarterly Report"
        ''.WriteValue xlsText, xlsFont1, xlsLeftAlign, xlsNormal, 2, 1, "Cool Guy Corporation"
        
        'write some data to the spreadsheet
        'Use the default format #3 "#,##0" (refer to the WriteDefaultFormats function)
        'The WriteDefaultFormats function is compliments of Dieter Hauk in Germany.
        ''.WriteValue xlsinteger, xlsFont0, xlsLeftAlign, xlsNormal, 6, 1, 2000, 3
        
        'write a cell with a shaded number with a bottom border
        ''.WriteValue xlsnumber, xlsFont1, xlsrightAlign + xlsBottomBorder + xlsShaded, xlsNormal, 7, 1, 12123.456, 4
        
        'write a normal left aligned string using font2 (bold & underline)
        ''.WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 8, 1, "This is a test string"
        
        'write a locked cell. The cell will not be able to be overwritten, BUT you
        'must set the sheet PROTECTION to on before it will take effect!!!
        ''.WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsLocked, 9, 1, "This cell is locked"
        
        'fill the cell with "F"'s
        ''.WriteValue xlsText, xlsFont0, xlsFillCell, xlsNormal, 10, 1, "F"
        
        'write a hidden cell to the spreadsheet. This only works for cells
        'that contain formula. Text, Number, Integer value text can not be hidden
        'using this feature. It is included here for the sake of completeness.
        ''.WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsHidden, 11, 1, "If this were a formula it would be hidden!"
        
        'write some dates to the file. NOTE: you need to write dates as xlsNumber
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 1, "N.Pedido", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 2, "C�d.OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 3, "R�tulo", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 4, "Descri��o do R�tulo", 12
        .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 5, "Prazo", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 6, "Qtde", 12
        .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 7, "Status", 12
        
        If lngQTDREGSNOVA > 0 Then
        
            '' Jogando os Dados na Planilha
            '' NOVALATA
            Frame11.Caption = "[ Aguarde ... Gerando arguivo EXCEL com dados da NOVALATA ]"
            Frame11.Refresh
            
            lngLINHA = 1
            prgBAR.Min = 0
            prgBAR.Max = UBound(arrDADOSTAB)
            
            For lngREGS = 1 To UBound(arrDADOSTAB) '' Novalata
                lngLINHA = (lngLINHA + 1)
                
                .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 1, arrDADOSTAB(lngREGS, 1), 1
                .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 2, arrDADOSTAB(lngREGS, 2), 1
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 3, arrDADOSTAB(lngREGS, 3), 12
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 4, arrDADOSTAB(lngREGS, 4), 12
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 5, arrDADOSTAB(lngREGS, 5), 12
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 6, arrDADOSTAB(lngREGS, 6), 12
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 7, arrDADOSTAB(lngREGS, 7), 12
        
                If lngCodUsuario = 0 Or lngCodUsuario = 16 Then
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 8, arrDADOSTAB(lngREGS, 8), 2
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 9, arrDADOSTAB(lngREGS, 9), 12
                
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 10, arrDADOSTAB(lngREGS, 10), 12
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 11, arrDADOSTAB(lngREGS, 11), 12
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 12, arrDADOSTAB(lngREGS, 12), 1
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 13, arrDADOSTAB(lngREGS, 13), 1
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 14, arrDADOSTAB(lngREGS, 14), 1
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 15, arrDADOSTAB(lngREGS, 15), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 16, arrDADOSTAB(lngREGS, 16), 12
                    
                    '' Verniz FF
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 17, arrDADOSTAB(lngREGS, 17), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 18, arrDADOSTAB(lngREGS, 18), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 19, arrDADOSTAB(lngREGS, 19), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 20, arrDADOSTAB(lngREGS, 20), 12
                
                    '' Fechamento
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 21, arrDADOSTAB(lngREGS, 21), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 22, arrDADOSTAB(lngREGS, 22), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 23, arrDADOSTAB(lngREGS, 23), 12
                
                Else
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 8, arrDADOSTAB(lngREGS, 10), 12
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 9, arrDADOSTAB(lngREGS, 11), 12
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 10, arrDADOSTAB(lngREGS, 12), 1
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 11, arrDADOSTAB(lngREGS, 13), 1
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 12, arrDADOSTAB(lngREGS, 14), 1
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 13, arrDADOSTAB(lngREGS, 15), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 14, arrDADOSTAB(lngREGS, 16), 12
                
                    '' Verniz FF
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 15, arrDADOSTAB(lngREGS, 17), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 16, arrDADOSTAB(lngREGS, 18), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 17, arrDADOSTAB(lngREGS, 19), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 18, arrDADOSTAB(lngREGS, 20), 12
                
                    '' Fechamento
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 19, arrDADOSTAB(lngREGS, 21), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 20, arrDADOSTAB(lngREGS, 22), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 21, arrDADOSTAB(lngREGS, 23), 12
                
                End If
                
                prgBAR.value = lngREGS
            Next lngREGS
        
        End If
        
        If lngQTDRESSTEEL > 0 Then
            
            Frame11.Caption = "[ Aguarde ... Gerando arguivo EXCEL com dados da STEEL ]"
            Frame11.Refresh
            
            prgBAR.Min = 0
            prgBAR.Max = UBound(arrDADOSTAB_STEEL)
            
            For lngREGS = 1 To UBound(arrDADOSTAB_STEEL) '' Steel
                lngLINHA = (lngLINHA + 1)
                
                .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 1, arrDADOSTAB_STEEL(lngREGS, 1), 1
                .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 2, arrDADOSTAB_STEEL(lngREGS, 2), 1
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 3, arrDADOSTAB_STEEL(lngREGS, 3), 12
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 4, arrDADOSTAB_STEEL(lngREGS, 4), 12
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 5, arrDADOSTAB_STEEL(lngREGS, 5), 12
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 6, arrDADOSTAB_STEEL(lngREGS, 6), 12
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 7, arrDADOSTAB_STEEL(lngREGS, 7), 12
                
                If lngCodUsuario = 0 Or lngCodUsuario = 16 Then
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 8, arrDADOSTAB_STEEL(lngREGS, 8), 2
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 9, arrDADOSTAB_STEEL(lngREGS, 9), 12
                
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 10, arrDADOSTAB_STEEL(lngREGS, 10), 12
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 11, arrDADOSTAB_STEEL(lngREGS, 11), 12
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 12, arrDADOSTAB_STEEL(lngREGS, 12), 1
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 13, arrDADOSTAB_STEEL(lngREGS, 13), 1
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 14, arrDADOSTAB_STEEL(lngREGS, 14), 1
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 15, arrDADOSTAB_STEEL(lngREGS, 15), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 16, arrDADOSTAB_STEEL(lngREGS, 16), 12
                
                    '' Verniz FF
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 17, arrDADOSTAB_STEEL(lngREGS, 17), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 18, arrDADOSTAB_STEEL(lngREGS, 18), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 19, arrDADOSTAB_STEEL(lngREGS, 19), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 20, arrDADOSTAB_STEEL(lngREGS, 20), 12
                
                    '' FEchamento
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 21, arrDADOSTAB_STEEL(lngREGS, 21), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 22, arrDADOSTAB_STEEL(lngREGS, 22), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 23, arrDADOSTAB_STEEL(lngREGS, 23), 12
                
                Else
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 8, arrDADOSTAB_STEEL(lngREGS, 10), 12
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 9, arrDADOSTAB_STEEL(lngREGS, 11), 12
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 10, arrDADOSTAB_STEEL(lngREGS, 12), 1
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 11, arrDADOSTAB_STEEL(lngREGS, 13), 1
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 12, arrDADOSTAB_STEEL(lngREGS, 14), 1
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 13, arrDADOSTAB_STEEL(lngREGS, 15), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 14, arrDADOSTAB_STEEL(lngREGS, 16), 12
                
                    '' Verniz FF
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 15, arrDADOSTAB_STEEL(lngREGS, 17), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 16, arrDADOSTAB_STEEL(lngREGS, 18), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 17, arrDADOSTAB_STEEL(lngREGS, 19), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 18, arrDADOSTAB_STEEL(lngREGS, 20), 12
                
                    '' FEchamento
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 19, arrDADOSTAB_STEEL(lngREGS, 21), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 20, arrDADOSTAB_STEEL(lngREGS, 22), 12
                    .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 21, arrDADOSTAB_STEEL(lngREGS, 23), 12
                
                End If
                
                 prgBAR.value = lngREGS
            Next lngREGS
        End If
        
        'PROTECT the spreadsheet so any cells specified as LOCKED will not be
        'overwritten. Also, all cells with HIDDEN set will hide their formula.
        'PROTECT does not use a password.
        .ProtectSpreadsheet = False 'False | True
        
        'Finally, close the spreadsheet
        .CloseFile
        Frame11.Visible = False
        
        MsgBox "Arquivo Excel : REMPEDCART.xls foi Criado !", vbInformation + vbOKOnly, "Aviso do Sistema"
    End With
    
    Exit Sub
    
Handle_Error:

    If BREC.State = 1 Then BREC.Close
    MsgBox "N�mero: " & Err.Number & vbCrLf & "Descri��o: " & Err.Description, vbOKOnly + vbCritical, "Aviso"

        
End Sub

