VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRELPP 
   Caption         =   "Relatório de Preços e Prazos"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   9855
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
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
      TabIndex        =   16
      Top             =   2160
      Width           =   9735
      Begin ComctlLib.ProgressBar prbDados 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   9735
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCIDCLIE"
         Top             =   255
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   1320
         Picture         =   "frmRELPP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDescCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Frame Frame4 
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
      Left            =   4920
      TabIndex        =   12
      Top             =   840
      Width           =   4815
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optEmpresa 
         Caption         =   "STEEL"
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
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optEmpresa 
         Caption         =   "NOVALATA E STEEL"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Ultima Compra ]"
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
      TabIndex        =   9
      Top             =   840
      Width           =   4935
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3600
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
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
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9735
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
         Picture         =   "frmRELPP.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
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
         Picture         =   "frmRELPP.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public lngCodUsuario    As Long

Dim objBLBFunc          As Object
Dim objRELCOMPPRECOS    As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object

Dim strCABEC1           As String
Dim strCABEC2           As String
Dim strNomRel           As String
Dim strEMPRESADESC      As String
Dim strEMPRESA          As String

Dim arrCLIENTESNOVALATA()   As PPCLIENTES
Dim arrCLIENTESSTEEL()      As PPCLIENTES
Dim arrDADOS()              As PPCLIENTES

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call GeraArquivoExcel
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

On Error GoTo Err_Command1_Click

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
    If Len(Trim(lblDescCliente.Caption)) = 0 Then
       txtCIDCLIE.Text = ""
       Call LimpaCamposLabel
    End If
        
    txtCIDCLIE.SetFocus

    Exit Sub
    
Err_Command1_Click:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : Command1_Click()", Me.Name, "Command1_Click()", strCAMARQERRO)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELCOMPPRECOS = CreateObject("RELCOMERCIAL.clsRELCOMPPRECOS")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    
    objRELCOMPPRECOS.FILIAL = FILIAL
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    Me.Caption = Me.Caption & " / " & Trim(Me.Name)


    optEmpresa(0).value = True
    Frame3.Caption = ""
    Frame3.Refresh
    prbDados.Min = 0

    Call LimpaCamposLabel

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELCOMPPRECOS = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub LimpaCamposLabel()
    lblDescCliente.Caption = ""
End Sub


Private Sub txtCIDCLIE_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, Me
End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCIDCLIE.Text
End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCIDCLIE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCIDCLIE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente)
    If Len(Trim(lblDescCliente.Caption)) = 0 Then
       txtCIDCLIE.Text = ""
       Call LimpaCamposLabel
       Cancel = True
       Exit Sub
    End If
    
    Exit Sub

End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label)

On Error GoTo Err_PegaDescTabelas

    lblLabel.Caption = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
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
    
    Exit Sub
    
Err_PegaDescTabelas:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
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

Private Sub GeraArquivoExcel()

On Error GoTo Err_Exporta

    Dim boolTemDados            As Boolean
    Dim boolTemDadosNOVA        As Boolean
    Dim boolTemDadosSTEEL       As Boolean
    
    Dim lngQTDREGS              As Long
    Dim lngTOTREGS              As Long
    Dim lngTOTREGSNOVA          As Long
    Dim lngTOTREGSSTEEL         As Long
    Dim lngREGDADOS             As Long
    
    Dim strEMPRESA              As String
    Dim strEMPARQ               As String
    Dim strNomRel               As String
    Dim i                       As Long
    Dim arrULTDATCOMPRA()       As String
    Dim arrITENSPEDIDO()        As PPPRODUTOS
    
    strEMPRESA = ""
    strEMPARQ = "NOVALATA"
    If (optEmpresa(1).value = True Or optEmpresa(2).value = True) Then
       strEMPRESA = "_STEEL"
       strEMPARQ = "STEEL"
    End If
    
    boolTemDados = False
    boolTemDadosNOVA = False
    boolTemDadosSTEEL = False
    
    prbDados.Visible = False
    
    '' Novalata
    lngTOTREGSNOVA = 0
    If (optEmpresa(0).value = True Or optEmpresa(2).value = True) Then
        sSql = ""
        
        sSql = "Select  Distinct" & vbCrLf
        sSql = sSql & "        SGI_CODCLI" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "        SGI_CADPEDVENDH" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'"
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            boolTemDadosNOVA = True
            lngTOTREGSNOVA = (lngTOTREGSNOVA + 1)
            BREC.MoveNext
        Loop
        BREC.Close
    End If
    
    '' Steel
    lngTOTREGSSTEEL = 0
    If (optEmpresa(1).value = True Or optEmpresa(2).value = True) Then

        sSql = ""
        
        sSql = "Select  Distinct" & vbCrLf
        sSql = sSql & "        SGI_CODCLI" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "        SGI_CADPEDVENDH" & strEMPRESA & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'"
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            boolTemDadosSTEEL = True
            lngTOTREGSSTEEL = (lngTOTREGSSTEEL + 1)
            BREC.MoveNext
        Loop
        BREC.Close

    End If
    
    If (boolTemDadosNOVA = True Or boolTemDadosSTEEL = True) Then boolTemDados = True
    
    If boolTemDados = False Then
        MsgBox "ATENÇÃO - Não há dados neste periodo !!!", vbOKOnly + vbExclamation, "Aviso do Sistema"
        Exit Sub
    End If
    
    
    '' Gerando Dados Novalata
    If boolTemDadosNOVA = True Then
       
        prbDados.Min = 0
        prbDados.Max = lngTOTREGSNOVA
        prbDados.Visible = True
     
        ReDim arrDADOS(1 To lngTOTREGSNOVA) As PPCLIENTES
       
        Frame3.Caption = "[ AGUARDE. Processando dados NOVALATA... ]"
        Frame3.Refresh
       
       
        sSql = ""
        
        sSql = "Select  Distinct" & vbCrLf
        sSql = sSql & "        PED.SGI_CODCLI" & vbCrLf
        sSql = sSql & "       ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "        SGI_CADPEDVENDH PED" & vbCrLf
        sSql = sSql & "       ,SGI_CADCLIENTE  CLIE" & vbCrLf
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "        PED.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "    And PED.SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "    And CLIE.SGI_FILIAL = PED.SGI_FILIAL" & vbCrLf
        sSql = sSql & "    And CLIE.SGI_CODIGO = PED.SGI_CODCLI"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        lngQTDREGS = 0
        Do While Not BREC.EOF()
                lngQTDREGS = (lngQTDREGS + 1)
                prbDados.value = lngQTDREGS
                
                arrDADOS(lngQTDREGS).lngCodClie = BREC!SGI_CODCLI
                arrDADOS(lngQTDREGS).strRAZAOSOC = BREC!SGI_RAZAOSOC
                
                arrULTDATCOMPRA = Split(PegaUltCompra(Str(BREC!SGI_CODCLI), ""), "|")
                
                If IsNumeric(arrULTDATCOMPRA(0)) Then arrDADOS(lngQTDREGS).strDescPgto = Trim(Str(arrULTDATCOMPRA(0)))
                If Not IsNumeric(arrULTDATCOMPRA(0)) Then arrDADOS(lngQTDREGS).strDescPgto = Trim(arrULTDATCOMPRA(0))
                
                arrDADOS(lngQTDREGS).DtUltimaCompra = arrULTDATCOMPRA(1)
                arrDADOS(lngQTDREGS).lngCODPEDIDO = CLng(arrULTDATCOMPRA(2))
                arrDADOS(lngQTDREGS).lngCodCondPgto = CLng(arrULTDATCOMPRA(3))
             
                '' Itens
                lngREGDADOS = 0
                sSql = ""
                
                sSql = "Select" & vbCrLf
                sSql = sSql & "       PEDI.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "      ,PEDI.SGI_CODPROD" & vbCrLf
                sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
                sSql = sSql & "      ,PEDI.SGI_VLUNIT" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDI PEDI" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO  PROD" & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       PEDI.SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & " And   PEDI.SGI_CODIGO    = " & Trim(arrULTDATCOMPRA(2)) & vbCrLf
                sSql = sSql & " And   PROD.SGI_FILIAL    = PEDI.SGI_FILIAL" & vbCrLf
                sSql = sSql & " And   PROD.SGI_IDPRODUTO = PEDI.SGI_IDPRODUTO"
                
                BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC3.EOF() Then
                    Do While Not BREC3.EOF()
                        lngREGDADOS = (lngREGDADOS + 1)
                        BREC3.MoveNext
                    Loop
                    
                    ReDim arrITENSPEDIDO(1 To lngREGDADOS) As PPPRODUTOS
                    
                    lngREGDADOS = 0
                    BREC3.MoveFirst
                    
                    Do While Not BREC3.EOF()
                        lngREGDADOS = (lngREGDADOS + 1)
                        
                        arrITENSPEDIDO(lngREGDADOS).lngIDProduto = BREC3!SGI_IDPRODUTO
                        arrITENSPEDIDO(lngREGDADOS).strCODPRODUTO = BREC3!SGI_CODPROD
                        arrITENSPEDIDO(lngREGDADOS).strDESCPROD = BREC3!SGI_DESCRICAO
                        arrITENSPEDIDO(lngREGDADOS).strPRECO = Format(BREC3!SGI_VLUNIT, "#,##0.00")
                        
                        BREC3.MoveNext
                    Loop
                End If
                BREC3.Close
                
                arrDADOS(lngQTDREGS).lngQTDITENS = lngREGDADOS
                If lngREGDADOS > 0 Then arrDADOS(lngQTDREGS).arrPRODUTOS = arrITENSPEDIDO
                
                
                BREC.MoveNext
        Loop
        BREC.Close
    
        arrCLIENTESNOVALATA = arrDADOS
    
    End If
    
    
    '' Gerando Dados Steel
    If boolTemDadosSTEEL = True Then
       
        prbDados.Min = 0
        prbDados.Max = lngTOTREGSSTEEL
        prbDados.Visible = True
     
        ReDim arrDADOS(1 To lngTOTREGSSTEEL) As PPCLIENTES
       
        Frame3.Caption = "[ AGUARDE. Processando dados STEEL... ]"
        Frame3.Refresh
       
        sSql = ""
        
        sSql = "Select  Distinct" & vbCrLf
        sSql = sSql & "        PED.SGI_CODCLI" & vbCrLf
        sSql = sSql & "       ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "        SGI_CADPEDVENDH" & strEMPRESA & " PED" & vbCrLf
        sSql = sSql & "       ,SGI_CADCLIENTE  CLIE" & vbCrLf
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "        PED.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "    And PED.SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "    And CLIE.SGI_FILIAL = PED.SGI_FILIAL" & vbCrLf
        sSql = sSql & "    And CLIE.SGI_CODIGO = PED.SGI_CODCLI"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        lngQTDREGS = 0
        Do While Not BREC.EOF()
                lngQTDREGS = (lngQTDREGS + 1)
                prbDados.value = lngQTDREGS
                
                arrDADOS(lngQTDREGS).lngCodClie = BREC!SGI_CODCLI
                arrDADOS(lngQTDREGS).strRAZAOSOC = BREC!SGI_RAZAOSOC
                
                arrULTDATCOMPRA = Split(PegaUltCompra(Str(BREC!SGI_CODCLI), strEMPRESA), "|")
                
                If IsNumeric(arrULTDATCOMPRA(0)) Then arrDADOS(lngQTDREGS).strDescPgto = Trim(Str(arrULTDATCOMPRA(0)))
                If Not IsNumeric(arrULTDATCOMPRA(0)) Then arrDADOS(lngQTDREGS).strDescPgto = Trim(arrULTDATCOMPRA(0))
                
                arrDADOS(lngQTDREGS).DtUltimaCompra = arrULTDATCOMPRA(1)
                arrDADOS(lngQTDREGS).lngCODPEDIDO = CLng(arrULTDATCOMPRA(2))
                arrDADOS(lngQTDREGS).lngCodCondPgto = CLng(arrULTDATCOMPRA(3))
             
                '' Itens
                lngREGDADOS = 0
                sSql = ""
                
                sSql = "Select" & vbCrLf
                sSql = sSql & "       PEDI.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "      ,PEDI.SGI_CODPROD" & vbCrLf
                sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
                sSql = sSql & "      ,PEDI.SGI_VLUNIT" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDI" & strEMPRESA & " PEDI" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO  PROD" & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       PEDI.SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & " And   PEDI.SGI_CODIGO    = " & Trim(arrULTDATCOMPRA(2)) & vbCrLf
                sSql = sSql & " And   PROD.SGI_FILIAL    = PEDI.SGI_FILIAL" & vbCrLf
                sSql = sSql & " And   PROD.SGI_IDPRODUTO = PEDI.SGI_IDPRODUTO"
                
                BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC3.EOF() Then
                    Do While Not BREC3.EOF()
                        lngREGDADOS = (lngREGDADOS + 1)
                        BREC3.MoveNext
                    Loop
                    
                    ReDim arrITENSPEDIDO(1 To lngREGDADOS) As PPPRODUTOS
                    
                    lngREGDADOS = 0
                    BREC3.MoveFirst
                    
                    Do While Not BREC3.EOF()
                        lngREGDADOS = (lngREGDADOS + 1)
                        
                        arrITENSPEDIDO(lngREGDADOS).lngIDProduto = BREC3!SGI_IDPRODUTO
                        arrITENSPEDIDO(lngREGDADOS).strCODPRODUTO = BREC3!SGI_CODPROD
                        arrITENSPEDIDO(lngREGDADOS).strDESCPROD = BREC3!SGI_DESCRICAO
                        arrITENSPEDIDO(lngREGDADOS).strPRECO = Format(BREC3!SGI_VLUNIT, "#,##0.00")
                        
                        BREC3.MoveNext
                    Loop
                End If
                BREC3.Close
                
                arrDADOS(lngQTDREGS).lngQTDITENS = lngREGDADOS
                If lngREGDADOS > 0 Then arrDADOS(lngQTDREGS).arrPRODUTOS = arrITENSPEDIDO
                
                
                BREC.MoveNext
        Loop
        BREC.Close
    
        arrCLIENTESSTEEL = arrDADOS
    
    End If
    
    
    '' Gerando o Arquivo XLS
    If optEmpresa(0).value = True Then strNomRel = "RELPP_NOVALATA.xls"
    If optEmpresa(1).value = True Then strNomRel = "RELPP_STEEL.xls"
    If optEmpresa(2).value = True Then strNomRel = "RELPP_NE.xls"
    
    '' Gerando o Arquivo Excel
    Call GeraArgExcel(strNomRel, boolTemDadosNOVA, boolTemDadosSTEEL)
    
    MsgBox "Arquivo gerado com sucesso !", vbOKOnly + vbInformation, "Aviso"
    
    Frame3.Caption = ""
    Frame3.Refresh
    prbDados.Visible = False
    
    Exit Sub
    
Err_Exporta:

    MsgBox "Erro      : " & Err.Number & vbCrLf & _
           "Descrição : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"

End Sub


Private Sub GeraArgExcel(strARQUIVO As String, boolTemDadosNOVA As Boolean, boolTemDadosSTEEL As Boolean)

On Error GoTo err_Excel

    Dim myExcelFile             As New clsExcelFile
    Dim FileName$
    Dim lngLINHA                As Long
    Dim lngREGS                 As Long
    Dim lngREGSCAPAC            As Long
    Dim arrDADOS                As Variant
    Dim lngCol                  As Long
    Dim lngColAno               As Long
    Dim i                       As Long
    Dim lngQTDANOS              As Long
    Dim dblPreco                As Double
    
    
    With myExcelFile
        
        FileName$ = strCamRelNovo & "RELPREPARA\" & strARQUIVO
        
        .CreateFile FileName$
        
        .PrintGridLines = False
        
        .SetMargin xlsTopMargin, 1.5   'set to 1.5 inches
        .SetMargin xlsLeftMargin, 1.5
        .SetMargin xlsRightMargin, 1.5
        .SetMargin xlsBottomMargin, 1.5
        
        .InsertHorizPageBreak 10
        .InsertHorizPageBreak 20
        
        .SetDefaultRowHeight 14
        
        .SetFont "Arial", 10, xlsNoFormat              'font0
        .SetFont "Arial", 10, xlsBold                  'font1
        .SetFont "Arial", 10, xlsBold + xlsUnderline   'font2
        .SetFont "Courier", 16, xlsBold + xlsItalic    'font3
        
        'Column widths are specified in Excel as 1/256th of a character.
        '               L,  C,  T
        
        'write some dates to the file. NOTE: you need to write dates as xlsNumber
        lngCol = 1
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 60
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Razão Social", 12
        
        lngCol = (lngCol + 1)
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 30
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Cód. Item", 12
        
        lngCol = (lngCol + 1)
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 60
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Descrição do Item", 12
        
        lngCol = (lngCol + 1)
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 30
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Prazo de Pagamento", 12
        
        lngCol = (lngCol + 1)
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 30
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Valor", 12
        
        lngCol = (lngCol + 1)
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 30
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Ultima Compra", 12
        
        lngCol = (lngCol + 1)
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 30
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Empresa", 12
        
        '' Dados da Novalata
        If boolTemDadosNOVA = True Then
             lngLINHA = 1
             Frame3.Caption = "[ Aguarde.... Gerando o Arguivo EXCEL com os dados da NOVALATA ! ]"
             Frame3.Refresh
             
             prbDados.Min = 0
             prbDados.Max = UBound(arrCLIENTESNOVALATA)
             
             For lngREGS = 1 To UBound(arrCLIENTESNOVALATA)
                prbDados.value = lngREGS
                
                For i = 1 To arrCLIENTESNOVALATA(lngREGS).lngQTDITENS
                    
                    lngCol = 1
                    lngLINHA = (lngLINHA + 1)
                    
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).strRAZAOSOC, 12                  '' RAZÃO SOCIAL
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).arrPRODUTOS(i).strCODPRODUTO, 12 '' Cód. Item
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).arrPRODUTOS(i).strDESCPROD, 12   '' Descrição Item
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).strDescPgto, 12                  '' Condição de Pagamento
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).arrPRODUTOS(i).strPRECO, 2    '' Valor
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).DtUltimaCompra, 12              '' Data Ultima Compra
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, "NOVALATA", 12                                                '' Empresa
                
                Next i
                
             
             Next lngREGS
        End If
    
    
        '' Dados da Steel
        If boolTemDadosNOVA = False Then lngLINHA = 1
        
        If boolTemDadosSTEEL = True Then
             Frame3.Caption = "[ Aguarde.... Gerando o Arguivo EXCEL com os dados da STEEL ! ]"
             Frame3.Refresh
             
             prbDados.Min = 0
             prbDados.Max = UBound(arrCLIENTESSTEEL)
             
             For lngREGS = 1 To UBound(arrCLIENTESSTEEL)
                prbDados.value = lngREGS
                
                For i = 1 To arrCLIENTESSTEEL(lngREGS).lngQTDITENS
                    
                    lngCol = 1
                    lngLINHA = (lngLINHA + 1)
                    
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).strRAZAOSOC, 12                  '' RAZÃO SOCIAL
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).arrPRODUTOS(i).strCODPRODUTO, 12 '' Cód. Item
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).arrPRODUTOS(i).strDESCPROD, 12   '' Descrição Item
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).strDescPgto, 12                  '' Condição de Pagamento
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).arrPRODUTOS(i).strPRECO, 2    '' Valor
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).DtUltimaCompra, 12              '' Data Ultima Compra
                    
                    lngCol = (lngCol + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, "STEEL", 12                                                '' Empresa
                
                Next i
                
             
             Next lngREGS
        End If
    
        .ProtectSpreadsheet = False 'False | True
        .CloseFile
    
    End With

    Exit Sub

err_Excel:

    MsgBox "ATENÇÃO" & vbCrLf & _
           "Erro Numero       : " & Err.Number & vbCrLf & _
           "Descrição do Erro : " & Err.Description, vbOKOnly + vbCritical, "Aviso"

End Sub

Private Function PegaUltCompra(strCODCLI As String, strNOMEMP As String) As String
    
    PegaUltCompra = ""
    
    sSql = ""
    
    sSql = " Select" & vbCrLf
    
    sSql = sSql & "        PED.SGI_CODIGO" & vbCrLf
    sSql = sSql & "       ,PED.SGI_DATAPED" & vbCrLf
    sSql = sSql & "       ,PED.SGI_CODCONDPGT" & vbCrLf
    sSql = sSql & "       ,PGTO.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "   From" & vbCrLf
    sSql = sSql & "        SGI_CADPEDVENDH" & strNOMEMP & " PED" & vbCrLf
    sSql = sSql & "       ,SGI_CADCONDPGTO PGTO" & vbCrLf
    
    sSql = sSql & "  Where" & vbCrLf
    sSql = sSql & "        PED.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "    And PED.SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "    And PED.SGI_CODCLI = " & strCODCLI & vbCrLf
    sSql = sSql & "    And PGTO.SGI_FILIAL = PED.SGI_FILIAL" & vbCrLf
    sSql = sSql & "    And PGTO.SGI_CODIGO = PED.SGI_CODCONDPGT" & vbCrLf
    sSql = sSql & "Order By PED.SGI_DATAPED Desc"
     
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then PegaUltCompra = BREC2!SGI_DESCRICAO & "|" & Format(BREC2!SGI_DATAPED, "DD/MM/YYYY") & "|" & BREC2!SGI_CODIGO & "|" & BREC2!SGI_CODCONDPGT
    BREC2.Close
    
End Function

