VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCURVABCLIENTE 
   Caption         =   "Curva ABC de Cliente"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   11205
   StartUpPosition =   1  'CenterOwner
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
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   1560
      Width           =   11175
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   1320
         Picture         =   "frmCURVABCLIENTE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   17
         Text            =   "txtCIDCLIE"
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label lblDescCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   11175
      Begin ComctlLib.ProgressBar prgBAR 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
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
      Left            =   8640
      TabIndex        =   11
      Top             =   960
      Width           =   2535
      Begin VB.OptionButton optEmpresa 
         Caption         =   "Steel"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optEmpresa 
         Caption         =   "Novalata"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Tipo ]"
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
      TabIndex        =   8
      Top             =   960
      Width           =   3015
      Begin VB.OptionButton optTIPO 
         Caption         =   "Quantidade"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optTIPO 
         Caption         =   "Valor"
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
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   975
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
      TabIndex        =   5
      Top             =   960
      Width           =   5415
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
         Left            =   1440
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11175
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
         Picture         =   "frmCURVABCLIENTE.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCURVABCLIENTE.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCURVABCLIENTE"
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
Dim objCURVABCLIENTE    As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String
Dim strNomRel           As String
Dim strEMPRESADESC      As String
Dim strEMPRESA          As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    If optTIPO(1).value = True Then Call ImprimeCurvABCQTDE
    If optTIPO(0).value = True Then Call ImprimeCurvABCVALOR
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
    If Len(Trim(lblDescCliente.Caption)) = 0 Then txtCIDCLIE.Text = ""
    
    txtCIDCLIE.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCURVABCLIENTE = CreateObject("RELCOMERCIAL.clsCURVABCLIENTE")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    
    objCURVABCLIENTE.FILIAL = FILIAL
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    optEmpresa(0).value = True
    optTIPO(1).value = True
    
    Me.Caption = Me.Caption & " / " & Trim(Me.Name)

    prgBAR.Min = 0
    Frame5.Caption = ""
    lblDescCliente.Caption = ""

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCURVABCLIENTE = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub


Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, Me
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, Me
End Sub

Private Sub optEmpresa_Click(Index As Integer)
    If Index = 0 Then strEMPRESA = ""
    If Index = 1 Then strEMPRESA = "_STEEL"
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
    
    If Year(CDate(mskDTINI.Text)) <> Year(CDate(mskDTFIN.Text)) Then
        MsgBox "O Ano não pode ser diferente !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTINI.SetFocus
        Exit Function
    End If
    
    ConfereCampos = True

End Function

Private Sub ImprimeCurvABCQTDE()
    
    Dim lngQTDEMES      As Long
    Dim lngQTDECLIE     As Long
    Dim lngQTDREGS      As Long
    Dim lngTOTCAPAC     As Long
    Dim lngMESINI       As Long
    Dim lngMESFIN       As Long
    Dim lngTOTMESES     As Long
    
    Dim arrCLIENTES()   As ABCCLIENTES
    Dim arrCAPACIDADE() As CAPACIDADE
    Dim arrMESES()      As MESES
    Dim arrDESCMES()    As String
    
    Dim i               As Long
    Dim J               As Long
    Dim K               As Long
    
    Dim strCAMPO01      As String
    Dim strCAMPO02      As String
    Dim strCAMPO03      As String
    Dim strCAMPO04      As String
    
    Frame5.Caption = ""
    Frame5.Refresh
    
    lngMESINI = 0
    If Month(CDate(mskDTINI.Text)) > 1 Then lngMESINI = Month(CDate(mskDTINI.Text))
    lngMESFIN = Month(CDate(mskDTFIN.Text))
    lngTOTMESES = (lngMESFIN - lngMESINI)
     
    ReDim arrMESES(1 To 12) As MESES
    ReDim arrDESCMES(1 To 12) As String
    arrDESCMES(1) = "Janeiro"
    arrDESCMES(2) = "Fevereiro"
    arrDESCMES(3) = "Marco"
    arrDESCMES(4) = "Abril"
    arrDESCMES(5) = "Maio"
    arrDESCMES(6) = "Junho"
    arrDESCMES(7) = "Julho"
    arrDESCMES(8) = "Agosto"
    arrDESCMES(9) = "Setembro"
    arrDESCMES(10) = "Outubro"
    arrDESCMES(11) = "Novembro"
    arrDESCMES(12) = "Dezembro"
    
    lngQTDECLIE = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Count(SGI_CODIGO) as Qtde_Regs" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        sSql = sSql & "   And SGI_CODIGO = " & txtCIDCLIE.Text
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then lngQTDECLIE = BREC!Qtde_Regs
    BREC.Close
    
    If lngQTDECLIE = 0 Then Exit Sub
    
    prgBAR.Visible = True
    prgBAR.Min = 0
    prgBAR.Max = lngQTDECLIE
    
    ReDim Preserve arrCLIENTES(1 To lngQTDECLIE) As ABCCLIENTES
    
    lngQTDECLIE = 0
    
    '' ----------------------------
    '' Pop Array Clientes
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_CPFCNPJ" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        sSql = sSql & "   And SGI_CODIGO = " & txtCIDCLIE.Text
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
        lngQTDECLIE = (lngQTDECLIE + 1)
        arrCLIENTES(lngQTDECLIE).lngCodClie = BREC!SGI_CODIGO
        arrCLIENTES(lngQTDECLIE).strRAZAOSOC = BREC!SGI_RAZAOSOC
        arrCLIENTES(lngQTDECLIE).strCNPJ = BREC!SGI_CPFCNPJ
        prgBAR.value = lngQTDECLIE
        
        Frame5.Caption = "[ " & Format(((lngQTDECLIE / prgBAR.Max) * 100), "##00") & "% - Concluido, Processando Aguarde !!! ]"
        Frame5.Refresh
        BREC.MoveNext
    Loop
    BREC.Close
    Frame5.Caption = ""
    Frame5.Refresh
    '' ----------------------------
    
    prgBAR.Min = 0
    prgBAR.Max = UBound(arrCLIENTES)
    
    For i = 1 To UBound(arrCLIENTES)
    
        prgBAR.value = i
        
        sSql = ""
    
        sSql = "Select Distinct" & vbCrLf
        sSql = sSql & "       Linha.SGI_CODIGO" & vbCrLf
        sSql = sSql & "      ,LINHA.SGI_CODLIN" & vbCrLf
        sSql = sSql & "      ,LINHA.SGI_DESCRI" & vbCrLf
        sSql = sSql & " From" & vbCrLf
        
        sSql = sSql & "       SGI_CADCLIENTE      CLIE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & "     VENDH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & "     VENDI" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINHA" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       CLIE.SGI_FILIAl     = " & FILIAL & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO     = " & arrCLIENTES(i).lngCodClie & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL     = VENDH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO     = VENDH.SGI_CODCLI" & vbCrLf
        sSql = sSql & "   And (VENDH.SGI_STATUS   = 'F' Or VENDH.SGI_STATUS = 'P')" & vbCrLf
        sSql = sSql & "   And VENDH.SGI_FILIAL    = VENDI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And VENDH.SGI_CODIGO    = VENDI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And VENDI.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And VENDI.SGI_IDPRODUTO = PROD.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL     = LINHA.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_CODLINPROD = LINHA.SGI_CODLIN" & vbCrLf
        sSql = sSql & "Order By LINHA.SGI_CODLIN"
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           lngQTDREGS = 0
           Do While Not BREC.EOF()
              lngQTDREGS = (lngQTDREGS + 1)
              ReDim Preserve arrCAPACIDADE(1 To lngQTDREGS) As CAPACIDADE
              arrCAPACIDADE(lngQTDREGS).lngCODIGO = BREC!SGI_CODIGO
              arrCAPACIDADE(lngQTDREGS).lngCodLinha = BREC!SGI_CODLIN
              arrCAPACIDADE(lngQTDREGS).strDESCLINHA = "'" & Replace(Replace(BREC!SGI_DESCRI, "Ã", "A"), "¼", "1/4") & "'"
              arrCAPACIDADE(lngQTDREGS).arrMESES = arrMESES
              BREC.MoveNext
           Loop
           arrCLIENTES(i).arrCAPACIDADE = arrCAPACIDADE
           arrCLIENTES(i).lngQTDECAPAC = lngQTDREGS
        End If
        BREC.Close
    
        Frame5.Caption = "[ " & Format(((i / prgBAR.Max) * 100), "##00") & "% - Concluido, Processando Aguarde !!! ]"
        Frame5.Refresh
    Next i
    
    prgBAR.Min = 0
    prgBAR.Max = UBound(arrCLIENTES)
    For i = 1 To UBound(arrCLIENTES)
    
        prgBAR.value = i
        If arrCLIENTES(i).lngQTDECAPAC > 0 Then
            For J = 1 To UBound(arrCLIENTES(i).arrCAPACIDADE)
                
                sSql = ""
                
                sSql = "Select" & vbCrLf
                sSql = sSql & "      Month(CONFH.SGI_DATACONF) As SGI_MES" & vbCrLf
                sSql = sSql & "     ,Sum(CONFI.SGI_QTDREAL)    As SGI_QTDREAL" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "      SGI_CADORDCONFH" & strEMPRESA & " CONFH" & vbCrLf
                sSql = sSql & "     ,SGI_CADORDCONFI" & strEMPRESA & " CONFI" & vbCrLf
                sSql = sSql & "     ,SGI_CADORDFATI" & strEMPRESA & "  FATI" & vbCrLf
                sSql = sSql & "     ,SGI_CADORDFATH" & strEMPRESA & "  FATH" & vbCrLf
                sSql = sSql & "     ,SGI_CADPEDVENDI" & strEMPRESA & " VENDI" & vbCrLf
                sSql = sSql & "     ,SGI_CADPEDVENDH" & strEMPRESA & " VENDH" & vbCrLf
                sSql = sSql & "     ,SGI_CADPRODUTO  PROD" & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "      CONFH.SGI_FILIAL          = " & FILIAL & vbCrLf
                sSql = sSql & "  And CONFH.SGI_DATACONF     Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "  And CONFH.SGI_FILIAL          = CONFI.SGI_FILIAL" & vbCrLf
                sSql = sSql & "  And CONFH.SGI_CODCONF         = CONFI.SGI_CODCONF" & vbCrLf
                sSql = sSql & "  And CONFH.SGI_FILIAL          = FATI.SGI_FILIAL" & vbCrLf
                sSql = sSql & "  And CONFH.SGI_CODORD          = FATI.SGI_CODORD" & vbCrLf
                sSql = sSql & "  And CONFI.SGI_IDPRODUTO       = FATI.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "  And FATI.SGI_FILIAL           = FATH.SGI_FILIAL" & vbCrLf
                sSql = sSql & "  And FATI.SGI_CODORD           = FATH.SGI_CODORD" & vbCrLf
                sSql = sSql & "  And FATH.SGI_FILIAL           = VENDI.SGI_FILIAL" & vbCrLf
                sSql = sSql & "  And FATH.SGI_CODPED           = VENDI.SGI_CODIGO" & vbCrLf
                sSql = sSql & "  And FATI.SGI_IDPRODUTO        = VENDI.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "  And VENDI.SGI_FILIAL          = VENDH.SGI_FILIAL" & vbCrLf
                sSql = sSql & "  And VENDI.SGI_CODIGO          = VENDH.SGI_CODIGO" & vbCrLf
                sSql = sSql & "  And VENDH.SGI_CODCLI          = " & arrCLIENTES(i).lngCodClie & vbCrLf
                sSql = sSql & "  And VENDI.SGI_FILIAL          = PROD.SGI_FILIAL" & vbCrLf
                sSql = sSql & "  And VENDI.SGI_IDPRODUTO       = PROD.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "  And PROD.SGI_CODLINPROD       = " & arrCLIENTES(i).arrCAPACIDADE(J).lngCodLinha & vbCrLf
                sSql = sSql & "Group By Month(CONFH.SGI_DATACONF)" & vbCrLf
                sSql = sSql & "Order BY SGI_MES"
    
                BREC.Open sSql, adoBanco_Dados, adOpenDynamic
                lngTOTCAPAC = 0
                Do While Not BREC.EOF()
                    arrCLIENTES(i).arrCAPACIDADE(J).arrMESES(BREC!SGI_MES).lngQTDE = BREC!SGI_QTDREAL
                    lngTOTCAPAC = (lngTOTCAPAC + BREC!SGI_QTDREAL)
                    BREC.MoveNext
                Loop
                arrCLIENTES(i).arrCAPACIDADE(J).lngTOTALCAPAC = lngTOTCAPAC
                If lngTOTMESES > 0 Then
                    arrCLIENTES(i).arrCAPACIDADE(J).lngMEDIACAPAC = Format((lngTOTCAPAC / lngTOTMESES), "##00")
                End If
                
                BREC.Close
            Next J
        End If
        
        Frame5.Caption = "[ " & Format(((i / prgBAR.Max) * 100), "##00") & "% - Concluido, Processando Aguarde !!! ]"
        Frame5.Refresh
    Next i
    
    
    
    '' ====================================
    '' Montando o Arquivo TXT
    strNomRel = "CURVAABCQUANT" & strEMPRESA & ".txt"
    
    strCAMPO01 = "RAZAO SOCIAL" & vbTab & _
                 "CNPJ" & vbTab & _
                 "LINHA"

    If lngMESINI = 0 Then lngMESINI = 1
    strCAMPO02 = ""
    
    For i = lngMESINI To lngMESFIN
        strCAMPO02 = strCAMPO02 & arrDESCMES(i)
        If i < lngMESFIN Then strCAMPO02 = strCAMPO02 & vbTab
    Next i
    
    strCAMPO03 = "Total" & vbTab & _
                 "Media"
    
    Open strCamRelNovo & strNomRel For Output As #1
    
    Print #1, strCAMPO01 & vbTab & _
              strCAMPO02 & vbTab & _
              strCAMPO03
              
    prgBAR.Min = 0
    prgBAR.Max = UBound(arrCLIENTES)
    For i = 1 To UBound(arrCLIENTES)
        prgBAR.value = i
        If arrCLIENTES(i).lngQTDECAPAC > 0 Then
            strCAMPO01 = ""
            For J = 1 To UBound(arrCLIENTES(i).arrCAPACIDADE)
                If arrCLIENTES(i).arrCAPACIDADE(J).lngTOTALCAPAC > 0 Then
                    
                    strCAMPO01 = arrCLIENTES(i).strRAZAOSOC & vbTab & _
                                 "'" & arrCLIENTES(i).strCNPJ & "'" & vbTab & _
                                 arrCLIENTES(i).arrCAPACIDADE(J).strDESCLINHA
                    
                    strCAMPO02 = ""
                    For K = lngMESINI To lngMESFIN
                        strCAMPO02 = strCAMPO02 & arrCLIENTES(i).arrCAPACIDADE(J).arrMESES(K).lngQTDE
                        If K < lngMESFIN Then strCAMPO02 = strCAMPO02 & vbTab
                    Next K
                    
                    strCAMPO03 = ""
                    strCAMPO03 = strCAMPO03 & arrCLIENTES(i).arrCAPACIDADE(J).lngTOTALCAPAC & vbTab & _
                                 strCAMPO03 & arrCLIENTES(i).arrCAPACIDADE(J).lngMEDIACAPAC
                    
                    Print #1, strCAMPO01 & vbTab & _
                              strCAMPO02 & vbTab & _
                              strCAMPO03
                
                End If
            Next J
        End If
    
        Frame5.Caption = "[ " & Format(((i / prgBAR.Max) * 100), "##00") & "% - Concluido, Geerando o Arquivo Aguarde !!! ]"
        Frame5.Refresh
    Next i
                 
    Close #1
    MsgBox "Arquivo Gerado com Sucesso !!!", vbOKOnly + vbExclamation, "Aviso"
    '' ------------------------------
    
    prgBAR.Visible = False
    Frame5.Caption = ""
    Frame5.Refresh
    
    
End Sub


Private Sub ImprimeCurvABCVALOR()

    Dim lngQTDECLIE     As Long
    Dim lngQTDEVEND     As Long
    Dim lngMESINI       As Long
    Dim lngMESFIN       As Long
    Dim lngTOTMESES     As Long
    Dim curTOTVALOR     As Currency
    
    Dim arrCLIENTES()   As ABCCLIVALOR
    Dim arrVENDEDOR()   As VENDEDORES
    Dim arrMESES()      As MESES
    Dim arrDESCMES()    As String

    Dim i               As Long
    Dim J               As Long
    Dim K               As Long
    
    Dim strCAMPO01      As String
    Dim strCAMPO02      As String
    Dim strCAMPO03      As String
    Dim strCAMPO04      As String
    
    Frame5.Caption = ""
    Frame5.Refresh
    
    lngMESINI = 0
    If Month(CDate(mskDTINI.Text)) > 1 Then lngMESINI = Month(CDate(mskDTINI.Text))
    lngMESFIN = Month(CDate(mskDTFIN.Text))
    lngTOTMESES = (lngMESFIN - lngMESINI)
    
    ReDim arrMESES(1 To 12) As MESES
    ReDim arrDESCMES(1 To 12) As String
    arrDESCMES(1) = "Janeiro"
    arrDESCMES(2) = "Fevereiro"
    arrDESCMES(3) = "Marco"
    arrDESCMES(4) = "Abril"
    arrDESCMES(5) = "Maio"
    arrDESCMES(6) = "Junho"
    arrDESCMES(7) = "Julho"
    arrDESCMES(8) = "Agosto"
    arrDESCMES(9) = "Setembro"
    arrDESCMES(10) = "Outubro"
    arrDESCMES(11) = "Novembro"
    arrDESCMES(12) = "Dezembro"
    
    lngQTDECLIE = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Count(SGI_CODIGO) as Qtde_Regs" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        sSql = sSql & "   And SGI_CODIGO = " & txtCIDCLIE.Text
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then lngQTDECLIE = BREC!Qtde_Regs
    BREC.Close
    
    If lngQTDECLIE = 0 Then Exit Sub

    prgBAR.Visible = True
    prgBAR.Min = 0
    prgBAR.Max = lngQTDECLIE
    
    ReDim Preserve arrCLIENTES(1 To lngQTDECLIE) As ABCCLIVALOR

    '' ----------------------------
    '' Pop Array Clientes
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_CPFCNPJ" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        sSql = sSql & "   And SGI_CODIGO = " & txtCIDCLIE.Text
    End If
    
    lngQTDECLIE = 0
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
        lngQTDECLIE = (lngQTDECLIE + 1)
        arrCLIENTES(lngQTDECLIE).lngCodClie = BREC!SGI_CODIGO
        arrCLIENTES(lngQTDECLIE).strRAZAOSOC = BREC!SGI_RAZAOSOC
        arrCLIENTES(lngQTDECLIE).strCNPJ = BREC!SGI_CPFCNPJ
        prgBAR.value = lngQTDECLIE
        
        Frame5.Caption = "[ " & Format(((lngQTDECLIE / prgBAR.Max) * 100), "##00") & "% - Concluido, Processando Aguarde !!! ]"
        Frame5.Refresh
        BREC.MoveNext
    Loop
    BREC.Close
    Frame5.Caption = ""
    Frame5.Refresh
    '' ----------------------------
    
    
    prgBAR.Min = 0
    prgBAR.Max = UBound(arrCLIENTES)
    
    For i = 1 To UBound(arrCLIENTES)
    
        prgBAR.value = i
        
        sSql = ""
        
        sSql = "Select Distinct" & vbCrLf
        sSql = sSql & "       VENDH.SGI_CODVEND" & vbCrLf
        sSql = sSql & "      ,VENDEDOR.SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADCLIENTE   CLIE" & vbCrLf
        sSql = sSql & "     , SGI_CADPEDVENDH" & strEMPRESA & "  VENDH" & vbCrLf
        sSql = sSql & "     , SGI_CADVENDEDOR  VENDEDOR" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       CLIE.SGI_FILIAL   = " & FILIAL & vbCrLf & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO   = " & arrCLIENTES(i).lngCodClie & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL   = VENDH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO   = VENDH.SGI_CODCLI" & vbCrLf
        sSql = sSql & "   And VENDH.SGI_FILIAL  = VENDEDOR.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And VENDH.SGI_CODVEND = VENDEDOR.SGI_CODIGO"
        
        lngQTDEVEND = 0
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
            Do While Not BREC.EOF()
                lngQTDEVEND = (lngQTDEVEND + 1)
                ReDim Preserve arrVENDEDOR(1 To lngQTDEVEND) As VENDEDORES
                arrVENDEDOR(lngQTDEVEND).lngCODVEND = BREC!SGI_CODVEND
                arrVENDEDOR(lngQTDEVEND).strDESCRICAO = BREC!SGI_DESCRICAO
                arrVENDEDOR(lngQTDEVEND).arrMESES = arrMESES
                BREC.MoveNext
            Loop
            arrCLIENTES(i).arrVENDEDOR = arrVENDEDOR
        End If
        BREC.Close
        arrCLIENTES(i).lngQTDVENDEDOR = lngQTDEVEND
        
        Frame5.Caption = "[ " & Format(((i / prgBAR.Max) * 100), "##00") & "% - Concluido, Processando Aguarde !!! ]"
        Frame5.Refresh
    
    Next i
    
    
    
    prgBAR.Min = 0
    prgBAR.Max = UBound(arrCLIENTES)

    For i = 1 To UBound(arrCLIENTES)
        
        prgBAR.value = i
        
        If arrCLIENTES(i).lngQTDVENDEDOR > 0 Then
            For J = 1 To UBound(arrCLIENTES(i).arrVENDEDOR)
                
                sSql = ""
                
                sSql = sSql & "Select" & vbCrLf
                sSql = sSql & "       Month(CONFH.SGI_DATACONF) As SGI_MES" & vbCrLf
                sSql = sSql & "      ,Sum((CONFI.SGI_VLUNIT * CONFI.SGI_QTDREAL))    As SGI_VLREAL" & vbCrLf
                
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_CADORDCONFH" & strEMPRESA & " CONFH" & vbCrLf
                sSql = sSql & "      ,SGI_CADORDCONFI" & strEMPRESA & " CONFI" & vbCrLf
                sSql = sSql & "      ,SGI_CADORDFATI" & strEMPRESA & "  FATI" & vbCrLf
                sSql = sSql & "      ,SGI_CADORDFATH" & strEMPRESA & "  FATH" & vbCrLf
                sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & " VENDI" & vbCrLf
                sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & " VENDH" & vbCrLf

                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       CONFH.SGI_FILIAL = " & FILIAL & vbCrLf
                sSql = sSql & "   And CONFH.SGI_DATACONF     Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "   And CONFH.SGI_FILIAL        = CONFI.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And CONFH.SGI_CODCONF       = CONFI.SGI_CODCONF" & vbCrLf
                sSql = sSql & "   And CONFH.SGI_FILIAL        = FATI.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And CONFH.SGI_CODORD        = FATI.SGI_CODORD" & vbCrLf
                sSql = sSql & "   And CONFI.SGI_IDPRODUTO     = FATI.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "   And FATI.SGI_FILIAL         = FATH.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And FATI.SGI_CODORD         = FATH.SGI_CODORD" & vbCrLf
                sSql = sSql & "   And FATH.SGI_FILIAL         = VENDI.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And FATH.SGI_CODPED         = VENDI.SGI_CODIGO" & vbCrLf
                sSql = sSql & "   And FATI.SGI_IDPRODUTO      = VENDI.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "   And VENDI.SGI_FILIAL        = VENDH.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And VENDI.SGI_CODIGO        = VENDH.SGI_CODIGO" & vbCrLf
                sSql = sSql & "   And VENDH.SGI_CODCLI        = " & arrCLIENTES(i).lngCodClie & vbCrLf
                sSql = sSql & "   And VENDH.SGI_CODVEND       = " & arrCLIENTES(i).arrVENDEDOR(J).lngCODVEND & vbCrLf
                sSql = sSql & "Group By Month(CONFH.SGI_DATACONF)" & vbCrLf
                sSql = sSql & "Order BY SGI_MES"
                   
                BREC.Open sSql, adoBanco_Dados, adOpenDynamic
                curTOTVALOR = 0
                Do While Not BREC.EOF()
                    arrCLIENTES(i).arrVENDEDOR(J).arrMESES(BREC!SGI_MES).curVALOR = BREC!SGI_VLREAL
                    curTOTVALOR = (curTOTVALOR + BREC!SGI_VLREAL)
                    BREC.MoveNext
                Loop
                BREC.Close
                arrCLIENTES(i).arrVENDEDOR(J).curTOTALVALOR = curTOTVALOR
                arrCLIENTES(i).arrVENDEDOR(J).curMEDIAVALOR = (curTOTVALOR / lngTOTMESES)
            Next J
        End If
        
        Frame5.Caption = "[ " & Format(((i / prgBAR.Max) * 100), "##00") & "% - Concluido, Processando Aguarde !!! ]"
        Frame5.Refresh
    
    Next i
    
    
    '' ====================================
    '' Montando o Arquivo TXT
    strNomRel = "CURVAABCVALOR" & strEMPRESA & ".txt"
    
    strCAMPO01 = "RAZAO SOCIAL" & vbTab & _
                 "CNPJ" & vbTab & _
                 "VENDEDOR"

    If lngMESINI = 0 Then lngMESINI = 1
    strCAMPO02 = ""
    
    For i = lngMESINI To lngMESFIN
        strCAMPO02 = strCAMPO02 & arrDESCMES(i)
        If i < lngMESFIN Then strCAMPO02 = strCAMPO02 & vbTab
    Next i
    
    strCAMPO03 = "Total" & vbTab & _
                 "Media"
    
    Open strCamRelNovo & strNomRel For Output As #1
    
    Print #1, strCAMPO01 & vbTab & _
              strCAMPO02 & vbTab & _
              strCAMPO03
              
    prgBAR.Min = 0
    prgBAR.Max = UBound(arrCLIENTES)
    For i = 1 To UBound(arrCLIENTES)
        prgBAR.value = i
        If arrCLIENTES(i).lngQTDVENDEDOR > 0 Then
            strCAMPO01 = ""
            For J = 1 To UBound(arrCLIENTES(i).arrVENDEDOR)
                If arrCLIENTES(i).arrVENDEDOR(J).curTOTALVALOR > 0 Then
                    
                    strCAMPO01 = arrCLIENTES(i).strRAZAOSOC & vbTab & _
                                 "'" & arrCLIENTES(i).strCNPJ & "'" & vbTab & _
                                 arrCLIENTES(i).arrVENDEDOR(J).strDESCRICAO
                    
                    strCAMPO02 = ""
                    For K = lngMESINI To lngMESFIN
                        strCAMPO02 = strCAMPO02 & Format(arrCLIENTES(i).arrVENDEDOR(J).arrMESES(K).curVALOR, "#,##0.00")
                        If K < lngMESFIN Then strCAMPO02 = strCAMPO02 & vbTab
                    Next K
                    
                    strCAMPO03 = ""
                    strCAMPO03 = strCAMPO03 & Format(arrCLIENTES(i).arrVENDEDOR(J).curTOTALVALOR, "#,##0.00") & vbTab & _
                                 strCAMPO03 & Format(arrCLIENTES(i).arrVENDEDOR(J).curMEDIAVALOR, "#,##0.00")
                    
                    Print #1, strCAMPO01 & vbTab & _
                              strCAMPO02 & vbTab & _
                              strCAMPO03
                
                End If
            Next J
        End If
    
        Frame5.Caption = "[ " & Format(((i / prgBAR.Max) * 100), "##00") & "% - Concluido, Geerando o Arquivo Aguarde !!! ]"
        Frame5.Refresh
    Next i
                 
    Close #1
    MsgBox "Arquivo Gerado com Sucesso !!!", vbOKOnly + vbExclamation, "Aviso"
    '' ------------------------------
    
    prgBAR.Min = 0
    prgBAR.Visible = False
    Frame5.Caption = ""
    Frame5.Refresh
    
    
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
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label)

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

End Sub

