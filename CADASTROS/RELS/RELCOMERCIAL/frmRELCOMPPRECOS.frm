VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRELCOMPPRECOS 
   Caption         =   "Relat�rio de Estudo de Pre�os"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   12825
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   15
      Top             =   1560
      Width           =   12735
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   1320
         Picture         =   "frmRELCOMPPRECOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
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
         TabIndex        =   10
         Top             =   240
         Width           =   7335
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
      Left            =   6480
      TabIndex        =   11
      Top             =   960
      Width           =   6255
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
         Left            =   3840
         TabIndex        =   14
         Top             =   240
         Width           =   2175
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
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   1335
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
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
      Top             =   2280
      Width           =   12735
      Begin ComctlLib.ProgressBar prbDados 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   6375
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
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
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   1320
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
         TabIndex        =   8
         Top             =   240
         Width           =   1095
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
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12735
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
         Picture         =   "frmRELCOMPPRECOS.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmRELCOMPPRECOS.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELCOMPPRECOS"
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
Dim objRELCOMPPRECOS    As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String
Dim strNomRel           As String
Dim strEMPRESADESC      As String
Dim strEMPRESA          As String

Dim arrCLIENTESNOVALATA()   As CPCLIENTES
Dim arrCLIENTESSTEEL()      As CPCLIENTES
Dim arrDADOSANOS()          As CPANOS


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
    If Len(Trim(lblDescCliente.Caption)) = 0 Then
       txtCIDCLIE.Text = ""
       Call LimpaCamposLabel
    End If
        
    txtCIDCLIE.SetFocus

    Exit Sub
    
Err_Command1_Click:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Fun��o : Command1_Click()", Me.Name, "Command1_Click()", strCAMARQERRO)

End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELCOMPPRECOS = CreateObject("RELCOMERCIAL.clsRELCOMPPRECOS")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "N�o foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
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

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, Me
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, Me
End Sub

Private Sub GeraArquivoExcel()

On Error GoTo Err_Exporta

    Dim boolTemDados            As Boolean
    Dim boolTemDadosNOVA        As Boolean
    Dim boolTemDadosSTEEL       As Boolean
    
    Dim lngQTDREGS              As Long
    Dim lngQTDREGSCAPAC         As Long
    Dim lngTOTREGS              As Long
    
    Dim arrDADOSCLIENOVA()      As CPCLIENTES
    Dim arrDADOSCLIESTEEL()     As CPCLIENTES
    
    Dim arrANOS()               As CPANOS
    Dim arrCAPACIDADE()         As CPCAPACIDADE
    
    Dim strEMPRESA              As String
    Dim strEMPARQ               As String
    Dim strNomRel               As String
    Dim i                       As Long
    Dim lngQTDEANOS             As Long
    
    strEMPRESA = ""
    strEMPARQ = "NOVALATA"
    If optEmpresa(1).value = True Then
       strEMPRESA = "_STEEL"
       strEMPARQ = "STEEL"
    End If
    
    boolTemDados = False
    boolTemDadosNOVA = False
    boolTemDadosSTEEL = False
    
    prbDados.Min = 0
    
    '' -----------------------
    '' Pega Anos
    
    If Year(mskDTINI.Text) = Year(mskDTFIN.Text) Then
        ReDim Preserve arrANOS(1 To 1) As CPANOS
        arrANOS(1).lngANO = Year(mskDTINI.Text)
    Else
        lngQTDEANOS = Year(mskDTFIN.Text) - Year(mskDTINI.Text)
        For i = 0 To lngQTDEANOS
            ReDim Preserve arrANOS(1 To (i + 1)) As CPANOS
            arrANOS(i + 1).lngANO = (Year(mskDTINI.Text) + i)
        Next i
    End If
    arrDADOSANOS = arrANOS
    
    '' -----------------------
    lngQTDREGS = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       CABEC.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & " CABEC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "        CABEC.SGI_FILIAL    = " & FILIAL & vbCrLf
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        sSql = sSql & "    And CABEC.SGI_CODCLI    = " & Trim(txtCIDCLIE.Text) & vbCrLf
    End If
    sSql = sSql & "    And CABEC.SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "    And CLIE.SGI_FILIAL     = CABEC.SGI_FILIAL" & vbCrLf
    sSql = sSql & "    And CLIE.SGI_CODIGO     = CABEC.SGI_CODCLI" & vbCrLf
    sSql = sSql & "Group By CABEC.SGI_CODCLI,CLIE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "Order By CABEC.SGI_CODCLI"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        With BREC
            boolTemDados = True
            lngTOTREGS = 0
            Do While Not .EOF()
                lngTOTREGS = (lngTOTREGS + 1)
                .MoveNext
            Loop
            
            prbDados.Max = lngTOTREGS
            lngQTDREGS = 0
            .MoveFirst
            
            Frame3.Caption = "[ AGUARDE. Processando dados " & strEMPARQ & "... ]"
            Frame3.Refresh
            
            Do While Not .EOF()
                lngQTDREGS = (lngQTDREGS + 1)
                
                If optEmpresa(0).value = True Or optEmpresa(2).value = True Then
                   ReDim Preserve arrDADOSCLIENOVA(1 To lngQTDREGS) As CPCLIENTES
                   boolTemDadosNOVA = True
                ElseIf optEmpresa(1).value = True Then
                   ReDim Preserve arrDADOSCLIESTEEL(1 To lngQTDREGS) As CPCLIENTES
                   boolTemDadosSTEEL = True
                End If
                
                prbDados.value = lngQTDREGS
                
                If optEmpresa(0).value = True Or optEmpresa(2).value = True Then
                    arrDADOSCLIENOVA(lngQTDREGS).lngCodClie = !SGI_CODCLI
                    arrDADOSCLIENOVA(lngQTDREGS).strRAZAOSOC = !SGI_RAZAOSOC
                ElseIf optEmpresa(1).value = True Then
                    arrDADOSCLIESTEEL(lngQTDREGS).lngCodClie = !SGI_CODCLI
                    arrDADOSCLIESTEEL(lngQTDREGS).strRAZAOSOC = !SGI_RAZAOSOC
                End If
                
                '' Pegando a Capacidade
                sSql = ""
                
                sSql = "Select" & vbCrLf
                sSql = sSql & "       PROD.SGI_CODLINPROD" & vbCrLf
                sSql = sSql & "      ,LINP.SGI_DESCRI" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & " CABEC" & vbCrLf
                sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & " ITENS" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
                sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINP" & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       CABEC.SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & "   And CABEC.SGI_CODCLI    = " & !SGI_CODCLI & vbCrLf
                sSql = sSql & "   And CABEC.SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "   And ITENS.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And ITENS.SGI_CODIGO    = CABEC.SGI_CODIGO" & vbCrLf
                sSql = sSql & "   And PROD.SGI_FILIAL     = ITENS.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ITENS.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "   And LINP.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And LINP.SGI_CODLIN     = PROD.SGI_CODLINPROD" & vbCrLf
                sSql = sSql & "Group By PROD.SGI_CODLINPROD,LINP.SGI_DESCRI" & vbCrLf
                sSql = sSql & "Order By PROD.SGI_CODLINPROD"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                lngQTDREGSCAPAC = 0
                Do While Not BREC2.EOF()
                    lngQTDREGSCAPAC = (lngQTDREGSCAPAC + 1)
                    ReDim Preserve arrCAPACIDADE(1 To lngQTDREGSCAPAC) As CPCAPACIDADE
                    
                    arrCAPACIDADE(lngQTDREGSCAPAC).lngCodLinProd = BREC2!SGI_CODLINPROD
                    arrCAPACIDADE(lngQTDREGSCAPAC).strDESCLIN = BREC2!SGI_DESCRI
                    
                    For i = 1 To UBound(arrANOS)
                    
                        sSql = ""
                    
                        sSql = "Select top 1" & vbCrLf
                        sSql = sSql & "       ITENS.SGI_VLUNIT" & vbCrLf
                        sSql = sSql & "  From" & vbCrLf
                        sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & "     CABEC" & vbCrLf
                        sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & "     ITENS" & vbCrLf
                        sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
                        sSql = sSql & " Where" & vbCrLf
                        sSql = sSql & "       CABEC.SGI_FILIAL = " & FILIAL & vbCrLf
                        sSql = sSql & "   And CABEC.SGI_CODCLI = " & !SGI_CODCLI & vbCrLf
                        sSql = sSql & "   And CABEC.SGI_DATAPED Between '" & Format(CDate("01/01/" & arrANOS(i).lngANO), "MM/DD/YYYY") & "' And '" & Format(CDate("31/12/" & arrANOS(i).lngANO), "MM/DD/YYYY") & "'" & vbCrLf
                        sSql = sSql & "   And ITENS.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
                        sSql = sSql & "   And ITENS.SGI_CODIGO    = CABEC.SGI_CODIGO" & vbCrLf
                        sSql = sSql & "   And PROD.SGI_FILIAL     = ITENS.SGI_FILIAL" & vbCrLf
                        sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ITENS.SGI_IDPRODUTO" & vbCrLf
                        sSql = sSql & "   And PROD.SGI_CODLINPROD = " & BREC2!SGI_CODLINPROD & vbCrLf
                        sSql = sSql & "Order By CABEC.SGI_DATAPED Desc"
                        
                        BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                        If Not BREC3.EOF() Then
                           arrANOS(i).strPRECO = Format(BREC3!SGI_VLUNIT, "#,##0.00")
                        Else
                           arrANOS(i).strPRECO = ""
                        End If
                        BREC3.Close
                    
                    Next i
                    
                    arrCAPACIDADE(lngQTDREGSCAPAC).arrANOS = arrANOS
                    
                    BREC2.MoveNext
                Loop
                BREC2.Close
                
                If optEmpresa(0).value = True Or optEmpresa(2).value = True Then
                    arrDADOSCLIENOVA(lngQTDREGS).arrCAPACIDADE = arrCAPACIDADE
                    arrDADOSCLIENOVA(lngQTDREGS).lngQTDECAPACIDADE = lngQTDREGSCAPAC
                ElseIf optEmpresa(1).value = True Then
                    arrDADOSCLIESTEEL(lngQTDREGS).arrCAPACIDADE = arrCAPACIDADE
                    arrDADOSCLIESTEEL(lngQTDREGS).lngQTDECAPACIDADE = lngQTDREGSCAPAC
                End If
                
                .MoveNext
                
            Loop
        End With
        
        arrCLIENTESNOVALATA = arrDADOSCLIENOVA
        arrCLIENTESSTEEL = arrDADOSCLIESTEEL
    
    End If
    BREC.Close
    
    If boolTemDados = False Then
        MsgBox "ATEN��O" & vbCrLf & _
               "N�o existe dados para carregar !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    
    '' Pega a Empresa STEEL
    If optEmpresa(2).value = True Then
       strEMPRESA = "_STEEL"
       strEMPARQ = "STEEL"
        
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       CABEC.SGI_CODCLI" & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & " CABEC" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "        CABEC.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "    And CABEC.SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
        If Len(Trim(txtCIDCLIE.Text)) > 0 Then
            sSql = sSql & "    And CABEC.SGI_CODCLI    = " & Trim(txtCIDCLIE.Text) & vbCrLf
        End If
        sSql = sSql & "    And CLIE.SGI_FILIAL     = CABEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "    And CLIE.SGI_CODIGO     = CABEC.SGI_CODCLI" & vbCrLf
        sSql = sSql & "Group By CABEC.SGI_CODCLI,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "Order By CABEC.SGI_CODCLI"
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
            With BREC
                boolTemDados = True
                lngTOTREGS = 0
                Do While Not .EOF()
                    lngTOTREGS = (lngTOTREGS + 1)
                    .MoveNext
                Loop
                
                prbDados.Max = lngTOTREGS
                lngQTDREGS = 0
                .MoveFirst
                
                Frame3.Caption = "[ AGUARDE. Processando dados " & strEMPARQ & "... ]"
                Frame3.Refresh
                boolTemDadosSTEEL = True
                
                
                Do While Not .EOF()
                    lngQTDREGS = (lngQTDREGS + 1)
                    
                    ReDim Preserve arrDADOSCLIESTEEL(1 To lngQTDREGS) As CPCLIENTES
                    
                    prbDados.value = lngQTDREGS
                    
                    arrDADOSCLIESTEEL(lngQTDREGS).lngCodClie = !SGI_CODCLI
                    arrDADOSCLIESTEEL(lngQTDREGS).strRAZAOSOC = !SGI_RAZAOSOC
                    
                    '' Pegando a Capacidade
                    sSql = ""
                    
                    sSql = "Select" & vbCrLf
                    sSql = sSql & "       PROD.SGI_CODLINPROD" & vbCrLf
                    sSql = sSql & "      ,LINP.SGI_DESCRI" & vbCrLf
                    sSql = sSql & "  From" & vbCrLf
                    sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & " CABEC" & vbCrLf
                    sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & " ITENS" & vbCrLf
                    sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
                    sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINP" & vbCrLf
                    sSql = sSql & " Where" & vbCrLf
                    sSql = sSql & "       CABEC.SGI_FILIAL    = " & FILIAL & vbCrLf
                    sSql = sSql & "   And CABEC.SGI_CODCLI    = " & !SGI_CODCLI & vbCrLf
                    sSql = sSql & "   And CABEC.SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'"
                    sSql = sSql & "   And ITENS.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And ITENS.SGI_CODIGO    = CABEC.SGI_CODIGO" & vbCrLf
                    sSql = sSql & "   And PROD.SGI_FILIAL     = ITENS.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ITENS.SGI_IDPRODUTO" & vbCrLf
                    sSql = sSql & "   And LINP.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And LINP.SGI_CODLIN     = PROD.SGI_CODLINPROD" & vbCrLf
                    sSql = sSql & "Group By PROD.SGI_CODLINPROD,LINP.SGI_DESCRI" & vbCrLf
                    sSql = sSql & "Order By PROD.SGI_CODLINPROD"
                    
                    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    lngQTDREGSCAPAC = 0
                    Do While Not BREC2.EOF()
                        lngQTDREGSCAPAC = (lngQTDREGSCAPAC + 1)
                        ReDim Preserve arrCAPACIDADE(1 To lngQTDREGSCAPAC) As CPCAPACIDADE
                        
                        arrCAPACIDADE(lngQTDREGSCAPAC).lngCodLinProd = BREC2!SGI_CODLINPROD
                        arrCAPACIDADE(lngQTDREGSCAPAC).strDESCLIN = BREC2!SGI_DESCRI
                        
                        
                        For i = 1 To UBound(arrANOS)
                        
                            sSql = ""
                        
                            sSql = "Select top 1" & vbCrLf
                            sSql = sSql & "       ITENS.SGI_VLUNIT" & vbCrLf
                            sSql = sSql & "  From" & vbCrLf
                            sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & "     CABEC" & vbCrLf
                            sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & "     ITENS" & vbCrLf
                            sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
                            sSql = sSql & " Where" & vbCrLf
                            sSql = sSql & "       CABEC.SGI_FILIAL = " & FILIAL & vbCrLf
                            sSql = sSql & "   And CABEC.SGI_CODCLI = " & !SGI_CODCLI & vbCrLf
                            sSql = sSql & "   And CABEC.SGI_DATAPED Between '" & Format(CDate("01/01/" & arrANOS(i).lngANO), "MM/DD/YYYY") & "' And '" & Format(CDate("31/12/" & arrANOS(i).lngANO), "MM/DD/YYYY") & "'" & vbCrLf
                            sSql = sSql & "   And ITENS.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
                            sSql = sSql & "   And ITENS.SGI_CODIGO    = CABEC.SGI_CODIGO" & vbCrLf
                            sSql = sSql & "   And PROD.SGI_FILIAL     = ITENS.SGI_FILIAL" & vbCrLf
                            sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ITENS.SGI_IDPRODUTO" & vbCrLf
                            sSql = sSql & "   And PROD.SGI_CODLINPROD = " & BREC2!SGI_CODLINPROD & vbCrLf
                            sSql = sSql & "Order By CABEC.SGI_DATAPED Desc"
                            
                            BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                            If Not BREC3.EOF() Then
                               arrANOS(i).strPRECO = Format(BREC3!SGI_VLUNIT, "#,##0.00")
                            Else
                               arrANOS(i).strPRECO = ""
                            End If
                            BREC3.Close
                        
                        Next i
                        
                        arrCAPACIDADE(lngQTDREGSCAPAC).arrANOS = arrANOS
                        
                        BREC2.MoveNext
                    Loop
                    BREC2.Close
                    
                    arrDADOSCLIESTEEL(lngQTDREGS).arrCAPACIDADE = arrCAPACIDADE
                    arrDADOSCLIESTEEL(lngQTDREGS).lngQTDECAPACIDADE = lngQTDREGSCAPAC
                    
                    .MoveNext
                Loop
            End With
            arrCLIENTESSTEEL = arrDADOSCLIESTEEL

        End If
        BREC.Close
    
    End If
    
    '' Gerando o Arquivo XLS
    If optEmpresa(0).value = True Then strNomRel = "RELCOMPPRECOS_NOVATALA.xls"
    If optEmpresa(1).value = True Then strNomRel = "RELCOMPPRECOS_STEEL.xls"
    If optEmpresa(2).value = True Then strNomRel = "RELCOMPPRECOS_NE.xls"
    
    '' Gerando o Arquivo Excel
    Call GeraArgExcel(strNomRel, boolTemDadosNOVA, boolTemDadosSTEEL)
    
    MsgBox "Dados Carregados com sucesso !", vbOKOnly + vbInformation, "Aviso"
    
    Exit Sub
    
Err_Exporta:

    MsgBox "Erro      : " & Err.Number & vbCrLf & _
           "Descri��o : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"

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
    
    
    lngQTDANOS = UBound(arrDADOSANOS)
    
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
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Raz�o Social", 12
        
        lngCol = (lngCol + 1)
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 30
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Lata", 12
        
        '' ------------------
        '' Dados do Anos
        For i = 1 To UBound(arrDADOSANOS)
            lngCol = (lngCol + 1)
            .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 18
           .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, Str(arrDADOSANOS(i).lngANO), 1
        Next i
        '' ------------------
        
        lngCol = (lngCol + 1)
        .SetColumnWidth CByte(Str(lngCol)), CByte(Str(lngCol)), 20
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, lngCol, "Empresa", 12
    
        '' Dados da Novalata
        lngLINHA = 1
        If boolTemDadosNOVA = True Then
             Frame3.Caption = "[ Aguarde.... Gerando o Arguivo EXCEL com os dados da NOVALATA ! ]"
             Frame3.Refresh
             
             prbDados.Min = 0
             prbDados.Max = UBound(arrCLIENTESNOVALATA)
             
             For lngREGS = 1 To UBound(arrCLIENTESNOVALATA)
                lngCol = 1
                lngLINHA = (lngLINHA + 1)
                
                prbDados.value = lngREGS
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).strRAZAOSOC, 12         '' RAZ�O SOCIAL
                
                '' Capacidade
                If arrCLIENTESNOVALATA(lngREGS).lngQTDECAPACIDADE > 0 Then
                   lngCol = (lngCol + 1)
                   .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).arrCAPACIDADE(1).strDESCLIN, 12         '' Descri��o da Linha
                   .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, ((lngCol + lngQTDANOS) + 1), "NOVALATA", 12                                                                '' Empresa
                  
                   '' Ano
                   lngColAno = lngCol
                   For i = 1 To UBound(arrCLIENTESNOVALATA(lngREGS).arrCAPACIDADE(1).arrANOS)
                       lngColAno = (lngColAno + 1)
                       .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, lngColAno, arrCLIENTESNOVALATA(lngREGS).arrCAPACIDADE(1).arrANOS(i).strPRECO, 2          '' Pegou o Anos
                   Next i
                  
                  
                   If arrCLIENTESNOVALATA(lngREGS).lngQTDECAPACIDADE > 1 Then
                      lngColAno = lngCol
                      For lngREGSCAPAC = 2 To arrCLIENTESNOVALATA(lngREGS).lngQTDECAPACIDADE
                            lngLINHA = (lngLINHA + 1)
                            .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESNOVALATA(lngREGS).arrCAPACIDADE(lngREGSCAPAC).strDESCLIN, 12         '' Descri��o da Linha
                            .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, ((lngCol + lngQTDANOS) + 1), "NOVALATA", 12                                             '' Empresa
                      
                            '' Ano
                            lngColAno = lngCol
                            For i = 1 To UBound(arrCLIENTESNOVALATA(lngREGS).arrCAPACIDADE(1).arrANOS)
                                lngColAno = (lngColAno + 1)
                                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, lngColAno, arrCLIENTESNOVALATA(lngREGS).arrCAPACIDADE(lngREGSCAPAC).arrANOS(i).strPRECO, 2         '' Pegou o Anos
                            Next i
                      Next lngREGSCAPAC
                   End If
                
                
                End If
                
            
             Next lngREGS
        End If
        
        '' Dados da STEEL
        If boolTemDadosSTEEL = True Then
             Frame3.Caption = "[ Aguarde.... Gerando o Arguivo EXCEL com os dados da STEEL ! ]"
             Frame3.Refresh
             
             prbDados.Min = 0
             prbDados.Max = UBound(arrCLIENTESSTEEL)
             
             For lngREGS = 1 To UBound(arrCLIENTESSTEEL)
                lngCol = 1
                lngLINHA = (lngLINHA + 1)
                
                prbDados.value = lngREGS
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).strRAZAOSOC, 12         '' RAZ�O SOCIAL
                
                '' Capacidade
                If arrCLIENTESSTEEL(lngREGS).lngQTDECAPACIDADE > 0 Then
                   lngCol = (lngCol + 1)
                  .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).arrCAPACIDADE(1).strDESCLIN, 12         '' DESCRI��O DA LINHA
                  .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, ((lngCol + lngQTDANOS) + 1), "STEEL", 12                                  '' Empresa
                  
                   '' Ano
                   lngColAno = lngCol
                   For i = 1 To UBound(arrCLIENTESSTEEL(lngREGS).arrCAPACIDADE(1).arrANOS)
                       lngColAno = (lngColAno + 1)
                       .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, lngColAno, arrCLIENTESSTEEL(lngREGS).arrCAPACIDADE(1).arrANOS(i).strPRECO, 2         '' Pegou o Anos
                   Next i
                  
                  If arrCLIENTESSTEEL(lngREGS).lngQTDECAPACIDADE > 1 Then
                     For lngREGSCAPAC = 2 To arrCLIENTESSTEEL(lngREGS).lngQTDECAPACIDADE
                         lngLINHA = (lngLINHA + 1)
                     
                         .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, lngCol, arrCLIENTESSTEEL(lngREGS).arrCAPACIDADE(lngREGSCAPAC).strDESCLIN, 12          '' RAZ�O SOCIAL
                         .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, ((lngCol + lngQTDANOS) + 1), "STEEL", 12                                              '' Empresa
                  
                         '' Ano
                         lngColAno = lngCol
                         For i = 1 To UBound(arrCLIENTESSTEEL(lngREGS).arrCAPACIDADE(1).arrANOS)
                             lngColAno = (lngColAno + 1)
                             .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, lngColAno, arrCLIENTESSTEEL(lngREGS).arrCAPACIDADE(lngREGSCAPAC).arrANOS(i).strPRECO, 2         '' Pegou o Anos
                         Next i
                     
                     Next lngREGSCAPAC
                  End If
                
                
                End If
            
             Next lngREGS
        End If
        
        .ProtectSpreadsheet = False 'False | True
        .CloseFile
    
    End With

    Frame3.Caption = ""
    Frame3.Refresh
    
    prbDados.value = 0

    Exit Sub

err_Excel:

    MsgBox "ATEN��O" & vbCrLf & _
           "Erro Numero       : " & Err.Number & vbCrLf & _
           "Descri��o do Erro : " & Err.Description, vbOKOnly + vbCritical, "Aviso"

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
       MsgBox "Somente � permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
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
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Fun��o : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub


Private Sub LimpaCamposLabel()
    lblDescCliente.Caption = ""
End Sub
