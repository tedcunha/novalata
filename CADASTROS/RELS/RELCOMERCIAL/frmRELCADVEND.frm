VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRELCADVEND 
   Caption         =   "Relatório de Vendedores"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   12375
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboESTNORM 
      Height          =   315
      Left            =   7440
      TabIndex        =   12
      Text            =   "cboESTNORM"
      Top             =   1200
      Width           =   750
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Vendedores ]"
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
      Top             =   1560
      Width           =   12375
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1320
         Picture         =   "frmRELCADVEND.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCODVEND 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "txtCODVEND"
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label lblDescVendedor 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendedor"
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Progresso....]"
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
      Top             =   2160
      Width           =   12375
      Begin ComctlLib.ProgressBar prgProg 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   6015
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
         TabIndex        =   6
         Top             =   240
         Width           =   2535
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
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   1215
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
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
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
         Picture         =   "frmRELCADVEND.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmRELCADVEND.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELCADVEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho             As String
Public Linha                As Variant
Public FILIAL               As Integer
Public strAcesso            As String

Dim arrDADOSNOVA()          As String
Dim arrDADOSSTEEL()         As String

Dim objBLBFunc              As Object
Dim objRELCADVEND           As Object
Dim objPESQPADRAO           As Object
Dim objREL                  As Object

Private Sub cmdImpressao_Click()
    Call GeraExcel
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

On Error GoTo Err_Command2_Click


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
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVEND.Text = varRETORNO
    
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedor)
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then txtCODVEND.Text = ""
    
    If txtCODVEND.Enabled = True Then txtCODVEND.SetFocus

    Exit Sub
    
Err_Command2_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : Command2_Click()", Me.Name, "Command2_Click()", strCAMARQERRO)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Public Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELCADVEND = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELCADVEND = CreateObject("RELCOMERCIAL.clsRELCADVEND")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    
    objRELCADVEND.FILIAL = FILIAL
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    Me.Caption = Me.Caption & " / " & Trim(Me.Name)
    
    optEmpresa(0).value = True
    Frame3.Visible = False
    prgProg.Min = 0

    Call LimpaCampos
    
    cboESTNORM.Visible = False
    objBLBFunc.Preenche_Estado cboESTNORM
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub GeraExcel()

    Dim lngQTDEMESEMAIS         As Long
    Dim lngQTDREGS              As Long
    Dim strNOMARQ               As String
    Dim strNOMEMPTAB            As String
    Dim boolTemDadosNovalata    As Boolean
    Dim boolTemDadosSTEEL       As Boolean
    
    Dim strDADOSFAT             As String
    Dim strDADOSPEDIDOS         As String
    
    Dim arrDADOSFAT()           As String
    Dim arrDADOSPED()           As String
    
    Dim i                       As Long
    
    Me.MousePointer = 11
    
    Frame3.Visible = True
    
    strNOMEMPTAB = ""
    strNOMARQ = "_NOVALATA"
    If optEmpresa(1).value = True Then
       strNOMEMPTAB = "_STEEL"
       strNOMARQ = "_STEEL"
    ElseIf optEmpresa(2).value = True Then
       strNOMARQ = "_NOVSTE"
    End If
    
    boolTemDadosNovalata = False
    boolTemDadosSTEEL = False
    
    prgProg.Min = 0
    
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       CADVE.SGI_DESCRICAO As SGI_DESCVEND" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_CODIGO     As SGI_CODCLIE" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_ESTNORM" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,CODL.SGI_DESCRI     As SGI_DECLINHA" & vbCrLf
    sSql = sSql & "      ,PEDVI.SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO  As SGI_NOMPROD" & vbCrLf
    sSql = sSql & "      ,PEDVI.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,PEDVH.SGI_CODVEND" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH" & strNOMEMPTAB & "     PEDVH" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI" & strNOMEMPTAB & "     PEDVI" & vbCrLf
    sSql = sSql & "      ,SGI_CADVENDEDOR     CADVE" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO CODL" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PEDVH.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDVI.SGI_FILIAL   = PEDVH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PEDVI.SGI_CODIGO   = PEDVH.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL    = PEDVI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = PEDVI.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And CODL.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CODL.SGI_CODLIN    = PROD.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "   And CADVE.SGI_FILIAL   = PEDVH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   ANd CADVE.SGI_CODIGO   = PEDVH.SGI_CODVEND" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL    = PEDVH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO    = PEDVH.SGI_CODCLI" & vbCrLf
    
    If Len(Trim(txtCODVEND.Text)) > 0 Then
        sSql = sSql & "   And PEDVH.SGI_CODVEND    = " & Trim(txtCODVEND.Text) & vbCrLf
    End If
    
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "       CADVE.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,CODL.SGI_DESCRI" & vbCrLf
    sSql = sSql & "      ,PEDVI.SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,PEDVI.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,PEDVH.SGI_CODVEND" & vbCrLf


    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC2.EOF() Then
        
        
        lngQTDREGS = 0
        Do While Not BREC2.EOF()
            lngQTDREGS = (lngQTDREGS + 1)
            BREC2.MoveNext
        Loop
    
        prgProg.Max = lngQTDREGS
        If optEmpresa(0).value = True Or optEmpresa(2).value = True Then ReDim arrDADOSNOVA(1 To lngQTDREGS, 1 To 13) As String
        If optEmpresa(1).value = True Then ReDim arrDADOSSTEEL(1 To lngQTDREGS, 1 To 13) As String
    
        BREC2.MoveFirst
        lngQTDREGS = 0
        
        If optEmpresa(0).value = True Or optEmpresa(2).value = True Then Frame3.Caption = "[ Aguarde gerando dados NOVALATA.... ]"
        If optEmpresa(1).value = True Then Frame3.Caption = "[ Aguarde gerando dados STEEL.... ]"
        Frame3.Refresh
        
        Do While Not BREC2.EOF()
            lngQTDREGS = (lngQTDREGS + 1)
            prgProg.value = lngQTDREGS
            DoEvents
            
            If optEmpresa(0).value = True Or optEmpresa(2).value = True Then
                arrDADOSNOVA(lngQTDREGS, 1) = Trim(BREC2!SGI_DESCVEND)
                arrDADOSNOVA(lngQTDREGS, 2) = Trim(BREC2!SGI_RAZAOSOC)
                arrDADOSNOVA(lngQTDREGS, 3) = Trim(BREC2!SGI_DECLINHA)
                arrDADOSNOVA(lngQTDREGS, 4) = Trim(BREC2!SGI_CODPROD)
                arrDADOSNOVA(lngQTDREGS, 5) = Trim(BREC2!SGI_NOMPROD)
                arrDADOSNOVA(lngQTDREGS, 6) = "NOVALATA"
                
                arrDADOSNOVA(lngQTDREGS, 7) = ""
                arrDADOSNOVA(lngQTDREGS, 8) = ""
                strDADOSFAT = PegaUltFat(BREC2!SGI_CODCLIE, BREC2!SGI_IDPRODUTO, strNOMEMPTAB, BREC2!SGI_CODVEND)
                strDADOSPEDIDOS = PegaUltMeses(BREC2!SGI_CODCLIE, BREC2!SGI_IDPRODUTO, strNOMEMPTAB, BREC2!SGI_CODVEND, Month(Now))
                
                If Len(Trim(strDADOSFAT)) > 0 Then
                   arrDADOSFAT = Split(strDADOSFAT, "|")
                   arrDADOSNOVA(lngQTDREGS, 7) = arrDADOSFAT(0)
                   arrDADOSNOVA(lngQTDREGS, 8) = arrDADOSFAT(1)
                End If
                
                arrDADOSPED = Split(strDADOSPEDIDOS, "|")
                
                arrDADOSNOVA(lngQTDREGS, 9) = arrDADOSPED(0)
                arrDADOSNOVA(lngQTDREGS, 10) = arrDADOSPED(1)
                arrDADOSNOVA(lngQTDREGS, 11) = arrDADOSPED(2)
                
                For i = 0 To (cboESTNORM.ListCount - 1)
                    If cboESTNORM.ItemData(i) = BREC2!SGI_ESTNORM Then
                        cboESTNORM.ListIndex = i
                        Exit For
                    End If
                Next i
                arrDADOSNOVA(lngQTDREGS, 12) = cboESTNORM.Text
                
                arrDADOSNOVA(lngQTDREGS, 13) = Trim(BREC2!SGI_CIDNORM)
                
                boolTemDadosNovalata = True
            ElseIf optEmpresa(1).value = True Then
                arrDADOSSTEEL(lngQTDREGS, 1) = Trim(BREC2!SGI_DESCVEND)
                arrDADOSSTEEL(lngQTDREGS, 2) = Trim(BREC2!SGI_RAZAOSOC)
                arrDADOSSTEEL(lngQTDREGS, 3) = Trim(BREC2!SGI_DECLINHA)
                arrDADOSSTEEL(lngQTDREGS, 4) = Trim(BREC2!SGI_CODPROD)
                arrDADOSSTEEL(lngQTDREGS, 5) = Trim(BREC2!SGI_NOMPROD)
                arrDADOSSTEEL(lngQTDREGS, 6) = "STEEL"
                
                arrDADOSSTEEL(lngQTDREGS, 7) = ""
                arrDADOSSTEEL(lngQTDREGS, 8) = ""
                strDADOSFAT = PegaUltFat(BREC2!SGI_CODCLIE, BREC2!SGI_IDPRODUTO, strNOMEMPTAB, BREC2!SGI_CODVEND)
                strDADOSPEDIDOS = PegaUltMeses(BREC2!SGI_CODCLIE, BREC2!SGI_IDPRODUTO, strNOMEMPTAB, BREC2!SGI_CODVEND, Month(Now))
                
                If Len(Trim(strDADOSFAT)) > 0 Then
                   arrDADOSFAT = Split(strDADOSFAT, "|")
                   arrDADOSSTEEL(lngQTDREGS, 7) = arrDADOSFAT(0)
                   arrDADOSSTEEL(lngQTDREGS, 8) = arrDADOSFAT(1)
                End If
                
                arrDADOSPED = Split(strDADOSPEDIDOS, "|")
                
                arrDADOSSTEEL(lngQTDREGS, 9) = arrDADOSPED(0)
                arrDADOSSTEEL(lngQTDREGS, 10) = arrDADOSPED(1)
                arrDADOSSTEEL(lngQTDREGS, 11) = arrDADOSPED(2)
                
                For i = 0 To (cboESTNORM.ListCount - 1)
                    If cboESTNORM.ItemData(i) = BREC2!SGI_ESTNORM Then
                       cboESTNORM.ListIndex = i
                       Exit For
                    End If
                Next i
                arrDADOSSTEEL(lngQTDREGS, 12) = cboESTNORM.Text
                
                arrDADOSSTEEL(lngQTDREGS, 13) = Trim(BREC2!SGI_CIDNORM)
                
                boolTemDadosSTEEL = True
            End If
            
            BREC2.MoveNext
        Loop
    
    End If
    BREC2.Close

    If optEmpresa(2).value = True Then
        
        '' Steel
        strNOMEMPTAB = "_STEEL"
        
        prgProg.Min = 0
        
        sSql = ""
        
        sSql = "Select Distinct" & vbCrLf
        sSql = sSql & "       CADVE.SGI_DESCRICAO As SGI_DESCVEND" & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_CODIGO     As SGI_CODCLIE" & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_ESTNORM" & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_CIDNORM" & vbCrLf
        sSql = sSql & "      ,CODL.SGI_DESCRI     As SGI_DECLINHA" & vbCrLf
        sSql = sSql & "      ,PEDVI.SGI_CODPROD" & vbCrLf
        sSql = sSql & "      ,PROD.SGI_DESCRICAO  As SGI_NOMPROD" & vbCrLf
        sSql = sSql & "      ,PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "      ,PEDVH.SGI_CODVEND" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDH" & strNOMEMPTAB & "     PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDI" & strNOMEMPTAB & "     PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_CADVENDEDOR     CADVE" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO CODL" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVH.SGI_FILIAL   = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_FILIAL   = PEDVH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO   = PEDVH.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And CODL.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CODL.SGI_CODLIN    = PROD.SGI_CODLINPROD" & vbCrLf
        sSql = sSql & "   And CADVE.SGI_FILIAL   = PEDVH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   ANd CADVE.SGI_CODIGO   = PEDVH.SGI_CODVEND" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL    = PEDVH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO    = PEDVH.SGI_CODCLI" & vbCrLf
        
        If Len(Trim(txtCODVEND.Text)) > 0 Then
            sSql = sSql & "   And PEDVH.SGI_CODVEND    = " & Trim(txtCODVEND.Text) & vbCrLf
        End If
        
        sSql = sSql & "Order By" & vbCrLf
        sSql = sSql & "       CADVE.SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "      ,CODL.SGI_DESCRI" & vbCrLf
        sSql = sSql & "      ,PEDVI.SGI_CODPROD" & vbCrLf
        sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "      ,PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "      ,PEDVH.SGI_CODVEND" & vbCrLf

        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        
        If Not BREC2.EOF() Then
            
            lngQTDREGS = 0
            Do While Not BREC2.EOF()
                lngQTDREGS = (lngQTDREGS + 1)
                BREC2.MoveNext
            Loop
        
            prgProg.Min = 0
            prgProg.Max = lngQTDREGS
            ReDim arrDADOSSTEEL(1 To lngQTDREGS, 1 To 13) As String
        
            BREC2.MoveFirst
            lngQTDREGS = 0
            
            Frame3.Caption = "[ Aguarde gerando dados STEEL.... ]"
            Frame3.Refresh
            
            Do While Not BREC2.EOF()
                lngQTDREGS = (lngQTDREGS + 1)
                prgProg.value = lngQTDREGS
                DoEvents
                
                arrDADOSSTEEL(lngQTDREGS, 1) = Trim(BREC2!SGI_DESCVEND)
                arrDADOSSTEEL(lngQTDREGS, 2) = Trim(BREC2!SGI_RAZAOSOC)
                arrDADOSSTEEL(lngQTDREGS, 3) = Trim(BREC2!SGI_DECLINHA)
                arrDADOSSTEEL(lngQTDREGS, 4) = Trim(BREC2!SGI_CODPROD)
                arrDADOSSTEEL(lngQTDREGS, 5) = Trim(BREC2!SGI_NOMPROD)
                arrDADOSSTEEL(lngQTDREGS, 6) = "STEEL"
                
                arrDADOSSTEEL(lngQTDREGS, 7) = ""
                arrDADOSSTEEL(lngQTDREGS, 8) = ""
                strDADOSFAT = PegaUltFat(BREC2!SGI_CODCLIE, BREC2!SGI_IDPRODUTO, strNOMEMPTAB, BREC2!SGI_CODVEND)
                strDADOSPEDIDOS = PegaUltMeses(BREC2!SGI_CODCLIE, BREC2!SGI_IDPRODUTO, strNOMEMPTAB, BREC2!SGI_CODVEND, Month(Now))
                
                If Len(Trim(strDADOSFAT)) > 0 Then
                   arrDADOSFAT = Split(strDADOSFAT, "|")
                   arrDADOSSTEEL(lngQTDREGS, 7) = arrDADOSFAT(0)
                   arrDADOSSTEEL(lngQTDREGS, 8) = arrDADOSFAT(1)
                End If
                
                arrDADOSPED = Split(strDADOSPEDIDOS, "|")
                
                arrDADOSSTEEL(lngQTDREGS, 9) = arrDADOSPED(0)
                arrDADOSSTEEL(lngQTDREGS, 10) = arrDADOSPED(1)
                arrDADOSSTEEL(lngQTDREGS, 11) = arrDADOSPED(2)
                
                For i = 0 To (cboESTNORM.ListCount - 1)
                    If cboESTNORM.ItemData(i) = BREC2!SGI_ESTNORM Then
                        cboESTNORM.ListIndex = i
                        Exit For
                    End If
                Next i
                arrDADOSSTEEL(lngQTDREGS, 12) = cboESTNORM.Text
                
                arrDADOSSTEEL(lngQTDREGS, 13) = Trim(BREC2!SGI_CIDNORM)
                
                boolTemDadosSTEEL = True
                
                BREC2.MoveNext
            Loop
        
        End If
        BREC2.Close
    
    End If

    Call GeraArgExcel("RELCADVEND" & strNOMARQ & ".xls", boolTemDadosNovalata, boolTemDadosSTEEL)

    MsgBox "Arquivo gerado com sucesso !!!", vbOKOnly + vbInformation, "Aviso"

    Frame3.Visible = False
    
    Me.MousePointer = 0


End Sub

Private Sub GeraArgExcel(strARQUIVO As String, boolTemDadosNOVA As Boolean, boolTemDadosSTEEL As Boolean)

On Error GoTo err_Excel

    Dim myExcelFile             As New clsExcelFile
    Dim FileName$
    Dim lngLINHA                As Long
    Dim lngREGS                 As Long
    Dim lngMES                  As Long
    
    Dim arrMESES(1 To 12) As String
    arrMESES(1) = "Janeiro"
    arrMESES(2) = "Fevereiro"
    arrMESES(3) = "Março"
    arrMESES(4) = "Abril"
    arrMESES(5) = "Maio"
    arrMESES(6) = "Junho"
    arrMESES(7) = "Julho"
    arrMESES(8) = "Agosto"
    arrMESES(9) = "Setembro"
    arrMESES(10) = "Outubro"
    arrMESES(11) = "Novembro"
    arrMESES(12) = "Dezembro"
    
    
    lngMES = Month(Now)
    
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
        .SetColumnWidth 1, 1, 60
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 1, "Nome do Vendedor", 12
        
        .SetColumnWidth 2, 2, 60
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 2, "Nome do Cliente", 12
        
        .SetColumnWidth 3, 3, 20
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 3, "Capacidade", 12
    
        .SetColumnWidth 4, 4, 20
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 4, "Código do Rótulo", 12
    
        .SetColumnWidth 5, 5, 60
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 5, "Descrição do Rótulo", 12
    
        .SetColumnWidth 6, 6, 20
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 6, "Dt.Ult.Fat", 12
    
        .SetColumnWidth 7, 7, 20
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 7, "Qtde.Ult.Fat", 12
    
        .SetColumnWidth 8, 8, 20
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 8, "Empresa", 12
    
        .SetColumnWidth 9, 9, 20
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 9, arrMESES(lngMES), 12
    
        If lngMES = 12 Then
            lngMES = 1
            .SetColumnWidth 10, 10, 20
            .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 10, arrMESES(lngMES), 12
        Else
            .SetColumnWidth 10, 10, 20
            .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 10, arrMESES((lngMES + 1)), 12
        End If
    
        If (lngMES + 1) = 12 Then
            lngMES = 1
            .SetColumnWidth 11, 11, 20
            .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 11, arrMESES(lngMES), 12
        ElseIf lngMES < 11 Then
            .SetColumnWidth 11, 11, 20
            .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 11, arrMESES((lngMES + 1)), 12
        End If
    
        .SetColumnWidth 12, 12, 10
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 12, "Estado", 12
        
        .SetColumnWidth 13, 13, 60
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 13, "Cidade", 12
        
        lngLINHA = 1
        '' Dados da Novalata
        If boolTemDadosNOVA = True Then
             Frame3.Caption = "[ Aguarde.... Gerando o Arguivo EXCEL com os dados da NOVALATA ! ]"
             Frame3.Refresh
             
             prgProg.Min = 0
             prgProg.Max = UBound(arrDADOSNOVA)
             
             For lngREGS = 1 To UBound(arrDADOSNOVA)
                lngLINHA = (lngLINHA + 1)
                
                prgProg.value = lngREGS
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 1, arrDADOSNOVA(lngREGS, 1), 12        '' Nome do vendedor
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 2, arrDADOSNOVA(lngREGS, 2), 12        '' Nome do Cliente
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 3, arrDADOSNOVA(lngREGS, 3), 12        '' Capacidade
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 4, arrDADOSNOVA(lngREGS, 4), 12        '' Código do Rótulo
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 5, arrDADOSNOVA(lngREGS, 5), 12        '' Descrição do Rótulo
                
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 6, arrDADOSNOVA(lngREGS, 7), 12        '' Dt.Ult.Fat
                
                If Len(Trim(arrDADOSNOVA(lngREGS, 8))) = 0 Then arrDADOSNOVA(lngREGS, 8) = 0
                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 7, arrDADOSNOVA(lngREGS, 8), 2      '' Qtde. Ult. Fat
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 8, arrDADOSNOVA(lngREGS, 6), 12        '' Empresa
                
                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 9, arrDADOSNOVA(lngREGS, 9), 1      '' Qtde. Ped
                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 10, arrDADOSNOVA(lngREGS, 10), 1    '' Qtde. Ped
                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 11, arrDADOSNOVA(lngREGS, 11), 1    '' Qtde. Ped
                
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 12, arrDADOSNOVA(lngREGS, 12), 12       '' Estado
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 13, arrDADOSNOVA(lngREGS, 13), 12       '' Cidade
                
                DoEvents
                
             Next lngREGS
        End If
        
        '' Dados da STEEL
        If boolTemDadosSTEEL = True Then
             Frame3.Caption = "[ Aguarde.... Gerando o Arguivo EXCEL com os dados da STEEL ! ]"
             Frame3.Refresh
             
             prgProg.Min = 0
             prgProg.Max = UBound(arrDADOSSTEEL)
             
             For lngREGS = 1 To UBound(arrDADOSSTEEL)
                lngLINHA = (lngLINHA + 1)
                
                prgProg.value = lngREGS
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 1, arrDADOSSTEEL(lngREGS, 1), 12        '' Nome do Vendedor
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 2, arrDADOSSTEEL(lngREGS, 2), 12        '' Nome do Cliente
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 3, arrDADOSSTEEL(lngREGS, 3), 12        '' Capacidade
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 4, arrDADOSSTEEL(lngREGS, 4), 12        '' Código do Rótulo
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 5, arrDADOSSTEEL(lngREGS, 5), 12        '' Descrição do Rótulo
                
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 6, arrDADOSSTEEL(lngREGS, 7), 12        '' Dt.Ult.Fat
                
                If Len(Trim(arrDADOSSTEEL(lngREGS, 8))) = 0 Then arrDADOSSTEEL(lngREGS, 8) = 0
                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 7, arrDADOSSTEEL(lngREGS, 8), 2      '' Qtde. Ult. Fat
                
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 8, arrDADOSSTEEL(lngREGS, 6), 12        '' Empresa
                
                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 9, arrDADOSSTEEL(lngREGS, 9), 1      '' Qtde Ped
                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 10, arrDADOSSTEEL(lngREGS, 10), 1    '' Qtde Ped
                .WriteValue xlsnumber, xlsFont0, xlsRightAlign, xlsNormal, lngLINHA, 11, arrDADOSSTEEL(lngREGS, 11), 1    '' Qtde Ped
                
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 12, arrDADOSSTEEL(lngREGS, 12), 12       '' Estado
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 13, arrDADOSSTEEL(lngREGS, 13), 12       '' Cidade
                
                DoEvents
                
             Next lngREGS
        End If
        
        .ProtectSpreadsheet = False 'False | True
        .CloseFile
    
    End With

    Frame3.Caption = ""
    Frame3.Refresh
    
    prgProg.value = 0

    Exit Sub

err_Excel:

    MsgBox "ATENÇÃO" & vbCrLf & _
           "Erro Numero       : " & Err.Number & vbCrLf & _
           "Descrição do Erro : " & Err.Description, vbOKOnly + vbCritical, "Aviso"

End Sub

Private Sub LimpaCampos()
    lblDescVendedor.Caption = ""
End Sub

Private Sub txtCODVEND_GotFocus()

On Error GoTo Err_txtCODVEND_GotFocus
    
    objBLBFunc.SelecionaCampos txtCODVEND.Name, Me

    Exit Sub
    
Err_txtCODVEND_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : txtCODVEND_GotFocus()", Me.Name, "txtCODVEND_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCODVEND_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODVEND_Validate

    Dim i As Integer
    
    If Len(Trim(txtCODVEND.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVEND.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVEND.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVEND.Text, lblDescVendedor)
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then
       txtCODVEND.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCODVEND_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : txtCODVEND_Validate()", Me.Name, "txtCODVEND_Validate()", strCAMARQERRO)

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

Private Function PegaUltFat(strCodclie As String, strIDPRODUTO As String, strNOMEMPTAB As String, strCODVEND As String) As String

    PegaUltFat = ""
    
    sSql = ""

    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       CONF.SGI_CODCONF" & vbCrLf
    sSql = sSql & "      ,CONF.SGI_DATACONF" & vbCrLf
    sSql = sSql & "      ,CONFI.SGI_QTDREAL" & vbCrLf
    sSql = sSql & "      ,CONF.SGI_CODFATURA" & vbCrLf
    sSql = sSql & "      ,OP.SGI_CODIGO      As SGI_CODOP" & vbCrLf
    sSql = sSql & "      ,PEDVH.SGI_CODIGO   As SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,PEDVH.SGI_CODVEND" & vbCrLf
    sSql = sSql & "      ,PROGE.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH" & strNOMEMPTAB & "  PEDVH" & vbCrLf
    sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMEMPTAB & " PROGE" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMEMPTAB & "    OP" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI" & strNOMEMPTAB & "   ORDF" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFH" & strNOMEMPTAB & "  CONF" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFI" & strNOMEMPTAB & "  CONFI" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PEDVH.SGI_FILIAl    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDVH.SGI_CODVEND   = " & strCODVEND & vbCrLf
    sSql = sSql & "   And PEDVH.SGI_CODCLI    = " & strCodclie & vbCrLf
    sSql = sSql & "   And PROGE.SGI_FILIAl    = PEDVH.SGI_FILIAl" & vbCrLf
    sSql = sSql & "   And PROGE.SGI_CODPED    = PEDVH.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And PROGE.SGI_IDPRODUTO = " & strIDPRODUTO & vbCrLf
    sSql = sSql & "   And OP.SGI_FILIAL       = PROGE.SGI_FILIAl" & vbCrLf
    sSql = sSql & "   And OP.SGI_CODPED       = PROGE.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And OP.SGI_IDPRODUTO    = PROGE.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And ORDF.SGI_FILIAL     = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ORDF.SGI_CODORDFAB  = OP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And ORDF.SGI_IDPRODUTO  = OP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And CONF.SGI_FILIAL     = ORDF.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CONF.SGI_CODORD     = ORDF.SGI_CODORD" & vbCrLf
    sSql = sSql & "   And CONFI.SGI_FILIAL    = CONF.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CONFI.SGI_CODCONF   = CONF.SGI_CODCONF" & vbCrLf
    sSql = sSql & "   And CONFI.SGI_IDPRODUTO = OP.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "         CONF.SGI_DATACONF Desc" & vbCrLf
    sSql = sSql & "        ,CONFI.SGI_QTDREAL"

    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then PegaUltFat = Format(BREC10!SGI_DATACONF, "DD/MM/YYYY") & "|" & BREC10!SGI_QTDREAL
    BREC10.Close

End Function


Private Function PegaUltMeses(strCodclie As String, strIDPRODUTO As String, strNOMEMPTAB As String, strCODVEND As String, lngMESVIG As Long) As String

    PegaUltMeses = ""
    
    Dim lngQTDMES   As Long
    Dim i           As Integer
    Dim strDTINI    As String
    Dim strDTFIN    As String

    lngQTDMES = 0
    For i = 0 To 2
        strDTINI = "'" & Format("01/" & (lngMESVIG + i) & "/" & Format(Year(Now), "####0000"), "MM/DD/YYYY") & "'"
        lngQTDMES = (lngQTDMES + 1)
        strDTFIN = "'" & Format((CDate("01/" & (lngMESVIG + lngQTDMES) & "/" & Format(Year(Now), "####0000")) - 1), "MM/DD/YYYY") & "'"

        sSql = ""
        
        sSql = "Select Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDH" & strNOMEMPTAB & "  PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMEMPTAB & " PROGE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVH.SGI_FILIAl    = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODVEND   = " & strCODVEND & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODCLI    = " & strCodclie & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA Between " & strDTINI & " And " & strDTFIN & vbCrLf
        sSql = sSql & "   And PROGE.SGI_FILIAl    = PEDVH.SGI_FILIAl" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED    = PEDVH.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO = " & strIDPRODUTO

        BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC10.EOF() Then
           If Not IsNull(BREC10!SGI_QTDE) Then
              PegaUltMeses = PegaUltMeses & BREC10!SGI_QTDE & "|"
           Else
              PegaUltMeses = PegaUltMeses & "0|"
           End If
        End If
        BREC10.Close

    Next i

End Function

