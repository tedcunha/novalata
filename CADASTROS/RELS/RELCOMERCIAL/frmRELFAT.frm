VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRELFAT 
   Caption         =   "Relatorio de Faturamentos"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCODVEND 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   10
      TabIndex        =   18
      Text            =   "txtCODVEND"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   2160
      Picture         =   "frmRELFAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2160
      Width           =   375
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Tipo do Rótulo ]"
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
      TabIndex        =   14
      Top             =   1440
      Width           =   4935
      Begin VB.OptionButton optTIPROT 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTIPROT 
         Caption         =   "Normal"
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
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTIPROT 
         Caption         =   "Homologado"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
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
      TabIndex        =   12
      Top             =   2520
      Width           =   9735
      Begin ComctlLib.ProgressBar prbDados 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9735
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
         Picture         =   "frmRELFAT.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
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
         Picture         =   "frmRELFAT.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   855
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
      TabIndex        =   4
      Top             =   840
      Width           =   4935
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3600
         TabIndex        =   5
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
         TabIndex        =   6
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
         Left            =   2640
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
      TabIndex        =   0
      Top             =   840
      Width           =   4815
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
         TabIndex        =   3
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
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   975
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
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label lblDescVendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDescVendedor"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2520
      TabIndex        =   20
      Top             =   2160
      Width           =   7215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor:"
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
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   2205
      Width           =   885
   End
End
Attribute VB_Name = "frmRELFAT"
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
Dim objRELFAT           As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object

Dim strCABEC1           As String
Dim strCABEC2           As String
Dim strNomRel           As String
Dim strEMPRESADESC      As String
Dim strEMPRESA          As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call Imprime
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
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ATIVO      = 1"
    
    ''If lngCodVendedor > 0 Then
    ''    sSql = sSql & "   And SGI_CODIGO = " & lngCodVendedor
    ''End If
    
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
    
    Call PegaDescTabelasVend("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedor, "Command2_Click()")
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then txtCODVEND.Text = ""
    
    If txtCODVEND.Enabled = True Then txtCODVEND.SetFocus

    Exit Sub
    
Err_Command2_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "P", "Função : Command2_Click()", Me.Name, "Command2_Click()", strCAMARQERRO)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELFAT = CreateObject("RELCOMERCIAL.clsRELFAT")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    
    objRELFAT.FILIAL = FILIAL
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    Me.Caption = Me.Caption & " / " & Trim(Me.Name)


    optEmpresa(0).value = True
    optTIPROT(2).value = True
    Frame3.Caption = ""
    Frame3.Refresh
    prbDados.Min = 0

    Call LimpaCamposLabel

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELFAT = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub LimpaCamposLabel()
    lblDescVendedor.Caption = ""
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
    
        sSql = "Select Distinct" & vbCrLf
        sSql = sSql & "       CH.SGI_DATACONF" & vbCrLf
        sSql = sSql & "      ,CF.SGI_CODORDPROD" & vbCrLf
        sSql = sSql & "      ,PD.SGI_CODTIPO" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADORDCONFH" & strEMPRESA & " CH" & vbCrLf
        sSql = sSql & "      ,SGI_CADORDCONFI" & strEMPRESA & " CF" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & "   OP" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PD" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & " PED" & vbCrLf
        
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "       CH.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And CH.SGI_DATACONF Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And CF.SGI_FILIAL    = CH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CF.SGI_CODCONF   = CH.SGI_CODCONF" & vbCrLf
        sSql = sSql & "   And OP.SGI_FILIAL    = CF.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And OP.SGI_CODIGO    = CF.SGI_CODORDPROD" & vbCrLf
        sSql = sSql & "   And PD.SGI_FILIAL    = OP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PD.SGI_IDPRODUTO = OP.SGI_IDPRODUTO" & vbCrLf
        
        If optTIPROT(0).value = True Then '' Homologada
            sSql = sSql & "   And PD.SGI_CODTIPO   = 2" & vbCrLf
        ElseIf optTIPROT(1).value = True Then '' Normal
            sSql = sSql & "   And PD.SGI_CODTIPO   = 1" & vbCrLf
        End If
        
        sSql = sSql & "   And PED.SGI_FILIAL   = OP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO   = OP.SGI_CODPED" & vbCrLf
        If Len(Trim(Trim(txtCODVEND.Text))) > 0 Then
            sSql = sSql & "   And PED.SGI_CODVEND  = " & Trim(txtCODVEND.Text) & vbCrLf
        End If
        
        sSql = sSql & "Order By CH.SGI_DATACONF,CF.SGI_CODORDPROD"
        
        strSQLNOVALATA = sSql
        
        If BREC.State = 1 Then BREC.Close
        
        BREC.Open sSql, adoBanco_Dados
        If Not BREC.EOF() Then
            boolTemRegs = True
            
                prbDados.Min = 0
                
                lngREGS = 0
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    BREC.MoveNext
                Loop
                lngQTDREGSNOVA = lngREGS
                
                prbDados.Max = lngREGS
                ReDim arrDADOSTAB(1 To lngREGS, 1 To 4) As String
                BREC.MoveFirst
                lngREGS = 0
            
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    prbDados.value = lngREGS
                    
                    arrDADOSTAB(lngREGS, 1) = Format(BREC!SGI_DATACONF, "DD/MM/YYYY")
                    arrDADOSTAB(lngREGS, 2) = BREC!SGI_CODORDPROD
                    If BREC!SGI_CODTIPO = 2 Then
                        arrDADOSTAB(lngREGS, 3) = "HOMOLOGADA"
                    Else
                        arrDADOSTAB(lngREGS, 3) = "NORMAL"
                    End If
                    arrDADOSTAB(lngREGS, 4) = "NOVALATA"
                    
                    BREC.MoveNext
                Loop
        End If
        BREC.Close
    End If
        
    
    
    '' Geração para NOVALATA e STEEL
    If (optEmpresa(1).value = True Or optEmpresa(2).value = True) Then
        
        strEMPRESA = "_STEEL"
        
        sSql = ""
        
        sSql = "Select Distinct" & vbCrLf
        sSql = sSql & "       CH.SGI_DATACONF" & vbCrLf
        sSql = sSql & "      ,CF.SGI_CODORDPROD" & vbCrLf
        sSql = sSql & "      ,PD.SGI_CODTIPO" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADORDCONFH" & strEMPRESA & " CH" & vbCrLf
        sSql = sSql & "      ,SGI_CADORDCONFI" & strEMPRESA & " CF" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strEMPRESA & "   OP" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PD" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & " PED" & vbCrLf
        
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "       CH.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And CH.SGI_DATACONF Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And CF.SGI_FILIAL    = CH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CF.SGI_CODCONF   = CH.SGI_CODCONF" & vbCrLf
        sSql = sSql & "   And OP.SGI_FILIAL    = CF.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And OP.SGI_CODIGO    = CF.SGI_CODORDPROD" & vbCrLf
        sSql = sSql & "   And PD.SGI_FILIAL    = OP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PD.SGI_IDPRODUTO = OP.SGI_IDPRODUTO" & vbCrLf
        
        If optTIPROT(0).value = True Then '' Homologada
            sSql = sSql & "   And PD.SGI_CODTIPO   = 2" & vbCrLf
        ElseIf optTIPROT(1).value = True Then '' Normal
            sSql = sSql & "   And PD.SGI_CODTIPO   = 1" & vbCrLf
        End If
        
        sSql = sSql & "   And PED.SGI_FILIAL   = OP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO   = OP.SGI_CODPED" & vbCrLf
        If Len(Trim(Trim(txtCODVEND.Text))) > 0 Then
            sSql = sSql & "   And PED.SGI_CODVEND  = " & Trim(txtCODVEND.Text) & vbCrLf
        End If
        
        sSql = sSql & "Order By CH.SGI_DATACONF,CF.SGI_CODORDPROD"
        
        strSQLSTEEL = sSql
        
        If BREC.State = 1 Then BREC.Close
        
        BREC.Open sSql, adoBanco_Dados
        If Not BREC.EOF() Then
            boolTemRegs = True
            
                '' Steel
            
                lngREGS = 0
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    BREC.MoveNext
                Loop
                lngQTDREGSTEEL = lngREGS
                
                BREC.MoveFirst
                prbDados.Max = lngREGS
                ReDim arrDADOSTAB_STEEL(1 To lngREGS, 1 To 4) As String
                lngREGS = 0
            
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    prbDados.value = lngREGS
                    
                    arrDADOSTAB_STEEL(lngREGS, 1) = Format(BREC!SGI_DATACONF, "DD/MM/YYYY")
                    arrDADOSTAB_STEEL(lngREGS, 2) = BREC!SGI_CODORDPROD
                    If BREC!SGI_CODTIPO = 2 Then
                        arrDADOSTAB_STEEL(lngREGS, 3) = "HOMOLOGADA"
                    Else
                        arrDADOSTAB_STEEL(lngREGS, 3) = "NORMAL"
                    End If
                    arrDADOSTAB_STEEL(lngREGS, 4) = "STEEL"
                    
                    BREC.MoveNext
                Loop
        End If
        BREC.Close
    End If
        
    
    If boolTemRegs = False Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    Call ExportaParaExcel(arrDADOSTAB, lngQTDREGSNOVA, arrDADOSTAB_STEEL, lngQTDREGSTEEL)

End Sub

Private Sub ExportaParaExcel(arrDADOSTAB() As String, lngQTDREGSNOVA As Long, arrDADOSTAB_STEEL() As String, lngQTDRESSTEEL As Long)

On Error GoTo Handle_Error

    Dim myExcelFile             As New clsExcelFile
    Dim FileName$
    Dim boolTemDados            As Boolean
    
    Dim lngREGS                 As Long
    Dim lngLINHA                As Long
    Dim lngQTDPED               As Long
    Dim lngQTDFAT               As Long
    Dim lngSALDO                As Long

    If lngQTDREGSNOVA = 0 And lngQTDRESSTEEL = 0 Then
        MsgBox "Atenção - Não há dados para gerar o arquivo !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If

    With myExcelFile
        'Create the new spreadsheet
        If optEmpresa(0).value = True Then
           FileName$ = strCamRelNovo & "RELPREPARA\RELFAT_NOVALATA.xls"
        ElseIf optEmpresa(1).value = True Then
           FileName$ = strCamRelNovo & "RELPREPARA\RELFAT_STEEL.xls"
        ElseIf optEmpresa(2).value = True Then
           FileName$ = strCamRelNovo & "RELPREPARA\RELFAT_NS.xls"
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
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 1, "Dt.Faturmaneto", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 2, "Cód.OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 3, "Tipo.OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 4, "Empresa", 12
        
        If lngQTDREGSNOVA > 0 Then
        
            '' Jogando os Dados na Planilha
            '' NOVALATA
            lngLINHA = 1
            prbDados.Min = 0
            prbDados.Max = UBound(arrDADOSTAB)
            
            For lngREGS = 1 To UBound(arrDADOSTAB) '' Novalata
                lngLINHA = (lngLINHA + 1)
                
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 1, arrDADOSTAB(lngREGS, 1), 1
                .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 2, arrDADOSTAB(lngREGS, 2), 1
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 3, arrDADOSTAB(lngREGS, 3), 12
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 4, arrDADOSTAB(lngREGS, 4), 12
        
                prbDados.value = lngREGS
            Next lngREGS
        
        End If
        
        If lngQTDRESSTEEL > 0 Then
            
            prbDados.Min = 0
            prbDados.Max = UBound(arrDADOSTAB_STEEL)
            
            For lngREGS = 1 To UBound(arrDADOSTAB_STEEL) '' Steel
                lngLINHA = (lngLINHA + 1)
                
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 1, arrDADOSTAB_STEEL(lngREGS, 1), 1
                .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 2, arrDADOSTAB_STEEL(lngREGS, 2), 1
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 3, arrDADOSTAB_STEEL(lngREGS, 3), 12
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 4, arrDADOSTAB_STEEL(lngREGS, 4), 12
                
                 prbDados.value = lngREGS
            Next lngREGS
        End If
        
        'PROTECT the spreadsheet so any cells specified as LOCKED will not be
        'overwritten. Also, all cells with HIDDEN set will hide their formula.
        'PROTECT does not use a password.
        .ProtectSpreadsheet = False 'False | True
        
        'Finally, close the spreadsheet
        .CloseFile
        
        MsgBox "Arquivo Excel : " & " foi Criado !", vbInformation + vbOKOnly, "Aviso do Sistema"
    End With
    
    Exit Sub
    
Handle_Error:

    If BREC.State = 1 Then BREC.Close
    MsgBox "Número: " & Err.Number & vbCrLf & "Descrição: " & Err.Description, vbOKOnly + vbCritical, "Aviso"

        
End Sub

Private Sub txtCODVEND_GotFocus()

On Error GoTo Err_txtCODVEND_GotFocus
    
    objBLBFunc.SelecionaCampos txtCODVEND.Name, Me

    Exit Sub
    
Err_txtCODVEND_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "P", "Função : txtCODVEND_GotFocus()", Me.Name, "txtCODVEND_GotFocus()", strCAMARQERRO)

End Sub

Private Sub PegaDescTabelasVend(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label, strFUNCAOPAI As String)

On Error GoTo Err_PegaDescTabelasVend

    lblLabel.Caption = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(Replace(Replace(strCODIGO, ".", ""), ",", ""))) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCODIGO) & vbCrLf
    sSql = sSql & "   And SGI_ATIVO = 1"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
    Exit Sub
    
Err_PegaDescTabelasVend:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "P", "Função : PegaDescTabelasVend()" & vbCrLf & "Função Pai :" & strFUNCAOPAI, Me.Name, "PegaDescTabelasVend()", strCAMARQERRO)

End Sub

Private Sub txtCODVEND_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCODVEND_KeyPress

    objBLBFunc.SoNumeroPonto KeyAscii, txtCODVEND.Text

    Exit Sub
    
Err_txtCODVEND_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "P", "Função : txtCODVEND_KeyPress()", Me.Name, "txtCODVEND_KeyPress()", strCAMARQERRO)

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
    
    txtCODVEND.Text = Trim(Replace(Replace(txtCODVEND.Text, ",", ""), ".", ""))
    
    Call PegaDescTabelasVend("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVEND.Text, lblDescVendedor, "txtCODVEND_Validate()")
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then
       txtCODVEND.Text = ""
       Cancel = True
    End If
    
    Exit Sub
    
Err_txtCODVEND_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "P", "Função : txtCODVEND_Validate()", Me.Name, "txtCODVEND_Validate()", strCAMARQERRO)

End Sub
