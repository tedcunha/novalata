VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCADLIMPACARTEIRA 
   Caption         =   "Limpa Carteira"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame27 
      Caption         =   "[ Motivo da Liquidação da OP ]"
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
      Height          =   2175
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   7575
      Begin VB.Frame Frame29 
         Caption         =   "[ Observação ]"
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
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtOBS_MotLiq 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   17
            Text            =   "frmCADLIMPACARTEIRA.frx":0000
            Top             =   240
            Width           =   7095
         End
      End
      Begin VB.Frame Frame28 
         Caption         =   "[ Motivo ]"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7335
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   1200
            Picture         =   "frmCADLIMPACARTEIRA.frx":000E
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtCODMOTLIQ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   13
            Text            =   "txtCODMOTL"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblDescMotLiq 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescMotLiq"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1560
            TabIndex        =   15
            Top             =   240
            Width           =   5655
         End
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
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Width           =   7575
      Begin ComctlLib.ProgressBar prgProgresso 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
         Min             =   1e-4
         Max             =   100
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
      TabIndex        =   4
      Top             =   840
      Width           =   7575
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   3015
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
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   10
            Top             =   0
            Width           =   855
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
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   9
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optEmpresa 
            Caption         =   "Todas"
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
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdExecuta 
         Caption         =   "&Executar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskDTLIMITE 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Até Data"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7575
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
         Picture         =   "frmCADLIMPACARTEIRA.frx":0110
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADLIMPACARTEIRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho             As String
Public Linha                As Variant
Public FILIAL               As Integer
Public strAcesso            As String
Public lngCodUsuario        As Long
Dim iCodigo                 As Integer

Dim objFuncoes              As Object
Dim objCADLIMPACARTEIRA     As Object
Dim objPESQPADRAO           As Object
Dim strEMPRESA              As String
Dim lngQTDEREG_NOVALATA     As Long
Dim lngQTDEREG_STEEL        As Long
Dim strDTENTREGA            As String
    

Private Sub cmdExecuta_Click()
    If ConfereCampos = False Then Exit Sub
    Call Executar
End Sub

Private Sub Command6_Click()

On Error GoTo Err_Command6_Click

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRI " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADMOTLIQOP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Motivos de Liquidação do OP")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODMOTLIQ.Text = varRETORNO
    
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMOTLIQOP", varRETORNO, lblDescMotLiq)
    If Len(Trim(lblDescMotLiq.Caption)) = 0 Then txtCODMOTLIQ.Text = ""
    
    If txtOBS_MotLiq.Enabled = True Then txtOBS_MotLiq.SetFocus

    Exit Sub
    
Err_Command6_Click:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : Command6_Click()", Me.Name, "Command6_Click()", strCAMARQERRO)

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

Private Sub Destroy_Objeto()
    Set objFuncoes = Nothing
    Set objCADLIMPACARTEIRA = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADLIMPACARTEIRA = CreateObject("CADLIMPACARTEIRA.clsCADLIMPACARTEIRA")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    
    objCADLIMPACARTEIRA.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos Me
    
    Call AbreBanco

    mskDTLIMITE.Text = Format(Now, "DD/MM/YYYY")
    Call PrgBar_Visivel(False)
    optEmpresa(0).Value = True
    lblDescMotLiq.Caption = ""

    strEMPRESA = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
        If Not IsDate(mskDTLIMITE.Text) Then
            MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTLIMITE.SetFocus
            Exit Function
        End If
        If Len(Trim(txtCODMOTLIQ.Text)) = 0 Then
            MsgBox "Informe o Motivo da Liquidação !!!", vbOKOnly + vbExclamation, "Aviso"
            txtCODMOTLIQ.SetFocus
            Exit Function
        End If
        If Len(Trim(txtOBS_MotLiq.Text)) = 0 Then
            MsgBox "Informe uma Observação para a liquidação !!!", vbOKOnly + vbExclamation, "Aviso"
            txtCODMOTLIQ.SetFocus
            Exit Function
        End If
        
    ConfereCampos = True

End Function


Private Sub Executar()

On Error GoTo Err_Executar
    
    Dim strNomeMod      As String
    Dim strNomeModPed   As String
    Dim I               As Long
    Dim arrDADOSLOG     As Variant
    
    strDTENTREGA = "'" & Format(CDate(mskDTLIMITE.Text), "MM/DD/YYYY") & "'"
    lngQTDEREG_NOVALATA = 0
    lngQTDEREG_STEEL = 0
    
    objCADLIMPACARTEIRA.OBSLIQ = "'" & Trim(Replace(txtOBS_MotLiq.Text, ",", " ")) & "'"
    objCADLIMPACARTEIRA.CODMOTLIQOP = Trim(txtCODMOTLIQ.Text)
    objCADLIMPACARTEIRA.DTENTREGA = strDTENTREGA
    
    
    If optEmpresa(0).Value = True Or optEmpresa(1).Value = True Then '' 0 - Todas as Empresas / 1 Somente Novalata
    
        sSql = ""
    
        sSql = "Select" & vbCrLf
        sSql = sSql & "      Count(*) as QtdeRegs" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_ORDEMPROD" & vbCrLf
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "      SGI_FILIAl      = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_DATENTREGA  < " & strDTENTREGA & vbCrLf
        sSql = sSql & "  And SGI_STATUS     In(0,1)"

        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then lngQTDEREG_NOVALATA = BREC!QtdeRegs
        BREC.Close
        
    End If
        
    If optEmpresa(0).Value = True Or optEmpresa(2).Value = True Then '' 0 - Todas as Empresas / 2 Somente Steel
        
        sSql = ""
    
        sSql = "Select" & vbCrLf
        sSql = sSql & "      Count(*) as QtdeRegs" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_ORDEMPROD_STEEL" & vbCrLf
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "      SGI_FILIAl      = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_DATENTREGA  < " & strDTENTREGA & vbCrLf
        sSql = sSql & "  And SGI_STATUS     In(0,1)"
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then lngQTDEREG_STEEL = BREC!QtdeRegs
        BREC.Close
    
    End If

    
    If lngQTDEREG_NOVALATA > 0 Then
        Call PrgBar_Visivel(True)
        Call Reseta_ProgresBar(lngQTDEREG_NOVALATA)
        Call ProcessaDados("", lngQTDEREG_NOVALATA)
    End If
    
    Call PrgBar_Visivel(False)
    If lngQTDEREG_STEEL > 0 Then
        Call PrgBar_Visivel(True)
        Call Reseta_ProgresBar(lngQTDEREG_STEEL)
        Call ProcessaDados("_STEEL", lngQTDEREG_STEEL)
    End If
    Call PrgBar_Visivel(False)
    
    '' Gravando os Dados
    If lngQTDEREG_NOVALATA > 0 Or lngQTDEREG_STEEL > 0 Then
        If objCADLIMPACARTEIRA.GravaDados(lngCodUsuario) = True Then
            MsgBox "Dados Processados com exito...", vbOKOnly + vbInformation, "Aviso"
        Else
            MsgBox "OS Dados não foram Processados...", vbOKOnly + vbInformation, "Aviso"
        End If
    Else
        MsgBox "Não existe dados para Processar !", vbOKOnly + vbExclamation, "Aviso"
    End If
    
    Exit Sub

Err_Executar:

    MsgBox "ATENÇÃO" & vbCrLf & _
           "Erro nº   : " & Err.Number & vbCrLf & _
           "Descrição : " & Err.Description & vbCrLf & _
           "Função    : Executar", vbOKOnly + vbCritical, "Aviso"

End Sub

Private Sub Reseta_ProgresBar(lngQTDE As Long)
    prgProgresso.Min = 0
    prgProgresso.Max = lngQTDE
End Sub

Private Sub PrgBar_Visivel(boolVisivel As Boolean)
    Frame3.Visible = boolVisivel
    prgProgresso.Visible = boolVisivel
End Sub

Private Sub ProcessaDados(strFILIAL As String, lngTOTGER_REGS As Long)
   
On Error GoTo Err_ProcessaDados
    
    Dim lngREGS      As Long
    Dim strSTATUS    As String
    Dim strSTATUSPED As String
    Dim arrDADOS     As Variant
    
    If lngTOTGER_REGS > 0 Then
        ReDim arrDADOS(1 To lngTOTGER_REGS, 1 To 5) As String
    End If
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       OP.SGI_DATENTREGA " & vbCrLf
    sSql = sSql & "      ,OP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,OP.SGI_IDPAI" & vbCrLf
    sSql = sSql & "      ,OP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,OP.SGI_CODPED" & vbCrLf

    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strFILIAL & " OP" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       OP.SGI_FILIAl      = " & FILIAL & vbCrLf
    sSql = sSql & "   And OP.SGI_DATENTREGA  < " & strDTENTREGA & vbCrLf
    sSql = sSql & "   And OP.SGI_STATUS      In(0,1)" & vbCrLf
    
    sSql = sSql & "Order By OP.SGI_CODIGO"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        lngREGS = 1
        Do While Not BREC.EOF()
        
            prgProgresso.Value = lngREGS
            Frame3.Caption = "[ " & Format(((lngREGS / lngTOTGER_REGS) * 100), "#0") & "% Completo... " & IIf(Len(Trim(strFILIAL)) = 0, "NOVALATA", "STEEL") & "]"
            Frame3.Refresh
            
            '' ==================================
            '' Pega Dados
            arrDADOS(lngREGS, 1) = "'" & Format(BREC!SGI_DATENTREGA, "MM/DD/YYYY") & "'"
            arrDADOS(lngREGS, 2) = BREC!SGI_CODIGO
            
            arrDADOS(lngREGS, 3) = ""
            If Not IsNull(BREC!SGI_IDPAI) Then arrDADOS(lngREGS, 3) = BREC!SGI_IDPAI
            
            arrDADOS(lngREGS, 4) = BREC!SGI_IDPRODUTO
            arrDADOS(lngREGS, 5) = BREC!SGI_CODPED
            '' ==================================
            
            lngREGS = lngREGS + 1
            DoEvents
            
            BREC.MoveNext
            
        Loop
    End If
    BREC.Close
    
    If Len(Trim(strFILIAL)) > 0 Then objCADLIMPACARTEIRA.DADOSSTEEL = arrDADOS
    If Len(Trim(strFILIAL)) = 0 Then objCADLIMPACARTEIRA.DADOSNOVA = arrDADOS
    
    Exit Sub
    
Err_ProcessaDados:

    If BREC.State = 1 Then BREC.Close
        
    MsgBox "ATENÇÃO" & vbCrLf & _
           "Erro nº   : " & Err.Number & vbCrLf & _
           "Descrição : " & Err.Description & vbCrLf & _
           "Função    : ProcessaDados" & vbCrLf & _
           "Registro  : " & lngREGS, vbOKOnly + vbCritical, "Aviso"

End Sub

Private Function PegaStatusOP(strCODOP As String, strNOMFILIAL As String) As Long

    PegaStatusOP = 2

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_STATUS" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCODOP
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then PegaStatusOP = BREC10!SGI_STATUS
    BREC10.Close
    
End Function

Private Sub mskDTLIMITE_GotFocus()
    objFuncoes.SelecionaCampos mskDTLIMITE.Name, Me
End Sub

Private Sub txtCODMOTLIQ_GotFocus()

On Error GoTo Err_txtCODMOTLIQ_GotFocus
    
    objFuncoes.SelecionaCampos txtCODMOTLIQ.Name, Me

    Exit Sub
    
Err_txtCODMOTLIQ_GotFocus:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : txtCODMOTLIQ_GotFocus()", Me.Name, "txtCODMOTLIQ_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCODMOTLIQ_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCODMOTLIQ_KeyPress

    objFuncoes.SoNumeroPonto KeyAscii, txtCODMOTLIQ.Text

    Exit Sub
    
Err_txtCODMOTLIQ_KeyPress:

    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : txtCODMOTLIQ_KeyPress()", Me.Name, "txtCODMOTLIQ_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCODMOTLIQ_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODMOTLIQ_Validate

    Dim I As Integer
    
    If Len(Trim(txtCODMOTLIQ.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODMOTLIQ.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODMOTLIQ.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMOTLIQOP", txtCODMOTLIQ.Text, lblDescMotLiq)
    If Len(Trim(lblDescMotLiq.Caption)) = 0 Then
       txtCODMOTLIQ.Text = ""
       Cancel = True
    Else
        txtOBS_MotLiq.SetFocus
    End If
    
    Exit Sub
    
Err_txtCODMOTLIQ_Validate:
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : txtCODMOTLIQ_Validate()", Me.Name, "txtCODMOTLIQ_Validate()", strCAMARQERRO)

End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label)

On Error GoTo Err_PegaDescTabelas

    lblLabel.Caption = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(Replace(Replace(strCODIGO, ".", ""), ",", ""))) = 0 Then Exit Sub
    
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
    
    Call objFuncoes.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub


Private Sub AbreBanco()
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If

End Sub
