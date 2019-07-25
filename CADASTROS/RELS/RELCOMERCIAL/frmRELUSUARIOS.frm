VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmRELUSUARIOS 
   Caption         =   "Relatório de Usuários do Sistema"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraVisua 
      Caption         =   "[ Visualização ]"
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
      Top             =   1440
      Width           =   8415
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "Em Excel"
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
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "Em tela"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   8415
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Usuários ]"
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
      Top             =   840
      Width           =   8415
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   1080
         Picture         =   "frmRELUSUARIOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCODUSU 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "txtCODUSU"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblDescUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescUsuario"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
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
         Picture         =   "frmRELUSUARIOS.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmRELUSUARIOS.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELUSUARIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strACESSO        As String
Public lngCodUsuario    As Long

Dim objBLBFunc          As Object
Dim objRELUSUARIOS      As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String
Dim strNomRel           As String
Dim strEMPRESADESC      As String
Dim lngCODIGO           As Long

Dim arrNivelMenuP()     As NivelMenuP

Private Type NivelMenuM
    strFILIAL           As String
    strCODUSUARIO       As String
    strCodMenu          As String
    strCodFilho         As String
    strCodNeto          As String
    strTIPO             As String
    strCIGLA            As String
    strACESSO           As String
    strTEXTO            As String
End Type

Private Type NivelMenuS
    strFILIAL           As String
    strCODUSUARIO       As String
    strCodMenu          As String
    strCodPai           As String
    strTIPO             As String
    strCIGLA            As String
    strCIGLA2           As String
    strTEXTO            As String
    arrNivelMenuM()     As NivelMenuM
End Type

Private Type NivelMenuP
    strFILIAL           As String
    strCODUSUARIO       As String
    strUSUARIO          As String
    strCodMenu          As String
    strTIPO             As String
    strCIGLA            As String
    strTEXTO            As String
    strSTATUS           As String
    strBLOQPRDIDO       As String
    arrNivelMenuS()     As NivelMenuS
End Type

Private Sub cmdImpressao_Click()
    Call Gera_Rel
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_NOME " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_USUARIO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_NOME"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Usuário"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_NOME"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Usuários", "", "", True)
    
    If Len(Trim(varRETORNO)) = 0 Then Exit Sub
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_NOME", "SGI_USUARIO", varRETORNO, lblDescUsuario)
    If Len(Trim(lblDescUsuario.Caption)) = 0 Then
        txtCODUSU.Text = ""
        lblDescUsuario.Caption = ""
        Exit Sub
    End If
    
    txtCODUSU.Text = varRETORNO
    lblDescUsuario.Caption = Trim(objBLBFunc.Crypt(lblDescUsuario.Caption))

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELUSUARIOS = CreateObject("RELCOMERCIAL.clsRELUSUARIOS")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call objBLBFunc.LimpaCampos(Me)
    
    objRELUSUARIOS.FILIAL = FILIAL
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    Me.Caption = Me.Caption & " / " & Me.Name
    
    Call LimpaCamposLabel
    
    Frame2.Visible = False
    optVisualizacao(0).value = True
    
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELUSUARIOS = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub txtCODUSU_GotFocus()
    objBLBFunc.SelecionaCampos txtCODUSU.Name, Me
End Sub

Private Sub txtCODUSU_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODUSU.Text
End Sub

Private Sub txtCODUSU_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODUSU.Text)) = 0 Then
        lblDescUsuario.Caption = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCODUSU.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODUSU.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_NOME", "SGI_USUARIO", txtCODUSU.Text, lblDescUsuario)
    If Len(Trim(lblDescUsuario.Caption)) = 0 Then
       txtCODUSU.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblDescUsuario.Caption = Trim(objBLBFunc.Crypt(lblDescUsuario.Caption))

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
    lblDescUsuario.Caption = ""
End Sub

Private Sub Gera_Rel()
    
    Dim intREGSP    As Integer
    Dim intREGSS    As Integer
    Dim intREGSM    As Integer
    Dim arrNivelS() As NivelMenuS
    Dim arrNivelM() As NivelMenuM
    
    sSql = ""
    intREGSP = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       USUA.SGI_NOME" & vbCrLf
    sSql = sSql & "     , USUA.SGI_ATIVO" & vbCrLf
    sSql = sSql & "     , USUA.SGI_PERMBLOQPED" & vbCrLf
    sSql = sSql & "     , MENU.* " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO USUA" & vbCrLf
    sSql = sSql & "     , SGI_MENUP   MENU" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       USUA.SGI_FILIAL        = " & FILIAL & vbCrLf
    
    If Len(Trim(txtCODUSU.Text)) > 0 Then
        sSql = sSql & "   And USUA.SGI_CODIGO = " & Trim(txtCODUSU.Text) & vbCrLf
    End If
    
    sSql = sSql & "   And MENU.SGI_FILIAL        = USUA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And MENU.SGI_CODUSUARIO    = USUA.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And MENU.SGI_TIPO          = 'P'" & vbCrLf
    sSql = sSql & "   And MENU.SGI_ATIVO         = 1" & vbCrLf
    
    sSql = sSql & "Order By USUA.SGI_CODIGO,MENU.SGI_CODIGO"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
        
        intREGSP = (intREGSP + 1)
        ReDim Preserve arrNivelMenuP(1 To intREGSP) As NivelMenuP
        arrNivelMenuP(intREGSP).strFILIAL = FILIAL
        arrNivelMenuP(intREGSP).strCODUSUARIO = BREC!SGI_CODUSUARIO
        arrNivelMenuP(intREGSP).strUSUARIO = Trim(objBLBFunc.Crypt(BREC!SGI_NOME))
        arrNivelMenuP(intREGSP).strCodMenu = BREC!SGI_CODIGO
        arrNivelMenuP(intREGSP).strTIPO = BREC!SGI_TIPO
        arrNivelMenuP(intREGSP).strCIGLA = BREC!SGI_CIGLA
        arrNivelMenuP(intREGSP).strTEXTO = BREC!SGI_TEXTO
        
        arrNivelMenuP(intREGSP).strSTATUS = "DESATIVADO"
        If BREC!SGI_ATIVO = 1 Then arrNivelMenuP(intREGSP).strSTATUS = "ATIVO"
        
        '' Permite Bloquear Pedido
        arrNivelMenuP(intREGSP).strBLOQPRDIDO = "NÃO"
        If BREC!SGI_PERMBLOQPED = 1 Then arrNivelMenuP(intREGSP).strBLOQPRDIDO = "SIM"
        
        '' Tipo S
        sSql = ""
        intREGSS = 0
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_MENUP " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODUSUARIO = " & BREC!SGI_CODUSUARIO & vbCrLf
        sSql = sSql & "   And SGI_TIPO       = 'S'" & vbCrLf
        sSql = sSql & "   And SGI_CIGLA      = '" & Trim(BREC!SGI_CIGLA) & "'" & vbCrLf
        sSql = sSql & "   And SGI_ATIVO      = 1" & vbCrLf
        sSql = sSql & "Order By SGI_CODIGO"
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC2.EOF()
            intREGSS = (intREGSS + 1)
            ReDim Preserve arrNivelS(1 To intREGSS) As NivelMenuS
            arrNivelS(intREGSS).strFILIAL = FILIAL
            arrNivelS(intREGSS).strCODUSUARIO = BREC2!SGI_CODUSUARIO
            arrNivelS(intREGSS).strCodMenu = BREC2!SGI_CODIGO
            arrNivelS(intREGSS).strCodPai = BREC!SGI_CODIGO
            arrNivelS(intREGSS).strTIPO = BREC2!SGI_TIPO
            arrNivelS(intREGSS).strCIGLA = BREC2!SGI_CIGLA
            arrNivelS(intREGSS).strCIGLA2 = BREC2!SGI_CIGLA2
            arrNivelS(intREGSS).strTEXTO = BREC2!SGI_TEXTO
            
            '' Tipo M
            sSql = ""
            intREGSM = 0
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_MENUP " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODUSUARIO = " & BREC2!SGI_CODUSUARIO & vbCrLf
            sSql = sSql & "   And SGI_TIPO       = 'M'" & vbCrLf
            sSql = sSql & "   And SGI_CIGLA      = '" & Trim(BREC2!SGI_CIGLA2) & "'" & vbCrLf
            sSql = sSql & "   And SGI_ATIVO      = 1" & vbCrLf
            sSql = sSql & "Order By SGI_CODIGO"

            BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
            Do While Not BREC3.EOF()
                intREGSM = (intREGSM + 1)
                ReDim Preserve arrNivelM(1 To intREGSM) As NivelMenuM
                arrNivelM(intREGSM).strFILIAL = FILIAL
                arrNivelM(intREGSM).strCODUSUARIO = BREC3!SGI_CODUSUARIO
                arrNivelM(intREGSM).strCodMenu = BREC3!SGI_CODIGO
                arrNivelM(intREGSM).strCodFilho = BREC2!SGI_CODIGO
                arrNivelM(intREGSM).strCodNeto = BREC!SGI_CODIGO
                arrNivelM(intREGSM).strTIPO = BREC3!SGI_TIPO
                arrNivelM(intREGSM).strCIGLA = BREC3!SGI_CIGLA
                arrNivelM(intREGSM).strTEXTO = BREC3!SGI_TEXTO
                arrNivelM(intREGSM).strACESSO = BREC3!SGI_ACESSO
                BREC3.MoveNext
            Loop
            BREC3.Close
            If intREGSM > 0 Then arrNivelS(intREGSS).arrNivelMenuM = arrNivelM
            
            BREC2.MoveNext
        Loop
        BREC2.Close
        If intREGSS > 0 Then arrNivelMenuP(intREGSP).arrNivelMenuS = arrNivelS
        
        BREC.MoveNext
    Loop
    BREC.Close
    
    If intREGSP = 0 Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Não há dados para Imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If optVisualizacao(0).value = True Then Gera_Em_Tela
    If optVisualizacao(1).value = True Then Gera_Em_Excel

End Sub

Private Sub Gera_Em_Tela()
    
    If Gera_Dados_Tabela("I") = False Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Não foi possivel gerar os dados para impressão !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim boolTemDados As Boolean
    
    boolTemDados = False
    
    sSql = ""

    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_VIRMENUP.SGI_COUSUARIO" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUP.SGI_USUARIO" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUP.SGI_CODMENU" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUP.SGI_TEXTO" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUP.SGI_STATUS" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUP.SGI_BLOQPED" & vbCrLf

    sSql = sSql & "      ,SGI_VIRMENUS.SGI_CODMENUS" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUS.SGI_TEXTOS" & vbCrLf

    sSql = sSql & "      ,SGI_VIRMENUM.SGI_CODMENUP" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUM.SGI_TEXTOP" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUM.SGI_ACESSOP" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_VIRMENUP SGI_VIRMENUP" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUS SGI_VIRMENUS" & vbCrLf
    sSql = sSql & "      ,SGI_VIRMENUM SGI_VIRMENUM" & vbCrLf

    sSql = sSql & " Where" & vbCrLf
    
    sSql = sSql & "       SGI_VIRMENUP.SGI_FILIAL     = " & FILIAL & vbCrLf
    
    If Len(Trim(txtCODUSU.Text)) > 0 Then
        sSql = sSql & " And   SGI_VIRMENUP.SGI_COUSUARIO  = " & Trim(txtCODUSU.Text) & vbCrLf
    End If
    sSql = sSql & " And   SGI_VIRMENUP.SGI_CODIGO     = " & objRELUSUARIOS.CODIGO & vbCrLf
    
    sSql = sSql & " And   SGI_VIRMENUS.SGI_FILIAL     = SGI_VIRMENUP.SGI_FILIAL" & vbCrLf
    sSql = sSql & " And   SGI_VIRMENUS.SGI_CODIGO     = SGI_VIRMENUP.SGI_CODIGO" & vbCrLf
    sSql = sSql & " And   SGI_VIRMENUS.SGI_CODPAI     = SGI_VIRMENUP.SGI_CODMENU" & vbCrLf
    sSql = sSql & " And   SGI_VIRMENUS.SGI_COUSUARIOS = SGI_VIRMENUP.SGI_COUSUARIO" & vbCrLf
    
    sSql = sSql & " And   SGI_VIRMENUM.SGI_FILIAL     = SGI_VIRMENUS.SGI_FILIAL" & vbCrLf
    sSql = sSql & " And   SGI_VIRMENUM.SGI_CODIGO     = SGI_VIRMENUS.SGI_CODIGO" & vbCrLf
    sSql = sSql & " And   SGI_VIRMENUM.SGI_CODNETO    = SGI_VIRMENUS.SGI_CODPAI" & vbCrLf
    sSql = sSql & " And   SGI_VIRMENUM.SGI_COUSUARIOP = SGI_VIRMENUS.SGI_COUSUARIOS" & vbCrLf
    sSql = sSql & " And   SGI_VIRMENUM.SGI_CODFILHO   = SGI_VIRMENUS.SGI_CODMENUS"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then boolTemDados = True
    BREC.Close
    
    If boolTemDados = False Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Não há dados para visualizar !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    strCABEC1 = "Relatório de Usuários "
    strCABEC2 = ""
    
    Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELACEUSUARIO.rpt", Linha, 1, strCABEC1, strCABEC2, True)
    
    Call Gera_Dados_Tabela("E")
    
    
End Sub

Private Function Gera_Dados_Tabela(strOPER As String) As Boolean

On Error GoTo err_grava

     Gera_Dados_Tabela = False

     Dim I As Integer
     Dim J As Integer
     Dim K As Integer

     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
        
        objRELUSUARIOS.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
        
        For I = 1 To UBound(arrNivelMenuP)
        
            sSql = ""
            
            sSql = "Insert Into SGI_VIRMENUP (" & vbCrLf
            sSql = sSql & "                            SGI_FILIAL" & vbCrLf
            sSql = sSql & "                           ,SGI_CODIGO" & vbCrLf
            sSql = sSql & "                           ,SGI_COUSUARIO" & vbCrLf
            sSql = sSql & "                           ,SGI_USUARIO" & vbCrLf
            sSql = sSql & "                           ,SGI_CODMENU" & vbCrLf
            sSql = sSql & "                           ,SGI_TIPO" & vbCrLf
            sSql = sSql & "                           ,SGI_CIGLA" & vbCrLf
            sSql = sSql & "                           ,SGI_TEXTO" & vbCrLf
            sSql = sSql & "                           ,SGI_STATUS" & vbCrLf
            sSql = sSql & "                           ,SGI_BLOQPED" & vbCrLf
            
            sSql = sSql & "              )   Values (" & vbCrLf
            sSql = sSql & "                          " & arrNivelMenuP(I).strFILIAL & vbCrLf
            sSql = sSql & "                         ," & objRELUSUARIOS.CODIGO & vbCrLf
            sSql = sSql & "                         ," & arrNivelMenuP(I).strCODUSUARIO & vbCrLf
            sSql = sSql & "                         ,'" & Trim(arrNivelMenuP(I).strUSUARIO) & "'" & vbCrLf
            sSql = sSql & "                         ," & arrNivelMenuP(I).strCodMenu & vbCrLf
            sSql = sSql & "                         ,'" & arrNivelMenuP(I).strTIPO & "'" & vbCrLf
            sSql = sSql & "                         ,'" & arrNivelMenuP(I).strCIGLA & "'" & vbCrLf
            sSql = sSql & "                         ,'" & arrNivelMenuP(I).strTEXTO & "'" & vbCrLf
            sSql = sSql & "                         ,'" & arrNivelMenuP(I).strSTATUS & "'" & vbCrLf
            sSql = sSql & "                         ,'" & arrNivelMenuP(I).strBLOQPRDIDO & "'" & vbCrLf
            
            sSql = sSql & "                         )"
           
            BGRV.CommandText = sSql
            BGRV.Execute
            
            '' Menu S
            For J = 1 To UBound(arrNivelMenuP(I).arrNivelMenuS)
            
                 sSql = ""
                 
                 sSql = "Insert Into SGI_VIRMENUS (" & vbCrLf
                 sSql = sSql & "                            SGI_FILIAL" & vbCrLf
                 sSql = sSql & "                           ,SGI_CODIGO" & vbCrLf
                 sSql = sSql & "                           ,SGI_COUSUARIOS" & vbCrLf
                 sSql = sSql & "                           ,SGI_CODMENUS" & vbCrLf
                 sSql = sSql & "                           ,SGI_CODPAI" & vbCrLf
                 sSql = sSql & "                           ,SGI_TIPOS" & vbCrLf
                 sSql = sSql & "                           ,SGI_CIGLAS" & vbCrLf
                 sSql = sSql & "                           ,SGI_CIGLA2S" & vbCrLf
                 sSql = sSql & "                           ,SGI_TEXTOS" & vbCrLf
                 sSql = sSql & "              )   Values (" & vbCrLf
                 sSql = sSql & "                          " & arrNivelMenuP(I).arrNivelMenuS(J).strFILIAL & vbCrLf
                 sSql = sSql & "                         ," & objRELUSUARIOS.CODIGO & vbCrLf
                 sSql = sSql & "                         ," & arrNivelMenuP(I).arrNivelMenuS(J).strCODUSUARIO & vbCrLf
                 sSql = sSql & "                         ," & arrNivelMenuP(I).arrNivelMenuS(J).strCodMenu & vbCrLf
                 sSql = sSql & "                         ," & arrNivelMenuP(I).arrNivelMenuS(J).strCodPai & vbCrLf
                 sSql = sSql & "                         ,'" & arrNivelMenuP(I).arrNivelMenuS(J).strTIPO & "'" & vbCrLf
                 sSql = sSql & "                         ,'" & arrNivelMenuP(I).arrNivelMenuS(J).strCIGLA & "'" & vbCrLf
                 sSql = sSql & "                         ,'" & arrNivelMenuP(I).arrNivelMenuS(J).strCIGLA2 & "'" & vbCrLf
                 sSql = sSql & "                         ,'" & arrNivelMenuP(I).arrNivelMenuS(J).strTEXTO & "'" & vbCrLf
                 sSql = sSql & "                         )"
            
                 BGRV.CommandText = sSql
                 BGRV.Execute
                 
                 '' Menu M
                 For K = 1 To UBound(arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM)
            
                     sSql = ""
                     
                     sSql = "Insert Into SGI_VIRMENUM (" & vbCrLf
                     sSql = sSql & "                            SGI_FILIAL" & vbCrLf
                     sSql = sSql & "                           ,SGI_CODIGO" & vbCrLf
                     sSql = sSql & "                           ,SGI_COUSUARIOP" & vbCrLf
                     sSql = sSql & "                           ,SGI_CODMENUP" & vbCrLf
                     sSql = sSql & "                           ,SGI_CODFILHO" & vbCrLf
                     sSql = sSql & "                           ,SGI_CODNETO" & vbCrLf
                     sSql = sSql & "                           ,SGI_TIPOP" & vbCrLf
                     sSql = sSql & "                           ,SGI_CIGLAP" & vbCrLf
                     sSql = sSql & "                           ,SGI_TEXTOP" & vbCrLf
                     sSql = sSql & "                           ,SGI_ACESSOP" & vbCrLf
                     sSql = sSql & "              )   Values (" & vbCrLf
                     sSql = sSql & "                          " & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strFILIAL & vbCrLf
                     sSql = sSql & "                         ," & objRELUSUARIOS.CODIGO & vbCrLf
                     sSql = sSql & "                         ," & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strCODUSUARIO & vbCrLf
                     sSql = sSql & "                         ," & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strCodMenu & vbCrLf
                     sSql = sSql & "                         ," & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strCodFilho & vbCrLf
                     sSql = sSql & "                         ," & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strCodNeto & vbCrLf
                     sSql = sSql & "                         ,'" & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strTIPO & "'" & vbCrLf
                     sSql = sSql & "                         ,'" & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strCIGLA & "'" & vbCrLf
                     sSql = sSql & "                         ,'" & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strTEXTO & "'" & vbCrLf
                     sSql = sSql & "                         ,'" & arrNivelMenuP(I).arrNivelMenuS(J).arrNivelMenuM(K).strACESSO & "'" & vbCrLf
                     sSql = sSql & "                         )"
                    
                     BGRV.CommandText = sSql
                     BGRV.Execute
                 
                 Next K
            Next J
        Next I
        
     ElseIf strOPER = "E" Then
     
        sSql = ""
        sSql = "Delete From SGI_VIRMENUM " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & objRELUSUARIOS.CODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        sSql = ""
        sSql = "Delete From SGI_VIRMENUS " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & objRELUSUARIOS.CODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        sSql = ""
        sSql = "Delete From SGI_VIRMENUP " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & objRELUSUARIOS.CODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     End If

     adoBanco_Dados.CommitTrans
     Gera_Dados_Tabela = True
     
     Exit Function

err_grava:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description & "Iten : " & I, strOPER, sSql)
     Set objErro = Nothing

End Function

Private Sub Gera_Em_Excel()


On Error GoTo Handle_Error

    Dim myExcelFile             As New clsExcelFile
    Dim FileName$
    Dim boolTemDados            As Boolean
    Dim I                       As Long
    Dim K                       As Long
    Dim L                       As Long
    Dim strCODUSUARIO           As String
    Dim lngLINHA                As Long
    
    With myExcelFile
        'Create the new spreadsheet
        FileName$ = strCamRelNovo & "RELPREPARA\ACESSOUSUARIOS.xls"
        
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
        
        .SetColumnWidth 1, 3, 30
        .SetColumnWidth 4, 4, 60
        ''.SetColumnWidth 6, 6, 25
        ''.SetColumnWidth 7, 7, 60
        
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
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 1, "Usuário", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 2, "Módulos", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 3, "Funções", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 4, "Telas", 12
        
        ''.WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 2, "Cód.OP", 12
        ''.WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 3, "Emissão", 12
        ''.WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 4, "Entrega", 12
        ''.WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 5, "Razão Social", 12
        ''.WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 6, "Rótulo", 12
        ''.WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 7, "Descrição do Rótulo", 12
        
            
        strCODUSUARIO = ""
        lngLINHA = 1
        For I = 1 To UBound(arrNivelMenuP)
            lngLINHA = (lngLINHA + 1)
            
            If Trim(arrNivelMenuP(I).strCODUSUARIO) <> Trim(strCODUSUARIO) Then
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 1, Trim(arrNivelMenuP(I).strUSUARIO), 12
                
                lngLINHA = (lngLINHA + 1)
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 2, Trim(arrNivelMenuP(I).strTEXTO), 12
                strCODUSUARIO = arrNivelMenuP(I).strCODUSUARIO
            Else
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 2, Trim(arrNivelMenuP(I).strTEXTO), 12
            End If
            
            '' Mewnu S
            For K = 1 To UBound(arrNivelMenuP(I).arrNivelMenuS)
                If Trim(arrNivelMenuP(I).arrNivelMenuS(K).strCodPai) = Trim(arrNivelMenuP(I).strCodMenu) Then
                    lngLINHA = (lngLINHA + 1)
                    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 3, Trim(arrNivelMenuP(I).arrNivelMenuS(K).strTEXTO), 12
                End If
                
                '' Menu M
                For L = 1 To UBound(arrNivelMenuP(I).arrNivelMenuS(K).arrNivelMenuM)
                    If Trim(arrNivelMenuP(I).arrNivelMenuS(K).arrNivelMenuM(L).strCodFilho) = Trim(arrNivelMenuP(I).arrNivelMenuS(K).strCodMenu) Then
                        lngLINHA = (lngLINHA + 1)
                        .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 4, Trim(arrNivelMenuP(I).arrNivelMenuS(K).arrNivelMenuM(L).strTEXTO), 12
                   End If
                Next L
            Next K
        Next I
        
        'PROTECT the spreadsheet so any cells specified as LOCKED will not be
        'overwritten. Also, all cells with HIDDEN set will hide their formula.
        'PROTECT does not use a password.
        .ProtectSpreadsheet = False 'False | True
        
        'Finally, close the spreadsheet
        .CloseFile
    
    End With
    
    Exit Sub
    
Handle_Error:

    If BREC.State = 1 Then BREC.Close
    MsgBox "Número: " & Err.Number & vbCrLf & "Descrição: " & Err.Description, vbOKOnly + vbCritical, "Aviso"

End Sub
