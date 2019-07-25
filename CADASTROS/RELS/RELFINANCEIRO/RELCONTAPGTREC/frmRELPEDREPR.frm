VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELPEDREPR 
   Caption         =   "Relatório de Vendas de Representantes"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   12795
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   "[ Filtra Data ]"
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
      TabIndex        =   24
      Top             =   960
      Width           =   2655
      Begin VB.OptionButton optSimNao 
         Caption         =   "Sim"
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
         Left            =   600
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optSimNao 
         Caption         =   "Não"
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
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
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
      Left            =   8880
      TabIndex        =   20
      Top             =   960
      Width           =   3855
      Begin VB.OptionButton optTipo 
         Caption         =   "Pago"
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
         Left            =   1560
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Em Aberto"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTipo 
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
         Index           =   0
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   5400
      TabIndex        =   15
      Top             =   1560
      Width           =   7335
      Begin VB.TextBox txtCodRepr 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Text            =   "txtCodRepr"
         Top             =   255
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   2280
         Picture         =   "frmRELPEDREPR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Representante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label lblDescRepr 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescRepr"
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   4575
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
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   5400
      TabIndex        =   12
      Top             =   960
      Width           =   3375
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   14
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Agrupamento ]"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   960
      Width           =   2655
      Begin VB.OptionButton optAgrup 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optAgrup 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optAgrup 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   5295
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
      TabIndex        =   4
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
         Picture         =   "frmRELPEDREPR.frx":0102
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
         Picture         =   "frmRELPEDREPR.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELPEDREPR"
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
Dim objRELPEDREPR       As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call ImpRel
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command4_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADREPRESENT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Representante")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCodRepr.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADREPRESENT", varRETORNO, lblDescRepr)
    End If
    txtCodRepr.SetFocus

End Sub

Private Sub Form_Load()


    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELPEDREPR = CreateObject("RELCOMERCIAL.clsRELPEDREPR")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELPEDREPR
    objRELPEDREPR.FILIAL = FILIAL

    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    optAgrup(0).Value = True
    optRELCOTAANSIN(0).Value = True
    optTipo(1).Value = True
    optSimNao(0).Value = True

    Call LimpaCamposLabel

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Destroy_Objetos()
    Set objBLBFunc = Nothing
    Set objRELPEDREPR = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objetos
End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
    If optSimNao(1).Value = True Then
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
    End If
    
    ConfereCampos = True

End Function

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELPEDREPR
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELPEDREPR
End Sub

Private Sub ImpRel()

    Dim strNomRel   As String
    Dim strTIPO     As String
    Dim strAGRUPA   As String
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "        SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "       ,SGI_CADPEDVENDH.SGI_CODREPRES" & vbCrLf
    sSql = sSql & "       ,SGI_CADPEDVENDH.SGI_CODCLI" & vbCrLf
    
    sSql = sSql & "       ,SGI_VALESHEAD.SGI_CODVALE" & vbCrLf
    sSql = sSql & "       ,SGI_VALESHEAD.SGI_DTVALE" & vbCrLf
    sSql = sSql & "       ,SGI_VALESHEAD.SGI_VALOR" & vbCrLf
    sSql = sSql & "       ,SGI_VALESHEAD.SGI_PORCDESC" & vbCrLf
    sSql = sSql & "       ,SGI_VALESHEAD.SGI_VLDESC" & vbCrLf
    sSql = sSql & "       ,SGI_VALESHEAD.SGI_VLPAGO" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "        SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "       ,SGI_CADPEDVENDH SGI_CADPEDVENDH" & vbCrLf
    sSql = sSql & "       ,SGI_CADREPRESENT SGI_CADREPRESENT" & vbCrLf
    sSql = sSql & "       ,SGI_VALESHEAD SGI_VALESHEAD" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_VALESHEAD.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If optSimNao(1).Value = True Then
        sSql = sSql & "   And SGI_VALESHEAD.SGI_DTVALE Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    End If
    
    sSql = sSql & "   And SGI_VALESHEAD.SGI_FILIAL      = SGI_CADPEDVENDH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_VALESHEAD.SGI_CODVALE     = SGI_CADPEDVENDH.SGI_CODIGO" & vbCrLf
    
    If optTipo(1).Value = True Then
        sSql = sSql & "   And SGI_VALESHEAD.SGI_STATUS  = 'A'" & vbCrLf
    ElseIf optTipo(2).Value = True Then
        sSql = sSql & "   And SGI_VALESHEAD.SGI_STATUS  = 'B'" & vbCrLf
    End If
    
    If Len(Trim(txtCodRepr.Text)) > 0 Then
        sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODREPRES = " & Trim(txtCodRepr.Text) & vbCrLf
    End If
    
    
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_FILIAL    = SGI_CADREPRESENT.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODREPRES = SGI_CADREPRESENT.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_FILIAL    = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODCLI    = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    
    
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If optAgrup(0).Value = True Then strAGRUPA = "Dia"
    If optAgrup(1).Value = True Then strAGRUPA = "Mês"
    If optAgrup(2).Value = True Then strAGRUPA = "Ano"
    
    If optTipo(0).Value = True Then strAGRUPA = strAGRUPA & "/Todos"
    If optTipo(1).Value = True Then strAGRUPA = strAGRUPA & "/Em Aberto"
    If optTipo(2).Value = True Then strAGRUPA = strAGRUPA & "/Pagos"
    
    
    strCABEC1 = "Relatório de vendas de representantes agrupado por " & strAGRUPA & " "
    
    If optRELCOTAANSIN(0).Value = True Then
        If optAgrup(0).Value = True Then strNomRel = "RELCOMREPR01.rpt"
        If optAgrup(1).Value = True Then strNomRel = "RELCOMREPR02.rpt"
        If optAgrup(2).Value = True Then strNomRel = "RELCOMREPR03.rpt"
        strCABEC2 = "Analitico"
    ElseIf optRELCOTAANSIN(1).Value = True Then
        If optAgrup(0).Value = True Then strNomRel = ""
        If optAgrup(1).Value = True Then strNomRel = ""
        If optAgrup(2).Value = True Then strNomRel = ""
        strCABEC2 = "Sintético"
    End If
    
    If optSimNao(1).Value = True Then
        If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then
            strCABEC2 = strCABEC2 & " - Dia : " & mskDTINI.Text
        ElseIf CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then
            strCABEC2 = strCABEC2 & " - no periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
        End If
    End If
    
    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If
    
    Exit Sub
    
End Sub

Private Sub txtCodRepr_GotFocus()
    objBLBFunc.SelecionaCampos txtCodRepr.Name, frmRELPEDREPR
End Sub

Private Sub txtCodRepr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then Call Command4_Click
End Sub

Private Sub txtCodRepr_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodRepr.Text
End Sub

Private Sub txtCodRepr_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCodRepr.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodRepr.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCodRepr.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADREPRESENT", txtCodRepr.Text, lblDescRepr)
    If Len(Trim(lblDescRepr.Caption)) = 0 Then
       txtCodRepr.Text = ""
       Cancel = True
    End If

End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCodigo As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(strCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCodigo)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub LimpaCamposLabel()
    lblDescRepr.Caption = ""
End Sub

