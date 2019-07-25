VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELORDFAB 
   Caption         =   "Ordem de produção"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraVendedor 
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
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   3360
      Width           =   6375
      Begin VB.CommandButton cmdPesqVend 
         Height          =   315
         Index           =   1
         Left            =   1200
         Picture         =   "frmRELORDFAB.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdPesqVend 
         Height          =   315
         Index           =   0
         Left            =   1200
         Picture         =   "frmRELORDFAB.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCODVENDFIN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Text            =   "txtCODVENDFIN"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtCODVENDINI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Text            =   "txtCODVENDINI"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblDescVendFin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendFin"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label lblDescVendIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendIni"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Tipo do Relatório ]"
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
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   2760
      Width           =   6375
      Begin VB.OptionButton optTipRel 
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
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTipRel 
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
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
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   2160
      Width           =   6375
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
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Fechado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Parcial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Aberto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Ordem ]"
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
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   6375
      Begin VB.OptionButton optOrdem 
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optOrdem 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optOrdem 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optOrdem 
         Caption         =   "Linha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optOrdem 
         Caption         =   "Rótulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   5
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
         ForeColor       =   &H8000000D&
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
      Width           =   6375
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
         Picture         =   "frmRELORDFAB.frx":0204
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
         Picture         =   "frmRELORDFAB.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELORDFAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public FILIAL        As Integer
Public strAcesso     As String
Public lngCodUsuario As Long
Dim objBLBFunc       As Object
Dim objRELPCP        As Object
Dim objPESQPADRAO    As Object
Dim objREL           As Object
Dim strCABEC1        As String
Dim strCABEC2        As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    If optOrdem(0).Value = True Then Call ImpRelRotulo
    If optOrdem(4).Value = True Then Call ImpRelVendedor
End Sub

Private Sub cmdPesqVend_Click(Index As Integer)

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Nome"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Vendedor")
    
    If Len(Trim(varRETORNO)) > 0 Then
       If Index = 0 Then
          txtCODVENDINI.Text = varRETORNO
          lblDescVendIni.Caption = PegaDescVend(varRETORNO)
          txtCODVENDINI.SetFocus
       ElseIf Index = 1 Then
          txtCODVENDFIN.Text = varRETORNO
          lblDescVendFin.Caption = PegaDescVend(varRETORNO)
          txtCODVENDFIN.SetFocus
       End If
    End If
    
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELPCP = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
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
    Set objRELPCP = CreateObject("RELPCP.clsRELORDFAB")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELORDFAB
    objRELPCP.FILIAL = FILIAL

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    optOrdem(0).Value = True
    optTipo(0).Value = True
    optTipRel(0).Value = True
    
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Now, "DD/MM/YYYY")

    fraVendedor.Visible = False
    
    Call LimpaCamposLabel

End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
        If Not IsDate(mskDTINI.Text) Then
            MsgBox "Data inicial inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
        If Not IsDate(mskDTFIN.Text) Then
            MsgBox "Data final inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTFIN.SetFocus
            Exit Function
        End If
        
        If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
            MsgBox "Data inicial não pode ser maior que data final !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
        
        If optOrdem(4).Value = True Then
            If Len(Trim(txtCODVENDINI.Text)) > 0 And Len(Trim(txtCODVENDFIN.Text)) > 0 Then
                If CLng(txtCODVENDINI.Text) > CLng(txtCODVENDFIN.Text) Then
                    MsgBox "Vendedor inicial não pode ser maior que vendedor final !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
            End If
            If Len(Trim(txtCODVENDFIN.Text)) > 0 And Len(Trim(txtCODVENDINI.Text)) = 0 Then
                MsgBox "Vendedor inicial não pode ser vázio !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        End If
    
    ConfereCampos = True

End Function

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELORDFAB
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELORDFAB
End Sub


Private Sub ImpRelRotulo()

    Dim strNomRel As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "      SGI_ORDEMPROD.SGI_CODPROD " & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD.SGI_DATAORDEM " & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD.SGI_STATUS " & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD.SGI_CODIGO " & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD.SGI_CODPED " & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD.SGI_QTDE " & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD.SGI_SALDO " & vbCrLf
    
    sSql = sSql & "     ,SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    sSql = sSql & "     ,SGI_CADPEDVENDH.SGI_CODIGO " & vbCrLf
    sSql = sSql & "     ,SGI_CADCLIENTE.SGI_RAZAOSOC " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    
    sSql = sSql & "      SGI_CADCLIENTE SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDH SGI_CADPEDVENDH " & vbCrLf
    sSql = sSql & "     ,SGI_CADPRODUTO SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & "     ,SGI_CADVENDEDOR SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & "     ,SGI_ORDEMPROD SGI_ORDEMPROD " & vbCrLf
        
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       SGI_ORDEMPROD.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And (SGI_ORDEMPROD.SGI_DATAORDEM Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "')" & vbCrLf
    
    If optTipo(0).Value = True Then
        sSql = sSql & "   And SGI_ORDEMPROD.SGI_STATUS = 0" & vbCrLf
    ElseIf optTipo(1).Value = True Then
        sSql = sSql & "   And SGI_ORDEMPROD.SGI_STATUS = 1" & vbCrLf
    ElseIf optTipo(2).Value = True Then
        sSql = sSql & "   And SGI_ORDEMPROD.SGI_STATUS = 2" & vbCrLf
    End If
    
    sSql = sSql & "   And SGI_ORDEMPROD.SGI_FILIAL = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD.SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD.SGI_FILIAL = SGI_CADPEDVENDH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD.SGI_CODPED = SGI_CADPEDVENDH.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_FILIAL = SGI_CADVENDEDOR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODVEND = SGI_CADVENDEDOR.SGI_CODIGO " & vbCrLf
    
    If optOrdem(0).Value = True Then
        sSql = sSql & "   Order By SGI_CODPROD"
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Relatório de ordem de produção por ordem de Rótulos"
    
    If optTipo(0).Value = True Then
       strCABEC1 = strCABEC1 & " [Aberto] "
    ElseIf optTipo(1).Value = True Then
       strCABEC1 = strCABEC1 & " [Parcial] "
    ElseIf optTipo(2).Value = True Then
       strCABEC1 = strCABEC1 & " [Fechado] "
    ElseIf optTipo(3).Value = True Then
       strCABEC1 = strCABEC1 & " [Todos] "
    End If
        
    If optTipRel(0).Value = True Then
        strCABEC1 = strCABEC1 & " - Análitico"
    ElseIf optTipRel(1).Value = True Then
        strCABEC1 = strCABEC1 & " - Sintético"
    End If
    
    If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text
    
    strNomRel = "RELPCPOFRAA.rpt"
    
    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, False)
    End If
    
End Sub



Private Sub ImpRelVendedor()

    Dim strNomRel As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD.SGI_DATAORDEM " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_CODPED " & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODVEND " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_DATAPED " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODCLI " & vbCrLf
    
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_CODPROD " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_QTDE " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_SALDO " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_DATAORDEM " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_DATENTREGA " & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD.SGI_STATUS " & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODLINPROD " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODLINPROD " & vbCrLf
    
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_CODLIN " & vbCrLf
    
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & "     , SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & "     , SGI_CADPEDVENDH SGI_CADPEDVENDH " & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & "     , SGI_CADVENDEDOR SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & "     , SGI_ORDEMPROD SGI_ORDEMPROD " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And (SGI_ORDEMPROD.SGI_DATAORDEM Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "')" & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD.SGI_FILIAL = SGI_CADPEDVENDH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD.SGI_CODPED = SGI_CADPEDVENDH.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD.SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD.SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL = SGI_CADLINHAPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN " & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_FILIAL  = SGI_CADVENDEDOR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODVEND = SGI_CADVENDEDOR.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_FILIAL  = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
    
    If Len(Trim(txtCODVENDINI.Text)) > 0 And Len(Trim(txtCODVENDFIN.Text)) > 0 Then
        sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODVEND >= " & txtCODVENDINI.Text & " And SGI_CADPEDVENDH.SGI_CODVEND <= " & txtCODVENDFIN.Text
    ElseIf Len(Trim(txtCODVENDINI.Text)) > 0 And Len(Trim(txtCODVENDFIN.Text)) = 0 Then
        sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODVEND = " & txtCODVENDINI.Text
    End If
    
    If optTipo(0).Value = True Then     '' Aberto
        sSql = sSql & "   And SGI_ORDEMPROD.SGI_STATUS = 0"
    ElseIf optTipo(1).Value = True Then '' Parcial
        sSql = sSql & "   And SGI_ORDEMPROD.SGI_STATUS = 1"
    ElseIf optTipo(2).Value = True Then '' Fechado
        sSql = sSql & "   And SGI_ORDEMPROD.SGI_STATUS = 2"
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Relatório de ordem de produção por ordem de Vendedor"
    
    If optTipo(0).Value = True Then
       strCABEC1 = strCABEC1 & " [Aberto] "
    ElseIf optTipo(1).Value = True Then
       strCABEC1 = strCABEC1 & " [Parcial] "
    ElseIf optTipo(2).Value = True Then
       strCABEC1 = strCABEC1 & " [Fechado] "
    ElseIf optTipo(3).Value = True Then
       strCABEC1 = strCABEC1 & " [Todos] "
    End If
        
    If optTipRel(0).Value = True Then
        strCABEC1 = strCABEC1 & " - Análitico"
    ElseIf optTipRel(1).Value = True Then
        strCABEC1 = strCABEC1 & " - Sintético"
    End If
    
    If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text
    
    strNomRel = "REOPVEND.RPT"
    
    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If


End Sub

Private Sub optOrdem_Click(Index As Integer)
    fraVendedor.Visible = True
    
    txtCODVENDINI.Text = ""
    txtCODVENDFIN.Text = ""
    
    lblDescVendIni.Caption = ""
    lblDescVendFin.Caption = ""
    
    If Index = 4 Then
        fraVendedor.Visible = True
    End If
End Sub

Private Sub txtCODVENDFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODVENDFIN.Name, frmRELORDFAB
End Sub

Private Sub txtCODVENDFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODVENDFIN.Text
End Sub

Private Sub txtCODVENDFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODVENDFIN.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCODVENDFIN.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVENDFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblDescVendFin.Caption = PegaDescVend(txtCODVENDFIN.Text)
    If Len(Trim(lblDescVendFin.Caption)) = 0 Then
        MsgBox "Vendedor não cadastrado !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODVENDFIN.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtCODVENDINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODVENDINI.Name, frmRELORDFAB
End Sub

Private Sub txtCODVENDINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODVENDINI.Text
End Sub

Private Sub LimpaCamposLabel()
    lblDescVendIni.Caption = ""
    lblDescVendFin.Caption = ""
End Sub

Private Function PegaDescVend(strCodVend As String) As String

    PegaDescVend = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "    And SGI_CODIGO = " & Trim(strCodVend)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescVend = BREC!SGI_DESCRICAO
    BREC.Close
    
End Function

Private Sub txtCODVENDINI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODVENDINI.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCODVENDINI.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVENDINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblDescVendIni.Caption = PegaDescVend(txtCODVENDINI.Text)
    If Len(Trim(lblDescVendIni.Caption)) = 0 Then
        MsgBox "Vendedor não cadastrado !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODVENDINI.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub
