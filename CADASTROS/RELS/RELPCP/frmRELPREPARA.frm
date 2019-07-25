VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmRELPREPARA 
   Caption         =   "Relatório de Preparação"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   10620
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   0
      TabIndex        =   30
      Top             =   3960
      Width           =   10575
      Begin MSComctlLib.ProgressBar prgPREP 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Max             =   100
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[ Tipo de Relatório ]"
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
      TabIndex        =   26
      Top             =   1560
      Width           =   5055
      Begin VB.OptionButton optAprRel 
         Caption         =   "Arquivo do EXCEL"
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
         Left            =   3000
         TabIndex        =   29
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optAprRel 
         Caption         =   "Arquivo TXT"
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
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAprRel 
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
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
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
      TabIndex        =   21
      Top             =   3240
      Width           =   10575
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   1320
         Picture         =   "frmRELPREPARA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   22
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
         TabIndex        =   24
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Familia de Produtos  ]"
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
      Height          =   1095
      Left            =   3240
      TabIndex        =   19
      Top             =   2160
      Width           =   7335
      Begin VB.ListBox lstFamProd 
         Height          =   735
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   240
         Width           =   6855
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
      Height          =   1095
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   3135
      Begin VB.OptionButton optTipo 
         Caption         =   "Somente Produtos [ Normais ]"
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
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Somente Produtos [ Rótulos ]"
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
         TabIndex        =   17
         Top             =   480
         Width           =   2895
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Conjugados ]"
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
      Top             =   1560
      Width           =   5415
      Begin VB.OptionButton optConj 
         Caption         =   "Pedidos + OP"
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
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optConj 
         Caption         =   "Somente Pedidos"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optConj 
         Caption         =   "Somente OP"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Filial ]"
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
      TabIndex        =   13
      Top             =   960
      Width           =   5055
      Begin VB.OptionButton optFilial 
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
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optFilial 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFilial 
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
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   10
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
         TabIndex        =   12
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10575
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
         Picture         =   "frmRELPREPARA.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmRELPREPARA.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELPREPARA"
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
Dim objRELPREPARA    As Object
Dim objPESQPADRAO    As Object
Dim objREL                  As Object

Dim strCABEC1               As String
Dim strCABEC2               As String

Dim lngPORC                 As Long
Dim arrDADOSNOVALATA()      As String
Dim arrDADOSSTEEL()         As String


Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    If optAprRel(1).value = True Then
        '' Em arquivo TXT
        If optTipo(0).value = True Then Call ImpRel
        If optTipo(1).value = True Then Call ImpRel2
        If optTipo(2).value = True Then Call ImpRel3
    ElseIf optAprRel(2).value = True Then
        '' Em Arquivo EXCEL
        If optTipo(0).value = True Then Call ImpRelEXCEL
    End If
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
    If Len(Trim(lblDescCliente.Caption)) = 0 Then txtCIDCLIE.Text = ""
    
    txtCIDCLIE.SetFocus

    Exit Sub
    
Err_Command1_Click:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "I", "Função : Command1_Click()", Me.Name, "Command1_Click()", strCAMARQERRO)

End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELPREPARA = CreateObject("RELPCP.clsRELPREPARA")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    objRELPREPARA.FILIAL = FILIAL

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7) & "RELPREPARA\"
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Now, "DD/MM/YYYY")

    optFilial(0).value = True
    prgPREP.Min = 0
    Frame9.Visible = False
    optConj(0).value = True
    optTipo(0).value = True
    
    Call LimpaListBox
    Call PopLSTBoxFam
    Call LimpaCamposLabel
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELPREPARA = Nothing
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
        
    ConfereCampos = True

End Function


Private Sub ImpRel()

On Error GoTo Err_Exporta

    Dim strNOMFILIAL    As String
    Dim strNomRel       As String
    Dim boolTemDados    As Boolean
    Dim lngQTDFOLHAS    As Long
    Dim dblPERDAPROC    As Double
    Dim lngQTDEFOLHAS   As Long
    Dim strESMALTE      As String
    Dim arrCORES        As Variant
    Dim intQTDCORES     As Integer
    Dim lngQTDREGS      As Long
    Dim lngQTDTOTAL     As Long
    Dim intCODFECHA     As Integer
    Dim lngQTDEOP       As Long
    Dim lngQTDETOTOP    As Long
    Dim strSTATUSOP     As String
    Dim lngCODOP        As Long
    
    Dim strCAMPO01      As String
    Dim strCAMPO02      As String
    Dim strCAMPO03      As String
    Dim strCAMPO04      As String
    
    Dim strDADOS01      As String
    Dim strDADOS02      As String
    Dim strDADOS03      As String
    Dim strDADOS04      As String
    
    Dim strVERNIZ01     As String
    Dim strVERNIZ02     As String
    Dim strVERNIZACAB   As String
    Dim strNECKIN       As String
    Dim strESTADO       As String
    Dim strFECHAGRAF    As String
    Dim strVERNCORPO    As String
    Dim strVERNTAMPA    As String
    Dim strVERNFUNDO    As String
    Dim strVERNARGOLA   As String
    Dim strOBSOP        As String
    Dim strTOTFAT       As String
    Dim strSTATUS2      As String
    
    
    Frame9.Visible = True
    If optFilial(1).value = True Then
        strNomRel = "RELPREPARA01_STEEL.TXT"
    ElseIf optFilial(0).value = True Then
        strNomRel = "RELPREPARA01_NOVA.TXT"
    End If
    
    boolTemDados = True
    
    strNOMFILIAL = ""
    If optFilial(1).value = True Then strNOMFILIAL = "_STEEL"
    
    sSql = ""
    
    If optConj(0).value = True Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODIGO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATENTREGA" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATAORDEM" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPED" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATAORDEM" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_QTDE" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_NOMEVEND" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPROD" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_STATUS" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FECHTPFU" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_OBSOP" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_QTDEPORFOLHA" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_QTDCORPSPADRAOSN" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechSoldaAgrafado" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernCorpo" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernTampa" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernFundo" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernArgola" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_QTDECORPOS" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_PERDPROC" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADESPPROD.SGI_DESCRICAO As SGI_DESCESP" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ESTNORM" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CIDNORM" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADTIPPROD.SGI_DESCRICAO As SGI_DESCTIPO" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_ESTENTRE" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " SGI_CADPEDVENDH" & strNOMFILIAL & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & " SGI_ORDEMPROD" & strNOMFILIAL & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADESPPROD SGI_CADESPPROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD SGI_CADTIPPROD" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATENTREGA Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'"
        
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
        
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL    = SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPED    = SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADESPPROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODESPECIE = SGI_CADESPPROD.SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADTIPPROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODTIPO    = SGI_CADTIPPROD.SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO"
    
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       
       Open strCamRelNovo & strNomRel For Output As #1
       
       strCAMPO01 = "VENDEDOR" & vbTab & _
                    "CLIENTE" & vbTab & _
                    "ROTULO" & vbTab & _
                    "Pedido de Venda" & vbTab & _
                    "Status OP" & vbTab & _
                    "DESCRICAO" & vbTab & _
                    "OP" & vbTab & _
                    "COD.CAPACIDADE" & vbTab & _
                    "CAPACIDADE" & vbTab & _
                    "TIPO" & vbTab & _
                    "DATA ENTRADA" & vbTab & _
                    "DATA PREPARACAO" & vbTab & _
                    "DATA LITOGRAFIA" & vbTab & _
                    "DATA MONSTAGEM" & vbTab & _
                    "DATA ENTREGA" & vbTab & _
                    "QUANTIDADE PEDIDO" & vbTab & _
                    "QTDE FOLHAS" & vbTab & _
                    "QTDE POR FOLHA" & vbTab & _
                    "VERNIZ.INT 01" & vbTab & _
                    "VERNIZ.INT 02"
        
        strCAMPO02 = "ESMALTE" & vbTab & _
                     "REVESTIMENTO" & vbTab & _
                     "VERNIZ.ACABAMENTO" & vbTab & _
                     "PREPARACAO" & vbTab & _
                     "1a.COR" & vbTab & _
                     "2a.COR" & vbTab & _
                     "3a.COR" & vbTab & _
                     "4a.COR" & vbTab & _
                     "5a.COR" & vbTab & _
                     "6a.COR" & vbTab & _
                     "7a.COR" & vbTab & _
                     "8a.COR" & vbTab & _
                     "LITOGRAFIA" & vbTab & _
                     "FECHAMENTO" & vbTab & _
                     "Neck IN"
                     
        strCAMPO03 = "Qtde.OP" & vbTab & _
                     "Qtde.Faturada" & vbTab & _
                     "Data Faturada" & vbTab & _
                     "Nota Fiscal" & vbTab & _
                     "Saldo.OP" & vbTab & _
                     "OP.Faturada" & vbTab & _
                     "Estado.Entrega" & vbTab & _
                     "Cidade Entrega" & vbTab & _
                     "OBS. da OP" & vbTab & _
                     "Fechamento" & vbTab & _
                     "Verniz CP" & vbTab & _
                     "Verniz TP" & vbTab & _
                     "Verniz FD" & vbTab & _
                     "Verniz ARG" & vbTab & _
                     "Total.Fat"
                     
       strCAMPO04 = "Status Pela Tolerancia"
       
       Print #1, strCAMPO01 & vbTab & _
                 strCAMPO02 & vbTab & _
                 strCAMPO03 & vbTab & _
                 strCAMPO04
                 
       lngQTDREGS = 0
       lngQTDTOTAL = 0
       prgPREP.Min = lngQTDREGS
       Do While Not BREC.EOF()
          lngQTDREGS = (lngQTDREGS + 1)
          BREC.MoveNext
       Loop
       If lngQTDREGS > 0 Then
          prgPREP.Max = lngQTDREGS
          lngQTDTOTAL = lngQTDREGS
      End If
        
        
        
       BREC.MoveFirst
       lngQTDREGS = 0
       Do While Not BREC.EOF()
       
            lngQTDREGS = (lngQTDREGS + 1)
            prgPREP.value = lngQTDREGS
            
            strOBSOP = ""
            If Not IsNull(BREC!SGI_OBSOP) Then strOBSOP = Trim(Replace(BREC!SGI_OBSOP, vbCrLf, " , "))
            
            If BREC!SGI_STATUS = 0 Then strSTATUSOP = "ABERTO"
            If BREC!SGI_STATUS = 1 Then strSTATUSOP = "Fat.Parcial"
            If BREC!SGI_STATUS = 2 Then strSTATUSOP = "Finalizada"
            If BREC!SGI_STATUS = 3 Then strSTATUSOP = "Bloqueada"
            
            strFECHAGRAF = ""
            If BREC!SGI_FechSoldaAgrafado = 0 Then strFECHAGRAF = "SOLDA"
            If BREC!SGI_FechSoldaAgrafado = 1 Then strFECHAGRAF = "AGRAFADO"
            If BREC!SGI_FechSoldaAgrafado = 2 Then strFECHAGRAF = "REPUXO"
            
            strVERNCORPO = ""
            If BREC!SGI_VernCorpo = 1 Then strVERNCORPO = "VEX"
            If BREC!SGI_VernCorpo = 2 Then strVERNCORPO = "VZ"
            If BREC!SGI_VernCorpo = 3 Then strVERNCORPO = "NAT"
            If BREC!SGI_VernCorpo = 4 Then strVERNCORPO = "VI"
            
            strVERNTAMPA = ""
            If BREC!SGI_VernTampa = 1 Then strVERNTAMPA = "VEX"
            If BREC!SGI_VernTampa = 2 Then strVERNTAMPA = "VZ"
            If BREC!SGI_VernTampa = 3 Then strVERNTAMPA = "NAT"
            If BREC!SGI_VernTampa = 4 Then strVERNTAMPA = "VI"
            
            strVERNFUNDO = ""
            If BREC!SGI_VernFundo = 1 Then strVERNFUNDO = "VEX"
            If BREC!SGI_VernFundo = 2 Then strVERNFUNDO = "VZ"
            If BREC!SGI_VernFundo = 3 Then strVERNFUNDO = "NAT"
            If BREC!SGI_VernFundo = 4 Then strVERNFUNDO = "VI"
            
            strVERNARGOLA = ""
            If BREC!SGI_VernArgola = 1 Then strVERNARGOLA = "VEX"
            If BREC!SGI_VernArgola = 2 Then strVERNARGOLA = "VZ"
            If BREC!SGI_VernArgola = 3 Then strVERNARGOLA = "NAT"
            If BREC!SGI_VernArgola = 4 Then strVERNARGOLA = "VI"
            
            '' Pegava o Estado de Entrega
            ''strESTADO = Pega_Estado(BREC!SGI_ESTENTRE)
            strESTADO = Pega_Estado(BREC!SGI_ESTNORM)
            
            lngQTDFOLHAS = 0
            dblPERDAPROC = 0
            lngQTDEFOLHAS = 0
            If BREC!SGI_QTDCORPSPADRAOSN = 0 Then
               If Not IsNull(BREC!SGI_QTDEPORFOLHA) Then lngQTDFOLHAS = BREC!SGI_QTDEPORFOLHA
               dblPERDAPROC = 1.05
            ElseIf BREC!SGI_QTDCORPSPADRAOSN = 1 Then
               If Not IsNull(BREC!SGI_QTDECORPOS) Then lngQTDFOLHAS = BREC!SGI_QTDECORPOS
               If Not IsNull(BREC!SGI_PERDPROC) Then dblPERDAPROC = BREC!SGI_PERDPROC
            End If
            If lngQTDFOLHAS > 0 Then lngQTDEFOLHAS = ((BREC!SGI_QTDE * dblPERDAPROC) / lngQTDFOLHAS)
       
            '' Verniz 01
            strVERNIZ01 = ""
            
            sSql = ""
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_VERNIZPROD VER" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO PRD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then strVERNIZ01 = BREC2!SGI_DESCRICAO
            BREC2.Close
            
            '' Verniz 02
            strVERNIZ02 = ""
            
            sSql = ""
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_VERNIZPROD02 VER" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO   PRD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then strVERNIZ02 = BREC2!SGI_DESCRICAO
            BREC2.Close
            
            '' Esmalte
            strESMALTE = ""
            
            sSql = ""
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_ESMALTEPROD  VER" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO   PRD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then strESMALTE = BREC2!SGI_DESCRICAO
            BREC2.Close
            
            '' Verniz Acabamento
            strVERNIZACAB = ""
            
            sSql = ""
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_VERNIZPRODACAB  VER" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO      PRD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then strVERNIZACAB = BREC2!SGI_DESCRICAO
            BREC2.Close
            
            
            '' --------------------------------
            '' Pega Cores
            ReDim arrCORES(1 To 8) As String
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PROD.SGI_DESCRICAO" & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CORESPROD CORES" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       CORES.SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And CORES.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And CORES.SGI_FILIAL    = PROD.SGI_FILIAL"
            sSql = sSql & "   And CORES.SGI_CODCOR    = PROD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then
                intQTDCORES = 1
                Do While Not BREC2.EOF()
                   arrCORES(intQTDCORES) = BREC2!SGI_DESCRICAO
                   intQTDCORES = (intQTDCORES + 1)
                   BREC2.MoveNext
                Loop
            End If
            BREC2.Close
            '' ----------------------------------
            
            intCODFECHA = 0
            If Not IsNull(BREC!SGI_FECHTPFU) Then intCODFECHA = BREC!SGI_FECHTPFU
       
            strNECKIN = "NÃO"
            If BREC!SGI_NECKIN = 1 Then strNECKIN = "SIM"
       
            strDADOS01 = BREC!SGI_NOMEVEND & vbTab & _
                         BREC!SGI_RAZAOSOC & vbTab & _
                         BREC!SGI_CODPROD & vbTab & _
                         BREC!SGI_CODPED & vbTab & _
                         strSTATUSOP & vbTab & _
                         Replace(Replace(BREC!SGI_DESCRICAO, "Ç", "C"), "Ã", "A") & vbTab & _
                         BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_CODLIN & vbTab & _
                         Trim(BREC!SGI_DESCRI) & "." & vbTab & _
                         BREC!SGI_DESCTIPO & vbTab & _
                         Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                         Format((BREC!SGI_DATENTREGA - 7), "DD/MM/YYYY") & vbTab & _
                         Format((BREC!SGI_DATENTREGA - 3), "DD/MM/YYYY") & vbTab & _
                         Format((BREC!SGI_DATENTREGA - 2), "DD/MM/YYYY") & vbTab & _
                         Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                         BREC!SGI_QTDE & vbTab & _
                         lngQTDEFOLHAS & vbTab & _
                         lngQTDFOLHAS & vbTab & _
                         strVERNIZ01 & vbTab & _
                         strVERNIZ02
       
            strDADOS02 = strESMALTE & vbTab & _
                         BREC!SGI_DESCESP & vbTab & _
                         strVERNIZACAB & vbTab & _
                         "" & vbTab & _
                         arrCORES(1) & vbTab & _
                         arrCORES(2) & vbTab & _
                         arrCORES(3) & vbTab & _
                         arrCORES(4) & vbTab & _
                         arrCORES(5) & vbTab & _
                         arrCORES(6) & vbTab & _
                         arrCORES(7) & vbTab & _
                         arrCORES(8) & vbTab & _
                         "" & vbTab & _
                         Fechamento(intCODFECHA) & vbTab & _
                         strNECKIN
       

            '' =======================================
            '' Pegando Faturamentos já realizado
            strTOTFAT = SomaFaturamento(Str(BREC!SGI_IDPRODUTO), Str(BREC!SGI_CODIGO), Trim(strNOMFILIAL))
            
            If BREC!SGI_STATUS = 2 Then
                strSTATUS2 = "Finalizada"
            Else
                strSTATUS2 = FechamentoOP(BREC!SGI_QTDE, CLng(strTOTFAT))
            End If
            
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       ITEN.*" & vbCrLf
            sSql = sSql & "      ,CABE.*" & vbCrLf
            sSql = sSql & "      ,ORDP.*" & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADORDCONFI" & strNOMFILIAL & " ITEN" & vbCrLf
            sSql = sSql & "     , SGI_CADORDCONFH" & strNOMFILIAL & " CABE" & vbCrLf
            sSql = sSql & "     , SGI_ORDEMPROD" & strNOMFILIAL & " ORDP" & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       ITEN.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And ITEN.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And ITEN.SGI_CODORDPROD = " & BREC!SGI_CODIGO & vbCrLf
            sSql = sSql & "   And ITEN.SGI_FILIAL     = CABE.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And ITEN.SGI_CODCONF    = CABE.SGI_CODCONF" & vbCrLf
            sSql = sSql & "   And ITEN.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And ITEN.SGI_CODORDPROD = ORDP.SGI_CODIGO" & vbCrLf
            
            sSql = sSql & "Order By CABE.SGI_DATACONF"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            
            lngQTDEOP = BREC!SGI_QTDE '' Qtde OP
            lngCODOP = 0
            If Not BREC2.EOF() Then
                Do While Not BREC2.EOF()
                    
                    lngQTDEOP = (lngQTDEOP - BREC2!SGI_QTDREAL)
                    
                    strDADOS03 = BREC2!SGI_QTDREAL & vbTab & _
                                 Format(BREC2!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                                 BREC2!SGI_CODFATURA & vbTab & _
                                 lngQTDEOP
                    
                    If lngCODOP <> BREC2!SGI_CODORDPROD Then
                        
                        lngCODOP = BREC2!SGI_CODORDPROD
                        
                        Print #1, strDADOS01 & vbTab & _
                                  strDADOS02 & vbTab & _
                                  BREC2!SGI_QTDE & vbTab & _
                                  strDADOS03 & vbTab & _
                                  BREC2!SGI_CODORDPROD & vbTab & _
                                  strESTADO & vbTab & _
                                  BREC!SGI_CIDNORM & vbTab & _
                                  strOBSOP & vbTab & _
                                  strFECHAGRAF & vbTab & _
                                  strVERNCORPO & vbTab & _
                                  strVERNTAMPA & vbTab & _
                                  strVERNFUNDO & vbTab & _
                                  strVERNARGOLA & vbTab & _
                                  strTOTFAT & vbTab & _
                                  strSTATUS2
                                  
                    Else
                        strDADOS01 = BREC!SGI_NOMEVEND & vbTab & _
                                     BREC!SGI_RAZAOSOC & vbTab & _
                                     BREC!SGI_CODPROD & vbTab & _
                                     BREC!SGI_CODPED & vbTab & _
                                     strSTATUSOP & vbTab & _
                                     Replace(Replace(BREC!SGI_DESCRICAO, "Ç", "C"), "Ã", "A") & vbTab & _
                                     BREC!SGI_CODIGO & vbTab & _
                                     BREC!SGI_CODLIN & vbTab & _
                                     Trim(BREC!SGI_DESCRI) & "." & vbTab & _
                                     BREC!SGI_DESCTIPO & vbTab & _
                                     Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                                     Format((BREC!SGI_DATENTREGA - 7), "DD/MM/YYYY") & vbTab & _
                                     Format((BREC!SGI_DATENTREGA - 3), "DD/MM/YYYY") & vbTab & _
                                     Format((BREC!SGI_DATENTREGA - 2), "DD/MM/YYYY") & vbTab & _
                                     Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     ""
        
                        strDADOS02 = "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     ""
                    
                        Print #1, strDADOS01 & vbTab & _
                                  strDADOS02 & vbTab & _
                                  "" & vbTab & _
                                  strDADOS03 & vbTab & _
                                  "" & vbTab & _
                                  strESTADO & vbTab & _
                                  BREC!SGI_CIDNORM & vbTab & _
                                  strOBSOP & vbTab & _
                                  strFECHAGRAF & vbTab & _
                                  strVERNCORPO & vbTab & _
                                  strVERNTAMPA & vbTab & _
                                  strVERNFUNDO & vbTab & _
                                  strVERNARGOLA & vbTab & _
                                  strTOTFAT & vbTab & _
                                  strSTATUS2
                    
                    End If
                    BREC2.MoveNext
                Loop
            Else
                
                strDADOS03 = "" & vbTab & _
                             "" & vbTab & _
                             "" & vbTab & _
                             BREC!SGI_QTDE & vbTab & _
                             "" & vbTab & _
                             strESTADO & vbTab & _
                             BREC!SGI_CIDNORM & vbTab & _
                             strOBSOP & vbTab & _
                             strFECHAGRAF & vbTab & _
                             strVERNCORPO & vbTab & _
                             strVERNTAMPA & vbTab & _
                             strVERNFUNDO & vbTab & _
                             strVERNARGOLA & vbTab & _
                             strTOTFAT & vbTab & _
                             strSTATUS2
                
                Print #1, strDADOS01 & vbTab & _
                          strDADOS02 & vbTab & _
                          BREC!SGI_QTDE & vbTab & _
                          strDADOS03
            End If
            BREC2.Close
            '' =======================================
         
         BREC.MoveNext
       Loop
       Close #1
    
       MsgBox "Arquivo Gerado com Exito !!!", vbOKOnly + vbInformation, "Aviso"
    
    Else
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       boolTemDados = False
    End If
    BREC.Close
    
    Frame9.Visible = False
    
    Exit Sub

Err_Exporta:
    
    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    Close #1
    
    
    If Err.Number = 70 Then
        MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
               "Erro Descr  : " & "Não pode inportar pois o Arguivo TXT eta aberto em outro programa !!!", vbOKOnly + vbCritical, "Aviso"
    
    Else
        MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
               "Erro Descr  : " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    End If

End Sub

Private Function Fechamento(intCODFECHA As Integer) As String

    Fechamento = ""
    If intCODFECHA = 0 Then Exit Function
    
    If intCODFECHA = 1 Then Fechamento = "Ø24"
    If intCODFECHA = 2 Then Fechamento = "Ø25"
    If intCODFECHA = 3 Then Fechamento = "Ø42"
    If intCODFECHA = 4 Then Fechamento = "Ø45"
    If intCODFECHA = 5 Then Fechamento = "Ø57"
    If intCODFECHA = 6 Then Fechamento = "Ø80"
    If intCODFECHA = 7 Then Fechamento = "Ø130"
    If intCODFECHA = 8 Then Fechamento = "Ø170"
    If intCODFECHA = 9 Then Fechamento = "Ø110"
    If intCODFECHA = 10 Then Fechamento = "Ø170 c/b Ø25"
    If intCODFECHA = 11 Then Fechamento = "Ø170 c/v Ø57"
    If intCODFECHA = 12 Then Fechamento = "TP"
    If intCODFECHA = 13 Then Fechamento = "TP2"
    If intCODFECHA = 14 Then Fechamento = "TP4"
    If intCODFECHA = 15 Then Fechamento = "FA"
    If intCODFECHA = 16 Then Fechamento = "A RECRAVAR"
    If intCODFECHA = 17 Then Fechamento = "FA - C/Visor"
    If intCODFECHA = 18 Then Fechamento = "COFRE"
    If intCODFECHA = 19 Then Fechamento = "Porta Canetas"
    If intCODFECHA = 20 Then Fechamento = "Ø32 Bico Ret."

End Function

Private Function Pega_Estado(lngINDICE As Long) As String

   Pega_Estado = ""

   Dim V_Estado As Variant
   
   V_Estado = Array("AM", "AC", "AL", "AP", "BA", "CE", "DF", "ES", _
                    "GO", "MA", "MG", "MT", "MS", "PE", "PA", "PB", "PI", "PR", "RJ", _
                    "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO", "EX")

   If (lngINDICE - 1) >= 0 Then Pega_Estado = V_Estado((lngINDICE - 1))

End Function

Private Function SomaFaturamento(strIDPRODUTO As String, strCODOP As String, strNOMFILIAL As String)
    
    SomaFaturamento = "0"
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      Sum(ITEN.SGI_QTDREAL) As SGI_QTDREAL" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADORDCONFI" & strNOMFILIAL & " ITEN" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITEN.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITEN.SGI_IDPRODUTO  = " & Trim(strIDPRODUTO) & vbCrLf
    sSql = sSql & "  And ITEN.SGI_CODORDPROD = " & Trim(strCODOP)
    
    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then
       If Not IsNull(BREC11!SGI_QTDREAL) Then SomaFaturamento = Trim(Str(BREC11!SGI_QTDREAL))
    End If
    BREC11.Close

End Function

Private Function FechamentoOP(lngQTDOP As Long, lngQTDFAT As Long) As String

    FechamentoOP = ""

    Dim currTOLERANCIA  As Currency
    Dim lngTOTOLETANCIA As Long
    Dim lngTOTOLOP      As Long

    If lngQTDOP <= 29000 Then currTOLERANCIA = 0.15
    If lngQTDOP > 29001 Then currTOLERANCIA = 0.1
    
    lngTOTOLETANCIA = (lngQTDOP * currTOLERANCIA)
    lngTOTOLOP = (lngQTDOP - lngTOTOLETANCIA)

    If lngQTDFAT >= lngTOTOLOP Then FechamentoOP = "Finalizada"
    If lngQTDFAT < lngTOTOLOP Then FechamentoOP = "ABERTO"

End Function


Private Sub ImpRel2()

On Error GoTo Err_Exporta

    Dim strNOMFILIAL    As String
    Dim strNomRel       As String
    Dim boolTemDados    As Boolean
    Dim lngQTDFOLHAS    As Long
    Dim dblPERDAPROC    As Double
    Dim lngQTDEFOLHAS   As Long
    Dim strESMALTE      As String
    Dim arrCORES        As Variant
    Dim intQTDCORES     As Integer
    Dim lngQTDREGS      As Long
    Dim lngQTDTOTAL     As Long
    Dim intCODFECHA     As Integer
    Dim lngQTDEOP       As Long
    Dim lngQTDETOTOP    As Long
    Dim strSTATUSOP     As String
    Dim lngCODOP        As Long
    
    Dim strCAMPO01      As String
    Dim strCAMPO02      As String
    Dim strCAMPO03      As String
    Dim strCAMPO04      As String
    
    Dim strDADOS01      As String
    Dim strDADOS02      As String
    Dim strDADOS03      As String
    Dim strDADOS04      As String
    
    Dim strVERNIZ01     As String
    Dim strVERNIZ02     As String
    Dim strVERNIZACAB   As String
    Dim strNECKIN       As String
    Dim strESTADO       As String
    Dim strFECHAGRAF    As String
    Dim strVERNCORPO    As String
    Dim strVERNTAMPA    As String
    Dim strVERNFUNDO    As String
    Dim strVERNARGOLA   As String
    Dim strOBSOP        As String
    Dim strTOTFAT       As String
    Dim strSTATUS2      As String
    Dim strVENDNOVA     As String
    Dim strVENDSTEEL    As String
    
    
    Frame9.Visible = True
    ''If optFilial(1).Value = True Then
    ''    strNomRel = "RELPREPARA01_STEEL.TXT"
    ''ElseIf optFilial(0).Value = True Then
        strNomRel = "RELPREPARAPROD.TXT"
    ''End If
    
    boolTemDados = True
    
    strNOMFILIAL = ""
    If optFilial(1).value = True Then strNOMFILIAL = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.*" & vbCrLf
    sSql = sSql & "      ,TIP.SGI_DESCRICAO As DESC_TIPO" & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO As DESC_ESP" & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD ESP" & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPPROD TIP" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE CLI" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_PRODUTOTIPO   = 1" & vbCrLf
    sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
    sSql = sSql & "   And PRO.SGI_STATUS = 1" & vbCrLf
    
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        sSql = sSql & "   And PRO.SGI_CODCLIE = " & Trim(txtCIDCLIE.Text) & vbCrLf
    End If
    
    sSql = sSql & "   And ESP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PRO.SGI_CODESPECIE" & vbCrLf
    sSql = sSql & "   And TIP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And TIP.SGI_CODIGO = PRO.SGI_CODTIPO" & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PRO.SGI_CODCLIE" & vbCrLf
    
    sSql = sSql & "Order by PRO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "        ,PRO.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "        ,PRO.SGI_CODROTULO" & vbCrLf
    sSql = sSql & "        ,PRO.SGI_DIGVERIF"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       
       Open strCamRelNovo & strNomRel For Output As #1
       
       strCAMPO01 = "SETOR" & vbTab & _
                    "VENDEDOR STEEL" & vbTab & _
                    "VENDEDOR NOVALATA" & vbTab & _
                    "CLIENTE" & vbTab & _
                    "ROTULO" & vbTab & _
                    "DESCRICAO" & vbTab & _
                    "TIPO" & vbTab & _
                    "VERNIZ.INT 01" & vbTab & _
                    "VERNIZ.INT 02"
        
        strCAMPO02 = "ESMALTE" & vbTab & _
                     "REVESTIMENTO" & vbTab & _
                     "VERNIZ.ACABAMENTO" & vbTab & _
                     "1a.COR" & vbTab & _
                     "2a.COR" & vbTab & _
                     "3a.COR" & vbTab & _
                     "4a.COR" & vbTab & _
                     "5a.COR" & vbTab & _
                     "6a.COR" & vbTab & _
                     "7a.COR" & vbTab & _
                     "8a.COR" & vbTab & _
                     "FECHAMENTO" & vbTab & _
                     "Neck IN"
                     
        strCAMPO03 = "Fecham.AGRAF" & vbTab & _
                     "Verniz CP" & vbTab & _
                     "Verniz TP" & vbTab & _
                     "Verniz FD" & vbTab & _
                     "Verniz ARG"
                     
       
       Print #1, strCAMPO01 & vbTab & _
                 strCAMPO02 & vbTab & _
                 strCAMPO03
                 
       lngQTDREGS = 0
       lngQTDTOTAL = 0
       prgPREP.Min = lngQTDREGS
       Do While Not BREC.EOF()
          lngQTDREGS = (lngQTDREGS + 1)
          BREC.MoveNext
       Loop
       If lngQTDREGS > 0 Then
          prgPREP.Max = lngQTDREGS
          lngQTDTOTAL = lngQTDREGS
       End If
        
        
        
       BREC.MoveFirst
       lngQTDREGS = 0
       Do While Not BREC.EOF()
       
            lngQTDREGS = (lngQTDREGS + 1)
            prgPREP.value = lngQTDREGS
            
            strFECHAGRAF = ""
            If BREC!SGI_FechSoldaAgrafado = 0 Then strFECHAGRAF = "SOLDA"
            If BREC!SGI_FechSoldaAgrafado = 1 Then strFECHAGRAF = "AGRAFADO"
            If BREC!SGI_FechSoldaAgrafado = 2 Then strFECHAGRAF = "REPUXO"
            
            strVERNCORPO = ""
            If BREC!SGI_VernCorpo = 1 Then strVERNCORPO = "VEX"
            If BREC!SGI_VernCorpo = 2 Then strVERNCORPO = "VZ"
            If BREC!SGI_VernCorpo = 3 Then strVERNCORPO = "NAT"
            If BREC!SGI_VernCorpo = 4 Then strVERNCORPO = "VI"
            
            strVERNTAMPA = ""
            If BREC!SGI_VernTampa = 1 Then strVERNTAMPA = "VEX"
            If BREC!SGI_VernTampa = 2 Then strVERNTAMPA = "VZ"
            If BREC!SGI_VernTampa = 3 Then strVERNTAMPA = "NAT"
            If BREC!SGI_VernTampa = 4 Then strVERNTAMPA = "VI"
            
            strVERNFUNDO = ""
            If BREC!SGI_VernFundo = 1 Then strVERNFUNDO = "VEX"
            If BREC!SGI_VernFundo = 2 Then strVERNFUNDO = "VZ"
            If BREC!SGI_VernFundo = 3 Then strVERNFUNDO = "NAT"
            If BREC!SGI_VernFundo = 4 Then strVERNFUNDO = "VI"
            
            strVERNARGOLA = ""
            If BREC!SGI_VernArgola = 1 Then strVERNARGOLA = "VEX"
            If BREC!SGI_VernArgola = 2 Then strVERNARGOLA = "VZ"
            If BREC!SGI_VernArgola = 3 Then strVERNARGOLA = "NAT"
            If BREC!SGI_VernArgola = 4 Then strVERNARGOLA = "VI"
       
            '' Verniz 01
            strVERNIZ01 = ""
            
            sSql = ""
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_VERNIZPROD VER" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO PRD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then strVERNIZ01 = BREC2!SGI_DESCRICAO
            BREC2.Close
            
            '' Verniz 02
            strVERNIZ02 = ""
            
            sSql = ""
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_VERNIZPROD02 VER" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO   PRD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then strVERNIZ02 = BREC2!SGI_DESCRICAO
            BREC2.Close
            
            '' Esmalte
            strESMALTE = ""
            
            sSql = ""
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_ESMALTEPROD  VER" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO   PRD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then strESMALTE = BREC2!SGI_DESCRICAO
            BREC2.Close
            
            '' Verniz Acabamento
            strVERNIZACAB = ""
            
            sSql = ""
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_VERNIZPRODACAB  VER" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO      PRD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then strVERNIZACAB = BREC2!SGI_DESCRICAO
            BREC2.Close
            
            
            '' --------------------------------
            '' Pega Cores
            ReDim arrCORES(1 To 8) As String
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PROD.SGI_DESCRICAO" & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CORESPROD CORES" & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       CORES.SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And CORES.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And CORES.SGI_FILIAL    = PROD.SGI_FILIAL"
            sSql = sSql & "   And CORES.SGI_CODCOR    = PROD.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then
                intQTDCORES = 1
                Do While Not BREC2.EOF()
                   arrCORES(intQTDCORES) = BREC2!SGI_DESCRICAO
                   intQTDCORES = (intQTDCORES + 1)
                   BREC2.MoveNext
                Loop
            End If
            BREC2.Close
            '' ----------------------------------
            
            intCODFECHA = 0
            If Not IsNull(BREC!SGI_FechTampaFuro) Then intCODFECHA = BREC!SGI_FechTampaFuro
       
            strNECKIN = "NÃO"
            If BREC!SGI_NECKIN = 1 Then strNECKIN = "SIM"
       
            strVENDSTEEL = ""
            If BREC!SGI_FILIALPED Then strVENDSTEEL = PegaVendedor(BREC!SGI_IDPRODUTO, "_STEEL")
            strVENDNOVA = PegaVendedor(BREC!SGI_IDPRODUTO, "")
       
       
            strDADOS01 = IIf(BREC!SGI_FILIALPED = 0, "STEEL", "NOVALATA") & vbTab & _
                         strVENDSTEEL & vbTab & _
                         strVENDNOVA & vbTab & _
                         Trim(BREC!SGI_RAZAOSOC) & vbTab & _
                         Format(IIf(IsNull(BREC!SGI_CODLINPROD), 0, BREC!SGI_CODLINPROD), "###000") & "." & _
                         Format(IIf(IsNull(BREC!SGI_CODCLIE), 0, BREC!SGI_CODCLIE), "####0000") & "." & _
                         Format(IIf(IsNull(BREC!SGI_CODROTULO), 0, BREC!SGI_CODROTULO), "##00") & "." & _
                         Format(IIf(IsNull(BREC!SGI_DIGVERIF), 0, BREC!SGI_DIGVERIF), "#0") & vbTab & _
                         Replace(Replace(BREC!SGI_DESCRICAO, "Ç", "C"), "Ã", "A") & vbTab & _
                         Trim(BREC!DESC_TIPO) & vbTab & _
                         strVERNIZ01 & vbTab & _
                         strVERNIZ02 & vbTab & _
                         strESMALTE & vbTab & _
                         Trim(BREC!DESC_ESP) & vbTab & _
                         strVERNIZACAB

            strDADOS02 = arrCORES(1) & vbTab & _
                         arrCORES(2) & vbTab & _
                         arrCORES(3) & vbTab & _
                         arrCORES(4) & vbTab & _
                         arrCORES(5) & vbTab & _
                         arrCORES(6) & vbTab & _
                         arrCORES(7) & vbTab & _
                         arrCORES(8) & vbTab & _
                         Fechamento(intCODFECHA) & vbTab & _
                         strNECKIN
       

            strDADOS03 = strFECHAGRAF & vbTab & _
                         strVERNCORPO & vbTab & _
                         strVERNTAMPA & vbTab & _
                         strVERNFUNDO & vbTab & _
                         strVERNARGOLA & vbTab & _
                         strTOTFAT & vbTab & _
                         strSTATUS2
            
            Print #1, strDADOS01 & vbTab & _
                      strDADOS02 & vbTab & _
                      strDADOS03
         
         BREC.MoveNext
       Loop
       Close #1
    
       MsgBox "Arquivo Gerado com Exito !!!", vbOKOnly + vbInformation, "Aviso"
    
    Else
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       boolTemDados = False
    End If
    BREC.Close
    
    Frame9.Visible = False
    
    Exit Sub

Err_Exporta:
    
    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    Close #1
    
    
    If Err.Number = 70 Then
        MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
               "Erro Descr  : " & "Não pode inportar pois o Arguivo TXT eta aberto em outro programa !!!", vbOKOnly + vbCritical, "Aviso"
    
    Else
        MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
               "Erro Descr  : " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    End If

End Sub



Public Function PegaVendedor(strIDPRODUTO As String, strEMPRESA As String) As String

    PegaVendedor = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       VEND.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI" & strEMPRESA & " ITEN" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strEMPRESA & " CAB" & vbCrLf
    sSql = sSql & "      ,SGI_CADVENDEDOR VEND" & vbCrLf
     
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And ITEN.SGI_IDPRODUTO = " & strIDPRODUTO & vbCrLf
    sSql = sSql & "   And CAB.SGI_FILIAL     = ITEN.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CAB.SGI_CODIGO     = ITEN.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And VEND.SGI_FILIAL    = CAB.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And VEND.SGI_CODIGO    = CAB.SGI_CODVEND"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then PegaVendedor = BREC10!SGI_DESCRICAO
    BREC10.Close
    
    
End Function

Private Sub LimpaListBox()
    lstFamProd.Clear
End Sub

Private Sub optTipo_Click(Index As Integer)
    Frame6.Enabled = False
    If Index = 2 Then Frame6.Enabled = True
End Sub

Private Sub PopLSTBoxFam()

    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
        lstFamProd.AddItem Trim(BREC!SGI_DESCRICAO)
        lstFamProd.ItemData(lstFamProd.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
    Loop
    BREC.Close

End Sub

Private Sub ImpRel3()

On Error GoTo Err_Exporta

    Dim strNOMFILIAL    As String
    Dim strNomRel       As String
    Dim boolTemDados    As Boolean
    Dim lngQTDFOLHAS    As Long
    Dim dblPERDAPROC    As Double
    Dim lngQTDEFOLHAS   As Long
    Dim strESMALTE      As String
    Dim arrCORES        As Variant
    Dim intQTDCORES     As Integer
    Dim lngQTDREGS      As Long
    Dim lngQTDTOTAL     As Long
    Dim intCODFECHA     As Integer
    Dim lngQTDEOP       As Long
    Dim lngQTDETOTOP    As Long
    Dim strSTATUSOP     As String
    Dim lngCODOP        As Long
    
    Dim strCAMPO01      As String
    Dim strCAMPO02      As String
    Dim strCAMPO03      As String
    Dim strCAMPO04      As String
    
    Dim strDADOS01      As String
    Dim strDADOS02      As String
    Dim strDADOS03      As String
    Dim strDADOS04      As String
    
    Dim strVERNIZ01     As String
    Dim strVERNIZ02     As String
    Dim strVERNIZACAB   As String
    Dim strNECKIN       As String
    Dim strESTADO       As String
    Dim strFECHAGRAF    As String
    Dim strVERNCORPO    As String
    Dim strVERNTAMPA    As String
    Dim strVERNFUNDO    As String
    Dim strVERNARGOLA   As String
    Dim strOBSOP        As String
    Dim strTOTFAT       As String
    Dim strSTATUS2      As String
    Dim strVENDNOVA     As String
    Dim strVENDSTEEL    As String
    Dim I               As Integer
    
    
    
    Frame9.Visible = True
    ''If optFilial(1).Value = True Then
    ''    strNomRel = "RELPREPARA01_STEEL.TXT"
    ''ElseIf optFilial(0).Value = True Then
        strNomRel = "RELPREPRODNORM.TXT"
    ''End If
    
    
    strDADOS01 = ""
    strDADOS02 = ""
    With lstFamProd
        For I = 0 To (.ListCount - 1)
            If .Selected(I) = True Then
                strDADOS01 = .ItemData(I)
                If I < (.ListCount - 1) Then
                    strDADOS02 = strDADOS02 & strDADOS01 & ";"
                End If
            End If
        Next I
    End With
    
    If Len(Trim(strDADOS02)) > 0 Then
        arrCORES = Split(strDADOS02, ";")
        strDADOS02 = ""
        For I = 0 To UBound(arrCORES)
            If Len(Trim(arrCORES(I))) > 0 Then
                strDADOS02 = strDADOS02 & arrCORES(I) & ","
            End If
        Next I
        strDADOS02 = Mid(strDADOS02, 1, (Len(Trim(strDADOS02)) - 1))
    End If
    
    boolTemDados = True
    
    strNOMFILIAL = ""
    If optFilial(1).value = True Then strNOMFILIAL = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.*" & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO As DESC_ESP" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPROD ESP" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_PRODUTOTIPO   = 0" & vbCrLf
    sSql = sSql & "   And PRO.SGI_PRODUTOESTILO = 0" & vbCrLf
    sSql = sSql & "   And PRO.SGI_STATUS = 1" & vbCrLf
    
    If Len(Trim(strDADOS02)) > 0 Then
        sSql = sSql & "   And PRO.SGI_CODGPROD IN(" & strDADOS02 & ")" & vbCrLf
    End If
    
    sSql = sSql & "   And ESP.SGI_FILIAL = PRO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO = PRO.SGI_CODGPROD" & vbCrLf
    
    sSql = sSql & "Order by PRO.SGI_CODGPROD,PRO.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       
       Open strCamRelNovo & strNomRel For Output As #1
       
       strCAMPO01 = "CÓDIGO" & vbTab & _
                    "DESCRICAO" & vbTab & _
                    "GRUPO"
        
       Print #1, strCAMPO01
                 
       lngQTDREGS = 0
       lngQTDTOTAL = 0
       prgPREP.Min = lngQTDREGS
       Do While Not BREC.EOF()
          lngQTDREGS = (lngQTDREGS + 1)
          BREC.MoveNext
       Loop
       If lngQTDREGS > 0 Then
          prgPREP.Max = lngQTDREGS
          lngQTDTOTAL = lngQTDREGS
       End If
        
        
        
       BREC.MoveFirst
       lngQTDREGS = 0
       Do While Not BREC.EOF()
       
            lngQTDREGS = (lngQTDREGS + 1)
            prgPREP.value = lngQTDREGS
            
       
            strDADOS01 = IIf(IsNull(BREC!SGI_CODIGO) = False, BREC!SGI_CODIGO, "") & vbTab & _
                         Replace(Replace(BREC!SGI_DESCRICAO, "Ç", "C"), "Ã", "A") & vbTab & _
                         Trim(BREC!DESC_ESP)
            
            Print #1, strDADOS01
         
         BREC.MoveNext
       Loop
       Close #1
    
       MsgBox "Arquivo Gerado com Exito !!!", vbOKOnly + vbInformation, "Aviso"
    
    Else
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       boolTemDados = False
    End If
    BREC.Close
    
    Frame9.Visible = False '
    
    Exit Sub

Err_Exporta:
    
    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    Close #1
    
    
    If Err.Number = 70 Then
        MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
               "Erro Descr  : " & "Não pode inportar pois o Arguivo TXT eta aberto em outro programa !!!", vbOKOnly + vbCritical, "Aviso"
    
    Else
        MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
               "Erro Descr  : " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    End If

End Sub


Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, StrCodigo As String, lblLabel As Label)

On Error GoTo Err_PegaDescTabelas

    lblLabel.Caption = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(StrCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(StrCodigo)
    
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
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "I", "Função : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub


Private Sub LimpaCamposLabel()
    lblDescCliente.Caption = ""
End Sub

Private Sub txtCIDCLIE_GotFocus()

On Error GoTo Err_txtCIDCLIE_GotFocus

    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, Me

    Exit Sub
    
Err_txtCIDCLIE_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "I", "Função : txtCIDCLIE_GotFocus()", Me.Name, "txtCIDCLIE_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)

On Error GoTo Err_txtCIDCLIE_KeyPress
    
    objBLBFunc.SoNumeroPonto KeyAscii, txtCIDCLIE.Text

    Exit Sub
    
Err_txtCIDCLIE_KeyPress:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "I", "Função : txtCIDCLIE_KeyPress()", Me.Name, "txtCIDCLIE_KeyPress()", strCAMARQERRO)

End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

On Error GoTo Err_txtCIDCLIE_Validate

    Dim I As Integer
    
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

    Exit Sub
    
Err_txtCIDCLIE_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "I", "Função : txtCIDCLIE_Validate()", Me.Name, "txtCIDCLIE_Validate()", strCAMARQERRO)

End Sub

Private Sub ImpRelEXCEL()

On Error GoTo Err_Exporta

    Dim strNOMFILIAL            As String
    Dim strNomRel               As String
    Dim boolTemDados            As Boolean
    Dim lngQTDFOLHAS            As Long
    Dim dblPERDAPROC            As Double
    Dim lngQTDEFOLHAS           As Long
    Dim strESMALTE              As String
    Dim arrCORES                As Variant
    Dim intQTDCORES             As Integer
    Dim lngQTDREGS              As Long
    Dim lngQTDTOTAL             As Long
    Dim intCODFECHA             As Integer
    Dim lngQTDEOP               As Long
    Dim lngQTDETOTOP            As Long
    Dim strSTATUSOP             As String
    Dim lngCODOP                As Long
    
    Dim strCAMPO01              As String
    Dim strCAMPO02              As String
    Dim strCAMPO03              As String
    Dim strCAMPO04              As String
    
    Dim strDADOS01              As String
    Dim strDADOS02              As String
    Dim strDADOS03              As String
    Dim strDADOS04              As String
    
    Dim strVERNIZ01             As String
    Dim strVERNIZ02             As String
    Dim strVERNIZACAB           As String
    Dim strNECKIN               As String
    Dim strESTADO               As String
    Dim strFECHAGRAF            As String
    Dim strVERNCORPO            As String
    Dim strVERNTAMPA            As String
    Dim strVERNFUNDO            As String
    Dim strVERNARGOLA           As String
    Dim strOBSOP                As String
    Dim strTOTFAT               As String
    Dim strSTATUS2              As String
    
    Dim lngDADOSNOVALATA        As Long
    Dim lngDADOSSTEEL           As Long
    
    
    Frame9.Visible = True
    If optFilial(1).value = True Then
        strNomRel = "RELPREPARA01_STEEL.xls"
    ElseIf optFilial(0).value = True Then
        strNomRel = "RELPREPARA01_NOVA.xls"
    ElseIf optFilial(2).value = True Then
        strNomRel = "RELPREPARA.xls"
    End If
    
    boolTemDados = False
    
    strNOMFILIAL = ""
    If optFilial(1).value = True Then strNOMFILIAL = "_STEEL"
    
    lngDADOSNOVALATA = 0
    lngDADOSSTEEL = 0
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATAORDEM" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATAORDEM" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_QTDE" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_NOMEVEND" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_STATUS" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FECHTPFU" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_OBSOP" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_QTDEPORFOLHA" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_QTDCORPSPADRAOSN" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechSoldaAgrafado" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernCorpo" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernTampa" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernFundo" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernArgola" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_QTDECORPOS" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_PERDPROC" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADESPPROD.SGI_DESCRICAO As SGI_DESCESP" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CODIGO     As SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ESTNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CIDNORM" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADTIPPROD.SGI_DESCRICAO As SGI_DESCTIPO" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_CODVEND" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_ESTENTRE" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " SGI_CADPEDVENDH" & strNOMFILIAL & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & " SGI_ORDEMPROD" & strNOMFILIAL & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD SGI_CADESPPROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPPROD SGI_CADTIPPROD" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATENTREGA Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL    = SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPED    = SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADESPPROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODESPECIE = SGI_CADESPPROD.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADTIPPROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODTIPO    = SGI_CADTIPPROD.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       
           boolTemDados = True
           
           lngQTDREGS = 0
           lngQTDTOTAL = 0
           prgPREP.Min = lngQTDREGS
           Do While Not BREC.EOF()
              lngQTDREGS = (lngQTDREGS + 1)
              BREC.MoveNext
           Loop
           If lngQTDREGS > 0 Then
              
              If optFilial(0).value = True Or optFilial(2).value = True Then
                ReDim arrDADOSNOVALATA(1 To lngQTDREGS) As String
                Frame9.Caption = "[ Aguarde... Carregando os dados NOVALATA ]"
                Frame9.Refresh
                lngDADOSNOVALATA = lngQTDREGS
              ElseIf optFilial(1).value = True Then
                ReDim arrDADOSSTEEL(1 To lngQTDREGS) As String
                Frame9.Caption = "[ Aguarde... Carregando os dados STEEL ]"
                Frame9.Refresh
                lngDADOSSTEEL = lngQTDREGS
              End If
              
              prgPREP.Max = lngQTDREGS
              lngQTDTOTAL = lngQTDREGS
          End If
            
           
           BREC.MoveFirst
           lngQTDREGS = 0
           
           Do While Not BREC.EOF()
           
                lngQTDREGS = (lngQTDREGS + 1)
                
                prgPREP.value = lngQTDREGS
                prgPREP.Refresh
                
                strOBSOP = ""
                If Not IsNull(BREC!SGI_OBSOP) Then strOBSOP = Trim(Replace(BREC!SGI_OBSOP, vbCrLf, " , "))
                
                If BREC!SGI_STATUS = 0 Then strSTATUSOP = "ABERTO"
                If BREC!SGI_STATUS = 1 Then strSTATUSOP = "Fat.Parcial"
                If BREC!SGI_STATUS = 2 Then strSTATUSOP = "Finalizada"
                If BREC!SGI_STATUS = 3 Then strSTATUSOP = "Bloqueada"
                
                strFECHAGRAF = ""
                If BREC!SGI_FechSoldaAgrafado = 0 Then strFECHAGRAF = "SOLDA"
                If BREC!SGI_FechSoldaAgrafado = 1 Then strFECHAGRAF = "AGRAFADO"
                If BREC!SGI_FechSoldaAgrafado = 2 Then strFECHAGRAF = "REPUXO"
                
                strVERNCORPO = ""
                If BREC!SGI_VernCorpo = 1 Then strVERNCORPO = "VEX"
                If BREC!SGI_VernCorpo = 2 Then strVERNCORPO = "VZ"
                If BREC!SGI_VernCorpo = 3 Then strVERNCORPO = "NAT"
                If BREC!SGI_VernCorpo = 4 Then strVERNCORPO = "VI"
                
                strVERNTAMPA = ""
                If BREC!SGI_VernTampa = 1 Then strVERNTAMPA = "VEX"
                If BREC!SGI_VernTampa = 2 Then strVERNTAMPA = "VZ"
                If BREC!SGI_VernTampa = 3 Then strVERNTAMPA = "NAT"
                If BREC!SGI_VernTampa = 4 Then strVERNTAMPA = "VI"
                
                strVERNFUNDO = ""
                If BREC!SGI_VernFundo = 1 Then strVERNFUNDO = "VEX"
                If BREC!SGI_VernFundo = 2 Then strVERNFUNDO = "VZ"
                If BREC!SGI_VernFundo = 3 Then strVERNFUNDO = "NAT"
                If BREC!SGI_VernFundo = 4 Then strVERNFUNDO = "VI"
                
                strVERNARGOLA = ""
                If BREC!SGI_VernArgola = 1 Then strVERNARGOLA = "VEX"
                If BREC!SGI_VernArgola = 2 Then strVERNARGOLA = "VZ"
                If BREC!SGI_VernArgola = 3 Then strVERNARGOLA = "NAT"
                If BREC!SGI_VernArgola = 4 Then strVERNARGOLA = "VI"
                
                '' Pegava o Estado de Entrega
                ''strESTADO = Pega_Estado(BREC!SGI_ESTENTRE)
                strESTADO = Pega_Estado(BREC!SGI_ESTNORM)
                
                lngQTDFOLHAS = 0
                dblPERDAPROC = 0
                lngQTDEFOLHAS = 0
                If BREC!SGI_QTDCORPSPADRAOSN = 0 Then
                   If Not IsNull(BREC!SGI_QTDEPORFOLHA) Then lngQTDFOLHAS = BREC!SGI_QTDEPORFOLHA
                   dblPERDAPROC = 1.05
                ElseIf BREC!SGI_QTDCORPSPADRAOSN = 1 Then
                   If Not IsNull(BREC!SGI_QTDECORPOS) Then lngQTDFOLHAS = BREC!SGI_QTDECORPOS
                   If Not IsNull(BREC!SGI_PERDPROC) Then dblPERDAPROC = BREC!SGI_PERDPROC
                End If
                If lngQTDFOLHAS > 0 Then lngQTDEFOLHAS = ((BREC!SGI_QTDE * dblPERDAPROC) / lngQTDFOLHAS)
           
                '' Verniz 01
                strVERNIZ01 = ""
                
                sSql = ""
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_VERNIZPROD VER" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO PRD" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
                sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF() Then strVERNIZ01 = BREC2!SGI_DESCRICAO
                BREC2.Close
                
                '' Verniz 02
                strVERNIZ02 = ""
                
                sSql = ""
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_VERNIZPROD02 VER" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO   PRD" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
                sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF() Then strVERNIZ02 = BREC2!SGI_DESCRICAO
                BREC2.Close
                
                '' Esmalte
                strESMALTE = ""
                
                sSql = ""
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_ESMALTEPROD  VER" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO   PRD" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
                sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF() Then strESMALTE = BREC2!SGI_DESCRICAO
                BREC2.Close
                
                '' Verniz Acabamento
                strVERNIZACAB = ""
                
                sSql = ""
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_VERNIZPRODACAB  VER" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO      PRD" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
                sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF() Then strVERNIZACAB = BREC2!SGI_DESCRICAO
                BREC2.Close
                
                
                '' --------------------------------
                '' Pega Cores
                ReDim arrCORES(1 To 8) As String
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PROD.SGI_DESCRICAO" & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CORESPROD CORES" & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CORES.SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & "   And CORES.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "   And CORES.SGI_FILIAL    = PROD.SGI_FILIAL"
                sSql = sSql & "   And CORES.SGI_CODCOR    = PROD.SGI_IDPRODUTO"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF() Then
                    intQTDCORES = 1
                    Do While Not BREC2.EOF()
                       arrCORES(intQTDCORES) = BREC2!SGI_DESCRICAO
                       intQTDCORES = (intQTDCORES + 1)
                       BREC2.MoveNext
                    Loop
                End If
                BREC2.Close
                '' ----------------------------------
                
                intCODFECHA = 0
                If Not IsNull(BREC!SGI_FECHTPFU) Then intCODFECHA = BREC!SGI_FECHTPFU
           
                strNECKIN = "NÃO"
                If BREC!SGI_NECKIN = 1 Then strNECKIN = "SIM"
           
                '' ---------------------------------------------
                '' Bloco de Dados 01
                strDADOS01 = BREC!SGI_CODVEND & vbTab & _
                             BREC!SGI_NOMEVEND & vbTab & _
                             BREC!SGI_CODCLI & vbTab & _
                             BREC!SGI_RAZAOSOC & vbTab & _
                             BREC!SGI_CODPROD & vbTab & _
                             BREC!SGI_CODPED & vbTab & _
                             strSTATUSOP & vbTab & _
                             Replace(Replace(BREC!SGI_DESCRICAO, "Ç", "C"), "Ã", "A") & vbTab & _
                             BREC!SGI_CODIGO & vbTab & _
                             BREC!SGI_CODLIN & vbTab & _
                             Trim(BREC!SGI_DESCRI) & "." & vbTab & _
                             BREC!SGI_DESCTIPO & vbTab & _
                             Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                             Format((BREC!SGI_DATENTREGA - 7), "DD/MM/YYYY") & vbTab & _
                             Format((BREC!SGI_DATENTREGA - 3), "DD/MM/YYYY") & vbTab & _
                             Format((BREC!SGI_DATENTREGA - 2), "DD/MM/YYYY") & vbTab & _
                             Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                             BREC!SGI_QTDE & vbTab & _
                             lngQTDEFOLHAS & vbTab & _
                             lngQTDFOLHAS & vbTab & _
                             strVERNIZ01 & vbTab & _
                             strVERNIZ02
           
                ''
                '' ----------------------------------------------------
                
                
                '' ---------------------------------------------
                '' Bloco de Dados 02
                strDADOS02 = strESMALTE & vbTab & _
                             BREC!SGI_DESCESP & vbTab & _
                             strVERNIZACAB & vbTab & _
                             "" & vbTab & _
                             arrCORES(1) & vbTab & _
                             arrCORES(2) & vbTab & _
                             arrCORES(3) & vbTab & _
                             arrCORES(4) & vbTab & _
                             arrCORES(5) & vbTab & _
                             arrCORES(6) & vbTab & _
                             arrCORES(7) & vbTab & _
                             arrCORES(8) & vbTab & _
                             "" & vbTab & _
                             Fechamento(intCODFECHA) & vbTab & _
                             strNECKIN
           
                
                '' =======================================
                '' Pegando Faturamentos já realizado
                strTOTFAT = SomaFaturamento(Str(BREC!SGI_IDPRODUTO), Str(BREC!SGI_CODIGO), Trim(strNOMFILIAL))
                
                If BREC!SGI_STATUS = 2 Then
                    strSTATUS2 = "Finalizada"
                Else
                    strSTATUS2 = FechamentoOP(BREC!SGI_QTDE, CLng(strTOTFAT))
                End If
                
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       ITEN.*" & vbCrLf
                sSql = sSql & "      ,CABE.*" & vbCrLf
                sSql = sSql & "      ,ORDP.*" & vbCrLf
                
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADORDCONFI" & strNOMFILIAL & " ITEN" & vbCrLf
                sSql = sSql & "     , SGI_CADORDCONFH" & strNOMFILIAL & " CABE" & vbCrLf
                sSql = sSql & "     , SGI_ORDEMPROD" & strNOMFILIAL & " ORDP" & vbCrLf
                
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       ITEN.SGI_FILIAL     = " & FILIAL & vbCrLf
                sSql = sSql & "   And ITEN.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "   And ITEN.SGI_CODORDPROD = " & BREC!SGI_CODIGO & vbCrLf
                sSql = sSql & "   And ITEN.SGI_FILIAL     = CABE.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And ITEN.SGI_CODCONF    = CABE.SGI_CODCONF" & vbCrLf
                sSql = sSql & "   And ITEN.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And ITEN.SGI_CODORDPROD = ORDP.SGI_CODIGO" & vbCrLf
                
                sSql = sSql & "Order By CABE.SGI_DATACONF"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                
                    lngQTDEOP = BREC!SGI_QTDE '' Qtde OP
                lngCODOP = 0
                If Not BREC2.EOF() Then
                    Do While Not BREC2.EOF()
                        
                        lngQTDEOP = (lngQTDEOP - BREC2!SGI_QTDREAL)
                        
                        strDADOS03 = BREC2!SGI_QTDREAL & vbTab & _
                                     Format(BREC2!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                                     BREC2!SGI_CODFATURA & vbTab & _
                                     lngQTDEOP
                        
                        If lngCODOP <> BREC2!SGI_CODORDPROD Then
                            
                            lngCODOP = BREC2!SGI_CODORDPROD
                            
                            If optFilial(0).value = True Or optFilial(2).value = True Then
                            
                                arrDADOSNOVALATA(lngQTDREGS) = strDADOS01 & vbTab & _
                                                               strDADOS02 & vbTab & _
                                                               BREC2!SGI_QTDE & vbTab & _
                                                               strDADOS03 & vbTab & _
                                                               BREC2!SGI_CODORDPROD & vbTab & _
                                                               strESTADO & vbTab & _
                                                               BREC!SGI_CIDNORM & vbTab & _
                                                               strOBSOP & vbTab & _
                                                               strFECHAGRAF & vbTab & _
                                                               strVERNCORPO & vbTab & _
                                                               strVERNTAMPA & vbTab & _
                                                               strVERNFUNDO & vbTab & _
                                                               strVERNARGOLA & vbTab & _
                                                               strTOTFAT & vbTab & _
                                                               strSTATUS2
                                      
                            ElseIf optFilial(1).value = True Then
                                
                                arrDADOSSTEEL(lngQTDREGS) = strDADOS01 & vbTab & _
                                                            strDADOS02 & vbTab & _
                                                            BREC2!SGI_QTDE & vbTab & _
                                                            strDADOS03 & vbTab & _
                                                            BREC2!SGI_CODORDPROD & vbTab & _
                                                            strESTADO & vbTab & _
                                                            BREC!SGI_CIDNORM & vbTab & _
                                                            strOBSOP & vbTab & _
                                                            strFECHAGRAF & vbTab & _
                                                            strVERNCORPO & vbTab & _
                                                            strVERNTAMPA & vbTab & _
                                                            strVERNFUNDO & vbTab & _
                                                            strVERNARGOLA & vbTab & _
                                                            strTOTFAT & vbTab & _
                                                            strSTATUS2
                            End If
                        
                        Else
                            strDADOS01 = BREC!SGI_CODVEND & vbTab & _
                                         BREC!SGI_NOMEVEND & vbTab & _
                                         BREC!SGI_CODCLI & vbTab & _
                                         BREC!SGI_RAZAOSOC & vbTab & _
                                         BREC!SGI_CODPROD & vbTab & _
                                         BREC!SGI_CODPED & vbTab & _
                                         strSTATUSOP & vbTab & _
                                         Replace(Replace(BREC!SGI_DESCRICAO, "Ç", "C"), "Ã", "A") & vbTab & _
                                         BREC!SGI_CODIGO & vbTab & _
                                         BREC!SGI_CODLIN & vbTab & _
                                         Trim(BREC!SGI_DESCRI) & "." & vbTab & _
                                         BREC!SGI_DESCTIPO & vbTab & _
                                         Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                                         Format((BREC!SGI_DATENTREGA - 7), "DD/MM/YYYY") & vbTab & _
                                         Format((BREC!SGI_DATENTREGA - 3), "DD/MM/YYYY") & vbTab & _
                                         Format((BREC!SGI_DATENTREGA - 2), "DD/MM/YYYY") & vbTab & _
                                         Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         ""
            
                            strDADOS02 = "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         "" & vbTab & _
                                         ""
                        
                            If optFilial(0).value = True Or optFilial(2).value = True Then
                                
                                arrDADOSNOVALATA(lngQTDREGS) = strDADOS01 & vbTab & _
                                                               strDADOS02 & vbTab & _
                                                               "" & vbTab & _
                                                               strDADOS03 & vbTab & _
                                                               "" & vbTab & _
                                                               strESTADO & vbTab & _
                                                               BREC!SGI_CIDNORM & vbTab & _
                                                               strOBSOP & vbTab & _
                                                               strFECHAGRAF & vbTab & _
                                                               strVERNCORPO & vbTab & _
                                                               strVERNTAMPA & vbTab & _
                                                               strVERNFUNDO & vbTab & _
                                                               strVERNARGOLA & vbTab & _
                                                               strTOTFAT & vbTab & _
                                                               strSTATUS2
                            ElseIf optFilial(1).value = True Then
                                
                                arrDADOSSTEEL(lngQTDREGS) = strDADOS01 & vbTab & _
                                                            strDADOS02 & vbTab & _
                                                            "" & vbTab & _
                                                            strDADOS03 & vbTab & _
                                                            "" & vbTab & _
                                                            strESTADO & vbTab & _
                                                            BREC!SGI_CIDNORM & vbTab & _
                                                            strOBSOP & vbTab & _
                                                            strFECHAGRAF & vbTab & _
                                                            strVERNCORPO & vbTab & _
                                                            strVERNTAMPA & vbTab & _
                                                            strVERNFUNDO & vbTab & _
                                                            strVERNARGOLA & vbTab & _
                                                            strTOTFAT & vbTab & _
                                                            strSTATUS2
                            End If
                        
                        End If
                        BREC2.MoveNext
                    Loop
                Else
                    
                    strDADOS03 = "" & vbTab & _
                                 "" & vbTab & _
                                 "" & vbTab & _
                                 BREC!SGI_QTDE & vbTab & _
                                 "" & vbTab & _
                                 strESTADO & vbTab & _
                                 BREC!SGI_CIDNORM & vbTab & _
                                 strOBSOP & vbTab & _
                                 strFECHAGRAF & vbTab & _
                                 strVERNCORPO & vbTab & _
                                 strVERNTAMPA & vbTab & _
                                 strVERNFUNDO & vbTab & _
                                 strVERNARGOLA & vbTab & _
                                 strTOTFAT & vbTab & _
                                 strSTATUS2
                    
                    If optFilial(0).value = True Or optFilial(2).value = True Then
                        arrDADOSNOVALATA(lngQTDREGS) = strDADOS01 & vbTab & _
                                                       strDADOS02 & vbTab & _
                                                       BREC!SGI_QTDE & vbTab & _
                                                       strDADOS03
                    ElseIf optFilial(1).value = True Then
                        arrDADOSSTEEL(lngQTDREGS) = strDADOS01 & vbTab & _
                                                    strDADOS02 & vbTab & _
                                                    BREC!SGI_QTDE & vbTab & _
                                                    strDADOS03
                    End If
                    
                End If
                BREC2.Close
                '' =======================================
             
             BREC.MoveNext
           Loop
    End If
    BREC.Close
    
    
    '' Dados da STEEL
    If optFilial(2).value = True Then
    
        strNOMFILIAL = "_STEEL"
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODIGO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATENTREGA" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATAORDEM" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPED" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATAORDEM" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_QTDE" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_NOMEVEND" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPROD" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_STATUS" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FECHTPFU" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_OBSOP" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_QTDEPORFOLHA" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_QTDCORPSPADRAOSN" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_FechSoldaAgrafado" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernCorpo" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernTampa" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernFundo" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_VernArgola" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_QTDECORPOS" & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_PERDPROC" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADESPPROD.SGI_DESCRICAO As SGI_DESCESP" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CODIGO As SGI_CODCLI" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_ESTNORM" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CIDNORM" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADTIPPROD.SGI_DESCRICAO As SGI_DESCTIPO" & vbCrLf
        
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_CODVEND" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_ESTENTRE" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " SGI_CADPEDVENDH" & strNOMFILIAL & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & " SGI_ORDEMPROD" & strNOMFILIAL & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADESPPROD SGI_CADESPPROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD SGI_CADTIPPROD" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_DATENTREGA Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
        
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
        
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_FILIAL    = SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_ORDEMPROD" & strNOMFILIAL & ".SGI_CODPED    = SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADESPPROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODESPECIE = SGI_CADESPPROD.SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADTIPPROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODTIPO    = SGI_CADTIPPROD.SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPEDVENDH" & strNOMFILIAL & ".SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           
               boolTemDados = True
               
               lngQTDREGS = 0
               lngQTDTOTAL = 0
               prgPREP.Min = lngQTDREGS
               Do While Not BREC.EOF()
                  lngQTDREGS = (lngQTDREGS + 1)
                  BREC.MoveNext
               Loop
               If lngQTDREGS > 0 Then
                  
                  ReDim arrDADOSSTEEL(1 To lngQTDREGS) As String
                  Frame9.Caption = "[ Aguarde... Carregando os dados STEEL ]"
                  Frame9.Refresh
                  lngDADOSSTEEL = lngQTDREGS
                  
                  prgPREP.Max = lngQTDREGS
                  lngQTDTOTAL = lngQTDREGS
              End If
                
               
               BREC.MoveFirst
               lngQTDREGS = 0
               
               Do While Not BREC.EOF()
               
                    lngQTDREGS = (lngQTDREGS + 1)
                    
                    prgPREP.value = lngQTDREGS
                    prgPREP.Refresh
                    
                    strOBSOP = ""
                    If Not IsNull(BREC!SGI_OBSOP) Then strOBSOP = Trim(Replace(BREC!SGI_OBSOP, vbCrLf, " , "))
                    
                    If BREC!SGI_STATUS = 0 Then strSTATUSOP = "ABERTO"
                    If BREC!SGI_STATUS = 1 Then strSTATUSOP = "Fat.Parcial"
                    If BREC!SGI_STATUS = 2 Then strSTATUSOP = "Finalizada"
                    If BREC!SGI_STATUS = 3 Then strSTATUSOP = "Bloqueada"
                    
                    strFECHAGRAF = ""
                    If BREC!SGI_FechSoldaAgrafado = 0 Then strFECHAGRAF = "SOLDA"
                    If BREC!SGI_FechSoldaAgrafado = 1 Then strFECHAGRAF = "AGRAFADO"
                    If BREC!SGI_FechSoldaAgrafado = 2 Then strFECHAGRAF = "REPUXO"
                    
                    strVERNCORPO = ""
                    If BREC!SGI_VernCorpo = 1 Then strVERNCORPO = "VEX"
                    If BREC!SGI_VernCorpo = 2 Then strVERNCORPO = "VZ"
                    If BREC!SGI_VernCorpo = 3 Then strVERNCORPO = "NAT"
                    If BREC!SGI_VernCorpo = 4 Then strVERNCORPO = "VI"
                    
                    strVERNTAMPA = ""
                    If BREC!SGI_VernTampa = 1 Then strVERNTAMPA = "VEX"
                    If BREC!SGI_VernTampa = 2 Then strVERNTAMPA = "VZ"
                    If BREC!SGI_VernTampa = 3 Then strVERNTAMPA = "NAT"
                    If BREC!SGI_VernTampa = 4 Then strVERNTAMPA = "VI"
                    
                    strVERNFUNDO = ""
                    If BREC!SGI_VernFundo = 1 Then strVERNFUNDO = "VEX"
                    If BREC!SGI_VernFundo = 2 Then strVERNFUNDO = "VZ"
                    If BREC!SGI_VernFundo = 3 Then strVERNFUNDO = "NAT"
                    If BREC!SGI_VernFundo = 4 Then strVERNFUNDO = "VI"
                    
                    strVERNARGOLA = ""
                    If BREC!SGI_VernArgola = 1 Then strVERNARGOLA = "VEX"
                    If BREC!SGI_VernArgola = 2 Then strVERNARGOLA = "VZ"
                    If BREC!SGI_VernArgola = 3 Then strVERNARGOLA = "NAT"
                    If BREC!SGI_VernArgola = 4 Then strVERNARGOLA = "VI"
                    
                    '' Pegava o Estado de Entrega
                    ''strESTADO = Pega_Estado(BREC!SGI_ESTENTRE)
                    strESTADO = Pega_Estado(BREC!SGI_ESTNORM)
                    
                    lngQTDFOLHAS = 0
                    dblPERDAPROC = 0
                    lngQTDEFOLHAS = 0
                    If BREC!SGI_QTDCORPSPADRAOSN = 0 Then
                       If Not IsNull(BREC!SGI_QTDEPORFOLHA) Then lngQTDFOLHAS = BREC!SGI_QTDEPORFOLHA
                       dblPERDAPROC = 1.05
                    ElseIf BREC!SGI_QTDCORPSPADRAOSN = 1 Then
                       If Not IsNull(BREC!SGI_QTDECORPOS) Then lngQTDFOLHAS = BREC!SGI_QTDECORPOS
                       If Not IsNull(BREC!SGI_PERDPROC) Then dblPERDAPROC = BREC!SGI_PERDPROC
                    End If
                    If lngQTDFOLHAS > 0 Then lngQTDEFOLHAS = ((BREC!SGI_QTDE * dblPERDAPROC) / lngQTDFOLHAS)
               
                    '' Verniz 01
                    strVERNIZ01 = ""
                    
                    sSql = ""
                    sSql = "Select " & vbCrLf
                    sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
                    sSql = sSql & "  From " & vbCrLf
                    sSql = sSql & "       SGI_VERNIZPROD VER" & vbCrLf
                    sSql = sSql & "      ,SGI_CADPRODUTO PRD" & vbCrLf
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
                    sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                    sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
                    
                    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    If Not BREC2.EOF() Then strVERNIZ01 = BREC2!SGI_DESCRICAO
                    BREC2.Close
                    
                    '' Verniz 02
                    strVERNIZ02 = ""
                    
                    sSql = ""
                    sSql = "Select " & vbCrLf
                    sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
                    sSql = sSql & "  From " & vbCrLf
                    sSql = sSql & "       SGI_VERNIZPROD02 VER" & vbCrLf
                    sSql = sSql & "      ,SGI_CADPRODUTO   PRD" & vbCrLf
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
                    sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                    sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
                    
                    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    If Not BREC2.EOF() Then strVERNIZ02 = BREC2!SGI_DESCRICAO
                    BREC2.Close
                    
                    '' Esmalte
                    strESMALTE = ""
                    
                    sSql = ""
                    sSql = "Select " & vbCrLf
                    sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
                    sSql = sSql & "  From " & vbCrLf
                    sSql = sSql & "       SGI_ESMALTEPROD  VER" & vbCrLf
                    sSql = sSql & "      ,SGI_CADPRODUTO   PRD" & vbCrLf
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
                    sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                    sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
                    
                    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    If Not BREC2.EOF() Then strESMALTE = BREC2!SGI_DESCRICAO
                    BREC2.Close
                    
                    '' Verniz Acabamento
                    strVERNIZACAB = ""
                    
                    sSql = ""
                    sSql = "Select " & vbCrLf
                    sSql = sSql & "       PRD.SGI_DESCRICAO " & vbCrLf
                    sSql = sSql & "  From " & vbCrLf
                    sSql = sSql & "       SGI_VERNIZPRODACAB  VER" & vbCrLf
                    sSql = sSql & "      ,SGI_CADPRODUTO      PRD" & vbCrLf
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       VER.SGI_FILIAL     = " & FILIAL & vbCrLf
                    sSql = sSql & "   And VER.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                    sSql = sSql & "   And VER.SGI_FILIAL     = PRD.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And VER.SGI_PRODUTO    = PRD.SGI_IDPRODUTO"
                    
                    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    If Not BREC2.EOF() Then strVERNIZACAB = BREC2!SGI_DESCRICAO
                    BREC2.Close
                    
                    
                    '' --------------------------------
                    '' Pega Cores
                    ReDim arrCORES(1 To 8) As String
                    sSql = ""
                    
                    sSql = "Select " & vbCrLf
                    sSql = sSql & "       PROD.SGI_DESCRICAO" & vbCrLf
                    sSql = sSql & "  From " & vbCrLf
                    sSql = sSql & "       SGI_CORESPROD CORES" & vbCrLf
                    sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       CORES.SGI_FILIAL    = " & FILIAL & vbCrLf
                    sSql = sSql & "   And CORES.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
                    sSql = sSql & "   And CORES.SGI_FILIAL    = PROD.SGI_FILIAL"
                    sSql = sSql & "   And CORES.SGI_CODCOR    = PROD.SGI_IDPRODUTO"
                    
                    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    If Not BREC2.EOF() Then
                        intQTDCORES = 1
                        Do While Not BREC2.EOF()
                           arrCORES(intQTDCORES) = BREC2!SGI_DESCRICAO
                           intQTDCORES = (intQTDCORES + 1)
                           BREC2.MoveNext
                        Loop
                    End If
                    BREC2.Close
                    '' ----------------------------------
                    
                    intCODFECHA = 0
                    If Not IsNull(BREC!SGI_FECHTPFU) Then intCODFECHA = BREC!SGI_FECHTPFU
               
                    strNECKIN = "NÃO"
                    If BREC!SGI_NECKIN = 1 Then strNECKIN = "SIM"
               
                    '' ---------------------------------------------
                    '' Bloco de Dados 01
                    
                    strDADOS01 = BREC!SGI_CODVEND & vbTab & _
                                 BREC!SGI_NOMEVEND & vbTab & _
                                 BREC!SGI_CODCLI & vbTab & _
                                 BREC!SGI_RAZAOSOC & vbTab & _
                                 BREC!SGI_CODPROD & vbTab & _
                                 BREC!SGI_CODPED & vbTab & _
                                 strSTATUSOP & vbTab & _
                                 Replace(Replace(BREC!SGI_DESCRICAO, "Ç", "C"), "Ã", "A") & vbTab & _
                                 BREC!SGI_CODIGO & vbTab & _
                                 BREC!SGI_CODLIN & vbTab & _
                                 Trim(BREC!SGI_DESCRI) & "." & vbTab & _
                                 BREC!SGI_DESCTIPO & vbTab & _
                                 Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                                 Format((BREC!SGI_DATENTREGA - 7), "DD/MM/YYYY") & vbTab & _
                                 Format((BREC!SGI_DATENTREGA - 3), "DD/MM/YYYY") & vbTab & _
                                 Format((BREC!SGI_DATENTREGA - 2), "DD/MM/YYYY") & vbTab & _
                                 Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                                 BREC!SGI_QTDE & vbTab & _
                                 lngQTDEFOLHAS & vbTab & _
                                 lngQTDFOLHAS & vbTab & _
                                 strVERNIZ01 & vbTab & _
                                 strVERNIZ02
               
                    ''
                    '' ----------------------------------------------------
                    
                    
                    '' ---------------------------------------------
                    '' Bloco de Dados 02
                    strDADOS02 = strESMALTE & vbTab & _
                                 BREC!SGI_DESCESP & vbTab & _
                                 strVERNIZACAB & vbTab & _
                                 "" & vbTab & _
                                 arrCORES(1) & vbTab & _
                                 arrCORES(2) & vbTab & _
                                 arrCORES(3) & vbTab & _
                                 arrCORES(4) & vbTab & _
                                 arrCORES(5) & vbTab & _
                                 arrCORES(6) & vbTab & _
                                 arrCORES(7) & vbTab & _
                                 arrCORES(8) & vbTab & _
                                 "" & vbTab & _
                                 Fechamento(intCODFECHA) & vbTab & _
                                 strNECKIN
               
                    
                    '' =======================================
                    '' Pegando Faturamentos já realizado
                    strTOTFAT = SomaFaturamento(Str(BREC!SGI_IDPRODUTO), Str(BREC!SGI_CODIGO), Trim(strNOMFILIAL))
                    
                    If BREC!SGI_STATUS = 2 Then
                        strSTATUS2 = "Finalizada"
                    Else
                        strSTATUS2 = FechamentoOP(BREC!SGI_QTDE, CLng(strTOTFAT))
                    End If
                    
                    sSql = ""
                    
                    sSql = "Select " & vbCrLf
                    sSql = sSql & "       ITEN.*" & vbCrLf
                    sSql = sSql & "      ,CABE.*" & vbCrLf
                    sSql = sSql & "      ,ORDP.*" & vbCrLf
                    
                    sSql = sSql & "  From " & vbCrLf
                    sSql = sSql & "       SGI_CADORDCONFI" & strNOMFILIAL & " ITEN" & vbCrLf
                    sSql = sSql & "     , SGI_CADORDCONFH" & strNOMFILIAL & " CABE" & vbCrLf
                    sSql = sSql & "     , SGI_ORDEMPROD" & strNOMFILIAL & " ORDP" & vbCrLf
                    
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       ITEN.SGI_FILIAL     = " & FILIAL & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_CODORDPROD = " & BREC!SGI_CODIGO & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_FILIAL     = CABE.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_CODCONF    = CABE.SGI_CODCONF" & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
                    sSql = sSql & "   And ITEN.SGI_CODORDPROD = ORDP.SGI_CODIGO" & vbCrLf
                    
                    sSql = sSql & "Order By CABE.SGI_DATACONF"
                    
                    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    
                    lngQTDEOP = BREC!SGI_QTDE '' Qtde OP
                    lngCODOP = 0
                    If Not BREC2.EOF() Then
                        Do While Not BREC2.EOF()
                            
                            lngQTDEOP = (lngQTDEOP - BREC2!SGI_QTDREAL)
                            
                            strDADOS03 = BREC2!SGI_QTDREAL & vbTab & _
                                         Format(BREC2!SGI_DATACONF, "DD/MM/YYYY") & vbTab & _
                                         BREC2!SGI_CODFATURA & vbTab & _
                                         lngQTDEOP
                            
                            If lngCODOP <> BREC2!SGI_CODORDPROD Then
                                
                                lngCODOP = BREC2!SGI_CODORDPROD
                                
                                arrDADOSSTEEL(lngQTDREGS) = strDADOS01 & vbTab & _
                                                            strDADOS02 & vbTab & _
                                                            BREC2!SGI_QTDE & vbTab & _
                                                            strDADOS03 & vbTab & _
                                                            BREC2!SGI_CODORDPROD & vbTab & _
                                                            strESTADO & vbTab & _
                                                            BREC!SGI_CIDNORM & vbTab & _
                                                            strOBSOP & vbTab & _
                                                            strFECHAGRAF & vbTab & _
                                                            strVERNCORPO & vbTab & _
                                                            strVERNTAMPA & vbTab & _
                                                            strVERNFUNDO & vbTab & _
                                                            strVERNARGOLA & vbTab & _
                                                            strTOTFAT & vbTab & _
                                                            strSTATUS2
                            
                            Else
                                strDADOS01 = BREC!SGI_CODVEND & vbTab & _
                                             BREC!SGI_NOMEVEND & vbTab & _
                                             BREC!SGI_CODCLI & vbTab & _
                                             BREC!SGI_RAZAOSOC & vbTab & _
                                             BREC!SGI_CODPROD & vbTab & _
                                             BREC!SGI_CODPED & vbTab & _
                                             strSTATUSOP & vbTab & _
                                             Replace(Replace(BREC!SGI_DESCRICAO, "Ç", "C"), "Ã", "A") & vbTab & _
                                             BREC!SGI_CODIGO & vbTab & _
                                             BREC!SGI_CODLIN & vbTab & _
                                             Trim(BREC!SGI_DESCRI) & "." & vbTab & _
                                             BREC!SGI_DESCTIPO & vbTab & _
                                             Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                                             Format((BREC!SGI_DATENTREGA - 7), "DD/MM/YYYY") & vbTab & _
                                             Format((BREC!SGI_DATENTREGA - 3), "DD/MM/YYYY") & vbTab & _
                                             Format((BREC!SGI_DATENTREGA - 2), "DD/MM/YYYY") & vbTab & _
                                             Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             ""
                
                                strDADOS02 = "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             "" & vbTab & _
                                             ""
                            
                                    
                                arrDADOSSTEEL(lngQTDREGS) = strDADOS01 & vbTab & _
                                                            strDADOS02 & vbTab & _
                                                            "" & vbTab & _
                                                            strDADOS03 & vbTab & _
                                                            "" & vbTab & _
                                                            strESTADO & vbTab & _
                                                            BREC!SGI_CIDNORM & vbTab & _
                                                            strOBSOP & vbTab & _
                                                            strFECHAGRAF & vbTab & _
                                                            strVERNCORPO & vbTab & _
                                                            strVERNTAMPA & vbTab & _
                                                            strVERNFUNDO & vbTab & _
                                                            strVERNARGOLA & vbTab & _
                                                            strTOTFAT & vbTab & _
                                                            strSTATUS2
                            End If
                            BREC2.MoveNext
                        Loop
                    Else
                        
                        strDADOS03 = "" & vbTab & _
                                     "" & vbTab & _
                                     "" & vbTab & _
                                     BREC!SGI_QTDE & vbTab & _
                                     "" & vbTab & _
                                     strESTADO & vbTab & _
                                     BREC!SGI_CIDNORM & vbTab & _
                                     strOBSOP & vbTab & _
                                     strFECHAGRAF & vbTab & _
                                     strVERNCORPO & vbTab & _
                                     strVERNTAMPA & vbTab & _
                                     strVERNFUNDO & vbTab & _
                                     strVERNARGOLA & vbTab & _
                                     strTOTFAT & vbTab & _
                                     strSTATUS2
                        
                        arrDADOSSTEEL(lngQTDREGS) = strDADOS01 & vbTab & _
                                                    strDADOS02 & vbTab & _
                                                    BREC!SGI_QTDE & vbTab & _
                                                    strDADOS03
                    End If
                    BREC2.Close
                    '' =======================================
                 
                 BREC.MoveNext
               Loop
        End If
        BREC.Close
    
    End If
    
    
    If boolTemDados = False Then
       MsgBox "ATENÇÃO - Não há dados para gerar o Aquivo !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
     '' Gerando o arquivo em excel
     Call GeraArgExcel(strCamRelNovo & strNomRel, arrDADOSNOVALATA, lngDADOSNOVALATA, arrDADOSSTEEL, lngDADOSSTEEL)
    
     MsgBox "Arquivo Gerado com Exito !!!", vbOKOnly + vbInformation, "Aviso"
    
    
    Frame9.Visible = False
    
    Exit Sub

Err_Exporta:
    
    If BREC.State = 1 Then BREC.Close
    If BREC2.State = 1 Then BREC2.Close
    Close #1
    
    
    If Err.Number = 70 Then
        MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
               "Erro Descr  : " & "Não pode inportar pois o Arguivo TXT eta aberto em outro programa !!!", vbOKOnly + vbCritical, "Aviso"
    
    Else
        MsgBox "Erro Numero : " & Err.Number & vbCrLf & _
               "Erro Descr  : " & Err.Description & vbCrLf & "Numero Regs : " & prgPREP.value, vbOKOnly + vbCritical, "Aviso"
    End If

End Sub



Private Sub GeraArgExcel(strARQUIVO As String, arrNOVALATA As Variant, lngQTDREGSNOVALATA As Long, arrSTEEL As Variant, lngQTDREGSSTEEL As Long)

On Error GoTo err_Excel

    Dim myExcelFile             As New clsExcelFile
    Dim FileName$
    Dim lngLINHA                As Long
    Dim lngREGS                 As Long
    Dim arrDADOS                As Variant
    Dim strDADOS                As String
    
    With myExcelFile
        
        FileName$ = strARQUIVO
        
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
        .SetColumnWidth 1, 1, 40
        .SetColumnWidth 2, 2, 40
        .SetColumnWidth 3, 3, 20
        .SetColumnWidth 4, 4, 70
        .SetColumnWidth 5, 7, 18
        .SetColumnWidth 8, 8, 70
        .SetColumnWidth 9, 13, 18
        .SetColumnWidth 14, 14, 20
        .SetColumnWidth 15, 17, 18
        .SetColumnWidth 18, 18, 20
        .SetColumnWidth 19, 20, 18
        .SetColumnWidth 21, 22, 40
        .SetColumnWidth 23, 34, 70
        .SetColumnWidth 35, 45, 18
        .SetColumnWidth 46, 46, 80
        .SetColumnWidth 47, 52, 18
        .SetColumnWidth 53, 53, 25
        .SetColumnWidth 54, 54, 18
        
        'write some dates to the file. NOTE: you need to write dates as xlsNumber
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 1, "Cód.Vendedor", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 2, "VENDEDOR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 3, "Cód.Clie", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 4, "CLIENTE", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 5, "ROTULO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 6, "Pedido de Venda", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 7, "Status OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 8, "DESCRICAO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 9, "OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 10, "COD.CAPACIDADE", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 11, "CAPACIDADE", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 12, "TIPO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 13, "DATA ENTRADA", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 14, "DATA PREPARACAO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 15, "DATA LITOGRAFIA", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 16, "DATA MONTAGEM", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 17, "DATA ENTREGA", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 18, "QUANTIDADE PEDIDO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 19, "QTDE FOLHAS", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 20, "QTDE POR FOLHA", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 21, "VERNIZ.INT 01", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 22, "VERNIZ.INT 02", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 23, "ESMALTE", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 24, "REVESTIMENTO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 25, "VERNIZ.ACABAMENTO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 26, "PREPARACAO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 27, "1a.COR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 28, "2a.COR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 29, "3a.COR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 30, "4a.COR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 31, "5a.COR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 32, "6a.COR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 33, "7a.COR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 34, "8a.COR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 35, "LITOGRAFIA", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 36, "FECHAMENTO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 37, "Neck IN", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 38, "Qtde.OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 39, "Qtde.Faturada", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 40, "Data Faturada", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 41, "Nota Fiscal", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 42, "Saldo.OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 43, "OP.Faturada", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 44, "Estado.Entrega", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 45, "Cidade Entrega", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 46, "OBS. da OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 47, "Fechamento", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 48, "Verniz CP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 49, "Verniz TP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 50, "Verniz FD", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 51, "Verniz ARG", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 52, "Total.Fat", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 53, "Status Pela Tolerancia", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 54, "Empresa", 12
    
         lngLINHA = 1
         
         
         strDADOS = ""
         '' Dados da Novalata
         If lngQTDREGSNOVALATA > 0 Then
         
             strDADOS = "NOVALATA"
             
             Frame9.Caption = "[ Aguarde.... Gerando o Arguivo EXCEL com os dados da NOVALATA ! ]"
             Frame9.Refresh
             
             prgPREP.Min = 0
             prgPREP.Max = UBound(arrNOVALATA)
             
             For lngREGS = 1 To UBound(arrNOVALATA)
                lngLINHA = (lngLINHA + 1)
                
                
                prgPREP.value = lngREGS
                prgPREP.Refresh
                
                arrDADOS = Split(arrNOVALATA(lngREGS), vbTab)
                
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 1, arrDADOS(0), 1      '' Cód Vencedor
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 2, arrDADOS(1), 12        '' VENDEDOR
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 3, arrDADOS(2), 1      '' Cód. Cliente
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 4, arrDADOS(3), 12        '' CLIENTE
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 5, arrDADOS(4), 12      '' ROTULO
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 6, arrDADOS(5), 1      '' Pedido de Venda
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 7, arrDADOS(6), 12      '' Status OP
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 8, arrDADOS(7), 12        '' DESCRICAO
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 9, arrDADOS(8), 1      '' OP
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 10, arrDADOS(9), 1     '' COD.CAPACIDADE
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 11, arrDADOS(10), 12      '' CAPACIDADE
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 12, arrDADOS(11), 12    '' TIPO
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 13, arrDADOS(12), 12    '' DATA ENTRADA
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 14, arrDADOS(13), 12    '' DATA PREPARACAO
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 15, arrDADOS(14), 12    '' DATA LITOGRAFIA
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 16, arrDADOS(15), 12    '' DATA MONTAGEM
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 17, arrDADOS(16), 12    '' DATA ENTREGA
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 18, arrDADOS(17), 1    '' QUANTIDADE PEDIDO
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 19, arrDADOS(18), 1    '' QTDE FOLHAS
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 20, arrDADOS(19), 1    '' QTDE POR FOLHA
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 21, arrDADOS(20), 1    '' QTDE POR FOLHA
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 22, arrDADOS(21), 12      '' VERNIZ.INT 01
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 23, arrDADOS(22), 12      '' VERNIZ.INT 02
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 24, arrDADOS(23), 12      '' ESMALTE
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 25, arrDADOS(24), 12      '' REVESTIMENTO
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 26, arrDADOS(25), 12      '' VERNIZ.ACABAMENTO
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 27, arrDADOS(26), 12      '' 1a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 28, arrDADOS(27), 12      '' 2a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 29, arrDADOS(28), 12      '' 3a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 30, arrDADOS(29), 12      '' 4a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 31, arrDADOS(30), 12      '' 5a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 32, arrDADOS(31), 12      '' 6a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 33, arrDADOS(32), 12      '' 7a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 34, arrDADOS(33), 12      '' 8a.COR
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 35, arrDADOS(34), 12    '' LITOGRAFIA
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 36, arrDADOS(35), 12    '' FECHAMENTO
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 37, arrDADOS(36), 12    '' Neck IN
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 38, arrDADOS(37), 1    '' Qtde.OP
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 39, arrDADOS(38), 1    '' Qtde.Faturada
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 40, arrDADOS(39), 12    '' Data Faturada
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 41, arrDADOS(40), 1    '' Nota Fiscal
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 42, arrDADOS(41), 1    '' Saldo.OP
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 43, arrDADOS(42), 1    '' OP.Faturada
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 44, arrDADOS(43), 12    '' Estado.Entrega
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 45, arrDADOS(44), 12      '' Cidade Entrega
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 46, arrDADOS(45), 12      '' OBS. da OP
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 47, arrDADOS(46), 12    '' FECHAMENTO
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 48, arrDADOS(47), 12    '' Verniz CP
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 49, arrDADOS(48), 12    '' Verniz TP
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 50, arrDADOS(49), 12    '' Verniz FD
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 51, arrDADOS(50), 12    '' Verniz ARG
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 52, arrDADOS(51), 1    '' Total Fat
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 53, arrDADOS(52), 12      '' Total Fat
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 54, "NOVALATA", 12      '' Empresa
            
             Next lngREGS
         
         End If
        
         '' Dados da STEEL
         If lngQTDREGSSTEEL > 0 Then
         
             strDADOS = "STEEL"
             
             Frame9.Caption = "[ Aguarde.... Gerando o Arguivo EXCEL com os dados da STEEL ! ]"
             Frame9.Refresh
             
             prgPREP.Min = 0
             prgPREP.Max = UBound(arrSTEEL)
             
             For lngREGS = 1 To UBound(arrSTEEL)
                lngLINHA = (lngLINHA + 1)
                
                prgPREP.value = lngREGS
                prgPREP.Refresh
                
                arrDADOS = Split(arrSTEEL(lngREGS), vbTab)
                
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 1, arrDADOS(0), 1      '' Cód. Vendedor
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 2, arrDADOS(1), 12        '' VENDEDOR
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 3, arrDADOS(2), 1      '' Cód. Cliente
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 4, arrDADOS(3), 12        '' CLIENTE
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 5, arrDADOS(4), 12      '' ROTULO
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 6, arrDADOS(5), 1      '' Pedido de Venda
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 7, arrDADOS(6), 12      '' Status OP
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 8, arrDADOS(7), 12        '' DESCRICAO
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 9, arrDADOS(8), 1      '' OP
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 10, arrDADOS(9), 1      '' COD.CAPACIDADE
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 11, arrDADOS(10), 12      '' CAPACIDADE
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 12, arrDADOS(11), 12    '' TIPO
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 13, arrDADOS(12), 12    '' DATA ENTRADA
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 14, arrDADOS(13), 12    '' DATA PREPARACAO
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 15, arrDADOS(14), 12    '' DATA LITOGRAFIA
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 16, arrDADOS(15), 12    '' DATA MONTAGEM
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 17, arrDADOS(16), 12    '' DATA ENTREGA
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 18, arrDADOS(17), 1    '' QUANTIDADE PEDIDO
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 19, arrDADOS(18), 1    '' QTDE FOLHAS
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 20, arrDADOS(19), 1    '' QTDE POR FOLHA
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 21, arrDADOS(20), 1    '' QTDE POR FOLHA
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 22, arrDADOS(21), 12      '' VERNIZ.INT 01
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 23, arrDADOS(22), 12      '' VERNIZ.INT 02
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 24, arrDADOS(23), 12      '' ESMALTE
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 25, arrDADOS(24), 12      '' REVESTIMENTO
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 26, arrDADOS(25), 12      '' VERNIZ.ACABAMENTO
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 27, arrDADOS(26), 12      '' 1a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 28, arrDADOS(27), 12      '' 2a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 29, arrDADOS(28), 12      '' 3a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 30, arrDADOS(29), 12      '' 4a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 31, arrDADOS(30), 12      '' 5a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 32, arrDADOS(31), 12      '' 6a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 33, arrDADOS(32), 12      '' 7a.COR
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 34, arrDADOS(33), 12      '' 8a.COR
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 35, arrDADOS(34), 12    '' LITOGRAFIA
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 36, arrDADOS(35), 12    '' FECHAMENTO
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 37, arrDADOS(36), 12    '' Neck IN
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 38, arrDADOS(37), 1    '' Qtde.OP
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 39, arrDADOS(38), 1    '' Qtde.Faturada
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 40, arrDADOS(39), 12    '' Data Faturada
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 41, arrDADOS(40), 1    '' Nota Fiscal
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 42, arrDADOS(41), 1    '' Saldo.OP
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 43, arrDADOS(42), 1    '' OP.Faturada
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 44, arrDADOS(43), 12    '' Estado.Entrega
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 45, arrDADOS(44), 12      '' Cidade Entrega
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 46, arrDADOS(45), 12      '' OBS. da OP
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 47, arrDADOS(46), 12    '' FECHAMENTO
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 48, arrDADOS(47), 12    '' Verniz CP
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 49, arrDADOS(48), 12    '' Verniz TP
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 50, arrDADOS(49), 12    '' Verniz FD
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 51, arrDADOS(50), 12    '' Verniz ARG
                .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 52, arrDADOS(51), 1    '' Total Fat
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 53, arrDADOS(52), 12      '' Total Fat
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 54, "STEEL", 12         '' Empresa
            
             Next lngREGS
         
         End If
        
        
        .ProtectSpreadsheet = False 'False | True
        .CloseFile
    
    End With

    Exit Sub

err_Excel:

    MsgBox "ATENÇÃO" & vbCrLf & _
           "Erro Numero       : " & Err.Number & vbCrLf & _
           "Descrição do Erro : " & Err.Description & vbCrLf & _
           "Linha : " & lngREGS & " Dados : " & strDADOS, vbOKOnly + vbCritical, "Aviso"

End Sub
