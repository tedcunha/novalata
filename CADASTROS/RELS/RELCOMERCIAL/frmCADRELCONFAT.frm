VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADRELCONFAT 
   Caption         =   "Relatório de Confirmação de Faturamento"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   6495
      Begin MSMask.MaskEdBox mskDataIni 
         Height          =   285
         Left            =   1200
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
      Begin MSMask.MaskEdBox mskDataFin 
         Height          =   285
         Left            =   3480
         TabIndex        =   7
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   975
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
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
         Picture         =   "frmCADRELCONFAT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADRELCONFAT.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADRELCONFAT"
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
Dim objRELCONFAT        As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String

Private Sub cmdImpressao_Click()
    If ValidaCampos = False Then Exit Sub
    Call ImpRelConFat
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELCONFAT = Nothing
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
    Set objRELCONFAT = CreateObject("RELCOMERCIAL.clsRELCONFAT")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELCONFAT.FILIAL = FILIAL
    
    mskDataIni.Text = Format(Now, "DD/MM/YYYY")
    mskDataFin.Text = Format(Now, "DD/MM/YYYY")

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub mskDataFin_GotFocus()
    objBLBFunc.SelecionaCampos mskDataFin.Name, frmCADRELCONFAT
End Sub

Private Sub mskDataIni_GotFocus()
    objBLBFunc.SelecionaCampos mskDataIni.Name, frmCADRELCONFAT
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = False
        
        
    If Not IsDate(mskDataIni.Text) Then
        MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDataIni.SetFocus
        Exit Function
    End If
    If Not IsDate(mskDataFin.Text) Then
        MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDataFin.SetFocus
        Exit Function
    End If
    
    If CDate(mskDataIni.Text) > CDate(mskDataFin.Text) Then
        MsgBox "Data Inicial não pode ser maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDataIni.SetFocus
        Exit Function
    End If
    
    ValidaCampos = True
End Function


Private Sub ImpRelConFat()

    Dim strNomRel As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "       SGI_CADORDCONFH.SGI_CODFATURA " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFH.SGI_DATACONF " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFH.SGI_VALTOTFAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFH.SGI_TOTALIPI " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFH.SGI_CODORD " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATH.SGI_CODORD " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATH.SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODCLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDCONFH SGI_CADORDCONFH " & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATH SGI_CADORDFATH " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH SGI_CADPEDVENDH " & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "      SGI_CADORDCONFH.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "  And (SGI_CADORDCONFH.SGI_DATACONF >= '" & Format(CDate(mskDataIni.Text), "MM/DD/YYYY") & "' And SGI_CADORDCONFH.SGI_DATACONF <= '" & Format(CDate(mskDataFin.Text), "MM/DD/YYYY") & "')" & vbCrLf
    
    sSql = sSql & "  And SGI_CADORDFATH.SGI_FILIAL     = SGI_CADORDCONFH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SGI_CADORDFATH.SGI_CODORD     = SGI_CADORDCONFH.SGI_CODORD " & vbCrLf
    
    sSql = sSql & "  And SGI_CADPEDVENDH.SGI_FILIAL    = SGI_CADORDFATH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SGI_CADPEDVENDH.SGI_CODIGO    = SGI_CADORDFATH.SGI_CODPED " & vbCrLf
    
    sSql = sSql & "  And SGI_CADPEDVENDH.SGI_FILIAL    = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SGI_CADPEDVENDH.SGI_CODCLI    = SGI_CADCLIENTE.SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Relatório de Confirmação de faturamento"
    
    strNomRel = "RELORDCONFH.rpt"
    
    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, False)
    End If
    
End Sub

