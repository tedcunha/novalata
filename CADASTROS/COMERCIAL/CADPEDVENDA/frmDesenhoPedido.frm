VERSION 5.00
Begin VB.Form frmDesenhoPedido 
   Caption         =   "Desenho do Rótulo"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   17130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   17055
      Begin VB.Image Assinatura 
         Height          =   6615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   16845
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17055
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
         Picture         =   "frmDesenhoPedido.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmDesenhoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngIDProduto     As Long
Public strDescProduto   As String

Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public FILIAL           As Integer
Public strACESSO        As String
Public strMODPAI        As String
Public strUSUARIO       As String
Dim objBLBFunc          As Object


Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
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
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   Call CarregaImagem
   
   Frame2.Caption = "[ Produto : " & Trim(strDescProduto) & "]"

End Sub

Private Sub CarregaImagem()
    
    Dim strNOMARQ   As String
    Dim strCAMINHO  As String
    
    sSql = "Select SGI_ARQUIVOIMG from SGI_CADPRODUTO Where SGI_FILIAL = " & FILIAL & " And SGI_IDPRODUTO = " & lngIDProduto
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic, adLockOptimistic
    If Not BREC2.EOF Then
       If Not IsNull(BREC2!SGI_ARQUIVOIMG) Then
          strCAMINHO = BREC2!SGI_ARQUIVOIMG
          If Len(Trim(strCAMINHO)) > 0 Then Assinatura.Picture = LoadPicture(strCAMINHO)
       End If
    End If
    BREC2.Close

End Sub

