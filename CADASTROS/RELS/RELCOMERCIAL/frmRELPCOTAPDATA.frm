VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRELPCOTAPDATA 
   Caption         =   "Gera P.Cota / P.Data"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar prgPREP 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
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
      Left            =   5400
      TabIndex        =   8
      Top             =   960
      Width           =   3015
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
         TabIndex        =   10
         Top             =   240
         Width           =   1335
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
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   5415
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
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
      Width           =   8415
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
         Picture         =   "frmRELPCOTAPDATA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmRELPCOTAPDATA.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELPCOTAPDATA"
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
Dim objRELPCOTAPDATA    As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object

Dim strCABEC1           As String
Dim strCABEC2           As String

Dim lngPORC             As Long

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call ImpRel
End Sub

Private Sub cmdVoltar_Click()
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
    Set objRELPCOTAPDATA = CreateObject("RELPCP.clsRELPCOTAPDATA")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    objRELPCOTAPDATA.FILIAL = FILIAL

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7) & "RELPCOTAPDATA\"
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Now, "DD/MM/YYYY")

    optFilial(0).value = True
    prgPREP.Min = 0
    prgPREP.Visible = False

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELPCOTAPDATA = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
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
    Dim boolTEMDADOS    As Boolean
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
    
    prgPREP.Min = 0
    prgPREP.Visible = True
    
    prgPREP.Visible = True
    If optFilial(1).value = True Then
        strNomRel = "RELPCOTAPDATA01_STEEL.TXT"
    ElseIf optFilial(0).value = True Then
        strNomRel = "RELPCOTAPDATA01_NOVA.TXT"
    End If
    
    boolTEMDADOS = True
    
    strNOMFILIAL = ""
    If optFilial(1).value = True Then strNOMFILIAL = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       VEND.SGI_DESCRICAO As SGI_NOMEVEND" & vbCrLf
    
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_ESTNORM" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_CIDNORM" & vbCrLf
     
    sSql = sSql & "      ,PEDVH.SGI_DATAPED" & vbCrLf
    
    sSql = sSql & "      ,PEDVI.SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,PEDVI.SGI_FECHTPFU" & vbCrLf
    
    sSql = sSql & "      ,PROGE.SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,PROGE.SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "      ,PROGE.SGI_QTDE" & vbCrLf
    sSql = sSql & "      ,PROGE.SGI_OBSOP" & vbCrLf
    sSql = sSql & "      ,PROGE.SGI_STATUS" & vbCrLf

    sSql = sSql & "      ,PROD.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO As SGI_DESCPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_QTDEPORFOLHA" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_QTDCORPSPADRAOSN" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_NECKIN" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_FechSoldaAgrafado" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernCorpo" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernTampa" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernFundo" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernArgola" & vbCrLf
    
    sSql = sSql & "      ,LINHA.SGI_CODIGO As SGI_CODLIN" & vbCrLf
    sSql = sSql & "      ,LINHA.SGI_DESCRI As SGI_DESCRILIN" & vbCrLf
    sSql = sSql & "      ,LINHA.SGI_CODLIN As SGI_LINHA" & vbCrLf
    sSql = sSql & "      ,LINHA.SGI_QTDECORPOS" & vbCrLf
    sSql = sSql & "      ,LINHA.SGI_PERDPROC" & vbCrLf
    
    sSql = sSql & "      ,TIP.SGI_DESCRICAO As SGI_DESCTIPO" & vbCrLf
    sSql = sSql & "      ,ESP.SGI_DESCRICAO As SGI_DESCESP" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI" & strNOMFILIAL & " PEDVI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDVH" & vbCrLf
    sSql = sSql & "      ,SGI_CADVENDEDOR       VEND" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE        CLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO        PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO   LINHA" & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPPROD        TIP" & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD        ESP" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PROGE.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROGE.SGI_DATENTREGA Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And (PROGE.SGI_STATUS    = 6 or PROGE.SGI_STATUS    = 7)" & vbCrLf
    
    sSql = sSql & "   And PEDVI.SGI_FILIAL    = PROGE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PEDVI.SGI_CODIGO    = PROGE.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And PEDVI.SGI_IDPRODUTO = PROGE.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And PEDVH.SGI_FILIAL    = PEDVI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PEDVH.SGI_CODIGO    = PEDVI.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And (PEDVH.SGI_STATUS   = 'C' or PEDVH.SGI_STATUS = '4')" & vbCrLf
    
    sSql = sSql & "   And VEND.SGI_FILIAL     = PEDVH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And VEND.SGI_CODIGO     = PEDVH.SGI_CODVEND" & vbCrLf
    
    sSql = sSql & "   And CLIE.SGI_FILIAL     = PEDVH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO     = PEDVH.SGI_CODCLI " & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL     = PROGE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = PROGE.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And LINHA.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINHA.SGI_CODLIN    = PROD.SGI_CODLINPROD" & vbCrLf
    
    sSql = sSql & "   And TIP.SGI_FILIAL      = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And TIP.SGI_CODIGO      = PROD.SGI_CODTIPO" & vbCrLf
    
    sSql = sSql & "   And ESP.SGI_FILIAL      = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ESP.SGI_CODIGO      = PROD.SGI_CODESPECIE" & vbCrLf
    
    sSql = sSql & "Order By PROGE.SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "       ,PROGE.SGI_CODPED"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
    
        lngQTDREGS = 0
        Do While Not BREC.EOF()
            lngQTDREGS = lngQTDREGS + 1
            BREC.MoveNext
        Loop
        BREC.MoveFirst
    
        prgPREP.Max = lngQTDREGS
    
       Open strCamRelNovo & strNomRel For Output As #1
       
       strCAMPO01 = "VENDEDOR" & vbTab & _
                    "CLIENTE" & vbTab & _
                    "ROTULO" & vbTab & _
                    "Pedido de Venda" & vbTab & _
                    "DESCRICAO" & vbTab & _
                    "COD.CAPACIDADE" & vbTab & _
                    "CAPACIDADE" & vbTab & _
                    "TIPO" & vbTab & _
                    "DATA PEDIDO" & vbTab & _
                    "DATA ENTREGA" & vbTab & _
                    "QUANTIDADE PEDIDO" & vbTab & _
                    "QTDE FOLHAS" & vbTab & _
                    "QTDE POR FOLHA" & vbTab & _
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
                     
        strCAMPO03 = "Estado.Entrega" & vbTab & _
                     "Cidade Entrega" & vbTab & _
                     "Observação" & vbTab & _
                     "Fechamento" & vbTab & _
                     "Verniz CP" & vbTab & _
                     "Verniz TP" & vbTab & _
                     "Verniz FD" & vbTab & _
                     "Verniz ARG" & vbTab & _
                     "Status"
                     
       Print #1, strCAMPO01 & vbTab & _
                 strCAMPO02 & vbTab & _
                 strCAMPO03
    
        lngQTDREGS = 0
        DoEvents
        
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
            
            
            '' Pegava o Estado de Entrega
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
            
            '' Dados
            strDADOS01 = BREC!SGI_NOMEVEND & vbTab & _
                         BREC!SGI_RAZAOSOC & vbTab & _
                         BREC!SGI_CODPROD & vbTab & _
                         BREC!SGI_CODPED & vbTab & _
                         Replace(Replace(BREC!SGI_DESCPROD, "Ç", "C"), "Ã", "A") & vbTab & _
                         BREC!SGI_LINHA & vbTab & _
                         Trim(BREC!SGI_DESCRILIN) & "." & vbTab & _
                         BREC!SGI_DESCTIPO & vbTab & _
                         Format(BREC!SGI_DATAPED, "DD/MM/YYYY") & vbTab & _
                         Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY") & vbTab & _
                         BREC!SGI_QTDE & vbTab & _
                         lngQTDEFOLHAS & vbTab & _
                         lngQTDFOLHAS & vbTab & _
                         strVERNIZ01 & vbTab & _
                         strVERNIZ02

            strDADOS02 = strESMALTE & vbTab & _
                         BREC!SGI_DESCESP & vbTab & _
                         strVERNIZACAB & vbTab & _
                         arrCORES(1) & vbTab & _
                         arrCORES(2) & vbTab & _
                         arrCORES(3) & vbTab & _
                         arrCORES(4) & vbTab & _
                         arrCORES(5) & vbTab & _
                         arrCORES(6) & vbTab & _
                         arrCORES(7) & vbTab & _
                         arrCORES(8) & vbTab & _
                         Fechamento(intCODFECHA) & vbTab & _
                         strNECKIN

            strDADOS03 = strESTADO & vbTab & _
                         BREC!SGI_CIDNORM & vbTab & _
                         strOBSOP & vbTab & _
                         strFECHAGRAF & vbTab & _
                         strVERNCORPO & vbTab & _
                         strVERNTAMPA & vbTab & _
                         strVERNFUNDO & vbTab & _
                         strVERNARGOLA & vbTab & _
                         PegaStatus(BREC!SGI_STATUS)
            
            Print #1, strDADOS01 & vbTab & _
                      strDADOS02 & vbTab & _
                      strDADOS03
            
            BREC.MoveNext
       Loop
       
       Close #1
    
       prgPREP.Visible = False
       MsgBox "Arquivo Gerado com Exito !!!", vbOKOnly + vbInformation, "Aviso"
    
    Else
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       boolTEMDADOS = False
    End If
    BREC.Close
    
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


Private Function Pega_Estado(lngINDICE As Long) As String

   Pega_Estado = ""

   Dim V_Estado As Variant
   
   V_Estado = Array("AM", "AC", "AL", "AP", "BA", "CE", "DF", "ES", _
                    "GO", "MA", "MG", "MT", "MS", "PE", "PA", "PB", "PI", "PR", "RJ", _
                    "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO", "EX")

   If (lngINDICE - 1) >= 0 Then Pega_Estado = V_Estado((lngINDICE - 1))

End Function


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

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, Me
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, Me
End Sub

Private Function PegaStatus(lngSTATUS As Long) As String

    PegaStatus = ""
    
    If lngSTATUS = 6 Then PegaStatus = "P.Cota"
    If lngSTATUS = 7 Then PegaStatus = "P.Data"

End Function

