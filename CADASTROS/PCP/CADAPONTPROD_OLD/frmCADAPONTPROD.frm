VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADAPONTPROD 
   Caption         =   "Apontamento de Produção"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   6840
      Width           =   4695
      Begin VB.OptionButton optFechaSN 
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
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   22
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton optFechaSN 
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
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   21
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha o Apontamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   2175
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid grdProcesso 
      Height          =   4455
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   10215
      _cx             =   18018
      _cy             =   7858
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   10215
      Begin VB.TextBox txtCODPROD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "txtCODPROD"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   2520
         Picture         =   "frmCADAPONTPROD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDescLinha 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescLinha"
         Height          =   285
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "Fluxo Produtivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   10215
      Begin VB.TextBox txtCODOP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   0
         Text            =   "txtCODOP"
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskDtDoc 
         Height          =   285
         Left            =   6840
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
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDescStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescStatus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8760
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   8160
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.OP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "N° Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Data do Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdAltera 
         Caption         =   "&Altera"
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
         Picture         =   "frmCADAPONTPROD.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
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
         Left            =   1800
         Picture         =   "frmCADAPONTPROD.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADAPONTPROD.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADAPONTPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCodVendedor   As Long
Public lngCodUsuario    As Long
Dim lngCodLog           As Long
Dim strCAPTION          As String
Dim arrITENSAPONT       As Variant

Dim objBLBFunc          As Object
Dim objCADAPONTPROD     As Object
Dim objPESQPADRAO       As Object

Const conCOL_SonProc_Ordem                     As Integer = 0
Const conCOL_SonProc_CodProc                   As Integer = 1
Const conCOL_SonProc_Desc                      As Integer = 2
Const conCOL_SonProc_HorIni                    As Integer = 3
Const conCOL_SonProc_HorFin                    As Integer = 4
Const conCOL_SonProc_Total                     As Integer = 5
Const conCOL_SonProc_FormatString              As String = "=Ordem|Cod.Processo|Descrição do Processo|Hora.Inicial|Hora.Final|Final"
Const conColumnsIn_SonProc                     As Integer = 6



Private Sub cmdAltera_Click()

    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(frmCADAPONTPROD, strCAPTION, cTipOper)

End Sub

Private Sub CmdSalva_Click()

On Error GoTo er_Salva

    Dim I As Integer
    
    Call objBLBFunc.RemoveLinhaVazia(grdProcesso, conCOL_SonProc_CodProc)
    
    If Verif_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADAPONTPROD.CODIGO = objBLBFunc.Gera_Codigo(Me.Name, FILIAL, Linha)
    objCADAPONTPROD.CODOP = CLng(txtCODOP.Text)
    objCADAPONTPROD.CODFP = CLng(txtCODPROD.Text)
    objCADAPONTPROD.DATDOC = Format(mskDtDoc.Text, "MM/DD/YYYY")
    If optFechaSN(1).Value = True Then objCADAPONTPROD.FECHAAPONT = 1
    If optFechaSN(0).Value = True Then objCADAPONTPROD.FECHAAPONT = 0

    If objCADAPONTPROD.FECHAAPONT = 0 Then objCADAPONTPROD.STATUS = 0
    If objCADAPONTPROD.FECHAAPONT = 1 Then objCADAPONTPROD.STATUS = 1

    '' Itens de Apontamento
    arrITENSAPONT = Empty
    With grdProcesso
        objCADAPONTPROD.QTDEAPONT = (.Rows - 1)
        If (.Rows - 1) > 0 Then
            ReDim arrITENSAPONT(1 To (.Rows - 1), 1 To 4) As String
            For I = 1 To (.Rows - 1)
                arrITENSAPONT(I, 1) = .Cell(flexcpText, I, conCOL_SonProc_CodProc)
                
                arrITENSAPONT(I, 2) = "Null"
                arrITENSAPONT(I, 3) = "Null"
                arrITENSAPONT(I, 4) = "Null"
                
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonProc_HorIni))) > 0 Then arrITENSAPONT(I, 2) = objBLBFunc.CONVHRMIN(.Cell(flexcpText, I, conCOL_SonProc_HorIni))
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonProc_HorFin))) > 0 Then arrITENSAPONT(I, 3) = objBLBFunc.CONVHRMIN(.Cell(flexcpText, I, conCOL_SonProc_HorFin))
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonProc_Total))) > 0 Then arrITENSAPONT(I, 4) = objBLBFunc.CONVHRMIN(.Cell(flexcpText, I, conCOL_SonProc_Total))
            Next I
        End If
    End With
    objCADAPONTPROD.ITENSAPONT = arrITENSAPONT


    '' Gravando as Informações no banco
    If objCADAPONTPROD.GRAVA(cTipOper) = False Then Exit Sub
    
    '' Atualizando os Dados
    If objBLBFunc.Atualiza(cTipOper, Str(objCADAPONTPROD.CODIGO), FILIAL, Me.Name, Linha) = False Then Exit Sub
    
    '' Gerando Log de Sistema
    lngCodLog = objBLBFunc.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objBLBFunc.GravaLogModulo(FILIAL, lngCodLog, Me.Name, cTipOper, lngCodUsuario, Str(objCADAPONTPROD.CODIGO), Linha)
    
    MsgBox "O Apontamento ( " & Trim(Str(objCADAPONTPROD.CODIGO)) & " ) foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    Unload Me

    Exit Sub
    
er_Salva:

    MsgBox "Erro N°   : " & Err.Number & vbCrLf & _
           "Descrição : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"

End Sub

Private Sub Command8_Click()

    If ConsisteFlx = False Then Exit Sub
    
    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       HEAD.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,HEAD.SGI_CODPROD " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRI " & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD       ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO      PRD" & vbCrLf
    sSql = sSql & "      ,SGI_CADFLUXPROD     HEAD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO PROD" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO   = " & Trim(txtCODOP.Text) & vbCrLf
    sSql = sSql & "   And PRD.SGI_FILIAL    = ORDP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PRD.SGI_IDPRODUTO = ORDP.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "   And HEAD.SGI_FILIAL   = PRD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And HEAD.SGI_CODPROD  = PRD.SGI_CODLINPROD " & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL   = HEAD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODIGO   = HEAD.SGI_IDPRODUTO " & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Cõdigo"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "HEAD.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CODPROD"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Cod.Linha"
    arrCAMPOS(2, 4) = "1000"
    arrCAMPOS(2, 5) = "HEAD.SGI_CODPROD"
    
    arrCAMPOS(3, 1) = "SGI_DESCRI"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Descrição"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "PROD.SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCODPROD.Text = varRETORNO
        lblDescLinha.Caption = DescLinhaProd(varRETORNO, "SGI_CODIGO")
        Call PopGrdProc(txtCODPROD.Text)
    End If
    txtCODPROD.SetFocus

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

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADAPONTPROD = CreateObject("CADAPONTPROD.clsCADAPONTPROD")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADAPONTPROD.FILIAL = FILIAL
   
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
    strCAPTION = Me.Caption & " - "
   
    Call IniciaForm


End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADAPONTPROD = Nothing
    Set objPESQPADRAO = Nothing
End Sub


Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(frmCADAPONTPROD, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(frmCADAPONTPROD)
    
    Call DesabilitaCampos(Trim(cTipOper))
    Call LimpaCamposLabel
    Call InitGridProc
    
    optFechaSN(0).Value = True
    
    objCADAPONTPROD.CODIGO = iCodigo
    If cTipOper = "I" Then mskDtDoc.Text = Format(Now, "DD/MM/YYYY")
    
    Call CarregaCampos
    Call IqualaCampoStatus
    
    
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Then
        Frame2.Enabled = True
        Frame3.Enabled = True
    ElseIf strTipOper = "C" Or strTipOper = "A" Then
        Frame2.Enabled = False
        Frame3.Enabled = False
    End If
End Sub

Private Sub LimpaCamposLabel()
    lblDescLinha.Caption = ""
    lblDescStatus.Caption = ""
End Sub


Private Function DescLinhaProd(strCodLinha As String, strCampo As String) As String
    
    DescLinhaProd = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       HEAD.*" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRI" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADFLUXPROD HEAD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO PROD" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       HEAD.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And HEAD." & strCampo & " = " & strCodLinha & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL = HEAD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODIGO = HEAD.SGI_IDPRODUTO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       DescLinhaProd = Trim(BREC!SGI_CODPROD) & " - " & Trim(BREC!SGI_DESCRI)
       txtCODPROD.Text = Trim(BREC!SGI_CODIGO)
    End If
    BREC.Close
    
End Function


Private Sub grdProcesso_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strTOTALPERIODO As String
    Dim dtTotalLiquido  As Date
    Dim lngMinutos      As Long
    
    With grdProcesso
        Select Case Col
            Case conCOL_SonProc_HorIni, _
                 conCOL_SonProc_HorFin
                 
                 If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonProc_HorIni), ":", ""))) > 0 And _
                    Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonProc_HorFin), ":", ""))) > 0 Then
                    
                    strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, .Row, conCOL_SonProc_HorIni), .Cell(flexcpText, .Row, conCOL_SonProc_HorFin))
                    
                    dtTotalLiquido = CDate(strTOTALPERIODO)
                    .Cell(flexcpText, .Row, conCOL_SonProc_Total) = Format(dtTotalLiquido, "HH:MM")
                    
                 End If
        End Select
    End With
End Sub

Private Sub grdProcesso_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonProc_Ordem, _
         conCOL_SonProc_CodProc, _
         conCOL_SonProc_Desc
         Cancel = True
    Case conCOL_SonProc_HorIni, conCOL_SonProc_HorFin
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdProcesso.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdProcesso_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim intHoras As Integer
    Dim intMinutos As Integer
    
    With grdProcesso
        Select Case Col
            Case conCOL_SonProc_HorIni, _
                 conCOL_SonProc_HorFin
                 If .EditText = "  :  " Then
                    Cancel = False
                    Exit Sub
                 End If
                 If Len(Trim(.EditText)) = 0 Then
                    Cancel = False
                    Exit Sub
                 End If
                 
                 '' ==================================
                 '' Validando Campo Horas
                 If Len(Trim(Mid(.EditText, 1, 2))) = 0 And Len(Trim(Mid(.EditText, 4, 2))) > 0 Then
                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 If Len(Trim(Mid(.EditText, 1, 2))) > 0 And Len(Trim(Mid(.EditText, 4, 2))) = 0 Then
                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 
                 
                 '' ==================================
                 '' Validando Horas
                 intHoras = CInt(Mid(.EditText, 1, 2))
                 intMinutos = CInt(Mid(.EditText, 4, 2))
                 If intHoras >= 24 Or intHoras < 0 Then
                    MsgBox "Hora Inválida o Dia vai somente até 24:00 !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 If intMinutos >= 60 Then
                    MsgBox "Minutos Inválido os minutos somente devem ser informados de 00 a 59 !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 
                 If Col = conCOL_SonProc_HorIni Then
                    If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonProc_HorFin))) > 0 Then
                        If CDate(.Cell(flexcpText, .Row, conCOL_SonProc_HorFin)) < CDate(.EditText) Then
                            MsgBox "A Hora final não pode ser maior que a Hora Inicial !!!", vbOKOnly + vbCritical, "Aviso"
                            Cancel = True
                        End If
                    End If
                 ElseIf Col = conCOL_SonProc_HorFin Then
                    If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonProc_HorIni))) > 0 Then
                        If CDate(.Cell(flexcpText, .Row, conCOL_SonProc_HorIni)) > CDate(Replace(.EditText, ";", ":")) Then
                            MsgBox "A Hora inicial não pode ser maior que a Hora final !!!", vbOKOnly + vbCritical, "Aviso"
                            Cancel = True
                        End If
                    End If
                 End If
        
        End Select
    End With
End Sub

Private Sub mskDtDoc_GotFocus()
    objBLBFunc.SelecionaCampos mskDtDoc.Name, frmCADAPONTPROD
End Sub

Private Sub txtCODOP_GotFocus()
    objBLBFunc.SelecionaCampos txtCODOP.Name, frmCADAPONTPROD
End Sub

Private Sub txtCODOP_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODOP.Text
End Sub

Private Sub txtCODOP_Validate(Cancel As Boolean)
    Cancel = PegaOP(txtCODOP.Text)
End Sub

Private Sub txtCODPROD_GotFocus()
    objBLBFunc.SelecionaCampos txtCODPROD.Name, frmCADAPONTPROD
End Sub

Private Sub txtCODPROD_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODPROD.Text
End Sub

Private Sub txtCODPROD_Validate(Cancel As Boolean)

    If Len(Trim(txtCODPROD.Text)) = 0 Then Exit Sub
   
    If ConsisteFlx = False Then
        Cancel = True
        Exit Sub
    End If
    
    lblDescLinha.Caption = DescLinhaProd(txtCODPROD.Text, "SGI_CODIGO")
    If Len(Trim(lblDescLinha.Caption)) = 0 Then
       txtCODPROD.Text = ""
       MsgBox "Este fluxo não existe não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If
    Call PopGrdProc(txtCODPROD.Text)
   
End Sub



Private Function ConsisteFlx() As Boolean
    ConsisteFlx = False
    If Len(Trim(txtCODOP.Text)) = 0 Then
        MsgBox "Primeiro informe a OP !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODOP.SetFocus
        Exit Function
    End If
    ConsisteFlx = True
End Function


Private Sub InitGridProc()

    With grdProcesso
    
       .Cols = conColumnsIn_SonProc
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProc_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProc_Ordem) = ""
       .ColDataType(conCOL_SonProc_Ordem) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProc_CodProc) = ""
       .ColDataType(conCOL_SonProc_CodProc) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProc_Desc) = ""
       .ColDataType(conCOL_SonProc_Desc) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProc_HorIni) = ""
       .ColDataType(conCOL_SonProc_HorIni) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProc_HorFin) = ""
       .ColDataType(conCOL_SonProc_HorFin) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProc_Total) = ""
       .ColDataType(conCOL_SonProc_Total) = flexDTString
       
       .ColWidth(conCOL_SonProc_Ordem) = 800
       .ColWidth(conCOL_SonProc_CodProc) = 1200
       .ColWidth(conCOL_SonProc_Desc) = 3500
       .ColWidth(conCOL_SonProc_HorIni) = 1500
       .ColWidth(conCOL_SonProc_HorFin) = 1500
       .ColWidth(conCOL_SonProc_Total) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PopGrdProc(strCodProc As String)

    If Len(Trim(strCodProc)) = 0 Then Exit Sub

    Call InitGridProc
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "        FLX.SGI_INDICE" & vbCrLf
    sSql = sSql & "       ,FLX.SGI_CODPROC" & vbCrLf
    sSql = sSql & "       ,PRC.SGI_DESCRI" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADFLXPRODPROCESSO FLX" & vbCrLf
    sSql = sSql & "       ,SGI_CADPROCESSO PRC" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        FLX.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "    And FLX.SGI_CODIGO = " & Trim(strCodProc) & vbCrLf
    sSql = sSql & "    And PRC.SGI_FILIAL = FLX.SGI_FILIAL" & vbCrLf
    sSql = sSql & "    And PRC.SGI_CODIGO = FLX.SGI_CODPROC" & vbCrLf
    sSql = sSql & " Order By FLX.SGI_INDICE"

    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC7.EOF() Then
        With grdProcesso
            Do While Not BREC7.EOF()
                .AddItem BREC7!SGI_INDICE & vbTab & _
                         BREC7!SGI_CODPROC & vbTab & _
                         Trim(BREC7!SGI_DESCRI) & vbTab & _
                         "" & vbTab & _
                         ""
                BREC7.MoveNext
            Loop
        End With
    End If
    BREC7.Close
    
End Sub

Private Function PegaOP(strCodOP As String) As Boolean
    
    PegaOP = False
    
    If Len(Trim(strCodOP)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodOP
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If IsNull(BREC!SGI_OPENVIADA) Then
            MsgBox "Atenção - Esta OP não foi enviada para produção !!!", vbOKOnly + vbExclamation, "Aviso"
            PegaOP = True
        ElseIf BREC!SGI_OPENVIADA = 2 Then
            MsgBox "Atenção - Esta OP esta sendo apontada !!!", vbOKOnly + vbExclamation, "Aviso"
            PegaOP = True
        ElseIf BREC!SGI_OPENVIADA = 3 Then
            MsgBox "Atenção - Esta OP já foi apontada !!!", vbOKOnly + vbExclamation, "Aviso"
            PegaOP = True
        End If
    Else
        MsgBox "Esta OP Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
        PegaOP = True
    End If
    BREC.Close
    
End Function

Private Function VerifCampo(strCampoIni As String, strCampoFin As String) As Boolean
    VerifCampo = False
    
    If Len(Trim(strCampoIni)) = 0 Then Exit Function
    If Len(Trim(strCampoFin)) = 0 Then Exit Function
    
    If CDate(strCampoIni) > CDate(strCampoFin) Then
        MsgBox "A hora inicial não pode ser maior que hora final !!!", vbOKOnly + vbExclamation, "Aviso"
        VerifCampo = True
        Exit Function
    End If
    
End Function

Private Sub IqualaCampoStatus()
    If cTipOper = "I" Then objCADAPONTPROD.STATUS = 0
    If objCADAPONTPROD.STATUS = 0 Then lblDescStatus.Caption = "Aberto"
    If objCADAPONTPROD.STATUS = 1 Then lblDescStatus.Caption = "Fechado"
End Sub

Private Function Verif_Campos() As Boolean
    
    Dim I               As Integer
    Dim lngTotProc      As Long
    Dim lngQTDEAPONT    As Long
    
    Verif_Campos = False
    
    With grdProcesso
        lngTotProc = (.Rows - 1)
        lngQTDEAPONT = 0
        For I = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, I, conCOL_SonProc_Total))) > 0 Then lngQTDEAPONT = (lngQTDEAPONT + 1)
        Next I
        If lngQTDEAPONT = lngTotProc Then optFechaSN(1).Value = True
        If lngQTDEAPONT = 0 Then
            MsgBox "Não foi informado nenhum apontamento !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End With
    
    Verif_Campos = True
End Function


Private Sub CarregaCampos()

    If objCADAPONTPROD.Carrega_Campos = True Then
        
        txtCodigo.Text = objCADAPONTPROD.CODIGO
        txtCODOP.Text = objCADAPONTPROD.CODOP
        mskDtDoc.Text = CDate(objCADAPONTPROD.DATDOC)
        optFechaSN(objCADAPONTPROD.FECHAAPONT).Value = True
        
        txtCODPROD.Text = objCADAPONTPROD.CODFP
        lblDescLinha.Caption = DescLinhaProd(txtCODPROD.Text, "SGI_CODIGO")
        
        Call PopGrd
        
    End If

End Sub


Private Sub PopGrd()
    
    Dim I            As Integer
    Dim strCampos0   As String
    Dim strCampos1   As String
    Dim strCampos2   As String
    
    arrITENSAPONT = objCADAPONTPROD.ITENSAPONT
    
    If IsArray(arrITENSAPONT) Then
        With grdProcesso
            For I = 1 To UBound(arrITENSAPONT)
            
                strCampos0 = ""
                strCampos1 = ""
                strCampos2 = ""
            
                If Len(Trim(arrITENSAPONT(I, 2))) > 0 Then strCampos0 = Mid(objBLBFunc.CONVMINHR(CLng(arrITENSAPONT(I, 2))), 1, 5)
                If Len(Trim(arrITENSAPONT(I, 3))) > 0 Then strCampos1 = Mid(objBLBFunc.CONVMINHR(CLng(arrITENSAPONT(I, 3))), 1, 5)
                If Len(Trim(arrITENSAPONT(I, 4))) > 0 Then strCampos2 = Mid(objBLBFunc.CONVMINHR(CLng(arrITENSAPONT(I, 4))), 1, 5)
            
                .AddItem (I - 1) & vbTab & _
                         arrITENSAPONT(I, 1) & vbTab & _
                         PegaDescProc(Str(arrITENSAPONT(I, 1))) & vbTab & _
                         strCampos0 & vbTab & _
                         strCampos1 & vbTab & _
                         strCampos2

                         
            Next I
        End With
    End If

End Sub

Private Function PegaDescProc(strCodProc As String) As String
    
    PegaDescProc = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCodProc)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaDescProc = BREC!SGI_DESCRI
    BREC.Close

End Function
