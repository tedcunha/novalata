VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADENTRMAT 
   Caption         =   "Entrada Materiais Manual"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13545
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   1680
      Width           =   13455
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3480
         Picture         =   "frmCADENTRMAT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCIDCLIE"
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txtCODMOTIVO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Text            =   "txtCODMOTIVO"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   3480
         Picture         =   "frmCADENTRMAT.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   23
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblDescCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         Height          =   285
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label lblDescMotEntSai 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescMotEntSai"
         Height          =   285
         Left            =   3840
         TabIndex        =   17
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo de Entrada.:"
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
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5415
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   13455
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   13080
         Picture         =   "frmCADENTRMAT.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   960
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton Command27 
         Height          =   300
         Left            =   13080
         Picture         =   "frmCADENTRMAT.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton Command26 
         Height          =   300
         Left            =   13080
         Picture         =   "frmCADENTRMAT.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Exclui a linha da Gride Selecionada"
         Top             =   600
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdPRODUTOS 
         Height          =   5055
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   12855
         _cx             =   22675
         _cy             =   8916
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
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   13455
      Begin VB.TextBox txtCODREQ 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "txtCODREQ"
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDATREQ 
         Height          =   285
         Left            =   12000
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Cód.Req."
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
         Left            =   5520
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCODREQ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODREQ"
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
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   0
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Entrada.:"
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
         Height          =   195
         Index           =   1
         Left            =   10575
         TabIndex        =   13
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Entrada.:"
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
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   12
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13455
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
         Left            =   1080
         Picture         =   "frmCADENTRMAT.frx":05E2
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   1920
         Picture         =   "frmCADENTRMAT.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmCADENTRMAT.frx":0C16
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCADENTRMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public cTipOper      As String
Public iCodigo       As Integer
Public FILIAL        As Integer
Public strAcesso     As String
Public strMODPAI     As String
Public strUSUARIO    As String

Dim objBLBFunc       As Object
Dim objCADENTRMAT    As Object
Dim objPESQPADRAO    As Object

Dim arrITENSREQ      As Variant
Dim arrITSAIREQANT   As Variant

Const conCOL_Prod_ID                              As Integer = 0
Const conCOL_Prod_Rotulo                          As Integer = 1
Const conCOL_Prod_PesqRot                         As Integer = 2
Const conCOL_Prod_DescrProd                       As Integer = 3
Const conCOL_Prod_Qtde                            As Integer = 4
Const conCOL_Prod_TemMov                          As Integer = 5
Const conCOL_Prod_QtdeKG                          As Integer = 6
Const conCOL_Prod_CODLOTE                         As Integer = 7
Const conCOL_Prod_INDICE                          As Integer = 8
Const conCOL_Prod_nf                              As Integer = 9
Const conCOL_Prod_FormatString                    As String = "=ID|Rótulo|...|Descrição|Qtde|TemMov|Qtde.KG|Código do Lote|Indice|NF nº"
Const conColumnsIn_Prod                           As Integer = 10

Private Sub cmdAltera_Click()
    
    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    cTipOper = "A"

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame5.Enabled = True
   
    Me.Caption = "Entrada Materiais Manual - [ ALTERAÇÃO ]"

End Sub

Private Sub CmdSalva_Click()

    Dim I       As Integer
    Dim sValor  As String
    
    Call objBLBFunc.RemoveLinhaVazia(grdPRODUTOS, conCOL_Prod_ID)
    
    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADENTRMAT.CADREQENTCOD = objCADENTRMAT.Gera_Codigo(Me.Name)
    
    objCADENTRMAT.CADDTREQ = CDate(mskDATREQ.Text)
    objCADENTRMAT.CODMOTIVO = CLng(txtCODMOTIVO.Text)
    objCADENTRMAT.CODCLIENTE = CLng(txtCIDCLIE.Text)
    
    With grdPRODUTOS
         arrITENSREQ = Empty
         If (.Rows - 1) > 0 Then
            ReDim arrITENSREQ(1 To (.Rows - 1), 1 To 9) As String
            For I = 1 To (.Rows - 1)
                arrITENSREQ(I, 1) = .Cell(flexcpText, I, conCOL_Prod_ID)
                arrITENSREQ(I, 2) = .Cell(flexcpText, I, conCOL_Prod_Rotulo)
                
                sValor = Replace(.Cell(flexcpText, I, conCOL_Prod_Qtde), ".", "")
                sValor = Replace(Trim(sValor), ",", ".")
                arrITENSREQ(I, 3) = sValor
                
                arrITENSREQ(I, 4) = .Cell(flexcpText, I, conCOL_Prod_TemMov)
                
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_Prod_QtdeKG))) > 0 Then
                    sValor = Replace(.Cell(flexcpText, I, conCOL_Prod_QtdeKG), ".", "")
                    sValor = Replace(Trim(sValor), ",", ".")
                End If
                arrITENSREQ(I, 5) = sValor
                
                arrITENSREQ(I, 6) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_Prod_CODLOTE))) > 0 Then arrITENSREQ(I, 6) = "'" & .Cell(flexcpText, I, conCOL_Prod_CODLOTE) & "'"
                
                arrITENSREQ(I, 7) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_Prod_INDICE))) > 0 Then arrITENSREQ(I, 7) = "'" & .Cell(flexcpText, I, conCOL_Prod_INDICE) & "'"
            
                arrITENSREQ(I, 8) = "'" & Format(CDate(mskDATREQ.Text), "MM/DD/YYYY") & "'"
            
                arrITENSREQ(I, 9) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_Prod_nf))) > 0 Then arrITENSREQ(I, 9) = "'" & .Cell(flexcpText, I, conCOL_Prod_nf) & "'"
            
            Next I
         End If
         objCADENTRMAT.ITENSREQ = arrITENSREQ
    End With
    
    If objCADENTRMAT.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A Entrada foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", IIf(cTipOper = "L", "liberado", ""))) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOTIVOS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCODMOTIVO.Text = varRETORNO
        lblDescMotEntSai.Caption = PegaDescrTipoEntSai(varRETORNO)
    End If
    txtCODMOTIVO.SetFocus

End Sub


Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 6, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_NOMFANTA" & vbCrLf
    sSql = sSql & "      ,SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CODREF" & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
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
    arrCAMPOS(3, 4) = "4500"
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
    
    arrCAMPOS(6, 1) = "SGI_CODREF"
    arrCAMPOS(6, 2) = "S"
    arrCAMPOS(6, 3) = "Cód.Antigo"
    arrCAMPOS(6, 4) = "1500"
    arrCAMPOS(6, 5) = "SGI_CODREF"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCIDCLIE.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescCliente)
    End If
    txtCIDCLIE.SetFocus

End Sub

Private Sub Command26_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdPRODUTOS, grdPRODUTOS.Row)
End Sub

Private Sub Command27_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGridProdtos
End Sub

Private Sub Command3_Click()

On Error GoTo Err_Prog
    
    If Len(Trim(txtCIDCLIE.Text)) = 0 Then
        MsgBox "ATENÇÃO - Informe o Cliente !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim intQUEST As Integer
    intQUEST = MsgBox("Tem Certeza que esta carregando a planilha correta ?", vbYesNo + vbQuestion + vbDefaultButton2, "ATENÇÂO")
    If intQUEST = 7 Then Exit Sub
    
    Dim oConn           As ADODB.Connection
    Dim oCmd            As ADODB.Command
    Dim oRS             As ADODB.Recordset
    Dim strCODPROD      As String
    Dim strINDICE       As String
    Dim intLINHA        As Integer
    Dim dblQTDETOTAL    As Double
    Dim dblPESOTOTAL    As Double
    
    Call ConfGridItReq

    ''"Data Source=C:\ricardo\PROGRAMAS\ENTRADAS.xls;"
    ''"Data Source=\\SRVLATA\PROGRAMAS\ENTRADAS.xls;" -- Antigo Novalata
    
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "C:\Ricardo\SGI\NOVALATA\DOCS-NOVALATA\Luigi\DePara_ExcelAntig.xls;" & _
                         "Extended Properties=""Excel 8.0;HDR=Yes;"";"
    
    ' cria o objecto command e define a conexao ativa
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oConn
    
    ' abre a planilha
    oCmd.CommandText = "SELECT * from [Plan1$]"
    
    ' cria o recordset com os dados
    Set oRS = New ADODB.Recordset
    oRS.Open oCmd, , adOpenKeyset, adLockOptimistic

    With grdPRODUTOS
        Do While Not oRS.EOF()
        
            strCODPROD = ""
            If Not IsNull(oRS(0).Value) Then strCODPROD = Trim(Replace(oRS(0).Value, " ", ""))
            
            If Len(Trim(strCODPROD)) > 0 Then
                
                strINDICE = PegaIDProduto(Trim(strCODPROD)) & Trim(Replace(oRS(3).Value, " ", ""))
                
                intLINHA = -1
                intLINHA = .FindRow(strINDICE, , conCOL_Prod_INDICE)
                
                If intLINHA = -1 Then
                
                    sSql = ""
                    
                    sSql = "Select " & vbCrLf
                    sSql = sSql & "       * " & vbCrLf
                    sSql = sSql & "  From " & vbCrLf
                    sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                    sSql = sSql & "   And SGI_CODIGO = '" & strCODPROD & "'"
                
                    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    If Not BREC2.EOF() Then
                        
                        .AddItem PegaIDProduto(Trim(strCODPROD)) & vbTab & _
                                 oRS(0).Value & vbTab & _
                                 "" & vbTab & _
                                 "" & vbTab & _
                                 oRS(1).Value & vbTab & _
                                 "" & vbTab & _
                                 Format(oRS(2).Value, "#,####0.0000") & vbTab & _
                                 Trim(Replace(oRS(3).Value, " ", "")) & vbTab & _
                                 strINDICE
                        
                        Call PesDescProduto(.Cell(flexcpText, (.Rows - 1), conCOL_Prod_ID), (.Rows - 1))
                    
                    End If
                    BREC2.Close
                Else
                    dblQTDETOTAL = CDbl(.Cell(flexcpText, intLINHA, conCOL_Prod_Qtde)) + CDbl(oRS(1).Value)
                    dblPESOTOTAL = CDbl(.Cell(flexcpText, intLINHA, conCOL_Prod_QtdeKG)) + CDbl(oRS(2).Value)
                    .Cell(flexcpText, intLINHA, conCOL_Prod_Qtde) = dblQTDETOTAL
                    .Cell(flexcpText, intLINHA, conCOL_Prod_QtdeKG) = Format(dblPESOTOTAL, "#,####0.0000")
                End If
                
                
            End If
            oRS.MoveNext
        Loop
    End With
    oRS.Close
    
    Set oCmd = Nothing
    
    Exit Sub
    
Err_Prog:

    MsgBox "Erro       : " & Err.Number & vbCrLf & _
           "Erro Desc. : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"
           
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADENTRMAT = CreateObject("CADENTRMAT.clsCADENTRMAT")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADENTRMAT.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame5.Enabled = True
    
    mskDATREQ.Enabled = True
    
    Me.Caption = "Entrada Materiais Manual - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADENTRMAT
    lblCODREQ(3).Caption = ""
    
    mskDATREQ.Text = Format(Date, "DD/MM/YYYY")
    
    Call ConfGridItReq
    Call LimpaCamposLabel
    
End Sub

Private Sub ConfGridItReq()

    With grdPRODUTOS
    
       .Cols = conColumnsIn_Prod
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Prod_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_Prod_ID) = ""
       .ColDataType(conCOL_Prod_ID) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_Rotulo) = ""
       .ColDataType(conCOL_Prod_Rotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Prod_PesqRot) = ""
       .ColDataType(conCOL_Prod_PesqRot) = flexDTString
       .ColComboList(conCOL_Prod_PesqRot) = "..."
       
       .Cell(flexcpData, 0, conCOL_Prod_DescrProd) = ""
       .ColDataType(conCOL_Prod_DescrProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Prod_Qtde) = ""
       .ColDataType(conCOL_Prod_Qtde) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_TemMov) = ""
       .ColDataType(conCOL_Prod_TemMov) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Prod_QtdeKG) = ""
       .ColDataType(conCOL_Prod_QtdeKG) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_Prod_CODLOTE) = ""
       .ColDataType(conCOL_Prod_CODLOTE) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Prod_INDICE) = ""
       .ColDataType(conCOL_Prod_INDICE) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Prod_nf) = ""
       .ColDataType(conCOL_Prod_nf) = flexDTString
       
       .ColWidth(conCOL_Prod_ID) = 0
       .ColWidth(conCOL_Prod_Rotulo) = 1500
       .ColWidth(conCOL_Prod_PesqRot) = 300
       .ColWidth(conCOL_Prod_DescrProd) = 4500
       .ColWidth(conCOL_Prod_Qtde) = 1000
       .ColWidth(conCOL_Prod_TemMov) = 0
       .ColWidth(conCOL_Prod_QtdeKG) = 1000
       .ColWidth(conCOL_Prod_CODLOTE) = 2000
       .ColWidth(conCOL_Prod_INDICE) = 0
       .ColWidth(conCOL_Prod_nf) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub grdPRODUTOS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdPRODUTOS
          Select Case Col
                 Case conCOL_Prod_Rotulo
                    If Len(Trim(.Cell(flexcpText, Row, Col))) > 0 Then
                        .Col = (Col + 3)
                        .EditCell
                    End If
                 Case conCOL_Prod_QtdeKG
                    If Len(Trim(.Cell(flexcpText, Row, Col))) > 0 Then .Cell(flexcpText, Row, Col) = Format(.Cell(flexcpText, Row, Col), "#,####0.0000")
          End Select
     End With
End Sub

Private Sub grdPRODUTOS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_Prod_Rotulo, _
         conCOL_Prod_DescrProd
         Cancel = True
    Case conCOL_Prod_Qtde, _
         conCOL_Prod_QtdeKG, _
         conCOL_Prod_CODLOTE, _
         conCOL_Prod_nf
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdPRODUTOS.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdPRODUTOS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If (grdPRODUTOS.Rows - 1) = 0 Then Exit Sub
    
    Dim strINDICE As String
    
    Select Case Col
        Case conCOL_Prod_PesqRot
    
            If cTipOper = "C" Then Exit Sub
            
            ReDim arrCAMPOS(1 To 5, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
            
            sSql = ""
            
            sSql = "Select" & vbCrLf
            sSql = sSql & "       PRO.SGI_IDPRODUTO" & vbCrLf
            sSql = sSql & "       ,Case When PRO.SGI_PRODUTOTIPO = 1 then" & vbCrLf
            sSql = sSql & "                  replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' +" & vbCrLf
            sSql = sSql & "                  replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' +" & vbCrLf
            sSql = sSql & "                  replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' +" & vbCrLf
            sSql = sSql & "                  (Case When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
            sSql = sSql & "                        When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End)" & vbCrLf
            sSql = sSql & "             Else" & vbCrLf
            sSql = sSql & "                  SGI_CODIGO" & vbCrLf
            sSql = sSql & "             End As SGI_CODIGO" & vbCrLf
            sSql = sSql & "       ,PRO.SGI_CODCLIE" & vbCrLf
            sSql = sSql & "       ,PRO.SGI_DESCRICAO" & vbCrLf
            sSql = sSql & "       ,PRO.SGI_COMPLEMENTO" & vbCrLf
            sSql = sSql & "  From" & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
            sSql = sSql & " Where" & vbCrLf
            sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_IDPRODUTO"
            arrCAMPOS(1, 2) = "N"
            arrCAMPOS(1, 3) = "ID"
            arrCAMPOS(1, 4) = "800"
            arrCAMPOS(1, 5) = "PRO.SGI_IDPRODUTO"
            
            arrCAMPOS(2, 1) = "SGI_CODIGO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Rótulo"
            arrCAMPOS(2, 4) = "1500"
            arrCAMPOS(2, 5) = "PRO.SGI_CODIGO"
            
            arrCAMPOS(3, 1) = "SGI_COMPLEMENTO"
            arrCAMPOS(3, 2) = "S"
            arrCAMPOS(3, 3) = "Complemento"
            arrCAMPOS(3, 4) = "2500"
            arrCAMPOS(3, 5) = "PRO.SGI_COMPLEMENTO"
            
            arrCAMPOS(4, 1) = "SGI_CODCLIE"
            arrCAMPOS(4, 2) = "N"
            arrCAMPOS(4, 3) = "Cliente"
            arrCAMPOS(4, 4) = "800"
            arrCAMPOS(4, 5) = "PRO.SGI_CODCLIE"
            
            arrCAMPOS(5, 1) = "SGI_DESCRICAO"
            arrCAMPOS(5, 2) = "S"
            arrCAMPOS(5, 3) = "Descrição"
            arrCAMPOS(5, 4) = "5000"
            arrCAMPOS(5, 5) = "PRO.SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
            
            If Len(Trim(varRETORNO)) > 0 Then
               Call LimpaCamposGrid(Row)
               grdPRODUTOS.Cell(flexcpText, Row, conCOL_Prod_ID) = varRETORNO
               Call PesDescProduto(varRETORNO, Row)
            End If
            
            strINDICE = Trim(grdPRODUTOS.Cell(flexcpText, Row, conCOL_Prod_ID)) & Trim(grdPRODUTOS.Cell(flexcpText, Row, conCOL_Prod_CODLOTE))
            
            If objBLBFunc.FcVerifItensRepetidos(grdPRODUTOS, Row, conCOL_Prod_INDICE, strINDICE) = False Then
               MsgBox "Este Produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               Call LimpaCamposGrid(Row)
               Exit Sub
            End If
    
            grdPRODUTOS.Col = (Col + 2)
            grdPRODUTOS.EditCell
    
    End Select

End Sub

Private Sub grdPRODUTOS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdPRODUTOS
          Select Case Col
                    Case conCOL_Prod_Rotulo
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_Prod_Qtde
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_Prod_QtdeKG
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 4, myvarAsDouble)
          End Select
     End With
End Sub

Private Sub grdPRODUTOS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     Dim strINDICE As String

     With grdPRODUTOS
          Select Case Col
                 Case conCOL_Prod_Rotulo
                        If .EditText = Empty Then Exit Sub
                        
                        .Cell(flexcpText, Row, conCOL_Prod_ID) = PegaIDProduto(Trim(.EditText))
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_Prod_ID))) = 0 Then
                           MsgBox "Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGrid(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        
                        strINDICE = Trim(.Cell(flexcpText, Row, conCOL_Prod_ID)) & Trim(.Cell(flexcpText, Row, conCOL_Prod_CODLOTE))
                        
                        If objBLBFunc.FcVerifItensRepetidos(grdPRODUTOS, Row, conCOL_Prod_INDICE, Trim(strINDICE)) = False Then
                           MsgBox "Este produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGrid(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        
                        Call PesDescProduto(.Cell(flexcpText, Row, conCOL_Prod_ID), Row)
                Case conCOL_Prod_Qtde
                        If .EditText = Empty Then Exit Sub
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Somente e permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                Case conCOL_Prod_QtdeKG
                        If .EditText = Empty Then Exit Sub
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Somente e permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                Case conCOL_Prod_CODLOTE
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(.EditText)) > 20 Then
                            MsgBox "Somente e permitido 20 digitos !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        strINDICE = Trim(.Cell(flexcpText, Row, conCOL_Prod_ID)) & Trim(.EditText)
                        If objBLBFunc.FcVerifItensRepetidos(grdPRODUTOS, Row, conCOL_Prod_INDICE, Trim(strINDICE)) = False Then
                           MsgBox "Este produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           Call LimpaCamposGrid(Row)
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_Prod_INDICE) = Trim(strINDICE)
          End Select
     End With

End Sub


Private Sub txtCIDCLIE_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, frmCADENTRMAT
End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCIDCLIE.Text
End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

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
    End If

End Sub

Private Sub txtCODMOTIVO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODMOTIVO.Name, frmCADENTRMAT
End Sub

Private Sub txtCODMOTIVO_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODMOTIVO.Text
End Sub

Private Sub txtCODMOTIVO_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCODMOTIVO.Text)) = 0 Then Exit Sub
   
   If Not IsNumeric(txtCODMOTIVO.Text) Then
        MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODMOTIVO.Text = ""
        lblDescMotEntSai.Caption = ""
        Exit Sub
   End If
   
   lblDescMotEntSai.Caption = PegaDescrTipoEntSai(txtCODMOTIVO.Text)
   If Len(Trim(lblDescMotEntSai.Caption)) = 0 Then
        MsgBox "Este motivo não existe !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODMOTIVO.Text = ""
        lblDescMotEntSai.Caption = ""
        Cancel = True
        Exit Sub
   End If
   
End Sub



Private Function Valida_Campos() As Boolean

     Valida_Campos = False
     
     If Not IsDate(mskDATREQ.Text) Then
        MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDATREQ.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtCODMOTIVO.Text)) = 0 Then
        MsgBox "Informe o motivo da entrada !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODMOTIVO.Text = ""
        txtCODMOTIVO.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtCIDCLIE.Text)) = 0 Then
        MsgBox "Informe o Cliente !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCIDCLIE.Text = ""
        txtCIDCLIE.SetFocus
        Exit Function
     End If
     
     If (grdPRODUTOS.Rows - 1) = 0 Then
        MsgBox "Não foi informado nenhum item para dar entrada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     Valida_Campos = True

End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame5.Enabled = False
   
    Me.Caption = "Entrada Materiais Manual - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADENTRMAT
    
    Call ConfGridItReq
    Call LimpaCamposLabel
    
    objCADENTRMAT.CADREQENTCOD = iCodigo
    
    If objCADENTRMAT.Carrega_campos = True Then
    
       lblCODREQ(3).Caption = objCADENTRMAT.CADREQENTCOD
       mskDATREQ.Text = Format(objCADENTRMAT.CADDTREQ, "DD/MM/YYYY")
       arrITENSREQ = objCADENTRMAT.ITENSREQ
       
       txtCODMOTIVO.Text = IIf(objCADENTRMAT.CODMOTIVO > 0, objCADENTRMAT.CODMOTIVO, "")
       lblDescMotEntSai.Caption = PegaDescrTipoEntSai(txtCODMOTIVO.Text)

       txtCIDCLIE.Text = objCADENTRMAT.CODCLIENTE
       Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente)

       Call PopGrd
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame5.Enabled = True
   
    Me.Caption = "Entrada Materiais Manual - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADENTRMAT
    
    ConfGridItReq
    Call LimpaCamposLabel
    
    objCADENTRMAT.CADREQENTCOD = iCodigo
    
    If objCADENTRMAT.Carrega_campos = True Then
       lblCODREQ(3).Caption = objCADENTRMAT.CADREQENTCOD
       mskDATREQ.Text = Format(objCADENTRMAT.CADDTREQ, "DD/MM/YYYY")
       arrITENSREQ = objCADENTRMAT.ITENSREQ
       arrITSAIREQANT = objCADENTRMAT.ITENSREQ '' Variavel Backup
      
       txtCODMOTIVO.Text = IIf(objCADENTRMAT.CODMOTIVO > 0, objCADENTRMAT.CODMOTIVO, "")
       lblDescMotEntSai.Caption = PegaDescrTipoEntSai(txtCODMOTIVO.Text)
       
       txtCIDCLIE.Text = objCADENTRMAT.CODCLIENTE
       Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente)
       
       Call PopGrd
    End If

End Sub


Private Function VerifNf() As Boolean

    VerifNf = True
    
    '' ----------------------------------------------------------------
    '' Verifica se existe NF
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_NFENTRADACABEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODREQENTR = " & objCADENTRMAT.CADREQENTCOD
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Este requisição está ligada a uma NF !!!", vbOKOnly + vbExclamation, "Aviso"
       VerifNf = False
    End If
    BREC.Close
    '' ----------------------------------------------------------------

End Function


Private Sub LimpaCamposLabel()
    lblDescMotEntSai.Caption = ""
    lblDescCliente.Caption = ""
End Sub

Private Function PegaDescrTipoEntSai(strCodigo As String) As String
    
    PegaDescrTipoEntSai = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOTIVOS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodigo
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaDescrTipoEntSai = Trim(BREC!SGI_DESCRI)
    BREC.Close
    
End Function

Private Sub IncRegGridProdtos()
   
    If objBLBFunc.FcExisteLinhaVazia(grdPRODUTOS, conCOL_Prod_ID) = False Then Exit Sub
    
    If Len(Trim(txtCIDCLIE.Text)) = 0 Then
        MsgBox "Informe o Cliente de destino do Produto !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtCODMOTIVO.Text)) = 0 Then
        MsgBox "Informe o motivo de entrada !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    grdPRODUTOS.AddItem "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        ""
                        
End Sub



Private Sub PesDescProduto(strID As String, lngRow As Long)

    If Len(Trim(strID)) = 0 Then Exit Sub
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PRO.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "       ,Case When PRO.SGI_PRODUTOTIPO = 1 then" & vbCrLf
    sSql = sSql & "                  replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' +" & vbCrLf
    sSql = sSql & "                  replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' +" & vbCrLf
    sSql = sSql & "                  replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' +" & vbCrLf
    sSql = sSql & "                  (Case When PRO.SGI_DIGVERIF Is Null Then '0'" & vbCrLf
    sSql = sSql & "                        When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End)" & vbCrLf
    sSql = sSql & "             Else" & vbCrLf
    sSql = sSql & "                  SGI_CODIGO" & vbCrLf
    sSql = sSql & "             End As SGI_CODIGO" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_COMPLEMENTO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL        = " & FILIAL
    sSql = sSql & "   And PRO.SGI_IDPRODUTO     = " & Trim(strID)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        grdPRODUTOS.Cell(flexcpText, lngRow, conCOL_Prod_Rotulo) = BREC!SGI_CODIGO
        grdPRODUTOS.Cell(flexcpText, lngRow, conCOL_Prod_DescrProd) = Trim(BREC!SGI_DESCRICAO)
        grdPRODUTOS.Cell(flexcpText, lngRow, conCOL_Prod_TemMov) = 0
        grdPRODUTOS.Cell(flexcpText, lngRow, conCOL_Prod_INDICE) = Trim(BREC!SGI_IDPRODUTO) & Trim(grdPRODUTOS.Cell(flexcpText, lngRow, conCOL_Prod_CODLOTE))
    End If
    BREC.Close
    
    grdPRODUTOS.Cell(flexcpText, lngRow, conCOL_Prod_TemMov) = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_PRODSALDOS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO  = " & strID & vbCrLf
    sSql = sSql & "   And SGI_CODCLIENTE = " & txtCIDCLIE.Text
    
    BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC5.EOF() Then grdPRODUTOS.Cell(flexcpText, lngRow, conCOL_Prod_TemMov) = 1
    BREC5.Close

End Sub

Private Sub LimpaCamposGrid(lngRow As Long)
    With grdPRODUTOS
        .Cell(flexcpText, lngRow, conCOL_Prod_ID) = Empty
        .Cell(flexcpText, lngRow, conCOL_Prod_Rotulo) = Empty
        .Cell(flexcpText, lngRow, conCOL_Prod_DescrProd) = Empty
        .Cell(flexcpText, lngRow, conCOL_Prod_Qtde) = Empty
        .Cell(flexcpText, lngRow, conCOL_Prod_QtdeKG) = Empty
        .Cell(flexcpText, lngRow, conCOL_Prod_CODLOTE) = Empty
        .Cell(flexcpText, lngRow, conCOL_Prod_INDICE) = Empty
    End With
End Sub

Private Function PegaIDProduto(strCodProduto As String) As String

    PegaIDProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_STATUS = 1" & vbCrLf

    sSql = sSql & "   And (Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End) = '" & Trim(strCodProduto) & "'"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If BREC!SGI_STATUS = 0 Then
           MsgBox "ATENÇÃO !!!" & vbCrLf & "O Produto " & Trim(strCodProduto) & " - " & Trim(BREC!SGI_DESCRICAO) & vbCrLf & "Não pode ser Utilizado está Desativado !!!", vbOKOnly + vbExclamation, "Aviso"
        Else
           PegaIDProduto = BREC!SGI_IDPRODUTO
        End If
    End If
    BREC.Close
    
End Function


Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADENTRMAT = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub PopGrd()

    Dim I As Integer
    
    If IsArray(arrITENSREQ) Then
        With grdPRODUTOS
            For I = 1 To UBound(arrITENSREQ)
                .AddItem arrITENSREQ(I, 1) & vbTab & _
                         arrITENSREQ(I, 2) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrITENSREQ(I, 3) & vbTab & _
                         "" & vbTab & _
                         arrITENSREQ(I, 4) & vbTab & _
                         arrITENSREQ(I, 5) & vbTab & _
                         arrITENSREQ(I, 6) & vbTab & _
                         arrITENSREQ(I, 7)
                
                Call PesDescProduto(.Cell(flexcpText, (.Rows - 1), conCOL_Prod_ID), (.Rows - 1))
            Next I
        End With
    End If

End Sub

Private Sub txtCODREQ_GotFocus()
    objBLBFunc.SelecionaCampos txtCODREQ.Name, frmCADENTRMAT
End Sub

Private Sub txtCODREQ_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODREQ.Text
End Sub

Private Sub txtCODREQ_Validate(Cancel As Boolean)

    If Len(Trim(txtCODREQ.Text)) = 0 Then Exit Sub
    
    Frame2.Enabled = True
    txtCIDCLIE.Enabled = True
    Command2.Enabled = True
    
    ''If PuxaReqProduto = False Then
    ''   MsgBox "Esta requisição não existe !!!", vbOKOnly + vbExclamation, "Aviso"
    ''   Cancel = True
    ''   Exit Sub
    ''End If

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

