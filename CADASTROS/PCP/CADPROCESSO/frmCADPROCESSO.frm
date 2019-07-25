VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADPROCESSO 
   Caption         =   "Cadastro de Processos"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7815
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
         Left            =   840
         Picture         =   "frmCADPROCESSO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   735
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
         Left            =   1560
         Picture         =   "frmCADPROCESSO.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   735
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
         Picture         =   "frmCADPROCESSO.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
   End
   Begin TabDlg.SSTab stProcesso 
      Height          =   5535
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "frmCADPROCESSO.frx":0736
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Operações"
      TabPicture(1)   =   "frmCADPROCESSO.frx":0752
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdOPERACOES"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command26"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command27"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Sub-Seções"
      TabPicture(2)   =   "frmCADPROCESSO.frx":076E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command27 
         Height          =   300
         Left            =   -67560
         Picture         =   "frmCADPROCESSO.frx":078A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton Command26 
         Height          =   300
         Left            =   -67560
         Picture         =   "frmCADPROCESSO.frx":08D4
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdOPERACOES 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   7215
         _cx             =   12726
         _cy             =   8705
         Appearance      =   1
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
         Height          =   495
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   7575
         Begin VB.CommandButton cmdGravProd 
            Height          =   315
            Left            =   7080
            Picture         =   "frmCADPROCESSO.frx":0A1E
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   1800
            Picture         =   "frmCADPROCESSO.frx":0B20
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   375
         End
         Begin VB.ComboBox cboSubSecao 
            Height          =   315
            Left            =   2160
            TabIndex        =   14
            Text            =   "cboSubSecao"
            Top             =   120
            Width           =   4935
         End
         Begin VB.TextBox txtCodSubSecao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            MaxLength       =   10
            TabIndex        =   13
            Text            =   "txtCodSubS"
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   660
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   10
         Top             =   840
         Width           =   7575
         Begin MSFlexGridLib.MSFlexGrid flxSUBSECAO 
            Height          =   4215
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   7435
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4935
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   7575
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1560
            TabIndex        =   19
            Top             =   960
            Width           =   1695
            Begin VB.OptionButton optAgregaValor 
               Caption         =   "Não"
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   21
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton optAgregaValor 
               Caption         =   "Sim"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.TextBox txtDescri 
            Height          =   285
            Left            =   1560
            TabIndex        =   1
            Text            =   "txtDescri"
            Top             =   600
            Width           =   5895
         End
         Begin VB.TextBox txtCODIGO 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   0
            Text            =   "txtCODIGO"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agrega Valor"
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
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
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
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
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
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   600
         End
      End
   End
End
Attribute VB_Name = "frmCADPROCESSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public cTipOper     As String
Public iCodigo      As Integer
Public FILIAL       As Integer
Public strAcesso    As String
Dim objBLBFunc      As Object
Dim objCADPROCESSO  As Object
Dim objPESQPADRAO   As Object
Dim arrSUBSETOR     As Variant
Dim arrOPERACOES    As Variant

Const conCOL_SonProcess_Ordem                       As Integer = 0
Const conCOL_SonProcess_CodOper                     As Integer = 1
Const conCOL_SonProcess_PesqOper                    As Integer = 2
Const conCOL_SonProcess_DescOper                    As Integer = 3
Const conCOL_SonProcess_Indice                      As Integer = 4
Const conCOL_SonProcess_FormatString                As String = "=Ordem|Cod.Operação|...|Descrição|Indice"
Const conColumnsIn_SonProcess                       As Integer = 5


Private Sub cboSubSecao_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboSubSecao, KeyAscii
End Sub

Private Sub cboSubSecao_Validate(Cancel As Boolean)
    If cboSubSecao.ListIndex > -1 Then txtCodSubSecao.Text = cboSubSecao.ItemData(cboSubSecao.ListIndex)
End Sub


Private Sub cmdAltera_Click()

    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    stProcesso.Tab = 0
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    txtDescri.SetFocus
    
    cTipOper = "A"
    
    Me.Caption = "Cadastro de Processos - [ ALTERAÇÃO ]"

End Sub

Private Sub cmdGravProd_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncSubSecao
End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If Verifica_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADPROCESSO.CODIGO = objCADPROCESSO.Gera_Codigo(Me.Name)
    
    objCADPROCESSO.DESCRI = txtDescri.Text
    objCADPROCESSO.TEMPMIN = 0
    If optAgregaValor(0).Value = True Then objCADPROCESSO.AGREGAVALOR = 0
    If optAgregaValor(1).Value = True Then objCADPROCESSO.AGREGAVALOR = 1
    
    arrSUBSETOR = Empty
    If (flxSUBSECAO.Rows - 1) > 0 Then
       ReDim arrSUBSETOR(1 To (flxSUBSECAO.Rows - 1)) As Variant
       For I = 1 To (flxSUBSECAO.Rows - 1)
           arrSUBSETOR(I) = flxSUBSECAO.TextMatrix(I, 1)
       Next I
    End If
    objCADPROCESSO.SUBSETOR = arrSUBSETOR
    
    '' Operações
    objCADPROCESSO.Operacao = Empty
    Call objBLBFunc.RemoveLinhaVazia(grdOPERACOES, conCOL_SonProcess_CodOper)
    
    With grdOPERACOES
        If (.Rows - 1) > 0 Then
            ReDim arrOPERACOES(1 To (.Rows - 1), 1 To 2) As String
            For I = 1 To (.Rows - 1)
                arrOPERACOES(I, 1) = .Cell(flexcpText, I, conCOL_SonProcess_Ordem)
                arrOPERACOES(I, 2) = .Cell(flexcpText, I, conCOL_SonProcess_CodOper)
            Next I
            objCADPROCESSO.Operacao = arrOPERACOES
        End If
    End With
    
    '' Grava as informações
    If objCADPROCESSO.GRAVA(cTipOper) = False Then Exit Sub
    
    MsgBox "O Processo Produtivo foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
    If cTipOper = "I" Then Unload Me

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
    sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Sub-Seção")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCodSubSecao.Text = varRETORNO
        
    cboSubSecao.ListIndex = -1
    txtCodSubSecao.SetFocus

End Sub



Private Sub Command26_Click()
    If cTipOper = "I" Or cTipOper = "A" Then
        Call objBLBFunc.ExclLinhaGrid(grdOPERACOES, grdOPERACOES.Row)
        Call RefazOrdem
        Call RefazIndice
    End If
End Sub

Private Sub Command27_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridOperacoes
End Sub

Private Sub flxSUBSECAO_KeyDown(KeyCode As Integer, Shift As Integer)
    If (flxSUBSECAO.Rows - 1) = 0 Then Exit Sub
    If cTipOper = "C" Then Exit Sub
    If KeyCode = vbKeyDelete Then
       If flxSUBSECAO.Rows = 2 Then flxSUBSECAO.Rows = 1
       If flxSUBSECAO.Rows > 2 Then flxSUBSECAO.RemoveItem (flxSUBSECAO.RowSel)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADPROCESSO = CreateObject("CADPROCESSO.clsCADPROCESSO")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADPROCESSO.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub


Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de Fluxo Produtivo - [ INCLUSÃO ]"
    
    stProcesso.Tab = 0
    
    objBLBFunc.LimpaCampos frmCADPROCESSO
    objCADPROCESSO.PreencheComboSubSecao cboSubSecao
    
    ConfGridSubSecao
    
    optAgregaValor(0).Value = True
   
    Call InitGridOperacoes
End Sub

Private Sub ConfGridSubSecao()

    flxSUBSECAO.Rows = 1
    flxSUBSECAO.Cols = 3
    
    flxSUBSECAO.TextMatrix(0, 0) = ""
    flxSUBSECAO.TextMatrix(0, 1) = "Código"
    flxSUBSECAO.TextMatrix(0, 2) = "Descrição"
    
    flxSUBSECAO.ColWidth(0) = 0
    flxSUBSECAO.ColWidth(1) = 1000
    flxSUBSECAO.ColWidth(2) = 4000

End Sub

Private Sub txtCelulaFis_GotFocus()
    objBLBFunc.SelecionaCampos txtDescri.Name, frmCADPROCESSO
End Sub

Private Sub txtCelulaFis_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyObjetos
End Sub

Private Sub grdOPERACOES_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonProcess_Ordem, _
         conCOL_SonProcess_DescOper
         Cancel = True
    Case conCOL_SonProcess_CodOper, _
         conCOL_SonProcess_PesqOper
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdOPERACOES.ComboList = ""
    End Select
    Exit Sub
End Sub


Private Sub grdOPERACOES_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdOPERACOES.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonProcess_PesqOper
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL  = " & FILIAL
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRI"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Nome"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "SGI_DESCRI"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Operacoes")
                        
            If Len(Trim(varRETORNO)) > 0 Then
               With grdOPERACOES
                    .Cell(flexcpText, Row, conCOL_SonProcess_CodOper) = varRETORNO
                    .Cell(flexcpText, Row, conCOL_SonProcess_DescOper) = PegaDescrOpercao(varRETORNO)
                    .Cell(flexcpText, Row, conCOL_SonProcess_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonProcess_Ordem) & varRETORNO)
               End With
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(grdOPERACOES, Row, conCOL_SonProcess_Indice, grdOPERACOES.Cell(flexcpText, Row, conCOL_SonProcess_Indice)) = False Then
               MsgBox "Esta Operação ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdOPERACOES.Cell(flexcpText, Row, conCOL_SonProcess_CodOper) = ""
               grdOPERACOES.Cell(flexcpText, Row, conCOL_SonProcess_DescOper) = ""
               grdOPERACOES.Cell(flexcpText, Row, conCOL_SonProcess_Indice) = ""
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdOPERACOES_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdOPERACOES
          Select Case Col
                    Case conCOL_SonProcess_CodOper
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdOPERACOES_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdOPERACOES
          Select Case Col
                 Case conCOL_SonProcess_CodOper
                        If .EditText = Empty Then Exit Sub
                        .Cell(flexcpText, Row, conCOL_SonProcess_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonProcess_Ordem)) & Trim(.EditText)
                        If objBLBFunc.FcVerifItensRepetidos(grdOPERACOES, Row, conCOL_SonProcess_Indice, .Cell(flexcpText, Row, conCOL_SonProcess_Indice)) = False Then
                           MsgBox "Esta Operação foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonProcess_CodOper) = Empty
                           .Cell(flexcpText, Row, conCOL_SonProcess_DescOper) = Empty
                           .Cell(flexcpText, Row, conCOL_SonProcess_Indice) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrOpercao(.EditText))) = 0 Then
                           MsgBox "Esta operação não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonProcess_DescOper) = Empty
                           .Cell(flexcpText, Row, conCOL_SonProcess_Indice) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonProcess_DescOper) = PegaDescrOpercao(.EditText)
          End Select
     End With
End Sub

Private Sub txtCodSubSecao_GotFocus()
    objBLBFunc.SelecionaCampos txtCodSubSecao.Name, frmCADPROCESSO
End Sub

Private Sub txtCodSubSecao_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blACHOU As Boolean
    
    If Len(Trim(txtCodSubSecao.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodSubSecao.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtCodSubSecao.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    For I = 0 To (cboSubSecao.ListCount - 1)
        If CInt(txtCodSubSecao.Text) = cboSubSecao.ItemData(I) Then cboSubSecao.ListIndex = I
    Next I
    
    If cboSubSecao.ListIndex = -1 Then
       MsgBox "Esta Sub-Seção não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtCodSubSecao.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub


Private Sub txtDescri_GotFocus()
    objBLBFunc.SelecionaCampos txtDescri.Name, frmCADPROCESSO
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub IncSubSecao()

    Dim I As Integer
    
    If Len(Trim(txtCodSubSecao.Text)) = 0 Or cboSubSecao.ListIndex = -1 Then
       MsgBox "Informe a Sub-Seção !!!", vbOKOnly + vbCritical, "aviso"
       txtCodSubSecao.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxSUBSECAO.Rows - 1)
        If txtCodSubSecao.Text = flxSUBSECAO.TextMatrix(I, 1) Then
           MsgBox "Esta Sub-Seção já foi inclusa !!!", vbOKOnly + vbCritical, "aviso"
           txtCodSubSecao.Text = ""
           cboSubSecao.ListIndex = -1
           txtCodSubSecao.SetFocus
           Exit Sub
        End If
    Next I
    
    flxSUBSECAO.AddItem "" & vbTab & _
                        txtCodSubSecao.Text & vbTab & _
                        cboSubSecao.Text
                        
    txtCodSubSecao.Text = ""
    cboSubSecao.ListIndex = -1
    txtCodSubSecao.SetFocus
    
End Sub


Private Function Verifica_Campos() As Boolean

    Verifica_Campos = False
    
    If Len(Trim(txtDescri.Text)) = 0 Then
       MsgBox "Informe a descrição do Fluxo Produtivo !!!", vbOKOnly + vbExclamation, "Aviso"
       stProcesso.Tab = 0
       txtDescri.SetFocus
       Exit Function
    End If
    
    'If (flxSUBSECAO.Rows - 1) = 0 Then
    '   MsgBox "Nenhuma Seção foi inclusa !!!", vbOKOnly + vbExclamation, "Aviso"
    '   stProcesso.Tab = 1
    '   txtCodSubSecao.SetFocus
    '   Exit Function
    'End If
    
    If cTipOper = "I" Then
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescri.Text) & "'"
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Está descrição do Fluxo Produtivo já existe !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          stProcesso.Tab = 0
          txtDescri.SetFocus
          Exit Function
       End If
       BREC.Close
       
    End If
    If cTipOper = "A" Then
    
       If objCADPROCESSO.DESCRI <> txtDescri.Text Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescri.Text) & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Está descrição do Fluxo Produtivo já existe !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtDescri.Text = objCADPROCESSO.DESCRI
             stProcesso.Tab = 0
             txtDescri.SetFocus
             Exit Function
          End If
          BREC.Close
       
       End If
    
    End If
    
    Verifica_Campos = True

End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    Me.Caption = "Cadastro de Fluxo Produtivo - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADPROCESSO
    
    objCADPROCESSO.CODIGO = iCodigo
    
    stProcesso.Tab = 0
    
    ConfGridSubSecao
    optAgregaValor(0).Value = True
    
    objCADPROCESSO.PreencheComboSubSecao cboSubSecao
    
    Call InitGridOperacoes
    
    If objCADPROCESSO.Carrega_campos = True Then
       
       txtCODIGO.Text = objCADPROCESSO.CODIGO
       txtDescri.Text = objCADPROCESSO.DESCRI
       optAgregaValor(objCADPROCESSO.AGREGAVALOR).Value = True
       
       arrSUBSETOR = objCADPROCESSO.SUBSETOR
       
       If IsArray(arrSUBSETOR) = True Then
          For I = 1 To UBound(arrSUBSETOR)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrSUBSETOR(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then
                 flxSUBSECAO.AddItem "" & vbTab & _
                                     BREC!SGI_CODIGO & vbTab & _
                                     BREC!SGI_DESCRI
              End If
              BREC.Close
              
          Next I
       End If
    
       Call PopGrdOperacao
    End If

End Sub


Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Me.Caption = "Cadastro de Fluxo Produtivo - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPROCESSO
    
    objCADPROCESSO.CODIGO = iCodigo
    
    stProcesso.Tab = 0
    
    ConfGridSubSecao
    optAgregaValor(0).Value = True
    
    objCADPROCESSO.PreencheComboSubSecao cboSubSecao
    Call InitGridOperacoes
    
    If objCADPROCESSO.Carrega_campos = True Then
       
       txtCODIGO.Text = objCADPROCESSO.CODIGO
       txtDescri.Text = objCADPROCESSO.DESCRI
       optAgregaValor(objCADPROCESSO.AGREGAVALOR).Value = True

       arrSUBSETOR = objCADPROCESSO.SUBSETOR
       
       If IsArray(arrSUBSETOR) = True Then
          For I = 1 To UBound(arrSUBSETOR)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrSUBSETOR(I)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then
                 flxSUBSECAO.AddItem "" & vbTab & _
                                     BREC!SGI_CODIGO & vbTab & _
                                     BREC!SGI_DESCRI
              End If
              BREC.Close
              
          Next I
       End If
       Call PopGrdOperacao
    End If

End Sub



Private Sub InitGridOperacoes()

    With grdOPERACOES
    
       .Cols = conColumnsIn_SonProcess
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProcess_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProcess_Ordem) = ""
       .ColDataType(conCOL_SonProcess_Ordem) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProcess_CodOper) = ""
       .ColDataType(conCOL_SonProcess_CodOper) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProcess_PesqOper) = ""
       .ColDataType(conCOL_SonProcess_PesqOper) = flexDTString
       .ColComboList(conCOL_SonProcess_PesqOper) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProcess_DescOper) = ""
       .ColDataType(conCOL_SonProcess_DescOper) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProcess_Indice) = ""
       .ColDataType(conCOL_SonProcess_Indice) = flexDTLong
       
       .ColWidth(conCOL_SonProcess_Ordem) = 1000
       .ColWidth(conCOL_SonProcess_CodOper) = 1200
       .ColWidth(conCOL_SonProcess_PesqOper) = 300
       .ColWidth(conCOL_SonProcess_DescOper) = 4000
       .ColWidth(conCOL_SonProcess_Indice) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncRegGridOperacoes()
   
    If objBLBFunc.FcExisteLinhaVazia(grdOPERACOES, conCOL_SonProcess_CodOper) = False Then Exit Sub
    
    With grdOPERACOES
            .AddItem "" & vbTab & _
                     "" & vbTab & _
                     "" & vbTab & _
                     "" & vbTab & _
                     ""
                    
            Call RefazOrdem
            Call RefazIndice
    End With
                            
End Sub


Private Function PegaDescrOpercao(strCodOper As String) As String
    PegaDescrOpercao = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCodOper)
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrOpercao = BREC2!SGI_DESCRI
    BREC2.Close
    
End Function


Private Sub RefazOrdem()

    Dim I As Integer
    
    With grdOPERACOES
        For I = 1 To (.Rows - 1)
            .Cell(flexcpText, I, conCOL_SonProcess_Ordem) = I
        Next I
    End With
End Sub

Private Sub RefazIndice()
    Dim I As Integer
    With grdOPERACOES
        For I = 1 To (.Rows - 1)
            .Cell(flexcpText, I, conCOL_SonProcess_Indice) = Trim(.Cell(flexcpText, I, conCOL_SonProcess_Ordem)) & Trim(.Cell(flexcpText, I, conCOL_SonProcess_CodOper))
        Next I
    End With
End Sub

Private Sub DestroyObjetos()
       Set objBLBFunc = Nothing
       Set objCADPROCESSO = Nothing
       Set objPESQPADRAO = Nothing
End Sub

Private Sub PopGrdOperacao()
    Dim I As Integer
    arrOPERACOES = objCADPROCESSO.Operacao
    If IsArray(arrOPERACOES) Then
        With grdOPERACOES
            For I = 1 To UBound(arrOPERACOES)
                .AddItem arrOPERACOES(I, 1) & vbTab & _
                         arrOPERACOES(I, 2) & vbTab & _
                         "" & vbTab & _
                         PegaDescrOpercao(Trim(Str(arrOPERACOES(I, 2)))) & vbTab & _
                         Trim(arrOPERACOES(I, 1) & arrOPERACOES(I, 2))
            Next I
        End With
    End If
End Sub
