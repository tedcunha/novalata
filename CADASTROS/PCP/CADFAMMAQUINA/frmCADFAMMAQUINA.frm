VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADFAMMAQUINA 
   Caption         =   "Cadastro de familia de máquinas"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Máquinas"
      TabPicture(0)   =   "frmCADFAMMAQUINA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdMAQUINAS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Fluxo Produtivo"
      TabPicture(1)   =   "frmCADFAMMAQUINA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdFluxo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Turnos"
      TabPicture(2)   =   "frmCADFAMMAQUINA.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdTurnos"
      Tab(2).Control(1)=   "Command4"
      Tab(2).Control(2)=   "Command3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Atributos a Recolher"
      TabPicture(3)   =   "frmCADFAMMAQUINA.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   -67200
         Picture         =   "frmCADFAMMAQUINA.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   -67200
         Picture         =   "frmCADFAMMAQUINA.frx":01AE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VSFlex8LCtl.VSFlexGrid grdFluxo 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   8055
         _cx             =   14208
         _cy             =   7011
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
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
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
      Begin VSFlex8LCtl.VSFlexGrid grdMAQUINAS 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   8055
         _cx             =   14208
         _cy             =   7011
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
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
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
      Begin VSFlex8LCtl.VSFlexGrid grdTurnos 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   7575
         _cx             =   13361
         _cy             =   7011
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
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8295
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
         Picture         =   "frmCADFAMMAQUINA.frx":0738
         Style           =   1  'Graphical
         TabIndex        =   8
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
         MaskColor       =   &H8000000F&
         Picture         =   "frmCADFAMMAQUINA.frx":0C6A
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADFAMMAQUINA.frx":0D6C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   8295
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   5655
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmCADFAMMAQUINA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Dim objBLBFunc          As Object
Dim objCADFAMMAQUINA    As Object
Dim objPESQPADRAO       As Object
Dim arrTURNOS           As Variant

Const conCOL_SonMaq_CodMaq                              As Integer = 0
Const conCOL_SonMaq_Desc                                As Integer = 1
Const conCOL_SonMaq_FormatString                        As String = "=Cód. Maq|Descrição Máquina"
Const conColumnsIn_SonMaq                               As Integer = 2

Const conCOL_SonFluxo_CodFluxo                          As Integer = 0
Const conCOL_SonFluxo_Desc                              As Integer = 1
Const conCOL_SonFluxo_FormatString                      As String = "=Cód. Fluxo|Descrição Fluxo"
Const conColumnsIn_SonFluxo                             As Integer = 2

Const conCOL_SonTurno_Codigo                            As Integer = 0
Const conCOL_SonTurno_Pesq                              As Integer = 1
Const conCOL_SonTurno_Descri                            As Integer = 2
Const conCOL_SonTurno_FormatString                      As String = "=Código|...|Descrição"
Const conColumnsIn_SonTurno                             As Integer = 3

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de familia de máquinas - [ ALTERAÇÃO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADFAMMAQUINA.CODIGO = objCADFAMMAQUINA.Gera_Codigo(Me.Name)
    
    objCADFAMMAQUINA.DESCRI = txtDescricao.Text
    
    '' =========================================
    '' Turnos
    arrTURNOS = Empty
    If (grdTurnos.Rows - 1) > 0 Then
        ReDim arrTURNOS(1 To (grdTurnos.Rows - 1)) As String
        For I = 1 To (grdTurnos.Rows - 1)
            arrTURNOS(I) = grdTurnos.Cell(flexcpText, I, conCOL_SonTurno_Codigo)
        Next I
    End If
    objCADFAMMAQUINA.TURNOS = arrTURNOS
    '' =========================================

    If objCADFAMMAQUINA.GRAVA(cTipOper) = False Then Exit Sub
    If objCADFAMMAQUINA.Atualiza(cTipOper, Str(objCADFAMMAQUINA.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
          
    MsgBox "A familia de máquinas foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
       
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADFAMMAQUINA = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADFAMMAQUINA = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command3_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGrid
End Sub

Private Sub Command4_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdTurnos, grdTurnos.Row)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADFAMMAQUINA = CreateObject("CADFAMMAQUINA.clsCADFAMMAQUINA")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
      
   objCADFAMMAQUINA.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub


Private Sub grdFluxo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonFluxo_CodFluxo, conCOL_SonFluxo_Desc
         Cancel = True
    End Select
    Exit Sub
End Sub

Private Sub grdMAQUINAS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonMaq_CodMaq, conCOL_SonMaq_Desc
         Cancel = True
    End Select
    Exit Sub
End Sub

Private Sub grdTurnos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonTurno_Descri
         Cancel = True
    Case conCOL_SonTurno_Codigo, conCOL_SonTurno_Pesq
         If cTipOper = "C" Then Cancel = True
    End Select
    Exit Sub
End Sub

Private Sub grdTurnos_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdTurnos.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonTurno_Pesq
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
            
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
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Turnos")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdTurnos.Cell(flexcpText, Row, conCOL_SonTurno_Codigo) = varRETORNO
               grdTurnos.Cell(flexcpText, Row, conCOL_SonTurno_Descri) = PegaDescrTurno(CLng(grdTurnos.Cell(flexcpText, Row, conCOL_SonTurno_Codigo)))
            End If
            
            If VerifItensRepetidos(Row, conCOL_SonTurno_Codigo, varRETORNO) = False Then
               MsgBox "Este turno já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdTurnos.Cell(flexcpText, Row, conCOL_SonTurno_Codigo) = Empty
               grdTurnos.Cell(flexcpText, Row, conCOL_SonTurno_Descri) = Empty
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdTurnos_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdTurnos
          Select Case Col
                 Case conCOL_SonTurno_Codigo
                        If .EditText = Empty Then Exit Sub
                        If VerifItensRepetidos(Row, conCOL_SonTurno_Codigo, .EditText) = False Then
                           MsgBox "Este turno ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonTurno_Codigo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonTurno_Descri) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrTurno(CLng(.EditText)))) = 0 Then
                           MsgBox "Este turno não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonTurno_Codigo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonTurno_Descri) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonTurno_Descri) = PegaDescrTurno(CLng(.EditText))
          End Select
     End With

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADFAMMAQUINA
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub


Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de familia de máquinas - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFAMMAQUINA
    
    txtCodigo.Text = ""
    
    Call InitGridMaq
    Call InitGridFluxo
    Call InitGridTurno
   
End Sub


Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Informe a Descrição !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADFAMMAQUINAS  " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        If Not BREC.EOF Then
           MsgBox "Esta familia de máquina já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADFAMMAQUINA.DESCRI <> txtDescricao.Text Then
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADFAMMAQUINAS  " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL    = " & FILIAL
           
           BREC.Open sSql, adoBanco_Dados
           If Not BREC.EOF Then
              MsgBox "Esta familia de máquina já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADFAMMAQUINA.DESCRI
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
           BREC.Close
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de familia de máquinas - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADFAMMAQUINA
    
    objCADFAMMAQUINA.CODIGO = iCodigo
    
    Call InitGridMaq
    Call InitGridFluxo
    Call InitGridTurno
    
    If objCADFAMMAQUINA.Carrega_campos = True Then
    
        txtCodigo.Text = objCADFAMMAQUINA.CODIGO
        txtDescricao.Text = objCADFAMMAQUINA.DESCRI
        Call PopGrdMaquinas
        Call PopGrdFluxo
        Call PopGrdTurnos
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de familia de máquinas - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFAMMAQUINA
        
    objCADFAMMAQUINA.CODIGO = iCodigo
    
    Call InitGridMaq
    Call InitGridFluxo
    Call InitGridTurno
    
    If objCADFAMMAQUINA.Carrega_campos = True Then
        txtCodigo.Text = objCADFAMMAQUINA.CODIGO
        txtDescricao.Text = objCADFAMMAQUINA.DESCRI
        Call PopGrdMaquinas
        Call PopGrdFluxo
        Call PopGrdTurnos
    End If
    
End Sub


Private Sub InitGridMaq()

    With grdMAQUINAS
    
       .Cols = conColumnsIn_SonMaq
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMaq_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMaq_CodMaq) = ""
       .ColDataType(conCOL_SonMaq_CodMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMaq_Desc) = ""
       .ColDataType(conCOL_SonMaq_Desc) = flexDTString
       
       .ColWidth(conCOL_SonMaq_CodMaq) = 1000
       .ColWidth(conCOL_SonMaq_Desc) = 3000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub InitGridFluxo()

    With grdFluxo
    
       .Cols = conColumnsIn_SonMaq
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonFluxo_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonFluxo_CodFluxo) = ""
       .ColDataType(conCOL_SonFluxo_CodFluxo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonFluxo_Desc) = ""
       .ColDataType(conCOL_SonFluxo_Desc) = flexDTString
       
       .ColWidth(conCOL_SonFluxo_CodFluxo) = 1000
       .ColWidth(conCOL_SonFluxo_Desc) = 3000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PopGrdMaquinas()
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODFAMILIA    = " & objCADFAMMAQUINA.CODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       grdMAQUINAS.AddItem BREC!SGI_CODIGO & vbTab & Trim(BREC!SGI_DESCRI)
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Private Sub PopGrdFluxo()
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODFAMILIA    = " & objCADFAMMAQUINA.CODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       grdFluxo.AddItem BREC!SGI_CODIGO & vbTab & Trim(BREC!SGI_DESCRI)
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Private Sub InitGridTurno()

    With grdTurnos
    
       .Cols = conColumnsIn_SonTurno
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonTurno_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonTurno_Codigo) = ""
       .ColDataType(conCOL_SonTurno_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonTurno_Pesq) = ""
       .ColDataType(conCOL_SonTurno_Pesq) = flexDTString
       .ColComboList(conCOL_SonTurno_Pesq) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonTurno_Descri) = ""
       .ColDataType(conCOL_SonTurno_Descri) = flexDTString
       
       .ColWidth(conCOL_SonTurno_Codigo) = 1500
       .ColWidth(conCOL_SonTurno_Pesq) = 300
       .ColWidth(conCOL_SonTurno_Descri) = 4500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Function VerifItensRepetidos(intRow As Long, intCol As Long, varCampo As Variant) As Boolean
    VerifItensRepetidos = False
    Dim I As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    
    For I = 1 To (grdTurnos.Rows - 1)
        If I <> intRow And grdTurnos.Cell(flexcpText, I, intCol) = varCampo Then Exit Function
    Next I
    VerifItensRepetidos = True
End Function

Private Function PegaDescrTurno(lngCODTURNO As Long) As String
    PegaDescrTurno = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODTURNO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescrTurno = BREC!SGI_DESCRI
    BREC.Close
    
End Function

Private Sub IncRegGrid()
   
    If ExisteLinhaVazia = False Then Exit Sub
    
    grdTurnos.AddItem "" & vbTab & _
                      "" & vbTab & _
                      ""
                            
End Sub


Private Function ExisteLinhaVazia() As Boolean
    ExisteLinhaVazia = False
    
    Dim I As Integer
    
    For I = 1 To (grdTurnos.Rows - 1)
        If grdTurnos.Cell(flexcpText, I, conCOL_SonTurno_Codigo) = Empty Then Exit Function
    Next I
    
    ExisteLinhaVazia = True
End Function

Private Sub PopGrdTurnos()

    Dim I As Integer
    
    arrTURNOS = objCADFAMMAQUINA.TURNOS
    If IsArray(arrTURNOS) Then
       For I = 1 To UBound(arrTURNOS)
           grdTurnos.AddItem arrTURNOS(I) & vbTab & _
                             "" & vbTab & _
                             PegaDescrTurno(CLng(arrTURNOS(I)))
       Next I
    End If
    
End Sub
