VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADOPSQUALI 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Cadastra OP's Qualidade"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   12840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   12360
      Picture         =   "frmCADOPSQUALI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   12360
      Picture         =   "frmCADOPSQUALI.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   300
   End
   Begin VSFlex8LCtl.VSFlexGrid grdOPS 
      Height          =   7215
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   12135
      _cx             =   21405
      _cy             =   12726
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
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   12615
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Filial"
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
         Left            =   5160
         TabIndex        =   12
         Top             =   240
         Width           =   4095
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
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   14
            Top             =   0
            Width           =   1335
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
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSMask.MaskEdBox mskDATA 
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   975
         TabIndex        =   5
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
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
         Left            =   4560
         TabIndex        =   15
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   420
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
         TabIndex        =   6
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12615
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
         Picture         =   "frmCADOPSQUALI.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
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
         Picture         =   "frmCADOPSQUALI.frx":07C6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
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
         Picture         =   "frmCADOPSQUALI.frx":08C8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADOPSQUALI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public iParcela         As Integer
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCodUsuario    As Long
Public intFILIALPED     As Integer
Public strFILIAL        As String

Dim lngCodLog           As Long
Dim strCAPTION          As String
Dim strNOMFILIAL        As String

Dim objBLBFunc          As Object
Dim objCADOPSQUALI      As Object

Dim arrOPS              As Variant

Const conCOL_SonMov_CodOP                                   As Integer = 0
Const conCOL_SonMov_CodPed                                  As Integer = 1
Const conCOL_SonMov_CodClie                                 As Integer = 2
Const conCOL_SonMov_RazaoSic                                As Integer = 3
Const conCOL_SonMov_CodProd                                 As Integer = 4
Const conCOL_SonMov_DescProd                                As Integer = 5
Const conCOL_SonMov_DtEntrega                               As Integer = 6
Const conCOL_SonMov_QtdeOP                                  As Integer = 7
Const conCOL_SonMov_FILIALOP                                As Integer = 8
Const conCOL_SonMov_FormatString                            As String = "=Cod.OP|Cód.Pedido|Cód.Cliente|Razão Social|Produto|Rótulo|Dt.Entrega|Qtde.OP|Filial.OP"
Const conColumnsIn_SonMov                                   As Integer = 9

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos

End Sub

Private Sub CmdSalva_Click()

    Dim i           As Integer
    
    Call objBLBFunc.RemoveLinhaVazia(grdOPS, conCOL_SonMov_CodOP)
    
    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADOPSQUALI.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)

    objCADOPSQUALI.DTLANC = "'" & Format(CDate(mskDATA.Text), "MM/DD/YYYY") & "'"

    '' Apontamento
    arrOPS = Empty
    With grdOPS
        If (.Rows - 1) > 0 Then
            ReDim arrOPS(1 To (.Rows - 1), 1 To 2) As String
            For i = 1 To (.Rows - 1)
                arrOPS(i, 1) = .Cell(flexcpText, i, conCOL_SonMov_CodOP)
                arrOPS(i, 2) = .Cell(flexcpText, i, conCOL_SonMov_FILIALOP)
            Next i
        End If
    End With
    objCADOPSQUALI.OPS = arrOPS

    If objCADOPSQUALI.GRAVA(cTipOper) = False Then Exit Sub

    MsgBox "As Op's foram " & IIf(cTipOper = "I", "inclusas", IIf(cTipOper = "A", "alteradas", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me
    If cTipOper = "A" Then
        Call LimpaCamposlabel
        Call InitGridMov
        Call CarregaCampos
    End If

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call IncRegGrid
End Sub

Private Sub Command3_Click()

On Error GoTo Err_Command3_Click
    
    If cTipOper = "C" Then Exit Sub
    If cTipOper = "I" Or cTipOper = "A" Then
        With grdOPS
            If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonMov_CodOP))) = 0 Then Exit Sub
            Call objBLBFunc.ExclLinhaGrid(grdOPS, grdOPS.Row)
        End With
    End If
    
    Exit Sub
    
Err_Command3_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command3_Click()", Me.Name, "Command3_Click()", strCAMARQERRO)

End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADOPSQUALI = CreateObject("CADOPSQUALI.clsCADOPSQUALI")
   
    objCADOPSQUALI.FILIAL = FILIAL
       
    strCAPTION = "Cadastra OP's Qualidade "
    
    Call IniciaForm

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub IniciaForm()

    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call LimpaCamposlabel
    Call DesabilitaCampos
    Call InitGridMov
    
    If cTipOper = "I" Then iCodigo = 0
    
    objCADOPSQUALI.CODIGO = iCodigo
    mskDATA.Text = Format(Now, "DD/MM/YYYY")
    optFilial(0).Value = True
    
    Call CarregaCampos
    
End Sub


Private Sub LimpaCamposlabel()
    txtCodigo.Text = ""
End Sub

Private Sub DesabilitaCampos()
    If cTipOper = "I" Then
        Frame2.Enabled = True
        txtCodigo.Enabled = False
        mskDATA.Enabled = True
    End If
    If cTipOper = "C" Or cTipOper = "A" Then
        Frame2.Enabled = True
        txtCodigo.Enabled = False
        mskDATA.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADOPSQUALI = Nothing
End Sub


Private Sub InitGridMov()

    With grdOPS
    
       .Cols = conColumnsIn_SonMov
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMov_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMov_CodOP) = ""
       .ColDataType(conCOL_SonMov_CodOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_CodPed) = ""
       .ColDataType(conCOL_SonMov_CodPed) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_CodClie) = ""
       .ColDataType(conCOL_SonMov_CodClie) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_RazaoSic) = ""
       .ColDataType(conCOL_SonMov_RazaoSic) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_CodProd) = ""
       .ColDataType(conCOL_SonMov_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_DescProd) = ""
       .ColDataType(conCOL_SonMov_DescProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_DtEntrega) = ""
       .ColDataType(conCOL_SonMov_DtEntrega) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonMov_QtdeOP) = ""
       .ColDataType(conCOL_SonMov_QtdeOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_FILIALOP) = ""
       .ColDataType(conCOL_SonMov_FILIALOP) = flexDTLong
       
       .ColWidth(conCOL_SonMov_CodOP) = 1000
       .ColWidth(conCOL_SonMov_CodPed) = 1000
       .ColWidth(conCOL_SonMov_CodClie) = 1000
       .ColWidth(conCOL_SonMov_RazaoSic) = 5000
       .ColWidth(conCOL_SonMov_CodProd) = 1400
       .ColWidth(conCOL_SonMov_DescProd) = 5000
       .ColWidth(conCOL_SonMov_DtEntrega) = 1200
       .ColWidth(conCOL_SonMov_QtdeOP) = 1200
       .ColWidth(conCOL_SonMov_FILIALOP) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub grdOPS_AfterEdit(ByVal Row As Long, ByVal Col As Long)

On Error GoTo Err_grdOPS_AfterEdit

     With grdOPS
          Select Case Col
                Case conCOL_SonMov_CodOP
          End Select
     End With
     
     Exit Sub

Err_grdOPS_AfterEdit:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdOPS_AfterEdit", Me.Name, "AfterEdit")

End Sub

Private Sub grdOPS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

On Error GoTo Err_grdOPS_BeforeEdit
    
    With grdOPS
        Select Case Col
            Case conCOL_SonMov_CodPed, _
                 conCOL_SonMov_CodClie, _
                 conCOL_SonMov_RazaoSic, _
                 conCOL_SonMov_CodProd, _
                 conCOL_SonMov_DescProd, _
                 conCOL_SonMov_DtEntrega, _
                 conCOL_SonMov_QtdeOP, _
                 conCOL_SonMov_FILIALOP
                 Cancel = True
            Case conCOL_SonMov_CodOP
                 If cTipOper = "C" Then Cancel = True
            Case Else
                .ComboList = ""
        End Select
    End With
    
    Exit Sub

Err_grdOPS_BeforeEdit:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdOPS_BeforeEdit", Me.Name, "BeforeEdit")

End Sub

Private Sub grdOPS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

On Error GoTo Err_grdOPS_KeyPressEdit
     
     With grdOPS
          Select Case Col
                    Case conCOL_SonMov_CodOP
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With

     Exit Sub

Err_grdOPS_KeyPressEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdOPS_KeyPressEdit", Me.Name, "KeyPressEdit")

End Sub

Private Sub grdOPS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

On Error GoTo Err_grdOPS_ValidateEdit
     
     With grdOPS
          Select Case Col
                 Case conCOL_SonMov_CodOP
                        If .EditText = Empty Then Exit Sub
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Códgo da OP inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If objBLBFunc.FcVerifItensRepetidos(grdOPS, Row, conCOL_SonMov_CodOP, .EditText) = False Then
                           MsgBox "Esta OP ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_CodPed) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_CodClie) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_RazaoSic) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_CodProd) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_DescProd) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_DtEntrega) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_QtdeOP) = Empty
                           
                           If optFilial(0).Value = True Then .Cell(flexcpText, Row, conCOL_SonMov_FILIALOP) = 0
                           If optFilial(1).Value = True Then .Cell(flexcpText, Row, conCOL_SonMov_FILIALOP) = 1
                           
                           Cancel = True
                           Exit Sub
                        End If
          
                        If VerifOP(.EditText, .Cell(flexcpText, .Row, conCOL_SonMov_FILIALOP)) = True Then
                            Cancel = True
                            Exit Sub
                        End If
          
                        If ConsisteOP(.EditText) = True Then
                            Cancel = True
                            Exit Sub
                        End If
          
          End Select
     End With
    
    Exit Sub

Err_grdOPS_ValidateEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdOPS_ValidateEdit", Me.Name, "ValidateEdit")

End Sub


Private Sub IncRegGrid()
   
    If cTipOper = "C" Then Exit Sub
    
    If objBLBFunc.FcExisteLinhaVazia(grdOPS, conCOL_SonMov_CodOP) = False Then Exit Sub
    
    With grdOPS
    
        .AddItem "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 ""
                 
        If optFilial(0).Value = True Then .Cell(flexcpText, (.Rows - 1), conCOL_SonMov_FILIALOP) = 0
        If optFilial(1).Value = True Then .Cell(flexcpText, (.Rows - 1), conCOL_SonMov_FILIALOP) = 1
        
    End With
    
End Sub



Private Function ConsisteOP(strCODOP As String) As Boolean
    
    ConsisteOP = True
    
    Dim strNOMFILIAL As String
    
    strNOMFILIAL = ""
    If optFilial(1).Value = True Then strNOMFILIAL = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       OP.SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,OP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,OP.SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,OP.SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "      ,OP.SGI_QTDEPED" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,PEDV.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & " OP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDV" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       OP.SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And OP.SGI_CODIGO      = " & Trim(strCODOP) & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL    = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = OP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And PEDV.SGI_FILIAL    = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PEDV.SGI_CODIGO    = OP.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL    = PEDV.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO    = PEDV.SGI_CODCLI"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        ConsisteOP = False
        With grdOPS
            .Cell(flexcpText, .Row, conCOL_SonMov_CodPed) = BREC!SGI_CODPED
            .Cell(flexcpText, .Row, conCOL_SonMov_CodClie) = BREC!SGI_CODCLI
            .Cell(flexcpText, .Row, conCOL_SonMov_RazaoSic) = Trim(BREC!SGI_RAZAOSOC)
            .Cell(flexcpText, .Row, conCOL_SonMov_CodProd) = Trim(BREC!SGI_CODPROD)
            .Cell(flexcpText, .Row, conCOL_SonMov_DescProd) = Trim(BREC!SGI_DESCRICAO)
            .Cell(flexcpText, .Row, conCOL_SonMov_DtEntrega) = Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY")
            .Cell(flexcpText, .Row, conCOL_SonMov_QtdeOP) = BREC!SGI_QTDEPED
        End With
    End If
    BREC.Close
    
    If ConsisteOP = True Then MsgBox "ATENÇÃO - A OP [" & Str(strCODOP) & "] Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
    
End Function

Private Function Valida_Campos() As Boolean

On Error GoTo Err_Valida_Campos
     
     Dim lngLINHA   As Long
     Dim i          As Long
     
     Valida_Campos = False
     
     If Not IsDate(mskDATA.Text) Then
        MsgBox "ATENÇÃO - Data de Lançamento inválido !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     If (grdOPS.Rows - 1) = 0 Then
        MsgBox "ATENÇÃO - Informe pelo menos 1 Apontamento !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     Valida_Campos = True

    Exit Function
    
Err_Valida_Campos:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Valida_Campos()", Me.Name, "Valida_Campos()", strCAMARQERRO)

End Function



Private Sub CarregaCampos()
    If objCADOPSQUALI.Carrega_Campos = True Then
    
        txtCodigo.Text = objCADOPSQUALI.CODIGO
        mskDATA.Text = objCADOPSQUALI.DTLANC
        arrOPS = objCADOPSQUALI.OPS
        
        Call PopGrd
        
    End If
End Sub

Private Sub PopGrd()
    
    Dim i As Long
    
    Call InitGridMov
    
    If IsArray(arrOPS) Then
        With grdOPS
            For i = 1 To UBound(arrOPS)
            
                .AddItem arrOPS(i, 1) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrOPS(i, 2)
                         
                Call PegaOP(Trim(Str(arrOPS(i, 1))), Trim(Str(arrOPS(i, 2))), (.Rows - 1))
            
            Next i
        End With
    End If
    
End Sub


Private Sub PegaOP(strCODOP As String, strFILIALOP As String, lngLINHA As Long)

    Dim strDESCFILIALOP As String

    strDESCFILIALOP = ""
    If CLng(strFILIALOP) = 1 Then strDESCFILIALOP = "_STEEL"
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       OP.SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,OP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,OP.SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,OP.SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "      ,OP.SGI_QTDEPED" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,PEDV.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strDESCFILIALOP & " OP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strDESCFILIALOP & " PEDV" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       OP.SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And OP.SGI_CODIGO      = " & strCODOP & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL    = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = OP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And PEDV.SGI_FILIAL    = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PEDV.SGI_CODIGO    = OP.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL    = PEDV.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO    = PEDV.SGI_CODCLI"
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        With grdOPS
            .Cell(flexcpText, lngLINHA, conCOL_SonMov_CodPed) = BREC6!SGI_CODPED
            .Cell(flexcpText, lngLINHA, conCOL_SonMov_CodClie) = BREC6!SGI_CODCLI
            .Cell(flexcpText, lngLINHA, conCOL_SonMov_RazaoSic) = Trim(BREC6!SGI_RAZAOSOC)
            .Cell(flexcpText, lngLINHA, conCOL_SonMov_CodProd) = Trim(BREC6!SGI_CODPROD)
            .Cell(flexcpText, lngLINHA, conCOL_SonMov_DescProd) = Trim(BREC6!SGI_DESCRICAO)
            .Cell(flexcpText, lngLINHA, conCOL_SonMov_DtEntrega) = Format(BREC6!SGI_DATENTREGA, "DD/MM/YYYY")
            .Cell(flexcpText, lngLINHA, conCOL_SonMov_QtdeOP) = BREC6!SGI_QTDEPED
        End With
    End If
    BREC6.Close
    

End Sub

Private Function VerifOP(strCODOP As String, strFILIALOP As String) As Boolean

    VerifOP = False
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADOPQUALI_IT" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODOP    = " & Trim(strCODOP) & vbCrLf
    sSql = sSql & "   And SGI_FILIALOP = " & Trim(strFILIALOP)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then VerifOP = True
    BREC.Close
    
    If VerifOP = True Then MsgBox "ATENÇÃO - A OP [" & Str(strCODOP) & "] já esta disponivel em outro lançamento !!!", vbOKOnly + vbExclamation, "Aviso"
    
End Function
