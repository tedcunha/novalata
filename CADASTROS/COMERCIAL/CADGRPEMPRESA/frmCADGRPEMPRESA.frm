VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADGRPEMPRESA 
   Caption         =   "Cadastro de Grupo de Empresas"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9900
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "[ Empresas do Grupo ]"
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
      Height          =   4935
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   9855
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   9480
         Picture         =   "frmCADGRPEMPRESA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   9480
         Picture         =   "frmCADGRPEMPRESA.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdGRPEMP 
         Height          =   4575
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   9255
         _cx             =   16325
         _cy             =   8070
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
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   9855
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   6015
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         TabIndex        =   7
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
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
         Picture         =   "frmCADGRPEMPRESA.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADGRPEMPRESA.frx":07C6
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADGRPEMPRESA.frx":08C8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADGRPEMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public strMODPAI        As String
Public FILIAL           As Integer
Public strAcesso        As String
Public strUsuario       As String
Public lngCodUsuario    As Long
Public strNOMETABELA    As String
Public strNOMEFILIAL    As String

Dim objBLBFunc          As Object
Dim objCADGRPEMPRESA    As Object
Dim objPESQPADRAO       As Object
Dim strCAPTION          As String
Dim arrGRPEMPRESA       As Variant

Const conCOL_GRPEMP_Codigo                      As Integer = 0
Const conCOL_GRPEMP_PesqEmp                     As Integer = 1
Const conCOL_GRPEMP_DescEmp                     As Integer = 2
Const conCOL_GRPEMP_CNPJ                        As Integer = 3
Const conCOL_GRPEMP_FormatString                As String = "=Cód.Empresa|...|Descrição da Empresa|CNPJ"
Const conColumnsIn_GRPEMP                       As Integer = 4

Private Sub cmdAltera_Click()

    cTipOper = "A"
    If objBLBFunc.ChecaAcesso2(cTipOper, strAcesso) = False Then Exit Sub
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub CmdSalva_Click()

    Dim I As Long
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADGRPEMPRESA.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
    
    objCADGRPEMPRESA.DESCRI = "'" & Trim(Replace(Replace(txtDescricao.Text, ",", ""), "'", "")) & "'"
    
    '' Empresas do Grupo
    arrGRPEMPRESA = Empty
    With grdGRPEMP
        If (.Rows - 1) > 0 Then
            ReDim arrGRPEMPRESA(1 To (.Rows - 1), 1 To 1) As String
            For I = 1 To (.Rows - 1)
                arrGRPEMPRESA(I, 1) = .Cell(flexcpText, I, conCOL_GRPEMP_Codigo)
            Next I
        End If
    End With
    objCADGRPEMPRESA.GRPEMPRESA = arrGRPEMPRESA
    
    If objCADGRPEMPRESA.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O Grupo de Empresa foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call IncRegGrid
End Sub

Private Sub Command3_Click()
    If cTipOper = "C" Then Exit Sub
    If grdGRPEMP.Row = 0 Then Exit Sub
    Call objBLBFunc.ExclLinhaGrid(grdGRPEMP, grdGRPEMP.Row)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADGRPEMPRESA = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADGRPEMPRESA = CreateObject("CADGRPEMPRESA.clsCADGRPEMPRESA")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strCAPTION = "Cadastro de Grupo de Empresas"
   
    objCADGRPEMPRESA.FILIAL = FILIAL
   
    Call IniciaForm

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub grdGRPEMP_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdGRPEMP
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        Select Case Col
               Case conCOL_GRPEMP_Codigo
        End Select
    End With
End Sub

Private Sub grdGRPEMP_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdGRPEMP
        Select Case Col
               Case conCOL_GRPEMP_DescEmp, _
                    conCOL_GRPEMP_CNPJ
                    Cancel = True
               Case conCOL_GRPEMP_Codigo
                    If cTipOper = "C" Then Cancel = True
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub
End Sub

Private Sub grdGRPEMP_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    With grdGRPEMP
        If (.Rows - 1) = 0 Then Exit Sub
        
        Select Case Col
            Case conCOL_GRPEMP_PesqEmp
                
                Call PesqClie(Row)
                
                Exit Sub
                
        End Select
    End With

End Sub

Private Sub grdGRPEMP_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdGRPEMP
          Select Case Col
                    Case conCOL_GRPEMP_Codigo
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdGRPEMP_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdGRPEMP
          Select Case Col
                 Case conCOL_GRPEMP_Codigo
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - Este código é inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If objBLBFunc.FcVerifItensRepetidos(grdGRPEMP, Row, conCOL_GRPEMP_Codigo, Trim(.EditText)) = False Then
                           MsgBox "ATENÇÃO" & vbCrLf & "O Código da Empresa já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdGRPEMP.Cell(flexcpText, Row, conCOL_GRPEMP_Codigo) = Empty
                           grdGRPEMP.Cell(flexcpText, Row, conCOL_GRPEMP_DescEmp) = Empty
                           grdGRPEMP.Cell(flexcpText, Row, conCOL_GRPEMP_CNPJ) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_GRPEMP_Codigo) = .EditText
                        Call PegaDadosClie(Trim(.EditText), Row)
                        
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_GRPEMP_DescEmp))) = 0 Then
                           MsgBox "ATENÇÃO" & vbCrLf & "Empresa não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_GRPEMP_DescEmp) = Empty
                           .Cell(flexcpText, Row, conCOL_GRPEMP_CNPJ) = Empty
                           Cancel = True
                           Exit Sub
                        End If
          End Select
     End With

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, Me
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

        ValidaCampos = False
     
        If Len(Trim(txtDescricao.Text)) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Campo Descrição não pode ser vázio !!!", vbOKOnly + vbExclamation, "Acviso"
                   Exit Function
        End If
     
        ValidaCampos = True
     
End Function


Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    Call ConfGrd
    
    If cTipOper = "I" Then iCodigo = 0
    objCADGRPEMPRESA.CODIGO = iCodigo
    
    Call CarregaCampos
    
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
    End If
End Sub

Private Sub CarregaCampos()

On Error GoTo Err_CarregaCampos
    
    If objCADGRPEMPRESA.Carrega_campos = True Then
        txtCodigo.Text = objCADGRPEMPRESA.CODIGO
        txtDescricao.Text = objCADGRPEMPRESA.DESCRI
    
        Call PopGrdEmp
    
    End If
    
    Exit Sub

Err_CarregaCampos:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CarregaCampos", Me.Name, "CarregaCampos")
    
End Sub


Private Sub ConfGrd()

    With grdGRPEMP

       .Cols = conColumnsIn_GRPEMP
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_GRPEMP_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_GRPEMP_Codigo) = ""
       .ColDataType(conCOL_GRPEMP_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_GRPEMP_PesqEmp) = ""
       .ColDataType(conCOL_GRPEMP_PesqEmp) = flexDTString
       .ColComboList(conCOL_GRPEMP_PesqEmp) = "..."
       
       .Cell(flexcpData, 0, conCOL_GRPEMP_DescEmp) = ""
       .ColDataType(conCOL_GRPEMP_DescEmp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_GRPEMP_CNPJ) = ""
       .ColDataType(conCOL_GRPEMP_CNPJ) = flexDTLong
       
       .ColWidth(conCOL_GRPEMP_Codigo) = 1100
       .ColWidth(conCOL_GRPEMP_PesqEmp) = 300
       .ColWidth(conCOL_GRPEMP_DescEmp) = 5000
       .ColWidth(conCOL_GRPEMP_CNPJ) = 1500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack

    End With
    
End Sub


Private Sub IncRegGrid()
   
    If cTipOper = "C" Then Exit Sub
    
    If objBLBFunc.FcExisteLinhaVazia(grdGRPEMP, conCOL_GRPEMP_Codigo) = False Then Exit Sub
    
    With grdGRPEMP
        .AddItem "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 ""
                 
    End With
   
End Sub



Private Sub PesqClie(lngROW As Long)
    With grdGRPEMP
        If (.Rows - 1) = 0 Then Exit Sub
            
        If cTipOper = "C" Then Exit Sub
        
        Dim lngLinha                    As Long
        Dim strINDICE                   As String
        ReDim arrCAMPOS(1 To 3, 1 To 5) As String
        ReDim arrTABELA(1 To 1) As String
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
        
        arrTABELA(1) = sSql
        
        arrCAMPOS(1, 1) = "SGI_CODIGO"
        arrCAMPOS(1, 2) = "N"
        arrCAMPOS(1, 3) = "Código"
        arrCAMPOS(1, 4) = "1100"
        arrCAMPOS(1, 5) = "SGI_CODIGO"
        
        arrCAMPOS(2, 1) = "SGI_RAZAOSOC"
        arrCAMPOS(2, 2) = "S"
        arrCAMPOS(2, 3) = "Razão Social"
        arrCAMPOS(2, 4) = "5000"
        arrCAMPOS(2, 5) = "SGI_RAZAOSOC"
        
        arrCAMPOS(3, 1) = "SGI_CPFCNPJ"
        arrCAMPOS(3, 2) = "N"
        arrCAMPOS(3, 3) = "CPF/CNPJ"
        arrCAMPOS(3, 4) = "1500"
        arrCAMPOS(3, 5) = "SGI_CPFCNPJ"
        
        varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Clientes")
        
        If Len(Trim(varRETORNO)) = 0 Then Exit Sub
        
        If objBLBFunc.FcVerifItensRepetidos(grdGRPEMP, lngROW, conCOL_GRPEMP_Codigo, Trim(varRETORNO)) = False Then
           MsgBox "ATENÇÃO" & vbCrLf & " O Cliente ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
           grdGRPEMP.Cell(flexcpText, lngROW, conCOL_GRPEMP_Codigo) = Empty
           grdGRPEMP.Cell(flexcpText, lngROW, conCOL_GRPEMP_DescEmp) = Empty
           grdGRPEMP.Cell(flexcpText, lngROW, conCOL_GRPEMP_CNPJ) = Empty
           Exit Sub
        End If
        
        .Cell(flexcpText, lngROW, conCOL_GRPEMP_Codigo) = varRETORNO
        Call PegaDadosClie(Trim(varRETORNO), lngROW)
    
    End With
End Sub

Private Sub PegaDadosClie(strCodClie As String, lngROW As Long)
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodClie
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then
        grdGRPEMP.Cell(flexcpText, lngROW, conCOL_GRPEMP_DescEmp) = BREC2!SGI_RAZAOSOC
        grdGRPEMP.Cell(flexcpText, lngROW, conCOL_GRPEMP_CNPJ) = BREC2!SGI_CPFCNPJ
    End If
    BREC2.Close
    
End Sub


Private Sub PopGrdEmp()

    Dim I As Integer
    
    arrGRPEMPRESA = objCADGRPEMPRESA.GRPEMPRESA
    If IsArray(arrGRPEMPRESA) Then
        With grdGRPEMP
            For I = 1 To UBound(arrGRPEMPRESA)
                .AddItem arrGRPEMPRESA(I, 1) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         ""
        
                Call PegaDadosClie(Trim(Str(arrGRPEMPRESA(I, 1))), (.Rows - 1))
            Next I
        End With
    End If

End Sub


