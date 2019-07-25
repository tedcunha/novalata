VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmPESQPADRAO 
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   13335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13335
      Begin VB.CommandButton Command3 
         Caption         =   "&Limpa"
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
         Left            =   3480
         Picture         =   "frmPESQPADRAO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdOrden 
         Caption         =   "Ordem"
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
         Left            =   2640
         Picture         =   "frmPESQPADRAO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   6240
         Top             =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Novo"
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
         Picture         =   "frmPESQPADRAO.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Desfas"
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
         Picture         =   "frmPESQPADRAO.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmPESQPADRAO.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Resultado dos Filtros ]"
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
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   13215
      Begin VSFlex8LCtl.VSFlexGrid grdCAMPOS 
         Height          =   4695
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   12975
         _cx             =   22886
         _cy             =   8281
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Filtros para Pesquisa ]"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   13215
      Begin VB.ComboBox cboFP 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Text            =   "cboFP"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VSFlex8LCtl.VSFlexGrid grdCAMPOSPESQ 
         Height          =   2175
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   12975
         _cx             =   22886
         _cy             =   3836
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
End
Attribute VB_Name = "frmPESQPADRAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public FILIAL           As Integer
Public strAcesso        As String
Public arrCAMPOS        As Variant
Public arrTABELA        As Variant
Public arrTABELA2       As Variant
Public strSqlCont       As String
Public strCABEC         As String
Public boolPesqUsuario  As Boolean
Public strForm          As String
Dim objBLBFunc          As Object
Dim objPESQPADRAO       As Object
Dim strNovo             As String

Const conCOL_SonGrdPesq_Campo00                     As Integer = 0
Const conCOL_SonGrdPesq_Campo01                     As Integer = 1
Const conCOL_SonGrdPesq_Campo02                     As Integer = 2
Const conCOL_SonGrdPesq_Campo03                     As Integer = 3
Const conCOL_SonGrdPesq_Campo04                     As Integer = 4
Const conCOL_SonGrdPesq_FormatString                As String = "=Campos de Pesquisa|Dados da Pesquisa|Campo 01|Campo 02|Campo 03"
Const conColumnsIn_SonGrdPesq                       As Integer = 5

Private Sub cboFP_LostFocus()

    ' hide date picker when user is done with it
    cboFP.Visible = False

End Sub

Private Sub cboFP_Validate(Cancel As Boolean)

    grdCAMPOSPESQ.Cell(flexcpText, grdCAMPOSPESQ.Row, conCOL_SonGrdPesq_Campo01) = Empty
    grdCAMPOSPESQ.Cell(flexcpData, grdCAMPOSPESQ.Row, conCOL_SonGrdPesq_Campo01) = Empty
    
    If Len(Trim(cboFP.Text)) > 0 Then
        With grdCAMPOSPESQ
            grdCAMPOSPESQ.Cell(flexcpText, grdCAMPOSPESQ.Row, conCOL_SonGrdPesq_Campo01) = cboFP.Text
            grdCAMPOSPESQ.Cell(flexcpData, grdCAMPOSPESQ.Row, conCOL_SonGrdPesq_Campo01) = cboFP.ItemData(cboFP.ListIndex)
        End With
    End If

End Sub

Private Sub cmdOrden_Click()
    Call ConfGrid
    Call PreenchGrid
End Sub

Private Sub cmdVoltar_Click()
    varRETORNO = ""
    Unload frmPESQPADRAO
End Sub

Private Sub Command1_Click()
    Call ConfGrid
    Call PreenchGrid
End Sub

Private Sub Command2_Click()

        If Len(Trim(strForm)) = 0 Then Exit Sub
        
        Dim objCham As Object
        Set objCham = CreateObject(Trim(strForm))
        
        strNovo = "I"
        objCham.cConnectNovo cCaminho, Linha, FILIAL, strAcesso, V_Usuario
        
        Set objCham = Nothing
End Sub

Private Sub Command3_Click()
    Call ConfGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    
    '' Campo do Tipo N Somente aceita Numericos
    
    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    Me.Caption = "Pesquisa " & strCABEC
    
    objBLBFunc.LimpaCampos Me
    objPESQPADRAO.FILIAL = FILIAL
    
    
    Call ConfGridCamposPesq
    Call ConfGrid
    
    strNovo = ""
    
End Sub

Public Sub ConfGrid()

    Dim I           As Integer
    Dim strCOLUNAS  As String
    
    
    strCOLUNAS = ""

    With grdCAMPOS
       
       If IsArray(arrCAMPOS) = True Then
           
            .Cols = UBound(arrCAMPOS)
            .Rows = 1
            .FixedCols = 0
           
            For I = 1 To UBound(arrCAMPOS)
                strCOLUNAS = strCOLUNAS & arrCAMPOS(I, 3)
                If I < UBound(arrCAMPOS) Then strCOLUNAS = strCOLUNAS & "|"
            Next I
           
            .FormatString = strCOLUNAS
            .AutoSizeMouse = False
    
            .Editable = flexEDNone
            .AllowUserResizing = flexResizeBoth
            .AllowSelection = False
            .HighLight = flexHighlightWithFocus
            .SelectionMode = flexSelectionByRow
            .BackColor = &H80000018
            .ForeColor = vbBlack
    
            For I = 1 To UBound(arrCAMPOS)
                .Cell(flexcpData, 0, (I - 1)) = ""
                If arrCAMPOS(I, 2) = "N" Then .Cell(flexcpData, 0, (I - 1)) = flexDTLong
                If arrCAMPOS(I, 2) = "S" Then .Cell(flexcpData, 0, (I - 1)) = flexDTString
                If arrCAMPOS(I, 2) = "D" Then .Cell(flexcpData, 0, (I - 1)) = flexDTDate
                If arrCAMPOS(I, 2) = "C" Then .Cell(flexcpData, 0, (I - 1)) = flexDTString
                .ColWidth((I - 1)) = arrCAMPOS(I, 4)
            Next I
        
        End If
    
    End With
    
End Sub


Private Sub PreenchGrid()

    Dim strCAMPOS   As String
    Dim I           As Integer
    
    If IsArray(arrCAMPOS) = True Then
    
        sSql = arrTABELA(1) & vbCrLf
       
        With grdCAMPOSPESQ
            For I = 0 To (.Rows - 1)
                If Len(Trim(.Cell(flexcpText, I, 1))) > 0 Then
                    If .Cell(flexcpText, I, 3) = "N" Then sSql = sSql & "   And " & .Cell(flexcpText, I, 2) & " = " & .Cell(flexcpText, I, 1) & vbCrLf
                    If .Cell(flexcpText, I, 3) = "S" Then sSql = sSql & "   And " & .Cell(flexcpText, I, 2) & " Like '%" & .Cell(flexcpText, I, 1) & "%'" & vbCrLf
                    If .Cell(flexcpText, I, 3) = "D" Then sSql = sSql & "   And " & .Cell(flexcpText, I, 2) & " = '" & Format(CDate(.Cell(flexcpText, I, 1)), "MM/DD/YYYY") & "'" & vbCrLf
                    If .Cell(flexcpText, I, 3) = "C" Then sSql = sSql & "   And " & .Cell(flexcpText, I, 2) & " = " & .Cell(flexcpData, I, 1) & vbCrLf
                End If
            Next I
        End With
       
        If IsArray(arrTABELA2) Then
        
            sSql = sSql & "Union" & vbCrLf
            sSql = sSql & arrTABELA2(1) & vbCrLf
            
            With grdCAMPOSPESQ
                For I = 0 To (.Rows - 1)
                    If Len(Trim(.Cell(flexcpText, I, 1))) > 0 Then
                        If .Cell(flexcpText, I, 3) = "N" Then sSql = sSql & "   And " & .Cell(flexcpText, I, 2) & " = " & .Cell(flexcpText, I, 1) & vbCrLf
                        If .Cell(flexcpText, I, 3) = "S" Then sSql = sSql & "   And " & .Cell(flexcpText, I, 2) & " Like '%" & .Cell(flexcpText, I, 1) & "%'" & vbCrLf
                        If .Cell(flexcpText, I, 3) = "D" Then sSql = sSql & "   And " & .Cell(flexcpText, I, 2) & " = '" & Format(CDate(.Cell(flexcpText, I, 1)), "MM/DD/YYYY") & "'" & vbCrLf
                        If .Cell(flexcpText, I, 3) = "C" Then sSql = sSql & "   And " & .Cell(flexcpText, I, 2) & " = " & .Cell(flexcpData, I, 1) & vbCrLf
                    End If
                Next I
            End With
        
        End If
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       Do While Not BREC.EOF
           
          strCAMPOS = ""
          For I = 1 To UBound(arrCAMPOS)
              If boolPesqUsuario = True Then
                If I = 2 Then
                   strCAMPOS = strCAMPOS & Trim(objBLBFunc.Crypt(BREC(arrCAMPOS(I, 1)))) & vbTab
                Else
                   strCAMPOS = strCAMPOS & BREC(arrCAMPOS(I, 1)) & vbTab
                End If
              Else
                strCAMPOS = strCAMPOS & BREC(arrCAMPOS(I, 1)) & vbTab
              End If
          Next I
          
          grdCAMPOS.AddItem strCAMPOS
          
          BREC.MoveNext
       Loop
       BREC.Close
       
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub grdCAMPOS_Click()
    If (grdCAMPOS.Rows - 1) > 0 And grdCAMPOS.RowSel > 0 Then varRETORNO = Trim(grdCAMPOS.Cell(flexcpText, grdCAMPOS.RowSel, 0))
End Sub

Private Sub grdCAMPOS_DblClick()
    varRETORNO = ""
    If (grdCAMPOS.Rows - 1) > 0 And grdCAMPOS.RowSel > 0 Then
       varRETORNO = Trim(grdCAMPOS.Cell(flexcpText, grdCAMPOS.RowSel, 0))
       Unload Me
    End If
End Sub

Private Sub grdCAMPOS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        varRETORNO = ""
        If (grdCAMPOS.Rows - 1) > 0 And grdCAMPOS.RowSel > 0 Then
           varRETORNO = Trim(grdCAMPOS.Cell(flexcpText, grdCAMPOS.RowSel, 0))
           Unload Me
        End If
    End If
End Sub

Private Sub grdCAMPOS_RowColChange()
    If (grdCAMPOS.Rows - 1) > 0 And grdCAMPOS.RowSel > 0 Then varRETORNO = Trim(grdCAMPOS.Cell(flexcpText, grdCAMPOS.RowSel, 0))
End Sub

Private Sub grdCAMPOSPESQ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Select Case Col
    Case 2, _
         3, _
         4
         Cancel = True
    Case Else
        grdCAMPOSPESQ.ComboList = ""
    End Select
    
    Exit Sub

End Sub

Private Sub grdCAMPOSPESQ_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    ' don't resize columns while editing dates
    If cboFP.Visible Then Cancel = True

End Sub

Private Sub grdCAMPOSPESQ_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' don't resize columns while editing dates
    If cboFP.Visible Then Cancel = True

End Sub

Private Sub grdCAMPOSPESQ_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

     If (grdCAMPOSPESQ.Rows - 1) - 0 And grdCAMPOSPESQ.RowSel = 0 Then Exit Sub
     
     With grdCAMPOSPESQ
          Select Case Col
                    Case 1
                         If grdCAMPOSPESQ.Cell(flexcpText, Row, 3) = "S" Then KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                         If grdCAMPOSPESQ.Cell(flexcpText, Row, 3) = "N" Then KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                         If grdCAMPOSPESQ.Cell(flexcpText, Row, 3) = "D" Then KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          End Select
     End With
     
     Exit Sub

End Sub

Private Sub grdCAMPOSPESQ_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If (grdCAMPOSPESQ.Rows - 1) - 0 And grdCAMPOSPESQ.RowSel = 0 Then Exit Sub
    
    With grdCAMPOSPESQ

        
        ' if this is a date column, edit it with the date picker control
        If .Col = conCOL_SonGrdPesq_Campo01 Then
            
            
            If .Cell(flexcpText, Row, conCOL_SonGrdPesq_Campo03) <> "C" Then Exit Sub
            
            Call PreenchCombo(.Cell(flexcpText, Row, conCOL_SonGrdPesq_Campo04))
            
            ' we'll handle the editing ourselves
            Cancel = True
            
            ' position date picker control over cell
            cboFP.Left = .Cell(flexcpLeft, Row, Col) + 100
            cboFP.Top = .Cell(flexcpTop, Row, Col) + 250
            cboFP.Width = .Cell(flexcpWidth, Row, Col)
            
            ' initialize value, save original in tag in case user hits escape
            ''cboFechTPFR.Value = cboFechTPFR
            ''cboFechTPFR.Tag = cboFechTPFR
            
            ' show and activate date picker control
            cboFP.Visible = True
            cboFP.SetFocus
            
            ' make it drop down the calendar
            ''SendKeys "{f4}"
            
        End If

    End With

End Sub

Private Sub grdCAMPOSPESQ_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If (grdCAMPOSPESQ.Rows - 1) - 0 And grdCAMPOSPESQ.RowSel = 0 Then Exit Sub
     
     With grdCAMPOSPESQ
          Select Case Col
                 Case 1
                        If .EditText = Empty Then Exit Sub
                        If .Cell(flexcpText, Row, 3) = "N" Then
                            If Not IsNumeric(.EditText) Then
                                MsgBox "ATENÇÃO" & vbCrLf & _
                                       "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
                                Cancel = True
                                Exit Sub
                            End If
                        ElseIf .Cell(flexcpText, Row, 3) = "D" Then
                            If Not IsDate(.EditText) Then
                                MsgBox "ATENÇÃO" & vbCrLf & _
                                       "Somente é permitido datas !!!", vbOKOnly + vbExclamation, "Aviso"
                                Cancel = True
                                Exit Sub
                            End If
                        End If
          End Select
     End With
     
     Exit Sub


End Sub


Private Sub Timer1_Timer()
    If strNovo = "I" Then
        ConfGrid
        PreenchGrid
        strNovo = ""
    End If
End Sub

Public Sub ConfGridCamposPesq()

    Dim I           As Integer
    Dim strCOLUNAS  As String
    Dim strCAMPOS   As String
    
    strCOLUNAS = ""
    strCAMPOS = ""

    With grdCAMPOSPESQ
       
       If IsArray(arrCAMPOS) = True Then
           
            .Cols = conColumnsIn_SonGrdPesq
            .Rows = 1
            .FixedCols = 1
            .FormatString = conCOL_SonGrdPesq_FormatString
            
            .AutoSizeMouse = False
            
            .AllowUserResizing = flexResizeBoth
           
            .AutoSizeMouse = False
            .AllowUserResizing = flexResizeNone
            
            .Cell(flexcpData, 0, conCOL_SonGrdPesq_Campo00) = ""
            .ColDataType(conCOL_SonGrdPesq_Campo00) = flexDTString
       
            .Cell(flexcpData, 0, conCOL_SonGrdPesq_Campo01) = ""
            .ColDataType(conCOL_SonGrdPesq_Campo01) = flexDTString
       
            .Cell(flexcpData, 0, conCOL_SonGrdPesq_Campo02) = ""
            .ColDataType(conCOL_SonGrdPesq_Campo02) = flexDTString
       
            .Cell(flexcpData, 0, conCOL_SonGrdPesq_Campo03) = ""
            .ColDataType(conCOL_SonGrdPesq_Campo03) = flexDTString
            
            .Cell(flexcpData, 0, conCOL_SonGrdPesq_Campo04) = ""
            .ColDataType(conCOL_SonGrdPesq_Campo04) = flexDTString
            
            .ColWidth(conCOL_SonGrdPesq_Campo00) = 3000
            .ColWidth(conCOL_SonGrdPesq_Campo01) = 7000
            .ColWidth(conCOL_SonGrdPesq_Campo02) = 0
            .ColWidth(conCOL_SonGrdPesq_Campo03) = 0
            .ColWidth(conCOL_SonGrdPesq_Campo04) = 0
        
            .Editable = flexEDKbdMouse
            .AllowSelection = False
            .HighLight = flexHighlightWithFocus
            .SelectionMode = flexSelectionByRow
            .BackColor = &H80000018
            .ForeColor = vbBlack
        
            '' ===========================================
            '' Inclui Campos
            For I = 1 To UBound(arrCAMPOS)
                If Trim(arrCAMPOS(I, 2)) <> "C" Then
                    .AddItem Trim(arrCAMPOS(I, 3)) & vbTab & _
                             "" & vbTab & _
                             Trim(arrCAMPOS(I, 5)) & vbTab & _
                             Trim(arrCAMPOS(I, 2)) & vbTab & _
                             ""
                Else
                    .AddItem Trim(arrCAMPOS(I, 3)) & vbTab & _
                             "" & vbTab & _
                             Trim(arrCAMPOS(I, 5)) & vbTab & _
                             Trim(arrCAMPOS(I, 2)) & vbTab & _
                             Trim(arrCAMPOS(I, 6))
                End If
            Next I
        
        
        End If
    
    End With
    
End Sub

Public Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub PreenchCombo(strDADOS As String)

    If BREC10.State = 1 Then BREC10.Close
    
    If Len(Trim(strDADOS)) = 0 Then Exit Sub
    
    Dim arrDADOS() As String
    
    cboFP.Clear
    
    arrDADOS = Split(strDADOS, "|")
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & arrDADOS(0) & vbCrLf
    sSql = sSql & "      ," & arrDADOS(1) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & arrDADOS(2) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "Order By " & arrDADOS(0) & "," & arrDADOS(1)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC10.EOF()
       cboFP.AddItem Trim(BREC10(arrDADOS(1)))
       cboFP.ItemData(cboFP.NewIndex) = BREC10(arrDADOS(0))
       BREC10.MoveNext
    Loop
    BREC10.Close
    
End Sub


