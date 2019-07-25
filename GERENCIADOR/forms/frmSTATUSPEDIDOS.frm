VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSTATUSPEDIDOS 
   Caption         =   "Gerenciado - Status dos Pedidos"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12810
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   0
      Top             =   6840
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12735
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16187393
         CurrentDate     =   42037
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16187393
         CurrentDate     =   42037
      End
      Begin VB.Label Label1 
         Caption         =   "Data Final"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid grdStatusPedidos 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   12735
      _cx             =   22463
      _cy             =   11245
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
Attribute VB_Name = "frmSTATUSPEDIDOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngLinPai       As Long
Dim lngLinFilho     As Long
Dim lngTotPodLib    As Long

Private Sub Command1_Click()
    Call ConfGrid
    Call ConsisteCampos
    Call ZeraTotais
    Call PopGridStatusPedido
End Sub

Private Sub Form_Load()

    Call ZeraTotais

    Call ConfGrid
    Call PopGridStatusPedido

End Sub

Private Sub ConfGrid()

    
    ' reset the control
    SetDefaults grdStatusPedidos
    
    With grdStatusPedidos
    
        
        .Redraw = False
        
        ' set the properties we want
        .Rows = 1
        .FixedRows = 1
        .Cols = 5
        .AllowUserResizing = flexResizeBoth
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarComplete
        .OutlineCol = 0
        .SubtotalPosition = flexSTAbove
        
        ' fill the control with data
        .Cell(flexcpText, 0, 0) = "Descrição"
        .Cell(flexcpText, 0, 1) = "Quantidade"
        .Cell(flexcpText, 0, 2) = "Pedido N."
        .Cell(flexcpText, 0, 3) = "Cód.Clie"
        .Cell(flexcpText, 0, 4) = "Cliente"
    
    End With

    
End Sub

Private Sub PopGridStatusPedido()

    Dim i%, j%
    
    With grdStatusPedidos
    
    
        ' add an item, make it a subtotal
        ''.AddItem "Novalata" & vbTab & ""
        ''.IsSubtotal(.Rows - 1) = True
        ''.Cell(flexcpPicture, .Rows - 1, 0) = imgFolder
    
        ''Pai
        .AddItem "Pedidos em aberto (Liberados)" & vbTab & "" & vbTab & ""
        .IsSubtotal(.Rows - 1) = True
        lngLinPai = (.Rows - 1)
        
        '' Filho
        .AddItem "Novalata" & vbTab & "" & vbTab & ""
        .IsSubtotal(.Rows - 1) = True
        .RowOutlineLevel(.Rows - 1) = 1
        lngLinFilho = (.Rows - 1)
            
        .Cell(flexcpText, lngLinFilho, 1) = PopulaStatus(grdStatusPedidos, "L", "")
            
        .AddItem "Steel" & vbTab & "" & vbTab & ""
        .IsSubtotal(.Rows - 1) = True
        .RowOutlineLevel(.Rows - 1) = 1
        lngLinFilho = (.Rows - 1)

        .Cell(flexcpText, lngLinFilho, 1) = PopulaStatus(grdStatusPedidos, "L", "_STEEL")
        .Cell(flexcpText, lngLinPai) = lngTotPodLib
        
        ''Pai
        .AddItem "Pedidos Aguardando Liberação" & vbTab & "" & vbTab & ""
        .IsSubtotal(.Rows - 1) = True
        lngLinPai = (.Rows - 1)
        
        .AutoSize 0, 2, , 300
        .Redraw = True

    End With


End Sub


Sub SetDefaults(fa As VSFlexGrid)
    
    With fa
        .BindToArray Null
        .Rows = 0
        .Cols = 0
        .ScrollTrack = False
        .ExplorerBar = flexExNone
        .AutoSearch = flexSearchNone
        .Editable = False
        .AllowUserResizing = flexResizeNone
        .SelectionMode = flexSelectionFree
        .OutlineBar = flexOutlineBarNone
        .OLEDragMode = flexOLEDragManual
        .OLEDropMode = flexOLEDropNone
        .ScrollTips = False
        .ToolTipText = ""
    End With
    
End Sub


Private Function PopulaStatus(fa As VSFlexGrid, strStat As String, strEmpresa As String) As Long
    
    PopulaStatus = 0
    
    With fa

        Call AbBanco(strSTRCONNECT)
        
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       PEDV.SGI_CODIGO" & vbCrLf
        sSql = sSql & "      ,PEDV.SGI_CODCLI" & vbCrLf
        sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDH" & strEmpresa & " PEDV" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE CLIE" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       PEDV.SGI_FILIAL = " & intFILIAL & vbCrLf
        sSql = sSql & "   And PEDV.SGI_STATUS = '" & strStat & "'" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL = PEDV.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO = PEDV.SGI_CODCLI" & vbCrLf
        
        If strStat = "L" Then
            sSql = sSql & "   And PEDV.SGI_LIBDATAHORA Between '" & Format(dtpInicial.Value, "MM/DD/YYYY") & " 00:00:00' And '" & Format(dtpFinal.Value, "MM/DD/YYYY") & " 23:59:59'"
        End If
        
        BREC.Open sSql, BD, adOpenDynamic
        Do While Not BREC.EOF()
            fa.AddItem "" & vbTab & _
                       "" & vbTab & _
                       BREC!SGI_CODIGO & vbTab & _
                       BREC!SGI_CODCLI & vbTab & _
                       BREC!SGI_RAZAOSOC
                       
            PopulaStatus = (PopulaStatus + 1)
            lngTotPodLib = (lngTotPodLib + 1)
            
            BREC.MoveNext
        Loop
        
        Call FcBanco
    
    
    End With
End Function

Private Sub Timer1_Timer()
    Call Command1_Click
End Sub

Private Function ConsisteCampos() As Boolean

    ConsisteCampos = False
    
    If dtpInicial.Value > dtpFinal.Value Then
        MsgBox "ATENCAO - A data Inicial nao pode ser maior que a data final !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
    
    ConsisteCampos = True

End Function

Private Sub ZeraTotais()
    lngTotPodLib = 0
End Sub
