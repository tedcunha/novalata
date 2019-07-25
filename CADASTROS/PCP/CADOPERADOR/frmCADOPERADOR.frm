VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADOPERADOR 
   Caption         =   "Cadastro de Operadores"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Turnos"
      TabPicture(0)   =   "frmCADOPERADOR.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdTURNOS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VSFlex8LCtl.VSFlexGrid grdTURNOS 
         Height          =   2535
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   5655
         _cx             =   9975
         _cy             =   4471
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
      TabIndex        =   7
      Top             =   0
      Width           =   5895
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
         Picture         =   "frmCADOPERADOR.frx":001C
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
         Picture         =   "frmCADOPERADOR.frx":054E
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
         Picture         =   "frmCADOPERADOR.frx":0650
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   5895
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   960
         Width           =   2415
         Begin VB.OptionButton optSimNao 
            Caption         =   "Sim"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optSimNao 
            Caption         =   "Não"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   13
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ativo:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   510
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   930
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmCADOPERADOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Public strUSUARIO   As String
Public iCodigo      As Integer
Public cTipOper     As String
Dim objBLBFunc      As Object
Dim objCADOPERADOR  As Object

Const conCOL_SonTur_CodTur                      As Integer = 0
Const conCOL_SonTur_Desc_Tur                    As Integer = 1
Const conCOL_SonTur_FormatString                As String = "=Cód. Turno|Descrição do Turno"
Const conColumnsIn_SonTur                       As Integer = 2


Private Sub cmdAltera_Click()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de Operadores - [ ALTERAÇÃO ]"
    
    txtDescricao.SetFocus
    
    cTipOper = "A"
    
End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
        
    If cTipOper = "I" Then objCADOPERADOR.CODIGO = objCADOPERADOR.Gera_Codigo(Me.Name)
           
    objCADOPERADOR.DESCRI = txtDescricao.Text
    If optSimNao(1).Value = True Then objCADOPERADOR.ATIVO = 1
    If optSimNao(0).Value = True Then objCADOPERADOR.ATIVO = 0
       
    If objCADOPERADOR.GRAVA(cTipOper) = False Then Exit Sub
    If objCADOPERADOR.Atualiza(cTipOper, Str(objCADOPERADOR.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
          
    MsgBox "O operador foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADOPERADOR = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADOPERADOR = Nothing
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
   Set objCADOPERADOR = CreateObject("CADOPERADOR.clsCADOPERADOR")
      
   objCADOPERADOR.FILIAL = FILIAL
   
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
   
    Me.Caption = "Cadastro de Operadores - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADOPERADOR
    
    txtCodigo.Text = ""
    
    optSimNao(1).Value = True
    
    Call InitGridTurnos
   
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADOPERADOR
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Operador inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADTIPALIM " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Operador já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
                
        If objCADOPERADOR.DESCRI <> txtDescricao.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADTIPALIM " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL =  " & FILIAL
           
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Opeerador já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADOPERADOR.DESCRI
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
   
    Me.Caption = "Cadastro de Operadores - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADOPERADOR
    objCADOPERADOR.CODIGO = iCodigo
    optSimNao(1).Value = True
    
    If objCADOPERADOR.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADOPERADOR.CODIGO)
       txtDescricao.Text = objCADOPERADOR.DESCRI
       optSimNao(objCADOPERADOR.ATIVO).Value = True
       
       Call InitGridTurnos
       Call PopGrdTurnos
       
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de Operadores - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADOPERADOR
    
    objCADOPERADOR.CODIGO = iCodigo
    
    optSimNao(1).Value = True
    
    If objCADOPERADOR.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADOPERADOR.CODIGO)
       txtDescricao.Text = objCADOPERADOR.DESCRI
       optSimNao(objCADOPERADOR.ATIVO).Value = True
       
       Call InitGridTurnos
       Call PopGrdTurnos
       
    End If
    
End Sub

Private Sub InitGridTurnos()

    With grdTURNOS
    
       .Cols = conColumnsIn_SonTur
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonTur_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonTur_CodTur) = ""
       .ColDataType(conCOL_SonTur_CodTur) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonTur_Desc_Tur) = ""
       .ColDataType(conCOL_SonTur_Desc_Tur) = flexDTString
       
       .ColWidth(conCOL_SonTur_CodTur) = 1500
       .ColWidth(conCOL_SonTur_Desc_Tur) = 4000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub PopGrdTurnos()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       HEA.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOVOPERMAQ MOV " & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN HEA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MOV.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And MOV.SGI_CODOPER = " & objCADOPERADOR.CODIGO & vbCrLf
    sSql = sSql & "   And HEA.SGI_FILIAL  = MOV.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And HEA.SGI_CODIGO  = MOV.SGI_FILIAL "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       
       grdTURNOS.AddItem BREC!SGI_CODIGO & vbTab & _
                         BREC!SGI_DESCRI
                         
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

