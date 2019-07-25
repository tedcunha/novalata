VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADMOTIVOSP 
   Caption         =   "Cadastro de Motivos"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxCADMOTIVOS 
         Height          =   3735
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6588
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   7935
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar"
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
         Picture         =   "frmCADMOTIVOSP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdInclui 
         Caption         =   "&Inclui"
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
         Picture         =   "frmCADMOTIVOSP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Inclui uma nova empresa"
         Top             =   120
         Width           =   735
      End
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
         Left            =   1560
         Picture         =   "frmCADMOTIVOSP.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Altera Empresa "
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdExclui 
         Caption         =   "&Exclui"
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
         Left            =   2280
         Picture         =   "frmCADMOTIVOSP.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdCanFiltro 
         Caption         =   "Desfas"
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
         Left            =   6360
         Picture         =   "frmCADMOTIVOSP.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
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
         Left            =   7080
         Picture         =   "frmCADMOTIVOSP.frx":119A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3240
         Top             =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   200
         Width           =   1695
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtCampos"
         Top             =   200
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmCADMOTIVOSP"
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
Dim objFuncoes      As Object
Dim objCADMOTIVOS   As Object
Dim iCodigo         As Integer

Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Operacao "A"
End Sub

Private Sub cmdCanFiltro_Click()
    txtCampos.Text = ""
    ConfGrid
    PreencheGrid
End Sub

Private Sub cmdExclui_Click()
  
  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub

  Dim iResp As Integer
  
  If Consistencia = False Then
     MsgBox "Existem requisi��es com este motivo relacionado !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  End If
    
  iResp = MsgBox("Confirma a exclus�o do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADMOTIVOS.GRAVA("E") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    If flxCADMOTIVOS.Rows > 1 Then Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADMOTIVOS = Nothing
    Unload Me
End Sub

Private Sub flxCADMOTIVOS_Click()
    If flxCADMOTIVOS.Rows > 1 Then objCADMOTIVOS.CODIGO = CInt(flxCADMOTIVOS.TextMatrix(flxCADMOTIVOS.RowSel, 1))
End Sub

Private Sub flxCADMOTIVOS_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCADMOTIVOS.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCADMOTIVOS_RowColChange()
    If flxCADMOTIVOS.Rows > 1 Then objCADMOTIVOS.CODIGO = CInt(flxCADMOTIVOS.TextMatrix(flxCADMOTIVOS.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
        
    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADMOTIVOS = CreateObject("CADMOTIVOS.clsCADMOTIVOS")
    
    objCADMOTIVOS.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADMOTIVOSP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "N�o foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    PreencheGrid
    
    cboFiltro.AddItem "C�digo"
    cboFiltro.AddItem "Descri��o"
    
    cboFiltro.ListIndex = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    

End Sub
Private Sub AbilitaCampos()
    
    If objCADMOTIVOS.Pesq_CadMotivos = False Then
       cmdAltera.Enabled = False
       cmdExclui.Enabled = False
       Frame1.Enabled = False
       Frame3.Enabled = False
    Else
       cmdAltera.Enabled = True
       cmdExclui.Enabled = True
       Frame1.Enabled = True
       Frame3.Enabled = True
    End If

End Sub

Private Sub ConfGrid()
    
    flxCADMOTIVOS.Rows = 1
    flxCADMOTIVOS.Cols = 3
    
    flxCADMOTIVOS.TextMatrix(0, 0) = ""
    flxCADMOTIVOS.TextMatrix(0, 1) = "C�digo"
    flxCADMOTIVOS.TextMatrix(0, 2) = "Descri��o"
    
    flxCADMOTIVOS.ColWidth(0) = 0
    flxCADMOTIVOS.ColWidth(1) = 1000
    flxCADMOTIVOS.ColWidth(2) = 4000
    
End Sub

Private Sub PreencheGrid()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOTIVOS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       
       flxCADMOTIVOS.AddItem "" & vbTab & _
                             BREC!SGI_CODIGO & vbTab & _
                             BREC!SGI_DESCRI
                               
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()
 
    If iCodigo > 0 Then
        Dim I As Integer
        
        For I = 1 To (flxCADMOTIVOS.Rows - 1)
             
            If flxCADMOTIVOS.TextMatrix(I, 1) = iCodigo Then
               flxCADMOTIVOS.Row = I
               flxCADMOTIVOS.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
    
  If flxCADMOTIVOS.Rows > 1 Then iCodigo = CInt(flxCADMOTIVOS.TextMatrix(flxCADMOTIVOS.RowSel, 1))
    
  frmCADMOTIVOS.cCaminho = cCaminho
  frmCADMOTIVOS.Linha = Linha
  frmCADMOTIVOS.iCodigo = iCodigo
  frmCADMOTIVOS.cTipOper = strOperacao
  frmCADMOTIVOS.FILIAL = FILIAL
  frmCADMOTIVOS.strAcesso = strAcesso
  frmCADMOTIVOS.Show vbModal
  
  AbilitaCampos
  ConfGrid
  PreencheGrid

End Sub


Private Sub Ordem()

  ConfGrid
  
  txtCampos.Text = ""
  
  sSql = ""
  
  If cboFiltro.ListIndex = 0 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADMOTIVOS " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_CODIGO "
  ElseIf cboFiltro.ListIndex = 1 Then
     sSql = " Select " & vbCrLf
     sSql = sSql & "        * " & vbCrLf
     sSql = sSql & "   from " & vbCrLf
     sSql = sSql & "        SGI_CADMOTIVOS " & vbCrLf
     sSql = sSql & "  Where " & vbCrLf
     sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & " Order by SGI_DESCRI "
  End If
  
  BREC.Open sSql, adoBanco_Dados
  Do While Not BREC.EOF
     flxCADMOTIVOS.AddItem "" & vbTab & _
                            BREC!SGI_CODIGO & vbTab & _
                            BREC!SGI_DESCRI
     BREC.MoveNext
  Loop
  BREC.Close

End Sub

Public Function Consistencia() As Boolean
    
    Consistencia = True
    
    '' Verificando se Existe Motivos
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADREQENTRMAT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODMOTIVO  = " & objCADMOTIVOS.CODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then Consistencia = False
    BREC.Close
    
End Function

