VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELGRAFICOS 
   Caption         =   "Gráficos"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6045
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
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
         Picture         =   "frmRELCOMGRAFICOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         Picture         =   "frmRELCOMGRAFICOS.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
   Begin TabDlg.SSTab stTabGrfs 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cotações"
      TabPicture(0)   =   "frmRELCOMGRAFICOS.frx":0204
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "frmRELCOMGRAFICOS.frx":0220
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   5655
         Begin VB.OptionButton optGRFCOTA 
            Caption         =   "Ano"
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
            Index           =   2
            Left            =   3960
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optGRFCOTA 
            Caption         =   "Mês"
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
            Index           =   1
            Left            =   2040
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optGRFCOTA 
            Caption         =   "Periodo"
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
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5655
         Begin MSMask.MaskEdBox mskDTFIN 
            Height          =   285
            Left            =   3960
            TabIndex        =   5
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDTINI 
            Height          =   285
            Left            =   1320
            TabIndex        =   6
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
            Caption         =   "Data Final"
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
            Left            =   2880
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Inicial"
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
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmRELGRAFICOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Dim objBLBFunc      As Object
Dim objRELGRFCOM    As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object
Dim strCABEC1       As String
Dim strCABEC2       As String

Private Sub cmdImpressao_Click()

    If stTabGrfs.Tab = 0 Then ImpGRFCOTA
    
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELGRFCOM = Nothing
    Set objPESQPADRAO = Nothing
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
    Set objRELGRFCOM = CreateObject("RELCOMERCIAL.clsRELGRFCOM")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELGRAFICOS
    
    objRELGRFCOM.FILIAL = FILIAL

    stTabGrfs.Tab = 0
    
    optGRFCOTA(0).Value = True
    
    mskDTINI.Text = Format(Date, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Date + 30, "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
End Sub

Private Sub ImpGRFCOTA()
    
    If Not IsDate(mskDTINI.Text) Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTINI.SetFocus
       Exit Sub
    End If
    If Not IsDate(mskDTFIN.Text) Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTFIN.SetFocus
       Exit Sub
    End If
    
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADCOTAVENDH.SGI_CODTIPORC "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_STATUS "
    sSql = sSql & "     , SGI_CADCOTAVENDH.SGI_VLTOT "
    sSql = sSql & "     , SGI_CADESPORCA.SGI_DESCRICAO "
    
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADCOTAVENDH SGI_CADCOTAVENDH "
    sSql = sSql & "     , SGI_CADESPORCA SGI_CADESPORCA "
    
    sSql = sSql & " Where "
    
    sSql = sSql & "        SGI_CADCOTAVENDH.SGI_FILIAL = SGI_CADESPORCA.SGI_FILIAL "
    sSql = sSql & "  And   SGI_CADCOTAVENDH.SGI_CODTIPORC = SGI_CADESPORCA.SGI_CODIGO "
    
    sSql = sSql & "  And   SGI_CADCOTAVENDH.SGI_FILIAL = " & FILIAL
    sSql = sSql & "  And  (SGI_CADCOTAVENDH.SGI_DATACOTA >= '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And SGI_CADCOTAVENDH.SGI_DATACOTA <= '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "')"
    sSql = sSql & " Order by SGI_CADESPORCA.SGI_DESCRICAO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Gráfico de Cotações "
    
    If optGRFCOTA(0).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text
    ElseIf optGRFCOTA(1).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Mês " & Format(Month(CDate(mskDTINI.Text)), "##00") & "/" & Year(CDate(mskDTINI.Text)) & " ao Mês " & Format(Month(CDate(mskDTFIN.Text)), "##00") & "/" & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "No Mês " & Format(Month(CDate(mskDTINI.Text)), "##00") & "/" & Year(CDate(mskDTFIN.Text))
    ElseIf optGRFCOTA(2).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "Do Ano " & Year(CDate(mskDTINI.Text)) & " ao Ano " & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "No Ano " & Year(CDate(mskDTFIN.Text))
    End If
    
    '' Chamada do Relatório
    If optGRFCOTA(0).Value = True Then
       Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RELGRFCOTP.rpt", Linha, 1, strCABEC1, strCABEC2, True)
    End If
    
End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELGRAFICOS
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELGRAFICOS
End Sub
