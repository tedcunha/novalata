VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELGRAFAPGREC 
   Caption         =   "Graficos de Contas a Pagar e Receber no Periodo"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5850
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   5775
      Begin VB.Frame Frame5 
         Caption         =   "[ Periodo ]"
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
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   5055
         Begin VB.OptionButton Option2 
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
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
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
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "[ Titulos ]"
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
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   5055
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
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
            Index           =   2
            Left            =   3480
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Fechado"
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
            Index           =   1
            Left            =   1800
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Aberto"
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
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   600
         Width           =   3855
         Begin VB.OptionButton optAPGAREC 
            Caption         =   "Á Receber"
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
            Index           =   1
            Left            =   1560
            TabIndex        =   6
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton optAPGAREC 
            Caption         =   "Á Pagar"
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
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmRELGRAFAPGREC.frx":0000
         Left            =   1560
         List            =   "frmRELGRAFAPGREC.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   3735
      End
      Begin MSMask.MaskEdBox mskDtFinal 
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDtInicial 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Finial:"
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
         Left            =   2880
         TabIndex        =   2
         Top             =   285
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
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
         Left            =   240
         TabIndex        =   0
         Top             =   285
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5775
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
         Picture         =   "frmRELGRAFAPGREC.frx":003D
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmRELGRAFAPGREC.frx":013F
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELGRAFAPGREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public FILIAL        As Integer
Public strAcesso     As String
Dim objBLBFunc       As Object
Dim objRELGRAFAPGREC As Object
Dim objREL           As Object
Dim strTitulo        As String
''Dim cCamRel          As String
Dim strCABEC2        As String
Dim lngCODOPERACAO   As Long


Private Sub cmdImpressao_Click()

    If Valida_Campos = False Then Exit Sub
    
    lngCODOPERACAO = objRELGRAFAPGREC.Gera_Codigo(Me.Name)
    
    If CDate(mskDtInicial.Text) <> CDate(mskDtFinal.Text) Then
       If optAPGAREC(0).Value = True Then
          strTitulo = "Gráfico de contas a pagar no periodo de " & mskDtInicial.Text & " á " & mskDtFinal.Text
       ElseIf optAPGAREC(1).Value = True Then
          strTitulo = "Gráfico de contas a receber no periodo de " & mskDtInicial.Text & " á " & mskDtFinal.Text
       End If
    Else
       If optAPGAREC(0).Value = True Then
          strTitulo = "Gráfico de contas a pagar no dia " & mskDtInicial.Text
       ElseIf optAPGAREC(1).Value = True Then
          strTitulo = "Gráfico de contas a receber no dia " & mskDtInicial.Text
       End If
    End If
    
    If optAPGAREC(0).Value = True Then PopulaContasAPG lngCODOPERACAO
    If optAPGAREC(1).Value = True Then PopulaContasAREC lngCODOPERACAO
    
    Call ChamaRel(cboTipo.ItemData(cboTipo.ListIndex))

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELGRAFAPGREC = Nothing
    Set objREL = Nothing
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
    Set objRELGRAFAPGREC = CreateObject("RELCONTAPGTREC.clsRELGRAFAPGREC")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")

    mskDtInicial.Text = Format(Now, "DD/MM/YYYY")
    mskDtFinal.Text = Format(Now, "DD/MM/YYYY")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELGRAFAPGREC
    
    objRELGRAFAPGREC.FILIAL = FILIAL
    
    optAPGAREC(0).Value = True
    
    PreenchComboTipo 1
    
    Option1(2).Value = True
    Option2(0).Value = True
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
End Sub

Private Sub mskDtFinal_GotFocus()
    objBLBFunc.SelecionaCampos mskDtFinal.Name, frmRELGRAFAPGREC
End Sub

Private Sub mskDtInicial_GotFocus()
    objBLBFunc.SelecionaCampos mskDtInicial.Name, frmRELGRAFAPGREC
End Sub

Private Function Valida_Campos() As Boolean

    Valida_Campos = False
    
    If Not IsDate(mskDtInicial.Text) Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDtInicial.SetFocus
       Exit Function
    End If
    If Not IsDate(mskDtFinal.Text) Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDtFinal.SetFocus
       Exit Function
    End If
    If CDate(mskDtInicial.Text) > CDate(mskDtFinal.Text) Then
       MsgBox "Data inicial não pode ser maior que data final !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDtFinal.SetFocus
       Exit Function
    End If
    If cboTipo.ListIndex = -1 Then
       MsgBox "Informe o tipo do gráfico !!!", vbOKOnly + vbExclamation, "Aviso"
       cboTipo.SetFocus
       Exit Function
    End If
    
    Valida_Campos = True

End Function

Private Sub PopulaContasAPG(lngCODOPERACAO As Long)
    
On Error GoTo err_TODOS
   
   Dim sValor As String
   
   adoBanco_Dados.BeginTrans
   BGRV.ActiveConnection = adoBanco_Dados
   
   '' Contas a Pagar
   sSql = "Select " & vbCrLf
   sSql = sSql & "       SGI_CONTASIAPG.SGI_DATAVENC   " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_NUMDOC     " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_PARCELA    " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLDOC      " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_CODIGO     " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_FILIAL     " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_QTDPARC    " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_CODFOR     " & vbCrLf
   sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC    " & vbCrLf
   sSql = sSql & "      ,SGI_CADGRUPDESP.SGI_DESCRICAO " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHAPG.SGI_GRPDESP    " & vbCrLf
   
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CONTASIAPG  " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHAPG  " & vbCrLf
   sSql = sSql & "      ,SGI_CADFORNEC   " & vbCrLf
   sSql = sSql & "      ,SGI_CADGRUPDESP " & vbCrLf
   
   sSql = sSql & " Where " & vbCrLf
   
   sSql = sSql & "        SGI_CONTASIAPG.SGI_FILIAL     = " & FILIAL & vbCrLf
   If Option1(0).Value = True Then
      sSql = sSql & "   And SGI_CONTASIAPG.SGI_STATUS     = 'A' " & vbCrLf
   ElseIf Option1(1).Value = True Then
      sSql = sSql & "   And SGI_CONTASIAPG.SGI_STATUS     = 'B' " & vbCrLf
   End If
   sSql = sSql & "   And SGI_CONTASIAPG.SGI_DATAVENC Between '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' " & vbCrLf
   If Option1(0).Value = True Then
      sSql = sSql & "   And SGI_CONTASIAPG.SGI_VLPAGO IS NULL " & vbCrLf
   ElseIf Option1(1).Value = True Then
      sSql = sSql & "   And SGI_CONTASIAPG.SGI_VLPAGO IS NOT NULL " & vbCrLf
   End If
   
   sSql = sSql & "   And SGI_CONTASHAPG.SGI_FILIAL     = SGI_CONTASIAPG.SGI_FILIAL  " & vbCrLf
   sSql = sSql & "   And SGI_CONTASHAPG.SGI_CODIGO     = SGI_CONTASIAPG.SGI_CODIGO  " & vbCrLf
   sSql = sSql & "   And SGI_CADFORNEC.SGI_FILIAL      = SGI_CONTASHAPG.SGI_FILIAL  " & vbCrLf
   sSql = sSql & "   And SGI_CADFORNEC.SGI_CODIGO      = SGI_CONTASHAPG.SGI_CODFOR  " & vbCrLf
   sSql = sSql & "   And SGI_CADGRUPDESP.SGI_FILIAL    = SGI_CONTASHAPG.SGI_FILIAL  " & vbCrLf
   sSql = sSql & "   And SGI_CADGRUPDESP.SGI_CODIGO    = SGI_CONTASHAPG.SGI_GRPDESP " & vbCrLf
   sSql = sSql & " Order by SGI_CONTASIAPG.SGI_DATAVENC "
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   Do While Not BREC.EOF
   
      sSql = " Insert into SGI_TEMPCONTAPGREC( " & vbCrLf
      sSql = sSql & "                                SGI_FILIAL" & vbCrLf
      sSql = sSql & "                               ,SGI_OPERACAO" & vbCrLf
      sSql = sSql & "                               ,SGI_NUMDOC" & vbCrLf
      sSql = sSql & "                               ,SGI_DATA" & vbCrLf
      sSql = sSql & "                               ,SGI_DATAVENC" & vbCrLf
      sSql = sSql & "                               ,SGI_DATAPGTO" & vbCrLf
      sSql = sSql & "                               ,SGI_CODFORNEC" & vbCrLf
      sSql = sSql & "                               ,SGI_CODCLI" & vbCrLf
      sSql = sSql & "                               ,SGI_CODGRPDSP" & vbCrLf
      sSql = sSql & "                               ,SGI_PARCELA" & vbCrLf
      sSql = sSql & "                               ,SGI_TOTPARC" & vbCrLf
      sSql = sSql & "                               ,SGI_VLDOC" & vbCrLf
      sSql = sSql & "                               ,SGI_VLPAGO" & vbCrLf
      sSql = sSql & "                               ,SGI_VLDESC" & vbCrLf
      sSql = sSql & "                               ,SGI_VLACRESC" & vbCrLf
      sSql = sSql & "                               ,SGI_STATUS" & vbCrLf
      sSql = sSql & "                               ,SGI_TIPREL" & vbCrLf
      sSql = sSql & "                               ,SGI_RAZAO)" & vbCrLf
      sSql = sSql & "                        Values ( " & vbCrLf
      sSql = sSql & "                                 " & FILIAL & vbCrLf
      sSql = sSql & "                                ," & lngCODOPERACAO & vbCrLf
      sSql = sSql & "                                ,'" & BREC!SGI_NUMDOC & "'" & vbCrLf
      sSql = sSql & "                                ,'" & Format(BREC!SGI_DATAVENC, "MM/DD/YYYY") & "'" & vbCrLf
      sSql = sSql & "                                ,'" & Format(BREC!SGI_DATAVENC, "MM/DD/YYYY") & "'" & vbCrLf
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ," & BREC!SGI_CODFOR & vbCrLf
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ," & BREC!SGI_GRPDESP & vbCrLf
      sSql = sSql & "                                ," & BREC!SGI_PARCELA & vbCrLf
      sSql = sSql & "                                ," & BREC!SGI_QTDPARC & vbCrLf
      
      sValor = Replace((BREC!SGI_VLDOC * -1), ".", "")
      sValor = Replace(sValor, ",", ".")
      sSql = sSql & "                                ," & sValor & vbCrLf
      
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ,'A'" & vbCrLf
      sSql = sSql & "                                ,1"
      sSql = sSql & "                                ,'" & BREC!SGI_RAZAOSOC & "')"
       
      BGRV.CommandText = sSql
      BGRV.Execute
      
      BREC.MoveNext
   Loop
   
   BREC.Close
   
   adoBanco_Dados.CommitTrans
   
   Exit Sub
   
err_TODOS:

    MsgBox "Erro Nº: " & Err.Number & " ]- Descrição : " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    adoBanco_Dados.RollbackTrans
    If BREC.State = 1 Then BREC.Close
   

End Sub


Private Sub ChamaRel(intTipo As Integer)

On Error GoTo Err_Imp

    Dim strArq As String
    Dim strTitulos As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_TEMPCONTAPGREC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_OPERACAO = " & lngCODOPERACAO & vbCrLf
    sSql = sSql & "       Order by SGI_DATA "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para a impressão !!!", vbOKOnly + vbCritical, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If optAPGAREC(0).Value = True Then
       If intTipo = 1 Then
          strCABEC2 = "(Por Ano)"
          objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasGRAF & "RELGRAFAPGANO2.rpt", Linha, 1, strTitulo, strCABEC2, False
       End If
       If intTipo = 2 Then
          strCABEC2 = "(Por Mês)"
          objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasGRAF & "RELGRAFAPGMES2.rpt", Linha, 1, strTitulo, strCABEC2, False
       End If
       If intTipo = 3 Then
          strCABEC2 = "(Por Fornecedor)"
          objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasGRAF & "RELGRAFAPGFORN2.rpt", Linha, 1, strTitulo, strCABEC2, False
       End If
       If intTipo = 4 Then
       
          If Option1(0).Value = True Then strTitulos = Option1(0).Caption
          If Option1(1).Value = True Then strTitulos = Option1(1).Caption
          If Option1(2).Value = True Then strTitulos = Option1(2).Caption
       
          strCABEC2 = "(Por Grupo de Despesas) - " & strTitulos
          
          If Option2(0).Value = True Then strArq = "RELGRAFGRPDSP2.rpt"
          If Option2(1).Value = True Then strArq = "RELGRAFGRPDSPMES.rpt"
          If Option2(2).Value = True Then strArq = "RELGRAFGRPDSPANO.rpt"
          
          objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasGRAF & strArq, Linha, 1, strTitulo, strCABEC2, False
       End If
    End If
    
    If optAPGAREC(1).Value = True Then
       If intTipo = 1 Then
          strCABEC2 = "(Por Ano)"
          objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasGRAF & "RELGRAFARECANO2.rpt", Linha, 1, strTitulo, strCABEC2, False
       End If
    
       If intTipo = 2 Then
          strCABEC2 = "(Por Mês)"
          objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasGRAF & "RELGRAFARECMES2.rpt", Linha, 1, strTitulo, strCABEC2, False
       End If
    
       If intTipo = 3 Then
          strCABEC2 = "(Por Cliente)"
          objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasGRAF & "RELGRAFARECCLI2.rpt", Linha, 1, strTitulo, strCABEC2, False
       End If
       
       If intTipo = 4 Then
       
          If Option1(0).Value = True Then strTitulos = Option1(0).Caption
          If Option1(1).Value = True Then strTitulos = Option1(1).Caption
          If Option1(2).Value = True Then strTitulos = Option1(2).Caption
          
          strCABEC2 = "(Por Grupo de Recebimento) - " & strTitulos
          
          If Option2(0).Value = True Then strArq = "RELGRAFGRPREC.rpt"
          If Option2(1).Value = True Then strArq = "RELGRAFGRPRECMES.rpt"
          If Option2(2).Value = True Then strArq = "RELGRAFGRPRECANO.rpt"
             
          objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasGRAF & strArq, Linha, 1, strTitulo, strCABEC2, False
       End If
    
    End If
    
    
    '' --------------------------------------------------------------------------
    '' Apagando Tabela
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = "Delete From SGI_TEMPCONTAPGREC " & vbCrLf
    sSql = sSql & "   Where SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "     And SGI_OPERACAO = " & lngCODOPERACAO
    
    BGRV.CommandText = sSql
    BGRV.Execute
    '' --------------------------------------------------------------------------
    
    adoBanco_Dados.CommitTrans
    
    Exit Sub
    
Err_Imp:
    
    MsgBox Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    If BREC.State = 1 Then BREC.Close

End Sub


Private Sub PreenchComboTipo(intTipo As Integer)

    cboTipo.Clear
    
    If intTipo = 1 Then
       cboTipo.AddItem "Ano"
       cboTipo.ItemData(cboTipo.NewIndex) = 1
    
       cboTipo.AddItem "Mês"
       cboTipo.ItemData(cboTipo.NewIndex) = 2
    
       cboTipo.AddItem "Fornecedores"
       cboTipo.ItemData(cboTipo.NewIndex) = 3
    
       cboTipo.AddItem "Grupo de Despesas"
       cboTipo.ItemData(cboTipo.NewIndex) = 4
    End If
    
    If intTipo = 2 Then
       cboTipo.AddItem "Ano"
       cboTipo.ItemData(cboTipo.NewIndex) = 1
    
       cboTipo.AddItem "Mês"
       cboTipo.ItemData(cboTipo.NewIndex) = 2
    
       cboTipo.AddItem "Cliente"
       cboTipo.ItemData(cboTipo.NewIndex) = 3
       
       cboTipo.AddItem "Grupo de Recebimento"
       cboTipo.ItemData(cboTipo.NewIndex) = 4
    End If

End Sub

Private Sub optAPGAREC_Click(Index As Integer)
    If Index = 0 Then PreenchComboTipo 1
    If Index = 1 Then PreenchComboTipo 2
End Sub

Private Sub PopulaContasAREC(lngCODOPERACAO As Long)
    
On Error GoTo err_TODOS
   
   Dim sValor As String
   
   Dim lngReg As Long
   
   adoBanco_Dados.BeginTrans
   BGRV.ActiveConnection = adoBanco_Dados
   
   '' Contas a Receber
   '' ----------------------------------------------------------------
   '' Pegando Titulos em aberto
    
   sSql = "Select " & vbCrLf
   sSql = sSql & "       SGI_CONTASIARC.SGI_NUMDOC      " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASIARC.SGI_DATAVENC    " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASIARC.SGI_PARCELA     " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASIARC.SGI_VLDOC       " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODIGO      " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHARC.SGI_FILIAL      " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHARC.SGI_QTDPARC     " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODCLI      " & vbCrLf
   sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC    " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHARC.SGI_CODGRPRECEB " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CONTASIARC " & vbCrLf
   sSql = sSql & "      ,SGI_CONTASHARC " & vbCrLf
   sSql = sSql & "      ,SGI_CADCLIENTE " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CONTASIARC.SGI_FILIAL = " & FILIAL & vbCrLf
   sSql = sSql & "   And SGI_CONTASIARC.SGI_DATAVENC >= '" & Format(CDate(mskDtInicial.Text), "MM/DD/YYYY") & "' And SGI_CONTASIARC.SGI_DATAVENC <= '" & Format(CDate(mskDtFinal.Text), "MM/DD/YYYY") & "' " & vbCrLf
   If Option1(0).Value = True Then
      sSql = sSql & "   And SGI_CONTASIARC.SGI_VLPAGO IS NULL " & vbCrLf
   ElseIf Option1(1).Value = True Then
      sSql = sSql & "   And SGI_CONTASIARC.SGI_VLPAGO IS NOT NULL " & vbCrLf
   End If
   sSql = sSql & "   And SGI_CONTASHARC.SGI_FILIAL   = SGI_CONTASIARC.SGI_FILIAL " & vbCrLf
   sSql = sSql & "   And SGI_CONTASHARC.SGI_CODIGO   = SGI_CONTASIARC.SGI_CODIGO " & vbCrLf
   sSql = sSql & "   And SGI_CADCLIENTE.SGI_FILIAL   = SGI_CONTASHARC.SGI_FILIAL " & vbCrLf
   sSql = sSql & "   And SGI_CADCLIENTE.SGI_CODIGO   = SGI_CONTASHARC.SGI_CODCLI "

    
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
   lngReg = 1
   Do While Not BREC.EOF
    
      sSql = " Insert into SGI_TEMPCONTAPGREC( " & vbCrLf
      sSql = sSql & "                                SGI_FILIAL" & vbCrLf
      sSql = sSql & "                               ,SGI_OPERACAO" & vbCrLf
      sSql = sSql & "                               ,SGI_NUMDOC" & vbCrLf
      sSql = sSql & "                               ,SGI_DATA" & vbCrLf
      sSql = sSql & "                               ,SGI_DATAVENC" & vbCrLf
      sSql = sSql & "                               ,SGI_DATAPGTO" & vbCrLf
      sSql = sSql & "                               ,SGI_CODFORNEC" & vbCrLf
      sSql = sSql & "                               ,SGI_CODCLI" & vbCrLf
      sSql = sSql & "                               ,SGI_CODGRPDSP" & vbCrLf
      sSql = sSql & "                               ,SGI_PARCELA" & vbCrLf
      sSql = sSql & "                               ,SGI_TOTPARC" & vbCrLf
      sSql = sSql & "                               ,SGI_VLDOC" & vbCrLf
      sSql = sSql & "                               ,SGI_VLPAGO" & vbCrLf
      sSql = sSql & "                               ,SGI_VLDESC" & vbCrLf
      sSql = sSql & "                               ,SGI_VLACRESC" & vbCrLf
      sSql = sSql & "                               ,SGI_STATUS" & vbCrLf
      sSql = sSql & "                               ,SGI_TIPREL" & vbCrLf
      sSql = sSql & "                               ,SGI_RAZAO)" & vbCrLf
      sSql = sSql & "                        Values ( " & vbCrLf
      sSql = sSql & "                                 " & FILIAL & vbCrLf
      sSql = sSql & "                                ," & lngCODOPERACAO & vbCrLf
      sSql = sSql & "                                ,'" & BREC!SGI_NUMDOC & "'" & vbCrLf
      sSql = sSql & "                                ,'" & Format(BREC!SGI_DATAVENC, "MM/DD/YYYY") & "'" & vbCrLf
      sSql = sSql & "                                ,'" & Format(BREC!SGI_DATAVENC, "MM/DD/YYYY") & "'" & vbCrLf
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ," & BREC!SGI_CODCLI & vbCrLf
      sSql = sSql & "                                ," & BREC!SGI_CODGRPRECEB & vbCrLf
      sSql = sSql & "                                ," & BREC!SGI_PARCELA & vbCrLf
      sSql = sSql & "                                ," & BREC!SGI_QTDPARC & vbCrLf
      
      sValor = Replace((BREC!SGI_VLDOC), ".", "")
      sValor = Replace(sValor, ",", ".")
      sSql = sSql & "                                ," & sValor & vbCrLf
      
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ,Null" & vbCrLf
      sSql = sSql & "                                ,'A'" & vbCrLf
      sSql = sSql & "                                ,2" & vbCrLf
      sSql = sSql & "                                ,'" & BREC!SGI_RAZAOSOC & "')"
      
      lngReg = lngReg + 1
      
      BGRV.CommandText = sSql
      BGRV.Execute
       
      BREC.MoveNext
   Loop
   
   BREC.Close
   
   adoBanco_Dados.CommitTrans
   
   Exit Sub
   
err_TODOS:

    MsgBox "Erro Nº: " & Err.Number & " - Dewscrição : " & Err.Description & " Seq " & lngReg, vbOKOnly + vbCritical, "Aviso"
    adoBanco_Dados.RollbackTrans
    If BREC.State = 1 Then BREC.Close
   

End Sub

