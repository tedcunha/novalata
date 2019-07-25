VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELMAPAPED 
   Caption         =   "Relatório mapa de vendas"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   11745
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      Caption         =   "[ Status ]"
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
      Height          =   1095
      Left            =   4320
      TabIndex        =   26
      Top             =   2160
      Width           =   7335
      Begin VB.OptionButton optSTATUS 
         Caption         =   "Para Estoque"
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
         Index           =   7
         Left            =   3480
         TabIndex        =   34
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton optSTATUS 
         Caption         =   "Aguardando Liberação do Fotolito"
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
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton optSTATUS 
         Caption         =   "Bloqueados"
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
         Index           =   5
         Left            =   4920
         TabIndex        =   32
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optSTATUS 
         Caption         =   "Faturados"
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
         Index           =   4
         Left            =   3480
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optSTATUS 
         Caption         =   "Reprovados"
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
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optSTATUS 
         Caption         =   "Aguardando Liberação"
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
         Index           =   2
         Left            =   4920
         TabIndex        =   29
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optSTATUS 
         Caption         =   "Liberados"
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
         Index           =   1
         Left            =   3480
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optSTATUS 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "[ Com NECK-IN ]"
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
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   2160
      Width           =   4335
      Begin VB.OptionButton optCOMNECKIN 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCOMNECKIN 
         Caption         =   "NÃO"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCOMNECKIN 
         Caption         =   "SIM"
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
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Filtro da Data ]"
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
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   1560
      Width           =   8415
      Begin VB.OptionButton optTipoFiltData 
         Caption         =   "Por Data de Entrega do Pedido"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   21
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton optTipoFiltData 
         Caption         =   "Por Data de Emissão do Pedido"
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
         Left            =   840
         TabIndex        =   20
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Empresa ]"
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
      Height          =   615
      Left            =   8400
      TabIndex        =   16
      Top             =   1560
      Width           =   3255
      Begin VB.OptionButton optEmpresa 
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optEmpresa 
         Caption         =   "STEEL ROL"
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
         Index           =   1
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Tipo ]"
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
      Height          =   615
      Left            =   9120
      TabIndex        =   13
      Top             =   960
      Width           =   2535
      Begin VB.OptionButton optTipo 
         Caption         =   "Sintético"
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
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Análitico"
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
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Agrupamento ]"
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
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   960
      Width           =   3735
      Begin VB.OptionButton optAgrup 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAgrup 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optAgrup 
         Caption         =   "Semana"
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
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optAgrup 
         Caption         =   "Dia"
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
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   5415
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
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
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2880
         TabIndex        =   7
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11655
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
         Picture         =   "frmRELMAPAPED.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmRELMAPAPED.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELMAPAPED"
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
Dim objRELMAPAPED   As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object
Dim strCABEC1       As String
Dim strCABEC2       As String
Dim strNomRel       As String
Dim strEMPRESADESC  As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    If optTipoFiltData(0).Value = True Then Call ImprimeMapaDiaDia
    If optTipoFiltData(1).Value = True Then Call ImprimeMapaDiaDia2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELMAPAPED = CreateObject("RELCOMERCIAL.clsRELMAPAPED")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELMAPAPED
    
    objRELMAPAPED.FILIAL = FILIAL

    optAgrup(0).Value = True
    optTipo(0).Value = True
    optTipoFiltData(0).Value = True
    optCOMNECKIN(2).Value = True
    optSTATUS(0).Value = True
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    optEmpresa(0).Value = True
    
    
    Me.Caption = Me.Caption & " / " & Trim(Me.Name)

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELMAPAPED = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
    If Not IsDate(mskDTINI.Text) Then
        MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTINI.SetFocus
        Exit Function
    End If
    If Not IsDate(mskDTFIN.Text) Then
        MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTFIN.SetFocus
        Exit Function
    End If
    
    If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
        MsgBox "Data Inicial não pode ser maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTINI.SetFocus
        Exit Function
    End If
    
    ConfereCampos = True

End Function


Private Sub ImprimeMapaDiaDia()

    Dim strTABELA       As String
    Dim strSTATUS       As String
    Dim strCABSTATUS    As String
    
    strTABELA = ""
    If optEmpresa(1).Value = True Then strTABELA = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "          SGI_CADPEDVENDI" & strTABELA & ".SGI_CODIGO" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strTABELA & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strTABELA & ".SGI_CODPROD" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strTABELA & ".SGI_QTDE" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strTABELA & ".SGI_VLUNIT" & vbCrLf
    
    sSql = sSql & "         ,SGI_CADPEDVENDH" & strTABELA & ".SGI_DATAPED" & vbCrLf
    sSql = sSql & "         ,SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "         ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "         ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
    sSql = sSql & "         ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "        ,SGI_CADPEDVENDH" & strTABELA & " SGI_CADPEDVENDH" & strTABELA & vbCrLf
    sSql = sSql & "        ,SGI_CADPEDVENDI" & strTABELA & " SGI_CADPEDVENDI" & strTABELA & vbCrLf
    sSql = sSql & "        ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_CADPEDVENDI" & strTABELA & ".SGI_FILIAL = " & FILIAL & vbCrLf
    
    If optEmpresa(0).Value = True Then
        strEMPRESADESC = "NOVALATA"
    ElseIf optEmpresa(1).Value = True Then
        strEMPRESADESC = "STEEL ROL"
    End If
    
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strTABELA & ".SGI_FILIAL = SGI_CADPEDVENDH" & strTABELA & ".SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strTABELA & ".SGI_CODIGO = SGI_CADPEDVENDH" & strTABELA & ".SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strTABELA & ".SGI_FILIAL = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strTABELA & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    sSql = sSql & "   And  SGI_CADPRODUTO.SGI_FILIAL = SGI_CADLINHAPRODUTO.SGI_FILIAL"
    sSql = sSql & "   And  SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN"
    
    If optCOMNECKIN(0).Value = True Then
       sSql = sSql & "   And  SGI_CADPRODUTO.SGI_NECKIN = 1" & vbCrLf   '' Com NECK-IN
    ElseIf optCOMNECKIN(0).Value = True Then
       sSql = sSql & "   And  SGI_CADPRODUTO.SGI_NECKIN = 0" & vbCrLf   '' Sem NECK-IN
    End If
        
    sSql = sSql & "   And  SGI_CADPEDVENDH" & strTABELA & ".SGI_DATAPED BETWEEN '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'"
    
     ''Status
    strSTATUS = ""
    strCABSTATUS = " Status : Todos"
    If optSTATUS(1).Value = True Then       '' Liberados
        strSTATUS = ".SGI_STATUS = 'L')"
        strCABSTATUS = " Status : " & optSTATUS(1).Caption
    ElseIf optSTATUS(2).Value = True Then   '' Aguardando Liberação
        strSTATUS = ".SGI_STATUS = 'B' Or SGI_CADPEDVENDH" & strTABELA & ".SGI_STATUS = 'N')"
        strCABSTATUS = " Status : " & optSTATUS(2).Caption
    ElseIf optSTATUS(3).Value = True Then   '' Reprovados
        strSTATUS = ".SGI_STATUS = 'R')"
        strCABSTATUS = " Status : " & optSTATUS(3).Caption
    ElseIf optSTATUS(4).Value = True Then   '' Faturados
        strSTATUS = ".SGI_STATUS = 'F' Or SGI_CADPEDVENDH" & strTABELA & ".SGI_STATUS = 'P' Or SGI_CADPEDVENDH" & strTABELA & ".SGI_STATUS = 'M')"
        strCABSTATUS = " Status : " & optSTATUS(4).Caption
    ElseIf optSTATUS(5).Value = True Then   '' Bloqueados
        strSTATUS = ".SGI_STATUS = 'S')"
        strCABSTATUS = " Status : " & optSTATUS(5).Caption
    ElseIf optSTATUS(6).Value = True Then   '' Aguardando Liberação do Fotolito
        strSTATUS = ".SGI_STATUS = 'V')"
        strCABSTATUS = " Status : " & optSTATUS(6).Caption
    ElseIf optSTATUS(7).Value = True Then   '' Para Estoque
        strSTATUS = ".SGI_STATUS = 'X')"
        strCABSTATUS = " Status : " & optSTATUS(7).Caption
    End If
    
    If Len(Trim(strSTATUS)) > 0 Then
        sSql = sSql & "   And  (SGI_CADPEDVENDH" & strTABELA & Trim(strSTATUS)
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF() Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Mapa de pedidos / " & strEMPRESADESC
    
    If optTipo(0).Value = True Then strCABEC1 = strCABEC1 & " [ Análitico/"
    If optTipo(1).Value = True Then strCABEC1 = strCABEC1 & " [ Sintético/"
    
    If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text

    strNomRel = ""
    If optTipo(0).Value = True Then '' Relatórios Análiticos
        If optEmpresa(0).Value = True Then '' Novalata
            If optAgrup(0).Value = True Then strNomRel = "RELMAPAVEND01.rpt"  '' Agrupamento por Dia
            If optAgrup(1).Value = True Then strNomRel = "RELMAPAVEND02.rpt"  '' Agrupamento por Semana
            If optAgrup(2).Value = True Then strNomRel = "RELMAPAVEND03.rpt"  '' Agrupamento por Mês
            If optAgrup(3).Value = True Then strNomRel = "RELMAPAVEND04.rpt"  '' Agrupamento por Ano
        ElseIf optEmpresa(1).Value = True Then '' Steel Rol
            If optAgrup(0).Value = True Then strNomRel = "RELMAPAVEND01_STEEL.rpt"  '' Agrupamento por Dia
            If optAgrup(1).Value = True Then strNomRel = "RELMAPAVEND02_STEEL.rpt"  '' Agrupamento por Semana
            If optAgrup(2).Value = True Then strNomRel = "RELMAPAVEND03_STEEL.rpt"  '' Agrupamento por Mês
            If optAgrup(3).Value = True Then strNomRel = "RELMAPAVEND04_STEEL.rpt"  '' Agrupamento por Ano
        End If
    End If

    If optAgrup(0).Value = True Then strCABEC1 = strCABEC1 & "Dia ]" & strCABSTATUS        '' Agrupamento por Dia
    If optAgrup(1).Value = True Then strCABEC1 = strCABEC1 & "Semana ]" & strCABSTATUS     '' Agrupamento por Semana
    If optAgrup(2).Value = True Then strCABEC1 = strCABEC1 & "Mês ]" & strCABSTATUS        '' Agrupamento por Mês
    If optAgrup(3).Value = True Then strCABEC1 = strCABEC1 & "Ano ]" & strCABSTATUS        '' Agrupamento por Ano

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If

End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELMAPAPED
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELMAPAPED
End Sub

Private Sub ImprimeMapaDiaDia2()

    Dim strTABELA As String
    
    strTABELA = ""
    If optEmpresa(1).Value = True Then strTABELA = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_ORDEMPROD" & strTABELA & ".SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "       ,SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "       ,SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "       ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
    sSql = sSql & "       ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "       ,SGI_ORDEMPROD" & strTABELA & ".SGI_CODPED" & vbCrLf
    sSql = sSql & "       ,SGI_ORDEMPROD" & strTABELA & ".SGI_CODPROD" & vbCrLf
    sSql = sSql & "       ,SGI_ORDEMPROD" & strTABELA & ".SGI_QTDEPED" & vbCrLf
    
    sSql = sSql & "       ,SGI_CADPEDVENDI" & strTABELA & ".SGI_VLUNIT" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "       ,SGI_CADPEDVENDI" & strTABELA & " SGI_CADPEDVENDI" & strTABELA & vbCrLf
    sSql = sSql & "       ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "       ,SGI_ORDEMPROD" & strTABELA & " SGI_ORDEMPROD" & strTABELA & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_ORDEMPROD" & strTABELA & ".SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And  SGI_ORDEMPROD" & strTABELA & ".SGI_DATENTREGA BETWEEN '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    
    sSql = sSql & "   And  SGI_ORDEMPROD" & strTABELA & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And  SGI_ORDEMPROD" & strTABELA & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And  SGI_ORDEMPROD" & strTABELA & ".SGI_FILIAL    = SGI_CADPEDVENDI" & strTABELA & ".SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And  SGI_ORDEMPROD" & strTABELA & ".SGI_IDPRODUTO = SGI_CADPEDVENDI" & strTABELA & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And  SGI_ORDEMPROD" & strTABELA & ".SGI_CODPED    = SGI_CADPEDVENDI" & strTABELA & ".SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And  SGI_CADPRODUTO.SGI_FILIAL      = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And  SGI_CADPRODUTO.SGI_CODLINPROD  = SGI_CADLINHAPRODUTO.SGI_CODLIN"
    
    If optEmpresa(0).Value = True Then
        strEMPRESADESC = "NOVALATA"
    ElseIf optEmpresa(1).Value = True Then
        strEMPRESADESC = "STEEL ROL"
    End If
    
    If optCOMNECKIN(0).Value = True Then
       sSql = sSql & "   And  SGI_CADPRODUTO.SGI_NECKIN = 1" & vbCrLf   '' Com NECK-IN
    ElseIf optCOMNECKIN(0).Value = True Then
       sSql = sSql & "   And  SGI_CADPRODUTO.SGI_NECKIN = 0" & vbCrLf   '' Sem NECK-IN
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Mapa de pedidos / " & strEMPRESADESC
    
    If optTipo(0).Value = True Then strCABEC1 = strCABEC1 & " [ Análitico/"
    If optTipo(1).Value = True Then strCABEC1 = strCABEC1 & " [ Sintético/"
    
    If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text

    strNomRel = ""
    If optTipo(0).Value = True Then '' Relatórios Análiticos
        If optEmpresa(0).Value = True Then '' Novalata
            If optAgrup(0).Value = True Then strNomRel = "RELMAPAVEND01_DTENT.rpt"  '' Agrupamento por Dia
            If optAgrup(1).Value = True Then strNomRel = "RELMAPAVEND02_DTENT.rpt"  '' Agrupamento por Semana
            If optAgrup(2).Value = True Then strNomRel = "RELMAPAVEND03_DTENT.rpt"  '' Agrupamento por Mês
            If optAgrup(3).Value = True Then strNomRel = "RELMAPAVEND04_DTENT.rpt"  '' Agrupamento por Ano
        ElseIf optEmpresa(1).Value = True Then '' Steel Rol
            If optAgrup(0).Value = True Then strNomRel = "RELMAPAVEND01_STEELDTE.rpt"  '' Agrupamento por Dia
            If optAgrup(1).Value = True Then strNomRel = "RELMAPAVEND02_STEELDTE.rpt"  '' Agrupamento por Semana
            If optAgrup(2).Value = True Then strNomRel = "RELMAPAVEND03_STEELDTE.rpt"  '' Agrupamento por Mês
            If optAgrup(3).Value = True Then strNomRel = "RELMAPAVEND04_STEELDTE.rpt"  '' Agrupamento por Ano
        End If
    End If

    If optAgrup(0).Value = True Then strCABEC1 = strCABEC1 & "Dia ]"        '' Agrupamento por Dia
    If optAgrup(1).Value = True Then strCABEC1 = strCABEC1 & "Semana ]"     '' Agrupamento por Semana
    If optAgrup(2).Value = True Then strCABEC1 = strCABEC1 & "Mês ]"        '' Agrupamento por Mês
    If optAgrup(3).Value = True Then strCABEC1 = strCABEC1 & "Ano ]"        '' Agrupamento por Ano

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If

End Sub

