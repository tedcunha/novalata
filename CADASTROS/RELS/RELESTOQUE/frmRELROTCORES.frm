VERSION 5.00
Begin VB.Form frmRELROTCORES 
   Caption         =   "Relatório de Rotulos"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   13890
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
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
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   3375
      Begin VB.OptionButton optStatus 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Ativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Inativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Com cores Cadastradas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   11280
      TabIndex        =   8
      Top             =   960
      Width           =   2535
      Begin VB.OptionButton optCoresCadSN 
         Caption         =   "Sim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optCoresCadSN 
         Caption         =   "Não"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   7815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Ordem ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3375
      Begin VB.OptionButton optOrdem 
         Caption         =   "RÓTULO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optOrdem 
         Caption         =   "CLIENTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
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
         Picture         =   "frmRELROTCORES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmRELROTCORES.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELROTCORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public lngCodUsuario    As Long

Dim strNomRel       As String
Dim strCABEC2       As String
Dim objBLBFunc      As Object
Dim objRELROTCORES  As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Private Sub cmdImpressao_Click()
    If optCoresCadSN(0).Value = True Then Call RotSemCores
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELROTCORES = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    
    
    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELROTCORES = CreateObject("RELESTOQUE.clsRELROTCORES")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELROTCORES.FILIAL = FILIAL
    objBLBFunc.LimpaCampos frmRELROTCORES
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
''    Call LimpaCamposLabel
    
    optOrdem(1).Value = True
    
    Frame3.Visible = True
    Frame3.Caption = "[ Rótulo ]"
    Frame4.Visible = False
    
    optCoresCadSN(0).Value = True
    optStatus(1).Value = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub optOrdem_Click(Index As Integer)
    If Index = 0 Then
        Frame3.Visible = False
        Frame4.Visible = True
        Frame4.Caption = "[ Cliente ]"
    ElseIf Index = 1 Then
        Frame3.Visible = True
        Frame3.Caption = "[ Rótulo ]"
        Frame4.Visible = False
    End If
End Sub

Private Sub RotSemCores()

    Dim lngQTDROTSEMCOR As Long
    Dim arrROTSEMCOR()  As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_PRODUTOTIPO   = 1" & vbCrLf
    sSql = sSql & "   And SGI_PRODUTOESTILO = 0" & vbCrLf
    
    If optStatus(0).Value = True Then sSql = sSql & "   And SGI_STATUS        = 0" '' Inativo
    If optStatus(1).Value = True Then sSql = sSql & "   And SGI_STATUS        = 1" '' Ativo
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        lngQTDROTSEMCOR = 0
        Do While Not BREC.EOF()
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       Count(C.SGI_IDPRODUTO) As QtdCores " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CORESPROD C " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       C.SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And C.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "Group By C.SGI_IDPRODUTO"
            
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF() Then
                
            Else
                lngQTDROTSEMCOR = (lngQTDROTSEMCOR + 1)
            End If
            BREC2.Close
            
            BREC.MoveNext
        Loop
    End If
    BREC.Close
    
    If optCoresCadSN(0).Value = True Then
        If lngQTDROTSEMCOR = 0 Then
            MsgBox "Não existe rótulos sem cores cadastradas !!!", vbOKOnly + vbExclamation, "aviso"
            Exit Sub
        End If
        
        ReDim arrROTSEMCOR(1 To lngQTDROTSEMCOR) As String
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_PRODUTOTIPO   = 1" & vbCrLf
        sSql = sSql & "   And SGI_PRODUTOESTILO = 0" & vbCrLf
        
        If optStatus(0).Value = True Then sSql = sSql & "   And SGI_STATUS        = 0" '' Inativo
        If optStatus(1).Value = True Then sSql = sSql & "   And SGI_STATUS        = 1" '' Ativo
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
            lngQTDROTSEMCOR = 0
            Do While Not BREC.EOF()
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       Count(C.SGI_IDPRODUTO) As QtdCores " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CORESPROD C " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       C.SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & "   And C.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
                sSql = sSql & "Group By C.SGI_IDPRODUTO"
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF() Then
                    
                Else
                    lngQTDROTSEMCOR = (lngQTDROTSEMCOR + 1)
                    arrROTSEMCOR(lngQTDROTSEMCOR) = Trim(Str(BREC!SGI_IDPRODUTO))
                End If
                BREC2.Close
                
                BREC.MoveNext
            Loop
        End If
        BREC.Close
    
        objRELROTCORES.RELROTSEMCORES = arrROTSEMCOR
    
        Call objRELROTCORES.GRAVA("E")
        Call objRELROTCORES.GRAVA("I")
    
        MsgBox "Relatório Criado com exito !!!", vbOKOnly + vbExclamation, "Aviso"
    
        strNomRel = "RELROTSCOR.rpt"
        strCABEC2 = "Relatório de Rótulos sem cores cadastradas"
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_RELROTULOS.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODCLIE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODROTULO" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DIGVERIF" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO " & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO " & vbCrLf
        sSql = sSql & "      ,SGI_RELROTULOS SGI_RELROTULOS " & vbCrLf
         
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_RELROTULOS.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_RELROTULOS.SGI_FILIAL     = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_RELROTULOS.SGI_IDPRODUTO  = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN"
        
        If optStatus(0).Value = True Then sSql = sSql & "   And SGI_CADPRODUTO.SGI_STATUS        = 0" '' Inativo
        If optStatus(1).Value = True Then sSql = sSql & "   And SGI_CADPRODUTO.SGI_STATUS        = 1" '' Ativo
        
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelEstoque & strNomRel, Linha, 1, strCABEC2, "", True)
        Call objRELROTCORES.GRAVA("E")
    
    End If

End Sub
