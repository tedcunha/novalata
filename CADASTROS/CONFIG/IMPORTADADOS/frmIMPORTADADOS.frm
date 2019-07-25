VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmIMPORTADADOS 
   Caption         =   "Importador de Dados"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   10860
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "QBG1"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Frame fraDados 
      Caption         =   "[ Carregando Arquivos ]"
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
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   10815
      Begin MSComctlLib.ProgressBar prbDADOS 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "De Para Cores"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   3135
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
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
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
         Picture         =   "frmIMPORTADADOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmIMPORTADADOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho  As String
Public Linha     As Variant
Public FILIAL    As Integer
Public strAcesso As String
Dim objBLBFunc   As Object

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    
On Error GoTo err_Trans
    
    Dim strCAMINHO      As String
    Dim strARQUIVO      As String
    Dim arrCAMPO()      As String
    Dim arrDADOS()      As String
    Dim lngDADOS        As Long
    Dim lngCONTADOR     As Long
    
    strCAMINHO = "C:\Ricardo\SGI\NOVALATA\DOCS-NOVALATA\Luigi"
    strARQUIVO = "De_Para_SIGE.txt"
    
    ''Open App.Path & "\" & "SIGE.txt" For Input As #1
    Open strCAMINHO & "\" & strARQUIVO For Input As #1
    
    lngDADOS = 1
    
    Do While Not EOF(1)
       
       ReDim Preserve arrDADOS(1 To lngDADOS) As String
       
       Input #1, arrDADOS(lngDADOS)
         
       lngDADOS = (lngDADOS + 1)
         
    Loop
      
    Close #1
    ' --------------------------------------
     
    If IsArray(arrDADOS) Then
    
        prbDADOS.Min = 0
        prbDADOS.Max = UBound(arrDADOS)
        
        Call AbiliDesCampos(True)
        
        '' Url Generaion Iron
        ''http://www.youtube.com/watch?v=1gmb8VDnwHA
        
        '' Inicia transação
        adoBanco_Dados.BeginTrans
        BGRV.ActiveConnection = adoBanco_Dados
        
        
        For lngCONTADOR = 1 To UBound(arrDADOS)
        
            prbDADOS.Value = lngCONTADOR
            
            arrCAMPO = Split(arrDADOS(lngCONTADOR), vbTab)
            
            If UBound(arrCAMPO) = 2 Then
                '' Trocando os Arquivos
                If Len(Trim(arrCAMPO(0))) > 0 And Len(Trim(arrCAMPO(2))) > 0 Then
                    
                    sSql = ""
                    
                    sSql = "Update SGI_CADPRODUTO Set " & vbCrLf
                    sSql = sSql & "                          SGI_CODIGO = '" & Trim(arrCAMPO(2)) & "'" & vbCrLf
                    sSql = sSql & "Where" & vbCrLf
                    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
                    sSql = sSql & "  And SGI_CODIGO = '" & Trim(arrCAMPO(0)) & "'"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    
                    
                Else
                        
                    ''sSql = "Update SGI_CADPRODUTO Set " & vbCrLf
                    ''sSql = sSql & "                          SGI_STATUS = 2" & vbCrLf
                    ''sSql = sSql & "Where" & vbCrLf
                    ''sSql = sSql & "      SGI_FILIAL    = " & FILIAL & vbCrLf
                    ''sSql = sSql & "  And SGI_IDPRODUTO = " & Trim(arrCAMPO(0)) & "'"
                
                End If
            End If
            
        
        Next lngCONTADOR
    
        adoBanco_Dados.CommitTrans
        
        Call AbiliDesCampos(False)
        
        MsgBox "Dados Importados com Exito !!!", vbOKOnly + vbInformation, "Aviso"
    
    Else
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Não há dados para importar !!!", vbOKOnly + vbExclamation, "Aviso"
               
    End If

    Exit Sub

err_Trans:
    
    adoBanco_Dados.RollbackTrans

    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "R", sSql)
    Set objErro = Nothing

End Sub

Private Sub Command2_Click()

On Error GoTo Err_Dados

    Dim lngQTDREGS As Long
    Dim strEMPRESA As String
    
    
    strEMPRESA = ""
    If optEmpresa(1).Value = True Then strEMPRESA = "_STEEL"
    
    lngQTDREGS = 0

    sSql = ""

    sSql = "Select" & vbCrLf
    sSql = sSql & "       ORDP.*" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "        SGI_PROGENTRPROD" & strEMPRESA & " PROGE" & vbCrLf
    sSql = sSql & "       ,SGI_ORDEMPROD" & strEMPRESA & "    ORDP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "        PROGE.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "  And   PROGE.SGI_STATUS   = 0" & vbCrLf
    sSql = sSql & "  And   ORDP.SGI_FILIAL    = PROGE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And   ORDP.SGI_IDPRODUTO = PROGE.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "  And   ORDP.SGI_CODPED    = PROGE.SGI_CODPED" & vbCrLf
    sSql = sSql & "  And   (ORDP.SGI_STATUS   = 6 or ORDP.SGI_STATUS = 7)" & vbCrLf
    sSql = sSql & "Order By ORDP.SGI_CODPED,ORDP.SGI_CODIGO"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        fraDados.Visible = True
        lngQTDREGS = 0
        Do While Not BREC.EOF()
            lngQTDREGS = (lngQTDREGS + 1)
            BREC.MoveNext
        Loop
        
        prbDADOS.Min = 0
        prbDADOS.Max = lngQTDREGS
        
        
        lngQTDREGS = 0
        BREC.MoveFirst
        
        '' Inicia transação
        adoBanco_Dados.BeginTrans
        BGRV.ActiveConnection = adoBanco_Dados
        
        Do While Not BREC.EOF()
            lngQTDREGS = (lngQTDREGS + 1)
            prbDADOS.Value = lngQTDREGS
            
            sSql = "Update SGI_ORDEMPROD" & strEMPRESA & vbCrLf
            sSql = sSql & "                     Set SGI_STATUS = 0" & vbCrLf
            sSql = sSql & "               Where " & vbCrLf
            sSql = sSql & "                     SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "                 And SGI_CODIGO = " & BREC!SGI_CODIGO
            
            BGRV.CommandText = sSql
            BGRV.Execute
            
            
            BREC.MoveNext
        Loop
        
        adoBanco_Dados.CommitTrans
        
        MsgBox "Dados Processados com exito !!!!", vbOKOnly + vbInformation, "Aviso"
        
    Else
        MsgBox "Não há dados para Processar !", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

    fraDados.Visible = False
    
    Exit Sub

Err_Dados:

    If adoBanco_Dados.State = 1 Then adoBanco_Dados.RollbackTrans

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")


   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
    
    
    Call AbiliDesCampos(False)
    
    prbDADOS.Min = 0
    optEmpresa(0).Value = True

End Sub

Private Sub AbiliDesCampos(AtivoSN As Boolean)
    fraDados.Visible = AtivoSN
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub
