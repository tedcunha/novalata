VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADDIASUTEIS 
   Caption         =   "Cadastro de Dias Úteis"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   6735
      Begin VB.TextBox txtTotMin 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Text            =   "txtTotMin"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtTotHoras 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Text            =   "txtTotHoras"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox totDias 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Text            =   "totDias"
         Top             =   1320
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDTENVENTO 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   975
      End
      Begin MSMask.MaskEdBox mskDtFinal 
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Top             =   960
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
         Caption         =   "Total em Minutos"
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
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total em Horas"
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
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total em Dias"
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
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Final :"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data do Evento:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1410
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   975
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6735
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
         Left            =   840
         Picture         =   "frmCADDIASUTEIS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   735
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
         Left            =   1560
         Picture         =   "frmCADDIASUTEIS.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   735
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
         Picture         =   "frmCADDIASUTEIS.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADDIASUTEIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public FILIAL           As Integer
Public strAcesso        As String
Dim objBLBFunc          As Object
Dim objCADDIASUTEIS     As Object


Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Dias Úteis - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
       
    If cTipOper = "I" Then objCADDIASUTEIS.CODIGO = objCADDIASUTEIS.Gera_Codigo(Me.Name)
       
    objCADDIASUTEIS.DESCRI = Trim(txtDescricao.Text)
    objCADDIASUTEIS.DIASEVENTO = CDate(mskDTENVENTO.Text)
    objCADDIASUTEIS.DATAFINAL = CDate(mskDtFinal.Text)
    objCADDIASUTEIS.TOTDIAS = CLng(TOTDIAS.Text)
    objCADDIASUTEIS.TOTHORAS = CLng(txtTotHoras.Text)
    objCADDIASUTEIS.TOTMIN = CLng(txtTotMin.Text)
       
    If objCADDIASUTEIS.GRAVA(cTipOper) = False Then Exit Sub
    If objCADDIASUTEIS.Atualiza(cTipOper, Str(objCADDIASUTEIS.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
          
    MsgBox "O dia útil foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADDIASUTEIS = Nothing
       Unload Me
    End If
          
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADDIASUTEIS = Nothing
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
    Set objCADDIASUTEIS = CreateObject("CADDIASUTEIS.clsCADDIASUTEIS")
    
    objCADDIASUTEIS.FILIAL = FILIAL
    
    If cTipOper = "I" Then Inclui
    If cTipOper = "A" Then Altera
    If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Dias Úteis - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADDIASUTEIS
    
    txtCodigo.Text = ""
   
End Sub
Private Sub mskDTENVENTO_GotFocus()
    objBLBFunc.SelecionaCampos mskDTENVENTO.Name, frmCADDIASUTEIS
End Sub

Private Sub mskDTENVENTO_Validate(Cancel As Boolean)
    TOTDIAS.Text = CalcDias(mskDtFinal.Text, mskDTENVENTO.Text)
    txtTotHoras.Text = CalcHoras(CLng(TOTDIAS.Text))
    txtTotMin.Text = CalcMin(CLng(txtTotHoras.Text))
End Sub

Private Sub mskDtFinal_GotFocus()
    objBLBFunc.SelecionaCampos mskDtFinal.Name, frmCADDIASUTEIS
End Sub

Private Sub mskDtFinal_Validate(Cancel As Boolean)
    TOTDIAS.Text = CalcDias(mskDtFinal.Text, mskDTENVENTO.Text)
    txtTotHoras.Text = CalcHoras(CLng(TOTDIAS.Text))
    txtTotMin.Text = CalcMin(CLng(txtTotHoras.Text))
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADDIASUTEIS
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub


Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição do evento Inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     If Not IsDate(mskDTENVENTO.Text) Then
        MsgBox "Data Inicial do evento inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTENVENTO.SetFocus
        Exit Function
     End If
     If Not IsDate(mskDtFinal.Text) Then
        MsgBox "Data Final do evento inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDtFinal.SetFocus
        Exit Function
     End If
     If CDate(mskDTENVENTO.Text) > CDate(mskDtFinal.Text) Then
        MsgBox "Data Inicial não pode ser maior que data final !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTENVENTO.SetFocus
        Exit Function
     End If
     
     
     If cTipOper = "I" Then
        
        '' -----------------------------------
        sSql = "Select * from SGI_CADDIASUTEIS Where SGI_DATENVENTO ='" & Format(CDate(mskDTENVENTO.Text), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & " And SGI_FILIAL = " & objCADDIASUTEIS.FILIAL
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "A data do evento já existe !!!", vbOKOnly + vbCritical, "Aviso"
           mskDTENVENTO.SetFocus
           BREC.Close
           Exit Function
        End If
        BREC.Close
        '' -----------------------------------
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADDIASUTEIS.DIASEVENTO <> CDate(mskDTENVENTO.Text) Then
           sSql = "Select * from SGI_CADDIASUTEIS Where SGI_DATENVENTO ='" & Format(CDate(mskDTENVENTO.Text), "MM/DD/YYYY") & "'" & vbCrLf
           sSql = sSql & " And SGI_FILIAL = " & objCADDIASUTEIS.FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "A data do evento já existe !!!", vbOKOnly + vbCritical, "Aviso"
              mskDTENVENTO.Text = Format(objCADDIASUTEIS.DIASEVENTO, "DD/MM/YYYY")
              mskDTENVENTO.SetFocus
              BREC.Close
              Exit Function
           End If
           BREC.Close
        End If
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de Dias Úteis - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADDIASUTEIS
    
    objCADDIASUTEIS.CODIGO = iCodigo
    
    If objCADDIASUTEIS.Carrega_campos = True Then
       txtCodigo.Text = objCADDIASUTEIS.CODIGO
       txtDescricao.Text = objCADDIASUTEIS.DESCRI
       mskDTENVENTO.Text = Format(objCADDIASUTEIS.DIASEVENTO, "DD/MM/YYYY")
       mskDtFinal.Text = Format(objCADDIASUTEIS.DATAFINAL, "DD/MM/YYYY")
    
       TOTDIAS.Text = CalcDias(mskDtFinal.Text, mskDTENVENTO.Text)
       txtTotHoras.Text = CalcHoras(CLng(TOTDIAS.Text))
       txtTotMin.Text = CalcMin(CLng(txtTotHoras.Text))
    
    End If

End Sub

Public Sub Altera()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Dias Úteis - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADDIASUTEIS
    
    objCADDIASUTEIS.CODIGO = iCodigo
    
    If objCADDIASUTEIS.Carrega_campos = True Then
       txtCodigo.Text = objCADDIASUTEIS.CODIGO
       txtDescricao.Text = objCADDIASUTEIS.DESCRI
       mskDTENVENTO.Text = Format(objCADDIASUTEIS.DIASEVENTO, "DD/MM/YYYY")
       mskDtFinal.Text = Format(objCADDIASUTEIS.DATAFINAL, "DD/MM/YYYY")
       
       TOTDIAS.Text = CalcDias(mskDtFinal.Text, mskDTENVENTO.Text)
       txtTotHoras.Text = CalcHoras(CLng(TOTDIAS.Text))
       txtTotMin.Text = CalcMin(CLng(txtTotHoras.Text))
    End If
    
End Sub

Private Function CalcDias(dtFinal As String, dtInicial As String) As Long
    
    CalcDias = 0
    
    If Not IsDate(dtInicial) Or Not IsDate(dtFinal) Then Exit Function
    If CDate(dtInicial) > CDate(dtFinal) Then
       MsgBox "Data Inicial não pode se maior que data Final !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    
    If CDate(dtFinal) = CDate(dtInicial) Then
       CalcDias = 1
    Else
       CalcDias = (CDate(dtFinal) - CDate(dtInicial))
    End If
    
End Function

Private Function CalcHoras(lngDias As Long) As Long
    
    CalcHoras = 0
    
    If lngDias = 0 Then Exit Function
    
    CalcHoras = (lngDias * 24)
    
End Function

Private Function CalcMin(lngHoras As Long) As Long
    
    CalcMin = 0
    
    If lngHoras = 0 Then Exit Function
    
    CalcMin = (lngHoras * 60)
    
End Function

