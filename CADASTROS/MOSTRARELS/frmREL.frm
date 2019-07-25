VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmREL 
   Caption         =   "Relatório"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   9135
      Begin VB.TextBox txtRel 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Text            =   "frmREL.frx":0000
         Top             =   120
         Width           =   9015
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         Picture         =   "frmREL.frx":0006
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmREL.frx":0108
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ProgressBar prgImpressao 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmREL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCampos         As String
Dim arrPaginas        As Variant
Dim lngPaginas        As Long
Dim intLinha          As Integer
Dim I                 As Integer
Dim J                 As Integer
Dim lngContPagi       As Long
Dim Impressora        As Printer
Public intOrientation As Integer
Public lngQuebLin     As Long
Public intFontSize    As Integer

Private Sub cmdImpressao_Click()
    Imprime
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_Load()
 
  txtRel.Text = ""
  ReDim arrPaginas(1 To 7000, 1 To 56) As String
    
  Open "C:\RICARDO\SGI\RELATORIOS\MOSTRAREL\Rel.txt" For Input As #1
  
  strCampos = ""
  lngPaginas = 1
  intLinha = 1
  Do While Not EOF(1)
     
     Input #1, strCampos
     
     arrPaginas(lngPaginas, intLinha) = strCampos
     intLinha = intLinha + 1
     If intLinha >= lngQuebLin Then
        lngPaginas = lngPaginas + 1
        intLinha = 1
     Else
        If strCampos = "R" Then
           lngPaginas = lngPaginas + 1
           intLinha = 1
        End If
     End If
     
  Loop
    
  Close #1
  
  ' --------------------------------------
  
  prgImpressao.Visible = True
  prgImpressao.Min = 0
  prgImpressao.Max = lngPaginas
  
  For I = 1 To UBound(arrPaginas)
      
      If Len(Trim(arrPaginas(I, 1))) = 0 Then Exit For
      
      For J = 1 To 56
          If Len(Trim(arrPaginas(I, J))) > 0 Then txtRel.Text = txtRel.Text & IIf((Trim(arrPaginas(I, J)) = "B") Or (Trim(arrPaginas(I, J)) = "R"), "", Replace(arrPaginas(I, J), ",", ".")) & vbCrLf
      Next J
      
      prgImpressao.Value = I
      
  Next I
  
  prgImpressao.Visible = False
  
  ' --------------------------------------
  
End Sub

Private Sub Imprime()
     
     If Printers.Count = 0 Then
        MsgBox "Não há impressora instalada !!!", vbOKOnly + vbCritical, "Aviso"
     End If
     
     Dim intLinha As Integer
     
     prgImpressao.Visible = True
     prgImpressao.Min = 0
     prgImpressao.Max = lngPaginas
     
     Printer.Font = "Courier New"
     Printer.FontSize = intFontSize
     Printer.Orientation = intOrientation
     
     For I = 1 To UBound(arrPaginas)
      
         If Len(Trim(arrPaginas(I, 1))) = 0 Then Exit For
      
         intLinha = 1
         For J = 1 To 56
             
             If Len(Trim(arrPaginas(I, J))) > 0 Then
                
                Printer.Print IIf(Trim(arrPaginas(I, J)) = "B", "", arrPaginas(I, J))
                intLinha = intLinha + 1
                
             End If
             
             If intLinha >= lngQuebLin Then
                Printer.NewPage
                intLinha = 0
             Else
                If Trim(arrPaginas(I, J)) = "R" Then
                   Printer.NewPage
                   intLinha = 0
                End If
             End If
             
         Next J
      
         prgImpressao.Value = I
      
     Next I
  
     prgImpressao.Visible = False
     
     Printer.EndDoc
  
     ' --------------------------------------
     
End Sub
