VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmRELSERV 
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   10605
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer cryRELVIEW 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      lastProp        =   600
      _cx             =   18653
      _cy             =   13573
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmRELSERV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intFilial        As Integer
Public strQUERYSQL      As String
Public strRELNOME       As String
Public Linha            As Variant
Public strCABEC1        As String
Public strCABEC2        As String
Public intORIENTATION   As Integer
Public boolArvoreSN     As Boolean
Public strConn          As String
Public boolView         As Boolean

Dim objFuncoes          As Object
Dim cryRELApplication   As New CRAXDRT.Application
Dim cryREL              As CRAXDRT.Report



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Set objFuncoes = Nothing
       Unload Me
    End If
End Sub

Private Sub Form_Load()

On Error GoTo Err_Desc

    Dim I As Integer
    Dim j As Integer
    Dim teste As String
    
    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set cryREL = cryRELApplication.OpenReport(strRELNOME, 1)
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
    
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Unload Me
       Exit Sub
    End If
    
    If BREC.State = 1 Then BREC.Close
    
    BREC.Open strQUERYSQL, adoBanco_Dados
    If BREC.EOF Then
       BREC.Close
       Exit Sub
       Unload Me
    End If
    
    Me.Caption = strCABEC2
    
    If Len(Trim(strCABEC1)) > 0 Then cryREL.ReportTitle = strCABEC1 & vbCrLf & strCABEC2
    
    'For I = 1 To (cryREL.Sections.Count)
    '    teste = cryREL.Sections.Item(I).Name
    '    For J = 1 To cryREL.Sections.Item(I).ReportObjects.Count
    '        ''teste = cryREL.Sections.Item(I).ReportObjects.Item(J).Name
    '        ''teste = cryREL.Sections.Item(I).AddPictureObject("C:\RICARDO\SGI\ARQUIVOS\ASSINATURA.jpeg", 0, 0)
    '        ''Call cryREL.Sections.Item(I).ReportObjects.Item(J).AddPictureObject("C:\RICARDO\SGI\ARQUIVOS\ASSINATURA.jpeg", 0, 0)
    '    Next J
    'Next I
    
    
    '' Passando o SQL para o relatório
    
    cryREL.Database.SetDataSource BREC
    
    cryRELVIEW.EnableExportButton = True
    cryRELVIEW.EnableRefreshButton = False
    cryRELVIEW.EnableCloseButton = False
    cryRELVIEW.EnableAnimationCtrl = False
    cryRELVIEW.EnableGroupTree = boolArvoreSN
    
    cryRELVIEW.ReportSource = cryREL
    cryRELVIEW.ViewReport
    
    Exit Sub
    
Err_Desc:
    
    If BREC.State = 1 Then BREC.Close
    
    If Err.Number = -2147206461 Then
       MsgBox "Erro Nº         : " & Err.Number & vbCrLf & _
              "Erro Desrcrição : Arquivo de relatório não encontrado !!!", vbOKOnly + vbCritical, "Erro"
    Else
       MsgBox "Erro Nº        : " & Err.Number & vbCrLf & _
              "Erro Desrcição : " & Err.Description, vbOKOnly + vbCritical, "Erro"
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Resize()
  
  cryRELVIEW.Top = 0
  cryRELVIEW.Left = 0
  cryRELVIEW.Height = ScaleHeight
  cryRELVIEW.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objFuncoes = Nothing
    Call Destroy_Objeto
End Sub

Private Sub Destroy_Objeto()
       Set objFuncoes = Nothing
End Sub
