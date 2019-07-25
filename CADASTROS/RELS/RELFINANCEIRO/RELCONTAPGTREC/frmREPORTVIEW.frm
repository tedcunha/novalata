VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Begin VB.Form frmREPORTVIEW 
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer RELVIEW 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _cx             =   12938
      _cy             =   10186
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1046
   End
End
Attribute VB_Name = "frmREPORTVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strRELNOME As String
Public strCabec1  As String
Public strCabec2  As String
Public strSQL     As String


Private Sub Form_Load()
    
    If strRELNOME = "rptRELAPGDTV" Then
         
        Dim rptRELAPGDTV  As New RELAPGDTV
        
        If BREC.State = 1 Then BREC.Close
       
        BREC.Open strSQL, adoBanco_Dados, adOpenDynamic
        
        If BREC.EOF Then
           BREC.Close
           Unload Me
           Exit Sub
        End If
        
        rptRELAPGDTV.Database.SetDataSource BREC
         
        rptRELAPGDTV.txtCABEC1.SetText strCabec1
        rptRELAPGDTV.txtCABEC2.SetText strCabec2
       
        RELVIEW.EnableRefreshButton = False
         
        RELVIEW.ReportSource = rptRELAPGDTV
        RELVIEW.ViewReport
         
    End If
    If strRELNOME = "rptRELAPGFORN" Then
    
       Dim rptRELAPGFORN As New RELAPGFORN
         
       rptRELAPGFORN.txtCABEC1.SetText strCabec1
       rptRELAPGFORN.txtCABEC2.SetText strCabec2
       
       If BREC.State = 1 Then BREC.Close
       
       BREC.Open strSQL, adoBanco_Dados, adOpenDynamic
        
       If BREC.EOF Then
          BREC.Close
          Unload Me
          Exit Sub
       End If
         
       rptRELAPGFORN.Database.SetDataSource BREC
         
       RELVIEW.EnableRefreshButton = False
         
       RELVIEW.ReportSource = rptRELAPGFORN
       RELVIEW.ViewReport
    
    End If
    If strRELNOME = "rptRELAPGDSP" Then
    
       Dim rptRELAPGDSP As New RELAPGDSP
    
       If BREC.State = 1 Then BREC.Close
       
       rptRELAPGDSP.txtCABEC1.SetText strCabec1
       rptRELAPGDSP.txtCABEC2.SetText strCabec2
       
       BREC.Open strSQL, adoBanco_Dados, adOpenDynamic
        
       If BREC.EOF Then
          BREC.Close
          Unload Me
          Exit Sub
       End If
         
       rptRELAPGDSP.Database.SetDataSource BREC
         
       RELVIEW.EnableRefreshButton = False
         
       RELVIEW.ReportSource = rptRELAPGDSP
       RELVIEW.ViewReport
    
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Resize()
 
    RELVIEW.Top = 0
    RELVIEW.Left = 0
    RELVIEW.Height = ScaleHeight
    RELVIEW.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If BREC.State = 1 Then BREC.Close
    
End Sub
