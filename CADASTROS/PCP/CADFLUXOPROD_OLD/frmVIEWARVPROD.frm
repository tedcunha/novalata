VERSION 5.00
Object = "{69ECBBD3-5C2A-4A84-ABEC-23937DBF1B54}#1.4#0"; "FlowChartPro.dll"
Begin VB.Form frmVIEWARVPROD 
   Caption         =   "Ver Arvore"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
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
         Picture         =   "frmVIEWARVPROD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
   End
   Begin FLOWCHARTLibCtl.Overview OverviewControl 
      Height          =   5055
      Left            =   0
      OleObjectBlob   =   "frmVIEWARVPROD.frx":0102
      TabIndex        =   1
      Top             =   840
      Width           =   7335
   End
End
Attribute VB_Name = "frmVIEWARVPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Parent As frmCADARVPROD

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Public Sub SetDocument(Document As FlowChart)

    OverviewControl.Document = Document
    OverviewControl.FitAll = True

End Sub

Public Sub SetParent(Form As frmCADARVPROD)

    Set Parent = Form

End Sub

Private Sub Form_Resize()
    ''If Me.Width > 140 Then OverviewControl.Width = Me.Width - 140
    ''If Me.Height > 800 Then OverviewControl.Height = Me.Height - 800
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Parent.fc.ZoomIn
    End If

    If Button.Index = 2 Then
        Parent.fc.ZoomOut
    End If

    If Button.Index = 3 Then
        Parent.fc.ZoomToFit
        Parent.fc.ScrollTo 0, 0
    End If

    If Button.Index = 4 Then
        Parent.fc.FitDocToObjects 5
    End If
End Sub
