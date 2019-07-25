VERSION 5.00
Object = "{69ECBBD3-5C2A-4A84-ABEC-23937DBF1B54}#1.4#0"; "FlowChartPro.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVIEWARVPROD 
   Caption         =   "Ver Arvore"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVIEWARVPROD.frx":0000
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVIEWARVPROD.frx":0352
            Key             =   "Minus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVIEWARVPROD.frx":06A4
            Key             =   "Fit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVIEWARVPROD.frx":09F6
            Key             =   "SizeDoc"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5970
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   635
      ButtonWidth     =   609
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Expande a Arvore"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Diminui a Arvore"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ajusta a Arvore"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ajusta o Documento Inteiro de modo que possa ver todo"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
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
         Picture         =   "frmVIEWARVPROD.frx":0D48
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
   Begin FLOWCHARTLibCtl.Overview OverviewControl 
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "frmVIEWARVPROD.frx":0E4A
      TabIndex        =   2
      Top             =   960
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
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
