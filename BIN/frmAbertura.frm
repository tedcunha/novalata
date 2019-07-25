VERSION 5.00
Begin VB.Form frmAbertura 
   BorderStyle     =   0  'None
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   4470
         Left            =   120
         Picture         =   "frmAbertura.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   9450
      End
      Begin VB.Label LblVersao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "LblVersao"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   4800
         Width           =   9495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "frmAbertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    LblVersao.Caption = "Versão : " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Timer1_Timer()
    Unload Me
    frmAcesso.Show
End Sub
