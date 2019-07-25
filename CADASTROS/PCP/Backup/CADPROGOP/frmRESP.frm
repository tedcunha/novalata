VERSION 5.00
Begin VB.Form frmRESP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[ ATENÇÃO ]"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   FillColor       =   &H0080FFFF&
   ForeColor       =   &H0080FFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      MaskColor       =   &H00800000&
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00800000&
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame fraRESP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   6855
      Begin VB.OptionButton optRESP 
         Caption         =   "Estoura a Linha e Remaneja OP's Manualmente ?"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.OptionButton optRESP 
         Caption         =   "Fraciona ?"
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton optRESP 
         Caption         =   "Estoura a Linha sem fracionar ?"
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
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.OptionButton optRESP 
         Caption         =   "Não faz nada ?"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Estourou a Capacidade da Linha,  O que desejá fazer com a OP ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmRESP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intRETORNO           As Integer
Public lngSALDODISP         As Long

Private Sub cmdCancel_Click()
    intRETORNO = vbNo
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If optRESP(0).Value = True Then intRETORNO = vbYes
    If optRESP(1).Value = True Then intRETORNO = 1
    If optRESP(2).Value = True Then intRETORNO = vbNo
    If optRESP(3).Value = True Then intRETORNO = 2
    Unload Me
End Sub

Private Sub Form_Load()
    optRESP(2).Value = True
    
    If lngSALDODISP < 0 Then Label3.Caption = "Estourou a Capacidade da Linha,  O que desejá fazer com a OP ?"
    If lngSALDODISP > 0 Then Label3.Caption = "Não Estourou a Capacidade da Linha,  O que dese fazer com as OP's ?"
    
End Sub

