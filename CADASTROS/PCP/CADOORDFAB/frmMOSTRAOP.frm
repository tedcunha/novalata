VERSION 5.00
Begin VB.Form frmMOSTRAOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mostra OP'S"
   ClientHeight    =   10605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   12165
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSAIR 
      Caption         =   "&Sair"
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
      Left            =   10320
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Assinatura 
      Height          =   9975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   12045
   End
End
Attribute VB_Name = "frmMOSTRAOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngCODOP As Long

Private Sub cmdSAIR_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Assinatura.Picture = LoadPicture("C:\ricardo\SGI\NOVALATA-ANTIGO\IMAGENS\Quimicolla Homologada 18L 150 dpi2.JPG")
End Sub
