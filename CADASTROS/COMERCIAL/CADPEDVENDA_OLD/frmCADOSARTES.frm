VERSION 5.00
Begin VB.Form frmCADOSARTES 
   Caption         =   "OS - Setor de Artes"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11280
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
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
         Picture         =   "frmCADOSARTES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADOSARTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strIDPRODUTO         As String
Public strNOMFILIAL         As String
Public FILIAL               As Integer
Public mskDTPED             As String
Public cTipOper             As String
Public intALTFILME          As Integer
Public intFOTNOVO           As Integer
Public strRETORNO           As String
Public intAction2Do         As Integer
Public intStatusOP          As Integer
Public strPRODCODLIN        As String
Public lngSALDOQTDENTR      As Long
Public lngCODPED            As Long
Public strGRPCOD            As String

Dim objBLBFunc3             As Object
Dim objCADPEDVENDAOS        As Object

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

