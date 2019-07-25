VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiSIGEGerenciador 
   BackColor       =   &H8000000C&
   Caption         =   "Gerenciador SIGE"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin MSComctlLib.ImageList ilsList 
      Left            =   720
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiSIGEGerenciador.frx":0000
            Key             =   "Key01"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TbrMenu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   7200
   End
   Begin VB.Timer tmrHora 
      Interval        =   1000
      Left            =   120
      Top             =   7800
   End
   Begin MSComctlLib.StatusBar stbMensagem 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8460
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16087
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "mdiSIGEGerenciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    
    strNOMCOMP = "'" & Trim(UCase(GetIPHostName())) & "'"
    
    stbMensagem.Panels(1).Text = ""
    stbMensagem.Panels(2).Text = strNOMCOMP
    stbMensagem.Panels(3).Text = GetIPAddress()
      
    strSTRCONNECT = CarregaStrConect
    
    If Len(Trim(strSTRCONNECT)) > 0 Then
        
        Call AbBanco(strSTRCONNECT)
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "        SGI_CODIGO" & vbCrLf
        sSql = sSql & "       ,SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADVENDEDOR" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_NOMCOMP = " & strNOMCOMP & vbCrLf
        
        BREC.Open sSql, BD, adOpenDynamic
        If Not BREC.EOF() Then
            lngCODVENDEDOR = BREC!SGI_CODIGO
            strNOMVENDEDOR = BREC!SGI_DESCRICAO
            stbMensagem.Panels(1).Text = "Vendedor : " & strNOMVENDEDOR
        End If
        BREC.Close
        
        Call FcBanco
        
    End If
    
    Call CriaToolBar(ilsList, TbrMenu)
    
    intFILIAL = 1
    
End Sub

Private Sub TbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    frmSTATUSPEDIDOS.Show
End Sub

Private Sub tmrHora_Timer()
  stbMensagem.Panels(4).Text = DateSerial(Year(Date), Month(Date), Day(Date))
  stbMensagem.Panels(5).Text = Format(Time, "HH:MM:SS")
End Sub


Public Sub CriaToolBar(ilsList As ImageList, TbrMenu As Toolbar)
   Dim btn As MSComctlLib.Button
   
   Set TbrMenu.ImageList = ilsList
   Set btn = TbrMenu.Buttons.Add(, "Dados de Pedidos", , , "Key01")
   btn.ToolTipText = "Status dos Pedidos"
   
   ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
   ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
   
   ''Set btn = TbrMenu.Buttons.Add(, "Incluir", , , "Key01")
   ''btn.ToolTipText = "Incluir Dados"
   
   ''Set btn = TbrMenu.Buttons.Add(, "Alterar", , , "Key01")
   ''btn.ToolTipText = "Alterar dados"

   ''Set btn = TbrMenu.Buttons.Add(, "Excluir", , , "Key01")
   ''btn.ToolTipText = "Excluir Dados"

   ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
   ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
   
   ''Set btn = TbrMenu.Buttons.Add(, "Pesquizar", , , "Key01")
   ''btn.ToolTipText = "Pesquisa Dados"

End Sub

