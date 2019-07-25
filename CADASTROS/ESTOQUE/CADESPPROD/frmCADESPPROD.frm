VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADESPPROD 
   Caption         =   "Cadastro de espécie de produtos"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "[Verniz / Esmalte ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   5895
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   480
         Width           =   3855
         Begin VB.OptionButton optVern02 
            Caption         =   "Sim"
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
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optVern02 
            Caption         =   "Não"
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
            Index           =   0
            Left            =   960
            TabIndex        =   22
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton optVern01 
            Caption         =   "Sim"
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
            Index           =   1
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optVern01 
            Caption         =   "Não"
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
            Index           =   0
            Left            =   960
            TabIndex        =   19
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Verniz Interno 02"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Verniz Interno 01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Folha de Flandres ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   5895
      Begin TabDlg.SSTab SSTab1 
         Height          =   2655
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4683
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Corpo"
         TabPicture(0)   =   "frmCADESPPROD.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstCorpo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tampa"
         TabPicture(1)   =   "frmCADESPPROD.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstTampa"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Fundo"
         TabPicture(2)   =   "frmCADESPPROD.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstFundo"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Argola"
         TabPicture(3)   =   "frmCADESPPROD.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lstArgola"
         Tab(3).ControlCount=   1
         Begin VB.ListBox lstArgola 
            Height          =   1860
            Left            =   -75000
            Style           =   1  'Checkbox
            TabIndex        =   14
            Top             =   360
            Width           =   5415
         End
         Begin VB.ListBox lstFundo 
            Height          =   1860
            Left            =   -75000
            Style           =   1  'Checkbox
            TabIndex        =   13
            Top             =   360
            Width           =   5415
         End
         Begin VB.ListBox lstTampa 
            Height          =   1860
            Left            =   -75000
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   360
            Width           =   5415
         End
         Begin VB.ListBox lstCorpo 
            Height          =   2085
            ItemData        =   "frmCADESPPROD.frx":0070
            Left            =   0
            List            =   "frmCADESPPROD.frx":0077
            Style           =   1  'Checkbox
            TabIndex        =   11
            Top             =   360
            Width           =   5415
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5895
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
         Picture         =   "frmCADESPPROD.frx":0085
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
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
         Left            =   1800
         Picture         =   "frmCADESPPROD.frx":0187
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAltera 
         Caption         =   "&Altera"
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
         Left            =   960
         Picture         =   "frmCADESPPROD.frx":0289
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         TabIndex        =   3
         Top             =   600
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmCADESPPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho   As String
Public Linha      As Variant
Public cTipOper   As String
Public iCodigo    As Integer
Public FILIAL     As Integer
Public strAcesso  As String
Public strUSUARIO As String
Dim objBLBFunc    As Object
Dim objCADESPPROD As Object
Dim arrCORPO      As Variant
Dim arrTAMPA      As Variant
Dim arrFUNDO      As Variant
Dim arrARGOLA     As Variant
Dim I             As Integer

Private Sub cmdAltera_Click()
    
    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de espécie de produtos - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim intITENSEL As Integer
    
    If ValidaCampos = True Then
       
       If cTipOper = "I" Then
          objCADESPPROD.ESPPRODCOD = objCADESPPROD.Gera_Codigo(Me.Name)
       End If
       
       objCADESPPROD.ESPPRODESC = txtDescricao.Text
       
       If optVern01(0).Value = True Then objCADESPPROD.Vern01 = 0
       If optVern01(1).Value = True Then objCADESPPROD.Vern01 = 1
        
       If optVern02(0).Value = True Then objCADESPPROD.Vern02 = 0
       If optVern02(1).Value = True Then objCADESPPROD.Vern02 = 1
        
        '' Corpo
        With lstCorpo
             intITENSEL = 0
             For I = 0 To (.ListCount - 1)
                    If .Selected(I) = True Then intITENSEL = intITENSEL + 1
             Next I
             
             arrCORPO = Empty
             If intITENSEL > 0 Then
                ReDim arrCORPO(1 To intITENSEL) As String
                intITENSEL = 0
                For I = 0 To (.ListCount - 1)
                       If .Selected(I) = True Then
                           intITENSEL = intITENSEL + 1
                           arrCORPO(intITENSEL) = .ItemData(I)
                       End If
                Next I
             End If
             objCADESPPROD.CORPO = arrCORPO
        End With
       
        '' Tampa
        With lstTampa
             intITENSEL = 0
             For I = 0 To (.ListCount - 1)
                    If .Selected(I) = True Then intITENSEL = intITENSEL + 1
             Next I
             
             arrTAMPA = Empty
             If intITENSEL > 0 Then
                ReDim arrTAMPA(1 To intITENSEL) As String
                intITENSEL = 0
                For I = 0 To (.ListCount - 1)
                       If .Selected(I) = True Then
                          intITENSEL = intITENSEL + 1
                          arrTAMPA(intITENSEL) = .ItemData(I)
                       End If
                Next I
             End If
             objCADESPPROD.TAMPA = arrTAMPA
        End With
       
        '' Fundo
        With lstFundo
             intITENSEL = 0
             For I = 0 To (.ListCount - 1)
                    If .Selected(I) = True Then intITENSEL = intITENSEL + 1
             Next I
             
             arrFUNDO = Empty
             If intITENSEL > 0 Then
                ReDim arrFUNDO(1 To intITENSEL) As String
                intITENSEL = 0
                For I = 0 To (.ListCount - 1)
                       If .Selected(I) = True Then
                          intITENSEL = intITENSEL + 1
                          arrFUNDO(intITENSEL) = .ItemData(I)
                       End If
                Next I
             End If
             objCADESPPROD.FUNDO = arrFUNDO
        End With
       
        '' Argola
        With lstArgola
             intITENSEL = 0
             For I = 0 To (.ListCount - 1)
                    If .Selected(I) = True Then intITENSEL = intITENSEL + 1
             Next I
             
             arrARGOLA = Empty
             If intITENSEL > 0 Then
                ReDim arrARGOLA(1 To intITENSEL) As String
                intITENSEL = 0
                For I = 0 To (.ListCount - 1)
                       If .Selected(I) = True Then
                          intITENSEL = intITENSEL + 1
                          arrARGOLA(intITENSEL) = .ItemData(I)
                       End If
                Next I
             End If
             objCADESPPROD.ARGOLA = arrARGOLA
        End With
       
       If objCADESPPROD.GRAVA(cTipOper) = True Then
          
          MsgBox "A espécie de produto foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objCADESPPROD = Nothing
             Unload Me
          End If
          
       End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADESPPROD = Nothing
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADESPPROD = CreateObject("CADESPPROD.clsCADESPPROD")
   
   objCADESPPROD.FILIAL = FILIAL
   
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de espécie de produtos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADESPPROD
    
    txtCodigo.Text = ""
    Call LimpaListBox
    Call PopList(lstCorpo)
    Call PopList(lstTampa)
    Call PopList(lstFundo)
    Call PopList(lstArgola)
   
    optVern01(1).Value = True
    optVern02(1).Value = True
   
End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    optVern01(1).Value = True
    optVern02(1).Value = True
    
    
    Me.Caption = "Cadastro de espécie de produtos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADESPPROD
    Call LimpaListBox
    Call PopList(lstCorpo)
    Call PopList(lstTampa)
    Call PopList(lstFundo)
    Call PopList(lstArgola)
    
    objCADESPPROD.ESPPRODCOD = iCodigo
    
    If objCADESPPROD.Carrega_campos = True Then
    
        txtCodigo.Text = Str(objCADESPPROD.ESPPRODCOD)
        txtDescricao.Text = objCADESPPROD.ESPPRODESC
       
        optVern01(objCADESPPROD.Vern01).Value = True
        optVern02(objCADESPPROD.Vern02).Value = True
       
        Call Seleciona_Corpo
        Call Seleciona_Tampa
        Call Seleciona_Fundo
        Call Seleciona_Argola
       
    End If
    
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    optVern01(1).Value = True
    optVern02(1).Value = True
    
    Me.Caption = "Cadastro de espécie de produtos - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADESPPROD
    
    objCADESPPROD.ESPPRODCOD = iCodigo
    Call LimpaListBox
    Call PopList(lstCorpo)
    Call PopList(lstTampa)
    Call PopList(lstFundo)
    Call PopList(lstArgola)
    
    If objCADESPPROD.Carrega_campos = True Then
    
        txtCodigo.Text = Str(objCADESPPROD.ESPPRODCOD)
        txtDescricao.Text = objCADESPPROD.ESPPRODESC
        
        optVern01(objCADESPPROD.Vern01).Value = True
        optVern02(objCADESPPROD.Vern02).Value = True
        
        Call Seleciona_Corpo
        Call Seleciona_Tampa
        Call Seleciona_Fundo
        Call Seleciona_Argola
    
    End If

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADESPPROD
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Espécie do tipo de produto inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select * from SGI_CADESPPROD Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
        sSql = sSql & " And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Espécie do tipo de produto já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADESPPROD.ESPPRODESC <> txtDescricao.Text Then
        
           sSql = "Select * from SGI_CADESPPROD Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
           sSql = sSql & " And SGI_FILIAL = " & FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Espécie do tipo de produto existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADESPPROD.ESPPRODESC
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function


Private Sub LimpaListBox()
    lstCorpo.Clear
    lstTampa.Clear
    lstFundo.Clear
    lstArgola.Clear
End Sub

Private Sub PopList(lstGenerico As ListBox)

    Dim arrDESC() As String
    ReDim arrDESC(1 To 4) As String

    arrDESC(1) = "VEX - Verniz Externo"
    arrDESC(2) = "VZ - Verniz dos 2 Lados"
    arrDESC(3) = "NAT - Natural"
    arrDESC(4) = "VI - Verniz interno Apenas"

    For I = 1 To UBound(arrDESC)
        lstGenerico.AddItem arrDESC(I)
        lstGenerico.ItemData(lstGenerico.NewIndex) = I
        lstGenerico.Selected(I - 1) = False
    Next I
    
End Sub

Private Sub Seleciona_Corpo()

        If Not IsArray(objCADESPPROD.CORPO) Then Exit Sub
        If lstCorpo.ListCount = 0 Then Exit Sub
       
        Dim J As Integer
       
        arrCORPO = objCADESPPROD.CORPO
        
        With lstCorpo
            
            For I = 0 To (.ListCount - 1)
                 For J = 1 To UBound(arrCORPO)
                    If .ItemData(I) = CInt(arrCORPO(J)) Then .Selected(I) = True
                 Next J
            Next I
        
        End With

End Sub

Private Sub Seleciona_Tampa()

        If Not IsArray(objCADESPPROD.TAMPA) Then Exit Sub
        If lstTampa.ListCount = 0 Then Exit Sub
       
        Dim J As Integer
        arrTAMPA = objCADESPPROD.TAMPA
       
        With lstTampa
            
            For I = 0 To (.ListCount - 1)
                 For J = 1 To UBound(arrTAMPA)
                    If .ItemData(I) = CInt(arrTAMPA(J)) Then .Selected(I) = True
                 Next J
            Next I
        
        End With

End Sub

Private Sub Seleciona_Fundo()

        If Not IsArray(objCADESPPROD.FUNDO) Then Exit Sub
        If lstFundo.ListCount = 0 Then Exit Sub
       
        Dim J As Integer
        arrFUNDO = objCADESPPROD.FUNDO
       
        With lstFundo
            
            For I = 0 To (.ListCount - 1)
                 For J = 1 To UBound(arrFUNDO)
                    If .ItemData(I) = CInt(arrFUNDO(J)) Then .Selected(I) = True
                 Next J
            Next I
        
        End With

End Sub

Private Sub Seleciona_Argola()

        If Not IsArray(objCADESPPROD.ARGOLA) Then Exit Sub
        If lstArgola.ListCount = 0 Then Exit Sub
       
        Dim J As Integer
        arrARGOLA = objCADESPPROD.ARGOLA
       
        With lstArgola
            
            For I = 0 To (.ListCount - 1)
                 For J = 1 To UBound(arrARGOLA)
                    If .ItemData(I) = CInt(arrARGOLA(J)) Then .Selected(I) = True
                 Next J
            Next I
        
        End With

End Sub

