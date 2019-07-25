VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADMOTLIQOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial    As Integer
Dim intCODIGO    As Integer
Dim strDESCRI    As String

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CODIGO() As Integer
    CODIGO = intCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Integer)
    intCODIGO = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property

Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUSUARIO As Long)
    frmCADMOTLIQOPP.cCaminho = cCaminho
    frmCADMOTLIQOPP.Linha = Linha
    frmCADMOTLIQOPP.FILIAL = iFilial
    frmCADMOTLIQOPP.strAcesso = strAcesso
    frmCADMOTLIQOPP.lngIDUSUARIO = lngIDUSUARIO
    frmCADMOTLIQOPP.strUsuario = strUsuario
    frmCADMOTLIQOPP.Show vbModal
End Sub

Public Function AtivoDesativo() As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMOTLIQOP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
    
        '' Grava Eventos
        sSql = "Insert Into SGI_CADMOTLIQOP (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                     ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                     ,SGI_DESCRI" & vbCrLf
        
        
        sSql = sSql & "                       ) Values (" & vbCrLf
        sSql = sSql & "                                 " & intFilial & vbCrLf
        sSql = sSql & "                                ," & intCODIGO & vbCrLf
        sSql = sSql & "                                ," & strDESCRI & vbCrLf
        sSql = sSql & "                                )"
     
     ElseIf strOPER = "A" Then
     
        '' Grava Eventos
        sSql = "Update SGI_CADMOTLIQOP Set " & vbCrLf
        sSql = sSql & "                     SGI_DESCRI     = " & strDESCRI & vbCrLf
        sSql = sSql & "              Where " & vbCrLf
        sSql = sSql & "                    SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "                And SGI_CODIGO = " & intCODIGO
     
     ElseIf strOPER = "E" Then
     
        sSql = "Delete from SGI_CADMOTLIQOP " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
     
     End If
     
     BGRV.CommandText = sSql
     BGRV.Execute
     
     adoBanco_Dados.CommitTrans
     
     GRAVA = True
     
     Exit Function
     
err_Trans:

     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
     Set objErro = Nothing

End Function

Public Function Carrega_Campos() As Boolean

    Carrega_Campos = False

    Dim I          As Integer
    Dim intQTDREG  As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOTLIQOP " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
        intCODIGO = BREC!SGI_CODIGO
        strDESCRI = Trim(BREC!SGI_DESCRI)
        Carrega_Campos = True
    End If
    BREC.Close
    ' ----------------------------------------------------------
    
End Function

