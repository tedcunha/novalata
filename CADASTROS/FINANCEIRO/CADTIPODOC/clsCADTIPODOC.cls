VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADTIPODOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial      As Long
Dim lngTIPDESPCOD  As Long
Dim strTIPDESPDES  As String

Public Property Get TIPDESPCOD() As Long
    TIPDESPCOD = lngTIPDESPCOD
End Property

Public Property Let TIPDESPCOD(ByVal vData As Long)
    lngTIPDESPCOD = vData
End Property

Public Property Get TIPDESPDES() As String
    TIPDESPDES = strTIPDESPDES
End Property

Public Property Let TIPDESPDES(ByVal vData As String)
    strTIPDESPDES = vData
End Property

Public Property Get Filial() As Long
    Filial = intFilial
End Property

Public Property Let Filial(ByVal vData As Long)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADTIPODOCP.cCaminho = cCaminho
    frmCADTIPODOCP.Linha = Linha
    frmCADTIPODOCP.Filial = iFilial
    frmCADTIPODOCP.strAcesso = strAcesso
    frmCADTIPODOCP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADTIPODOC.cCaminho = cCaminho
    frmCADTIPODOC.Linha = Linha
    frmCADTIPODOC.Filial = iFilial
    frmCADTIPODOC.strAcesso = strAcesso
    frmCADTIPODOC.strUSUARIO = strUSUARIO
    frmCADTIPODOC.cTipOper = "I"
    frmCADTIPODOC.Show vbModal
End Sub

Public Function Pesq_CADTIPODOC() As Boolean

   Pesq_CADTIPODOC = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADTIPODOC " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CADTIPODOC = True
   BREC.Close

End Function

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADTIPODOC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngTIPDESPCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
     
      lngTIPDESPCOD = BREC!SGI_CODIGO
      strTIPDESPDES = BREC!SGI_DESCRICAO
      
      Carrega_campos = True
   End If
   
   BREC.Close

End Function


Public Function Gera_Codigo(sModulo As String) As Integer

    Gera_Codigo = 1
    
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql + "       (Max(SGI_NUMERO) + 1) As SGI_NUMERO " & vbCrLf
    sSql = sSql + "  From " & vbCrLf
    sSql = sSql + "       SGI_NUMERO " & vbCrLf
    sSql = sSql + " Where " & vbCrLf
    sSql = sSql + "       SGI_MODULO = '" & sModulo & "'"
    sSql = sSql + "   And SGI_FILIAL = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados
    
    If Not BREC.EOF Then
    
       If IsNull(BREC!SGI_NUMERO) = True Then
          
          Gera_Codigo = 1
          
          sSql = "Insert into SGI_NUMERO (SGI_FILIAL,SGI_NUMERO,SGI_MODULO) Values(" & vbCrLf
          sSql = sSql + "                                              " & intFilial & vbCrLf
          sSql = sSql + "                                            ,1" & vbCrLf
          sSql = sSql + "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql + "                                          )" & vbCrLf
          
       ElseIf BREC!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql + "           SGI_NUMERO = " & BREC!SGI_NUMERO & vbCrLf
          sSql = sSql + "         Where " & vbCrLf
          sSql = sSql + "               SGI_MODULO = '" & sModulo & "'" & vbCrLf
          sSql = sSql + "           And SGI_FILIAL =  " & intFilial
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       
    End If
    
    BREC.Close
    
End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
        
        '' Grava especificação técnica do produto
        sSql = "Insert Into SGI_CADTIPODOC (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & lngTIPDESPCOD
        sSql = sSql & "                           ,'" & strTIPDESPDES & "'" & vbCrLf
        sSql = sSql & "                           )"
     
     ElseIf strOPER = "A" Then
       
        '' altera especificação técnica do produto
        sSql = "Update SGI_CADTIPODOC Set " & vbCrLf
        sSql = sSql & "                SGI_DESCRICAO = '" & strTIPDESPDES & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & lngTIPDESPCOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
     ElseIf strOPER = "E" Then
     
        '' Exclui especificação técnica do produto
        sSql = "Delete from SGI_CADTIPODOC " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngTIPDESPCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
     End If
     
     BGRV.CommandText = sSql
     BGRV.Execute
     
     adoBanco_Dados.CommitTrans
     
     GRAVA = True
     
     Exit Function
     
err_grava:
     
     adoBanco_Dados.RollbackTrans

     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
     Set objErro = Nothing

End Function


