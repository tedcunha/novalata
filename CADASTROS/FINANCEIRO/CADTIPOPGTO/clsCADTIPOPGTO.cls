VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADTIPOPGTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial      As Long
Dim lngTIPPGTOCOD  As Long
Dim strTIPPGTODES  As String
Dim intOperacao    As Integer

Public Property Get OPERACAO() As Integer
    OPERACAO = intOperacao
End Property

Public Property Let OPERACAO(ByVal vData As Integer)
    intOperacao = vData
End Property

Public Property Get TIPPGTOCOD() As Long
    TIPPGTOCOD = lngTIPPGTOCOD
End Property

Public Property Let TIPPGTOCOD(ByVal vData As Long)
    lngTIPPGTOCOD = vData
End Property

Public Property Get TIPPGTODES() As String
    TIPPGTODES = strTIPPGTODES
End Property

Public Property Let TIPPGTODES(ByVal vData As String)
    strTIPPGTODES = vData
End Property

Public Property Get Filial() As Long
    Filial = intFilial
End Property

Public Property Let Filial(ByVal vData As Long)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADTIPOPGTOP.cCaminho = cCaminho
    frmCADTIPOPGTOP.Linha = Linha
    frmCADTIPOPGTOP.Filial = iFilial
    frmCADTIPOPGTOP.strAcesso = strAcesso
    frmCADTIPOPGTOP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADTIPOPGTO.cCaminho = cCaminho
    frmCADTIPOPGTO.Linha = Linha
    frmCADTIPOPGTO.Filial = iFilial
    frmCADTIPOPGTO.strAcesso = strAcesso
    frmCADTIPOPGTO.strUSUARIO = strUSUARIO
    frmCADTIPOPGTO.cTipOper = "I"
    frmCADTIPOPGTO.Show vbModal
End Sub

Public Function Pesq_CADTIPOPGTO() As Boolean

   Pesq_CADTIPOPGTO = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADTIPOPGTO " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CADTIPOPGTO = True
   BREC.Close

End Function

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADTIPOPGTO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngTIPPGTOCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
     
      lngTIPPGTOCOD = BREC!SGI_CODIGO
      strTIPPGTODES = BREC!SGI_DESCRICAO
      intOperacao = BREC!SGI_OPERACAO
      
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
        sSql = "Insert Into SGI_CADTIPOPGTO (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO,SGI_OPERACAO,SGI_SINAL) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & lngTIPPGTOCOD
        sSql = sSql & "                           ,'" & strTIPPGTODES & "'" & vbCrLf
        sSql = sSql & "                           ," & intOperacao & vbCrLf
        If intOperacao = 1 Then sSql = sSql & "                           ,'-'" & vbCrLf
        If intOperacao = 2 Then sSql = sSql & "                           ,'+'" & vbCrLf
        sSql = sSql & "                           )"
     
     ElseIf strOPER = "A" Then
       
        '' altera especificação técnica do produto
        sSql = "Update SGI_CADTIPOPGTO Set " & vbCrLf
        sSql = sSql & "                SGI_DESCRICAO = '" & strTIPPGTODES & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & lngTIPPGTOCOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
     ElseIf strOPER = "E" Then
     
        '' Exclui especificação técnica do produto
        sSql = "Delete from SGI_CADTIPOPGTO " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngTIPPGTOCOD & vbCrLf
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

