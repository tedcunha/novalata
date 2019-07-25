VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADESPORCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial      As Integer
Dim intESPORCACOD  As Integer
Dim strESPORCADES  As String
Dim boolATIVASERV  As Boolean

Public Property Get ATIVASERV() As Boolean
    ATIVASERV = boolATIVASERV
End Property

Public Property Let ATIVASERV(ByVal vData As Boolean)
    boolATIVASERV = vData
End Property

Public Property Get ESPORCADES() As String
    ESPORCADES = strESPORCADES
End Property

Public Property Let ESPORCADES(ByVal vData As String)
    strESPORCADES = vData
End Property

Public Property Get ESPORCACOD() As Integer
    ESPORCACOD = intESPORCACOD
End Property

Public Property Let ESPORCACOD(ByVal vData As Integer)
    intESPORCACOD = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADESPORCAP.cCaminho = cCaminho
    frmCADESPORCAP.Linha = Linha
    frmCADESPORCAP.FILIAL = iFilial
    frmCADESPORCAP.strAcesso = strAcesso
    frmCADESPORCAP.strUSUARIO = strUSUARIO
    frmCADESPORCAP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADESPORCA.cCaminho = cCaminho
    frmCADESPORCA.Linha = Linha
    frmCADESPORCA.FILIAL = iFilial
    frmCADESPORCA.strAcesso = strAcesso
    frmCADESPORCA.strUSUARIO = strUSUARIO
    frmCADESPORCA.cTipOper = "I"
    frmCADESPORCA.Show vbModal
End Sub


Public Function Pesq_CadEsporca() As Boolean

   Pesq_CadEsporca = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADESPORCA " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadEsporca = True
   BREC.Close

End Function


Public Function Gera_Codigo(sModulo As String) As Long

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
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava especie de orçamento
        sSql = "Insert Into SGI_CADESPORCA (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO,SGI_ATIVASERV) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intESPORCACOD
        sSql = sSql & "                           ,'" & strESPORCADES & "'" & vbCrLf
        sSql = sSql & "                           ,'" & boolATIVASERV & "'" & vbCrLf
        sSql = sSql & "                           )"
     
     ElseIf strOPER = "A" Then
     
        '' Grava especie de orçamento
        sSql = "Update SGI_CADESPORCA Set " & vbCrLf
        sSql = sSql & "               SGI_DESCRICAO = '" & strESPORCADES & "'" & vbCrLf
        sSql = sSql & "              ,SGI_ATIVASERV = '" & boolATIVASERV & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intESPORCACOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
     ElseIf strOPER = "E" Then
     
        '' Apaga especie de orçamento
        sSql = "Delete from SGI_CADESPORCA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intESPORCACOD & vbCrLf
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


Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADESPORCA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intESPORCACOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intESPORCACOD = BREC!SGI_CODIGO
      strESPORCADES = BREC!SGI_DESCRICAO
      boolATIVASERV = BREC!SGI_ATIVASERV
      
      Carrega_campos = True
   End If
   
   BREC.Close

End Function

