VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADTIPPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intTIPPRODCOD As Integer
Dim strTIPPRODESC As String
Dim intFilial     As Integer
Dim intTemLst     As Integer
Dim intCompLst    As Integer
Dim lngHOMOLOG    As Long

Public Property Get HOMOLOG() As Long
    HOMOLOG = lngHOMOLOG
End Property

Public Property Let HOMOLOG(ByVal vData As Long)
    lngHOMOLOG = vData
End Property

Public Property Get TEMLST() As Integer
    TEMLST = intTemLst
End Property

Public Property Let TEMLST(ByVal vData As Integer)
    intTemLst = vData
End Property

Public Property Get COMPLST() As Integer
    COMPLST = intCompLst
End Property

Public Property Let COMPLST(ByVal vData As Integer)
    intCompLst = vData
End Property

Public Property Get TIPPRODCOD() As Integer
    TIPPRODCOD = intTIPPRODCOD
End Property

Public Property Let TIPPRODCOD(ByVal vData As Integer)
    intTIPPRODCOD = vData
End Property

Public Property Get TIPPRODESC() As String
    TIPPRODESC = strTIPPRODESC
End Property

Public Property Let TIPPRODESC(ByVal vData As String)
    strTIPPRODESC = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADTIPPROP.cCaminho = cCaminho
    frmCADTIPPROP.Linha = Linha
    frmCADTIPPROP.FILIAL = iFilial
    frmCADTIPPROP.strAcesso = strAcesso
    frmCADTIPPROP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADTIPPRO.cCaminho = cCaminho
    frmCADTIPPRO.Linha = Linha
    frmCADTIPPRO.FILIAL = iFilial
    frmCADTIPPRO.strAcesso = strAcesso
    frmCADTIPPRO.strUSUARIO = strUSUARIO
    frmCADTIPPRO.cTipOper = "I"
    frmCADTIPPRO.Show vbModal
End Sub

Public Function Pesq_CadTipProd() As Boolean

   Pesq_CadTipProd = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADTIPPROD " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
   
      intTIPPRODCOD = BREC!SGI_CODIGO
      strTIPPRODESC = BREC!SGI_DESCRICAO
      intTemLst = IIf(IsNull(BREC!SGI_TEMLST) = False, BREC!SGI_TEMLST, 0)
      intCompLst = IIf(IsNull(BREC!SGI_COMPLST) = False, BREC!SGI_COMPLST, 0)
   
      Pesq_CadTipProd = True
   End If
   
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
        
        '' Grava Tipo de Produto
        sSql = "Insert Into SGI_CADTIPPROD (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO,SGI_TEMLST,SGI_COMPLST,SGI_HOMOLOGSN) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intTIPPRODCOD
        sSql = sSql & "                           ,'" & strTIPPRODESC & "'" & vbCrLf
        sSql = sSql & "                           ," & intTemLst & vbCrLf
        sSql = sSql & "                           ," & intCompLst & vbCrLf
        sSql = sSql & "                           ," & lngHOMOLOG & vbCrLf
        sSql = sSql & "                           )"
     
     ElseIf strOPER = "A" Then
     
        '' Grava Apaga tipo de Produto
        sSql = "Update SGI_CADTIPPROD Set " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intTIPPRODCOD & vbCrLf
        sSql = sSql & "              ,SGI_DESCRICAO = '" & strTIPPRODESC & "'" & vbCrLf
        sSql = sSql & "              ,SGI_TEMLST    = " & intTemLst & vbCrLf
        sSql = sSql & "              ,SGI_COMPLST   = " & intCompLst & vbCrLf
        sSql = sSql & "              ,SGI_HOMOLOGSN = " & lngHOMOLOG & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intTIPPRODCOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Tipo de Produto
        sSql = "Delete from SGI_CADTIPPROD " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intTIPPRODCOD & vbCrLf
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
   sSql = sSql & "       SGI_CADTIPPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intTIPPRODCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intTIPPRODCOD = BREC!SGI_CODIGO
      strTIPPRODESC = BREC!SGI_DESCRICAO
      intTemLst = IIf(IsNull(BREC!SGI_TEMLST) = False, BREC!SGI_TEMLST, 0)
      intCompLst = IIf(IsNull(BREC!SGI_COMPLST) = False, BREC!SGI_COMPLST, 0)
      lngHOMOLOG = BREC!SGI_HOMOLOGSN
      
      Carrega_campos = True
   End If
   
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

