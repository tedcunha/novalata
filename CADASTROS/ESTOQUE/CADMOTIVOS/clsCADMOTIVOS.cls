VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADMOTIVOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial  As Integer
Dim intCODIGO  As Integer
Dim strDESCRI  As String

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property

Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property

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

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADMOTIVOSP.cCaminho = cCaminho
    frmCADMOTIVOSP.Linha = Linha
    frmCADMOTIVOSP.FILIAL = iFilial
    frmCADMOTIVOSP.strAcesso = strAcesso
    frmCADMOTIVOSP.strUSUARIO = strUSUARIO
    frmCADMOTIVOSP.Show vbModal
End Sub

Public Function Pesq_CadMotivos() As Boolean

   Pesq_CadMotivos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMOTIVOS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadMotivos = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
       
        '' Grava especificação técnica do produto
        sSql = "Insert Into SGI_CADMOTIVOS (SGI_FILIAL,SGI_CODIGO,SGI_DESCRI) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intCODIGO
        sSql = sSql & "                           ,'" & strDESCRI & "'" & vbCrLf
        sSql = sSql & "                           )"
     
     ElseIf strOPER = "A" Then
       
        '' altera especificação técnica do produto
        sSql = "Update SGI_CADMOTIVOS Set " & vbCrLf
        sSql = sSql & "               SGI_DESCRI = '" & strDESCRI & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intCODIGO & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
     ElseIf strOPER = "E" Then
     
        '' Exclui especificação técnica do produto
        sSql = "Delete from SGI_CADMOTIVOS " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
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

Public Function Gera_Codigo(sModulo As String) As Integer

    Gera_Codigo = 1
    
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       (Max(SGI_NUMERO) + 1) As SGI_NUMERO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_NUMERO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_MODULO = '" & sModulo & "'"
    sSql = sSql & "   And SGI_FILIAL = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados
    
    If Not BREC.EOF Then
    
       If IsNull(BREC!SGI_NUMERO) = True Then
          
          Gera_Codigo = 1
          
          sSql = "Insert into SGI_NUMERO (SGI_FILIAL,SGI_NUMERO,SGI_MODULO) Values(" & vbCrLf
          sSql = sSql & "                                              " & intFilial & vbCrLf
          sSql = sSql & "                                            ,1" & vbCrLf
          sSql = sSql & "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql & "                                          )" & vbCrLf
          
       ElseIf BREC!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql & "           SGI_NUMERO = " & BREC!SGI_NUMERO & vbCrLf
          sSql = sSql & "         Where " & vbCrLf
          sSql = sSql & "               SGI_MODULO = '" & sModulo & "'" & vbCrLf
          sSql = sSql & "           And SGI_FILIAL =  " & intFilial
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
              
    End If
    
    BREC.Close
    
End Function

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMOTIVOS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRI
      
      Carrega_campos = True
   End If
   
   BREC.Close

End Function

