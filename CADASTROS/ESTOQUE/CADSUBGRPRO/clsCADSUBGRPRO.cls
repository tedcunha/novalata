VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADSUBGRPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intSUBGRPRODCOD As Integer
Dim strSUBGRPRODESC As String
Dim vriPRODESPECIE  As Variant
Dim intFilial       As Integer

Public Property Get SUBGRPRODCOD() As Integer
    SUBGRPRODCOD = intSUBGRPRODCOD
End Property

Public Property Let SUBGRPRODCOD(ByVal vData As Integer)
    intSUBGRPRODCOD = vData
End Property

Public Property Get SUBGRPRODESC() As String
    SUBGRPRODESC = strSUBGRPRODESC
End Property

Public Property Let SUBGRPRODESC(ByVal vData As String)
    strSUBGRPRODESC = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get PRODESPECIE() As Variant
    PRODESPECIE = vriPRODESPECIE
End Property

Public Property Let PRODESPECIE(ByVal vData As Variant)
    vriPRODESPECIE = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADSUBGRPROP.cCaminho = cCaminho
    frmCADSUBGRPROP.Linha = Linha
    frmCADSUBGRPROP.FILIAL = iFilial
    frmCADSUBGRPROP.strAcesso = strAcesso
    frmCADSUBGRPROP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADSUBGRPRO.cCaminho = cCaminho
    frmCADSUBGRPRO.Linha = Linha
    frmCADSUBGRPRO.FILIAL = iFilial
    frmCADSUBGRPRO.strAcesso = strAcesso
    frmCADSUBGRPRO.strUSUARIO = strUSUARIO
    frmCADSUBGRPRO.cTipOper = "I"
    frmCADSUBGRPRO.Show vbModal
End Sub

Public Function Pesq_CadSubGruProd() As Boolean

   Pesq_CadSubGruProd = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       SUB.SGI_CODIGO     " & vbCrLf
   sSql = sSql + "      ,SUB.SGI_DESCRICAO  " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADSUBGRPROD SUB " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SUB.SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
   
      intSUBGRPRODCOD = BREC!SGI_CODIGO
      strSUBGRPRODESC = BREC!SGI_DESCRICAO
   
      Pesq_CadSubGruProd = True
   End If
   
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim I As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava Sybgrupo de Produto
        sSql = "Insert Into SGI_CADSUBGRPROD (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intSUBGRPRODCOD
        sSql = sSql & "                           ,'" & strSUBGRPRODESC & "'" & vbCrLf
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava os Sub itens Especie do Produto
        If IsArray(vriPRODESPECIE) = True Then
           For I = 1 To UBound(vriPRODESPECIE)
               
               sSql = "Insert Into SGI_SUBGRUPRODITEN (SGI_FILIAL,SGI_CODIGO,SGI_CODESPECIE) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial
               sSql = sSql & "                           ," & intSUBGRPRODCOD
               sSql = sSql & "                           ," & vriPRODESPECIE(I) & vbCrLf
               sSql = sSql & "                           )"
           
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
        
     
     ElseIf strOPER = "A" Then
     
        '' Apaga Itens de Sub grupo de produtos
        sSql = "Delete from SGI_SUBGRUPRODITEN " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intSUBGRPRODCOD & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava Sub grupo de Produtos
        sSql = "Update SGI_CADSUBGRPROD Set " & vbCrLf
        sSql = sSql & "               SGI_DESCRICAO  = '" & strSUBGRPRODESC & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intSUBGRPRODCOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava os Sub itens Especie do Produto
        If IsArray(vriPRODESPECIE) = True Then
           For I = 1 To UBound(vriPRODESPECIE)
               
               sSql = "Insert Into SGI_SUBGRUPRODITEN (SGI_FILIAL,SGI_CODIGO,SGI_CODESPECIE) " & vbCrLf
               sSql = sSql & "                 Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intSUBGRPRODCOD & vbCrLf
               sSql = sSql & "                           ," & vriPRODESPECIE(I) & vbCrLf
               sSql = sSql & "                           )"
           
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
        
     ElseIf strOPER = "E" Then
     
        '' Apaga Sub grupos de Produto
        sSql = "Delete from SGI_CADSUBGRPROD " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intSUBGRPRODCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Apaga Itens de Sub grupo de produtos
        sSql = "Delete from SGI_SUBGRUPRODITEN " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intSUBGRPRODCOD & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     End If
     
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

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim intReg As Integer
   Dim I      As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADSUBGRPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intSUBGRPRODCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      intSUBGRPRODCOD = BREC!SGI_CODIGO
      strSUBGRPRODESC = BREC!SGI_DESCRICAO
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   If Carrega_campos = False Then Exit Function
   
   '' ---------------
   
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_SUBGRUPRODITEN " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intSUBGRPRODCOD & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim vriPRODESPECIE(1 To intReg)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         vriPRODESPECIE(intReg) = BREC2!SGI_CODESPECIE
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

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

Public Sub PreenchComboEspecie(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADESPPROD " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
    BREC.Close

End Sub
