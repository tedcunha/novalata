VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADGRUPDESP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial      As Long
Dim lngGRPDESPCCOD As Long
Dim strGRPDESPDESC As String
Dim arrFORNEC      As Variant
Dim arrFORNECBKP   As Variant

Public Property Get FORNECBKP() As Variant
    FORNECBKP = arrFORNECBKP
End Property

Public Property Let FORNECBKP(ByVal vData As Variant)
    arrFORNECBKP = vData
End Property


Public Property Get FORNEC() As Variant
    FORNEC = arrFORNEC
End Property

Public Property Let FORNEC(ByVal vData As Variant)
    arrFORNEC = vData
End Property

Public Property Get GRPDESPCCOD() As Long
    GRPDESPCCOD = lngGRPDESPCCOD
End Property

Public Property Let GRPDESPCCOD(ByVal vData As Long)
    lngGRPDESPCCOD = vData
End Property

Public Property Get GRPDESPDESC() As String
    GRPDESPDESC = strGRPDESPDESC
End Property

Public Property Let GRPDESPDESC(ByVal vData As String)
    strGRPDESPDESC = vData
End Property

Public Property Get Filial() As Long
    Filial = intFilial
End Property

Public Property Let Filial(ByVal vData As Long)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADGRUPDESPP.cCaminho = cCaminho
    frmCADGRUPDESPP.Linha = Linha
    frmCADGRUPDESPP.Filial = iFilial
    frmCADGRUPDESPP.strAcesso = strAcesso
    frmCADGRUPDESPP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADGRUPDESP.cCaminho = cCaminho
    frmCADGRUPDESP.Linha = Linha
    frmCADGRUPDESP.Filial = iFilial
    frmCADGRUPDESP.strAcesso = strAcesso
    frmCADGRUPDESP.strUSUARIO = strUSUARIO
    frmCADGRUPDESP.cTipOper = "I"
    frmCADGRUPDESP.Show vbModal
End Sub

Public Function Pesq_CadGrupDesp() As Boolean

   Pesq_CadGrupDesp = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADGRUPDESP " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadGrupDesp = True
   BREC.Close

End Function

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim intREG As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADGRUPDESP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngGRPDESPCCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      lngGRPDESPCCOD = BREC!SGI_CODIGO
      strGRPDESPDESC = BREC!SGI_DESCRICAO
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   '' -----------------------------------------------------------
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_GRPDESPFORN " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODGRPDSP = " & lngGRPDESPCCOD
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
      intREG = 0
      Do While Not BREC2.EOF
         intREG = intREG + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrFORNEC(1 To intREG) As String
      ReDim arrFORNECBKP(1 To intREG) As String
      BREC2.MoveFirst
      intREG = 1
      Do While Not BREC2.EOF
         arrFORNEC(intREG) = Trim(Str(BREC2!SGI_CODFORN))
         arrFORNECBKP(intREG) = Trim(Str(BREC2!SGI_CODFORN))
         intREG = intREG + 1
         BREC2.MoveNext
      Loop
   End If
   BREC2.Close
   '' -----------------------------------------------------------

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
     
     Dim I As Integer
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
       
        '' Grava especificação técnica do produto
        sSql = "Insert Into SGI_CADGRUPDESP (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & lngGRPDESPCCOD
        sSql = sSql & "                           ,'" & strGRPDESPDESC & "'" & vbCrLf
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrFORNEC) Then
           For I = 1 To UBound(arrFORNEC)
               sSql = "Insert into SGI_GRPDESPFORN (SGI_FILIAL,SGI_CODGRPDSP,SGI_CODFORN) " & vbCrLf
               sSql = sSql & "                     Values (" & vbCrLf
               sSql = sSql & "                              " & intFilial & vbCrLf
               sSql = sSql & "                             ," & lngGRPDESPCCOD & vbCrLf
               sSql = sSql & "                             ," & arrFORNEC(I) & ")"
           
               BGRV.CommandText = sSql
               BGRV.Execute
           Next I
        End If
     
     ElseIf strOPER = "A" Then
       
        '' ----------------------------------------------
        sSql = "Delete From SGI_GRPDESPFORN " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODGRPDSP = " & lngGRPDESPCCOD
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' altera especificação técnica do produto
        sSql = "Update SGI_CADGRUPDESP Set " & vbCrLf
        sSql = sSql & "                SGI_DESCRICAO = '" & strGRPDESPDESC & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & lngGRPDESPCCOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ----------------------------------------------
        If IsArray(arrFORNEC) Then
           For I = 1 To UBound(arrFORNEC)
               sSql = "Insert into SGI_GRPDESPFORN (SGI_FILIAL,SGI_CODGRPDSP,SGI_CODFORN) " & vbCrLf
               sSql = sSql & "                     Values (" & vbCrLf
               sSql = sSql & "                              " & intFilial & vbCrLf
               sSql = sSql & "                             ," & lngGRPDESPCCOD & vbCrLf
               sSql = sSql & "                             ," & arrFORNEC(I) & ")"
               BGRV.CommandText = sSql
               BGRV.Execute
           Next I
        End If
     
     ElseIf strOPER = "E" Then
     
        '' Exclui especificação técnica do produto
        sSql = "Delete From SGI_GRPDESPFORN " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODGRPDSP = " & lngGRPDESPCCOD
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Delete from SGI_CADGRUPDESP " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngGRPDESPCCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ----------------------------------------------
     
     End If
     
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

Public Sub PreencheComboFornec(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       FORN.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADFORNEC  FORN " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       FORN.SGI_FILIAL  = " & intFilial & vbCrLf
     sSql = sSql & "   And FORN.SGI_STATUS  = 2" & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_RAZAOSOC)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

