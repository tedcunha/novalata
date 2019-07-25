VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADCONDPAGTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial  As Integer
Dim intCODPGTO As Integer
Dim strDESPGTO As String
Dim intPARPGTO As Integer
Dim arrNPCPGTO As Variant

Public Property Get CODPGTO() As Integer
    CODPGTO = intCODPGTO
End Property

Public Property Let CODPGTO(ByVal vData As Integer)
    intCODPGTO = vData
End Property

Public Property Get DESPGTO() As String
    DESPGTO = strDESPGTO
End Property

Public Property Let DESPGTO(ByVal vData As String)
    strDESPGTO = vData
End Property

Public Property Get PARPGTO() As Integer
    PARPGTO = intPARPGTO
End Property

Public Property Let PARPGTO(ByVal vData As Integer)
    intPARPGTO = vData
End Property

Public Property Get NPCPGTO() As Variant
    NPCPGTO = arrNPCPGTO
End Property

Public Property Let NPCPGTO(ByVal vData As Variant)
    arrNPCPGTO = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADCONDPAGTOP.cCaminho = cCaminho
    frmCADCONDPAGTOP.Linha = Linha
    frmCADCONDPAGTOP.FILIAL = iFilial
    frmCADCONDPAGTOP.strAcesso = strAcesso
    frmCADCONDPAGTOP.strUSUARIO = strUSUARIO
    frmCADCONDPAGTOP.Show vbModal
End Sub

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Function Pesq_CadCondPgto() As Boolean

   Pesq_CadCondPgto = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADCONDPGTO " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadCondPgto = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo grav_err
     
     GRAVA = False
     
     Dim I As Integer
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        ''' -------------------------------------------------
        ''' Grava condição de pagamento
        sSql = "Insert Into SGI_CADCONDPGTO (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO,SGI_PARCELAS) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intCODPGTO
        sSql = sSql & "                           ,'" & strDESPGTO & "'" & vbCrLf
        sSql = sSql & "                           ," & intPARPGTO & vbCrLf
        sSql = sSql & "                           )"
         
        BGRV.CommandText = sSql
        BGRV.Execute
        
        ''' -------------------------------------------------
        ''' Grava itens de pagamento
        If IsArray(arrNPCPGTO) = True Then
           For I = 1 To UBound(arrNPCPGTO)
            
               sSql = "Insert Into SGI_CONDPGTOPARC (SGI_FILIAL,SGI_CODIGO,SGI_PARC,SGI_DIAS,SGI_PORCPARC) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intCODPGTO & vbCrLf
               sSql = sSql & "                           ," & arrNPCPGTO(I, 1) & vbCrLf
               sSql = sSql & "                           ," & arrNPCPGTO(I, 3) & vbCrLf
               
               If Len(Trim(arrNPCPGTO(I, 4))) > 0 Then
                  sSql = sSql & "                           ," & Replace(arrNPCPGTO(I, 4), ",", ".") & vbCrLf
               Else
                  sSql = sSql & "                           ,0" & vbCrLf
               End If
               
               sSql = sSql & "                           )" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next I
        End If
     
     ElseIf strOPER = "A" Then
     
        ''' -------------------------------------------------
        ''' Apaga Intens de pagamento
        sSql = "Delete from SGI_CONDPGTOPARC " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODPGTO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        ''' -------------------------------------------------
        ''' Grava Condição de pagamento
        sSql = "Update SGI_CADCONDPGTO Set " & vbCrLf
        sSql = sSql & "                SGI_DESCRICAO = '" & strDESPGTO & "'" & vbCrLf
        sSql = sSql & "               ,SGI_PARCELAS  =  " & intPARPGTO & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intCODPGTO & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
     
        ''' -------------------------------------------------
        ''' Grava itens de pagamento
        If IsArray(arrNPCPGTO) = True Then
           For I = 1 To UBound(arrNPCPGTO)
            
               sSql = "Insert Into SGI_CONDPGTOPARC (SGI_FILIAL,SGI_CODIGO,SGI_PARC,SGI_DIAS,SGI_PORCPARC) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intCODPGTO & vbCrLf
               sSql = sSql & "                           ," & arrNPCPGTO(I, 1) & vbCrLf
               sSql = sSql & "                           ," & arrNPCPGTO(I, 3) & vbCrLf
               sSql = sSql & "                           ," & Replace(arrNPCPGTO(I, 4), ",", ".") & vbCrLf
               sSql = sSql & "                           )" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next I
        End If
     
     ElseIf strOPER = "E" Then
     
        ''' -------------------------------------------------
        ''' Apga Intens de condição de pagamento
        sSql = "Delete from SGI_CONDPGTOPARC " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODPGTO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        ''' -------------------------------------------------
        ''' Apaga condição de pagamento
        sSql = "Delete from SGI_CADCONDPGTO " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODPGTO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     End If
    
     adoBanco_Dados.CommitTrans
     GRAVA = True
     
     Exit Function
     
grav_err:
     
     adoBanco_Dados.RollbackTrans
    
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
     Set objErro = Nothing
    
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


Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim intReg As Integer
   Dim I      As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADCONDPGTO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODPGTO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      intCODPGTO = BREC!SGI_CODIGO
      strDESPGTO = BREC!SGI_DESCRICAO
      intPARPGTO = BREC!SGI_PARCELAS
       
      Carrega_campos = True
   End If
   
   BREC.Close
   
   If Carrega_campos = False Then Exit Function
   
   '' ---------------
   
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CONDPGTOPARC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intCODPGTO & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrNPCPGTO(1 To intReg, 1 To 3)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrNPCPGTO(intReg, 1) = BREC2!SGI_PARC
         arrNPCPGTO(intReg, 2) = BREC2!SGI_DIAS
         arrNPCPGTO(intReg, 3) = BREC2!SGI_PORCPARC
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

End Function

Public Function Atualiza(strAcao As String, lngCODIGO As Long, lngFilial_Atu As Integer, strModulo As String) As Boolean
    
On Error GoTo Erro_Atualiza

    Atualiza = False
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_ATUALIZA" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & lngFilial_Atu & vbCrLf
    sSql = sSql & "  And SGI_MODULO = '" & Trim(strModulo) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       
        sSql = "Insert Into SGI_ATUALIZA (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                  ,SGI_MODULO" & vbCrLf
        sSql = sSql & "                  ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                  ,SGI_CODIGO)" & vbCrLf
        sSql = sSql & "           Values (" & vbCrLf
        sSql = sSql & "                   " & lngFilial_Atu & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strModulo) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strAcao) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(Str(lngCODIGO)) & "'" & vbCrLf
        sSql = sSql & "                   )"
    
        BGRV.CommandText = sSql
        BGRV.Execute
    
    Else
    
        sSql = sSql & "Update SGI_ATUALIZA Set" & vbCrLf
        sSql = sSql & "           SGI_ACAO   = '" & Trim(strAcao) & "'" & vbCrLf
        sSql = sSql & "          ,SGI_CODIGO = '" & Trim(Str(lngCODIGO)) & "'" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & lngFilial_Atu & vbCrLf
        sSql = sSql & "  And SGI_MODULO = '" & Trim(strModulo) & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
    
    End If
    BREC.Close
    
    Atualiza = True
    
    adoBanco_Dados.CommitTrans
    
    Exit Function
        
Erro_Atualiza:
    
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strAcao, sSql)
     Set objErro = Nothing
    
End Function

