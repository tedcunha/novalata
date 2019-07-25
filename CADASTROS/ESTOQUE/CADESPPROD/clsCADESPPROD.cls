VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADESPPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intESPPRODCOD   As Integer
Dim strESPPRODESC   As String
Dim intFilial       As Integer
Dim arrCORPO        As Variant
Dim arrTAMPA        As Variant
Dim arrFUNDO        As Variant
Dim arrARGOLA       As Variant
Dim intVern01       As Integer
Dim intVern02       As Integer

Public Property Get Vern02() As Integer
    Vern02 = intVern02
End Property

Public Property Let Vern02(ByVal vData As Integer)
    intVern02 = vData
End Property

Public Property Get Vern01() As Integer
    Vern01 = intVern01
End Property

Public Property Let Vern01(ByVal vData As Integer)
    intVern01 = vData
End Property

Public Property Get CORPO() As Variant
    CORPO = arrCORPO
End Property

Public Property Let CORPO(ByVal vData As Variant)
    arrCORPO = vData
End Property

Public Property Get TAMPA() As Variant
    TAMPA = arrTAMPA
End Property

Public Property Let TAMPA(ByVal vData As Variant)
    arrTAMPA = vData
End Property

Public Property Get FUNDO() As Variant
    FUNDO = arrFUNDO
End Property

Public Property Let FUNDO(ByVal vData As Variant)
    arrFUNDO = vData
End Property

Public Property Get ARGOLA() As Variant
    ARGOLA = arrARGOLA
End Property

Public Property Let ARGOLA(ByVal vData As Variant)
    arrARGOLA = vData
End Property

Public Property Get ESPPRODCOD() As Integer
    ESPPRODCOD = intESPPRODCOD
End Property

Public Property Let ESPPRODCOD(ByVal vData As Integer)
    intESPPRODCOD = vData
End Property

Public Property Get ESPPRODESC() As String
    ESPPRODESC = strESPPRODESC
End Property

Public Property Let ESPPRODESC(ByVal vData As String)
    strESPPRODESC = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADESPPRODP.cCaminho = cCaminho
    frmCADESPPRODP.Linha = Linha
    frmCADESPPRODP.FILIAL = iFilial
    frmCADESPPRODP.strAcesso = strAcesso
    frmCADESPPRODP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADESPPROD.cCaminho = cCaminho
    frmCADESPPROD.Linha = Linha
    frmCADESPPROD.FILIAL = iFilial
    frmCADESPPROD.strAcesso = strAcesso
    frmCADESPPROD.strUSUARIO = strUSUARIO
    frmCADESPPROD.cTipOper = "I"
    frmCADESPPROD.Show vbModal
End Sub

Public Function Pesq_CadEspProd() As Boolean

   Pesq_CadEspProd = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADESPPROD " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
   
      intESPPRODCOD = BREC!SGI_CODIGO
      strESPPRODESC = BREC!SGI_DESCRICAO
   
      Pesq_CadEspProd = True
   End If
   
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava

     Dim I As Integer

     GRAVA = False
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava Tipo de Produto
        sSql = ""
        
        sSql = "Insert Into SGI_CADESPPROD (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO,SGI_Vern01,SGI_Vern02) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intESPPRODCOD
        sSql = sSql & "                           ,'" & strESPPRODESC & "'" & vbCrLf
        sSql = sSql & "                           ," & intVern01 & vbCrLf
        sSql = sSql & "                           ," & intVern02 & vbCrLf
        sSql = sSql & "                           )"
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
Voltar:
     
        '' Corpo
        If IsArray(arrCORPO) Then
            For I = 1 To UBound(arrCORPO)
     
                sSql = ""
                
                sSql = "Insert Into SGI_CADESPROD_CORPO (" & vbCrLf
                sSql = sSql & "                                 SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                ,SGI_COD" & vbCrLf
                sSql = sSql & "                       ) Values (" & vbCrLf
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intESPPRODCOD & vbCrLf
                sSql = sSql & "                                ," & arrCORPO(I) & vbCrLf
                sSql = sSql & "                                )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
     
        '' Tampa
        If IsArray(arrTAMPA) Then
            For I = 1 To UBound(arrTAMPA)
     
                sSql = ""
                
                sSql = "Insert Into SGI_CADESPROD_TAMPA (" & vbCrLf
                sSql = sSql & "                                 SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                ,SGI_COD" & vbCrLf
                sSql = sSql & "                       ) Values (" & vbCrLf
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intESPPRODCOD & vbCrLf
                sSql = sSql & "                                ," & arrTAMPA(I) & vbCrLf
                sSql = sSql & "                                )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
     
        '' Fundo
        If IsArray(arrFUNDO) Then
            For I = 1 To UBound(arrFUNDO)
     
                sSql = ""
                
                sSql = "Insert Into SGI_CADESPROD_FUNDO (" & vbCrLf
                sSql = sSql & "                                 SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                ,SGI_COD" & vbCrLf
                sSql = sSql & "                       ) Values (" & vbCrLf
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intESPPRODCOD & vbCrLf
                sSql = sSql & "                                ," & arrFUNDO(I) & vbCrLf
                sSql = sSql & "                                )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
     
        '' Argola
        If IsArray(arrARGOLA) Then
            For I = 1 To UBound(arrARGOLA)
     
                sSql = ""
                
                sSql = "Insert Into SGI_CADESPROD_ARGOLA (" & vbCrLf
                sSql = sSql & "                                 SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                ,SGI_COD" & vbCrLf
                sSql = sSql & "                       ) Values (" & vbCrLf
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intESPPRODCOD & vbCrLf
                sSql = sSql & "                                ," & arrARGOLA(I) & vbCrLf
                sSql = sSql & "                                )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
     
     
     ElseIf strOPER = "A" Or strOPER = "E" Then
     
        '' Grava Apaga tipo de Produto
        sSql = "Update SGI_CADESPPROD Set " & vbCrLf
        sSql = sSql & "               SGI_DESCRICAO = '" & strESPPRODESC & "'" & vbCrLf
        sSql = sSql & "              ,SGI_Vern01 = " & intVern01 & vbCrLf
        sSql = sSql & "              ,SGI_Vern02 = " & intVern02 & vbCrLf
        
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intESPPRODCOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Corpo
        sSql = "Delete from SGI_CADESPROD_CORPO " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intESPPRODCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Tampa
        sSql = "Delete from SGI_CADESPROD_TAMPA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intESPPRODCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Fundo
        sSql = "Delete from SGI_CADESPROD_FUNDO " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intESPPRODCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Argola
        sSql = "Delete from SGI_CADESPROD_ARGOLA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intESPPRODCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo Voltar
     
        '' Apaga Tipo de Produto
        sSql = "Delete from SGI_CADESPPROD " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intESPPRODCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
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

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim intQTDREG As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADESPPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intESPPRODCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intESPPRODCOD = BREC!SGI_CODIGO
      strESPPRODESC = BREC!SGI_DESCRICAO
      
      intVern01 = BREC!SGI_Vern01
      intVern02 = BREC!SGI_Vern02
      
      Carrega_campos = True
   End If
   
   BREC.Close

    '' ==================================================================
    '' Corpo
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADESPROD_CORPO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intESPPRODCOD
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrCORPO(1 To intQTDREG) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrCORPO(intQTDREG) = BREC6!SGI_COD
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================

    '' ==================================================================
    '' Tampa
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADESPROD_TAMPA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intESPPRODCOD
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrTAMPA(1 To intQTDREG) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrTAMPA(intQTDREG) = BREC6!SGI_COD
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================

    '' ==================================================================
    '' Fundo
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADESPROD_FUNDO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intESPPRODCOD
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrFUNDO(1 To intQTDREG) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrFUNDO(intQTDREG) = BREC6!SGI_COD
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================

    '' ==================================================================
    '' Argola
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADESPROD_ARGOLA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intESPPRODCOD
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrARGOLA(1 To intQTDREG) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrARGOLA(intQTDREG) = BREC6!SGI_COD
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================

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

