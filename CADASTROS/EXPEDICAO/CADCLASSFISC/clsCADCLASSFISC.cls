VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADCLASSFISC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim lngCODIGO       As Long
Dim strLETRA        As String
Dim strNOMECLA      As String
Dim arrFAMPROD      As Variant
Dim arrNATOPER      As Variant
Dim curALIQUOTA     As Currency
Dim intTEMST        As Integer
Dim strIPITRANSF    As String
Dim strALIQII       As String

Public Property Get ALIQII() As String
    ALIQII = strALIQII
End Property

Public Property Let ALIQII(ByVal vData As String)
    strALIQII = vData
End Property

Public Property Get IPITRANSF() As String
    IPITRANSF = strIPITRANSF
End Property

Public Property Let IPITRANSF(ByVal vData As String)
    strIPITRANSF = vData
End Property


Public Property Get TEMST() As Integer
    TEMST = intTEMST
End Property

Public Property Let TEMST(ByVal vData As Integer)
    intTEMST = vData
End Property

Public Property Get NATOPER() As Variant
    NATOPER = arrNATOPER
End Property

Public Property Let NATOPER(ByVal vData As Variant)
    arrNATOPER = vData
End Property

Public Property Get ALIQUOTA() As Currency
    ALIQUOTA = curALIQUOTA
End Property

Public Property Let ALIQUOTA(ByVal vData As Currency)
    curALIQUOTA = vData
End Property

Public Property Get FAMPROD() As Variant
    FAMPROD = arrFAMPROD
End Property

Public Property Let FAMPROD(ByVal vData As Variant)
    arrFAMPROD = vData
End Property

Public Property Get NOMECLA() As String
    NOMECLA = strNOMECLA
End Property

Public Property Let NOMECLA(ByVal vData As String)
    strNOMECLA = vData
End Property

Public Property Get LETRA() As String
    LETRA = strLETRA
End Property

Public Property Let LETRA(ByVal vData As String)
    strLETRA = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property


Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADCLASSFISCP.cCaminho = cCaminho
    frmCADCLASSFISCP.Linha = Linha
    frmCADCLASSFISCP.FILIAL = iFilial
    frmCADCLASSFISCP.strAcesso = strAcesso
    frmCADCLASSFISCP.strUSUARIO = strUSUARIO
    frmCADCLASSFISCP.Show vbModal
End Sub

Public Function Pesq_CadClassFisc() As Boolean

   Pesq_CadClassFisc = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADCLASSFIS " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadClassFisc = True
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

     Dim I          As Integer
     Dim strValor   As String
     
     GRAVA = False
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava Classificação
        sSql = sSql & "Insert Into SGI_CADCLASSFIS (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                            ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                            ,SGI_CODCLASS" & vbCrLf
        sSql = sSql & "                            ,SGI_NOMECLA" & vbCrLf
        sSql = sSql & "                            ,SGI_ALIQUOTA" & vbCrLf
        sSql = sSql & "                            ,SGI_ALIQIPITRANSF" & vbCrLf
        sSql = sSql & "                            ,SGI_TEMST" & vbCrLf
        sSql = sSql & "                            ,SGI_ALIQII" & vbCrLf
        
        sSql = sSql & "                   ) Values (" & intFilial & vbCrLf
        sSql = sSql & "                           ," & lngCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & strLETRA & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strNOMECLA & "'" & vbCrLf
        
        If curALIQUOTA > 0 Then
            strValor = Replace(Format(curALIQUOTA, "#,##0.00"), ".", "")
            strValor = Replace(strValor, ",", ".")
            sSql = sSql & "                           ," & strValor & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                           ," & strIPITRANSF & vbCrLf
        sSql = sSql & "                           ," & intTEMST & vbCrLf
        sSql = sSql & "                           ," & strALIQII & vbCrLf
        
        sSql = sSql & "                           )"
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Familia de Produtos
        If IsArray(arrFAMPROD) Then
        For I = 1 To UBound(arrFAMPROD)
            
            sSql = "Insert Into SGI_CADFAMPRODCLAS (" & vbCrLf
            sSql = sSql & "                                SGI_FILIAL" & vbCrLf
            sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
            sSql = sSql & "                               ,SGI_CODFAMPROD" & vbCrLf
            sSql = sSql & "                                ) Values (" & vbCrLf
            sSql = sSql & "                                " & intFilial & vbCrLf
            sSql = sSql & "                               ," & lngCODIGO & vbCrLf
            sSql = sSql & "                               ," & arrFAMPROD(I) & vbCrLf
            sSql = sSql & "                                )"
            
            BGRV.CommandText = sSql
            BGRV.Execute
            
        Next I
        End If

        '' Natureza de Operação
        If IsArray(arrNATOPER) Then
            For I = 1 To UBound(arrNATOPER)
                
                sSql = "Insert Into SGI_CADNATOPERCLA (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_CODNATOPER" & vbCrLf
                sSql = sSql & "                                ) Values (" & vbCrLf
                sSql = sSql & "                                " & intFilial & vbCrLf
                sSql = sSql & "                               ," & lngCODIGO & vbCrLf
                sSql = sSql & "                               ," & arrNATOPER(I) & vbCrLf
                sSql = sSql & "                                )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If

     ElseIf strOPER = "A" Then
     
        '' Delete Familia de Produtos
        sSql = "Delete from SGI_CADFAMPRODCLAS " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
                
        '' Delete Natureza de Operação
        sSql = "Delete from SGI_CADNATOPERCLA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
                
                
        '' Altera Classificação Fiscal
        sSql = "Update SGI_CADCLASSFIS Set " & vbCrLf
        sSql = sSql & "               SGI_CODCLASS   = '" & strLETRA & "'" & vbCrLf
        sSql = sSql & "              ,SGI_NOMECLA    = '" & strNOMECLA & "'" & vbCrLf
        sSql = sSql & "              ,SGI_TEMST      = " & intTEMST & vbCrLf
        
        If curALIQUOTA > 0 Then
            strValor = Replace(Format(curALIQUOTA, "#,##0.00"), ".", "")
            strValor = Replace(strValor, ",", ".")
            sSql = sSql & "                           ,SGI_ALIQUOTA = " & strValor & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_ALIQUOTA = Null" & vbCrLf
        End If
         
        sSql = sSql & "              ,SGI_ALIQIPITRANSF = " & strIPITRANSF & vbCrLf
        sSql = sSql & "              ,SGI_ALIQII        = " & strALIQII & vbCrLf
        
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & lngCODIGO & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Familia de Produtos
        If IsArray(arrFAMPROD) Then
            For I = 1 To UBound(arrFAMPROD)
                
                sSql = "Insert Into SGI_CADFAMPRODCLAS (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_CODFAMPROD" & vbCrLf
                sSql = sSql & "                                ) Values (" & vbCrLf
                sSql = sSql & "                                " & intFilial & vbCrLf
                sSql = sSql & "                               ," & lngCODIGO & vbCrLf
                sSql = sSql & "                               ," & arrFAMPROD(I) & vbCrLf
                sSql = sSql & "                                )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
     
        '' Natureza de Operação
        If IsArray(arrNATOPER) Then
            For I = 1 To UBound(arrNATOPER)
                
                sSql = "Insert Into SGI_CADNATOPERCLA (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_CODNATOPER" & vbCrLf
                sSql = sSql & "                                ) Values (" & vbCrLf
                sSql = sSql & "                                " & intFilial & vbCrLf
                sSql = sSql & "                               ," & lngCODIGO & vbCrLf
                sSql = sSql & "                               ," & arrNATOPER(I) & vbCrLf
                sSql = sSql & "                                )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Familia de Produtos
        sSql = "Delete from SGI_CADFAMPRODCLAS " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Delete Natureza de Operação
        sSql = "Delete from SGI_CADNATOPERCLA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Classificação Fiscal
        sSql = "Delete from SGI_CADCLASSFIS " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
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
   sSql = sSql & "       SGI_CADCLASSFIS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      lngCODIGO = BREC!SGI_CODIGO
      strLETRA = BREC!SGI_CODCLASS
      strNOMECLA = BREC!SGI_NOMECLA
      If Not IsNull(BREC!SGI_ALIQUOTA) Then curALIQUOTA = BREC!SGI_ALIQUOTA
      intTEMST = BREC!SGI_TEMST
      If Not IsNull(BREC!SGI_ALIQIPITRANSF) Then strIPITRANSF = Format(BREC!SGI_ALIQIPITRANSF, "#,##0.00")
      If Not IsNull(BREC!SGI_ALIQII) Then strALIQII = Format(BREC!SGI_ALIQII, "#,##0.00")
      Carrega_campos = True
   End If
   
   BREC.Close

    '' ---------------------------------------------------------
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFAMPRODCLAS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO & vbCrLf
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrFAMPROD(1 To intQTDREG) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF
          
          arrFAMPROD(intQTDREG) = BREC3!SGI_CODFAMPROD
          
          intQTDREG = intQTDREG + 1
          
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close

    '' ---------------------------------------------------------
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADNATOPERCLA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO & vbCrLf
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrNATOPER(1 To intQTDREG) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF
          
          arrNATOPER(intQTDREG) = BREC3!SGI_CODNATOPER
          
          intQTDREG = intQTDREG + 1
          
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close

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


