VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADDADOSOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial        As Integer
Dim lngCODIGO        As Long
Dim lngCODOP         As Long
Dim lngIDPRODUTO     As Long
Dim dtDATALOTE       As Date
Dim arrVERNIZ        As Variant
Dim arrCORES         As Variant

Public Property Get IDPRODUTO() As Long
    IDPRODUTO = lngIDPRODUTO
End Property

Public Property Let IDPRODUTO(ByVal vData As Long)
    lngIDPRODUTO = vData
End Property

Public Property Get CORES() As Variant
    CORES = arrCORES
End Property

Public Property Let CORES(ByVal vData As Variant)
    arrCORES = vData
End Property

Public Property Get VERNIZ() As Variant
    VERNIZ = arrVERNIZ
End Property

Public Property Let VERNIZ(ByVal vData As Variant)
    arrVERNIZ = vData
End Property

Public Property Get DATALOTE() As Date
    DATALOTE = dtDATALOTE
End Property

Public Property Let DATALOTE(ByVal vData As Date)
    dtDATALOTE = vData
End Property

Public Property Get CODOP() As Long
    CODOP = lngCODOP
End Property

Public Property Let CODOP(ByVal vData As Long)
    lngCODOP = vData
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

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long)
    frmCADDADOSOPP.cCaminho = cCaminho
    frmCADDADOSOPP.Linha = Linha
    frmCADDADOSOPP.FILIAL = iFilial
    frmCADDADOSOPP.strAcesso = strAcesso
    frmCADDADOSOPP.strUsuario = strUsuario
    frmCADDADOSOPP.lngCodUsuaro = lngIDUsuario
    frmCADDADOSOPP.Show vbModal
End Sub

Public Function AtivoDesativo() As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADDADOSOP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
     
     Dim I  As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
      
     If strOPER = "I" Then '' Inclui
        
Volta_Grava:

        sSql = "Insert Into SGI_CADDADOSOP (" & vbCrLf
        sSql = sSql & "                            SGI_FILIAL" & vbCrLf
        sSql = sSql & "                           ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                           ,SGI_DATALOTE" & vbCrLf
        sSql = sSql & "                           ,SGI_CODOP" & vbCrLf
        sSql = sSql & "                           ,SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "                  ) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & lngCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & Format(dtDATALOTE, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                           ," & lngCODOP & vbCrLf
        sSql = sSql & "                           ," & lngIDPRODUTO & vbCrLf
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' ---------------------------------
        '' Dados OP Verniz
        If IsArray(arrVERNIZ) Then
            For I = 1 To UBound(arrVERNIZ)
            
                If Len(Trim(arrVERNIZ(I, 1))) > 0 Then
            
                sSql = "Insert Into SGI_DADOSOPVERNIZ (" & vbCrLf
                sSql = sSql & "                               SGI_FILIAL" & vbCrLf
                sSql = sSql & "                              ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                              ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                              ,SGI_LOTE" & vbCrLf
                sSql = sSql & "                              ,SGI_PRODUCAO" & vbCrLf
                sSql = sSql & "                              ,SGI_DATPROD" & vbCrLf
                sSql = sSql & "                              ,SGI_CODMAQ" & vbCrLf
                sSql = sSql & "                              ,SGI_CODTUR" & vbCrLf
                sSql = sSql & "                              ,SGI_TIPO" & vbCrLf
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & lngCODIGO & vbCrLf
                sSql = sSql & "                              ," & arrVERNIZ(I, 1) & vbCrLf
                sSql = sSql & "                              ,'" & Trim(arrVERNIZ(I, 2)) & "'" & vbCrLf
                sSql = sSql & "                              ," & Trim(arrVERNIZ(I, 3)) & vbCrLf
                sSql = sSql & "                              ,'" & Trim(arrVERNIZ(I, 4)) & "'" & vbCrLf
                sSql = sSql & "                              ," & Trim(arrVERNIZ(I, 5)) & vbCrLf
                sSql = sSql & "                              ," & Trim(arrVERNIZ(I, 6)) & vbCrLf
                sSql = sSql & "                              ," & arrVERNIZ(I, 7) & vbCrLf
                sSql = sSql & "                              )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                End If
            
            Next I
        End If
     
        '' ---------------------------------
        '' Dados OP Cores
        If IsArray(arrCORES) Then
            For I = 1 To UBound(arrCORES)
            
                sSql = "Insert Into SGI_DADOSOPCORES (" & vbCrLf
                sSql = sSql & "                               SGI_FILIAL" & vbCrLf
                sSql = sSql & "                              ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                              ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                              ,SGI_LOTE" & vbCrLf
                sSql = sSql & "                              ,SGI_PRODUCAO" & vbCrLf
                sSql = sSql & "                              ,SGI_DATPROD" & vbCrLf
                sSql = sSql & "                              ,SGI_CODMAQ" & vbCrLf
                sSql = sSql & "                              ,SGI_CODTUR" & vbCrLf
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & lngCODIGO & vbCrLf
                sSql = sSql & "                              ," & arrCORES(I, 1) & vbCrLf
                sSql = sSql & "                              ,'" & Trim(arrCORES(I, 2)) & "'" & vbCrLf
                sSql = sSql & "                              ," & Trim(arrCORES(I, 3)) & vbCrLf
                sSql = sSql & "                              ,'" & Trim(arrCORES(I, 4)) & "'" & vbCrLf
                sSql = sSql & "                              ," & Trim(arrCORES(I, 5)) & vbCrLf
                sSql = sSql & "                              ," & Trim(arrCORES(I, 6)) & vbCrLf
                sSql = sSql & "                              )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
        
        
     ElseIf strOPER = "A" Then '' Altera
     
            '' Apagando Cores
            sSql = "Delete From SGI_DADOSOPCORES " & vbCrLf
            sSql = sSql & "      Where " & vbCrLf
            sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
            BGRV.CommandText = sSql
            BGRV.Execute
        
            '' Apagando Verniz
            sSql = "Delete From SGI_DADOSOPVERNIZ " & vbCrLf
            sSql = sSql & "      Where " & vbCrLf
            sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
            BGRV.CommandText = sSql
            BGRV.Execute
        
            '' Apagando Dados
            sSql = "Delete From SGI_CADDADOSOP " & vbCrLf
            sSql = sSql & "      Where " & vbCrLf
            sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
            BGRV.CommandText = sSql
            BGRV.Execute
            
            If strOPER = "A" Then GoTo Volta_Grava
     
     ElseIf strOPER = "E" Then '' Exclui
        
            '' Apagando Cores
            sSql = "Delete From SGI_DADOSOPCORES " & vbCrLf
            sSql = sSql & "      Where " & vbCrLf
            sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
            BGRV.CommandText = sSql
            BGRV.Execute
        
            '' Apagando Verniz
            sSql = "Delete From SGI_DADOSOPVERNIZ " & vbCrLf
            sSql = sSql & "      Where " & vbCrLf
            sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
            BGRV.CommandText = sSql
            BGRV.Execute
        
            '' Apagando Dados
            sSql = "Delete From SGI_CADDADOSOP " & vbCrLf
            sSql = sSql & "      Where " & vbCrLf
            sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
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


Public Function Carrega_Campos() As Boolean

    Carrega_Campos = False

    Dim I               As Integer
    Dim intQTDREG       As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDADOSOP  " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       BREC.Close
       Exit Function
    End If
    
    lngCODIGO = BREC!SGI_CODIGO
    dtDATALOTE = BREC!SGI_DATALOTE
    lngCODOP = BREC!SGI_CODOP
    lngIDPRODUTO = BREC!SGI_IDPRODUTO
    
    BREC.Close
    
    Carrega_Campos = True
    ' ----------------------------------------------------------
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_DADOSOPVERNIZ " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO  = " & lngCODIGO
    sSql = sSql & " Order By SGI_TIPO"
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF() Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrVERNIZ(1 To intQTDREG, 1 To 7) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF()
       
            arrVERNIZ(intQTDREG, 1) = BREC3!SGI_IDPRODUTO
            arrVERNIZ(intQTDREG, 2) = BREC3!SGI_LOTE
            arrVERNIZ(intQTDREG, 3) = BREC3!SGI_PRODUCAO
            arrVERNIZ(intQTDREG, 4) = Format(BREC3!SGI_DATPROD, "DD/MM/YYYY")
            arrVERNIZ(intQTDREG, 5) = BREC3!SGI_CODMAQ
            arrVERNIZ(intQTDREG, 6) = BREC3!SGI_CODTUR
            arrVERNIZ(intQTDREG, 7) = BREC3!SGI_TIPO
          
            BREC3.MoveNext
            intQTDREG = (intQTDREG + 1)
       Loop
    End If
    BREC3.Close
    
    
    ' ----------------------------------------------------------
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_DADOSOPCORES " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO  = " & lngCODIGO
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF() Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrCORES(1 To intQTDREG, 1 To 6) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF()
       
            arrCORES(intQTDREG, 1) = BREC3!SGI_IDPRODUTO
            arrCORES(intQTDREG, 2) = BREC3!SGI_LOTE
            arrCORES(intQTDREG, 3) = BREC3!SGI_PRODUCAO
            arrCORES(intQTDREG, 4) = Format(BREC3!SGI_DATPROD, "DD/MM/YYYY")
            arrCORES(intQTDREG, 5) = BREC3!SGI_CODMAQ
            arrCORES(intQTDREG, 6) = BREC3!SGI_CODTUR
          
            BREC3.MoveNext
            intQTDREG = (intQTDREG + 1)
       Loop
    End If
    BREC3.Close
    
End Function

