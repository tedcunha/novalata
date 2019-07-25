VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADLIMPACARTEIRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim strCODMOTLIQOP  As String
Dim strOBSLIQ       As String
Dim strDTENTREGA    As String
Dim arrDADOSNOVA    As Variant
Dim arrDADOSSTEEL   As Variant

Public Property Get DADOSSTEEL() As Variant
    DADOSSTEEL = arrDADOSSTEEL
End Property

Public Property Let DADOSSTEEL(ByVal vData As Variant)
    arrDADOSSTEEL = vData
End Property

Public Property Get DADOSNOVA() As Variant
    DADOSNOVA = arrDADOSNOVA
End Property

Public Property Let DADOSNOVA(ByVal vData As Variant)
    arrDADOSNOVA = vData
End Property

Public Property Get DTENTREGA() As String
    OBSLIQ = strDTENTREGA
End Property

Public Property Let DTENTREGA(ByVal vData As String)
    strDTENTREGA = vData
End Property

Public Property Get OBSLIQ() As String
    OBSLIQ = strOBSLIQ
End Property

Public Property Let OBSLIQ(ByVal vData As String)
    strOBSLIQ = vData
End Property

Public Property Get CODMOTLIQOP() As String
    CODMOTLIQOP = strCODMOTLIQOP
End Property

Public Property Let CODMOTLIQOP(ByVal vData As String)
    strCODMOTLIQOP = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADLIMPACARTEIRA.cCaminho = cCaminho
    frmCADLIMPACARTEIRA.Linha = Linha
    frmCADLIMPACARTEIRA.FILIAL = iFilial
    frmCADLIMPACARTEIRA.strAcesso = strAcesso
    frmCADLIMPACARTEIRA.Show vbModal
End Sub


Public Function GravaDados(lngCodUs As Long) As Boolean
   
On Error GoTo Err_ProcessaDados
    
    GravaDados = False
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    Dim I            As Long
    Dim lngREGS      As Long
    Dim strSTATUS    As String
    Dim strSTATUSPED As String
    
    '' Grava Dados Novalata
    If IsArray(arrDADOSNOVA) Then
        For I = 1 To UBound(arrDADOSNOVA)
        
            '' =====================================================
            '' Dando Baixa na OP
            sSql = ""
            
            sSql = "Update SGI_ORDEMPROD Set" & vbCrLf
            sSql = sSql & "                                    SGI_STATUS      = 9" & vbCrLf
            sSql = sSql & "                                  , SGI_OBSLIQ      = " & strOBSLIQ & vbCrLf
            sSql = sSql & "                                  , SGI_CODMOTLIQOP = " & strCODMOTLIQOP & vbCrLf
            sSql = sSql & "Where" & vbCrLf
            sSql = sSql & "      SGI_FILIAL    = " & intFilial & vbCrLf
            sSql = sSql & "  And SGI_CODIGO    = " & arrDADOSNOVA(I, 2) & vbCrLf
            sSql = sSql & "  And SGI_IDPRODUTO = " & arrDADOSNOVA(I, 4) & vbCrLf
            If Len(Trim(arrDADOSNOVA(I, 3))) > 0 Then sSql = sSql & "  And SGI_IDPAI  = " & arrDADOSNOVA(I, 3)
        
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
                            
            '' =====================================================
            '' Gera Log das OP
            sSql = ""
            sSql = "Insert Into SGI_LOGFORM (" & vbCrLf
            sSql = sSql & "                              SGI_FILIAL" & vbCrLf
            sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
            sSql = sSql & "                             ,SGI_DATA" & vbCrLf
            sSql = sSql & "                             ,SGI_HORA" & vbCrLf
            sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
            sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
            sSql = sSql & "                             ,SGI_FORM" & vbCrLf
            sSql = sSql & "                    ) Values (" & vbCrLf
            sSql = sSql & "                              " & intFilial & vbCrLf
            sSql = sSql & "                             ," & arrDADOSNOVA(I, 2) & vbCrLf
            sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
            sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
            sSql = sSql & "                             ," & lngCodUs & vbCrLf
            sSql = sSql & "                             ,'BX'" & vbCrLf
            sSql = sSql & "                             ,'frmCADOORDFAB'" & vbCrLf
            sSql = sSql & "                             )"
            
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
                            
            
            '' =====================================================
            '' Dando Baixa no Pedido
            sSql = ""
            
            sSql = "Update SGI_CADPEDVENDH Set SGI_STATUS = 'M'" & vbCrLf
            sSql = sSql & "Where" & vbCrLf
            sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "  And SGI_CODIGO = " & arrDADOSNOVA(I, 5)
        
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
            
            
            '' =====================================================
            '' Gera Log daos Pedidos
            sSql = ""
            
            sSql = "Insert Into SGI_LOGPED (" & vbCrLf
            sSql = sSql & "                              SGI_FILIAL" & vbCrLf
            sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
            sSql = sSql & "                             ,SGI_DATA" & vbCrLf
            sSql = sSql & "                             ,SGI_HORA" & vbCrLf
            sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
            sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
            sSql = sSql & "                    ) Values (" & vbCrLf
            sSql = sSql & "                              " & intFilial & vbCrLf
            sSql = sSql & "                             ," & arrDADOSNOVA(I, 5) & vbCrLf
            sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
            sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
            sSql = sSql & "                             ," & lngCodUs & vbCrLf
            sSql = sSql & "                             ,'M'" & vbCrLf
            sSql = sSql & "                             )"
            
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
            
            
            
            '' =====================================================
            '' Atualizando Programação de Entrega
            sSql = ""
            
            sSql = "Update SGI_PROGENTRPROD Set SGI_STATUS = 9" & vbCrLf
            sSql = sSql & "Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL    = " & intFilial & vbCrLf
            sSql = sSql & "  And SGI_CODPED    = " & arrDADOSNOVA(I, 5) & vbCrLf
            sSql = sSql & "  And SGI_IDPRODUTO = " & arrDADOSNOVA(I, 4) & vbCrLf
            If Len(Trim(arrDADOSNOVA(I, 3))) > 0 Then sSql = sSql & "  And SGI_IDINTERNO = " & arrDADOSNOVA(I, 3)
            
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
            
        Next I
    End If
    
    '' Grava Dados Steel
    If IsArray(arrDADOSSTEEL) Then
        For I = 1 To UBound(arrDADOSSTEEL)
        
            '' =====================================================
            '' Dando Baixa na OP
            sSql = ""
            
            sSql = "Update SGI_ORDEMPROD_STEEL Set" & vbCrLf
            sSql = sSql & "                                    SGI_STATUS      = 9" & vbCrLf
            sSql = sSql & "                                  , SGI_OBSLIQ      = " & strOBSLIQ & vbCrLf
            sSql = sSql & "                                  , SGI_CODMOTLIQOP = " & strCODMOTLIQOP & vbCrLf
            sSql = sSql & "Where" & vbCrLf
            sSql = sSql & "      SGI_FILIAL    = " & intFilial & vbCrLf
            sSql = sSql & "  And SGI_CODIGO    = " & arrDADOSSTEEL(I, 2) & vbCrLf
            sSql = sSql & "  And SGI_IDPRODUTO = " & arrDADOSSTEEL(I, 4) & vbCrLf
            If Len(Trim(arrDADOSSTEEL(I, 3))) > 0 Then sSql = sSql & "  And SGI_IDPAI  = " & arrDADOSSTEEL(I, 3)
        
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
                            
            
            '' =====================================================
            '' Gera Log das OP
            sSql = ""
            sSql = "Insert Into SGI_LOGFORM (" & vbCrLf
            sSql = sSql & "                              SGI_FILIAL" & vbCrLf
            sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
            sSql = sSql & "                             ,SGI_DATA" & vbCrLf
            sSql = sSql & "                             ,SGI_HORA" & vbCrLf
            sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
            sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
            sSql = sSql & "                             ,SGI_FORM" & vbCrLf
            sSql = sSql & "                    ) Values (" & vbCrLf
            sSql = sSql & "                              " & intFilial & vbCrLf
            sSql = sSql & "                             ," & arrDADOSSTEEL(I, 2) & vbCrLf
            sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
            sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
            sSql = sSql & "                             ," & lngCodUs & vbCrLf
            sSql = sSql & "                             ,'BX'" & vbCrLf
            sSql = sSql & "                             ,'frmCADOORDFAB_STEEL'" & vbCrLf
            sSql = sSql & "                             )"
            
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
            
            
            '' =====================================================
            '' Dando Baixa no Pedido
            sSql = ""
            
            sSql = "Update SGI_CADPEDVENDH_STEEL Set SGI_STATUS = 'M'" & vbCrLf
            sSql = sSql & "Where" & vbCrLf
            sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "  And SGI_CODIGO = " & arrDADOSSTEEL(I, 5)
        
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
            
            '' =====================================================
            '' Gera Log daos Pedidos
            sSql = ""
            
            sSql = "Insert Into SGI_LOGPED_STEEL (" & vbCrLf
            sSql = sSql & "                              SGI_FILIAL" & vbCrLf
            sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
            sSql = sSql & "                             ,SGI_DATA" & vbCrLf
            sSql = sSql & "                             ,SGI_HORA" & vbCrLf
            sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
            sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
            sSql = sSql & "                    ) Values (" & vbCrLf
            sSql = sSql & "                              " & intFilial & vbCrLf
            sSql = sSql & "                             ," & arrDADOSSTEEL(I, 5) & vbCrLf
            sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
            sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
            sSql = sSql & "                             ," & lngCodUs & vbCrLf
            sSql = sSql & "                             ,'M'" & vbCrLf
            sSql = sSql & "                             )"
            
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
            
            
            '' =====================================================
            '' Atualizando Programação de Entrega
            sSql = ""
            
            sSql = "Update SGI_PROGENTRPROD_STEEL Set SGI_STATUS = 9" & vbCrLf
            sSql = sSql & "Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL    = " & intFilial & vbCrLf
            sSql = sSql & "  And SGI_CODPED    = " & arrDADOSSTEEL(I, 5) & vbCrLf
            sSql = sSql & "  And SGI_IDPRODUTO = " & arrDADOSSTEEL(I, 4) & vbCrLf
            If Len(Trim(arrDADOSSTEEL(I, 3))) > 0 Then sSql = sSql & "  And SGI_IDINTERNO = " & arrDADOSSTEEL(I, 3)
            
            BGRV.CommandText = sSql
            BGRV.Execute
            '' =====================================================
            
        Next I
    End If
    
    adoBanco_Dados.CommitTrans
    GravaDados = True
    
    Exit Function
    
Err_ProcessaDados:

    adoBanco_Dados.RollbackTrans

    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "G", sSql)
    Set objErro = Nothing


End Function


