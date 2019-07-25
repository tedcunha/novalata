VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADVENDEDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial       As Integer
Dim intVENDCODIGO   As Integer
Dim strVENDDESCRI   As String
Dim strEMAIL        As String
Dim varNIVELCOMIS   As Variant
Dim varMETAS        As Variant
Dim varCLIENTES     As Variant
Dim strNOMEARQ      As String
Dim strCAMINHO      As String
Dim varCONTARQ      As Variant
Dim strSKYPE        As String
Dim strMSN          As String
Dim lngCodUsuario   As Long
Dim arrTIPOORCA     As Variant
Dim strNOMCOMP      As String
Dim intATIVO        As Integer
Dim arrVENDEDORES   As Variant

Public Property Get VENDEDORES() As Variant
    VENDEDORES = arrVENDEDORES
End Property

Public Property Let VENDEDORES(ByVal vData As Variant)
    arrVENDEDORES = vData
End Property

Public Property Get ATIVO() As Integer
    ATIVO = intATIVO
End Property

Public Property Let ATIVO(ByVal vData As Integer)
    intATIVO = vData
End Property

Public Property Get NOMCOMP() As String
    NOMCOMP = strNOMCOMP
End Property

Public Property Let NOMCOMP(ByVal vData As String)
    strNOMCOMP = vData
End Property

Public Property Get TIPOORCA() As Variant
    TIPOORCA = arrTIPOORCA
End Property

Public Property Let TIPOORCA(ByVal vData As Variant)
    arrTIPOORCA = vData
End Property

Public Property Get CodUsuario() As Long
    CodUsuario = lngCodUsuario
End Property

Public Property Let CodUsuario(ByVal vData As Long)
    lngCodUsuario = vData
End Property

Public Property Get SKYPE() As String
    SKYPE = strSKYPE
End Property

Public Property Let SKYPE(ByVal vData As String)
    strSKYPE = vData
End Property

Public Property Get MSN() As String
    MSN = strMSN
End Property

Public Property Let MSN(ByVal vData As String)
    strMSN = vData
End Property

Public Property Get CONTARQ() As Variant
    CONTARQ = varCONTARQ
End Property

Public Property Let CONTARQ(ByVal vData As Variant)
    varCONTARQ = vData
End Property

Public Property Get CAMINHO() As String
    CAMINHO = strCAMINHO
End Property

Public Property Let CAMINHO(ByVal vData As String)
    strCAMINHO = vData
End Property

Public Property Get NOMEARQ() As String
    NOMEARQ = strNOMEARQ
End Property

Public Property Let NOMEARQ(ByVal vData As String)
    strNOMEARQ = vData
End Property

Public Property Get EMAIL() As String
    EMAIL = strEMAIL
End Property

Public Property Let EMAIL(ByVal vData As String)
    strEMAIL = vData
End Property

Public Property Get CLIENTES() As Variant
    CLIENTES = varCLIENTES
End Property

Public Property Let CLIENTES(ByVal vData As Variant)
    varCLIENTES = vData
End Property

Public Property Get METAS() As Variant
    METAS = varMETAS
End Property

Public Property Let METAS(ByVal vData As Variant)
    varMETAS = vData
End Property

Public Property Get NIVELCOMIS() As Variant
    NIVELCOMIS = varNIVELCOMIS
End Property

Public Property Let NIVELCOMIS(ByVal vData As Variant)
    varNIVELCOMIS = vData
End Property

Public Property Get VENDDESCRI() As String
    VENDDESCRI = strVENDDESCRI
End Property

Public Property Let VENDDESCRI(ByVal vData As String)
    strVENDDESCRI = vData
End Property

Public Property Get VENDCODIGO() As Integer
    VENDCODIGO = intVENDCODIGO
End Property

Public Property Let VENDCODIGO(ByVal vData As Integer)
    intVENDCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADVENDEDORP.cCaminho = cCaminho
    frmCADVENDEDORP.Linha = Linha
    frmCADVENDEDORP.FILIAL = iFilial
    frmCADVENDEDORP.strAcesso = strAcesso
    frmCADVENDEDORP.strUSUARIO = strUSUARIO
    frmCADVENDEDORP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADVENDEDOR.cCaminho = cCaminho
    frmCADVENDEDOR.Linha = Linha
    frmCADVENDEDOR.FILIAL = iFilial
    frmCADVENDEDOR.strAcesso = strAcesso
    frmCADVENDEDOR.strUSUARIO = strUSUARIO
    frmCADVENDEDOR.cTipOper = "I"
    frmCADVENDEDOR.Show vbModal
End Sub

Public Function Pesq_CadVendedor() As Boolean

   Pesq_CadVendedor = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADVENDEDOR " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadVendedor = True
   BREC.Close

End Function

Public Sub PreenchComboCliente(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_RAZAOSOC)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
    BREC.Close

End Sub


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim I        As Integer
     Dim strVALOR As String
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava Vendedor
        sSql = "Insert Into SGI_CADVENDEDOR (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO,SGI_EMAIL,SGI_NOMEARQ,SGI_CAMINHOARQ,SGI_SKYPE,SGI_MSN,SGI_CODUSUARIO,SGI_NOMCOMP,SGI_ATIVO) Values (" & vbCrLf
        sSql = sSql & "                             " & intFilial & vbCrLf
        sSql = sSql & "                           , " & intVENDCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & strVENDDESCRI & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMAIL & "'" & vbCrLf
        
        If Len(Trim(strNOMEARQ)) > 0 Then
           sSql = sSql & "                           ,'" & strNOMEARQ & "'" & vbCrLf
        Else
           sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        If Len(Trim(strCAMINHO)) > 0 Then
           sSql = sSql & "                           ,'" & strCAMINHO & "'" & vbCrLf
        Else
           sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                           ,'" & strSKYPE & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strMSN & "'" & vbCrLf
        
        If lngCodUsuario > 0 Then
            sSql = sSql & "                           ," & lngCodUsuario & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                           ," & strNOMCOMP & vbCrLf
        sSql = sSql & "                           ," & intATIVO & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Gravando Niveis de desconto
        If IsArray(varNIVELCOMIS) = True Then
           
           For I = 1 To UBound(varNIVELCOMIS)
            
               sSql = "Insert Into SGI_CADNIVDESVEND (SGI_FILIAL,SGI_CODIGO,SGI_DESC1,SGI_DESC2,SGI_COMISS) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intVENDCODIGO & vbCrLf
               
               strVALOR = Replace(varNIVELCOMIS(I, 1), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               strVALOR = Replace(varNIVELCOMIS(I, 2), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               strVALOR = Replace(varNIVELCOMIS(I, 3), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
     
        '' Gravando Metas
        If IsArray(varMETAS) = True Then
           
           For I = 1 To UBound(varMETAS)
            
               sSql = "Insert Into SGI_CADMETASVEND (SGI_FILIAL,SGI_CODIGO,SGI_VALO1,SGI_VALO2,SGI_COMISS) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intVENDCODIGO & vbCrLf
               
               strVALOR = Replace(varMETAS(I, 1), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               strVALOR = Replace(varMETAS(I, 2), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               strVALOR = Replace(varMETAS(I, 3), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
     
        '' Gravando Cliente
        If IsArray(varCLIENTES) = True Then
           
           For I = 1 To UBound(varCLIENTES)
            
               sSql = "Insert Into SGI_CADCLIEVEND (SGI_FILIAL,SGI_CODIGO,SGI_CODCLI) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intVENDCODIGO & vbCrLf
               sSql = sSql & "                           ," & varCLIENTES(I) & vbCrLf
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
        
        
        '' Gravando Tipos de Orçamentos
        If IsArray(arrTIPOORCA) Then
            For I = 1 To UBound(arrTIPOORCA)
            
                sSql = "Insert Into SGI_VENDTIPORCA (SGI_FILIAL,SGI_CODVEND,SGI_CODTIPORCA) Values (" & vbCrLf
                sSql = sSql & "                             " & intFilial & vbCrLf
                sSql = sSql & "                            ," & intVENDCODIGO & vbCrLf
                sSql = sSql & "                            ," & arrTIPOORCA(I) & vbCrLf
                sSql = sSql & "                            )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
     
        '' Gravando Vendedores Para Vendedores
        If IsArray(arrVENDEDORES) Then
            For I = 1 To UBound(arrVENDEDORES)
            
                sSql = "Insert Into SGI_VENDTOVEND (" & vbCrLf
                sSql = sSql & "                     SGI_FILIAL" & vbCrLf
                sSql = sSql & "                    ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                    ,SGI_CODVEND" & vbCrLf
                
                sSql = sSql & "           ) Values (" & vbCrLf
                sSql = sSql & "                     " & intFilial & vbCrLf
                sSql = sSql & "                    ," & intVENDCODIGO & vbCrLf
                sSql = sSql & "                    ," & arrVENDEDORES(I) & vbCrLf
                
                sSql = sSql & "                            )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
     
     ElseIf strOPER = "A" Then
     
        '' Apagando Niveis de Desconto
        sSql = "Delete from SGI_CADNIVDESVEND " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Apagando Metas
        sSql = "Delete from SGI_CADMETASVEND " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Apagando Clientes
        sSql = "Delete from SGI_CADCLIEVEND " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Apagando Tipos de Orcamentos
        sSql = "Delete from SGI_VENDTIPORCA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODVEND = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apagando Vendedores Para Vendedores
        sSql = "Delete from SGI_VENDTOVEND" & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Alterando Vendedor
        sSql = "Update SGI_CADVENDEDOR Set " & vbCrLf
        sSql = sSql & "                    SGI_DESCRICAO = '" & strVENDDESCRI & "'" & vbCrLf
        sSql = sSql & "                   ,SGI_EMAIL     = '" & strEMAIL & "'" & vbCrLf
        sSql = sSql & "                   ,SGI_NOMEARQ    = '" & strNOMEARQ & "'" & vbCrLf
        sSql = sSql & "                   ,SGI_CAMINHOARQ = '" & strCAMINHO & "'" & vbCrLf
        sSql = sSql & "                   ,SGI_SKYPE      = '" & strSKYPE & "'" & vbCrLf
        sSql = sSql & "                   ,SGI_MSN        = '" & strMSN & "'" & vbCrLf
        
        If lngCodUsuario > 0 Then
            sSql = sSql & "                   ,SGI_CODUSUARIO = " & lngCodUsuario & vbCrLf
        Else
            sSql = sSql & "                   ,SGI_CODUSUARIO = Null" & vbCrLf
        End If
        
        sSql = sSql & "                   ,SGI_NOMCOMP    = " & strNOMCOMP & vbCrLf
        sSql = sSql & "                   ,SGI_ATIVO      = " & intATIVO & vbCrLf
        
        sSql = sSql & "          Where "
        sSql = sSql & "                SGI_FILIAL = " & intFilial
        sSql = sSql & "            And SGI_CODIGO = " & intVENDCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Gravando Niveis de desconto
        If IsArray(varNIVELCOMIS) = True Then
           
           For I = 1 To UBound(varNIVELCOMIS)
            
               sSql = "Insert Into SGI_CADNIVDESVEND (SGI_FILIAL,SGI_CODIGO,SGI_DESC1,SGI_DESC2,SGI_COMISS) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intVENDCODIGO & vbCrLf
               
               strVALOR = Replace(varNIVELCOMIS(I, 1), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               strVALOR = Replace(varNIVELCOMIS(I, 2), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               strVALOR = Replace(varNIVELCOMIS(I, 3), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
        
        '' Gravando Metas
        If IsArray(varMETAS) = True Then
           
           For I = 1 To UBound(varMETAS)
            
               sSql = "Insert Into SGI_CADMETASVEND (SGI_FILIAL,SGI_CODIGO,SGI_VALO1,SGI_VALO2,SGI_COMISS) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intVENDCODIGO & vbCrLf
               
               strVALOR = Replace(varMETAS(I, 1), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               strVALOR = Replace(varMETAS(I, 2), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               strVALOR = Replace(varMETAS(I, 3), ".", "")
               strVALOR = Replace(strVALOR, ",", ".")
               sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
               
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
        
        '' Gravando Cliente
        If IsArray(varCLIENTES) = True Then
           
           For I = 1 To UBound(varCLIENTES)
            
               sSql = "Insert Into SGI_CADCLIEVEND (SGI_FILIAL,SGI_CODIGO,SGI_CODCLI) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intVENDCODIGO & vbCrLf
               sSql = sSql & "                           ," & varCLIENTES(I) & vbCrLf
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
     
        '' Gravando Tipos de Orçamentos
        If IsArray(arrTIPOORCA) Then
            For I = 1 To UBound(arrTIPOORCA)
            
                sSql = "Insert Into SGI_VENDTIPORCA (SGI_FILIAL,SGI_CODVEND,SGI_CODTIPORCA) Values (" & vbCrLf
                sSql = sSql & "                             " & intFilial & vbCrLf
                sSql = sSql & "                            ," & intVENDCODIGO & vbCrLf
                sSql = sSql & "                            ," & arrTIPOORCA(I) & vbCrLf
                sSql = sSql & "                            )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
     
        '' Gravando Vendedores Para Vendedores
        If IsArray(arrVENDEDORES) Then
            For I = 1 To UBound(arrVENDEDORES)
            
                sSql = "Insert Into SGI_VENDTOVEND (" & vbCrLf
                sSql = sSql & "                     SGI_FILIAL" & vbCrLf
                sSql = sSql & "                    ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                    ,SGI_CODVEND" & vbCrLf
                
                sSql = sSql & "           ) Values (" & vbCrLf
                sSql = sSql & "                     " & intFilial & vbCrLf
                sSql = sSql & "                    ," & intVENDCODIGO & vbCrLf
                sSql = sSql & "                    ," & arrVENDEDORES(I) & vbCrLf
                
                sSql = sSql & "                            )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
     
     ElseIf strOPER = "E" Then
     
        '' Apagando Niveis de Desconto
        sSql = "Delete from SGI_CADNIVDESVEND " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Apagando Metas
        sSql = "Delete from SGI_CADMETASVEND " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Apagando Clientes
        sSql = "Delete from SGI_CADCLIEVEND " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apagando Tipos de Orcamentos
        sSql = "Delete from SGI_VENDTIPORCA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODVEND = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apagando Vendedores para vendedores
        sSql = "Delete from SGI_VENDTOVEND " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODVEND = " & intVENDCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apagando o vendedor
        sSql = "Delete from SGI_CADVENDEDOR " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intVENDCODIGO
     
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


Public Function Pesq_Vendedor() As Boolean

   Pesq_Vendedor = False
   
   Dim intQtdReg As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
   sSql = sSql & " Where "
   sSql = sSql & "       SGI_FILIAL  = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intVENDCODIGO
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intVENDCODIGO = BREC!SGI_CODIGO
      strVENDDESCRI = BREC!SGI_DESCRICAO
      If Not IsNull(BREC!SGI_EMAIL) Then strEMAIL = BREC!SGI_EMAIL
      If Not IsNull(BREC!SGI_SKYPE) Then strSKYPE = BREC!SGI_SKYPE
      If Not IsNull(BREC!SGI_MSN) Then strMSN = BREC!SGI_MSN
      
      If Not IsNull(BREC!SGI_NOMEARQ) Then strNOMEARQ = Trim(BREC!SGI_NOMEARQ)
      If Not IsNull(BREC!SGI_CAMINHOARQ) Then strCAMINHO = Trim(BREC!SGI_CAMINHOARQ)
      If Not IsNull(BREC!SGI_CODUSUARIO) Then lngCodUsuario = BREC!SGI_CODUSUARIO
      If Not IsNull(BREC!SGI_NOMCOMP) Then strNOMCOMP = BREC!SGI_NOMCOMP
      If Not IsNull(BREC!SGI_ATIVO) Then intATIVO = BREC!SGI_ATIVO
   
      Pesq_Vendedor = True
   End If
   
   BREC.Close
   
   If Pesq_Vendedor = False Then Exit Function
   
   '' Pesquisa Niveis de Desconto
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADNIVDESVEND " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO = " & intVENDCODIGO
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
      
      intQtdReg = 0
      Do While Not BREC2.EOF
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim varNIVELCOMIS(1 To intQtdReg, 1 To 3) As String
      
      intQtdReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
      
         varNIVELCOMIS(intQtdReg, 1) = Format(BREC2!SGI_DESC1, "#,##0.00")
         varNIVELCOMIS(intQtdReg, 2) = Format(BREC2!SGI_DESC2, "#,##0.00")
         varNIVELCOMIS(intQtdReg, 3) = Format(BREC2!SGI_COMISS, "#,##0.00")
         
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

   '' Pesquisa Metas
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMETASVEND " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO = " & intVENDCODIGO
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
      
      intQtdReg = 0
      Do While Not BREC2.EOF
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim varMETAS(1 To intQtdReg, 1 To 3) As String
      
      intQtdReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
      
         varMETAS(intQtdReg, 1) = Format(BREC2!SGI_VALO1, "#,##0.00")
         varMETAS(intQtdReg, 2) = Format(BREC2!SGI_VALO2, "#,##0.00")
         varMETAS(intQtdReg, 3) = Format(BREC2!SGI_COMISS, "#,##0.00")
         
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' Pesquisa Cliente
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADCLIEVEND " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO = " & intVENDCODIGO
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
      
      intQtdReg = 0
      Do While Not BREC2.EOF
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim varCLIENTES(1 To intQtdReg) As String
      
      intQtdReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
      
         varCLIENTES(intQtdReg) = Str(BREC2!SGI_CODCLI)
         
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   
   '' Pesquisa Cliente
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_VENDTIPORCA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL     = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODVEND    = " & intVENDCODIGO
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
      
      intQtdReg = 0
      Do While Not BREC2.EOF
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrTIPOORCA(1 To intQtdReg) As String
      
      intQtdReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
      
         arrTIPOORCA(intQtdReg) = Str(BREC2!SGI_CODTIPORCA)
         
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

   '' Pesquisa vendedor para vendedor
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_VENDTOVEND " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL     = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO     = " & intVENDCODIGO
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
      
      intQtdReg = 0
      Do While Not BREC2.EOF
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrVENDEDORES(1 To intQtdReg) As String
      
      intQtdReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
      
         arrVENDEDORES(intQtdReg) = Trim(Str(BREC2!SGI_CODVEND))
         
         intQtdReg = intQtdReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

End Function


Public Sub PreenchComboUsuario(Combo As Variant)

    Combo.Clear
    
    Dim objFunc As Object
    Set objFunc = CreateObject("BLBCWS.clsFuncoes")
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_USUARIO " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(objFunc.Crypt(BREC!SGI_NOME))
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close
    
    Combo.AddItem "VÁZIO"
    Combo.ItemData(Combo.NewIndex) = 0
    
    
    Set objFunc = Nothing

End Sub

