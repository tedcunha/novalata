VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADMAQUINA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim intCODIGO           As Integer
Dim strDESCRI           As String
Dim intOPTATIVO         As Integer
Dim strANOFABR          As String
Dim strNUMATIVO         As String
Dim curCAPMINUTO        As Currency
Dim lngUNIMED           As Long
Dim lngFAMILIA          As Long
Dim arrGRDFONTALIM      As Variant
Dim arrGRDTURNOS        As Variant
Dim arrGRDPRODUTOS      As Variant
Dim arrGRDOPERADOR      As Variant
Dim intQtdOperadores    As Integer
Dim curFraOperadores    As Currency
Dim arrPARAMETROS       As Variant
Dim arrDIASPARAMETROS   As Variant
Dim arrLINHACAP         As Variant

Public Property Get LINHACAP() As Variant
    LINHACAP = arrLINHACAP
End Property
Public Property Let LINHACAP(ByVal vData As Variant)
    arrLINHACAP = vData
End Property

Public Property Get DIASPARAMETROS() As Variant
    DIASPARAMETROS = arrDIASPARAMETROS
End Property
Public Property Let DIASPARAMETROS(ByVal vData As Variant)
    arrDIASPARAMETROS = vData
End Property

Public Property Get PARAMETROS() As Variant
    PARAMETROS = arrPARAMETROS
End Property
Public Property Let PARAMETROS(ByVal vData As Variant)
    arrPARAMETROS = vData
End Property

Public Property Get FraOperadores() As Currency
    FraOperadores = curFraOperadores
End Property
Public Property Let FraOperadores(ByVal vData As Currency)
    curFraOperadores = vData
End Property

Public Property Get QtdOperadores() As Integer
    QtdOperadores = intQtdOperadores
End Property
Public Property Let QtdOperadores(ByVal vData As Integer)
    intQtdOperadores = vData
End Property

Public Property Get FAMILIA() As Long
    FAMILIA = lngFAMILIA
End Property
Public Property Let FAMILIA(ByVal vData As Long)
    lngFAMILIA = vData
End Property

Public Property Get GRDFONTALIM() As Variant
    GRDFONTALIM = arrGRDFONTALIM
End Property
Public Property Let GRDFONTALIM(ByVal vData As Variant)
    arrGRDFONTALIM = vData
End Property

Public Property Get GRDTURNOS() As Variant
    GRDTURNOS = arrGRDTURNOS
End Property
Public Property Let GRDTURNOS(ByVal vData As Variant)
    arrGRDTURNOS = vData
End Property

Public Property Get GRDPRODUTOS() As Variant
    GRDPRODUTOS = arrGRDPRODUTOS
End Property
Public Property Let GRDPRODUTOS(ByVal vData As Variant)
    arrGRDPRODUTOS = vData
End Property

Public Property Get GRDOPERADOR() As Variant
    GRDOPERADOR = arrGRDOPERADOR
End Property
Public Property Let GRDOPERADOR(ByVal vData As Variant)
    arrGRDOPERADOR = vData
End Property

Public Property Get UNIMED() As Long
    UNIMED = lngUNIMED
End Property
Public Property Let UNIMED(ByVal vData As Long)
    lngUNIMED = vData
End Property

Public Property Get CAPMINUTO() As Currency
    CAPMINUTO = curCAPMINUTO
End Property
Public Property Let CAPMINUTO(ByVal vData As Currency)
    curCAPMINUTO = vData
End Property

Public Property Get NUMATIVO() As String
    NUMATIVO = strNUMATIVO
End Property
Public Property Let NUMATIVO(ByVal vData As String)
    strNUMATIVO = vData
End Property

Public Property Get ANOFABR() As String
    ANOFABR = strANOFABR
End Property
Public Property Let ANOFABR(ByVal vData As String)
    strANOFABR = vData
End Property

Public Property Get OPTATIVO() As Integer
    OPTATIVO = intOPTATIVO
End Property
Public Property Let OPTATIVO(ByVal vData As Integer)
    intOPTATIVO = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property
Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property

Public Property Get CODIGO() As Integer
    CODIGO = intCODIGO
End Property
Public Property Let CODIGO(ByVal vData As Integer)
    intCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property
Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADMAQUINAP.cCaminho = cCaminho
    frmCADMAQUINAP.Linha = Linha
    frmCADMAQUINAP.FILIAL = iFilial
    frmCADMAQUINAP.strAcesso = strAcesso
    frmCADMAQUINAP.Show vbModal
End Sub

Public Function Pesq_CadMaquina() As Boolean

   Pesq_CadMaquina = False
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       SGI_CODIGO  " & vbCrLf
   sSql = sSql + "      ,SGI_DESCRI  " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADMAQUINA " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadMaquina = True
   BREC.Close

End Function

Public Sub PreencheComboTipAlim(Combo As Variant)

     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADTIPALIM " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRI)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboTurnos(Combo As Variant)

     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRI)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboProd(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL      = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreenchComboUnidade(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_UNIDADE)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub PreenchComboFamilia(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFAMMAQUINAS " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub


Public Sub PreencheComboOperadores(Combo As Variant)

     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADOPERADOR " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRI)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Function Gera_Codigo(sModulo As String) As Long

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


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
     
     Dim I      As Integer
     Dim sValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
        
        '' Grava especificação técnica do produto
        sSql = "Insert Into SGI_CADMAQUINA (SGI_FILIAL,SGI_CODIGO,SGI_DESCRI,SGI_ATIVA,SGI_ANOFABR,SGI_NATIVO,SGI_CAPMINU,SGI_CODUNIMED,SGI_CODFAMILIA,SGI_QTDOPERADORES,SGI_FRAOPERADORES) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & intCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & strDESCRI & "'" & vbCrLf
        sSql = sSql & "                           ," & intOPTATIVO & vbCrLf
        sSql = sSql & "                           ,'" & strANOFABR & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strNUMATIVO & "'" & vbCrLf
        
        sValor = Replace(curCAPMINUTO, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                           ," & sValor & vbCrLf
        
        sSql = sSql & "                           ," & lngUNIMED & vbCrLf
        sSql = sSql & "                           ," & lngFAMILIA & vbCrLf
        
        sSql = sSql & "                           ," & intQtdOperadores & vbCrLf
        
        sValor = Replace(curFraOperadores, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                           ," & sValor & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute

        '' -- Fonte de Alimentação
        If IsArray(arrGRDFONTALIM) Then
           For I = 1 To UBound(arrGRDFONTALIM)
               
               sSql = "Insert Into SGI_CADMAQFALIM (SGI_FILIAL,SGI_CODIGO,SGI_CODTIPAL) Values (" & vbCrLf
               sSql = sSql & "                     " & intFilial & vbCrLf
               sSql = sSql & "                    ," & intCODIGO & vbCrLf
               sSql = sSql & "                    ," & arrGRDFONTALIM(I) & vbCrLf
               sSql = sSql & "                    )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
        
        '' -- Turnos
        If IsArray(arrGRDTURNOS) Then
           For I = 1 To UBound(arrGRDTURNOS)
               
               sSql = "Insert Into SGI_CADMAQTURN (SGI_FILIAL,SGI_CODIGO,SGI_CODTURN,SGI_PORCUSO) Values (" & vbCrLf
               sSql = sSql & "                     " & intFilial & vbCrLf
               sSql = sSql & "                    ," & intCODIGO & vbCrLf
               sSql = sSql & "                    ," & arrGRDTURNOS(I, 1) & vbCrLf
               
               sValor = Replace(arrGRDTURNOS(I, 2), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               sSql = sSql & "                    ," & sValor & vbCrLf
               
               sSql = sSql & "                    )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
        
        '' -- Produtos
        If IsArray(arrGRDPRODUTOS) Then
           For I = 1 To UBound(arrGRDPRODUTOS)
               
               sSql = "Insert Into SGI_CADMAQPROD (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                    ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                    ,SGI_CODPROD" & vbCrLf
               sSql = sSql & "                    ,SGI_QTDPCHORA" & vbCrLf
               sSql = sSql & "                    ,SGI_QTDPCMIN" & vbCrLf
               sSql = sSql & "                    ,SGI_CUSTOHORA" & vbCrLf
               sSql = sSql & "                    ,SGI_CUSTOMIN)" & vbCrLf
               sSql = sSql & "             Values (" & vbCrLf
               sSql = sSql & "                     " & intFilial & vbCrLf
               sSql = sSql & "                    ," & intCODIGO & vbCrLf
               sSql = sSql & "                    ,'" & arrGRDPRODUTOS(I, 1) & "'" & vbCrLf
               
               sValor = Replace(arrGRDPRODUTOS(I, 2), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               If Len(Trim(sValor)) > 0 Then
                  sSql = sSql & "                           ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sValor = Replace(arrGRDPRODUTOS(I, 3), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               If Len(Trim(sValor)) > 0 Then
                  sSql = sSql & "                           ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sValor = Replace(arrGRDPRODUTOS(I, 4), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               If Len(Trim(sValor)) > 0 Then
                  sSql = sSql & "                           ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sValor = Replace(arrGRDPRODUTOS(I, 5), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               If Len(Trim(sValor)) > 0 Then
                  sSql = sSql & "                           ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sSql = sSql & "                    )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
        
        '' -- Operadores
        If IsArray(arrGRDOPERADOR) Then
           For I = 1 To UBound(arrGRDOPERADOR)
               
               sSql = "Insert Into SGI_CADMAQOPER (SGI_FILIAL,SGI_CODIGO,SGI_CODOPER,SGI_PORCUTIL) Values (" & vbCrLf
               sSql = sSql & "                     " & intFilial & vbCrLf
               sSql = sSql & "                    ," & intCODIGO & vbCrLf
               sSql = sSql & "                    ," & arrGRDOPERADOR(I, 1) & vbCrLf
               
               sValor = Replace(arrGRDOPERADOR(I, 2), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               sSql = sSql & "                           ," & sValor & vbCrLf
               
               sSql = sSql & "                    )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
     
        '' -- Paradas para manutenção
        If IsArray(arrPARAMETROS) Then
           For I = 1 To UBound(arrPARAMETROS)
           
                sSql = "Insert Into SGI_CADPARAMMANUT (SGI_FILIAL,SGI_CODMAQUINA,SGI_CODPERIODO,SGI_DATAMANUT,SGI_DATAFINAL,SGI_HORAINI,SGI_HORAFIN,SGI_TOTALUSADO,SGI_MAQUSO,SGI_ATIVO,SGI_EMCONJUNTO) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & intCODIGO & vbCrLf
                sSql = sSql & "                             ," & arrPARAMETROS(I, 1) & vbCrLf
                sSql = sSql & "                             ,'" & Format(CDate(arrPARAMETROS(I, 2)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & Format(CDate(arrPARAMETROS(I, 3)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & arrPARAMETROS(I, 4) & ":00'" & vbCrLf
                sSql = sSql & "                             ,'" & arrPARAMETROS(I, 5) & ":00'" & vbCrLf
                sSql = sSql & "                             ,'" & arrPARAMETROS(I, 6) & ":00'" & vbCrLf
                sSql = sSql & "                             ," & arrPARAMETROS(I, 7) & vbCrLf
                sSql = sSql & "                             ," & arrPARAMETROS(I, 8) & vbCrLf
                sSql = sSql & "                             ," & arrPARAMETROS(I, 9) & vbCrLf
                sSql = sSql & ")"
           
                BGRV.CommandText = sSql
                BGRV.Execute
           
           Next I
        End If
        
        '' -- Dias da manutenção
        If IsArray(arrDIASPARAMETROS) Then
           For I = 1 To UBound(arrDIASPARAMETROS)
           
                sSql = "Insert Into SGI_CADPERIODOMANUT (SGI_FILIAL,SGI_CODMAQUINA,SGI_CODPERIODO,SGI_DATAPARAM,SGI_DATAMANUT,SGI_HORAINI,SGI_HORAFIN,SGI_TOTALUSADO,SGI_MAQUSO,SGI_ATIVO,SGI_EMCONJUNTO) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & intCODIGO & vbCrLf
                sSql = sSql & "                             ," & arrDIASPARAMETROS(I, 8) & vbCrLf
                sSql = sSql & "                             ,'" & Format(CDate(arrDIASPARAMETROS(I, 9)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & Format(CDate(arrDIASPARAMETROS(I, 1)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & arrDIASPARAMETROS(I, 2) & ":00'" & vbCrLf
                sSql = sSql & "                             ,'" & arrDIASPARAMETROS(I, 3) & ":00'" & vbCrLf
                sSql = sSql & "                             ,'" & arrDIASPARAMETROS(I, 4) & ":00'" & vbCrLf
                sSql = sSql & "                             ," & arrDIASPARAMETROS(I, 5) & vbCrLf
                sSql = sSql & "                             ," & arrDIASPARAMETROS(I, 6) & vbCrLf
                sSql = sSql & "                             ," & arrDIASPARAMETROS(I, 7) & vbCrLf
                sSql = sSql & ")"
           
                BGRV.CommandText = sSql
                BGRV.Execute
           
           Next I
        End If
        
        '' linha
        If IsArray(arrLINHACAP) Then
            For I = 1 To UBound(arrLINHACAP)
                sSql = ""
                
                sSql = "Insert Into SGI_LINHAMAQPROD Into (" & vbCrLf
                sSql = sSql & "                                    SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                   ,SGI_ITEN" & vbCrLf
                sSql = sSql & "                                   ,SGI_IDLIN" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODLIN" & vbCrLf
                sSql = sSql & "                                   ,SGI_SEQCORT" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODCORT" & vbCrLf
                sSql = sSql & "                                   ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                         ) Values (" & vbCrLf
                sSql = sSql & "                                   " & intFilial & vbCrLf
                sSql = sSql & "                                  ," & intCODIGO & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 1) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 2) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 3) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 4) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 5) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 6) & vbCrLf
                sSql = sSql & "                                  )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next I
        End If
     
     ElseIf strOPER = "A" Then
     
        '' Operadores
        sSql = "Delete from SGI_CADMAQOPER " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Produtos
        sSql = "Delete from SGI_CADMAQPROD " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Turnos
        sSql = "Delete from SGI_CADMAQTURN " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Tipo de Alimentação
        sSql = "Delete from SGI_CADMAQFALIM " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Parâmetros para manutenção
        sSql = "Delete from SGI_CADPERIODOMANUT " & vbCrLf
        sSql = sSql & "      Where SGI_CODMAQUINA = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL     = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Delete from SGI_CADPARAMMANUT " & vbCrLf
        sSql = sSql & "      Where SGI_CODMAQUINA = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL     = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        sSql = "Delete from SGI_LINHAMAQPROD " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
     
        '' -- Fonte de Alimentação
        If IsArray(arrGRDFONTALIM) Then
           For I = 1 To UBound(arrGRDFONTALIM)
               
               sSql = "Insert Into SGI_CADMAQFALIM (SGI_FILIAL,SGI_CODIGO,SGI_CODTIPAL) Values (" & vbCrLf
               sSql = sSql & "                     " & intFilial & vbCrLf
               sSql = sSql & "                    ," & intCODIGO & vbCrLf
               sSql = sSql & "                    ," & arrGRDFONTALIM(I) & vbCrLf
               sSql = sSql & "                    )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
        
        '' -- Turnos
        If IsArray(arrGRDTURNOS) Then
           For I = 1 To UBound(arrGRDTURNOS)
               
               sSql = "Insert Into SGI_CADMAQTURN (SGI_FILIAL,SGI_CODIGO,SGI_CODTURN,SGI_PORCUSO) Values (" & vbCrLf
               sSql = sSql & "                     " & intFilial & vbCrLf
               sSql = sSql & "                    ," & intCODIGO & vbCrLf
               sSql = sSql & "                    ," & arrGRDTURNOS(I, 1) & vbCrLf
               
               sValor = Replace(arrGRDTURNOS(I, 2), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               sSql = sSql & "                    ," & sValor & vbCrLf
               
               sSql = sSql & "                    )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
        
        '' -- Produtos
        If IsArray(arrGRDPRODUTOS) Then
           For I = 1 To UBound(arrGRDPRODUTOS)
               
               sSql = "Insert Into SGI_CADMAQPROD (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                    ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                    ,SGI_CODPROD" & vbCrLf
               sSql = sSql & "                    ,SGI_QTDPCHORA" & vbCrLf
               sSql = sSql & "                    ,SGI_QTDPCMIN" & vbCrLf
               sSql = sSql & "                    ,SGI_CUSTOHORA" & vbCrLf
               sSql = sSql & "                    ,SGI_CUSTOMIN)" & vbCrLf
               sSql = sSql & "             Values (" & vbCrLf
               sSql = sSql & "                     " & intFilial & vbCrLf
               sSql = sSql & "                    ," & intCODIGO & vbCrLf
               sSql = sSql & "                    ,'" & arrGRDPRODUTOS(I, 1) & "'" & vbCrLf
               
               sValor = Replace(arrGRDPRODUTOS(I, 2), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               If Len(Trim(sValor)) > 0 Then
                  sSql = sSql & "                           ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sValor = Replace(arrGRDPRODUTOS(I, 3), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               If Len(Trim(sValor)) > 0 Then
                  sSql = sSql & "                           ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sValor = Replace(arrGRDPRODUTOS(I, 4), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               If Len(Trim(sValor)) > 0 Then
                  sSql = sSql & "                           ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sValor = Replace(arrGRDPRODUTOS(I, 5), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               If Len(Trim(sValor)) > 0 Then
                  sSql = sSql & "                           ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sSql = sSql & "                    )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
        
        '' -- Operadores
        If IsArray(arrGRDOPERADOR) Then
           For I = 1 To UBound(arrGRDOPERADOR)
               
               sSql = "Insert Into SGI_CADMAQOPER (SGI_FILIAL,SGI_CODIGO,SGI_CODOPER,SGI_PORCUTIL) Values (" & vbCrLf
               sSql = sSql & "                     " & intFilial & vbCrLf
               sSql = sSql & "                    ," & intCODIGO & vbCrLf
               sSql = sSql & "                    ," & arrGRDOPERADOR(I, 1) & vbCrLf
               
               sValor = Replace(arrGRDOPERADOR(I, 2), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               sSql = sSql & "                           ," & sValor & vbCrLf
               
               sSql = sSql & "                    )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
     
        '' -- Paradas para manutenção
        If IsArray(arrPARAMETROS) Then
           For I = 1 To UBound(arrPARAMETROS)
           
                sSql = "Insert Into SGI_CADPARAMMANUT (SGI_FILIAL,SGI_CODMAQUINA,SGI_CODPERIODO,SGI_DATAMANUT,SGI_DATAFINAL,SGI_HORAINI,SGI_HORAFIN,SGI_TOTALUSADO,SGI_MAQUSO,SGI_ATIVO,SGI_EMCONJUNTO) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & intCODIGO & vbCrLf
                sSql = sSql & "                             ," & arrPARAMETROS(I, 1) & vbCrLf
                sSql = sSql & "                             ,'" & Format(CDate(arrPARAMETROS(I, 2)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & Format(CDate(arrPARAMETROS(I, 3)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & arrPARAMETROS(I, 4) & ":00'" & vbCrLf
                sSql = sSql & "                             ,'" & arrPARAMETROS(I, 5) & ":00'" & vbCrLf
                sSql = sSql & "                             ,'" & arrPARAMETROS(I, 6) & ":00'" & vbCrLf
                sSql = sSql & "                             ," & arrPARAMETROS(I, 7) & vbCrLf
                sSql = sSql & "                             ," & arrPARAMETROS(I, 8) & vbCrLf
                sSql = sSql & "                             ," & arrPARAMETROS(I, 9) & vbCrLf
                sSql = sSql & ")"
           
                BGRV.CommandText = sSql
                BGRV.Execute
           
           Next I
        End If
        
        '' -- Dias da manutenção
        If IsArray(arrDIASPARAMETROS) Then
           For I = 1 To UBound(arrDIASPARAMETROS)
           
                sSql = "Insert Into SGI_CADPERIODOMANUT (SGI_FILIAL,SGI_CODMAQUINA,SGI_CODPERIODO,SGI_DATAPARAM,SGI_DATAMANUT,SGI_HORAINI,SGI_HORAFIN,SGI_TOTALUSADO,SGI_MAQUSO,SGI_ATIVO,SGI_EMCONJUNTO) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & intCODIGO & vbCrLf
                sSql = sSql & "                             ," & arrDIASPARAMETROS(I, 8) & vbCrLf
                sSql = sSql & "                             ,'" & Format(CDate(arrDIASPARAMETROS(I, 9)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & Format(CDate(arrDIASPARAMETROS(I, 1)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & arrDIASPARAMETROS(I, 2) & ":00'" & vbCrLf
                sSql = sSql & "                             ,'" & arrDIASPARAMETROS(I, 3) & ":00'" & vbCrLf
                sSql = sSql & "                             ,'" & arrDIASPARAMETROS(I, 4) & ":00'" & vbCrLf
                sSql = sSql & "                             ," & arrDIASPARAMETROS(I, 5) & vbCrLf
                sSql = sSql & "                             ," & arrDIASPARAMETROS(I, 6) & vbCrLf
                sSql = sSql & "                             ," & arrDIASPARAMETROS(I, 7) & vbCrLf
                sSql = sSql & ")"
           
                BGRV.CommandText = sSql
                BGRV.Execute
           
           Next I
        End If
        
        
        '' linha
        If IsArray(arrLINHACAP) Then
            For I = 1 To UBound(arrLINHACAP)
                sSql = ""
                
                sSql = "Insert Into SGI_LINHAMAQPROD (" & vbCrLf
                sSql = sSql & "                                    SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                   ,SGI_ITEN" & vbCrLf
                sSql = sSql & "                                   ,SGI_IDLIN" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODLIN" & vbCrLf
                sSql = sSql & "                                   ,SGI_SEQCORT" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODCORT" & vbCrLf
                sSql = sSql & "                                   ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                         ) Values (" & vbCrLf
                sSql = sSql & "                                   " & intFilial & vbCrLf
                sSql = sSql & "                                  ," & intCODIGO & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 1) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 2) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 3) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 4) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 5) & vbCrLf
                sSql = sSql & "                                  ," & arrLINHACAP(I, 6) & vbCrLf
                sSql = sSql & "                                  )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next I
        End If
        
        sSql = "Update SGI_CADMAQUINA Set SGI_DESCRI = '" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_ATIVA  = " & intOPTATIVO & vbCrLf
        sSql = sSql & "                  ,SGI_NATIVO = '" & Trim(strNUMATIVO) & "'" & vbCrLf
        
        If lngFAMILIA > 0 Then
            sSql = sSql & "                  ,SGI_CODFAMILIA = " & lngFAMILIA & vbCrLf
        Else
            sSql = sSql & "                  ,SGI_CODFAMILIA = Null" & vbCrLf
        End If
        
        sValor = Replace(curCAPMINUTO, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                  ,SGI_CAPMINU =" & sValor & vbCrLf
        
        sSql = sSql & "                Where " & vbCrLf
        sSql = sSql & "                      SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "                  And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "E" Then
     
        '' Operadores
        sSql = "Delete from SGI_CADMAQOPER " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Produtos
        sSql = "Delete from SGI_CADMAQPROD " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Turnos
        sSql = "Delete from SGI_CADMAQTURN " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Tipo de Alimentação
        sSql = "Delete from SGI_CADMAQFALIM " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Parâmetros para manutenção
        sSql = "Delete from SGI_CADPERIODOMANUT " & vbCrLf
        sSql = sSql & "      Where SGI_CODMAQUINA = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL     = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Delete from SGI_CADPARAMMANUT " & vbCrLf
        sSql = sSql & "      Where SGI_CODMAQUINA = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL     = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Máquina
        sSql = "Delete from SGI_CADMAQUINA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
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
   
   Dim intREGS As Integer
   
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRI
      intOPTATIVO = BREC!SGI_ATIVA
      strANOFABR = BREC!SGI_ANOFABR
      strNUMATIVO = BREC!SGI_NATIVO
      curCAPMINUTO = BREC!SGI_CAPMINU
      If Not IsNull(BREC!SGI_CODUNIMED) Then lngUNIMED = BREC!SGI_CODUNIMED
      If Not IsNull(BREC!SGI_CODFAMILIA) Then lngFAMILIA = BREC!SGI_CODFAMILIA
      If Not IsNull(BREC!SGI_QTDOPERADORES) Then intQtdOperadores = BREC!SGI_QTDOPERADORES
      If Not IsNull(BREC!SGI_FRAOPERADORES) Then curFraOperadores = BREC!SGI_FRAOPERADORES
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   '' --------------------------------------------
   '' Tipo de Alimentação
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMAQFALIM " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC2.Open sSql, adoBanco_Dados
   
   If Not BREC2.EOF Then
      
      intREGS = 0
      Do While Not BREC2.EOF
         intREGS = intREGS + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrGRDFONTALIM(1 To intREGS) As Variant
      
      BREC2.MoveFirst
      intREGS = 1
      Do While Not BREC2.EOF
         arrGRDFONTALIM(intREGS) = BREC2!SGI_CODTIPAL
         intREGS = intREGS + 1
         BREC2.MoveNext
      Loop
      
   End If
   BREC2.Close
   
   '' --------------------------------------------
   '' Turnos
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMAQTURN " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC3.Open sSql, adoBanco_Dados
   
   If Not BREC3.EOF Then
      
      intREGS = 0
      Do While Not BREC3.EOF
         intREGS = intREGS + 1
         BREC3.MoveNext
      Loop
      
      ReDim arrGRDTURNOS(1 To intREGS, 1 To 2) As Variant
      
      BREC3.MoveFirst
      intREGS = 1
      Do While Not BREC3.EOF
         arrGRDTURNOS(intREGS, 1) = BREC3!SGI_CODTURN
         arrGRDTURNOS(intREGS, 2) = BREC3!SGI_PORCUSO
         intREGS = intREGS + 1
         BREC3.MoveNext
      Loop
      
   End If
   BREC3.Close

   '' --------------------------------------------
   '' Produtos
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMAQPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC4.Open sSql, adoBanco_Dados
   
   If Not BREC4.EOF Then
      
      intREGS = 0
      Do While Not BREC4.EOF
         intREGS = intREGS + 1
         BREC4.MoveNext
      Loop
      
      ReDim arrGRDPRODUTOS(1 To intREGS, 1 To 5) As Variant
      
      BREC4.MoveFirst
      intREGS = 1
      Do While Not BREC4.EOF
         arrGRDPRODUTOS(intREGS, 1) = Trim(BREC4!SGI_CODPROD)
         arrGRDPRODUTOS(intREGS, 2) = ""
         arrGRDPRODUTOS(intREGS, 3) = ""
         arrGRDPRODUTOS(intREGS, 4) = ""
         arrGRDPRODUTOS(intREGS, 5) = ""
         If Not IsNull(BREC4!SGI_QTDPCHORA) Then arrGRDPRODUTOS(intREGS, 2) = BREC4!SGI_QTDPCHORA
         If Not IsNull(BREC4!SGI_QTDPCMIN) Then arrGRDPRODUTOS(intREGS, 3) = BREC4!SGI_QTDPCMIN
         If Not IsNull(BREC4!SGI_CUSTOHORA) Then arrGRDPRODUTOS(intREGS, 4) = BREC4!SGI_CUSTOHORA
         If Not IsNull(BREC4!SGI_CUSTOMIN) Then arrGRDPRODUTOS(intREGS, 5) = BREC4!SGI_CUSTOMIN
         intREGS = intREGS + 1
         BREC4.MoveNext
      Loop
      
   End If
   BREC4.Close


   '' --------------------------------------------
   '' Operadores
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMAQOPER " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC4.Open sSql, adoBanco_Dados
   
   If Not BREC4.EOF Then
      
      intREGS = 0
      Do While Not BREC4.EOF
         intREGS = intREGS + 1
         BREC4.MoveNext
      Loop
      
      ReDim arrGRDOPERADOR(1 To intREGS, 1 To 2) As Variant
      
      BREC4.MoveFirst
      intREGS = 1
      Do While Not BREC4.EOF
         arrGRDOPERADOR(intREGS, 1) = BREC4!SGI_CODOPER
         arrGRDOPERADOR(intREGS, 2) = BREC4!SGI_PORCUTIL
         intREGS = intREGS + 1
         BREC4.MoveNext
      Loop
      
   End If
   BREC4.Close
   
   '' --------------------------------------------
   '' Parâmetros para manutenção
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPARAMMANUT " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODMAQUINA = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL     = " & intFilial
   
   BREC5.Open sSql, adoBanco_Dados
   
   If Not BREC5.EOF Then
      
      intREGS = 0
      Do While Not BREC5.EOF
         intREGS = intREGS + 1
         BREC5.MoveNext
      Loop
      
      ReDim arrPARAMETROS(1 To intREGS, 1 To 9) As Variant
      
      BREC5.MoveFirst
      intREGS = 1
      Do While Not BREC5.EOF
         arrPARAMETROS(intREGS, 1) = BREC5!SGI_CODPERIODO
         arrPARAMETROS(intREGS, 2) = Format(BREC5!SGI_DATAMANUT, "DD/MM/YYYY")
         arrPARAMETROS(intREGS, 3) = Format(BREC5!SGI_DATAFINAL, "DD/MM/YYYY")
         arrPARAMETROS(intREGS, 4) = Format(BREC5!SGI_HORAINI, "HH:MM")
         arrPARAMETROS(intREGS, 5) = Format(BREC5!SGI_HORAFIN, "HH:MM")
         arrPARAMETROS(intREGS, 6) = Format(BREC5!SGI_TOTALUSADO, "HH:MM")
         arrPARAMETROS(intREGS, 7) = BREC5!SGI_MAQUSO
         arrPARAMETROS(intREGS, 8) = BREC5!SGI_ATIVO
         arrPARAMETROS(intREGS, 9) = BREC5!SGI_EMCONJUNTO
         intREGS = intREGS + 1
         BREC5.MoveNext
      Loop
      
   End If
   BREC5.Close


   '' --------------------------------------------
   '' Parâmetros para Dias
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPERIODOMANUT " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODMAQUINA = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL     = " & intFilial
   
   BREC6.Open sSql, adoBanco_Dados
   
   If Not BREC6.EOF Then
      
      intREGS = 0
      Do While Not BREC6.EOF
         intREGS = intREGS + 1
         BREC6.MoveNext
      Loop
      
      ReDim arrDIASPARAMETROS(1 To intREGS, 1 To 9) As Variant
      
      BREC6.MoveFirst
      intREGS = 1
      Do While Not BREC6.EOF
         arrDIASPARAMETROS(intREGS, 1) = Format(BREC6!SGI_DATAMANUT, "DD/MM/YYYY")
         arrDIASPARAMETROS(intREGS, 2) = Format(BREC6!SGI_HORAINI, "HH:MM")
         arrDIASPARAMETROS(intREGS, 3) = Format(BREC6!SGI_HORAFIN, "HH:MM")
         arrDIASPARAMETROS(intREGS, 4) = Format(BREC6!SGI_TOTALUSADO, "HH:MM")
         arrDIASPARAMETROS(intREGS, 5) = BREC6!SGI_MAQUSO
         arrDIASPARAMETROS(intREGS, 6) = BREC6!SGI_ATIVO
         arrDIASPARAMETROS(intREGS, 7) = BREC6!SGI_EMCONJUNTO
         arrDIASPARAMETROS(intREGS, 8) = BREC6!SGI_CODPERIODO
         arrDIASPARAMETROS(intREGS, 9) = Format(BREC6!SGI_DATAPARAM, "DD/MM/YYYY")
         intREGS = intREGS + 1
         BREC6.MoveNext
      Loop
      
   End If
   BREC6.Close

   '' --------------------------------------------
   '' Tipo de Alimentação
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_LINHAMAQPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "Order By SGI_ITEN,SGI_SEQCORT"
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF() Then
      
      intREGS = 0
      Do While Not BREC2.EOF
         intREGS = intREGS + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrLINHACAP(1 To intREGS, 1 To 6) As Variant
      
      BREC2.MoveFirst
      intREGS = 1
      Do While Not BREC2.EOF
         arrLINHACAP(intREGS, 1) = BREC2!SGI_ITEN
         arrLINHACAP(intREGS, 2) = BREC2!SGI_IDLIN
         arrLINHACAP(intREGS, 3) = BREC2!SGI_CODLIN
         arrLINHACAP(intREGS, 4) = BREC2!SGI_SEQCORT
         arrLINHACAP(intREGS, 5) = BREC2!SGI_CODCORT
         arrLINHACAP(intREGS, 6) = BREC2!SGI_INDICE
         intREGS = intREGS + 1
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

