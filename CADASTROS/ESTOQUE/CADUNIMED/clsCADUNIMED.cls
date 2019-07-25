VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADUNIMED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intUNIMEDCOD    As Integer
Dim strUNIMEDUNI    As String
Dim strUNIMEDDES    As String
Dim intFilial       As Integer
Dim lngCODFAMUNID   As Long
Dim intPADRAO       As Integer
Dim curFATOR        As Currency

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADUNIMEDP.cCaminho = cCaminho
    frmCADUNIMEDP.Linha = Linha
    frmCADUNIMEDP.FILIAL = iFilial
    frmCADUNIMEDP.strAcesso = strAcesso
    frmCADUNIMEDP.Show vbModal
End Sub

Public Property Get FATOR() As Currency
    FATOR = curFATOR
End Property

Public Property Let FATOR(ByVal vData As Currency)
    curFATOR = vData
End Property

Public Property Get PADRAO() As Integer
    PADRAO = intPADRAO
End Property

Public Property Let PADRAO(ByVal vData As Integer)
    intPADRAO = vData
End Property

Public Property Get CODFAMUNID() As Long
    CODFAMUNID = lngCODFAMUNID
End Property

Public Property Let CODFAMUNID(ByVal vData As Long)
    lngCODFAMUNID = vData
End Property

Public Property Get UNIMEDCOD() As Integer
    UNIMEDCOD = intUNIMEDCOD
End Property

Public Property Let UNIMEDCOD(ByVal vData As Integer)
    intUNIMEDCOD = vData
End Property

Public Property Get UNIMEDDES() As String
    UNIMEDDES = strUNIMEDDES
End Property

Public Property Let UNIMEDDES(ByVal vData As String)
    strUNIMEDDES = vData
End Property

Public Property Get UNIMEDUNI() As String
    UNIMEDUNI = strUNIMEDUNI
End Property

Public Property Let UNIMEDUNI(ByVal vData As String)
    strUNIMEDUNI = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property


Public Function Pesq_Cadunimed() As Boolean

   Pesq_Cadunimed = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADUNIMED " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intUNIMEDCOD = BREC!SGI_CODIGO
      strUNIMEDUNI = BREC!SGI_UNIDADE
      strUNIMEDDES = BREC!SGI_DESCRICAO
   
      Pesq_Cadunimed = True
   End If
   
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans

     GRAVA = False
     
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
                
        '' Grava Tipo de Produto
        sSql = "Insert Into SGI_CADUNIMED (SGI_FILIAL,SGI_CODIGO,SGI_UNIDADE,SGI_DESCRICAO,SGI_CODFAMUNID,SGI_FATOR,SGI_PADRAO) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intUNIMEDCOD
        sSql = sSql & "                           ,'" & strUNIMEDUNI & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strUNIMEDDES & "'" & vbCrLf
        sSql = sSql & "                           ," & lngCODFAMUNID & vbCrLf
        
        strValor = Replace(Format(curFATOR, "#,####0.0000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        sSql = sSql & "                           ," & intPADRAO & vbCrLf
        
        sSql = sSql & "                           )"
     
     ElseIf strOPER = "A" Then
     
        '' Grava Apaga tipo de Produto
        sSql = "Update SGI_CADUNIMED Set " & vbCrLf
        sSql = sSql & "               SGI_UNIDADE   = '" & strUNIMEDUNI & "'" & vbCrLf
        sSql = sSql & "              ,SGI_DESCRICAO = '" & strUNIMEDDES & "'" & vbCrLf
        
        
        sSql = sSql & "              ,SGI_CODFAMUNID = " & lngCODFAMUNID & vbCrLf
        strValor = Replace(Format(curFATOR, "#,####0.0000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "              ,SGI_FATOR     = " & strValor & vbCrLf
        sSql = sSql & "              ,SGI_PADRAO    = " & intPADRAO & vbCrLf
        
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intUNIMEDCOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Tipo de Produto
        sSql = "Delete from SGI_CADUNIMED " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intUNIMEDCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
     End If
     
     BGRV.CommandText = sSql
     BGRV.Execute
     
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
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intUNIMEDCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intUNIMEDCOD = BREC!SGI_CODIGO
      strUNIMEDUNI = BREC!SGI_UNIDADE
      strUNIMEDDES = BREC!SGI_DESCRICAO
      
      If Not IsNull(BREC!SGI_CODFAMUNID) Then lngCODFAMUNID = BREC!SGI_CODFAMUNID
      If Not IsNull(BREC!SGI_FATOR) Then curFATOR = BREC!SGI_FATOR
      intPADRAO = BREC!SGI_PADRAO
      
      Carrega_campos = True
   End If
   
   BREC.Close

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

Public Sub PreenchComboFamiliaUnidade(Combo As Variant)
    
    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFAMUNIDADE " & vbCrLf
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

