VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADTRANSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial     As Integer
Dim lngCODIGO     As Long
Dim strRAZAOSOC   As String
Dim strCPFCNPJ    As String
Dim strENDERECO   As String
Dim strBAIRRO     As String
Dim strCIDADE     As String
Dim intESTADO     As Integer
Dim strTELEFONE   As String
Dim strCONTATO    As String
Dim strCEP        As String

Public Property Get CONTATO() As String
    CONTATO = strCONTATO
End Property

Public Property Let CONTATO(ByVal vData As String)
    strCONTATO = vData
End Property

Public Property Get TELEFONE() As String
    TELEFONE = strTELEFONE
End Property

Public Property Let TELEFONE(ByVal vData As String)
    strTELEFONE = vData
End Property

Public Property Get CEP() As String
    CEP = strCEP
End Property

Public Property Let CEP(ByVal vData As String)
    strCEP = vData
End Property

Public Property Get ESTADO() As Integer
    ESTADO = intESTADO
End Property

Public Property Let ESTADO(ByVal vData As Integer)
    intESTADO = vData
End Property

Public Property Get CIDADE() As String
    CIDADE = strCIDADE
End Property

Public Property Let CIDADE(ByVal vData As String)
    strCIDADE = vData
End Property

Public Property Get BAIRRO() As String
    BAIRRO = strBAIRRO
End Property

Public Property Let BAIRRO(ByVal vData As String)
    strBAIRRO = vData
End Property

Public Property Get ENDERECO() As String
    ENDERECO = strENDERECO
End Property

Public Property Let ENDERECO(ByVal vData As String)
    strENDERECO = vData
End Property

Public Property Get CPFCNPJ() As String
    CPFCNPJ = strCPFCNPJ
End Property

Public Property Let CPFCNPJ(ByVal vData As String)
    strCPFCNPJ = vData
End Property

Public Property Get RAZAOSOC() As String
    RAZAOSOC = strRAZAOSOC
End Property

Public Property Let RAZAOSOC(ByVal vData As String)
    strRAZAOSOC = vData
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
    frmCADTRANSPP.cCaminho = cCaminho
    frmCADTRANSPP.Linha = Linha
    frmCADTRANSPP.FILIAL = iFilial
    frmCADTRANSPP.strAcesso = strAcesso
    frmCADTRANSPP.Show vbModal
End Sub

Public Function Pesq_CadTransp() As Boolean

   Pesq_CadTransp = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
   sSql = sSql & " Where "
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadTransp = True
   BREC.Close

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

On Error GoTo err_Trans

     GRAVA = False
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     sSql = ""
     
     If strOPER = "I" Then
        
        '' Grava Tipo de Produto
        sSql = "Insert Into SGI_CADTRANSP (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                   ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                   ,SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "                   ,SGI_ENDERECO" & vbCrLf
        sSql = sSql & "                   ,SGI_CPFCNPJ" & vbCrLf
        sSql = sSql & "                   ,SGI_BAIRRO" & vbCrLf
        sSql = sSql & "                   ,SGI_CIDADE" & vbCrLf
        sSql = sSql & "                   ,SGI_ESTADO" & vbCrLf
        sSql = sSql & "                   ,SGI_CEP" & vbCrLf
        sSql = sSql & "                   ,SGI_TELEFONE" & vbCrLf
        sSql = sSql & "                   ,SGI_CONTATO" & vbCrLf
        sSql = sSql & "                   )" & vbCrLf
        sSql = sSql & "            Values (" & intFilial & vbCrLf
        sSql = sSql & "                   ," & lngCODIGO & vbCrLf
        sSql = sSql & "                   ,'" & strRAZAOSOC & "'" & vbCrLf
        sSql = sSql & "                   ,'" & strENDERECO & "'" & vbCrLf
        sSql = sSql & "                   ,'" & strCPFCNPJ & "'" & vbCrLf
        sSql = sSql & "                   ,'" & strBAIRRO & "'" & vbCrLf
        sSql = sSql & "                   ,'" & strCIDADE & "'" & vbCrLf
        sSql = sSql & "                   ," & intESTADO & vbCrLf
        sSql = sSql & "                   ,'" & strCEP & "'" & vbCrLf
        sSql = sSql & "                   ,'" & strTELEFONE & "'" & vbCrLf
        sSql = sSql & "                   ,'" & strCONTATO & "'" & vbCrLf
        sSql = sSql & "                    )"
     
     ElseIf strOPER = "A" Then
     
        '' Grava Apaga tipo de Produto
        sSql = "Update SGI_CADTRANSP Set " & vbCrLf
        sSql = sSql & "                           SGI_DESCRICAO = '" & strRAZAOSOC & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_ENDERECO  = '" & strENDERECO & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_CPFCNPJ   = '" & strCPFCNPJ & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_BAIRRO    = '" & strBAIRRO & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_CIDADE    = '" & strCIDADE & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_ESTADO    = " & intESTADO & vbCrLf
        sSql = sSql & "                          ,SGI_CEP       = '" & strCEP & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_TELEFONE  = '" & strTELEFONE & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_CONTATO   = '" & strCONTATO & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & lngCODIGO & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Tipo de Produto
        sSql = "Delete from SGI_CADTRANSP " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
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
   sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      lngCODIGO = BREC!SGI_CODIGO
      strRAZAOSOC = BREC!SGI_DESCRICAO
      strCPFCNPJ = BREC!SGI_CPFCNPJ
      strENDERECO = BREC!SGI_ENDERECO
      strBAIRRO = BREC!SGI_BAIRRO
      strCIDADE = BREC!SGI_CIDADE
      intESTADO = BREC!SGI_ESTADO
      strTELEFONE = BREC!SGI_TELEFONE
      strCONTATO = BREC!SGI_CONTATO
      strCEP = BREC!SGI_CEP
      
      Carrega_campos = True
   End If
   
   BREC.Close

End Function

