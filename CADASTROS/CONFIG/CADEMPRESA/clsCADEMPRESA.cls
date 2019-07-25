VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADEMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intEMPCOD       As Integer
Dim strEMPDESC      As String
Dim strEMPCNPJ      As String
Dim strEMPINSCEST   As String
Dim intFILIAL       As Integer
Dim intPradrao      As Integer
Dim strEMPENDCOB    As String
Dim strEMPBAICOB    As String
Dim strEMPESTCOB    As String
Dim strEMPCIDCOB    As String
Dim strEMPCEPCOB    As String
Dim strEMPFONCOB    As String
Dim strEMPFAXCOB    As String
Dim strEMPENDENT    As String
Dim strEMPBAIENT    As String
Dim strEMPESTENT    As String
Dim strEMPCIDENT    As String
Dim strEMPCEPENT    As String
Dim strEMPFONENT    As String
Dim strEMPFAXENT    As String
Dim strCAMINHO      As String
Dim strARQUIVO      As String
Dim strDATEXP       As String

Dim strEMAIL        As String
Dim strSENHAEMAIL   As String
Dim strSMTP         As String
Dim strPORTASMTP    As String
Dim strEMAILCC      As String

Public Property Get EMAIL() As String
    EMAIL = strEMAIL
End Property

Public Property Let EMAIL(ByVal vData As String)
    strEMAIL = vData
End Property

Public Property Get SENHAEMAIL() As String
    SENHAEMAIL = strSENHAEMAIL
End Property

Public Property Let SENHAEMAIL(ByVal vData As String)
    strSENHAEMAIL = vData
End Property

Public Property Get SMTP() As String
    SMTP = strSMTP
End Property

Public Property Let SMTP(ByVal vData As String)
    strSMTP = vData
End Property

Public Property Get PORTASMTP() As String
    PORTASMTP = strPORTASMTP
End Property

Public Property Let PORTASMTP(ByVal vData As String)
    strPORTASMTP = vData
End Property

Public Property Get EMAILCC() As String
    EMAILCC = strEMAILCC
End Property

Public Property Let EMAILCC(ByVal vData As String)
    strEMAILCC = vData
End Property

Public Property Get DATEXP() As String
    DATEXP = strDATEXP
End Property

Public Property Let DATEXP(ByVal vData As String)
    strDATEXP = vData
End Property

Public Property Get ARQUIVO() As String
    ARQUIVO = strARQUIVO
End Property

Public Property Let ARQUIVO(ByVal vData As String)
    strARQUIVO = vData
End Property


Public Property Get CAMINHO() As String
    CAMINHO = strCAMINHO
End Property

Public Property Let CAMINHO(ByVal vData As String)
    strCAMINHO = vData
End Property

Public Property Get EMPINSCEST() As String
    EMPINSCEST = strEMPINSCEST
End Property

Public Property Let EMPINSCEST(ByVal vData As String)
    strEMPINSCEST = vData
End Property

Public Property Get EMPESTENT() As String
    EMPESTENT = strEMPESTENT
End Property

Public Property Let EMPESTENT(ByVal vData As String)
    strEMPESTENT = vData
End Property

Public Property Get EMPESTCOB() As String
    EMPESTCOB = strEMPESTCOB
End Property

Public Property Let EMPESTCOB(ByVal vData As String)
    strEMPESTCOB = vData
End Property

Public Property Get EMPFAXCOB() As String
    EMPFAXCOB = strEMPFAXCOB
End Property

Public Property Let EMPFAXCOB(ByVal vData As String)
    strEMPFAXCOB = vData
End Property

Public Property Get EMPFONCOB() As String
    EMPFONCOB = strEMPFONCOB
End Property

Public Property Let EMPFONCOB(ByVal vData As String)
    strEMPFONCOB = vData
End Property

Public Property Get EMPCEPCOB() As String
    EMPCEPCOB = strEMPCEPCOB
End Property

Public Property Let EMPCEPCOB(ByVal vData As String)
    strEMPCEPCOB = vData
End Property

Public Property Get EMPCIDCOB() As String
    EMPCIDCOB = strEMPCIDCOB
End Property

Public Property Let EMPCIDCOB(ByVal vData As String)
    strEMPCIDCOB = vData
End Property

Public Property Get EMPBAICOB() As String
    EMPBAICOB = strEMPBAICOB
End Property

Public Property Let EMPBAICOB(ByVal vData As String)
    strEMPBAICOB = vData
End Property

Public Property Get EMPENDCOB() As String
    EMPENDCOB = strEMPENDCOB
End Property

Public Property Let EMPENDCOB(ByVal vData As String)
    strEMPENDCOB = vData
End Property

Public Property Get EMPFAXENT() As String
    EMPFAXENT = strEMPFAXENT
End Property

Public Property Let EMPFAXENT(ByVal vData As String)
    strEMPFAXENT = vData
End Property

Public Property Get EMPFONENT() As String
    EMPFONENT = strEMPFONENT
End Property

Public Property Let EMPFONENT(ByVal vData As String)
    strEMPFONENT = vData
End Property

Public Property Get EMPCEPENT() As String
    EMPCEPENT = strEMPCEPENT
End Property

Public Property Let EMPCEPENT(ByVal vData As String)
    strEMPCEPENT = vData
End Property

Public Property Get EMPCIDENT() As String
    EMPCIDENT = strEMPCIDENT
End Property

Public Property Let EMPCIDENT(ByVal vData As String)
    strEMPCIDENT = vData
End Property

Public Property Get EMPBAIENT() As String
    EMPBAIENT = strEMPBAIENT
End Property

Public Property Let EMPBAIENT(ByVal vData As String)
    strEMPBAIENT = vData
End Property

Public Property Get EMPENDENT() As String
    EMPENDENT = strEMPENDENT
End Property

Public Property Let EMPENDENT(ByVal vData As String)
    strEMPENDENT = vData
End Property

Public Property Get Pradrao() As Integer
    Pradrao = intPradrao
End Property

Public Property Let Pradrao(ByVal vData As Integer)
    intPradrao = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFILIAL
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFILIAL = vData
End Property

Public Property Get EMPCOD() As Integer
    EMPCOD = intEMPCOD
End Property

Public Property Let EMPCOD(ByVal vData As Integer)
    intEMPCOD = vData
End Property

Public Property Get EMPDESC() As String
    EMPDESC = strEMPDESC
End Property

Public Property Let EMPDESC(ByVal vData As String)
    strEMPDESC = vData
End Property

Public Property Get EMPCNPJ() As String
    EMPCNPJ = strEMPCNPJ
End Property

Public Property Let EMPCNPJ(ByVal vData As String)
    strEMPCNPJ = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADEMPRTP.cCaminho = cCaminho
    frmCADEMPRTP.Linha = Linha
    frmCADEMPRTP.FILIAL = iFilial
    frmCADEMPRTP.strAcesso = strAcesso
    frmCADEMPRTP.Show vbModal
End Sub

Public Function Pesq_Empresa() As Boolean

   Pesq_Empresa = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_FILIAL " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL <> " & intFILIAL
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
   
      intEMPCOD = BREC!SGI_FILIAL
      strEMPDESC = BREC!SGI_DESCRICAO
      strEMPCNPJ = BREC!SGI_CNPJ
      
      
   
      Pesq_Empresa = True
   End If
   
   BREC.Close
   

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
    
        '' Grava Empresa
        sSql = "Insert Into SGI_FILIAL (SGI_FILIAL,SGI_DESCRICAO,SGI_CNPJ,SGI_PADRAO," & vbCrLf
        sSql = sSql & "                SGI_ENDCOBR,SGI_BAICOBR,SGI_ESTCOBR,SGI_CIDCOBR,SGI_CEPCOBR,SGI_FONDCOBR,SGI_FAXDCOBR," & vbCrLf
        sSql = sSql & "                SGI_ENDENTR,SGI_BAIENTR,SGI_ESTENTR,SGI_CIDENTR,SGI_CEPENTR,SGI_FONDENTR,SGI_FAXDENTR,SGI_INSCEST,SGI_NOMEARQ,SGI_CAMINHOARQ,SGI_SENHA,SGI_SENHA2," & vbCrLf
        sSql = sSql & "                SGI_EMAIL,SGI_SENHAEMAIL,SGI_SMTP,SGI_PORTASMTP,SGI_EMAILCC" & vbCrLf
        sSql = sSql & "    ) Values (" & vbCrLf
        
        sSql = sSql & "                             " & intEMPCOD
        sSql = sSql & "                           ,'" & strEMPDESC & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPCNPJ & "'" & vbCrLf
        sSql = sSql & "                           ," & intPradrao & vbCrLf
        
        sSql = sSql & "                           ,'" & strEMPENDCOB & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPBAICOB & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPESTCOB & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPCIDCOB & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPCEPCOB & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPFONCOB & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPFAXCOB & "'" & vbCrLf
        
        sSql = sSql & "                           ,'" & strEMPENDENT & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPBAIENT & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPESTENT & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPCIDENT & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPCEPENT & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPFONENT & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPFAXENT & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strEMPINSCEST & "'" & vbCrLf
        
        sSql = sSql & "                           ,'" & strCAMINHO & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strARQUIVO & "'" & vbCrLf
        sSql = sSql & "                           ," & strDATEXP & vbCrLf
        sSql = sSql & "                           ,Null" & vbCrLf
        sSql = sSql & "                           ," & strEMAIL & vbCrLf
        sSql = sSql & "                           ," & strSENHAEMAIL & vbCrLf
        sSql = sSql & "                           ," & strSMTP & vbCrLf
        sSql = sSql & "                           ," & strPORTASMTP & vbCrLf
        sSql = sSql & "                           ," & strEMAILCC & vbCrLf
        
        sSql = sSql & "                           )"
     
     ElseIf strOPER = "A" Then
     
        '' Grava Empresa
        sSql = "Update SGI_FILIAL Set " & vbCrLf
        sSql = sSql & "               SGI_DESCRICAO  = '" & strEMPDESC & "'" & vbCrLf
        sSql = sSql & "              ,SGI_CNPJ       = '" & strEMPCNPJ & "'" & vbCrLf
        sSql = sSql & "              ,SGI_PADRAO     = " & intPradrao & vbCrLf
        
        sSql = sSql & "              ,SGI_ENDCOBR    = '" & strEMPENDCOB & "'" & vbCrLf
        sSql = sSql & "              ,SGI_BAICOBR    = '" & strEMPBAICOB & "'" & vbCrLf
        sSql = sSql & "              ,SGI_ESTCOBR    = '" & strEMPESTCOB & "'" & vbCrLf
        sSql = sSql & "              ,SGI_CIDCOBR    = '" & strEMPCIDCOB & "'" & vbCrLf
        sSql = sSql & "              ,SGI_CEPCOBR    = '" & strEMPCEPCOB & "'" & vbCrLf
        sSql = sSql & "              ,SGI_FONDCOBR   = '" & strEMPFONCOB & "'" & vbCrLf
        sSql = sSql & "              ,SGI_FAXDCOBR   = '" & strEMPFAXCOB & "'" & vbCrLf
        
        sSql = sSql & "              ,SGI_ENDENTR    = '" & strEMPENDENT & "'" & vbCrLf
        sSql = sSql & "              ,SGI_BAIENTR    = '" & strEMPBAIENT & "'" & vbCrLf
        sSql = sSql & "              ,SGI_ESTENTR    = '" & strEMPESTENT & "'" & vbCrLf
        sSql = sSql & "              ,SGI_CIDENTR    = '" & strEMPCIDENT & "'" & vbCrLf
        sSql = sSql & "              ,SGI_CEPENTR    = '" & strEMPCEPENT & "'" & vbCrLf
        sSql = sSql & "              ,SGI_FONDENTR   = '" & strEMPFONENT & "'" & vbCrLf
        sSql = sSql & "              ,SGI_FAXDENTR   = '" & strEMPFAXENT & "'" & vbCrLf
        sSql = sSql & "              ,SGI_INSCEST    = '" & strEMPINSCEST & "'" & vbCrLf
        
        sSql = sSql & "              ,SGI_CAMINHOARQ = '" & Trim(strCAMINHO) & "'" & vbCrLf
        sSql = sSql & "              ,SGI_NOMEARQ    = '" & Trim(strARQUIVO) & "'" & vbCrLf
        sSql = sSql & "              ,SGI_SENHA      = " & strDATEXP & vbCrLf
        sSql = sSql & "              ,SGI_SENHA2     = Null" & vbCrLf
        
''        sSql = sSql & "              ,SGI_EMAIL      = " & strEMAIL & vbCrLf
''        sSql = sSql & "              ,SGI_SENHAEMAIL = " & strSENHAEMAIL & vbCrLf
''        sSql = sSql & "              ,SGI_SMTP       = " & strSMTP & vbCrLf
''        sSql = sSql & "              ,SGI_PORTASMTP  = " & strPORTASMTP & vbCrLf
''        sSql = sSql & "              ,SGI_EMAILCC    = " & strEMAILCC & vbCrLf
       
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_FILIAL    = " & intEMPCOD
     
     ElseIf strOPER = "E" Then
     
        sSql = "Delete from SGI_FILIAL "
        sSql = sSql & "      Where SGI_FILIAL = " & intEMPCOD
     
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
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_FILIAL " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intEMPCOD
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
   
      intEMPCOD = BREC!SGI_FILIAL
      strEMPDESC = BREC!SGI_DESCRICAO
      strEMPCNPJ = BREC!SGI_CNPJ
      intFILIAL = BREC!SGI_FILIAL
      If Not IsNull(BREC!SGI_SENHA) Then strDATEXP = Format(DateSerial(Year(BREC!SGI_SENHA), Month(BREC!SGI_SENHA), Day(BREC!SGI_SENHA)), "DD/MM/YYYY")
      
      intPradrao = 0
      If Not IsNull(BREC!SGI_PADRAO) Then intPradrao = BREC!SGI_PADRAO
      
      If Not IsNull(BREC!SGI_ENDCOBR) Then strEMPENDCOB = BREC!SGI_ENDCOBR
      If Not IsNull(BREC!SGI_BAICOBR) Then strEMPBAICOB = BREC!SGI_BAICOBR
      If Not IsNull(BREC!SGI_ESTCOBR) Then strEMPESTCOB = BREC!SGI_ESTCOBR
      If Not IsNull(BREC!SGI_CIDCOBR) Then strEMPCIDCOB = BREC!SGI_CIDCOBR
      If Not IsNull(BREC!SGI_CEPCOBR) Then strEMPCEPCOB = BREC!SGI_CEPCOBR
      If Not IsNull(BREC!SGI_FONDCOBR) Then strEMPFONCOB = BREC!SGI_FONDCOBR
      If Not IsNull(BREC!SGI_FAXDCOBR) Then strEMPFAXCOB = BREC!SGI_FAXDCOBR
      
      If Not IsNull(BREC!SGI_ENDENTR) Then strEMPENDENT = BREC!SGI_ENDENTR
      If Not IsNull(BREC!SGI_BAIENTR) Then strEMPBAIENT = BREC!SGI_BAIENTR
      If Not IsNull(BREC!SGI_ESTENTR) Then strEMPESTENT = BREC!SGI_ESTENTR
      If Not IsNull(BREC!SGI_CIDENTR) Then strEMPCIDENT = BREC!SGI_CIDENTR
      If Not IsNull(BREC!SGI_CEPENTR) Then strEMPCEPENT = BREC!SGI_CEPENTR
      If Not IsNull(BREC!SGI_FONDENTR) Then strEMPFONENT = BREC!SGI_FONDENTR
      If Not IsNull(BREC!SGI_FAXDENTR) Then strEMPFAXENT = BREC!SGI_FAXDENTR
      
      If Not IsNull(BREC!SGI_INSCEST) Then strEMPINSCEST = BREC!SGI_INSCEST
      
      If Dir(BREC!SGI_CAMINHOARQ) <> "" Then
        If Not IsNull(BREC!SGI_CAMINHOARQ) Then strCAMINHO = BREC!SGI_CAMINHOARQ
        If Not IsNull(BREC!SGI_NOMEARQ) Then strARQUIVO = BREC!SGI_NOMEARQ
      Else
        strCAMINHO = ""
        strARQUIVO = ""
      End If
   
''      If Not IsNull(BREC!SGI_EMAIL) Then strEMAIL = BREC!SGI_EMAIL
''      If Not IsNull(BREC!SGI_SENHAEMAIL) Then strSENHAEMAIL = BREC!SGI_SENHAEMAIL
''      If Not IsNull(BREC!SGI_SMTP) Then strSMTP = BREC!SGI_SMTP
''      If Not IsNull(BREC!SGI_PORTASMTP) Then strPORTASMTP = BREC!SGI_PORTASMTP
''      If Not IsNull(BREC!SGI_EMAILCC) Then strEMAILCC = BREC!SGI_EMAILCC
      
      Carrega_campos = True
   End If
   
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
    
    BREC.Open sSql, adoBanco_Dados
    
    If Not BREC.EOF Then
    
       If IsNull(BREC!SGI_NUMERO) = True Then
          
          Gera_Codigo = 1
          
          sSql = "Insert into SGI_NUMERO (SGI_NUMERO,SGI_MODULO) Values(" & vbCrLf
          sSql = sSql + "                                            1" & vbCrLf
          sSql = sSql + "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql + "                                          )" & vbCrLf
          
       ElseIf BREC!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql + "           SGI_NUMERO = " & BREC!SGI_NUMERO & vbCrLf
          sSql = sSql + "         Where " & vbCrLf
          sSql = sSql + "               SGI_MODULO = '" & sModulo & "'"
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
    End If
    
    BREC.Close
    
End Function

