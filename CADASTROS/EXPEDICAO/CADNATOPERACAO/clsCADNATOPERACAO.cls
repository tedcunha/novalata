VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADNATOPERACAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim lngCODIGO       As Long
Dim strNOMECLA      As String
Dim strDESCRICAO    As String
Dim intENTSAI       As Integer
Dim intDENFOREST    As Integer
Dim intIMPEXT       As Integer
Dim arrALIQICMS     As Variant
Dim intDefault      As Integer
Dim intPessoaFJ     As Integer
Dim strSITTRIB      As String
Dim intESPECIAL     As Integer
Dim intESPECIAL02   As Integer
Dim strALIQPIS      As String
Dim strALIQCOFINS   As String
Dim intREGIMESTESP  As Integer

Public Property Get REGIMESTESP() As Integer
    REGIMESTESP = intREGIMESTESP
End Property

Public Property Let REGIMESTESP(ByVal vData As Integer)
    intREGIMESTESP = vData
End Property

Public Property Get ALIQCOFINS() As String
    ALIQCOFINS = strALIQCOFINS
End Property

Public Property Let ALIQCOFINS(ByVal vData As String)
    strALIQCOFINS = vData
End Property

Public Property Get ALIQPIS() As String
    ALIQPIS = strALIQPIS
End Property

Public Property Let ALIQPIS(ByVal vData As String)
    strALIQPIS = vData
End Property

Public Property Get ESPECIAL02() As Integer
    ESPECIAL02 = intESPECIAL02
End Property

Public Property Let ESPECIAL02(ByVal vData As Integer)
    intESPECIAL02 = vData
End Property

Public Property Get ESPECIAL() As Integer
    ESPECIAL = intESPECIAL
End Property

Public Property Let ESPECIAL(ByVal vData As Integer)
    intESPECIAL = vData
End Property

Public Property Get SITTRIB() As String
    SITTRIB = strSITTRIB
End Property

Public Property Let SITTRIB(ByVal vData As String)
    strSITTRIB = vData
End Property

Public Property Get PessoaFJ() As Integer
    PessoaFJ = intPessoaFJ
End Property

Public Property Let PessoaFJ(ByVal vData As Integer)
    intPessoaFJ = vData
End Property


Public Property Get DEFAULT() As Integer
    DEFAULT = intDefault
End Property

Public Property Let DEFAULT(ByVal vData As Integer)
    intDefault = vData
End Property

Public Property Get ALIQICMS() As Variant
    ALIQICMS = arrALIQICMS
End Property

Public Property Let ALIQICMS(ByVal vData As Variant)
    arrALIQICMS = vData
End Property

Public Property Get IMPEXT() As Integer
    IMPEXT = intIMPEXT
End Property

Public Property Let IMPEXT(ByVal vData As Integer)
    intIMPEXT = vData
End Property

Public Property Get DENFOREST() As Integer
    DENFOREST = intDENFOREST
End Property

Public Property Let DENFOREST(ByVal vData As Integer)
    intDENFOREST = vData
End Property

Public Property Get ENTSAI() As Integer
    ENTSAI = intENTSAI
End Property

Public Property Let ENTSAI(ByVal vData As Integer)
    intENTSAI = vData
End Property

Public Property Get NOMECLA() As String
    NOMECLA = strNOMECLA
End Property

Public Property Let NOMECLA(ByVal vData As String)
    strNOMECLA = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Property Get DESCRICAO() As String
    DESCRICAO = strDESCRICAO
End Property

Public Property Let DESCRICAO(ByVal vData As String)
    strDESCRICAO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property


Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long)
    frmCADNATOPERACAOP.cCaminho = cCaminho
    frmCADNATOPERACAOP.Linha = Linha
    frmCADNATOPERACAOP.FILIAL = iFilial
    frmCADNATOPERACAOP.strAcesso = strAcesso
    frmCADNATOPERACAOP.lngCodUsuaro = lngIDUsuario
    frmCADNATOPERACAOP.Show vbModal
End Sub

Public Function Pesq_CadNatOperacao() As Boolean

   Pesq_CadNatOperacao = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadNatOperacao = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
         
     Dim sValor As String
     Dim I      As Integer
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
        
        '' Grava Natureza de Operação...
        sSql = "Insert Into SGI_CADNATOPERACAO (" & vbCrLf
        sSql = sSql & "                          SGI_FILIAL      " & vbCrLf
        sSql = sSql & "                         ,SGI_CODIGO      " & vbCrLf
        sSql = sSql & "                         ,SGI_NOMECLCOD   " & vbCrLf
        sSql = sSql & "                         ,SGI_DESCRICAO   " & vbCrLf
        sSql = sSql & "                         ,SGI_ENTSAI      " & vbCrLf
        sSql = sSql & "                         ,SGI_DENTFORAEST " & vbCrLf
        sSql = sSql & "                         ,SGI_IMPEXT      " & vbCrLf
        sSql = sSql & "                         ,SGI_DEFAULT     " & vbCrLf
        sSql = sSql & "                         ,SGI_PESSOAFJ    " & vbCrLf
        sSql = sSql & "                         ,SGI_SITTRIB     " & vbCrLf
        sSql = sSql & "                         ,SGI_EXPECIAL    " & vbCrLf
        sSql = sSql & "                         ,SGI_EXPECIAL02  " & vbCrLf
        sSql = sSql & "                         ,SGI_ALIQPIS     " & vbCrLf
        sSql = sSql & "                         ,SGI_ALIQCOFINS  " & vbCrLf
        sSql = sSql & "                         ,SGI_REGIMESTESP " & vbCrLf
        
        sSql = sSql & "               ) Values (" & vbCrLf
        
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & lngCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & Trim(strNOMECLA) & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strDESCRICAO & "'" & vbCrLf
        sSql = sSql & "                           ," & intENTSAI & vbCrLf
        sSql = sSql & "                           ," & intDENFOREST & vbCrLf
        sSql = sSql & "                           ," & intIMPEXT & vbCrLf
        sSql = sSql & "                           ," & intDefault & vbCrLf
        sSql = sSql & "                           ," & intPessoaFJ & vbCrLf
        sSql = sSql & "                           ,'" & Trim(strSITTRIB) & "'" & vbCrLf
        sSql = sSql & "                           ," & intESPECIAL & vbCrLf
        sSql = sSql & "                           ," & intESPECIAL02 & vbCrLf
        sSql = sSql & "                           ," & strALIQPIS & vbCrLf
        sSql = sSql & "                           ," & strALIQCOFINS & vbCrLf
        sSql = sSql & "                           ," & intREGIMESTESP & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrALIQICMS) = True Then
           For I = 1 To UBound(arrALIQICMS)
           
               sSql = ""
               
               sSql = "Insert Into SGI_CADALIQICMS (" & vbCrLf
               
               sSql = sSql & "                              SGI_FILIAL  " & vbCrLf
               sSql = sSql & "                             ,SGI_CODIGO  " & vbCrLf
               sSql = sSql & "                             ,SGI_ESTORIG " & vbCrLf
               sSql = sSql & "                             ,SGI_ESTDEST " & vbCrLf
               sSql = sSql & "                             ,SGI_PORCICMS" & vbCrLf
               sSql = sSql & "                             ,SGI_TEMSTSN" & vbCrLf
               sSql = sSql & "                             ,SGI_ALIQST" & vbCrLf
               sSql = sSql & "                             ,SGI_ALIICMSINT" & vbCrLf
               sSql = sSql & "                             ,SGI_ALIIVAORIG" & vbCrLf
               sSql = sSql & "                             ,SGI_PROTOCOLOS" & vbCrLf
               sSql = sSql & "                             ,SGI_OPTSIMPSN" & vbCrLf
               sSql = sSql & "                             ,SGI_PROTSIMPL" & vbCrLf
               sSql = sSql & "                             ,SGI_PORCICMSSIMPL" & vbCrLf
               sSql = sSql & "                             ,SGI_ALIIVAORIGSIMPL" & vbCrLf
               sSql = sSql & "                             ,SGI_ALIICMSINTSIMPL" & vbCrLf
               sSql = sSql & "                             ,SGI_ALIQSTSIMPL" & vbCrLf
               
               sSql = sSql & "                   ) Values (" & vbCrLf
               
               sSql = sSql & "                             " & intFilial & vbCrLf
               sSql = sSql & "                            ," & lngCODIGO & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 1) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 2) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 3) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 4) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 5) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 6) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 7) & vbCrLf
               sSql = sSql & "                            ,'" & arrALIQICMS(I, 8) & "'" & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 9) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 10) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 11) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 12) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 13) & vbCrLf
               sSql = sSql & "                            ," & arrALIQICMS(I, 14) & vbCrLf
               
               sSql = sSql & "                             )"
               
               
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
        
     
     ElseIf strOPER = "A" Then
     
        '' Aliquota de ICMS...
        sSql = "Delete from SGI_CADALIQICMS " & vbCrLf
        sSql = sSql & "                      Where " & vbCrLf
        sSql = sSql & "                            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "                        And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Altera Natureza de Operação...
        sSql = "Update SGI_CADNATOPERACAO Set " & vbCrLf
        sSql = sSql & "                      SGI_NOMECLCOD   = '" & Trim(strNOMECLA) & "'" & vbCrLf
        sSql = sSql & "                     ,SGI_DESCRICAO   = '" & strDESCRICAO & "'" & vbCrLf
        sSql = sSql & "                     ,SGI_ENTSAI      = " & intENTSAI & vbCrLf
        sSql = sSql & "                     ,SGI_DENTFORAEST = " & intDENFOREST & vbCrLf
        sSql = sSql & "                     ,SGI_IMPEXT      = " & intIMPEXT & vbCrLf
        sSql = sSql & "                     ,SGI_DEFAULT     = " & intDefault & vbCrLf
        sSql = sSql & "                     ,SGI_PESSOAFJ    = " & intPessoaFJ & vbCrLf
        sSql = sSql & "                     ,SGI_SITTRIB     = '" & Trim(strSITTRIB) & "'" & vbCrLf
        sSql = sSql & "                     ,SGI_EXPECIAL    = " & intESPECIAL & vbCrLf
        sSql = sSql & "                     ,SGI_EXPECIAL02  = " & intESPECIAL02 & vbCrLf
        sSql = sSql & "                     ,SGI_ALIQPIS     = " & strALIQPIS & vbCrLf
        sSql = sSql & "                     ,SGI_ALIQCOFINS  = " & strALIQCOFINS & vbCrLf
        sSql = sSql & "                     ,SGI_REGIMESTESP = " & intREGIMESTESP & vbCrLf
        
        sSql = sSql & "                 Where " & vbCrLf
        sSql = sSql & "                       SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "                   And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
    
        If IsArray(arrALIQICMS) = True Then
           For I = 1 To UBound(arrALIQICMS)
           
                If Len(Trim(arrALIQICMS(I, 1))) <> 0 And Len(Trim(arrALIQICMS(I, 2))) <> 0 And Len(Trim(arrALIQICMS(I, 3))) <> 0 Then
                    sSql = "Insert Into SGI_CADALIQICMS (" & vbCrLf
                    
                    sSql = sSql & "                              SGI_FILIAL  " & vbCrLf
                    sSql = sSql & "                             ,SGI_CODIGO  " & vbCrLf
                    sSql = sSql & "                             ,SGI_ESTORIG " & vbCrLf
                    sSql = sSql & "                             ,SGI_ESTDEST " & vbCrLf
                    sSql = sSql & "                             ,SGI_PORCICMS" & vbCrLf
                    sSql = sSql & "                             ,SGI_TEMSTSN" & vbCrLf
                    sSql = sSql & "                             ,SGI_ALIQST" & vbCrLf
                    sSql = sSql & "                             ,SGI_ALIICMSINT" & vbCrLf
                    sSql = sSql & "                             ,SGI_ALIIVAORIG" & vbCrLf
                    sSql = sSql & "                             ,SGI_PROTOCOLOS" & vbCrLf
                    sSql = sSql & "                             ,SGI_OPTSIMPSN" & vbCrLf
                    sSql = sSql & "                             ,SGI_PROTSIMPL" & vbCrLf
                    sSql = sSql & "                             ,SGI_PORCICMSSIMPL" & vbCrLf
                    sSql = sSql & "                             ,SGI_ALIIVAORIGSIMPL" & vbCrLf
                    sSql = sSql & "                             ,SGI_ALIICMSINTSIMPL" & vbCrLf
                    sSql = sSql & "                             ,SGI_ALIQSTSIMPL" & vbCrLf
                    
                    sSql = sSql & "                   ) Values (" & vbCrLf
                    
                    sSql = sSql & "                             " & intFilial & vbCrLf
                    sSql = sSql & "                            ," & lngCODIGO & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 1) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 2) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 3) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 4) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 5) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 6) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 7) & vbCrLf
                    sSql = sSql & "                            ,'" & arrALIQICMS(I, 8) & "'" & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 9) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 10) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 11) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 12) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 13) & vbCrLf
                    sSql = sSql & "                            ," & arrALIQICMS(I, 14) & vbCrLf
                    
                    sSql = sSql & "                             )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                End If
            
           Next I
        End If
     
     ElseIf strOPER = "E" Then
     
        sSql = "Delete from SGI_CADALIQICMS " & vbCrLf
        sSql = sSql & "                      Where " & vbCrLf
        sSql = sSql & "                            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "                        And SGI_CODIGO = " & lngCODIGO
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Exclui Natureza de Operação...
        sSql = "Delete from SGI_CADNATOPERACAO " & vbCrLf
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
   
   Dim intReg As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF() Then
   
      lngCODIGO = BREC!SGI_CODIGO
      strNOMECLA = BREC!SGI_NOMECLCOD
      strDESCRICAO = BREC!SGI_DESCRICAO
      intENTSAI = BREC!SGI_ENTSAI
      intDENFOREST = BREC!SGI_DENTFORAEST
      intIMPEXT = BREC!SGI_IMPEXT
      intDefault = BREC!SGI_DEFAULT
      If Not IsNull(BREC!SGI_PESSOAFJ) Then intPessoaFJ = BREC!SGI_PESSOAFJ
      If Not IsNull(BREC!SGI_SITTRIB) Then strSITTRIB = BREC!SGI_SITTRIB
      
      If Not IsNull(BREC!SGI_EXPECIAL) Then intESPECIAL = BREC!SGI_EXPECIAL
      If Not IsNull(BREC!SGI_EXPECIAL02) Then intESPECIAL02 = BREC!SGI_EXPECIAL02
      
      If Not IsNull(BREC!SGI_ALIQPIS) Then strALIQPIS = Format(BREC!SGI_ALIQPIS, "#,##0.00")
      If Not IsNull(BREC!SGI_ALIQCOFINS) Then strALIQCOFINS = Format(BREC!SGI_ALIQCOFINS, "#,##0.00")
      
      intREGIMESTESP = BREC!SGI_REGIMESTESP
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   
   If Carrega_campos = False Then Exit Function
   
   '' -----------------------------------------------------
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADALIQICMS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrALIQICMS(1 To intReg, 1 To 14) As String
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF()
      
         arrALIQICMS(intReg, 1) = Trim(Str(BREC2!SGI_ESTORIG))
         arrALIQICMS(intReg, 2) = Trim(Str(BREC2!SGI_ESTDEST))
         If Not IsNull(BREC2!SGI_PORCICMS) Then arrALIQICMS(intReg, 3) = BREC2!SGI_PORCICMS
         arrALIQICMS(intReg, 4) = BREC2!SGI_TEMSTSN
         If Not IsNull(BREC2!SGI_ALIQST) Then arrALIQICMS(intReg, 5) = BREC2!SGI_ALIQST
         If Not IsNull(BREC2!SGI_ALIICMSINT) Then arrALIQICMS(intReg, 6) = BREC2!SGI_ALIICMSINT
         If Not IsNull(BREC2!SGI_ALIIVAORIG) Then arrALIQICMS(intReg, 7) = BREC2!SGI_ALIIVAORIG
         If Not IsNull(BREC2!SGI_PROTOCOLOS) Then arrALIQICMS(intReg, 8) = BREC2!SGI_PROTOCOLOS
         
         If Not IsNull(BREC2!SGI_OPTSIMPSN) Then arrALIQICMS(intReg, 9) = BREC2!SGI_OPTSIMPSN
         If Not IsNull(BREC2!SGI_PROTSIMPL) Then arrALIQICMS(intReg, 10) = BREC2!SGI_PROTSIMPL
         If Not IsNull(BREC2!SGI_PORCICMSSIMPL) Then arrALIQICMS(intReg, 11) = BREC2!SGI_PORCICMSSIMPL
         If Not IsNull(BREC2!SGI_ALIIVAORIGSIMPL) Then arrALIQICMS(intReg, 12) = BREC2!SGI_ALIIVAORIGSIMPL
         If Not IsNull(BREC2!SGI_ALIICMSINTSIMPL) Then arrALIQICMS(intReg, 13) = BREC2!SGI_ALIICMSINTSIMPL
         If Not IsNull(BREC2!SGI_ALIQSTSIMPL) Then arrALIQICMS(intReg, 14) = BREC2!SGI_ALIQSTSIMPL
         
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   

End Function


