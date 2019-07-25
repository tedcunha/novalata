VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADSENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intSENCODIGO        As Integer
Dim strSENNOME          As String
Dim intSENDEPTO         As Integer
Dim strSENSENHA         As String
Dim intSENACESSO        As Integer
Dim intFilial           As Integer
Dim intFUNCAO           As Integer
Dim intSECAO            As Integer
Dim intSETOR            As Integer
Dim arrEMAIL            As Variant
Dim intPERMBLOQPED      As Integer
Dim intLIBFINANCEIRO    As Integer
Dim intLIBCOMERCIAL     As Integer
Dim intREPEDSN          As Integer
Dim intLIQPEDSN         As Integer
Dim intLIBPEDBLOQSN     As Integer
Dim intLIBPEDFOTSN      As Integer
Dim arrMENUPAI          As Variant
Dim arrMENUFILHO        As Variant
Dim arrMENUNETO         As Variant
Dim intNOVO             As Integer
Dim intDESABPROD        As Integer
Dim intPERMLIBFOT       As Integer
Dim intPERMFAT10POR     As Integer
Dim intLIBPDATAPCOTA    As Integer
Dim intPERMFATROTDIFSN  As Integer
Dim intPVCLIE           As Integer
Dim intEVENDEDOR        As Integer
Dim intATIVO            As Integer
Dim intBLOQCRED         As Integer
Dim intPERMALTPEDFAT    As Integer
Dim intMOP              As Integer
Dim intPermExcPedSN     As Integer

Public Property Get PermExcPedSN() As Integer
    PermExcPedSN = intPermExcPedSN
End Property

Public Property Let PermExcPedSN(ByVal vData As Integer)
    intPermExcPedSN = vData
End Property


Public Property Get MOP() As Integer
    MOP = intMOP
End Property

Public Property Let MOP(ByVal vData As Integer)
    intMOP = vData
End Property

Public Property Get PERMALTPEDFAT() As Integer
    PERMALTPEDFAT = intPERMALTPEDFAT
End Property

Public Property Let PERMALTPEDFAT(ByVal vData As Integer)
    intPERMALTPEDFAT = vData
End Property

Public Property Get BLOQCRED() As Integer
    BLOQCRED = intBLOQCRED
End Property

Public Property Let BLOQCRED(ByVal vData As Integer)
    intBLOQCRED = vData
End Property

Public Property Get ATIVO() As Integer
    ATIVO = intATIVO
End Property

Public Property Let ATIVO(ByVal vData As Integer)
    intATIVO = vData
End Property

Public Property Get EVENDEDOR() As Integer
    EVENDEDOR = intEVENDEDOR
End Property

Public Property Let EVENDEDOR(ByVal vData As Integer)
    intEVENDEDOR = vData
End Property

Public Property Get PVCLIE() As Integer
    PVCLIE = intPVCLIE
End Property

Public Property Let PVCLIE(ByVal vData As Integer)
    intPVCLIE = vData
End Property

Public Property Get PERMFATROTDIFSN() As Integer
    PERMFATROTDIFSN = intPERMFATROTDIFSN
End Property

Public Property Let PERMFATROTDIFSN(ByVal vData As Integer)
    intPERMFATROTDIFSN = vData
End Property

Public Property Get LIBPDATAPCOTA() As Integer
    LIBPDATAPCOTA = intLIBPDATAPCOTA
End Property

Public Property Let LIBPDATAPCOTA(ByVal vData As Integer)
    intLIBPDATAPCOTA = vData
End Property

Public Property Get PERMFAT10POR() As Integer
    PERMFAT10POR = intPERMFAT10POR
End Property

Public Property Let PERMFAT10POR(ByVal vData As Integer)
    intPERMFAT10POR = vData
End Property

Public Property Get PERMLIBFOT() As Integer
    PERMLIBFOT = intPERMLIBFOT
End Property

Public Property Let PERMLIBFOT(ByVal vData As Integer)
    intPERMLIBFOT = vData
End Property

Public Property Get DESABPROD() As Integer
    DESABPROD = intDESABPROD
End Property

Public Property Let DESABPROD(ByVal vData As Integer)
    intDESABPROD = vData
End Property

Public Property Get NOVO() As Integer
    NOVO = intNOVO
End Property

Public Property Let NOVO(ByVal vData As Integer)
    intNOVO = vData
End Property

Public Property Get MENUNETO() As Variant
    MENUNETO = arrMENUNETO
End Property

Public Property Let MENUNETO(ByVal vData As Variant)
    arrMENUNETO = vData
End Property

Public Property Get MENUFILHO() As Variant
    MENUFILHO = arrMENUFILHO
End Property

Public Property Let MENUFILHO(ByVal vData As Variant)
    arrMENUFILHO = vData
End Property

Public Property Get MENUPAI() As Variant
    MENUPAI = arrMENUPAI
End Property

Public Property Let MENUPAI(ByVal vData As Variant)
    arrMENUPAI = vData
End Property

Public Property Get LIBPEDFOTSN() As Integer
    LIBPEDFOTSN = intLIBPEDFOTSN
End Property

Public Property Let LIBPEDFOTSN(ByVal vData As Integer)
    intLIBPEDFOTSN = vData
End Property

Public Property Get LIBPEDBLOQSN() As Integer
    LIBPEDBLOQSN = intLIBPEDBLOQSN
End Property

Public Property Let LIBPEDBLOQSN(ByVal vData As Integer)
    intLIBPEDBLOQSN = vData
End Property

Public Property Get LIQPEDSN() As Integer
    LIQPEDSN = intLIQPEDSN
End Property

Public Property Let LIQPEDSN(ByVal vData As Integer)
    intLIQPEDSN = vData
End Property

Public Property Get REPEDSN() As Integer
    REPEDSN = intREPEDSN
End Property

Public Property Let REPEDSN(ByVal vData As Integer)
    intREPEDSN = vData
End Property

Public Property Get LIBCOMERCIAL() As Integer
    LIBCOMERCIAL = intLIBCOMERCIAL
End Property

Public Property Let LIBCOMERCIAL(ByVal vData As Integer)
    intLIBCOMERCIAL = vData
End Property

Public Property Get LIBFINANCEIRO() As Integer
    LIBFINANCEIRO = intLIBFINANCEIRO
End Property

Public Property Let LIBFINANCEIRO(ByVal vData As Integer)
    intLIBFINANCEIRO = vData
End Property

Public Property Get PERMBLOQPED() As Integer
    PERMBLOQPED = intPERMBLOQPED
End Property

Public Property Let PERMBLOQPED(ByVal vData As Integer)
    intPERMBLOQPED = vData
End Property

Public Property Get FUNCAO() As Integer
    FUNCAO = intFUNCAO
End Property

Public Property Let FUNCAO(ByVal vData As Integer)
    intFUNCAO = vData
End Property

Public Property Get SECAO() As Integer
    SECAO = intSECAO
End Property

Public Property Let SECAO(ByVal vData As Integer)
    intSECAO = vData
End Property

Public Property Get SETOR() As Integer
    SETOR = intSETOR
End Property

Public Property Let SETOR(ByVal vData As Integer)
    intSETOR = vData
End Property

Public Property Get EMAIL() As Variant
    EMAIL = arrEMAIL
End Property

Public Property Let EMAIL(ByVal vData As Variant)
    arrEMAIL = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADSENHAP.cCaminho = cCaminho
    frmCADSENHAP.Linha = Linha
    frmCADSENHAP.FILIAL = iFilial
    frmCADSENHAP.strAcesso = strAcesso
    frmCADSENHAP.Show vbModal
End Sub

Public Function Pesq_CadSenha() As Boolean

   Pesq_CadSenha = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_USUARIO " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
   
      intSENCODIGO = BREC!SGI_CODIGO
      strSENNOME = BREC!SGI_NOME
      If Not IsNull(BREC!SGI_DEPARTAMENTO) Then intSENDEPTO = BREC!SGI_DEPARTAMENTO
      strSENSENHA = BREC!SGI_SENHA
   
      Pesq_CadSenha = True
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
     
     Dim I As Integer
     
     GRAVA = False
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
    
        '' Grava Usuario
        sSql = "Insert Into SGI_USUARIO (SGI_FILIAL,SGI_CODIGO,SGI_NOME,SGI_DEPARTAMENTO,SGI_SENHA,SGI_ACESSO,SGI_CODFUNCAO,SGI_CODSETOR,SGI_CODSECAO,SGI_PERMBLOQPED" & vbCrLf
        sSql = sSql & "                           ,SGI_LIBFINSN" & vbCrLf
        sSql = sSql & "                           ,SGI_LIBCOMSN" & vbCrLf
        sSql = sSql & "                           ,SGI_REPEDSN" & vbCrLf
        sSql = sSql & "                           ,SGI_LIQPEDSN" & vbCrLf
        sSql = sSql & "                           ,SGI_LIBPEDBLOQSN" & vbCrLf
        sSql = sSql & "                           ,SGI_LIBPEDFOTSN" & vbCrLf
        sSql = sSql & "                           ,SGI_NOVO" & vbCrLf
        sSql = sSql & "                           ,SGI_DESABPROD" & vbCrLf
        sSql = sSql & "                           ,SGI_PERMLIBFOT" & vbCrLf
        sSql = sSql & "                           ,SGI_PERMFAT10POR" & vbCrLf
        sSql = sSql & "                           ,SGI_LIBPDATAPCOTA" & vbCrLf
        sSql = sSql & "                           ,SGI_PERMFATROTDIFSN" & vbCrLf
        sSql = sSql & "                           ,SGI_PVCLIE" & vbCrLf
        sSql = sSql & "                           ,SGI_EVENDEDOR" & vbCrLf
        sSql = sSql & "                           ,SGI_ATIVO" & vbCrLf
        sSql = sSql & "                           ,SGI_BLOQCRED" & vbCrLf
        sSql = sSql & "                           ,SGI_PERMALTPEDFAT" & vbCrLf
        sSql = sSql & "                           ,SGI_MOP" & vbCrLf
        sSql = sSql & "                           ,SGI_PERMEXCPED" & vbCrLf
        
        sSql = sSql & "                  ) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intSENCODIGO
        sSql = sSql & "                           ,'" & strSENNOME & "'" & vbCrLf
        sSql = sSql & "                           ,0" & vbCrLf
        sSql = sSql & "                           ,'" & strSENSENHA & "'" & vbCrLf
        sSql = sSql & "                           ," & intSENACESSO & vbCrLf
        sSql = sSql & "                           ," & intFUNCAO & vbCrLf
        sSql = sSql & "                           ," & intSETOR & vbCrLf
        sSql = sSql & "                           ," & intSECAO & vbCrLf
        sSql = sSql & "                           ," & intPERMBLOQPED & vbCrLf
        sSql = sSql & "                           ," & intLIBFINANCEIRO & vbCrLf
        sSql = sSql & "                           ," & intLIBCOMERCIAL & vbCrLf
        sSql = sSql & "                           ," & intREPEDSN & vbCrLf
        sSql = sSql & "                           ," & intLIQPEDSN & vbCrLf
        sSql = sSql & "                           ," & intLIBPEDBLOQSN & vbCrLf
        sSql = sSql & "                           ," & intLIBPEDFOTSN & vbCrLf
        sSql = sSql & "                           ," & intNOVO & vbCrLf
        sSql = sSql & "                           ," & intDESABPROD & vbCrLf
        sSql = sSql & "                           ," & intPERMLIBFOT & vbCrLf
        sSql = sSql & "                           ," & intPERMFAT10POR & vbCrLf
        sSql = sSql & "                           ," & intLIBPDATAPCOTA & vbCrLf
        sSql = sSql & "                           ," & intPERMFATROTDIFSN & vbCrLf
        sSql = sSql & "                           ," & intPVCLIE & vbCrLf
        sSql = sSql & "                           ," & intEVENDEDOR & vbCrLf
        sSql = sSql & "                           ," & intATIVO & vbCrLf
        sSql = sSql & "                           ," & intBLOQCRED & vbCrLf
        sSql = sSql & "                           ," & intPERMALTPEDFAT & vbCrLf
        sSql = sSql & "                           ," & intMOP & vbCrLf
        sSql = sSql & "                           ," & intPermExcPedSN & vbCrLf
        
        sSql = sSql & "                           )"
                
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrEMAIL) Then
           For I = 1 To UBound(arrEMAIL)
                sSql = "Insert into SGI_CADEMAILUSUARIO (SGI_FILIAL,SGI_CODIGO,SGI_EMAIL,SGI_ATIVO) " & vbCrLf
                sSql = sSql & "                         Values (" & vbCrLf
                sSql = sSql & "                                  " & intFilial & vbCrLf
                sSql = sSql & "                                 ," & intSENCODIGO & vbCrLf
                sSql = sSql & "                                 ,'" & arrEMAIL(I, 1) & "'" & vbCrLf
                sSql = sSql & "                                 ," & arrEMAIL(I, 2) & vbCrLf
                sSql = sSql & "                                 )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
           Next I
        End If
     
MENU_NOVO:

        If intNOVO = 0 Then
        
            '' ============================
            '' Menu Pai
            If IsArray(arrMENUPAI) Then
                For I = 1 To UBound(arrMENUPAI)
                    sSql = ""
                    
                    sSql = "Insert Into SGI_MENUP (" & vbCrLf
                    sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                      ,SGI_TEXTO" & vbCrLf
                    sSql = sSql & "                      ,SGI_TIPO" & vbCrLf
                    sSql = sSql & "                      ,SGI_CIGLA" & vbCrLf
                    sSql = sSql & "                      ,SGI_CIGLA2" & vbCrLf
                    sSql = sSql & "                      ,SGI_MODULO" & vbCrLf
                    sSql = sSql & "                      ,SGI_ACESSO" & vbCrLf
                    sSql = sSql & "                      ,SGI_NIVEL" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODGER" & vbCrLf
                    sSql = sSql & "                      ,SGI_DESNIV" & vbCrLf
                    sSql = sSql & "                      ,SGI_DEPTO" & vbCrLf
                    sSql = sSql & "                      ,SGI_IMAGEN" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODUSUARIO" & vbCrLf
                    sSql = sSql & "                      ,SGI_ATIVO" & vbCrLf
                    sSql = sSql & "            ) Values (" & vbCrLf
                    sSql = sSql & "                      " & arrMENUPAI(I, 1) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUPAI(I, 2) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUPAI(I, 3) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUPAI(I, 4) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUPAI(I, 5) & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,0" & vbCrLf
                    sSql = sSql & "                     ,0" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ," & intSENCODIGO & vbCrLf
                    sSql = sSql & "                     ," & arrMENUPAI(I, 6) & vbCrLf
                    sSql = sSql & "                     )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    
                Next I
            End If
        
            '' ============================
            '' Menu Filho
            If IsArray(arrMENUFILHO) Then
                For I = 1 To UBound(arrMENUFILHO)
                
                    sSql = ""
                    
                    sSql = "Insert Into SGI_MENUP (" & vbCrLf
                    sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                      ,SGI_TEXTO" & vbCrLf
                    sSql = sSql & "                      ,SGI_TIPO" & vbCrLf
                    sSql = sSql & "                      ,SGI_CIGLA" & vbCrLf
                    sSql = sSql & "                      ,SGI_CIGLA2" & vbCrLf
                    sSql = sSql & "                      ,SGI_MODULO" & vbCrLf
                    sSql = sSql & "                      ,SGI_ACESSO" & vbCrLf
                    sSql = sSql & "                      ,SGI_NIVEL" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODGER" & vbCrLf
                    sSql = sSql & "                      ,SGI_DESNIV" & vbCrLf
                    sSql = sSql & "                      ,SGI_DEPTO" & vbCrLf
                    sSql = sSql & "                      ,SGI_IMAGEN" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODUSUARIO" & vbCrLf
                    sSql = sSql & "                      ,SGI_ATIVO" & vbCrLf
                    sSql = sSql & "            ) Values (" & vbCrLf
                    sSql = sSql & "                      " & arrMENUFILHO(I, 1) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUFILHO(I, 2) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUFILHO(I, 3) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUFILHO(I, 4) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUFILHO(I, 5) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUFILHO(I, 6) & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,0" & vbCrLf
                    sSql = sSql & "                     ,0" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ," & intSENCODIGO & vbCrLf
                    sSql = sSql & "                     ," & arrMENUFILHO(I, 7) & vbCrLf
                    sSql = sSql & "                     )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                Next I
            End If
     
            '' ============================
            '' Menu Neto
            If IsArray(arrMENUNETO) Then
                For I = 1 To UBound(arrMENUNETO)
                
                    sSql = ""
                    
                    sSql = "Insert Into SGI_MENUP (" & vbCrLf
                    sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                      ,SGI_TEXTO" & vbCrLf
                    sSql = sSql & "                      ,SGI_TIPO" & vbCrLf
                    sSql = sSql & "                      ,SGI_CIGLA" & vbCrLf
                    sSql = sSql & "                      ,SGI_CIGLA2" & vbCrLf
                    sSql = sSql & "                      ,SGI_MODULO" & vbCrLf
                    sSql = sSql & "                      ,SGI_ACESSO" & vbCrLf
                    sSql = sSql & "                      ,SGI_NIVEL" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODGER" & vbCrLf
                    sSql = sSql & "                      ,SGI_DESNIV" & vbCrLf
                    sSql = sSql & "                      ,SGI_DEPTO" & vbCrLf
                    sSql = sSql & "                      ,SGI_IMAGEN" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODUSUARIO" & vbCrLf
                    sSql = sSql & "                      ,SGI_ATIVO" & vbCrLf
                    sSql = sSql & "            ) Values (" & vbCrLf
                    sSql = sSql & "                      " & arrMENUNETO(I, 1) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUNETO(I, 2) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUNETO(I, 3) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUNETO(I, 4) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUNETO(I, 5) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUNETO(I, 6) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUNETO(I, 7) & vbCrLf
                    sSql = sSql & "                     ," & arrMENUNETO(I, 9) & vbCrLf
                    sSql = sSql & "                     ,0" & vbCrLf
                    sSql = sSql & "                     ,0" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ,Null" & vbCrLf
                    sSql = sSql & "                     ," & intSENCODIGO & vbCrLf
                    sSql = sSql & "                     ," & arrMENUNETO(I, 8) & vbCrLf
                    sSql = sSql & "                     )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                Next I
            End If
        
        End If
     
     ElseIf strOPER = "A" Then
     
        If intNOVO = 0 Then
        
            sSql = "Delete From SGI_MENUP " & vbCrLf
            sSql = sSql & "      Where " & vbCrLf
            sSql = sSql & "            SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "        And SGI_CODUSUARIO = " & intSENCODIGO
        
            BGRV.CommandText = sSql
            BGRV.Execute
        
        End If
     
        sSql = "Delete from SGI_CADEMAILUSUARIO "
        sSql = sSql & "      Where SGI_CODIGO = " & intSENCODIGO
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrEMAIL) Then
           For I = 1 To UBound(arrEMAIL)
                sSql = "Insert into SGI_CADEMAILUSUARIO (SGI_FILIAL,SGI_CODIGO,SGI_EMAIL,SGI_ATIVO) " & vbCrLf
                sSql = sSql & "                         Values (" & vbCrLf
                sSql = sSql & "                                  " & intFilial & vbCrLf
                sSql = sSql & "                                 ," & intSENCODIGO & vbCrLf
                sSql = sSql & "                                 ,'" & arrEMAIL(I, 1) & "'" & vbCrLf
                sSql = sSql & "                                 ," & arrEMAIL(I, 2) & vbCrLf
                sSql = sSql & "                                 )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
           Next I
        End If
        
        '' Grava Usuario
        sSql = "Update SGI_USUARIO Set " & vbCrLf
        sSql = sSql & "               SGI_NOME            = '" & strSENNOME & "'" & vbCrLf
        sSql = sSql & "              ,SGI_DEPARTAMENTO    = 0" & vbCrLf
        sSql = sSql & "              ,SGI_SENHA           = '" & strSENSENHA & "'" & vbCrLf
        sSql = sSql & "              ,SGI_ACESSO          =  " & intSENACESSO & vbCrLf
        sSql = sSql & "              ,SGI_CODFUNCAO       =  " & intFUNCAO & vbCrLf
        sSql = sSql & "              ,SGI_CODSETOR        =  " & intSETOR & vbCrLf
        sSql = sSql & "              ,SGI_CODSECAO        =  " & intSECAO & vbCrLf
        sSql = sSql & "              ,SGI_PERMBLOQPED     =  " & intPERMBLOQPED & vbCrLf
        sSql = sSql & "              ,SGI_LIBFINSN        =  " & intLIBFINANCEIRO & vbCrLf
        sSql = sSql & "              ,SGI_LIBCOMSN        =  " & intLIBCOMERCIAL & vbCrLf
        sSql = sSql & "              ,SGI_REPEDSN         =  " & intREPEDSN & vbCrLf
        sSql = sSql & "              ,SGI_LIQPEDSN        =  " & intLIQPEDSN & vbCrLf
        sSql = sSql & "              ,SGI_LIBPEDBLOQSN    =  " & intLIBPEDBLOQSN & vbCrLf
        sSql = sSql & "              ,SGI_LIBPEDFOTSN     =  " & intLIBPEDFOTSN & vbCrLf
        sSql = sSql & "              ,SGI_NOVO            =  " & intNOVO & vbCrLf
        sSql = sSql & "              ,SGI_DESABPROD       =  " & intDESABPROD & vbCrLf
        sSql = sSql & "              ,SGI_PERMLIBFOT      =  " & intPERMLIBFOT & vbCrLf
        sSql = sSql & "              ,SGI_PERMFAT10POR    =  " & intPERMFAT10POR & vbCrLf
        sSql = sSql & "              ,SGI_LIBPDATAPCOTA   =  " & intLIBPDATAPCOTA & vbCrLf
        sSql = sSql & "              ,SGI_PERMFATROTDIFSN =  " & intPERMFATROTDIFSN & vbCrLf
        sSql = sSql & "              ,SGI_PVCLIE          =  " & intPVCLIE & vbCrLf
        sSql = sSql & "              ,SGI_EVENDEDOR       =  " & intEVENDEDOR & vbCrLf
        sSql = sSql & "              ,SGI_ATIVO           =  " & intATIVO & vbCrLf
        sSql = sSql & "              ,SGI_BLOQCRED        =  " & intBLOQCRED & vbCrLf
        sSql = sSql & "              ,SGI_PERMALTPEDFAT   =  " & intPERMALTPEDFAT & vbCrLf
        sSql = sSql & "              ,SGI_MOP             =  " & intMOP & vbCrLf
        sSql = sSql & "              ,SGI_PERMEXCPED      =  " & intPermExcPedSN & vbCrLf
        
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intSENCODIGO
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        If intNOVO = 0 Then GoTo MENU_NOVO
     
     ElseIf strOPER = "E" Then
     
        If intNOVO = 0 Then
        
            sSql = "Delete From SGI_MENUP " & vbCrLf
            sSql = sSql & "      Where " & vbCrLf
            sSql = sSql & "            SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "        And SGI_CODUSUARIO = " & intSENCODIGO
            
            BGRV.CommandText = sSql
            BGRV.Execute
        
        End If
     
        sSql = "Delete from SGI_CADEMAILUSUARIO "
        sSql = sSql & "      Where SGI_CODIGO = " & intSENCODIGO
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Usuario
        sSql = "Delete from SGI_USUARIO "
        sSql = sSql & "      Where SGI_CODIGO = " & intSENCODIGO
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
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description & "Iten : " & I, strOPER, sSql)
     Set objErro = Nothing

End Function

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get SENCODIGO() As Integer
    SENCODIGO = intSENCODIGO
End Property

Public Property Let SENCODIGO(ByVal vData As Integer)
    intSENCODIGO = vData
End Property

Public Property Get SENNOME() As String
    SENNOME = strSENNOME
End Property

Public Property Let SENNOME(ByVal vData As String)
    strSENNOME = vData
End Property

Public Property Get SENDEPTO() As Integer
    SENDEPTO = intSENDEPTO
End Property

Public Property Let SENDEPTO(ByVal vData As Integer)
    intSENDEPTO = vData
End Property

Public Property Get SENSENHA() As String
    SENSENHA = strSENSENHA
End Property

Public Property Let SENSENHA(ByVal vData As String)
    strSENSENHA = vData
End Property

Public Property Get SENACESSO() As Integer
    SENACESSO = intSENACESSO
End Property

Public Property Let SENACESSO(ByVal vData As Integer)
    intSENACESSO = vData
End Property


Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim qtdQTDREGS As Long
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_USUARIO " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_CODIGO = " & intSENCODIGO
   sSql = sSql + "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
   
      intSENCODIGO = BREC!SGI_CODIGO
      strSENNOME = BREC!SGI_NOME
      If Not IsNull(BREC!SGI_DEPARTAMENTO) Then intSENDEPTO = BREC!SGI_DEPARTAMENTO
      strSENSENHA = BREC!SGI_SENHA
      intSENACESSO = BREC!SGI_ACESSO
      intFUNCAO = BREC!SGI_CODFUNCAO
      intSETOR = BREC!SGI_CODSETOR
      intSECAO = BREC!SGI_CODSECAO
      intPERMBLOQPED = BREC!SGI_PERMBLOQPED
      intLIBFINANCEIRO = BREC!SGI_LIBFINSN
      intLIBCOMERCIAL = BREC!SGI_LIBCOMSN
      intREPEDSN = BREC!SGI_REPEDSN
      intLIQPEDSN = BREC!SGI_LIQPEDSN
      intLIBPEDBLOQSN = BREC!SGI_LIBPEDBLOQSN
      intLIBPEDFOTSN = BREC!SGI_LIBPEDFOTSN
      intNOVO = BREC!SGI_NOVO
      intDESABPROD = BREC!SGI_DESABPROD
      intPERMLIBFOT = BREC!SGI_PERMLIBFOT
      intPERMFAT10POR = BREC!SGI_PERMFAT10POR
      intLIBPDATAPCOTA = BREC!SGI_LIBPDATAPCOTA
      intPERMFATROTDIFSN = BREC!SGI_PERMFATROTDIFSN
      intPVCLIE = BREC!SGI_PVCLIE
      intEVENDEDOR = BREC!SGI_EVENDEDOR
      If Not IsNull(BREC!SGI_ATIVO) Then intATIVO = BREC!SGI_ATIVO
      intBLOQCRED = BREC!SGI_BLOQCRED
      intPERMALTPEDFAT = BREC!SGI_PERMALTPEDFAT
      intMOP = BREC!SGI_MOP
      intPermExcPedSN = BREC!SGI_PERMEXCPED
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   sSql = ""
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADEMAILUSUARIO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO = " & intSENCODIGO
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC2.EOF Then
      qtdQTDREGS = 0
      Do While Not BREC2.EOF
         qtdQTDREGS = qtdQTDREGS + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrEMAIL(1 To qtdQTDREGS, 1 To 2) As Variant
      BREC2.MoveFirst
      qtdQTDREGS = 1
      Do While Not BREC2.EOF
         arrEMAIL(qtdQTDREGS, 1) = BREC2!SGI_EMAIL
         arrEMAIL(qtdQTDREGS, 2) = BREC2!SGI_ATIVO
         qtdQTDREGS = (qtdQTDREGS + 1)
         BREC2.MoveNext
      Loop
   End If
   BREC2.Close
   
   '' ========================
   '' Menu Pai
   sSql = ""
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_MENUP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL     = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODUSUARIO = " & intSENCODIGO & vbCrLf
   sSql = sSql & "   And SGI_TIPO       = 'P'" & vbCrLf
   sSql = sSql & "Order By SGI_CODIGO"
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC2.EOF Then
      qtdQTDREGS = 0
      Do While Not BREC2.EOF
         qtdQTDREGS = qtdQTDREGS + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrMENUPAI(1 To qtdQTDREGS, 1 To 6) As String
      BREC2.MoveFirst
      qtdQTDREGS = 1
      Do While Not BREC2.EOF
         arrMENUPAI(qtdQTDREGS, 1) = BREC2!SGI_FILIAL
         arrMENUPAI(qtdQTDREGS, 2) = BREC2!SGI_CODIGO
         arrMENUPAI(qtdQTDREGS, 3) = BREC2!SGI_TEXTO
         arrMENUPAI(qtdQTDREGS, 4) = BREC2!SGI_TIPO
         arrMENUPAI(qtdQTDREGS, 5) = BREC2!SGI_CIGLA
         arrMENUPAI(qtdQTDREGS, 6) = BREC2!SGI_ATIVO
         qtdQTDREGS = (qtdQTDREGS + 1)
         BREC2.MoveNext
      Loop
   End If
   BREC2.Close
   
   '' ========================
   '' Menu Filho
   sSql = ""
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_MENUP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL     = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODUSUARIO = " & intSENCODIGO & vbCrLf
   sSql = sSql & "   And SGI_TIPO       = 'S'" & vbCrLf
   sSql = sSql & "Order By SGI_CODIGO"
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC2.EOF Then
      qtdQTDREGS = 0
      Do While Not BREC2.EOF
         qtdQTDREGS = qtdQTDREGS + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrMENUFILHO(1 To qtdQTDREGS, 1 To 7) As String
      BREC2.MoveFirst
      qtdQTDREGS = 1
      Do While Not BREC2.EOF
         arrMENUFILHO(qtdQTDREGS, 1) = BREC2!SGI_FILIAL
         arrMENUFILHO(qtdQTDREGS, 2) = BREC2!SGI_CODIGO
         arrMENUFILHO(qtdQTDREGS, 3) = BREC2!SGI_TEXTO
         arrMENUFILHO(qtdQTDREGS, 4) = BREC2!SGI_TIPO
         arrMENUFILHO(qtdQTDREGS, 5) = BREC2!SGI_CIGLA
         arrMENUFILHO(qtdQTDREGS, 6) = BREC2!SGI_CIGLA2
         arrMENUFILHO(qtdQTDREGS, 7) = BREC2!SGI_ATIVO
         qtdQTDREGS = (qtdQTDREGS + 1)
         BREC2.MoveNext
      Loop
   End If
   BREC2.Close
   
   
   '' ========================
   '' Menu Neto
   sSql = ""
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_MENUP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL     = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODUSUARIO = " & intSENCODIGO & vbCrLf
   sSql = sSql & "   And SGI_TIPO       = 'M'" & vbCrLf
   sSql = sSql & "Order By SGI_CODIGO"
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC2.EOF Then
      qtdQTDREGS = 0
      Do While Not BREC2.EOF
         qtdQTDREGS = qtdQTDREGS + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrMENUNETO(1 To qtdQTDREGS, 1 To 9) As String
      BREC2.MoveFirst
      qtdQTDREGS = 1
      Do While Not BREC2.EOF
      
         arrMENUNETO(qtdQTDREGS, 1) = BREC2!SGI_FILIAL
         arrMENUNETO(qtdQTDREGS, 2) = BREC2!SGI_CODIGO
         arrMENUNETO(qtdQTDREGS, 3) = BREC2!SGI_TEXTO
         arrMENUNETO(qtdQTDREGS, 4) = BREC2!SGI_TIPO
         arrMENUNETO(qtdQTDREGS, 5) = BREC2!SGI_CIGLA
         arrMENUNETO(qtdQTDREGS, 6) = BREC2!SGI_CIGLA2
         arrMENUNETO(qtdQTDREGS, 7) = BREC2!SGI_MODULO
         arrMENUNETO(qtdQTDREGS, 8) = BREC2!SGI_ACESSO
         arrMENUNETO(qtdQTDREGS, 9) = BREC2!SGI_ATIVO
         qtdQTDREGS = (qtdQTDREGS + 1)
         BREC2.MoveNext
      Loop
   End If
   BREC2.Close
   
End Function

Public Sub PreencheComboAcesso(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MN.SGI_CODGER  " & vbCrLf
    sSql = sSql & "      ,MN.SGI_DESNIV  " & vbCrLf
    sSql = sSql & "      ,DP.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_MENUP    MN" & vbCrLf
    sSql = sSql & "      ,SGI_CADDEPTO DP" & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    ''sSql = sSql & "      MN.SGI_CODGER   > 0 " & vbCrLf
    sSql = sSql & "      MN.SGI_FILIAL   = " & intFilial & vbCrLf
    sSql = sSql & "  And DP.SGI_FILIAL   = MN.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And DP.SGI_CODDEPTO = MN.SGI_DEPTO  " & vbCrLf
    
    sSql = sSql & "Group by MN.SGI_CODGER,MN.SGI_DESNIV,DP.SGI_DESCRICAO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Format(Str(BREC!SGI_CODGER), "00") & " - " & Trim(BREC!SGI_DESNIV) & " - " & Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODGER
       BREC.MoveNext
    Loop
    
    If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
    BREC.Close

End Sub

 
Public Sub PreencheComboSetor(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "      SGI_CADSETOR " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL   = " & intFilial & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

Public Sub PreencheComboSecao(Combo As Variant, lngCODSETOR As Long)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "      SE.*  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "      SGI_CADITESET IT " & vbCrLf
    sSql = sSql & "     ,SGI_CADSECAO  SE " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      IT.SGI_FILIAL   = " & intFilial & vbCrLf
    sSql = sSql & "  And IT.SGI_CODIGO   = " & lngCODSETOR & vbCrLf
    sSql = sSql & "  And SE.SGI_FILIAL   = IT.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SE.SGI_CODIGO   = IT.SGI_CODSECAO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

Public Sub PreencheComboFuncao(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "      SGI_CADFUNCAO " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL   = " & intFilial & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    BREC.Close

End Sub

Public Function PreenchComboAtivo() As String
        
    PreenchComboAtivo = PreenchComboAtivo & "|#" & 1 & ";SIM"
    PreenchComboAtivo = PreenchComboAtivo & "|#" & 0 & ";NÃO"

End Function

