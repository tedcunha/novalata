VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intMENCODIGO As Long
Dim intMENNIVEL  As Integer
Dim varMENTIPACS As Variant
Dim trrMENMENU   As Variant
Dim strMENDESNIV As String
Dim intMENDEPTO  As Long
Dim iCodigo      As Long
Dim sAcesso      As String
Dim sAcesso2     As String
Dim intFilial    As Long

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get MENCODIGO() As Integer
    MENCODIGO = intMENCODIGO
End Property

Public Property Let MENCODIGO(ByVal vData As Integer)
    intMENCODIGO = vData
End Property

Public Property Get MENNIVEL() As Integer
    MENNIVEL = intMENNIVEL
End Property

Public Property Let MENNIVEL(ByVal vData As Integer)
    intMENNIVEL = vData
End Property

Public Property Get MENTIPACS() As Variant
    MENTIPACS = varMENTIPACS
End Property

Public Property Let MENTIPACS(ByVal vData As Variant)
    varMENTIPACS = vData
End Property

Public Property Get MENMENU() As Variant
    MENMENU = trrMENMENU
End Property

Public Property Let MENMENU(ByVal vData As Variant)
    trrMENMENU = vData
End Property

Public Property Get MENDESNIV() As String
    MENDESNIV = strMENDESNIV
End Property

Public Property Let MENDESNIV(ByVal vData As String)
    strMENDESNIV = vData
End Property

Public Property Get MENDEPTO() As Integer
    MENDEPTO = intMENDEPTO
End Property

Public Property Let MENDEPTO(ByVal vData As Integer)
    intMENDEPTO = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strACESSO As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADMENUP.cCaminho = cCaminho
    frmCADMENUP.Linha = Linha
    frmCADMENUP.FILIAL = iFilial
    frmCADMENUP.strACESSO = strACESSO
    frmCADMENUP.Show vbModal
End Sub

Public Function Pesq_Menu() As Boolean

   Pesq_Menu = False
   

   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_MENUP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then Pesq_Menu = True
   
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
          sSql = sSql + "                                             " & intFilial & vbCrLf
          sSql = sSql + "                                            ,1" & vbCrLf
          sSql = sSql + "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql + "                                          )" & vbCrLf
          
       ElseIf BREC!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql + "           SGI_NUMERO = " & BREC!SGI_NUMERO & vbCrLf
          sSql = sSql + "         Where " & vbCrLf
          sSql = sSql + "               SGI_MODULO = '" & sModulo & "'" & vbCrLf
          sSql = sSql + "           And SGI_FILIAL = " & intFilial & vbCrLf
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
    End If
    
    BREC.Close
    
End Function


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava

     GRAVA = False
     
     Dim I As Integer
     Dim j As Integer
     Dim K As Integer
     
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
        
Operacao_Altera:

        sAcesso = ""
        For I = 1 To UBound(varMENTIPACS)
            sAcesso = sAcesso + Mid(varMENTIPACS(I), 1, 1)
        Next I
        
        For I = 1 To UBound(arrMenu)
               
               '' Grava Usuario
               sSql = "Insert Into SGI_MENUP (SGI_FILIAL" & vbCrLf
               sSql = sSql & "               ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "               ,SGI_TEXTO" & vbCrLf
               sSql = sSql & "               ,SGI_TIPO" & vbCrLf
               sSql = sSql & "               ,SGI_CIGLA" & vbCrLf
               sSql = sSql & "               ,SGI_CIGLA2" & vbCrLf
               sSql = sSql & "               ,SGI_MODULO" & vbCrLf
               sSql = sSql & "               ,SGI_ACESSO" & vbCrLf
               sSql = sSql & "               ,SGI_NIVEL" & vbCrLf
               sSql = sSql & "               ,SGI_CODGER" & vbCrLf
               sSql = sSql & "               ,SGI_DESNIV" & vbCrLf
               sSql = sSql & "               ,SGI_DEPTO)" & vbCrLf
               sSql = sSql & "       Values (" & vbCrLf
               sSql = sSql & "                " & intFilial & vbCrLf
               sSql = sSql & "               ," & arrMenu(I).intCODIGO & vbCrLf
               sSql = sSql & "               ,'" & Trim(arrMenu(I).strTEXTO) & "'" & vbCrLf
               sSql = sSql & "               ,'" & Trim(arrMenu(I).strTIPO) & "'" & vbCrLf
               sSql = sSql & "               ,'" & IIf(Len(Trim(arrMenu(I).strCIGLA)) > 0, Trim(arrMenu(I).strCIGLA), "Null") & "'" & vbCrLf
               sSql = sSql & "               ,'" & IIf(Len(Trim(arrMenu(I).strCIGLA2)) > 0, Trim(arrMenu(I).strCIGLA2), "Null") & "'" & vbCrLf
               sSql = sSql & "               ,'" & Trim(arrMenu(I).strModulo) & "'" & vbCrLf
               sSql = sSql & "               ,'" & sAcesso & "'" & vbCrLf
               sSql = sSql & "               ," & intMENNIVEL & vbCrLf
               sSql = sSql & "               ," & intMENCODIGO & vbCrLf
               sSql = sSql & "               ,'" & strMENDESNIV & "'" & vbCrLf
               If intMENDEPTO > 0 Then
                  sSql = sSql & "               ," & intMENDEPTO & vbCrLf
               Else
                  sSql = sSql & "               ,Null" & vbCrLf
               End If
               sSql = sSql & "               )"
        
               BGRV.CommandText = sSql
               BGRV.Execute
               
               If arrMenu(I).intQTDNIVEL > 0 Then
               
                  For j = 1 To UBound(arrMenu(I).arrNIVEL_S)
               
                       '' Grava Usuario
                       sSql = "Insert Into SGI_MENUP (SGI_FILIAL" & vbCrLf
                       sSql = sSql & "               ,SGI_CODIGO" & vbCrLf
                       sSql = sSql & "               ,SGI_TEXTO" & vbCrLf
                       sSql = sSql & "               ,SGI_TIPO" & vbCrLf
                       sSql = sSql & "               ,SGI_CIGLA" & vbCrLf
                       sSql = sSql & "               ,SGI_CIGLA2" & vbCrLf
                       sSql = sSql & "               ,SGI_MODULO" & vbCrLf
                       sSql = sSql & "               ,SGI_ACESSO" & vbCrLf
                       sSql = sSql & "               ,SGI_NIVEL" & vbCrLf
                       sSql = sSql & "               ,SGI_CODGER" & vbCrLf
                       sSql = sSql & "               ,SGI_DESNIV" & vbCrLf
                       sSql = sSql & "               ,SGI_DEPTO)" & vbCrLf
                       sSql = sSql & "       Values (" & vbCrLf
                       sSql = sSql & "                " & intFilial & vbCrLf
                       sSql = sSql & "               ," & arrMenu(I).arrNIVEL_S(j).intCODIGO & vbCrLf
                       sSql = sSql & "               ,'" & Trim(arrMenu(I).arrNIVEL_S(j).strTEXTO) & "'" & vbCrLf
                       sSql = sSql & "               ,'" & Trim(arrMenu(I).arrNIVEL_S(j).strTIPO) & "'" & vbCrLf
                       sSql = sSql & "               ,'" & IIf(Len(Trim(arrMenu(I).arrNIVEL_S(j).strCIGLA)) > 0, Trim(arrMenu(I).arrNIVEL_S(j).strCIGLA), "Null") & "'" & vbCrLf
                       sSql = sSql & "               ,'" & IIf(Len(Trim(arrMenu(I).arrNIVEL_S(j).strCIGLA2)) > 0, Trim(arrMenu(I).arrNIVEL_S(j).strCIGLA2), "Null") & "'" & vbCrLf
                       sSql = sSql & "               ,'" & Trim(arrMenu(I).arrNIVEL_S(j).strModulo) & "'" & vbCrLf
                       sSql = sSql & "               ,'" & sAcesso & "'" & vbCrLf
                       sSql = sSql & "               ," & intMENNIVEL & vbCrLf
                       sSql = sSql & "               ," & intMENCODIGO & vbCrLf
                       sSql = sSql & "               ,'" & strMENDESNIV & "'" & vbCrLf
                       If intMENDEPTO > 0 Then
                          sSql = sSql & "               ," & intMENDEPTO & vbCrLf
                       Else
                          sSql = sSql & "               ,Null" & vbCrLf
                       End If
                       sSql = sSql & "               )"
                
                       BGRV.CommandText = sSql
                       BGRV.Execute
                       
                       If arrMenu(I).arrNIVEL_S(j).intQTDNIVEL > 0 Then
                          For K = 1 To UBound(arrMenu(I).arrNIVEL_S(j).arrNIVEL_M)
                          
                               '' Grava Usuario
                               sSql = "Insert Into SGI_MENUP (SGI_FILIAL" & vbCrLf
                               sSql = sSql & "               ,SGI_CODIGO" & vbCrLf
                               sSql = sSql & "               ,SGI_TEXTO" & vbCrLf
                               sSql = sSql & "               ,SGI_TIPO" & vbCrLf
                               sSql = sSql & "               ,SGI_CIGLA" & vbCrLf
                               sSql = sSql & "               ,SGI_CIGLA2" & vbCrLf
                               sSql = sSql & "               ,SGI_MODULO" & vbCrLf
                               sSql = sSql & "               ,SGI_ACESSO" & vbCrLf
                               sSql = sSql & "               ,SGI_NIVEL" & vbCrLf
                               sSql = sSql & "               ,SGI_CODGER" & vbCrLf
                               sSql = sSql & "               ,SGI_DESNIV" & vbCrLf
                               sSql = sSql & "               ,SGI_DEPTO)" & vbCrLf
                               sSql = sSql & "       Values (" & vbCrLf
                               sSql = sSql & "                " & intFilial & vbCrLf
                               sSql = sSql & "               ," & arrMenu(I).arrNIVEL_S(j).arrNIVEL_M(K).intCODIGO & vbCrLf
                               sSql = sSql & "               ,'" & Trim(arrMenu(I).arrNIVEL_S(j).arrNIVEL_M(K).strTEXTO) & "'" & vbCrLf
                               sSql = sSql & "               ,'" & Trim(arrMenu(I).arrNIVEL_S(j).arrNIVEL_M(K).strTIPO) & "'" & vbCrLf
                               sSql = sSql & "               ,'" & IIf(Len(Trim(arrMenu(I).arrNIVEL_S(j).arrNIVEL_M(K).strCIGLA)) > 0, Trim(arrMenu(I).arrNIVEL_S(j).arrNIVEL_M(K).strCIGLA), "Null") & "'" & vbCrLf
                               sSql = sSql & "               ,'" & IIf(Len(Trim(arrMenu(I).arrNIVEL_S(j).arrNIVEL_M(K).strCIGLA2)) > 0, Trim(arrMenu(I).arrNIVEL_S(j).arrNIVEL_M(K).strCIGLA2), "Null") & "'" & vbCrLf
                               sSql = sSql & "               ,'" & Trim(arrMenu(I).arrNIVEL_S(j).arrNIVEL_M(K).strModulo) & "'" & vbCrLf
                               sSql = sSql & "               ,'" & sAcesso & "'" & vbCrLf
                               sSql = sSql & "               ," & intMENNIVEL & vbCrLf
                               sSql = sSql & "               ," & intMENCODIGO & vbCrLf
                               sSql = sSql & "               ,'" & strMENDESNIV & "'" & vbCrLf
                               If intMENDEPTO > 0 Then
                                  sSql = sSql & "               ," & intMENDEPTO & vbCrLf
                               Else
                                  sSql = sSql & "               ,Null" & vbCrLf
                               End If
                               sSql = sSql & "               )"
                        
                               BGRV.CommandText = sSql
                               BGRV.Execute
                          
                          Next K
                       End If
               
                  Next j
               
               End If
            
        Next I
     
     ElseIf strOPER = "A" Or strOPER = "E" Then
     
        '' Apaga o Menu
        sSql = "Delete from SGI_MENUP "
        sSql = sSql & "      Where SGI_CODGER = " & intMENCODIGO
        sSql = sSql & "        And SGI_FILIAL = " & intFilial & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo Operacao_Altera
     
     End If
     
     GRAVA = True
     adoBanco_Dados.CommitTrans
     
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
   
   Dim I As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_MENUP " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_CODGER = " & intMENCODIGO
   sSql = sSql + "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   Do While Not BREC.EOF
   
      intMENCODIGO = BREC!SGI_CODGER
      intMENNIVEL = BREC!SGI_NIVEL
      sAcesso = BREC!SGI_ACESSO
      intMENDEPTO = BREC!SGI_DEPTO
      
      Carrega_campos = True
      BREC.MoveNext
      
   Loop
   
   ReDim varMENTIPACS(1 To Len(Trim(sAcesso)))
   
   For I = 1 To Len(Trim(sAcesso))
       If Mid(sAcesso, I, 1) = "I" Then
          sAcesso2 = sAcesso2 & "I-INCLUSÃO"
          varMENTIPACS(I) = "I-INCLUSÃO"
       ElseIf Mid(sAcesso, I, 1) = "A" Then
          sAcesso2 = sAcesso2 & "A-ALTERAÇÃO"
          varMENTIPACS(I) = "A-ALTERAÇÃO"
       ElseIf Mid(sAcesso, I, 1) = "E" Then
          sAcesso2 = sAcesso2 & "E-EXCLUSÃO"
          varMENTIPACS(I) = "E-EXCLUSÃO"
       ElseIf Mid(sAcesso, I, 1) = "C" Then
          sAcesso2 = sAcesso2 & "C-CONSULTA"
          varMENTIPACS(I) = "C-CONSULTA"
       ElseIf Mid(sAcesso, I, 1) = "R" Then
          sAcesso2 = sAcesso2 & "R-RELATÓRIO"
          varMENTIPACS(I) = "R-RELATÓRIO"
       ElseIf Mid(sAcesso, I, 1) = "P" Then
          sAcesso2 = sAcesso2 & "P-IMPRESSÃO"
          varMENTIPACS(I) = "P-IMPRESSÃO"
       ElseIf Mid(sAcesso, I, 1) = "L" Then
          sAcesso2 = sAcesso2 & "L-LIBERA"
          varMENTIPACS(I) = "L-LIBERA"
       ElseIf Mid(sAcesso, I, 1) = "B" Then
          sAcesso2 = sAcesso2 & "B-BLOQUEIA"
          varMENTIPACS(I) = "B-BLOQUEIA"
       ElseIf Mid(sAcesso, I, 1) = "V" Then
          sAcesso2 = sAcesso2 & "V-REPROVA"
          varMENTIPACS(I) = "V-REPROVA"
       End If
       If I < Len(Trim(sAcesso)) Then sAcesso2 = sAcesso2 & ","
   Next I
   
   BREC.Close

End Function

