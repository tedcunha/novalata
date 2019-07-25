VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADQTDTURNOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial      As Integer
Dim intCODIGO      As Integer
Dim strDESCRI      As String
Dim intATIVO       As Integer
Dim arrDIASSEMANA  As Variant
Dim arrOPERADORES  As Variant
Dim arrPARADAS     As Variant

Public Property Get PARADAS() As Variant
    PARADAS = arrPARADAS
End Property

Public Property Let PARADAS(ByVal vData As Variant)
    arrPARADAS = vData
End Property

Public Property Let ATIVO(ByVal vData As Integer)
    intATIVO = vData
End Property

Public Property Get ATIVO() As Integer
    ATIVO = intATIVO
End Property

Public Property Get OPERADORES() As Variant
    OPERADORES = arrOPERADORES
End Property

Public Property Let OPERADORES(ByVal vData As Variant)
    arrOPERADORES = vData
End Property

Public Property Get DIASSEMANA() As Variant
    DIASSEMANA = arrDIASSEMANA
End Property

Public Property Let DIASSEMANA(ByVal vData As Variant)
    arrDIASSEMANA = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CODIGO() As Integer
    CODIGO = intCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Integer)
    intCODIGO = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property

Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADQTDTURNOSP.cCaminho = cCaminho
    frmCADQTDTURNOSP.Linha = Linha
    frmCADQTDTURNOSP.FILIAL = iFilial
    frmCADQTDTURNOSP.strAcesso = strAcesso
    frmCADQTDTURNOSP.strUSUARIO = strUSUARIO
    frmCADQTDTURNOSP.Show vbModal
End Sub

Public Function Pesq_CadQtdTurnos() As Boolean

   Pesq_CadQtdTurnos = False
   
   sSql = ""
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADQTDETURN " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadQtdTurnos = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
     
     Dim I As Integer
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
       
RETORNO:
        
        '' Grava turnos
        sSql = "Insert Into SGI_CADQTDETURN (SGI_FILIAL,SGI_CODIGO,SGI_DESCRI,SGI_QTDTURN,SGI_ATIVO) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intCODIGO
        sSql = sSql & "                           ,'" & Mid(strDESCRI, 1, 30) & "'" & vbCrLf
        sSql = sSql & "                           ,1" & vbCrLf
        sSql = sSql & "                           ," & intATIVO & vbCrLf
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrDIASSEMANA) = True Then
           For I = 1 To UBound(arrDIASSEMANA)
                
               sSql = "Insert into SGI_CADTURNSEM (SGI_FILIAL,SGI_CODIGO,SGI_DIASEM,SGI_HORAINI,SGI_HORAFIN,SGI_PARADA,SGI_TOTHORA,SGI_HORALIQ) Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ," & intCODIGO & vbCrLf
               sSql = sSql & "                         ," & arrDIASSEMANA(I, 1) & vbCrLf
               sSql = sSql & "                         ,'" & arrDIASSEMANA(I, 2) & "'" & vbCrLf
               sSql = sSql & "                         ,'" & arrDIASSEMANA(I, 3) & "'" & vbCrLf
               sSql = sSql & "                         ," & arrDIASSEMANA(I, 5) & vbCrLf
               sSql = sSql & "                         ,'" & arrDIASSEMANA(I, 4) & "'" & vbCrLf
               sSql = sSql & "                         ,'" & arrDIASSEMANA(I, 6) & "')"
               
               BGRV.CommandText = sSql
               BGRV.Execute
     
           Next I
        End If
     
        If IsArray(arrPARADAS) = True Then
           For I = 1 To UBound(arrPARADAS)
                
               sSql = "Insert into SGI_CADTURNPARADAS (SGI_FILIAL,SGI_CODIGO,SGI_PERIODO,SGI_PARADA,SGI_HORAINI,SGI_HORAFIN,SGI_TOTAL,SGI_COMPARADA) Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ," & intCODIGO & vbCrLf
               sSql = sSql & "                         ," & arrPARADAS(I, 6) & vbCrLf
               sSql = sSql & "                         ," & arrPARADAS(I, 1) & vbCrLf
               sSql = sSql & "                         ,'" & arrPARADAS(I, 2) & "'" & vbCrLf
               sSql = sSql & "                         ,'" & arrPARADAS(I, 3) & "'" & vbCrLf
               sSql = sSql & "                         ,'" & arrPARADAS(I, 4) & "'" & vbCrLf
               sSql = sSql & "                         ," & arrPARADAS(I, 5) & ")" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
     
           Next I
        End If
        
        If IsArray(arrOPERADORES) = True Then
           For I = 1 To UBound(arrOPERADORES)
                
               sSql = "Insert into SGI_CADMOVOPERMAQ (SGI_FILIAL,SGI_CODIGO,SGI_CODOPER) Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ," & intCODIGO & vbCrLf
               sSql = sSql & "                         ," & arrOPERADORES(I) & vbCrLf
               sSql = sSql & ")"
               
               BGRV.CommandText = sSql
               BGRV.Execute
     
           Next I
        End If
     
     ElseIf strOPER = "A" Then
     
        sSql = "Delete from SGI_CADMOVOPERMAQ " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Delete from SGI_CADTURNPARADAS " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Delete from SGI_CADTURNSEM " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrDIASSEMANA) = True Then
           For I = 1 To UBound(arrDIASSEMANA)
                
               sSql = "Insert into SGI_CADTURNSEM (SGI_FILIAL,SGI_CODIGO,SGI_DIASEM,SGI_HORAINI,SGI_HORAFIN,SGI_PARADA,SGI_TOTHORA,SGI_HORALIQ) Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ," & intCODIGO & vbCrLf
               sSql = sSql & "                         ," & arrDIASSEMANA(I, 1) & vbCrLf
               sSql = sSql & "                         ,'" & arrDIASSEMANA(I, 2) & "'" & vbCrLf
               sSql = sSql & "                         ,'" & arrDIASSEMANA(I, 3) & "'" & vbCrLf
               sSql = sSql & "                         ," & arrDIASSEMANA(I, 5) & vbCrLf
               sSql = sSql & "                         ,'" & arrDIASSEMANA(I, 4) & "'" & vbCrLf
               sSql = sSql & "                         ,'" & arrDIASSEMANA(I, 6) & "')"
               
               BGRV.CommandText = sSql
               BGRV.Execute
     
           Next I
        End If
     
        If IsArray(arrPARADAS) = True Then
           For I = 1 To UBound(arrPARADAS)
                
               sSql = "Insert into SGI_CADTURNPARADAS (SGI_FILIAL,SGI_CODIGO,SGI_PERIODO,SGI_PARADA,SGI_HORAINI,SGI_HORAFIN,SGI_TOTAL,SGI_COMPARADA) Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ," & intCODIGO & vbCrLf
               sSql = sSql & "                         ," & arrPARADAS(I, 6) & vbCrLf
               sSql = sSql & "                         ," & arrPARADAS(I, 1) & vbCrLf
               sSql = sSql & "                         ,'" & arrPARADAS(I, 2) & "'" & vbCrLf
               sSql = sSql & "                         ,'" & arrPARADAS(I, 3) & "'" & vbCrLf
               sSql = sSql & "                         ,'" & arrPARADAS(I, 4) & "'" & vbCrLf
               sSql = sSql & "                         ," & arrPARADAS(I, 5) & ")" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
     
           Next I
        End If
     
     
        If IsArray(arrOPERADORES) = True Then
           For I = 1 To UBound(arrOPERADORES)
                
               sSql = "Insert into SGI_CADMOVOPERMAQ (SGI_FILIAL,SGI_CODIGO,SGI_CODOPER) Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ," & intCODIGO & vbCrLf
               sSql = sSql & "                         ," & arrOPERADORES(I) & vbCrLf
               sSql = sSql & ")"
               
               BGRV.CommandText = sSql
               BGRV.Execute
     
           Next I
        End If
        
        sSql = "Update SGI_CADQTDETURN Set "
        sSql = sSql & "                    SGI_DESCRI = '" & strDESCRI & "'" & vbCrLf
        sSql = sSql & "                   ,SGI_ATIVO  = " & intATIVO & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
     ElseIf strOPER = "E" Then
        
     
        sSql = "Delete from SGI_CADMOVOPERMAQ " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Delete from SGI_CADTURNPARADAS " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        sSql = "Delete from SGI_CADTURNSEM " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        sSql = "Delete from SGI_CADQTDETURN " & vbCrLf
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


Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim nreg As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRI
      intATIVO = BREC!SGI_ATIVO
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   '' Itens -- Semenas dos turnos
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADTURNSEM " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      nreg = 0
      Do While Not BREC2.EOF
         nreg = nreg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrDIASSEMANA(1 To nreg, 1 To 5) As Variant
      nreg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrDIASSEMANA(nreg, 1) = Str(BREC2!SGI_DIASEM)
         arrDIASSEMANA(nreg, 2) = Format(BREC2!SGI_HORAINI, "HH:MM")
         arrDIASSEMANA(nreg, 3) = Format(BREC2!SGI_HORAFIN, "HH:MM")
         arrDIASSEMANA(nreg, 4) = Format(BREC2!SGI_TOTHORA, "HH:MM")
         arrDIASSEMANA(nreg, 5) = Format(BREC2!SGI_PARADA, "##00")
         nreg = nreg + 1
         BREC2.MoveNext
      Loop
      
   End If
   BREC2.Close
   
   '' Itens -- Semenas dos turnos
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADTURNPARADAS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
   
   BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC3.EOF Then
   
      nreg = 0
      Do While Not BREC3.EOF
         nreg = nreg + 1
         BREC3.MoveNext
      Loop
      
      ReDim arrPARADAS(1 To nreg, 1 To 6) As Variant
      nreg = 1
      BREC3.MoveFirst
      Do While Not BREC3.EOF
         arrPARADAS(nreg, 1) = Str(BREC3!SGI_PARADA)
         arrPARADAS(nreg, 2) = Format(BREC3!SGI_HORAINI, "HH:MM")
         arrPARADAS(nreg, 3) = Format(BREC3!SGI_HORAFIN, "HH:MM")
         arrPARADAS(nreg, 4) = Format(BREC3!SGI_TOTAL, "HH:MM")
         arrPARADAS(nreg, 5) = BREC3!SGI_COMPARADA
         arrPARADAS(nreg, 6) = BREC3!SGI_PERIODO
         nreg = nreg + 1
         BREC3.MoveNext
      Loop
      
   End If
   BREC3.Close
   
   
   '' Operadores
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMOVOPERMAQ " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
   
   BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC4.EOF Then
   
      nreg = 0
      Do While Not BREC4.EOF
         nreg = nreg + 1
         BREC4.MoveNext
      Loop
      
      ReDim arrOPERADORES(1 To nreg) As String
      nreg = 1
      BREC4.MoveFirst
      Do While Not BREC4.EOF
         arrOPERADORES(nreg) = Str(BREC4!SGI_CODOPER)
         nreg = nreg + 1
         BREC4.MoveNext
      Loop
      
   End If
   BREC4.Close
   
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

