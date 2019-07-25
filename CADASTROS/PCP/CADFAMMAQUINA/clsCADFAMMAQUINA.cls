VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADFAMMAQUINA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial      As Integer
Dim lngCODIGO      As Long
Dim strDESCRI      As String
Dim arrTURNOS      As Variant

Public Property Get TURNOS() As Variant
    TURNOS = arrTURNOS
End Property
Public Property Let TURNOS(ByVal vData As Variant)
    arrTURNOS = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property
Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
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
    frmCADFAMMAQUINAP.cCaminho = cCaminho
    frmCADFAMMAQUINAP.Linha = Linha
    frmCADFAMMAQUINAP.FILIAL = iFilial
    frmCADFAMMAQUINAP.strAcesso = strAcesso
    frmCADFAMMAQUINAP.strUSUARIO = strUSUARIO
    frmCADFAMMAQUINAP.Show vbModal
End Sub

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

Public Function Pesq_CadFamilia() As Boolean

   Pesq_CadFamilia = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADFAMMAQUINAS " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadFamilia = True
   BREC.Close

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

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     Dim I As Integer
     
     GRAVA = False
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava Familia de Máquinas
        sSql = "Insert Into SGI_CADFAMMAQUINAS (SGI_FILIAL,SGI_CODIGO,SGI_DESCRI) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & lngCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & "                           )"
               
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrTURNOS) = True Then
           For I = 1 To UBound(arrTURNOS)
                
               sSql = "Insert into SGI_CADFAMTURNO (SGI_FILIAL,SGI_CODIGO,SGI_CODTURNO) Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ," & lngCODIGO & vbCrLf
               sSql = sSql & "                         ," & arrTURNOS(I) & ")" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
     
           Next I
        End If
        
     
     ElseIf strOPER = "A" Then
     
        '' Apaga Turnos da Familia de Máquinas
        sSql = "Delete from SGI_CADFAMTURNO " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Altera Familia de Máquinas
        sSql = "Update SGI_CADFAMMAQUINAS Set " & vbCrLf
        sSql = sSql & "               SGI_DESCRI = '" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & lngCODIGO & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrTURNOS) = True Then
           For I = 1 To UBound(arrTURNOS)
                
               sSql = "Insert into SGI_CADFAMTURNO (SGI_FILIAL,SGI_CODIGO,SGI_CODTURNO) Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ," & lngCODIGO & vbCrLf
               sSql = sSql & "                         ," & arrTURNOS(I) & ")" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
     
           Next I
        End If
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Familia de Máquinas
        sSql = "Delete from SGI_CADFAMMAQUINAS " & vbCrLf
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
   
   Dim nreg As Long
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFAMMAQUINAS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      
      lngCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRI
      
      Carrega_campos = True
   End If
   BREC.Close


   '' Turnos
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFAMTURNO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      nreg = 0
      Do While Not BREC2.EOF
         nreg = nreg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrTURNOS(1 To nreg) As String
      nreg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrTURNOS(nreg) = Str(BREC2!SGI_CODTURNO)
         nreg = nreg + 1
         BREC2.MoveNext
      Loop
      
   End If
   BREC2.Close

End Function

