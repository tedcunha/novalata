VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADPROCESSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim intCODIGO       As Integer
Dim strDESCRI       As String
Dim lngTEMPMIN      As Long
Dim strCELFISICA    As String
Dim arrSUBSETOR     As Variant
Dim intAGREGAVALOR  As Integer
Dim lngFAMILIA      As Long
Dim lngTIPOPER      As Long
Dim arrOPERACAO     As Variant

Public Property Get Operacao() As Variant
    Operacao = arrOPERACAO
End Property
Public Property Let Operacao(ByVal vData As Variant)
    arrOPERACAO = vData
End Property


Public Property Get TIPOPER() As Long
    TIPOPER = lngTIPOPER
End Property
Public Property Let TIPOPER(ByVal vData As Long)
    lngTIPOPER = vData
End Property

Public Property Get FAMILIA() As Long
    FAMILIA = lngFAMILIA
End Property
Public Property Let FAMILIA(ByVal vData As Long)
    lngFAMILIA = vData
End Property


Public Property Get AGREGAVALOR() As Integer
    AGREGAVALOR = intAGREGAVALOR
End Property
Public Property Let AGREGAVALOR(ByVal vData As Integer)
    intAGREGAVALOR = vData
End Property

Public Property Get TEMPMIN() As Long
    TEMPMIN = lngTEMPMIN
End Property
Public Property Let TEMPMIN(ByVal vData As Long)
    lngTEMPMIN = vData
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

Public Property Get CELFISICA() As String
    CELFISICA = strCELFISICA
End Property

Public Property Let CELFISICA(ByVal vData As String)
    strCELFISICA = vData
End Property

Public Property Get SUBSETOR() As Variant
    SUBSETOR = arrSUBSETOR
End Property
Public Property Let SUBSETOR(ByVal vData As Variant)
    arrSUBSETOR = vData
End Property


Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADPROCESSOP.cCaminho = cCaminho
    frmCADPROCESSOP.Linha = Linha
    frmCADPROCESSOP.FILIAL = iFilial
    frmCADPROCESSOP.strAcesso = strAcesso
    frmCADPROCESSOP.Show vbModal
End Sub
Public Function Pesq_CadProcesso() As Boolean

   Pesq_CadProcesso = False
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       SGI_CODIGO  " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADPROCESSO " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadProcesso = True
   BREC.Close

End Function

Public Sub PreencheComboSubSecao(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADSUBSECAO " & vbCrLf
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

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim I As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava Sub-Seção
        sSql = "Insert Into SGI_CADPROCESSO (SGI_FILIAL,SGI_CODIGO,SGI_DESCRI,SGI_TEMPMIN,SGI_CELFISICA,SGI_AGREGAVALOR) Values (" & vbCrLf
        sSql = sSql & "                   " & intFilial & vbCrLf
        sSql = sSql & "                  ," & intCODIGO & vbCrLf
        sSql = sSql & "                  ,'" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & "                  ," & lngTEMPMIN & vbCrLf
        sSql = sSql & "                  ,''" & vbCrLf
        sSql = sSql & "                  ," & intAGREGAVALOR & vbCrLf
        sSql = sSql & "                   )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava os itens da Sub-Seção
        If IsArray(arrSUBSETOR) = True Then
           For I = 1 To UBound(arrSUBSETOR)
               
               sSql = "Insert Into SGI_CADITPROCESSO (SGI_FILIAL,SGI_CODIGO,SGI_CODSETOR) Values (" & vbCrLf
               sSql = sSql & "                    " & intFilial & vbCrLf
               sSql = sSql & "                   ," & intCODIGO & vbCrLf
               sSql = sSql & "                   ," & arrSUBSETOR(I) & vbCrLf
               sSql = sSql & "                   )"
           
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
        
        
        '' Gravando Operações
        If IsArray(arrOPERACAO) Then
            For I = 1 To UBound(arrOPERACAO)
            
                sSql = "Insert Into SGI_CADPROCOPERA (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_ORDEM" & vbCrLf
                sSql = sSql & "                             ,SGI_CODOPER" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & intCODIGO & vbCrLf
                sSql = sSql & "                             ," & arrOPERACAO(I, 1) & vbCrLf
                sSql = sSql & "                             ," & arrOPERACAO(I, 2) & vbCrLf
                sSql = sSql & "                             )"
               
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
        
     
     ElseIf (strOPER = "A") Then
     
        '' Altera Processo
        sSql = "Update SGI_CADPROCESSO Set " & vbCrLf
        
        sSql = sSql & " SGI_DESCRI      = '" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & ",SGI_AGREGAVALOR = " & intAGREGAVALOR & vbCrLf
        
        sSql = sSql & "       Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Apaga Sub-Seção
        sSql = "Delete from SGI_CADITPROCESSO " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Operações
        sSql = "Delete from SGI_CADPROCOPERA " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava os itens da Sub-Seção
        If IsArray(arrSUBSETOR) = True Then
           For I = 1 To UBound(arrSUBSETOR)
               
               sSql = "Insert Into SGI_CADITPROCESSO (SGI_FILIAL,SGI_CODIGO,SGI_CODSETOR) Values (" & vbCrLf
               sSql = sSql & "                    " & intFilial & vbCrLf
               sSql = sSql & "                   ," & intCODIGO & vbCrLf
               sSql = sSql & "                   ," & arrSUBSETOR(I) & vbCrLf
               sSql = sSql & "                   )"
           
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
        
        
        '' Gravando Operações
        If IsArray(arrOPERACAO) Then
            For I = 1 To UBound(arrOPERACAO)
            
                sSql = "Insert Into SGI_CADPROCOPERA (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_ORDEM" & vbCrLf
                sSql = sSql & "                             ,SGI_CODOPER" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & intCODIGO & vbCrLf
                sSql = sSql & "                             ," & arrOPERACAO(I, 1) & vbCrLf
                sSql = sSql & "                             ," & arrOPERACAO(I, 2) & vbCrLf
                sSql = sSql & "                             )"
               
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
        
     ElseIf (strOPER = "E") Then
        
        '' Apaga Sub-Seção
        sSql = "Delete from SGI_CADITPROCESSO " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Apaga Processo
        sSql = "Delete from SGI_CADPROCESSO " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Apaga Operações
        sSql = "Delete from SGI_CADPROCOPERA " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
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

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim intReg As Integer
   Dim I      As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRI
      lngTEMPMIN = BREC!SGI_TEMPMIN
      intAGREGAVALOR = BREC!SGI_AGREGAVALOR
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   If Carrega_campos = False Then Exit Function
   
   '' ---------------
   
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADITPROCESSO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intCODIGO & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrSUBSETOR(1 To intReg)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrSUBSETOR(intReg) = BREC2!SGI_CODSETOR
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   BREC2.Close


   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPROCOPERA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "Order By SGI_ORDEM"
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrOPERACAO(1 To intReg, 1 To 2) As String
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrOPERACAO(intReg, 1) = Trim(Str(BREC2!SGI_ORDEM))
         arrOPERACAO(intReg, 2) = Trim(Str(BREC2!SGI_CODOPER))
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   BREC2.Close

End Function

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


Public Sub PreenchComboTipOper(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
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

