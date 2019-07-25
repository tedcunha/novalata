VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADORDEMCORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial        As Integer
Dim lngCODIGO        As Long
Dim strCODMAQ        As String
Dim intDEFALT        As Integer
Dim arrCORES         As Variant

Public Property Get CORES() As Variant
    CORES = arrCORES
End Property

Public Property Let CORES(ByVal vData As Variant)
    arrCORES = vData
End Property

Public Property Get DEFALT() As Integer
    DEFALT = intDEFALT
End Property

Public Property Let DEFALT(ByVal vData As Integer)
    intDEFALT = vData
End Property

Public Property Get CODMAQ() As String
    CODMAQ = strCODMAQ
End Property

Public Property Let CODMAQ(ByVal vData As String)
    strCODMAQ = vData
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

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long)
    frmCADORDEMCORESP.cCaminho = cCaminho
    frmCADORDEMCORESP.Linha = Linha
    frmCADORDEMCORESP.FILIAL = iFilial
    frmCADORDEMCORESP.strAcesso = strAcesso
    frmCADORDEMCORESP.strUsuario = strUsuario
    frmCADORDEMCORESP.lngCodUsuaro = lngIDUsuario
    frmCADORDEMCORESP.intFILIALPED = 0
    frmCADORDEMCORESP.Show vbModal
End Sub

Public Function AtivoDesativo() As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADORDEMCORES " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim I As Long
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
Altera_Registro:
        
        '' Grava Cabeçalho de Ordem de Cores
        sSql = "Insert Into SGI_CADORDEMCORES (" & vbCrLf
        
        sSql = sSql & "                        SGI_FILIAL" & vbCrLf
        sSql = sSql & "                       ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                       ,SGI_DEFALT" & vbCrLf

        sSql = sSql & "              ) Values (" & vbCrLf
        
        sSql = sSql & "                        " & intFilial & vbCrLf
        sSql = sSql & "                       ," & lngCODIGO & vbCrLf
        sSql = sSql & "                       ," & intDEFALT & vbCrLf
        
        sSql = sSql & "                                )"
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrCORES) Then
            For I = 1 To UBound(arrCORES)
        
                '' Grava Ordem de Cores
                sSql = "Insert Into SGI_CADORDEMCORES_ORDEM (" & vbCrLf
                
                sSql = sSql & "                        SGI_FILIAL" & vbCrLf
                sSql = sSql & "                       ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                       ,SGI_ORDEM" & vbCrLf
                sSql = sSql & "                       ,SGI_IDPRODUTO" & vbCrLf
        
                sSql = sSql & "              ) Values (" & vbCrLf
                
                sSql = sSql & "                        " & intFilial & vbCrLf
                sSql = sSql & "                       ," & lngCODIGO & vbCrLf
                sSql = sSql & "                       ," & arrCORES(I, 1) & vbCrLf
                sSql = sSql & "                       ," & arrCORES(I, 2) & vbCrLf
                
                sSql = sSql & "                                )"
             
                BGRV.CommandText = sSql
                BGRV.Execute
        
            Next I
        End If
     
     ElseIf strOPER = "E" Or strOPER = "A" Then
     
        '' Item Ordem
        sSql = "Delete from SGI_CADORDEMCORES_ORDEM " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Cabeçalho
        sSql = "Delete from SGI_CADORDEMCORES " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
        If strOPER = "A" Then GoTo Altera_Registro
     
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

Public Function Carrega_campos() As Boolean

On Error GoTo Err_Carrega_campos
    
    Carrega_campos = False
    
    Dim intReg As Integer
    Dim I      As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDEMCORES " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
    sSql = sSql & "   And SGI_FILIAL   = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
        
       lngCODIGO = BREC!SGI_CODIGO
       intDEFALT = BREC!SGI_DEFALT
       
       Carrega_campos = True
    End If
    BREC.Close
   
    If Carrega_campos = False Then Exit Function
   
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ITEN.SGI_ORDEM" & vbCrLf
    sSql = sSql & "      ,ITEN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDEMCORES_ORDEM ITEN" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PROD"
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ITEN.SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And ITEN.SGI_CODIGO    = " & lngCODIGO & vbCrLf
    sSql = sSql & "   And ITEN.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ITEN.SGI_IDPRODUTO = PROD.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "Order By ITEN.SGI_ORDEM"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
      intReg = 0
      Do While Not BREC.EOF
         intReg = intReg + 1
         BREC.MoveNext
      Loop
      
      ReDim arrCORES(1 To intReg, 1 To 4) As String
      BREC.MoveFirst
      intReg = 1
      Do While Not BREC.EOF
        
         arrCORES(intReg, 1) = BREC!SGI_ORDEM
         arrCORES(intReg, 2) = BREC!SGI_CODIGO
         arrCORES(intReg, 3) = BREC!SGI_DESCRICAO
         arrCORES(intReg, 4) = BREC!SGI_IDPRODUTO
         
         intReg = intReg + 1
         BREC.MoveNext
      Loop
    End If
    BREC.Close

    Exit Function
    
Err_Carrega_campos:

    Carrega_campos = False
    If BREC.State = 1 Then BREC.Close
    
    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description & "Iten : " & I, "C", sSql, "frmCADMOVCXGER", "Carrega_campos")
    Set objErro = Nothing


End Function

