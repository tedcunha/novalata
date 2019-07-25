VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADOPSQUALI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial    As Integer
Dim intCODIGO    As Integer
Dim strDTLANC    As String
Dim arrOPS       As Variant

Public Property Get OPS() As Variant
    OPS = arrOPS
End Property

Public Property Let OPS(ByVal vData As Variant)
    arrOPS = vData
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

Public Property Get DTLANC() As String
    DTLANC = strDTLANC
End Property

Public Property Let DTLANC(ByVal vData As String)
    strDTLANC = vData
End Property

Public Function AtivoDesativo() As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADOPQUALIH" & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function


Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUSUARIO As Long)
    frmCADOPSQUALIP.cCaminho = cCaminho
    frmCADOPSQUALIP.Linha = Linha
    frmCADOPSQUALIP.FILIAL = iFilial
    frmCADOPSQUALIP.strAcesso = strAcesso
    frmCADOPSQUALIP.lngIDUSUARIO = lngIDUSUARIO
    frmCADOPSQUALIP.strUsuario = strUsuario
    frmCADOPSQUALIP.Show vbModal
End Sub

Public Function Carrega_Campos() As Boolean

    Carrega_Campos = False

    Dim i          As Integer
    Dim intQTDREG  As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADOPQUALIH" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
        intCODIGO = BREC!SGI_CODIGO
        strDTLANC = Format(BREC!SGI_DTLANC, "DD/MM/YYYY")
        Carrega_Campos = True
    End If
    BREC.Close
    ' ----------------------------------------------------------
    
    If Carrega_Campos = False Then Exit Function
    
    '' -----------------------------------------------------------
    '' Itens
    sSql = ""
    arrOPS = Empty
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADOPQUALI_IT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrOPS(1 To intQTDREG, 1 To 2) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF()
          
          arrOPS(intQTDREG, 1) = BREC3!SGI_CODOP
          arrOPS(intQTDREG, 2) = BREC3!SGI_FILIALOP
          
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close
    
    
End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
     
     Dim i As Long
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
    
        '' Grava Op's Header
        sSql = ""
        sSql = "Insert Into SGI_CADOPQUALIH (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                     ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                     ,SGI_DTLANC" & vbCrLf
        
        sSql = sSql & "                       ) Values (" & vbCrLf
        
        sSql = sSql & "                                 " & intFilial & vbCrLf
        sSql = sSql & "                                ," & intCODIGO & vbCrLf
        sSql = sSql & "                                ," & strDTLANC & vbCrLf
        sSql = sSql & "                                )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
Alt_OPS:

        If IsArray(arrOPS) Then
        
            '' Grava Op's Itens
            For i = 1 To UBound(arrOPS)
                
                sSql = ""
                sSql = "Insert Into SGI_CADOPQUALI_IT (SGI_FILIAL" & vbCrLf
                sSql = sSql & "                     ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                     ,SGI_CODOP" & vbCrLf
                sSql = sSql & "                     ,SGI_FILIALOP" & vbCrLf
                
                sSql = sSql & "                       ) Values (" & vbCrLf
                
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intCODIGO & vbCrLf
                sSql = sSql & "                                ," & Trim(arrOPS(i, 1)) & vbCrLf
                sSql = sSql & "                                ," & Trim(arrOPS(i, 2)) & vbCrLf
                
                sSql = sSql & "                                )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next i
        End If
     
     ElseIf strOPER = "A" Then
            
        '' Altara Op's Header
        sSql = ""
        
        sSql = "Update SGI_CADOPQUALIH Set " & vbCrLf
        sSql = sSql & "                               SGI_DTLANC = " & strDTLANC & vbCrLf
        sSql = sSql & "                     Where" & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Op's Header
        sSql = ""
        sSql = "Delete from SGI_CADOPQUALI_IT" & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        GoTo Alt_OPS
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Op's Item
        sSql = "Delete from SGI_CADOPQUALI_IT" & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Op's Header
        sSql = "Delete from SGI_CADOPQUALIH " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
     
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

