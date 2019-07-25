VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADGRPEMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim lngCODIGO           As Long
Dim strDESCRI           As String
Dim arrGRPEMPRESA       As Variant

Public Property Get GRPEMPRESA() As Variant
    GRPEMPRESA = arrGRPEMPRESA
End Property

Public Property Let GRPEMPRESA(ByVal vData As Variant)
    arrGRPEMPRESA = vData
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

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long)
    frmCADGRPEMPRESAP.cCaminho = cCaminho
    frmCADGRPEMPRESAP.Linha = Linha
    frmCADGRPEMPRESAP.FILIAL = iFilial
    frmCADGRPEMPRESAP.strAcesso = strAcesso
    frmCADGRPEMPRESAP.strUsuario = strUsuario
    frmCADGRPEMPRESAP.lngCodUsuaro = lngIDUsuario
    frmCADGRPEMPRESAP.intFILIALPED = 0
    frmCADGRPEMPRESAP.Show vbModal
End Sub

Public Function AtivoDesativo(strNOMETABELA As String) As Boolean

   AtivoDesativo = False
   
   If Len(Trim(strNOMETABELA)) = 0 Then Exit Function
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       " & strNOMETABELA & vbCrLf
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
        
        
        '' Grava Cabeçalho de Grupo de Cores
        sSql = "Insert Into SGI_CADGRPEMPRESA (" & vbCrLf
        
        sSql = sSql & "                        SGI_FILIAL" & vbCrLf
        sSql = sSql & "                       ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                       ,SGI_DESCRI" & vbCrLf

        sSql = sSql & "              ) Values (" & vbCrLf
        
        sSql = sSql & "                        " & intFilial & vbCrLf
        sSql = sSql & "                       ," & lngCODIGO & vbCrLf
        sSql = sSql & "                       ," & strDESCRI & vbCrLf
        
        sSql = sSql & "                                )"
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
Altera_Registro:
     
        If IsArray(arrGRPEMPRESA) Then
            For I = 1 To UBound(arrGRPEMPRESA)
        
                sSql = ""
                
                sSql = "Insert Into SGI_CADGRPEMPRESAIT" & " (" & vbCrLf
                sSql = sSql & "                                                     SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODEMP" & vbCrLf
                
                sSql = sSql & "                                           ) Values (" & vbCrLf
                sSql = sSql & "                                                     " & intFilial & vbCrLf
                sSql = sSql & "                                                    ," & lngCODIGO & vbCrLf
                sSql = sSql & "                                                    ," & arrGRPEMPRESA(I, 1) & vbCrLf
                
                sSql = sSql & "                                                     )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
        
     ElseIf strOPER = "A" Then
     
        
        '' Itens
        '' ==================
        sSql = ""
        
        sSql = "Delete from SGI_CADGRPEMPRESAIT" & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ==================
        sSql = ""
        
        sSql = "Update SGI_CADGRPEMPRESA Set" & vbCrLf
        sSql = sSql & "                           SGI_DESCRI = " & strDESCRI & vbCrLf
        
        sSql = sSql & "Where" & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "  And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo Altera_Registro
     
     ElseIf strOPER = "E" Then
     
        
        '' Itens
        '' ==================
        sSql = ""
        
        sSql = "Delete from SGI_CADGRPEMPRESAIT" & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Cabeçalho
        '' ==================
        sSql = ""
        
        sSql = "Delete from SGI_CADGRPEMPRESA" & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
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


Public Function Carrega_campos() As Boolean

On Error GoTo Err_Carrega_campos
    
    Carrega_campos = False
    
    Dim intReg As Integer
    Dim I      As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRPEMPRESA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
    sSql = sSql & "   And SGI_FILIAL   = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       lngCODIGO = BREC!SGI_CODIGO
       strDESCRI = BREC!SGI_DESCRI
       Carrega_campos = True
    End If
    BREC.Close
   
    If Carrega_campos = False Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRPEMPRESAIT" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCODIGO & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
      intReg = 0
      Do While Not BREC.EOF
         intReg = intReg + 1
         BREC.MoveNext
      Loop
      
      ReDim arrGRPEMPRESA(1 To intReg, 1 To 1) As String
      BREC.MoveFirst
      intReg = 1
      Do While Not BREC.EOF
        
         arrGRPEMPRESA(intReg, 1) = BREC!SGI_CODEMP
         
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

