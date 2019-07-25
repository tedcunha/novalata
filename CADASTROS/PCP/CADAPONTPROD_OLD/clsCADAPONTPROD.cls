VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADAPONTPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial        As Integer
Dim lngCODIGO        As Long
Dim lngCODOP         As Long
Dim lngCODFP         As Long
Dim strDATDOC        As String
Dim lngSTATUS        As Long
Dim intFECHAAPONT    As Integer
Dim arrITENSAPONT    As Variant
Dim lngQTDEAPONT     As Long

Public Property Get CODFP() As Long
    CODFP = lngCODFP
End Property

Public Property Let CODFP(ByVal vData As Long)
    lngCODFP = vData
End Property

Public Property Get QTDEAPONT() As Long
    QTDEAPONT = lngQTDEAPONT
End Property

Public Property Let QTDEAPONT(ByVal vData As Long)
    lngQTDEAPONT = vData
End Property

Public Property Get ITENSAPONT() As Variant
    ITENSAPONT = arrITENSAPONT
End Property

Public Property Let ITENSAPONT(ByVal vData As Variant)
    arrITENSAPONT = vData
End Property

Public Property Get DATDOC() As String
    DATDOC = strDATDOC
End Property

Public Property Let DATDOC(ByVal vData As String)
    strDATDOC = vData
End Property

Public Property Get FECHAAPONT() As Integer
    FECHAAPONT = intFECHAAPONT
End Property

Public Property Let FECHAAPONT(ByVal vData As Integer)
    intFECHAAPONT = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Property Get STATUS() As Long
    STATUS = lngSTATUS
End Property

Public Property Let STATUS(ByVal vData As Long)
    lngSTATUS = vData
End Property

Public Property Get CODOP() As Long
    CODOP = lngCODOP
End Property

Public Property Let CODOP(ByVal vData As Long)
    lngCODOP = vData
End Property


Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long)
    frmCADAPONTPRODP.cCaminho = cCaminho
    frmCADAPONTPRODP.Linha = Linha
    frmCADAPONTPRODP.FILIAL = iFilial
    frmCADAPONTPRODP.strAcesso = strAcesso
    frmCADAPONTPRODP.strUsuario = strUsuario
    frmCADAPONTPRODP.lngCodUsuario = lngIDUsuario
    frmCADAPONTPRODP.Show vbModal
End Sub

Public Function AtivoDesativo() As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADAPONTPRDH " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
     
     Dim I  As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
      
     If strOPER = "I" Then '' Inclui
     
Status_Altera:
        
        '' Inserindo OP Enviada
        sSql = "Insert Into SGI_CADAPONTPRDH (" & vbCrLf
        sSql = sSql & "                               SGI_FILIAL" & vbCrLf
        sSql = sSql & "                              ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                              ,SGI_CODOP" & vbCrLf
        sSql = sSql & "                              ,SGI_CODFP" & vbCrLf
        sSql = sSql & "                              ,SGI_DTDOC" & vbCrLf
        sSql = sSql & "                              ,SGI_STATUS" & vbCrLf
        sSql = sSql & "                              ,SGI_QTDEAPONT" & vbCrLf
        sSql = sSql & "                              ,SGI_FECHAPONT" & vbCrLf
        sSql = sSql & "                     ) Values (" & vbCrLf
        sSql = sSql & "                               " & intFilial & vbCrLf
        sSql = sSql & "                              ," & lngCODIGO & vbCrLf
        sSql = sSql & "                              ," & lngCODOP & vbCrLf
        sSql = sSql & "                              ," & lngCODFP & vbCrLf
        sSql = sSql & "                              ,'" & Trim(strDATDOC) & "'" & vbCrLf
        sSql = sSql & "                              ," & lngSTATUS & vbCrLf
        sSql = sSql & "                              ," & lngQTDEAPONT & vbCrLf
        sSql = sSql & "                              ," & intFECHAAPONT & vbCrLf
        sSql = sSql & "                              )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If IsArray(arrITENSAPONT) Then
            For I = 1 To UBound(arrITENSAPONT)
                
                sSql = "Insert Into SGI_CADAPONTPRDI (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_CODPROC" & vbCrLf
                sSql = sSql & "                             ,SGI_ORDEM" & vbCrLf
                sSql = sSql & "                             ,SGI_HORINI" & vbCrLf
                sSql = sSql & "                             ,SGI_HORFIN" & vbCrLf
                sSql = sSql & "                             ,SGI_TOTAL" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngCODIGO & vbCrLf
                sSql = sSql & "                             ," & arrITENSAPONT(I, 1) & vbCrLf
                sSql = sSql & "                             ," & (I - 1) & vbCrLf
                sSql = sSql & "                             ," & arrITENSAPONT(I, 2) & vbCrLf
                sSql = sSql & "                             ," & arrITENSAPONT(I, 3) & vbCrLf
                sSql = sSql & "                             ," & arrITENSAPONT(I, 4) & vbCrLf
                sSql = sSql & "                             )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
        
        If intFECHAAPONT = 0 Then
        
            sSql = "Update SGI_ORDEMPROD Set " & vbCrLf
            sSql = sSql & "                         SGI_OPENVIADA = 2" & vbCrLf
            sSql = sSql & "                   Where " & vbCrLf
            sSql = sSql & "                         SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "                     And SGI_CODIGO = " & lngCODOP
            
            BGRV.CommandText = sSql
            BGRV.Execute
        
        ElseIf intFECHAAPONT = 1 Then
        
            sSql = "Update SGI_ORDEMPROD Set " & vbCrLf
            sSql = sSql & "                         SGI_OPENVIADA = 3" & vbCrLf
            sSql = sSql & "                   Where " & vbCrLf
            sSql = sSql & "                         SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "                     And SGI_CODIGO = " & lngCODOP
            
            BGRV.CommandText = sSql
            BGRV.Execute
        
        End If
     
     ElseIf strOPER = "A" Or strOPER = "E" Then '' Altera/Exclui
     
        '' -----------------------------------------------
        
        sSql = "Delete From SGI_CADAPONTPRDI " & vbCrLf
        sSql = sSql & "                        Where " & vbCrLf
        sSql = sSql & "                              SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "                          And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' -----------------------------------------------
        
        sSql = "Delete From SGI_CADAPONTPRDH " & vbCrLf
        sSql = sSql & "                        Where " & vbCrLf
        sSql = sSql & "                              SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "                          And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' -----------------------------------------------
        
        If intFECHAAPONT = 1 Or intFECHAAPONT = 0 Then
        
            sSql = "Update SGI_ORDEMPROD Set " & vbCrLf
            sSql = sSql & "                         SGI_OPENVIADA = 1" & vbCrLf
            sSql = sSql & "                   Where " & vbCrLf
            sSql = sSql & "                         SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "                     And SGI_CODIGO = " & lngCODOP
            
            BGRV.CommandText = sSql
            BGRV.Execute
        
        End If
        
        
        If strOPER = "A" Then GoTo Status_Altera
        
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


Public Function Carrega_Campos() As Boolean

    Carrega_Campos = False

    Dim I         As Integer
    Dim intQTDREG As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRDH.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPRDH PRDH" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRDH.SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And PRDH.SGI_CODIGO = " & lngCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       BREC.Close
       Exit Function
    End If
    
    lngCODIGO = BREC!SGI_CODIGO
    lngCODOP = BREC!SGI_CODOP
    lngCODFP = BREC!SGI_CODFP
    strDATDOC = Format(BREC!SGI_DTDOC, "DD/MM/YYYY")
    lngSTATUS = BREC!SGI_STATUS
    lngQTDEAPONT = BREC!SGI_QTDEAPONT
    intFECHAAPONT = BREC!SGI_FECHAPONT
    
    BREC.Close
    
    Carrega_Campos = True
    ' ----------------------------------------------------------
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRDI.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPRDI PRDI" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRDI.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRDI.SGI_CODIGO = " & lngCODIGO & vbCrLf
    sSql = sSql & "Order By PRDI.SGI_ORDEM "
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF() Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrITENSAPONT(1 To intQTDREG, 1 To 4) As String
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF()
       
          arrITENSAPONT(intQTDREG, 1) = BREC3!SGI_CODPROC
          If Not IsNull(BREC3!SGI_HORINI) Then arrITENSAPONT(intQTDREG, 2) = BREC3!SGI_HORINI
          If Not IsNull(BREC3!SGI_HORFIN) Then arrITENSAPONT(intQTDREG, 3) = BREC3!SGI_HORFIN
          If Not IsNull(BREC3!SGI_TOTAL) Then arrITENSAPONT(intQTDREG, 4) = BREC3!SGI_TOTAL
          
          BREC3.MoveNext
          intQTDREG = (intQTDREG + 1)
       Loop
    End If
    BREC3.Close
    
End Function

