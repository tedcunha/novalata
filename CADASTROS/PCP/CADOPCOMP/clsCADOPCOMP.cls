VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADOPCOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial        As Integer
Dim lngCODIGO        As Long
Dim dtDATAOP         As String
Dim lngQTDOP         As Long
Dim lngCODCAPACID    As Long
Dim lngIDPRODUTO     As Long
Dim intSTATUS        As Integer
Dim strOBS           As String
Dim lngCODUSUARIO    As Long
Dim strNOMUSUARIO    As String
Dim strDTCRIA        As String

Public Property Get CODCAPACID() As Long
    CODCAPACID = lngCODCAPACID
End Property

Public Property Let CODCAPACID(ByVal vData As Long)
    lngCODCAPACID = vData
End Property

Public Property Get DTCRIA() As String
    DTCRIA = strDTCRIA
End Property

Public Property Let DTCRIA(ByVal vData As String)
    strDTCRIA = vData
End Property

Public Property Get NOMUSUARIO() As String
    NOMUSUARIO = strNOMUSUARIO
End Property

Public Property Let NOMUSUARIO(ByVal vData As String)
    strNOMUSUARIO = vData
End Property

Public Property Get CODUSUARIO() As Long
    CODUSUARIO = lngCODUSUARIO
End Property

Public Property Let CODUSUARIO(ByVal vData As Long)
    lngCODUSUARIO = vData
End Property

Public Property Get OBS() As String
    OBS = strOBS
End Property

Public Property Let OBS(ByVal vData As String)
    strOBS = vData
End Property

Public Property Get STATUS() As Integer
    STATUS = intSTATUS
End Property

Public Property Let STATUS(ByVal vData As Integer)
    intSTATUS = vData
End Property

Public Property Get IDPRODUTO() As Long
    IDPRODUTO = lngIDPRODUTO
End Property

Public Property Let IDPRODUTO(ByVal vData As Long)
    lngIDPRODUTO = vData
End Property

Public Property Get QTDOP() As Long
    QTDOP = lngQTDOP
End Property

Public Property Let QTDOP(ByVal vData As Long)
    lngQTDOP = vData
End Property

Public Property Get DATAOP() As String
    DATAOP = dtDATAOP
End Property

Public Property Let DATAOP(ByVal vData As String)
    dtDATAOP = vData
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

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADOPCOMPP.cCaminho = cCaminho
    frmCADOPCOMPP.Linha = Linha
    frmCADOPCOMPP.FILIAL = iFilial
    frmCADOPCOMPP.strAcesso = strAcesso
    frmCADOPCOMPP.strUsuario = strUsuario
    frmCADOPCOMPP.lngCodUsuaro = lngIDUsuario
    frmCADOPCOMPP.intFILIALPED = intFILIAL_PED
    frmCADOPCOMPP.Show vbModal
End Sub

Public Function PreenchComboAtivo() As String
    PreenchComboAtivo = PreenchComboAtivo & "|#" & 1 & ";SIM"
    PreenchComboAtivo = PreenchComboAtivo & "|#" & 0 & ";NÃO"
End Function

Public Function GRAVA(strOPER As String, strNOMFILIAL As String) As Boolean

On Error GoTo grav_err
     
     GRAVA = False
     
     Dim I As Integer
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Or strOPER = "A" Then
     
        If strOPER = "I" Then
            
            sSql = ""
            
            sSql = "Insert Into SGI_CADOPCOMPNT" & strNOMFILIAL & " (" & vbCrLf
            sSql = sSql & "                               SGI_FILIAL" & vbCrLf
            sSql = sSql & "                              ,SGI_CODIGO" & vbCrLf
            sSql = sSql & "                              ,SGI_STATUS" & vbCrLf
            sSql = sSql & "                              ,SGI_DATOP" & vbCrLf
            sSql = sSql & "                              ,SGI_QTDOP" & vbCrLf
            sSql = sSql & "                              ,SGI_CODLIN" & vbCrLf
            sSql = sSql & "                              ,SGI_IDPRODUTO" & vbCrLf
            sSql = sSql & "                              ,SGI_OBS" & vbCrLf
            sSql = sSql & "                              ,SGI_CODUSUARIO" & vbCrLf
            sSql = sSql & "                              ,SGI_USUARIO" & vbCrLf
            sSql = sSql & "                              ,SGI_DTCRIA" & vbCrLf
            
            sSql = sSql & "                     ) Values (" & vbCrLf
            sSql = sSql & "                               " & intFilial & vbCrLf
            sSql = sSql & "                              ," & lngCODIGO & vbCrLf
            sSql = sSql & "                              ," & intSTATUS & vbCrLf
            sSql = sSql & "                              ," & dtDATAOP & vbCrLf
            sSql = sSql & "                              ," & lngQTDOP & vbCrLf
            sSql = sSql & "                              ," & lngCODCAPACID & vbCrLf
            sSql = sSql & "                              ," & lngIDPRODUTO & vbCrLf
            sSql = sSql & "                              ," & strOBS & vbCrLf
            sSql = sSql & "                              ," & lngCODUSUARIO & vbCrLf
            sSql = sSql & "                              ," & strNOMUSUARIO & vbCrLf
            sSql = sSql & "                              ," & strDTCRIA & vbCrLf
            
            sSql = sSql & "                              )"
            
            BGRV.CommandText = sSql
            BGRV.Execute
        
        ElseIf strOPER = "A" Then
             
            sSql = ""
            
            sSql = "Update SGI_CADOPCOMPNT" & strNOMFILIAL & " Set " & vbCrLf
            sSql = sSql & "                               SGI_DATOP      = " & dtDATAOP & vbCrLf
            sSql = sSql & "                              ,SGI_QTDOP      = " & lngQTDOP & vbCrLf
            sSql = sSql & "                              ,SGI_IDPRODUTO  = " & lngIDPRODUTO & vbCrLf
            sSql = sSql & "                              ,SGI_OBS        = " & strOBS & vbCrLf
            
            sSql = sSql & "                     Where" & vbCrLf
            sSql = sSql & "                           SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "                       And SGI_CODIGO = " & lngCODIGO
            
            BGRV.CommandText = sSql
            BGRV.Execute
         
        End If
        
     ElseIf strOPER = "E" Then
        
        sSql = ""
        
        sSql = "Delete From SGI_CADOPCOMPNT" & strNOMFILIAL & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO    = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
               
     
     End If
    
     adoBanco_Dados.CommitTrans
     GRAVA = True
     
     Exit Function
     
grav_err:
     
     adoBanco_Dados.RollbackTrans
    
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
     Set objErro = Nothing
    
End Function

Public Function AtivoDesativo(strNOMFILIAL As String) As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADOPCOMPNT" & strNOMFILIAL & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function


Public Function Carrega_Campos(strNOMFILIAL As String) As Boolean

    Carrega_Campos = False

    Dim I          As Integer
    Dim intQTDREG  As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADOPCOMPNT" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
        
        intSTATUS = BREC!SGI_STATUS
        dtDATAOP = Format(BREC!SGI_DATOP, "DD/MM/YYYY")
        lngQTDOP = BREC!SGI_QTDOP
        lngIDPRODUTO = BREC!SGI_IDPRODUTO
        lngCODCAPACID = BREC!SGI_CODLIN
        
        If Not IsNull(BREC!SGI_OBS) Then strOBS = BREC!SGI_OBS
        If Not IsNull(BREC!SGI_CODUSUARIO) Then lngCODUSUARIO = BREC!SGI_CODUSUARIO
        If Not IsNull(BREC!SGI_USUARIO) Then strNOMUSUARIO = BREC!SGI_USUARIO
        If Not IsNull(BREC!SGI_DTCRIA) Then strDTCRIA = Format(BREC!SGI_DTCRIA, "DD/MM/YYYY HH:MM:SS")
        
        Carrega_Campos = True
    End If
    BREC.Close
    ' ----------------------------------------------------------
    
    If Carrega_Campos = False Then Exit Function
    
End Function

