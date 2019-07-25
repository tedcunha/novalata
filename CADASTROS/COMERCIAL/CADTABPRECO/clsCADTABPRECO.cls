VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADTABPRECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial      As Integer
Dim intCODIGO      As Integer
Dim strCODPROD     As String
Dim dtDATATAB      As Date
Dim strVIGSIMNAO   As String
Dim arrTABPRECO    As Variant

Public Property Get VIGSIMNAO() As String
    VIGSIMNAO = strVIGSIMNAO
End Property

Public Property Let VIGSIMNAO(ByVal vData As String)
    strVIGSIMNAO = vData
End Property

Public Property Get DATATAB() As Date
    DATATAB = dtDATATAB
End Property

Public Property Let DATATAB(ByVal vData As Date)
    dtDATATAB = vData
End Property

Public Property Get TABPRECO() As Variant
    TABPRECO = arrTABPRECO
End Property

Public Property Let TABPRECO(ByVal vData As Variant)
    arrTABPRECO = vData
End Property

Public Property Get CODPROD() As String
    CODPROD = strCODPROD
End Property

Public Property Let CODPROD(ByVal vData As String)
    strCODPROD = vData
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

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADTABPRECOP.cCaminho = cCaminho
    frmCADTABPRECOP.Linha = Linha
    frmCADTABPRECOP.FILIAL = iFilial
    frmCADTABPRECOP.strAcesso = strAcesso
    frmCADTABPRECOP.strUSUARIO = strUSUARIO
    frmCADTABPRECOP.Show vbModal
End Sub

Public Function Pesq_CadTabPreco() As Boolean

   Pesq_CadTabPreco = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_TABPRECO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadTabPreco = True
   BREC.Close

End Function

Public Sub PreencheComboProd(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       PRODUT.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPRODUTO  PRODUT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       PRODUT.SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     Do While Not BREC.EOF
        If Not IsNull(BREC!SGI_CODIGO) Then
            Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
        End If
        BREC.MoveNext
     Loop
     BREC.Close
    
End Sub

Public Sub PreencheComboCondPgto(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       PGTO.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADCONDPGTO  PGTO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       PGTO.SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
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


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava

     GRAVA = False
     
     Dim strValor As String
     Dim I        As Integer
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
grvAltera:

        If IsArray(arrTABPRECO) = True Then
           For I = 1 To UBound(arrTABPRECO)
        
               '' Grava tipo de despesa
               sSql = "Insert Into SGI_TABPRECO (" & vbCrLf
               sSql = sSql & "                   SGI_FILIAL" & vbCrLf
               sSql = sSql & "                  ,SGI_CODTAB" & vbCrLf
               sSql = sSql & "                  ,SGI_DATATAB" & vbCrLf
               sSql = sSql & "                  ,SGI_CODPROD" & vbCrLf
               sSql = sSql & "                  ,SGI_CODPGTO" & vbCrLf
               sSql = sSql & "                  ,SGI_VALOR" & vbCrLf
               sSql = sSql & "                  ,SGI_PORCACRE" & vbCrLf
               sSql = sSql & "                  ,SGI_VLVENDA" & vbCrLf
               sSql = sSql & "                  ,SGI_VIGENTE)" & vbCrLf
               sSql = sSql & "                    Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intCODIGO & vbCrLf
               sSql = sSql & "                           ,'" & Format(dtDATATAB, "MM/DD/YYYY") & "'" & vbCrLf
               sSql = sSql & "                           ,'" & strCODPROD & "'" & vbCrLf
               sSql = sSql & "                           ," & arrTABPRECO(I, 1) & vbCrLf
        
               strValor = Replace(arrTABPRECO(I, 2), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                           , " & Trim(strValor) & vbCrLf
        
               strValor = Replace(arrTABPRECO(I, 3), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                           , " & Trim(strValor) & vbCrLf
               
               strValor = Replace(arrTABPRECO(I, 4), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                           , " & Trim(strValor) & vbCrLf
               
               sSql = sSql & "                           ,'" & strVIGSIMNAO & "'" & vbCrLf
               sSql = sSql & "                           )"
            
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
     
     ElseIf strOPER = "A" Then
     
        sSql = "Delete from SGI_TABPRECO " & vbCrLf
        sSql = sSql & "        Where " & vbCrLf
        sSql = sSql & "              SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "          And SGI_CODTAB = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        GoTo grvAltera
     
     ElseIf strOPER = "E" Then
     
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
   
   Dim intTOTREG As Integer
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_TABPRECO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODTAB = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
   
      intCODIGO = BREC!SGI_CODTAB
      strCODPROD = BREC!SGI_CODPROD
      
      dtDATATAB = BREC!SGI_DATATAB
      strVIGSIMNAO = BREC!SGI_VIGENTE
      
      intTOTREG = 0
      Do While Not BREC.EOF
         intTOTREG = intTOTREG + 1
         BREC.MoveNext
      Loop
      
      BREC.MoveFirst
      
      ReDim arrTABPRECO(1 To intTOTREG, 1 To 4) As String
      intTOTREG = 1
      Do While Not BREC.EOF
         arrTABPRECO(intTOTREG, 1) = BREC!SGI_CODPGTO
         arrTABPRECO(intTOTREG, 2) = BREC!SGI_VALOR
         arrTABPRECO(intTOTREG, 3) = BREC!SGI_PORCACRE
         arrTABPRECO(intTOTREG, 4) = BREC!SGI_VLVENDA
         intTOTREG = intTOTREG + 1
         BREC.MoveNext
      Loop
      
      Carrega_campos = True
   End If
   
   BREC.Close

End Function

