VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADMOEDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial      As Integer
Dim intCODIGO      As Integer
Dim strDESCRI      As String
Dim arrINDICE      As Variant

Public Property Get INDICE() As Variant
    INDICE = arrINDICE
End Property
Public Property Let INDICE(ByVal vData As Variant)
    arrINDICE = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property
Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property

Public Property Get CODIGO() As Integer
    CODIGO = intCODIGO
End Property
Public Property Let CODIGO(ByVal vData As Integer)
    intCODIGO = vData
End Property
Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property
Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADMOEDAP.cCaminho = cCaminho
    frmCADMOEDAP.Linha = Linha
    frmCADMOEDAP.FILIAL = iFilial
    frmCADMOEDAP.strAcesso = strAcesso
    frmCADMOEDAP.Show vbModal
End Sub

Public Function Pesq_CadMoeda() As Boolean

   Pesq_CadMoeda = False
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *  " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMOEDA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadMoeda = True
   BREC.Close

End Function

Public Function Gera_Codigo(sModulo As String) As Integer

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

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
     
     Dim I      As Integer
     Dim sValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
     
Altera:
       
        '' Grava especificação técnica do produto
        sSql = "Insert Into SGI_CADMOEDA (SGI_FILIAL,SGI_CODIGO,SGI_DESCRICAO) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & intCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & strDESCRI & "'" & vbCrLf
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Gravando Indice da Moeda
        If UBound(arrINDICE) > 0 Then
           For I = 1 To UBound(arrINDICE)
               
               sSql = "Insert Into SGI_CADMOEDAINDICE (SGI_FILIAL,SGI_CODIGO,SGI_DTINDICE,SGI_VLINDICE) Values (" & vbCrLf
               sSql = sSql & "                      " & intFilial & vbCrLf
               sSql = sSql & "                     ," & intCODIGO & vbCrLf
               sSql = sSql & "                     ,'" & Format(CDate(arrINDICE(I, 1)), "MM/DD/YYYY") & "'" & vbCrLf
        
               sValor = Replace(arrINDICE(I, 2), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               sSql = sSql & "                           ," & sValor & vbCrLf
               
               sSql = sSql & "                           )"
        
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next I
        End If
             
     ElseIf strOPER = "A" Or strOPER = "E" Then
       
        sSql = "Delete from SGI_CADMOEDA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Delete from SGI_CADMOEDAINDICE " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo Altera
        
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
   
   Dim intREGS As Integer
   
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMOEDA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRICAO
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   '' ------------------------------------------------
   '' Carrega Indices
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMOEDAINDICE " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
         
         intREGS = 0
         Do While Not BREC.EOF
            intREGS = (intREGS + 1)
            BREC.MoveNext
         Loop
        
         ReDim arrINDICE(1 To intREGS, 1 To 2) As String
         
         BREC.MoveFirst
         intREGS = 1
         Do While Not BREC.EOF
            arrINDICE(intREGS, 1) = Format(BREC!SGI_DTINDICE, "DD/MM/YYYY")
            arrINDICE(intREGS, 2) = Format(BREC!SGI_VLINDICE, "#,##0.00")
            intREGS = (intREGS + 1)
            BREC.MoveNext
         Loop
        
   End If
   BREC.Close
   '' ------------------------------------------------

End Function

