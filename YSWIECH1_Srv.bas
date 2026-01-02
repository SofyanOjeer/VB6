Attribute VB_Name = "srvYSWIECH1"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'Dim rsSabX As New ADODB.Recordset
Dim rsADO As ADODB.Recordset

Type typeYSWIECH1
 
      SWIEC1SWID   As Long
      SWIEC1SEQ1   As Long
      SWIEC1SEQ0   As Long
      
      SWIEC1INFO   As String
      
      SWIEC1YAMJ   As Long
      SWIEC1YHMS   As Long
      SWIEC1YUSR   As String
      SWIEC1YVER   As String
     
   
End Type


Public Function sqlYSWIECH1_Delete(oldY As typeYSWIECH1)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWIECH1_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWIEC1SWID = " & oldY.SWIEC1SWID & " and SWIEC1SEQ1 = " & oldY.SWIEC1SEQ1


'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YSWIECH1" & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWIECH1_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYSWIECH1_Delete = Error
End Function

Public Function sqlYSWIECH1_Update(newY As typeYSWIECH1, oldY As typeYSWIECH1)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWIECH1_Update = Null

'===================================================================================

xWhere = " where SWIEC1SWID = " & oldY.SWIEC1SWID & " and SWIEC1SEQ1 = " & newY.SWIEC1SEQ1 & " and SWIEC1YVER = " & newY.SWIEC1YVER
xSet = " set"
blnUpdate = False
newY.SWIEC1YVER = newY.SWIEC1YVER + 1

' Détecter les modifications
'===================================================================================
'If newY.SWIEC1SWID <> oldY.SWIEC1SWID Then blnUpdate = True:  xSet = xSet & " , SWIEC1SWID = " & newY.SWIEC1SWID
'If newY.SWIEC1SEQ0 <> oldY.SWIEC1SEQ0 Then blnUpdate = True:  xSet = xSet & " , SWIEC1SEQ0 = " & newY.SWIEC1SEQ0

If newY.SWIEC1SEQ0 <> oldY.SWIEC1SEQ0 Then blnUpdate = True:  xSet = xSet & " , SWIEC1SEQ0 = " & newY.SWIEC1SEQ0

If newY.SWIEC1YAMJ <> oldY.SWIEC1YAMJ Then blnUpdate = True:  xSet = xSet & " , SWIEC1YAMJ = " & newY.SWIEC1YAMJ
If newY.SWIEC1YHMS <> oldY.SWIEC1YHMS Then blnUpdate = True:  xSet = xSet & " , SWIEC1YHMS = " & newY.SWIEC1YHMS

If newY.SWIEC1INFO <> oldY.SWIEC1INFO Then blnUpdate = True:  xSet = xSet & " , SWIEC1INFO = '" & newY.SWIEC1INFO & "'"
If newY.SWIEC1YUSR <> oldY.SWIEC1YUSR Then blnUpdate = True:  xSet = xSet & " , SWIEC1YUSR = '" & newY.SWIEC1YUSR & "'"

If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWIECH1" & xSet & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWIECH1_Update = "Erreur màj : " & newY.SWIEC1SEQ1
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSWIECH1_Update = Error
End Function
Public Function sqlYSWIECH1_Update_Field(oldY As typeYSWIECH1, lSQL_Set As String)
Dim xSql As String, Nb As Long

On Error GoTo Error_Handler
sqlYSWIECH1_Update_Field = Null



xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWIECH1 " & lSQL_Set & "" _
     & " where SWIEC1SWID = '" & oldY.SWIEC1SWID & "'" _
     & " and SWIEC1SEQ1 = " & oldY.SWIEC1SEQ1 _
     & " and SWIEC1YVER = " & oldY.SWIEC1YVER
     
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIECH1_Update_Field = "Erreur màj : " & oldY.SWIEC1SWID & " - " & oldY.SWIEC1SEQ1
    Exit Function
End If
    

Exit Function
Error_Handler:
    sqlYSWIECH1_Update_Field = Error
End Function


Public Function sqlYSWIECH1_Insert(newY As typeYSWIECH1)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSWIECH1_Insert = Null
xSet = " (SWIEC1SWID "
xValues = " values(" & newY.SWIEC1SWID

' Détecter les modifications
'===================================================================================
'If newY.SWIEC1SWID <> 0 Then xSet = xSet & ",SWIEC1SWID": xValues = xValues & " ," & newY.SWIEC1SWID
If newY.SWIEC1SEQ1 <> 0 Then xSet = xSet & ",SWIEC1SEQ1": xValues = xValues & " ," & newY.SWIEC1SEQ1
If newY.SWIEC1SEQ0 <> 0 Then xSet = xSet & ",SWIEC1SEQ0": xValues = xValues & " ," & newY.SWIEC1SEQ0
If newY.SWIEC1YAMJ <> 0 Then xSet = xSet & ",SWIEC1YAMJ": xValues = xValues & " ," & newY.SWIEC1YAMJ
If newY.SWIEC1YHMS <> 0 Then xSet = xSet & ",SWIEC1YHMS": xValues = xValues & " ," & newY.SWIEC1YHMS
If newY.SWIEC1YVER <> 0 Then xSet = xSet & ",SWIEC1YVER": xValues = xValues & " ," & newY.SWIEC1YVER

If Trim(newY.SWIEC1INFO) <> "" Then xSet = xSet & ",SWIEC1INFO": xValues = xValues & " ,'" & Replace(Trim(newY.SWIEC1INFO), "'", "''") & "'"

If Trim(newY.SWIEC1YUSR) <> "" Then xSet = xSet & ",SWIEC1YUSR": xValues = xValues & " ,'" & newY.SWIEC1YUSR & "'"
      
       
      

xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YSWIECH1" & xSet & ")" & xValues & ")"
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIECH1_Insert = "Erreur màj : " & newY.SWIEC1SWID & " - " & newY.SWIEC1SEQ1
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWIECH1_Insert = Error
End Function

Public Function rsYSWIECH1_GetBuffer(rsADO As ADODB.Recordset, lYSWIECH1 As typeYSWIECH1)
On Error GoTo Error_Handler
rsYSWIECH1_GetBuffer = Null


lYSWIECH1.SWIEC1SWID = rsADO("SWIEC1SWID")
lYSWIECH1.SWIEC1SEQ1 = rsADO("SWIEC1SEQ1")
lYSWIECH1.SWIEC1SEQ0 = rsADO("SWIEC1SEQ0")

lYSWIECH1.SWIEC1INFO = rsADO("SWIEC1INFO")

lYSWIECH1.SWIEC1YAMJ = rsADO("SWIEC1YAMJ")
lYSWIECH1.SWIEC1YHMS = rsADO("SWIEC1YHMS")
lYSWIECH1.SWIEC1YVER = rsADO("SWIEC1YVER")
lYSWIECH1.SWIEC1YUSR = rsADO("SWIEC1YUSR")

Exit Function
Error_Handler:
rsYSWIECH1_GetBuffer = Error


End Function
Public Function rsYSWIECH1_Init(lYSWIECH1 As typeYSWIECH1)


lYSWIECH1.SWIEC1SWID = 0
lYSWIECH1.SWIEC1SEQ1 = 0
lYSWIECH1.SWIEC1SEQ0 = 0

lYSWIECH1.SWIEC1INFO = ""

lYSWIECH1.SWIEC1YAMJ = 0
lYSWIECH1.SWIEC1YHMS = 0
lYSWIECH1.SWIEC1YVER = ""
lYSWIECH1.SWIEC1YUSR = ""

End Function





















