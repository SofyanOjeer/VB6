Attribute VB_Name = "srvYSWI950"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'Dim rsSabX As New ADODB.Recordset
Dim rsADO As ADODB.Recordset
Public mYSWI950_SWISABSWID As Long, mYSWI950_SQL_Set As String
Public arrBIC_Loro() As String, arrBIC_Loro_Nb As Integer
Public arrBIC_Nostro() As String, arrBIC_Nostro_Nb As Integer

Type typeYSWI950
 
      SWI950SWID   As Long
      SWI950SWIL   As Long
      SWI950WES    As String
      SWI950WBIC   As String
      SWI950WVAL   As Long
      SWI950SENS   As String
      SWI950WDEV   As String
      SWI950WMTD   As Currency
      SWI950WN20   As String
      SWI950WL20   As String
      SWI950SWIX   As Long
          
   
End Type


Public Function sqlYSWI950_Delete(oldY As typeYSWI950)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWI950_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWI950SWID = " & oldY.SWI950SWID & " and SWI950SWIL = " & oldY.SWI950SWIL


'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YSWI950" & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWI950_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYSWI950_Delete = Error
End Function

Public Function sqlYSWI950_Update(newY As typeYSWI950, oldY As typeYSWI950)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWI950_Update = Null

'===================================================================================

xWhere = " where SWI950SWID = " & oldY.SWI950SWID & " and SWI950SWIL = " & newY.SWI950SWIL
xSet = " set"
blnUpdate = False

' Détecter les modifications
'===================================================================================
'If newY.SWI950SWID <> oldY.SWI950SWID Then blnUpdate = True:  xSet = xSet & " , SWI950SWID = " & newY.SWI950SWID
If newY.SWI950WVAL <> oldY.SWI950WVAL Then blnUpdate = True:  xSet = xSet & " , SWI950WVAL = " & newY.SWI950WVAL
If newY.SWI950SWIX <> oldY.SWI950SWIX Then blnUpdate = True:  xSet = xSet & " , SWI950SWIX = " & newY.SWI950SWIX
'If newY.SWI950SWIL <> oldY.SWI950SWIL Then blnUpdate = True:  xSet = xSet & " , SWI950SWIL = " & newY.SWI950SWIL
If newY.SWI950WMTD <> oldY.SWI950WMTD Then blnUpdate = True:  xSet = xSet & " , SWI950WMTD = " & cur_P(newY.SWI950WMTD)

If newY.SWI950WES <> oldY.SWI950WES Then blnUpdate = True:  xSet = xSet & " , SWI950WES = '" & newY.SWI950WES & "'"
If newY.SWI950WBIC <> oldY.SWI950WBIC Then blnUpdate = True:  xSet = xSet & " , SWI950WBIC = '" & newY.SWI950WBIC & "'"
If newY.SWI950WDEV <> oldY.SWI950WDEV Then blnUpdate = True:  xSet = xSet & " , SWI950WDEV = '" & Replace(newY.SWI950WDEV, "'", "''") & "'"
If newY.SWI950WN20 <> oldY.SWI950WN20 Then blnUpdate = True:  xSet = xSet & " , SWI950WN20 = '" & newY.SWI950WN20 & "'"
If newY.SWI950WL20 <> oldY.SWI950WL20 Then blnUpdate = True:  xSet = xSet & " , SWI950WL20 = '" & newY.SWI950WL20 & "'"
If newY.SWI950SENS <> oldY.SWI950SENS Then blnUpdate = True:  xSet = xSet & " , SWI950SENS = '" & newY.SWI950SENS & "'"

If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWI950" & xSet & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWI950_Update = "Erreur màj : " & newY.SWI950SWID & " " & newY.SWI950SWIL
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSWI950_Update = Error
End Function
Public Function sqlYSWI950_Update_Field(oldY As typeYSWI950, lSQL_Set As String)
Dim xSql As String, Nb As Long

On Error GoTo Error_Handler
sqlYSWI950_Update_Field = Null



xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWI950 " & lSQL_Set & "" _
     & " where SWI950SWID = " & oldY.SWI950SWID _
     & " and SWI950SWIL = " & oldY.SWI950SWIL
     
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWI950_Update_Field = "Erreur màj : " & oldY.SWI950SWID & " - " & oldY.SWI950SWIL
    Exit Function
End If
    

Exit Function
Error_Handler:
    sqlYSWI950_Update_Field = Error
End Function


Public Function sqlYSWI950_Update_Table(lSQL_Set As String)
Dim xSql As String, Nb As Long

On Error GoTo Error_Handler
sqlYSWI950_Update_Table = Null



xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWI950 " & lSQL_Set & ""
     
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWI950_Update_Table = "Erreur màj : " & lSQL_Set
    Exit Function
End If
    

Exit Function
Error_Handler:
    sqlYSWI950_Update_Table = Error
End Function

Public Function sqlYSWI950_Insert(newY As typeYSWI950)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSWI950_Insert = Null
xSet = " (SWI950SWID "
xValues = " values(" & newY.SWI950SWID

' Détecter les modifications
'===================================================================================
'If newY.SWI950SWID <> 0 Then xSet = xSet & ",SWI950SWID": xValues = xValues & " ," & newY.SWI950SWID
If newY.SWI950WVAL <> 0 Then xSet = xSet & ",SWI950WVAL": xValues = xValues & " ," & newY.SWI950WVAL
If newY.SWI950SWIX <> 0 Then xSet = xSet & ",SWI950SWIX": xValues = xValues & " ," & newY.SWI950SWIX
If newY.SWI950SWIL <> 0 Then xSet = xSet & ",SWI950SWIL": xValues = xValues & " ," & newY.SWI950SWIL
If newY.SWI950WMTD <> 0 Then xSet = xSet & ",SWI950WMTD": xValues = xValues & " ," & cur_P(newY.SWI950WMTD)

If Trim(newY.SWI950WES) <> "" Then xSet = xSet & ",SWI950WES": xValues = xValues & " ,'" & newY.SWI950WES & "'"
If Trim(newY.SWI950WBIC) <> "" Then xSet = xSet & ",SWI950WBIC": xValues = xValues & " ,'" & newY.SWI950WBIC & "'"
If Trim(newY.SWI950WDEV) <> "" Then xSet = xSet & ",SWI950WDEV": xValues = xValues & " ,'" & newY.SWI950WDEV & "'"

If Trim(newY.SWI950WN20) <> "" Then xSet = xSet & ",SWI950WN20": xValues = xValues & " ,'" & newY.SWI950WN20 & "'"
If Trim(newY.SWI950WL20) <> "" Then xSet = xSet & ",SWI950WL20": xValues = xValues & " ,'" & newY.SWI950WL20 & "'"
If Trim(newY.SWI950SENS) <> "" Then xSet = xSet & ",SWI950SENS": xValues = xValues & " ,'" & newY.SWI950SENS & "'"

      
       
      

xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YSWI950" & xSet & ")" & xValues & ")"
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWI950_Insert = "Erreur màj : " & newY.SWI950SWID & " - " & newY.SWI950SWIL
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWI950_Insert = Error
End Function

Public Function rsYSWI950_GetBuffer(rsADO As ADODB.Recordset, lYSWI950 As typeYSWI950)
On Error GoTo Error_Handler
rsYSWI950_GetBuffer = Null


lYSWI950.SWI950SWID = rsADO("SWI950SWID")
lYSWI950.SWI950WVAL = rsADO("SWI950WVAL")

lYSWI950.SWI950WES = rsADO("SWI950WES")
lYSWI950.SWI950WBIC = rsADO("SWI950WBIC")
lYSWI950.SWI950SWIX = rsADO("SWI950SWIX")

lYSWI950.SWI950WDEV = rsADO("SWI950WDEV")
lYSWI950.SWI950WMTD = rsADO("SWI950WMTD")
lYSWI950.SWI950WN20 = rsADO("SWI950WN20")
lYSWI950.SWI950WL20 = rsADO("SWI950WL20")
lYSWI950.SWI950SENS = rsADO("SWI950SENS")


Exit Function
Error_Handler:
rsYSWI950_GetBuffer = Error


End Function
Public Function rsYSWI950_Init(lYSWI950 As typeYSWI950)


lYSWI950.SWI950SWID = 0
lYSWI950.SWI950SWIL = 0
lYSWI950.SWI950WVAL = 0
lYSWI950.SWI950WES = ""
lYSWI950.SWI950WBIC = ""
lYSWI950.SWI950SWIX = 0

lYSWI950.SWI950WDEV = ""
lYSWI950.SWI950WMTD = 0
lYSWI950.SWI950WN20 = ""
lYSWI950.SWI950WL20 = ""
lYSWI950.SWI950SENS = ""


End Function





















