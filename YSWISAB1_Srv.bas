Attribute VB_Name = "srvYSWISAB1"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
Type typeYSWISAB1
 
      SWISAB1ID   As Long
      
      SWISABW50P   As String
      SWISABW50Z   As String
      SWISABW52A   As String
      
      SWISABW59P   As String
      SWISABW59Z   As String
      SWISABW57A   As String
      SWISABWEBA   As String
      SWISABW71A   As String

'____________________________________________________ Journalisation
End Type
Public xYSWISAB1 As typeYSWISAB1
Public Function sqlYSWISAB1_Delete(oldY As typeYSWISAB1)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWISAB1_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWISAB1ID = " & oldY.SWISAB1ID


'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YSWISAB1" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWISAB1_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYSWISAB1_Delete = Error
End Function

Public Function sqlYSWISAB1_Update(newY As typeYSWISAB1, oldY As typeYSWISAB1)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWISAB1_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SWISAB1ID <> newY.SWISAB1ID Then
    sqlYSWISAB1_Update = "Erreur SWISAB1ID"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWISAB1ID = " & oldY.SWISAB1ID

xSet = " set"
blnUpdate = False


' Détecter les modifications
'===================================================================================
If newY.SWISABW50P <> oldY.SWISABW50P Then blnUpdate = True:  xSet = xSet & " , SWISABW50P = '" & Replace(newY.SWISABW50P, "'", "''") & "'"
If newY.SWISABW50Z <> oldY.SWISABW50Z Then blnUpdate = True:  xSet = xSet & " , SWISABW50Z = '" & Replace(newY.SWISABW50Z, "'", "''") & "'"
If newY.SWISABW52A <> oldY.SWISABW52A Then blnUpdate = True:  xSet = xSet & " , SWISABW52A = '" & newY.SWISABW52A & "'"

If newY.SWISABW59P <> oldY.SWISABW59P Then blnUpdate = True:  xSet = xSet & " , SWISABW59P = '" & newY.SWISABW59P & "'"
If newY.SWISABW59Z <> oldY.SWISABW59Z Then blnUpdate = True:  xSet = xSet & " , SWISABW59Z = '" & newY.SWISABW59Z & "'"
If newY.SWISABW57A <> oldY.SWISABW57A Then blnUpdate = True:  xSet = xSet & " , SWISABW57A = '" & newY.SWISABW57A & "'"
If newY.SWISABWEBA <> oldY.SWISABWEBA Then blnUpdate = True:  xSet = xSet & " , SWISABWEBA = '" & newY.SWISABWEBA & "'"
If newY.SWISABW71A <> oldY.SWISABW71A Then blnUpdate = True:  xSet = xSet & " , SWISABW71A = '" & newY.SWISABW71A & "'"


If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB1" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWISAB1_Update = "Erreur màj : " & newY.SWISAB1ID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSWISAB1_Update = Error
End Function

Public Function sqlYSWISAB1_Insert(newY As typeYSWISAB1)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSWISAB1_Insert = Null
xSet = " (SWISAB1ID "
xValues = " values(" & newY.SWISAB1ID

' Détecter les modifications
'===================================================================================
If Trim(newY.SWISABW50P) <> "" Then xSet = xSet & ",SWISABW50P": xValues = xValues & " ,'" & Replace(newY.SWISABW50P, "'", "''") & "'"
If Trim(newY.SWISABW50Z) <> "" Then xSet = xSet & ",SWISABW50Z": xValues = xValues & " ,'" & Replace(newY.SWISABW50Z, "'", "''") & "'"
If Trim(newY.SWISABW52A) <> "" Then
    If Len(newY.SWISABW52A) > 11 Then newY.SWISABW52A = Mid$(newY.SWISABW52A, 1, 11)
    xSet = xSet & ",SWISABW52A": xValues = xValues & " ,'" & newY.SWISABW52A & "'"
End If
If Trim(newY.SWISABW59P) <> "" Then xSet = xSet & ",SWISABW59P": xValues = xValues & " ,'" & newY.SWISABW59P & "'"
If Trim(newY.SWISABW59Z) <> "" Then xSet = xSet & ",SWISABW59Z": xValues = xValues & " ,'" & newY.SWISABW59Z & "'"
If Trim(newY.SWISABW57A) <> "" Then
    If Len(newY.SWISABW57A) > 11 Then newY.SWISABW57A = Mid$(newY.SWISABW57A, 1, 11)
    xSet = xSet & ",SWISABW57A": xValues = xValues & " ,'" & newY.SWISABW57A & "'"
End If
If Trim(newY.SWISABWEBA) <> "" Then xSet = xSet & ",SWISABWEBA": xValues = xValues & " ,'" & newY.SWISABWEBA & "'"
If Trim(newY.SWISABW71A) <> "" Then xSet = xSet & ",SWISABW71A": xValues = xValues & " ,'" & newY.SWISABW71A & "'"
xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YSWISAB1" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWISAB1_Insert = "Erreur màj : " & newY.SWISAB1ID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWISAB1_Insert = Error
End Function

Public Function rsYSWISAB1_GetBuffer(rsADO As ADODB.Recordset, lYSWISAB1 As typeYSWISAB1)
On Error GoTo Error_Handler
rsYSWISAB1_GetBuffer = Null

lYSWISAB1.SWISAB1ID = rsADO("SWISAB1ID")

lYSWISAB1.SWISABW50P = rsADO("SWISABW50P")
lYSWISAB1.SWISABW50Z = rsADO("SWISABW50Z")
lYSWISAB1.SWISABW52A = rsADO("SWISABW52A")

lYSWISAB1.SWISABW59P = rsADO("SWISABW59P")
lYSWISAB1.SWISABW59Z = rsADO("SWISABW59Z")
lYSWISAB1.SWISABW57A = rsADO("SWISABW57A")

lYSWISAB1.SWISABWEBA = rsADO("SWISABWEBA")
lYSWISAB1.SWISABW71A = rsADO("SWISABW71A")

Exit Function
Error_Handler:
rsYSWISAB1_GetBuffer = Error


End Function
Public Function rsYSWISAB1_Init(lYSWISAB1 As typeYSWISAB1)

lYSWISAB1.SWISAB1ID = 0

lYSWISAB1.SWISABW50P = ""
lYSWISAB1.SWISABW50Z = ""
lYSWISAB1.SWISABW52A = ""

lYSWISAB1.SWISABW59P = ""
lYSWISAB1.SWISABW59Z = ""
lYSWISAB1.SWISABW57A = ""

lYSWISAB1.SWISABWEBA = ""
lYSWISAB1.SWISABW71A = ""

End Function
