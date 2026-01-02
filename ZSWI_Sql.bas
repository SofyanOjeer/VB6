Attribute VB_Name = "sqlZSWI"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset

Public Function sqlZSWIENA0_Update(newY As typeZSWIENA0, oldY As typeZSWIENA0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlZSWIENA0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SWIENAINT <> newY.SWIENAINT Then
    sqlZSWIENA0_Update = "Erreur SWIENAINT : " & newY.SWIENAINT
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWIENAINT = " & oldY.SWIENAINT

xSet = " set" 'xSet & " set SWIENAWIDH = " & newY.SWIENAWIDH
blnUpdate = False


' Détecter les modifications
'===================================================================================

If newY.SWIENACET <> oldY.SWIENACET Then blnUpdate = True:  xSet = xSet & " , SWIENACET = '" & newY.SWIENACET & "'"
If newY.SWIENAREF <> oldY.SWIENAREF Then blnUpdate = True:  xSet = xSet & " , SWIENAREF = '" & newY.SWIENAREF & "'"


If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  '"
    xSql = "update " & paramIBM_Library_SAB & ".ZSWIENA0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlZSWIENA0_Update = "Erreur màj : " & newY.SWIENAINT
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlZSWIENA0_Update = Error
End Function

















