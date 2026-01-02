Attribute VB_Name = "sqlZMNUMEN0"
Option Explicit

Public Function sqlZMNUMEN0_Insert(newY As typeZMNUMEN0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZMNUMEN0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

'If newY.MNUMENETB <> 0 Then
xSet = xSet & ",MNUMENETB": xValues = xValues & " ," & newY.MNUMENETB
'If newY.MNUMENCGR <> 0 Then
xSet = xSet & ",MNUMENREF": xValues = xValues & " ," & newY.MNUMENREF
xSet = xSet & ",MNUMENGRP": xValues = xValues & " ,'" & newY.MNUMENGRP & "'"
'If newY.MNUMENPRE <> 0 Then
xSet = xSet & ",MNUMENPRE": xValues = xValues & " ," & newY.MNUMENPRE
'If newY.MNUMENORD <> 0 Then
xSet = xSet & ",MNUMENORD": xValues = xValues & " ," & newY.MNUMENORD
'If newY.MNUMENCOD <> 0 Then
xSet = xSet & ",MNUMENCOD": xValues = xValues & " ," & newY.MNUMENCOD
If Trim(newY.MNUMENOIA) <> "" Then xSet = xSet & ",MNUMENOIA": xValues = xValues & " ,'" & newY.MNUMENOIA & "'"
If Trim(newY.MNUMENJOQ) <> "" Then xSet = xSet & ",MNUMENJOQ": xValues = xValues & " ,'" & Text_Apostrophe(newY.MNUMENJOQ) & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "

xSQL = "Insert into " & paramIBM_Library_SAB & ".ZMNUMEN0" & xSet & ")" & xValues & ")"
Debug.Print xSQL
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZMNUMEN0_Insert = "Erreur màj : " & newY.MNUMENCOD
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZMNUMEN0_Insert = Error
End Function
Public Function sqlZMNUMEN0_Update(newY As typeZMNUMEN0, oldY As typeZMNUMEN0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlZMNUMEN0_Update = Null

xWhere = " where MNUMENETB = " & oldY.MNUMENETB _
         & " and MNUMENREF = " & oldY.MNUMENREF _
         & " and MNUMENGRP  = '" & Replace(Trim(newY.MNUMENGRP), "'", "''") & "'" _
         & " and MNUMENPRE = " & oldY.MNUMENPRE _
         & " and MNUMENORD = " & oldY.MNUMENORD
xSet = xSet & " set"

' Détecter les modifications
'===================================================================================
If newY.MNUMENETB <> oldY.MNUMENETB Then xSet = xSet & " , MNUMENETB = " & newY.MNUMENETB
If newY.MNUMENREF <> oldY.MNUMENREF Then xSet = xSet & " , MNUMENREF = " & newY.MNUMENREF
If newY.MNUMENGRP <> oldY.MNUMENGRP Then xSet = xSet & " , MNUMENGRP = '" & Replace(Trim(newY.MNUMENGRP), "'", "''") & "'"
If newY.MNUMENPRE <> oldY.MNUMENPRE Then xSet = xSet & " , MNUMENPRE = " & newY.MNUMENPRE
If newY.MNUMENORD <> oldY.MNUMENORD Then xSet = xSet & " , MNUMENORD = " & newY.MNUMENORD
If newY.MNUMENCOD <> oldY.MNUMENCOD Then xSet = xSet & " , MNUMENCOD = " & newY.MNUMENCOD
If newY.MNUMENOIA <> oldY.MNUMENOIA Then xSet = xSet & " , MNUMENOIA = '" & newY.MNUMENOIA & "'"
If newY.MNUMENJOQ <> oldY.MNUMENJOQ Then xSet = xSet & " , MNUMENJOQ = '" & Replace(Trim(newY.MNUMENJOQ), "'", "''") & "'"

Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SAB & ".ZMNUMEN0" & xSet & xWhere
Debug.Print xSQL
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZMNUMEN0_Update = "Erreur màj : " & newY.MNUMENCOD
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZMNUMEN0_Update = Error
End Function


