Attribute VB_Name = "sqlZMNUHLB0"
Option Explicit

Public Function sqlZMNUHLB0_Update(newY As typeZMNUHLB0, oldY As typeZMNUHLB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlZMNUHLB0_Update = Null

xWhere = " where MNUHLBETB = " & oldY.MNUHLBETB _
         & " and MNUHLBREF = " & oldY.MNUHLBREF _
         & " and MNUHLBCLA = '" & oldY.MNUHLBCLA & "'" _
         & " and MNUHLBNOM = '" & oldY.MNUHLBNOM & "'"
xSet = xSet & " set"

' Détecter les modifications
'===================================================================================
If newY.MNUHLBETB <> oldY.MNUHLBETB Then xSet = xSet & " , MNUHLBETB = " & newY.MNUHLBETB
If newY.MNUHLBREF <> oldY.MNUHLBREF Then xSet = xSet & " , MNUHLBREF = " & newY.MNUHLBREF
If newY.MNUHLBCLA <> oldY.MNUHLBCLA Then xSet = xSet & " , MNUHLBCLA = '" & Text_Apostrophe(newY.MNUHLBCLA) & "'"
If newY.MNUHLBNOM <> oldY.MNUHLBNOM Then xSet = xSet & " , MNUHLBNOM = '" & Text_Apostrophe(newY.MNUHLBNOM) & "'"
If newY.MNUHLBVAL <> oldY.MNUHLBVAL Then xSet = xSet & " , MNUHLBVAL = '" & Text_Apostrophe(newY.MNUHLBVAL) & "'"
If newY.MNUHLBDBD <> oldY.MNUHLBDBD Then xSet = xSet & " , MNUHLBDBD = " & newY.MNUHLBDBD
If newY.MNUHLBDBH <> oldY.MNUHLBDBH Then xSet = xSet & " , MNUHLBDBH = " & newY.MNUHLBDBH
If newY.MNUHLBFID <> oldY.MNUHLBFID Then xSet = xSet & " , MNUHLBFID = " & newY.MNUHLBFID
If newY.MNUHLBFIH <> oldY.MNUHLBFIH Then xSet = xSet & " , MNUHLBFIH = " & newY.MNUHLBFIH
If newY.MNUHLBSUS <> oldY.MNUHLBSUS Then xSet = xSet & " , MNUHLBSUS = " & newY.MNUHLBSUS
If newY.MNUHLBSDT <> oldY.MNUHLBSDT Then xSet = xSet & " , MNUHLBSDT = " & newY.MNUHLBSDT
If newY.MNUHLBSHE <> oldY.MNUHLBSHE Then xSet = xSet & " , MNUHLBSHE = " & newY.MNUHLBSHE

Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SAB & ".ZMNUHLB0" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZMNUHLB0_Update = "Erreur màj : " & newY.MNUHLBCLA & newY.MNUHLBNOM
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZMNUHLB0_Update = Error
End Function


