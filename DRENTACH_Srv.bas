Attribute VB_Name = "srvDRENTACH"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsado As ADODB.Recordset
 
Type typeDRENTACH
 
      DRCHSTA     As String * 1   ' STATUT
      DRCHVER     As Integer      ' No VERSION
      DRCHPER     As Long         ' PERIODE TRAITEMENT
      DRCHETA     As String * 2   ' CODE ETABLISSEMENT
      DRCHCLIA    As String * 1   ' BLANC / T
      DRCHCLIB    As Long         ' CODE MATRICULE
      DRCHCRTA    As Long         ' CODE RENTA
      DRCHCGRP    As Long         ' CODE RENTA REGROUPEMENT
      DRCHCTR     As Long         ' COMPTAGE
      DRCHMMRB    As Currency     ' MONTANT MARGE RENTA -BASE
      DRCHMAJ     As Long         ' Sequence mise à jour

End Type
Public xDRENTACH As typeDRENTACH

Public Function sqlDRENTACH_Insert(newY As typeDRENTACH, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDRENTACH_Insert = Null

xSet = " (DRCHVER"
xValues = " values(" & newY.DRCHVER

' Insertion : Normalement tout est à créer :
'    CHARGES : MONTANT SYSTEMATIQUEMENT MULTIPLIE PAR (-1) dans cmdDRENTACH_Charger
'===================================================================================
If Trim(newY.DRCHSTA) <> "" Then xSet = xSet & ",DRCHSTA": xValues = xValues & ", '" & newY.DRCHSTA & "'"
'  If newY.DRCHVER <> 0 Then xSet = xSet & ",DRCHVER": xValues = xValues & " ," & newY.DRCHVER
If newY.DRCHPER <> 0 Then xSet = xSet & ",DRCHPER": xValues = xValues & ", " & newY.DRCHPER
If Trim(newY.DRCHETA) <> "" Then xSet = xSet & ",DRCHETA": xValues = xValues & ", '" & newY.DRCHETA & "'"
If Trim(newY.DRCHCLIA) <> "" Then xSet = xSet & ",DRCHCLIA": xValues = xValues & ", '" & newY.DRCHCLIA & "'"
If newY.DRCHCLIB <> 0 Then xSet = xSet & ",DRCHCLIB": xValues = xValues & ", " & newY.DRCHCLIB
If newY.DRCHCRTA <> 0 Then xSet = xSet & ",DRCHCRTA": xValues = xValues & ", " & newY.DRCHCRTA
If newY.DRCHCGRP <> 0 Then xSet = xSet & ",DRCHCGRP": xValues = xValues & ", " & newY.DRCHCGRP
If newY.DRCHCTR <> 0 Then xSet = xSet & ",DRCHCTR": xValues = xValues & ", " & newY.DRCHCTR
If newY.DRCHMMRB <> 0 Then xSet = xSet & ",DRCHMMRB": xValues = xValues & ", " & cur_P(newY.DRCHMMRB)

xSql = "Insert into BODWH.DRENTACH" & xSet & ")" & xValues & ")"

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDRENTACH_Insert = "Erreur màj : " & newY.DRCHVER & newY.DRCHPER & newY.DRCHETA & newY.DRCHCLIA & newY.DRCHCLIB & newY.DRCHCRTA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDRENTACH_Insert = Error
End Function

Public Function sqlDRENTACH_Delete(newY As typeDRENTACH, oldY As typeDRENTACH, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDRENTACH_Delete = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DRCHVER <> newY.DRCHVER Or oldY.DRCHPER <> newY.DRCHPER Or oldY.DRCHETA <> newY.DRCHETA Or _
oldY.DRCHCLIA <> newY.DRCHCLIA Or oldY.DRCHCLIB <> newY.DRCHCLIB Or oldY.DRCHCRTA <> newY.DRCHCRTA Then
    sqlDRENTACH_Delete = "Clé erronnée lors de la suppression !"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Delete'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DRCHVER = " & oldY.DRCHVER & " and DRCHPER = " & oldY.DRCHPER _
         & " and DRCHETA = '" & oldY.DRCHETA & "' and DRCHCLIA = '" & oldY.DRCHCLIA & "'" _
         & " and DRCHCLIB = " & oldY.DRCHCLIB & " and DRCHCRTA = " & oldY.DRCHCRTA _
         & " and DRCHMAJ = " & oldY.DRCHMAJ

' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DRENTACH" & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDRENTACH_Delete = "Erreur SUPP : " & oldY.DRCHVER & oldY.DRCHPER & oldY.DRCHETA & oldY.DRCHCLIA & oldY.DRCHCLIB & oldY.DRCHCRTA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDRENTACH_Delete = Error
End Function

Public Function sqlDRENTACH_Update(newY As typeDRENTACH, oldY As typeDRENTACH, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDRENTACH_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DRCHVER <> newY.DRCHVER Or oldY.DRCHPER <> newY.DRCHPER Or oldY.DRCHETA <> newY.DRCHETA Or _
oldY.DRCHCLIA <> newY.DRCHCLIA Or oldY.DRCHCLIB <> newY.DRCHCLIB Or oldY.DRCHCRTA <> newY.DRCHCRTA Then
    sqlDRENTACH_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DRCHVER = " & oldY.DRCHVER & " and DRCHPER = " & oldY.DRCHPER _
         & " and DRCHETA = '" & oldY.DRCHETA & "' and DRCHCLIA = '" & oldY.DRCHCLIA & "'" _
         & " and DRCHCLIB = " & oldY.DRCHCLIB & " and DRCHCRTA = " & oldY.DRCHCRTA _
         & " and DRCHMAJ = " & oldY.DRCHMAJ

newY.DRCHMAJ = newY.DRCHMAJ + 1
xSet = xSet & " set DRCHMAJ = " & newY.DRCHMAJ

' Détecter les modifications
'    CHARGES : MONTANT SYSTEMATIQUEMENT MULTIPLIE PAR (-1) dans cmdDRENTACH_Charger
'===================================================================================
If Trim(newY.DRCHSTA) <> "" Then xSet = xSet & ", DRCHSTA='" & newY.DRCHSTA & "'"
If newY.DRCHCGRP <> 0 Then xSet = xSet & ", DRCHCGRP=" & newY.DRCHCGRP
If newY.DRCHCTR <> 0 Then xSet = xSet & ", DRCHCTR=" & newY.DRCHCTR
If newY.DRCHMMRB <> 0 Then xSet = xSet & ", DRCHMMRB=" & cur_P(newY.DRCHMMRB)

xSql = "update BODWH.DRENTACH" & xSet & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDRENTACH_Update = "Erreur màj : " & newY.DRCHVER & newY.DRCHPER & newY.DRCHETA & newY.DRCHCLIA & newY.DRCHCLIB & newY.DRCHCRTA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDRENTACH_Update = Error
End Function


Public Function srvDRENTACH_GetBuffer_ODBC(rsado As ADODB.Recordset, lDRENTACH As typeDRENTACH)

On Error GoTo Error_Handler

srvDRENTACH_GetBuffer_ODBC = Null

lDRENTACH.DRCHSTA = rsado("DRCHSTA")
lDRENTACH.DRCHVER = rsado("DRCHVER")
lDRENTACH.DRCHPER = rsado("DRCHPER")
lDRENTACH.DRCHETA = rsado("DRCHETA")
lDRENTACH.DRCHCLIA = rsado("DRCHCLIA")
lDRENTACH.DRCHCLIB = rsado("DRCHCLIB")
lDRENTACH.DRCHCRTA = rsado("DRCHCRTA")
lDRENTACH.DRCHCGRP = rsado("DRCHCGRP")
lDRENTACH.DRCHCTR = rsado("DRCHCTR")
lDRENTACH.DRCHMMRB = rsado("DRCHMMRB")
lDRENTACH.DRCHMAJ = rsado("DRCHMAJ")

Exit Function
Error_Handler:
srvDRENTACH_GetBuffer_ODBC = Error

End Function

Public Function srvDRENTACH_Init(lDRENTACH As typeDRENTACH)

lDRENTACH.DRCHSTA = ""
lDRENTACH.DRCHVER = 0
lDRENTACH.DRCHPER = 0
lDRENTACH.DRCHETA = ""
lDRENTACH.DRCHCLIA = ""
lDRENTACH.DRCHCLIB = 0
lDRENTACH.DRCHCRTA = 0
lDRENTACH.DRCHCGRP = 0
lDRENTACH.DRCHCTR = 0
lDRENTACH.DRCHMMRB = 0
lDRENTACH.DRCHMAJ = 0

End Function

Public Sub srvDRENTACH_fgDisplay(lDRENTACH As typeDRENTACH, fgDisplay As MSFlexGrid)

fgDisplay.Rows = 12

fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "DRCHSTA     1A"
fgDisplay.Col = 1: fgDisplay = "Statut"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHSTA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "DRCHVER     1S0"
fgDisplay.Col = 1: fgDisplay = "Version"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHVER
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "DRCHPER     8S0"
fgDisplay.Col = 1: fgDisplay = "Période de traitement"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHPER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "DRCHETA     2A"
fgDisplay.Col = 1: fgDisplay = "Code établissement"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHETA
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "DRCHCLIA    1A"
fgDisplay.Col = 1: fgDisplay = "Blanc / T"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHCLIA
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "DRCHCLIB    7S0"
fgDisplay.Col = 1: fgDisplay = "No Matricule"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHCLIB
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "DRCHCRTA    5S0"
fgDisplay.Col = 1: fgDisplay = "Code renta"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHCRTA
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "DRCHCGRP    5S0"
fgDisplay.Col = 1: fgDisplay = "Code regroupement"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHCGRP
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "DRCHCTR     5S0"
fgDisplay.Col = 1: fgDisplay = "Comptage"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHCTR
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "DRCHMMRB   18P2"
fgDisplay.Col = 1: fgDisplay = "Montant marge renta -BASE"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHMMRB
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "DRCHMAJ     5S0"
fgDisplay.Col = 1: fgDisplay = "Séquence mise à jour"
fgDisplay.Col = 2: fgDisplay = lDRENTACH.DRCHMAJ

End Sub


