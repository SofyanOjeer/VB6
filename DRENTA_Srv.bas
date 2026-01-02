Attribute VB_Name = "srvDRENTA"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsado As ADODB.Recordset
 
Type typeDRENTA
 
      DRTASTA     As String * 1   ' STATUT
      DRTAVER     As Integer      ' No VERSION
      DRTAPER     As Long         ' PERIODE TRAITEMENT
      DRTAETA     As String * 2   ' CODE ETABLISSEMENT
      DRTACLIA    As String * 1   ' BLANC / T
      DRTACLIB    As Long         ' CODE MATRICULE
      DRTACRTA    As Long         ' CODE RENTA
      DRTACGRP    As Long         ' CODE RENTA REGROUPEMENT
      DRTAMOYB    As Currency     ' MONTANT MOYENNE -BASE
      DRTACTR     As Long         ' COMPTAGE
      DRTAMMRB    As Currency     ' MONTANT MARGE RENTA -BASE
      DRTATXM     As Double       ' Taux de marge
      DRTAMAJ     As Long         ' Sequence mise à jour

End Type
Public xDRENTA As typeDRENTA
Public Function sqlDRENTA_Insert(newY As typeDRENTA, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDRENTA_Insert = Null

xSet = " (DRTAVER"
xValues = " values(" & newY.DRTAVER

' Insertion : Normalement tout est à créer
'===================================================================================
If Trim(newY.DRTASTA) <> "" Then xSet = xSet & ",DRTASTA": xValues = xValues & ", '" & newY.DRTASTA & "'"
'  If newY.DRTAVER <> 0 Then xSet = xSet & ",DRTAVER": xValues = xValues & " ," & newY.DRTAVER
If newY.DRTAPER <> 0 Then xSet = xSet & ",DRTAPER": xValues = xValues & ", " & newY.DRTAPER
If Trim(newY.DRTAETA) <> "" Then xSet = xSet & ",DRTAETA": xValues = xValues & ", '" & newY.DRTAETA & "'"
If Trim(newY.DRTACLIA) <> "" Then xSet = xSet & ",DRTACLIA": xValues = xValues & ", '" & newY.DRTACLIA & "'"
If newY.DRTACLIB <> 0 Then xSet = xSet & ",DRTACLIB": xValues = xValues & ", " & newY.DRTACLIB
If newY.DRTACRTA <> 0 Then xSet = xSet & ",DRTACRTA": xValues = xValues & ", " & newY.DRTACRTA
If newY.DRTACGRP <> 0 Then xSet = xSet & ",DRTACGRP": xValues = xValues & ", " & newY.DRTACGRP
If newY.DRTAMOYB <> 0 Then xSet = xSet & ",DRTAMOYB": xValues = xValues & ", " & cur_P(newY.DRTAMOYB)
If newY.DRTACTR <> 0 Then xSet = xSet & ",DRTACTR": xValues = xValues & ", " & newY.DRTACTR
If newY.DRTAMMRB <> 0 Then xSet = xSet & ",DRTAMMRB": xValues = xValues & ", " & cur_P(newY.DRTAMMRB)
If newY.DRTATXM <> 0 Then xSet = xSet & ",DRTATXM": xValues = xValues & ", " & Comma_Point(newY.DRTATXM)

xSql = "Insert into BODWH.DRENTA" & xSet & ")" & xValues & ")"

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDRENTA_Insert = "Erreur màj : " & newY.DRTAVER & newY.DRTAPER & newY.DRTAETA & newY.DRTACLIA & newY.DRTACLIB & newY.DRTACRTA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDRENTA_Insert = Error
End Function

Public Function sqlDRENTA_Delete(newY As typeDRENTA, oldY As typeDRENTA, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDRENTA_Delete = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DRTAVER <> newY.DRTAVER Or oldY.DRTAPER <> newY.DRTAPER Or oldY.DRTAETA <> newY.DRTAETA Or _
oldY.DRTACLIA <> newY.DRTACLIA Or oldY.DRTACLIB <> newY.DRTACLIB Or oldY.DRTACRTA <> newY.DRTACRTA Then
    sqlDRENTA_Delete = "Clé erronnée lors de la suppression !"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Delete'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DRTAVER = " & oldY.DRTAVER & " and DRTAPER = " & oldY.DRTAPER _
         & " and DRTAETA = '" & oldY.DRTAETA & "' and DRTACLIA = '" & oldY.DRTACLIA & "'" _
         & " and DRTACLIB = " & oldY.DRTACLIB & " and DRTACRTA = " & oldY.DRTACRTA _
         & " and DRTAMAJ = " & oldY.DRTAMAJ

' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DRENTA" & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDRENTA_Delete = "Erreur SUPP : " & oldY.DRTAVER & oldY.DRTAPER & oldY.DRTAETA & oldY.DRTACLIA & oldY.DRTACLIB & oldY.DRTACRTA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDRENTA_Delete = Error
End Function


Public Function sqlDRENTA_Delete_ForRename(oldY As typeDRENTA, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDRENTA_Delete_ForRename = Null

' Contrôle  : Même clé d'accès old / new  >>>>>> NON NECESSAIRE POUR FONCTION -RENOMMER-
'=======================================================================================
'If oldY.DRTAVER <> newY.DRTAVER Or oldY.DRTAPER <> newY.DRTAPER Or oldY.DRTAETA <> newY.DRTAETA Or _
'oldY.DRTACLIA <> newY.DRTACLIA Or oldY.DRTACLIB <> newY.DRTACLIB Or oldY.DRTACRTA <> newY.DRTACRTA Then
'    sqlDRENTA_Delete = "Clé erronnée lors de la suppression !"
'    Exit Function
'End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Delete'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DRTAVER = " & oldY.DRTAVER & " and DRTAPER = " & oldY.DRTAPER _
         & " and DRTAETA = '" & oldY.DRTAETA & "' and DRTACLIA = '" & oldY.DRTACLIA & "'" _
         & " and DRTACLIB = " & oldY.DRTACLIB & " and DRTACRTA = " & oldY.DRTACRTA _
         & " and DRTAMAJ = " & oldY.DRTAMAJ

' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DRENTA" & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDRENTA_Delete_ForRename = "Erreur SUPP : " & oldY.DRTAVER & oldY.DRTAPER & oldY.DRTAETA & oldY.DRTACLIA & oldY.DRTACLIB & oldY.DRTACRTA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDRENTA_Delete_ForRename = Error
End Function


Public Function sqlDRENTA_Update(newY As typeDRENTA, oldY As typeDRENTA, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDRENTA_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DRTAVER <> newY.DRTAVER Or oldY.DRTAPER <> newY.DRTAPER Or oldY.DRTAETA <> newY.DRTAETA Or _
oldY.DRTACLIA <> newY.DRTACLIA Or oldY.DRTACLIB <> newY.DRTACLIB Or oldY.DRTACRTA <> newY.DRTACRTA Then
    sqlDRENTA_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DRTAVER = " & oldY.DRTAVER & " and DRTAPER = " & oldY.DRTAPER _
         & " and DRTAETA = '" & oldY.DRTAETA & "' and DRTACLIA = '" & oldY.DRTACLIA & "'" _
         & " and DRTACLIB = " & oldY.DRTACLIB & " and DRTACRTA = " & oldY.DRTACRTA _
         & " and DRTAMAJ = " & oldY.DRTAMAJ

newY.DRTAMAJ = newY.DRTAMAJ + 1
xSet = xSet & " set DRTAMAJ = " & newY.DRTAMAJ

' Détecter les modifications
'===================================================================================
If Trim(newY.DRTASTA) <> "" Then xSet = xSet & ", DRTASTA='" & newY.DRTASTA & "'"
If newY.DRTACGRP <> 0 Then xSet = xSet & ", DRTACGRP=" & newY.DRTACGRP
If newY.DRTAMOYB <> 0 Then xSet = xSet & ", DRTAMOYB=" & cur_P(newY.DRTAMOYB)
If newY.DRTACTR <> 0 Then xSet = xSet & ", DRTACTR=" & newY.DRTACTR
If newY.DRTAMMRB <> 0 Then xSet = xSet & ", DRTAMMRB=" & cur_P(newY.DRTAMMRB)
If newY.DRTATXM <> 0 Then xSet = xSet & ", DRTATXM=" & Comma_Point(newY.DRTATXM)

xSql = "update BODWH.DRENTA" & xSet & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDRENTA_Update = "Erreur màj : " & newY.DRTAVER & newY.DRTAPER & newY.DRTAETA & newY.DRTACLIA & newY.DRTACLIB & newY.DRTACRTA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDRENTA_Update = Error
End Function


Public Function srvDRENTA_GetBuffer_ODBC(rsado As ADODB.Recordset, lDRENTA As typeDRENTA)

On Error GoTo Error_Handler

srvDRENTA_GetBuffer_ODBC = Null

lDRENTA.DRTASTA = rsado("DRTASTA")
lDRENTA.DRTAVER = rsado("DRTAVER")
lDRENTA.DRTAPER = rsado("DRTAPER")
lDRENTA.DRTAETA = rsado("DRTAETA")
lDRENTA.DRTACLIA = rsado("DRTACLIA")
lDRENTA.DRTACLIB = rsado("DRTACLIB")
lDRENTA.DRTACRTA = rsado("DRTACRTA")
lDRENTA.DRTACGRP = rsado("DRTACGRP")
lDRENTA.DRTAMOYB = rsado("DRTAMOYB")
lDRENTA.DRTACTR = rsado("DRTACTR")
lDRENTA.DRTAMMRB = rsado("DRTAMMRB")
lDRENTA.DRTATXM = rsado("DRTATXM")
lDRENTA.DRTAMAJ = rsado("DRTAMAJ")

Exit Function
Error_Handler:
srvDRENTA_GetBuffer_ODBC = Error

End Function

Public Function srvDRENTA_Init(lDRENTA As typeDRENTA)

lDRENTA.DRTASTA = ""
lDRENTA.DRTAVER = 0
lDRENTA.DRTAPER = 0
lDRENTA.DRTAETA = ""
lDRENTA.DRTACLIA = ""
lDRENTA.DRTACLIB = 0
lDRENTA.DRTACRTA = 0
lDRENTA.DRTACGRP = 0
lDRENTA.DRTAMOYB = 0
lDRENTA.DRTACTR = 0
lDRENTA.DRTAMMRB = 0
lDRENTA.DRTATXM = 0
lDRENTA.DRTAMAJ = 0

End Function

Public Sub srvDRENTA_fgDisplay(lDRENTA As typeDRENTA, fgDisplay As MSFlexGrid)

fgDisplay.Rows = 14

fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "DRTASTA     1A"
fgDisplay.Col = 1: fgDisplay = "Statut"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTASTA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "DRTAVER     1S0"
fgDisplay.Col = 1: fgDisplay = "Version"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTAVER
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "DRTAPER     8S0"
fgDisplay.Col = 1: fgDisplay = "Période de traitement"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTAPER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "DRTAETA     2A"
fgDisplay.Col = 1: fgDisplay = "Code établissement"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTAETA
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "DRTACLIA    1A"
fgDisplay.Col = 1: fgDisplay = "Blanc / T"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTACLIA
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "DRTACLIB    7S0"
fgDisplay.Col = 1: fgDisplay = "No Matricule"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTACLIB
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "DRTACRTA    5S0"
fgDisplay.Col = 1: fgDisplay = "Code renta"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTACRTA
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "DRTACGRP    5S0"
fgDisplay.Col = 1: fgDisplay = "Code regroupement"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTACGRP
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "DRTAMOYB   18P2"
fgDisplay.Col = 1: fgDisplay = "Montant moyenne -BASE"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTAMOYB
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "DRTACTR     5S0"
fgDisplay.Col = 1: fgDisplay = "Comptage"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTACTR
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "DRTAMMRB   18P2"
fgDisplay.Col = 1: fgDisplay = "Montant marge renta -BASE"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTAMMRB
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "DRTATXM   14P9"
fgDisplay.Col = 1: fgDisplay = "Taux de marge -BASE"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTATXM
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "DRTAMAJ     5S0"
fgDisplay.Col = 1: fgDisplay = "Séquence mise à jour"
fgDisplay.Col = 2: fgDisplay = lDRENTA.DRTAMAJ

End Sub


