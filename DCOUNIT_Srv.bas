Attribute VB_Name = "srvDCOUNIT"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsado As ADODB.Recordset
 
Type typeDCOUNIT
 
      DCOUSTA     As String * 1   ' STATUT
      DCOUVER     As Integer      ' No VERSION
      DCOUPER     As Long         ' PERIODE TRAITEMENT
      DCOUCRTA    As Long         ' CODE RENTA
      DCOUNUO     As Double       ' NOMBRE UNITAIRE OEUVRE
      DCOUCOU     As Currency     ' MONTANT COUT UNITAIRE
      DCOUMAJ     As Long         ' Sequence mise à jour

End Type
Public xDCOUNIT As typeDCOUNIT
Public Function sqlDCOUNIT_Insert(newY As typeDCOUNIT, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDCOUNIT_Insert = Null

xSet = " (DCOUVER"
xValues = " values(" & newY.DCOUVER

' Insertion : Normalement tout est à créer
'===================================================================================
If Trim(newY.DCOUSTA) <> "" Then xSet = xSet & ",DCOUSTA": xValues = xValues & ", '" & newY.DCOUSTA & "'"
If newY.DCOUPER <> 0 Then xSet = xSet & ",DCOUPER": xValues = xValues & ", " & newY.DCOUPER
If newY.DCOUCRTA <> 0 Then xSet = xSet & ",DCOUCRTA": xValues = xValues & ", " & newY.DCOUCRTA
If newY.DCOUNUO <> 0 Then xSet = xSet & ",DCOUNUO": xValues = xValues & ", " & Comma_Point(newY.DCOUNUO)
If newY.DCOUCOU <> 0 Then xSet = xSet & ",DCOUCOU": xValues = xValues & ", " & cur_P(newY.DCOUCOU)

xSql = "Insert into BODWH.DCOUNIT" & xSet & ")" & xValues & ")"

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCOUNIT_Insert = "Erreur màj : " & newY.DCOUVER & newY.DCOUPER & newY.DCOUCRTA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCOUNIT_Insert = Error
End Function

Public Function sqlDCOUNIT_Delete(newY As typeDCOUNIT, oldY As typeDCOUNIT, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDCOUNIT_Delete = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DCOUVER <> newY.DCOUVER Or oldY.DCOUPER <> newY.DCOUPER Or oldY.DCOUCRTA <> newY.DCOUCRTA Then
    sqlDCOUNIT_Delete = "Clé erronnée lors de la suppression !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Delete'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DCOUVER = " & oldY.DCOUVER & " and DCOUPER = " & oldY.DCOUPER _
         & " and DCOUCRTA = " & oldY.DCOUCRTA & " and DCOUMAJ = " & oldY.DCOUMAJ

' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DCOUNIT" & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCOUNIT_Delete = "Erreur SUPP : " & oldY.DCOUVER & oldY.DCOUPER & oldY.DCOUCRTA & oldY.DCOUMAJ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCOUNIT_Delete = Error
End Function

Public Function sqlDCOUNIT_Update(newY As typeDCOUNIT, oldY As typeDCOUNIT, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDCOUNIT_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DCOUVER <> newY.DCOUVER Or oldY.DCOUPER <> newY.DCOUPER Or oldY.DCOUCRTA <> newY.DCOUCRTA Then
    sqlDCOUNIT_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DCOUVER = " & oldY.DCOUVER & " and DCOUPER = " & oldY.DCOUPER _
         & " and DCOUCRTA = " & oldY.DCOUCRTA & " and DCOUMAJ = " & oldY.DCOUMAJ

newY.DCOUMAJ = newY.DCOUMAJ + 1
xSet = xSet & " set DCOUMAJ = " & newY.DCOUMAJ

' Détecter les modifications
'===================================================================================
If Trim(newY.DCOUSTA) <> "" Then xSet = xSet & ", DCOUSTA='" & newY.DCOUSTA & "'"
If newY.DCOUNUO <> 0 Then xSet = xSet & ", DCOUNUO=" & Comma_Point(newY.DCOUNUO)
If newY.DCOUCOU <> 0 Then xSet = xSet & ", DCOUCOU=" & cur_P(newY.DCOUCOU)

xSql = "update BODWH.DCOUNIT" & xSet & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCOUNIT_Update = "Erreur màj : " & newY.DCOUVER & newY.DCOUPER & newY.DCOUCRTA & newY.DCOUMAJ

    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCOUNIT_Update = Error
End Function


Public Function srvDCOUNIT_GetBuffer_ODBC(rsado As ADODB.Recordset, lDCOUNIT As typeDCOUNIT)

On Error GoTo Error_Handler

srvDCOUNIT_GetBuffer_ODBC = Null

lDCOUNIT.DCOUSTA = rsado("DCOUSTA")
lDCOUNIT.DCOUVER = rsado("DCOUVER")
lDCOUNIT.DCOUPER = rsado("DCOUPER")
lDCOUNIT.DCOUCRTA = rsado("DCOUCRTA")
lDCOUNIT.DCOUNUO = rsado("DCOUNUO")
lDCOUNIT.DCOUCOU = rsado("DCOUCOU")
lDCOUNIT.DCOUMAJ = rsado("DCOUMAJ")

Exit Function
Error_Handler:
srvDCOUNIT_GetBuffer_ODBC = Error

End Function

Public Function srvDCOUNIT_Init(lDCOUNIT As typeDCOUNIT)

lDCOUNIT.DCOUSTA = ""
lDCOUNIT.DCOUVER = 0
lDCOUNIT.DCOUPER = 0
lDCOUNIT.DCOUCRTA = 0
lDCOUNIT.DCOUNUO = 0
lDCOUNIT.DCOUCOU = 0
lDCOUNIT.DCOUMAJ = 0

End Function

Public Sub srvDCOUNIT_fgDisplay(lDCOUNIT As typeDCOUNIT, fgDisplay As MSFlexGrid)

fgDisplay.Rows = 8

fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "DCOUSTA     1A"
fgDisplay.Col = 1: fgDisplay = "Statut"
fgDisplay.Col = 2: fgDisplay = lDCOUNIT.DCOUSTA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "DCOUVER     1S0"
fgDisplay.Col = 1: fgDisplay = "Version"
fgDisplay.Col = 2: fgDisplay = lDCOUNIT.DCOUVER
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "DCOUPER     8S0"
fgDisplay.Col = 1: fgDisplay = "Période de traitement"
fgDisplay.Col = 2: fgDisplay = lDCOUNIT.DCOUPER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "DCOUCRTA    5S0"
fgDisplay.Col = 1: fgDisplay = "Code renta"
fgDisplay.Col = 2: fgDisplay = lDCOUNIT.DCOUCRTA
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "DCOUNUO     5S2"
fgDisplay.Col = 1: fgDisplay = "Nombre d'Unité d'oeuvre"
fgDisplay.Col = 2: fgDisplay = lDCOUNIT.DCOUNUO
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "DCOUCOU    18P2"
fgDisplay.Col = 1: fgDisplay = "Montant coût unitaire"
fgDisplay.Col = 2: fgDisplay = lDCOUNIT.DCOUCOU
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "DCOUMAJ     5S0"
fgDisplay.Col = 1: fgDisplay = "Séquence mise à jour"
fgDisplay.Col = 2: fgDisplay = lDCOUNIT.DCOUMAJ

End Sub


