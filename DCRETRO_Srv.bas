Attribute VB_Name = "srvDCRETRO"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsado As ADODB.Recordset
 
Type typeDCRETRO
 
      DRETSTA     As String * 1   ' STATUT
      DRETVER     As Integer      ' No VERSION
      DRETPER     As Long         ' PERIODE TRAITEMENT
      DRETETB     As String * 2   ' CODE ETABLISSEMENT
      DRETAGE     As String * 2   ' CODE AGENCE
      DRETSER     As String * 2   ' CODE SERVICE
      DRETSSE     As String * 2   ' CODE SOUS-SERVICE
      DRETOPE     As String * 3   ' CODE OPERATION
      DRETNUM     As Long         ' NO OPERATION
      DRETDTR     As Long         ' DATE DE TRAITEMENT
      DRETSEQ     As Integer      ' No SEQUENCE
      DRETEVT     As String * 3   ' CODE EVENEMENT
      DRETNAT     As String * 6   ' CODE NATURE
      DRETREF     As String * 15  ' NOTRE REFERENCE
      DRETCLI     As Long         ' NO CLIENT
      DRETDEV     As String * 3   ' CODE DEVISE
      DRETMNT1    As Currency     ' MONTANT COMMISSION 1
      DRETMNT2    As Currency     ' MONTANT COMMISSION 2
      DRETCLR     As Long         ' CLIENT RENTA
      DRETCRTA    As Long         ' CODE RENTA
      DRETCTG     As Integer      ' COMPTAGE
      DRETMAJ     As Long         ' Sequence mise à jour
      
End Type
Public xDCRETRO As typeDCRETRO
Public Function sqlDCRETRO_Insert(newY As typeDCRETRO, cnADO As ADODB.Connection)
  Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDCRETRO_Insert = Null

xSet = " (DRETVER"
xValues = " values(" & newY.DRETVER

' Insertion : Normalement tout est à créer
'===================================================================================
If Trim(newY.DRETSTA) <> "" Then xSet = xSet & ",DRETSTA": xValues = xValues & ", '" & newY.DRETSTA & "'"
If newY.DRETPER <> 0 Then xSet = xSet & ",DRETPER": xValues = xValues & ", " & newY.DRETPER
If Trim(newY.DRETETB) <> "" Then xSet = xSet & ",DRETETB": xValues = xValues & ", '" & newY.DRETETB & "'"
If Trim(newY.DRETAGE) <> "" Then xSet = xSet & ",DRETAGE": xValues = xValues & ", '" & newY.DRETAGE & "'"
If Trim(newY.DRETSER) <> "" Then xSet = xSet & ",DRETSER": xValues = xValues & ", '" & newY.DRETSER & "'"
If Trim(newY.DRETSSE) <> "" Then xSet = xSet & ",DRETSSE": xValues = xValues & ", '" & newY.DRETSSE & "'"
If Trim(newY.DRETOPE) <> "" Then xSet = xSet & ",DRETOPE": xValues = xValues & ", '" & newY.DRETOPE & "'"
If newY.DRETNUM <> 0 Then xSet = xSet & ",DRETNUM": xValues = xValues & ", " & newY.DRETNUM
If newY.DRETDTR <> 0 Then xSet = xSet & ",DRETDTR": xValues = xValues & ", " & newY.DRETDTR
If newY.DRETSEQ <> 0 Then xSet = xSet & ",DRETSEQ": xValues = xValues & ", " & newY.DRETSEQ
If Trim(newY.DRETEVT) <> "" Then xSet = xSet & ",DRETEVT": xValues = xValues & ", '" & newY.DRETEVT & "'"
If Trim(newY.DRETNAT) <> "" Then xSet = xSet & ",DRETNAT": xValues = xValues & ", '" & newY.DRETNAT & "'"
If Trim(newY.DRETREF) <> "" Then xSet = xSet & ",DRETREF": xValues = xValues & ", '" & newY.DRETREF & "'"
If newY.DRETCLI <> 0 Then xSet = xSet & ",DRETCLI": xValues = xValues & ", " & newY.DRETCLI
If Trim(newY.DRETDEV) <> "" Then xSet = xSet & ",DRETDEV": xValues = xValues & ", '" & newY.DRETDEV & "'"
If newY.DRETMNT1 <> 0 Then xSet = xSet & ",DRETMNT1": xValues = xValues & ", " & cur_P(newY.DRETMNT1)
If newY.DRETMNT2 <> 0 Then xSet = xSet & ",DRETMNT2": xValues = xValues & ", " & cur_P(newY.DRETMNT2)
If newY.DRETCLR <> 0 Then xSet = xSet & ",DRETCLR": xValues = xValues & ", " & newY.DRETCLR
If newY.DRETCRTA <> 0 Then xSet = xSet & ",DRETCRTA": xValues = xValues & ", " & newY.DRETCRTA
If newY.DRETCTG <> 0 Then xSet = xSet & ",DRETCTG": xValues = xValues & ", " & newY.DRETCTG

xSql = "Insert into BODWH.DCRETRO" & xSet & ")" & xValues & ")"

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCRETRO_Insert = "Erreur màj : " & newY.DRETVER & newY.DRETPER & newY.DRETETB & newY.DRETAGE & newY.DRETSER & newY.DRETSSE & newY.DRETOPE & newY.DRETNUM & newY.DRETDTR & newY.DRETSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCRETRO_Insert = Error
End Function

Public Function sqlDCRETRO_Delete(newY As typeDCRETRO, oldY As typeDCRETRO, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDCRETRO_Delete = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DRETVER <> newY.DRETVER Or oldY.DRETPER <> newY.DRETPER Or oldY.DRETETB <> newY.DRETETB Or _
oldY.DRETAGE <> newY.DRETAGE Or oldY.DRETSER <> newY.DRETSER Or oldY.DRETSSE <> newY.DRETSSE Or _
oldY.DRETOPE <> newY.DRETOPE Or oldY.DRETNUM <> newY.DRETNUM Or oldY.DRETDTR <> newY.DRETDTR Or _
oldY.DRETSEQ <> newY.DRETSEQ Then
    sqlDCRETRO_Delete = "Clé erronnée lors de la suppression !"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Delete'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DRETVER = " & oldY.DRETVER & " and DRETPER = " & oldY.DRETPER _
         & " and DRETETB = '" & oldY.DRETETB & "' and DRETAGE = '" & oldY.DRETAGE & "'" _
         & " and DRETSER = '" & oldY.DRETSER & "' and DRETSSE = '" & oldY.DRETSSE & "'" _
         & " and DRETOPE = '" & oldY.DRETOPE & "' and DRETNUM = " & oldY.DRETNUM _
         & " and DRETDTR = " & oldY.DRETDTR & " and DRETSEQ = " & oldY.DRETSEQ _
         & " and DRETMAJ = " & oldY.DRETMAJ

' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DCRETRO" & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCRETRO_Delete = "Erreur SUPP : " & oldY.DRETVER & oldY.DRETPER & oldY.DRETETB & oldY.DRETAGE & oldY.DRETSER & oldY.DRETSSE & oldY.DRETOPE & oldY.DRETNUM & oldY.DRETDTR & oldY.DRETSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCRETRO_Delete = Error
End Function

Public Function sqlDCRETRO_Update(newY As typeDCRETRO, oldY As typeDCRETRO, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDCRETRO_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DRETVER <> newY.DRETVER Or oldY.DRETPER <> newY.DRETPER Or oldY.DRETETB <> newY.DRETETB Or _
oldY.DRETAGE <> newY.DRETAGE Or oldY.DRETSER <> newY.DRETSER Or oldY.DRETSSE <> newY.DRETSSE Or _
oldY.DRETOPE <> newY.DRETOPE Or oldY.DRETNUM <> newY.DRETNUM Or oldY.DRETDTR <> newY.DRETDTR Or _
oldY.DRETSEQ <> newY.DRETSEQ Then
    sqlDCRETRO_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DRETVER = " & oldY.DRETVER & " and DRETPER = " & oldY.DRETPER _
         & " and DRETETB = '" & oldY.DRETETB & "' and DRETAGE = '" & oldY.DRETAGE & "'" _
         & " and DRETSER = '" & oldY.DRETSER & "' and DRETSSE = '" & oldY.DRETSSE & "'" _
         & " and DRETOPE = '" & oldY.DRETOPE & "' and DRETNUM = " & oldY.DRETNUM _
         & " and DRETDTR = " & oldY.DRETDTR & " and DRETSEQ = " & oldY.DRETSEQ _
         & " and DRETMAJ = " & oldY.DRETMAJ

newY.DRETMAJ = newY.DRETMAJ + 1
xSet = xSet & " set DRETMAJ = " & newY.DRETMAJ

' Détecter les modifications
'===================================================================================
If Trim(newY.DRETSTA) <> "" Then xSet = xSet & ", DRETSTA='" & newY.DRETSTA & "'"
If Trim(newY.DRETEVT) <> "" Then xSet = xSet & ", DRETEVT='" & newY.DRETEVT & "'"
If Trim(newY.DRETNAT) <> "" Then xSet = xSet & ", DRETNAT='" & newY.DRETNAT & "'"
If Trim(newY.DRETREF) <> "" Then xSet = xSet & ", DRETREF='" & newY.DRETREF & "'"
If newY.DRETCLI <> 0 Then xSet = xSet & ", DRETCLI=" & newY.DRETCLI
If Trim(newY.DRETDEV) <> "" Then xSet = xSet & ", DRETDEV='" & newY.DRETDEV & "'"
' If newY.DRETMNT1 <> 0 Then xSet = xSet & ", DRETMNT1=" & cur_P(newY.DRETMNT1)
' If newY.DRETMNT2 <> 0 Then xSet = xSet & ", DRETMNT2=" & cur_P(newY.DRETMNT2)
xSet = xSet & ", DRETMNT1=" & cur_P(newY.DRETMNT1)
xSet = xSet & ", DRETMNT2=" & cur_P(newY.DRETMNT2)
If newY.DRETCLR <> 0 Then xSet = xSet & ", DRETCLR=" & newY.DRETCLR
If newY.DRETCRTA <> 0 Then xSet = xSet & ", DRETCRTA=" & newY.DRETCRTA
If newY.DRETCTG <> 0 Then xSet = xSet & ", DRETCTG=" & newY.DRETCTG

xSql = "update BODWH.DCRETRO" & xSet & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCRETRO_Update = "Erreur màj : " & newY.DRETVER & newY.DRETPER & newY.DRETETB & newY.DRETAGE & newY.DRETSER & newY.DRETSSE & newY.DRETOPE & newY.DRETNUM & newY.DRETDTR & newY.DRETSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCRETRO_Update = Error
End Function


Public Function srvDCRETRO_GetBuffer_ODBC(rsado As ADODB.Recordset, lDCRETRO As typeDCRETRO)

On Error GoTo Error_Handler

srvDCRETRO_GetBuffer_ODBC = Null

lDCRETRO.DRETSTA = rsado("DRETSTA")
lDCRETRO.DRETVER = rsado("DRETVER")
lDCRETRO.DRETPER = rsado("DRETPER")
lDCRETRO.DRETETB = rsado("DRETETB")
lDCRETRO.DRETAGE = rsado("DRETAGE")
lDCRETRO.DRETSER = rsado("DRETSER")
lDCRETRO.DRETSSE = rsado("DRETSSE")
lDCRETRO.DRETOPE = rsado("DRETOPE")
lDCRETRO.DRETNUM = rsado("DRETNUM")
lDCRETRO.DRETDTR = rsado("DRETDTR")
lDCRETRO.DRETSEQ = rsado("DRETSEQ")
lDCRETRO.DRETEVT = rsado("DRETEVT")
lDCRETRO.DRETNAT = rsado("DRETNAT")
lDCRETRO.DRETREF = rsado("DRETREF")
lDCRETRO.DRETCLI = rsado("DRETCLI")
lDCRETRO.DRETDEV = rsado("DRETDEV")
lDCRETRO.DRETMNT1 = rsado("DRETMNT1")
lDCRETRO.DRETMNT2 = rsado("DRETMNT2")
lDCRETRO.DRETCLR = rsado("DRETCLR")
lDCRETRO.DRETCRTA = rsado("DRETCRTA")
lDCRETRO.DRETCTG = rsado("DRETCTG")
lDCRETRO.DRETMAJ = rsado("DRETMAJ")

Exit Function
Error_Handler:
srvDCRETRO_GetBuffer_ODBC = Error

End Function

Public Function srvDCRETRO_Init(lDCRETRO As typeDCRETRO)

lDCRETRO.DRETSTA = ""
lDCRETRO.DRETVER = 0
lDCRETRO.DRETPER = 0
lDCRETRO.DRETETB = ""
lDCRETRO.DRETAGE = ""
lDCRETRO.DRETSER = ""
lDCRETRO.DRETSSE = ""
lDCRETRO.DRETOPE = ""
lDCRETRO.DRETNUM = 0
lDCRETRO.DRETDTR = 0
lDCRETRO.DRETSEQ = 0
lDCRETRO.DRETEVT = ""
lDCRETRO.DRETNAT = ""
lDCRETRO.DRETREF = ""
lDCRETRO.DRETCLI = 0
lDCRETRO.DRETDEV = ""
lDCRETRO.DRETMNT1 = 0
lDCRETRO.DRETMNT2 = 0
lDCRETRO.DRETCLR = 0
lDCRETRO.DRETCRTA = 0
lDCRETRO.DRETCTG = 0
lDCRETRO.DRETMAJ = 0

End Function

Public Sub srvDCRETRO_fgDisplay(lDCRETRO As typeDCRETRO, fgDisplay As MSFlexGrid)

fgDisplay.Rows = 23

fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "DRETSTA     1A"
fgDisplay.Col = 1: fgDisplay = "Statut"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETSTA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "DRETVER     1S0"
fgDisplay.Col = 1: fgDisplay = "Version"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETVER
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "DRETPER     8S0"
fgDisplay.Col = 1: fgDisplay = "Période de traitement"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETPER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "DRETETB     2A"
fgDisplay.Col = 1: fgDisplay = "Code établissement"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETETB
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "DRETAGE     2A"
fgDisplay.Col = 1: fgDisplay = "Code agence"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETAGE
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "DRETSER     2A"
fgDisplay.Col = 1: fgDisplay = "Code service"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETSER
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "DRETSSE     2A"
fgDisplay.Col = 1: fgDisplay = "Code sous-service"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETSSE
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "DRETOPE     3A"
fgDisplay.Col = 1: fgDisplay = "Code opération"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETOPE
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "DRETNUM     9P0"
fgDisplay.Col = 1: fgDisplay = "No opération"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETNUM
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "DRETDTR     8S0"
fgDisplay.Col = 1: fgDisplay = "Date de traitement"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETDTR
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "DRETSEQ     3S0"
fgDisplay.Col = 1: fgDisplay = "No séquence"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETSEQ
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "DRETEVT     3A"
fgDisplay.Col = 1: fgDisplay = "Code évènement"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETOPE
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "DRETNAT     6A"
fgDisplay.Col = 1: fgDisplay = "Code nature"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETNAT
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "DRETREF    15A"
fgDisplay.Col = 1: fgDisplay = "Notre référence"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETREF
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "DRETCLI     7S0"
fgDisplay.Col = 1: fgDisplay = "No client"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETCLI
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "DRETDEV     3A"
fgDisplay.Col = 1: fgDisplay = "Code devise"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETDEV
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "DRETMNT1   18P2"
fgDisplay.Col = 1: fgDisplay = "Montant commission 1"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETMNT1
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "DRETMNT2   18P2"
fgDisplay.Col = 1: fgDisplay = "Montant commission 2"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETMNT2
fgDisplay.Row = 19
fgDisplay.Col = 0: fgDisplay = "DRETCLR     7S0"
fgDisplay.Col = 1: fgDisplay = "No client renta"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETCLR
fgDisplay.Row = 20
fgDisplay.Col = 0: fgDisplay = "DRETCRTA    5S0"
fgDisplay.Col = 1: fgDisplay = "Code renta"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETCRTA
fgDisplay.Row = 21
fgDisplay.Col = 0: fgDisplay = "DRETCTG     1S0"
fgDisplay.Col = 1: fgDisplay = "Comptage"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETCTG
fgDisplay.Row = 22
fgDisplay.Col = 0: fgDisplay = "DRETMAJ     5S0"
fgDisplay.Col = 1: fgDisplay = "Séquence mise à jour"
fgDisplay.Col = 2: fgDisplay = lDCRETRO.DRETMAJ

End Sub


