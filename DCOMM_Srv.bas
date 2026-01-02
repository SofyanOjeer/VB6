Attribute VB_Name = "srvDCOMM"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsado As ADODB.Recordset
 
Type typeDCOMM
 
      DCOMSTA     As String * 1   ' STATUT
      DCOMVER     As Integer      ' No VERSION
      DCOMPER     As Long         ' PERIODE TRAITEMENT
      DCOMETA     As String * 2   ' CODE ETABLISSEMENT
      DCOMAGE     As String * 2   ' CODE AGENCE
      DCOMSER     As String * 2   ' CODE SERVICE
      DCOMSES     As String * 2   ' CODE SOUS-SERVICE
      DCOMOPE     As String * 3   ' CODE OPERATION
      DCOMNAT     As String * 6   ' CODE NATURE
      DCOMNUM     As Long         ' NO OPERATION
      DCOMSEN     As String * 1   ' SENS
      DCOMSEQ     As Integer      ' No SEQUENCE
      DCOMCLE     As String * 20  ' CLE = COM+NUM+SEQ
      DCOMCLIA    As String * 1   '   / T
      DCOMCLIB    As Long         ' NO CLIENT
      DCOMCOM     As String * 6   ' CODE COMMISSION
      DCOMCRTA    As Long         ' CODE RENTA
      DCOMCTL1    As String * 10   ' CODE CONTROLE 1
      DCOMCTL2    As String * 10   ' CODE CONTROLE 2
      DCOMMONB    As Currency     ' MONTANT COMMISSION BASE
      DCOMMOND    As Currency     ' MONTANT COMMISSION DEVISE
      DCOMDEV     As String * 3   ' CODE DEVISE
      DCOMMAJ     As Long         ' Sequence mise à jour
      
End Type
Public xDCOMM As typeDCOMM
Public Function sqlDCOMM_Insert(newY As typeDCOMM, cnADO As ADODB.Connection)
  Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDCOMM_Insert = Null

xSet = " (DCOMVER"
xValues = " values(" & newY.DCOMVER

' Insertion : Normalement tout est à créer
'===================================================================================
If Trim(newY.DCOMSTA) <> "" Then xSet = xSet & ",DCOMSTA": xValues = xValues & ", '" & newY.DCOMSTA & "'"
If newY.DCOMPER <> 0 Then xSet = xSet & ",DCOMPER": xValues = xValues & ", " & newY.DCOMPER
If Trim(newY.DCOMETA) <> "" Then xSet = xSet & ",DCOMETA": xValues = xValues & ", '" & newY.DCOMETA & "'"
If Trim(newY.DCOMAGE) <> "" Then xSet = xSet & ",DCOMAGE": xValues = xValues & ", '" & newY.DCOMAGE & "'"
If Trim(newY.DCOMSER) <> "" Then xSet = xSet & ",DCOMSER": xValues = xValues & ", '" & newY.DCOMSER & "'"
If Trim(newY.DCOMSES) <> "" Then xSet = xSet & ",DCOMSES": xValues = xValues & ", '" & newY.DCOMSES & "'"
If Trim(newY.DCOMOPE) <> "" Then xSet = xSet & ",DCOMOPE": xValues = xValues & ", '" & newY.DCOMOPE & "'"
If Trim(newY.DCOMNAT) <> "" Then xSet = xSet & ",DCOMNAT": xValues = xValues & ", '" & newY.DCOMNAT & "'"
If newY.DCOMNUM <> 0 Then xSet = xSet & ",DCOMNUM": xValues = xValues & ", " & newY.DCOMNUM
If Trim(newY.DCOMSEN) <> "" Then xSet = xSet & ",DCOMSEN": xValues = xValues & ", '" & newY.DCOMSEN & "'"
If newY.DCOMSEQ <> 0 Then xSet = xSet & ",DCOMSEQ": xValues = xValues & ", " & newY.DCOMSEQ
If Trim(newY.DCOMCLE) <> "" Then xSet = xSet & ",DCOMCLE": xValues = xValues & ", '" & newY.DCOMCLE & "'"
If Trim(newY.DCOMCLIA) <> "" Then xSet = xSet & ",DCOMCLIA": xValues = xValues & ", '" & newY.DCOMCLIA & "'"
If newY.DCOMCLIB <> 0 Then xSet = xSet & ",DCOMCLIB": xValues = xValues & ", " & newY.DCOMCLIB
If Trim(newY.DCOMCOM) <> "" Then xSet = xSet & ",DCOMCOM": xValues = xValues & ", '" & newY.DCOMCOM & "'"
If Trim(newY.DCOMCTL1) <> "" Then xSet = xSet & ",DCOMCTL1": xValues = xValues & ", '" & newY.DCOMCTL1 & "'"
If Trim(newY.DCOMCTL2) <> "" Then xSet = xSet & ",DCOMCTL2": xValues = xValues & ", '" & newY.DCOMCTL2 & "'"
If newY.DCOMMONB <> 0 Then xSet = xSet & ",DCOMMONB": xValues = xValues & ", " & cur_P(newY.DCOMMONB)
If newY.DCOMMOND <> 0 Then xSet = xSet & ",DCOMMOND": xValues = xValues & ", " & cur_P(newY.DCOMMOND)
If Trim(newY.DCOMDEV) <> "" Then xSet = xSet & ",DCOMDEV": xValues = xValues & ", '" & newY.DCOMDEV & "'"
If newY.DCOMCRTA <> 0 Then xSet = xSet & ",DCOMCRTA": xValues = xValues & ", " & newY.DCOMCRTA

xSql = "Insert into BODWH.DCOMM" & xSet & ")" & xValues & ")"

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCOMM_Insert = "Erreur màj : " & newY.DCOMVER & newY.DCOMPER & newY.DCOMETA & newY.DCOMAGE & newY.DCOMSER & newY.DCOMSES & newY.DCOMOPE & newY.DCOMNAT & newY.DCOMNUM & newY.DCOMSEN & newY.DCOMSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCOMM_Insert = Error
End Function

Public Function sqlDCOMM_Delete(newY As typeDCOMM, oldY As typeDCOMM, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDCOMM_Delete = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DCOMVER <> newY.DCOMVER Or oldY.DCOMPER <> newY.DCOMPER Or oldY.DCOMETA <> newY.DCOMETA Or _
oldY.DCOMAGE <> newY.DCOMAGE Or oldY.DCOMSER <> newY.DCOMSER Or oldY.DCOMSES <> newY.DCOMSES Or _
oldY.DCOMOPE <> newY.DCOMOPE Or oldY.DCOMNAT <> newY.DCOMNAT Or oldY.DCOMNUM <> newY.DCOMNUM Or _
oldY.DCOMSEN <> newY.DCOMSEN Or oldY.DCOMSEQ <> newY.DCOMSEQ Then
    sqlDCOMM_Delete = "Clé erronnée lors de la suppression !"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Delete'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DCOMVER = " & oldY.DCOMVER & " and DCOMPER = " & oldY.DCOMPER _
         & " and DCOMETA = '" & oldY.DCOMETA & "' and DCOMAGE = '" & oldY.DCOMAGE & "'" _
         & " and DCOMSER = '" & oldY.DCOMSER & "' and DCOMSES = '" & oldY.DCOMSES & "'" _
         & " and DCOMOPE = '" & oldY.DCOMOPE & "' and DCOMNAT = '" & oldY.DCOMNAT & "'" _
         & " and DCOMNUM = " & oldY.DCOMNUM & " and DCOMSEN = '" & oldY.DCOMSEN & "'" _
         & " and DCOMSEQ = " & oldY.DCOMSEQ & " and DCOMMAJ = " & oldY.DCOMMAJ

' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DCOMM" & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCOMM_Delete = "Erreur SUPP : " & oldY.DCOMVER & oldY.DCOMPER & oldY.DCOMETA & oldY.DCOMAGE & oldY.DCOMSER & oldY.DCOMSES & oldY.DCOMOPE & oldY.DCOMNAT & oldY.DCOMNUM & oldY.DCOMSEN & oldY.DCOMSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCOMM_Delete = Error
End Function

Public Function sqlDCOMM_Update(newY As typeDCOMM, oldY As typeDCOMM, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDCOMM_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DCOMVER <> newY.DCOMVER Or oldY.DCOMPER <> newY.DCOMPER Or oldY.DCOMETA <> newY.DCOMETA Or _
oldY.DCOMAGE <> newY.DCOMAGE Or oldY.DCOMSER <> newY.DCOMSER Or oldY.DCOMSES <> newY.DCOMSES Or _
oldY.DCOMOPE <> newY.DCOMOPE Or oldY.DCOMNAT <> newY.DCOMNAT Or oldY.DCOMNUM <> newY.DCOMNUM Or _
oldY.DCOMSEN <> newY.DCOMSEN Or oldY.DCOMSEQ <> newY.DCOMSEQ Then
    sqlDCOMM_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DCOMVER = " & oldY.DCOMVER & " and DCOMPER = " & oldY.DCOMPER _
         & " and DCOMETA = '" & oldY.DCOMETA & "' and DCOMAGE = '" & oldY.DCOMAGE & "'" _
         & " and DCOMSER = '" & oldY.DCOMSER & "' and DCOMSES = '" & oldY.DCOMSES & "'" _
         & " and DCOMOPE = '" & oldY.DCOMOPE & "' and DCOMNAT = '" & oldY.DCOMNAT & "'" _
         & " and DCOMNUM = " & oldY.DCOMNUM & " and DCOMSEN = '" & oldY.DCOMSEN & "'" _
         & " and DCOMSEQ = " & oldY.DCOMSEQ & " and DCOMMAJ = " & oldY.DCOMMAJ

newY.DCOMMAJ = newY.DCOMMAJ + 1
xSet = xSet & " set DCOMMAJ = " & newY.DCOMMAJ

' Détecter les modifications
'===================================================================================
If Trim(newY.DCOMSTA) <> "" Then xSet = xSet & ", DCOMSTA='" & newY.DCOMSTA & "'"
If Trim(newY.DCOMCLE) <> "" Then xSet = xSet & ", DCOMCLE='" & newY.DCOMCLE & "'"
If Trim(newY.DCOMCLIA) <> "" Then xSet = xSet & ", DCOMCLIA='" & newY.DCOMCLIA & "'"
If newY.DCOMCLIB <> 0 Then xSet = xSet & ", DCOMCLIB=" & newY.DCOMCLIB
If Trim(newY.DCOMCOM) <> "" Then xSet = xSet & ", DCOMCOM='" & newY.DCOMCOM & "'"
If newY.DCOMCRTA <> 0 Then xSet = xSet & ", DCOMCRTA=" & newY.DCOMCRTA
xSet = xSet & ", DCOMCTL1='" & newY.DCOMCTL1 & "'"
xSet = xSet & ", DCOMCTL2='" & newY.DCOMCTL2 & "'"
xSet = xSet & ", DCOMMONB=" & cur_P(newY.DCOMMONB)
xSet = xSet & ", DCOMMOND=" & cur_P(newY.DCOMMOND)
If Trim(newY.DCOMDEV) <> "" Then xSet = xSet & ", DCOMDEV='" & newY.DCOMDEV & "'"

xSql = "update BODWH.DCOMM" & xSet & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCOMM_Update = "Erreur màj : " & newY.DCOMVER & newY.DCOMPER & newY.DCOMETA & newY.DCOMAGE & newY.DCOMSER & newY.DCOMSES & newY.DCOMOPE & newY.DCOMNAT & newY.DCOMNUM & newY.DCOMSEN & newY.DCOMSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCOMM_Update = Error
End Function


Public Function srvDCOMM_GetBuffer_ODBC(rsado As ADODB.Recordset, lDCOMM As typeDCOMM)

On Error GoTo Error_Handler

srvDCOMM_GetBuffer_ODBC = Null

lDCOMM.DCOMSTA = rsado("DCOMSTA")
lDCOMM.DCOMVER = rsado("DCOMVER")
lDCOMM.DCOMPER = rsado("DCOMPER")
lDCOMM.DCOMETA = rsado("DCOMETA")
lDCOMM.DCOMAGE = rsado("DCOMAGE")
lDCOMM.DCOMSER = rsado("DCOMSER")
lDCOMM.DCOMSES = rsado("DCOMSES")
lDCOMM.DCOMOPE = rsado("DCOMOPE")
lDCOMM.DCOMNAT = rsado("DCOMNAT")
lDCOMM.DCOMNUM = rsado("DCOMNUM")
lDCOMM.DCOMSEN = rsado("DCOMSEN")
lDCOMM.DCOMSEQ = rsado("DCOMSEQ")
lDCOMM.DCOMCLE = rsado("DCOMCLE")
lDCOMM.DCOMCLIA = rsado("DCOMCLIA")
lDCOMM.DCOMCLIB = rsado("DCOMCLIB")
lDCOMM.DCOMCOM = rsado("DCOMCOM")
lDCOMM.DCOMCRTA = rsado("DCOMCRTA")
lDCOMM.DCOMCTL1 = rsado("DCOMCTL1")
lDCOMM.DCOMCTL2 = rsado("DCOMCTL2")
lDCOMM.DCOMMONB = rsado("DCOMMONB")
lDCOMM.DCOMMOND = rsado("DCOMMOND")
lDCOMM.DCOMDEV = rsado("DCOMDEV")
lDCOMM.DCOMMAJ = rsado("DCOMMAJ")

Exit Function
Error_Handler:
srvDCOMM_GetBuffer_ODBC = Error

End Function

Public Function srvDCOMM_Init(lDCOMM As typeDCOMM)

lDCOMM.DCOMSTA = ""
lDCOMM.DCOMVER = 0
lDCOMM.DCOMPER = 0
lDCOMM.DCOMETA = ""
lDCOMM.DCOMAGE = ""
lDCOMM.DCOMSER = ""
lDCOMM.DCOMSES = ""
lDCOMM.DCOMOPE = ""
lDCOMM.DCOMNAT = ""
lDCOMM.DCOMNUM = 0
lDCOMM.DCOMSEN = ""
lDCOMM.DCOMSEQ = 0
lDCOMM.DCOMCLE = ""
lDCOMM.DCOMCLIA = ""
lDCOMM.DCOMCLIB = 0
lDCOMM.DCOMCOM = ""
lDCOMM.DCOMCRTA = 0
lDCOMM.DCOMCTL1 = ""
lDCOMM.DCOMCTL2 = ""
lDCOMM.DCOMMONB = 0
lDCOMM.DCOMMOND = 0
lDCOMM.DCOMDEV = ""
lDCOMM.DCOMMAJ = 0

End Function

Public Sub srvDCOMM_fgDisplay(lDCOMM As typeDCOMM, fgDisplay As MSFlexGrid)

fgDisplay.Rows = 24

fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "DCOMSTA     1A"
fgDisplay.Col = 1: fgDisplay = "Statut"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMSTA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "DCOMVER     1S0"
fgDisplay.Col = 1: fgDisplay = "Version"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMVER
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "DCOMPER     8S0"
fgDisplay.Col = 1: fgDisplay = "Période de traitement"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMPER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "DCOMETA     2A"
fgDisplay.Col = 1: fgDisplay = "Code établissement"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMETA
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "DCOMAGE     2A"
fgDisplay.Col = 1: fgDisplay = "Code agence"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMAGE
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "DCOMSER     2A"
fgDisplay.Col = 1: fgDisplay = "Code service"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMSER
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "DCOMSES     2A"
fgDisplay.Col = 1: fgDisplay = "Code sous-service"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMSES
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "DCOMOPE     3A"
fgDisplay.Col = 1: fgDisplay = "Code opération"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMOPE
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "DCOMNAT     6A"
fgDisplay.Col = 1: fgDisplay = "Nature"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMNAT
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "DCOMNUM     9P0"
fgDisplay.Col = 1: fgDisplay = "No opération"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMNUM
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "DCOMSEN     1A"
fgDisplay.Col = 1: fgDisplay = "Sens"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMSEN
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "DCOMSEQ     3S0"
fgDisplay.Col = 1: fgDisplay = "No séquence"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMSEQ
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "DCOMCLE    20A"
fgDisplay.Col = 1: fgDisplay = "CLE"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMCLE
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "DCOMCLIA   15A"
fgDisplay.Col = 1: fgDisplay = " /T"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMCLIA
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "DCOMCLIB    7S0"
fgDisplay.Col = 1: fgDisplay = "No client"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMCLIB
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "DCOMCOM     6A"
fgDisplay.Col = 1: fgDisplay = "Code commission"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMCOM
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "DCOMCRTA    5S0"
fgDisplay.Col = 1: fgDisplay = "Code renta"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMCRTA
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "DCOMCTL1   10A"
fgDisplay.Col = 1: fgDisplay = "Code contrôle 1"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMCTL1
fgDisplay.Row = 19
fgDisplay.Col = 0: fgDisplay = "DCOMCTL1   10A"
fgDisplay.Col = 1: fgDisplay = "Code contrôle 1"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMCTL1
fgDisplay.Row = 20
fgDisplay.Col = 0: fgDisplay = "DCOMMONB   18P2"
fgDisplay.Col = 1: fgDisplay = "Montant commission Base"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMMONB
fgDisplay.Row = 21
fgDisplay.Col = 0: fgDisplay = "DCOMMOND   18P2"
fgDisplay.Col = 1: fgDisplay = "Montant commission Devise"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMMOND
fgDisplay.Row = 22
fgDisplay.Col = 0: fgDisplay = "DCOMDEV     3A"
fgDisplay.Col = 1: fgDisplay = "Code devise"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMDEV
fgDisplay.Row = 23
fgDisplay.Col = 0: fgDisplay = "DCOMMAJ     5S0"
fgDisplay.Col = 1: fgDisplay = "Séquence mise à jour"
fgDisplay.Col = 2: fgDisplay = lDCOMM.DCOMMAJ

End Sub


