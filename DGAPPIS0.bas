Attribute VB_Name = "srvZGAPPIS0"
'---------------------------------------------------------
Option Explicit
Type typeDGAPPIS0

    DGAPPISSTA As String * 1  ' statut
    DGAPPISVER As Integer     ' n° version
    DGAPPISPER As String * 8  ' période de traitement
    DGAPPISETA As String * 2  ' code établissement
    DGAPPISSEQ As Long        ' N° séquence (info)
    DGAPPISKAM As String * 3  ' Auto / Manuel
    DGAPPISCLI As Long        ' client
    DGAPPISDEC As Long        ' date échéance
    DGAPPISMTE As Currency    ' cv EUR
    DGAPPISNBJ As Long        ' nb jours 'situation / échéance
    DGAPPISVEC As String * 3  ' classe de durée (01M ..... 10A 99A)



    GAPPISTAB As Integer     ' CODE ETAT
    GAPPISECH As Integer     ' NUMÉRO ÉCHÉANCIER
    GAPPISCLA As Integer     ' N° CLASSE ECHEANCIER
    GAPPISETA As Integer     ' CODE ÉTABLISSEMENT
    GAPPISAGE As Integer     ' CODE AGENCE
    GAPPISSER As String * 2  ' CODE SERVICE
    GAPPISSSE As String * 2  ' CODE SOUS-SERVICE
    GAPPISOPE As String * 3  ' CODE OPÉRATION
    GAPPISNAT As String * 3  ' CODE NATURE
    GAPPISNUO As Long        ' NUMÉRO OPÉRATION
    GAPPISDEV As String * 3  ' DEVISE
    GAPPISSEN As String * 1  ' SENS
    GAPPISDEC As Long        ' DATE ÉCHÉANCE
    GAPPISRUB As String * 10 ' RUBRIQUE COMPTABLE
    GAPPISTPR As String * 9  ' TYPE PRODUIT
    GAPPISCLI As String * 7  ' NUMÉRO CLIENT
    GAPPISMON As Currency    ' MONTANT DU FLUX
    GAPPISTTI As String * 1  ' TYPE DE TAUX INTERNE
    GAPPISTTE As String * 1  ' TYPE DE TAUX EXTERNE
    GAPPISRTV As String * 6  ' CODE TAUX
    GAPPISTAU As Double      ' VALEUR DU TAUX
    GAPPISSOL As Currency    ' SOLDE RUBRI COMPTABL
    GAPPISPOU As Double      ' POURCENTAGE
    GAPPISSIG As String * 13 ' SIGLE DU CLIENT
    GAPPISVIL As String * 12 ' VILLE
    
    GAPPISTMC As Double      ' marge client
    GAPPISDAR As Long        ' DATE ÉCHÉANCE REVISE
    GAPPISVAT As Currency    ' VALEUR ACTUELLE
    GAPPISVAP As Currency    ' VALEUR ACT. PONDEREE

    GAPPISTVF As String         ' code taux BIA F V R
    GAPPISTP1 As Currency    ' variation taux +1%
    GAPPISTP2 As Currency    ' variation taux +2%
    GAPPISTM1 As Currency    ' variation taux -1%
    GAPPISTM2 As Currency    ' variation taux -2%
    GAPPISMAR As Currency    '
End Type

Public ddsDGAPPIS0(46) As String * 50

Type typeALM

    Lib         As String
    Row         As Integer
    Col         As Integer
    Mt1         As Currency
    Mt2         As Currency
    Mt3         As Currency
    Mt4         As Currency
    Mt5         As Currency
End Type

Public Function sqlDGAPPIS0_Delete(oldY As typeDGAPPIS0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDGAPPIS0_Delete = Null


xWhere = " where DGAPPISSTA = '" & oldY.DGAPPISSTA & "'" _
       & " and   DGAPPISVER = " & oldY.DGAPPISVER _
       & " and   DGAPPISPER = '" & oldY.DGAPPISPER & "'" _
       & " and   DGAPPISETA = '" & oldY.DGAPPISETA & "'" _
       & " and   DGAPPISSEQ = " & oldY.DGAPPISSEQ _


' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DGAPPIS0" & xWhere

Set rsSab = cnsab.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDGAPPIS0_Delete = "Erreur SUP : " & oldY.DGAPPISSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDGAPPIS0_Delete = Error
End Function


Public Function sqlDGAPPIS0_Insert(newY As typeDGAPPIS0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDGAPPIS0_Insert = Null

xSet = " (DGAPPISSEQ"
xValues = " values('" & Replace(Trim(newY.DGAPPISSEQ), "'", "''") & "'"

' Insertion :
'===================================================================================
If newY.DGAPPISVER <> 0 Then xSet = xSet & ",DGAPPISVER": xValues = xValues & ", " & newY.DGAPPISVER
If newY.DGAPPISPER <> 0 Then xSet = xSet & ",DGAPPISPER": xValues = xValues & ", " & newY.DGAPPISPER
If newY.DGAPPISCLI <> 0 Then xSet = xSet & ",DGAPPISCLI": xValues = xValues & ", " & newY.DGAPPISCLI
If newY.DGAPPISDEC <> 0 Then xSet = xSet & ",DGAPPISDEC": xValues = xValues & ", " & newY.DGAPPISDEC
If newY.DGAPPISMTE <> 0 Then xSet = xSet & ",DGAPPISMTE": xValues = xValues & ", " & cur_P(newY.DGAPPISMTE)
If newY.DGAPPISNBJ <> 0 Then xSet = xSet & ",DGAPPISNBJ": xValues = xValues & ", " & newY.DGAPPISNBJ
If newY.GAPPISTAB <> 0 Then xSet = xSet & ",GAPPISTAB": xValues = xValues & ", " & newY.GAPPISTAB
If newY.GAPPISECH <> 0 Then xSet = xSet & ",GAPPISECH": xValues = xValues & ", " & newY.GAPPISECH
If newY.GAPPISCLA <> 0 Then xSet = xSet & ",GAPPISCLA": xValues = xValues & ", " & newY.GAPPISCLA
If newY.GAPPISETA <> 0 Then xSet = xSet & ",GAPPISETA": xValues = xValues & ", " & newY.GAPPISETA
If newY.GAPPISAGE <> 0 Then xSet = xSet & ",GAPPISAGE": xValues = xValues & ", " & newY.GAPPISAGE
If newY.GAPPISNUO <> 0 Then xSet = xSet & ",GAPPISNUO": xValues = xValues & ", " & newY.GAPPISNUO
If newY.GAPPISDEC <> 0 Then xSet = xSet & ",GAPPISDEC": xValues = xValues & ", " & newY.GAPPISDEC
If newY.GAPPISMON <> 0 Then xSet = xSet & ",GAPPISMON": xValues = xValues & ", " & cur_P(newY.GAPPISMON)
If newY.GAPPISTAU <> 0 Then xSet = xSet & ",GAPPISTAU": xValues = xValues & ", " & Comma_Point(newY.GAPPISTAU)
If newY.GAPPISSOL <> 0 Then xSet = xSet & ",GAPPISSOL": xValues = xValues & ", " & cur_P(newY.GAPPISSOL)
If newY.GAPPISPOU <> 0 Then xSet = xSet & ",GAPPISPOU": xValues = xValues & ", " & Comma_Point(newY.GAPPISPOU)
If newY.GAPPISDAR <> 0 Then xSet = xSet & ",GAPPISDAR": xValues = xValues & ", " & newY.GAPPISDAR
If newY.GAPPISTMC <> 0 Then xSet = xSet & ",GAPPISTMC": xValues = xValues & ", " & Comma_Point(newY.GAPPISTMC)
If newY.GAPPISVAT <> 0 Then xSet = xSet & ",GAPPISVAT": xValues = xValues & ", " & cur_P(newY.GAPPISVAT)
If newY.GAPPISVAP <> 0 Then xSet = xSet & ",GAPPISVAP": xValues = xValues & ", " & cur_P(newY.GAPPISVAP)

'===================================================================================
If Trim(newY.DGAPPISSTA) <> "" Then xSet = xSet & ",DGAPPISSTA": xValues = xValues & ", '" & Replace(Trim(newY.DGAPPISSTA), "'", "''") & "'"
If Trim(newY.DGAPPISETA) <> "" Then xSet = xSet & ",DGAPPISETA": xValues = xValues & ", '" & Replace(Trim(newY.DGAPPISETA), "'", "''") & "'"
If Trim(newY.DGAPPISKAM) <> "" Then xSet = xSet & ",DGAPPISKAM": xValues = xValues & ", '" & Replace(Trim(newY.DGAPPISKAM), "'", "''") & "'"
If Trim(newY.DGAPPISVEC) <> "" Then xSet = xSet & ",DGAPPISVEC": xValues = xValues & ", '" & Replace(Trim(newY.DGAPPISVEC), "'", "''") & "'"
If Trim(newY.GAPPISSER) <> "" Then xSet = xSet & ",GAPPISSER": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISSER), "'", "''") & "'"
If Trim(newY.GAPPISSSE) <> "" Then xSet = xSet & ",GAPPISSSE": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISSSE), "'", "''") & "'"
If Trim(newY.GAPPISOPE) <> "" Then xSet = xSet & ",GAPPISOPE": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISOPE), "'", "''") & "'"
If Trim(newY.GAPPISNAT) <> "" Then xSet = xSet & ",GAPPISNAT": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISNAT), "'", "''") & "'"
If Trim(newY.GAPPISDEV) <> "" Then xSet = xSet & ",GAPPISDEV": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISDEV), "'", "''") & "'"
If Trim(newY.GAPPISSEN) <> "" Then xSet = xSet & ",GAPPISSEN": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISSEN), "'", "''") & "'"
If Trim(newY.GAPPISRUB) <> "" Then xSet = xSet & ",GAPPISRUB": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISRUB), "'", "''") & "'"
If Trim(newY.GAPPISTPR) <> "" Then xSet = xSet & ",GAPPISTPR": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISTPR), "'", "''") & "'"
If Trim(newY.GAPPISCLI) <> "" Then xSet = xSet & ",GAPPISCLI": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISCLI), "'", "''") & "'"
If Trim(newY.GAPPISTTI) <> "" Then xSet = xSet & ",GAPPISTTI": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISTTI), "'", "''") & "'"
If Trim(newY.GAPPISTTE) <> "" Then xSet = xSet & ",GAPPISTTE": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISTTE), "'", "''") & "'"
If Trim(newY.GAPPISRTV) <> "" Then xSet = xSet & ",GAPPISRTV": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISRTV), "'", "''") & "'"
If Trim(newY.GAPPISSIG) <> "" Then xSet = xSet & ",GAPPISSIG": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISSIG), "'", "''") & "'"
If Trim(newY.GAPPISVIL) <> "" Then xSet = xSet & ",GAPPISVIL": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISVIL), "'", "''") & "'"
If Trim(newY.GAPPISTVF) <> "" Then xSet = xSet & ",GAPPISTVF": xValues = xValues & ", '" & Replace(Trim(newY.GAPPISTVF), "'", "''") & "'"

xSql = "Insert into " & paramIBM_Library_BODWH & ".DGAPPIS0" & xSet & ")" & xValues & ")"

Set rsSab = cnsab.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDGAPPIS0_Insert = "Erreur màj : " & newY.DGAPPISSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDGAPPIS0_Insert = Error
End Function
Public Function sqlDGAPPIS0_Read(oldY As typeDGAPPIS0)
Dim X As String, xSql As String, Nb As Long
Dim V

On Error GoTo Error_Handler
sqlDGAPPIS0_Read = Null

xSql = "select * from " & paramIBM_Library_BODWH & ".DGAPPIS0 " _
       & " where DGAPPISSTA = '" & oldY.DGAPPISSTA & "'" _
       & " and   DGAPPISVER = " & oldY.DGAPPISVER _
       & " and   DGAPPISPER = " & oldY.DGAPPISPER _
       & " and   DGAPPISETA = '" & oldY.DGAPPISETA & "'" _
       & " and   DGAPPISSEQ = " & oldY.DGAPPISSEQ

Set rsSab = cnsab.Execute(xSql)

If rsSab.EOF Then
    sqlDGAPPIS0_Read = "? inconnu"
Else
    V = rsDGAPPIS0_GetBuffer(rsSab, oldY)
    If Not IsNull(V) Then sqlDGAPPIS0_Read = "? srvDGAPPIS0_GetBuffer"
End If
 
Exit Function
Error_Handler:
    sqlDGAPPIS0_Read = Error
End Function


Public Function sqlDGAPPIS0_Update(newY As typeDGAPPIS0, oldY As typeDGAPPIS0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
On Error GoTo Error_Handler
sqlDGAPPIS0_Update = Null
blnUpdate = False

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DGAPPISSEQ <> newY.DGAPPISSEQ Then
    sqlDGAPPIS0_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================
xWhere = " where DGAPPISSTA = '" & oldY.DGAPPISSTA & "'" _
       & " and   DGAPPISVER = " & oldY.DGAPPISVER _
       & " and   DGAPPISPER = " & oldY.DGAPPISPER _
       & " and   DGAPPISETA = '" & oldY.DGAPPISETA & "'" _
       & " and   DGAPPISSEQ = " & oldY.DGAPPISSEQ


xSet = " set"

' Détecter les modifications
'===================================================================================
If newY.DGAPPISCLI <> oldY.DGAPPISCLI Then blnUpdate = True:  xSet = xSet & " , DGAPPISCLI = " & newY.DGAPPISCLI
If newY.DGAPPISDEC <> oldY.DGAPPISDEC Then blnUpdate = True:  xSet = xSet & " , DGAPPISDEC = " & newY.DGAPPISDEC
If newY.DGAPPISMTE <> oldY.DGAPPISMTE Then blnUpdate = True:  xSet = xSet & " , dgappismte = " & cur_P(newY.DGAPPISMTE)
If newY.DGAPPISNBJ <> oldY.DGAPPISNBJ Then blnUpdate = True:  xSet = xSet & " , DGAPPISNBJ = " & newY.DGAPPISNBJ

If newY.GAPPISTAB <> oldY.GAPPISTAB Then blnUpdate = True:  xSet = xSet & " , GAPPISTAB = " & newY.GAPPISTAB
If newY.GAPPISECH <> oldY.GAPPISECH Then blnUpdate = True:  xSet = xSet & " , GAPPISECH = " & newY.GAPPISECH
If newY.GAPPISCLA <> oldY.GAPPISCLA Then blnUpdate = True:  xSet = xSet & " , GAPPISCLA = " & newY.GAPPISCLA
If newY.GAPPISETA <> oldY.GAPPISETA Then blnUpdate = True:  xSet = xSet & " , GAPPISETA = " & newY.GAPPISETA
If newY.GAPPISAGE <> oldY.GAPPISAGE Then blnUpdate = True:  xSet = xSet & " , GAPPISAGE = " & newY.GAPPISAGE
If newY.GAPPISNUO <> oldY.GAPPISNUO Then blnUpdate = True:  xSet = xSet & " , GAPPISNUO = " & newY.GAPPISNUO
If newY.GAPPISDEC <> oldY.GAPPISDEC Then blnUpdate = True:  xSet = xSet & " , GAPPISDEC = " & newY.GAPPISDEC
If newY.GAPPISMON <> oldY.GAPPISMON Then blnUpdate = True:  xSet = xSet & " , gappisMON = " & cur_P(newY.GAPPISMON)
If newY.GAPPISTAU <> oldY.GAPPISTAU Then blnUpdate = True:  xSet = xSet & " , GAPPISTAU = " & Comma_Point(newY.GAPPISTAU)
If newY.GAPPISSOL <> oldY.GAPPISSOL Then blnUpdate = True:  xSet = xSet & " , gappisSOL = " & cur_P(newY.GAPPISSOL)
If newY.GAPPISPOU <> oldY.GAPPISPOU Then blnUpdate = True:  xSet = xSet & " , GAPPISPOU = " & Comma_Point(newY.GAPPISPOU)
If newY.GAPPISTMC <> oldY.GAPPISTMC Then blnUpdate = True:  xSet = xSet & " , GAPPISTMC = " & Comma_Point(newY.GAPPISTMC)
If newY.GAPPISDAR <> oldY.GAPPISDAR Then blnUpdate = True:  xSet = xSet & " , GAPPISDAR = " & newY.GAPPISDAR
If newY.GAPPISVAT <> oldY.GAPPISVAT Then blnUpdate = True:  xSet = xSet & " , GAPPISVAT = " & cur_P(newY.GAPPISVAT)
If newY.GAPPISVAP <> oldY.GAPPISVAP Then blnUpdate = True:  xSet = xSet & " , GAPPISVAP = " & cur_P(newY.GAPPISVAP)

If Trim(newY.DGAPPISVEC) <> Trim(oldY.DGAPPISVEC) Then blnUpdate = True: xSet = xSet & ",DGAPPISVEC = '" & Replace(Trim(newY.DGAPPISVEC), "'", "''") & "'"
If Trim(newY.GAPPISSER) <> Trim(oldY.GAPPISSER) Then blnUpdate = True: xSet = xSet & ",GAPPISSER = '" & Replace(Trim(newY.GAPPISSER), "'", "''") & "'"
If Trim(newY.GAPPISSSE) <> Trim(oldY.GAPPISSSE) Then blnUpdate = True: xSet = xSet & ",GAPPISSSE = '" & Replace(Trim(newY.GAPPISSSE), "'", "''") & "'"
If Trim(newY.GAPPISOPE) <> Trim(oldY.GAPPISOPE) Then blnUpdate = True: xSet = xSet & ",GAPPISOPE = '" & Replace(Trim(newY.GAPPISOPE), "'", "''") & "'"
If Trim(newY.GAPPISNAT) <> Trim(oldY.GAPPISNAT) Then blnUpdate = True: xSet = xSet & ",GAPPISNAT = '" & Replace(Trim(newY.GAPPISNAT), "'", "''") & "'"
If Trim(newY.GAPPISDEV) <> Trim(oldY.GAPPISDEV) Then blnUpdate = True: xSet = xSet & ",GAPPISDEV = '" & Replace(Trim(newY.GAPPISDEV), "'", "''") & "'"
If Trim(newY.GAPPISSEN) <> Trim(oldY.GAPPISSEN) Then blnUpdate = True: xSet = xSet & ",GAPPISSEN = '" & Replace(Trim(newY.GAPPISSEN), "'", "''") & "'"
If Trim(newY.GAPPISRUB) <> Trim(oldY.GAPPISRUB) Then blnUpdate = True: xSet = xSet & ",GAPPISRUB = '" & Replace(Trim(newY.GAPPISRUB), "'", "''") & "'"
If Trim(newY.GAPPISTPR) <> Trim(oldY.GAPPISTPR) Then blnUpdate = True: xSet = xSet & ",GAPPISTPR = '" & Replace(Trim(newY.GAPPISTPR), "'", "''") & "'"
If Trim(newY.GAPPISCLI) <> Trim(oldY.GAPPISCLI) Then blnUpdate = True: xSet = xSet & ",GAPPISCLI = '" & Replace(Trim(newY.GAPPISCLI), "'", "''") & "'"
If Trim(newY.GAPPISTTI) <> Trim(oldY.GAPPISTTI) Then blnUpdate = True: xSet = xSet & ",GAPPISTTI = '" & Replace(Trim(newY.GAPPISTTI), "'", "''") & "'"
If Trim(newY.GAPPISTTE) <> Trim(oldY.GAPPISTTE) Then blnUpdate = True: xSet = xSet & ",GAPPISTTE = '" & Replace(Trim(newY.GAPPISTTE), "'", "''") & "'"
If Trim(newY.GAPPISRTV) <> Trim(oldY.GAPPISRTV) Then blnUpdate = True: xSet = xSet & ",GAPPISRTV = '" & Replace(Trim(newY.GAPPISRTV), "'", "''") & "'"
If Trim(newY.GAPPISSIG) <> Trim(oldY.GAPPISSIG) Then blnUpdate = True: xSet = xSet & ",GAPPISSIG = '" & Replace(Trim(newY.GAPPISSIG), "'", "''") & "'"
If Trim(newY.GAPPISVIL) <> Trim(oldY.GAPPISVIL) Then blnUpdate = True: xSet = xSet & ",GAPPISVIL = '" & Replace(Trim(newY.GAPPISVIL), "'", "''") & "'"
If Trim(newY.GAPPISTVF) <> Trim(oldY.GAPPISTVF) Then blnUpdate = True: xSet = xSet & ",GAPPISTVF = '" & Replace(Trim(newY.GAPPISTVF), "'", "''") & "'"

If blnUpdate Then
    If Trim(newY.DGAPPISKAM) = "" Then newY.DGAPPISKAM = "M"
    If newY.DGAPPISKAM <> oldY.DGAPPISKAM Then blnUpdate = True: xSet = xSet & ",DGAPPISKAM = '" & Replace(Trim(newY.DGAPPISKAM), "'", "''") & "'"

    Mid$(xSet, InStr(xSet, ","), 1) = " "
    xSql = "update " & paramIBM_Library_BODWH & ".DGAPPIS0" & xSet & xWhere
    
    Set rsSab = cnsab.Execute(xSql, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlDGAPPIS0_Update = "Erreur màj : " & newY.DGAPPISSEQ
    
        Exit Function
    End If
End If
Exit Function
Error_Handler:
    sqlDGAPPIS0_Update = Error
End Function




'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsDGAPPIS0_GetBuffer(rsado As ADODB.Recordset, rsDGAPPIS0 As typeDGAPPIS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsDGAPPIS0_GetBuffer = Null

rsDGAPPIS0.DGAPPISSTA = rsado("DGAPPISSTA")
rsDGAPPIS0.DGAPPISVER = rsado("DGAPPISVER")
rsDGAPPIS0.DGAPPISPER = rsado("DGAPPISPER")
rsDGAPPIS0.DGAPPISETA = rsado("DGAPPISETA")
rsDGAPPIS0.DGAPPISSEQ = rsado("DGAPPISSEQ")
rsDGAPPIS0.DGAPPISKAM = rsado("DGAPPISKAM")
rsDGAPPIS0.DGAPPISCLI = rsado("DGAPPISCLI")
rsDGAPPIS0.DGAPPISDEC = rsado("DGAPPISDEC")
rsDGAPPIS0.DGAPPISMTE = rsado("DGAPPISMTE")
rsDGAPPIS0.DGAPPISNBJ = rsado("DGAPPISNBJ")
rsDGAPPIS0.DGAPPISVEC = rsado("DGAPPISVEC")


rsDGAPPIS0.GAPPISTAB = rsado("GAPPISTAB")
rsDGAPPIS0.GAPPISECH = rsado("GAPPISECH")
rsDGAPPIS0.GAPPISCLA = rsado("GAPPISCLA")
rsDGAPPIS0.GAPPISETA = rsado("GAPPISETA")
rsDGAPPIS0.GAPPISAGE = rsado("GAPPISAGE")
rsDGAPPIS0.GAPPISSER = rsado("GAPPISSER")
rsDGAPPIS0.GAPPISSSE = rsado("GAPPISSSE")
rsDGAPPIS0.GAPPISOPE = rsado("GAPPISOPE")
rsDGAPPIS0.GAPPISNAT = rsado("GAPPISNAT")
rsDGAPPIS0.GAPPISNUO = rsado("GAPPISNUO")
rsDGAPPIS0.GAPPISDEV = rsado("GAPPISDEV")
rsDGAPPIS0.GAPPISSEN = rsado("GAPPISSEN")
rsDGAPPIS0.GAPPISDEC = rsado("GAPPISDEC")
rsDGAPPIS0.GAPPISRUB = rsado("GAPPISRUB")
rsDGAPPIS0.GAPPISTPR = rsado("GAPPISTPR")
rsDGAPPIS0.GAPPISCLI = rsado("GAPPISCLI")
rsDGAPPIS0.GAPPISMON = rsado("GAPPISMON")
rsDGAPPIS0.GAPPISTTI = rsado("GAPPISTTI")
rsDGAPPIS0.GAPPISTTE = rsado("GAPPISTTE")
rsDGAPPIS0.GAPPISRTV = rsado("GAPPISRTV")
rsDGAPPIS0.GAPPISTAU = rsado("GAPPISTAU")
rsDGAPPIS0.GAPPISSOL = rsado("GAPPISSOL")
rsDGAPPIS0.GAPPISPOU = rsado("GAPPISPOU")
rsDGAPPIS0.GAPPISSIG = rsado("GAPPISSIG")
rsDGAPPIS0.GAPPISVIL = rsado("GAPPISVIL")

rsDGAPPIS0.GAPPISTMC = rsado("GAPPISTMC")
rsDGAPPIS0.GAPPISDAR = rsado("GAPPISDAR")
rsDGAPPIS0.GAPPISVAT = rsado("GAPPISVAT")
rsDGAPPIS0.GAPPISVAP = rsado("GAPPISVAP")
rsDGAPPIS0.GAPPISTVF = rsado("GAPPISTVF")
rsDGAPPIS0.GAPPISTP1 = rsado("GAPPISTP1")
rsDGAPPIS0.GAPPISTP2 = rsado("GAPPISTP2")
rsDGAPPIS0.GAPPISTM1 = rsado("GAPPISTM1")
rsDGAPPIS0.GAPPISTM2 = rsado("GAPPISTM2")
rsDGAPPIS0.GAPPISMAR = rsado("GAPPISMAR")

Exit Function

Error_Handler:

rsDGAPPIS0_GetBuffer = Error

End Function




Public Sub ddsGAPPIS0_Init()

ddsDGAPPIS0(1) = "A* DGAPPISSTA statut"
ddsDGAPPIS0(2) = "N* DGAPPISVER n° version"
ddsDGAPPIS0(3) = "N* DGAPPISPER période de traitement"
ddsDGAPPIS0(4) = "A* DGAPPISETA code établissement"
ddsDGAPPIS0(5) = "N* DGAPPISSEQ N° séquence (info)"
ddsDGAPPIS0(6) = "A* DGAPPISKAM Auto / Manuel"
ddsDGAPPIS0(7) = "N  DGAPPISCLI client"
ddsDGAPPIS0(8) = "N  DGAPPISDEC date échéance"
ddsDGAPPIS0(9) = "C  DGAPPISMTE cv EUR"
ddsDGAPPIS0(10) = "N* DGAPPISNBJ nb jours 'situation / échéance"
ddsDGAPPIS0(11) = "A* DGAPPISVEC classe de durée (01M ..... 10A 99A)"
ddsDGAPPIS0(12) = "N  GAPPISTAB  CODE ETAT"
ddsDGAPPIS0(13) = "N  GAPPISECH  NUMÉRO ÉCHÉANCIER"
ddsDGAPPIS0(14) = "N  GAPPISCLA  N° CLASSE ECHEANCIER"
ddsDGAPPIS0(15) = "N* GAPPISETA  CODE ÉTABLISSEMENT"
ddsDGAPPIS0(16) = "N* GAPPISAGE  CODE AGENCE"
ddsDGAPPIS0(17) = "A* GAPPISSER  CODE SERVICE"
ddsDGAPPIS0(18) = "A* GAPPISSSE  CODE SOUS-SERVICE"
ddsDGAPPIS0(19) = "A* GAPPISOPE  CODE OPÉRATION"
ddsDGAPPIS0(20) = "A* GAPPISNAT  CODE NATURE"
ddsDGAPPIS0(21) = "N* GAPPISNUO  NUMÉRO OPÉRATION"
ddsDGAPPIS0(22) = "A  GAPPISDEV  DEVISE"
ddsDGAPPIS0(23) = "A  GAPPISSEN  SENS"
ddsDGAPPIS0(24) = "N* GAPPISDEC  DATE ÉCHÉANCE"
ddsDGAPPIS0(25) = "A  GAPPISRUB  RUBRIQUE COMPTABLE"
ddsDGAPPIS0(26) = "A  GAPPISTPR  TYPE PRODUIT"
ddsDGAPPIS0(27) = "A* GAPPISCLI  NUMÉRO CLIENT"
ddsDGAPPIS0(28) = "C  GAPPISMON  MONTANT DU FLUX"
ddsDGAPPIS0(29) = "A  GAPPISTTI  TYPE DE TAUX INTERNE"
ddsDGAPPIS0(30) = "A  GAPPISTTE  TYPE DE TAUX EXTERNE"
ddsDGAPPIS0(31) = "A  GAPPISRTV  CODE TAUX"
ddsDGAPPIS0(32) = "D  GAPPISTAU  VALEUR DU TAUX"
ddsDGAPPIS0(33) = "C  GAPPISSOL  SOLDE RUBRI COMPTABL"
ddsDGAPPIS0(34) = "D  GAPPISPOU  POURCENTAGE"
ddsDGAPPIS0(35) = "A  GAPPISSIG  SIGLE DU CLIENT"
ddsDGAPPIS0(36) = "A  GAPPISVIL  VILLE"
ddsDGAPPIS0(37) = "D  GAPPISTMC  marge client"
ddsDGAPPIS0(38) = "N  GAPPISDAR  DATE ÉCHÉANCE REVISE"
ddsDGAPPIS0(39) = "C  GAPPISVAT  VALEUR ACTUELLE"
ddsDGAPPIS0(40) = "C  GAPPISVAP  VALEUR ACT. PONDEREE"
ddsDGAPPIS0(41) = "A  GAPPISTVF  TYPE DE TAUX BIA"
ddsDGAPPIS0(42) = "C  GAPPISTP1  variation Taux +1%"
ddsDGAPPIS0(43) = "C  GAPPISTP2  variation Taux +2%"
ddsDGAPPIS0(44) = "C  GAPPISTM1  variation Taux -1%"
ddsDGAPPIS0(45) = "C  GAPPISTM2  variation Taux -2%"
ddsDGAPPIS0(46) = "C  GAPPISMAR  marge"

End Sub
