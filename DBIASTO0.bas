Attribute VB_Name = "srvDBIASTO0"
'---------------------------------------------------------
Option Explicit
Type typeDBIASTO0

    DBIASTOSTA As String * 1  ' statut
    DBIASTOVER As Integer     ' n° version
    DBIASTOPER As String * 8  ' période de traitement
    DBIASTOETA As String * 2  ' code établissement
    DBIASTOSEQ As Long        ' N° séquence (info)
    DBIASTOKAM As String * 3  ' Auto / Manuel
    DBIASTOCLI As Long        ' client
    DBIASTOMTE As Currency    ' cv EUR
    DBIASTOAUT As String * 20 ' autorisation
    DBIASTOAU0 As String * 20 ' autorisation chapeau
    DBIASTOCPT As String * 20 ' compte

    YSTOETA As String * 2  ' CODE ÉTABLISSEMENT
    YSTOAGE As String * 2  ' CODE AGENCE
    YSTOSER As String * 2  ' CODE SERVICE
    YSTOSSE As String * 2  ' CODE SOUS-SERVICE
    YSTOOPE As String * 3  ' CODE OPÉRATION
    YSTONUM As Long        ' NUMÉRO OPÉRATION
    YSTOSEQ As Long        ' n° séquence
    YSTOPCI As String * 10 ' RUBRIQUE COMPTABLE
    YSTOCCL As String * 1  ' client / tiers
    YSTOCLI As Long        ' NUMÉRO CLIENT
    YSTODEV As String * 3  ' DEVISE
    YSTOMON As Currency    ' MONTANT disponible
    YSTODEB As Long        ' DATE début
    YSTOFIN As Long        ' date fin
    YSTOAPP As String * 3  ' application origine
    YSTONAT As String * 6  ' CODE NATURE
    YSTOCC1 As String * 1  ' client / tiers
    YSTOCL1 As Long        ' NUMÉRO CLIENT 1
    YSTOCC2 As String * 1  ' client / tiers
    YSTOCL2 As Long        ' NUMÉRO CLIENT 2
    YSTOCTX As String * 6  ' CODE TAUX
    YSTOTAU As Double      ' taux / marge
    
End Type

Public ddsDBIASTO0(33) As String * 50

Public Function sqlDBIASTO0_Delete(oldY As typeDBIASTO0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDBIASTO0_Delete = Null


xWhere = " where DBIASTOSTA = '" & oldY.DBIASTOSTA & "'" _
       & " and   DBIASTOVER = " & oldY.DBIASTOVER _
       & " and   DBIASTOPER = '" & oldY.DBIASTOPER & "'" _
       & " and   DBIASTOETA = '" & oldY.DBIASTOETA & "'" _
       & " and   DBIASTOSEQ = " & oldY.DBIASTOSEQ _


' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DBIASTO0" & xWhere

Set rsSab = cnsab.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDBIASTO0_Delete = "Erreur SUP : " & oldY.DBIASTOSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDBIASTO0_Delete = Error
End Function


Public Function sqlDBIASTO0_Insert(newY As typeDBIASTO0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDBIASTO0_Insert = Null

xSet = " (DBIASTOSEQ"
xValues = " values('" & Replace(Trim(newY.DBIASTOSEQ), "'", "''") & "'"

' Insertion :
'===================================================================================
If newY.DBIASTOVER <> 0 Then xSet = xSet & ",DBIASTOVER": xValues = xValues & ", " & newY.DBIASTOVER
If newY.DBIASTOPER <> 0 Then xSet = xSet & ",DBIASTOPER": xValues = xValues & ", " & newY.DBIASTOPER
If newY.DBIASTOCLI <> 0 Then xSet = xSet & ",DBIASTOCLI": xValues = xValues & ", " & newY.DBIASTOCLI
If newY.DBIASTOMTE <> 0 Then xSet = xSet & ",DBIASTOMTE": xValues = xValues & ", " & cur_P(newY.DBIASTOMTE)
If newY.YSTONUM <> 0 Then xSet = xSet & ",YSTONUM": xValues = xValues & ", " & newY.YSTONUM
If newY.YSTOSEQ <> 0 Then xSet = xSet & ",YSTOSEQ": xValues = xValues & ", " & newY.YSTOSEQ
If newY.YSTOCLI <> 0 Then xSet = xSet & ",YSTOCLI": xValues = xValues & ", " & newY.YSTOCLI
If newY.YSTOMON <> 0 Then xSet = xSet & ",YSTOMON": xValues = xValues & ", " & cur_P(newY.YSTOMON)
If newY.YSTODEB <> 0 Then xSet = xSet & ",YSTODEB": xValues = xValues & ", " & newY.YSTODEB
If newY.YSTOFIN <> 0 Then xSet = xSet & ",YSTOFIN": xValues = xValues & ", " & newY.YSTOFIN
If newY.YSTOCL1 <> 0 Then xSet = xSet & ",YSTOCL1": xValues = xValues & ", " & newY.YSTOCL1
If newY.YSTOCL2 <> 0 Then xSet = xSet & ",YSTOCL2": xValues = xValues & ", " & newY.YSTOCL2
If newY.YSTOTAU <> 0 Then xSet = xSet & ",YSTOTAU": xValues = xValues & ", " & Comma_Point(newY.YSTOTAU)

'===================================================================================

If Trim(newY.DBIASTOSTA) <> "" Then xSet = xSet & ",DBIASTOSTA": xValues = xValues & ", '" & Replace(Trim(newY.DBIASTOSTA), "'", "''") & "'"
If Trim(newY.DBIASTOETA) <> "" Then xSet = xSet & ",DBIASTOETA": xValues = xValues & ", '" & Replace(Trim(newY.DBIASTOETA), "'", "''") & "'"
If Trim(newY.DBIASTOKAM) <> "" Then xSet = xSet & ",DBIASTOKAM": xValues = xValues & ", '" & Replace(Trim(newY.DBIASTOKAM), "'", "''") & "'"
If Trim(newY.DBIASTOAUT) <> "" Then xSet = xSet & ",DBIASTOAUT": xValues = xValues & ", '" & Replace(Trim(newY.DBIASTOAUT), "'", "''") & "'"
If Trim(newY.DBIASTOAU0) <> "" Then xSet = xSet & ",DBIASTOAU0": xValues = xValues & ", '" & Replace(Trim(newY.DBIASTOAU0), "'", "''") & "'"
If Trim(newY.DBIASTOCPT) <> "" Then xSet = xSet & ",DBIASTOCPT": xValues = xValues & ", '" & Replace(Trim(newY.DBIASTOCPT), "'", "''") & "'"

If Trim(newY.YSTOETA) <> "" Then xSet = xSet & ",YSTOETA": xValues = xValues & ", '" & Replace(Trim(newY.YSTOETA), "'", "''") & "'"
If Trim(newY.YSTOAGE) <> "" Then xSet = xSet & ",YSTOAGE": xValues = xValues & ", '" & Replace(Trim(newY.YSTOAGE), "'", "''") & "'"
If Trim(newY.YSTOSER) <> "" Then xSet = xSet & ",YSTOSER": xValues = xValues & ", '" & Replace(Trim(newY.YSTOSER), "'", "''") & "'"
If Trim(newY.YSTOSSE) <> "" Then xSet = xSet & ",YSTOSSE": xValues = xValues & ", '" & Replace(Trim(newY.YSTOSSE), "'", "''") & "'"
If Trim(newY.YSTOOPE) <> "" Then xSet = xSet & ",YSTOOPE": xValues = xValues & ", '" & Replace(Trim(newY.YSTOOPE), "'", "''") & "'"
If Trim(newY.YSTOPCI) <> "" Then xSet = xSet & ",YSTOPCI": xValues = xValues & ", '" & Replace(Trim(newY.YSTOPCI), "'", "''") & "'"
If Trim(newY.YSTOCCL) <> "" Then xSet = xSet & ",YSTOCCL": xValues = xValues & ", '" & Replace(Trim(newY.YSTOCCL), "'", "''") & "'"
If Trim(newY.YSTODEV) <> "" Then xSet = xSet & ",YSTODEV": xValues = xValues & ", '" & Replace(Trim(newY.YSTODEV), "'", "''") & "'"
If Trim(newY.YSTOAPP) <> "" Then xSet = xSet & ",YSTOAPP": xValues = xValues & ", '" & Replace(Trim(newY.YSTOAPP), "'", "''") & "'"
If Trim(newY.YSTONAT) <> "" Then xSet = xSet & ",YSTONAT": xValues = xValues & ", '" & Replace(Trim(newY.YSTONAT), "'", "''") & "'"
If Trim(newY.YSTOCC1) <> "" Then xSet = xSet & ",YSTOCC1": xValues = xValues & ", '" & Replace(Trim(newY.YSTOCC1), "'", "''") & "'"
If Trim(newY.YSTOCC2) <> "" Then xSet = xSet & ",YSTOCC2": xValues = xValues & ", '" & Replace(Trim(newY.YSTOCC2), "'", "''") & "'"
If Trim(newY.YSTOCTX) <> "" Then xSet = xSet & ",YSTOCTX": xValues = xValues & ", '" & Replace(Trim(newY.YSTOCTX), "'", "''") & "'"

xSql = "Insert into " & paramIBM_Library_BODWH & ".DBIASTO0" & xSet & ")" & xValues & ")"

Set rsSab = cnsab.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDBIASTO0_Insert = "Erreur màj : " & newY.DBIASTOSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDBIASTO0_Insert = Error
End Function
Public Function sqlDBIASTO0_Read(oldY As typeDBIASTO0)
Dim X As String, xSql As String, Nb As Long
Dim V

On Error GoTo Error_Handler
sqlDBIASTO0_Read = Null

xSql = "select * from " & paramIBM_Library_BODWH & ".DBIASTO0 " _
       & " where DBIASTOSTA = '" & oldY.DBIASTOSTA & "'" _
       & " and   DBIASTOVER = " & oldY.DBIASTOVER _
       & " and   DBIASTOPER = " & oldY.DBIASTOPER _
       & " and   DBIASTOETA = '" & oldY.DBIASTOETA & "'" _
       & " and   DBIASTOSEQ = " & oldY.DBIASTOSEQ

Set rsSab = cnsab.Execute(xSql)

If rsSab.EOF Then
    sqlDBIASTO0_Read = "? inconnu"
Else
    V = rsDBIASTO0_GetBuffer(rsSab, oldY)
    If Not IsNull(V) Then sqlDBIASTO0_Read = "? srvDBIASTO0_GetBuffer"
End If
 
Exit Function
Error_Handler:
    sqlDBIASTO0_Read = Error
End Function


Public Function sqlDBIASTO0_Update(newY As typeDBIASTO0, oldY As typeDBIASTO0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
On Error GoTo Error_Handler
sqlDBIASTO0_Update = Null
blnUpdate = False

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DBIASTOSEQ <> newY.DBIASTOSEQ Then
    sqlDBIASTO0_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================
xWhere = " where DBIASTOSTA = '" & oldY.DBIASTOSTA & "'" _
       & " and   DBIASTOVER = " & oldY.DBIASTOVER _
       & " and   DBIASTOPER = " & oldY.DBIASTOPER _
       & " and   DBIASTOETA = '" & oldY.DBIASTOETA & "'" _
       & " and   DBIASTOSEQ = " & oldY.DBIASTOSEQ


xSet = " set"

' Détecter les modifications
'===================================================================================
If newY.DBIASTOCLI <> oldY.DBIASTOCLI Then blnUpdate = True:  xSet = xSet & " , DBIASTOCLI = " & newY.DBIASTOCLI
If newY.DBIASTOMTE <> oldY.DBIASTOMTE Then blnUpdate = True:  xSet = xSet & " , DBIASTOmte = " & cur_P(newY.DBIASTOMTE)

If newY.YSTONUM <> oldY.YSTONUM Then blnUpdate = True:  xSet = xSet & " , YSTONUM = " & newY.YSTONUM
If newY.YSTOSEQ <> oldY.YSTOSEQ Then blnUpdate = True:  xSet = xSet & " , YSTOSEQ= " & newY.YSTOSEQ
If newY.YSTOMON <> oldY.YSTOMON Then blnUpdate = True:  xSet = xSet & " , YSTOMON = " & cur_P(newY.YSTOMON)
If newY.YSTOCLI <> oldY.YSTOCL1 Then blnUpdate = True:  xSet = xSet & " , YSTOCLI = " & newY.YSTOCLI
If newY.YSTODEB <> oldY.YSTODEB Then blnUpdate = True:  xSet = xSet & " , YSTODEB = " & newY.YSTODEB
If newY.YSTOFIN <> oldY.YSTOFIN Then blnUpdate = True:  xSet = xSet & " , YSTOFIN = " & newY.YSTOFIN
If newY.YSTOCL1 <> oldY.YSTOCL1 Then blnUpdate = True:  xSet = xSet & " , YSTOCL1 = " & newY.YSTOCL1
If newY.YSTOCL2 <> oldY.YSTOCL2 Then blnUpdate = True:  xSet = xSet & " , YSTOCL2 = " & newY.YSTOCL2
If newY.YSTOTAU <> oldY.YSTOTAU Then blnUpdate = True:  xSet = xSet & " , YSTOTAU = " & Comma_Point(newY.YSTOTAU)

If Trim(newY.DBIASTOAUT) <> Trim(oldY.DBIASTOAUT) Then blnUpdate = True: xSet = xSet & ",DBIASTOAUT = '" & Replace(Trim(newY.DBIASTOAUT), "'", "''") & "'"
If Trim(newY.DBIASTOAU0) <> Trim(oldY.DBIASTOAU0) Then blnUpdate = True: xSet = xSet & ",DBIASTOAU0 = '" & Replace(Trim(newY.DBIASTOAU0), "'", "''") & "'"
If Trim(newY.DBIASTOCPT) <> Trim(oldY.DBIASTOCPT) Then blnUpdate = True: xSet = xSet & ",DBIASTOCPT = '" & Replace(Trim(newY.DBIASTOCPT), "'", "''") & "'"
If Trim(newY.YSTOETA) <> Trim(oldY.YSTOETA) Then blnUpdate = True: xSet = xSet & ",YSTOETA = '" & Replace(Trim(newY.YSTOETA), "'", "''") & "'"
If Trim(newY.YSTOAGE) <> Trim(oldY.YSTOAGE) Then blnUpdate = True: xSet = xSet & ",YSTOAGE = '" & Replace(Trim(newY.YSTOAGE), "'", "''") & "'"
If Trim(newY.YSTOSER) <> Trim(oldY.YSTOSER) Then blnUpdate = True: xSet = xSet & ",YSTOSER = '" & Replace(Trim(newY.YSTOSER), "'", "''") & "'"
If Trim(newY.YSTOSSE) <> Trim(oldY.YSTOSSE) Then blnUpdate = True: xSet = xSet & ",YSTOSSE = '" & Replace(Trim(newY.YSTOSSE), "'", "''") & "'"
If Trim(newY.YSTOOPE) <> Trim(oldY.YSTOOPE) Then blnUpdate = True: xSet = xSet & ",YSTOOPE = '" & Replace(Trim(newY.YSTOOPE), "'", "''") & "'"
If Trim(newY.YSTOPCI) <> Trim(oldY.YSTOPCI) Then blnUpdate = True: xSet = xSet & ",YSTOPCI = '" & Replace(Trim(newY.YSTOPCI), "'", "''") & "'"
If Trim(newY.YSTOCCL) <> Trim(oldY.YSTOCCL) Then blnUpdate = True: xSet = xSet & ",YSTOCCL = '" & Replace(Trim(newY.YSTOCCL), "'", "''") & "'"
If Trim(newY.YSTODEV) <> Trim(oldY.YSTODEV) Then blnUpdate = True: xSet = xSet & ",YSTODEV = '" & Replace(Trim(newY.YSTODEV), "'", "''") & "'"
If Trim(newY.YSTOAPP) <> Trim(oldY.YSTOAPP) Then blnUpdate = True: xSet = xSet & ",YSTOAPP = '" & Replace(Trim(newY.YSTOAPP), "'", "''") & "'"
If Trim(newY.YSTONAT) <> Trim(oldY.YSTONAT) Then blnUpdate = True: xSet = xSet & ",YSTONAT = '" & Replace(Trim(newY.YSTONAT), "'", "''") & "'"
If Trim(newY.YSTOCC1) <> Trim(oldY.YSTOCC1) Then blnUpdate = True: xSet = xSet & ",YSTOCC1 = '" & Replace(Trim(newY.YSTOCC1), "'", "''") & "'"
If Trim(newY.YSTOCC2) <> Trim(oldY.YSTOCC2) Then blnUpdate = True: xSet = xSet & ",YSTOCC2 = '" & Replace(Trim(newY.YSTOCC2), "'", "''") & "'"
If Trim(newY.YSTOCTX) <> Trim(oldY.YSTOCTX) Then blnUpdate = True: xSet = xSet & ",YSTOCTX = '" & Replace(Trim(newY.YSTOCTX), "'", "''") & "'"

If blnUpdate Then
    If Trim(newY.DBIASTOKAM) = "" Then newY.DBIASTOKAM = "M"
    If newY.DBIASTOKAM <> oldY.DBIASTOKAM Then blnUpdate = True: xSet = xSet & ",DBIASTOKAM = '" & Replace(Trim(newY.DBIASTOKAM), "'", "''") & "'"

    Mid$(xSet, InStr(xSet, ","), 1) = " "
    xSql = "update " & paramIBM_Library_BODWH & ".DBIASTO0" & xSet & xWhere
    
    Set rsSab = cnsab.Execute(xSql, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlDBIASTO0_Update = "Erreur màj : " & newY.DBIASTOSEQ
    
        Exit Function
    End If
End If
Exit Function
Error_Handler:
    sqlDBIASTO0_Update = Error
End Function




'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsDBIASTO0_GetBuffer(rsADO As ADODB.Recordset, rsDBIASTO0 As typeDBIASTO0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsDBIASTO0_GetBuffer = Null

rsDBIASTO0.DBIASTOSTA = rsADO("DBIASTOSTA")
rsDBIASTO0.DBIASTOVER = rsADO("DBIASTOVER")
rsDBIASTO0.DBIASTOPER = rsADO("DBIASTOPER")
rsDBIASTO0.DBIASTOETA = rsADO("DBIASTOETA")
rsDBIASTO0.DBIASTOSEQ = rsADO("DBIASTOSEQ")
rsDBIASTO0.DBIASTOKAM = rsADO("DBIASTOKAM")
rsDBIASTO0.DBIASTOCLI = rsADO("DBIASTOCLI")
rsDBIASTO0.DBIASTOMTE = rsADO("DBIASTOMTE")
rsDBIASTO0.DBIASTOAUT = rsADO("DBIASTOAUT")
rsDBIASTO0.DBIASTOAU0 = rsADO("DBIASTOAU0")
rsDBIASTO0.DBIASTOCPT = rsADO("DBIASTOCPT")


rsDBIASTO0.YSTOETA = rsADO("YSTOETA")
rsDBIASTO0.YSTOAGE = rsADO("YSTOAGE")
rsDBIASTO0.YSTOSER = rsADO("YSTOSER")
rsDBIASTO0.YSTOSSE = rsADO("YSTOSSE")
rsDBIASTO0.YSTOOPE = rsADO("YSTOOPE")
rsDBIASTO0.YSTONUM = rsADO("YSTONUM")
rsDBIASTO0.YSTOSEQ = rsADO("YSTOSEQ")
rsDBIASTO0.YSTOPCI = rsADO("YSTOPCI")
rsDBIASTO0.YSTOCCL = rsADO("YSTOCCL")
rsDBIASTO0.YSTOCLI = rsADO("YSTOCLI")
rsDBIASTO0.YSTODEV = rsADO("YSTODEV")
rsDBIASTO0.YSTOMON = rsADO("YSTOMON")
rsDBIASTO0.YSTODEB = rsADO("YSTODEB")
rsDBIASTO0.YSTOFIN = rsADO("YSTOFIN")
rsDBIASTO0.YSTOAPP = rsADO("YSTOAPP")
rsDBIASTO0.YSTONAT = rsADO("YSTONAT")
rsDBIASTO0.YSTOCC1 = rsADO("YSTOCC1")
rsDBIASTO0.YSTOCL1 = rsADO("YSTOCL1")
rsDBIASTO0.YSTOCC2 = rsADO("YSTOCC2")
rsDBIASTO0.YSTOCL2 = rsADO("YSTOCL2")
rsDBIASTO0.YSTOCTX = rsADO("YSTOCTX")
rsDBIASTO0.YSTOTAU = rsADO("YSTOTAU")

Exit Function

Error_Handler:

rsDBIASTO0_GetBuffer = Error

End Function




Public Sub ddsYSTO0_Init()

ddsDBIASTO0(1) = "A* DBIASTOSTA statut"
ddsDBIASTO0(2) = "N* DBIASTOVER n° version"
ddsDBIASTO0(3) = "N* DBIASTOPER période de traitement"
ddsDBIASTO0(4) = "A* DBIASTOETA code établissement"
ddsDBIASTO0(5) = "N* DBIASTOSEQ N° séquence (info)"
ddsDBIASTO0(6) = "A* DBIASTOKAM Auto / Manuel"
ddsDBIASTO0(7) = "N  DBIASTOCLI client"
ddsDBIASTO0(8) = "C  DBIASTOMTE cv EUR"
ddsDBIASTO0(9) = "A  DBIASTOAUT autorisation"
ddsDBIASTO0(10) = "A  DBIASTOAU0 autorisation plafond"
ddsDBIASTO0(11) = "A  DBIASTOCPT compte"


ddsDBIASTO0(12) = "A* YSTOETA     Code établissement"
ddsDBIASTO0(13) = "A* YSTOAGE     Code agence"
ddsDBIASTO0(14) = "A* YSTOSER     Code service"
ddsDBIASTO0(15) = "A* YSTOSSE     Code sous-service"
ddsDBIASTO0(16) = "A* YSTOOPE     Code opération"
ddsDBIASTO0(17) = "N* YSTONUM     N°opération"
ddsDBIASTO0(18) = "N* YSTOSEQ     N° séquence (prêt)"
ddsDBIASTO0(19) = "A  YSTOPCI     Rubrique comptable"
ddsDBIASTO0(20) = "A* YSTOCCL     client / tiers"
ddsDBIASTO0(21) = "A* YSTOCLI     N° client"
ddsDBIASTO0(22) = "A  YSTODEV     Devise"
ddsDBIASTO0(23) = "C  YSTOMON     Montant"
ddsDBIASTO0(24) = "N  YSTODEB     Date début"
ddsDBIASTO0(25) = "N  YSTOFIN     Date fin"
ddsDBIASTO0(26) = "A  YSTOAPP     application origine"
ddsDBIASTO0(27) = "A  YSTONAT     Code nature"
ddsDBIASTO0(28) = "A  YSTOCC1     client / tiers 1"
ddsDBIASTO0(29) = "N  YSTOCL1     N° client 1"
ddsDBIASTO0(30) = "A  YSTOCC2     client / tiers 2"
ddsDBIASTO0(31) = "N  YSTOCL2     N° client 2"
ddsDBIASTO0(32) = "A  YSTOCTX     Code taux"
ddsDBIASTO0(33) = "D  YSTOTAU     Taux / Marge"




End Sub


