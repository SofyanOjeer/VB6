Attribute VB_Name = "srvDCREINT0"
'---------------------------------------------------------
Option Explicit
Type typeDCREINT0

    CREINTSTA As String * 1  ' statut
    CREINTVER As Integer     ' n° version
    CREINTPER As String * 8  ' période de traitement
    
    CREINTETA As String * 2  ' code établissement
    CREINTAGE As String * 2  ' agence
    CREINTSER As String * 2  ' CODE SERVICE
    CREINTSSE As String * 2  ' CODE SOUS-SERVICE
    CREINTDOS As Long        ' NUMÉRO dossier
    CREINTPRE As Integer     ' no prêt
    CREINTNAT As String * 6  ' CODE NATURE
    CREINTNAP As String * 6  '
    CREINTCLI As Long        ' NUMÉRO CLIENT
    CREINTMT0 As Currency    ' capital
    CREINTDEV As String * 3  ' DEVISE
    
    CREINTECH As Long        ' DATE ÉCHÉANCE
    CREINTMTX As Currency    ' capital restant dû
    CREINTTOF As Double      ' taux
    CREINTTOM As Double      ' marge
    CREINTPERK As String * 1 ' périodicité
    CREINTPERN As Integer    ' nb période
    CREINTUAMJ As Long       ' DATE maj TA
    CREINTUHMS As Long       ' heure maj TA
    
    CREINTT01 As Currency    ' intérêts Trim 01
    CREINTM01 As Currency    ' marge Trim 01
    CREINTT02 As Currency    ' intérêts Trim 02
    CREINTM02 As Currency    ' marge Trim 02
    CREINTT03 As Currency    ' intérêts Trim 03
    CREINTM03 As Currency    ' marge Trim 03
    CREINTT04 As Currency    ' intérêts Trim 04
    CREINTM04 As Currency    ' marge Trim 04
   
    
    CREINTT11 As Currency    ' intérêts Trim 11
    CREINTM11 As Currency    ' marge Trim 11
    CREINTT12 As Currency    ' intérêts Trim 12
    CREINTM12 As Currency    ' marge Trim 12
    CREINTT13 As Currency    ' intérêts Trim 13
    CREINTM13 As Currency    ' marge Trim 13
    CREINTT14 As Currency    ' intérêts Trim 14
    CREINTM14 As Currency    ' marge Trim 14
    
    CREINTT21 As Currency    ' intérêts Trim 21
    CREINTM21 As Currency    ' marge Trim 21
    CREINTT22 As Currency    ' intérêts Trim 22
    CREINTM22 As Currency    ' marge Trim 22
    CREINTT23 As Currency    ' intérêts Trim 23
    CREINTM23 As Currency    ' marge Trim 23
    CREINTT24 As Currency    ' intérêts Trim 24
    CREINTM24 As Currency    ' marge Trim 24
    
    CREINTT31 As Currency    ' intérêts Trim 31
    CREINTM31 As Currency    ' marge Trim 31
    CREINTT32 As Currency    ' intérêts Trim 32
    CREINTM32 As Currency    ' marge Trim 32
    CREINTT33 As Currency    ' intérêts Trim 33
    CREINTM33 As Currency    ' marge Trim 33
    CREINTT34 As Currency    ' intérêts Trim 34
    CREINTM34 As Currency    ' marge Trim 34
    
    CREINTT41 As Currency    ' intérêts Trim 41
    CREINTM41 As Currency    ' marge Trim 41
    CREINTT42 As Currency    ' intérêts Trim 42
    CREINTM42 As Currency    ' marge Trim 42
    CREINTT43 As Currency    ' intérêts Trim 43
    CREINTM43 As Currency    ' marge Trim 43
    CREINTT44 As Currency    ' intérêts Trim 44
    CREINTM44 As Currency    ' marge Trim 44

    
End Type

Public ddsDCREINT0(62) As String * 50

Type typeDCRETA

    CRETADEB      As String
    CRETAFIN      As String
    CRETAMIN     As Currency
    CRETAMARGE   As Currency
    CRETATAU    As Double
End Type

Public Function sqlDCREINT0_Delete(oldY As typeDCREINT0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDCREINT0_Delete = Null


xWhere = " where CREINTSTA = '" & oldY.CREINTSTA & "'" _
       & " and   CREINTVER = " & oldY.CREINTVER _
       & " and   CREINTPER = '" & oldY.CREINTPER & "'" _
       & " and   CREINTETA = '" & oldY.CREINTETA & "'" _
       & " and   CREINTAGE = '" & oldY.CREINTAGE & "'" _
       & " and   CREINTSER = '" & oldY.CREINTSER & "'" _
       & " and   CREINTSSE = '" & oldY.CREINTSSE & "'" _
       & " and   CREINTDOS = " & oldY.CREINTDOS _
       & " and   CREINTPRE = " & oldY.CREINTPRE


' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DCREINT0" & xWhere

Set rsSab = cnsab.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCREINT0_Delete = "Erreur SUP : " & oldY.CREINTDOS
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCREINT0_Delete = Error
End Function


Public Function sqlDCREINT0_Insert(newY As typeDCREINT0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDCREINT0_Insert = Null

xSet = " (CREINTSTA"
xValues = " values('" & newY.CREINTSTA & "'"

' Insertion :
'===================================================================================
'If Trim(newY.CREINTSTA) <> "" Then xSet = xSet & ",CREINTSTA": xValues = xValues & ", '" & Replace(Trim(newY.CREINTSTA), "'", "''") & "'"
If newY.CREINTVER <> 0 Then xSet = xSet & ",CREINTVER": xValues = xValues & ", " & newY.CREINTVER
If newY.CREINTPER <> 0 Then xSet = xSet & ",CREINTPER": xValues = xValues & ", " & newY.CREINTPER

If Trim(newY.CREINTETA) <> "" Then xSet = xSet & ",CREINTETA": xValues = xValues & ", '" & Replace(Trim(newY.CREINTETA), "'", "''") & "'"
If Trim(newY.CREINTAGE) <> "" Then xSet = xSet & ",CREINTAGE": xValues = xValues & ", '" & Replace(Trim(newY.CREINTAGE), "'", "''") & "'"
If Trim(newY.CREINTSER) <> "" Then xSet = xSet & ",CREINTSER": xValues = xValues & ", '" & Replace(Trim(newY.CREINTSER), "'", "''") & "'"
If Trim(newY.CREINTSSE) <> "" Then xSet = xSet & ",CREINTSSE": xValues = xValues & ", '" & Replace(Trim(newY.CREINTSSE), "'", "''") & "'"
If newY.CREINTDOS <> 0 Then xSet = xSet & ",CREINTDOS": xValues = xValues & ", " & newY.CREINTDOS
If newY.CREINTPRE <> 0 Then xSet = xSet & ",CREINTPRE": xValues = xValues & ", " & newY.CREINTPRE
If Trim(newY.CREINTNAT) <> "" Then xSet = xSet & ",CREINTNAT": xValues = xValues & ", '" & Replace(Trim(newY.CREINTNAT), "'", "''") & "'"
If Trim(newY.CREINTNAP) <> "" Then xSet = xSet & ",CREINTNAP": xValues = xValues & ", '" & Replace(Trim(newY.CREINTNAP), "'", "''") & "'"
If newY.CREINTCLI <> 0 Then xSet = xSet & ",CREINTCLI": xValues = xValues & ", " & newY.CREINTCLI
If newY.CREINTMT0 <> 0 Then xSet = xSet & ",CREINTMT0": xValues = xValues & ", " & cur_P(newY.CREINTMT0)
If Trim(newY.CREINTDEV) <> "" Then xSet = xSet & ",CREINTDEV": xValues = xValues & ", '" & Replace(Trim(newY.CREINTDEV), "'", "''") & "'"

If newY.CREINTECH <> 0 Then xSet = xSet & ",CREINTECH": xValues = xValues & ", " & newY.CREINTECH
If newY.CREINTMTX <> 0 Then xSet = xSet & ",CREINTMTX": xValues = xValues & ", " & cur_P(newY.CREINTMTX)
If newY.CREINTTOF <> 0 Then xSet = xSet & ",CREINTTOF": xValues = xValues & ", " & Comma_Point(newY.CREINTTOF)
If newY.CREINTTOM <> 0 Then xSet = xSet & ",CREINTTOM": xValues = xValues & ", " & Comma_Point(newY.CREINTTOM)
If Trim(newY.CREINTPERK) <> "" Then xSet = xSet & ",CREINTPERK": xValues = xValues & ", '" & Replace(Trim(newY.CREINTPERK), "'", "''") & "'"
If newY.CREINTPERN <> 0 Then xSet = xSet & ",CREINTPERN": xValues = xValues & ", " & newY.CREINTPERN
If newY.CREINTUAMJ <> 0 Then xSet = xSet & ",CREINTUAMJ": xValues = xValues & ", " & newY.CREINTUAMJ
If newY.CREINTUHMS <> 0 Then xSet = xSet & ",CREINTUHMS": xValues = xValues & ", " & newY.CREINTUHMS

If newY.CREINTT01 <> 0 Then xSet = xSet & ",CREINTT01": xValues = xValues & ", " & cur_P(newY.CREINTT01)
If newY.CREINTM01 <> 0 Then xSet = xSet & ",CREINTM01": xValues = xValues & ", " & cur_P(newY.CREINTM01)
If newY.CREINTT02 <> 0 Then xSet = xSet & ",CREINTT02": xValues = xValues & ", " & cur_P(newY.CREINTT02)
If newY.CREINTM02 <> 0 Then xSet = xSet & ",CREINTM02": xValues = xValues & ", " & cur_P(newY.CREINTM02)
If newY.CREINTT03 <> 0 Then xSet = xSet & ",CREINTT03": xValues = xValues & ", " & cur_P(newY.CREINTT03)
If newY.CREINTM03 <> 0 Then xSet = xSet & ",CREINTM03": xValues = xValues & ", " & cur_P(newY.CREINTM03)
If newY.CREINTT04 <> 0 Then xSet = xSet & ",CREINTT04": xValues = xValues & ", " & cur_P(newY.CREINTT04)
If newY.CREINTM04 <> 0 Then xSet = xSet & ",CREINTM04": xValues = xValues & ", " & cur_P(newY.CREINTM04)

If newY.CREINTT11 <> 0 Then xSet = xSet & ",CREINTT11": xValues = xValues & ", " & cur_P(newY.CREINTT11)
If newY.CREINTM11 <> 0 Then xSet = xSet & ",CREINTM11": xValues = xValues & ", " & cur_P(newY.CREINTM11)
If newY.CREINTT12 <> 0 Then xSet = xSet & ",CREINTT12": xValues = xValues & ", " & cur_P(newY.CREINTT12)
If newY.CREINTM12 <> 0 Then xSet = xSet & ",CREINTM12": xValues = xValues & ", " & cur_P(newY.CREINTM12)
If newY.CREINTT13 <> 0 Then xSet = xSet & ",CREINTT13": xValues = xValues & ", " & cur_P(newY.CREINTT13)
If newY.CREINTM13 <> 0 Then xSet = xSet & ",CREINTM13": xValues = xValues & ", " & cur_P(newY.CREINTM13)
If newY.CREINTT14 <> 0 Then xSet = xSet & ",CREINTT14": xValues = xValues & ", " & cur_P(newY.CREINTT14)
If newY.CREINTM14 <> 0 Then xSet = xSet & ",CREINTM14": xValues = xValues & ", " & cur_P(newY.CREINTM14)

If newY.CREINTT21 <> 0 Then xSet = xSet & ",CREINTT21": xValues = xValues & ", " & cur_P(newY.CREINTT21)
If newY.CREINTM21 <> 0 Then xSet = xSet & ",CREINTM21": xValues = xValues & ", " & cur_P(newY.CREINTM21)
If newY.CREINTT22 <> 0 Then xSet = xSet & ",CREINTT22": xValues = xValues & ", " & cur_P(newY.CREINTT22)
If newY.CREINTM22 <> 0 Then xSet = xSet & ",CREINTM22": xValues = xValues & ", " & cur_P(newY.CREINTM22)
If newY.CREINTT23 <> 0 Then xSet = xSet & ",CREINTT23": xValues = xValues & ", " & cur_P(newY.CREINTT23)
If newY.CREINTM23 <> 0 Then xSet = xSet & ",CREINTM23": xValues = xValues & ", " & cur_P(newY.CREINTM23)
If newY.CREINTT24 <> 0 Then xSet = xSet & ",CREINTT24": xValues = xValues & ", " & cur_P(newY.CREINTT24)
If newY.CREINTM24 <> 0 Then xSet = xSet & ",CREINTM24": xValues = xValues & ", " & cur_P(newY.CREINTM24)

If newY.CREINTT31 <> 0 Then xSet = xSet & ",CREINTT31": xValues = xValues & ", " & cur_P(newY.CREINTT31)
If newY.CREINTM31 <> 0 Then xSet = xSet & ",CREINTM31": xValues = xValues & ", " & cur_P(newY.CREINTM31)
If newY.CREINTT32 <> 0 Then xSet = xSet & ",CREINTT32": xValues = xValues & ", " & cur_P(newY.CREINTT32)
If newY.CREINTM32 <> 0 Then xSet = xSet & ",CREINTM32": xValues = xValues & ", " & cur_P(newY.CREINTM32)
If newY.CREINTT33 <> 0 Then xSet = xSet & ",CREINTT33": xValues = xValues & ", " & cur_P(newY.CREINTT33)
If newY.CREINTM33 <> 0 Then xSet = xSet & ",CREINTM33": xValues = xValues & ", " & cur_P(newY.CREINTM33)
If newY.CREINTT34 <> 0 Then xSet = xSet & ",CREINTT34": xValues = xValues & ", " & cur_P(newY.CREINTT34)
If newY.CREINTM34 <> 0 Then xSet = xSet & ",CREINTM34": xValues = xValues & ", " & cur_P(newY.CREINTM34)

If newY.CREINTT41 <> 0 Then xSet = xSet & ",CREINTT41": xValues = xValues & ", " & cur_P(newY.CREINTT41)
If newY.CREINTM41 <> 0 Then xSet = xSet & ",CREINTM41": xValues = xValues & ", " & cur_P(newY.CREINTM41)
If newY.CREINTT42 <> 0 Then xSet = xSet & ",CREINTT42": xValues = xValues & ", " & cur_P(newY.CREINTT42)
If newY.CREINTM42 <> 0 Then xSet = xSet & ",CREINTM42": xValues = xValues & ", " & cur_P(newY.CREINTM42)
If newY.CREINTT43 <> 0 Then xSet = xSet & ",CREINTT43": xValues = xValues & ", " & cur_P(newY.CREINTT43)
If newY.CREINTM43 <> 0 Then xSet = xSet & ",CREINTM43": xValues = xValues & ", " & cur_P(newY.CREINTM43)
If newY.CREINTT44 <> 0 Then xSet = xSet & ",CREINTT44": xValues = xValues & ", " & cur_P(newY.CREINTT44)
If newY.CREINTM44 <> 0 Then xSet = xSet & ",CREINTM44": xValues = xValues & ", " & cur_P(newY.CREINTM44)



xSql = "Insert into " & paramIBM_Library_BODWH & ".DCREINT0" & xSet & ")" & xValues & ")"

Set rsSab = cnsab.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDCREINT0_Insert = "Erreur màj : " & newY.CREINTDOS
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDCREINT0_Insert = Error
End Function
Public Function sqlDCREINT0_Read(oldY As typeDCREINT0)
Dim X As String, xSql As String, Nb As Long
Dim V

On Error GoTo Error_Handler
sqlDCREINT0_Read = Null

xSql = "select * from " & paramIBM_Library_BODWH & ".DCREINT0 " _
       & " where CREINTSTA = '" & oldY.CREINTSTA & "'" _
       & " and   CREINTVER = " & oldY.CREINTVER _
       & " and   CREINTPER = " & oldY.CREINTPER _
       & " and   CREINTETA = '" & oldY.CREINTETA & "'" _
       & " and   CREINTAGE = '" & oldY.CREINTAGE & "'" _
       & " and   CREINTSER = '" & oldY.CREINTSER & "'" _
       & " and   CREINTSSE = '" & oldY.CREINTSSE & "'" _
       & " and   CREINTDOS = " & oldY.CREINTDOS _
       & " and   CREINTPRE = " & oldY.CREINTPRE

Set rsSab = cnsab.Execute(xSql)

If rsSab.EOF Then
    sqlDCREINT0_Read = "? inconnu"
Else
    V = rsDCREINT0_GetBuffer(rsSab, oldY)
    If Not IsNull(V) Then sqlDCREINT0_Read = "? srvDCREINT0_GetBuffer"
End If
 
Exit Function
Error_Handler:
    sqlDCREINT0_Read = Error
End Function


Public Function sqlDCREINT0_Update(newY As typeDCREINT0, oldY As typeDCREINT0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
On Error GoTo Error_Handler
sqlDCREINT0_Update = Null
blnUpdate = False

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.CREINTDOS <> newY.CREINTDOS Then
    sqlDCREINT0_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================
xWhere = " where CREINTSTA = '" & oldY.CREINTSTA & "'" _
       & " and   CREINTVER = " & oldY.CREINTVER _
       & " and   CREINTPER = " & oldY.CREINTPER _
       & " and   CREINTETA = '" & oldY.CREINTETA & "'" _
       & " and   CREINTAGE = '" & oldY.CREINTAGE & "'" _
       & " and   CREINTSER = '" & oldY.CREINTSER & "'" _
       & " and   CREINTSSE = '" & oldY.CREINTSSE & "'" _
       & " and   CREINTDOS = " & oldY.CREINTDOS _
       & " and   CREINTPRE = " & oldY.CREINTPRE


xSet = " set"

' Détecter les modifications
'===================================================================================
If Trim(newY.CREINTNAT) <> Trim(oldY.CREINTNAT) Then blnUpdate = True: xSet = xSet & ",CREINTNAT = '" & Replace(Trim(newY.CREINTNAT), "'", "''") & "'"
If Trim(newY.CREINTNAP) <> Trim(oldY.CREINTNAP) Then blnUpdate = True: xSet = xSet & ",CREINTNAP = '" & Replace(Trim(newY.CREINTNAP), "'", "''") & "'"
If newY.CREINTCLI <> oldY.CREINTCLI Then blnUpdate = True:  xSet = xSet & " , CREINTCLI = " & newY.CREINTCLI
If newY.CREINTMT0 <> oldY.CREINTMT0 Then blnUpdate = True:  xSet = xSet & " , CREINTMT0 = " & cur_P(newY.CREINTMT0)
If Trim(newY.CREINTDEV) <> Trim(oldY.CREINTDEV) Then blnUpdate = True: xSet = xSet & ",CREINTDEV = '" & Replace(Trim(newY.CREINTDEV), "'", "''") & "'"
If newY.CREINTECH <> oldY.CREINTECH Then blnUpdate = True:  xSet = xSet & " , CREINTECH = " & newY.CREINTECH
If newY.CREINTMTX <> oldY.CREINTMTX Then blnUpdate = True:  xSet = xSet & " , CREINTMTX = " & cur_P(newY.CREINTMTX)
If newY.CREINTTOF <> oldY.CREINTTOF Then blnUpdate = True:  xSet = xSet & " , CREINTTOF = " & Comma_Point(newY.CREINTTOF)
If newY.CREINTTOM <> oldY.CREINTTOM Then blnUpdate = True:  xSet = xSet & " , CREINTTOM = " & Comma_Point(newY.CREINTTOM)
If Trim(newY.CREINTPERK) <> Trim(oldY.CREINTPERK) Then blnUpdate = True: xSet = xSet & ",CREINTPERK = '" & Replace(Trim(newY.CREINTPERK), "'", "''") & "'"
If newY.CREINTPERN <> oldY.CREINTPERN Then blnUpdate = True:  xSet = xSet & " , CREINTPERN = " & newY.CREINTPERN
If newY.CREINTUAMJ <> oldY.CREINTUAMJ Then blnUpdate = True:  xSet = xSet & " , CREINTUAMJ = " & newY.CREINTUAMJ
If newY.CREINTUHMS <> oldY.CREINTUHMS Then blnUpdate = True:  xSet = xSet & " , CREINTUHMS = " & newY.CREINTUHMS

If newY.CREINTT01 <> oldY.CREINTT01 Then blnUpdate = True:  xSet = xSet & " , CREINTT01 = " & cur_P(newY.CREINTT01)
If newY.CREINTM01 <> oldY.CREINTM01 Then blnUpdate = True:  xSet = xSet & " , CREINTM01 = " & cur_P(newY.CREINTM01)
If newY.CREINTT02 <> oldY.CREINTT02 Then blnUpdate = True:  xSet = xSet & " , CREINTT02 = " & cur_P(newY.CREINTT02)
If newY.CREINTM02 <> oldY.CREINTM02 Then blnUpdate = True:  xSet = xSet & " , CREINTM02 = " & cur_P(newY.CREINTM02)
If newY.CREINTT03 <> oldY.CREINTT03 Then blnUpdate = True:  xSet = xSet & " , CREINTT03 = " & cur_P(newY.CREINTT03)
If newY.CREINTM03 <> oldY.CREINTM03 Then blnUpdate = True:  xSet = xSet & " , CREINTM03 = " & cur_P(newY.CREINTM03)
If newY.CREINTT04 <> oldY.CREINTT04 Then blnUpdate = True:  xSet = xSet & " , CREINTT04 = " & cur_P(newY.CREINTT04)
If newY.CREINTM04 <> oldY.CREINTM04 Then blnUpdate = True:  xSet = xSet & " , CREINTM04 = " & cur_P(newY.CREINTM04)

If newY.CREINTT11 <> oldY.CREINTT11 Then blnUpdate = True:  xSet = xSet & " , CREINTT11 = " & cur_P(newY.CREINTT11)
If newY.CREINTM11 <> oldY.CREINTM11 Then blnUpdate = True:  xSet = xSet & " , CREINTM11 = " & cur_P(newY.CREINTM11)
If newY.CREINTT12 <> oldY.CREINTT12 Then blnUpdate = True:  xSet = xSet & " , CREINTT12 = " & cur_P(newY.CREINTT12)
If newY.CREINTM12 <> oldY.CREINTM12 Then blnUpdate = True:  xSet = xSet & " , CREINTM12 = " & cur_P(newY.CREINTM12)
If newY.CREINTT13 <> oldY.CREINTT13 Then blnUpdate = True:  xSet = xSet & " , CREINTT13 = " & cur_P(newY.CREINTT13)
If newY.CREINTM13 <> oldY.CREINTM13 Then blnUpdate = True:  xSet = xSet & " , CREINTM13 = " & cur_P(newY.CREINTM13)
If newY.CREINTT14 <> oldY.CREINTT14 Then blnUpdate = True:  xSet = xSet & " , CREINTT14 = " & cur_P(newY.CREINTT14)
If newY.CREINTM14 <> oldY.CREINTM14 Then blnUpdate = True:  xSet = xSet & " , CREINTM14 = " & cur_P(newY.CREINTM14)

If newY.CREINTT21 <> oldY.CREINTT21 Then blnUpdate = True:  xSet = xSet & " , CREINTT21 = " & cur_P(newY.CREINTT21)
If newY.CREINTM21 <> oldY.CREINTM21 Then blnUpdate = True:  xSet = xSet & " , CREINTM21 = " & cur_P(newY.CREINTM21)
If newY.CREINTT22 <> oldY.CREINTT22 Then blnUpdate = True:  xSet = xSet & " , CREINTT22 = " & cur_P(newY.CREINTT22)
If newY.CREINTM22 <> oldY.CREINTM22 Then blnUpdate = True:  xSet = xSet & " , CREINTM22 = " & cur_P(newY.CREINTM22)
If newY.CREINTT23 <> oldY.CREINTT23 Then blnUpdate = True:  xSet = xSet & " , CREINTT23 = " & cur_P(newY.CREINTT23)
If newY.CREINTM23 <> oldY.CREINTM23 Then blnUpdate = True:  xSet = xSet & " , CREINTM23 = " & cur_P(newY.CREINTM23)
If newY.CREINTT24 <> oldY.CREINTT24 Then blnUpdate = True:  xSet = xSet & " , CREINTT24 = " & cur_P(newY.CREINTT24)
If newY.CREINTM24 <> oldY.CREINTM24 Then blnUpdate = True:  xSet = xSet & " , CREINTM24 = " & cur_P(newY.CREINTM24)

If newY.CREINTT31 <> oldY.CREINTT31 Then blnUpdate = True:  xSet = xSet & " , CREINTT31 = " & cur_P(newY.CREINTT31)
If newY.CREINTM31 <> oldY.CREINTM31 Then blnUpdate = True:  xSet = xSet & " , CREINTM31 = " & cur_P(newY.CREINTM31)
If newY.CREINTT32 <> oldY.CREINTT32 Then blnUpdate = True:  xSet = xSet & " , CREINTT32 = " & cur_P(newY.CREINTT32)
If newY.CREINTM32 <> oldY.CREINTM32 Then blnUpdate = True:  xSet = xSet & " , CREINTM32 = " & cur_P(newY.CREINTM32)
If newY.CREINTT33 <> oldY.CREINTT33 Then blnUpdate = True:  xSet = xSet & " , CREINTT33 = " & cur_P(newY.CREINTT33)
If newY.CREINTM33 <> oldY.CREINTM33 Then blnUpdate = True:  xSet = xSet & " , CREINTM33 = " & cur_P(newY.CREINTM33)
If newY.CREINTT34 <> oldY.CREINTT34 Then blnUpdate = True:  xSet = xSet & " , CREINTT34 = " & cur_P(newY.CREINTT34)
If newY.CREINTM34 <> oldY.CREINTM34 Then blnUpdate = True:  xSet = xSet & " , CREINTM34 = " & cur_P(newY.CREINTM34)

If newY.CREINTT41 <> oldY.CREINTT41 Then blnUpdate = True:  xSet = xSet & " , CREINTT41 = " & cur_P(newY.CREINTT41)
If newY.CREINTM41 <> oldY.CREINTM41 Then blnUpdate = True:  xSet = xSet & " , CREINTM41 = " & cur_P(newY.CREINTM41)
If newY.CREINTT42 <> oldY.CREINTT42 Then blnUpdate = True:  xSet = xSet & " , CREINTT42 = " & cur_P(newY.CREINTT42)
If newY.CREINTM42 <> oldY.CREINTM42 Then blnUpdate = True:  xSet = xSet & " , CREINTM42 = " & cur_P(newY.CREINTM42)
If newY.CREINTT43 <> oldY.CREINTT43 Then blnUpdate = True:  xSet = xSet & " , CREINTT43 = " & cur_P(newY.CREINTT43)
If newY.CREINTM43 <> oldY.CREINTM43 Then blnUpdate = True:  xSet = xSet & " , CREINTM43 = " & cur_P(newY.CREINTM43)
If newY.CREINTT44 <> oldY.CREINTT44 Then blnUpdate = True:  xSet = xSet & " , CREINTT44 = " & cur_P(newY.CREINTT44)
If newY.CREINTM44 <> oldY.CREINTM44 Then blnUpdate = True:  xSet = xSet & " , CREINTM44 = " & cur_P(newY.CREINTM44)



If blnUpdate Then

    Mid$(xSet, InStr(xSet, ","), 1) = " "
    xSql = "update " & paramIBM_Library_BODWH & ".DCREINT0" & xSet & xWhere
    
    Set rsSab = cnsab.Execute(xSql, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlDCREINT0_Update = "Erreur màj : " & newY.CREINTDOS
    
        Exit Function
    End If
End If
Exit Function
Error_Handler:
    sqlDCREINT0_Update = Error
End Function




'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsDCREINT0_GetBuffer(rsADO As ADODB.Recordset, rsDCREINT0 As typeDCREINT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsDCREINT0_GetBuffer = Null

rsDCREINT0.CREINTSTA = rsADO("CREINTSTA")
rsDCREINT0.CREINTVER = rsADO("CREINTVER")
rsDCREINT0.CREINTPER = rsADO("CREINTPER")

rsDCREINT0.CREINTETA = rsADO("CREINTETA")
rsDCREINT0.CREINTAGE = rsADO("CREINTAGE")
rsDCREINT0.CREINTSER = rsADO("CREINTSER")
rsDCREINT0.CREINTSSE = rsADO("CREINTSSE")
rsDCREINT0.CREINTDOS = rsADO("CREINTDOS")
rsDCREINT0.CREINTPRE = rsADO("CREINTPRE")
rsDCREINT0.CREINTNAT = rsADO("CREINTNAT")
rsDCREINT0.CREINTNAP = rsADO("CREINTNAP")
rsDCREINT0.CREINTCLI = rsADO("CREINTCLI")
rsDCREINT0.CREINTMT0 = rsADO("CREINTMT0")
rsDCREINT0.CREINTDEV = rsADO("CREINTDEV")

rsDCREINT0.CREINTECH = rsADO("CREINTECH")
rsDCREINT0.CREINTMTX = rsADO("CREINTMTX")
rsDCREINT0.CREINTTOF = rsADO("CREINTTOF")
rsDCREINT0.CREINTTOM = rsADO("CREINTTOM")
rsDCREINT0.CREINTPERK = rsADO("CREINTPERK")
rsDCREINT0.CREINTPERN = rsADO("CREINTPERN")
rsDCREINT0.CREINTUAMJ = rsADO("CREINTUAMJ")
rsDCREINT0.CREINTUHMS = rsADO("CREINTUHMS")

rsDCREINT0.CREINTT01 = rsADO("CREINTT01")
rsDCREINT0.CREINTM01 = rsADO("CREINTM01")
rsDCREINT0.CREINTT02 = rsADO("CREINTT02")
rsDCREINT0.CREINTM02 = rsADO("CREINTM02")
rsDCREINT0.CREINTT03 = rsADO("CREINTT03")
rsDCREINT0.CREINTM03 = rsADO("CREINTM03")
rsDCREINT0.CREINTT04 = rsADO("CREINTT04")
rsDCREINT0.CREINTM04 = rsADO("CREINTM04")


rsDCREINT0.CREINTT11 = rsADO("CREINTT11")
rsDCREINT0.CREINTM11 = rsADO("CREINTM11")
rsDCREINT0.CREINTT12 = rsADO("CREINTT12")
rsDCREINT0.CREINTM12 = rsADO("CREINTM12")
rsDCREINT0.CREINTT13 = rsADO("CREINTT13")
rsDCREINT0.CREINTM13 = rsADO("CREINTM13")
rsDCREINT0.CREINTT14 = rsADO("CREINTT14")
rsDCREINT0.CREINTM14 = rsADO("CREINTM14")

rsDCREINT0.CREINTT21 = rsADO("CREINTT21")
rsDCREINT0.CREINTM21 = rsADO("CREINTM21")
rsDCREINT0.CREINTT22 = rsADO("CREINTT22")
rsDCREINT0.CREINTM22 = rsADO("CREINTM22")
rsDCREINT0.CREINTT23 = rsADO("CREINTT23")
rsDCREINT0.CREINTM23 = rsADO("CREINTM23")
rsDCREINT0.CREINTT24 = rsADO("CREINTT24")
rsDCREINT0.CREINTM24 = rsADO("CREINTM24")

rsDCREINT0.CREINTT31 = rsADO("CREINTT31")
rsDCREINT0.CREINTM31 = rsADO("CREINTM31")
rsDCREINT0.CREINTT32 = rsADO("CREINTT32")
rsDCREINT0.CREINTM32 = rsADO("CREINTM32")
rsDCREINT0.CREINTT33 = rsADO("CREINTT33")
rsDCREINT0.CREINTM33 = rsADO("CREINTM33")
rsDCREINT0.CREINTT34 = rsADO("CREINTT34")
rsDCREINT0.CREINTM34 = rsADO("CREINTM34")

rsDCREINT0.CREINTT41 = rsADO("CREINTT41")
rsDCREINT0.CREINTM41 = rsADO("CREINTM41")
rsDCREINT0.CREINTT42 = rsADO("CREINTT42")
rsDCREINT0.CREINTM42 = rsADO("CREINTM42")
rsDCREINT0.CREINTT43 = rsADO("CREINTT43")
rsDCREINT0.CREINTM43 = rsADO("CREINTM43")
rsDCREINT0.CREINTT44 = rsADO("CREINTT44")
rsDCREINT0.CREINTM44 = rsADO("CREINTM44")


Exit Function

Error_Handler:

rsDCREINT0_GetBuffer = Error

End Function




Public Sub ddsCREINT0_Init()

ddsDCREINT0(1) = "A* CREINTSTA  statut"
ddsDCREINT0(2) = "N* CREINTVER  n° version"
ddsDCREINT0(3) = "N* CREINTPER  période de traitement"

ddsDCREINT0(4) = "A* CREINTETA  code établissement"
ddsDCREINT0(5) = "N* CREINTAGE  CODE AGENCE"
ddsDCREINT0(6) = "A* CREINTSER  CODE SERVICE"
ddsDCREINT0(7) = "A* CREINTSSE  CODE SOUS-SERVICE"
ddsDCREINT0(8) = "N* CREINTDOS  NUMÉRO OPÉRATION"
ddsDCREINT0(9) = "N* CREINTPRE  NUMERO PRET"
ddsDCREINT0(10) = "A* CREINTNAT  NAT/DOS"
ddsDCREINT0(11) = "A* CREINTNAP  NAT/PRE"
ddsDCREINT0(12) = "A* CREINTCLI  NUMÉRO CLIENT"
ddsDCREINT0(13) = "C* CREINTMT0  CAPITAL"
ddsDCREINT0(14) = "A* CREINTDEV  DEVISE"

ddsDCREINT0(15) = "D* CREINTECH  ÉCHÉANCE"
ddsDCREINT0(16) = "C* CREINTMTX  CAPITAL RESTANT DU"
ddsDCREINT0(17) = "N* CREINTTOF  TAUX"
ddsDCREINT0(18) = "N* CREINTTOM  MARGE"
ddsDCREINT0(19) = "A* CREINTPERK  PERIODICITE"
ddsDCREINT0(20) = "N* CREINTPERN  NB PERIODE"
ddsDCREINT0(21) = "D* CREINTUAMJ  DATE MAJ TA"
ddsDCREINT0(22) = "N* CREINTUHMS  HEURE MAJ TA"

ddsDCREINT0(23) = "N  CREINTT01  Intérêts Trim 01"
ddsDCREINT0(24) = "N  CREINTM01  Marge Trim 01"
ddsDCREINT0(25) = "N  CREINTT02  Intérêts Trim 02"
ddsDCREINT0(26) = "N  CREINTM02  Marge Trim 02"
ddsDCREINT0(27) = "N  CREINTT03  Intérêts Trim 03"
ddsDCREINT0(28) = "N  CREINTM03  Marge Trim 03"
ddsDCREINT0(29) = "N  CREINTT04  Intérêts Trim 04"
ddsDCREINT0(30) = "N  CREINTM04  Marge Trim 04"

ddsDCREINT0(31) = "N  CREINTT11  Intérêts Trim 11"
ddsDCREINT0(32) = "N  CREINTM11  Marge Trim 11"
ddsDCREINT0(33) = "N  CREINTT12  Intérêts Trim 12"
ddsDCREINT0(34) = "N  CREINTM12  Marge Trim 12"
ddsDCREINT0(35) = "N  CREINTT13  Intérêts Trim 13"
ddsDCREINT0(36) = "N  CREINTM13  Marge Trim 13"
ddsDCREINT0(37) = "N  CREINTT14  Intérêts Trim 14"
ddsDCREINT0(38) = "N  CREINTM14  Marge Trim 14"

ddsDCREINT0(39) = "N  CREINTT21  Intérêts Trim 21"
ddsDCREINT0(40) = "N  CREINTM21  Marge Trim 21"
ddsDCREINT0(41) = "N  CREINTT22  Intérêts Trim 22"
ddsDCREINT0(42) = "N  CREINTM22  Marge Trim 22"
ddsDCREINT0(43) = "N  CREINTT23  Intérêts Trim 23"
ddsDCREINT0(44) = "N  CREINTM23  Marge Trim 23"
ddsDCREINT0(45) = "N  CREINTT24  Intérêts Trim 24"
ddsDCREINT0(46) = "N  CREINTM24  Marge Trim 24"

ddsDCREINT0(47) = "N  CREINTT31  Intérêts Trim 31"
ddsDCREINT0(48) = "N  CREINTM31  Marge Trim 31"
ddsDCREINT0(49) = "N  CREINTT32  Intérêts Trim 32"
ddsDCREINT0(50) = "N  CREINTM32  Marge Trim 32"
ddsDCREINT0(51) = "N  CREINTT33  Intérêts Trim 33"
ddsDCREINT0(52) = "N  CREINTM33  Marge Trim 33"
ddsDCREINT0(53) = "N  CREINTT34  Intérêts Trim 34"
ddsDCREINT0(54) = "N  CREINTM34  Marge Trim 34"

ddsDCREINT0(55) = "N  CREINTT41  Intérêts Trim 41"
ddsDCREINT0(56) = "N  CREINTM41  Marge Trim 41"
ddsDCREINT0(57) = "N  CREINTT42  Intérêts Trim 42"
ddsDCREINT0(58) = "N  CREINTM42  Marge Trim 42"
ddsDCREINT0(59) = "N  CREINTT43  Intérêts Trim 43"
ddsDCREINT0(60) = "N  CREINTM43  Marge Trim 43"
ddsDCREINT0(61) = "N  CREINTT44  Intérêts Trim 44"
ddsDCREINT0(62) = "N  CREINTM44  Marge Trim 44"

End Sub

