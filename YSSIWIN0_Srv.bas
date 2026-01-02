Attribute VB_Name = "srvYSSIWIN0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSIWIN0
    SSIWINNAT    As String
    SSIWINUIDX   As String
    
    SSIWINSTAK   As String
    SSIWINUIDD   As Long
    SSIWINPRFX   As String
    SSIWINPRFK   As String
    SSIWINUNOM    As String
    
    SSIWINTLNK   As Long
    SSIWINYFCT   As String
    SSIWINYUSR   As String
    SSIWINYAMJ   As Long
    SSIWINYHMS   As Long
    SSIWINYVER   As Long
    
    
    SSIWINMAIL   As String
    SSIWINGUID   As String
    SSIWININFO   As String
End Type


'Public xYSSIWIN0 As typeYSSIWIN0
Public Function sqlYSSIWIN0_Insert(newY As typeYSSIWIN0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIWIN0_Insert = Null

xSet = " (SSIWINNAT,SSIWINGUID"
xValues = " values('" & newY.SSIWINNAT & "','" & newY.SSIWINGUID & "'"

' Détecter les modifications
'===================================================================================
If newY.SSIWINUIDD <> 0 Then xSet = xSet & ",SSIWINUIDD": xValues = xValues & " ," & newY.SSIWINUIDD
If newY.SSIWINTLNK <> 0 Then xSet = xSet & ",SSIWINTLNK": xValues = xValues & " ," & newY.SSIWINTLNK
If newY.SSIWINYVER <> 0 Then xSet = xSet & ",SSIWINYVER": xValues = xValues & " ," & newY.SSIWINYVER
If newY.SSIWINYAMJ <> 0 Then xSet = xSet & ",SSIWINYAMJ": xValues = xValues & " ," & newY.SSIWINYAMJ
If newY.SSIWINYHMS <> 0 Then xSet = xSet & ",SSIWINYHMS": xValues = xValues & " ," & newY.SSIWINYHMS

If Trim(newY.SSIWINSTAK) <> "" Then xSet = xSet & ",SSIWINSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINSTAK), "'", "''") & "'"
If Trim(newY.SSIWINPRFX) <> "" Then xSet = xSet & ",SSIWINPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINPRFX), "'", "''") & "'"
If Trim(newY.SSIWINPRFK) <> "" Then xSet = xSet & ",SSIWINPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINPRFK), "'", "''") & "'"
If Trim(newY.SSIWINUNOM) <> "" Then xSet = xSet & ",SSIWINUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINUNOM), "'", "''") & "'"
If Trim(newY.SSIWINYFCT) <> "" Then xSet = xSet & ",SSIWINYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINYFCT), "'", "''") & "'"
If Trim(newY.SSIWINYUSR) <> "" Then xSet = xSet & ",SSIWINYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINYUSR), "'", "''") & "'"
If Trim(newY.SSIWINMAIL) <> "" Then xSet = xSet & ",SSIWINMAIL": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINMAIL), "'", "''") & "'"
If Trim(newY.SSIWINUIDX) <> "" Then xSet = xSet & ",SSIWINUIDX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINUIDX), "'", "''") & "'"
If Trim(newY.SSIWININFO) <> "" Then xSet = xSet & ",SSIWININFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWININFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIWIN0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIWIN0_Insert = "Erreur màj : " & newY.SSIWINPRFK
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIWIN0_Insert = "sqlYSSIWIN0_Insert " & vbCrLf & Error
End Function
Public Function sqlYSSIWINH_Insert(newY As typeYSSIWIN0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIWINH_Insert = Null

xSet = " (SSIWINNAT,SSIWINGUID"
xValues = " values('" & newY.SSIWINNAT & "','" & newY.SSIWINGUID & "'"

' Détecter les modifications
'===================================================================================
If newY.SSIWINUIDD <> 0 Then xSet = xSet & ",SSIWINUIDD": xValues = xValues & " ," & newY.SSIWINUIDD
If newY.SSIWINTLNK <> 0 Then xSet = xSet & ",SSIWINTLNK": xValues = xValues & " ," & newY.SSIWINTLNK
If newY.SSIWINYVER <> 0 Then xSet = xSet & ",SSIWINYVER": xValues = xValues & " ," & newY.SSIWINYVER
If newY.SSIWINYAMJ <> 0 Then xSet = xSet & ",SSIWINYAMJ": xValues = xValues & " ," & newY.SSIWINYAMJ
If newY.SSIWINYHMS <> 0 Then xSet = xSet & ",SSIWINYHMS": xValues = xValues & " ," & newY.SSIWINYHMS

If Trim(newY.SSIWINSTAK) <> "" Then xSet = xSet & ",SSIWINSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINSTAK), "'", "''") & "'"
If Trim(newY.SSIWINPRFX) <> "" Then xSet = xSet & ",SSIWINPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINPRFX), "'", "''") & "'"
If Trim(newY.SSIWINPRFK) <> "" Then xSet = xSet & ",SSIWINPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINPRFK), "'", "''") & "'"
If Trim(newY.SSIWINUNOM) <> "" Then xSet = xSet & ",SSIWINUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINUNOM), "'", "''") & "'"
If Trim(newY.SSIWINYFCT) <> "" Then xSet = xSet & ",SSIWINYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINYFCT), "'", "''") & "'"
If Trim(newY.SSIWINYUSR) <> "" Then xSet = xSet & ",SSIWINYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINYUSR), "'", "''") & "'"
If Trim(newY.SSIWINMAIL) <> "" Then xSet = xSet & ",SSIWINMAIL": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINMAIL), "'", "''") & "'"
If Trim(newY.SSIWINUIDX) <> "" Then xSet = xSet & ",SSIWINUIDX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWINUIDX), "'", "''") & "'"
If Trim(newY.SSIWININFO) <> "" Then xSet = xSet & ",SSIWININFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSIWININFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIWINH" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIWINH_Insert = "Erreur màj : " & newY.SSIWINPRFK
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIWINH_Insert = "sqlYSSIWINH_Insert " & vbCrLf & Error
End Function


Public Function sqlYSSIWIN0_Update(newY As typeYSSIWIN0, oldY As typeYSSIWIN0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIWIN0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIWINGUID <> newY.SSIWINGUID Then
    sqlYSSIWIN0_Update = "Erreur SSIWINGUID : " & newY.SSIWINGUID & " / " & oldY.SSIWINGUID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIWINNAT = '" & oldY.SSIWINNAT & "'" _
       & " and SSIWINGUID = '" & oldY.SSIWINGUID & "'" _
       & " and SSIWINYVER = " & oldY.SSIWINYVER
       
newY.SSIWINYVER = newY.SSIWINYVER + 1
xSet = xSet & " set SSIWINYVER = " & newY.SSIWINYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIWINUIDD <> oldY.SSIWINUIDD Then blnUpdate = True: xSet = xSet & " , SSIWINUIDD = " & newY.SSIWINUIDD
If newY.SSIWINTLNK <> oldY.SSIWINTLNK Then blnUpdate = True: xSet = xSet & " , SSIWINTLNK = " & newY.SSIWINTLNK
If newY.SSIWINYAMJ <> oldY.SSIWINYAMJ Then blnUpdate = True: xSet = xSet & " , SSIWINYAMJ = " & newY.SSIWINYAMJ
If newY.SSIWINYHMS <> oldY.SSIWINYHMS Then blnUpdate = True: xSet = xSet & " , SSIWINYHMS = " & newY.SSIWINYHMS

If newY.SSIWINSTAK <> oldY.SSIWINSTAK Then blnUpdate = True:  xSet = xSet & " , SSIWINSTAK = '" & Replace(Trim(newY.SSIWINSTAK), "'", "''") & "'"
If newY.SSIWINPRFX <> oldY.SSIWINPRFX Then blnUpdate = True:  xSet = xSet & " , SSIWINPRFX = '" & Replace(Trim(newY.SSIWINPRFX), "'", "''") & "'"
If newY.SSIWINPRFK <> oldY.SSIWINPRFK Then blnUpdate = True:  xSet = xSet & " , SSIWINPRFK = '" & Replace(Trim(newY.SSIWINPRFK), "'", "''") & "'"
If newY.SSIWINUNOM <> oldY.SSIWINUNOM Then blnUpdate = True:  xSet = xSet & " , SSIWINUNOM = '" & Replace(Trim(newY.SSIWINUNOM), "'", "''") & "'"
If newY.SSIWINYFCT <> oldY.SSIWINYFCT Then blnUpdate = True:  xSet = xSet & " , SSIWINYFCT = '" & Replace(Trim(newY.SSIWINYFCT), "'", "''") & "'"
If newY.SSIWINYUSR <> oldY.SSIWINYUSR Then blnUpdate = True:  xSet = xSet & " , SSIWINYUSR = '" & Replace(Trim(newY.SSIWINYUSR), "'", "''") & "'"
If newY.SSIWINMAIL <> oldY.SSIWINMAIL Then blnUpdate = True:  xSet = xSet & " , SSIWINMAIL= '" & Replace(Trim(newY.SSIWINMAIL), "'", "''") & "'"
If newY.SSIWINUIDX <> oldY.SSIWINUIDX Then blnUpdate = True:  xSet = xSet & " , SSIWINUIDX= '" & Replace(Trim(newY.SSIWINUIDX), "'", "''") & "'"
If newY.SSIWININFO <> oldY.SSIWININFO Then blnUpdate = True:  xSet = xSet & " , SSIWININFO= '" & Replace(Trim(newY.SSIWININFO), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIWIN0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIWIN0_Update = "Erreur màj : " & newY.SSIWINUIDX
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIWIN0_Update = "sqlYSSIWIN0_Update " & vbCrLf & Error
End Function
Public Function sqlYSSIWINH_Update(newY As typeYSSIWIN0, oldY As typeYSSIWIN0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIWINH_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIWINGUID <> newY.SSIWINGUID Then
    sqlYSSIWINH_Update = "Erreur SSIWINGUID : " & newY.SSIWINGUID & " / " & oldY.SSIWINGUID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIWINNAT = '" & oldY.SSIWINNAT & "'" _
       & " and SSIWINGUID = '" & oldY.SSIWINGUID & "'" _
       & " and SSIWINYVER = " & oldY.SSIWINYVER
       
xSet = xSet & " set SSIWINYVER = " & newY.SSIWINYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIWINUIDD <> oldY.SSIWINUIDD Then blnUpdate = True: xSet = xSet & " , SSIWINUIDD = " & newY.SSIWINUIDD
If newY.SSIWINTLNK <> oldY.SSIWINTLNK Then blnUpdate = True: xSet = xSet & " , SSIWINTLNK = " & newY.SSIWINTLNK
If newY.SSIWINYAMJ <> oldY.SSIWINYAMJ Then blnUpdate = True: xSet = xSet & " , SSIWINYAMJ = " & newY.SSIWINYAMJ
If newY.SSIWINYHMS <> oldY.SSIWINYHMS Then blnUpdate = True: xSet = xSet & " , SSIWINYHMS = " & newY.SSIWINYHMS

If newY.SSIWINSTAK <> oldY.SSIWINSTAK Then blnUpdate = True:  xSet = xSet & " , SSIWINSTAK = '" & Replace(Trim(newY.SSIWINSTAK), "'", "''") & "'"
If newY.SSIWINPRFX <> oldY.SSIWINPRFX Then blnUpdate = True:  xSet = xSet & " , SSIWINPRFX = '" & Replace(Trim(newY.SSIWINPRFX), "'", "''") & "'"
If newY.SSIWINPRFK <> oldY.SSIWINPRFK Then blnUpdate = True:  xSet = xSet & " , SSIWINPRFK = '" & Replace(Trim(newY.SSIWINPRFK), "'", "''") & "'"
If newY.SSIWINUNOM <> oldY.SSIWINUNOM Then blnUpdate = True:  xSet = xSet & " , SSIWINUNOM = '" & Replace(Trim(newY.SSIWINUNOM), "'", "''") & "'"
If newY.SSIWINYFCT <> oldY.SSIWINYFCT Then blnUpdate = True:  xSet = xSet & " , SSIWINYFCT = '" & Replace(Trim(newY.SSIWINYFCT), "'", "''") & "'"
If newY.SSIWINYUSR <> oldY.SSIWINYUSR Then blnUpdate = True:  xSet = xSet & " , SSIWINYUSR = '" & Replace(Trim(newY.SSIWINYUSR), "'", "''") & "'"
If newY.SSIWINMAIL <> oldY.SSIWINMAIL Then blnUpdate = True:  xSet = xSet & " , SSIWINMAIL= '" & Replace(Trim(newY.SSIWINMAIL), "'", "''") & "'"
If newY.SSIWINUIDX <> oldY.SSIWINUIDX Then blnUpdate = True:  xSet = xSet & " , SSIWINUIDX= '" & Replace(Trim(newY.SSIWINUIDX), "'", "''") & "'"
If newY.SSIWININFO <> oldY.SSIWININFO Then blnUpdate = True:  xSet = xSet & " , SSIWININFO= '" & Replace(Trim(newY.SSIWININFO), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIWINH" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIWINH_Update = "Erreur màj : " & newY.SSIWINUIDX
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIWINH_Update = "sqlYSSIWINH_Update " & vbCrLf & Error
End Function


Public Function rsYSSIWIN0_GetBuffer(rsAdo As ADODB.Recordset, lYSSIWIN0 As typeYSSIWIN0)
On Error GoTo Error_Handler
rsYSSIWIN0_GetBuffer = Null
lYSSIWIN0.SSIWINNAT = rsAdo("SSIWINNAT")
lYSSIWIN0.SSIWINUIDX = Trim(rsAdo("SSIWINUIDX"))

lYSSIWIN0.SSIWINSTAK = rsAdo("SSIWINSTAK")
lYSSIWIN0.SSIWINUIDD = rsAdo("SSIWINUIDD")
lYSSIWIN0.SSIWINPRFX = Trim(rsAdo("SSIWINPRFX"))
lYSSIWIN0.SSIWINPRFK = rsAdo("SSIWINPRFK")
lYSSIWIN0.SSIWINUNOM = Trim(rsAdo("SSIWINUNOM"))

lYSSIWIN0.SSIWINTLNK = rsAdo("SSIWINTLNK")
lYSSIWIN0.SSIWINYFCT = Trim(rsAdo("SSIWINYFCT"))
lYSSIWIN0.SSIWINYUSR = Trim(rsAdo("SSIWINYUSR"))
lYSSIWIN0.SSIWINYAMJ = rsAdo("SSIWINYAMJ")
lYSSIWIN0.SSIWINYHMS = rsAdo("SSIWINYHMS")
lYSSIWIN0.SSIWINYVER = rsAdo("SSIWINYVER")

lYSSIWIN0.SSIWINMAIL = Trim(rsAdo("SSIWINMAIL"))
lYSSIWIN0.SSIWINGUID = Trim(rsAdo("SSIWINGUID"))
lYSSIWIN0.SSIWININFO = Trim(rsAdo("SSIWININFO"))

Exit Function
Error_Handler:
rsYSSIWIN0_GetBuffer = Error


End Function

Public Function rsYSSIWIN0_Init(lYSSIWIN0 As typeYSSIWIN0)
lYSSIWIN0.SSIWINNAT = ""
lYSSIWIN0.SSIWINUIDX = ""

lYSSIWIN0.SSIWINSTAK = ""
lYSSIWIN0.SSIWINUIDD = 0
lYSSIWIN0.SSIWINPRFX = ""
lYSSIWIN0.SSIWINPRFK = ""
lYSSIWIN0.SSIWINUNOM = ""

lYSSIWIN0.SSIWINTLNK = 0
lYSSIWIN0.SSIWINYFCT = ""
lYSSIWIN0.SSIWINYUSR = ""
lYSSIWIN0.SSIWINYAMJ = 0
lYSSIWIN0.SSIWINYHMS = 0
lYSSIWIN0.SSIWINYVER = 0
lYSSIWIN0.SSIWINMAIL = ""
lYSSIWIN0.SSIWINGUID = ""
lYSSIWIN0.SSIWININFO = ""
End Function











