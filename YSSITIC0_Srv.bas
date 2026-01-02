Attribute VB_Name = "srvYSSITIC0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSITIC0
    SSITICNAT    As String
    SSITICUIDX   As String
       
    SSITICSTAK   As String
    SSITICUIDD   As Long
    SSITICPRFX   As String
    SSITICPRFK   As String
    SSITICUNOM    As String
    
    SSITICTLNK   As Long
    SSITICYFCT   As String
    SSITICYUSR   As String
    SSITICYAMJ   As Long
    SSITICYHMS   As Long
    SSITICYVER   As Long
    
    SSITICINFO   As String
End Type


'Public xYSSITIC0 As typeYSSITIC0
Public Function sqlYSSITIC0_Insert(newY As typeYSSITIC0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSITIC0_Insert = Null

xSet = " (SSITICNAT,SSITICUIDX,SSITICUIDD"
xValues = " values('" & newY.SSITICNAT & "','" & newY.SSITICUIDX & "'," & newY.SSITICUIDD

' Détecter les modifications
'===================================================================================
If newY.SSITICTLNK <> 0 Then xSet = xSet & ",SSITICTLNK": xValues = xValues & " ," & newY.SSITICTLNK
If newY.SSITICYVER <> 0 Then xSet = xSet & ",SSITICYVER": xValues = xValues & " ," & newY.SSITICYVER
If newY.SSITICYAMJ <> 0 Then xSet = xSet & ",SSITICYAMJ": xValues = xValues & " ," & newY.SSITICYAMJ
If newY.SSITICYHMS <> 0 Then xSet = xSet & ",SSITICYHMS": xValues = xValues & " ," & newY.SSITICYHMS

If Trim(newY.SSITICSTAK) <> "" Then xSet = xSet & ",SSITICSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICSTAK), "'", "''") & "'"
If Trim(newY.SSITICPRFX) <> "" Then xSet = xSet & ",SSITICPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICPRFX), "'", "''") & "'"
If Trim(newY.SSITICPRFK) <> "" Then xSet = xSet & ",SSITICPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICPRFK), "'", "''") & "'"
If Trim(newY.SSITICUNOM) <> "" Then xSet = xSet & ",SSITICUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICUNOM), "'", "''") & "'"
If Trim(newY.SSITICYFCT) <> "" Then xSet = xSet & ",SSITICYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICYFCT), "'", "''") & "'"
If Trim(newY.SSITICYUSR) <> "" Then xSet = xSet & ",SSITICYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICYUSR), "'", "''") & "'"
If Trim(newY.SSITICINFO) <> "" Then xSet = xSet & ",SSITICINFO": xValues = xValues & " ,'" & Replace(RTrim(newY.SSITICINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSITIC0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSITIC0_Insert = "Erreur màj : " & newY.SSITICNAT & " " & newY.SSITICUIDX
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSITIC0_Insert = "sqlYSSITIC0_Insert " & newY.SSITICNAT & " " & newY.SSITICUIDX & vbCrLf & Error
End Function
Public Function sqlYSSITICH_Insert(newY As typeYSSITIC0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSITICH_Insert = Null

xSet = " (SSITICNAT,SSITICUIDX,SSITICUIDD"
xValues = " values('" & newY.SSITICNAT & "','" & newY.SSITICUIDX & "'," & newY.SSITICUIDD

' Détecter les modifications
'===================================================================================
If newY.SSITICTLNK <> 0 Then xSet = xSet & ",SSITICTLNK": xValues = xValues & " ," & newY.SSITICTLNK
If newY.SSITICYVER <> 0 Then xSet = xSet & ",SSITICYVER": xValues = xValues & " ," & newY.SSITICYVER
If newY.SSITICYAMJ <> 0 Then xSet = xSet & ",SSITICYAMJ": xValues = xValues & " ," & newY.SSITICYAMJ
If newY.SSITICYHMS <> 0 Then xSet = xSet & ",SSITICYHMS": xValues = xValues & " ," & newY.SSITICYHMS

If Trim(newY.SSITICSTAK) <> "" Then xSet = xSet & ",SSITICSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICSTAK), "'", "''") & "'"
If Trim(newY.SSITICPRFX) <> "" Then xSet = xSet & ",SSITICPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICPRFX), "'", "''") & "'"
If Trim(newY.SSITICPRFK) <> "" Then xSet = xSet & ",SSITICPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICPRFK), "'", "''") & "'"
If Trim(newY.SSITICUNOM) <> "" Then xSet = xSet & ",SSITICUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICUNOM), "'", "''") & "'"
If Trim(newY.SSITICYFCT) <> "" Then xSet = xSet & ",SSITICYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICYFCT), "'", "''") & "'"
If Trim(newY.SSITICYUSR) <> "" Then xSet = xSet & ",SSITICYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSITICYUSR), "'", "''") & "'"
If Trim(newY.SSITICINFO) <> "" Then xSet = xSet & ",SSITICINFO": xValues = xValues & " ,'" & Replace(RTrim(newY.SSITICINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSITICH" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSITICH_Insert = "Erreur màj : " & newY.SSITICNAT & " " & newY.SSITICUIDX & " " & newY.SSITICYVER
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSITICH_Insert = "sqlYSSITICH_Insert " & newY.SSITICUIDX & " " & newY.SSITICYVER & vbCrLf & Error
End Function


Public Function sqlYSSITIC0_Update(newY As typeYSSITIC0, oldY As typeYSSITIC0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSITIC0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSITICUIDX <> newY.SSITICUIDX Then
    sqlYSSITIC0_Update = "Erreur SSITICUIDX : " & newY.SSITICUIDX & " / " & oldY.SSITICUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSITICNAT = '" & oldY.SSITICNAT & "'" _
       & " and SSITICUIDX = '" & oldY.SSITICUIDX & "'" _
       & " and SSITICUIDD = " & oldY.SSITICUIDD _
       & " and SSITICYVER = " & oldY.SSITICYVER
       
newY.SSITICYVER = newY.SSITICYVER + 1
xSet = xSet & " set SSITICYVER = " & newY.SSITICYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSITICTLNK <> oldY.SSITICTLNK Then blnUpdate = True: xSet = xSet & " , SSITICTLNK = " & newY.SSITICTLNK
If newY.SSITICYAMJ <> oldY.SSITICYAMJ Then blnUpdate = True: xSet = xSet & " , SSITICYAMJ = " & newY.SSITICYAMJ
If newY.SSITICYHMS <> oldY.SSITICYHMS Then blnUpdate = True: xSet = xSet & " , SSITICYHMS = " & newY.SSITICYHMS

If newY.SSITICSTAK <> oldY.SSITICSTAK Then blnUpdate = True:  xSet = xSet & " , SSITICSTAK = '" & Replace(Trim(newY.SSITICSTAK), "'", "''") & "'"
If newY.SSITICPRFX <> oldY.SSITICPRFX Then blnUpdate = True:  xSet = xSet & " , SSITICPRFX = '" & Replace(Trim(newY.SSITICPRFX), "'", "''") & "'"
If newY.SSITICPRFK <> oldY.SSITICPRFK Then blnUpdate = True:  xSet = xSet & " , SSITICPRFK = '" & Replace(Trim(newY.SSITICPRFK), "'", "''") & "'"
If newY.SSITICUNOM <> oldY.SSITICUNOM Then blnUpdate = True:  xSet = xSet & " , SSITICUNOM = '" & Replace(Trim(newY.SSITICUNOM), "'", "''") & "'"
If newY.SSITICYFCT <> oldY.SSITICYFCT Then blnUpdate = True:  xSet = xSet & " , SSITICYFCT = '" & Replace(Trim(newY.SSITICYFCT), "'", "''") & "'"
If newY.SSITICYUSR <> oldY.SSITICYUSR Then blnUpdate = True:  xSet = xSet & " , SSITICYUSR = '" & Replace(Trim(newY.SSITICYUSR), "'", "''") & "'"
If newY.SSITICINFO <> oldY.SSITICINFO Then blnUpdate = True:  xSet = xSet & " , SSITICINFO= '" & Replace(RTrim(newY.SSITICINFO), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSITIC0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSITIC0_Update = "Erreur màj : " & newY.SSITICNAT & " " & newY.SSITICUIDX
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSITIC0_Update = "sqlYSSITIC0_Update " & vbCrLf & Error
End Function
Public Function sqlYSSITICH_Update(newY As typeYSSITIC0, oldY As typeYSSITIC0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSITICH_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSITICUIDX <> newY.SSITICUIDX Then
    sqlYSSITICH_Update = "Erreur SSITICUIDX : " & newY.SSITICUIDX & " / " & oldY.SSITICUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSITICNAT = '" & oldY.SSITICNAT & "'" _
       & " and SSITICUIDX = '" & oldY.SSITICUIDX & "'" _
       & " and SSITICUIDD = " & oldY.SSITICUIDD _
       & " and SSITICYVER = " & oldY.SSITICYVER
       
xSet = xSet & " set SSITICYVER = " & newY.SSITICYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSITICTLNK <> oldY.SSITICTLNK Then blnUpdate = True: xSet = xSet & " , SSITICTLNK = " & newY.SSITICTLNK
If newY.SSITICYAMJ <> oldY.SSITICYAMJ Then blnUpdate = True: xSet = xSet & " , SSITICYAMJ = " & newY.SSITICYAMJ
If newY.SSITICYHMS <> oldY.SSITICYHMS Then blnUpdate = True: xSet = xSet & " , SSITICYHMS = " & newY.SSITICYHMS

If newY.SSITICSTAK <> oldY.SSITICSTAK Then blnUpdate = True:  xSet = xSet & " , SSITICSTAK = '" & Replace(Trim(newY.SSITICSTAK), "'", "''") & "'"
If newY.SSITICPRFX <> oldY.SSITICPRFX Then blnUpdate = True:  xSet = xSet & " , SSITICPRFX = '" & Replace(Trim(newY.SSITICPRFX), "'", "''") & "'"
If newY.SSITICPRFK <> oldY.SSITICPRFK Then blnUpdate = True:  xSet = xSet & " , SSITICPRFK = '" & Replace(Trim(newY.SSITICPRFK), "'", "''") & "'"
If newY.SSITICUNOM <> oldY.SSITICUNOM Then blnUpdate = True:  xSet = xSet & " , SSITICUNOM = '" & Replace(Trim(newY.SSITICUNOM), "'", "''") & "'"
If newY.SSITICYFCT <> oldY.SSITICYFCT Then blnUpdate = True:  xSet = xSet & " , SSITICYFCT = '" & Replace(Trim(newY.SSITICYFCT), "'", "''") & "'"
If newY.SSITICYUSR <> oldY.SSITICYUSR Then blnUpdate = True:  xSet = xSet & " , SSITICYUSR = '" & Replace(Trim(newY.SSITICYUSR), "'", "''") & "'"
If newY.SSITICINFO <> oldY.SSITICINFO Then blnUpdate = True:  xSet = xSet & " , SSITICINFO= '" & Replace(RTrim(newY.SSITICINFO), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSITICH" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
       sqlYSSITICH_Update = "Erreur màj : " & newY.SSITICNAT & " " & newY.SSITICUIDX & " " & newY.SSITICYVER
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSITICH_Update = "sqlYSSITICH_Update " & vbCrLf & Error
End Function


Public Function rsYSSITIC0_GetBuffer(rsAdo As ADODB.Recordset, lYSSITIC0 As typeYSSITIC0)
On Error GoTo Error_Handler
rsYSSITIC0_GetBuffer = Null
lYSSITIC0.SSITICNAT = rsAdo("SSITICNAT")
lYSSITIC0.SSITICUIDX = Trim(rsAdo("SSITICUIDX"))

lYSSITIC0.SSITICSTAK = rsAdo("SSITICSTAK")
lYSSITIC0.SSITICUIDD = rsAdo("SSITICUIDD")
lYSSITIC0.SSITICPRFX = Trim(rsAdo("SSITICPRFX"))
lYSSITIC0.SSITICPRFK = rsAdo("SSITICPRFK")
lYSSITIC0.SSITICUNOM = Trim(rsAdo("SSITICUNOM"))

lYSSITIC0.SSITICTLNK = rsAdo("SSITICTLNK")
lYSSITIC0.SSITICYFCT = Trim(rsAdo("SSITICYFCT"))
lYSSITIC0.SSITICYUSR = Trim(rsAdo("SSITICYUSR"))
lYSSITIC0.SSITICYAMJ = rsAdo("SSITICYAMJ")
lYSSITIC0.SSITICYHMS = rsAdo("SSITICYHMS")
lYSSITIC0.SSITICYVER = rsAdo("SSITICYVER")

lYSSITIC0.SSITICINFO = RTrim(rsAdo("SSITICINFO")) '!!!!!! RTrim , ne pas faire Trim

Exit Function
Error_Handler:
rsYSSITIC0_GetBuffer = Error


End Function

Public Function rsYSSITIC0_Init(lYSSITIC0 As typeYSSITIC0)
lYSSITIC0.SSITICNAT = ""
lYSSITIC0.SSITICUIDX = ""

lYSSITIC0.SSITICSTAK = ""
lYSSITIC0.SSITICUIDD = 0
lYSSITIC0.SSITICPRFX = ""
lYSSITIC0.SSITICPRFK = ""
lYSSITIC0.SSITICUNOM = ""

lYSSITIC0.SSITICTLNK = 0
lYSSITIC0.SSITICYFCT = ""
lYSSITIC0.SSITICYUSR = ""
lYSSITIC0.SSITICYAMJ = 0
lYSSITIC0.SSITICYHMS = 0
lYSSITIC0.SSITICYVER = 0
lYSSITIC0.SSITICINFO = ""
End Function

