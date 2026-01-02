Attribute VB_Name = "srvYSSIDIV0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSIDIV0
    SSIDIVNAT    As String
    SSIDIVUIDX   As String
    
    SSIDIVDIDK   As String
    
    SSIDIVSTAK   As String
    SSIDIVUIDD   As Long
    SSIDIVPRFX   As String
    SSIDIVPRFK   As String
    SSIDIVUNOM    As String
    
    SSIDIVTLNK   As Long
    SSIDIVYFCT   As String
    SSIDIVYUSR   As String
    SSIDIVYAMJ   As Long
    SSIDIVYHMS   As Long
    SSIDIVYVER   As Long
    
    SSIDIVINFO   As String
End Type


'Public xYSSIDIV0 As typeYSSIDIV0
Public Function sqlYSSIDIV0_Insert(newY As typeYSSIDIV0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIDIV0_Insert = Null

xSet = " (SSIDIVNAT,SSIDIVUIDX,SSIDIVUIDD"
xValues = " values('" & newY.SSIDIVNAT & "','" & newY.SSIDIVUIDX & "'," & newY.SSIDIVUIDD

' Détecter les modifications
'===================================================================================
If newY.SSIDIVTLNK <> 0 Then xSet = xSet & ",SSIDIVTLNK": xValues = xValues & " ," & newY.SSIDIVTLNK
If newY.SSIDIVYVER <> 0 Then xSet = xSet & ",SSIDIVYVER": xValues = xValues & " ," & newY.SSIDIVYVER
If newY.SSIDIVYAMJ <> 0 Then xSet = xSet & ",SSIDIVYAMJ": xValues = xValues & " ," & newY.SSIDIVYAMJ
If newY.SSIDIVYHMS <> 0 Then xSet = xSet & ",SSIDIVYHMS": xValues = xValues & " ," & newY.SSIDIVYHMS

If Trim(newY.SSIDIVSTAK) <> "" Then xSet = xSet & ",SSIDIVSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVSTAK), "'", "''") & "'"
If Trim(newY.SSIDIVPRFX) <> "" Then xSet = xSet & ",SSIDIVPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVPRFX), "'", "''") & "'"
If Trim(newY.SSIDIVPRFK) <> "" Then xSet = xSet & ",SSIDIVPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVPRFK), "'", "''") & "'"
If Trim(newY.SSIDIVUNOM) <> "" Then xSet = xSet & ",SSIDIVUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVUNOM), "'", "''") & "'"
If Trim(newY.SSIDIVYFCT) <> "" Then xSet = xSet & ",SSIDIVYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVYFCT), "'", "''") & "'"
If Trim(newY.SSIDIVYUSR) <> "" Then xSet = xSet & ",SSIDIVYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVYUSR), "'", "''") & "'"
If Trim(newY.SSIDIVDIDK) <> "" Then xSet = xSet & ",SSIDIVDIDK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVDIDK), "'", "''") & "'"
If Trim(newY.SSIDIVINFO) <> "" Then xSet = xSet & ",SSIDIVINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIDIV0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIDIV0_Insert = "Erreur màj : " & newY.SSIDIVPRFK
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIDIV0_Insert = "sqlYSSIDIV0_Insert " & vbCrLf & Error
End Function
Public Function sqlYSSIDIVH_Insert(newY As typeYSSIDIV0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIDIVH_Insert = Null

xSet = " (SSIDIVNAT,SSIDIVUIDX,SSIDIVUIDD"
xValues = " values('" & newY.SSIDIVNAT & "','" & newY.SSIDIVUIDX & "'," & newY.SSIDIVUIDD

' Détecter les modifications
'===================================================================================
If newY.SSIDIVTLNK <> 0 Then xSet = xSet & ",SSIDIVTLNK": xValues = xValues & " ," & newY.SSIDIVTLNK
If newY.SSIDIVYVER <> 0 Then xSet = xSet & ",SSIDIVYVER": xValues = xValues & " ," & newY.SSIDIVYVER
If newY.SSIDIVYAMJ <> 0 Then xSet = xSet & ",SSIDIVYAMJ": xValues = xValues & " ," & newY.SSIDIVYAMJ
If newY.SSIDIVYHMS <> 0 Then xSet = xSet & ",SSIDIVYHMS": xValues = xValues & " ," & newY.SSIDIVYHMS

If Trim(newY.SSIDIVSTAK) <> "" Then xSet = xSet & ",SSIDIVSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVSTAK), "'", "''") & "'"
If Trim(newY.SSIDIVPRFX) <> "" Then xSet = xSet & ",SSIDIVPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVPRFX), "'", "''") & "'"
If Trim(newY.SSIDIVPRFK) <> "" Then xSet = xSet & ",SSIDIVPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVPRFK), "'", "''") & "'"
If Trim(newY.SSIDIVUNOM) <> "" Then xSet = xSet & ",SSIDIVUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVUNOM), "'", "''") & "'"
If Trim(newY.SSIDIVYFCT) <> "" Then xSet = xSet & ",SSIDIVYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVYFCT), "'", "''") & "'"
If Trim(newY.SSIDIVYUSR) <> "" Then xSet = xSet & ",SSIDIVYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVYUSR), "'", "''") & "'"
If Trim(newY.SSIDIVDIDK) <> "" Then xSet = xSet & ",SSIDIVDIDK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVDIDK), "'", "''") & "'"
If Trim(newY.SSIDIVINFO) <> "" Then xSet = xSet & ",SSIDIVINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDIVINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIDIVH" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIDIVH_Insert = "Erreur màj : " & newY.SSIDIVPRFK
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIDIVH_Insert = "sqlYSSIDIVH_Insert " & vbCrLf & Error
End Function


Public Function sqlYSSIDIV0_Update(newY As typeYSSIDIV0, oldY As typeYSSIDIV0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIDIV0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIDIVUIDX <> newY.SSIDIVUIDX Then
    sqlYSSIDIV0_Update = "Erreur SSIDIVUIDX : " & newY.SSIDIVUIDX & " / " & oldY.SSIDIVUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIDIVNAT = '" & oldY.SSIDIVNAT & "'" _
       & " and SSIDIVUIDX = '" & oldY.SSIDIVUIDX & "'" _
       & " and SSIDIVUIDD = " & oldY.SSIDIVUIDD _
       & " and SSIDIVYVER = " & oldY.SSIDIVYVER
       
newY.SSIDIVYVER = newY.SSIDIVYVER + 1
xSet = xSet & " set SSIDIVYVER = " & newY.SSIDIVYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIDIVTLNK <> oldY.SSIDIVTLNK Then blnUpdate = True: xSet = xSet & " , SSIDIVTLNK = " & newY.SSIDIVTLNK
If newY.SSIDIVYAMJ <> oldY.SSIDIVYAMJ Then blnUpdate = True: xSet = xSet & " , SSIDIVYAMJ = " & newY.SSIDIVYAMJ
If newY.SSIDIVYHMS <> oldY.SSIDIVYHMS Then blnUpdate = True: xSet = xSet & " , SSIDIVYHMS = " & newY.SSIDIVYHMS

If newY.SSIDIVSTAK <> oldY.SSIDIVSTAK Then blnUpdate = True:  xSet = xSet & " , SSIDIVSTAK = '" & Replace(Trim(newY.SSIDIVSTAK), "'", "''") & "'"
If newY.SSIDIVPRFX <> oldY.SSIDIVPRFX Then blnUpdate = True:  xSet = xSet & " , SSIDIVPRFX = '" & Replace(Trim(newY.SSIDIVPRFX), "'", "''") & "'"
If newY.SSIDIVPRFK <> oldY.SSIDIVPRFK Then blnUpdate = True:  xSet = xSet & " , SSIDIVPRFK = '" & Replace(Trim(newY.SSIDIVPRFK), "'", "''") & "'"
If newY.SSIDIVUNOM <> oldY.SSIDIVUNOM Then blnUpdate = True:  xSet = xSet & " , SSIDIVUNOM = '" & Replace(Trim(newY.SSIDIVUNOM), "'", "''") & "'"
If newY.SSIDIVYFCT <> oldY.SSIDIVYFCT Then blnUpdate = True:  xSet = xSet & " , SSIDIVYFCT = '" & Replace(Trim(newY.SSIDIVYFCT), "'", "''") & "'"
If newY.SSIDIVYUSR <> oldY.SSIDIVYUSR Then blnUpdate = True:  xSet = xSet & " , SSIDIVYUSR = '" & Replace(Trim(newY.SSIDIVYUSR), "'", "''") & "'"
If newY.SSIDIVDIDK <> oldY.SSIDIVDIDK Then blnUpdate = True:  xSet = xSet & " , SSIDIVDIDK= '" & Replace(Trim(newY.SSIDIVDIDK), "'", "''") & "'"
If newY.SSIDIVINFO <> oldY.SSIDIVINFO Then blnUpdate = True:  xSet = xSet & " , SSIDIVINFO= '" & Replace(Trim(newY.SSIDIVINFO), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIDIV0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIDIV0_Update = "Erreur màj : " & newY.SSIDIVDIDK
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIDIV0_Update = "sqlYSSIDIV0_Update " & vbCrLf & Error
End Function
Public Function sqlYSSIDIVH_Update(newY As typeYSSIDIV0, oldY As typeYSSIDIV0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIDIVH_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIDIVUIDX <> newY.SSIDIVUIDX Then
    sqlYSSIDIVH_Update = "Erreur SSIDIVUIDX : " & newY.SSIDIVUIDX & " / " & oldY.SSIDIVUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIDIVNAT = '" & oldY.SSIDIVNAT & "'" _
       & " and SSIDIVUIDX = '" & oldY.SSIDIVUIDX & "'" _
       & " and SSIDIVUIDD = " & oldY.SSIDIVUIDD _
       & " and SSIDIVYVER = " & oldY.SSIDIVYVER
       
xSet = xSet & " set SSIDIVYVER = " & newY.SSIDIVYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIDIVTLNK <> oldY.SSIDIVTLNK Then blnUpdate = True: xSet = xSet & " , SSIDIVTLNK = " & newY.SSIDIVTLNK
If newY.SSIDIVYAMJ <> oldY.SSIDIVYAMJ Then blnUpdate = True: xSet = xSet & " , SSIDIVYAMJ = " & newY.SSIDIVYAMJ
If newY.SSIDIVYHMS <> oldY.SSIDIVYHMS Then blnUpdate = True: xSet = xSet & " , SSIDIVYHMS = " & newY.SSIDIVYHMS

If newY.SSIDIVSTAK <> oldY.SSIDIVSTAK Then blnUpdate = True:  xSet = xSet & " , SSIDIVSTAK = '" & Replace(Trim(newY.SSIDIVSTAK), "'", "''") & "'"
If newY.SSIDIVPRFX <> oldY.SSIDIVPRFX Then blnUpdate = True:  xSet = xSet & " , SSIDIVPRFX = '" & Replace(Trim(newY.SSIDIVPRFX), "'", "''") & "'"
If newY.SSIDIVPRFK <> oldY.SSIDIVPRFK Then blnUpdate = True:  xSet = xSet & " , SSIDIVPRFK = '" & Replace(Trim(newY.SSIDIVPRFK), "'", "''") & "'"
If newY.SSIDIVUNOM <> oldY.SSIDIVUNOM Then blnUpdate = True:  xSet = xSet & " , SSIDIVUNOM = '" & Replace(Trim(newY.SSIDIVUNOM), "'", "''") & "'"
If newY.SSIDIVYFCT <> oldY.SSIDIVYFCT Then blnUpdate = True:  xSet = xSet & " , SSIDIVYFCT = '" & Replace(Trim(newY.SSIDIVYFCT), "'", "''") & "'"
If newY.SSIDIVYUSR <> oldY.SSIDIVYUSR Then blnUpdate = True:  xSet = xSet & " , SSIDIVYUSR = '" & Replace(Trim(newY.SSIDIVYUSR), "'", "''") & "'"
If newY.SSIDIVDIDK <> oldY.SSIDIVDIDK Then blnUpdate = True:  xSet = xSet & " , SSIDIVDIDK= '" & Replace(Trim(newY.SSIDIVDIDK), "'", "''") & "'"
If newY.SSIDIVINFO <> oldY.SSIDIVINFO Then blnUpdate = True:  xSet = xSet & " , SSIDIVINFO= '" & Replace(Trim(newY.SSIDIVINFO), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIDIVH" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIDIVH_Update = "Erreur màj : " & newY.SSIDIVDIDK
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIDIVH_Update = "sqlYSSIDIVH_Update " & vbCrLf & Error
End Function


Public Function rsYSSIDIV0_GetBuffer(rsAdo As ADODB.Recordset, lYSSIDIV0 As typeYSSIDIV0)
On Error GoTo Error_Handler
rsYSSIDIV0_GetBuffer = Null
lYSSIDIV0.SSIDIVNAT = rsAdo("SSIDIVNAT")
lYSSIDIV0.SSIDIVUIDX = Trim(rsAdo("SSIDIVUIDX"))
lYSSIDIV0.SSIDIVDIDK = Trim(rsAdo("SSIDIVDIDK"))

lYSSIDIV0.SSIDIVSTAK = rsAdo("SSIDIVSTAK")
lYSSIDIV0.SSIDIVUIDD = rsAdo("SSIDIVUIDD")
lYSSIDIV0.SSIDIVPRFX = Trim(rsAdo("SSIDIVPRFX"))
lYSSIDIV0.SSIDIVPRFK = rsAdo("SSIDIVPRFK")
lYSSIDIV0.SSIDIVUNOM = Trim(rsAdo("SSIDIVUNOM"))

lYSSIDIV0.SSIDIVTLNK = rsAdo("SSIDIVTLNK")
lYSSIDIV0.SSIDIVYFCT = Trim(rsAdo("SSIDIVYFCT"))
lYSSIDIV0.SSIDIVYUSR = Trim(rsAdo("SSIDIVYUSR"))
lYSSIDIV0.SSIDIVYAMJ = rsAdo("SSIDIVYAMJ")
lYSSIDIV0.SSIDIVYHMS = rsAdo("SSIDIVYHMS")
lYSSIDIV0.SSIDIVYVER = rsAdo("SSIDIVYVER")

lYSSIDIV0.SSIDIVINFO = Trim(rsAdo("SSIDIVINFO"))

Exit Function
Error_Handler:
rsYSSIDIV0_GetBuffer = Error


End Function

Public Function rsYSSIDIV0_Init(lYSSIDIV0 As typeYSSIDIV0)
lYSSIDIV0.SSIDIVNAT = ""
lYSSIDIV0.SSIDIVUIDX = ""
lYSSIDIV0.SSIDIVDIDK = ""

lYSSIDIV0.SSIDIVSTAK = ""
lYSSIDIV0.SSIDIVUIDD = 0
lYSSIDIV0.SSIDIVPRFX = ""
lYSSIDIV0.SSIDIVPRFK = ""
lYSSIDIV0.SSIDIVUNOM = ""

lYSSIDIV0.SSIDIVTLNK = 0
lYSSIDIV0.SSIDIVYFCT = ""
lYSSIDIV0.SSIDIVYUSR = ""
lYSSIDIV0.SSIDIVYAMJ = 0
lYSSIDIV0.SSIDIVYHMS = 0
lYSSIDIV0.SSIDIVYVER = 0
lYSSIDIV0.SSIDIVINFO = ""
End Function













