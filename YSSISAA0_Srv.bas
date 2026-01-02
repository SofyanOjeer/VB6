Attribute VB_Name = "srvYSSISAA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSISAA0
    SSISAANAT    As String
    SSISAAUIDX   As String
    SSISAAUSEQ   As Long
    
    SSISAASTAK   As String
    SSISAAUIDD   As Long
    SSISAAPRFX   As String
    SSISAAPRFK   As String
    SSISAAUNOM    As String
    
    SSISAATLNK   As Long
    SSISAAYFCT   As String
    SSISAAYUSR   As String
    SSISAAYAMJ   As Long
    SSISAAYHMS   As Long
    SSISAAYVER   As Long
    SSISAAINFO   As String
End Type
'Public xYSSISAA0 As typeYSSISAA0
Public Function sqlYSSISAA0_Insert(newY As typeYSSISAA0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSISAA0_Insert = Null

xSet = " (SSISAANAT,SSISAAUIDX,SSISAAUSEQ"
xValues = " values('" & newY.SSISAANAT & "','" & newY.SSISAAUIDX & "'," & newY.SSISAAUSEQ

' Détecter les modifications
'===================================================================================
If newY.SSISAAUIDD <> 0 Then xSet = xSet & ",SSISAAUIDD": xValues = xValues & " ," & newY.SSISAAUIDD
If newY.SSISAATLNK <> 0 Then xSet = xSet & ",SSISAATLNK": xValues = xValues & " ," & newY.SSISAATLNK
If newY.SSISAAYVER <> 0 Then xSet = xSet & ",SSISAAYVER": xValues = xValues & " ," & newY.SSISAAYVER
If newY.SSISAAYAMJ <> 0 Then xSet = xSet & ",SSISAAYAMJ": xValues = xValues & " ," & newY.SSISAAYAMJ
If newY.SSISAAYHMS <> 0 Then xSet = xSet & ",SSISAAYHMS": xValues = xValues & " ," & newY.SSISAAYHMS

If Trim(newY.SSISAASTAK) <> "" Then xSet = xSet & ",SSISAASTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAASTAK), "'", "''") & "'"
If Trim(newY.SSISAAPRFX) <> "" Then xSet = xSet & ",SSISAAPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAPRFX), "'", "''") & "'"
If Trim(newY.SSISAAPRFK) <> "" Then xSet = xSet & ",SSISAAPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAPRFK), "'", "''") & "'"
If Trim(newY.SSISAAUNOM) <> "" Then xSet = xSet & ",SSISAAUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAUNOM), "'", "''") & "'"
If Trim(newY.SSISAAYFCT) <> "" Then xSet = xSet & ",SSISAAYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAYFCT), "'", "''") & "'"
If Trim(newY.SSISAAYUSR) <> "" Then xSet = xSet & ",SSISAAYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAYUSR), "'", "''") & "'"
If Trim(newY.SSISAAINFO) <> "" Then xSet = xSet & ",SSISAAINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSISAA0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSISAA0_Insert = "Erreur màj : " & newY.SSISAAPRFK
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSISAA0_Insert = "sqlYSSISAA0_Insert " & vbCrLf & Error
End Function
Public Function sqlYSSISAAH_Insert(newY As typeYSSISAA0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSISAAH_Insert = Null

xSet = " (SSISAANAT,SSISAAUIDX,SSISAAUSEQ"
xValues = " values('" & newY.SSISAANAT & "','" & newY.SSISAAUIDX & "'," & newY.SSISAAUSEQ

' Détecter les modifications
'===================================================================================
If newY.SSISAAUIDD <> 0 Then xSet = xSet & ",SSISAAUIDD": xValues = xValues & " ," & newY.SSISAAUIDD
If newY.SSISAATLNK <> 0 Then xSet = xSet & ",SSISAATLNK": xValues = xValues & " ," & newY.SSISAATLNK
If newY.SSISAAYVER <> 0 Then xSet = xSet & ",SSISAAYVER": xValues = xValues & " ," & newY.SSISAAYVER
If newY.SSISAAYAMJ <> 0 Then xSet = xSet & ",SSISAAYAMJ": xValues = xValues & " ," & newY.SSISAAYAMJ
If newY.SSISAAYHMS <> 0 Then xSet = xSet & ",SSISAAYHMS": xValues = xValues & " ," & newY.SSISAAYHMS

If Trim(newY.SSISAASTAK) <> "" Then xSet = xSet & ",SSISAASTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAASTAK), "'", "''") & "'"
If Trim(newY.SSISAAPRFX) <> "" Then xSet = xSet & ",SSISAAPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAPRFX), "'", "''") & "'"
If Trim(newY.SSISAAPRFK) <> "" Then xSet = xSet & ",SSISAAPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAPRFK), "'", "''") & "'"
If Trim(newY.SSISAAUNOM) <> "" Then xSet = xSet & ",SSISAAUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAUNOM), "'", "''") & "'"
If Trim(newY.SSISAAYFCT) <> "" Then xSet = xSet & ",SSISAAYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAYFCT), "'", "''") & "'"
If Trim(newY.SSISAAYUSR) <> "" Then xSet = xSet & ",SSISAAYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAYUSR), "'", "''") & "'"
If Trim(newY.SSISAAINFO) <> "" Then xSet = xSet & ",SSISAAINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSISAAINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSISAAH" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSISAAH_Insert = "Erreur màj : " & newY.SSISAAPRFK
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSISAAH_Insert = "sqlYSSISAAH_Insert " & vbCrLf & Error
End Function


Public Function sqlYSSISAA0_Update(newY As typeYSSISAA0, oldY As typeYSSISAA0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSISAA0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSISAAUIDX <> newY.SSISAAUIDX Then
    sqlYSSISAA0_Update = "Erreur SSISAAUIDX : " & newY.SSISAAUIDX & " / " & oldY.SSISAAUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSISAANAT = '" & oldY.SSISAANAT & "'" _
       & " and SSISAAUIDX = '" & oldY.SSISAAUIDX & "'" _
       & " and SSISAAUSEQ = " & oldY.SSISAAUSEQ _
       & " and SSISAAYVER = " & oldY.SSISAAYVER
       
newY.SSISAAYVER = newY.SSISAAYVER + 1
xSet = xSet & " set SSISAAYVER = " & newY.SSISAAYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSISAAUIDD <> oldY.SSISAAUIDD Then blnUpdate = True: xSet = xSet & " , SSISAAUIDD = " & newY.SSISAAUIDD
If newY.SSISAATLNK <> oldY.SSISAATLNK Then blnUpdate = True: xSet = xSet & " , SSISAATLNK = " & newY.SSISAATLNK
If newY.SSISAAYAMJ <> oldY.SSISAAYAMJ Then blnUpdate = True: xSet = xSet & " , SSISAAYAMJ = " & newY.SSISAAYAMJ
If newY.SSISAAYHMS <> oldY.SSISAAYHMS Then blnUpdate = True: xSet = xSet & " , SSISAAYHMS = " & newY.SSISAAYHMS

If newY.SSISAASTAK <> oldY.SSISAASTAK Then blnUpdate = True:  xSet = xSet & " , SSISAASTAK = '" & Replace(Trim(newY.SSISAASTAK), "'", "''") & "'"
If newY.SSISAAPRFX <> oldY.SSISAAPRFX Then blnUpdate = True:  xSet = xSet & " , SSISAAPRFX = '" & Replace(Trim(newY.SSISAAPRFX), "'", "''") & "'"
If newY.SSISAAPRFK <> oldY.SSISAAPRFK Then blnUpdate = True:  xSet = xSet & " , SSISAAPRFK = '" & Replace(Trim(newY.SSISAAPRFK), "'", "''") & "'"
If newY.SSISAAUNOM <> oldY.SSISAAUNOM Then blnUpdate = True:  xSet = xSet & " , SSISAAUNOM = '" & Replace(Trim(newY.SSISAAUNOM), "'", "''") & "'"
If newY.SSISAAYFCT <> oldY.SSISAAYFCT Then blnUpdate = True:  xSet = xSet & " , SSISAAYFCT = '" & Replace(Trim(newY.SSISAAYFCT), "'", "''") & "'"
If newY.SSISAAINFO <> oldY.SSISAAINFO Then blnUpdate = True:  xSet = xSet & " , SSISAAINFO= '" & Replace(Trim(newY.SSISAAINFO), "'", "''") & "'"
If newY.SSISAAYUSR <> oldY.SSISAAYUSR Then blnUpdate = True:  xSet = xSet & " , SSISAAYUSR = '" & Replace(Trim(newY.SSISAAYUSR), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSISAA0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSISAA0_Update = "Erreur màj : " & newY.SSISAAUIDX
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSISAA0_Update = "sqlYSSISAA0_Update " & vbCrLf & Error
End Function
Public Function sqlYSSISAAH_Update(newY As typeYSSISAA0, oldY As typeYSSISAA0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSISAAH_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSISAAUIDX <> newY.SSISAAUIDX Then
    sqlYSSISAAH_Update = "Erreur SSISAAUIDX : " & newY.SSISAAUIDX & " / " & oldY.SSISAAUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSISAANAT = '" & oldY.SSISAANAT & "'" _
       & " and SSISAAUIDX = '" & oldY.SSISAAUIDX & "'" _
       & " and SSISAAUSEQ = " & oldY.SSISAAUSEQ _
       & " and SSISAAYVER = " & oldY.SSISAAYVER
       
xSet = xSet & " set SSISAAYVER = " & newY.SSISAAYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSISAAUIDD <> oldY.SSISAAUIDD Then blnUpdate = True: xSet = xSet & " , SSISAAUIDD = " & newY.SSISAAUIDD
If newY.SSISAATLNK <> oldY.SSISAATLNK Then blnUpdate = True: xSet = xSet & " , SSISAATLNK = " & newY.SSISAATLNK
If newY.SSISAAYAMJ <> oldY.SSISAAYAMJ Then blnUpdate = True: xSet = xSet & " , SSISAAYAMJ = " & newY.SSISAAYAMJ
If newY.SSISAAYHMS <> oldY.SSISAAYHMS Then blnUpdate = True: xSet = xSet & " , SSISAAYHMS = " & newY.SSISAAYHMS

If newY.SSISAASTAK <> oldY.SSISAASTAK Then blnUpdate = True:  xSet = xSet & " , SSISAASTAK = '" & Replace(Trim(newY.SSISAASTAK), "'", "''") & "'"
If newY.SSISAAPRFX <> oldY.SSISAAPRFX Then blnUpdate = True:  xSet = xSet & " , SSISAAPRFX = '" & Replace(Trim(newY.SSISAAPRFX), "'", "''") & "'"
If newY.SSISAAPRFK <> oldY.SSISAAPRFK Then blnUpdate = True:  xSet = xSet & " , SSISAAPRFK = '" & Replace(Trim(newY.SSISAAPRFK), "'", "''") & "'"
If newY.SSISAAUNOM <> oldY.SSISAAUNOM Then blnUpdate = True:  xSet = xSet & " , SSISAAUNOM = '" & Replace(Trim(newY.SSISAAUNOM), "'", "''") & "'"
If newY.SSISAAYFCT <> oldY.SSISAAYFCT Then blnUpdate = True:  xSet = xSet & " , SSISAAYFCT = '" & Replace(Trim(newY.SSISAAYFCT), "'", "''") & "'"
If newY.SSISAAINFO <> oldY.SSISAAINFO Then blnUpdate = True:  xSet = xSet & " , SSISAAINFO= '" & Replace(Trim(newY.SSISAAINFO), "'", "''") & "'"
If newY.SSISAAYUSR <> oldY.SSISAAYUSR Then blnUpdate = True:  xSet = xSet & " , SSISAAYUSR = '" & Replace(Trim(newY.SSISAAYUSR), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSISAAH" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSISAAH_Update = "Erreur màj : " & newY.SSISAAUIDX
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSISAAH_Update = "sqlYSSISAAH_Update " & vbCrLf & Error
End Function


Public Function rsYSSISAA0_GetBuffer(rsAdo As ADODB.Recordset, lYSSISAA0 As typeYSSISAA0)
On Error GoTo Error_Handler
rsYSSISAA0_GetBuffer = Null
lYSSISAA0.SSISAANAT = rsAdo("SSISAANAT")
lYSSISAA0.SSISAAUIDX = Trim(rsAdo("SSISAAUIDX"))
lYSSISAA0.SSISAAUSEQ = rsAdo("SSISAAUSEQ")

lYSSISAA0.SSISAASTAK = rsAdo("SSISAASTAK")
lYSSISAA0.SSISAAUIDD = rsAdo("SSISAAUIDD")
lYSSISAA0.SSISAAPRFX = Trim(rsAdo("SSISAAPRFX"))
lYSSISAA0.SSISAAPRFK = rsAdo("SSISAAPRFK")
lYSSISAA0.SSISAAUNOM = Trim(rsAdo("SSISAAUNOM"))

lYSSISAA0.SSISAATLNK = rsAdo("SSISAATLNK")
lYSSISAA0.SSISAAYFCT = Trim(rsAdo("SSISAAYFCT"))
lYSSISAA0.SSISAAYUSR = Trim(rsAdo("SSISAAYUSR"))
lYSSISAA0.SSISAAYAMJ = rsAdo("SSISAAYAMJ")

lYSSISAA0.SSISAAYHMS = rsAdo("SSISAAYHMS")
lYSSISAA0.SSISAAYVER = rsAdo("SSISAAYVER")
lYSSISAA0.SSISAAINFO = Trim(rsAdo("SSISAAINFO"))

Exit Function
Error_Handler:
rsYSSISAA0_GetBuffer = Error


End Function

Public Function rsYSSISAA0_Init(lYSSISAA0 As typeYSSISAA0)
lYSSISAA0.SSISAANAT = ""
lYSSISAA0.SSISAAUIDX = ""
lYSSISAA0.SSISAAUSEQ = 0

lYSSISAA0.SSISAASTAK = ""
lYSSISAA0.SSISAAUIDD = 0
lYSSISAA0.SSISAAPRFX = ""
lYSSISAA0.SSISAAPRFK = ""
lYSSISAA0.SSISAAUNOM = ""

lYSSISAA0.SSISAATLNK = 0
lYSSISAA0.SSISAAYFCT = ""
lYSSISAA0.SSISAAYUSR = ""
lYSSISAA0.SSISAAYAMJ = 0
lYSSISAA0.SSISAAYHMS = 0
lYSSISAA0.SSISAAYVER = 0
lYSSISAA0.SSISAAINFO = ""
End Function










