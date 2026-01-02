Attribute VB_Name = "srvYSSIMEL0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSIMEL0
    SSIMELNAT    As String
    SSIMELUIDX   As String
       
    SSIMELSTAK   As String
    SSIMELUIDD   As Long
    SSIMELPRFX   As String
    SSIMELPRFK   As String
    SSIMELUNOM    As String
    
    SSIMELTLNK   As Long
    SSIMELYFCT   As String
    SSIMELYUSR   As String
    SSIMELYAMJ   As Long
    SSIMELYHMS   As Long
    SSIMELYVER   As Long
    
    SSIMELINFO   As String
End Type


'Public xYSSIMEL0 As typeYSSIMEL0
Public Function sqlYSSIMEL0_Insert(newY As typeYSSIMEL0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIMEL0_Insert = Null

xSet = " (SSIMELNAT,SSIMELUIDX,SSIMELUIDD"
xValues = " values('" & newY.SSIMELNAT & "','" & newY.SSIMELUIDX & "'," & newY.SSIMELUIDD

' Détecter les modifications
'===================================================================================
If newY.SSIMELTLNK <> 0 Then xSet = xSet & ",SSIMELTLNK": xValues = xValues & " ," & newY.SSIMELTLNK
If newY.SSIMELYVER <> 0 Then xSet = xSet & ",SSIMELYVER": xValues = xValues & " ," & newY.SSIMELYVER
If newY.SSIMELYAMJ <> 0 Then xSet = xSet & ",SSIMELYAMJ": xValues = xValues & " ," & newY.SSIMELYAMJ
If newY.SSIMELYHMS <> 0 Then xSet = xSet & ",SSIMELYHMS": xValues = xValues & " ," & newY.SSIMELYHMS

If Trim(newY.SSIMELSTAK) <> "" Then xSet = xSet & ",SSIMELSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELSTAK), "'", "''") & "'"
If Trim(newY.SSIMELPRFX) <> "" Then xSet = xSet & ",SSIMELPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELPRFX), "'", "''") & "'"
If Trim(newY.SSIMELPRFK) <> "" Then xSet = xSet & ",SSIMELPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELPRFK), "'", "''") & "'"
If Trim(newY.SSIMELUNOM) <> "" Then xSet = xSet & ",SSIMELUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELUNOM), "'", "''") & "'"
If Trim(newY.SSIMELYFCT) <> "" Then xSet = xSet & ",SSIMELYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELYFCT), "'", "''") & "'"
If Trim(newY.SSIMELYUSR) <> "" Then xSet = xSet & ",SSIMELYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELYUSR), "'", "''") & "'"
If Trim(newY.SSIMELINFO) <> "" Then xSet = xSet & ",SSIMELINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIMEL0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIMEL0_Insert = "Erreur màj : " & newY.SSIMELNAT & " " & newY.SSIMELUIDX
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIMEL0_Insert = "sqlYSSIMEL0_Insert " & newY.SSIMELNAT & " " & newY.SSIMELUIDX & vbCrLf & Error
End Function
Public Function sqlYSSIMELH_Insert(newY As typeYSSIMEL0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIMELH_Insert = Null

xSet = " (SSIMELNAT,SSIMELUIDX,SSIMELUIDD"
xValues = " values('" & newY.SSIMELNAT & "','" & newY.SSIMELUIDX & "'," & newY.SSIMELUIDD

' Détecter les modifications
'===================================================================================
If newY.SSIMELTLNK <> 0 Then xSet = xSet & ",SSIMELTLNK": xValues = xValues & " ," & newY.SSIMELTLNK
If newY.SSIMELYVER <> 0 Then xSet = xSet & ",SSIMELYVER": xValues = xValues & " ," & newY.SSIMELYVER
If newY.SSIMELYAMJ <> 0 Then xSet = xSet & ",SSIMELYAMJ": xValues = xValues & " ," & newY.SSIMELYAMJ
If newY.SSIMELYHMS <> 0 Then xSet = xSet & ",SSIMELYHMS": xValues = xValues & " ," & newY.SSIMELYHMS

If Trim(newY.SSIMELSTAK) <> "" Then xSet = xSet & ",SSIMELSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELSTAK), "'", "''") & "'"
If Trim(newY.SSIMELPRFX) <> "" Then xSet = xSet & ",SSIMELPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELPRFX), "'", "''") & "'"
If Trim(newY.SSIMELPRFK) <> "" Then xSet = xSet & ",SSIMELPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELPRFK), "'", "''") & "'"
If Trim(newY.SSIMELUNOM) <> "" Then xSet = xSet & ",SSIMELUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELUNOM), "'", "''") & "'"
If Trim(newY.SSIMELYFCT) <> "" Then xSet = xSet & ",SSIMELYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELYFCT), "'", "''") & "'"
If Trim(newY.SSIMELYUSR) <> "" Then xSet = xSet & ",SSIMELYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELYUSR), "'", "''") & "'"
If Trim(newY.SSIMELINFO) <> "" Then xSet = xSet & ",SSIMELINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSIMELINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIMELH" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIMELH_Insert = "Erreur màj : " & newY.SSIMELNAT & " " & newY.SSIMELUIDX & " " & newY.SSIMELYVER
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIMELH_Insert = "sqlYSSIMELH_Insert " & newY.SSIMELUIDX & " " & newY.SSIMELYVER & vbCrLf & Error
End Function


Public Function sqlYSSIMEL0_Update(newY As typeYSSIMEL0, oldY As typeYSSIMEL0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIMEL0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIMELUIDX <> newY.SSIMELUIDX Then
    sqlYSSIMEL0_Update = "Erreur SSIMELUIDX : " & newY.SSIMELUIDX & " / " & oldY.SSIMELUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIMELNAT = '" & oldY.SSIMELNAT & "'" _
       & " and SSIMELUIDX = '" & oldY.SSIMELUIDX & "'" _
       & " and SSIMELUIDD = " & oldY.SSIMELUIDD _
       & " and SSIMELYVER = " & oldY.SSIMELYVER
       
newY.SSIMELYVER = newY.SSIMELYVER + 1
xSet = xSet & " set SSIMELYVER = " & newY.SSIMELYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIMELTLNK <> oldY.SSIMELTLNK Then blnUpdate = True: xSet = xSet & " , SSIMELTLNK = " & newY.SSIMELTLNK
If newY.SSIMELYAMJ <> oldY.SSIMELYAMJ Then blnUpdate = True: xSet = xSet & " , SSIMELYAMJ = " & newY.SSIMELYAMJ
If newY.SSIMELYHMS <> oldY.SSIMELYHMS Then blnUpdate = True: xSet = xSet & " , SSIMELYHMS = " & newY.SSIMELYHMS

If newY.SSIMELSTAK <> oldY.SSIMELSTAK Then blnUpdate = True:  xSet = xSet & " , SSIMELSTAK = '" & Replace(Trim(newY.SSIMELSTAK), "'", "''") & "'"
If newY.SSIMELPRFX <> oldY.SSIMELPRFX Then blnUpdate = True:  xSet = xSet & " , SSIMELPRFX = '" & Replace(Trim(newY.SSIMELPRFX), "'", "''") & "'"
If newY.SSIMELPRFK <> oldY.SSIMELPRFK Then blnUpdate = True:  xSet = xSet & " , SSIMELPRFK = '" & Replace(Trim(newY.SSIMELPRFK), "'", "''") & "'"
If newY.SSIMELUNOM <> oldY.SSIMELUNOM Then blnUpdate = True:  xSet = xSet & " , SSIMELUNOM = '" & Replace(Trim(newY.SSIMELUNOM), "'", "''") & "'"
If newY.SSIMELYFCT <> oldY.SSIMELYFCT Then blnUpdate = True:  xSet = xSet & " , SSIMELYFCT = '" & Replace(Trim(newY.SSIMELYFCT), "'", "''") & "'"
If newY.SSIMELYUSR <> oldY.SSIMELYUSR Then blnUpdate = True:  xSet = xSet & " , SSIMELYUSR = '" & Replace(Trim(newY.SSIMELYUSR), "'", "''") & "'"
If newY.SSIMELINFO <> oldY.SSIMELINFO Then blnUpdate = True:  xSet = xSet & " , SSIMELINFO= '" & Replace(Trim(newY.SSIMELINFO), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIMEL0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIMEL0_Update = "Erreur màj : " & newY.SSIMELNAT & " " & newY.SSIMELUIDX
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIMEL0_Update = "sqlYSSIMEL0_Update " & vbCrLf & Error
End Function
Public Function sqlYSSIMELH_Update(newY As typeYSSIMEL0, oldY As typeYSSIMEL0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIMELH_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIMELUIDX <> newY.SSIMELUIDX Then
    sqlYSSIMELH_Update = "Erreur SSIMELUIDX : " & newY.SSIMELUIDX & " / " & oldY.SSIMELUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIMELNAT = '" & oldY.SSIMELNAT & "'" _
       & " and SSIMELUIDX = '" & oldY.SSIMELUIDX & "'" _
       & " and SSIMELUIDD = " & oldY.SSIMELUIDD _
       & " and SSIMELYVER = " & oldY.SSIMELYVER
       
xSet = xSet & " set SSIMELYVER = " & newY.SSIMELYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIMELTLNK <> oldY.SSIMELTLNK Then blnUpdate = True: xSet = xSet & " , SSIMELTLNK = " & newY.SSIMELTLNK
If newY.SSIMELYAMJ <> oldY.SSIMELYAMJ Then blnUpdate = True: xSet = xSet & " , SSIMELYAMJ = " & newY.SSIMELYAMJ
If newY.SSIMELYHMS <> oldY.SSIMELYHMS Then blnUpdate = True: xSet = xSet & " , SSIMELYHMS = " & newY.SSIMELYHMS

If newY.SSIMELSTAK <> oldY.SSIMELSTAK Then blnUpdate = True:  xSet = xSet & " , SSIMELSTAK = '" & Replace(Trim(newY.SSIMELSTAK), "'", "''") & "'"
If newY.SSIMELPRFX <> oldY.SSIMELPRFX Then blnUpdate = True:  xSet = xSet & " , SSIMELPRFX = '" & Replace(Trim(newY.SSIMELPRFX), "'", "''") & "'"
If newY.SSIMELPRFK <> oldY.SSIMELPRFK Then blnUpdate = True:  xSet = xSet & " , SSIMELPRFK = '" & Replace(Trim(newY.SSIMELPRFK), "'", "''") & "'"
If newY.SSIMELUNOM <> oldY.SSIMELUNOM Then blnUpdate = True:  xSet = xSet & " , SSIMELUNOM = '" & Replace(Trim(newY.SSIMELUNOM), "'", "''") & "'"
If newY.SSIMELYFCT <> oldY.SSIMELYFCT Then blnUpdate = True:  xSet = xSet & " , SSIMELYFCT = '" & Replace(Trim(newY.SSIMELYFCT), "'", "''") & "'"
If newY.SSIMELYUSR <> oldY.SSIMELYUSR Then blnUpdate = True:  xSet = xSet & " , SSIMELYUSR = '" & Replace(Trim(newY.SSIMELYUSR), "'", "''") & "'"
If newY.SSIMELINFO <> oldY.SSIMELINFO Then blnUpdate = True:  xSet = xSet & " , SSIMELINFO= '" & Replace(Trim(newY.SSIMELINFO), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIMELH" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
       sqlYSSIMELH_Update = "Erreur màj : " & newY.SSIMELNAT & " " & newY.SSIMELUIDX & " " & newY.SSIMELYVER
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIMELH_Update = "sqlYSSIMELH_Update " & vbCrLf & Error
End Function


Public Function rsYSSIMEL0_GetBuffer(rsAdo As ADODB.Recordset, lYSSIMEL0 As typeYSSIMEL0)
On Error GoTo Error_Handler
rsYSSIMEL0_GetBuffer = Null
lYSSIMEL0.SSIMELNAT = rsAdo("SSIMELNAT")
lYSSIMEL0.SSIMELUIDX = Trim(rsAdo("SSIMELUIDX"))

lYSSIMEL0.SSIMELSTAK = rsAdo("SSIMELSTAK")
lYSSIMEL0.SSIMELUIDD = rsAdo("SSIMELUIDD")
lYSSIMEL0.SSIMELPRFX = Trim(rsAdo("SSIMELPRFX"))
lYSSIMEL0.SSIMELPRFK = rsAdo("SSIMELPRFK")
lYSSIMEL0.SSIMELUNOM = Trim(rsAdo("SSIMELUNOM"))

lYSSIMEL0.SSIMELTLNK = rsAdo("SSIMELTLNK")
lYSSIMEL0.SSIMELYFCT = Trim(rsAdo("SSIMELYFCT"))
lYSSIMEL0.SSIMELYUSR = Trim(rsAdo("SSIMELYUSR"))
lYSSIMEL0.SSIMELYAMJ = rsAdo("SSIMELYAMJ")
lYSSIMEL0.SSIMELYHMS = rsAdo("SSIMELYHMS")
lYSSIMEL0.SSIMELYVER = rsAdo("SSIMELYVER")

lYSSIMEL0.SSIMELINFO = Trim(rsAdo("SSIMELINFO"))

Exit Function
Error_Handler:
rsYSSIMEL0_GetBuffer = Error


End Function

Public Function rsYSSIMEL0_Init(lYSSIMEL0 As typeYSSIMEL0)
lYSSIMEL0.SSIMELNAT = ""
lYSSIMEL0.SSIMELUIDX = ""

lYSSIMEL0.SSIMELSTAK = ""
lYSSIMEL0.SSIMELUIDD = 0
lYSSIMEL0.SSIMELPRFX = ""
lYSSIMEL0.SSIMELPRFK = ""
lYSSIMEL0.SSIMELUNOM = ""

lYSSIMEL0.SSIMELTLNK = 0
lYSSIMEL0.SSIMELYFCT = ""
lYSSIMEL0.SSIMELYUSR = ""
lYSSIMEL0.SSIMELYAMJ = 0
lYSSIMEL0.SSIMELYHMS = 0
lYSSIMEL0.SSIMELYVER = 0
lYSSIMEL0.SSIMELINFO = ""
End Function
