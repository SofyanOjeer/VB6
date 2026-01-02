Attribute VB_Name = "srvYSSISAB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSISAB0
    SSISABNAT    As String
    SSISABUIDX   As String
    SSISABULOT   As Long
    
    SSISABSTAK   As String
    SSISABUIDD   As Long
    SSISABPRFX   As String
    SSISABPRFK   As String
    SSISABUNOM    As String
    
    SSISABTLNK   As Long
    SSISABYFCT   As String
    SSISABYUSR   As String
    SSISABYAMJ   As Long
    SSISABYHMS   As Long
    SSISABYVER   As Long
    SSISABINFO   As String
End Type

Type typeWSSIMNU0
    SSIMNUCOD    As Long
    SSIMNULIB    As String
    SSIMNUENS    As String
    
    SSIMNUPRE    As Long
    SSIMNUORD    As Long
    SSIMNUARB    As String
    SSIMNUINFO   As String
End Type

'Public xYSSISAB0 As typeYSSISAB0
Public Function sqlYSSISAB0_Insert(newY As typeYSSISAB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSISAB0_Insert = Null

xSet = " (SSISABNAT,SSISABUIDX,SSISABULOT"
xValues = " values('" & newY.SSISABNAT & "','" & newY.SSISABUIDX & "'," & newY.SSISABULOT

' Détecter les modifications
'===================================================================================
If newY.SSISABUIDD <> 0 Then xSet = xSet & ",SSISABUIDD": xValues = xValues & " ," & newY.SSISABUIDD
If newY.SSISABTLNK <> 0 Then xSet = xSet & ",SSISABTLNK": xValues = xValues & " ," & newY.SSISABTLNK
If newY.SSISABYVER <> 0 Then xSet = xSet & ",SSISABYVER": xValues = xValues & " ," & newY.SSISABYVER
If newY.SSISABYAMJ <> 0 Then xSet = xSet & ",SSISABYAMJ": xValues = xValues & " ," & newY.SSISABYAMJ
If newY.SSISABYHMS <> 0 Then xSet = xSet & ",SSISABYHMS": xValues = xValues & " ," & newY.SSISABYHMS

If Trim(newY.SSISABSTAK) <> "" Then xSet = xSet & ",SSISABSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABSTAK), "'", "''") & "'"
If Trim(newY.SSISABPRFX) <> "" Then xSet = xSet & ",SSISABPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABPRFX), "'", "''") & "'"
If Trim(newY.SSISABPRFK) <> "" Then xSet = xSet & ",SSISABPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABPRFK), "'", "''") & "'"
If Trim(newY.SSISABUNOM) <> "" Then xSet = xSet & ",SSISABUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABUNOM), "'", "''") & "'"
If Trim(newY.SSISABYFCT) <> "" Then xSet = xSet & ",SSISABYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABYFCT), "'", "''") & "'"
If Trim(newY.SSISABYUSR) <> "" Then xSet = xSet & ",SSISABYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABYUSR), "'", "''") & "'"
If Trim(newY.SSISABINFO) <> "" Then xSet = xSet & ",SSISABINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSISAB0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSISAB0_Insert = "Erreur màj : " & newY.SSISABPRFK
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSISAB0_Insert = "sqlYSSISAB0_Insert " & vbCrLf & Error
End Function
Public Function sqlYSSISABH_Insert(newY As typeYSSISAB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSISABH_Insert = Null

xSet = " (SSISABNAT,SSISABUIDX,SSISABULOT"
xValues = " values('" & newY.SSISABNAT & "','" & newY.SSISABUIDX & "'," & newY.SSISABULOT

' Détecter les modifications
'===================================================================================
If newY.SSISABUIDD <> 0 Then xSet = xSet & ",SSISABUIDD": xValues = xValues & " ," & newY.SSISABUIDD
If newY.SSISABTLNK <> 0 Then xSet = xSet & ",SSISABTLNK": xValues = xValues & " ," & newY.SSISABTLNK
If newY.SSISABYVER <> 0 Then xSet = xSet & ",SSISABYVER": xValues = xValues & " ," & newY.SSISABYVER
If newY.SSISABYAMJ <> 0 Then xSet = xSet & ",SSISABYAMJ": xValues = xValues & " ," & newY.SSISABYAMJ
If newY.SSISABYHMS <> 0 Then xSet = xSet & ",SSISABYHMS": xValues = xValues & " ," & newY.SSISABYHMS

If Trim(newY.SSISABSTAK) <> "" Then xSet = xSet & ",SSISABSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABSTAK), "'", "''") & "'"
If Trim(newY.SSISABPRFX) <> "" Then xSet = xSet & ",SSISABPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABPRFX), "'", "''") & "'"
If Trim(newY.SSISABPRFK) <> "" Then xSet = xSet & ",SSISABPRFK": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABPRFK), "'", "''") & "'"
If Trim(newY.SSISABUNOM) <> "" Then xSet = xSet & ",SSISABUNOM": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABUNOM), "'", "''") & "'"
If Trim(newY.SSISABYFCT) <> "" Then xSet = xSet & ",SSISABYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABYFCT), "'", "''") & "'"
If Trim(newY.SSISABYUSR) <> "" Then xSet = xSet & ",SSISABYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABYUSR), "'", "''") & "'"
If Trim(newY.SSISABINFO) <> "" Then xSet = xSet & ",SSISABINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSISABINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSISABH" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSISABH_Insert = "Erreur màj : " & newY.SSISABPRFK
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSISABH_Insert = "sqlYSSISABH_Insert " & vbCrLf & Error
End Function


Public Function sqlYSSISAB0_Update(newY As typeYSSISAB0, oldY As typeYSSISAB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSISAB0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSISABUIDX <> newY.SSISABUIDX Then
    sqlYSSISAB0_Update = "Erreur SSISABUIDX : " & newY.SSISABUIDX & " / " & oldY.SSISABUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSISABNAT = '" & oldY.SSISABNAT & "'" _
       & " and SSISABUIDX = '" & oldY.SSISABUIDX & "'" _
       & " and SSISABULOT = " & oldY.SSISABULOT _
       & " and SSISABYVER = " & oldY.SSISABYVER
       
newY.SSISABYVER = newY.SSISABYVER + 1
xSet = xSet & " set SSISABYVER = " & newY.SSISABYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSISABUIDD <> oldY.SSISABUIDD Then blnUpdate = True: xSet = xSet & " , SSISABUIDD = " & newY.SSISABUIDD
If newY.SSISABTLNK <> oldY.SSISABTLNK Then blnUpdate = True: xSet = xSet & " , SSISABTLNK = " & newY.SSISABTLNK
If newY.SSISABYAMJ <> oldY.SSISABYAMJ Then blnUpdate = True: xSet = xSet & " , SSISABYAMJ = " & newY.SSISABYAMJ
If newY.SSISABYHMS <> oldY.SSISABYHMS Then blnUpdate = True: xSet = xSet & " , SSISABYHMS = " & newY.SSISABYHMS

If newY.SSISABSTAK <> oldY.SSISABSTAK Then blnUpdate = True:  xSet = xSet & " , SSISABSTAK = '" & Replace(Trim(newY.SSISABSTAK), "'", "''") & "'"
If newY.SSISABPRFX <> oldY.SSISABPRFX Then blnUpdate = True:  xSet = xSet & " , SSISABPRFX = '" & Replace(Trim(newY.SSISABPRFX), "'", "''") & "'"
If newY.SSISABPRFK <> oldY.SSISABPRFK Then blnUpdate = True:  xSet = xSet & " , SSISABPRFK = '" & Replace(Trim(newY.SSISABPRFK), "'", "''") & "'"
If newY.SSISABUNOM <> oldY.SSISABUNOM Then blnUpdate = True:  xSet = xSet & " , SSISABUNOM = '" & Replace(Trim(newY.SSISABUNOM), "'", "''") & "'"
If newY.SSISABYFCT <> oldY.SSISABYFCT Then blnUpdate = True:  xSet = xSet & " , SSISABYFCT = '" & Replace(Trim(newY.SSISABYFCT), "'", "''") & "'"
If newY.SSISABINFO <> oldY.SSISABINFO Then blnUpdate = True:  xSet = xSet & " , SSISABINFO= '" & Replace(Trim(newY.SSISABINFO), "'", "''") & "'"
If newY.SSISABYUSR <> oldY.SSISABYUSR Then blnUpdate = True:  xSet = xSet & " , SSISABYUSR = '" & Replace(Trim(newY.SSISABYUSR), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSISAB0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSISAB0_Update = "Erreur màj : " & newY.SSISABUIDX
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSISAB0_Update = "sqlYSSISAB0_Update " & vbCrLf & Error
End Function
Public Function sqlYSSISABH_Update(newY As typeYSSISAB0, oldY As typeYSSISAB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSISABH_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSISABUIDX <> newY.SSISABUIDX Then
    sqlYSSISABH_Update = "Erreur SSISABUIDX : " & newY.SSISABUIDX & " / " & oldY.SSISABUIDX
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSISABNAT = '" & oldY.SSISABNAT & "'" _
       & " and SSISABUIDX = '" & oldY.SSISABUIDX & "'" _
       & " and SSISABULOT = " & oldY.SSISABULOT _
       & " and SSISABYVER = " & oldY.SSISABYVER
       
xSet = xSet & " set SSISABYVER = " & newY.SSISABYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSISABUIDD <> oldY.SSISABUIDD Then blnUpdate = True: xSet = xSet & " , SSISABUIDD = " & newY.SSISABUIDD
If newY.SSISABTLNK <> oldY.SSISABTLNK Then blnUpdate = True: xSet = xSet & " , SSISABTLNK = " & newY.SSISABTLNK
If newY.SSISABYAMJ <> oldY.SSISABYAMJ Then blnUpdate = True: xSet = xSet & " , SSISABYAMJ = " & newY.SSISABYAMJ
If newY.SSISABYHMS <> oldY.SSISABYHMS Then blnUpdate = True: xSet = xSet & " , SSISABYHMS = " & newY.SSISABYHMS

If newY.SSISABSTAK <> oldY.SSISABSTAK Then blnUpdate = True:  xSet = xSet & " , SSISABSTAK = '" & Replace(Trim(newY.SSISABSTAK), "'", "''") & "'"
If newY.SSISABPRFX <> oldY.SSISABPRFX Then blnUpdate = True:  xSet = xSet & " , SSISABPRFX = '" & Replace(Trim(newY.SSISABPRFX), "'", "''") & "'"
If newY.SSISABPRFK <> oldY.SSISABPRFK Then blnUpdate = True:  xSet = xSet & " , SSISABPRFK = '" & Replace(Trim(newY.SSISABPRFK), "'", "''") & "'"
If newY.SSISABUNOM <> oldY.SSISABUNOM Then blnUpdate = True:  xSet = xSet & " , SSISABUNOM = '" & Replace(Trim(newY.SSISABUNOM), "'", "''") & "'"
If newY.SSISABYFCT <> oldY.SSISABYFCT Then blnUpdate = True:  xSet = xSet & " , SSISABYFCT = '" & Replace(Trim(newY.SSISABYFCT), "'", "''") & "'"
If newY.SSISABINFO <> oldY.SSISABINFO Then blnUpdate = True:  xSet = xSet & " , SSISABINFO= '" & Replace(Trim(newY.SSISABINFO), "'", "''") & "'"
If newY.SSISABYUSR <> oldY.SSISABYUSR Then blnUpdate = True:  xSet = xSet & " , SSISABYUSR = '" & Replace(Trim(newY.SSISABYUSR), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSISABH" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSISABH_Update = "Erreur màj : " & newY.SSISABUIDX
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSISABH_Update = "sqlYSSISABH_Update " & vbCrLf & Error
End Function


Public Function rsYSSISAB0_GetBuffer(rsAdo As ADODB.Recordset, lYSSISAB0 As typeYSSISAB0)
On Error GoTo Error_Handler
rsYSSISAB0_GetBuffer = Null
lYSSISAB0.SSISABNAT = rsAdo("SSISABNAT")
lYSSISAB0.SSISABUIDX = Trim(rsAdo("SSISABUIDX"))
lYSSISAB0.SSISABULOT = rsAdo("SSISABULOT")

lYSSISAB0.SSISABSTAK = rsAdo("SSISABSTAK")
lYSSISAB0.SSISABUIDD = rsAdo("SSISABUIDD")
lYSSISAB0.SSISABPRFX = Trim(rsAdo("SSISABPRFX"))
lYSSISAB0.SSISABPRFK = rsAdo("SSISABPRFK")
lYSSISAB0.SSISABUNOM = Trim(rsAdo("SSISABUNOM"))

lYSSISAB0.SSISABTLNK = rsAdo("SSISABTLNK")
lYSSISAB0.SSISABYFCT = Trim(rsAdo("SSISABYFCT"))
lYSSISAB0.SSISABYUSR = Trim(rsAdo("SSISABYUSR"))
lYSSISAB0.SSISABYAMJ = rsAdo("SSISABYAMJ")

lYSSISAB0.SSISABYHMS = rsAdo("SSISABYHMS")
lYSSISAB0.SSISABYVER = rsAdo("SSISABYVER")
lYSSISAB0.SSISABINFO = RTrim(rsAdo("SSISABINFO"))

Exit Function
Error_Handler:
rsYSSISAB0_GetBuffer = Error


End Function

Public Function rsYSSISAB0_Init(lYSSISAB0 As typeYSSISAB0)
lYSSISAB0.SSISABNAT = ""
lYSSISAB0.SSISABUIDX = ""
lYSSISAB0.SSISABULOT = 0

lYSSISAB0.SSISABSTAK = ""
lYSSISAB0.SSISABUIDD = 0
lYSSISAB0.SSISABPRFX = ""
lYSSISAB0.SSISABPRFK = ""
lYSSISAB0.SSISABUNOM = ""

lYSSISAB0.SSISABTLNK = 0
lYSSISAB0.SSISABYFCT = ""
lYSSISAB0.SSISABYUSR = ""
lYSSISAB0.SSISABYAMJ = 0
lYSSISAB0.SSISABYHMS = 0
lYSSISAB0.SSISABYVER = 0
lYSSISAB0.SSISABINFO = ""
End Function












