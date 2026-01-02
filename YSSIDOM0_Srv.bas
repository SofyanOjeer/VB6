Attribute VB_Name = "srvYSSIDOM0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSIDOM0
    SSIDOMNAT    As String
    SSIDOMUIDN   As Long
    SSIDOMDIDX   As String
    SSIDOMUIDD   As Long
    SSIDOMUIDX   As String
    SSIDOMUNIT   As String
    SSIDOMSTAK   As String
    SSIDOMDECH   As Long
    SSIDOMPRFX   As String
    SSIDOMPRFK   As String
    SSIDOMPRFD   As Long
    SSIDOMPRFH   As Long
    SSIDOMTLNK   As Long
    SSIDOMYFCT   As String
    SSIDOMYUSR   As String
    SSIDOMYAMJ   As Long
    SSIDOMYHMS   As Long
    SSIDOMYVER   As Long

End Type
Public xYSSIDOM0 As typeYSSIDOM0
Public Function sqlYSSIDOM0_Insert(newY As typeYSSIDOM0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIDOM0_Insert = Null

xSet = " (SSIDOMNAT,SSIDOMUIDN,SSIDOMDIDX,SSIDOMUIDX,SSIDOMUIDD"
xValues = " values('" & newY.SSIDOMNAT & "'," & newY.SSIDOMUIDN & ",'" & newY.SSIDOMDIDX & "','" & Replace(Trim(newY.SSIDOMUIDX), "'", "''") & "'," & newY.SSIDOMUIDD

' Détecter les modifications
'===================================================================================
'If newY.SSIDOMUIDD <> 0 Then xSet = xSet & ",SSIDOMUIDD": xValues = xValues & " ," & newY.SSIDOMUIDD
If newY.SSIDOMYVER <> 0 Then xSet = xSet & ",SSIDOMYVER": xValues = xValues & " ," & newY.SSIDOMYVER
If newY.SSIDOMDECH <> 0 Then xSet = xSet & ",SSIDOMDECH": xValues = xValues & " ," & newY.SSIDOMDECH
If newY.SSIDOMPRFD <> 0 Then xSet = xSet & ",SSIDOMPRFD": xValues = xValues & " ," & newY.SSIDOMPRFD
If newY.SSIDOMPRFH <> 0 Then xSet = xSet & ",SSIDOMPRFH": xValues = xValues & " ," & newY.SSIDOMPRFH
If newY.SSIDOMTLNK <> 0 Then xSet = xSet & ",SSIDOMTLNK": xValues = xValues & " ," & newY.SSIDOMTLNK
If newY.SSIDOMYAMJ <> 0 Then xSet = xSet & ",SSIDOMYAMJ": xValues = xValues & " ," & newY.SSIDOMYAMJ
If newY.SSIDOMYHMS <> 0 Then xSet = xSet & ",SSIDOMYHMS": xValues = xValues & " ," & newY.SSIDOMYHMS

If Trim(newY.SSIDOMPRFK) <> "" Then xSet = xSet & ",SSIDOMPRFK": xValues = xValues & " ,'" & newY.SSIDOMPRFK & "'"
If Trim(newY.SSIDOMSTAK) <> "" Then xSet = xSet & ",SSIDOMSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMSTAK), "'", "''") & "'"
If Trim(newY.SSIDOMPRFX) <> "" Then xSet = xSet & ",SSIDOMPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMPRFX), "'", "''") & "'"
If Trim(newY.SSIDOMYUSR) <> "" Then xSet = xSet & ",SSIDOMYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMYUSR), "'", "''") & "'"
If Trim(newY.SSIDOMYFCT) <> "" Then xSet = xSet & ",SSIDOMYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMYFCT), "'", "''") & "'"
If Trim(newY.SSIDOMUNIT) <> "" Then xSet = xSet & ",SSIDOMUNIT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMUNIT), "'", "''") & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIDOM0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIDOM0_Insert = "Erreur màj : " & newY.SSIDOMUIDN
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIDOM0_Insert = "sqlYSSIDOM0_Insert " & vbCrLf & Error
End Function

Public Function sqlYSSIDOMH_Insert(newY As typeYSSIDOM0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIDOMH_Insert = Null

xSet = " (SSIDOMNAT,SSIDOMUIDN,SSIDOMDIDX,SSIDOMUIDX,SSIDOMUIDD,SSIDOMYVER"
xValues = " values('" & newY.SSIDOMNAT & "'," & newY.SSIDOMUIDN & ",'" & newY.SSIDOMDIDX & "','" & Replace(Trim(newY.SSIDOMUIDX), "'", "''") & "'," & newY.SSIDOMUIDD & "," & newY.SSIDOMYVER

' Détecter les modifications
'===================================================================================
'If newY.SSIDOMUIDD <> 0 Then xSet = xSet & ",SSIDOMUIDD": xValues = xValues & " ," & newY.SSIDOMUIDD
If newY.SSIDOMDECH <> 0 Then xSet = xSet & ",SSIDOMDECH": xValues = xValues & " ," & newY.SSIDOMDECH
If newY.SSIDOMPRFD <> 0 Then xSet = xSet & ",SSIDOMPRFD": xValues = xValues & " ," & newY.SSIDOMPRFD
If newY.SSIDOMPRFH <> 0 Then xSet = xSet & ",SSIDOMPRFH": xValues = xValues & " ," & newY.SSIDOMPRFH
If newY.SSIDOMTLNK <> 0 Then xSet = xSet & ",SSIDOMTLNK": xValues = xValues & " ," & newY.SSIDOMTLNK
If newY.SSIDOMYAMJ <> 0 Then xSet = xSet & ",SSIDOMYAMJ": xValues = xValues & " ," & newY.SSIDOMYAMJ
If newY.SSIDOMYHMS <> 0 Then xSet = xSet & ",SSIDOMYHMS": xValues = xValues & " ," & newY.SSIDOMYHMS

If Trim(newY.SSIDOMPRFK) <> "" Then xSet = xSet & ",SSIDOMPRFK": xValues = xValues & " ,'" & newY.SSIDOMPRFK & "'"
If Trim(newY.SSIDOMSTAK) <> "" Then xSet = xSet & ",SSIDOMSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMSTAK), "'", "''") & "'"
If Trim(newY.SSIDOMPRFX) <> "" Then xSet = xSet & ",SSIDOMPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMPRFX), "'", "''") & "'"
If Trim(newY.SSIDOMYUSR) <> "" Then xSet = xSet & ",SSIDOMYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMYUSR), "'", "''") & "'"
If Trim(newY.SSIDOMYFCT) <> "" Then xSet = xSet & ",SSIDOMYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMYFCT), "'", "''") & "'"
If Trim(newY.SSIDOMUNIT) <> "" Then xSet = xSet & ",SSIDOMUNIT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIDOMUNIT), "'", "''") & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIDOMH" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIDOMH_Insert = "Erreur màj : " & xSQL
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIDOMH_Insert = "sqlYSSIDOMH_Insert " & vbCrLf & Error
End Function


Public Function sqlYSSIDOM0_Update(newY As typeYSSIDOM0, oldY As typeYSSIDOM0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIDOM0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIDOMUIDN <> newY.SSIDOMUIDN Then
    sqlYSSIDOM0_Update = "Erreur SSIDOMUIDN : " & newY.SSIDOMUIDN & " / " & oldY.SSIDOMUIDN
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIDOMNAT = '" & oldY.SSIDOMNAT & "'" _
       & " and SSIDOMUIDN = " & oldY.SSIDOMUIDN _
       & " and SSIDOMDIDX = '" & oldY.SSIDOMDIDX & "'" _
       & " and SSIDOMUIDX = '" & Replace(Trim(oldY.SSIDOMUIDX), "'", "''") & "'" _
       & " and SSIDOMUIDD = " & oldY.SSIDOMUIDD _
       & " and SSIDOMYVER = " & oldY.SSIDOMYVER
       
newY.SSIDOMYVER = newY.SSIDOMYVER + 1
xSet = xSet & " set SSIDOMYVER = " & newY.SSIDOMYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIDOMUIDD <> oldY.SSIDOMUIDD Then blnUpdate = True: xSet = xSet & " , SSIDOMUIDD = " & newY.SSIDOMUIDD
If newY.SSIDOMDECH <> oldY.SSIDOMDECH Then blnUpdate = True: xSet = xSet & " , SSIDOMDECH = " & newY.SSIDOMDECH
If newY.SSIDOMPRFD <> oldY.SSIDOMPRFD Then blnUpdate = True: xSet = xSet & " , SSIDOMPRFD = " & newY.SSIDOMPRFD
If newY.SSIDOMPRFH <> oldY.SSIDOMPRFH Then blnUpdate = True: xSet = xSet & " , SSIDOMPRFH = " & newY.SSIDOMPRFH
If newY.SSIDOMTLNK <> oldY.SSIDOMTLNK Then blnUpdate = True: xSet = xSet & " , SSIDOMTLNK = " & newY.SSIDOMTLNK
If newY.SSIDOMYAMJ <> oldY.SSIDOMYAMJ Then blnUpdate = True: xSet = xSet & " , SSIDOMYAMJ = " & newY.SSIDOMYAMJ
If newY.SSIDOMYHMS <> oldY.SSIDOMYHMS Then blnUpdate = True: xSet = xSet & " , SSIDOMYHMS = " & newY.SSIDOMYHMS

If newY.SSIDOMUIDX <> oldY.SSIDOMUIDX Then blnUpdate = True:  xSet = xSet & " , SSIDOMUIDX = '" & Replace(Trim(newY.SSIDOMUIDX), "'", "''") & "'"
If newY.SSIDOMPRFK <> oldY.SSIDOMPRFK Then blnUpdate = True:  xSet = xSet & " , SSIDOMPRFK= '" & newY.SSIDOMPRFK & "'"
If newY.SSIDOMSTAK <> oldY.SSIDOMSTAK Then blnUpdate = True:  xSet = xSet & " , SSIDOMSTAK = '" & Replace(Trim(newY.SSIDOMSTAK), "'", "''") & "'"
If newY.SSIDOMPRFX <> oldY.SSIDOMPRFX Then blnUpdate = True:  xSet = xSet & " , SSIDOMPRFX = '" & Replace(Trim(newY.SSIDOMPRFX), "'", "''") & "'"
If newY.SSIDOMYUSR <> oldY.SSIDOMYUSR Then blnUpdate = True:  xSet = xSet & " , SSIDOMYUSR = '" & Replace(Trim(newY.SSIDOMYUSR), "'", "''") & "'"
If newY.SSIDOMYFCT <> oldY.SSIDOMYFCT Then blnUpdate = True:  xSet = xSet & " , SSIDOMYFCT = '" & Replace(Trim(newY.SSIDOMYFCT), "'", "''") & "'"
If newY.SSIDOMUNIT <> oldY.SSIDOMUNIT Then blnUpdate = True:  xSet = xSet & " , SSIDOMUNIT = '" & Replace(Trim(newY.SSIDOMUNIT), "'", "''") & "'"


If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIDOM0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIDOM0_Update = "Erreur màj : " & xSQL
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIDOM0_Update = "sqlYSSIDOM0_Update " & vbCrLf & Error
End Function
Public Function sqlYSSIDOM0_Delete(oldY As typeYSSIDOM0)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYSSIDOM0_Delete = Null

xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
       & " where SSIDOMNAT = '" & oldY.SSIDOMNAT & "'" _
       & " and SSIDOMUIDN = " & oldY.SSIDOMUIDN _
       & " and SSIDOMDIDX = '" & oldY.SSIDOMDIDX & "'" _
       & " and SSIDOMUIDX = '" & oldY.SSIDOMUIDX & "'" _
       & " and SSIDOMUIDD = " & oldY.SSIDOMUIDD _
       & " and SSIDOMYVER = " & oldY.SSIDOMYVER
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
If Nb = 0 Then
    sqlYSSIDOM0_Delete = xSQL
    Exit Function
End If
    


Exit Function
Error_Handler:
    sqlYSSIDOM0_Delete = "sqlYSSIDOM0_Delete " & vbCrLf & Error
End Function

Public Function sqlYSSIDOM0_Update_CMD(lSQL As String)
Dim X As String, Nb As Long

On Error GoTo Error_Handler
sqlYSSIDOM0_Update_CMD = Null

Set rsAdo = cnSab_Update.Execute(lSQL, Nb)

'If Nb = 0 Then
'    sqlYSSIDOM0_Update_CMD = "" 'lSQL
'    Exit Function
'End If
    


Exit Function
Error_Handler:
    sqlYSSIDOM0_Update_CMD = "sqlYSSIDOM0_Update_CMD " & vbCrLf & Error
End Function

Public Function rsYSSIDOM0_GetBuffer(rsAdo As ADODB.Recordset, lYSSIDOM0 As typeYSSIDOM0)
On Error GoTo Error_Handler
rsYSSIDOM0_GetBuffer = Null
lYSSIDOM0.SSIDOMNAT = rsAdo("SSIDOMNAT")
lYSSIDOM0.SSIDOMUIDN = rsAdo("SSIDOMUIDN")
lYSSIDOM0.SSIDOMDIDX = Trim(rsAdo("SSIDOMDIDX"))
lYSSIDOM0.SSIDOMUNIT = Trim(rsAdo("SSIDOMUNIT"))
lYSSIDOM0.SSIDOMUIDD = rsAdo("SSIDOMUIDD")
lYSSIDOM0.SSIDOMUIDX = Trim(rsAdo("SSIDOMUIDX"))
lYSSIDOM0.SSIDOMSTAK = rsAdo("SSIDOMSTAK")
lYSSIDOM0.SSIDOMDECH = rsAdo("SSIDOMDECH")
lYSSIDOM0.SSIDOMPRFX = Trim(rsAdo("SSIDOMPRFX"))
lYSSIDOM0.SSIDOMPRFK = rsAdo("SSIDOMPRFK")
lYSSIDOM0.SSIDOMPRFD = rsAdo("SSIDOMPRFD")
lYSSIDOM0.SSIDOMPRFH = rsAdo("SSIDOMPRFH")
lYSSIDOM0.SSIDOMTLNK = rsAdo("SSIDOMTLNK")
lYSSIDOM0.SSIDOMYUSR = Trim(rsAdo("SSIDOMYUSR"))
lYSSIDOM0.SSIDOMYFCT = Trim(rsAdo("SSIDOMYFCT"))
lYSSIDOM0.SSIDOMYAMJ = rsAdo("SSIDOMYAMJ")

lYSSIDOM0.SSIDOMYHMS = rsAdo("SSIDOMYHMS")
lYSSIDOM0.SSIDOMYVER = rsAdo("SSIDOMYVER")

Exit Function
Error_Handler:
rsYSSIDOM0_GetBuffer = Error


End Function

Public Function rsYSSIDOM0_Init(lYSSIDOM0 As typeYSSIDOM0)
lYSSIDOM0.SSIDOMUIDN = 0
lYSSIDOM0.SSIDOMYVER = 0
lYSSIDOM0.SSIDOMNAT = ""
lYSSIDOM0.SSIDOMDIDX = ""
lYSSIDOM0.SSIDOMUNIT = ""
lYSSIDOM0.SSIDOMUIDD = 0
lYSSIDOM0.SSIDOMUIDX = ""
lYSSIDOM0.SSIDOMSTAK = ""
lYSSIDOM0.SSIDOMDECH = 0
lYSSIDOM0.SSIDOMPRFD = 0
lYSSIDOM0.SSIDOMPRFH = 0
lYSSIDOM0.SSIDOMPRFX = ""
lYSSIDOM0.SSIDOMTLNK = 0
lYSSIDOM0.SSIDOMYAMJ = 0
lYSSIDOM0.SSIDOMYFCT = ""
lYSSIDOM0.SSIDOMYUSR = ""
lYSSIDOM0.SSIDOMYHMS = 0
lYSSIDOM0.SSIDOMPRFK = ""

End Function







