Attribute VB_Name = "srvYSSIUSR0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSIUSR0
    SSIUSRNAT    As String
    SSIUSRUIDN   As Long
    SSIUSRUIDX   As String
    SSIUSRUNIT   As String
    SSIUSRSTAK   As String
    SSIUSRDECH   As Long
    SSIUSRPRFX   As String
    SSIUSRPRFK   As String
    SSIUSRPRFD   As Long
    SSIUSRPRFH   As Long
    SSIUSRTLNK   As Long
    SSIUSRYFCT   As String
    SSIUSRYUSR   As String
    SSIUSRYAMJ   As Long
    SSIUSRYHMS   As Long
    SSIUSRYVER   As Long

End Type
Public xYSSIUSR0 As typeYSSIUSR0
Public Function sqlYSSIUSR0_Update_CMD(lSQL As String)
Dim X As String, Nb As Long

On Error GoTo Error_Handler
sqlYSSIUSR0_Update_CMD = Null

Set rsAdo = cnSab_Update.Execute(lSQL, Nb)

Exit Function
Error_Handler:
    sqlYSSIUSR0_Update_CMD = "sqlYSSIUSR0_Update_CMD " & vbCrLf & Error
End Function

Public Function sqlYSSIUSR0_Delete(oldY As typeYSSIUSR0)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYSSIUSR0_Delete = Null

xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
       & " where SSIUSRNAT = '" & oldY.SSIUSRNAT & "'" _
       & " and SSIUSRUIDN = " & oldY.SSIUSRUIDN
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
If Nb = 0 Then
    sqlYSSIUSR0_Delete = xSQL
    Exit Function
End If
    


Exit Function
Error_Handler:
    sqlYSSIUSR0_Delete = "sqlYSSIUSR0_Delete " & vbCrLf & Error
End Function


Public Function sqlYSSIUSR0_Insert(newY As typeYSSIUSR0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIUSR0_Insert = Null

xSet = " (SSIUSRNAT,SSIUSRUIDN"
xValues = " values('" & newY.SSIUSRNAT & "'," & newY.SSIUSRUIDN

' Détecter les modifications
'===================================================================================
If newY.SSIUSRYVER <> 0 Then xSet = xSet & ",SSIUSRYVER": xValues = xValues & " ," & newY.SSIUSRYVER
If newY.SSIUSRDECH <> 0 Then xSet = xSet & ",SSIUSRDECH": xValues = xValues & " ," & newY.SSIUSRDECH
If newY.SSIUSRPRFD <> 0 Then xSet = xSet & ",SSIUSRPRFD": xValues = xValues & " ," & newY.SSIUSRPRFD
If newY.SSIUSRPRFH <> 0 Then xSet = xSet & ",SSIUSRPRFH": xValues = xValues & " ," & newY.SSIUSRPRFH
If newY.SSIUSRTLNK <> 0 Then xSet = xSet & ",SSIUSRTLNK": xValues = xValues & " ," & newY.SSIUSRTLNK
If newY.SSIUSRYAMJ <> 0 Then xSet = xSet & ",SSIUSRYAMJ": xValues = xValues & " ," & newY.SSIUSRYAMJ
If newY.SSIUSRYHMS <> 0 Then xSet = xSet & ",SSIUSRYHMS": xValues = xValues & " ," & newY.SSIUSRYHMS

If Trim(newY.SSIUSRPRFK) <> "" Then xSet = xSet & ",SSIUSRPRFK": xValues = xValues & " ,'" & newY.SSIUSRPRFK & "'"
If Trim(newY.SSIUSRUIDX) <> "" Then xSet = xSet & ",SSIUSRUIDX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRUIDX), "'", "''") & "'"
If Trim(newY.SSIUSRSTAK) <> "" Then xSet = xSet & ",SSIUSRSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRSTAK), "'", "''") & "'"
If Trim(newY.SSIUSRPRFX) <> "" Then xSet = xSet & ",SSIUSRPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRPRFX), "'", "''") & "'"
If Trim(newY.SSIUSRYUSR) <> "" Then xSet = xSet & ",SSIUSRYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRYUSR), "'", "''") & "'"
If Trim(newY.SSIUSRYFCT) <> "" Then xSet = xSet & ",SSIUSRYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRYFCT), "'", "''") & "'"
If Trim(newY.SSIUSRUNIT) <> "" Then xSet = xSet & ",SSIUSRUNIT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRUNIT), "'", "''") & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIUSR0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIUSR0_Insert = "Erreur màj : " & newY.SSIUSRUIDN
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIUSR0_Insert = "sqlYSSIUSR0_Insert " & vbCrLf & Error
End Function

Public Function sqlYSSIUSRH_Insert(newY As typeYSSIUSR0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSSIUSRH_Insert = Null

xSet = " (SSIUSRNAT,SSIUSRUIDN,SSIUSRYVER"
xValues = " values('" & newY.SSIUSRNAT & "'," & newY.SSIUSRUIDN & " ," & newY.SSIUSRYVER

' Détecter les modifications
'===================================================================================
If newY.SSIUSRDECH <> 0 Then xSet = xSet & ",SSIUSRDECH": xValues = xValues & " ," & newY.SSIUSRDECH
If newY.SSIUSRPRFD <> 0 Then xSet = xSet & ",SSIUSRPRFD": xValues = xValues & " ," & newY.SSIUSRPRFD
If newY.SSIUSRPRFH <> 0 Then xSet = xSet & ",SSIUSRPRFH": xValues = xValues & " ," & newY.SSIUSRPRFH
If newY.SSIUSRTLNK <> 0 Then xSet = xSet & ",SSIUSRTLNK": xValues = xValues & " ," & newY.SSIUSRTLNK
If newY.SSIUSRYAMJ <> 0 Then xSet = xSet & ",SSIUSRYAMJ": xValues = xValues & " ," & newY.SSIUSRYAMJ
If newY.SSIUSRYHMS <> 0 Then xSet = xSet & ",SSIUSRYHMS": xValues = xValues & " ," & newY.SSIUSRYHMS

If Trim(newY.SSIUSRPRFK) <> "" Then xSet = xSet & ",SSIUSRPRFK": xValues = xValues & " ,'" & newY.SSIUSRPRFK & "'"
If Trim(newY.SSIUSRUIDX) <> "" Then xSet = xSet & ",SSIUSRUIDX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRUIDX), "'", "''") & "'"
If Trim(newY.SSIUSRSTAK) <> "" Then xSet = xSet & ",SSIUSRSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRSTAK), "'", "''") & "'"
If Trim(newY.SSIUSRPRFX) <> "" Then xSet = xSet & ",SSIUSRPRFX": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRPRFX), "'", "''") & "'"
If Trim(newY.SSIUSRYUSR) <> "" Then xSet = xSet & ",SSIUSRYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRYUSR), "'", "''") & "'"
If Trim(newY.SSIUSRYFCT) <> "" Then xSet = xSet & ",SSIUSRYFCT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRYFCT), "'", "''") & "'"
If Trim(newY.SSIUSRUNIT) <> "" Then xSet = xSet & ",SSIUSRUNIT": xValues = xValues & " ,'" & Replace(Trim(newY.SSIUSRUNIT), "'", "''") & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIUSRH" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSIUSRH_Insert = "Erreur màj : " & newY.SSIUSRUIDN
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSIUSRH_Insert = "sqlYSSIUSRH_Insert " & vbCrLf & Error
End Function


Public Function sqlYSSIUSR0_Update(newY As typeYSSIUSR0, oldY As typeYSSIUSR0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIUSR0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIUSRUIDN <> newY.SSIUSRUIDN Then
    sqlYSSIUSR0_Update = "Erreur SSIUSRUIDN : " & newY.SSIUSRUIDN & " / " & oldY.SSIUSRUIDN
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIUSRNAT = '" & oldY.SSIUSRNAT & "'" _
       & " and SSIUSRUIDN = " & oldY.SSIUSRUIDN _
       & " and SSIUSRYVER = " & oldY.SSIUSRYVER
       
newY.SSIUSRYVER = newY.SSIUSRYVER + 1
xSet = xSet & " set SSIUSRYVER = " & newY.SSIUSRYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIUSRDECH <> oldY.SSIUSRDECH Then blnUpdate = True: xSet = xSet & " , SSIUSRDECH = " & newY.SSIUSRDECH
If newY.SSIUSRPRFD <> oldY.SSIUSRPRFD Then blnUpdate = True: xSet = xSet & " , SSIUSRPRFD = " & newY.SSIUSRPRFD
If newY.SSIUSRPRFH <> oldY.SSIUSRPRFH Then blnUpdate = True: xSet = xSet & " , SSIUSRPRFH = " & newY.SSIUSRPRFH
If newY.SSIUSRTLNK <> oldY.SSIUSRTLNK Then blnUpdate = True: xSet = xSet & " , SSIUSRTLNK = " & newY.SSIUSRTLNK
If newY.SSIUSRYAMJ <> oldY.SSIUSRYAMJ Then blnUpdate = True: xSet = xSet & " , SSIUSRYAMJ = " & newY.SSIUSRYAMJ
If newY.SSIUSRYHMS <> oldY.SSIUSRYHMS Then blnUpdate = True: xSet = xSet & " , SSIUSRYHMS = " & newY.SSIUSRYHMS

If newY.SSIUSRPRFK <> oldY.SSIUSRPRFK Then blnUpdate = True:  xSet = xSet & " , SSIUSRPRFK= '" & newY.SSIUSRPRFK & "'"
If newY.SSIUSRUIDX <> oldY.SSIUSRUIDX Then blnUpdate = True:  xSet = xSet & " , SSIUSRUIDX = '" & Replace(Trim(newY.SSIUSRUIDX), "'", "''") & "'"
If newY.SSIUSRSTAK <> oldY.SSIUSRSTAK Then blnUpdate = True:  xSet = xSet & " , SSIUSRSTAK = '" & Replace(Trim(newY.SSIUSRSTAK), "'", "''") & "'"
If newY.SSIUSRPRFX <> oldY.SSIUSRPRFX Then blnUpdate = True:  xSet = xSet & " , SSIUSRPRFX = '" & Replace(Trim(newY.SSIUSRPRFX), "'", "''") & "'"
If newY.SSIUSRYUSR <> oldY.SSIUSRYUSR Then blnUpdate = True:  xSet = xSet & " , SSIUSRYUSR = '" & Replace(Trim(newY.SSIUSRYUSR), "'", "''") & "'"
If newY.SSIUSRYFCT <> oldY.SSIUSRYFCT Then blnUpdate = True:  xSet = xSet & " , SSIUSRYFCT = '" & Replace(Trim(newY.SSIUSRYFCT), "'", "''") & "'"
If newY.SSIUSRUNIT <> oldY.SSIUSRUNIT Then blnUpdate = True:  xSet = xSet & " , SSIUSRUNIT = '" & Replace(Trim(newY.SSIUSRUNIT), "'", "''") & "'"
If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIUSR0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIUSR0_Update = "Erreur màj : " & newY.SSIUSRUIDN
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIUSR0_Update = "sqlYSSIUSR0_Update " & vbCrLf & Error
End Function

Public Function rsYSSIUSR0_GetBuffer(rsAdo As ADODB.Recordset, lYSSIUSR0 As typeYSSIUSR0)
On Error GoTo Error_Handler
rsYSSIUSR0_GetBuffer = Null
lYSSIUSR0.SSIUSRNAT = rsAdo("SSIUSRNAT")
lYSSIUSR0.SSIUSRUIDN = rsAdo("SSIUSRUIDN")

lYSSIUSR0.SSIUSRUIDX = Trim(rsAdo("SSIUSRUIDX"))
lYSSIUSR0.SSIUSRUNIT = Trim(rsAdo("SSIUSRUNIT"))
lYSSIUSR0.SSIUSRSTAK = rsAdo("SSIUSRSTAK")
lYSSIUSR0.SSIUSRDECH = rsAdo("SSIUSRDECH")
lYSSIUSR0.SSIUSRPRFX = Trim(rsAdo("SSIUSRPRFX"))
lYSSIUSR0.SSIUSRPRFK = rsAdo("SSIUSRPRFK")
lYSSIUSR0.SSIUSRPRFD = rsAdo("SSIUSRPRFD")
lYSSIUSR0.SSIUSRPRFH = rsAdo("SSIUSRPRFH")
lYSSIUSR0.SSIUSRTLNK = rsAdo("SSIUSRTLNK")

lYSSIUSR0.SSIUSRYFCT = rsAdo("SSIUSRYFCT")
lYSSIUSR0.SSIUSRYUSR = Trim(rsAdo("SSIUSRYUSR"))
lYSSIUSR0.SSIUSRYAMJ = rsAdo("SSIUSRYAMJ")
lYSSIUSR0.SSIUSRYHMS = rsAdo("SSIUSRYHMS")
lYSSIUSR0.SSIUSRYVER = rsAdo("SSIUSRYVER")

Exit Function
Error_Handler:
rsYSSIUSR0_GetBuffer = Error


End Function

Public Function rsYSSIUSR0_Init(lYSSIUSR0 As typeYSSIUSR0)
lYSSIUSR0.SSIUSRUIDN = 0
lYSSIUSR0.SSIUSRYVER = 0
lYSSIUSR0.SSIUSRNAT = ""
lYSSIUSR0.SSIUSRUIDX = ""
lYSSIUSR0.SSIUSRUNIT = ""
lYSSIUSR0.SSIUSRSTAK = ""
lYSSIUSR0.SSIUSRDECH = 0
lYSSIUSR0.SSIUSRPRFD = 0
lYSSIUSR0.SSIUSRPRFH = 0
lYSSIUSR0.SSIUSRPRFX = ""
lYSSIUSR0.SSIUSRTLNK = 0
lYSSIUSR0.SSIUSRYFCT = ""
lYSSIUSR0.SSIUSRYAMJ = 0
lYSSIUSR0.SSIUSRYUSR = ""
lYSSIUSR0.SSIUSRYHMS = 0
lYSSIUSR0.SSIUSRPRFK = ""

End Function





