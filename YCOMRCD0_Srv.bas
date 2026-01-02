Attribute VB_Name = "srvYCOMRCD0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYCOMRCD0
    COMRCDNAT    As String
    COMRCDPIE   As Long
    COMRCDECR   As Long
    COMRCDSER   As String
    COMRCDSSE   As String
    COMRCDCLI   As String
    COMRCDOPE   As String
    COMRCDNUM   As Long
    COMRCDDTR   As Long
    COMRCDPCI   As String
    COMRCDDEV   As String
    COMRCDMTD   As Currency
    COMRCDMTR   As Currency
    COMRCDSTA   As String
    COMRCDRLV   As Long
    COMRCDYUSR   As String
    COMRCDYAMJ   As Long
    COMRCDYHMS   As Long
    COMRCDYVER   As Long
    
    COMRCDZTYP   As String
    COMRCDZORD   As String
    COMRCDZCOM   As String

End Type
Public xYCOMRCD0 As typeYCOMRCD0
Public Function sqlYCOMRCD0_Insert(newY As typeYCOMRCD0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCOMRCD0_Insert = Null

xSet = " (COMRCDNAT,COMRCDPIE,COMRCDECR"
xValues = " values('" & newY.COMRCDNAT & "'," & newY.COMRCDPIE & "," & newY.COMRCDECR

newY.COMRCDYAMJ = DSys
newY.COMRCDYHMS = time_Hms
newY.COMRCDYUSR = usrName_UCase

' Détecter les modifications
'===================================================================================
If newY.COMRCDRLV <> 0 Then xSet = xSet & ",COMRCDRLV": xValues = xValues & " ," & newY.COMRCDRLV
If newY.COMRCDYVER <> 0 Then xSet = xSet & ",COMRCDYVER": xValues = xValues & " ," & newY.COMRCDYVER
If newY.COMRCDNUM <> 0 Then xSet = xSet & ",COMRCDNUM": xValues = xValues & " ," & newY.COMRCDNUM
If newY.COMRCDDTR <> 0 Then xSet = xSet & ",COMRCDDTR": xValues = xValues & " ," & newY.COMRCDDTR
If newY.COMRCDMTD <> 0 Then xSet = xSet & ",COMRCDMTD": xValues = xValues & " ," & Replace(newY.COMRCDMTD, ",", ".")
If newY.COMRCDMTR <> 0 Then xSet = xSet & ",COMRCDMTR": xValues = xValues & " ," & Replace(newY.COMRCDMTR, ",", ".")
If newY.COMRCDYAMJ <> 0 Then xSet = xSet & ",COMRCDYAMJ": xValues = xValues & " ," & newY.COMRCDYAMJ
If newY.COMRCDYHMS <> 0 Then xSet = xSet & ",COMRCDYHMS": xValues = xValues & " ," & newY.COMRCDYHMS

If Trim(newY.COMRCDSER) <> "" Then xSet = xSet & ",COMRCDSER": xValues = xValues & " ,'" & newY.COMRCDSER & "'"
If Trim(newY.COMRCDSSE) <> "" Then xSet = xSet & ",COMRCDSSE": xValues = xValues & " ,'" & newY.COMRCDSSE & "'"
If Trim(newY.COMRCDDEV) <> "" Then xSet = xSet & ",COMRCDDEV": xValues = xValues & " ,'" & newY.COMRCDDEV & "'"
If Trim(newY.COMRCDOPE) <> "" Then xSet = xSet & ",COMRCDOPE": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDOPE), "'", "''") & "'"
If Trim(newY.COMRCDPCI) <> "" Then xSet = xSet & ",COMRCDPCI": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDPCI), "'", "''") & "'"
If Trim(newY.COMRCDYUSR) <> "" Then xSet = xSet & ",COMRCDYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDYUSR), "'", "''") & "'"
If Trim(newY.COMRCDSTA) <> "" Then xSet = xSet & ",COMRCDSTA": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDSTA), "'", "''") & "'"
If Trim(newY.COMRCDCLI) <> "" Then xSet = xSet & ",COMRCDCLI": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDCLI), "'", "''") & "'"

If Trim(newY.COMRCDZTYP) <> "" Then xSet = xSet & ",COMRCDZTYP": xValues = xValues & " ,'" & newY.COMRCDZTYP & "'"
If Trim(newY.COMRCDZORD) <> "" Then xSet = xSet & ",COMRCDZORD": xValues = xValues & " ,'" & newY.COMRCDZORD & "'"
If Trim(newY.COMRCDZCOM) <> "" Then xSet = xSet & ",COMRCDZCOM": xValues = xValues & " ,'" & newY.COMRCDZCOM & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YCOMRCD0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCOMRCD0_Insert = "Erreur màj : " & newY.COMRCDPIE
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCOMRCD0_Insert = "sqlYCOMRCD0_Insert " & vbCrLf & Error
End Function

Public Function sqlYCOMRCDH_Insert(newY As typeYCOMRCD0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCOMRCDH_Insert = Null

xSet = " (COMRCDNAT,COMRCDPIE,COMRCDECR,COMRCDYVER"
xValues = " values('" & newY.COMRCDNAT & "'," & newY.COMRCDPIE & "," & newY.COMRCDECR & "," & newY.COMRCDYVER

' Détecter les modifications
'===================================================================================
If newY.COMRCDRLV <> 0 Then xSet = xSet & ",COMRCDRLV": xValues = xValues & " ," & newY.COMRCDRLV
If newY.COMRCDNUM <> 0 Then xSet = xSet & ",COMRCDNUM": xValues = xValues & " ," & newY.COMRCDNUM
If newY.COMRCDDTR <> 0 Then xSet = xSet & ",COMRCDDTR": xValues = xValues & " ," & newY.COMRCDDTR
If newY.COMRCDMTD <> 0 Then xSet = xSet & ",COMRCDMTD": xValues = xValues & " ," & cur_P(newY.COMRCDMTD)
If newY.COMRCDMTR <> 0 Then xSet = xSet & ",COMRCDMTR": xValues = xValues & " ," & cur_P(newY.COMRCDMTR)
If newY.COMRCDYAMJ <> 0 Then xSet = xSet & ",COMRCDYAMJ": xValues = xValues & " ," & newY.COMRCDYAMJ
If newY.COMRCDYHMS <> 0 Then xSet = xSet & ",COMRCDYHMS": xValues = xValues & " ," & newY.COMRCDYHMS

If Trim(newY.COMRCDSER) <> "" Then xSet = xSet & ",COMRCDSER": xValues = xValues & " ,'" & newY.COMRCDSER & "'"
If Trim(newY.COMRCDSSE) <> "" Then xSet = xSet & ",COMRCDSSE": xValues = xValues & " ,'" & newY.COMRCDSSE & "'"
If Trim(newY.COMRCDDEV) <> "" Then xSet = xSet & ",COMRCDDEV": xValues = xValues & " ,'" & newY.COMRCDDEV & "'"
If Trim(newY.COMRCDOPE) <> "" Then xSet = xSet & ",COMRCDOPE": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDOPE), "'", "''") & "'"
If Trim(newY.COMRCDPCI) <> "" Then xSet = xSet & ",COMRCDPCI": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDPCI), "'", "''") & "'"
If Trim(newY.COMRCDYUSR) <> "" Then xSet = xSet & ",COMRCDYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDYUSR), "'", "''") & "'"
If Trim(newY.COMRCDSTA) <> "" Then xSet = xSet & ",COMRCDSTA": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDSTA), "'", "''") & "'"
If Trim(newY.COMRCDCLI) <> "" Then xSet = xSet & ",COMRCDCLI": xValues = xValues & " ,'" & Replace(Trim(newY.COMRCDCLI), "'", "''") & "'"

If Trim(newY.COMRCDZTYP) <> "" Then xSet = xSet & ",COMRCDZTYP": xValues = xValues & " ,'" & newY.COMRCDZTYP & "'"
If Trim(newY.COMRCDZORD) <> "" Then xSet = xSet & ",COMRCDZORD": xValues = xValues & " ,'" & newY.COMRCDZORD & "'"
If Trim(newY.COMRCDZCOM) <> "" Then xSet = xSet & ",COMRCDZCOM": xValues = xValues & " ,'" & newY.COMRCDZCOM & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YCOMRCDH" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCOMRCDH_Insert = "Erreur màj : " & xSQL
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCOMRCDH_Insert = "sqlYCOMRCDH_Insert " & vbCrLf & Error
End Function


Public Function sqlYCOMRCD0_Update(newY As typeYCOMRCD0, oldY As typeYCOMRCD0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYCOMRCD0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.COMRCDPIE <> newY.COMRCDPIE Then
    sqlYCOMRCD0_Update = "Erreur COMRCDPIE : " & newY.COMRCDPIE & " / " & oldY.COMRCDPIE
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where COMRCDNAT = '" & oldY.COMRCDNAT & "'" _
       & " and COMRCDPIE = " & oldY.COMRCDPIE _
       & " and COMRCDECR = " & oldY.COMRCDECR _
       & " and COMRCDYVER = " & oldY.COMRCDYVER
       
newY.COMRCDYVER = newY.COMRCDYVER + 1
xSet = xSet & " set COMRCDYVER = " & newY.COMRCDYVER
newY.COMRCDYAMJ = DSys
newY.COMRCDYHMS = time_Hms
newY.COMRCDYUSR = usrName_UCase

blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.COMRCDRLV <> oldY.COMRCDRLV Then blnUpdate = True: xSet = xSet & " , COMRCDRLV = " & newY.COMRCDRLV
If newY.COMRCDNUM <> oldY.COMRCDNUM Then blnUpdate = True: xSet = xSet & " , COMRCDNUM = " & newY.COMRCDNUM
If newY.COMRCDDTR <> oldY.COMRCDDTR Then blnUpdate = True: xSet = xSet & " , COMRCDDTR = " & newY.COMRCDDTR
If newY.COMRCDMTD <> oldY.COMRCDMTD Then blnUpdate = True: xSet = xSet & " , COMRCDMTD = " & cur_P(newY.COMRCDMTD)
If newY.COMRCDMTR <> oldY.COMRCDMTR Then blnUpdate = True: xSet = xSet & " , COMRCDMTR = " & cur_P(newY.COMRCDMTR)
If newY.COMRCDYAMJ <> oldY.COMRCDYAMJ Then blnUpdate = True: xSet = xSet & " , COMRCDYAMJ = " & newY.COMRCDYAMJ
If newY.COMRCDYHMS <> oldY.COMRCDYHMS Then blnUpdate = True: xSet = xSet & " , COMRCDYHMS = " & newY.COMRCDYHMS

If newY.COMRCDSER <> oldY.COMRCDSER Then blnUpdate = True:  xSet = xSet & " , COMRCDSER = '" & Replace(Trim(newY.COMRCDSER), "'", "''") & "'"
If newY.COMRCDSSE <> oldY.COMRCDSSE Then blnUpdate = True:  xSet = xSet & " , COMRCDSSE = '" & Replace(Trim(newY.COMRCDSSE), "'", "''") & "'"
If newY.COMRCDDEV <> oldY.COMRCDDEV Then blnUpdate = True:  xSet = xSet & " , COMRCDDEV= '" & newY.COMRCDDEV & "'"
If newY.COMRCDOPE <> oldY.COMRCDOPE Then blnUpdate = True:  xSet = xSet & " , COMRCDOPE = '" & Replace(Trim(newY.COMRCDOPE), "'", "''") & "'"
If newY.COMRCDPCI <> oldY.COMRCDPCI Then blnUpdate = True:  xSet = xSet & " , COMRCDPCI = '" & Replace(Trim(newY.COMRCDPCI), "'", "''") & "'"
If newY.COMRCDYUSR <> oldY.COMRCDYUSR Then blnUpdate = True:  xSet = xSet & " , COMRCDYUSR = '" & Replace(Trim(newY.COMRCDYUSR), "'", "''") & "'"
If newY.COMRCDSTA <> oldY.COMRCDSTA Then blnUpdate = True:  xSet = xSet & " , COMRCDSTA = '" & Replace(Trim(newY.COMRCDSTA), "'", "''") & "'"
If newY.COMRCDCLI <> oldY.COMRCDCLI Then blnUpdate = True:  xSet = xSet & " , COMRCDCLI = '" & Replace(Trim(newY.COMRCDCLI), "'", "''") & "'"

If newY.COMRCDZTYP <> oldY.COMRCDZTYP Then blnUpdate = True:  xSet = xSet & " , COMRCDZTYP = '" & Replace(Trim(newY.COMRCDZTYP), "'", "''") & "'"
If newY.COMRCDZORD <> oldY.COMRCDZORD Then blnUpdate = True:  xSet = xSet & " , COMRCDZORD = '" & Replace(Trim(newY.COMRCDZORD), "'", "''") & "'"
If newY.COMRCDZCOM <> oldY.COMRCDZCOM Then blnUpdate = True:  xSet = xSet & " , COMRCDZCOM = '" & Replace(Trim(newY.COMRCDZCOM), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YCOMRCD0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYCOMRCD0_Update = "Erreur màj : " & xSQL
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYCOMRCD0_Update = "sqlYCOMRCD0_Update " & vbCrLf & Error
End Function
Public Function sqlYCOMRCD0_Delete(oldY As typeYCOMRCD0)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYCOMRCD0_Delete = Null

xSQL = "delete from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
       & " where COMRCDNAT = '" & oldY.COMRCDNAT & "'" _
       & " and COMRCDPIE = " & oldY.COMRCDPIE _
       & " and COMRCDECR = " & oldY.COMRCDECR _
       & " and COMRCDYVER = " & oldY.COMRCDYVER
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
If Nb = 0 Then
    sqlYCOMRCD0_Delete = xSQL
    Exit Function
End If
    


Exit Function
Error_Handler:
    sqlYCOMRCD0_Delete = "sqlYCOMRCD0_Delete " & vbCrLf & Error
End Function

Public Function sqlYCOMRCD0_Update_CMD(lSQL As String)
Dim X As String, Nb As Long

On Error GoTo Error_Handler
sqlYCOMRCD0_Update_CMD = Null

Set rsAdo = cnSab_Update.Execute(lSQL, Nb)

'If Nb = 0 Then
'    sqlYCOMRCD0_Update_CMD = "" 'lSQL
'    Exit Function
'End If
    


Exit Function
Error_Handler:
    sqlYCOMRCD0_Update_CMD = "sqlYCOMRCD0_Update_CMD " & vbCrLf & Error
End Function

Public Function rsYCOMRCD0_GetBuffer(rsAdo As ADODB.Recordset, lYCOMRCD0 As typeYCOMRCD0)
On Error GoTo Error_Handler
rsYCOMRCD0_GetBuffer = Null
lYCOMRCD0.COMRCDNAT = rsAdo("COMRCDNAT")
lYCOMRCD0.COMRCDPIE = rsAdo("COMRCDPIE")
lYCOMRCD0.COMRCDSER = Trim(rsAdo("COMRCDSER"))
lYCOMRCD0.COMRCDCLI = Trim(rsAdo("COMRCDCLI"))
lYCOMRCD0.COMRCDECR = rsAdo("COMRCDECR")
lYCOMRCD0.COMRCDSSE = Trim(rsAdo("COMRCDSSE"))
lYCOMRCD0.COMRCDOPE = rsAdo("COMRCDOPE")
lYCOMRCD0.COMRCDNUM = rsAdo("COMRCDNUM")
lYCOMRCD0.COMRCDPCI = Trim(rsAdo("COMRCDPCI"))
lYCOMRCD0.COMRCDDEV = rsAdo("COMRCDDEV")
lYCOMRCD0.COMRCDDTR = rsAdo("COMRCDDTR")
lYCOMRCD0.COMRCDMTD = rsAdo("COMRCDMTD")
lYCOMRCD0.COMRCDMTR = rsAdo("COMRCDMTR")
lYCOMRCD0.COMRCDYUSR = Trim(rsAdo("COMRCDYUSR"))
lYCOMRCD0.COMRCDSTA = Trim(rsAdo("COMRCDSTA"))
lYCOMRCD0.COMRCDYAMJ = rsAdo("COMRCDYAMJ")

lYCOMRCD0.COMRCDYHMS = rsAdo("COMRCDYHMS")
lYCOMRCD0.COMRCDYVER = rsAdo("COMRCDYVER")
lYCOMRCD0.COMRCDRLV = rsAdo("COMRCDRLV")

lYCOMRCD0.COMRCDZTYP = rsAdo("COMRCDZTYP")
lYCOMRCD0.COMRCDZORD = rsAdo("COMRCDZORD")
lYCOMRCD0.COMRCDZCOM = rsAdo("COMRCDZCOM")

Exit Function
Error_Handler:
rsYCOMRCD0_GetBuffer = Error


End Function

Public Function rsYCOMRCD0_Init(lYCOMRCD0 As typeYCOMRCD0)
lYCOMRCD0.COMRCDPIE = 0
lYCOMRCD0.COMRCDYVER = 0
lYCOMRCD0.COMRCDNAT = ""
lYCOMRCD0.COMRCDSER = ""
lYCOMRCD0.COMRCDCLI = ""
lYCOMRCD0.COMRCDECR = 0
lYCOMRCD0.COMRCDSSE = ""
lYCOMRCD0.COMRCDOPE = ""
lYCOMRCD0.COMRCDNUM = 0
lYCOMRCD0.COMRCDDTR = 0
lYCOMRCD0.COMRCDMTD = 0
lYCOMRCD0.COMRCDPCI = ""
lYCOMRCD0.COMRCDMTR = 0
lYCOMRCD0.COMRCDYAMJ = 0
lYCOMRCD0.COMRCDSTA = ""
lYCOMRCD0.COMRCDYUSR = ""
lYCOMRCD0.COMRCDYHMS = 0
lYCOMRCD0.COMRCDDEV = ""
lYCOMRCD0.COMRCDRLV = 0
lYCOMRCD0.COMRCDZTYP = ""
lYCOMRCD0.COMRCDZORD = ""
lYCOMRCD0.COMRCDZCOM = ""

End Function







