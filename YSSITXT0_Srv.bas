Attribute VB_Name = "srvYSSITXT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSITXT0
    SSITXTNAT    As String
    SSITXTUIDN   As Long
    SSITXTDIDX   As String
    SSITXTUIDX   As String
    SSITXTUIDD   As Long
    SSITXTTLNK   As Long
    SSITXTYUSR   As String
    SSITXTYAMJ   As Long
    SSITXTYHMS   As Long
    SSITXTYVER   As Long
    SSITXTINFO   As String
End Type
Public xYSSITXT0 As typeYSSITXT0
Public Function sqlYSSITXT0_Delete(oldY As typeYSSITXT0)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYSSITXT0_Delete = Null

xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSITXT0" _
       & " where SSITXTNAT = '" & oldY.SSITXTNAT & "'" _
       & " and SSITXTUIDN = " & oldY.SSITXTUIDN _
       & " and SSITXTDIDX = '" & oldY.SSITXTDIDX & "'" _
       & " and SSITXTUIDX = '" & oldY.SSITXTUIDX & "'" _
       & " and SSITXTUIDD = " & oldY.SSITXTUIDD _
       & " and SSITXTTLNK = " & oldY.SSITXTTLNK
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
If Nb = 0 Then
    sqlYSSITXT0_Delete = xSQL
    Exit Function
End If
    


Exit Function
Error_Handler:
    sqlYSSITXT0_Delete = "sqlYSSITXT0_Delete " & vbCrLf & Error
End Function


Public Function sqlYSSITXT0_Insert(newY As typeYSSITXT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String
Dim wSSITXTUIDX As String

On Error GoTo Error_Handler
sqlYSSITXT0_Insert = Null

If Len(newY.SSITXTUIDX) > 20 Then
   wSSITXTUIDX = Mid$(newY.SSITXTUIDX, 1, 20)
Else
   wSSITXTUIDX = newY.SSITXTUIDX
End If


xSet = " (SSITXTNAT,SSITXTUIDN,SSITXTDIDX,SSITXTUIDX,SSITXTUIDD,SSITXTTLNK"
xValues = " values('" & newY.SSITXTNAT & "'," & newY.SSITXTUIDN & ",'" & newY.SSITXTDIDX & "','" & Replace(wSSITXTUIDX, "'", "''") & "'," & newY.SSITXTUIDD & "," & newY.SSITXTTLNK

' Détecter les modifications
'===================================================================================
If newY.SSITXTYVER <> 0 Then xSet = xSet & ",SSITXTYVER": xValues = xValues & " ," & newY.SSITXTYVER
If newY.SSITXTYAMJ <> 0 Then xSet = xSet & ",SSITXTYAMJ": xValues = xValues & " ," & newY.SSITXTYAMJ
If newY.SSITXTYHMS <> 0 Then xSet = xSet & ",SSITXTYHMS": xValues = xValues & " ," & newY.SSITXTYHMS

If Trim(newY.SSITXTYUSR) <> "" Then xSet = xSet & ",SSITXTYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.SSITXTYUSR), "'", "''") & "'"
If Trim(newY.SSITXTINFO) <> "" Then xSet = xSet & ",SSITXTINFO": xValues = xValues & " ,'" & Replace(Trim(newY.SSITXTINFO), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSITXT0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSSITXT0_Insert = "Erreur màj : " & newY.SSITXTUIDN
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSSITXT0_Insert = "sqlYSSITXT0_Insert " & vbCrLf & Error
End Function

Public Function sqlYSSITXT0_Update(newY As typeYSSITXT0, oldY As typeYSSITXT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSITXT0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSITXTUIDN <> newY.SSITXTUIDN Then
    sqlYSSITXT0_Update = "Erreur SSITXTUIDN : " & newY.SSITXTUIDN & " / " & oldY.SSITXTUIDN
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSITXTNAT = '" & oldY.SSITXTNAT & "'" _
       & " and SSITXTUIDN = " & oldY.SSITXTUIDN _
       & " and SSITXTDIDX = '" & oldY.SSITXTDIDX & "'" _
       & " and SSITXTUIDX = '" & Replace(oldY.SSITXTUIDX, "'", "''") & "'" _
       & " and SSITXTTLNK = " & oldY.SSITXTTLNK _
       & " and SSITXTUIDD = " & oldY.SSITXTUIDD _
       & " and SSITXTYVER = " & oldY.SSITXTYVER
       
newY.SSITXTYVER = newY.SSITXTYVER + 1
xSet = xSet & " set SSITXTYVER = " & newY.SSITXTYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSITXTUIDD <> oldY.SSITXTUIDD Then blnUpdate = True: xSet = xSet & " , SSITXTUIDD = " & newY.SSITXTUIDD
If newY.SSITXTYAMJ <> oldY.SSITXTYAMJ Then blnUpdate = True: xSet = xSet & " , SSITXTYAMJ = " & newY.SSITXTYAMJ
If newY.SSITXTYHMS <> oldY.SSITXTYHMS Then blnUpdate = True: xSet = xSet & " , SSITXTYHMS = " & newY.SSITXTYHMS

If newY.SSITXTINFO <> oldY.SSITXTINFO Then blnUpdate = True:  xSet = xSet & " , SSITXTINFO= '" & Replace(Trim(newY.SSITXTINFO), "'", "''") & "'"
If newY.SSITXTYUSR <> oldY.SSITXTYUSR Then blnUpdate = True:  xSet = xSet & " , SSITXTYUSR = '" & Replace(Trim(newY.SSITXTYUSR), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSITXT0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSITXT0_Update = "Erreur màj : " & newY.SSITXTUIDN
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSITXT0_Update = "sqlYSSITXT0_Update " & vbCrLf & Error
End Function

Public Function rsYSSITXT0_GetBuffer(rsAdo As ADODB.Recordset, lYSSITXT0 As typeYSSITXT0)
On Error GoTo Error_Handler
rsYSSITXT0_GetBuffer = Null
lYSSITXT0.SSITXTNAT = rsAdo("SSITXTNAT")
lYSSITXT0.SSITXTUIDN = rsAdo("SSITXTUIDN")
lYSSITXT0.SSITXTDIDX = Trim(rsAdo("SSITXTDIDX"))
lYSSITXT0.SSITXTUIDX = rsAdo("SSITXTUIDX")
lYSSITXT0.SSITXTUIDD = rsAdo("SSITXTUIDD")
lYSSITXT0.SSITXTINFO = Trim(rsAdo("SSITXTINFO"))
lYSSITXT0.SSITXTTLNK = rsAdo("SSITXTTLNK")
lYSSITXT0.SSITXTYUSR = Trim(rsAdo("SSITXTYUSR"))
lYSSITXT0.SSITXTYAMJ = rsAdo("SSITXTYAMJ")

lYSSITXT0.SSITXTYHMS = rsAdo("SSITXTYHMS")
lYSSITXT0.SSITXTYVER = rsAdo("SSITXTYVER")

Exit Function
Error_Handler:
rsYSSITXT0_GetBuffer = Error


End Function

Public Function rsYSSITXT0_Init(lYSSITXT0 As typeYSSITXT0)
lYSSITXT0.SSITXTUIDN = 0
lYSSITXT0.SSITXTYVER = 0
lYSSITXT0.SSITXTNAT = ""
lYSSITXT0.SSITXTDIDX = ""
lYSSITXT0.SSITXTUIDX = ""
lYSSITXT0.SSITXTINFO = ""
lYSSITXT0.SSITXTUIDD = 0
lYSSITXT0.SSITXTTLNK = 0
lYSSITXT0.SSITXTYAMJ = 0
lYSSITXT0.SSITXTYUSR = ""
lYSSITXT0.SSITXTYHMS = 0

End Function








