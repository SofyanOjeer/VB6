Attribute VB_Name = "srvYROPINF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYROPINF0
 
      ROPINFID     As Long         'identification
      ROPINFIDP    As Long         'sequence processus
      ROPINFIDT    As Long         'sequence tâche
      ROPINFIDT2   As Long         'sequence suivi
      ROPINFIDTL   As Long         'sequence tâche liée
      ROPINFMAIL   As String * 5   'codes envoi mail
      ROPINFSTA    As String * 1   'état du dossier
      ROPINFSTAK    As String * 1   'état alerte
      ROPINFSTAD    As String * 1   'état du dossier
      ROPINFCUSR   As String * 12  'utilisateur création
      ROPINFCAMJ   As String * 8   'date création
      ROPINFUUSR   As String * 12  'utilisateur màj
      ROPINFUAMJ   As String * 8   'date maàj
      ROPINFUHMS   As String * 6   'heure màj
      ROPINFUVER   As Long         'version
      ROPINFGUO    As Long      'unité d'oeuvre
      ROPINFGECH   As String * 8   'échéeance
      ROPINFGUSR   As String * 12  'gestionnaire responsable
      ROPINFGSRV   As String * 4   'service gestionnaire responsable
      ROPINFGNAT   As String * 1   'nature
      ROPINFGPRV   As String * 1   'confidentialité
      ROPINFGTXT   As String       'memo

End Type
Public xYROPINF0 As typeYROPINF0
Public Function sqlYROPINF0_Delete(oldY As typeYROPINF0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYROPINF0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where ROPINFID = " & oldY.ROPINFID _
       & " and ROPINFIDP  = " & oldY.ROPINFIDP _
       & " and ROPINFIDT  = " & oldY.ROPINFIDT _
       & " and ROPINFIDT2  = " & oldY.ROPINFIDT2 _
       & " and ROPINFUVER = " & oldY.ROPINFUVER

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YROPINF0" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYROPINF0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYROPINF0_Delete = Error
End Function

Public Function sqlYROPINF0_Delete_GE(oldY As typeYROPINF0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYROPINF0_Delete_GE = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where ROPINFID = " & oldY.ROPINFID _
       & " and ROPINFIDP  = " & oldY.ROPINFIDP _
       & " and ROPINFIDT  >= " & oldY.ROPINFIDT

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YROPINF0" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYROPINF0_Delete_GE = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYROPINF0_Delete_GE = Error
End Function

Public Function sqlYROPINF0_Update(newY As typeYROPINF0, oldY As typeYROPINF0, blnUUSR As Boolean)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYROPINF0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.ROPINFID <> newY.ROPINFID _
Or oldY.ROPINFIDP <> newY.ROPINFIDP _
Or oldY.ROPINFIDT <> newY.ROPINFIDT _
Or oldY.ROPINFIDT2 <> newY.ROPINFIDT2 _
Or oldY.ROPINFUVER <> newY.ROPINFUVER Then
    sqlYROPINF0_Update = "Erreur ROPINFID : " & newY.ROPINFID & "." & oldY.ROPINFIDT & "." & oldY.ROPINFIDT2 & "." & oldY.ROPINFUVER
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where ROPINFID = " & oldY.ROPINFID _
       & " and ROPINFIDP  = " & oldY.ROPINFIDP _
       & " and ROPINFIDT  = " & oldY.ROPINFIDT _
       & " and ROPINFIDT2  = " & oldY.ROPINFIDT2 _
       & " and ROPINFUVER = " & oldY.ROPINFUVER

newY.ROPINFUVER = newY.ROPINFUVER + 1
xSet = xSet & " set ROPINFUVER = " & newY.ROPINFUVER
blnUpdate = False

If blnUUSR Then
    newY.ROPINFUUSR = usrName_UCase
    newY.ROPINFUAMJ = DSys
    newY.ROPINFUHMS = time_Hms
End If

' Détecter les modifications
'===================================================================================
If newY.ROPINFMAIL <> oldY.ROPINFMAIL Then blnUpdate = True:  xSet = xSet & " , ROPINFMAIL = '" & newY.ROPINFMAIL & "'"
If newY.ROPINFSTA <> oldY.ROPINFSTA Then blnUpdate = True:  xSet = xSet & " , ROPINFSTA = '" & newY.ROPINFSTA & "'"
If newY.ROPINFSTAK <> oldY.ROPINFSTAK Then blnUpdate = True:  xSet = xSet & " , ROPINFSTAK = '" & newY.ROPINFSTAK & "'"
If newY.ROPINFSTAD <> oldY.ROPINFSTAD Then blnUpdate = True:  xSet = xSet & " , ROPINFSTAD = '" & newY.ROPINFSTAD & "'"
If newY.ROPINFCUSR <> oldY.ROPINFCUSR Then blnUpdate = True:  xSet = xSet & " , ROPINFCUSR = '" & newY.ROPINFCUSR & "'"
If newY.ROPINFCAMJ <> oldY.ROPINFCAMJ Then blnUpdate = True:  xSet = xSet & " , ROPINFCAMJ = '" & newY.ROPINFCAMJ & "'"
If newY.ROPINFUUSR <> oldY.ROPINFUUSR Then blnUpdate = True:  xSet = xSet & " , ROPINFUUSR = '" & newY.ROPINFUUSR & "'"
If newY.ROPINFUAMJ <> oldY.ROPINFUAMJ Then blnUpdate = True:  xSet = xSet & " , ROPINFUAMJ = '" & newY.ROPINFUAMJ & "'"
If newY.ROPINFUHMS <> oldY.ROPINFUHMS Then blnUpdate = True:  xSet = xSet & " , ROPINFUHMS = '" & newY.ROPINFUHMS & "'"
If newY.ROPINFGECH <> oldY.ROPINFGECH Then blnUpdate = True:  xSet = xSet & " , ROPINFGECH = '" & newY.ROPINFGECH & "'"
If newY.ROPINFGUSR <> oldY.ROPINFGUSR Then blnUpdate = True:  xSet = xSet & " , ROPINFGUSR = '" & newY.ROPINFGUSR & "'"
If newY.ROPINFGSRV <> oldY.ROPINFGSRV Then blnUpdate = True:  xSet = xSet & " , ROPINFGSRV = '" & newY.ROPINFGSRV & "'"
If newY.ROPINFGNAT <> oldY.ROPINFGNAT Then blnUpdate = True:  xSet = xSet & " , ROPINFGNAT = '" & newY.ROPINFGNAT & "'"
If newY.ROPINFGPRV <> oldY.ROPINFGPRV Then blnUpdate = True:  xSet = xSet & " , ROPINFGPRV = '" & newY.ROPINFGPRV & "'"
If newY.ROPINFGTXT <> oldY.ROPINFGTXT Then blnUpdate = True:  xSet = xSet & " , ROPINFGTXT = '" & Replace(Trim(newY.ROPINFGTXT), "'", "''") & "'"
If newY.ROPINFGUO <> oldY.ROPINFGUO Then blnUpdate = True:  xSet = xSet & ",ROPINFGUO = " & newY.ROPINFGUO
If newY.ROPINFIDTL <> oldY.ROPINFIDTL Then blnUpdate = True:  xSet = xSet & ",ROPINFIDTL = " & newY.ROPINFIDTL

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YROPINF0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYROPINF0_Update = "Erreur màj : " & newY.ROPINFID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYROPINF0_Update = Error
End Function

Public Function sqlYROPINF0_Requête(xFct As String, xSet As String, xWhere As String)
Dim X As String, xSQL As String, Nb As Long
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYROPINF0_Requête = Null

    
xSQL = xFct & paramIBM_Library_SABSPE & ".YROPINF0" & xSet & xWhere
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT


Exit Function
Error_Handler:
    sqlYROPINF0_Requête = Error
End Function

Public Function sqlYROPINF0_Insert(newY As typeYROPINF0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYROPINF0_Insert = Null

xSet = " (ROPINFID"
xValues = " values(" & newY.ROPINFID

newY.ROPINFCUSR = usrName_UCase
newY.ROPINFCAMJ = DSys

newY.ROPINFUUSR = usrName_UCase
newY.ROPINFUAMJ = DSys
newY.ROPINFUHMS = time_Hms

' Détecter les modifications
'===================================================================================
If newY.ROPINFIDP <> 0 Then xSet = xSet & ",ROPINFIDP": xValues = xValues & " ," & newY.ROPINFIDP
If newY.ROPINFIDT <> 0 Then xSet = xSet & ",ROPINFIDT": xValues = xValues & " ," & newY.ROPINFIDT
If newY.ROPINFIDT2 <> 0 Then xSet = xSet & ",ROPINFIDT2": xValues = xValues & " ," & newY.ROPINFIDT2
If newY.ROPINFIDTL <> 0 Then xSet = xSet & ",ROPINFIDTL": xValues = xValues & " ," & newY.ROPINFIDTL
If newY.ROPINFUVER <> 0 Then xSet = xSet & ",ROPINFUVER": xValues = xValues & " ," & newY.ROPINFUVER
If newY.ROPINFGUO <> 0 Then xSet = xSet & ",ROPINFGUO": xValues = xValues & " ," & newY.ROPINFGUO

If Trim(newY.ROPINFMAIL) <> "" Then xSet = xSet & ",ROPINFMAIL": xValues = xValues & " ,'" & newY.ROPINFMAIL & "'"
If Trim(newY.ROPINFSTA) <> "" Then xSet = xSet & ",ROPINFSTA": xValues = xValues & " ,'" & newY.ROPINFSTA & "'"
If Trim(newY.ROPINFSTAK) <> "" Then xSet = xSet & ",ROPINFSTAK": xValues = xValues & " ,'" & newY.ROPINFSTAK & "'"
If Trim(newY.ROPINFSTAD) <> "" Then xSet = xSet & ",ROPINFSTAD": xValues = xValues & " ,'" & newY.ROPINFSTAD & "'"
If Trim(newY.ROPINFCUSR) <> "" Then xSet = xSet & ",ROPINFCUSR": xValues = xValues & " ,'" & newY.ROPINFCUSR & "'"
If Trim(newY.ROPINFCAMJ) <> "" Then xSet = xSet & ",ROPINFCAMJ": xValues = xValues & " ,'" & newY.ROPINFCAMJ & "'"
If Trim(newY.ROPINFUUSR) <> "" Then xSet = xSet & ",ROPINFUUSR": xValues = xValues & " ,'" & newY.ROPINFUUSR & "'"
If Trim(newY.ROPINFUAMJ) <> "" Then xSet = xSet & ",ROPINFUAMJ": xValues = xValues & " ,'" & newY.ROPINFUAMJ & "'"
If Trim(newY.ROPINFUHMS) <> "" Then xSet = xSet & ",ROPINFUHMS": xValues = xValues & " ,'" & newY.ROPINFUHMS & "'"
If Trim(newY.ROPINFGECH) <> "" Then xSet = xSet & ",ROPINFGECH": xValues = xValues & " ,'" & newY.ROPINFGECH & "'"
If Trim(newY.ROPINFGUSR) <> "" Then xSet = xSet & ",ROPINFGUSR": xValues = xValues & " ,'" & newY.ROPINFGUSR & "'"
If Trim(newY.ROPINFGSRV) <> "" Then xSet = xSet & ",ROPINFGSRV": xValues = xValues & " ,'" & newY.ROPINFGSRV & "'"
If Trim(newY.ROPINFGNAT) <> "" Then xSet = xSet & ",ROPINFGNAT": xValues = xValues & " ,'" & newY.ROPINFGNAT & "'"
If Trim(newY.ROPINFGPRV) <> "" Then xSet = xSet & ",ROPINFGPRV": xValues = xValues & " ,'" & newY.ROPINFGPRV & "'"
If Trim(newY.ROPINFGTXT) <> "" Then xSet = xSet & ",ROPINFGTXT": xValues = xValues & " ,'" & Replace(Trim(newY.ROPINFGTXT), "'", "''") & "'"

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YROPINF0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYROPINF0_Insert = "Erreur màj : " & newY.ROPINFID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYROPINF0_Insert = Error
End Function

Public Function rsYROPINF0_GetBuffer(rsAdo As ADODB.Recordset, lYROPINF0 As typeYROPINF0)
On Error GoTo Error_Handler
rsYROPINF0_GetBuffer = Null

lYROPINF0.ROPINFID = rsAdo("ROPINFID")
lYROPINF0.ROPINFIDP = rsAdo("ROPINFIDP")
lYROPINF0.ROPINFIDT = rsAdo("ROPINFIDT")
lYROPINF0.ROPINFIDT2 = rsAdo("ROPINFIDT2")
lYROPINF0.ROPINFIDTL = rsAdo("ROPINFIDTL")
lYROPINF0.ROPINFMAIL = rsAdo("ROPINFMAIL")
lYROPINF0.ROPINFSTA = rsAdo("ROPINFSTA")
lYROPINF0.ROPINFSTAK = rsAdo("ROPINFSTAK")
lYROPINF0.ROPINFSTAD = rsAdo("ROPINFSTAD")
lYROPINF0.ROPINFCUSR = rsAdo("ROPINFCUSR")
lYROPINF0.ROPINFCAMJ = rsAdo("ROPINFCAMJ")
lYROPINF0.ROPINFUUSR = rsAdo("ROPINFUUSR")
lYROPINF0.ROPINFUAMJ = rsAdo("ROPINFUAMJ")
lYROPINF0.ROPINFUHMS = rsAdo("ROPINFUHMS")
lYROPINF0.ROPINFUVER = rsAdo("ROPINFUVER")
lYROPINF0.ROPINFGUO = rsAdo("ROPINFGUO")
lYROPINF0.ROPINFGECH = rsAdo("ROPINFGECH")
lYROPINF0.ROPINFGUSR = rsAdo("ROPINFGUSR")
lYROPINF0.ROPINFGSRV = rsAdo("ROPINFGSRV")
lYROPINF0.ROPINFGNAT = rsAdo("ROPINFGNAT")
lYROPINF0.ROPINFGPRV = rsAdo("ROPINFGPRV")
lYROPINF0.ROPINFGTXT = rsAdo("ROPINFGTXT")

Exit Function
Error_Handler:
rsYROPINF0_GetBuffer = Error


End Function

Public Function rsYROPINF0_Init(lYROPINF0 As typeYROPINF0)

lYROPINF0.ROPINFID = 0
lYROPINF0.ROPINFIDP = 1
lYROPINF0.ROPINFIDT = 0
lYROPINF0.ROPINFIDT2 = 1
lYROPINF0.ROPINFIDTL = 0
lYROPINF0.ROPINFMAIL = "G"
lYROPINF0.ROPINFSTA = " "
lYROPINF0.ROPINFSTAK = " "
lYROPINF0.ROPINFSTAD = " "
lYROPINF0.ROPINFCUSR = usrName_UCase
lYROPINF0.ROPINFCAMJ = DSys
lYROPINF0.ROPINFUUSR = usrName_UCase
lYROPINF0.ROPINFGSRV = ""
lYROPINF0.ROPINFUAMJ = DSys
lYROPINF0.ROPINFUHMS = time_Hms
lYROPINF0.ROPINFUVER = 0
lYROPINF0.ROPINFGUO = 0
lYROPINF0.ROPINFGECH = DSys
lYROPINF0.ROPINFGUSR = "" 'usrName_UCase
lYROPINF0.ROPINFGNAT = "I"
lYROPINF0.ROPINFGPRV = " "
lYROPINF0.ROPINFGTXT = ""

End Function



