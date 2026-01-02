Attribute VB_Name = "srvYROPDOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public paramROPDOS_Path As String

Dim rsAdo As ADODB.Recordset
 
Type typeYROPDOS0
 
      ROPDOSID     As Long         'identification
      ROPDOSSTA    As String * 1   'état du dossier
      ROPDOSSTAK   As String * 1   'état alerte
      ROPDOSCUSR   As String * 12  'utilisateur création
      ROPDOSCAMJ   As String * 8   'date création
      ROPDOSUUSR   As String * 12  'utilisateur màj
      ROPDOSUAMJ   As String * 8   'date maàj
      ROPDOSUHMS   As String * 6   'heure màj
      ROPDOSUVER   As Long         'version
      ROPDOSGECH   As String * 8   'échéeance
      ROPDOSGUSR   As String * 12  'gestionnaire responsable
      ROPDOSGSRV   As String * 4   'service gestionnaire responsable
      ROPDOSGNAT   As String * 1   'nature I D E M
      ROPDOSGPRV   As String * 1   'confidentialité
      ROPDOSGGRA   As String * 1   'gravité
      ROPDOSGPRI   As String * 1   'priorité
      ROPDOSGCOU   As Long         'coût en K €
      ROPDOSIAMJ   As String * 8   'date du constat
      ROPDOSISRV   As String * 4   'service initiateur
      ROPDOSIUSR   As String * 12  'utilisateur initiateur
      ROPDOSIREF   As String * 20  'référence de l'opération
      ROPDOSXDOM   As String * 12  'domaine
      ROPDOSXAPP   As String * 12  'application
      ROPDOSXID    As String * 20  'référence prestataire
      ROPDOSQUAL   As String * 3   'qualification
End Type
Public xYROPDOS0 As typeYROPDOS0
Public Function sqlYROPDOS0_Delete(oldY As typeYROPDOS0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYROPDOS0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where ROPDOSID = " & oldY.ROPDOSID _
       & " and ROPDOSUVER = " & oldY.ROPDOSUVER

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YROPDOS0" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYROPDOS0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYROPDOS0_Delete = Error
End Function

Public Function sqlYROPDOS0_Update(newY As typeYROPDOS0, oldY As typeYROPDOS0, blnUUSR As Boolean)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYROPDOS0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.ROPDOSID <> newY.ROPDOSID _
Or oldY.ROPDOSUVER <> newY.ROPDOSUVER Then
    sqlYROPDOS0_Update = "Erreur ROPDOSID : " & newY.ROPDOSID & "." & oldY.ROPDOSUVER
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where ROPDOSID = " & oldY.ROPDOSID _
       & " and ROPDOSUVER = " & oldY.ROPDOSUVER

newY.ROPDOSUVER = newY.ROPDOSUVER + 1
xSet = xSet & " set ROPDOSUVER = " & newY.ROPDOSUVER
blnUpdate = False

If blnUUSR Then
    newY.ROPDOSUUSR = usrName_UCase
    newY.ROPDOSUAMJ = DSys
    newY.ROPDOSUHMS = time_Hms
End If

' Détecter les modifications
'===================================================================================
If newY.ROPDOSSTA <> oldY.ROPDOSSTA Then blnUpdate = True:  xSet = xSet & " , ROPDOSSTA = '" & newY.ROPDOSSTA & "'"
If newY.ROPDOSSTAK <> oldY.ROPDOSSTAK Then blnUpdate = True:  xSet = xSet & " , ROPDOSSTAK = '" & newY.ROPDOSSTAK & "'"
If newY.ROPDOSCUSR <> oldY.ROPDOSCUSR Then blnUpdate = True:  xSet = xSet & " , ROPDOSCUSR = '" & newY.ROPDOSCUSR & "'"
If newY.ROPDOSCAMJ <> oldY.ROPDOSCAMJ Then blnUpdate = True:  xSet = xSet & " , ROPDOSCAMJ = '" & newY.ROPDOSCAMJ & "'"
If newY.ROPDOSUUSR <> oldY.ROPDOSUUSR Then blnUpdate = True:  xSet = xSet & " , ROPDOSUUSR = '" & newY.ROPDOSUUSR & "'"
If newY.ROPDOSUAMJ <> oldY.ROPDOSUAMJ Then blnUpdate = True:  xSet = xSet & " , ROPDOSUAMJ = '" & newY.ROPDOSUAMJ & "'"
If newY.ROPDOSUHMS <> oldY.ROPDOSUHMS Then blnUpdate = True:  xSet = xSet & " , ROPDOSUHMS = '" & newY.ROPDOSUHMS & "'"
If newY.ROPDOSGECH <> oldY.ROPDOSGECH Then blnUpdate = True:  xSet = xSet & " , ROPDOSGECH = '" & newY.ROPDOSGECH & "'"
If newY.ROPDOSGUSR <> oldY.ROPDOSGUSR Then blnUpdate = True:  xSet = xSet & " , ROPDOSGUSR = '" & newY.ROPDOSGUSR & "'"
If newY.ROPDOSGSRV <> oldY.ROPDOSGSRV Then blnUpdate = True:  xSet = xSet & " , ROPDOSGSRV = '" & newY.ROPDOSGSRV & "'"
If newY.ROPDOSGNAT <> oldY.ROPDOSGNAT Then blnUpdate = True:  xSet = xSet & " , ROPDOSGNAT = '" & newY.ROPDOSGNAT & "'"
If newY.ROPDOSGPRV <> oldY.ROPDOSGPRV Then blnUpdate = True:  xSet = xSet & " , ROPDOSGPRV = '" & newY.ROPDOSGPRV & "'"
If newY.ROPDOSGGRA <> oldY.ROPDOSGGRA Then blnUpdate = True:  xSet = xSet & " , ROPDOSGGRA = '" & newY.ROPDOSGGRA & "'"
If newY.ROPDOSGPRI <> oldY.ROPDOSGPRI Then blnUpdate = True:  xSet = xSet & " , ROPDOSGPRI = '" & newY.ROPDOSGPRI & "'"
If newY.ROPDOSIAMJ <> oldY.ROPDOSIAMJ Then blnUpdate = True:  xSet = xSet & " , ROPDOSIAMJ = '" & newY.ROPDOSIAMJ & "'"
If newY.ROPDOSISRV <> oldY.ROPDOSISRV Then blnUpdate = True:  xSet = xSet & " , ROPDOSISRV = '" & newY.ROPDOSISRV & "'"
If newY.ROPDOSIUSR <> oldY.ROPDOSIUSR Then blnUpdate = True:  xSet = xSet & " , ROPDOSIUSR = '" & newY.ROPDOSIUSR & "'"
If newY.ROPDOSIREF <> oldY.ROPDOSIREF Then blnUpdate = True:  xSet = xSet & " , ROPDOSIREF = '" & Replace(Trim(newY.ROPDOSIREF), "'", "''") & "'"
If newY.ROPDOSXDOM <> oldY.ROPDOSXDOM Then blnUpdate = True:  xSet = xSet & " , ROPDOSXDOM = '" & newY.ROPDOSXDOM & "'"
If newY.ROPDOSXAPP <> oldY.ROPDOSXAPP Then blnUpdate = True:  xSet = xSet & " , ROPDOSXAPP = '" & newY.ROPDOSXAPP & "'"
If newY.ROPDOSXID <> oldY.ROPDOSXID Then blnUpdate = True:  xSet = xSet & " , ROPDOSXID = '" & Replace(Trim(newY.ROPDOSXID), "'", "''") & "'"
If newY.ROPDOSGCOU <> oldY.ROPDOSGCOU Then blnUpdate = True:  xSet = xSet & " , ROPDOSGCOU = " & newY.ROPDOSGCOU
If newY.ROPDOSQUAL <> oldY.ROPDOSQUAL Then blnUpdate = True:  xSet = xSet & " , ROPDOSQUAL = '" & Replace(Trim(newY.ROPDOSQUAL), "'", "''") & "'"

If newY.ROPDOSID < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YROPDOS0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYROPDOS0_Update = "Erreur màj : " & newY.ROPDOSID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYROPDOS0_Update = Error
End Function

Public Function sqlYROPDOS0_Insert(newY As typeYROPDOS0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYROPDOS0_Insert = Null
xSet = " (ROPDOSID"
xValues = " values(" & newY.ROPDOSID

newY.ROPDOSCUSR = usrName_UCase
newY.ROPDOSCAMJ = DSys

newY.ROPDOSUUSR = usrName_UCase
newY.ROPDOSUAMJ = DSys
newY.ROPDOSUHMS = time_Hms

' Détecter les modifications
'===================================================================================
If newY.ROPDOSUVER <> 0 Then xSet = xSet & ",ROPDOSUVER": xValues = xValues & " ," & newY.ROPDOSUVER
If newY.ROPDOSGCOU <> 0 Then xSet = xSet & ",ROPDOSGCOU": xValues = xValues & " ," & newY.ROPDOSGCOU

If Trim(newY.ROPDOSSTA) <> "" Then xSet = xSet & ",ROPDOSSTA": xValues = xValues & " ,'" & newY.ROPDOSSTA & "'"
If Trim(newY.ROPDOSSTAK) <> "" Then xSet = xSet & ",ROPDOSSTAK": xValues = xValues & " ,'" & newY.ROPDOSSTAK & "'"
If Trim(newY.ROPDOSCUSR) <> "" Then xSet = xSet & ",ROPDOSCUSR": xValues = xValues & " ,'" & newY.ROPDOSCUSR & "'"
If Trim(newY.ROPDOSCAMJ) <> "" Then xSet = xSet & ",ROPDOSCAMJ": xValues = xValues & " ,'" & newY.ROPDOSCAMJ & "'"
If Trim(newY.ROPDOSUUSR) <> "" Then xSet = xSet & ",ROPDOSUUSR": xValues = xValues & " ,'" & newY.ROPDOSUUSR & "'"
If Trim(newY.ROPDOSUAMJ) <> "" Then xSet = xSet & ",ROPDOSUAMJ": xValues = xValues & " ,'" & newY.ROPDOSUAMJ & "'"
If Trim(newY.ROPDOSUHMS) <> "" Then xSet = xSet & ",ROPDOSUHMS": xValues = xValues & " ,'" & newY.ROPDOSUHMS & "'"
If Trim(newY.ROPDOSGECH) <> "" Then xSet = xSet & ",ROPDOSGECH": xValues = xValues & " ,'" & newY.ROPDOSGECH & "'"
If Trim(newY.ROPDOSGUSR) <> "" Then xSet = xSet & ",ROPDOSGUSR": xValues = xValues & " ,'" & newY.ROPDOSGUSR & "'"
If Trim(newY.ROPDOSGSRV) <> "" Then xSet = xSet & ",ROPDOSGSRV": xValues = xValues & " ,'" & newY.ROPDOSGSRV & "'"
If Trim(newY.ROPDOSGNAT) <> "" Then xSet = xSet & ",ROPDOSGNAT": xValues = xValues & " ,'" & newY.ROPDOSGNAT & "'"
If Trim(newY.ROPDOSGPRV) <> "" Then xSet = xSet & ",ROPDOSGPRV": xValues = xValues & " ,'" & newY.ROPDOSGPRV & "'"
If Trim(newY.ROPDOSGGRA) <> "" Then xSet = xSet & ",ROPDOSGGRA": xValues = xValues & " ,'" & newY.ROPDOSGGRA & "'"
If Trim(newY.ROPDOSGPRI) <> "" Then xSet = xSet & ",ROPDOSGPRI": xValues = xValues & " ,'" & newY.ROPDOSGPRI & "'"
If Trim(newY.ROPDOSIAMJ) <> "" Then xSet = xSet & ",ROPDOSIAMJ": xValues = xValues & " ,'" & newY.ROPDOSIAMJ & "'"
If Trim(newY.ROPDOSISRV) <> "" Then xSet = xSet & ",ROPDOSISRV": xValues = xValues & " ,'" & newY.ROPDOSISRV & "'"
If Trim(newY.ROPDOSIUSR) <> "" Then xSet = xSet & ",ROPDOSIUSR": xValues = xValues & " ,'" & newY.ROPDOSIUSR & "'"
If Trim(newY.ROPDOSIREF) <> "" Then xSet = xSet & ",ROPDOSIREF": xValues = xValues & " ,'" & Replace(Trim(newY.ROPDOSIREF), "'", "''") & "'"
If Trim(newY.ROPDOSXDOM) <> "" Then xSet = xSet & ",ROPDOSXDOM": xValues = xValues & " ,'" & newY.ROPDOSXDOM & "'"
If Trim(newY.ROPDOSXAPP) <> "" Then xSet = xSet & ",ROPDOSXAPP": xValues = xValues & " ,'" & newY.ROPDOSXAPP & "'"
If Trim(newY.ROPDOSXID) <> "" Then xSet = xSet & ",ROPDOSXID": xValues = xValues & " ,'" & Replace(Trim(newY.ROPDOSXID), "'", "''") & "'"
If Trim(newY.ROPDOSQUAL) <> "" Then xSet = xSet & ",ROPDOSQUAL": xValues = xValues & " ,'" & Replace(Trim(newY.ROPDOSQUAL), "'", "''") & "'"

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YROPDOS0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYROPDOS0_Insert = "Erreur màj : " & newY.ROPDOSID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYROPDOS0_Insert = Error
End Function

Public Function sqlROPDOSID_Init(lBIATABID As String, lROPDOSID As Long)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Static mROPDOSID As Long
On Error GoTo Error_Handler
sqlROPDOSID_Init = Null

xSQL = "select ROPDOSID from " & paramIBM_Library_SABSPE & ".YROPDOS0 " _
     & "  Where ROPDOSID >= " & mROPDOSID & " order by ROPDOSID desc"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    sqlROPDOSID_Init = "EOF"
Else
    mROPDOSID = rsSab("ROPDOSID")
    lROPDOSID = mROPDOSID + 1
End If

 
Exit Function
Error_Handler:
    sqlROPDOSID_Init = Error
End Function


Public Function rsYROPDOS0_GetBuffer(rsAdo As ADODB.Recordset, lYROPDOS0 As typeYROPDOS0)
On Error GoTo Error_Handler
rsYROPDOS0_GetBuffer = Null

lYROPDOS0.ROPDOSID = rsAdo("ROPDOSID")
lYROPDOS0.ROPDOSSTA = rsAdo("ROPDOSSTA")
lYROPDOS0.ROPDOSSTAK = rsAdo("ROPDOSSTAK")
lYROPDOS0.ROPDOSCUSR = rsAdo("ROPDOSCUSR")
lYROPDOS0.ROPDOSCAMJ = rsAdo("ROPDOSCAMJ")
lYROPDOS0.ROPDOSUUSR = rsAdo("ROPDOSUUSR")
lYROPDOS0.ROPDOSUAMJ = rsAdo("ROPDOSUAMJ")
lYROPDOS0.ROPDOSUHMS = rsAdo("ROPDOSUHMS")
lYROPDOS0.ROPDOSUVER = rsAdo("ROPDOSUVER")
lYROPDOS0.ROPDOSGECH = rsAdo("ROPDOSGECH")
lYROPDOS0.ROPDOSGUSR = rsAdo("ROPDOSGUSR")
lYROPDOS0.ROPDOSGSRV = rsAdo("ROPDOSGSRV")
lYROPDOS0.ROPDOSGNAT = rsAdo("ROPDOSGNAT")
lYROPDOS0.ROPDOSGPRV = rsAdo("ROPDOSGPRV")
lYROPDOS0.ROPDOSGGRA = rsAdo("ROPDOSGGRA")
lYROPDOS0.ROPDOSGCOU = rsAdo("ROPDOSGCOU")
lYROPDOS0.ROPDOSGPRI = rsAdo("ROPDOSGPRI")
lYROPDOS0.ROPDOSIAMJ = rsAdo("ROPDOSIAMJ")
lYROPDOS0.ROPDOSISRV = rsAdo("ROPDOSISRV")
lYROPDOS0.ROPDOSIUSR = rsAdo("ROPDOSIUSR")
lYROPDOS0.ROPDOSIREF = rsAdo("ROPDOSIREF")
lYROPDOS0.ROPDOSXDOM = rsAdo("ROPDOSXDOM")
lYROPDOS0.ROPDOSXAPP = rsAdo("ROPDOSXAPP")
lYROPDOS0.ROPDOSXID = rsAdo("ROPDOSXID")
lYROPDOS0.ROPDOSQUAL = rsAdo("ROPDOSQUAL")

Exit Function
Error_Handler:
rsYROPDOS0_GetBuffer = Error


End Function

Public Function rsYROPDOS0_Init(lYROPDOS0 As typeYROPDOS0)

lYROPDOS0.ROPDOSID = 0
lYROPDOS0.ROPDOSSTA = " "
lYROPDOS0.ROPDOSSTAK = " "
lYROPDOS0.ROPDOSCUSR = usrName_UCase
lYROPDOS0.ROPDOSCAMJ = DSys
lYROPDOS0.ROPDOSUUSR = usrName_UCase
lYROPDOS0.ROPDOSUAMJ = DSys
lYROPDOS0.ROPDOSUHMS = time_Hms
lYROPDOS0.ROPDOSUVER = 0
lYROPDOS0.ROPDOSGECH = DSys
lYROPDOS0.ROPDOSGUSR = usrName_UCase
lYROPDOS0.ROPDOSGSRV = ""
lYROPDOS0.ROPDOSGNAT = "I"
lYROPDOS0.ROPDOSGPRV = "W"
lYROPDOS0.ROPDOSGNAT = "I"
lYROPDOS0.ROPDOSGGRA = " "
lYROPDOS0.ROPDOSGPRI = "0"
lYROPDOS0.ROPDOSGCOU = 0
lYROPDOS0.ROPDOSIAMJ = DSys
lYROPDOS0.ROPDOSISRV = "XXXX"
lYROPDOS0.ROPDOSIUSR = usrName_UCase
lYROPDOS0.ROPDOSIREF = ""
lYROPDOS0.ROPDOSXDOM = ""
lYROPDOS0.ROPDOSXAPP = ""
lYROPDOS0.ROPDOSXID = ""
lYROPDOS0.ROPDOSQUAL = ""


End Function



