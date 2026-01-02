Attribute VB_Name = "srvYKYCDOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYKYCDOS0
 
      KYCDOSNAT    As String    'nature D J *
      KYCDOSID     As String    'identification
      KYCDOSSEQ    As Long      'séquence
      KYCDOSSEQ2    As Long      'séquence
      
      KYCDOSSTAK   As String    'statut / nat
      KYCDOSPJ     As String
      KYCDOSDECH   As Long  'échéance
      KYCDOSDAMJ   As Long   'date du document
      KYCDOSDLIB   As String  'libellé
      
      KYCDOSUUSR   As String  'utilisateur màj
      KYCDOSUAMJ   As Long   'date maàj
      KYCDOSUHMS   As Long  'heure màj
      KYCDOSUVER   As Long         'version
      KYCDOSUFCT   As String
End Type

Type typeWKYCCTL
 
      CLIENACLI     As String    'identification
      CLIENARA      As String
      CLIENARES     As String
      CLIENCPIE     As String
      CLILIBDA1     As Long
      KYCDOSSEQ2    As Long      'séquence
      
      KYCDOSDAMJ   As Long   'date du document
      KYCDOSDECH   As Long
      KYCDOSDLIB   As String  'libellé
      STA          As String
End Type

Public Function sqlYKYCDOS0_Delete(oldY As typeYKYCDOS0, blnKYCDOSUVER As Boolean)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYKYCDOS0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================
If blnKYCDOSUVER Then
    xWhere = " where KYCDOSID = '" & oldY.KYCDOSID & "'" _
           & " and KYCDOSNAT = '" & oldY.KYCDOSNAT & "'" _
           & " and KYCDOSSEQ = " & oldY.KYCDOSSEQ _
           & " and KYCDOSSEQ2 = " & oldY.KYCDOSSEQ2 _
           & " and KYCDOSUVER = " & oldY.KYCDOSUVER
Else
    xWhere = " where KYCDOSID = '" & oldY.KYCDOSID & "'" _
           & " and KYCDOSNAT = '" & oldY.KYCDOSNAT & "'" _
           & " and KYCDOSSEQ = " & oldY.KYCDOSSEQ _
           & " and KYCDOSSEQ2 = " & oldY.KYCDOSSEQ2
End If

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YKYCDOS0" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYKYCDOS0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    
    
oldY.KYCDOSUUSR = usrName_UCase
oldY.KYCDOSUAMJ = DSys
oldY.KYCDOSUHMS = time_Hms
oldY.KYCDOSUVER = -oldY.KYCDOSUVER
sqlYKYCDOS0_Delete = sqlYKYCDOSH_Insert(oldY)
    


Exit Function
Error_Handler:
    sqlYKYCDOS0_Delete = Error
End Function

Public Function sqlYKYCDOS0_Delete_Where(lWhere As String)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYKYCDOS0_Delete_Where = Null

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YKYCDOS0" & lWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYKYCDOS0_Delete_Where = "Erreur màj : " & lWhere
        Exit Function
    End If
    

Exit Function
Error_Handler:
    sqlYKYCDOS0_Delete_Where = Error
End Function

Public Function sqlYKYCDOS0_Update(newY As typeYKYCDOS0, oldY As typeYKYCDOS0, blnUUSR As Boolean)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYKYCDOS0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.KYCDOSID <> newY.KYCDOSID _
Or oldY.KYCDOSUVER <> newY.KYCDOSUVER Then
    sqlYKYCDOS0_Update = "Erreur KYCDOSID : " & newY.KYCDOSID & "." & oldY.KYCDOSUVER
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where KYCDOSNAT= '" & oldY.KYCDOSNAT & "'" _
       & " and KYCDOSID = '" & oldY.KYCDOSID & "'" _
       & " and KYCDOSSEQ= " & oldY.KYCDOSSEQ _
       & " and KYCDOSSEQ2= " & oldY.KYCDOSSEQ2 _
       & " and KYCDOSUVER = " & oldY.KYCDOSUVER

newY.KYCDOSUVER = newY.KYCDOSUVER + 1
xSet = xSet & " set KYCDOSUVER = " & newY.KYCDOSUVER
blnUpdate = False

If blnUUSR Then
    newY.KYCDOSUUSR = usrName_UCase
    newY.KYCDOSUAMJ = DSys
    newY.KYCDOSUHMS = time_Hms
End If

' Détecter les modifications
'===================================================================================
If newY.KYCDOSSTAK <> oldY.KYCDOSSTAK Then blnUpdate = True:  xSet = xSet & " , KYCDOSSTAK = '" & newY.KYCDOSSTAK & "'"
If newY.KYCDOSDECH <> oldY.KYCDOSDECH Then blnUpdate = True:  xSet = xSet & " , KYCDOSDECH = " & newY.KYCDOSDECH
If newY.KYCDOSDAMJ <> oldY.KYCDOSDAMJ Then blnUpdate = True:  xSet = xSet & " , KYCDOSDAMJ = " & newY.KYCDOSDAMJ
If newY.KYCDOSDLIB <> oldY.KYCDOSDLIB Then blnUpdate = True:  xSet = xSet & " , KYCDOSDLIB = '" & Replace(Trim(newY.KYCDOSDLIB), "'", "''") & "'"
If newY.KYCDOSUUSR <> oldY.KYCDOSUUSR Then xSet = xSet & " , KYCDOSUUSR = '" & newY.KYCDOSUUSR & "'"
If newY.KYCDOSUAMJ <> oldY.KYCDOSUAMJ Then xSet = xSet & " , KYCDOSUAMJ = " & newY.KYCDOSUAMJ
If newY.KYCDOSUHMS <> oldY.KYCDOSUHMS Then xSet = xSet & " , KYCDOSUHMS = " & newY.KYCDOSUHMS

If newY.KYCDOSPJ <> oldY.KYCDOSPJ Then blnUpdate = True:  xSet = xSet & " , KYCDOSPJ = '" & newY.KYCDOSPJ & "'"
If newY.KYCDOSUFCT <> oldY.KYCDOSUFCT Then blnUpdate = True:  xSet = xSet & " , KYCDOSUFCT = '" & newY.KYCDOSUFCT & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YKYCDOS0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYKYCDOS0_Update = "Erreur màj : " & newY.KYCDOSID
        Exit Function
    End If
    sqlYKYCDOS0_Update = sqlYKYCDOSH_Insert(newY)
    
End If

Exit Function
Error_Handler:
    sqlYKYCDOS0_Update = Error
End Function

Public Function sqlYKYCDOS0_Insert(newY As typeYKYCDOS0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYKYCDOS0_Insert = Null
xSet = " (KYCDOSNAT , KYCDOSID , KYCDOSSEQ , KYCDOSSEQ2 "
xValues = " values('" & newY.KYCDOSNAT & "' ,'" & newY.KYCDOSID & "' ," & newY.KYCDOSSEQ & " ," & newY.KYCDOSSEQ2

newY.KYCDOSUUSR = usrName_UCase
newY.KYCDOSUAMJ = DSys
newY.KYCDOSUHMS = time_Hms

' Détecter les modifications
'===================================================================================
If newY.KYCDOSUVER <> 0 Then xSet = xSet & ",KYCDOSUVER": xValues = xValues & " ," & newY.KYCDOSUVER
If Trim(newY.KYCDOSUAMJ) <> "" Then xSet = xSet & ",KYCDOSUAMJ": xValues = xValues & " ," & newY.KYCDOSUAMJ
If Trim(newY.KYCDOSUHMS) <> "" Then xSet = xSet & ",KYCDOSUHMS": xValues = xValues & " ," & newY.KYCDOSUHMS
If Trim(newY.KYCDOSDECH) <> "" Then xSet = xSet & ",KYCDOSDECH": xValues = xValues & " ," & newY.KYCDOSDECH
If Trim(newY.KYCDOSDAMJ) <> "" Then xSet = xSet & ",KYCDOSDAMJ": xValues = xValues & " ," & newY.KYCDOSDAMJ

If Trim(newY.KYCDOSSTAK) <> "" Then xSet = xSet & ",KYCDOSSTAK": xValues = xValues & " ,'" & newY.KYCDOSSTAK & "'"
If Trim(newY.KYCDOSUUSR) <> "" Then xSet = xSet & ",KYCDOSUUSR": xValues = xValues & " ,'" & newY.KYCDOSUUSR & "'"
If Trim(newY.KYCDOSDLIB) <> "" Then xSet = xSet & ",KYCDOSDLIB": xValues = xValues & " ,'" & Replace(Trim(newY.KYCDOSDLIB), "'", "''") & "'"

If Trim(newY.KYCDOSPJ) <> "" Then xSet = xSet & ",KYCDOSPJ": xValues = xValues & " ,'" & newY.KYCDOSPJ & "'"
If Trim(newY.KYCDOSUFCT) <> "" Then xSet = xSet & ",KYCDOSUFCT": xValues = xValues & " ,'" & newY.KYCDOSUFCT & "'"

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YKYCDOS0" & xSet & ")" & xValues & ")"

Set rsADO = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYKYCDOS0_Insert = "Erreur màj : " & newY.KYCDOSID
    Exit Function
End If

sqlYKYCDOS0_Insert = sqlYKYCDOSH_Insert(newY)

Exit Function
Error_Handler:
    sqlYKYCDOS0_Insert = Error
End Function

Public Function sqlYKYCDOSH_Insert(newY As typeYKYCDOS0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYKYCDOSH_Insert = Null
xSet = " (KYCDOSNAT , KYCDOSID , KYCDOSSEQ , KYCDOSSEQ2 "
xValues = " values('" & newY.KYCDOSNAT & "' ,'" & newY.KYCDOSID & "' ," & newY.KYCDOSSEQ & " ," & newY.KYCDOSSEQ2

' Détecter les modifications
'===================================================================================
If newY.KYCDOSUVER <> 0 Then xSet = xSet & ",KYCDOSUVER": xValues = xValues & " ," & newY.KYCDOSUVER
If Trim(newY.KYCDOSUAMJ) <> "" Then xSet = xSet & ",KYCDOSUAMJ": xValues = xValues & " ," & newY.KYCDOSUAMJ
If Trim(newY.KYCDOSUHMS) <> "" Then xSet = xSet & ",KYCDOSUHMS": xValues = xValues & " ," & newY.KYCDOSUHMS
If Trim(newY.KYCDOSDECH) <> "" Then xSet = xSet & ",KYCDOSDECH": xValues = xValues & " ," & newY.KYCDOSDECH
If Trim(newY.KYCDOSDAMJ) <> "" Then xSet = xSet & ",KYCDOSDAMJ": xValues = xValues & " ," & newY.KYCDOSDAMJ

If Trim(newY.KYCDOSSTAK) <> "" Then xSet = xSet & ",KYCDOSSTAK": xValues = xValues & " ,'" & newY.KYCDOSSTAK & "'"
If Trim(newY.KYCDOSUUSR) <> "" Then xSet = xSet & ",KYCDOSUUSR": xValues = xValues & " ,'" & newY.KYCDOSUUSR & "'"
If Trim(newY.KYCDOSDLIB) <> "" Then xSet = xSet & ",KYCDOSDLIB": xValues = xValues & " ,'" & Replace(Trim(newY.KYCDOSDLIB), "'", "''") & "'"

If Trim(newY.KYCDOSPJ) <> "" Then xSet = xSet & ",KYCDOSPJ": xValues = xValues & " ,'" & newY.KYCDOSPJ & "'"
If Trim(newY.KYCDOSUFCT) <> "" Then xSet = xSet & ",KYCDOSUFCT": xValues = xValues & " ,'" & newY.KYCDOSUFCT & "'"

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YKYCDOSH" & xSet & ")" & xValues & ")"

Set rsADO = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYKYCDOSH_Insert = "Erreur màj : " & newY.KYCDOSID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYKYCDOSH_Insert = Error
End Function

Public Function rsYKYCDOS0_GetBuffer(rsADO As ADODB.Recordset, lYKYCDOS0 As typeYKYCDOS0)
On Error GoTo Error_Handler
rsYKYCDOS0_GetBuffer = Null

lYKYCDOS0.KYCDOSNAT = rsADO("KYCDOSNAT")
lYKYCDOS0.KYCDOSID = Trim(rsADO("KYCDOSID"))
lYKYCDOS0.KYCDOSSEQ = rsADO("KYCDOSSEQ")
lYKYCDOS0.KYCDOSSEQ2 = rsADO("KYCDOSSEQ2")

lYKYCDOS0.KYCDOSSTAK = rsADO("KYCDOSSTAK")
lYKYCDOS0.KYCDOSPJ = rsADO("KYCDOSPJ")
lYKYCDOS0.KYCDOSDECH = rsADO("KYCDOSDECH")
lYKYCDOS0.KYCDOSDAMJ = rsADO("KYCDOSDAMJ")
lYKYCDOS0.KYCDOSDLIB = Trim(rsADO("KYCDOSDLIB"))

lYKYCDOS0.KYCDOSUUSR = Trim(rsADO("KYCDOSUUSR"))
lYKYCDOS0.KYCDOSUAMJ = rsADO("KYCDOSUAMJ")
lYKYCDOS0.KYCDOSUHMS = rsADO("KYCDOSUHMS")
lYKYCDOS0.KYCDOSUVER = rsADO("KYCDOSUVER")
lYKYCDOS0.KYCDOSUFCT = rsADO("KYCDOSUFCT")

Exit Function
Error_Handler:
rsYKYCDOS0_GetBuffer = Error


End Function

Public Function rsYKYCDOS0_Init(lYKYCDOS0 As typeYKYCDOS0)

lYKYCDOS0.KYCDOSNAT = " "
lYKYCDOS0.KYCDOSID = ""
lYKYCDOS0.KYCDOSSEQ = 0
lYKYCDOS0.KYCDOSSEQ2 = 0

lYKYCDOS0.KYCDOSSTAK = " "
lYKYCDOS0.KYCDOSPJ = " "
lYKYCDOS0.KYCDOSDECH = 0
lYKYCDOS0.KYCDOSDAMJ = 0
lYKYCDOS0.KYCDOSDLIB = ""

lYKYCDOS0.KYCDOSUUSR = usrName_UCase
lYKYCDOS0.KYCDOSUAMJ = DSys
lYKYCDOS0.KYCDOSUHMS = time_Hms
lYKYCDOS0.KYCDOSUVER = 0
lYKYCDOS0.KYCDOSUFCT = " "

End Function




