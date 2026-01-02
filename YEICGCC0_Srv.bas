Attribute VB_Name = "srvYEICGCC0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYEICGCC0
 
      EICGCCID     As Long
      EICGCCSTA    As String * 1
      EICGCCSTAK   As String * 1
      
      EICGCCETB    As Long
      EICGCCAGE    As Long
      EICGCCSER    As String * 2
      EICGCCSSE    As String * 2
      EICGCCOPE    As String * 3
      EICGCCDOS    As Long
      EICGCCAMJ    As Long
     
      EICGCCECLI   As String * 7
      EICGCCECPT   As String * 20
      EICGCCEMT    As Currency
      EICGCCECHQ   As String * 7
      EICGCCEIND   As String * 1
      EICGCCEAMJ   As Long
      
      EICGCCXBQ    As String * 5
      EICGCCXCPT   As String * 20
      EICGCCXNOM   As String * 32
      EICGCCXID    As Long
      EICGCCXECO   As String * 32
      
      EICGCCKLAB   As String * 1
      EICGCCKSIG   As String * 1
      EICGCCKEND   As String * 1
      EICGCCKMT    As String * 1
      
      EICGCCVAMJ   As Long
      EICGCCVJPG   As Long
      EICGCCVREM   As Long
      EICGCCVINT   As String * 16
      EICGCCVEXT   As String * 16
    
      EICGCCUUSR   As String * 10
      EICGCCUAMJ   As Long
      EICGCCUHMS   As Long
      EICGCCUSEQ   As Long

    
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    
    JOENTT          As String * 2
    JODATE          As String * 6

'____________________________________________________ Journalisation
End Type
Public xYEICGCC0 As typeYEICGCC0
Public Function getSeuil1YEICGCC0() As Currency
Dim xSQL As String
Dim retour As Currency
Dim rsSeuil As New ADODB.Recordset

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
         & " where BIATABID = 'YEICGCC0' and BIATABK1 = 'seuil1'"
    Set rsSeuil = cnsab.Execute(xSQL)

    If Not rsSeuil.EOF Then
        retour = CCur(Trim(rsSeuil("BIATABK2")))
    Else
        retour = 10000 'par défaut
    End If
    rsSeuil.Close
    Set rsSeuil = Nothing

    getSeuil1YEICGCC0 = retour
    
End Function

Public Function getSeuil2YEICGCC0() As Currency
Dim xSQL As String
Dim retour As Currency
Dim rsSeuil As New ADODB.Recordset

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
         & " where BIATABID = 'YEICGCC0' and BIATABK1 = 'seuil2'"
    Set rsSeuil = cnsab.Execute(xSQL)

    If Not rsSeuil.EOF Then
        retour = CCur(Trim(rsSeuil("BIATABK2")))
    Else
        retour = 150000 'par défaut
    End If
    rsSeuil.Close
    Set rsSeuil = Nothing

    getSeuil2YEICGCC0 = retour
    
End Function

Public Function sqlYEICGCC0_Delete(oldY As typeYEICGCC0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYEICGCC0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where EICGCCID = " & oldY.EICGCCID _
       & " and EICGCCUSEQ = " & oldY.EICGCCUSEQ

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYEICGCC0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYEICGCC0_Delete = Error
End Function

Public Function sqlYEICGCC0_Update(newY As typeYEICGCC0, oldY As typeYEICGCC0, blnUUSR As Boolean)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYEICGCC0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.EICGCCID <> newY.EICGCCID _
Or oldY.EICGCCUSEQ <> newY.EICGCCUSEQ Then
    sqlYEICGCC0_Update = "Erreur EICGCCID : " & newY.EICGCCID & "." & oldY.EICGCCUSEQ
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where EICGCCID = " & oldY.EICGCCID _
       & " and EICGCCUSEQ = " & oldY.EICGCCUSEQ

newY.EICGCCUSEQ = newY.EICGCCUSEQ + 1
xSet = xSet & " set EICGCCUSEQ = " & newY.EICGCCUSEQ
blnUpdate = False

If blnUUSR Then
    newY.EICGCCUUSR = usrName_UCase
    newY.EICGCCUAMJ = DSys
    newY.EICGCCUHMS = time_Hms
End If

' Détecter les modifications
'===================================================================================
If newY.EICGCCSTA <> oldY.EICGCCSTA Then blnUpdate = True:  xSet = xSet & " , EICGCCSTA = '" & newY.EICGCCSTA & "'"
If newY.EICGCCSTAK <> oldY.EICGCCSTAK Then blnUpdate = True:  xSet = xSet & " , EICGCCSTAK = '" & newY.EICGCCSTAK & "'"

If newY.EICGCCETB <> oldY.EICGCCETB Then blnUpdate = True:  xSet = xSet & " , EICGCCETB = " & newY.EICGCCETB
If newY.EICGCCAGE <> oldY.EICGCCAGE Then blnUpdate = True:  xSet = xSet & " , EICGCCAGE = " & newY.EICGCCAGE
If newY.EICGCCSER <> oldY.EICGCCSER Then blnUpdate = True:  xSet = xSet & " , EICGCCSER = '" & newY.EICGCCSER & "'"
If newY.EICGCCSSE <> oldY.EICGCCSSE Then blnUpdate = True:  xSet = xSet & " , EICGCCSSE = '" & newY.EICGCCSSE & "'"
If newY.EICGCCOPE <> oldY.EICGCCOPE Then blnUpdate = True:  xSet = xSet & " , EICGCCOPE = '" & newY.EICGCCOPE & "'"
If newY.EICGCCDOS <> oldY.EICGCCDOS Then blnUpdate = True:  xSet = xSet & " , EICGCCDOS = " & newY.EICGCCDOS
If newY.EICGCCAMJ <> oldY.EICGCCAMJ Then blnUpdate = True:  xSet = xSet & " , EICGCCAMJ = " & newY.EICGCCAMJ

If newY.EICGCCECLI <> oldY.EICGCCECLI Then blnUpdate = True:  xSet = xSet & " , EICGCCECLI = '" & newY.EICGCCECLI & "'"
If newY.EICGCCECPT <> oldY.EICGCCECPT Then blnUpdate = True:  xSet = xSet & " , EICGCCECPT = '" & newY.EICGCCECPT & "'"
If newY.EICGCCEMT <> oldY.EICGCCEMT Then blnUpdate = True:  xSet = xSet & " , EICGCCEMT = " & cur_P(newY.EICGCCEMT)
If newY.EICGCCECHQ <> oldY.EICGCCECHQ Then blnUpdate = True:  xSet = xSet & " , EICGCCECHQ = '" & newY.EICGCCECHQ & "'"
If newY.EICGCCEIND <> oldY.EICGCCEIND Then blnUpdate = True:  xSet = xSet & " , EICGCCEIND = '" & newY.EICGCCEIND & "'"
If newY.EICGCCEAMJ <> oldY.EICGCCEAMJ Then blnUpdate = True:  xSet = xSet & " , EICGCCEAMJ = " & newY.EICGCCEAMJ

If newY.EICGCCXBQ <> oldY.EICGCCXBQ Then blnUpdate = True:  xSet = xSet & " , EICGCCXBQ = '" & Replace(Trim(newY.EICGCCXBQ), "'", "''") & "'"
If newY.EICGCCXCPT <> oldY.EICGCCXCPT Then blnUpdate = True:  xSet = xSet & " , EICGCCXCPT = '" & Replace(Trim(newY.EICGCCXCPT), "'", "''") & "'"
If newY.EICGCCXNOM <> oldY.EICGCCXNOM Then blnUpdate = True:  xSet = xSet & " , EICGCCXNOM = '" & Replace(Trim(newY.EICGCCXNOM), "'", "''") & "'"
If newY.EICGCCXID <> oldY.EICGCCXID Then blnUpdate = True:  xSet = xSet & " , EICGCCXID = " & newY.EICGCCXID
If newY.EICGCCXECO <> oldY.EICGCCXECO Then blnUpdate = True:  xSet = xSet & " , EICGCCXECO = '" & Replace(Trim(newY.EICGCCXECO), "'", "''") & "'"

If newY.EICGCCKLAB <> oldY.EICGCCKLAB Then blnUpdate = True:  xSet = xSet & " , EICGCCKLAB = '" & newY.EICGCCKLAB & "'"
If newY.EICGCCKSIG <> oldY.EICGCCKSIG Then blnUpdate = True:  xSet = xSet & " , EICGCCKSIG = '" & Replace(Trim(newY.EICGCCKSIG), "'", "''") & "'"
If newY.EICGCCKEND <> oldY.EICGCCKEND Then blnUpdate = True:  xSet = xSet & " , EICGCCKEND = '" & Replace(Trim(newY.EICGCCKEND), "'", "''") & "'"
If newY.EICGCCKMT <> oldY.EICGCCKMT Then blnUpdate = True:  xSet = xSet & " , EICGCCKMT = '" & newY.EICGCCKMT & "'"

If newY.EICGCCVAMJ <> oldY.EICGCCVAMJ Then blnUpdate = True:  xSet = xSet & " , EICGCCVAMJ = " & newY.EICGCCVAMJ
If newY.EICGCCVJPG <> oldY.EICGCCVJPG Then blnUpdate = True:  xSet = xSet & " , EICGCCVJPG = " & newY.EICGCCVJPG
If newY.EICGCCVREM <> oldY.EICGCCVREM Then blnUpdate = True:  xSet = xSet & " , EICGCCVREM = " & newY.EICGCCVREM
If newY.EICGCCVINT <> oldY.EICGCCVINT Then blnUpdate = True:  xSet = xSet & " , EICGCCVINT = '" & Replace(Trim(newY.EICGCCVINT), "'", "''") & "'"
If newY.EICGCCVEXT <> oldY.EICGCCVEXT Then blnUpdate = True:  xSet = xSet & " , EICGCCVEXT = '" & Replace(Trim(newY.EICGCCVEXT), "'", "''") & "'"


If newY.EICGCCUUSR <> oldY.EICGCCUUSR Then blnUpdate = True:  xSet = xSet & " , EICGCCUUSR = '" & newY.EICGCCUUSR & "'"
If newY.EICGCCUAMJ <> oldY.EICGCCUAMJ Then blnUpdate = True:  xSet = xSet & " , EICGCCUAMJ = " & newY.EICGCCUAMJ
If newY.EICGCCUHMS <> oldY.EICGCCUHMS Then blnUpdate = True:  xSet = xSet & " , EICGCCUHMS = " & newY.EICGCCUHMS

'If newY.EICGCCID < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYEICGCC0_Update = "Erreur màj : " & newY.EICGCCID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYEICGCC0_Update = Error
End Function

Public Function sqlYEICGCC0_Insert(newY As typeYEICGCC0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYEICGCC0_Insert = Null
xSet = " (EICGCCID"
xValues = " values(" & newY.EICGCCID

newY.EICGCCUUSR = usrName_UCase
newY.EICGCCUAMJ = DSys
newY.EICGCCUHMS = time_Hms

' Détecter les modifications
'===================================================================================
If newY.EICGCCETB <> 0 Then xSet = xSet & ",EICGCCETB": xValues = xValues & " ," & newY.EICGCCETB
If newY.EICGCCAGE <> 0 Then xSet = xSet & ",EICGCCAGE": xValues = xValues & " ," & newY.EICGCCAGE
If newY.EICGCCDOS <> 0 Then xSet = xSet & ",EICGCCDOS": xValues = xValues & " ," & newY.EICGCCDOS
If newY.EICGCCAMJ <> 0 Then xSet = xSet & ",EICGCCAMJ": xValues = xValues & " ," & newY.EICGCCAMJ
If newY.EICGCCEMT <> 0 Then xSet = xSet & ",EICGCCEMT": xValues = xValues & " ," & cur_P(newY.EICGCCEMT)
If newY.EICGCCEAMJ <> 0 Then xSet = xSet & ",EICGCCEAMJ": xValues = xValues & " ," & newY.EICGCCEAMJ
If newY.EICGCCXID <> 0 Then xSet = xSet & ",EICGCCXID": xValues = xValues & " ," & newY.EICGCCXID
If newY.EICGCCVAMJ <> 0 Then xSet = xSet & ",EICGCCVAMJ": xValues = xValues & " ," & newY.EICGCCVAMJ
If newY.EICGCCVJPG <> 0 Then xSet = xSet & ",EICGCCVJPG": xValues = xValues & " ," & newY.EICGCCVJPG
If newY.EICGCCVREM <> 0 Then xSet = xSet & ",EICGCCVREM": xValues = xValues & " ," & newY.EICGCCVREM
If newY.EICGCCUSEQ <> 0 Then xSet = xSet & ",EICGCCUSEQ": xValues = xValues & " ," & newY.EICGCCUSEQ
If newY.EICGCCUAMJ <> 0 Then xSet = xSet & ",EICGCCUAMJ": xValues = xValues & " ," & newY.EICGCCUAMJ
If newY.EICGCCUHMS <> 0 Then xSet = xSet & ",EICGCCUHMS": xValues = xValues & " ," & newY.EICGCCUHMS

If Trim(newY.EICGCCSTA) <> "" Then xSet = xSet & ",EICGCCSTA": xValues = xValues & " ,'" & newY.EICGCCSTA & "'"
If Trim(newY.EICGCCSTAK) <> "" Then xSet = xSet & ",EICGCCSTAK": xValues = xValues & " ,'" & newY.EICGCCSTAK & "'"

If Trim(newY.EICGCCSER) <> "" Then xSet = xSet & ",EICGCCSER": xValues = xValues & " ,'" & newY.EICGCCSER & "'"
If Trim(newY.EICGCCSSE) <> "" Then xSet = xSet & ",EICGCCSSE": xValues = xValues & " ,'" & newY.EICGCCSSE & "'"
If Trim(newY.EICGCCOPE) <> "" Then xSet = xSet & ",EICGCCOPE": xValues = xValues & " ,'" & newY.EICGCCOPE & "'"

If Trim(newY.EICGCCECLI) <> "" Then xSet = xSet & ",EICGCCECLI": xValues = xValues & " ,'" & newY.EICGCCECLI & "'"
If Trim(newY.EICGCCECPT) <> "" Then xSet = xSet & ",EICGCCECPT": xValues = xValues & " ,'" & newY.EICGCCECPT & "'"
If Trim(newY.EICGCCECHQ) <> "" Then xSet = xSet & ",EICGCCECHQ": xValues = xValues & " ,'" & newY.EICGCCECHQ & "'"
If Trim(newY.EICGCCEIND) <> "" Then xSet = xSet & ",EICGCCEIND": xValues = xValues & " ,'" & newY.EICGCCEIND & "'"

If Trim(newY.EICGCCXBQ) <> "" Then xSet = xSet & ",EICGCCXBQ": xValues = xValues & " ,'" & Replace(Trim(newY.EICGCCXBQ), "'", "''") & "'"
If Trim(newY.EICGCCXCPT) <> "" Then xSet = xSet & ",EICGCCXCPT": xValues = xValues & " ,'" & Replace(Trim(newY.EICGCCXCPT), "'", "''") & "'"
If Trim(newY.EICGCCXNOM) <> "" Then xSet = xSet & ",EICGCCXNOM": xValues = xValues & " ,'" & Replace(Trim(newY.EICGCCXNOM), "'", "''") & "'"
If Trim(newY.EICGCCXECO) <> "" Then xSet = xSet & ",EICGCCXECO": xValues = xValues & " ,'" & Replace(Trim(newY.EICGCCXECO), "'", "''") & "'"

If Trim(newY.EICGCCKLAB) <> "" Then xSet = xSet & ",EICGCCKLAB": xValues = xValues & " ,'" & newY.EICGCCKLAB & "'"
If Trim(newY.EICGCCKSIG) <> "" Then xSet = xSet & ",EICGCCKSIG": xValues = xValues & " ,'" & newY.EICGCCKSIG & "'"
If Trim(newY.EICGCCKEND) <> "" Then xSet = xSet & ",EICGCCKEND": xValues = xValues & " ,'" & newY.EICGCCKEND & "'"
If Trim(newY.EICGCCKMT) <> "" Then xSet = xSet & ",EICGCCKMT": xValues = xValues & " ,'" & newY.EICGCCKMT & "'"

If Trim(newY.EICGCCVINT) <> "" Then xSet = xSet & ",EICGCCVINT": xValues = xValues & " ,'" & Replace(Trim(newY.EICGCCVINT), "'", "''") & "'"
If Trim(newY.EICGCCVEXT) <> "" Then xSet = xSet & ",EICGCCVEXT": xValues = xValues & " ,'" & Replace(Trim(newY.EICGCCVEXT), "'", "''") & "'"

If Trim(newY.EICGCCUUSR) <> "" Then xSet = xSet & ",EICGCCUUSR": xValues = xValues & " ,'" & newY.EICGCCUUSR & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" & xSet & ")" & xValues & ")"

Set rsADO = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYEICGCC0_Insert = "Erreur màj : " & newY.EICGCCID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYEICGCC0_Insert = Error
End Function

Public Function rsYEICGCC0_GetBuffer(rsADO As ADODB.Recordset, lYEICGCC0 As typeYEICGCC0)
On Error GoTo Error_Handler
rsYEICGCC0_GetBuffer = Null

lYEICGCC0.JORCV = 0
lYEICGCC0.JOSEQN = 0
lYEICGCC0.JRNBIATRN = 0
lYEICGCC0.JOENTT = ""
lYEICGCC0.JODATE = ""

lYEICGCC0.EICGCCID = rsADO("EICGCCID")
lYEICGCC0.EICGCCSTA = rsADO("EICGCCSTA")
lYEICGCC0.EICGCCSTAK = rsADO("EICGCCSTAK")

lYEICGCC0.EICGCCETB = rsADO("EICGCCETB")
lYEICGCC0.EICGCCAGE = rsADO("EICGCCAGE")
lYEICGCC0.EICGCCSER = rsADO("EICGCCSER")
lYEICGCC0.EICGCCSSE = rsADO("EICGCCSSE")
lYEICGCC0.EICGCCOPE = rsADO("EICGCCOPE")
lYEICGCC0.EICGCCDOS = rsADO("EICGCCDOS")
lYEICGCC0.EICGCCAMJ = rsADO("EICGCCAMJ")

lYEICGCC0.EICGCCECLI = rsADO("EICGCCECLI")
lYEICGCC0.EICGCCECPT = rsADO("EICGCCECPT")
lYEICGCC0.EICGCCEMT = rsADO("EICGCCEMT")
lYEICGCC0.EICGCCECHQ = rsADO("EICGCCECHQ")
lYEICGCC0.EICGCCEIND = rsADO("EICGCCEIND")
lYEICGCC0.EICGCCEAMJ = rsADO("EICGCCEAMJ")

lYEICGCC0.EICGCCXBQ = rsADO("EICGCCXBQ")
lYEICGCC0.EICGCCXCPT = rsADO("EICGCCXCPT")
lYEICGCC0.EICGCCXNOM = rsADO("EICGCCXNOM")
lYEICGCC0.EICGCCXID = rsADO("EICGCCXID")
lYEICGCC0.EICGCCXECO = rsADO("EICGCCXECO")

lYEICGCC0.EICGCCKLAB = rsADO("EICGCCKLAB")
lYEICGCC0.EICGCCKSIG = rsADO("EICGCCKSIG")
lYEICGCC0.EICGCCKEND = rsADO("EICGCCKEND")
lYEICGCC0.EICGCCKMT = rsADO("EICGCCKMT")

lYEICGCC0.EICGCCVAMJ = rsADO("EICGCCVAMJ")
lYEICGCC0.EICGCCVJPG = rsADO("EICGCCVJPG")
lYEICGCC0.EICGCCVREM = rsADO("EICGCCVREM")
lYEICGCC0.EICGCCVINT = rsADO("EICGCCVINT")
lYEICGCC0.EICGCCVEXT = rsADO("EICGCCVEXT")

lYEICGCC0.EICGCCUUSR = rsADO("EICGCCUUSR")
lYEICGCC0.EICGCCUAMJ = rsADO("EICGCCUAMJ")
lYEICGCC0.EICGCCUHMS = rsADO("EICGCCUHMS")
lYEICGCC0.EICGCCUSEQ = rsADO("EICGCCUSEQ")

Exit Function
Error_Handler:
rsYEICGCC0_GetBuffer = Error


End Function
'---------------------------------------------------------
Public Function rsJEICGCC0_GetBuffer(rsADO As ADODB.Recordset, rsYEICGCC0 As typeYEICGCC0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJEICGCC0_GetBuffer = Null

rsJEICGCC0_GetBuffer = rsYEICGCC0_GetBuffer(rsADO, rsYEICGCC0)
rsYEICGCC0.JORCV = rsADO("JORCV")
rsYEICGCC0.JOSEQN = rsADO("JOSEQN")
rsYEICGCC0.JRNBIATRN = rsADO("JRNBIATRN")
rsYEICGCC0.JOENTT = rsADO("JOENTT")
rsYEICGCC0.JODATE = rsADO("JODATE")

Exit Function

Error_Handler:

rsJEICGCC0_GetBuffer = Error

End Function


Public Function rsYEICGCC0_Init(lYEICGCC0 As typeYEICGCC0)

lYEICGCC0.EICGCCID = 0
lYEICGCC0.EICGCCSTA = ""
lYEICGCC0.EICGCCSTAK = ""
      
lYEICGCC0.EICGCCETB = 0
lYEICGCC0.EICGCCAGE = 0
lYEICGCC0.EICGCCSER = ""
lYEICGCC0.EICGCCSSE = ""
lYEICGCC0.EICGCCOPE = ""
lYEICGCC0.EICGCCDOS = 0
lYEICGCC0.EICGCCAMJ = 0
     
lYEICGCC0.EICGCCECLI = ""
lYEICGCC0.EICGCCECPT = ""
lYEICGCC0.EICGCCEMT = 0
lYEICGCC0.EICGCCECHQ = ""
lYEICGCC0.EICGCCEIND = ""
lYEICGCC0.EICGCCEAMJ = 0
      
lYEICGCC0.EICGCCXBQ = ""
lYEICGCC0.EICGCCXCPT = ""
lYEICGCC0.EICGCCXNOM = ""
lYEICGCC0.EICGCCXID = 0
lYEICGCC0.EICGCCXECO = ""
      
lYEICGCC0.EICGCCKLAB = ""
lYEICGCC0.EICGCCKSIG = ""
lYEICGCC0.EICGCCKEND = ""
lYEICGCC0.EICGCCKMT = ""

lYEICGCC0.EICGCCVAMJ = 0
lYEICGCC0.EICGCCVJPG = 0
lYEICGCC0.EICGCCVREM = 0
lYEICGCC0.EICGCCVINT = ""
lYEICGCC0.EICGCCVEXT = ""
    
lYEICGCC0.EICGCCUUSR = ""
lYEICGCC0.EICGCCUAMJ = 0
lYEICGCC0.EICGCCUHMS = 0
lYEICGCC0.EICGCCUSEQ = 0
End Function





