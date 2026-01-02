Attribute VB_Name = "srvYGOSDOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYGOSDOS0
 
      GOSDOSIDD    As Long
      GOSDOSWBIC   As String
      GOSDOSWES    As String
      GOSDOSWMTK   As String
      GOSDOSWTRN   As String
      GOSDOSWMTD   As Currency
      GOSDOSWDEV   As String
      GOSDOSWID1   As Long
      GOSDOSWIDL   As Long
      GOSDOSWIDH   As Long
      
      GOSDOSRCOM   As String
      GOSDOSCLI    As String
      
      GOSDOSPAYS   As String
      GOSDOSLABK   As String
      GOSDOSSTAG   As String
      GOSDOSSTAD   As String
      GOSDOSECHD   As Long
      GOSDOSISRV   As String
      GOSDOSIAMJ   As Long
      GOSDOSITOP   As String
      GOSDOSGSRV   As String
      
      GOSDOSUSRV   As String
      GOSDOSUUSR   As String
      GOSDOSUAMJ   As Long
      GOSDOSUHMS   As Long
      GOSDOSUSEQ   As Long

    
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    
    JOENTT          As String * 2
    JODATE          As String * 6

'____________________________________________________ Journalisation
End Type
Public xYGOSDOS0 As typeYGOSDOS0
Public Function sqlYGOSDOS0_Delete(oldY As typeYGOSDOS0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlYGOSDOS0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where GOSDOSIDD = " & oldY.GOSDOSIDD _
       & " and GOSDOSUSEQ = " & oldY.GOSDOSUSEQ

'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YGOSDOS0" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYGOSDOS0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYGOSDOS0_Delete = Error
End Function

Public Function sqlYGOSDOS0_Update(newY As typeYGOSDOS0, oldY As typeYGOSDOS0, blnUUSR As Boolean)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlYGOSDOS0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.GOSDOSIDD <> newY.GOSDOSIDD _
Or oldY.GOSDOSUSEQ <> newY.GOSDOSUSEQ Then
    sqlYGOSDOS0_Update = "Erreur GOSDOSIDD : " & newY.GOSDOSIDD & "." & oldY.GOSDOSUSEQ
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where GOSDOSIDD = " & oldY.GOSDOSIDD _
       & " and GOSDOSUSEQ = " & oldY.GOSDOSUSEQ

newY.GOSDOSUSEQ = newY.GOSDOSUSEQ + 1
xSet = xSet & " set GOSDOSUSEQ = " & newY.GOSDOSUSEQ
blnUpdate = False

If blnUUSR Then
    newY.GOSDOSUUSR = usrName_UCase
    newY.GOSDOSUSRV = currentSSIWINUNIT
    newY.GOSDOSUAMJ = DSys
    newY.GOSDOSUHMS = time_Hms
End If

' Détecter les modifications
'===================================================================================
If newY.GOSDOSWID1 <> oldY.GOSDOSWID1 Then blnUpdate = True:  xSet = xSet & " , GOSDOSWID1 = " & newY.GOSDOSWID1
If newY.GOSDOSWIDL <> oldY.GOSDOSWIDL Then blnUpdate = True:  xSet = xSet & " , GOSDOSWIDL = " & newY.GOSDOSWIDL
If newY.GOSDOSWIDH <> oldY.GOSDOSWIDH Then blnUpdate = True:  xSet = xSet & " , GOSDOSWIDH = " & newY.GOSDOSWIDH
If newY.GOSDOSWMTD <> oldY.GOSDOSWMTD Then blnUpdate = True:  xSet = xSet & " , GOSDOSWMTD = " & newY.GOSDOSWMTD
If newY.GOSDOSECHD <> oldY.GOSDOSECHD Then blnUpdate = True:  xSet = xSet & " , GOSDOSECHD = " & newY.GOSDOSECHD
If newY.GOSDOSUAMJ <> oldY.GOSDOSUAMJ Then blnUpdate = True:  xSet = xSet & " , GOSDOSUAMJ = " & newY.GOSDOSUAMJ
If newY.GOSDOSUHMS <> oldY.GOSDOSUHMS Then blnUpdate = True:  xSet = xSet & " , GOSDOSUHMS = " & newY.GOSDOSUHMS
If newY.GOSDOSIAMJ <> oldY.GOSDOSIAMJ Then blnUpdate = True:  xSet = xSet & " , GOSDOSIAMJ = " & newY.GOSDOSIAMJ

If newY.GOSDOSWBIC <> oldY.GOSDOSWBIC Then blnUpdate = True:  xSet = xSet & " , GOSDOSWBIC = '" & newY.GOSDOSWBIC & "'"
If newY.GOSDOSWES <> oldY.GOSDOSWES Then blnUpdate = True:  xSet = xSet & " , GOSDOSWES = '" & newY.GOSDOSWES & "'"
If newY.GOSDOSWMTK <> oldY.GOSDOSWMTK Then blnUpdate = True:  xSet = xSet & " , GOSDOSWMTK = '" & newY.GOSDOSWMTK & "'"
If newY.GOSDOSWTRN <> oldY.GOSDOSWTRN Then blnUpdate = True:  xSet = xSet & " , GOSDOSWTRN = '" & newY.GOSDOSWTRN & "'"
If newY.GOSDOSWDEV <> oldY.GOSDOSWDEV Then blnUpdate = True:  xSet = xSet & " , GOSDOSWDEV = '" & newY.GOSDOSWDEV & "'"
If newY.GOSDOSRCOM <> oldY.GOSDOSRCOM Then blnUpdate = True:  xSet = xSet & " , GOSDOSRCOM = '" & newY.GOSDOSRCOM & "'"
If newY.GOSDOSCLI <> oldY.GOSDOSCLI Then blnUpdate = True:  xSet = xSet & " , GOSDOSCLI = '" & newY.GOSDOSCLI & "'"
If newY.GOSDOSPAYS <> oldY.GOSDOSPAYS Then blnUpdate = True:  xSet = xSet & " , GOSDOSPAYS = '" & newY.GOSDOSPAYS & "'"
If newY.GOSDOSLABK <> oldY.GOSDOSLABK Then blnUpdate = True:  xSet = xSet & " , GOSDOSLABK = '" & newY.GOSDOSLABK & "'"
If newY.GOSDOSSTAG <> oldY.GOSDOSSTAG Then blnUpdate = True:  xSet = xSet & " , GOSDOSSTAG = '" & newY.GOSDOSSTAG & "'"
If newY.GOSDOSSTAD <> oldY.GOSDOSSTAD Then blnUpdate = True:  xSet = xSet & " , GOSDOSSTAD = '" & newY.GOSDOSSTAD & "'"
If newY.GOSDOSISRV <> oldY.GOSDOSISRV Then blnUpdate = True:  xSet = xSet & " , GOSDOSISRV = '" & newY.GOSDOSISRV & "'"
If newY.GOSDOSITOP <> oldY.GOSDOSITOP Then blnUpdate = True:  xSet = xSet & " , GOSDOSITOP = '" & newY.GOSDOSITOP & "'"
If newY.GOSDOSGSRV <> oldY.GOSDOSGSRV Then blnUpdate = True:  xSet = xSet & " , GOSDOSGSRV = '" & newY.GOSDOSGSRV & "'"
If newY.GOSDOSUSRV <> oldY.GOSDOSUSRV Then blnUpdate = True:  xSet = xSet & " , GOSDOSUSRV = '" & newY.GOSDOSUSRV & "'"


If newY.GOSDOSUUSR <> oldY.GOSDOSUUSR Then blnUpdate = True:  xSet = xSet & " , GOSDOSUUSR = '" & newY.GOSDOSUUSR & "'"


If blnUpdate Then
    
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YGOSDOS0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYGOSDOS0_Update = "Erreur màj : " & newY.GOSDOSIDD
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYGOSDOS0_Update = Error
End Function

Public Function sqlYGOSDOS0_Insert(newY As typeYGOSDOS0)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

'On Error GoTo Error_Handler
sqlYGOSDOS0_Insert = Null
xSet = " (GOSDOSIDD"
xValues = " values(" & newY.GOSDOSIDD

newY.GOSDOSUUSR = usrName_UCase
newY.GOSDOSUSRV = currentSSIWINUNIT
newY.GOSDOSUAMJ = DSys
newY.GOSDOSUHMS = time_Hms

' Détecter les modifications
'===================================================================================
If newY.GOSDOSWID1 <> 0 Then xSet = xSet & ",GOSDOSWID1": xValues = xValues & " ," & newY.GOSDOSWID1
If newY.GOSDOSWIDL <> 0 Then xSet = xSet & ",GOSDOSWIDL": xValues = xValues & " ," & newY.GOSDOSWIDL
If newY.GOSDOSWIDH <> 0 Then xSet = xSet & ",GOSDOSWIDH": xValues = xValues & " ," & newY.GOSDOSWIDH
If newY.GOSDOSWMTD <> 0 Then xSet = xSet & ",GOSDOSWMTD": xValues = xValues & " ," & cur_P(newY.GOSDOSWMTD)
If newY.GOSDOSECHD <> 0 Then xSet = xSet & ",GOSDOSECHD": xValues = xValues & " ," & newY.GOSDOSECHD
If newY.GOSDOSIAMJ <> 0 Then xSet = xSet & ",GOSDOSIAMJ": xValues = xValues & " ," & newY.GOSDOSIAMJ
If newY.GOSDOSUSEQ <> 0 Then xSet = xSet & ",GOSDOSUSEQ": xValues = xValues & " ," & newY.GOSDOSUSEQ
If newY.GOSDOSUAMJ <> 0 Then xSet = xSet & ",GOSDOSUAMJ": xValues = xValues & " ," & newY.GOSDOSUAMJ
If newY.GOSDOSUHMS <> 0 Then xSet = xSet & ",GOSDOSUHMS": xValues = xValues & " ," & newY.GOSDOSUHMS

If Trim(newY.GOSDOSWBIC) <> "" Then xSet = xSet & ",GOSDOSWBIC": xValues = xValues & " ,'" & newY.GOSDOSWBIC & "'"
If Trim(newY.GOSDOSWES) <> "" Then xSet = xSet & ",GOSDOSWES": xValues = xValues & " ,'" & newY.GOSDOSWES & "'"
If Trim(newY.GOSDOSWMTK) <> "" Then xSet = xSet & ",GOSDOSWMTK": xValues = xValues & " ,'" & newY.GOSDOSWMTK & "'"
If Trim(newY.GOSDOSWTRN) <> "" Then xSet = xSet & ",GOSDOSWTRN": xValues = xValues & " ,'" & newY.GOSDOSWTRN & "'"
If Trim(newY.GOSDOSWDEV) <> "" Then xSet = xSet & ",GOSDOSWDEV": xValues = xValues & " ,'" & newY.GOSDOSWDEV & "'"
If Trim(newY.GOSDOSRCOM) <> "" Then xSet = xSet & ",GOSDOSRCOM": xValues = xValues & " ,'" & newY.GOSDOSRCOM & "'"
If Trim(newY.GOSDOSCLI) <> "" Then xSet = xSet & ",GOSDOSCLI": xValues = xValues & " ,'" & newY.GOSDOSCLI & "'"
If Trim(newY.GOSDOSPAYS) <> "" Then xSet = xSet & ",GOSDOSPAYS": xValues = xValues & " ,'" & newY.GOSDOSPAYS & "'"
If Trim(newY.GOSDOSLABK) <> "" Then xSet = xSet & ",GOSDOSLABK": xValues = xValues & " ,'" & newY.GOSDOSLABK & "'"
If Trim(newY.GOSDOSSTAG) <> "" Then xSet = xSet & ",GOSDOSSTAG": xValues = xValues & " ,'" & newY.GOSDOSSTAG & "'"
If Trim(newY.GOSDOSSTAD) <> "" Then xSet = xSet & ",GOSDOSSTAD": xValues = xValues & " ,'" & newY.GOSDOSSTAD & "'"
If Trim(newY.GOSDOSISRV) <> "" Then xSet = xSet & ",GOSDOSISRV": xValues = xValues & " ,'" & newY.GOSDOSISRV & "'"
If Trim(newY.GOSDOSITOP) <> "" Then xSet = xSet & ",GOSDOSITOP": xValues = xValues & " ,'" & newY.GOSDOSITOP & "'"
If Trim(newY.GOSDOSGSRV) <> "" Then xSet = xSet & ",GOSDOSGSRV": xValues = xValues & " ,'" & newY.GOSDOSGSRV & "'"
If Trim(newY.GOSDOSUSRV) <> "" Then xSet = xSet & ",GOSDOSUSRV": xValues = xValues & " ,'" & newY.GOSDOSUSRV & "'"

If Trim(newY.GOSDOSUUSR) <> "" Then xSet = xSet & ",GOSDOSUUSR": xValues = xValues & " ,'" & newY.GOSDOSUUSR & "'"

xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YGOSDOS0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYGOSDOS0_Insert = "Erreur màj : " & newY.GOSDOSIDD
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYGOSDOS0_Insert = Error
End Function

Public Function rsYGOSDOS0_GetBuffer(rsADO As ADODB.Recordset, lYGOSDOS0 As typeYGOSDOS0)
On Error GoTo Error_Handler
rsYGOSDOS0_GetBuffer = Null

lYGOSDOS0.JORCV = 0
lYGOSDOS0.JOSEQN = 0
lYGOSDOS0.JRNBIATRN = 0
lYGOSDOS0.JOENTT = ""
lYGOSDOS0.JODATE = ""

lYGOSDOS0.GOSDOSIDD = rsADO("GOSDOSIDD")
lYGOSDOS0.GOSDOSWBIC = rsADO("GOSDOSWBIC")
lYGOSDOS0.GOSDOSWES = rsADO("GOSDOSWES")
lYGOSDOS0.GOSDOSWMTK = rsADO("GOSDOSWMTK")
lYGOSDOS0.GOSDOSWTRN = rsADO("GOSDOSWTRN")
lYGOSDOS0.GOSDOSWMTD = rsADO("GOSDOSWMTD")
lYGOSDOS0.GOSDOSWDEV = rsADO("GOSDOSWDEV")
lYGOSDOS0.GOSDOSWID1 = rsADO("GOSDOSWID1")
lYGOSDOS0.GOSDOSWIDL = rsADO("GOSDOSWIDL")
lYGOSDOS0.GOSDOSWIDH = rsADO("GOSDOSWIDH")
lYGOSDOS0.GOSDOSRCOM = rsADO("GOSDOSRCOM")
lYGOSDOS0.GOSDOSCLI = rsADO("GOSDOSCLI")
lYGOSDOS0.GOSDOSPAYS = rsADO("GOSDOSPAYS")
lYGOSDOS0.GOSDOSLABK = rsADO("GOSDOSLABK")
lYGOSDOS0.GOSDOSSTAG = rsADO("GOSDOSSTAG")
lYGOSDOS0.GOSDOSSTAD = rsADO("GOSDOSSTAD")
lYGOSDOS0.GOSDOSECHD = rsADO("GOSDOSECHD")
lYGOSDOS0.GOSDOSISRV = rsADO("GOSDOSISRV")
lYGOSDOS0.GOSDOSGSRV = rsADO("GOSDOSGSRV")
lYGOSDOS0.GOSDOSIAMJ = rsADO("GOSDOSIAMJ")
lYGOSDOS0.GOSDOSITOP = rsADO("GOSDOSITOP")
lYGOSDOS0.GOSDOSUSRV = rsADO("GOSDOSUSRV")

lYGOSDOS0.GOSDOSUUSR = rsADO("GOSDOSUUSR")
lYGOSDOS0.GOSDOSUAMJ = rsADO("GOSDOSUAMJ")
lYGOSDOS0.GOSDOSUHMS = rsADO("GOSDOSUHMS")
lYGOSDOS0.GOSDOSUSEQ = rsADO("GOSDOSUSEQ")

Exit Function
Error_Handler:
rsYGOSDOS0_GetBuffer = Error


End Function
'---------------------------------------------------------
Public Function rsJGOSDOS0_GetBuffer(rsADO As ADODB.Recordset, rsYGOSDOS0 As typeYGOSDOS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJGOSDOS0_GetBuffer = Null

rsJGOSDOS0_GetBuffer = rsYGOSDOS0_GetBuffer(rsADO, rsYGOSDOS0)
rsYGOSDOS0.JORCV = rsADO("JORCV")
rsYGOSDOS0.JOSEQN = rsADO("JOSEQN")
rsYGOSDOS0.JRNBIATRN = rsADO("JRNBIATRN")
rsYGOSDOS0.JOENTT = rsADO("JOENTT")
rsYGOSDOS0.JODATE = rsADO("JODATE")

Exit Function

Error_Handler:

rsJGOSDOS0_GetBuffer = Error

End Function


Public Function rsYGOSDOS0_Init(lYGOSDOS0 As typeYGOSDOS0)

lYGOSDOS0.GOSDOSIDD = 0
lYGOSDOS0.GOSDOSIDD = 0
lYGOSDOS0.GOSDOSWID1 = 0
lYGOSDOS0.GOSDOSWIDL = 0
lYGOSDOS0.GOSDOSWIDH = 0
lYGOSDOS0.GOSDOSWBIC = ""
lYGOSDOS0.GOSDOSWES = ""
lYGOSDOS0.GOSDOSWMTK = ""
lYGOSDOS0.GOSDOSWTRN = ""
lYGOSDOS0.GOSDOSWMTD = 0
lYGOSDOS0.GOSDOSWDEV = ""
lYGOSDOS0.GOSDOSRCOM = ""
lYGOSDOS0.GOSDOSCLI = ""
lYGOSDOS0.GOSDOSPAYS = ""
lYGOSDOS0.GOSDOSLABK = ""
lYGOSDOS0.GOSDOSSTAG = ""
lYGOSDOS0.GOSDOSSTAD = ""
lYGOSDOS0.GOSDOSECHD = 0
lYGOSDOS0.GOSDOSISRV = ""
lYGOSDOS0.GOSDOSIAMJ = 0
lYGOSDOS0.GOSDOSITOP = ""
lYGOSDOS0.GOSDOSGSRV = ""
lYGOSDOS0.GOSDOSUSRV = ""
lYGOSDOS0.GOSDOSUUSR = ""
lYGOSDOS0.GOSDOSUAMJ = 0
lYGOSDOS0.GOSDOSUHMS = 0
lYGOSDOS0.GOSDOSUSEQ = 0
    
lYGOSDOS0.GOSDOSUUSR = ""
lYGOSDOS0.GOSDOSUAMJ = 0
lYGOSDOS0.GOSDOSUHMS = 0
lYGOSDOS0.GOSDOSUSEQ = 0


lYGOSDOS0.JORCV = 0
lYGOSDOS0.JOSEQN = 0
lYGOSDOS0.JRNBIATRN = 0
    
lYGOSDOS0.JOENTT = ""
lYGOSDOS0.JODATE = ""

End Function







