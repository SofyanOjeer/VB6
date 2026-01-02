Attribute VB_Name = "srvYGOSEVE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYGOSEVE0
 
      GOSEVEIDD    As Long
      GOSEVEIDE    As Long
      
      GOSEVESWID   As Long
      GOSEVESTAE   As String
      GOSEVEGSRV   As String
      GOSEVEUSRV   As String
      GOSEVEUUSR   As String
      GOSEVEUAMJ   As Long
      GOSEVEUHMS   As Long
      GOSEVEUSEQ   As Long
      
      GOSEVENAT   As String
      GOSEVETXT   As String

    
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    
    JOENTT          As String * 2
    JODATE          As String * 6

'____________________________________________________ Journalisation
End Type
Public xYGOSEVE0 As typeYGOSEVE0
Public Function sqlYGOSEVE0_Delete(oldY As typeYGOSEVE0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlYGOSEVE0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where GOSEVEIDD = " & oldY.GOSEVEIDD _
       & " and GOSEVEIDE = " & oldY.GOSEVEIDE _
       & " and GOSEVEUSEQ = " & oldY.GOSEVEUSEQ

'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YGOSEVE0" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYGOSEVE0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYGOSEVE0_Delete = Error
End Function

Public Function sqlYGOSEVE0_Update(newY As typeYGOSEVE0, oldY As typeYGOSEVE0, blnUUSR As Boolean)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlYGOSEVE0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.GOSEVEIDD <> newY.GOSEVEIDD _
Or oldY.GOSEVEIDE <> newY.GOSEVEIDE _
Or oldY.GOSEVEUSEQ <> newY.GOSEVEUSEQ Then
    sqlYGOSEVE0_Update = "Erreur GOSEVEIDD : " & newY.GOSEVEIDD & "." & oldY.GOSEVEIDE
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where GOSEVEIDD = " & oldY.GOSEVEIDD _
       & " and GOSEVEIDE = " & oldY.GOSEVEIDE _
       & " and GOSEVEUSEQ = " & oldY.GOSEVEUSEQ

newY.GOSEVEUSEQ = newY.GOSEVEUSEQ + 1
xSet = xSet & " set GOSEVEUSEQ = " & newY.GOSEVEUSEQ
blnUpdate = False

If blnUUSR Then
    newY.GOSEVEUUSR = usrName_UCase
    newY.GOSEVEUSRV = currentSSIWINUNIT
    newY.GOSEVEUAMJ = DSys
    newY.GOSEVEUHMS = time_Hms
End If

' Détecter les modifications
'===================================================================================
If newY.GOSEVESWID <> oldY.GOSEVESWID Then blnUpdate = True:  xSet = xSet & " , GOSEVESWID = " & newY.GOSEVESWID
If newY.GOSEVEUAMJ <> oldY.GOSEVEUAMJ Then blnUpdate = True:  xSet = xSet & " , GOSEVEUAMJ = " & newY.GOSEVEUAMJ
If newY.GOSEVEUHMS <> oldY.GOSEVEUHMS Then blnUpdate = True:  xSet = xSet & " , GOSEVEUHMS = " & newY.GOSEVEUHMS

If newY.GOSEVESTAE <> oldY.GOSEVESTAE Then blnUpdate = True:  xSet = xSet & " , GOSEVESTAE = '" & newY.GOSEVESTAE & "'"
If newY.GOSEVEGSRV <> oldY.GOSEVEGSRV Then blnUpdate = True:  xSet = xSet & " , GOSEVEGSRV = '" & newY.GOSEVEGSRV & "'"
If newY.GOSEVEUSRV <> oldY.GOSEVEUSRV Then blnUpdate = True:  xSet = xSet & " , GOSEVEUSRV = '" & newY.GOSEVEUSRV & "'"
If newY.GOSEVEUUSR <> oldY.GOSEVEUUSR Then blnUpdate = True:  xSet = xSet & " , GOSEVEUUSR = '" & newY.GOSEVEUUSR & "'"
If newY.GOSEVENAT <> oldY.GOSEVENAT Then blnUpdate = True:  xSet = xSet & " , GOSEVENAT = '" & newY.GOSEVENAT & "'"
If newY.GOSEVETXT <> oldY.GOSEVETXT Then blnUpdate = True:  xSet = xSet & " , GOSEVETXT = '" & Replace(newY.GOSEVETXT, "'", "''") & "'"


If blnUpdate Then
    
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YGOSEVE0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYGOSEVE0_Update = "Erreur màj : " & newY.GOSEVEIDD & "." & oldY.GOSEVEIDE
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYGOSEVE0_Update = Error
End Function

Public Function sqlYGOSEVE0_Insert(newY As typeYGOSEVE0)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

'On Error GoTo Error_Handler
sqlYGOSEVE0_Insert = Null
xSet = " (GOSEVEIDD,GOSEVEIDE"
xValues = " values(" & newY.GOSEVEIDD & " ," & newY.GOSEVEIDE

newY.GOSEVEUUSR = usrName_UCase
newY.GOSEVEUSRV = currentSSIWINUNIT
newY.GOSEVEUAMJ = DSys
newY.GOSEVEUHMS = time_Hms

' Détecter les modifications
'===================================================================================

If newY.GOSEVESWID <> 0 Then xSet = xSet & ",GOSEVESWID": xValues = xValues & " ," & newY.GOSEVESWID
If newY.GOSEVEUSEQ <> 0 Then xSet = xSet & ",GOSEVEUSEQ": xValues = xValues & " ," & newY.GOSEVEUSEQ
If newY.GOSEVEUAMJ <> 0 Then xSet = xSet & ",GOSEVEUAMJ": xValues = xValues & " ," & newY.GOSEVEUAMJ
If newY.GOSEVEUHMS <> 0 Then xSet = xSet & ",GOSEVEUHMS": xValues = xValues & " ," & newY.GOSEVEUHMS


If Trim(newY.GOSEVESTAE) <> "" Then xSet = xSet & ",GOSEVESTAE": xValues = xValues & " ,'" & newY.GOSEVESTAE & "'"
If Trim(newY.GOSEVEGSRV) <> "" Then xSet = xSet & ",GOSEVEGSRV": xValues = xValues & " ,'" & newY.GOSEVEGSRV & "'"
If Trim(newY.GOSEVEUSRV) <> "" Then xSet = xSet & ",GOSEVEUSRV": xValues = xValues & " ,'" & newY.GOSEVEUSRV & "'"

If Trim(newY.GOSEVEUUSR) <> "" Then xSet = xSet & ",GOSEVEUUSR": xValues = xValues & " ,'" & newY.GOSEVEUUSR & "'"
If Trim(newY.GOSEVENAT) <> "" Then xSet = xSet & ",GOSEVENAT": xValues = xValues & " ,'" & newY.GOSEVENAT & "'"
If Trim(newY.GOSEVETXT) <> "" Then xSet = xSet & ",GOSEVETXT": xValues = xValues & " ,'" & Replace(newY.GOSEVETXT, "'", "''") & "'"

xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YGOSEVE0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYGOSEVE0_Insert = "Erreur màj : " & newY.GOSEVEIDD
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYGOSEVE0_Insert = Error
End Function

Public Function rsYGOSEVE0_GetBuffer(rsADO As ADODB.Recordset, lYGOSEVE0 As typeYGOSEVE0)
On Error GoTo Error_Handler
rsYGOSEVE0_GetBuffer = Null

lYGOSEVE0.JORCV = 0
lYGOSEVE0.JOSEQN = 0
lYGOSEVE0.JRNBIATRN = 0
lYGOSEVE0.JOENTT = ""
lYGOSEVE0.JODATE = ""

lYGOSEVE0.GOSEVEIDD = rsADO("GOSEVEIDD")
lYGOSEVE0.GOSEVEIDE = rsADO("GOSEVEIDE")

lYGOSEVE0.GOSEVESWID = rsADO("GOSEVESWID")
lYGOSEVE0.GOSEVESTAE = rsADO("GOSEVESTAE")
lYGOSEVE0.GOSEVEGSRV = rsADO("GOSEVEGSRV")
lYGOSEVE0.GOSEVEUSRV = rsADO("GOSEVEUSRV")

lYGOSEVE0.GOSEVEUUSR = rsADO("GOSEVEUUSR")
lYGOSEVE0.GOSEVEUAMJ = rsADO("GOSEVEUAMJ")
lYGOSEVE0.GOSEVEUHMS = rsADO("GOSEVEUHMS")
lYGOSEVE0.GOSEVEUSEQ = rsADO("GOSEVEUSEQ")

lYGOSEVE0.GOSEVENAT = rsADO("GOSEVENAT")
lYGOSEVE0.GOSEVETXT = rsADO("GOSEVETXT")

Exit Function
Error_Handler:
rsYGOSEVE0_GetBuffer = Error


End Function
'---------------------------------------------------------
Public Function rsJGOSEVE0_GetBuffer(rsADO As ADODB.Recordset, rsYGOSEVE0 As typeYGOSEVE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJGOSEVE0_GetBuffer = Null

rsJGOSEVE0_GetBuffer = rsYGOSEVE0_GetBuffer(rsADO, rsYGOSEVE0)
rsYGOSEVE0.JORCV = rsADO("JORCV")
rsYGOSEVE0.JOSEQN = rsADO("JOSEQN")
rsYGOSEVE0.JRNBIATRN = rsADO("JRNBIATRN")
rsYGOSEVE0.JOENTT = rsADO("JOENTT")
rsYGOSEVE0.JODATE = rsADO("JODATE")

Exit Function

Error_Handler:

rsJGOSEVE0_GetBuffer = Error

End Function


Public Function rsYGOSEVE0_Init(lYGOSEVE0 As typeYGOSEVE0)


lYGOSEVE0.GOSEVEIDD = 0
lYGOSEVE0.GOSEVEIDE = 0
      
lYGOSEVE0.GOSEVESWID = 0
lYGOSEVE0.GOSEVESTAE = ""
lYGOSEVE0.GOSEVEGSRV = ""
lYGOSEVE0.GOSEVEUSRV = ""
lYGOSEVE0.GOSEVEUUSR = ""
lYGOSEVE0.GOSEVEUAMJ = 0
lYGOSEVE0.GOSEVEUHMS = 0
lYGOSEVE0.GOSEVEUSEQ = 0
      
lYGOSEVE0.GOSEVENAT = ""
lYGOSEVE0.GOSEVETXT = ""




lYGOSEVE0.JORCV = 0
lYGOSEVE0.JOSEQN = 0
lYGOSEVE0.JRNBIATRN = 0
    
lYGOSEVE0.JOENTT = ""
lYGOSEVE0.JODATE = ""

End Function



Public Sub rsYGOSEVE0_Init_Param(lYGOSEVE0 As typeYGOSEVE0)
Call rsYGOSEVE0_Init(lYGOSEVE0)
lYGOSEVE0.GOSEVEIDD = -1
lYGOSEVE0.GOSEVESTAE = "I"
lYGOSEVE0.GOSEVEGSRV = currentSSIWINUNIT
End Sub






