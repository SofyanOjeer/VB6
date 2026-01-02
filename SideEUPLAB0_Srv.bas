Attribute VB_Name = "srvSideEUPLAB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeSideEUPLAB0
 
      EUPLABID    As String
      EUPLABBICE  As String
      EUPLABNOME  As String
      EUPLABNOM2  As String
      EUPLABLIB   As String
      EUPLABMONT  As Currency
      EUPLABDEVI  As String
      EUPLABSTAI  As String
      EUPLABSTAS1  As String
      EUPLABSTAS2  As String
      EUPLABSTAS3  As String
      EUPLABSTAS4  As String
      EUPLABSTAS5  As String
      EUPLABSTAS6  As String
      EUPLABSTAS7  As String
      EUPLABSTAS8  As String

End Type
Public xSideEUPLAB0 As typeSideEUPLAB0
Public Function sqlSideEUPLAB0_Insert(newY As typeSideEUPLAB0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlSideEUPLAB0_Insert = Null

xSet = " (EUPLABID"
xValues = " values('" & newY.EUPLABID & "'"

' Détecter les modifications
'===================================================================================
If Trim(newY.EUPLABBICE) <> "" Then xSet = xSet & ",EUPLABBICE": xValues = xValues & " ,'" & Replace(Trim(newY.EUPLABBICE), "'", "''") & "'"
If Trim(newY.EUPLABNOME) <> "" Then xSet = xSet & ",EUPLABNOME": xValues = xValues & " ,'" & Replace(Trim(newY.EUPLABNOME), "'", "''") & "'"
If Trim(newY.EUPLABNOM2) <> "" Then xSet = xSet & ",EUPLABNOM2": xValues = xValues & " ,'" & Replace(Trim(newY.EUPLABNOM2), "'", "''") & "'"
If Trim(newY.EUPLABLIB) <> "" Then xSet = xSet & ",EUPLABLIB": xValues = xValues & " ,'" & Replace(Trim(newY.EUPLABLIB), "'", "''") & "'"
If newY.EUPLABMONT <> 0 Then xSet = xSet & ",EUPLABMONT": xValues = xValues & " ," & cur_P(newY.EUPLABMONT)
If Trim(newY.EUPLABDEVI) <> "" Then xSet = xSet & ",EUPLABDEVI": xValues = xValues & " ,'" & Replace(Trim(newY.EUPLABDEVI), "'", "''") & "'"
If Trim(newY.EUPLABSTAI) <> "" Then xSet = xSet & ",EUPLABSTAI": xValues = xValues & " ,'" & Replace(Trim(newY.EUPLABSTAI), "'", "''") & "'"
If Trim(newY.EUPLABSTAS1) <> "" Then xSet = xSet & ",EUPLABSTAS1": xValues = xValues & " ,'" & Replace(Trim(newY.EUPLABSTAS1), "'", "''") & "'"
If Trim(newY.EUPLABSTAS2) <> "" Then xSet = xSet & ",EUPLABSTAS2": xValues = xValues & " ,'" & Replace(Trim(newY.EUPLABSTAS2), "'", "''") & "'"

xSql = "Insert into   " & paramODBC_SideEUPLAB0 & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnAdo.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlSideEUPLAB0_Insert = "Erreur màj : " & newY.EUPLABID
    Exit Function
End If
 
Exit Function
Error_Handler:
    
    sqlSideEUPLAB0_Insert = Error
End Function

Public Function sqlSideEUPLAB0_Delete(oldY As typeSideEUPLAB0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlSideEUPLAB0_Delete = Null
xWhere = " where EUPLABID = '" & oldY.EUPLABID & "'"

xSql = "Delete from " & paramODBC_SideEUPLAB0 & xWhere
Call FEU_ROUGE
Set rsADO = cnAdo.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlSideEUPLAB0_Delete = "Erreur màj : " & oldY.EUPLABID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlSideEUPLAB0_Delete = Error
End Function

Public Function sqlSideEUPLAB0_Update(newY As typeSideEUPLAB0, oldY As typeSideEUPLAB0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlSideEUPLAB0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.EUPLABID <> newY.EUPLABID Then
    sqlSideEUPLAB0_Update = "Erreur EUPLABID : " & newY.EUPLABID & " / " & oldY.EUPLABID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where EUPLABID = " & oldY.EUPLABID '& " and EUPLABUPDS = " & oldY.EUPLABUPDS

'newY.EUPLABUPDS = newY.EUPLABUPDS + 1
'xSet = xSet & " set EUPLABUPDS = " & newY.EUPLABUPDS

' Détecter les modifications
'==================================================================================== '" & Replace(Trim(newY.EUPLABBICE), "'", "''") & "'"
If newY.EUPLABBICE <> oldY.EUPLABBICE Then xSet = xSet & " , EUPLABBICE = '" & Replace(Trim(newY.EUPLABBICE), "'", "''") & "'"
If newY.EUPLABNOME <> oldY.EUPLABNOME Then xSet = xSet & " , EUPLABNOME = '" & Replace(Trim(newY.EUPLABNOME), "'", "''") & "'"
If newY.EUPLABNOM2 <> oldY.EUPLABNOM2 Then xSet = xSet & " , EUPLABNOM2 = '" & Replace(Trim(newY.EUPLABNOM2), "'", "''") & "'"
If newY.EUPLABLIB <> oldY.EUPLABLIB Then xSet = xSet & " , EUPLABLIB = '" & Replace(Trim(newY.EUPLABLIB), "'", "''") & "'"
If newY.EUPLABMONT <> oldY.EUPLABMONT Then xSet = xSet & " , EUPLABMONT = " & cur_P(newY.EUPLABMONT)
If newY.EUPLABDEVI <> oldY.EUPLABDEVI Then xSet = xSet & " , EUPLADEVI = '" & Replace(Trim(newY.EUPLABDEVI), "'", "''") & "'"
If newY.EUPLABSTAI <> oldY.EUPLABSTAI Then xSet = xSet & " , EUPLABSTAI = '" & Replace(Trim(newY.EUPLABSTAI), "'", "''") & "'"
If newY.EUPLABSTAS2 <> oldY.EUPLABSTAS2 Then xSet = xSet & " , EUPLABSTAS2 = '" & Replace(Trim(newY.EUPLABSTAS2), "'", "''") & "'"

xSql = "update  " & paramODBC_SideEUPLAB0 & xSet & xWhere
Call FEU_ROUGE
Set rsADO = cnAdo.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlSideEUPLAB0_Update = "Erreur màj : " & newY.EUPLABID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlSideEUPLAB0_Update = Error
End Function

Public Function rsSideEUPLAB0_GetBuffer(rsADO As ADODB.Recordset, lSideEUPLAB0 As typeSideEUPLAB0)
On Error GoTo Error_Handler
Dim V
rsSideEUPLAB0_GetBuffer = Null
rsSideEUPLAB0_Init lSideEUPLAB0
lSideEUPLAB0.EUPLABID = rsADO("EUPLABID")
V = rsADO("EUPLABBICE"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABBICE = V
V = rsADO("EUPLABNOME"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABNOME = V
V = rsADO("EUPLABNOM2"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABNOM2 = V
V = rsADO("EUPLABLIB"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABLIB = V
lSideEUPLAB0.EUPLABMONT = rsADO("EUPLABMONT")
V = rsADO("EUPLABDEVI"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABDEVI = V
V = rsADO("EUPLABSTAI"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAI = V
V = rsADO("EUPLABSTAS1"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAS1 = V
V = rsADO("EUPLABSTAS2"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAS2 = V
V = rsADO("EUPLABSTAS3"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAS3 = V
V = rsADO("EUPLABSTAS4"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAS4 = V
V = rsADO("EUPLABSTAS5"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAS5 = V
V = rsADO("EUPLABSTAS6"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAS6 = V
V = rsADO("EUPLABSTAS7"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAS7 = V
V = rsADO("EUPLABSTAS8"): If Not IsNull(V) Then lSideEUPLAB0.EUPLABSTAS8 = V

Exit Function
Error_Handler:
rsSideEUPLAB0_GetBuffer = Error


End Function

Public Function rsSideEUPLAB0_Init(lSideEUPLAB0 As typeSideEUPLAB0)
lSideEUPLAB0.EUPLABID = ""
lSideEUPLAB0.EUPLABBICE = ""
lSideEUPLAB0.EUPLABNOME = ""
lSideEUPLAB0.EUPLABNOM2 = ""
lSideEUPLAB0.EUPLABLIB = ""
lSideEUPLAB0.EUPLABMONT = 0
lSideEUPLAB0.EUPLABDEVI = ""
lSideEUPLAB0.EUPLABSTAI = ""
lSideEUPLAB0.EUPLABSTAS1 = "0"
lSideEUPLAB0.EUPLABSTAS2 = ""
lSideEUPLAB0.EUPLABSTAS3 = ""
lSideEUPLAB0.EUPLABSTAS4 = ""
lSideEUPLAB0.EUPLABSTAS5 = ""
lSideEUPLAB0.EUPLABSTAS6 = ""
lSideEUPLAB0.EUPLABSTAS7 = ""
lSideEUPLAB0.EUPLABSTAS8 = ""

End Function



