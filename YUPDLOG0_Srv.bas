Attribute VB_Name = "srvYUPDLOG0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYUPDLOG0
 
      UPDLOGID    As Long         'IDENTIFICATION
      UPDLOGAMJ   As Long         'date maj
      UPDLOGHMS   As Long         'date hms
      UPDLOGUSR   As String * 12  'utilisateur
      UPDLOGAPP   As String * 12  'application
      UPDLOGFCT   As String * 12  'fonction
      UPDLOGTXT   As String * 32  'informations
      UPDLOGUPDS  As Long         'Sequence mise à jour

End Type
Public xYUPDLOG0 As typeYUPDLOG0
Public Function sqlYUPDLOG0_Insert(newY As typeYUPDLOG0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYUPDLOG0_Insert = Null

xSet = " (UPDLOGID"
xValues = " values(" & newY.UPDLOGID

' Détecter les modifications
'===================================================================================
If newY.UPDLOGAMJ <> 0 Then xSet = xSet & ",UPDLOGAMJ": xValues = xValues & " ," & newY.UPDLOGAMJ
If newY.UPDLOGHMS <> 0 Then xSet = xSet & ",UPDLOGhms": xValues = xValues & " ," & newY.UPDLOGHMS
If Trim(newY.UPDLOGUSR) <> "" Then xSet = xSet & ",UPDLOGUSR": xValues = xValues & " ,'" & newY.UPDLOGUSR & "'"
If Trim(newY.UPDLOGAPP) <> "" Then xSet = xSet & ",UPDLOGAPP": xValues = xValues & " ,'" & newY.UPDLOGAPP & "'"
If Trim(newY.UPDLOGFCT) <> "" Then xSet = xSet & ",UPDLOGFCT": xValues = xValues & " ,'" & newY.UPDLOGFCT & "'"
If Trim(newY.UPDLOGTXT) <> "" Then xSet = xSet & ",UPDLOGTXT": xValues = xValues & " ,'" & newY.UPDLOGTXT & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YUPDLOG0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYUPDLOG0_Insert = "Erreur màj : " & newY.UPDLOGID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYUPDLOG0_Insert = Error
End Function

Public Function sqlYUPDLOG0_Update(newY As typeYUPDLOG0, oldY As typeYUPDLOG0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYUPDLOG0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.UPDLOGID <> newY.UPDLOGID Then
    sqlYUPDLOG0_Update = "Erreur UPDLOGID : " & newY.UPDLOGID & " / " & oldY.UPDLOGID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where UPDLOGID = " & oldY.UPDLOGID & " and UPDLOGUPDS = " & oldY.UPDLOGUPDS

newY.UPDLOGUPDS = newY.UPDLOGUPDS + 1
xSet = xSet & " set UPDLOGUPDS = " & newY.UPDLOGUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.UPDLOGAMJ <> oldY.UPDLOGAMJ Then blnUpdate = True: xSet = xSet & " , UPDLOGAMJ = " & newY.UPDLOGAMJ
If newY.UPDLOGHMS <> oldY.UPDLOGHMS Then blnUpdate = True: xSet = xSet & " , UPDLOGhms = " & newY.UPDLOGHMS
If newY.UPDLOGUSR <> oldY.UPDLOGUSR Then blnUpdate = True:  xSet = xSet & " , UPDLOGUSR = '" & newY.UPDLOGUSR & "'"
If newY.UPDLOGAPP <> oldY.UPDLOGAPP Then blnUpdate = True:  xSet = xSet & " , UPDLOGAPP = '" & newY.UPDLOGAPP & "'"
If newY.UPDLOGFCT <> oldY.UPDLOGFCT Then blnUpdate = True:  xSet = xSet & " , UPDLOGFCT= '" & newY.UPDLOGFCT & "'"
If newY.UPDLOGTXT <> oldY.UPDLOGTXT Then blnUpdate = True:  xSet = xSet & " , UPDLOGTXT = '" & newY.UPDLOGTXT & "'"

If newY.UPDLOGID < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YUPDLOG0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYUPDLOG0_Update = "Erreur màj : " & newY.UPDLOGID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYUPDLOG0_Update = Error
End Function

Public Function sqlYUPDLOG0_Init(newY As typeYUPDLOG0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim xxx As typeYUPDLOG0

On Error GoTo Error_Handler
sqlYUPDLOG0_Init = Null

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YUPDLOG0" & " where  UPDLOGID =  -1"
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

xxx.UPDLOGUPDS = rsAdo("UPDLOGUPDS")
newY.UPDLOGID = rsAdo("UPDLOGAMJ") + 1
newY.UPDLOGAMJ = DSys
newY.UPDLOGHMS = time_Hms
newY.UPDLOGUPDS = 0

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where UPDLOGID = -1" & " and UPDLOGUPDS = " & xxx.UPDLOGUPDS

xSet = " set UPDLOGUPDS = " & xxx.UPDLOGUPDS + 1 & " , UPDLOGAMJ = " & newY.UPDLOGID


xSQL = "update " & paramIBM_Library_SABSPE & ".YUPDLOG0" & xSet & xWhere
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYUPDLOG0_Init = "Erreur màj : " & newY.UPDLOGID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYUPDLOG0_Init = Error
End Function

Public Function srvYUPDLOG0_GetBuffer_ODBC(rsAdo As ADODB.Recordset, lYUPDLOG0 As typeYUPDLOG0)
On Error GoTo Error_Handler
srvYUPDLOG0_GetBuffer_ODBC = Null
lYUPDLOG0.UPDLOGID = rsAdo("UPDLOGID")
lYUPDLOG0.UPDLOGAMJ = rsAdo("UPDLOGAMJ")
lYUPDLOG0.UPDLOGHMS = rsAdo("UPDLOGHMS")
lYUPDLOG0.UPDLOGUSR = rsAdo("UPDLOGUSR")
lYUPDLOG0.UPDLOGAPP = rsAdo("UPDLOGAPP")
lYUPDLOG0.UPDLOGFCT = rsAdo("UPDLOGFCT")
lYUPDLOG0.UPDLOGTXT = rsAdo("UPDLOGTXT")
lYUPDLOG0.UPDLOGUPDS = rsAdo("UPDLOGUPDS")

Exit Function
Error_Handler:
srvYUPDLOG0_GetBuffer_ODBC = Error


End Function

Public Function srvYUPDLOG0_Init(lYUPDLOG0 As typeYUPDLOG0)
lYUPDLOG0.UPDLOGID = 0
lYUPDLOG0.UPDLOGAMJ = 0
lYUPDLOG0.UPDLOGHMS = 0
lYUPDLOG0.UPDLOGUSR = ""
lYUPDLOG0.UPDLOGAPP = ""
lYUPDLOG0.UPDLOGFCT = ""
lYUPDLOG0.UPDLOGTXT = ""
lYUPDLOG0.UPDLOGUPDS = 0

End Function

