Attribute VB_Name = "srvYSWILNK0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYSWILNK0
 
      SWILNKSWID   As Long
      
      SWILNKAPPC    As String
      SWILNKAPPN    As Long
      SWILNKSTA    As String
    
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    
    JOENTT          As String * 2
    JODATE          As String * 6

'____________________________________________________ Journalisation
End Type
Public xYSWILNK0 As typeYSWILNK0
Public Function sqlYSWILNK0_Delete(oldY As typeYSWILNK0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlYSWILNK0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWILNKSWID = " & oldY.SWILNKSWID

'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YSWILNK0" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWILNK0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYSWILNK0_Delete = Error
End Function

Public Function sqlYSWILNK0_Update(newY As typeYSWILNK0, oldY As typeYSWILNK0, blnUUSR As Boolean)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlYSWILNK0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SWILNKSWID <> newY.SWILNKSWID Then
    sqlYSWILNK0_Update = "Erreur SWILNKSWID : " & newY.SWILNKSWID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWILNKSWID = " & oldY.SWILNKSWID

xSet = " set" 'xSet & " set SWILNKWIDH = " & newY.SWILNKWIDH
blnUpdate = False


' Détecter les modifications
'===================================================================================

If newY.SWILNKAPPC <> oldY.SWILNKAPPC Then blnUpdate = True:  xSet = xSet & " , SWILNKAPPC = '" & newY.SWILNKAPPC & "'"
If newY.SWILNKAPPN <> oldY.SWILNKAPPN Then blnUpdate = True:  xSet = xSet & " , SWILNKAPPN = " & newY.SWILNKAPPN
If newY.SWILNKSTA <> oldY.SWILNKSTA Then blnUpdate = True:  xSet = xSet & " , SWILNKSTA = '" & newY.SWILNKSTA & "'"


If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWILNK0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWILNK0_Update = "Erreur màj : " & newY.SWILNKSWID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSWILNK0_Update = Error
End Function

Public Function sqlYSWILNK0_Insert(newY As typeYSWILNK0)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

'On Error GoTo Error_Handler
sqlYSWILNK0_Insert = Null
xSet = " (SWILNKSWID "
xValues = " values(" & newY.SWILNKSWID

' Détecter les modifications
'===================================================================================

If newY.SWILNKAPPN <> 0 Then xSet = xSet & ",SWILNKAPPN": xValues = xValues & " ," & newY.SWILNKAPPN


If Trim(newY.SWILNKAPPC) <> "" Then xSet = xSet & ",SWILNKAPPC": xValues = xValues & " ,'" & newY.SWILNKAPPC & "'"
If Trim(newY.SWILNKSTA) <> "" Then xSet = xSet & ",SWILNKSTA": xValues = xValues & " ,'" & newY.SWILNKSTA & "'"
xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YSWILNK0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWILNK0_Insert = "Erreur màj : " & newY.SWILNKSWID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWILNK0_Insert = Error
End Function

Public Function rsYSWILNK0_GetBuffer(rsADO As ADODB.Recordset, lYSWILNK0 As typeYSWILNK0)
On Error GoTo Error_Handler
rsYSWILNK0_GetBuffer = Null

lYSWILNK0.JORCV = 0
lYSWILNK0.JOSEQN = 0
lYSWILNK0.JRNBIATRN = 0
lYSWILNK0.JOENTT = ""
lYSWILNK0.JODATE = ""

lYSWILNK0.SWILNKSWID = rsADO("SWILNKSWID")

lYSWILNK0.SWILNKAPPC = rsADO("SWILNKAPPC")
lYSWILNK0.SWILNKAPPN = rsADO("SWILNKAPPN")

lYSWILNK0.SWILNKSTA = rsADO("SWILNKSTA")

Exit Function
Error_Handler:
rsYSWILNK0_GetBuffer = Error


End Function
'---------------------------------------------------------
Public Function rsJSWILNK0_GetBuffer(rsADO As ADODB.Recordset, rsYSWILNK0 As typeYSWILNK0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJSWILNK0_GetBuffer = Null

rsJSWILNK0_GetBuffer = rsYSWILNK0_GetBuffer(rsADO, rsYSWILNK0)
rsYSWILNK0.JORCV = rsADO("JORCV")
rsYSWILNK0.JOSEQN = rsADO("JOSEQN")
rsYSWILNK0.JRNBIATRN = rsADO("JRNBIATRN")
rsYSWILNK0.JOENTT = rsADO("JOENTT")
rsYSWILNK0.JODATE = rsADO("JODATE")

Exit Function

Error_Handler:

rsJSWILNK0_GetBuffer = Error

End Function


Public Function rsYSWILNK0_Init(lYSWILNK0 As typeYSWILNK0)


lYSWILNK0.SWILNKSWID = 0
      
lYSWILNK0.SWILNKAPPC = ""
lYSWILNK0.SWILNKAPPN = 0
lYSWILNK0.SWILNKSTA = ""



lYSWILNK0.JORCV = 0
lYSWILNK0.JOSEQN = 0
lYSWILNK0.JRNBIATRN = 0
    
lYSWILNK0.JOENTT = ""
lYSWILNK0.JODATE = ""

End Function











