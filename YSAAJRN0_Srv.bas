Attribute VB_Name = "srvYSAAJRN0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
Type typeYSAAJRN0
 
      SAAJRNAID   As Long
      
      SAAJRNAMJH   As Double
      SAAJRNSEQ    As Double
      SAAJRNEVEC   As String
      
      SAAJRNEVEN   As Long
      SAAJRNTOPK   As String
      SAAJRNTOPX   As String
      SAAJRNSUFX    As Double
'____________________________________________________ Journalisation
End Type
Public xYSAAJRN0 As typeYSAAJRN0
Public Function sqlYSAAJRN0_Delete(oldY As typeYSAAJRN0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSAAJRN0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SAAJRNAID = " & oldY.SAAJRNAID _
       & " and   SAAJRNAMJH = " & oldY.SAAJRNAMJH _
       & " and   SAAJRNSEQ =  " & oldY.SAAJRNSEQ


'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YSAAJRN0" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSAAJRN0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYSAAJRN0_Delete = Error
End Function

Public Function sqlYSAAJRN0_Update(newY As typeYSAAJRN0, oldY As typeYSAAJRN0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSAAJRN0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SAAJRNAID <> newY.SAAJRNAID _
Or oldY.SAAJRNAMJH <> newY.SAAJRNAMJH _
Or oldY.SAAJRNSEQ <> newY.SAAJRNSEQ Then
    sqlYSAAJRN0_Update = "Erreur SAAJRNAID"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SAAJRNAID = " & oldY.SAAJRNAID

xSet = " set"
blnUpdate = False


' Détecter les modifications
'===================================================================================
If newY.SAAJRNAMJH <> oldY.SAAJRNAMJH Then blnUpdate = True:  xSet = xSet & " , SAAJRNAMJH = " & newY.SAAJRNAMJH
If newY.SAAJRNSEQ <> oldY.SAAJRNSEQ Then blnUpdate = True:  xSet = xSet & " , SAAJRNSEQ = " & newY.SAAJRNSEQ
If newY.SAAJRNEVEN <> oldY.SAAJRNEVEN Then blnUpdate = True:  xSet = xSet & " , SAAJRNEVEN = " & newY.SAAJRNEVEN
If newY.SAAJRNSUFX <> oldY.SAAJRNSUFX Then blnUpdate = True:  xSet = xSet & " , SAAJRNSUFX = " & newY.SAAJRNSUFX

If newY.SAAJRNEVEC <> oldY.SAAJRNEVEC Then blnUpdate = True:  xSet = xSet & " , SAAJRNEVEC = '" & newY.SAAJRNEVEC & "'"
If newY.SAAJRNTOPX <> oldY.SAAJRNTOPX Then blnUpdate = True:  xSet = xSet & " , SAAJRNTOPX = '" & newY.SAAJRNTOPX & "'"
If newY.SAAJRNTOPX <> oldY.SAAJRNTOPX Then blnUpdate = True:  xSet = xSet & " , SAAJRNTOPX = '" & Replace(newY.SAAJRNTOPX, "'", "''") & "'"

If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSAAJRN0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSAAJRN0_Update = "Erreur màj : " & newY.SAAJRNAID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSAAJRN0_Update = Error
End Function

Public Function sqlYSAAJRN0_Insert(newY As typeYSAAJRN0)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSAAJRN0_Insert = Null
xSet = " (SAAJRNAID "
xValues = " values(" & newY.SAAJRNAID

' Détecter les modifications
'===================================================================================
If Trim(newY.SAAJRNAMJH) <> "" Then xSet = xSet & ",SAAJRNAMJH": xValues = xValues & " ," & newY.SAAJRNAMJH
If Trim(newY.SAAJRNSEQ) <> "" Then xSet = xSet & ",SAAJRNSEQ": xValues = xValues & " ," & newY.SAAJRNSEQ
If Trim(newY.SAAJRNEVEN) <> "" Then xSet = xSet & ",SAAJRNEVEN": xValues = xValues & " ," & newY.SAAJRNEVEN
If Trim(newY.SAAJRNSUFX) <> "" Then xSet = xSet & ",SAAJRNSUFX": xValues = xValues & " ," & newY.SAAJRNSUFX

If Trim(newY.SAAJRNEVEC) <> "" Then xSet = xSet & ",SAAJRNEVEC": xValues = xValues & " ,'" & newY.SAAJRNEVEC & "'"

If Trim(newY.SAAJRNTOPX) <> "" Then xSet = xSet & ",SAAJRNTOPX": xValues = xValues & " ,'" & Replace(newY.SAAJRNTOPX, "'", "''") & "'"
If Trim(newY.SAAJRNTOPK) <> "" Then xSet = xSet & ",SAAJRNTOPK": xValues = xValues & " ,'" & newY.SAAJRNTOPK & "'"

xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YSAAJRN0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSAAJRN0_Insert = "Erreur màj : " & newY.SAAJRNAID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSAAJRN0_Insert = Error
End Function

Public Function rsYSAAJRN0_GetBuffer(rsADO As ADODB.Recordset, lYSAAJRN0 As typeYSAAJRN0)
On Error GoTo Error_Handler
rsYSAAJRN0_GetBuffer = Null

lYSAAJRN0.SAAJRNAID = rsADO("SAAJRNAID")

lYSAAJRN0.SAAJRNAMJH = rsADO("SAAJRNAMJH")
lYSAAJRN0.SAAJRNSEQ = rsADO("SAAJRNSEQ")
lYSAAJRN0.SAAJRNEVEC = rsADO("SAAJRNEVEC")

lYSAAJRN0.SAAJRNEVEN = rsADO("SAAJRNEVEN")
lYSAAJRN0.SAAJRNTOPK = rsADO("SAAJRNTOPK")
lYSAAJRN0.SAAJRNTOPX = rsADO("SAAJRNTOPX")

lYSAAJRN0.SAAJRNSUFX = rsADO("SAAJRNSUFX")

Exit Function
Error_Handler:
rsYSAAJRN0_GetBuffer = Error


End Function
Public Function rsYSAAJRN0_Init(lYSAAJRN0 As typeYSAAJRN0)

lYSAAJRN0.SAAJRNAID = 0

lYSAAJRN0.SAAJRNAMJH = 0
lYSAAJRN0.SAAJRNSEQ = 0
lYSAAJRN0.SAAJRNEVEC = ""

lYSAAJRN0.SAAJRNEVEN = 0
lYSAAJRN0.SAAJRNTOPK = ""
lYSAAJRN0.SAAJRNTOPX = ""

lYSAAJRN0.SAAJRNSUFX = 0

End Function


