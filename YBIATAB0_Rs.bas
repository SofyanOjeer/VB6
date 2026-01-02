Attribute VB_Name = "rsYBIATAB0"
'---------------------------------------------------------
Option Explicit
Dim rsAdo As ADODB.Recordset
Type typeYBIATAB0

    BIATABID        As String * 12
    BIATABK1        As String * 12
    BIATABK2        As String * 12
    BIATABTXT      As String * 128

End Type


'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIATAB0_GetBuffer(rsAdo As ADODB.Recordset, rsYBIATAB0 As typeYBIATAB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIATAB0_GetBuffer = Null

rsYBIATAB0.BIATABID = rsAdo("BIATABID")
rsYBIATAB0.BIATABK1 = rsAdo("BIATABK1")
rsYBIATAB0.BIATABK2 = rsAdo("BIATABK2")
rsYBIATAB0.BIATABTXT = rsAdo("BIATABTXT")

Exit Function

Error_Handler:

rsYBIATAB0_GetBuffer = Error

End Function

Public Sub rsYBIATAB0_Réplication()
Dim xSQL As String
Dim xYBIATAB0 As typeYBIATAB0
Dim X As String, V
On Error Resume Next
'Set rsMDB = Nothing
rsMDB.Close
xSQL = "delete * from YBIATAB0"
Set rsMDB = cnMDB.Execute(xSQL)
rsMDB.Open "select * from YBIATAB0", cnMDB, , adLockOptimistic

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0"
    Set rsSab = cnsab.Execute(xSQL)
    Do While Not rsSab.EOF
        
        V = rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
        'Debug.Print xYBIATAB0.BIATABID; xYBIATAB0.BIATABK1; xYBIATAB0.BIATABK2
        If IsNull(V) Then
            rsMDB.AddNew
            V = rsYBIATAB0_PutBuffer(rsMDB, xYBIATAB0)
            rsMDB.Update

            If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & " : " & "rsYBIATAB0_Réplication"
        End If
        rsSab.MoveNext
        DoEvents
    Loop

Set rsMDB = Nothing
Set rsSab = Nothing
End Sub

Public Function sqlYBIATAB0_Insert(newY As typeYBIATAB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYBIATAB0_Insert = Null

xSet = " (BIATABID"
xValues = " values('" & newY.BIATABID & "'"
' Détecter les modifications
'===================================================================================

If Trim(newY.BIATABK1) <> "" Then xSet = xSet & ",BIATABk1": xValues = xValues & " ,'" & Replace(newY.BIATABK1, "'", "''") & "'"
If Trim(newY.BIATABK2) <> "" Then xSet = xSet & ",BIATABk2": xValues = xValues & " ,'" & Replace(newY.BIATABK2, "'", "''") & "'"
If Trim(newY.BIATABTXT) <> "" Then xSet = xSet & ",BIATABtxt": xValues = xValues & " ,'" & Replace(newY.BIATABTXT, "'", "''") & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YBIATAB0" & xSet & ")" & xValues & ")"

'Set rsADO = cnSab_Update.Execute(xSql, Nb)
Set rsAdo = cnsab.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBIATAB0_Insert = "Erreur màj : " & newY.BIATABID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBIATAB0_Insert = Error
End Function
Public Function sqlYBIATAB0_Delete(oldY As typeYBIATAB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYBIATAB0_Delete = Null
xWhere = " where BIATABID = '" & Trim(oldY.BIATABID) & "'" _
       & " and BIATABK1 = '" & Trim(oldY.BIATABK1) & "'" _
       & " and BIATABK2 = '" & Trim(oldY.BIATABK2) & "'"
       

xSQL = "Delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 " & xWhere

'Set rsADO = cnSab_Update.Execute(xSql, Nb)
Set rsAdo = cnsab.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBIATAB0_Delete = "Erreur màj : " & oldY.BIATABID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBIATAB0_Delete = Error
End Function

Public Function sqlYBIATAB0_Delete_ID_K1(lBIATABID As String, lBIATABK1 As String)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYBIATAB0_Delete_ID_K1 = Null
xWhere = " where BIATABID = '" & lBIATABID & "'" _
       & " and BIATABK1 = '" & lBIATABK1 & "'"
       

xSQL = "Delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 " & xWhere

'Set rsADO = cnSab_Update.Execute(xSql, Nb)
Set rsAdo = cnsab.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBIATAB0_Delete_ID_K1 = "Erreur màj : " & lBIATABID & " " & lBIATABK1
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBIATAB0_Delete_ID_K1 = Error
End Function

Public Function sqlYBIATAB0_Transaction(lFct As String, newYBIATAB0 As typeYBIATAB0, oldYBIATAB0 As typeYBIATAB0)
Dim xSQL As String
On Error GoTo Error_Handler

Dim V
App_Debug = "sqlYBIATAB0_Transaction : " & lFct

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
Call FEU_ORANGE
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lFct
    Case "Update":
                Call FEU_ROUGE
                V = sqlYBIATAB0_Update(newYBIATAB0, oldYBIATAB0)
    Case "New": V = sqlYBIATAB0_Insert(newYBIATAB0)
    Case "Delete": V = sqlYBIATAB0_Delete(oldYBIATAB0)
    Case "Delete_#SAB": V = sqlYBIATAB0_Delete_ID_K1("CREDOC_#SAB", oldYBIATAB0.BIATABK2)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
Call FEU_VERT
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    sqlYBIATAB0_Transaction = V
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function


Public Function sqlYBIATAB0_Update(newY As typeYBIATAB0, oldY As typeYBIATAB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYBIATAB0_Update = Null

xWhere = " where BIATABID = '" & Trim(oldY.BIATABID) & "'" _
       & " and BIATABK1 = '" & Trim(oldY.BIATABK1) & "'" _
       & " and BIATABK2 = '" & Trim(oldY.BIATABK2) & "'"

xSet = " SET"
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.BIATABK1 <> oldY.BIATABK1 Then blnUpdate = True:  xSet = xSet & " , BIATABK1 = '" & Replace(newY.BIATABK1, "'", "''") & "'"
If newY.BIATABK2 <> oldY.BIATABK2 Then blnUpdate = True:  xSet = xSet & " , BIATABK2 = '" & Replace(newY.BIATABK2, "'", "''") & "'"
If newY.BIATABTXT <> oldY.BIATABTXT Then blnUpdate = True:  xSet = xSet & " , BIATABTXT = '" & Replace(newY.BIATABTXT, "'", "''") & "'"


If blnUpdate Then
    Mid$(xSet, 6, 1) = " "
    xSQL = "update " & paramIBM_Library_SABSPE & ".YBIATAB0" & xSet & xWhere
    
   ' Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_ROUGE
    Set rsAdo = cnsab.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYBIATAB0_Update = "Erreur màj : " & newY.BIATABK1
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYBIATAB0_Update = Error
End Function

'---------------------------------------------------------
Public Function rsYBIATAB0_Read(lId As String, lK1 As String, lK2 As String, lMemo As String)
'---------------------------------------------------------
Dim xYBIATAB0 As typeYBIATAB0
Dim X As String, V
Dim rsMDB As New ADODB.Recordset
On Error GoTo Error_Handler

rsYBIATAB0_Read = Null
lMemo = ""

X = "select * from YBIATAB0 where" _
    & " BIATABID = '" & lId & "'" _
    & " and BIATABK1 = '" & lK1 & "'" _
    & " and BIATABK2 = '" & lK2 & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    V = rsMDB("BIATABTXT")
    If Not IsNull(V) Then lMemo = Trim(V)
Else
    rsYBIATAB0_Read = "? rsYBIATAB0_Read : " & lId & "_" & lK1 & "_" & lK2
End If
Exit Function

Error_Handler:
'-------------
    rsYBIATAB0_Read = " rsYBIATAB0_Read : " & Error
End Function
'---------------------------------------------------------
Public Function sqlYBIATAB0_Read(lId As String, lK1 As String, lK2 As String, lMemo As String)
'---------------------------------------------------------
Dim xYBIATAB0 As typeYBIATAB0
Dim X As String, V
Dim rsSab As New ADODB.Recordset
On Error GoTo Error_Handler

sqlYBIATAB0_Read = Null
lMemo = ""

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & "where BIATABID = '" & lId & "'"
If lK1 <> "" Then X = X & " and BIATABK1 = '" & lK1 & "'"
If lK2 <> "" Then X = X & " and BIATABK2 = '" & lK2 & "'"
    
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    V = rsSab("BIATABTXT")
    If Not IsNull(V) Then lMemo = CStr(V) 'Trim(V)
Else
    sqlYBIATAB0_Read = "? rsYBIATAB0_Read : " & lId & "_" & lK1 & "_" & lK2
End If
Exit Function

Error_Handler:
'-------------
    sqlYBIATAB0_Read = " rsYBIATAB0_Read : " & Error
End Function

Public Sub rsYBIATAB0_cboK2(lId As String, lK1 As String, cbo As Control)  'ComboBox)
Dim X As String, K1 As Integer
Dim blnSAB As Boolean
K1 = 0
blnSAB = False
Select Case lK1
    Case "CLIENAPAY": blnSAB = True: K1 = 16
    Case "CLIENACAT", "CLIENAETA": blnSAB = True: K1 = 13
    Case "PLANCOPRO":  K1 = 12
End Select

cbo.Clear
X = "select * from YBIATAB0 " _
    & " where BIATABID = '" & lId & "'" _
    & " and BIATABK1 = '" & lK1 & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    If blnSAB Then
        cbo.AddItem Trim(Mid$(rsMDB("BIATABK2"), 4, 9)) & "    " & Trim(Mid$(rsMDB("BIATABTXT"), K1, 24))
    Else
        If K1 = 0 Then
            cbo.AddItem rsMDB("BIATABK2")
        Else
            cbo.AddItem Trim(rsMDB("BIATABK2")) & " " & Trim(Mid$(rsMDB("BIATABTXT"), K1, 24))
        End If
    End If
    rsMDB.MoveNext
Loop

End Sub
Public Sub sqlYBIATAB0_cboID(lId As String, cbo As ComboBox)
Dim X As String
Dim kLen As Integer

Select Case Trim(lId)
    Case "ROPDOSSTA", "ROPDOSGNAT", "ROPDOSGPRV", "ROPDOSGGRA", "ROPDOSGPRI": kLen = 1
    Case "ROPDOSXAPP", "ROPDOSXDOM": kLen = 12
    Case "ROPINFSTA", "ROPINFGNAT", "ROPINFMAIL": kLen = 1
    Case Else:  kLen = 12
End Select

cbo.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = '" & lId & "'"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    cbo.AddItem Mid$(rsSab("BIATABK1"), 1, kLen) & " - " & Trim(Mid$(rsSab("BIATABTXT"), 1, 24))
    rsSab.MoveNext
Loop

If cbo.ListCount > 0 Then cbo.ListIndex = 0
End Sub

Public Sub sqlYBIATAB0_cboID_K1(lId As String, lK1 As String, cbo As ComboBox)
Dim X As String
Dim kLen As Integer

Select Case Trim(lId)
    Case Else:  kLen = 12
End Select

cbo.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = '" & lId & "'" _
    & " and BIATABK1 = '" & lK1 & "'"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    cbo.AddItem Mid$(rsSab("BIATABK2"), 1, kLen) & " - " & Trim(Mid$(rsSab("BIATABTXT"), 1, 24))
    rsSab.MoveNext
Loop

If cbo.ListCount > 0 Then cbo.ListIndex = 0
End Sub

'---------------------------------------------------------
Public Sub rsYBIATAB0_Init(rsYBIATAB0 As typeYBIATAB0)
'---------------------------------------------------------
rsYBIATAB0.BIATABID = ""
rsYBIATAB0.BIATABK1 = ""
rsYBIATAB0.BIATABK2 = ""
rsYBIATAB0.BIATABTXT = ""

End Sub


'








