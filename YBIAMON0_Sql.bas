Attribute VB_Name = "sqlYBIAMON0"
Option Explicit

'______________________________________________________________
Public blnSAB_TAUX_Auto As Boolean
Public blnALERTE_TAU As Boolean, blnALERTE_FIX As Boolean
Public blnUPDATE_TAU As Boolean, blnUPDATE_FIX As Boolean
Public blnCONTROL_TAU As Boolean, blnCONTROL_FIX As Boolean
Public blnEURJ1M As Boolean, blnEURUSD As Boolean
Public sabEURJ1M As Double, sabEURUSD As Double
Public newEURJ1M As Double, newEURUSD As Double

Public Function fctExploitation_Auto_Control(lYBIAMON0 As typeYBIAMON0)
Dim oldYBIAMON0 As typeYBIAMON0
Dim V
On Error GoTo Error_Handler

fctExploitation_Auto_Control = Null
oldYBIAMON0 = lYBIAMON0

V = rsYBIAMON0_Read(lYBIAMON0)
If Not IsNull(V) Then GoTo Error_MsgBox
If lYBIAMON0.MONSTATUS <> oldYBIAMON0.MONSTATUS Then
    V = "Action précédente en cours : " & lYBIAMON0.MONAPP & ">" & lYBIAMON0.MONFLUX & " > " & lYBIAMON0.MONSTATUS
    GoTo Error_MsgBox
End If
If Trim(lYBIAMON0.MONFILE) >= YBIATAB0_DATE_CPT_J Then
    V = "Déjà traité à cette date : " & lYBIAMON0.MONAPP & ">" & lYBIAMON0.MONFLUX & " > " & lYBIAMON0.MONFILE
    GoTo Error_MsgBox
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'--------------------------------------------------------------
Call ECRIT_LOG2008("Avant V = cnSAB_Transaction(""BeginTrans"") dans fctExploitation_Auto_Control")
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
oldYBIAMON0 = lYBIAMON0
lYBIAMON0.MONSTATUS = "MONITOR"
lYBIAMON0.MONNUM = lYBIAMON0.MONNUM + 1
V = sqlYBIAMON0_Update(lYBIAMON0, oldYBIAMON0, True)
If Not IsNull(V) Then GoTo Error_MsgBox
'------------------------------------------------------------------------------------

Exit Function

Error_Handler:
    V = Error & " " & lYBIAMON0.MONFLUX & " --> " & lYBIAMON0.MONSTATUS
Error_MsgBox:
    fctExploitation_Auto_Control = V
    If blnAuto_Form_Show Then MsgBox V, vbCritical, frmElp_Caption & App_Debug
    
End Function
Public Function fctExploitation_Transaction_Control(lYBIAMON0 As typeYBIAMON0)
Dim oldYBIAMON0 As typeYBIAMON0
Dim V
On Error GoTo Error_Handler

fctExploitation_Transaction_Control = Null
oldYBIAMON0 = lYBIAMON0

V = rsYBIAMON0_Read(lYBIAMON0)
If Not IsNull(V) Then GoTo Error_MsgBox

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'--------------------------------------------------------------
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
oldYBIAMON0 = lYBIAMON0
lYBIAMON0.MONSTATUS = "MONITOR"
lYBIAMON0.MONNUM = lYBIAMON0.MONNUM + 1
V = sqlYBIAMON0_Update(lYBIAMON0, oldYBIAMON0, True)
If Not IsNull(V) Then GoTo Error_MsgBox
'------------------------------------------------------------------------------------

Exit Function

Error_Handler:
    V = Error & " " & lYBIAMON0.MONFLUX & " --> " & lYBIAMON0.MONSTATUS
Error_MsgBox:
    fctExploitation_Transaction_Control = V
    If blnAuto_Form_Show Then MsgBox V, vbCritical, frmElp_Caption & App_Debug
    
End Function


Public Function fctExploitation_Transaction_End(newY As typeYBIAMON0, oldY As typeYBIAMON0)
Dim newYBIAMON0 As typeYBIAMON0
On Error GoTo Error_Handler
Dim V
fctExploitation_Transaction_End = Null

'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'------------------------------------------------------------------------------------
V = sqlYBIAMON0_Update(newY, oldY, True)
If Not IsNull(V) Then GoTo Error_MsgBox
'------------------------------------------------------------------------------------

    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        newYBIAMON0.MONSTATUS = "Rollback"
    Else
        V = cnSAB_Transaction("Commit")
        newYBIAMON0.MONSTATUS = ""
    End If

'------------------------------------------------------------------------------------

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    fctExploitation_Transaction_End = V
    If Not blnAuto_Form_Show Then MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Function


Public Function fctExploitation_Auto_End(lYBIAMON0 As typeYBIAMON0)
Dim newYBIAMON0 As typeYBIAMON0
On Error GoTo Error_Handler
Dim V
fctExploitation_Auto_End = Null

'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'------------------------------------------------------------------------------------
newYBIAMON0 = lYBIAMON0
newYBIAMON0.MONSTATUS = ""
newYBIAMON0.MONFILE = YBIATAB0_DATE_CPT_J
V = sqlYBIAMON0_Update(newYBIAMON0, lYBIAMON0, True)
If Not IsNull(V) Then GoTo Error_MsgBox
'------------------------------------------------------------------------------------

    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        newYBIAMON0.MONSTATUS = "Rollback"
    Else
        V = cnSAB_Transaction("Commit")
        newYBIAMON0.MONSTATUS = ""
    End If

'------------------------------------------------------------------------------------

Exit Function

Error_Handler:
    V = Error & " " & lYBIAMON0.MONFLUX & " --> " & lYBIAMON0.MONSTATUS
Error_MsgBox:
    fctExploitation_Auto_End = V
    If Not blnAuto_Form_Show Then MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Function

Public Function sqlYBIAMON0_Insert(newY As typeYBIAMON0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYBIAMON0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

xSet = xSet & ",MONAPP": xValues = xValues & " ,'" & Trim(newY.MONAPP) & "'"
xSet = xSet & ",MONFLUX": xValues = xValues & " ,'" & newY.MONFLUX & "'"
xSet = xSet & ",MONSTATUS": xValues = xValues & " ,'" & newY.MONSTATUS & "'"
xSet = xSet & ",MONNUM": xValues = xValues & " ," & newY.MONNUM
xSet = xSet & ",MONJOB": xValues = xValues & " ,'" & newY.MONJOB & "'"
xSet = xSet & ",MONPGM": xValues = xValues & " ,'" & newY.MONPGM & "'"
xSet = xSet & ",MONUSR": xValues = xValues & " ,'" & newY.MONUSR & "'"
xSet = xSet & ",MONAMJ": xValues = xValues & " ," & newY.MONAMJ
xSet = xSet & ",MONHMS": xValues = xValues & " ," & newY.MONHMS
xSet = xSet & ",MONFILE": xValues = xValues & " ,'" & newY.MONFILE & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "
Call FEU_ROUGE
xSql = "Insert into " & paramIBM_Library_SABSPE & ".YBIAMON7" & xSet & ")" & xValues & ")"

Set rsSab_Update = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBIAMON0_Insert = "Erreur màj : " & newY.MONPGM
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBIAMON0_Insert = Error
End Function
Public Function sqlYBIAMON0_Update(newY As typeYBIAMON0, oldY As typeYBIAMON0, blnStamp As Boolean)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlYBIAMON0_Update = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where MONAPP = '" & oldY.MONAPP & "'" _
         & " and MONFLUX = '" & oldY.MONFLUX & "'" _
         & " and MONUPDS = " & oldY.MONUPDS
         
newY.MONUPDS = newY.MONUPDS + 1
xSet = xSet & " set MONUPDS = " & newY.MONUPDS

If blnStamp Then
    newY.MONUSR = usrName_UCase
    newY.MONAMJ = DSys
    newY.MONHMS = time_Hms
End If


' Détecter les modifications
'===================================================================================
If newY.MONAPP <> oldY.MONAPP Then xSet = xSet & " , MONAPP = '" & Trim(newY.MONAPP) & "'"
If newY.MONFLUX <> oldY.MONFLUX Then xSet = xSet & " , MONFLUX = '" & newY.MONFLUX & "'"
If newY.MONSTATUS <> oldY.MONSTATUS Then xSet = xSet & " , MONSTATUS = '" & Trim(newY.MONSTATUS) & "'"
If newY.MONNUM <> oldY.MONNUM Then xSet = xSet & " , MONNUM = " & newY.MONNUM
If newY.MONJOB <> oldY.MONJOB Then xSet = xSet & " , MONJOB = '" & newY.MONJOB & "'"
If newY.MONPGM <> oldY.MONPGM Then xSet = xSet & " , MONPGM = '" & newY.MONPGM & "'"
If newY.MONUSR <> oldY.MONUSR Then xSet = xSet & " , MONUSR = '" & newY.MONUSR & "'"
If newY.MONAMJ <> oldY.MONAMJ Then xSet = xSet & " , MONAMJ = " & newY.MONAMJ
If newY.MONHMS <> oldY.MONHMS Then xSet = xSet & " , MONHMS = " & newY.MONHMS
If newY.MONFILE <> oldY.MONFILE Then xSet = xSet & " , MONFILE = '" & newY.MONFILE & "'"


'''Mid$(xSet, 6, 1) = " "
xSql = "update " & paramIBM_Library_SABSPE & ".YBIAMON7" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBIAMON0_Update = "Erreur màj : " & newY.MONPGM
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBIAMON0_Update = Error
End Function



