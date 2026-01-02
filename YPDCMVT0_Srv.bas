Attribute VB_Name = "srvYPDCMVT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public paramPDCMVT_Path As String

Dim rsAdo As ADODB.Recordset
 
Type typeYPDCMVT0
 
      PDCMVTDTR     As String * 8   'date comptable
      PDCMVTPIE     As Long         'N° pièce
      PDCMVTECR     As Long         'N° écriture
      PDCMVTOPEC    As String * 3   'code opération
      PDCMVTOPEN    As Long         'n° opération
      PDCMVTCPT     As String * 20  'compte PDC
      PDCMVTDEV     As String * 3   'devise
      PDCMVTMTD     As Currency     'montant en devise
      PDCMVTMTE     As Currency     'montant en euro
      PDCMVTTAUX    As Double       'taux
      PDCMVTDVA     As String * 8   'date valeur
      PDCMVTCLI     As String * 7   'client
      PDCMVTSTA     As String * 1   'statut
      PDCMVTSTA2    As String * 1   'statut
      PDCMVTKCUT    As String * 1   'position non coupée
      PDCMVTSER     As String * 2   'service
      PDCMVTSSE     As String * 2   'sous-service
      PDCMVTUPDS    As Long         'séq màj
End Type
Public xYPDCMVT0 As typeYPDCMVT0
Public Function sqlYPDCMVT0_DeleteW(lWhere As String, Nb As Long)
Dim X As String, xSQL As String

On Error GoTo Error_Handler
sqlYPDCMVT0_DeleteW = Null
    
xSQL = "delete from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 " & lWhere
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
Exit Function
Error_Handler:
    sqlYPDCMVT0_DeleteW = Error
End Function

Public Function sqlYPDCMVT0_Insert(newY As typeYPDCMVT0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYPDCMVT0_Insert = Null
xSet = " (PDCMVTDTR"
xValues = " values('" & newY.PDCMVTDTR & "'"


' Détecter les modifications
'===================================================================================
If newY.PDCMVTOPEN <> 0 Then xSet = xSet & ",PDCMVTOPEN": xValues = xValues & " ," & newY.PDCMVTOPEN
If newY.PDCMVTPIE <> 0 Then xSet = xSet & ",PDCMVTPIE": xValues = xValues & " ," & newY.PDCMVTPIE
If newY.PDCMVTECR <> 0 Then xSet = xSet & ",PDCMVTECR": xValues = xValues & " ," & newY.PDCMVTECR
If newY.PDCMVTMTD <> 0 Then xSet = xSet & ",PDCMVTMTD": xValues = xValues & " ," & Replace(newY.PDCMVTMTD, ",", ".")
If newY.PDCMVTMTE <> 0 Then xSet = xSet & ",PDCMVTMTE": xValues = xValues & " ," & Replace(newY.PDCMVTMTE, ",", ".")
If newY.PDCMVTTAUX <> 0 Then xSet = xSet & ",PDCMVTTAUX": xValues = xValues & " ," & Replace(newY.PDCMVTTAUX, ",", ".")

If Trim(newY.PDCMVTOPEC) <> "" Then xSet = xSet & ",PDCMVTOPEC": xValues = xValues & " ,'" & newY.PDCMVTOPEC & "'"
If Trim(newY.PDCMVTCPT) <> "" Then xSet = xSet & ",PDCMVTCPT": xValues = xValues & " ,'" & newY.PDCMVTCPT & "'"
If Trim(newY.PDCMVTDEV) <> "" Then xSet = xSet & ",PDCMVTDEV": xValues = xValues & " ,'" & newY.PDCMVTDEV & "'"
If Trim(newY.PDCMVTDVA) <> "" Then xSet = xSet & ",PDCMVTDVA": xValues = xValues & " ,'" & newY.PDCMVTDVA & "'"
If Trim(newY.PDCMVTCLI) <> "" Then xSet = xSet & ",PDCMVTCLI": xValues = xValues & " ,'" & newY.PDCMVTCLI & "'"
If Trim(newY.PDCMVTSTA) <> "" Then xSet = xSet & ",PDCMVTSTA": xValues = xValues & " ,'" & newY.PDCMVTSTA & "'"
If Trim(newY.PDCMVTSTA2) <> "" Then xSet = xSet & ",PDCMVTSTA2": xValues = xValues & " ,'" & newY.PDCMVTSTA2 & "'"
If Trim(newY.PDCMVTKCUT) <> "" Then xSet = xSet & ",PDCMVTKCUT": xValues = xValues & " ,'" & newY.PDCMVTKCUT & "'"
If Trim(newY.PDCMVTSER) <> "" Then xSet = xSet & ",PDCMVTSER": xValues = xValues & " ,'" & newY.PDCMVTSER & "'"
If Trim(newY.PDCMVTSSE) <> "" Then xSet = xSet & ",PDCMVTSSE": xValues = xValues & " ,'" & newY.PDCMVTSSE & "'"

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0" & xSet & ")" & xValues & ")"
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYPDCMVT0_Insert = "Erreur màj : " & newY.PDCMVTDTR & " " & newY.PDCMVTPIE & " " & newY.PDCMVTECR
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYPDCMVT0_Insert = Error
End Function

Public Function sqlYPDCMVT0_Update(newY As typeYPDCMVT0, oldY As typeYPDCMVT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYPDCMVT0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.PDCMVTDTR <> newY.PDCMVTDTR _
Or oldY.PDCMVTPIE <> newY.PDCMVTPIE _
Or oldY.PDCMVTECR <> newY.PDCMVTECR Then
    sqlYPDCMVT0_Update = "Erreur Clé: " & newY.PDCMVTDTR & "." & oldY.PDCMVTPIE & "." & oldY.PDCMVTECR
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where PDCMVTDTR = '" & oldY.PDCMVTDTR & "'" _
       & " and PDCMVTPIE = " & oldY.PDCMVTPIE _
       & " and PDCMVTECR = " & oldY.PDCMVTECR _
       & " and PDCMVTUPDS = " & oldY.PDCMVTUPDS

newY.PDCMVTUPDS = newY.PDCMVTUPDS + 1
xSet = xSet & " set PDCMVTUPDS = " & newY.PDCMVTUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.PDCMVTOPEN <> oldY.PDCMVTOPEN Then blnUpdate = True:  xSet = xSet & " , PDCMVTOPEN = " & cur_P(newY.PDCMVTOPEN)
If newY.PDCMVTMTD <> oldY.PDCMVTMTD Then blnUpdate = True:  xSet = xSet & " , PDCMVTMTD = " & cur_P(newY.PDCMVTMTD)
If newY.PDCMVTMTE <> oldY.PDCMVTMTE Then blnUpdate = True:  xSet = xSet & " , PDCMVTMTE = " & cur_P(newY.PDCMVTMTE)
If newY.PDCMVTTAUX <> oldY.PDCMVTTAUX Then blnUpdate = True:  xSet = xSet & " , PDCMVTTAUX = " & cur_P(newY.PDCMVTTAUX)

If newY.PDCMVTOPEC <> oldY.PDCMVTOPEC Then blnUpdate = True:  xSet = xSet & " , PDCMVTOPEC = '" & newY.PDCMVTOPEC & "'"
If newY.PDCMVTCPT <> oldY.PDCMVTCPT Then blnUpdate = True:  xSet = xSet & " , PDCMVTCPT = '" & newY.PDCMVTCPT & "'"
If newY.PDCMVTDEV <> oldY.PDCMVTDEV Then blnUpdate = True:  xSet = xSet & " , PDCMVTDEV = '" & newY.PDCMVTDEV & "'"
If newY.PDCMVTDVA <> oldY.PDCMVTDVA Then blnUpdate = True:  xSet = xSet & " , PDCMVTDVA = '" & newY.PDCMVTDVA & "'"
If newY.PDCMVTCLI <> oldY.PDCMVTCLI Then blnUpdate = True:  xSet = xSet & " , PDCMVTCLI = '" & newY.PDCMVTCLI & "'"
If newY.PDCMVTSTA <> oldY.PDCMVTSTA Then blnUpdate = True:  xSet = xSet & " , PDCMVTSTA = '" & newY.PDCMVTSTA & "'"
If newY.PDCMVTSTA2 <> oldY.PDCMVTSTA2 Then blnUpdate = True:  xSet = xSet & " , PDCMVTSTA2 = '" & newY.PDCMVTSTA2 & "'"
If newY.PDCMVTKCUT <> oldY.PDCMVTKCUT Then blnUpdate = True:  xSet = xSet & " , PDCMVTKCUT = '" & newY.PDCMVTKCUT & "'"
If newY.PDCMVTSER <> oldY.PDCMVTSER Then blnUpdate = True:  xSet = xSet & " , PDCMVTSER = '" & newY.PDCMVTSER & "'"
If newY.PDCMVTSSE <> oldY.PDCMVTSSE Then blnUpdate = True:  xSet = xSet & " , PDCMVTSSE = '" & newY.PDCMVTSSE & "'"


If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYPDCMVT0_Update = "Erreur màj : " & newY.PDCMVTDEV
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYPDCMVT0_Update = Error
End Function





Public Function rsYPDCMVT0_GetBuffer(rsAdo As ADODB.Recordset, lYPDCMVT0 As typeYPDCMVT0)
On Error GoTo Error_Handler
rsYPDCMVT0_GetBuffer = Null

lYPDCMVT0.PDCMVTDTR = rsAdo("PDCMVTDTR")
lYPDCMVT0.PDCMVTPIE = rsAdo("PDCMVTPIE")
lYPDCMVT0.PDCMVTECR = rsAdo("PDCMVTECR")
lYPDCMVT0.PDCMVTOPEC = rsAdo("PDCMVTOPEC")
lYPDCMVT0.PDCMVTOPEN = rsAdo("PDCMVTOPEN")
lYPDCMVT0.PDCMVTCPT = rsAdo("PDCMVTCPT")
lYPDCMVT0.PDCMVTDEV = rsAdo("PDCMVTDEV")
lYPDCMVT0.PDCMVTMTD = rsAdo("PDCMVTMTD")
lYPDCMVT0.PDCMVTMTE = rsAdo("PDCMVTMTE")
lYPDCMVT0.PDCMVTTAUX = rsAdo("PDCMVTTAUX")
lYPDCMVT0.PDCMVTDVA = rsAdo("PDCMVTDVA")
lYPDCMVT0.PDCMVTCLI = rsAdo("PDCMVTCLI")
lYPDCMVT0.PDCMVTSTA = rsAdo("PDCMVTSTA")
lYPDCMVT0.PDCMVTSTA2 = rsAdo("PDCMVTSTA2")
lYPDCMVT0.PDCMVTKCUT = rsAdo("PDCMVTKCUT")
lYPDCMVT0.PDCMVTSER = rsAdo("PDCMVTSER")
lYPDCMVT0.PDCMVTSSE = rsAdo("PDCMVTSSE")
lYPDCMVT0.PDCMVTUPDS = rsAdo("PDCMVTUPDS")

Exit Function
Error_Handler:
rsYPDCMVT0_GetBuffer = Error


End Function

Public Function rsYPDCMVT0_Init(lYPDCMVT0 As typeYPDCMVT0)

lYPDCMVT0.PDCMVTDTR = ""
lYPDCMVT0.PDCMVTPIE = 0
lYPDCMVT0.PDCMVTECR = 0
lYPDCMVT0.PDCMVTOPEC = ""
lYPDCMVT0.PDCMVTOPEN = 0
lYPDCMVT0.PDCMVTCPT = ""
lYPDCMVT0.PDCMVTDEV = ""
lYPDCMVT0.PDCMVTMTD = 0
lYPDCMVT0.PDCMVTMTE = 0
lYPDCMVT0.PDCMVTTAUX = 0
lYPDCMVT0.PDCMVTDVA = ""
lYPDCMVT0.PDCMVTCLI = ""
lYPDCMVT0.PDCMVTSTA = ""
lYPDCMVT0.PDCMVTSTA2 = ""
lYPDCMVT0.PDCMVTKCUT = ""
lYPDCMVT0.PDCMVTSER = ""
lYPDCMVT0.PDCMVTSSE = ""
lYPDCMVT0.PDCMVTUPDS = 0

End Function




