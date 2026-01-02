Attribute VB_Name = "srvYICCCPT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYICCCPT0

    ICCCPTETA   As Integer       ' établissement
    ICCCPTAGE   As Integer       ' agence
    ICCCPTCOM   As String * 20   ' compte
    ICCCPTDEV   As String * 3  ' devise
    ICCCPTGRP   As Integer       ' groupe
    ICCCPTSTA   As String * 1  ' statut
    ICCCPTUUSR   As String * 10 ' utilisateur maj
    ICCCPTUAMJ   As Long        ' DATE maj
    ICCCPTUHMS   As Long        ' heure maj
    ICCCPTUSEQ   As Long        ' N° séquence (info)
    
End Type
Public Function sqlYICCCPT0_Insert(newY As typeYICCCPT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYICCCPT0_Insert = Null

xSet = " (ICCCPTUSEQ"
xValues = " values(" & newY.ICCCPTUSEQ

' Insertion :
'===================================================================================
If newY.ICCCPTUAMJ <> 0 Then xSet = xSet & ",ICCCPTUAMJ": xValues = xValues & ", " & newY.ICCCPTUAMJ
If newY.ICCCPTUHMS <> 0 Then xSet = xSet & ",ICCCPTUHMS": xValues = xValues & ", " & newY.ICCCPTUHMS
If newY.ICCCPTETA <> 0 Then xSet = xSet & ",ICCCPTETA": xValues = xValues & ", " & newY.ICCCPTETA
If newY.ICCCPTAGE <> 0 Then xSet = xSet & ",ICCCPTAGE": xValues = xValues & ", " & newY.ICCCPTAGE
If newY.ICCCPTGRP <> 0 Then xSet = xSet & ",ICCCPTGRP": xValues = xValues & ", " & newY.ICCCPTGRP

'===================================================================================

If Trim(newY.ICCCPTCOM) <> "" Then xSet = xSet & ",ICCCPTCOM": xValues = xValues & ", '" & Replace(Trim(newY.ICCCPTCOM), "'", "''") & "'"
If Trim(newY.ICCCPTSTA) <> "" Then xSet = xSet & ",ICCCPTSTA": xValues = xValues & ", '" & Replace(Trim(newY.ICCCPTSTA), "'", "''") & "'"
If Trim(newY.ICCCPTDEV) <> "" Then xSet = xSet & ",ICCCPTDEV": xValues = xValues & ", '" & Replace(Trim(newY.ICCCPTDEV), "'", "''") & "'"
If newY.ICCCPTUUSR <> "" Then xSet = xSet & ",ICCCPTUUSR": xValues = xValues & ", '" & Replace(newY.ICCCPTUUSR, "'", "''") & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YICCCPT0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYICCCPT0_Insert = "Erreur màj : " & newY.ICCCPTDEV & " " & newY.ICCCPTUSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYICCCPT0_Insert = Error
End Function
Public Function sqlYICCCPT0_Update(newY As typeYICCCPT0, oldY As typeYICCCPT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
Dim K As Integer

On Error GoTo Error_Handler
sqlYICCCPT0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.ICCCPTCOM <> newY.ICCCPTCOM _
Or oldY.ICCCPTETA <> newY.ICCCPTETA _
Or oldY.ICCCPTAGE <> newY.ICCCPTAGE _
Or oldY.ICCCPTUHMS <> newY.ICCCPTUHMS _
Or oldY.ICCCPTUUSR <> newY.ICCCPTUUSR _
Or oldY.ICCCPTUSEQ <> newY.ICCCPTUSEQ Then
    sqlYICCCPT0_Update = "Erreur ICCCPTUAMJ : " & newY.ICCCPTUAMJ & "." & oldY.ICCCPTUSEQ
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where ICCCPTETA = " & oldY.ICCCPTETA _
       & " and ICCCPTAGE = " & oldY.ICCCPTAGE _
       & " and ICCCPTCOM = '" & oldY.ICCCPTCOM & "'" _
       & " and ICCCPTUAMJ = " & oldY.ICCCPTUAMJ _
       & " and ICCCPTUHMS = " & oldY.ICCCPTUHMS _
       & " and ICCCPTUUSR = '" & oldY.ICCCPTUUSR & "'" _
       & " and ICCCPTUSEQ = " & oldY.ICCCPTUSEQ

xSet = " set "
blnUpdate = False


' Détecter les modifications
'===================================================================================
If newY.ICCCPTSTA <> oldY.ICCCPTSTA Then blnUpdate = True:  xSet = xSet & " , ICCCPTSTA = '" & newY.ICCCPTSTA & "'"
If newY.ICCCPTGRP <> oldY.ICCCPTGRP Then blnUpdate = True:  xSet = xSet & " , ICCCPTGRP = " & newY.ICCCPTGRP

If blnUpdate Then
    K = InStr(xSet, ",")
    If K > 0 Then Mid$(xSet, K, 1) = " "
    xSQL = "update " & paramIBM_Library_SABSPE & ".YICCCPT0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYICCCPT0_Update = "Erreur màj : " & newY.ICCCPTUAMJ
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYICCCPT0_Update = Error
End Function


'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYICCCPT0_GetBuffer(rsAdo As ADODB.Recordset, rsYICCCPT0 As typeYICCCPT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYICCCPT0_GetBuffer = Null

rsYICCCPT0.ICCCPTETA = rsAdo("ICCCPTETA")
rsYICCCPT0.ICCCPTAGE = rsAdo("ICCCPTAGE")
rsYICCCPT0.ICCCPTCOM = rsAdo("ICCCPTCOM")
rsYICCCPT0.ICCCPTDEV = rsAdo("ICCCPTDEV")
rsYICCCPT0.ICCCPTGRP = rsAdo("ICCCPTGRP")
rsYICCCPT0.ICCCPTSTA = rsAdo("ICCCPTSTA")

rsYICCCPT0.ICCCPTUUSR = rsAdo("ICCCPTUUSR")
rsYICCCPT0.ICCCPTUAMJ = rsAdo("ICCCPTUAMJ")
rsYICCCPT0.ICCCPTUHMS = rsAdo("ICCCPTUHMS")
rsYICCCPT0.ICCCPTUSEQ = rsAdo("ICCCPTUSEQ")

Exit Function

Error_Handler:

rsYICCCPT0_GetBuffer = Error

End Function









Public Sub rsYICCCPT0_Init(lYICCCPT0 As typeYICCCPT0)
lYICCCPT0.ICCCPTETA = 0                  ' chèque
lYICCCPT0.ICCCPTAGE = 0
lYICCCPT0.ICCCPTCOM = ""
lYICCCPT0.ICCCPTDEV = ""
lYICCCPT0.ICCCPTSTA = ""                   ' statut
lYICCCPT0.ICCCPTGRP = ""

lYICCCPT0.ICCCPTUUSR = usrName_UCase       ' utilisateur maj
lYICCCPT0.ICCCPTUAMJ = DSys                ' DATE maj
lYICCCPT0.ICCCPTUHMS = time_Hms            ' heure maj
lYICCCPT0.ICCCPTUSEQ = 0

End Sub


