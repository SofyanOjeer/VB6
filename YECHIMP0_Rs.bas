Attribute VB_Name = "rsYECHIMP0"
'---------------------------------------------------------
Option Explicit
Type typeYECHIMP0

    ECHIMPJOB      As Long      '                 ')
    ECHIMPJOBS     As Long       '                 ')
    ECHIMPSEQ      As Long       '                 ')
    ECHIMPCPT      As String * 20         'COMPTE           ')
    ECHIMPDEV      As String * 3         'DEVISE           ')
    ECHIMPDTRT     As Long       '                 ')
    ECHIMPDOPE     As Long       '                 ')
    ECHIMPDDEB     As Long       '                 ')
    ECHIMPDFIN     As Long       '                 ')
    ECHIMPIDEM     As Currency       'INT DEB MONTANT  ')
    ECHIMPIDES     As String * 1         'INT DEB SENS     ')
    ECHIMPIDEV     As Long       'INT DEB VALEUR   ')
    ECHIMPIDET     As Double       'INT DEB TAUX     ')
    ECHIMPICRM     As Currency       'INT CRE MONTANT  '
    ECHIMPICRS     As String * 1         'INT CRE SENS     '
    ECHIMPICRV     As Long       'INT CRE VALEUR   '
    ECHIMPICRT     As Double       'INT CRE TAUX     '
    ECHIMPCPFD     As Currency       '                 '
    ECHIMPCMVT     As Currency       '                 '
    ECHIMPCCPT     As Currency       '                 '
    ECHIMPMON      As Currency       'MONTANT TOTAL    '
    ECHIMPMONS     As String * 1         'SENS             '
    ECHIMPNREF     As String * 10         'NOTRE REF        '
    ECHIMPAD1      As String * 32         'ADRESSE          '
    ECHIMPAD2      As String * 32         'ADRESSE          '
    ECHIMPAD3      As String * 32         'ADRESSE          '
    ECHIMPAD4      As String * 32         'ADRESSE          '
    ECHIMPAD5      As String * 32         'ADRESSE          '
    ECHIMPAD6      As String * 32         'ADRESSE          '

End Type
'---------------------------------------------------------
Public Function rsYCHQMON0_GetBuffer(rsSab As ADODB.Recordset, rsYCHQMON0 As typeYCHQMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYCHQMON0_GetBuffer = Null

rsYCHQMON0.CHQRC1ETA = rsSab("CHQRC1ETA")
rsYCHQMON0.CHQRC1AGE = rsSab("CHQRC1AGE")
rsYCHQMON0.CHQRC1SER = rsSab("CHQRC1SER")
rsYCHQMON0.CHQRC1SSE = rsSab("CHQRC1SSE")
rsYCHQMON0.CHQRC1OPE = rsSab("CHQRC1OPE")
rsYCHQMON0.CHQRC1DOS = rsSab("CHQRC1DOS")
rsYCHQMON0.CHQRC1DCR = rsSab("CHQRC1DCR")
rsYCHQMON0.CHQDATE = rsSab("CHQDATE")
rsYCHQMON0.CHQCOMPTE = rsSab("CHQCOMPTE")
rsYCHQMON0.CHQCREM = rsSab("CHQCREM")
rsYCHQMON0.CHQDEVISE = rsSab("CHQDEVISE")
rsYCHQMON0.CHQMONTANT = rsSab("CHQMONTANT")
rsYCHQMON0.CHQNB = rsSab("CHQNB")
rsYCHQMON0.CHQMONSTA = rsSab("CHQMONSTA")
rsYCHQMON0.CHQMONUPDS = rsSab("CHQMONUPDS")

Exit Function

Error_Handler:

rsYCHQMON0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsYCHQMON0_Init(rsYCHQMON0 As typeYCHQMON0)
rsYCHQMON0.CHQRC1ETA = 0
rsYCHQMON0.CHQRC1AGE = 0
rsYCHQMON0.CHQRC1SER = ""
rsYCHQMON0.CHQRC1SSE = ""
rsYCHQMON0.CHQRC1OPE = ""
rsYCHQMON0.CHQRC1DOS = 0
rsYCHQMON0.CHQRC1DCR = 0
rsYCHQMON0.CHQDATE = 0
rsYCHQMON0.CHQCOMPTE = ""
rsYCHQMON0.CHQCREM = ""
rsYCHQMON0.CHQDEVISE = ""
rsYCHQMON0.CHQMONTANT = 0
rsYCHQMON0.CHQNB = 0
rsYCHQMON0.CHQMONSTA = ""
rsYCHQMON0.CHQMONUPDS = 0
'---------------------------------------------------------

End Sub

Public Function sqlYUPDLOG0_Insert(newY As typeYUPDLOG0)
Dim X As String, xSql As String, Nb As Long
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

xSql = "Insert into " & paramIBM_Library_SABSPE & ".YUPDLOG0" & xSet & ")" & xValues & ")"

Set rsADO = cnSab_Update.Execute(xSql, Nb)

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
Dim X As String, xSql As String, Nb As Long
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
    
    xSql = "update " & paramIBM_Library_SABSPE & ".YUPDLOG0" & xSet & xWhere
    
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    
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



