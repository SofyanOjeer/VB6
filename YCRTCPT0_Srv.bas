Attribute VB_Name = "srvYCRTCPT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYCRTCPT0
 
      CRTCPTCPT   As String       'devise
      CRTCPTRUB   As String        'rubriqueCRT
      CRTCPTSTA   As String       'statut
      

End Type
Public xYCRTCPT0 As typeYCRTCPT0

Type typeYCRTRUB0

    Code        As String
    PCEC        As String
    PCEC_Len    As String
    Lib         As String

End Type
Public arrCRT_PCEC_Rub() As typeYCRTRUB0, arrCRT_PCEC_Rub_Nb As Integer, arrCRT_PCEC_Rub_Nb0 As Integer
Public arrCRT_Rub() As typeYCRTRUB0, arrCRT_Rub_Nb As Integer, arrCRT_Rub_K As Integer

'---------------------------------------------------------
Public Sub arrYCPTRUB0_Load()
'---------------------------------------------------------
Dim V, xSQL As String, X As String
On Error GoTo Error_Handler


xSQL = "select count(Distinct BIATABK1) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Rubrique'"
Set rsSab = cnsab.Execute(xSQL)

ReDim arrCRT_Rub(rsSab(0) + 2)

arrCRT_PCEC_Rub_Nb = 0
arrCRT_Rub_Nb = 0: arrCRT_Rub_K = 0
arrCRT_Rub(0).Code = "": arrCRT_Rub(0).Lib = ""
X = "?$"

xSQL = "select distinct BIATABK1 , BIATABTXT from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Rubrique'" _
     & " order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    If X <> Trim(rsSab("BIATABK1")) Then
        X = Trim(rsSab("BIATABK1"))
        arrCRT_Rub_Nb = arrCRT_Rub_Nb + 1
        arrCRT_Rub(arrCRT_Rub_Nb).Code = X
        arrCRT_Rub(arrCRT_Rub_Nb).Lib = Trim(rsSab("BIATABTXT"))
    End If
    rsSab.MoveNext
Loop
'_______________________________________________________________________________________________________

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Rubrique'"
Set rsSab = cnsab.Execute(xSQL)
ReDim arrCRT_PCEC_Rub(rsSab(0) + 1)



xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Rubrique'" _
     & " order by BIATABK2 , BIATABK1"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrCRT_PCEC_Rub_Nb = arrCRT_PCEC_Rub_Nb + 1
    X = Trim(rsSab("BIATABK1"))
    arrCRT_PCEC_Rub(arrCRT_PCEC_Rub_Nb).Code = X
    arrCRT_PCEC_Rub(arrCRT_PCEC_Rub_Nb).PCEC = Trim(rsSab("BIATABK2"))
    arrCRT_PCEC_Rub(arrCRT_PCEC_Rub_Nb).PCEC_Len = Len(arrCRT_PCEC_Rub(arrCRT_PCEC_Rub_Nb).PCEC)
    
    If arrCRT_PCEC_Rub(arrCRT_PCEC_Rub_Nb).PCEC_Len > 0 Then
        If arrCRT_PCEC_Rub_Nb0 = 0 Then arrCRT_PCEC_Rub_Nb0 = arrCRT_PCEC_Rub_Nb
    End If
    arrCRT_PCEC_Rub(arrCRT_PCEC_Rub_Nb).Lib = Trim(rsSab("BIATABTXT"))
      
    rsSab.MoveNext
Loop

'_______________________________________________________________________________________________________


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug


End Sub


Public Function sqlYCRTCPT0_Update(newY As typeYCRTCPT0, oldY As typeYCRTCPT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYCRTCPT0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.CRTCPTCPT <> newY.CRTCPTCPT Then
    sqlYCRTCPT0_Update = "Erreur CRTCPTPIE : " & newY.CRTCPTCPT
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where CRTCPTCPT = '" & oldY.CRTCPTCPT & "'"
xSet = xSet & " set CRTCPTSTA = '" & newY.CRTCPTSTA & "'"
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.CRTCPTRUB <> oldY.CRTCPTRUB Then blnUpdate = True:  xSet = xSet & " , CRTCPTRUB = '" & newY.CRTCPTRUB & "'"
If newY.CRTCPTSTA <> oldY.CRTCPTSTA Then blnUpdate = True


If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YCRTCPT0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYCRTCPT0_Update = "Erreur màj : " & newY.CRTCPTCPT
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYCRTCPT0_Update = Error
End Function
Public Function sqlYCRTCPT0_Insert(newY As typeYCRTCPT0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCRTCPT0_Insert = Null
xSet = " (CRTCPTCPT"
xValues = " values('" & newY.CRTCPTCPT & "'"

If Trim(newY.CRTCPTRUB) <> "" Then xSet = xSet & ",CRTCPTRUB": xValues = xValues & " ,'" & newY.CRTCPTRUB & "'"
If Trim(newY.CRTCPTSTA) <> "" Then xSet = xSet & ",CRTCPTSTA": xValues = xValues & " ,'" & newY.CRTCPTSTA & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YCRTCPT0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCRTCPT0_Insert = "Erreur màj : " & newY.CRTCPTCPT
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCRTCPT0_Insert = Error
End Function


Public Function rsYCRTCPT0_GetBuffer(rsAdo As ADODB.Recordset, lYCRTCPT0 As typeYCRTCPT0)
On Error GoTo Error_Handler
rsYCRTCPT0_GetBuffer = Null

lYCRTCPT0.CRTCPTCPT = rsAdo("CRTCPTCPT")
lYCRTCPT0.CRTCPTRUB = rsAdo("CRTCPTRUB")
lYCRTCPT0.CRTCPTSTA = rsAdo("CRTCPTSTA")


Exit Function
Error_Handler:
rsYCRTCPT0_GetBuffer = Error


End Function

Public Function rsYCRTCPT0_Init(lYCRTCPT0 As typeYCRTCPT0)

lYCRTCPT0.CRTCPTCPT = ""      ' 3   'devise
lYCRTCPT0.CRTCPTRUB = ""
lYCRTCPT0.CRTCPTSTA = ""      ' 1   'statut



End Function




