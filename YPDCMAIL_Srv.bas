Attribute VB_Name = "srvYPDCMAIL"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public paramPDCMAIL_Path As String

Dim rsAdo As ADODB.Recordset
 
Type typeYPDCMAIL
 
      PDCMAILDTR     As String * 8   'date comptable
      PDCMAILSEQ     As Integer      'séquence
      PDCMAILTXT     As String  'texte
End Type
Public xYPDCMAIL As typeYPDCMAIL
Public Function sqlYPDCMAIL_DeleteW(lWhere As String)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYPDCMAIL_DeleteW = Null
    
xSQL = "delete from " & paramIBM_Library_SABSPE_XXX & ".YPDCMAIL " & lWhere
Call FEU_ROUGE
Set rsAdo = cnsab.Execute(xSQL, Nb)
 Call FEU_VERT
Exit Function
Error_Handler:
    sqlYPDCMAIL_DeleteW = Error
End Function

Public Function sqlYPDCMAIL_Insert(newY As typeYPDCMAIL)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String
Dim K As Long, lenX As Long, xPDCMAILTXT As String
On Error GoTo Error_Handler

sqlYPDCMAIL_Insert = Null

newY.PDCMAILSEQ = 0
xPDCMAILTXT = Replace(Trim(newY.PDCMAILTXT), "|", " ")
xPDCMAILTXT = Replace(xPDCMAILTXT, "'", "|")
lenX = Len(xPDCMAILTXT)
For K = 1 To lenX Step 1000
    If newY.PDCMAILDTR > 20100000 Then newY.PDCMAILSEQ = newY.PDCMAILSEQ + 1
    xSet = " (PDCMAILDTR,PDCMAILSEQ,PDCMAILTXT)"
    xValues = " values('" & newY.PDCMAILDTR & "' ," & newY.PDCMAILSEQ & " ,'" & Mid$(xPDCMAILTXT, K, 1000) & "')"
    Call FEU_ROUGE
    xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YPDCMAIL" & xSet & xValues
    'Set rsAdo = cnSab_Update.Execute(xSql, Nb)
    Set rsAdo = cnsab.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYPDCMAIL_Insert = "Erreur màj : " & newY.PDCMAILDTR
        Exit Function
    End If
Next K

Exit Function
Error_Handler:
    sqlYPDCMAIL_Insert = Error
End Function



Public Function rsYPDCMAIL_GetBuffer(rsAdo As ADODB.Recordset, lYPDCMAIL As typeYPDCMAIL)
On Error GoTo Error_Handler
rsYPDCMAIL_GetBuffer = Null

lYPDCMAIL.PDCMAILDTR = rsAdo("PDCMAILDTR")
lYPDCMAIL.PDCMAILSEQ = rsAdo("PDCMAILSEQ")
lYPDCMAIL.PDCMAILTXT = rsAdo("PDCMAILTXT")

Exit Function
Error_Handler:
rsYPDCMAIL_GetBuffer = Error


End Function

Public Function rsYPDCMAIL_Init(lYPDCMAIL As typeYPDCMAIL)

lYPDCMAIL.PDCMAILDTR = ""
lYPDCMAIL.PDCMAILSEQ = 0
lYPDCMAIL.PDCMAILTXT = ""

End Function







