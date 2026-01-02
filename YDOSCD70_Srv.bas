Attribute VB_Name = "srvYDOSCD70"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYDOSCD70
    DOSCD7DSIT   As Long
    DOSCD7OPE    As String
    DOSCD7NUM    As Long
    DOSCD7KCN    As String
    DOSCD7KNAT   As String
    DOSCD7PCI    As String
    DOSCD7CLI    As String
    DOSCD7MTD    As Currency
    DOSCD7DEV    As String
    DOSCD7STA    As String
    DOSCD7DDEB   As Long
    DOSCD7DFIN   As Long
    DOSCD7DAMJ   As Long
End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYDOSCD70_GetBuffer(rsADO As ADODB.Recordset, rsYDOSCD70 As typeYDOSCD70)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYDOSCD70_GetBuffer = Null
rsYDOSCD70.DOSCD7DSIT = rsADO("DOSCD7DSIT")
rsYDOSCD70.DOSCD7OPE = rsADO("DOSCD7OPE")
rsYDOSCD70.DOSCD7NUM = rsADO("DOSCD7NUM")
rsYDOSCD70.DOSCD7KCN = rsADO("DOSCD7KCN")
rsYDOSCD70.DOSCD7KNAT = rsADO("DOSCD7KNAT")
rsYDOSCD70.DOSCD7PCI = rsADO("DOSCD7PCI")
rsYDOSCD70.DOSCD7CLI = rsADO("DOSCD7CLI")
rsYDOSCD70.DOSCD7MTD = rsADO("DOSCD7MTD")
rsYDOSCD70.DOSCD7DEV = rsADO("DOSCD7DEV")
rsYDOSCD70.DOSCD7STA = rsADO("DOSCD7STA")
rsYDOSCD70.DOSCD7DDEB = rsADO("DOSCD7DDEB")
rsYDOSCD70.DOSCD7DFIN = rsADO("DOSCD7DFIN")
rsYDOSCD70.DOSCD7DAMJ = rsADO("DOSCD7DAMJ")

Exit Function

Error_Handler:

rsYDOSCD70_GetBuffer = Error

End Function


Public Function sqlYDOSCD70_Update_Field(oldY As typeYDOSCD70, lSQL_Set As String)
Dim xSql As String, Nb As Long

On Error GoTo Error_Handler
sqlYDOSCD70_Update_Field = Null


xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YDOSCD70 " & lSQL_Set & "" _
     & " where DOSCD7DSIT= " & oldY.DOSCD7DSIT _
     & " and DOSCD7OPE = '" & oldY.DOSCD7OPE & "'" _
     & " and DOSCD7NUM = " & oldY.DOSCD7NUM _
     & " and DOSCD7KCN = '" & oldY.DOSCD7KCN & "'" _
     & " and DOSCD7KNAT = '" & oldY.DOSCD7KNAT & "'" _
     & " and DOSCD7PCI = '" & oldY.DOSCD7PCI & "'" _
     & " and DOSCD7DDEB = " & oldY.DOSCD7DDEB _
     & " and DOSCD7DFIN = " & oldY.DOSCD7DFIN

'Call FEU_ROUGE
'Set rsADO = cnSab_Update.Execute(xSql, Nb)  'FICHIER NON JOURNALISé
Set rsADO = cnsab.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYDOSCD70_Update_Field = "Erreur màj : " & oldY.DOSCD7DSIT & " - " & oldY.DOSCD7NUM
    Exit Function
End If
    

Exit Function
Error_Handler:
    sqlYDOSCD70_Update_Field = Error
End Function







Public Sub rsYDOSCD70_Init(lYDOSCD70 As typeYDOSCD70)

lYDOSCD70.DOSCD7DSIT = 0
lYDOSCD70.DOSCD7OPE = ""
lYDOSCD70.DOSCD7NUM = 0
lYDOSCD70.DOSCD7KCN = ""
lYDOSCD70.DOSCD7KNAT = ""
lYDOSCD70.DOSCD7CLI = ""
lYDOSCD70.DOSCD7PCI = ""
lYDOSCD70.DOSCD7MTD = 0
lYDOSCD70.DOSCD7DEV = ""
lYDOSCD70.DOSCD7STA = ""
lYDOSCD70.DOSCD7DDEB = 0
lYDOSCD70.DOSCD7DFIN = 0
lYDOSCD70.DOSCD7DAMJ = 0

End Sub





