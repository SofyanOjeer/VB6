Attribute VB_Name = "srvYDOSXOD0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYDOSXOD0

    DOSXODDTR   As Long
    DOSXODPIE   As Long
    DOSXODECR   As Long
    DOSXODKDC   As String
    DOSXODOPE   As String
    DOSXODNUM   As Long
    DOSXODLIB   As String
    DOSXODUUSR  As String
    DOSXODUAMJ   As String
    DOSXODUHMS   As String
End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYDOSXOD0_GetBuffer(rsADO As ADODB.Recordset, rsYDOSXOD0 As typeYDOSXOD0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYDOSXOD0_GetBuffer = Null

rsYDOSXOD0.DOSXODOPE = rsADO("DOSXODOPE")
rsYDOSXOD0.DOSXODNUM = rsADO("DOSXODNUM")
rsYDOSXOD0.DOSXODLIB = rsADO("DOSXODLIB")
rsYDOSXOD0.DOSXODUUSR = rsADO("DOSXODUUSR")
rsYDOSXOD0.DOSXODUAMJ = rsADO("DOSXODUAMJ")

rsYDOSXOD0.DOSXODUHMS = rsADO("DOSXODUHMS")
rsYDOSXOD0.DOSXODDTR = rsADO("DOSXODDTR")
rsYDOSXOD0.DOSXODPIE = rsADO("DOSXODPIE")
rsYDOSXOD0.DOSXODECR = rsADO("DOSXODECR")
rsYDOSXOD0.DOSXODKDC = rsADO("DOSXODKDC")

Exit Function

Error_Handler:

rsYDOSXOD0_GetBuffer = Error

End Function









Public Sub rsYDOSXOD0_Init(lYDOSXOD0 As typeYDOSXOD0)
lYDOSXOD0.DOSXODUUSR = ""
lYDOSXOD0.DOSXODLIB = ""
lYDOSXOD0.DOSXODUAMJ = 0
lYDOSXOD0.DOSXODOPE = ""
lYDOSXOD0.DOSXODUHMS = 0
lYDOSXOD0.DOSXODDTR = 0
lYDOSXOD0.DOSXODNUM = 0
lYDOSXOD0.DOSXODPIE = 0
lYDOSXOD0.DOSXODECR = 0
lYDOSXOD0.DOSXODKDC = ""

End Sub





Public Function sqlYDOSXOD0_Update(newY As typeYDOSXOD0, oldY As typeYDOSXOD0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYDOSXOD0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DOSXODDTR <> newY.DOSXODDTR _
Or oldY.DOSXODPIE <> newY.DOSXODPIE _
Or oldY.DOSXODECR <> newY.DOSXODECR Then
    sqlYDOSXOD0_Update = "Erreur DOSXODDTR : " & newY.DOSXODDTR & "." & oldY.DOSXODPIE & "." & oldY.DOSXODECR
    Exit Function
End If
'===================================================================================

xWhere = " where DOSXODDTR = " & oldY.DOSXODDTR _
       & " and DOSXODPIE = " & oldY.DOSXODPIE & " and DOSXODECR = " & oldY.DOSXODECR

newY.DOSXODUUSR = usrName_UCase
newY.DOSXODUAMJ = DSys
newY.DOSXODUHMS = time_Hms

xSet = xSet & " set DOSXODUUSR ='" & Trim(newY.DOSXODUUSR) & "'"
blnUpdate = False


' Détecter les modifications
'===================================================================================
If newY.DOSXODNUM <> oldY.DOSXODNUM Then blnUpdate = True:  xSet = xSet & " , DOSXODNUM = " & newY.DOSXODNUM
If newY.DOSXODUAMJ <> oldY.DOSXODUAMJ Then blnUpdate = True:  xSet = xSet & " , DOSXODUAMJ = " & newY.DOSXODUAMJ
If newY.DOSXODUHMS <> oldY.DOSXODUHMS Then blnUpdate = True:  xSet = xSet & " , DOSXODUHMS = " & newY.DOSXODUHMS

If newY.DOSXODOPE <> oldY.DOSXODOPE Then blnUpdate = True:  xSet = xSet & " , DOSXODOPE = '" & newY.DOSXODOPE & "'"


If blnUpdate Then
    
    xSql = "update " & paramIBM_Library_SABSPE & ".YDOSXOD0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnsab.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYDOSXOD0_Update = "Erreur màj : " & newY.DOSXODDTR
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYDOSXOD0_Update = Error
End Function


