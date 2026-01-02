Attribute VB_Name = "srvZADRESS0"
Option Explicit

Public Function sqlZADRESS0_Update(newY As typeYADRESS0, oldY As typeYADRESS0, cnAdo As ADODB.Connection)
' attention
'$$$$$$$$$$$$ ne gère pas tous les champs, par ex : datedébut,tel,telex .....
'===============================================================================

Dim X As String, xSql As String, Nb As Long, K As Integer
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
Dim rsADO As New ADODB.Recordset
On Error GoTo Error_Handler

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.ADRESSETA = newY.ADRESSETA _
And oldY.ADRESSTYP = newY.ADRESSTYP _
And oldY.ADRESSPLA = newY.ADRESSPLA _
And oldY.ADRESSNUM = newY.ADRESSNUM _
And oldY.ADRESSCOA = newY.ADRESSCOA Then
    sqlZADRESS0_Update = Null
Else
    sqlZADRESS0_Update = "Erreur ADRESSNUM : " & newY.ADRESSNUM & " / " & oldY.ADRESSNUM
    Exit Function
End If
'===================================================================================

xWhere = " where ADRESSETA = " & oldY.ADRESSETA _
       & " and ADRESSTYP = '" & oldY.ADRESSTYP & "'" _
       & " and ADRESSPLA = " & oldY.ADRESSPLA _
       & " and ADRESSNUM = '" & oldY.ADRESSNUM & "'" _
       & " and ADRESSCOA = '" & oldY.ADRESSCOA & "'"

xSet = " set"
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.ADRESSRA1 <> oldY.ADRESSRA1 Then blnUpdate = True: xSet = xSet & " , ADRESSRA1 = '" & newY.ADRESSRA1 & "'"
If newY.ADRESSRA2 <> oldY.ADRESSRA2 Then blnUpdate = True:  xSet = xSet & " , ADRESSRA2 = '" & newY.ADRESSRA2 & "'"
If newY.ADRESSAD1 <> oldY.ADRESSAD1 Then blnUpdate = True:  xSet = xSet & " , ADRESSAD1 = '" & newY.ADRESSAD1 & "'"
If newY.ADRESSAD2 <> oldY.ADRESSAD2 Then blnUpdate = True:  xSet = xSet & " , ADRESSAD2 = '" & newY.ADRESSAD2 & "'"
If newY.ADRESSAD3 <> oldY.ADRESSAD3 Then blnUpdate = True:  xSet = xSet & " , ADRESSAD3 = '" & newY.ADRESSAD3 & "'"
If newY.ADRESSCOP <> oldY.ADRESSCOP Then blnUpdate = True:  xSet = xSet & " , ADRESSCOP = '" & newY.ADRESSCOP & "'"
If newY.ADRESSVIL <> oldY.ADRESSVIL Then blnUpdate = True:  xSet = xSet & " , ADRESSVIL = '" & newY.ADRESSVIL & "'"
If newY.ADRESSPAY <> oldY.ADRESSPAY Then blnUpdate = True:  xSet = xSet & " , ADRESSPAY = '" & newY.ADRESSPAY & "'"

If blnUpdate Then
    'Si modification , supprimer la première virgule
    K = InStr(5, xSet, ","): If K > 0 Then Mid$(xSet, K, 1) = " "

    xSql = "update " & paramIBM_Library_SAB & ".ZADRESS0" & xSet & xWhere
    
    Set rsADO = cnAdo.Execute(xSql, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlZADRESS0_Update = "Erreur màj : " & newY.ADRESSNUM
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlZADRESS0_Update = Error
End Function


'---------------------------------------------------------
Public Function srvZADRESS0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvZADRESS0_GetBuffer_ODBC = Null

    recYADRESS0.ADRESSETA = rsADO("ADRESSETA")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYADRESS0.ADRESSTYP = rsADO("ADRESSTYP")    'mId$(MsgTxt, K + 6, 1)
    recYADRESS0.ADRESSPLA = rsADO("ADRESSPLA")    'CLng(Val(mId$(MsgTxt, K + 7, 4)))
    recYADRESS0.ADRESSNUM = rsADO("ADRESSNUM")    'mId$(MsgTxt, K + 11, 20)
    recYADRESS0.ADRESSCOA = rsADO("ADRESSCOA")    'mId$(MsgTxt, K + 31, 2)
    recYADRESS0.ADRESSDLI = rsADO("ADRESSDLI")    'CLng(Val(mId$(MsgTxt, K + 33, 8)))
    recYADRESS0.ADRESSDDE = rsADO("ADRESSDDE")    'CLng(Val(mId$(MsgTxt, K + 41, 8)))
    recYADRESS0.ADRESSRA1 = rsADO("ADRESSRA1")    'mId$(MsgTxt, K + 49, 32)
    recYADRESS0.ADRESSRA2 = rsADO("ADRESSRA2")    'mId$(MsgTxt, K + 81, 32)
    recYADRESS0.ADRESSAD1 = rsADO("ADRESSAD1")    'mId$(MsgTxt, K + 113, 32)
    recYADRESS0.ADRESSAD2 = rsADO("ADRESSAD2")    'mId$(MsgTxt, K + 145, 32)
    recYADRESS0.ADRESSAD3 = rsADO("ADRESSAD3")    'mId$(MsgTxt, K + 177, 32)
    recYADRESS0.ADRESSCOP = rsADO("ADRESSCOP")    'mId$(MsgTxt, K + 209, 6)
    recYADRESS0.ADRESSVIL = rsADO("ADRESSVIL")    'mId$(MsgTxt, K + 215, 25)
    recYADRESS0.ADRESSPAY = rsADO("ADRESSPAY")    'mId$(MsgTxt, K + 240, 25)
    recYADRESS0.ADRESSTEL = rsADO("ADRESSTEL")    'mId$(MsgTxt, K + 265, 20)
    recYADRESS0.ADRESSFAX = rsADO("ADRESSFAX")    'mId$(MsgTxt, K + 285, 20)
    recYADRESS0.ADRESSTEX = rsADO("ADRESSTEX")    'mId$(MsgTxt, K + 305, 20)


Exit Function

Error_Handler:
srvZADRESS0_GetBuffer_ODBC = Error

End Function


