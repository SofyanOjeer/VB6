Attribute VB_Name = "srvYCLIGRP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYCLIGRP0
 
      CLIGRPETB     As Long
      CLIGRPCLI     As String * 7
      CLIGRPREG     As String * 7
      CLIGRPREL     As String * 3
      CLIGRPCOM     As String * 28
      CLIGRPAUT     As String * 1
      CLIGRPRAT     As String * 1
      CLIGRPTAU     As Double
      CLIGRPPAR     As Long
      
      CLIGRPCLI_RA1 As String
      CLIGRPREG_RA1 As String
End Type
Public xYCLIGRP0 As typeYCLIGRP0
Public Function sqlYCLIGRP0_Insert(newY As typeYCLIGRP0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCLIGRP0_Insert = Null

xSet = " (CLIGRPETB"
xValues = " values(" & newY.CLIGRPETB

' Détecter les modifications
'===================================================================================
xSql = "Insert into " & paramIBM_Library_SABSPE & ".YCLIGRP0" & xSet & ")" & xValues & ")"

Set rsADO = cnAdo.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCLIGRP0_Insert = "Erreur màj : " & newY.CLIGRPETB
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCLIGRP0_Insert = Error
End Function

Public Function sqlYCLIGRP0_Update(newY As typeYCLIGRP0, oldY As typeYCLIGRP0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlYCLIGRP0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.CLIGRPETB <> newY.CLIGRPETB Then
    sqlYCLIGRP0_Update = "Erreur CLIGRPETB : " & newY.CLIGRPETB & " / " & oldY.CLIGRPETB
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================


xSql = "update " & paramIBM_Library_SABSPE & ".YCLIGRP0" & xSet & xWhere

Set rsADO = cnAdo.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCLIGRP0_Update = "Erreur màj : " & newY.CLIGRPETB
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCLIGRP0_Update = Error
End Function

Public Function sqlYCLIGRP0_Init(newY As typeYCLIGRP0, cnAdo As ADODB.Connection, rsADO As ADODB.Recordset)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim xxx As typeYCLIGRP0

On Error GoTo Error_Handler
sqlYCLIGRP0_Init = Null
xSql = "update " & paramIBM_Library_SABSPE & ".YCLIGRP0" & xSet & xWhere

Set rsADO = cnAdo.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCLIGRP0_Init = "Erreur màj : " & newY.CLIGRPETB
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCLIGRP0_Init = Error
End Function

Public Function srvYCLIGRP0_GetBuffer_ODBC(rsADO As ADODB.Recordset, lYCLIGRP0 As typeYCLIGRP0)
On Error GoTo Error_Handler
srvYCLIGRP0_GetBuffer_ODBC = Null
lYCLIGRP0.CLIGRPETB = rsADO("CLIGRPETB")
lYCLIGRP0.CLIGRPCLI = rsADO("CLIGRPCLI")
lYCLIGRP0.CLIGRPREG = rsADO("CLIGRPREG")
lYCLIGRP0.CLIGRPREL = rsADO("CLIGRPREL")
lYCLIGRP0.CLIGRPCOM = rsADO("CLIGRPCOM")
lYCLIGRP0.CLIGRPAUT = rsADO("CLIGRPAUT")
lYCLIGRP0.CLIGRPRAT = rsADO("CLIGRPRAT")
lYCLIGRP0.CLIGRPTAU = rsADO("CLIGRPTAU")
lYCLIGRP0.CLIGRPPAR = rsADO("CLIGRPPAR")

Exit Function
Error_Handler:
srvYCLIGRP0_GetBuffer_ODBC = Error


End Function

Public Function srvYCLIGRP0_Init(lYCLIGRP0 As typeYCLIGRP0)
lYCLIGRP0.CLIGRPETB = 0
lYCLIGRP0.CLIGRPCLI = ""
lYCLIGRP0.CLIGRPREG = ""
lYCLIGRP0.CLIGRPREL = ""
lYCLIGRP0.CLIGRPCOM = ""
lYCLIGRP0.CLIGRPAUT = ""
lYCLIGRP0.CLIGRPRAT = ""
lYCLIGRP0.CLIGRPTAU = 0
lYCLIGRP0.CLIGRPPAR = 0

lYCLIGRP0.CLIGRPREG_RA1 = ""
lYCLIGRP0.CLIGRPCLI_RA1 = ""
End Function

Public Sub srvYCLIGRP0_fgDisplay(lYCLIGRP0 As typeYCLIGRP0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 12
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "CLIGRPCLI   9P"
fgDisplay.Col = 1: fgDisplay = "Identification"
fgDisplay.Col = 2: fgDisplay = lYCLIGRP0.CLIGRPCLI
End Sub



Public Function libCLIGRPREL(lCLIGRPREL As String) As String
Dim X As String
X = lCLIGRPREL
Select Case lCLIGRPREL
    Case "ADM": X = "Administrateurs"
    Case "DIR": X = "Dirigeants"
    Case "FIL": X = "Filiales"
    Case "GGR": X = "Groupes"
 End Select
libCLIGRPREL = X

End Function
