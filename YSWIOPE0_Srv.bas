Attribute VB_Name = "srvYSWIOPE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYSWIOPE0
 
      SWIOPEID    As Long         'IDENTIFICATION1
      SWISABCOP   As String * 3   'SAB CODE OPE
      SWISABDOS   As Long         'SAB NUMERO DOS
      SWISABCPTD  As Long         'Date compta
      SAAAID      As Long         'SAA ID-1
      SAAUMIDL    As Long         'SAA ID-2
      SAAUMIDH    As Long         'SAA ID-3
      SWIOPESTA   As String * 4   'STATUT EN COURS
      SWIOPESTAD  As Long         'STA:DATE MAJ
      SWIOPESTAH  As Long         'STA:HEURE MAJ
      SWIOPEFLUD  As Long         'FLUX : DATE TRT
      SWIOPEFLUH  As Long         'FLUX : HEURE TRT
      SWIOPEXMT   As String * 3   'TYPE MSG
      SWIOPEXBIC  As String * 11  'BIC Sender
      SWIOPEXTRN   As String * 16  'CHAMP 20
      SWIOPEX32A  As Currency      'MONTANT
      SWIOPEX32D  As String * 3   'DEVISE
      SWIOPEX32V  As Long         'DATE VALEUR
      SWIOPEUPDS  As Long         'Sequence mise à jour

End Type
Public xYSWIOPE0 As typeYSWIOPE0
Public Function sqlYSWIOPE0_Insert(newY As typeYSWIOPE0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSWIOPE0_Insert = Null

xSet = " (SWIOPEID"
xValues = " values(" & newY.SWIOPEID

' Détecter les modifications
'===================================================================================
If Trim(newY.SWISABCOP) <> "" Then xSet = xSet & ",SWISABCOP": xValues = xValues & " ,'" & newY.SWISABCOP & "'"
If newY.SWISABDOS <> 0 Then xSet = xSet & ",SWISABDOS": xValues = xValues & " ," & newY.SWISABDOS
If newY.SWISABCPTD <> 0 Then xSet = xSet & ",SWISABCPTD": xValues = xValues & " ," & newY.SWISABCPTD
If newY.SAAAID <> 0 Then xSet = xSet & ",SAAAID": xValues = xValues & " ," & newY.SAAAID
If newY.SAAUMIDL <> 0 Then xSet = xSet & ",SAAUMIDL": xValues = xValues & " ," & newY.SAAUMIDL
If newY.SAAUMIDH <> 0 Then xSet = xSet & ",SAAUMIDH": xValues = xValues & " ," & newY.SAAUMIDH
If Trim(newY.SWIOPESTA) <> "" Then xSet = xSet & ",SWIOPESTA": xValues = xValues & " ,'" & newY.SWIOPESTA & "'"
If newY.SWIOPESTAD <> 0 Then xSet = xSet & ",SWIOPESTAD": xValues = xValues & " ," & newY.SWIOPESTAD
If newY.SWIOPESTAH <> 0 Then xSet = xSet & ",SWIOPESTAH": xValues = xValues & " ," & newY.SWIOPESTAH
If newY.SWIOPEFLUD <> 0 Then xSet = xSet & ",SWIOPEFLUD": xValues = xValues & " ," & newY.SWIOPEFLUD
If newY.SWIOPEFLUH <> 0 Then xSet = xSet & ",SWIOPEFLUH": xValues = xValues & " ," & newY.SWIOPEFLUH
If Trim(newY.SWIOPEXMT) <> "" Then xSet = xSet & ",SWIOPEXMT": xValues = xValues & " ,'" & newY.SWIOPEXMT & "'"
If Trim(newY.SWIOPEXBIC) <> "" Then xSet = xSet & ",SWIOPEXBIC": xValues = xValues & " ,'" & newY.SWIOPEXBIC & "'"
If Trim(newY.SWIOPEXTRN) <> "" Then xSet = xSet & ",SWIOPEXTRN": xValues = xValues & " ,'" & newY.SWIOPEXTRN & "'"
If newY.SWIOPEX32A <> 0 Then xSet = xSet & ",SWIOPEX32A": xValues = xValues & " ," & cur_P(newY.SWIOPEX32A)
If Trim(newY.SWIOPEX32D) <> "" Then xSet = xSet & ",SWIOPEX32D": xValues = xValues & " ,'" & newY.SWIOPEX32D & "'"
If newY.SWIOPEX32V <> 0 Then xSet = xSet & ",SWIOPEX32V": xValues = xValues & " ," & newY.SWIOPEX32V

xSql = "Insert into " & paramIBM_Library_SABSPE & ".YSWIOPE0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnAdo.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIOPE0_Insert = "Erreur màj : " & newY.SWIOPEID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWIOPE0_Insert = Error
End Function

Public Function sqlYSWIOPE0_Update(newY As typeYSWIOPE0, oldY As typeYSWIOPE0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlYSWIOPE0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SWIOPEID <> newY.SWIOPEID Then
    sqlYSWIOPE0_Update = "Erreur SWIOPEID : " & newY.SWIOPEID & " / " & oldY.SWIOPEID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWIOPEID = " & oldY.SWIOPEID & " and SWIOPEUPDS = " & oldY.SWIOPEUPDS

newY.SWIOPEUPDS = newY.SWIOPEUPDS + 1
xSet = xSet & " set SWIOPEUPDS = " & newY.SWIOPEUPDS

' Détecter les modifications
'===================================================================================
If newY.SWISABCOP <> oldY.SWISABCOP Then xSet = xSet & " , SWISABCOP = '" & newY.SWISABCOP & "'"
If newY.SWISABDOS <> oldY.SWISABDOS Then xSet = xSet & " , SWISABDOS = " & newY.SWISABDOS
If newY.SWISABCPTD <> oldY.SWISABCPTD Then xSet = xSet & " , SWISABCPTD = " & newY.SWISABCPTD
If newY.SAAAID <> oldY.SAAAID Then xSet = xSet & " , SAAAID = " & newY.SAAAID
If newY.SAAUMIDL <> oldY.SAAUMIDL Then xSet = xSet & " , SAAUMIDL = " & newY.SAAUMIDL
If newY.SAAUMIDH <> oldY.SAAUMIDH Then xSet = xSet & " , SAAUMIDH = " & newY.SAAUMIDH
If newY.SWIOPESTA <> oldY.SWIOPESTA Then xSet = xSet & " , SWIOPESTA = '" & newY.SWIOPESTA & "'"
If newY.SWIOPESTAD <> oldY.SWIOPESTAD Then xSet = xSet & " , SWIOPESTAD = " & newY.SWIOPESTAD
If newY.SWIOPESTAH <> oldY.SWIOPESTAH Then xSet = xSet & " , SWIOPESTAH = " & newY.SWIOPESTAH
If newY.SWIOPEFLUD <> oldY.SWIOPEFLUD Then xSet = xSet & " , SWIOPEFLUD = " & newY.SWIOPEFLUD
If newY.SWIOPEFLUH <> oldY.SWIOPEFLUH Then xSet = xSet & " , SWIOPEFLUH = " & newY.SWIOPEFLUH
If newY.SWIOPEXMT <> oldY.SWIOPEXMT Then xSet = xSet & " , SWIOPEXMT = '" & newY.SWIOPEXMT & "'"
If newY.SWIOPEXBIC <> oldY.SWIOPEXBIC Then xSet = xSet & " , SWIOPEXBIC = '" & newY.SWIOPEXBIC & "'"
If newY.SWIOPEXTRN <> oldY.SWIOPEXTRN Then xSet = xSet & " , SWIOPEXTRN = '" & newY.SWIOPEXTRN & "'"
If newY.SWIOPEX32A <> oldY.SWIOPEX32A Then xSet = xSet & " , SWIOPEX32A = " & cur_P(newY.SWIOPEX32A)
If newY.SWIOPEX32D <> oldY.SWIOPEX32D Then xSet = xSet & " , SWIOPEX32D = '" & newY.SWIOPEX32D & "'"
If newY.SWIOPEX32V <> oldY.SWIOPEX32V Then xSet = xSet & " , SWIOPEX32V = " & newY.SWIOPEX32V

xSql = "update " & paramIBM_Library_SABSPE & ".YSWIOPE0" & xSet & xWhere
Call FEU_ROUGE
Set rsADO = cnAdo.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIOPE0_Update = "Erreur màj : " & newY.SWIOPEID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWIOPE0_Update = Error
End Function

Public Function sqlYSWIOPE0_Init(newY As typeYSWIOPE0, cnAdo As ADODB.Connection, rsADO As ADODB.Recordset)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim xxx As typeYSWIOPE0

On Error GoTo Error_Handler
sqlYSWIOPE0_Init = Null

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIOPE0" & " where  SWIOPEID =  -1"
Set rsADO = cnAdo.Execute(xSql, Nb)

xxx.SWIOPEUPDS = rsADO("SWIOPEUPDS")
newY.SWIOPEID = rsADO("SWISABDOS") + 1
newY.SWIOPESTAD = DSys
newY.SWIOPESTAH = time_Hms
newY.SWIOPEUPDS = 0

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWIOPEID = -1" & " and SWIOPEUPDS = " & xxx.SWIOPEUPDS

xSet = " set SWIOPEUPDS = " & xxx.SWIOPEUPDS + 1 & " , SWISABDOS = " & newY.SWIOPEID & " , SWIOPESTAD = " & newY.SWIOPESTAD & " , SWIOPESTAH = " & newY.SWIOPESTAH


xSql = "update " & paramIBM_Library_SABSPE & ".YSWIOPE0" & xSet & xWhere
Call FEU_ROUGE
Set rsADO = cnAdo.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIOPE0_Init = "Erreur màj : " & newY.SWIOPEID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWIOPE0_Init = Error
End Function

Public Function srvYSWIOPE0_GetBuffer_ODBC(rsADO As ADODB.Recordset, lYSWIOPE0 As typeYSWIOPE0)
On Error GoTo Error_Handler
srvYSWIOPE0_GetBuffer_ODBC = Null
lYSWIOPE0.SWIOPEID = rsADO("SWIOPEID")
lYSWIOPE0.SWISABCOP = rsADO("SWISABCOP")
lYSWIOPE0.SWISABDOS = rsADO("SWISABDOS")
lYSWIOPE0.SWISABCPTD = rsADO("SWISABCPTD")
lYSWIOPE0.SAAAID = rsADO("SAAAID")
lYSWIOPE0.SAAUMIDL = rsADO("SAAUMIDL")
lYSWIOPE0.SAAUMIDH = rsADO("SAAUMIDH")
lYSWIOPE0.SWIOPESTA = rsADO("SWIOPESTA")
lYSWIOPE0.SWIOPESTAD = rsADO("SWIOPESTAD")
lYSWIOPE0.SWIOPESTAH = rsADO("SWIOPESTAH")
lYSWIOPE0.SWIOPEFLUD = rsADO("SWIOPEFLUD")
lYSWIOPE0.SWIOPEFLUH = rsADO("SWIOPEFLUH")
lYSWIOPE0.SWIOPEXMT = rsADO("SWIOPEXMT")
lYSWIOPE0.SWIOPEXBIC = rsADO("SWIOPEXBIC")
lYSWIOPE0.SWIOPEXTRN = rsADO("SWIOPEXTRN")
lYSWIOPE0.SWIOPEX32A = rsADO("SWIOPEX32A")
lYSWIOPE0.SWIOPEX32D = rsADO("SWIOPEX32D")
lYSWIOPE0.SWIOPEX32V = rsADO("SWIOPEX32V")
lYSWIOPE0.SWIOPEUPDS = rsADO("SWIOPEUPDS")

Exit Function
Error_Handler:
srvYSWIOPE0_GetBuffer_ODBC = Error


End Function

Public Function srvYSWIOPE0_Init(lYSWIOPE0 As typeYSWIOPE0)
lYSWIOPE0.SWIOPEID = 0
lYSWIOPE0.SWISABCOP = ""
lYSWIOPE0.SWISABDOS = 0
lYSWIOPE0.SWISABCPTD = 0
lYSWIOPE0.SAAAID = 0
lYSWIOPE0.SAAUMIDL = 0
lYSWIOPE0.SAAUMIDH = 0
lYSWIOPE0.SWIOPESTA = ""
lYSWIOPE0.SWIOPESTAD = 0
lYSWIOPE0.SWIOPESTAH = 0
lYSWIOPE0.SWIOPEFLUD = 0
lYSWIOPE0.SWIOPEFLUH = 0
lYSWIOPE0.SWIOPEXMT = ""
lYSWIOPE0.SWIOPEXBIC = ""
lYSWIOPE0.SWIOPEXTRN = ""
lYSWIOPE0.SWIOPEX32A = 0
lYSWIOPE0.SWIOPEX32D = ""
lYSWIOPE0.SWIOPEX32V = 0
lYSWIOPE0.SWIOPEUPDS = 0

End Function

Public Sub srvYSWIOPE0_fgDisplay(lYSWIOPE0 As typeYSWIOPE0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 12
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIOPEID   9P"
fgDisplay.Col = 1: fgDisplay = "Identification"
fgDisplay.Col = 2: fgDisplay = lYSWIOPE0.SWIOPEID
End Sub


