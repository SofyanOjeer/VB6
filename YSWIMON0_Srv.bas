Attribute VB_Name = "srvYSWIMON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYSWIMON0
 
      SWIMONID    As Long         'IDENTIFICATION                        1       9      0     P  *
      SWISABNUM   As Long         'SAB NUMERO INTERN                     6       9      0     P  *
      SWISABCOP   As String * 3   'SAB CODE OPE                         11       3            A  *
      SWISABDOS   As Long         'SAB NUMERO DOS                       14       9      0     P  *
      SAAAID      As Long         'SAA ID-1                             19       7      0     P  *
      SAAUMIDL    As Long         'SAA ID-2                             23      15      0     P  *
      SAAUMIDH    As Long         'SAA ID-3                             31      15      0     P  *
      SAAQUEUE    As String * 20  'QUEUE EN COURS                       39      20            A  *
      SAAQMOD     As Integer
      SAAQOFAC    As Integer
      SAAUNIT     As String * 4   'UNITE EN COURS                       59       4            A  *
      SWIMONSTA   As String * 4   'STATUT EN COURS                      63       4            A  *
      SWIMONSTAD  As Long         'STA:DATE MAJ                         67       8      0     S  *
      SWIMONSTAH  As Long         'STA:HEURE MAJ                        75       6      0     S  *
      SWIMONFLUX  As String * 1   'FLUX : E / S                         81       1            A  *
      SWIMONFLUQ  As String * 2   'FLUX : QUEUE IN                      82       2            A  *
      SWIMONFLUD  As Long         'FLUX : DATE TRT                      84       8      0     S  *
      SWIMONFLUH  As Long         'FLUX : HEURE TRT                     92       6      0     S  *
      SWIMONFLUS  As Long         'FLUX : SEQUENCE                      98       5      0     P  *
      SWIMONXMT   As String * 3   'TYPE MSG                            101       3            A  *
      SWIMONX20   As String * 16  'CHAMP 20 :                          104      16            A  *
      SWIMONX21   As String * 16  'CHAMP 21                            120      16            A  *
      SWIMONX32A  As Currency         'MONTANT                             136      15      2     P  *
      SWIMONX32D  As String * 3   'DEVISE                              144       3            A  *
      SWIMONX32V  As Long         'DATE VALEUR                         147       9      0     S  *
      SWIMONUPDS  As Long         'Sequence mise à jour

End Type
Public xYSWIMON0 As typeYSWIMON0
Public Function sqlYSWIMON0_Insert(newY As typeYSWIMON0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSWIMON0_Insert = Null

xSet = " (SWIMONID"
xValues = " values(" & newY.SWIMONID

' Détecter les modifications
'===================================================================================
If newY.SWISABNUM <> 0 Then xSet = xSet & ",SWISABNUM": xValues = xValues & " ," & newY.SWISABNUM
If Trim(newY.SWISABCOP) <> "" Then xSet = xSet & ",SWISABCOP": xValues = xValues & " ,'" & newY.SWISABCOP & "'"
If newY.SWISABDOS <> 0 Then xSet = xSet & ",SWISABDOS": xValues = xValues & " ," & newY.SWISABDOS
If newY.SAAAID <> 0 Then xSet = xSet & ",SAAAID": xValues = xValues & " ," & newY.SAAAID
If newY.SAAUMIDL <> 0 Then xSet = xSet & ",SAAUMIDL": xValues = xValues & " ," & newY.SAAUMIDL
If newY.SAAUMIDH <> 0 Then xSet = xSet & ",SAAUMIDH": xValues = xValues & " ," & newY.SAAUMIDH
If Trim(newY.SAAQUEUE) <> "" Then xSet = xSet & ",SAAQUEUE": xValues = xValues & " ,'" & Trim(newY.SAAQUEUE) & "'"
If newY.SAAQMOD <> 0 Then xSet = xSet & ",SAAqmod": xValues = xValues & " ," & newY.SAAQMOD
If newY.SAAQOFAC <> 0 Then xSet = xSet & ",SAAQOFAC": xValues = xValues & " ," & newY.SAAQOFAC
If Trim(newY.SAAUNIT) <> "" Then xSet = xSet & ",SAAUNIT": xValues = xValues & " ,'" & newY.SAAUNIT & "'"
If Trim(newY.SWIMONSTA) <> "" Then xSet = xSet & ",SWIMONSTA": xValues = xValues & " ,'" & newY.SWIMONSTA & "'"
If newY.SWIMONSTAD <> 0 Then xSet = xSet & ",SWIMONSTAD": xValues = xValues & " ," & newY.SWIMONSTAD
If newY.SWIMONSTAH <> 0 Then xSet = xSet & ",SWIMONSTAH": xValues = xValues & " ," & newY.SWIMONSTAH
If Trim(newY.SWIMONFLUX) <> "" Then xSet = xSet & ",SWIMONFLUX": xValues = xValues & " ,'" & newY.SWIMONFLUX & "'"
If Trim(newY.SWIMONFLUQ) <> "" Then xSet = xSet & ",SWIMONFLUQ": xValues = xValues & " ,'" & newY.SWIMONFLUQ & "'"
If newY.SWIMONFLUD <> 0 Then xSet = xSet & ",SWIMONFLUD": xValues = xValues & " ," & newY.SWIMONFLUD
If newY.SWIMONFLUH <> 0 Then xSet = xSet & ",SWIMONFLUH": xValues = xValues & " ," & newY.SWIMONFLUH
If newY.SWIMONFLUS <> 0 Then xSet = xSet & ",SWIMONFLUS": xValues = xValues & " ," & newY.SWIMONFLUS
If Trim(newY.SWIMONXMT) <> "" Then xSet = xSet & ",SWIMONXMT": xValues = xValues & " ,'" & newY.SWIMONXMT & "'"
If Trim(newY.SWIMONX20) <> "" Then xSet = xSet & ",SWIMONX20": xValues = xValues & " ,'" & Replace(Trim(newY.SWIMONX20), "'", "''") & "'"
If Trim(newY.SWIMONX21) <> "" Then xSet = xSet & ",SWIMONX21": xValues = xValues & " ,'" & Replace(Trim(newY.SWIMONX21), "'", "''") & "'"
If newY.SWIMONX32A <> 0 Then xSet = xSet & ",SWIMONX32A": xValues = xValues & " ," & cur_P(newY.SWIMONX32A)
If Trim(newY.SWIMONX32D) <> "" Then xSet = xSet & ",SWIMONX32D": xValues = xValues & " ,'" & newY.SWIMONX32D & "'"
If newY.SWIMONX32V <> 0 Then xSet = xSet & ",SWIMONX32V": xValues = xValues & " ," & newY.SWIMONX32V

xSql = "Insert into " & paramIBM_Library_SABSPE & ".YSWIMON0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnAdo.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIMON0_Insert = "Erreur màj : " & newY.SWIMONID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWIMON0_Insert = Error
End Function

Public Function sqlYSWIMON0_Update(newY As typeYSWIMON0, oldY As typeYSWIMON0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWIMON0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SWIMONID <> newY.SWIMONID Then
    sqlYSWIMON0_Update = "Erreur SWIMONID : " & newY.SWIMONID & " / " & oldY.SWIMONID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWIMONID = " & oldY.SWIMONID & " and SWIMONUPDS = " & oldY.SWIMONUPDS

newY.SWIMONUPDS = newY.SWIMONUPDS + 1
xSet = xSet & " set SWIMONUPDS = " & newY.SWIMONUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SWISABNUM <> oldY.SWISABNUM Then blnUpdate = True: xSet = xSet & " , SWISABNUM = " & newY.SWISABNUM
If newY.SWISABCOP <> oldY.SWISABCOP Then blnUpdate = True:  xSet = xSet & " , SWISABCOP = '" & newY.SWISABCOP & "'"
If newY.SWISABDOS <> oldY.SWISABDOS Then blnUpdate = True:  xSet = xSet & " , SWISABDOS = " & newY.SWISABDOS
If newY.SAAAID <> oldY.SAAAID Then blnUpdate = True:  xSet = xSet & " , SAAAID = " & newY.SAAAID
If newY.SAAUMIDL <> oldY.SAAUMIDL Then blnUpdate = True:  xSet = xSet & " , SAAUMIDL = " & newY.SAAUMIDL
If newY.SAAUMIDH <> oldY.SAAUMIDH Then blnUpdate = True:  xSet = xSet & " , SAAUMIDH = " & newY.SAAUMIDH
If newY.SAAQUEUE <> oldY.SAAQUEUE Then blnUpdate = True:  xSet = xSet & " , SAAQUEUE = '" & Trim(newY.SAAQUEUE) & "'"
If newY.SAAQMOD <> oldY.SAAQMOD Then blnUpdate = True:  xSet = xSet & " , SAAqmod = " & newY.SAAQMOD
If newY.SAAQOFAC <> oldY.SAAQOFAC Then blnUpdate = True:  xSet = xSet & " , SAAQOFAC = " & newY.SAAQOFAC
If newY.SAAUNIT <> oldY.SAAUNIT Then blnUpdate = True:  xSet = xSet & " , SAAUNIT = '" & newY.SAAUNIT & "'"
If newY.SWIMONSTA <> oldY.SWIMONSTA Then blnUpdate = True:  xSet = xSet & " , SWIMONSTA = '" & newY.SWIMONSTA & "'"
If newY.SWIMONFLUX <> oldY.SWIMONFLUX Then blnUpdate = True:  xSet = xSet & " , SWIMONFLUX = '" & newY.SWIMONFLUX & "'"
If newY.SWIMONFLUQ <> oldY.SWIMONFLUQ Then blnUpdate = True:  xSet = xSet & " , SWIMONFLUQ = '" & newY.SWIMONFLUQ & "'"
If newY.SWIMONFLUD <> oldY.SWIMONFLUD Then blnUpdate = True:  xSet = xSet & " , SWIMONFLUD = " & newY.SWIMONFLUD
If newY.SWIMONFLUH <> oldY.SWIMONFLUH Then blnUpdate = True:  xSet = xSet & " , SWIMONFLUH = " & newY.SWIMONFLUH
If newY.SWIMONFLUS <> oldY.SWIMONFLUS Then blnUpdate = True:  xSet = xSet & " , SWIMONFLUS = " & newY.SWIMONFLUS
If newY.SWIMONXMT <> oldY.SWIMONXMT Then blnUpdate = True:  xSet = xSet & " , SWIMONXMT = '" & newY.SWIMONXMT & "'"
If newY.SWIMONX20 <> oldY.SWIMONX20 Then blnUpdate = True:  xSet = xSet & " , SWIMONX20 = '" & newY.SWIMONX20 & "'"
If newY.SWIMONX21 <> oldY.SWIMONX21 Then blnUpdate = True:  xSet = xSet & " , SWIMONX21 = '" & newY.SWIMONX21 & "'"
If newY.SWIMONX32A <> oldY.SWIMONX32A Then blnUpdate = True:  xSet = xSet & " , SWIMONX32A = " & cur_P(newY.SWIMONX32A)
If newY.SWIMONX32D <> oldY.SWIMONX32D Then blnUpdate = True:  xSet = xSet & " , SWIMONX32D = '" & newY.SWIMONX32D & "'"
If newY.SWIMONX32V <> oldY.SWIMONX32V Then blnUpdate = True:  xSet = xSet & " , SWIMONX32V = " & newY.SWIMONX32V

If newY.SWIMONID < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    If newY.SWIMONSTAD <> oldY.SWIMONSTAD Then xSet = xSet & " , SWIMONSTAD = " & newY.SWIMONSTAD
    If newY.SWIMONSTAH <> oldY.SWIMONSTAH Then xSet = xSet & " , SWIMONSTAH = " & newY.SWIMONSTAH
    
    xSql = "update " & paramIBM_Library_SABSPE & ".YSWIMON0" & xSet & xWhere
    Call FEU_ROUGE

    Set rsADO = cnAdo.Execute(xSql, Nb)
    Call FEU_VERT

    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWIMON0_Update = "Erreur màj : " & newY.SWIMONID
        Exit Function
    End If
    
    If newY.SWIMONSTA <> oldY.SWIMONSTA Then
        If newY.SWIMONSTA = "S998" Then srvYSWIMON0_SendMail newY
            'Case Is = "S200", "S800"
            'Case Is > "S900":
            'Case Else: srvYSWIMON0_SendMail newY
        'End Select
    End If
End If

Exit Function
Error_Handler:
    sqlYSWIMON0_Update = Error
End Function

Public Function sqlYSWIMON0_Init(newY As typeYSWIMON0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim xxx As typeYSWIMON0

On Error GoTo Error_Handler
sqlYSWIMON0_Init = Null

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" & " where  SWIMONID =  -1"
Set rsADO = cnAdo.Execute(xSql, Nb)

xxx.SWIMONUPDS = rsADO("SWIMONUPDS")
newY.SWIMONID = rsADO("SWISABNUM") + 1
newY.SWIMONSTAD = DSys
newY.SWIMONSTAH = time_Hms
newY.SWIMONUPDS = 0

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWIMONID = -1" & " and SWIMONUPDS = " & xxx.SWIMONUPDS

xSet = " set SWIMONUPDS = " & xxx.SWIMONUPDS + 1 & " , SWISABNUM = " & newY.SWIMONID & " , SWIMONSTAD = " & newY.SWIMONSTAD & " , SWIMONSTAH = " & newY.SWIMONSTAH


xSql = "update " & paramIBM_Library_SABSPE & ".YSWIMON0" & xSet & xWhere
        Call FEU_ROUGE

Set rsADO = cnAdo.Execute(xSql, Nb)
        Call FEU_VERT

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIMON0_Init = "Erreur màj : " & newY.SWIMONID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWIMON0_Init = Error
End Function

Public Function SAA_from_SAB_TRN(lTRN_SAB As String, lSWISABUnit As String) As String

' A REMPLACER PAR srvSWISABREF

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' copie SAA_Srv.bas pour test de reprise
'$$$$$$$$$$$$$$$$$$$$$$$$$$


Dim wTRN As String, wTRN_Unit As String
Dim K As Integer

wTRN = Trim(lTRN_SAB) & Space$(20)
wTRN_Unit = "": K = 1

Select Case Mid$(wTRN, 1, 2)
    Case "TC": K = 3: wTRN_Unit = "BOTC"
    Case "00":
            Select Case Mid$(wTRN, 3, 3)
                Case "CDE", "CDI", "RDE", "RDI": wTRN_Unit = "SOBI": K = 3
                Case "CPT", "TRF": wTRN_Unit = "ORPA": K = 3
            End Select
    Case Else
            Select Case Mid$(wTRN, 1, 3)
                Case "CDE", "CDI", "RDE", "RDI": wTRN_Unit = "SOBI"
                Case "CPT", "TRF": wTRN_Unit = "ORPA"
               Case "PRE", "EMP", "SWP": wTRN_Unit = "BOTC"
               'Case "CSO", "GSO": wTRN_Unit = "CSOP"
            End Select
End Select

If Trim(lSWISABUnit) <> "" Then wTRN_Unit = lSWISABUnit   'Forçage Service d'origine

If wTRN_Unit = "" Then
    SAA_from_SAB_TRN = Mid$(Trim(wTRN), K, 16)
Else
    SAA_from_SAB_TRN = Trim(wTRN_Unit & Mid$(wTRN, K, 12))
End If

End Function

Public Function srvYSWIMON0_GetBuffer_ODBC(rsADO As ADODB.Recordset, lYSWIMON0 As typeYSWIMON0)
On Error GoTo Error_Handler
srvYSWIMON0_GetBuffer_ODBC = Null
lYSWIMON0.SWIMONID = rsADO("SWIMONID")
lYSWIMON0.SWISABNUM = rsADO("SWISABNUM")
lYSWIMON0.SWISABCOP = rsADO("SWISABCOP")
lYSWIMON0.SWISABDOS = rsADO("SWISABDOS")
lYSWIMON0.SAAAID = rsADO("SAAAID")
lYSWIMON0.SAAUMIDL = rsADO("SAAUMIDL")
lYSWIMON0.SAAUMIDH = rsADO("SAAUMIDH")
lYSWIMON0.SAAQUEUE = rsADO("SAAQUEUE")
lYSWIMON0.SAAQMOD = rsADO("SAAQMOD")
lYSWIMON0.SAAQOFAC = rsADO("SAAQOFAC")
lYSWIMON0.SAAUNIT = rsADO("SAAUNIT")
lYSWIMON0.SWIMONSTA = rsADO("SWIMONSTA")
lYSWIMON0.SWIMONSTAD = rsADO("SWIMONSTAD")
lYSWIMON0.SWIMONSTAH = rsADO("SWIMONSTAH")
lYSWIMON0.SWIMONFLUX = rsADO("SWIMONFLUX")
lYSWIMON0.SWIMONFLUQ = rsADO("SWIMONFLUQ")
lYSWIMON0.SWIMONFLUD = rsADO("SWIMONFLUD")
lYSWIMON0.SWIMONFLUH = rsADO("SWIMONFLUH")
lYSWIMON0.SWIMONFLUS = rsADO("SWIMONFLUS")
lYSWIMON0.SWIMONXMT = rsADO("SWIMONXMT")
lYSWIMON0.SWIMONX20 = rsADO("SWIMONX20")
lYSWIMON0.SWIMONX21 = rsADO("SWIMONX21")
lYSWIMON0.SWIMONX32A = rsADO("SWIMONX32A")
lYSWIMON0.SWIMONX32D = rsADO("SWIMONX32D")
lYSWIMON0.SWIMONX32V = rsADO("SWIMONX32V")
lYSWIMON0.SWIMONUPDS = rsADO("SWIMONUPDS")

Exit Function
Error_Handler:
srvYSWIMON0_GetBuffer_ODBC = Error


End Function

Public Function srvYSWIMON0_Init(lYSWIMON0 As typeYSWIMON0)
lYSWIMON0.SWIMONID = 0
lYSWIMON0.SWISABNUM = 0
lYSWIMON0.SWISABCOP = ""
lYSWIMON0.SWISABDOS = 0
lYSWIMON0.SAAAID = 0
lYSWIMON0.SAAUMIDL = 0
lYSWIMON0.SAAUMIDH = 0
lYSWIMON0.SAAQUEUE = ""
lYSWIMON0.SAAQMOD = 0
lYSWIMON0.SAAQOFAC = 0
lYSWIMON0.SAAUNIT = ""
lYSWIMON0.SWIMONSTA = ""
lYSWIMON0.SWIMONSTAD = 0
lYSWIMON0.SWIMONSTAH = 0
lYSWIMON0.SWIMONFLUX = ""
lYSWIMON0.SWIMONFLUQ = ""
lYSWIMON0.SWIMONFLUD = 0
lYSWIMON0.SWIMONFLUH = 0
lYSWIMON0.SWIMONFLUS = 0
lYSWIMON0.SWIMONXMT = ""
lYSWIMON0.SWIMONX20 = ""
lYSWIMON0.SWIMONX21 = ""
lYSWIMON0.SWIMONX32A = 0
lYSWIMON0.SWIMONX32D = ""
lYSWIMON0.SWIMONX32V = 0
lYSWIMON0.SWIMONUPDS = 0

End Function

Public Sub srvYSWIMON0_fgDisplay(lYSWIMON0 As typeYSWIMON0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 12
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIMONID   9P"
fgDisplay.Col = 1: fgDisplay = "Identification"
fgDisplay.Col = 2: fgDisplay = lYSWIMON0.SWIMONID
End Sub


Public Sub srvSWISABREF(lTRN_SAB As String, lTRN_SAA As String, lUnit As String, lDOS As Long)
Dim wTRN As String
Dim K As Integer, X As String, KScan As Integer


'  pour REMPLACER SAA_from_SAB_TRN(   à faire JPL 2004.10.11


wTRN = Trim(lTRN_SAB) & Space$(20)
lUnit = "": K = 1

Select Case Mid$(wTRN, 1, 2)
    Case "TC": K = 3: lUnit = "BOTC"
    Case "00": K = 3
            Select Case Mid$(wTRN, 3, 3)
                Case "CDE", "CDI", "RDE", "RDI": lUnit = "SOBI"
                Case "CPT", "TRF": lUnit = "ORPA"
            End Select
    Case Else
            Select Case Mid$(wTRN, 1, 3)
                Case "PRE", "EMP", "SWP": lUnit = "BOTC"
            End Select
End Select

X = Mid$(Trim(wTRN), K, 12)
If lUnit <> "" Then
    lTRN_SAA = Trim(lUnit & X)
Else
    lTRN_SAA = X
    KScan = InStr(1, X, "BOTC"): If KScan > 0 Then lUnit = "BOTC": K = KScan + 4
    KScan = InStr(1, X, "ORPA"): If KScan > 0 Then lUnit = "ORPA": K = KScan + 4
    KScan = InStr(1, X, "SOBI"): If KScan > 0 Then lUnit = "SOBI": K = KScan + 4
    KScan = InStr(1, X, "SOBF"): If KScan > 0 Then lUnit = "SOBF": K = KScan + 4
    KScan = InStr(1, X, "DCOM"): If KScan > 0 Then lUnit = "DCOM": K = KScan + 4
End If
If lUnit = "" Then
    Select Case Mid$(X, 1, 3)
       Case "CDE", "CDI", "RDE", "RDI": lUnit = "SOBI"
       Case "CPT", "TRF", "ORP": lUnit = "ORPA"
       Case "PRE", "EMP", "SWP": lUnit = "BOTC"
       Case "CSO", "GSO": lUnit = "CSOP"
    End Select
End If
X = Mid$(Trim(wTRN), K, 12)
lDOS = CLng(Val(X))

End Sub

Public Sub srvYSWIMON0_SendMail(newY As typeYSWIMON0)
Dim wSendMail As typeSendMail
Dim wSAAQUEUE As String
Dim bgColor As String
wSendMail.FromDisplayName = "SAB=>SWIFT"

If newY.SAAUNIT <> "0   " Then
    wSendMail.RecipientDisplayName = newY.SAAUNIT
Else
    wSendMail.RecipientDisplayName = Mid$(newY.SWIMONX20, 1, 4)
End If
wSAAQUEUE = Trim(newY.SAAQUEUE)

bgColor = "YELLOW"
Select Case wSAAQUEUE
    Case "OFCS_Validate": wSendMail.CcDisplayName = "DEON": bgColor = "CYAN"
    Case Is = "": wSAAQUEUE = "???????": bgColor = "MAGENTA"
End Select
wSendMail.Subject = "MT" & newY.SWIMONXMT & " " & Trim(newY.SWIMONX20) & " => " & wSAAQUEUE
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                    & htmlFontColor("BLUE") & "<CENTER>MT" & newY.SWIMONXMT _
                    & "&#160;&#160;&#160;&#160;&#160;" & Format$(newY.SWIMONX32A, "### ### ### ###.00") & " " & newY.SWIMONX32D & "</CENTER>" _
                    & htmlFontColor("BLACK") & "<BR>Le message SWIFT " _
                    & htmlFontColor("RED") & Trim(newY.SWIMONX20) & htmlFontColor("BLACK") & "   est actuellement en file d'attente : " & htmlFontColor("RED") & wSAAQUEUE _
                    & htmlFontColor("BLACK") & "<BR><BR>Le message a été envoyé de SAB vers ALLIANCE le " & htmlFontColor("BLUE") & dateImp10_S(newY.SWIMONFLUD) & " à  " & timeImp8(newY.SWIMONFLUH) _
                    & htmlFontColor("BLACK") & "<BR><BR>SABNUM : " & newY.SWISABNUM  ' & "  _   Statut :" & htmlFontColor("BLUE") & newY.SWIMONSTA

If newY.SWIMONSTA = "S998" Then wSendMail.Message = wSendMail.Message _
    & htmlFontColor("MAGENTA") & "<BR><BR>Ce message n'a pas pu être traité par SWIFT ALLIANCE," _
    & "<BR>ceci est généralement dû à un code BIC erroné ou une erreur de syntaxe." _
    & "<BR>Vérifiez dans 'Message File' les messages de type 'XXX'."
    
wSendMail.Attachment = ""
wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub
