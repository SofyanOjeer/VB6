Attribute VB_Name = "srvYECHIMP0"
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
    ECHIMPAD7      As String * 32         'ADRESSE          '
End Type
Type typeYECHREL0

    ECHRELDVAL    As Long
    ECHRELMDB     As Currency
    ECHRELMCR     As Currency
    ECHRELSD      As Currency
    ECHRELSDS     As String * 1
    ECHRELNBJ     As Long
    ECHRELNBR     As Currency
    ECHRELTAUX    As Double

End Type

'---------------------------------------------------------
Public Function rsYECHIMP0_GetBuffer(rsSab As ADODB.Recordset, rsYECHIMP0 As typeYECHIMP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYECHIMP0_GetBuffer = Null

rsYECHIMP0.ECHIMPJOB = rsSab("ECHIMPJOB")
rsYECHIMP0.ECHIMPJOBS = rsSab("ECHIMPJOBS")
rsYECHIMP0.ECHIMPSEQ = rsSab("ECHIMPSEQ")
rsYECHIMP0.ECHIMPCPT = rsSab("ECHIMPCPT")
rsYECHIMP0.ECHIMPDEV = rsSab("ECHIMPDEV")
rsYECHIMP0.ECHIMPDTRT = rsSab("ECHIMPDTRT")
rsYECHIMP0.ECHIMPDOPE = rsSab("ECHIMPDOPE")
rsYECHIMP0.ECHIMPDDEB = rsSab("ECHIMPDDEB")
rsYECHIMP0.ECHIMPDFIN = rsSab("ECHIMPDFIN")
rsYECHIMP0.ECHIMPIDEM = rsSab("ECHIMPIDEM")
rsYECHIMP0.ECHIMPIDES = rsSab("ECHIMPIDES")
rsYECHIMP0.ECHIMPIDEV = rsSab("ECHIMPIDEV")
rsYECHIMP0.ECHIMPIDET = rsSab("ECHIMPIDET")
rsYECHIMP0.ECHIMPICRM = rsSab("ECHIMPICRM")
rsYECHIMP0.ECHIMPICRS = rsSab("ECHIMPICRS")
rsYECHIMP0.ECHIMPICRV = rsSab("ECHIMPICRV")
rsYECHIMP0.ECHIMPICRT = rsSab("ECHIMPICRT")
rsYECHIMP0.ECHIMPCPFD = rsSab("ECHIMPCPFD")
rsYECHIMP0.ECHIMPCMVT = rsSab("ECHIMPCMVT")
rsYECHIMP0.ECHIMPCCPT = rsSab("ECHIMPCCPT")
rsYECHIMP0.ECHIMPMON = rsSab("ECHIMPMON")
rsYECHIMP0.ECHIMPMONS = rsSab("ECHIMPMONS")
rsYECHIMP0.ECHIMPNREF = rsSab("ECHIMPNREF")
rsYECHIMP0.ECHIMPAD1 = rsSab("ECHIMPAD1")
rsYECHIMP0.ECHIMPAD2 = rsSab("ECHIMPAD2")
rsYECHIMP0.ECHIMPAD3 = rsSab("ECHIMPAD3")
rsYECHIMP0.ECHIMPAD4 = rsSab("ECHIMPAD4")
rsYECHIMP0.ECHIMPAD5 = rsSab("ECHIMPAD5")
rsYECHIMP0.ECHIMPAD6 = rsSab("ECHIMPAD6")
rsYECHIMP0.ECHIMPAD7 = rsSab("ECHIMPAD7")
Exit Function

Error_Handler:

rsYECHIMP0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsYECHIMP0_Init(rsYECHIMP0 As typeYECHIMP0)

rsYECHIMP0.ECHIMPJOB = 0
rsYECHIMP0.ECHIMPJOBS = 0
rsYECHIMP0.ECHIMPSEQ = 0
rsYECHIMP0.ECHIMPCPT = ""
rsYECHIMP0.ECHIMPDEV = ""
rsYECHIMP0.ECHIMPDTRT = 0
rsYECHIMP0.ECHIMPDOPE = 0
rsYECHIMP0.ECHIMPDDEB = 0
rsYECHIMP0.ECHIMPDFIN = 0
rsYECHIMP0.ECHIMPIDEM = 0
rsYECHIMP0.ECHIMPIDES = ""
rsYECHIMP0.ECHIMPIDEV = 0
rsYECHIMP0.ECHIMPIDET = 0
rsYECHIMP0.ECHIMPICRM = 0
rsYECHIMP0.ECHIMPICRS = ""
rsYECHIMP0.ECHIMPICRV = 0
rsYECHIMP0.ECHIMPICRT = 0
rsYECHIMP0.ECHIMPCPFD = 0
rsYECHIMP0.ECHIMPCMVT = 0
rsYECHIMP0.ECHIMPCCPT = 0
rsYECHIMP0.ECHIMPMON = 0
rsYECHIMP0.ECHIMPMONS = ""
rsYECHIMP0.ECHIMPNREF = ""
rsYECHIMP0.ECHIMPAD1 = ""
rsYECHIMP0.ECHIMPAD2 = ""
rsYECHIMP0.ECHIMPAD3 = ""
rsYECHIMP0.ECHIMPAD4 = ""
rsYECHIMP0.ECHIMPAD5 = ""
rsYECHIMP0.ECHIMPAD6 = ""
rsYECHIMP0.ECHIMPAD7 = ""
'---------------------------------------------------------

End Sub

Public Function sqlYECHIMP0_Insert(newY As typeYECHIMP0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYECHIMP0_Insert = Null

xSet = " (ECHIMPJOB"
xValues = " values(" & newY.ECHIMPJOB

' Détecter les modifications
'===================================================================================
If newY.ECHIMPJOBS <> 0 Then xSet = xSet & ",ECHIMPJOBS": xValues = xValues & " ," & newY.ECHIMPJOBS
If newY.ECHIMPSEQ <> 0 Then xSet = xSet & ",ECHIMPSEQ": xValues = xValues & " ," & newY.ECHIMPSEQ
If newY.ECHIMPDTRT <> 0 Then xSet = xSet & ",ECHIMPDTRT": xValues = xValues & " ," & newY.ECHIMPDTRT
If newY.ECHIMPDOPE <> 0 Then xSet = xSet & ",ECHIMPDOPE": xValues = xValues & " ," & newY.ECHIMPDOPE
If newY.ECHIMPDDEB <> 0 Then xSet = xSet & ",ECHIMPDDEB": xValues = xValues & " ," & newY.ECHIMPDDEB
If newY.ECHIMPDFIN <> 0 Then xSet = xSet & ",ECHIMPDFIN": xValues = xValues & " ," & newY.ECHIMPDFIN
If newY.ECHIMPIDEM <> 0 Then xSet = xSet & ",ECHIMPIDEM": xValues = xValues & " ," & cur_P(newY.ECHIMPIDEM)
If newY.ECHIMPIDEV <> 0 Then xSet = xSet & ",ECHIMPIDEV": xValues = xValues & " ," & newY.ECHIMPIDEV
If newY.ECHIMPIDET <> 0 Then xSet = xSet & ",ECHIMPIDET": xValues = xValues & " ," & Replace(CStr(newY.ECHIMPIDET), ",", ".")
If newY.ECHIMPICRM <> 0 Then xSet = xSet & ",ECHIMPICRM": xValues = xValues & " ," & cur_P(newY.ECHIMPICRM)
If newY.ECHIMPICRV <> 0 Then xSet = xSet & ",ECHIMPICRV": xValues = xValues & " ," & newY.ECHIMPICRV
If newY.ECHIMPICRT <> 0 Then xSet = xSet & ",ECHIMPICRT": xValues = xValues & " ," & Replace(CStr(newY.ECHIMPICRT), ",", ".")
If newY.ECHIMPCPFD <> 0 Then xSet = xSet & ",ECHIMPCPFD": xValues = xValues & " ," & cur_P(newY.ECHIMPCPFD)
If newY.ECHIMPCMVT <> 0 Then xSet = xSet & ",ECHIMPCMVT": xValues = xValues & " ," & cur_P(newY.ECHIMPCMVT)
If newY.ECHIMPCCPT <> 0 Then xSet = xSet & ",ECHIMPCCPT": xValues = xValues & " ," & cur_P(newY.ECHIMPCCPT)
If newY.ECHIMPMON <> 0 Then xSet = xSet & ",ECHIMPMON": xValues = xValues & " ," & cur_P(newY.ECHIMPMON)


If Trim(newY.ECHIMPCPT) <> "" Then xSet = xSet & ",ECHIMPCPT": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPCPT), "'", "''") & "'"
If Trim(newY.ECHIMPDEV) <> "" Then xSet = xSet & ",ECHIMPDEV": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPDEV), "'", "''") & "'"
If Trim(newY.ECHIMPIDES) <> "" Then xSet = xSet & ",ECHIMPIDES": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPIDES), "'", "''") & "'"
If Trim(newY.ECHIMPICRS) <> "" Then xSet = xSet & ",ECHIMPICRS": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPICRS), "'", "''") & "'"
If Trim(newY.ECHIMPMONS) <> "" Then xSet = xSet & ",ECHIMPMONS": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPMONS), "'", "''") & "'"
If Trim(newY.ECHIMPNREF) <> "" Then xSet = xSet & ",ECHIMPNREF": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPNREF), "'", "''") & "'"

newY.ECHIMPAD1 = Replace(newY.ECHIMPAD1, "BANQUE BANQUE", "BANQUE")
newY.ECHIMPAD1 = Replace(newY.ECHIMPAD1, "STE SOCIETE", "SOCIETE")

If Trim(newY.ECHIMPAD1) <> "" Then xSet = xSet & ",ECHIMPAD1": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPAD1), "'", "''") & "'"
If Trim(newY.ECHIMPAD2) <> "" Then xSet = xSet & ",ECHIMPAD2": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPAD2), "'", "''") & "'"
If Trim(newY.ECHIMPAD3) <> "" Then xSet = xSet & ",ECHIMPAD3": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPAD3), "'", "''") & "'"
If Trim(newY.ECHIMPAD4) <> "" Then xSet = xSet & ",ECHIMPAD4": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPAD4), "'", "''") & "'"
If Trim(newY.ECHIMPAD5) <> "" Then xSet = xSet & ",ECHIMPAD5": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPAD5), "'", "''") & "'"
If Trim(newY.ECHIMPAD6) <> "" Then xSet = xSet & ",ECHIMPAD6": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPAD6), "'", "''") & "'"
If Trim(newY.ECHIMPAD7) <> "" Then xSet = xSet & ",ECHIMPAD7": xValues = xValues & " ,'" & Replace(Trim(newY.ECHIMPAD7), "'", "''") & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YECHIMP0" & xSet & ")" & xValues & ")"

Set rsSab = cnsab.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYECHIMP0_Insert = "Erreur màj : " & newY.ECHIMPJOB
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYECHIMP0_Insert = Error
End Function

