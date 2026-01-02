Attribute VB_Name = "srvYCDOSWI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOSWI0Len = 589 ' 34 + 555
Public Const recYCDOSWI0_Block = 200
Public Const constYCDOSWI0 = "YCDOSWI0"
Dim meYbase As typeYBase
Dim paramYCDOSWI0_Import As String

Type typeYCDOSWI0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CDOSWIETB       As Integer                        ' CODE ETABLISSEMENT
    CDOSWIAGE       As Integer                        ' AGENCE
    CDOSWISER       As String * 2                     ' SERVICE
    CDOSWISSE       As String * 2                     ' SOUS-SERVICE
    CDOSWICOP       As String * 3                     ' CODE OPERATION
    CDOSWIDOS       As Long                           ' NUMERO DOSSIER
    CDOSWINUR       As Long                           ' N° RENOUVELLEMENT
    CDOSWIUTI       As Long                           ' N° UTILISATION
    CDOSWIPAI       As Long                           ' N° PAIEMENT
    CDOSWIREG       As Long                           ' N° REGLEMENT/ENCAIS
    CDOSWIBER       As String * 1                     ' BENEFICIAIR CLI/TIE
    CDOSWIBEN       As String * 7                     ' BENEFICIAIRE EXPORT
    CDOSWIBAR       As String * 1                     ' BANQU.BENEF.CLI/TIE
    CDOSWIBAB       As String * 7                     ' BANQUE BENEF
    CDOSWIBDE       As String * 12                    ' BIC BQDES
    CDOSWIBIN       As String * 12                    ' BIC BQINT
    CDOSWIBBD       As String * 12                    ' BIC BQBAD
    CDOSWIBBE       As String * 12                    ' BIC BQBEN
    CDOSWIBBA       As String * 12                    ' BIC BQBAN
    CDOSWIDDR       As Long                           ' DT DEM RBT
    CDOSWIDAV       As Long                           ' DT AVIS PAIE
    CDOSWILI1       As String * 79                    ' LIBEL AVI
    CDOSWILI2       As String * 79                    ' LIBEL AVI
    CDOSWILI3       As String * 79                    ' LIBEL AVI
    CDOSWILI4       As String * 79                    ' LIBEL AVI
    CDOSWIIBD       As String * 34                    ' IBAN BQ EMETT/DESTI
    CDOSWIIBB       As String * 34                    ' IBAN BQ BENEF
    CDOSWICBE       As String * 1                     ' CODE IBAN BENEF
    CDOSWIIBE       As String * 34                    ' IBAN BENEF.
    CDOSWICHA       As String * 1                     ' CHARGES O/B/S
End Type
    
'---------------------------------------------------------
Public Function srvYCDOSWI0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOSWI0 As typeYCDOSWI0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOSWI0_GetBuffer_ODBC = Null

    recYCDOSWI0.CDOSWIETB = rsADO("CDOSWIETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOSWI0.CDOSWIAGE = rsADO("CDOSWIAGE")    'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOSWI0.CDOSWISER = rsADO("CDOSWISER")    'mId$(MsgTxt, K + 11, 2)
    recYCDOSWI0.CDOSWISSE = rsADO("CDOSWISSE")    'mId$(MsgTxt, K + 13, 2)
    recYCDOSWI0.CDOSWICOP = rsADO("CDOSWICOP")    'mId$(MsgTxt, K + 15, 3)
    recYCDOSWI0.CDOSWIDOS = rsADO("CDOSWIDOS")    'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOSWI0.CDOSWINUR = rsADO("CDOSWINUR")    'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOSWI0.CDOSWIUTI = rsADO("CDOSWIUTI")    'CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOSWI0.CDOSWIPAI = rsADO("CDOSWIPAI")    'CLng(Val(mId$(MsgTxt, K + 38, 2)))
    recYCDOSWI0.CDOSWIREG = rsADO("CDOSWIREG")    'CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOSWI0.CDOSWIBER = rsADO("CDOSWIBER")    'mId$(MsgTxt, K + 44, 1)
    recYCDOSWI0.CDOSWIBEN = rsADO("CDOSWIBEN")    'mId$(MsgTxt, K + 45, 7)
    recYCDOSWI0.CDOSWIBAR = rsADO("CDOSWIBAR")    'mId$(MsgTxt, K + 52, 1)
    recYCDOSWI0.CDOSWIBAB = rsADO("CDOSWIBAB")    'mId$(MsgTxt, K + 53, 7)
    recYCDOSWI0.CDOSWIBDE = rsADO("CDOSWIBDE")    'mId$(MsgTxt, K + 60, 12)
    recYCDOSWI0.CDOSWIBIN = rsADO("CDOSWIBIN")    'mId$(MsgTxt, K + 72, 12)
    recYCDOSWI0.CDOSWIBBD = rsADO("CDOSWIBBD")    'mId$(MsgTxt, K + 84, 12)
    recYCDOSWI0.CDOSWIBBE = rsADO("CDOSWIBBE")    'mId$(MsgTxt, K + 96, 12)
    recYCDOSWI0.CDOSWIBBA = rsADO("CDOSWIBBA")    'mId$(MsgTxt, K + 108, 12)
    recYCDOSWI0.CDOSWIDDR = rsADO("CDOSWIDDR")    'CLng(Val(mId$(MsgTxt, K + 120, 8)))
    recYCDOSWI0.CDOSWIDAV = rsADO("CDOSWIDAV")    'CLng(Val(mId$(MsgTxt, K + 128, 8)))
    recYCDOSWI0.CDOSWILI1 = rsADO("CDOSWILI1")    'mId$(MsgTxt, K + 136, 79)
    recYCDOSWI0.CDOSWILI2 = rsADO("CDOSWILI2")    'mId$(MsgTxt, K + 215, 79)
    recYCDOSWI0.CDOSWILI3 = rsADO("CDOSWILI3")    'mId$(MsgTxt, K + 294, 79)
    recYCDOSWI0.CDOSWILI4 = rsADO("CDOSWILI4")    'mId$(MsgTxt, K + 373, 79)
    recYCDOSWI0.CDOSWIIBD = rsADO("CDOSWIIBD")    'mId$(MsgTxt, K + 452, 34)
    recYCDOSWI0.CDOSWIIBB = rsADO("CDOSWIIBB")    'mId$(MsgTxt, K + 486, 34)
    recYCDOSWI0.CDOSWICBE = rsADO("CDOSWICBE")    'mId$(MsgTxt, K + 520, 1)
    recYCDOSWI0.CDOSWIIBE = rsADO("CDOSWIIBE")    'mId$(MsgTxt, K + 521, 34)
    recYCDOSWI0.CDOSWICHA = rsADO("CDOSWICHA")    'mId$(MsgTxt, K + 555, 1)

Exit Function

Error_Handler:
srvYCDOSWI0_GetBuffer_ODBC = Error

End Function

Public Function srvYCDOSWI0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOSWI0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOSWI0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDOSWI0_Import = "?"

paramYCDOSWI0_Import = paramYBase_DataF & Trim(constYCDOSWI0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOSWI0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOSWI0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOSWI0
            meYbase.K1 = mId$(xIn, 15, 29) 'recYCDOSWI0.CDODOSCOP & recYCDOSWI0.CDODOSDOS .........
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOSWI0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOSWI0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOSWI0_Import" & xIn, vbCritical, Error
Close

srvYCDOSWI0_Import = Error
End Function

Public Function srvYCDOSWI0_Import_Read(lId As String, lYCDOSWI0 As typeYCDOSWI0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDOSWI0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOSWI0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOSWI0_GetBuffer lYCDOSWI0
    srvYCDOSWI0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOSWI0_Import_Read" & xIn, vbCritical, Error
srvYCDOSWI0_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYCDOSWI0_Monitor(recYCDOSWI0 As typeYCDOSWI0)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDOSWI0.Method), 1, 4)
    Case "Seek"
                srvYCDOSWI0_Monitor = srvYCDOSWI0_Seek(recYCDOSWI0)
    Case Else
                recYCDOSWI0.Err = recYCDOSWI0.Method
                Call srvYCDOSWI0_Error(recYCDOSWI0)
                srvYCDOSWI0_Monitor = recYCDOSWI0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDOSWI0_Error(recYCDOSWI0 As typeYCDOSWI0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOSWI0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOSWI0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOSWI0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDOSWI0s.bas  ( " _
                & Trim(recYCDOSWI0.obj) & " : " & Trim(recYCDOSWI0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDOSWI0_GetBuffer(recYCDOSWI0 As typeYCDOSWI0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOSWI0_GetBuffer = Null
recYCDOSWI0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOSWI0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOSWI0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOSWI0.Err = Space$(10) Then

    recYCDOSWI0.CDOSWIETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOSWI0.CDOSWIAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOSWI0.CDOSWISER = mId$(MsgTxt, K + 11, 2)
    recYCDOSWI0.CDOSWISSE = mId$(MsgTxt, K + 13, 2)
    recYCDOSWI0.CDOSWICOP = mId$(MsgTxt, K + 15, 3)
    recYCDOSWI0.CDOSWIDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOSWI0.CDOSWINUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOSWI0.CDOSWIUTI = CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOSWI0.CDOSWIPAI = CLng(Val(mId$(MsgTxt, K + 38, 2)))
    recYCDOSWI0.CDOSWIREG = CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOSWI0.CDOSWIBER = mId$(MsgTxt, K + 44, 1)
    recYCDOSWI0.CDOSWIBEN = mId$(MsgTxt, K + 45, 7)
    recYCDOSWI0.CDOSWIBAR = mId$(MsgTxt, K + 52, 1)
    recYCDOSWI0.CDOSWIBAB = mId$(MsgTxt, K + 53, 7)
    recYCDOSWI0.CDOSWIBDE = mId$(MsgTxt, K + 60, 12)
    recYCDOSWI0.CDOSWIBIN = mId$(MsgTxt, K + 72, 12)
    recYCDOSWI0.CDOSWIBBD = mId$(MsgTxt, K + 84, 12)
    recYCDOSWI0.CDOSWIBBE = mId$(MsgTxt, K + 96, 12)
    recYCDOSWI0.CDOSWIBBA = mId$(MsgTxt, K + 108, 12)
    recYCDOSWI0.CDOSWIDDR = CLng(Val(mId$(MsgTxt, K + 120, 8)))
    recYCDOSWI0.CDOSWIDAV = CLng(Val(mId$(MsgTxt, K + 128, 8)))
    recYCDOSWI0.CDOSWILI1 = mId$(MsgTxt, K + 136, 79)
    recYCDOSWI0.CDOSWILI2 = mId$(MsgTxt, K + 215, 79)
    recYCDOSWI0.CDOSWILI3 = mId$(MsgTxt, K + 294, 79)
    recYCDOSWI0.CDOSWILI4 = mId$(MsgTxt, K + 373, 79)
    recYCDOSWI0.CDOSWIIBD = mId$(MsgTxt, K + 452, 34)
    recYCDOSWI0.CDOSWIIBB = mId$(MsgTxt, K + 486, 34)
    recYCDOSWI0.CDOSWICBE = mId$(MsgTxt, K + 520, 1)
    recYCDOSWI0.CDOSWIIBE = mId$(MsgTxt, K + 521, 34)
    recYCDOSWI0.CDOSWICHA = mId$(MsgTxt, K + 555, 1)

Else
    srvYCDOSWI0_GetBuffer = recYCDOSWI0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOSWI0Len

End Function

'---------------------------------------------------------
Private Sub srvYCDOSWI0_PutBuffer(recYCDOSWI0 As typeYCDOSWI0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDOSWI0Len) = Space$(recYCDOSWI0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOSWI0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOSWI0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOSWI0.CDOSWIETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDOSWI0.CDOSWIAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDOSWI0.CDOSWISER
    Mid$(MsgTxt, K + 13, 2) = recYCDOSWI0.CDOSWISSE
    Mid$(MsgTxt, K + 15, 3) = recYCDOSWI0.CDOSWICOP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDOSWI0.CDOSWIDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDOSWI0.CDOSWINUR, "000 ")
    Mid$(MsgTxt, K + 32, 6) = Format$(recYCDOSWI0.CDOSWIUTI, "00000 ")
    Mid$(MsgTxt, K + 38, 2) = Format$(recYCDOSWI0.CDOSWIPAI, "0 ")
    Mid$(MsgTxt, K + 40, 4) = Format$(recYCDOSWI0.CDOSWIREG, "000 ")
    Mid$(MsgTxt, K + 44, 1) = recYCDOSWI0.CDOSWIBER
    Mid$(MsgTxt, K + 45, 7) = recYCDOSWI0.CDOSWIBEN
    Mid$(MsgTxt, K + 52, 1) = recYCDOSWI0.CDOSWIBAR
    Mid$(MsgTxt, K + 53, 7) = recYCDOSWI0.CDOSWIBAB
    Mid$(MsgTxt, K + 60, 12) = recYCDOSWI0.CDOSWIBDE
    Mid$(MsgTxt, K + 72, 12) = recYCDOSWI0.CDOSWIBIN
    Mid$(MsgTxt, K + 84, 12) = recYCDOSWI0.CDOSWIBBD
    Mid$(MsgTxt, K + 96, 12) = recYCDOSWI0.CDOSWIBBE
    Mid$(MsgTxt, K + 108, 12) = recYCDOSWI0.CDOSWIBBA
    Mid$(MsgTxt, K + 120, 8) = Format$(recYCDOSWI0.CDOSWIDDR, "0000000 ")
    Mid$(MsgTxt, K + 128, 8) = Format$(recYCDOSWI0.CDOSWIDAV, "0000000 ")
    Mid$(MsgTxt, K + 136, 79) = recYCDOSWI0.CDOSWILI1
    Mid$(MsgTxt, K + 215, 79) = recYCDOSWI0.CDOSWILI2
    Mid$(MsgTxt, K + 294, 79) = recYCDOSWI0.CDOSWILI3
    Mid$(MsgTxt, K + 373, 79) = recYCDOSWI0.CDOSWILI4
    Mid$(MsgTxt, K + 452, 34) = recYCDOSWI0.CDOSWIIBD
    Mid$(MsgTxt, K + 486, 34) = recYCDOSWI0.CDOSWIIBB
    Mid$(MsgTxt, K + 520, 1) = recYCDOSWI0.CDOSWICBE
    Mid$(MsgTxt, K + 521, 34) = recYCDOSWI0.CDOSWIIBE
    Mid$(MsgTxt, K + 555, 1) = recYCDOSWI0.CDOSWICHA


End Sub


Public Sub srvYCDOSWI0_ElpDisplay(recYCDOSWI0 As typeYCDOSWI0)
frmElpDisplay.fgData.Rows = 31
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWISER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWISER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWISSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWISSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWICOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWICOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWINUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWINUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIUTI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° UTILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIUTI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIPAI    1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° PAIEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIPAI
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIREG    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° REGLEMENT/ENCAIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIREG
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIR CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBER
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBEN    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIRE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBEN
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBAR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BANQU.BENEF.CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBAR
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBAB    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BANQUE BENEF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBAB
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBDE   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BIC BQDES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBDE
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBIN   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BIC BQINT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBIN
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBBD   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BIC BQBAD"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBBD
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBBE   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BIC BQBEN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBBE
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIBBA   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BIC BQBAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIBBA
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIDDR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DT DEM RBT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIDDR
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIDAV    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DT AVIS PAIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIDAV
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWILI1   79A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL AVI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWILI1
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWILI2   79A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL AVI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWILI2
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWILI3   79A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL AVI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWILI3
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWILI4   79A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL AVI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWILI4
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIIBD   34A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "IBAN BQ EMETT/DESTI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIIBD
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIIBB   34A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "IBAN BQ BENEF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIIBB
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWICBE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE IBAN BENEF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWICBE
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWIIBE   34A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "IBAN BENEF."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWIIBE
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOSWICHA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CHARGES O/B/S"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOSWI0.CDOSWICHA
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvYCDOSWI0_Seek(recYCDOSWI0 As typeYCDOSWI0)
'---------------------------------------------------------

srvYCDOSWI0_Seek = "?"
MsgTxtLen = 0
Call srvYCDOSWI0_PutBuffer(recYCDOSWI0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDOSWI0_GetBuffer(recYCDOSWI0)) Then
            srvYCDOSWI0_Seek = Null
        Else
            Call srvYCDOSWI0_Error(recYCDOSWI0)
        End If
    End If
End If

End Function
'-----------------------------------------------------
Function srvYCDOSWI0_Update(recYCDOSWI0 As typeYCDOSWI0)
'-----------------------------------------------------

srvYCDOSWI0_Update = "?"

MsgTxtLen = 0
Call srvYCDOSWI0_PutBuffer(recYCDOSWI0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOSWI0_GetBuffer(recYCDOSWI0)) Then
        Call srvYCDOSWI0_Error(recYCDOSWI0)
        srvYCDOSWI0_Update = recYCDOSWI0.Err
        Exit Function
    Else
        srvYCDOSWI0_Update = Null
    End If
Else
    recYCDOSWI0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDOSWI0_Init(recYCDOSWI0 As typeYCDOSWI0)
'---------------------------------------------------------
MsgTxt = Space$(recYCDOSWI0Len)
MsgTxtIndex = 0
Call srvYCDOSWI0_GetBuffer(recYCDOSWI0)
recYCDOSWI0.obj = "ZCDODOS0_S"
recYCDOSWI0.CDOSWIETB = 0      'As Integer                        ' CODE ETABLISSEMENT
recYCDOSWI0.CDOSWIAGE = 0      'As Integer                        ' AGENCE
recYCDOSWI0.CDOSWISER = ""      'As String * 2                     ' SERVICE
recYCDOSWI0.CDOSWISSE = ""      'As String * 2                     ' SOUS-SERVICE
recYCDOSWI0.CDOSWICOP = ""      'As String * 3                     ' CODE OPERATION
recYCDOSWI0.CDOSWIDOS = 0      'As Long                           ' NUMERO DOSSIER
recYCDOSWI0.CDOSWINUR = 0      'As Long                           ' N° RENOUVELLEMENT
recYCDOSWI0.CDOSWIUTI = 0      'As Long                           ' N° UTILISATION
recYCDOSWI0.CDOSWIPAI = 0      'As Long                           ' N° PAIEMENT
recYCDOSWI0.CDOSWIREG = 0      'As Long                           ' N° REGLEMENT/ENCAIS
recYCDOSWI0.CDOSWIBER = ""      'As String * 1                     ' BENEFICIAIR CLI/TIE
recYCDOSWI0.CDOSWIBEN = ""      'As String * 7                     ' BENEFICIAIRE EXPORT
recYCDOSWI0.CDOSWIBAR = ""      'As String * 1                     ' BANQU.BENEF.CLI/TIE
recYCDOSWI0.CDOSWIBAB = ""      'As String * 7                     ' BANQUE BENEF
recYCDOSWI0.CDOSWIBDE = ""      'As String * 12                    ' BIC BQDES
recYCDOSWI0.CDOSWIBIN = ""      'As String * 12                    ' BIC BQINT
recYCDOSWI0.CDOSWIBBD = ""      'As String * 12                    ' BIC BQBAD
recYCDOSWI0.CDOSWIBBE = ""      'As String * 12                    ' BIC BQBEN
recYCDOSWI0.CDOSWIBBA = ""      'As String * 12                    ' BIC BQBAN
recYCDOSWI0.CDOSWIDDR = 0      'As Long                           ' DT DEM RBT
recYCDOSWI0.CDOSWIDAV = 0      'As Long                           ' DT AVIS PAIE
recYCDOSWI0.CDOSWILI1 = ""      'As String * 79                    ' LIBEL AVI
recYCDOSWI0.CDOSWILI2 = ""      'As String * 79                    ' LIBEL AVI
recYCDOSWI0.CDOSWILI3 = ""      'As String * 79                    ' LIBEL AVI
recYCDOSWI0.CDOSWILI4 = ""      'As String * 79                    ' LIBEL AVI
recYCDOSWI0.CDOSWIIBD = ""      'As String * 34                    ' IBAN BQ EMETT/DESTI
recYCDOSWI0.CDOSWIIBB = ""      'As String * 34                    ' IBAN BQ BENEF
recYCDOSWI0.CDOSWICBE = ""      'As String * 1                     ' CODE IBAN BENEF
recYCDOSWI0.CDOSWIIBE = ""      'As String * 34                    ' IBAN BENEF.
recYCDOSWI0.CDOSWICHA = ""      'As String * 1                     ' CHARGES O/B/S
End Sub







