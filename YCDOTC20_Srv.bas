Attribute VB_Name = "srvYCDOTC20"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOTC20Len = 295 ' 34 + 261
Public Const recYCDOTC20_Block = 100
Public Const constYCDOTC20 = "YCDOTC20"
Dim meYbase As typeYBase
Dim paramYCDOTC20_Import As String

Type typeYCDOTC20
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CDOTC2ETB       As Integer                        ' CODE ETABLISSEMENT
    CDOTC2AGE       As Integer                        ' AGENCE
    CDOTC2SER       As String * 2                     ' SERVICE
    CDOTC2SSE       As String * 2                     ' SOUS-SERVICE
    CDOTC2COP       As String * 3                     ' CODE OPERATION
    CDOTC2DOS       As Long                           ' NUMERO DOSSIER
    CDOTC2NUR       As Long                           ' N° RENOUVELLEMENT
    CDOTC2UTI       As Long                           ' N° UTILILSAT°./MODIF
    CDOTC2EVE       As String * 2                     ' EVENEMENT
    CDOTC2SEQ       As Long                           ' N° SEQUENCE
    CDOTC2COM       As String * 6                     ' Code commission
    CDOTC2DEV       As String * 3                     ' Devise
    CDOTC2CAT       As String * 3                     ' Catégorie client
    CDOTC2CLI       As String * 7                     ' N° Client
    CDOTC2DEB       As Long                           ' Date début effet
    CDOTC2FIN       As Long                           ' Date fin effet
    CDOTC2TVA       As String * 1                     ' TVA (O/N)
    CDOTC2PER       As String * 1                     ' Périodicité
    CDOTC2CUM       As String * 1                     ' Cumulable (O/N)
    CDOTC2MTF       As Currency                       ' Montant fixe
    CDOTC2IND       As String * 1                     ' Indivisibilité (O/N)
    CDOTC2AVE       As String * 1                     ' Avis à échéance
    CDOTC2MT1       As Long                           ' Montant tranche 1
    CDOTC2MT2       As Long                           ' Montant tranche 2
    CDOTC2MT3       As Long                           ' Montant tranche 3
    CDOTC2MT4       As Long                           ' Montant tranche 4
    CDOTC2MT5       As Long                           ' Montant tranche 5
    CDOTC2MT6       As Long                           ' Montant tranche 6
    CDOTC2TX1       As Double                         ' Taux tranche 1
    CDOTC2TX2       As Double                         ' Taux tranche 2
    CDOTC2TX3       As Double                         ' Taux tranche 3
    CDOTC2TX4       As Double                         ' Taux tranche 4
    CDOTC2TX5       As Double                         ' Taux tranche 5
    CDOTC2TX6       As Double                         ' Taux tranche 6
End Type
    
'---------------------------------------------------------
Public Function srvYCDOTC20_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOTC20 As typeYCDOTC20)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOTC20_GetBuffer_ODBC = Null

    recYCDOTC20.CDOTC2ETB = rsADO("CDOTC2ETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOTC20.CDOTC2AGE = rsADO("CDOTC2AGE")    'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOTC20.CDOTC2SER = rsADO("CDOTC2SER")    'mId$(MsgTxt, K + 11, 2)
    recYCDOTC20.CDOTC2SSE = rsADO("CDOTC2SSE")    'mId$(MsgTxt, K + 13, 2)
    recYCDOTC20.CDOTC2COP = rsADO("CDOTC2COP")    'mId$(MsgTxt, K + 15, 3)
    recYCDOTC20.CDOTC2DOS = rsADO("CDOTC2DOS")    'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOTC20.CDOTC2NUR = rsADO("CDOTC2NUR")    'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOTC20.CDOTC2UTI = rsADO("CDOTC2UTI")    'CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOTC20.CDOTC2EVE = rsADO("CDOTC2EVE")    'mId$(MsgTxt, K + 38, 2)
    recYCDOTC20.CDOTC2SEQ = rsADO("CDOTC2SEQ")    'CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOTC20.CDOTC2COM = rsADO("CDOTC2COM")    'mId$(MsgTxt, K + 44, 6)
    recYCDOTC20.CDOTC2DEV = rsADO("CDOTC2DEV")    'mId$(MsgTxt, K + 50, 3)
    recYCDOTC20.CDOTC2CAT = rsADO("CDOTC2CAT")    'mId$(MsgTxt, K + 53, 3)
    recYCDOTC20.CDOTC2CLI = rsADO("CDOTC2CLI")    'mId$(MsgTxt, K + 56, 7)
    recYCDOTC20.CDOTC2DEB = rsADO("CDOTC2DEB")    'CLng(Val(mId$(MsgTxt, K + 63, 8)))
    recYCDOTC20.CDOTC2FIN = rsADO("CDOTC2FIN")    'CLng(Val(mId$(MsgTxt, K + 71, 8)))
    recYCDOTC20.CDOTC2TVA = rsADO("CDOTC2TVA")    'mId$(MsgTxt, K + 79, 1)
    recYCDOTC20.CDOTC2PER = rsADO("CDOTC2PER")    'mId$(MsgTxt, K + 80, 1)
    recYCDOTC20.CDOTC2CUM = rsADO("CDOTC2CUM")    'mId$(MsgTxt, K + 81, 1)
    recYCDOTC20.CDOTC2MTF = rsADO("CDOTC2MTF")    'CCur(Val(mId$(MsgTxt, K + 82, 16))) / 100
    recYCDOTC20.CDOTC2IND = rsADO("CDOTC2IND")    'mId$(MsgTxt, K + 98, 1)
    recYCDOTC20.CDOTC2AVE = rsADO("CDOTC2AVE")    'mId$(MsgTxt, K + 99, 1)
    recYCDOTC20.CDOTC2MT1 = rsADO("CDOTC2MT1")    'CLng(Val(mId$(MsgTxt, K + 100, 14)))
    recYCDOTC20.CDOTC2MT2 = rsADO("CDOTC2MT2")    'CLng(Val(mId$(MsgTxt, K + 114, 14)))
    recYCDOTC20.CDOTC2MT3 = rsADO("CDOTC2MT3")    'CLng(Val(mId$(MsgTxt, K + 128, 14)))
    recYCDOTC20.CDOTC2MT4 = rsADO("CDOTC2MT4")    'CLng(Val(mId$(MsgTxt, K + 142, 14)))
    recYCDOTC20.CDOTC2MT5 = rsADO("CDOTC2MT5")    'CLng(Val(mId$(MsgTxt, K + 156, 14)))
    recYCDOTC20.CDOTC2MT6 = rsADO("CDOTC2MT6")    'CLng(Val(mId$(MsgTxt, K + 170, 14)))
    recYCDOTC20.CDOTC2TX1 = rsADO("CDOTC2TX1")    'CDbl(Val(mId$(MsgTxt, K + 184, 13))) / 1000000
    recYCDOTC20.CDOTC2TX2 = rsADO("CDOTC2TX2")    'CDbl(Val(mId$(MsgTxt, K + 197, 13))) / 1000000
    recYCDOTC20.CDOTC2TX3 = rsADO("CDOTC2TX3")    'CDbl(Val(mId$(MsgTxt, K + 210, 13))) / 1000000
    recYCDOTC20.CDOTC2TX4 = rsADO("CDOTC2TX4")    'CDbl(Val(mId$(MsgTxt, K + 223, 13))) / 1000000
    recYCDOTC20.CDOTC2TX5 = rsADO("CDOTC2TX5")    'CDbl(Val(mId$(MsgTxt, K + 236, 13))) / 1000000
    recYCDOTC20.CDOTC2TX6 = rsADO("CDOTC2TX6")    'CDbl(Val(mId$(MsgTxt, K + 249, 13))) / 1000000

Exit Function

Error_Handler:
srvYCDOTC20_GetBuffer_ODBC = Error

End Function

Public Function srvYCDOTC20_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOTC20
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOTC20_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDOTC20_Import = "?"

paramYCDOTC20_Import = paramYBase_DataF & Trim(constYCDOTC20) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOTC20_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOTC20) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOTC20
            meYbase.K1 = mId$(xIn, 15, 29) 'recYCDOTC20.CDODOSCOP & recYCDOTC20.CDODOSDOS  .......
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOTC20_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOTC20
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOTC20_Import" & xIn, vbCritical, Error
Close

srvYCDOTC20_Import = Error
End Function

Public Function srvYCDOTC20_Import_Read(lId As String, lYCDOTC20 As typeYCDOTC20)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDOTC20_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOTC20
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOTC20_GetBuffer lYCDOTC20
    srvYCDOTC20_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOTC20_Import_Read" & xIn, vbCritical, Error
srvYCDOTC20_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYCDOTC20_Monitor(recYCDOTC20 As typeYCDOTC20)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDOTC20.Method), 1, 4)
    Case "Seek"
                srvYCDOTC20_Monitor = srvYCDOTC20_Seek(recYCDOTC20)
    Case Else
                recYCDOTC20.Err = recYCDOTC20.Method
                Call srvYCDOTC20_Error(recYCDOTC20)
                srvYCDOTC20_Monitor = recYCDOTC20.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDOTC20_Error(recYCDOTC20 As typeYCDOTC20)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOTC20" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOTC20.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOTC20.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDOTC20s.bas  ( " _
                & Trim(recYCDOTC20.Obj) & " : " & Trim(recYCDOTC20.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDOTC20_GetBuffer(recYCDOTC20 As typeYCDOTC20)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOTC20_GetBuffer = Null
recYCDOTC20.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOTC20.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOTC20.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOTC20.Err = Space$(10) Then
    recYCDOTC20.CDOTC2ETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOTC20.CDOTC2AGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOTC20.CDOTC2SER = mId$(MsgTxt, K + 11, 2)
    recYCDOTC20.CDOTC2SSE = mId$(MsgTxt, K + 13, 2)
    recYCDOTC20.CDOTC2COP = mId$(MsgTxt, K + 15, 3)
    recYCDOTC20.CDOTC2DOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOTC20.CDOTC2NUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOTC20.CDOTC2UTI = CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOTC20.CDOTC2EVE = mId$(MsgTxt, K + 38, 2)
    recYCDOTC20.CDOTC2SEQ = CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOTC20.CDOTC2COM = mId$(MsgTxt, K + 44, 6)
    recYCDOTC20.CDOTC2DEV = mId$(MsgTxt, K + 50, 3)
    recYCDOTC20.CDOTC2CAT = mId$(MsgTxt, K + 53, 3)
    recYCDOTC20.CDOTC2CLI = mId$(MsgTxt, K + 56, 7)
    recYCDOTC20.CDOTC2DEB = CLng(Val(mId$(MsgTxt, K + 63, 8)))
    recYCDOTC20.CDOTC2FIN = CLng(Val(mId$(MsgTxt, K + 71, 8)))
    recYCDOTC20.CDOTC2TVA = mId$(MsgTxt, K + 79, 1)
    recYCDOTC20.CDOTC2PER = mId$(MsgTxt, K + 80, 1)
    recYCDOTC20.CDOTC2CUM = mId$(MsgTxt, K + 81, 1)
    recYCDOTC20.CDOTC2MTF = CCur(Val(mId$(MsgTxt, K + 82, 16))) / 100
    recYCDOTC20.CDOTC2IND = mId$(MsgTxt, K + 98, 1)
    recYCDOTC20.CDOTC2AVE = mId$(MsgTxt, K + 99, 1)
    recYCDOTC20.CDOTC2MT1 = CLng(Val(mId$(MsgTxt, K + 100, 14)))
    recYCDOTC20.CDOTC2MT2 = CLng(Val(mId$(MsgTxt, K + 114, 14)))
    recYCDOTC20.CDOTC2MT3 = CLng(Val(mId$(MsgTxt, K + 128, 14)))
    recYCDOTC20.CDOTC2MT4 = CLng(Val(mId$(MsgTxt, K + 142, 14)))
    recYCDOTC20.CDOTC2MT5 = CLng(Val(mId$(MsgTxt, K + 156, 14)))
    recYCDOTC20.CDOTC2MT6 = CLng(Val(mId$(MsgTxt, K + 170, 14)))
    recYCDOTC20.CDOTC2TX1 = CDbl(Val(mId$(MsgTxt, K + 184, 13))) / 1000000
    recYCDOTC20.CDOTC2TX2 = CDbl(Val(mId$(MsgTxt, K + 197, 13))) / 1000000
    recYCDOTC20.CDOTC2TX3 = CDbl(Val(mId$(MsgTxt, K + 210, 13))) / 1000000
    recYCDOTC20.CDOTC2TX4 = CDbl(Val(mId$(MsgTxt, K + 223, 13))) / 1000000
    recYCDOTC20.CDOTC2TX5 = CDbl(Val(mId$(MsgTxt, K + 236, 13))) / 1000000
    recYCDOTC20.CDOTC2TX6 = CDbl(Val(mId$(MsgTxt, K + 249, 13))) / 1000000
Else
    srvYCDOTC20_GetBuffer = recYCDOTC20.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOTC20Len

End Function

'---------------------------------------------------------
Private Sub srvYCDOTC20_PutBuffer(recYCDOTC20 As typeYCDOTC20)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDOTC20Len) = Space$(recYCDOTC20Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOTC20.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOTC20.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOTC20.CDOTC2ETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDOTC20.CDOTC2AGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDOTC20.CDOTC2SER
    Mid$(MsgTxt, K + 13, 2) = recYCDOTC20.CDOTC2SSE
    Mid$(MsgTxt, K + 15, 3) = recYCDOTC20.CDOTC2COP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDOTC20.CDOTC2DOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDOTC20.CDOTC2NUR, "000 ")
    Mid$(MsgTxt, K + 32, 6) = Format$(recYCDOTC20.CDOTC2UTI, "00000 ")
    Mid$(MsgTxt, K + 38, 2) = recYCDOTC20.CDOTC2EVE
    Mid$(MsgTxt, K + 40, 4) = Format$(recYCDOTC20.CDOTC2SEQ, "000 ")
    Mid$(MsgTxt, K + 44, 6) = recYCDOTC20.CDOTC2COM
    Mid$(MsgTxt, K + 50, 3) = recYCDOTC20.CDOTC2DEV
    Mid$(MsgTxt, K + 53, 3) = recYCDOTC20.CDOTC2CAT
    Mid$(MsgTxt, K + 56, 7) = recYCDOTC20.CDOTC2CLI
    Mid$(MsgTxt, K + 63, 8) = Format$(recYCDOTC20.CDOTC2DEB, "0000000 ")
    Mid$(MsgTxt, K + 71, 8) = Format$(recYCDOTC20.CDOTC2FIN, "0000000 ")
    Mid$(MsgTxt, K + 79, 1) = recYCDOTC20.CDOTC2TVA
    Mid$(MsgTxt, K + 80, 1) = recYCDOTC20.CDOTC2PER
    Mid$(MsgTxt, K + 81, 1) = recYCDOTC20.CDOTC2CUM
    Mid$(MsgTxt, K + 82, 16) = Format$(recYCDOTC20.CDOTC2MTF * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 98, 1) = recYCDOTC20.CDOTC2IND
    Mid$(MsgTxt, K + 99, 1) = recYCDOTC20.CDOTC2AVE
    Mid$(MsgTxt, K + 100, 14) = Format$(recYCDOTC20.CDOTC2MT1, "0000000000000 ")
    Mid$(MsgTxt, K + 114, 14) = Format$(recYCDOTC20.CDOTC2MT2, "0000000000000 ")
    Mid$(MsgTxt, K + 128, 14) = Format$(recYCDOTC20.CDOTC2MT3, "0000000000000 ")
    Mid$(MsgTxt, K + 142, 14) = Format$(recYCDOTC20.CDOTC2MT4, "0000000000000 ")
    Mid$(MsgTxt, K + 156, 14) = Format$(recYCDOTC20.CDOTC2MT5, "0000000000000 ")
    Mid$(MsgTxt, K + 170, 14) = Format$(recYCDOTC20.CDOTC2MT6, "0000000000000 ")
    Mid$(MsgTxt, K + 184, 13) = Format$(recYCDOTC20.CDOTC2TX1 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 197, 13) = Format$(recYCDOTC20.CDOTC2TX2 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 210, 13) = Format$(recYCDOTC20.CDOTC2TX3 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 223, 13) = Format$(recYCDOTC20.CDOTC2TX4 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 236, 13) = Format$(recYCDOTC20.CDOTC2TX5 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 249, 13) = Format$(recYCDOTC20.CDOTC2TX6 * 1000000, "000000000000 ")
    
MsgTxtLen = MsgTxtLen + recYCDOTC20Len
End Sub



Public Sub srvYCDOTC20_ElpDisplay(recYCDOTC20 As typeYCDOTC20)
frmElpDisplay.fgData.Rows = 35
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2ETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2ETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2AGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2AGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2SER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2SER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2SSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2SSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2COP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2COP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2DOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2DOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2NUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2NUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2UTI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° UTILILSAT°./MODIF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2UTI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2EVE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2EVE
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2SEQ    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° SEQUENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2SEQ
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2COM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Code commission"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2COM
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2DEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Devise"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2DEV
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2CAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Catégorie client"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2CAT
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2CLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° Client"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2CLI
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2DEB    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Date début effet"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2DEB
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2FIN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Date fin effet"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2FIN
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2TVA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TVA (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2TVA
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2PER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Périodicité"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2PER
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2CUM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Cumulable (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2CUM
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2MTF 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant fixe"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2MTF
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2IND    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Indivisibilité (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2IND
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2AVE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Avis à échéance"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2AVE
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2MT1   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant tranche 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2MT1
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2MT2   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant tranche 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2MT2
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2MT3   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant tranche 3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2MT3
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2MT4   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant tranche 4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2MT4
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2MT5   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant tranche 5"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2MT5
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2MT6   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant tranche 6"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2MT6
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2TX1 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2TX1
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2TX2 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2TX2
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2TX3 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2TX3
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2TX4 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2TX4
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2TX5 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 5"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2TX5
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTC2TX6 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 6"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTC20.CDOTC2TX6
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYCDOTC20_Export_CSV()
Dim xIn As String
Open "C:\Temp\YCDOTC20.txt" For Input As #1
Open "C:\Temp\YCDOTC20.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "CDOTC2ETB;CDOTC2AGE;CDOTC2SER;CDOTC2SSE;CDOTC2COP;CDOTC2DOS;CDOTC2NUR;CDOTC2UTI;CDOTC2EVE;CDOTC2SEQ;CDOTC2COM;CDOTC2DEV;CDOTC2CAT;CDOTC2CLI;CDOTC2DEB;CDOTC2FIN;CDOTC2TVA;CDOTC2PER;CDOTC2CUM;CDOTC2MTF;CDOTC2IND;CDOTC2AVE;CDOTC2MT1;CDOTC2MT2;CDOTC2MT3;CDOTC2MT4;CDOTC2MT5;CDOTC2MT6;CDOTC2TX1;CDOTC2TX2;CDOTC2TX3;CDOTC2TX4;CDOTC2TX5;CDOTC2TX6;"
    Print #2, "CODE ETABLISSEMENT;AGENCE;SERVICE;SOUS-SERVICE;CODE OPERATION;NUMERO DOSSIER;N° RENOUVELLEMENT;N° UTILILSAT°./MODIF;EVENEMENT;N° SEQUENCE;Code commission;Devise;Catégorie client;N° Client;Date début effet;Date fin effet;TVA (O/N);Périodicité;Cumulable (O/N);Montant fixe;Indivisibilité (O/N);Avis à échéance;Montant tranche 1;Montant tranche 2;Montant tranche 3;Montant tranche 4;Montant tranche 5;Montant tranche 6;Taux tranche 1;Taux tranche 2;Taux tranche 3;Taux tranche 4;Taux tranche 5;Taux tranche 6;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" & mId$(xIn, 6, 5) & ";" _
      & mId$(xIn, 11, 2) & ";" & mId$(xIn, 13, 2) & ";" _
      & mId$(xIn, 15, 3) & ";" & mId$(xIn, 18, 10) & ";" _
      & mId$(xIn, 28, 4) & ";" & mId$(xIn, 32, 6) & ";" _
      & mId$(xIn, 38, 2) & ";" & mId$(xIn, 40, 4) & ";" _
      & mId$(xIn, 44, 6) & ";" & mId$(xIn, 50, 3) & ";" _
      & mId$(xIn, 53, 3) & ";" & mId$(xIn, 56, 7) & ";" _
      & mId$(xIn, 63, 8) & ";" & mId$(xIn, 71, 8) & ";" _
      & mId$(xIn, 79, 1) & ";" & mId$(xIn, 80, 1) & ";" _
      & mId$(xIn, 81, 1) & ";" & mId$(xIn, 82, 16) & ";" _
      & mId$(xIn, 98, 1) & ";" & mId$(xIn, 99, 1) & ";" _
      & mId$(xIn, 100, 14) & ";" & mId$(xIn, 114, 14) & ";" _
      & mId$(xIn, 128, 14) & ";" & mId$(xIn, 142, 14) & ";" _
      & mId$(xIn, 156, 14) & ";" & mId$(xIn, 170, 14) & ";" _
      & mId$(xIn, 184, 13) & ";" & mId$(xIn, 197, 13) & ";" _
      & mId$(xIn, 210, 13) & ";" & mId$(xIn, 223, 13) & ";" _
      & mId$(xIn, 236, 13) & ";" & mId$(xIn, 249, 13) & ";"
Loop
Close
End Sub

'---------------------------------------------------------
Private Function srvYCDOTC20_Seek(recYCDOTC20 As typeYCDOTC20)
'---------------------------------------------------------

srvYCDOTC20_Seek = "?"
MsgTxtLen = 0
Call srvYCDOTC20_PutBuffer(recYCDOTC20)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDOTC20_GetBuffer(recYCDOTC20)) Then
            srvYCDOTC20_Seek = Null
        Else
            Call srvYCDOTC20_Error(recYCDOTC20)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvYCDOTC20_Update(recYCDOTC20 As typeYCDOTC20)
'-----------------------------------------------------

srvYCDOTC20_Update = "?"

MsgTxtLen = 0
Call srvYCDOTC20_PutBuffer(recYCDOTC20)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOTC20_GetBuffer(recYCDOTC20)) Then
        Call srvYCDOTC20_Error(recYCDOTC20)
        srvYCDOTC20_Update = recYCDOTC20.Err
        Exit Function
    Else
        srvYCDOTC20_Update = Null
    End If
Else
    recYCDOTC20.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDOTC20_Init(recYCDOTC20 As typeYCDOTC20)
'---------------------------------------------------------
MsgTxt = Space$(recYCDOTC20Len)
MsgTxtIndex = 0
Call srvYCDOTC20_GetBuffer(recYCDOTC20)
recYCDOTC20.Obj = "ZCDODOS0_S"
recYCDOTC20.CDOTC2ETB = 1    'As String Integer                        ' CODE ETABLISSEMENT
recYCDOTC20.CDOTC2AGE = 1     'As String Integer                        ' AGENCE
recYCDOTC20.CDOTC2SER = ""    'As String * 2                     ' SERVICE
recYCDOTC20.CDOTC2SSE = ""    'As String * 2                     ' SOUS-SERVICE
recYCDOTC20.CDOTC2COP = ""    'As String * 3                     ' CODE OPERATION
recYCDOTC20.CDOTC2DOS = 0     'As String Long                           ' NUMERO DOSSIER
recYCDOTC20.CDOTC2NUR = 0     'As String Long                           ' N° RENOUVELLEMENT
recYCDOTC20.CDOTC2UTI = 0     'As String Long                           ' N° UTILILSAT°./MODIF
recYCDOTC20.CDOTC2EVE = ""    'As String * 2                     ' EVENEMENT
recYCDOTC20.CDOTC2SEQ = 0     'As String Long                           ' N° SEQUENCE
recYCDOTC20.CDOTC2COM = ""    'As String * 6                     ' Code commission
recYCDOTC20.CDOTC2DEV = ""    'As String * 3                     ' Devise
recYCDOTC20.CDOTC2CAT = ""    'As String * 3                     ' Catégorie client
recYCDOTC20.CDOTC2CLI = ""    'As String * 7                     ' N° Client
recYCDOTC20.CDOTC2DEB = 0     'As String Long                           ' Date début effet
recYCDOTC20.CDOTC2FIN = 0     'As String Long                           ' Date fin effet
recYCDOTC20.CDOTC2TVA = ""    'As String * 1                     ' TVA (O/N)
recYCDOTC20.CDOTC2PER = ""    'As String * 1                     ' Périodicité
recYCDOTC20.CDOTC2CUM = ""    'As String * 1                     ' Cumulable (O/N)
recYCDOTC20.CDOTC2MTF = 0     'As String Currency                       ' Montant fixe
recYCDOTC20.CDOTC2IND = ""    'As String * 1                     ' Indivisibilité (O/N)
recYCDOTC20.CDOTC2AVE = ""    'As String * 1                     ' Avis à échéance
recYCDOTC20.CDOTC2MT1 = 0     'As String Long                           ' Montant tranche 1
recYCDOTC20.CDOTC2MT2 = 0     'As String Long                           ' Montant tranche 2
recYCDOTC20.CDOTC2MT3 = 0     'As String Long                           ' Montant tranche 3
recYCDOTC20.CDOTC2MT4 = 0     'As String Long                           ' Montant tranche 4
recYCDOTC20.CDOTC2MT5 = 0     'As String Long                           ' Montant tranche 5
recYCDOTC20.CDOTC2MT6 = 0     'As String Long                           ' Montant tranche 6
recYCDOTC20.CDOTC2TX1 = 0     'As String Double                         ' Taux tranche 1
recYCDOTC20.CDOTC2TX2 = 0     'As String Double                         ' Taux tranche 2
recYCDOTC20.CDOTC2TX3 = 0     'As String Double                         ' Taux tranche 3
recYCDOTC20.CDOTC2TX4 = 0     'As String Double                         ' Taux tranche 4
recYCDOTC20.CDOTC2TX5 = 0     'As String Double                         ' Taux tranche 5
recYCDOTC20.CDOTC2TX6 = 0     'As String Double                         ' Taux tranche 6

End Sub





