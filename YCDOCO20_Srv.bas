Attribute VB_Name = "srvYCDOCO20"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOCO20Len = 367 ' 34 + 333
Public Const recYCDOCO20_Block = 20
Public Const constYCDOCO20 = "YCDOCO20"
Dim meYbase As typeYBase
Dim paramYCDOCO20_Import As String

Type typeYCDOCO20
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    CDOCO2ETB       As Integer                        ' Etablissement
    CDOCO2AGE       As Integer                        ' Agence
    CDOCO2SER       As String * 2                     ' Service
    CDOCO2SSE       As String * 2                     ' Sous service
    CDOCO2COP       As String * 3                     ' Code Opération
    CDOCO2DOS       As Long                           ' N° Dossier
    CDOCO2NUR       As Long                           ' N° Renouv
    CDOCO2UTI       As Long                           ' N° Utilisation
    CDOCO2EVE       As String * 2                     ' Evénement
    CDOCO2SEQ       As Long                           ' N° Séquence
    CDOCO2SPE       As Long                           ' N° Séq Pério
    CDOCO2TVA       As String * 1                     ' TVA O/N
    CDOCO2PER       As String * 1                     ' Périodicité
    CDOCO2CUM       As String * 1                     ' Cumulable (O/N)
    CDOCO2IND       As String * 1                     ' Indivisibilité
    CDOCO2AVE       As String * 1                     ' Avis à échéance
    CDOCO2TYA       As String * 2                     ' Type Assiette
    CDOCO2MTA       As Currency 'Long                           ' Mt Assiette
    CDOCO2JRB       As String * 1                     ' Jours Reel/Banc
    CDOCO2ANN       As String * 1                     ' Type année
    CDOCO2NBJ       As Long                           ' Nb jours
    CDOCO2MIN       As Long                           ' Montant minimum
    CDOCO2MAX       As Long                           ' Montant maximum
    CDOCO2SEU       As Long                           ' Seuil exonérat°
    CDOCO2MT1       As Long                           ' Mt tranche 1
    CDOCO2MT2       As Long                           ' Mt tranche 2
    CDOCO2MT3       As Long                           ' Mt tranche 3
    CDOCO2MT4       As Long                           ' Mt tranche 4
    CDOCO2MT5       As Long                           ' Mt tranche 5
    CDOCO2MT6       As Long                           ' Mt tranche 6
    CDOCO2TX1       As Double                         ' Taux tranche 1
    CDOCO2TX2       As Double                         ' Taux tranche 2
    CDOCO2TX3       As Double                         ' Taux tranche 3
    CDOCO2TX4       As Double                         ' Taux tranche 4
    CDOCO2TX5       As Double                         ' Taux tranche 5
    CDOCO2TX6       As Double                         ' Taux tranche 6
    CDOCO2MON       As Long                           ' Montant Calculé
    CDOCO2MTV       As Long                           ' Montant TVA
    CDOCO2MTE       As Long                           ' Montant Av.Extr
End Type
    
'---------------------------------------------------------
Public Function srvYCDOCO20_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOCO20 As typeYCDOCO20)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOCO20_GetBuffer_ODBC = Null

    recYCDOCO20.CDOCO2ETB = rsADO("CDOCO2ETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOCO20.CDOCO2AGE = rsADO("CDOCO2AGE")    'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOCO20.CDOCO2SER = rsADO("CDOCO2SER")    'mId$(MsgTxt, K + 11, 2)
    recYCDOCO20.CDOCO2SSE = rsADO("CDOCO2SSE")    'mId$(MsgTxt, K + 13, 2)
    recYCDOCO20.CDOCO2COP = rsADO("CDOCO2COP")    'mId$(MsgTxt, K + 15, 3)
    recYCDOCO20.CDOCO2DOS = rsADO("CDOCO2DOS")    'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOCO20.CDOCO2NUR = rsADO("CDOCO2NUR")    'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOCO20.CDOCO2UTI = rsADO("CDOCO2UTI")    'CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOCO20.CDOCO2EVE = rsADO("CDOCO2EVE")    'mId$(MsgTxt, K + 38, 2)
    recYCDOCO20.CDOCO2SEQ = rsADO("CDOCO2SEQ")    'CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOCO20.CDOCO2SPE = rsADO("CDOCO2SPE")    'CLng(Val(mId$(MsgTxt, K + 44, 4)))
    recYCDOCO20.CDOCO2TVA = rsADO("CDOCO2TVA")    'mId$(MsgTxt, K + 48, 1)
    recYCDOCO20.CDOCO2PER = rsADO("CDOCO2PER")    'mId$(MsgTxt, K + 49, 1)
    recYCDOCO20.CDOCO2CUM = rsADO("CDOCO2CUM")    'mId$(MsgTxt, K + 50, 1)
    recYCDOCO20.CDOCO2IND = rsADO("CDOCO2IND")    'mId$(MsgTxt, K + 51, 1)
    recYCDOCO20.CDOCO2AVE = rsADO("CDOCO2AVE")    'mId$(MsgTxt, K + 52, 1)
    recYCDOCO20.CDOCO2TYA = rsADO("CDOCO2TYA")    'mId$(MsgTxt, K + 53, 2)
    recYCDOCO20.CDOCO2MTA = rsADO("CDOCO2MTA")    'CLng(Val(mId$(MsgTxt, K + 55, 16)))
    recYCDOCO20.CDOCO2JRB = rsADO("CDOCO2JRB")    'mId$(MsgTxt, K + 71, 1)
    recYCDOCO20.CDOCO2ANN = rsADO("CDOCO2ANN")    'mId$(MsgTxt, K + 72, 1)
    recYCDOCO20.CDOCO2NBJ = rsADO("CDOCO2NBJ")    'CLng(Val(mId$(MsgTxt, K + 73, 5)))
    recYCDOCO20.CDOCO2MIN = rsADO("CDOCO2MIN")    'CLng(Val(mId$(MsgTxt, K + 78, 16)))
    recYCDOCO20.CDOCO2MAX = rsADO("CDOCO2MAX")    'CLng(Val(mId$(MsgTxt, K + 94, 16)))
    recYCDOCO20.CDOCO2SEU = rsADO("CDOCO2SEU")    'CLng(Val(mId$(MsgTxt, K + 110, 14)))
    recYCDOCO20.CDOCO2MT1 = rsADO("CDOCO2MT1")    'CLng(Val(mId$(MsgTxt, K + 124, 14)))
    recYCDOCO20.CDOCO2MT2 = rsADO("CDOCO2MT2")    'CLng(Val(mId$(MsgTxt, K + 138, 14)))
    recYCDOCO20.CDOCO2MT3 = rsADO("CDOCO2MT3")    'CLng(Val(mId$(MsgTxt, K + 152, 14)))
    recYCDOCO20.CDOCO2MT4 = rsADO("CDOCO2MT4")    'CLng(Val(mId$(MsgTxt, K + 166, 14)))
    recYCDOCO20.CDOCO2MT5 = rsADO("CDOCO2MT5")    'CLng(Val(mId$(MsgTxt, K + 180, 14)))
    recYCDOCO20.CDOCO2MT6 = rsADO("CDOCO2MT6")    'CLng(Val(mId$(MsgTxt, K + 194, 14)))
    recYCDOCO20.CDOCO2TX1 = rsADO("CDOCO2TX1")    'CDbl(Val(mId$(MsgTxt, K + 208, 13))) / 1000000
    recYCDOCO20.CDOCO2TX2 = rsADO("CDOCO2TX2")    'CDbl(Val(mId$(MsgTxt, K + 221, 13))) / 1000000
    recYCDOCO20.CDOCO2TX3 = rsADO("CDOCO2TX3")    'CDbl(Val(mId$(MsgTxt, K + 234, 13))) / 1000000
    recYCDOCO20.CDOCO2TX4 = rsADO("CDOCO2TX4")    'CDbl(Val(mId$(MsgTxt, K + 247, 13))) / 1000000
    recYCDOCO20.CDOCO2TX5 = rsADO("CDOCO2TX5")    'CDbl(Val(mId$(MsgTxt, K + 260, 13))) / 1000000
    recYCDOCO20.CDOCO2TX6 = rsADO("CDOCO2TX6")    'CDbl(Val(mId$(MsgTxt, K + 273, 13))) / 1000000
    recYCDOCO20.CDOCO2MON = rsADO("CDOCO2MON")    'CLng(Val(mId$(MsgTxt, K + 286, 16)))
    recYCDOCO20.CDOCO2MTV = rsADO("CDOCO2MTV")    'CLng(Val(mId$(MsgTxt, K + 302, 16)))
    recYCDOCO20.CDOCO2MTE = rsADO("CDOCO2MTE")    'CLng(Val(mId$(MsgTxt, K + 318, 16)))

Exit Function

Error_Handler:
srvYCDOCO20_GetBuffer_ODBC = Error

End Function



Public Function srvYCDOCO20_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOCO20
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOCO20_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDOCO20_Import = "?"

paramYCDOCO20_Import = paramYBase_DataF & Trim(constYCDOCO20) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOCO20_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOCO20) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOCO20
            meYbase.K1 = mId$(xIn, 15, 33) 'recYCDOCO20.CDODOSCOP & recYCDOCO20.CDODOSDOS .........
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOCO20_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOCO20
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOCO20_Import" & xIn, vbCritical, Error
Close

srvYCDOCO20_Import = Error
End Function

Public Function srvYCDOCO20_Import_Read(lId As String, lYCDOCO20 As typeYCDOCO20)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDOCO20_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOCO20
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOCO20_GetBuffer lYCDOCO20
    srvYCDOCO20_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOCO20_Import_Read" & xIn, vbCritical, Error
srvYCDOCO20_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYCDOCO20_Monitor(recYCDOCO20 As typeYCDOCO20)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDOCO20.Method), 1, 4)
    Case "Seek"
                srvYCDOCO20_Monitor = srvYCDOCO20_Seek(recYCDOCO20)
    Case Else
                recYCDOCO20.Err = recYCDOCO20.Method
                Call srvYCDOCO20_Error(recYCDOCO20)
                srvYCDOCO20_Monitor = recYCDOCO20.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDOCO20_Error(recYCDOCO20 As typeYCDOCO20)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOCO20" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOCO20.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOCO20.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDOCO20s.bas  ( " _
                & Trim(recYCDOCO20.Obj) & " : " & Trim(recYCDOCO20.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDOCO20_GetBuffer(recYCDOCO20 As typeYCDOCO20)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOCO20_GetBuffer = Null
recYCDOCO20.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOCO20.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOCO20.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOCO20.Err = Space$(10) Then
    recYCDOCO20.CDOCO2ETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOCO20.CDOCO2AGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOCO20.CDOCO2SER = mId$(MsgTxt, K + 11, 2)
    recYCDOCO20.CDOCO2SSE = mId$(MsgTxt, K + 13, 2)
    recYCDOCO20.CDOCO2COP = mId$(MsgTxt, K + 15, 3)
    recYCDOCO20.CDOCO2DOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOCO20.CDOCO2NUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOCO20.CDOCO2UTI = CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOCO20.CDOCO2EVE = mId$(MsgTxt, K + 38, 2)
    recYCDOCO20.CDOCO2SEQ = CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOCO20.CDOCO2SPE = CLng(Val(mId$(MsgTxt, K + 44, 4)))
    recYCDOCO20.CDOCO2TVA = mId$(MsgTxt, K + 48, 1)
    recYCDOCO20.CDOCO2PER = mId$(MsgTxt, K + 49, 1)
    recYCDOCO20.CDOCO2CUM = mId$(MsgTxt, K + 50, 1)
    recYCDOCO20.CDOCO2IND = mId$(MsgTxt, K + 51, 1)
    recYCDOCO20.CDOCO2AVE = mId$(MsgTxt, K + 52, 1)
    recYCDOCO20.CDOCO2TYA = mId$(MsgTxt, K + 53, 2)
    recYCDOCO20.CDOCO2MTA = CCur(Val(mId$(MsgTxt, K + 55, 16)))
    recYCDOCO20.CDOCO2JRB = mId$(MsgTxt, K + 71, 1)
    recYCDOCO20.CDOCO2ANN = mId$(MsgTxt, K + 72, 1)
    recYCDOCO20.CDOCO2NBJ = CLng(Val(mId$(MsgTxt, K + 73, 5)))
    recYCDOCO20.CDOCO2MIN = CLng(Val(mId$(MsgTxt, K + 78, 16)))
    recYCDOCO20.CDOCO2MAX = CLng(Val(mId$(MsgTxt, K + 94, 16)))
    recYCDOCO20.CDOCO2SEU = CLng(Val(mId$(MsgTxt, K + 110, 14)))
    recYCDOCO20.CDOCO2MT1 = CLng(Val(mId$(MsgTxt, K + 124, 14)))
    recYCDOCO20.CDOCO2MT2 = CLng(Val(mId$(MsgTxt, K + 138, 14)))
    recYCDOCO20.CDOCO2MT3 = CLng(Val(mId$(MsgTxt, K + 152, 14)))
    recYCDOCO20.CDOCO2MT4 = CLng(Val(mId$(MsgTxt, K + 166, 14)))
    recYCDOCO20.CDOCO2MT5 = CLng(Val(mId$(MsgTxt, K + 180, 14)))
    recYCDOCO20.CDOCO2MT6 = CLng(Val(mId$(MsgTxt, K + 194, 14)))
    recYCDOCO20.CDOCO2TX1 = CDbl(Val(mId$(MsgTxt, K + 208, 13))) / 1000000
    recYCDOCO20.CDOCO2TX2 = CDbl(Val(mId$(MsgTxt, K + 221, 13))) / 1000000
    recYCDOCO20.CDOCO2TX3 = CDbl(Val(mId$(MsgTxt, K + 234, 13))) / 1000000
    recYCDOCO20.CDOCO2TX4 = CDbl(Val(mId$(MsgTxt, K + 247, 13))) / 1000000
    recYCDOCO20.CDOCO2TX5 = CDbl(Val(mId$(MsgTxt, K + 260, 13))) / 1000000
    recYCDOCO20.CDOCO2TX6 = CDbl(Val(mId$(MsgTxt, K + 273, 13))) / 1000000
    recYCDOCO20.CDOCO2MON = CLng(Val(mId$(MsgTxt, K + 286, 16)))
    recYCDOCO20.CDOCO2MTV = CLng(Val(mId$(MsgTxt, K + 302, 16)))
    recYCDOCO20.CDOCO2MTE = CLng(Val(mId$(MsgTxt, K + 318, 16)))
Else
    srvYCDOCO20_GetBuffer = recYCDOCO20.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOCO20Len

End Function

'---------------------------------------------------------
Private Sub srvYCDOCO20_PutBuffer(recYCDOCO20 As typeYCDOCO20)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDOCO20Len) = Space$(recYCDOCO20Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOCO20.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOCO20.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOCO20.CDOCO2ETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDOCO20.CDOCO2AGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDOCO20.CDOCO2SER
    Mid$(MsgTxt, K + 13, 2) = recYCDOCO20.CDOCO2SSE
    Mid$(MsgTxt, K + 15, 3) = recYCDOCO20.CDOCO2COP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDOCO20.CDOCO2DOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDOCO20.CDOCO2NUR, "000 ")
    Mid$(MsgTxt, K + 32, 6) = Format$(recYCDOCO20.CDOCO2UTI, "00000 ")
    Mid$(MsgTxt, K + 38, 2) = recYCDOCO20.CDOCO2EVE
    Mid$(MsgTxt, K + 40, 4) = Format$(recYCDOCO20.CDOCO2SEQ, "000 ")
    Mid$(MsgTxt, K + 44, 4) = Format$(recYCDOCO20.CDOCO2SPE, "000 ")
    Mid$(MsgTxt, K + 48, 1) = recYCDOCO20.CDOCO2TVA
    Mid$(MsgTxt, K + 49, 1) = recYCDOCO20.CDOCO2PER
    Mid$(MsgTxt, K + 50, 1) = recYCDOCO20.CDOCO2CUM
    Mid$(MsgTxt, K + 51, 1) = recYCDOCO20.CDOCO2IND
    Mid$(MsgTxt, K + 52, 1) = recYCDOCO20.CDOCO2AVE
    Mid$(MsgTxt, K + 53, 2) = recYCDOCO20.CDOCO2TYA
    Mid$(MsgTxt, K + 55, 16) = Format$(recYCDOCO20.CDOCO2MTA, "000000000000000 ")
    Mid$(MsgTxt, K + 71, 1) = recYCDOCO20.CDOCO2JRB
    Mid$(MsgTxt, K + 72, 1) = recYCDOCO20.CDOCO2ANN
    Mid$(MsgTxt, K + 73, 5) = Format$(recYCDOCO20.CDOCO2NBJ, "0000 ")
    Mid$(MsgTxt, K + 78, 16) = Format$(recYCDOCO20.CDOCO2MIN, "000000000000000 ")
    Mid$(MsgTxt, K + 94, 16) = Format$(recYCDOCO20.CDOCO2MAX, "000000000000000 ")
    Mid$(MsgTxt, K + 110, 14) = Format$(recYCDOCO20.CDOCO2SEU, "0000000000000 ")
    Mid$(MsgTxt, K + 124, 14) = Format$(recYCDOCO20.CDOCO2MT1, "0000000000000 ")
    Mid$(MsgTxt, K + 138, 14) = Format$(recYCDOCO20.CDOCO2MT2, "0000000000000 ")
    Mid$(MsgTxt, K + 152, 14) = Format$(recYCDOCO20.CDOCO2MT3, "0000000000000 ")
    Mid$(MsgTxt, K + 166, 14) = Format$(recYCDOCO20.CDOCO2MT4, "0000000000000 ")
    Mid$(MsgTxt, K + 180, 14) = Format$(recYCDOCO20.CDOCO2MT5, "0000000000000 ")
    Mid$(MsgTxt, K + 194, 14) = Format$(recYCDOCO20.CDOCO2MT6, "0000000000000 ")
    Mid$(MsgTxt, K + 208, 13) = Format$(recYCDOCO20.CDOCO2TX1 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 221, 13) = Format$(recYCDOCO20.CDOCO2TX2 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 234, 13) = Format$(recYCDOCO20.CDOCO2TX3 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 247, 13) = Format$(recYCDOCO20.CDOCO2TX4 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 260, 13) = Format$(recYCDOCO20.CDOCO2TX5 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 273, 13) = Format$(recYCDOCO20.CDOCO2TX6 * 1000000, "000000000000 ")
    Mid$(MsgTxt, K + 286, 16) = Format$(recYCDOCO20.CDOCO2MON, "000000000000000 ")
    Mid$(MsgTxt, K + 302, 16) = Format$(recYCDOCO20.CDOCO2MTV, "000000000000000 ")
    Mid$(MsgTxt, K + 318, 16) = Format$(recYCDOCO20.CDOCO2MTE, "000000000000000 ")
    
    MsgTxtLen = MsgTxtLen + recYCDOCO20Len
End Sub

Public Sub srvYCDOCO20_ElpDisplay(recYCDOCO20 As typeYCDOCO20)
frmElpDisplay.fgData.Rows = 40
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2ETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Etablissement"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2ETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2AGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Agence"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2AGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2SER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Service"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2SER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2SSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Sous service"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2SSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2COP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Code Opération"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2COP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2DOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° Dossier"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2DOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2NUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° Renouv"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2NUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2UTI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° Utilisation"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2UTI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2EVE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Evénement"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2EVE
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2SEQ    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° Séquence"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2SEQ
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2SPE    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° Séq Pério"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2SPE
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2TVA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TVA O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2TVA
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2PER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Périodicité"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2PER
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2CUM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Cumulable (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2CUM
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2IND    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Indivisibilité"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2IND
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2AVE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Avis à échéance"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2AVE
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2TYA    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Type Assiette"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2TYA
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MTA   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mt Assiette"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MTA
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2JRB    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Jours Reel/Banc"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2JRB
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2ANN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Type année"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2ANN
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2NBJ    4P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Nb jours"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2NBJ
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MIN   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant minimum"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MIN
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MAX   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant maximum"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MAX
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2SEU   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Seuil exonérat°"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2SEU
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MT1   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mt tranche 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MT1
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MT2   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mt tranche 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MT2
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MT3   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mt tranche 3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MT3
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MT4   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mt tranche 4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MT4
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MT5   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mt tranche 5"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MT5
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MT6   13P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mt tranche 6"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MT6
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2TX1 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2TX1
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2TX2 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2TX2
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2TX3 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2TX3
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2TX4 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2TX4
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2TX5 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 5"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2TX5
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2TX6 12.6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Taux tranche 6"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2TX6
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MON   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant Calculé"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MON
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MTV   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MTV
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCO2MTE   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Montant Av.Extr"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCO20.CDOCO2MTE
frmElpDisplay.Show vbModal
End Sub

Public Sub srvYCDOCO20_Export_CSV()
Dim xIn As String
Open "C:\Temp\YCDOCO20.txt" For Input As #1
Open "C:\Temp\YCDOCO20.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "CDOCO2ETB;CDOCO2AGE;CDOCO2SER;CDOCO2SSE;CDOCO2COP;CDOCO2DOS;CDOCO2NUR;CDOCO2UTI;CDOCO2EVE;CDOCO2SEQ;CDOCO2SPE;CDOCO2TVA;CDOCO2PER;CDOCO2CUM;CDOCO2IND;CDOCO2AVE;CDOCO2TYA;CDOCO2MTA;CDOCO2JRB;CDOCO2ANN;CDOCO2NBJ;CDOCO2MIN;CDOCO2MAX;CDOCO2SEU;CDOCO2MT1;CDOCO2MT2;CDOCO2MT3;CDOCO2MT4;CDOCO2MT5;CDOCO2MT6;CDOCO2TX1;CDOCO2TX2;CDOCO2TX3;CDOCO2TX4;CDOCO2TX5;CDOCO2TX6;CDOCO2MON;CDOCO2MTV;CDOCO2MTE;"
    Print #2, "Etablissement;Agence;Service;Sous service;Code Opération;N° Dossier;N° Renouv;N° Utilisation;Evénement;N° Séquence;N° Séq Pério;TVA O/N;Périodicité;Cumulable (O/N);Indivisibilité;Avis à échéance;Type Assiette;Mt Assiette;Jours Reel/Banc;Type année;Nb jours;Montant minimum;Montant maximum;Seuil exonérat°;Mt tranche 1;Mt tranche 2;Mt tranche 3;Mt tranche 4;Mt tranche 5;Mt tranche 6;Taux tranche 1;Taux tranche 2;Taux tranche 3;Taux tranche 4;Taux tranche 5;Taux tranche 6;Montant Calculé;Montant TVA;Montant Av.Extr;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 5) & ";" & mId$(xIn, 11, 2) & ";" _
      & mId$(xIn, 13, 2) & ";" & mId$(xIn, 15, 3) & ";" _
      & mId$(xIn, 18, 10) & ";" & mId$(xIn, 28, 4) & ";" _
      & mId$(xIn, 32, 6) & ";" & mId$(xIn, 38, 2) & ";" _
      & mId$(xIn, 40, 4) & ";" & mId$(xIn, 44, 4) & ";" _
      & mId$(xIn, 48, 1) & ";" & mId$(xIn, 49, 1) & ";" _
      & mId$(xIn, 50, 1) & ";" & mId$(xIn, 51, 1) & ";" _
      & mId$(xIn, 52, 1) & ";" & mId$(xIn, 53, 2) & ";" _
      & mId$(xIn, 55, 16) & ";" & mId$(xIn, 71, 1) & ";" & mId$(xIn, 72, 1) & ";" _
      & mId$(xIn, 73, 5) & ";" & mId$(xIn, 78, 16) & ";" _
      & mId$(xIn, 94, 16) & ";" & mId$(xIn, 110, 14) & ";" & mId$(xIn, 124, 14) & ";" _
      & mId$(xIn, 138, 14) & ";" & mId$(xIn, 152, 14) & ";" _
      & mId$(xIn, 166, 14) & ";" & mId$(xIn, 180, 14) & ";" _
      & mId$(xIn, 194, 14) & ";" & mId$(xIn, 208, 13) & ";" _
      & mId$(xIn, 221, 13) & ";" & mId$(xIn, 234, 13) & ";" _
      & mId$(xIn, 247, 13) & ";" & mId$(xIn, 260, 13) & ";" _
      & mId$(xIn, 273, 13) & ";" & mId$(xIn, 286, 16) & ";" _
      & mId$(xIn, 302, 16) & ";" & mId$(xIn, 318, 16) & ";"
Loop
Close
End Sub

'---------------------------------------------------------
Private Function srvYCDOCO20_Seek(recYCDOCO20 As typeYCDOCO20)
'---------------------------------------------------------

srvYCDOCO20_Seek = "?"
MsgTxtLen = 0
Call srvYCDOCO20_PutBuffer(recYCDOCO20)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDOCO20_GetBuffer(recYCDOCO20)) Then
            srvYCDOCO20_Seek = Null
        Else
            Call srvYCDOCO20_Error(recYCDOCO20)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvYCDOCO20_Update(recYCDOCO20 As typeYCDOCO20)
'-----------------------------------------------------

srvYCDOCO20_Update = "?"

MsgTxtLen = 0
Call srvYCDOCO20_PutBuffer(recYCDOCO20)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOCO20_GetBuffer(recYCDOCO20)) Then
        Call srvYCDOCO20_Error(recYCDOCO20)
        srvYCDOCO20_Update = recYCDOCO20.Err
        Exit Function
    Else
        srvYCDOCO20_Update = Null
    End If
Else
    recYCDOCO20.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDOCO20_Init(recYCDOCO20 As typeYCDOCO20)
'---------------------------------------------------------
MsgTxt = Space$(recYCDOCO20Len)
MsgTxtIndex = 0
Call srvYCDOCO20_GetBuffer(recYCDOCO20)
recYCDOCO20.Obj = "ZCDODOS0_S"
recYCDOCO20.CDOCO2ETB = 1       'As Integer                        ' Etablissement
recYCDOCO20.CDOCO2AGE = 1       'As Integer                        ' Agence
recYCDOCO20.CDOCO2SER = ""       'As String * 2                     ' Service
recYCDOCO20.CDOCO2SSE = ""       'As String * 2                     ' Sous service
recYCDOCO20.CDOCO2COP = ""       'As String * 3                     ' Code Opération
recYCDOCO20.CDOCO2DOS = 0       'As Long                           ' N° Dossier
recYCDOCO20.CDOCO2NUR = 0       'As Long                           ' N° Renouv
recYCDOCO20.CDOCO2UTI = 0       'As Long                           ' N° Utilisation
recYCDOCO20.CDOCO2EVE = ""       'As String * 2                     ' Evénement
recYCDOCO20.CDOCO2SEQ = 0       'As Long                           ' N° Séquence
recYCDOCO20.CDOCO2SPE = 0       'As Long                           ' N° Séq Pério
recYCDOCO20.CDOCO2TVA = ""       'As String * 1                     ' TVA O/N
recYCDOCO20.CDOCO2PER = ""       'As String * 1                     ' Périodicité
recYCDOCO20.CDOCO2CUM = ""       'As String * 1                     ' Cumulable (O/N)
recYCDOCO20.CDOCO2IND = ""       'As String * 1                     ' Indivisibilité
recYCDOCO20.CDOCO2AVE = ""       'As String * 1                     ' Avis à échéance
recYCDOCO20.CDOCO2TYA = ""       'As String * 2                     ' Type Assiette
recYCDOCO20.CDOCO2MTA = 0       'As Long                           ' Mt Assiette
recYCDOCO20.CDOCO2JRB = ""       'As String * 1                     ' Jours Reel/Banc
recYCDOCO20.CDOCO2ANN = ""       'As String * 1                     ' Type année
recYCDOCO20.CDOCO2NBJ = 0       'As Long                           ' Nb jours
recYCDOCO20.CDOCO2MIN = 0       'As Long                           ' Montant minimum
recYCDOCO20.CDOCO2MAX = 0       'As Long                           ' Montant maximum
recYCDOCO20.CDOCO2SEU = 0       'As Long                           ' Seuil exonérat°
recYCDOCO20.CDOCO2MT1 = 0       'As Long                           ' Mt tranche 1
recYCDOCO20.CDOCO2MT2 = 0       'As Long                           ' Mt tranche 2
recYCDOCO20.CDOCO2MT3 = 0       'As Long                           ' Mt tranche 3
recYCDOCO20.CDOCO2MT4 = 0       'As Long                           ' Mt tranche 4
recYCDOCO20.CDOCO2MT5 = 0       'As Long                           ' Mt tranche 5
recYCDOCO20.CDOCO2MT6 = 0       'As Long                           ' Mt tranche 6
recYCDOCO20.CDOCO2TX1 = 0       'As Double                         ' Taux tranche 1
recYCDOCO20.CDOCO2TX2 = 0       'As Double                         ' Taux tranche 2
recYCDOCO20.CDOCO2TX3 = 0       'As Double                         ' Taux tranche 3
recYCDOCO20.CDOCO2TX4 = 0       'As Double                         ' Taux tranche 4
recYCDOCO20.CDOCO2TX5 = 0       'As Double                         ' Taux tranche 5
recYCDOCO20.CDOCO2TX6 = 0       'As Double                         ' Taux tranche 6
recYCDOCO20.CDOCO2MON = 0       'As Long                           ' Montant Calculé
recYCDOCO20.CDOCO2MTV = 0       'As Long                           ' Montant TVA
recYCDOCO20.CDOCO2MTE = 0       'As Long                           ' Montant Av.Extr

End Sub





