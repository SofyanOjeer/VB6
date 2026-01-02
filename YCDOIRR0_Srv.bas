Attribute VB_Name = "srvYCDOIRR0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOIRR0Len = 140 ' 34 + 106
Public Const recYCDOIRR0_Block = 200
Public Const constYCDOIRR0 = "YCDOIRR0"
Dim meYbase As typeYBase
Dim paramYCDOIRR0_Import As String

Type typeYCDOIRR0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CDOIRRETB       As Integer                        ' CODE ETABLISSEMENT
    CDOIRRAGE       As Integer                        ' AGENCE
    CDOIRRSER       As String * 2                     ' SERVICE
    CDOIRRSSE       As String * 2                     ' SOUS-SERVICE
    CDOIRRCOP       As String * 3                     ' CODE OPERATION
    CDOIRRDOS       As Long                           ' NUMERO DOSSIER
    CDOIRRNUR       As Long                           ' N° RENOUVELLEMENT
    CDOIRRUTI       As Long                           ' N° UTILISATION
    CDOIRRSEQ       As Long                           ' N° SEQUENCE
    CDOIRRTEX       As String * 75                    ' TEXTE
End Type
    
'---------------------------------------------------------
Public Function srvYCDOIRR0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOIRR0 As typeYCDOIRR0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOIRR0_GetBuffer_ODBC = Null

    recYCDOIRR0.CDOIRRETB = rsADO("CDOIRRETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOIRR0.CDOIRRAGE = rsADO("CDOIRRAGE")    'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOIRR0.CDOIRRSER = rsADO("CDOIRRSER")    'mId$(MsgTxt, K + 11, 2)
    recYCDOIRR0.CDOIRRSSE = rsADO("CDOIRRSSE")    'mId$(MsgTxt, K + 13, 2)
    recYCDOIRR0.CDOIRRCOP = rsADO("CDOIRRCOP")    'mId$(MsgTxt, K + 15, 3)
    recYCDOIRR0.CDOIRRDOS = rsADO("CDOIRRDOS")    'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOIRR0.CDOIRRNUR = rsADO("CDOIRRNUR")    'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOIRR0.CDOIRRUTI = rsADO("CDOIRRUTI")    'CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOIRR0.CDOIRRSEQ = rsADO("CDOIRRSEQ")    'CLng(Val(mId$(MsgTxt, K + 38, 4)))
    recYCDOIRR0.CDOIRRTEX = rsADO("CDOIRRTEX")    'mId$(MsgTxt, K + 42, 75)

Exit Function

Error_Handler:
srvYCDOIRR0_GetBuffer_ODBC = Error

End Function


Public Function srvYCDOIRR0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOIRR0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOIRR0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDOIRR0_Import = "?"

paramYCDOIRR0_Import = paramYBase_DataF & Trim(constYCDOIRR0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOIRR0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOIRR0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOIRR0
            meYbase.K1 = mId$(xIn, 15, 27) 'recYCDOIRR0.CDODOSCOP & recYCDOIRR0.CDODOSDOS .........
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOIRR0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOIRR0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOIRR0_Import" & xIn, vbCritical, Error
Close

srvYCDOIRR0_Import = Error
End Function

Public Function srvYCDOIRR0_Import_Read(lId As String, lYCDOIRR0 As typeYCDOIRR0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDOIRR0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOIRR0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOIRR0_GetBuffer lYCDOIRR0
    srvYCDOIRR0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOIRR0_Import_Read" & xIn, vbCritical, Error
srvYCDOIRR0_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYCDOIRR0_Monitor(recYCDOIRR0 As typeYCDOIRR0)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDOIRR0.Method), 1, 4)
    Case "Seek"
                srvYCDOIRR0_Monitor = srvYCDOIRR0_Seek(recYCDOIRR0)
    Case Else
                recYCDOIRR0.Err = recYCDOIRR0.Method
                Call srvYCDOIRR0_Error(recYCDOIRR0)
                srvYCDOIRR0_Monitor = recYCDOIRR0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDOIRR0_Error(recYCDOIRR0 As typeYCDOIRR0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOIRR0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOIRR0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOIRR0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDOIRR0s.bas  ( " _
                & Trim(recYCDOIRR0.obj) & " : " & Trim(recYCDOIRR0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDOIRR0_GetBuffer(recYCDOIRR0 As typeYCDOIRR0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOIRR0_GetBuffer = Null
recYCDOIRR0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOIRR0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOIRR0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOIRR0.Err = Space$(10) Then

    recYCDOIRR0.CDOIRRETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOIRR0.CDOIRRAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOIRR0.CDOIRRSER = mId$(MsgTxt, K + 11, 2)
    recYCDOIRR0.CDOIRRSSE = mId$(MsgTxt, K + 13, 2)
    recYCDOIRR0.CDOIRRCOP = mId$(MsgTxt, K + 15, 3)
    recYCDOIRR0.CDOIRRDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOIRR0.CDOIRRNUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOIRR0.CDOIRRUTI = CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOIRR0.CDOIRRSEQ = CLng(Val(mId$(MsgTxt, K + 38, 4)))
    recYCDOIRR0.CDOIRRTEX = mId$(MsgTxt, K + 42, 75)

Else
    srvYCDOIRR0_GetBuffer = recYCDOIRR0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOIRR0Len

End Function

'---------------------------------------------------------
Private Sub srvYCDOIRR0_PutBuffer(recYCDOIRR0 As typeYCDOIRR0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDOIRR0Len) = Space$(recYCDOIRR0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOIRR0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOIRR0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOIRR0.CDOIRRETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDOIRR0.CDOIRRAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDOIRR0.CDOIRRSER
    Mid$(MsgTxt, K + 13, 2) = recYCDOIRR0.CDOIRRSSE
    Mid$(MsgTxt, K + 15, 3) = recYCDOIRR0.CDOIRRCOP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDOIRR0.CDOIRRDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDOIRR0.CDOIRRNUR, "000 ")
    Mid$(MsgTxt, K + 32, 6) = Format$(recYCDOIRR0.CDOIRRUTI, "00000 ")
    Mid$(MsgTxt, K + 38, 4) = Format$(recYCDOIRR0.CDOIRRSEQ, "000 ")
    Mid$(MsgTxt, K + 42, 75) = recYCDOIRR0.CDOIRRTEX

End Sub


'---------------------------------------------------------
Private Function srvYCDOIRR0_Seek(recYCDOIRR0 As typeYCDOIRR0)
'---------------------------------------------------------

srvYCDOIRR0_Seek = "?"
MsgTxtLen = 0
Call srvYCDOIRR0_PutBuffer(recYCDOIRR0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDOIRR0_GetBuffer(recYCDOIRR0)) Then
            srvYCDOIRR0_Seek = Null
        Else
            Call srvYCDOIRR0_Error(recYCDOIRR0)
        End If
    End If
End If

End Function

Public Sub srvYCDOIRR0_ElpDisplay(recYCDOIRR0 As typeYCDOIRR0)

frmElpDisplay.fgData.Rows = 11
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRCOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRCOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRNUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRNUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRUTI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° UTILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRUTI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRSEQ    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° SEQUENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRSEQ
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOIRRTEX   75A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TEXTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOIRR0.CDOIRRTEX
frmElpDisplay.Show vbModal
End Sub

'-----------------------------------------------------
Function srvYCDOIRR0_Update(recYCDOIRR0 As typeYCDOIRR0)
'-----------------------------------------------------

srvYCDOIRR0_Update = "?"

MsgTxtLen = 0
Call srvYCDOIRR0_PutBuffer(recYCDOIRR0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOIRR0_GetBuffer(recYCDOIRR0)) Then
        Call srvYCDOIRR0_Error(recYCDOIRR0)
        srvYCDOIRR0_Update = recYCDOIRR0.Err
        Exit Function
    Else
        srvYCDOIRR0_Update = Null
    End If
Else
    recYCDOIRR0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDOIRR0_Init(recYCDOIRR0 As typeYCDOIRR0)
'---------------------------------------------------------
MsgTxt = Space$(recYCDOIRR0Len)
MsgTxtIndex = 0
Call srvYCDOIRR0_GetBuffer(recYCDOIRR0)
recYCDOIRR0.obj = "ZCDODOS0_S"

recYCDOIRR0.CDOIRRETB = 0     'As String Integer                  ' CODE ETABLISSEMENT
recYCDOIRR0.CDOIRRAGE = 0     'As String Integer                  ' AGENCE
recYCDOIRR0.CDOIRRSER = ""     'As String * 2                     ' SERVICE
recYCDOIRR0.CDOIRRSSE = ""     'As String * 2                     ' SOUS-SERVICE
recYCDOIRR0.CDOIRRCOP = ""     'As String * 3                     ' CODE OPERATION
recYCDOIRR0.CDOIRRDOS = 0     'As String Long                     ' NUMERO DOSSIER
recYCDOIRR0.CDOIRRNUR = 0     'As String Long                     ' N° RENOUVELLEMENT
recYCDOIRR0.CDOIRRUTI = 0     'As String Long                     ' N° UTILISATION
recYCDOIRR0.CDOIRRSEQ = 0     'As String Long                     ' N° SEQUENCE
recYCDOIRR0.CDOIRRTEX = ""     'As String * 75                    ' TEXTE

End Sub


