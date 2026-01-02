Attribute VB_Name = "srvYCDODES0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDODES0Len = 140 ' 34 + 106
Public Const recYCDODES0_Block = 200
Public Const constYCDODES0 = "YCDODES0"
Dim meYbase As typeYBase
Dim paramYCDODES0_Import As String

Type typeYCDODES0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CDODESETB       As Integer                        ' CODE ETABLISSEMENT
    CDODESAGE       As Integer                        ' AGENCE
    CDODESSER       As String * 2                     ' SERVICE
    CDODESSSE       As String * 2                     ' SOUS-SERVICE
    CDODESCOP       As String * 3                     ' CODE OPERATION
    CDODESDOS       As Long                           ' NUMERO DOSSIER
    CDODESNUR       As Long                           ' N° RENOUVELLEMENT
    CDODESUTI       As Long                           ' N° UTILISATION
    CDODESSEQ       As Long                           ' N° SEQUENCE
    CDODESTEX       As String * 65                    ' TEXTE
End Type
    
'---------------------------------------------------------
Public Function srvYCDODES0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDODES0 As typeYCDODES0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDODES0_GetBuffer_ODBC = Null

    recYCDODES0.CDODESETB = rsADO("CDODESETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDODES0.CDODESAGE = rsADO("CDODESAGE")    'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDODES0.CDODESSER = rsADO("CDODESSER")    'mId$(MsgTxt, K + 11, 2)
    recYCDODES0.CDODESSSE = rsADO("CDODESSSE")    'mId$(MsgTxt, K + 13, 2)
    recYCDODES0.CDODESCOP = rsADO("CDODESCOP")    'mId$(MsgTxt, K + 15, 3)
    recYCDODES0.CDODESDOS = rsADO("CDODESDOS")    'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDODES0.CDODESNUR = rsADO("CDODESNUR")    'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDODES0.CDODESUTI = rsADO("CDODESUTI")    'CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDODES0.CDODESSEQ = rsADO("CDODESSEQ")    'CLng(Val(mId$(MsgTxt, K + 38, 4)))
    recYCDODES0.CDODESTEX = rsADO("CDODESTEX")    'mId$(MsgTxt, K + 42, 65)

Exit Function

Error_Handler:
srvYCDODES0_GetBuffer_ODBC = Error

End Function


Public Function srvYCDODES0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDODES0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDODES0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDODES0_Import = "?"

paramYCDODES0_Import = paramYBase_DataF & Trim(constYCDODES0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDODES0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDODES0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDODES0
            meYbase.K1 = mId$(xIn, 15, 27) 'recYCDODES0.CDODOSCOP & recYCDODES0.CDODOSDOS .........
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDODES0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDODES0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDODES0_Import" & xIn, vbCritical, Error
Close

srvYCDODES0_Import = Error
End Function

Public Function srvYCDODES0_Import_Read(lId As String, lYCDODES0 As typeYCDODES0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDODES0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDODES0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDODES0_GetBuffer lYCDODES0
    srvYCDODES0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDODES0_Import_Read" & xIn, vbCritical, Error
srvYCDODES0_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYCDODES0_Monitor(recYCDODES0 As typeYCDODES0)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDODES0.Method), 1, 4)
    Case "Seek"
                srvYCDODES0_Monitor = srvYCDODES0_Seek(recYCDODES0)
    Case Else
                recYCDODES0.Err = recYCDODES0.Method
                Call srvYCDODES0_Error(recYCDODES0)
                srvYCDODES0_Monitor = recYCDODES0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDODES0_Error(recYCDODES0 As typeYCDODES0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDODES0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDODES0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDODES0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDODES0s.bas  ( " _
                & Trim(recYCDODES0.obj) & " : " & Trim(recYCDODES0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDODES0_GetBuffer(recYCDODES0 As typeYCDODES0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDODES0_GetBuffer = Null
recYCDODES0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDODES0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDODES0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDODES0.Err = Space$(10) Then

    recYCDODES0.CDODESETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDODES0.CDODESAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDODES0.CDODESSER = mId$(MsgTxt, K + 11, 2)
    recYCDODES0.CDODESSSE = mId$(MsgTxt, K + 13, 2)
    recYCDODES0.CDODESCOP = mId$(MsgTxt, K + 15, 3)
    recYCDODES0.CDODESDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDODES0.CDODESNUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDODES0.CDODESUTI = CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDODES0.CDODESSEQ = CLng(Val(mId$(MsgTxt, K + 38, 4)))
    recYCDODES0.CDODESTEX = mId$(MsgTxt, K + 42, 65)

Else
    srvYCDODES0_GetBuffer = recYCDODES0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDODES0Len

End Function

'---------------------------------------------------------
Private Sub srvYCDODES0_PutBuffer(recYCDODES0 As typeYCDODES0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDODES0Len) = Space$(recYCDODES0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDODES0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDODES0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDODES0.CDODESETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDODES0.CDODESAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDODES0.CDODESSER
    Mid$(MsgTxt, K + 13, 2) = recYCDODES0.CDODESSSE
    Mid$(MsgTxt, K + 15, 3) = recYCDODES0.CDODESCOP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDODES0.CDODESDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDODES0.CDODESNUR, "000 ")
    Mid$(MsgTxt, K + 32, 6) = Format$(recYCDODES0.CDODESUTI, "00000 ")
    Mid$(MsgTxt, K + 38, 4) = Format$(recYCDODES0.CDODESSEQ, "000 ")
    Mid$(MsgTxt, K + 42, 65) = recYCDODES0.CDODESTEX

End Sub


'---------------------------------------------------------
Private Function srvYCDODES0_Seek(recYCDODES0 As typeYCDODES0)
'---------------------------------------------------------

srvYCDODES0_Seek = "?"
MsgTxtLen = 0
Call srvYCDODES0_PutBuffer(recYCDODES0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDODES0_GetBuffer(recYCDODES0)) Then
            srvYCDODES0_Seek = Null
        Else
            Call srvYCDODES0_Error(recYCDODES0)
        End If
    End If
End If

End Function
Public Sub srvYCDODES0_ElpDisplay(recYCDODES0 As typeYCDODES0)
frmElpDisplay.fgData.Rows = 11
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESCOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESCOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESNUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESNUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESUTI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° UTILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESUTI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESSEQ    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° SEQUENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESSEQ
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODESTEX   65A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TEXTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODES0.CDODESTEX
frmElpDisplay.Show vbModal
End Sub

'-----------------------------------------------------
Function srvYCDODES0_Update(recYCDODES0 As typeYCDODES0)
'-----------------------------------------------------

srvYCDODES0_Update = "?"

MsgTxtLen = 0
Call srvYCDODES0_PutBuffer(recYCDODES0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDODES0_GetBuffer(recYCDODES0)) Then
        Call srvYCDODES0_Error(recYCDODES0)
        srvYCDODES0_Update = recYCDODES0.Err
        Exit Function
    Else
        srvYCDODES0_Update = Null
    End If
Else
    recYCDODES0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDODES0_Init(recYCDODES0 As typeYCDODES0)
'---------------------------------------------------------
MsgTxt = Space$(recYCDODES0Len)
MsgTxtIndex = 0
Call srvYCDODES0_GetBuffer(recYCDODES0)
recYCDODES0.obj = "ZCDODOS0_S"

recYCDODES0.CDODESETB = 0     'As String Integer                        ' CODE ETABLISSEMENT
recYCDODES0.CDODESAGE = 0     'As String Integer                        ' AGENCE
recYCDODES0.CDODESSER = ""     'As String * 2                     ' SERVICE
recYCDODES0.CDODESSSE = ""     'As String * 2                     ' SOUS-SERVICE
recYCDODES0.CDODESCOP = ""     'As String * 3                     ' CODE OPERATION
recYCDODES0.CDODESDOS = 0     'As String Long                           ' NUMERO DOSSIER
recYCDODES0.CDODESNUR = 0     'As String Long                           ' N° RENOUVELLEMENT
recYCDODES0.CDODESUTI = 0     'As String Long                           ' N° UTILISATION
recYCDODES0.CDODESSEQ = 0     'As String Long                           ' N° SEQUENCE
recYCDODES0.CDODESTEX = ""     'As String * 65                    ' TEXTE
End Sub







