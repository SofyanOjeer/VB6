Attribute VB_Name = "srvYSOLDE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYSOLDE0Len = 264 ' 34 +230
Public Const recYSOLDE0_Block = 100
Public Const memoYSOLDE0Len = 230
Public Const constYSOLDE0 = "YSOLDE0"
Public paramYSOLDE0_Import As String
Dim meYbase As typeYBase

Type typeYSOLDE0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SOLDEETA        As Integer                        ' ETABLISSEMENT
    SOLDEPLA        As Long                           ' NUMERO PLAN
    SOLDECOM        As String * 20                    ' NUMERO COMPTE
    SOLDEDMO        As Long                           ' DATE DERNIER MVT
    SOLDEDAN        As Long                           ' DATE ANTERIEUR
    SOLDECEN        As Currency                         ' SOLDE ENCOURS
    SOLDECAN        As Currency                         ' SOLDE ANTERIEUR
    SOLDEC01        As Currency                         ' SOLDE M
    SOLDEC02        As Currency                         ' SOLDE M -1
    SOLDEC03        As Currency                         ' SOLDE M -2
    SOLDEC04        As Currency                         ' SOLDE M -3
    SOLDEC05        As Currency                         ' SOLDE M -4
    SOLDEC06        As Currency                         ' SOLDE M -5
    SOLDEC07        As Currency                         ' SOLDE M -6
    SOLDEC08        As Currency                         ' SOLDE M -7
    SOLDEC09        As Currency                         ' SOLDE M -8
    SOLDEC10        As Currency                         ' SOLDE M -9
    SOLDEC11        As Currency                         ' SOLDE M -10
    SOLDEC12        As Currency                         ' SOLDE M -11
    SOLDEVEN        As Currency                         ' SOLDE VAL. ENCOURS
    SOLDEVAN        As Currency                         ' SOLDE VAL. ANTERIEUR
    SOLDEV01        As Currency                         ' SOLDE VAL. M
    SOLDEV02        As Currency                         ' SOLDE VAL. M -1
    SOLDEV03        As Currency                         ' SOLDE VAL. M -2
    SOLDEV04        As Currency                         ' SOLDE VAL. M -3
    SOLDEV05        As Currency                         ' SOLDE VAL. M -4
    SOLDEV06        As Currency                         ' SOLDE VAL. M -5
    SOLDEV07        As Currency                         ' SOLDE VAL. M -6
    SOLDEV08        As Currency                         ' SOLDE VAL. M -7
    SOLDEV09        As Currency                         ' SOLDE VAL. M -8
    SOLDEV10        As Currency                         ' SOLDE VAL. M -9
    SOLDEV11        As Currency                         ' SOLDE VAL. M -10
    SOLDEV12        As Currency                         ' SOLDE VAL. M -11
    
End Type
    
    
Public arrYSOLDE0() As typeYSOLDE0
Public arrYSOLDE0_NB As Integer
Public arrYSOLDE0_NBMax As Integer
Public arrYSOLDE0_Index As Integer
Public arrYSOLDE0_Suite As Boolean


Public Function srvYSOLDE0_Import_Read(lId As String, lYSOLDE0 As typeYSOLDE0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYSOLDE0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYSOLDE0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYSOLDE0_GetBuffer lYSOLDE0
    srvYSOLDE0_Import_Read = Null
Else
    recYSOLDE0_Init lYSOLDE0
    'lYSOLDE0.COMPTECOM = lId
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYSOLDE0_Import_Read" & xIn, vbCritical, Error
srvYSOLDE0_Import_Read = Error
End Function

Public Function srvYSOLDE0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYSOLDE0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYSOLDE0_Import = Null
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If


srvYSOLDE0_Import = "?"

paramYSOLDE0_Import = paramYBase_DataF & Trim(constYSOLDE0) & paramYBase_Data_ExtensionP

Open Trim(paramYSOLDE0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYSOLDE0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYSOLDE0
            meYbase.K1 = mId$(xIn, 10, 20)
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYSOLDE0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYSOLDE0
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYSOLDE0_Import" & xIn, vbCritical, Error
Close

srvYSOLDE0_Import = Error
End Function



'-----------------------------------------------------
Function srvYSOLDE0_Update(recYSOLDE0 As typeYSOLDE0)
'-----------------------------------------------------

srvYSOLDE0_Update = "?"

MsgTxtLen = 0
Call srvYSOLDE0_PutBuffer(recYSOLDE0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYSOLDE0_GetBuffer(recYSOLDE0)) Then
        Call srvYSOLDE0_Error(recYSOLDE0)
        srvYSOLDE0_Update = recYSOLDE0.Err
        Exit Function
    Else
        srvYSOLDE0_Update = Null
    End If
Else
    recYSOLDE0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYSOLDE0_Error(recYSOLDE0 As typeYSOLDE0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YSOLDE0" & Chr$(10) & Chr$(13)

Select Case mId$(recYSOLDE0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYSOLDE0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YSOLDE0s.bas  ( " & Trim(recYSOLDE0.obj) & " : " & Trim(recYSOLDE0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYSOLDE0_Monitor(recYSOLDE0 As typeYSOLDE0)
'-----------------------------------------------------

arrYSOLDE0_Suite = False
Select Case mId$(Trim(recYSOLDE0.Method), 1, 4)
    Case "Snap"
              srvYSOLDE0_Monitor = srvYSOLDE0_Snap(recYSOLDE0)
    Case Else
            srvYSOLDE0_Monitor = srvYSOLDE0_Seek(recYSOLDE0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYSOLDE0_GetBuffer(recYSOLDE0 As typeYSOLDE0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYSOLDE0_GetBuffer = Null
recYSOLDE0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYSOLDE0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYSOLDE0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYSOLDE0.Err = Space$(10) Then
    recYSOLDE0.SOLDEETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYSOLDE0.SOLDEPLA = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYSOLDE0.SOLDECOM = mId$(MsgTxt, K + 10, 20)
    recYSOLDE0.SOLDEDMO = CLng(Val(mId$(MsgTxt, K + 30, 8)))
    recYSOLDE0.SOLDEDAN = CLng(Val(mId$(MsgTxt, K + 38, 8)))
    recYSOLDE0.SOLDECEN = CCur(mId$(MsgTxt, K + 46, 19)) / 1000
    recYSOLDE0.SOLDECAN = CCur(mId$(MsgTxt, K + 65, 19)) / 1000
    recYSOLDE0.SOLDEC01 = CCur(mId$(MsgTxt, K + 84, 19)) / 1000
    recYSOLDE0.SOLDEC02 = CCur(mId$(MsgTxt, K + 103, 19)) / 1000
    recYSOLDE0.SOLDEC03 = CCur(mId$(MsgTxt, K + 122, 19)) / 1000
    recYSOLDE0.SOLDEC04 = CCur(mId$(MsgTxt, K + 141, 19)) / 1000
    recYSOLDE0.SOLDEC05 = CCur(mId$(MsgTxt, K + 160, 19)) / 1000
    recYSOLDE0.SOLDEC06 = CCur(mId$(MsgTxt, K + 179, 19)) / 1000
    recYSOLDE0.SOLDEC07 = CCur(mId$(MsgTxt, K + 198, 19)) / 1000
    recYSOLDE0.SOLDEC08 = CCur(mId$(MsgTxt, K + 217, 19)) / 1000
    recYSOLDE0.SOLDEC09 = CCur(mId$(MsgTxt, K + 236, 19)) / 1000
    recYSOLDE0.SOLDEC10 = CCur(mId$(MsgTxt, K + 255, 19)) / 1000
    recYSOLDE0.SOLDEC11 = CCur(mId$(MsgTxt, K + 274, 19)) / 1000
    recYSOLDE0.SOLDEC12 = CCur(mId$(MsgTxt, K + 293, 19)) / 1000
    recYSOLDE0.SOLDEVEN = CCur(mId$(MsgTxt, K + 312, 19)) / 1000
    recYSOLDE0.SOLDEVAN = CCur(mId$(MsgTxt, K + 331, 19)) / 1000
    recYSOLDE0.SOLDEV01 = CCur(mId$(MsgTxt, K + 350, 19)) / 1000
    recYSOLDE0.SOLDEV02 = CCur(mId$(MsgTxt, K + 369, 19)) / 1000
    recYSOLDE0.SOLDEV03 = CCur(mId$(MsgTxt, K + 388, 19)) / 1000
    recYSOLDE0.SOLDEV04 = CCur(mId$(MsgTxt, K + 407, 19)) / 1000
    recYSOLDE0.SOLDEV05 = CCur(mId$(MsgTxt, K + 426, 19)) / 1000
    recYSOLDE0.SOLDEV06 = CCur(mId$(MsgTxt, K + 445, 19)) / 1000
    recYSOLDE0.SOLDEV07 = CCur(mId$(MsgTxt, K + 464, 19)) / 1000
    recYSOLDE0.SOLDEV08 = CCur(mId$(MsgTxt, K + 483, 19)) / 1000
    recYSOLDE0.SOLDEV09 = CCur(mId$(MsgTxt, K + 502, 19)) / 1000
    recYSOLDE0.SOLDEV10 = CCur(mId$(MsgTxt, K + 521, 19)) / 1000
    recYSOLDE0.SOLDEV11 = CCur(mId$(MsgTxt, K + 540, 19)) / 1000
    recYSOLDE0.SOLDEV12 = CCur(mId$(MsgTxt, K + 559, 19)) / 1000
Else
    srvYSOLDE0_GetBuffer = recYSOLDE0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYSOLDE0Len

End Function

'---------------------------------------------------------
Public Sub srvYSOLDE0_PutBuffer(recYSOLDE0 As typeYSOLDE0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYSOLDE0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYSOLDE0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYSOLDE0.SOLDEETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYSOLDE0.SOLDEPLA, "000 ")
    Mid$(MsgTxt, K + 10, 20) = recYSOLDE0.SOLDECOM
    Mid$(MsgTxt, K + 30, 8) = Format$(recYSOLDE0.SOLDEDMO, "0000000 ")
    Mid$(MsgTxt, K + 38, 8) = Format$(recYSOLDE0.SOLDEDAN, "0000000 ")
    Mid$(MsgTxt, K + 46, 19) = Format$(recYSOLDE0.SOLDECEN * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 65, 19) = Format$(recYSOLDE0.SOLDECAN * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 84, 19) = Format$(recYSOLDE0.SOLDEC01 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 103, 19) = Format$(recYSOLDE0.SOLDEC02 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 122, 19) = Format$(recYSOLDE0.SOLDEC03 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 141, 19) = Format$(recYSOLDE0.SOLDEC04 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 160, 19) = Format$(recYSOLDE0.SOLDEC05 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 179, 19) = Format$(recYSOLDE0.SOLDEC06 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 198, 19) = Format$(recYSOLDE0.SOLDEC07 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 217, 19) = Format$(recYSOLDE0.SOLDEC08 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 236, 19) = Format$(recYSOLDE0.SOLDEC09 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 255, 19) = Format$(recYSOLDE0.SOLDEC10 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 274, 19) = Format$(recYSOLDE0.SOLDEC11 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 293, 19) = Format$(recYSOLDE0.SOLDEC12 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 312, 19) = Format$(recYSOLDE0.SOLDEVEN * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 331, 19) = Format$(recYSOLDE0.SOLDEVAN * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 350, 19) = Format$(recYSOLDE0.SOLDEV01 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 369, 19) = Format$(recYSOLDE0.SOLDEV02 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 388, 19) = Format$(recYSOLDE0.SOLDEV03 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 407, 19) = Format$(recYSOLDE0.SOLDEV04 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 426, 19) = Format$(recYSOLDE0.SOLDEV05 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 445, 19) = Format$(recYSOLDE0.SOLDEV06 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 464, 19) = Format$(recYSOLDE0.SOLDEV07 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 483, 19) = Format$(recYSOLDE0.SOLDEV08 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 502, 19) = Format$(recYSOLDE0.SOLDEV09 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 521, 19) = Format$(recYSOLDE0.SOLDEV10 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 540, 19) = Format$(recYSOLDE0.SOLDEV11 * 1000, "000000000000000000 ")
    Mid$(MsgTxt, K + 559, 19) = Format$(recYSOLDE0.SOLDEV12 * 1000, "000000000000000000 ")
MsgTxtLen = MsgTxtLen + recYSOLDE0Len
End Sub



'---------------------------------------------------------
Private Function srvYSOLDE0_Seek(recYSOLDE0 As typeYSOLDE0)
'---------------------------------------------------------

srvYSOLDE0_Seek = "?"
MsgTxtLen = 0
Call srvYSOLDE0_PutBuffer(recYSOLDE0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYSOLDE0_GetBuffer(recYSOLDE0)) Then
        srvYSOLDE0_Seek = Null
    Else
        Call srvYSOLDE0_Error(recYSOLDE0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYSOLDE0_Snap(recYSOLDE0 As typeYSOLDE0)
'---------------------------------------------------------
srvYSOLDE0_Snap = "?"
MsgTxtLen = 0
Call srvYSOLDE0_PutBuffer(recYSOLDE0)
Call srvYSOLDE0_PutBuffer(arrYSOLDE0(0))
If IsNull(SndRcv()) Then
    srvYSOLDE0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYSOLDE0_GetBuffer(recYSOLDE0)) Then
            Call arrYSOLDE0_AddItem(recYSOLDE0)
            arrYSOLDE0_Suite = True
        Else
            arrYSOLDE0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
Public Sub srvYSOLDE0_ElpDisplay(recYSOLDE0 As typeYSOLDE0)
frmElpDisplay.fgData.Rows = 34
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEPLA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDECOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDECOM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEDMO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DERNIER MVT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEDMO
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEDAN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ANTERIEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEDAN
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDECEN 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE ENCOURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDECEN
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDECAN 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE ANTERIEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDECAN
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC01 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC01
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC02 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC02
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC03 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC03
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC04 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC04
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC05 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC05
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC06 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -5"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC06
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC07 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -6"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC07
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC08 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -7"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC08
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC09 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -8"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC09
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC10 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -9"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC10
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC11 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -10"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC11
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEC12 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE M -11"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEC12
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEVEN 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. ENCOURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEVEN
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEVAN 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. ANTERIEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEVAN
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV01 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV01
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV02 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV02
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV03 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV03
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV04 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV04
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV05 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV05
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV06 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -5"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV06
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV07 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -6"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV07
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV08 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -7"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV08
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV09 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -8"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV09
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV10 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -9"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV10
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV11 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -10"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV11
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEV12 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE VAL. M -11"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSOLDE0.SOLDEV12
frmElpDisplay.Show vbModal
End Sub


Public Sub srvYSOLDE0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #lIdFile_Destination, "SOLDEETA;SOLDEPLA;SOLDECOM;SOLDEDMO;SOLDEDAN;SOLDECEN;SOLDECAN;SOLDEC01;SOLDEC02;SOLDEC03;SOLDEC04;SOLDEC05;SOLDEC06;SOLDEC07;SOLDEC08;SOLDEC09;SOLDEC10;SOLDEC11;SOLDEC12;SOLDEVEN;SOLDEVAN;SOLDEV01;SOLDEV02;SOLDEV03;SOLDEV04;SOLDEV05;SOLDEV06;SOLDEV07;SOLDEV08;SOLDEV09;SOLDEV10;SOLDEV11;SOLDEV12;"
    Print #lIdFile_Destination, "ETABLISSEMENT;NUMERO PLAN;NUMERO COMPTE;DATE DERNIER MVT;DATE ANTERIEUR;SOLDE ENCOURS;SOLDE ANTERIEUR;SOLDE M;SOLDE M -1;SOLDE M -2;SOLDE M -3;SOLDE M -4;SOLDE M -5;SOLDE M -6;SOLDE M -7;SOLDE M -8;SOLDE M -9;SOLDE M -10;SOLDE M -11;SOLDE VAL. ENCOURS;SOLDE VAL. ANTERIEUR;SOLDE VAL. M;SOLDE VAL. M -1;SOLDE VAL. M -2;SOLDE VAL. M -3;SOLDE VAL. M -4;SOLDE VAL. M -5;SOLDE VAL. M -6;SOLDE VAL. M -7;SOLDE VAL. M -8;SOLDE VAL. M -9;SOLDE VAL. M -10;SOLDE VAL. M -11;"
    Print #lIdFile_Destination, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 4) & ";" & mId$(xIn, 10, 20) & ";" _
      & Val(mId$(xIn, 30, 8)) + 19000000 & ";" & Val(mId$(xIn, 38, 8)) + 19000000 & ";" _
      & cur_19V(CCur(mId$(xIn, 46, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 65, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 84, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 103, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 122, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 141, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 160, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 179, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 198, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 217, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 236, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 255, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 274, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 293, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 312, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 331, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 350, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 369, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 388, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 407, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 426, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 445, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 464, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 483, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 502, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 521, 19)) / 1000) & ";" & cur_19V(CCur(mId$(xIn, 540, 19)) / 1000) & ";" _
      & cur_19V(CCur(mId$(xIn, 559, 19)) / 1000) & ";"
Loop
End Sub

'---------------------------------------------------------
Public Sub arrYSOLDE0_AddItem(recYSOLDE0 As typeYSOLDE0)
'---------------------------------------------------------
          
arrYSOLDE0_NB = arrYSOLDE0_NB + 1
    
If arrYSOLDE0_NB > arrYSOLDE0_NBMax Then
    arrYSOLDE0_NBMax = arrYSOLDE0_NBMax + recYSOLDE0_Block
    ReDim Preserve arrYSOLDE0(arrYSOLDE0_NBMax)
End If
            
arrYSOLDE0(arrYSOLDE0_NB) = recYSOLDE0
End Sub



'---------------------------------------------------------
Public Sub recYSOLDE0_Init(recYSOLDE0 As typeYSOLDE0)
'---------------------------------------------------------
recYSOLDE0.obj = "ZSOLDE0_S"
recYSOLDE0.Method = ""
recYSOLDE0.Err = ""
recYSOLDE0.SOLDEETA = 0 'Integer                        ' ETABLISSEMENT
recYSOLDE0.SOLDEPLA = 0 'Long                           ' NUMERO PLAN
recYSOLDE0.SOLDECOM = "" 'String * 20                    ' NUMERO COMPTE"
recYSOLDE0.SOLDEDMO = 0 'Long                           ' DATE DERNIER MVT
recYSOLDE0.SOLDEDAN = 0 'Long                           ' DATE ANTERIEUR
recYSOLDE0.SOLDECEN = 0 'Currency                         ' SOLDE ENCOURS
recYSOLDE0.SOLDECAN = 0 'Currency                         ' SOLDE ANTERIEUR
recYSOLDE0.SOLDEC01 = 0 'Currency                         ' SOLDE M
recYSOLDE0.SOLDEC02 = 0 'Currency                         ' SOLDE M -1
recYSOLDE0.SOLDEC03 = 0 'Currency                         ' SOLDE M -2
recYSOLDE0.SOLDEC04 = 0 'Currency                         ' SOLDE M -3
recYSOLDE0.SOLDEC05 = 0 'Currency                         ' SOLDE M -4
recYSOLDE0.SOLDEC06 = 0 'Currency                         ' SOLDE M -5
recYSOLDE0.SOLDEC07 = 0 'Currency                         ' SOLDE M -6
recYSOLDE0.SOLDEC08 = 0 'Currency                         ' SOLDE M -7
recYSOLDE0.SOLDEC09 = 0 'Currency                         ' SOLDE M -8
recYSOLDE0.SOLDEC10 = 0 'Currency                         ' SOLDE M -9
recYSOLDE0.SOLDEC11 = 0 'Currency                         ' SOLDE M -10
recYSOLDE0.SOLDEC12 = 0 'Currency                         ' SOLDE M -11
recYSOLDE0.SOLDEVEN = 0 'Currency                         ' SOLDE VAL. ENCOURS
recYSOLDE0.SOLDEVAN = 0 'Currency                         ' SOLDE VAL. ANTERIEUR
recYSOLDE0.SOLDEV01 = 0 'Currency                         ' SOLDE VAL. M
recYSOLDE0.SOLDEV02 = 0 'Currency                         ' SOLDE VAL. M -1
recYSOLDE0.SOLDEV03 = 0 'Currency                         ' SOLDE VAL. M -2
recYSOLDE0.SOLDEV04 = 0 'Currency                         ' SOLDE VAL. M -3
recYSOLDE0.SOLDEV05 = 0 'Currency                         ' SOLDE VAL. M -4
recYSOLDE0.SOLDEV06 = 0 'Currency                         ' SOLDE VAL. M -5
recYSOLDE0.SOLDEV07 = 0 'Currency                         ' SOLDE VAL. M -6
recYSOLDE0.SOLDEV08 = 0 'Currency                         ' SOLDE VAL. M -7
recYSOLDE0.SOLDEV09 = 0 'Currency                         ' SOLDE VAL. M -8
recYSOLDE0.SOLDEV10 = 0 'Currency                         ' SOLDE VAL. M -9
recYSOLDE0.SOLDEV11 = 0 'Currency                         ' SOLDE VAL. M -10
recYSOLDE0.SOLDEV12 = 0 'Currency                         ' SOLDE VAL. M -11

End Sub







