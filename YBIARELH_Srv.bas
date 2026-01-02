Attribute VB_Name = "srvYBIARELH"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYBIARELHLen = 157 ' 34 +123
Public Const recYBIARELH_Block = 100
Public Const memoYBIARELHLen = 123
Public Const constYBIARELH = "YBIARELH"
Public paramYBIARELH_Import As String
Dim meYbase As typeYBase

Type typeYBIARELH
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    BIARELCOM       As String * 20                    ' NUMERO COMPTE
    BIARELREL       As String * 1                     '
    BIARELID        As Long                           '
    BIARELNUM       As Long                           '
    BIARELSD0       As String * 20                    '
    BIARELD0       As String * 8                    '
    BIAMVTID0       As Long                           '
    BIARELSD1       As Currency                       '
    BIARELD1        As String * 8                    '
    BIAMVTID1       As Long                           '
    BIAOLDCOM       As String * 11                    '
    BIAOLDDEV       As String * 3                    '

End Type
    
    
Public arrYBIARELH() As typeYBIARELH
Public arrYBIARELH_NB As Integer
Public arrYBIARELH_NBMax As Integer
Public arrYBIARELH_Index As Integer
Public arrYBIARELH_Suite As Boolean

'-----------------------------------------------------
Function srvYBIARELH_Update(recYBIARELH As typeYBIARELH)
'-----------------------------------------------------

srvYBIARELH_Update = "?"
recYBIARELH.obj = "YBIARELH_S"
MsgTxtLen = 0
Call srvYBIARELH_PutBuffer(recYBIARELH)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYBIARELH_GetBuffer(recYBIARELH)) Then
        Call srvYBIARELH_Error(recYBIARELH)
        srvYBIARELH_Update = recYBIARELH.Err
        Exit Function
    Else
        srvYBIARELH_Update = Null
    End If
Else
    recYBIARELH.Err = "srv"
End If
End Function

Public Function srvYBIARELH_Import_old(lnb As Long)
Dim paramYBIARELH_Import As String
Dim xIn As String, X As String
Dim meMVTP0 As typeMvtP0

On Error GoTo Error_Handle

srvYBIARELH_Import_old = "?"

paramYBIARELH_Import = paramYBase_DataF & Trim(constYBIARELH) & paramYBase_Data_ExtensionP

Open Trim(paramYBIARELH_Import) For Input As #1

lnb = 0

recMvtP0_Init meMVTP0
meMVTP0.Method = constAddNew

mdbMvtP0.tableMvtP0_Open

Do Until EOF(1)
    lnb = lnb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If mId$(xIn, 22, 6) = "00000 " Then                 ' en tête
            meMVTP0.ID = constYBIARELH & mId$(xIn, 1, 21)   ' code et compte
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
        End If
            
    End If
Loop


Close
srvYBIARELH_Import_old = Null
Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIARELH_Import" & xIn, vbCritical, Error
Close

srvYBIARELH_Import_old = Error
End Function

Public Function srvYBIARELH_Import_Read_Old(lYBIARELH As typeYBIARELH)
Dim meMVTP0 As typeMvtP0

srvYBIARELH_Import_Read_Old = Null
meMVTP0.ID = constYBIARELH & lYBIARELH.BIARELCOM & lYBIARELH.BIARELREL
meMVTP0.Method = "Seek="
If tableMvtP0_Read(meMVTP0) = 0 Then
    MsgTxt = Space$(34) & meMVTP0.Text
    MsgTxtIndex = 0
                
    srvYBIARELH_GetBuffer lYBIARELH


Else
    lYBIARELH.Err = 9998
    srvYBIARELH_Import_Read_Old = "? " & meMVTP0.ID
End If

End Function

'-----------------------------------------------------
Sub srvYBIARELH_Error(recYBIARELH As typeYBIARELH)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YBIARELH" & Chr$(10) & Chr$(13)

Select Case mId$(recYBIARELH.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYBIARELH.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YBIARELHs.bas  ( " & Trim(recYBIARELH.obj) & " : " & Trim(recYBIARELH.Method) & " )"

End Sub


Public Function srvYBIARELH_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle


recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYBIARELH
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYBIARELH_Import = Null
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If




srvYBIARELH_Import = "?"

paramYBIARELH_Import = paramYBase_DataF & Trim(constYBIARELH) & paramYBase_Data_ExtensionP

Open Trim(paramYBIARELH_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYBIARELH) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If mId$(xIn, 22, 6) = "00000 " Then                 ' en tête
            meYbase.ID = constYBIARELH
            meYbase.K1 = mId$(xIn, 1, 21)   '   compte code
            meYbase.Text = xIn
            dbYBase_Update meYbase
        End If
    End If
Loop


Close
srvYBIARELH_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYBIARELH
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIARELH_Import" & xIn, vbCritical, Error
Close

srvYBIARELH_Import = Error
End Function

'=====================================================



'-----------------------------------------------------
Public Function srvYBIARELH_Monitor(recYBIARELH As typeYBIARELH)
'-----------------------------------------------------

arrYBIARELH_Suite = False
Select Case mId$(Trim(recYBIARELH.Method), 1, 4)
    Case "Snap"
              srvYBIARELH_Monitor = srvYBIARELH_Snap(recYBIARELH)
    Case Else
            srvYBIARELH_Monitor = srvYBIARELH_Seek(recYBIARELH)
End Select

End Function
Public Function srvYBIARELH_Import_Read(lId As String, lYBIARELH As typeYBIARELH)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYBIARELH_Import_Read = "?"

meYbase.Method = "Seek>="
meYbase.ID = constYBIARELH
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    If Trim(mId$(meYbase.K1, 1, 20)) = Trim(mId$(lId, 1, 20)) Then
        MsgTxt = Space$(34) & meYbase.Text
        MsgTxtIndex = 0
        srvYBIARELH_GetBuffer lYBIARELH
        srvYBIARELH_Import_Read = Null
    End If
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIARELH_Import_Read" & xIn, vbCritical, Error
srvYBIARELH_Import_Read = Error
End Function



'---------------------------------------------------------
Public Function srvYBIARELH_GetBuffer(recYBIARELH As typeYBIARELH)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYBIARELH_GetBuffer = Null
recYBIARELH.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYBIARELH.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYBIARELH.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYBIARELH.Err = Space$(10) Then
    recYBIARELH.BIARELCOM = mId$(MsgTxt, K + 1, 20)
    recYBIARELH.BIARELREL = mId$(MsgTxt, K + 21, 1)
    recYBIARELH.BIARELID = CLng(Val(mId$(MsgTxt, K + 22, 6)))
    recYBIARELH.BIARELNUM = CLng(Val(mId$(MsgTxt, K + 28, 6)))
    recYBIARELH.BIARELSD0 = CCur(mId$(MsgTxt, K + 34, 19) / 1000)
    recYBIARELH.BIARELD0 = mId$(MsgTxt, K + 53, 8)
    recYBIARELH.BIAMVTID0 = CLng(Val(mId$(MsgTxt, K + 61, 11)))
    recYBIARELH.BIARELSD1 = CCur(mId$(MsgTxt, K + 72, 19) / 1000)
    recYBIARELH.BIARELD1 = mId$(MsgTxt, K + 91, 8)
    recYBIARELH.BIAMVTID1 = CLng(Val(mId$(MsgTxt, K + 99, 11)))
    recYBIARELH.BIAOLDCOM = mId$(MsgTxt, K + 110, 11)
    recYBIARELH.BIAOLDDEV = mId$(MsgTxt, K + 121, 3)
   
Else
    srvYBIARELH_GetBuffer = recYBIARELH.Err
End If

MsgTxtIndex = MsgTxtIndex + recYBIARELHLen

End Function

'---------------------------------------------------------
Public Sub srvYBIARELH_PutBuffer(recYBIARELH As typeYBIARELH)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYBIARELH.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYBIARELH.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 20) = recYBIARELH.BIARELCOM
    Mid$(MsgTxt, K + 21, 1) = recYBIARELH.BIARELREL
    Mid$(MsgTxt, K + 22, 6) = Format$(recYBIARELH.BIARELID, "00000 ")
    Mid$(MsgTxt, K + 28, 6) = Format$(recYBIARELH.BIARELNUM, "00000 ")
    Mid$(MsgTxt, K + 34, 18) = Format$(Abs(recYBIARELH.BIARELSD0) * 1000, "000000000000000000")
    Mid$(MsgTxt, K + 52, 1) = IIf(recYBIARELH.BIARELSD0 > 0, "-", " ")
    Mid$(MsgTxt, K + 53, 8) = Format$(recYBIARELH.BIARELD0, "00000000")
    Mid$(MsgTxt, K + 61, 11) = Format$(recYBIARELH.BIAMVTID0, "0000000000 ")
    Mid$(MsgTxt, K + 72, 18) = Format$(Abs(recYBIARELH.BIARELSD1) * 1000, "000000000000000000")
    Mid$(MsgTxt, K + 90, 1) = IIf(recYBIARELH.BIARELSD1 > 0, "-", " ")
    Mid$(MsgTxt, K + 91, 8) = Format$(recYBIARELH.BIARELD1, "00000000")
    Mid$(MsgTxt, K + 99, 11) = Format$(recYBIARELH.BIAMVTID1, "0000000000 ")
    Mid$(MsgTxt, K + 110, 11) = recYBIARELH.BIAOLDCOM
    Mid$(MsgTxt, K + 121, 3) = recYBIARELH.BIAOLDDEV
    
MsgTxtLen = MsgTxtLen + recYBIARELHLen
End Sub



Private Function srvYBIARELH_Seek(recYBIARELH As typeYBIARELH)
'---------------------------------------------------------

srvYBIARELH_Seek = "?"
MsgTxtLen = 0
Call srvYBIARELH_PutBuffer(recYBIARELH)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYBIARELH_GetBuffer(recYBIARELH)) Then
        srvYBIARELH_Seek = Null
    Else
        Call srvYBIARELH_Error(recYBIARELH)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYBIARELH_Snap(recYBIARELH As typeYBIARELH)
'---------------------------------------------------------
srvYBIARELH_Snap = "?"
MsgTxtLen = 0
Call srvYBIARELH_PutBuffer(recYBIARELH)
Call srvYBIARELH_PutBuffer(arrYBIARELH(0))
If IsNull(SndRcv()) Then
    srvYBIARELH_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYBIARELH_GetBuffer(recYBIARELH)) Then
            Call arrYBIARELH_AddItem(recYBIARELH)
            arrYBIARELH_Suite = True
        Else
            arrYBIARELH_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYBIARELH_AddItem(recYBIARELH As typeYBIARELH)
'---------------------------------------------------------
          
arrYBIARELH_NB = arrYBIARELH_NB + 1
    
If arrYBIARELH_NB > arrYBIARELH_NBMax Then
    arrYBIARELH_NBMax = arrYBIARELH_NBMax + recYBIARELH_Block
    ReDim Preserve arrYBIARELH(arrYBIARELH_NBMax)
End If
            
arrYBIARELH(arrYBIARELH_NB) = recYBIARELH
End Sub



'---------------------------------------------------------
Public Sub recYBIARELH_Init(recYBIARELH As typeYBIARELH)
'---------------------------------------------------------
recYBIARELH.obj = "YBIARELH_S"
recYBIARELH.Method = ""
recYBIARELH.Err = ""
recYBIARELH.BIARELCOM = ""        ' As String * 20                    ' NUMERO COMPTE
recYBIARELH.BIARELREL = ""        ' As String * 1                     '
recYBIARELH.BIARELID = 0         '  As Long                           '
recYBIARELH.BIARELNUM = 0        ' As Long                           '
recYBIARELH.BIARELSD0 = ""        ' As String * 20                    '
recYBIARELH.BIARELD0 = ""        ' As String * 8                    '
recYBIARELH.BIAMVTID0 = 0        ' As Long                           '
recYBIARELH.BIARELSD1 = 0        ' As Currency                       '
recYBIARELH.BIARELD1 = ""        '  As String * 8                    '
recYBIARELH.BIAMVTID1 = 0        ' As Long                           '
recYBIARELH.BIAOLDCOM = ""        ' As String * 11                    '
recYBIARELH.BIAOLDDEV = ""        ' As String * 3                    '

End Sub









