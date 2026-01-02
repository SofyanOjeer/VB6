Attribute VB_Name = "srvYBIAMON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYBIAMON0Len = 130 ' 34 +96
Public Const recYBIAMON0_Block = 100
Public Const memoYBIAMON0Len = 96
Public Const constYBIAMON0 = "YBIAMON0  "

Type typeYBIAMON0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MONAPP          As String * 10
    MONFLUX         As String * 10
    MONSTATUS       As String * 10
    MONNUM          As Long
    MONJOB          As String * 10
    MONPGM          As String * 10
    MONUSR          As String * 10
    MONAMJ          As String * 8
    MONHMS          As String * 8
    MONFILE         As String * 10

End Type
    
    
Public arrYBIAMON0() As typeYBIAMON0
Public arrYBIAMON0_NB As Integer
Public arrYBIAMON0_NBMax As Integer
Public arrYBIAMON0_Index As Integer
Public arrYBIAMON0_Suite As Boolean

'-----------------------------------------------------
Function srvYBIAMON0_Update(recYBIAMON0 As typeYBIAMON0)
'-----------------------------------------------------

srvYBIAMON0_Update = "?"
recYBIAMON0.obj = "YBIAMON0_S"
MsgTxtLen = 0
Call srvYBIAMON0_PutBuffer(recYBIAMON0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYBIAMON0_GetBuffer(recYBIAMON0)) Then
        Call srvYBIAMON0_Error(recYBIAMON0)
        srvYBIAMON0_Update = recYBIAMON0.Err
        Exit Function
    Else
        srvYBIAMON0_Update = Null
    End If
Else
    recYBIAMON0.Err = "srv"
End If
End Function

Public Function srvYBIAMON0_Import(lnb As Long)
Dim paramYBIAMON0_Import As String
Dim xIn As String, X As String
Dim meMVTP0 As typeMvtP0

On Error GoTo Error_Handle

srvYBIAMON0_Import = "?"

paramYBIAMON0_Import = paramYBase_DataF & Trim(constYBIAMON0) & paramYBase_Data_ExtensionP

Open Trim(paramYBIAMON0_Import) For Input As #1

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
            meMVTP0.ID = constYBIAMON0 & mId$(xIn, 1, 21)   ' code et compte
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
        End If
            
    End If
Loop


Close
srvYBIAMON0_Import = Null
Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIAMON0_Import" & xIn, vbCritical, Error
Close

srvYBIAMON0_Import = Error
End Function

Public Function srvYBIAMON0_R_Mdb(lYBIAMON0 As typeYBIAMON0)
Dim meMVTP0 As typeMvtP0

srvYBIAMON0_R_Mdb = Null
meMVTP0.ID = constYBIAMON0 & lYBIAMON0.MONAPP & lYBIAMON0.MONFLUX
meMVTP0.Method = "Seek="
If tableMvtP0_Read(meMVTP0) = 0 Then
    MsgTxt = Space$(34) & meMVTP0.Text
    MsgTxtIndex = 0
                
    srvYBIAMON0_GetBuffer lYBIAMON0


Else
    lYBIAMON0.Err = 9998
    srvYBIAMON0_R_Mdb = "? " & meMVTP0.ID
End If

End Function

'-----------------------------------------------------
Sub srvYBIAMON0_Error(recYBIAMON0 As typeYBIAMON0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YBIAMON0" & Chr$(10) & Chr$(13)

Select Case mId$(recYBIAMON0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYBIAMON0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YBIAMON0s.bas  ( " & Trim(recYBIAMON0.obj) & " : " & Trim(recYBIAMON0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYBIAMON0_Monitor(recYBIAMON0 As typeYBIAMON0)
'-----------------------------------------------------

arrYBIAMON0_Suite = False
Select Case mId$(Trim(recYBIAMON0.Method), 1, 4)
    Case "Snap"
              srvYBIAMON0_Monitor = srvYBIAMON0_Snap(recYBIAMON0)
    Case Else
            srvYBIAMON0_Monitor = srvYBIAMON0_Seek(recYBIAMON0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYBIAMON0_GetBuffer(recYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYBIAMON0_GetBuffer = Null
recYBIAMON0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYBIAMON0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYBIAMON0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYBIAMON0.Err = Space$(10) Then
    recYBIAMON0.MONAPP = mId$(MsgTxt, K + 1, 10)
    recYBIAMON0.MONFLUX = mId$(MsgTxt, K + 11, 10)
    recYBIAMON0.MONSTATUS = mId$(MsgTxt, K + 21, 10)
    recYBIAMON0.MONNUM = CLng(Val(mId$(MsgTxt, K + 31, 10)))
    recYBIAMON0.MONJOB = mId$(MsgTxt, K + 41, 10)
    recYBIAMON0.MONPGM = mId$(MsgTxt, K + 51, 10)
    recYBIAMON0.MONUSR = mId$(MsgTxt, K + 61, 10)
    recYBIAMON0.MONAMJ = mId$(MsgTxt, K + 71, 8)
    recYBIAMON0.MONHMS = mId$(MsgTxt, K + 79, 8)
    recYBIAMON0.MONFILE = mId$(MsgTxt, K + 87, 10)
   
Else
    srvYBIAMON0_GetBuffer = recYBIAMON0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYBIAMON0Len

End Function

'---------------------------------------------------------
Public Sub srvYBIAMON0_PutBuffer(recYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYBIAMON0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYBIAMON0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 10) = recYBIAMON0.MONAPP
    Mid$(MsgTxt, K + 11, 10) = recYBIAMON0.MONFLUX
    Mid$(MsgTxt, K + 21, 10) = recYBIAMON0.MONSTATUS
    Mid$(MsgTxt, K + 31, 10) = Format$(recYBIAMON0.MONNUM, "0000000000")
    Mid$(MsgTxt, K + 41, 10) = recYBIAMON0.MONJOB
    Mid$(MsgTxt, K + 51, 10) = recYBIAMON0.MONPGM
    Mid$(MsgTxt, K + 61, 10) = recYBIAMON0.MONUSR
    Mid$(MsgTxt, K + 71, 8) = Format$(recYBIAMON0.MONAMJ, "00000000")
    Mid$(MsgTxt, K + 79, 8) = Format$(recYBIAMON0.MONHMS, "000000") & "00"
    Mid$(MsgTxt, K + 87, 10) = recYBIAMON0.MONFILE
    
MsgTxtLen = MsgTxtLen + recYBIAMON0Len
End Sub



Private Function srvYBIAMON0_Seek(recYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------

srvYBIAMON0_Seek = "?"
MsgTxtLen = 0
Call srvYBIAMON0_PutBuffer(recYBIAMON0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYBIAMON0_GetBuffer(recYBIAMON0)) Then
        srvYBIAMON0_Seek = Null
    Else
        Call srvYBIAMON0_Error(recYBIAMON0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYBIAMON0_Snap(recYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
srvYBIAMON0_Snap = "?"
MsgTxtLen = 0
Call srvYBIAMON0_PutBuffer(recYBIAMON0)
Call srvYBIAMON0_PutBuffer(arrYBIAMON0(0))
If IsNull(SndRcv()) Then
    srvYBIAMON0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYBIAMON0_GetBuffer(recYBIAMON0)) Then
            Call arrYBIAMON0_AddItem(recYBIAMON0)
            arrYBIAMON0_Suite = True
        Else
            arrYBIAMON0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYBIAMON0_AddItem(recYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
          
arrYBIAMON0_NB = arrYBIAMON0_NB + 1
    
If arrYBIAMON0_NB > arrYBIAMON0_NBMax Then
    arrYBIAMON0_NBMax = arrYBIAMON0_NBMax + recYBIAMON0_Block
    ReDim Preserve arrYBIAMON0(arrYBIAMON0_NBMax)
End If
            
arrYBIAMON0(arrYBIAMON0_NB) = recYBIAMON0
End Sub



'---------------------------------------------------------
Public Sub recYBIAMON0_Init(recYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
recYBIAMON0.obj = "YBIAMON0_S"
recYBIAMON0.Method = ""
recYBIAMON0.Err = ""
recYBIAMON0.MONAPP = ""
recYBIAMON0.MONFLUX = ""
recYBIAMON0.MONSTATUS = ""
recYBIAMON0.MONNUM = 0
recYBIAMON0.MONJOB = ""
recYBIAMON0.MONPGM = ""
recYBIAMON0.MONUSR = ""
recYBIAMON0.MONAMJ = "00000000"
recYBIAMON0.MONHMS = "00000000"
recYBIAMON0.MONFILE = ""
End Sub










