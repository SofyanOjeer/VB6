Attribute VB_Name = "srvYCLIREF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCLIREF0Len = 255 ' 34 +221
Public Const recYCLIREF0_Block = 50
Public Const memoYCLIREF0Len = 221
Public Const constYCLIREF0 = "YCLIREF0  "
Public paramYCLIREF0_Import As String

Type typeYCLIREF0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CLIREFETA       As Integer                        ' ETABLISSEMENT
    CLIREFCLI       As String * 7                     ' NUMERO CLIENT
    CLIREFCOR       As String * 2                     ' CODE REFERENCE
    CLIREFREF       As String * 15                    ' REFERENCE CLIENT
End Type
    
    
Public arrYCLIREF0() As typeYCLIREF0
Public arrYCLIREF0_NB As Integer
Public arrYCLIREF0_NBMax As Integer
Public arrYCLIREF0_Index As Integer
Public arrYCLIREF0_Suite As Boolean

Public Function srvYCLIREF0_Import(lnb As Long)
Dim xIn As String, x As String
Dim meMVTP0 As typeMvtP0

On Error GoTo Error_Handle


srvYCLIREF0_Import = "?"

paramYCLIREF0_Import = paramYBase_DataF & Trim(constYCLIREF0) & paramYBase_Data_ExtensionP

Open Trim(paramYCLIREF0_Import) For Input As #1
lnb = 0

recMvtP0_Init meMVTP0
meMVTP0.Method = constAddNew

mdbMvtP0.tableMvtP0_Open

Do Until EOF(1)
    lnb = lnb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meMVTP0.Id = constYCLIREF0 & mId$(xIn, 6, 7)
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
            
    End If
        
Loop


Close
srvYCLIREF0_Import = Null
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCLIREF0_Import" & xIn, vbCritical, Error
Close

srvYCLIREF0_Import = Error
End Function



'-----------------------------------------------------
Function srvYCLIREF0_Update(recYCLIREF0 As typeYCLIREF0)
'-----------------------------------------------------

srvYCLIREF0_Update = "?"

MsgTxtLen = 0
Call srvYCLIREF0_PutBuffer(recYCLIREF0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCLIREF0_GetBuffer(recYCLIREF0)) Then
        Call srvYCLIREF0_Error(recYCLIREF0)
        srvYCLIREF0_Update = recYCLIREF0.Err
        Exit Function
    Else
        srvYCLIREF0_Update = Null
    End If
Else
    recYCLIREF0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYCLIREF0_Error(recYCLIREF0 As typeYCLIREF0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCLIREF0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCLIREF0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCLIREF0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YCLIREF0s.bas  ( " & Trim(recYCLIREF0.obj) & " : " & Trim(recYCLIREF0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYCLIREF0_Monitor(recYCLIREF0 As typeYCLIREF0)
'-----------------------------------------------------

arrYCLIREF0_Suite = False
Select Case mId$(Trim(recYCLIREF0.Method), 1, 4)
    Case "Snap"
              srvYCLIREF0_Monitor = srvYCLIREF0_Snap(recYCLIREF0)
    Case Else
            srvYCLIREF0_Monitor = srvYCLIREF0_Seek(recYCLIREF0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYCLIREF0_GetBuffer(recYCLIREF0 As typeYCLIREF0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCLIREF0_GetBuffer = Null
recYCLIREF0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCLIREF0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCLIREF0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCLIREF0.Err = Space$(10) Then
    recYCLIREF0.CLIREFETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCLIREF0.CLIREFCLI = mId$(MsgTxt, K + 6, 7)
    recYCLIREF0.CLIREFCOR = mId$(MsgTxt, K + 13, 2)
    recYCLIREF0.CLIREFREF = mId$(MsgTxt, K + 15, 15)

Else
    srvYCLIREF0_GetBuffer = recYCLIREF0.Err
    recYCLIREF0.CLIREFCOR = "?"
    recYCLIREF0.CLIREFREF = "? cliref"
End If
MsgTxtIndex = MsgTxtIndex + recYCLIREF0Len

End Function

Public Sub srvYCLIREF0_ElpDisplay(recYCLIREF0 As typeYCLIREF0)
frmElpDisplay.fgData.Rows = 5
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIREFETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIREF0.CLIREFETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIREFCLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIREF0.CLIREFCLI
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIREFCOR    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE REFERENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIREF0.CLIREFCOR
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIREFREF   15A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCE CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIREF0.CLIREFREF
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Public Sub srvYCLIREF0_PutBuffer(recYCLIREF0 As typeYCLIREF0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCLIREF0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCLIREF0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCLIREF0.CLIREFETA, "0000 ")
    Mid$(MsgTxt, K + 6, 7) = recYCLIREF0.CLIREFCLI
    Mid$(MsgTxt, K + 13, 2) = recYCLIREF0.CLIREFCOR
    Mid$(MsgTxt, K + 15, 15) = recYCLIREF0.CLIREFREF

MsgTxtLen = MsgTxtLen + recYCLIREF0Len
End Sub



'---------------------------------------------------------
Private Function srvYCLIREF0_Seek(recYCLIREF0 As typeYCLIREF0)
'---------------------------------------------------------

srvYCLIREF0_Seek = "?"
MsgTxtLen = 0
Call srvYCLIREF0_PutBuffer(recYCLIREF0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYCLIREF0_GetBuffer(recYCLIREF0)) Then
        srvYCLIREF0_Seek = Null
    Else
        Call srvYCLIREF0_Error(recYCLIREF0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYCLIREF0_Snap(recYCLIREF0 As typeYCLIREF0)
'---------------------------------------------------------
srvYCLIREF0_Snap = "?"
MsgTxtLen = 0
Call srvYCLIREF0_PutBuffer(recYCLIREF0)
Call srvYCLIREF0_PutBuffer(arrYCLIREF0(0))
If IsNull(SndRcv()) Then
    srvYCLIREF0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYCLIREF0_GetBuffer(recYCLIREF0)) Then
            Call arrYCLIREF0_AddItem(recYCLIREF0)
            arrYCLIREF0_Suite = True
        Else
            arrYCLIREF0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYCLIREF0_AddItem(recYCLIREF0 As typeYCLIREF0)
'---------------------------------------------------------
          
arrYCLIREF0_NB = arrYCLIREF0_NB + 1
    
If arrYCLIREF0_NB > arrYCLIREF0_NBMax Then
    arrYCLIREF0_NBMax = arrYCLIREF0_NBMax + recYCLIREF0_Block
    ReDim Preserve arrYCLIREF0(arrYCLIREF0_NBMax)
End If
            
arrYCLIREF0(arrYCLIREF0_NB) = recYCLIREF0
End Sub



'---------------------------------------------------------
Public Sub recYCLIREF0_Init(recYCLIREF0 As typeYCLIREF0)
'---------------------------------------------------------
recYCLIREF0.obj = "ZCLIREF0_S"
recYCLIREF0.Method = ""
recYCLIREF0.Err = ""

End Sub








