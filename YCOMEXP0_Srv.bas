Attribute VB_Name = "srvYCOMEXP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCOMEXP0Len = 554 ' 34 +520
Public Const recYCOMEXP0_Block = 50
Public Const memoYCOMEXP0Len = 520

Type typeYCOMEXP0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    COMEXPETA       As Integer                        ' ETABLISSEMENT
    COMEXPTRA       As String * 6                     ' CODE TRAITEMENT
    COMEXPOPT       As String * 3                     ' CODE OPTION
    COMEXPARG       As String * 12                    ' ARGUMENT
    COMEXPDON       As String * 100                   ' DONNEE
End Type
    
    
Public arrYCOMEXP0() As typeYCOMEXP0
Public arrYCOMEXP0_NB As Integer
Public arrYCOMEXP0_NBMax As Integer
Public arrYCOMEXP0_Index As Integer
Public arrYCOMEXP0_Suite As Boolean

'-----------------------------------------------------
Function srvYCOMEXP0_Update(recYCOMEXP0 As typeYCOMEXP0)
'-----------------------------------------------------

srvYCOMEXP0_Update = "?"

MsgTxtLen = 0
Call srvYCOMEXP0_PutBuffer(recYCOMEXP0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCOMEXP0_GetBuffer(recYCOMEXP0)) Then
        Call srvYCOMEXP0_Error(recYCOMEXP0)
        srvYCOMEXP0_Update = recYCOMEXP0.Err
        Exit Function
    Else
        srvYCOMEXP0_Update = Null
    End If
Else
    recYCOMEXP0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYCOMEXP0_Error(recYCOMEXP0 As typeYCOMEXP0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCOMEXP0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCOMEXP0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCOMEXP0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YCOMEXP0s.bas  ( " & Trim(recYCOMEXP0.Obj) & " : " & Trim(recYCOMEXP0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYCOMEXP0_Monitor(recYCOMEXP0 As typeYCOMEXP0)
'-----------------------------------------------------

arrYCOMEXP0_Suite = False
Select Case mId$(Trim(recYCOMEXP0.Method), 1, 4)
    Case "Snap"
              srvYCOMEXP0_Monitor = srvYCOMEXP0_Snap(recYCOMEXP0)
    Case Else
            srvYCOMEXP0_Monitor = srvYCOMEXP0_Seek(recYCOMEXP0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYCOMEXP0_GetBuffer(recYCOMEXP0 As typeYCOMEXP0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCOMEXP0_GetBuffer = Null
recYCOMEXP0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCOMEXP0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCOMEXP0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCOMEXP0.Err = Space$(10) Then
    recYCOMEXP0.COMEXPETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCOMEXP0.COMEXPTRA = mId$(MsgTxt, K + 6, 6)
    recYCOMEXP0.COMEXPOPT = mId$(MsgTxt, K + 12, 3)
    recYCOMEXP0.COMEXPARG = mId$(MsgTxt, K + 15, 12)
    recYCOMEXP0.COMEXPDON = mId$(MsgTxt, K + 27, 100)
Else
    srvYCOMEXP0_GetBuffer = recYCOMEXP0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCOMEXP0Len

End Function

'---------------------------------------------------------
Public Sub srvYCOMEXP0_PutBuffer(recYCOMEXP0 As typeYCOMEXP0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCOMEXP0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCOMEXP0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 5) = Format$(recYCOMEXP0.COMEXPETA, "0000 ")
    Mid$(MsgTxt, K + 6, 6) = recYCOMEXP0.COMEXPTRA
    Mid$(MsgTxt, K + 12, 3) = recYCOMEXP0.COMEXPOPT
    Mid$(MsgTxt, K + 15, 12) = recYCOMEXP0.COMEXPARG
    Mid$(MsgTxt, K + 27, 100) = recYCOMEXP0.COMEXPDON
MsgTxtLen = MsgTxtLen + recYCOMEXP0Len
End Sub


Public Sub srvYCOMEXP0_ElpDisplay(recYCOMEXP0 As typeYCOMEXP0)
frmElpDisplay.fgData.Rows = 6
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMEXPETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMEXP0.COMEXPETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMEXPTRA    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TRAITEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMEXP0.COMEXPTRA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMEXPOPT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPTION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMEXP0.COMEXPOPT
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMEXPARG   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ARGUMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMEXP0.COMEXPARG
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMEXPDON  100A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONNEE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMEXP0.COMEXPDON
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYCOMEXP0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YCOMEXP0.txt" For Input As #1
Open "C:\Temp\YCOMEXP0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "COMEXPETA;COMEXPTRA;COMEXPOPT;COMEXPARG;COMEXPDON;"
    Print #2, "ETABLISSEMENT;CODE TRAITEMENT;CODE OPTION;ARGUMENT;DONNEE;"
    Print #2, ";;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 6) & ";" _
      & mId$(xIn, 12, 3) & ";" _
      & mId$(xIn, 15, 12) & ";" _
      & mId$(xIn, 27, 100) & ";"
Loop
Close
End Sub

'---------------------------------------------------------
Private Function srvYCOMEXP0_Seek(recYCOMEXP0 As typeYCOMEXP0)
'---------------------------------------------------------

srvYCOMEXP0_Seek = "?"
MsgTxtLen = 0
Call srvYCOMEXP0_PutBuffer(recYCOMEXP0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYCOMEXP0_GetBuffer(recYCOMEXP0)) Then
        srvYCOMEXP0_Seek = Null
    Else
        Call srvYCOMEXP0_Error(recYCOMEXP0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYCOMEXP0_Snap(recYCOMEXP0 As typeYCOMEXP0)
'---------------------------------------------------------
srvYCOMEXP0_Snap = "?"
MsgTxtLen = 0
Call srvYCOMEXP0_PutBuffer(recYCOMEXP0)
Call srvYCOMEXP0_PutBuffer(arrYCOMEXP0(0))
If IsNull(SndRcv()) Then
    srvYCOMEXP0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYCOMEXP0_GetBuffer(recYCOMEXP0)) Then
            Call arrYCOMEXP0_AddItem(recYCOMEXP0)
            arrYCOMEXP0_Suite = True
        Else
            arrYCOMEXP0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYCOMEXP0_AddItem(recYCOMEXP0 As typeYCOMEXP0)
'---------------------------------------------------------
          
arrYCOMEXP0_NB = arrYCOMEXP0_NB + 1
    
If arrYCOMEXP0_NB > arrYCOMEXP0_NBMax Then
    arrYCOMEXP0_NBMax = arrYCOMEXP0_NBMax + recYCOMEXP0_Block
    ReDim Preserve arrYCOMEXP0(arrYCOMEXP0_NBMax)
End If
            
arrYCOMEXP0(arrYCOMEXP0_NB) = recYCOMEXP0
End Sub



'---------------------------------------------------------
Public Sub recYCOMEXP0_Init(recYCOMEXP0 As typeYCOMEXP0)
'---------------------------------------------------------
recYCOMEXP0.Obj = "ZCOMEXP0_S"
recYCOMEXP0.Method = ""
recYCOMEXP0.Err = ""
recYCOMEXP0.COMEXPETA = 1
recYCOMEXP0.COMEXPTRA = ""
recYCOMEXP0.COMEXPOPT = ""
recYCOMEXP0.COMEXPARG = ""
recYCOMEXP0.COMEXPDON = ""
End Sub





