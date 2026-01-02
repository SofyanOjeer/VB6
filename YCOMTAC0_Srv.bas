Attribute VB_Name = "srvYCOMTAC0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCOMTAC0Len = 554 ' 34 +520
Public Const recYCOMTAC0_Block = 50
Public Const memoYCOMTAC0Len = 520

Type typeYCOMTAC0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    COMTACETA       As Integer                        ' ETABLISSEMENT
    COMTACTRA       As String * 6                     ' CODE TRAITEMENT
    COMTACNUM       As Long                           ' NUMERO DE TACHE
    COMTACOPT       As String * 3                     ' OPTION
    COMTACPER       As String * 1                     ' PERIODICITE
    
End Type
    
    
Public arrYCOMTAC0() As typeYCOMTAC0
Public arrYCOMTAC0_NB As Integer
Public arrYCOMTAC0_NBMax As Integer
Public arrYCOMTAC0_Index As Integer
Public arrYCOMTAC0_Suite As Boolean

'-----------------------------------------------------
Function srvYCOMTAC0_Update(recYCOMTAC0 As typeYCOMTAC0)
'-----------------------------------------------------

srvYCOMTAC0_Update = "?"

MsgTxtLen = 0
Call srvYCOMTAC0_PutBuffer(recYCOMTAC0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCOMTAC0_GetBuffer(recYCOMTAC0)) Then
        Call srvYCOMTAC0_Error(recYCOMTAC0)
        srvYCOMTAC0_Update = recYCOMTAC0.Err
        Exit Function
    Else
        srvYCOMTAC0_Update = Null
    End If
Else
    recYCOMTAC0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYCOMTAC0_Error(recYCOMTAC0 As typeYCOMTAC0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCOMTAC0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCOMTAC0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCOMTAC0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YCOMTAC0s.bas  ( " & Trim(recYCOMTAC0.Obj) & " : " & Trim(recYCOMTAC0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYCOMTAC0_Monitor(recYCOMTAC0 As typeYCOMTAC0)
'-----------------------------------------------------

arrYCOMTAC0_Suite = False
Select Case mId$(Trim(recYCOMTAC0.Method), 1, 4)
    Case "Snap"
              srvYCOMTAC0_Monitor = srvYCOMTAC0_Snap(recYCOMTAC0)
    Case Else
            srvYCOMTAC0_Monitor = srvYCOMTAC0_Seek(recYCOMTAC0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYCOMTAC0_GetBuffer(recYCOMTAC0 As typeYCOMTAC0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCOMTAC0_GetBuffer = Null
recYCOMTAC0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCOMTAC0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCOMTAC0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCOMTAC0.Err = Space$(10) Then
    recYCOMTAC0.COMTACETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCOMTAC0.COMTACTRA = mId$(MsgTxt, K + 6, 6)
    recYCOMTAC0.COMTACNUM = CLng(Val(mId$(MsgTxt, K + 12, 6)))
    recYCOMTAC0.COMTACOPT = mId$(MsgTxt, K + 18, 3)
    recYCOMTAC0.COMTACPER = mId$(MsgTxt, K + 21, 1)
Else
    srvYCOMTAC0_GetBuffer = recYCOMTAC0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCOMTAC0Len

End Function

'---------------------------------------------------------
Public Sub srvYCOMTAC0_PutBuffer(recYCOMTAC0 As typeYCOMTAC0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCOMTAC0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCOMTAC0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCOMTAC0.COMTACETA, "0000 ")
    Mid$(MsgTxt, K + 6, 6) = recYCOMTAC0.COMTACTRA
    Mid$(MsgTxt, K + 12, 6) = Format$(recYCOMTAC0.COMTACNUM, "00000 ")
    Mid$(MsgTxt, K + 18, 3) = recYCOMTAC0.COMTACOPT
    Mid$(MsgTxt, K + 21, 1) = recYCOMTAC0.COMTACPER

MsgTxtLen = MsgTxtLen + recYCOMTAC0Len
End Sub



Public Sub srvYCOMTAC0_ElpDisplay(recYCOMTAC0 As typeYCOMTAC0)
frmElpDisplay.fgData.Rows = 6
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMTACETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMTAC0.COMTACETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMTACTRA    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TRAITEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMTAC0.COMTACTRA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMTACNUM    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DE TACHE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMTAC0.COMTACNUM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMTACOPT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OPTION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMTAC0.COMTACOPT
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMTACPER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PERIODICITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMTAC0.COMTACPER
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYCOMTAC0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YCOMTAC0.txt" For Input As #1
Open "C:\Temp\YCOMTAC0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "COMTACETA;COMTACTRA;COMTACNUM;COMTACOPT;COMTACPER;"
    Print #2, "ETABLISSEMENT;CODE TRAITEMENT;NUMERO DE TACHE;OPTION;PERIODICITE;"
    Print #2, ";;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 6) & ";" _
      & mId$(xIn, 12, 6) & ";" _
      & mId$(xIn, 18, 3) & ";" _
      & mId$(xIn, 21, 1) & ";"
Loop
Close
End Sub

'---------------------------------------------------------
Private Function srvYCOMTAC0_Seek(recYCOMTAC0 As typeYCOMTAC0)
'---------------------------------------------------------

srvYCOMTAC0_Seek = "?"
MsgTxtLen = 0
Call srvYCOMTAC0_PutBuffer(recYCOMTAC0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYCOMTAC0_GetBuffer(recYCOMTAC0)) Then
        srvYCOMTAC0_Seek = Null
    Else
        Call srvYCOMTAC0_Error(recYCOMTAC0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYCOMTAC0_Snap(recYCOMTAC0 As typeYCOMTAC0)
'---------------------------------------------------------
srvYCOMTAC0_Snap = "?"
MsgTxtLen = 0
Call srvYCOMTAC0_PutBuffer(recYCOMTAC0)
Call srvYCOMTAC0_PutBuffer(arrYCOMTAC0(0))
If IsNull(SndRcv()) Then
    srvYCOMTAC0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYCOMTAC0_GetBuffer(recYCOMTAC0)) Then
            Call arrYCOMTAC0_AddItem(recYCOMTAC0)
            arrYCOMTAC0_Suite = True
        Else
            arrYCOMTAC0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYCOMTAC0_AddItem(recYCOMTAC0 As typeYCOMTAC0)
'---------------------------------------------------------
          
arrYCOMTAC0_NB = arrYCOMTAC0_NB + 1
    
If arrYCOMTAC0_NB > arrYCOMTAC0_NBMax Then
    arrYCOMTAC0_NBMax = arrYCOMTAC0_NBMax + recYCOMTAC0_Block
    ReDim Preserve arrYCOMTAC0(arrYCOMTAC0_NBMax)
End If
            
arrYCOMTAC0(arrYCOMTAC0_NB) = recYCOMTAC0
End Sub



'---------------------------------------------------------
Public Sub recYCOMTAC0_Init(recYCOMTAC0 As typeYCOMTAC0)
'---------------------------------------------------------
recYCOMTAC0.Obj = "ZCOMTAC0_S"
recYCOMTAC0.Method = ""
recYCOMTAC0.Err = ""
recYCOMTAC0.COMTACETA = 1
recYCOMTAC0.COMTACNUM = 0
recYCOMTAC0.COMTACOPT = ""
recYCOMTAC0.COMTACTRA = ""
recYCOMTAC0.COMTACPER = ""

End Sub





