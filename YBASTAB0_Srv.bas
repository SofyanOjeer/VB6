Attribute VB_Name = "srvYBASTAB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYBASTAB0Len = 339 ' 34 +305
Public Const recYBASTAB0_Block = 50
Public Const memoYBASTAB0Len = 305
Public Const constYBASTAB0 = "YBASTAB0  "

Type typeYBASTAB0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    BASTABETA       As Integer                        ' ETABLISSEMENT
    BASTABNUM       As Long                           ' NUMERO TABLE
    BASTABARG       As String * 16                    ' ARGUMENT
    BASTABAMJ       As Long                           ' date      1aammjj
    BASTABVAL       As Double                         ' valeur 5v9s
    BASTABDON       As String * 256                   ' DONNEES
End Type
    
    
Public arrYBASTAB0() As typeYBASTAB0
Public arrYBASTAB0_NB As Integer
Public arrYBASTAB0_NBMax As Integer
Public arrYBASTAB0_Index As Integer
Public arrYBASTAB0_Suite As Boolean

Public Function srvYBASTAB0_Import(lnb As Long)
Dim xIn As String, X As String
Dim meMVTP0 As typeMvtP0

On Error GoTo Error_Handle


srvYBASTAB0_Import = "?"

Open Trim(paramTemp_Folder & "FTP\YBASTAB0.txt") For Input As #1
lnb = 0

recMvtP0_Init meMVTP0
meMVTP0.Method = constAddNew

mdbMvtP0.tableMvtP0_Open

Do Until EOF(1)
    lnb = lnb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meMVTP0.ID = constYBASTAB0 & mId$(xIn, 6, 7)
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
            
    End If
        
Loop


Close
srvYBASTAB0_Import = Null
Exit Function

Error_Handle:
 MsgBox "erreur : srvYBASTAB0_Import" & xIn, vbCritical, Error
Close

srvYBASTAB0_Import = Error
End Function



'-----------------------------------------------------
Function srvYBASTAB0_Update(recYBASTAB0 As typeYBASTAB0, blnMsgBox_Error As Boolean)
'-----------------------------------------------------

srvYBASTAB0_Update = "?"

MsgTxtLen = 0
Call srvYBASTAB0_PutBuffer(recYBASTAB0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYBASTAB0_GetBuffer(recYBASTAB0)) Then
        If blnMsgBox_Error Then Call srvYBASTAB0_Error(recYBASTAB0)
        srvYBASTAB0_Update = recYBASTAB0.Err
        Exit Function
    Else
        srvYBASTAB0_Update = Null
    End If
Else
    recYBASTAB0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYBASTAB0_Error(recYBASTAB0 As typeYBASTAB0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YBASTAB0" & Chr$(10) & Chr$(13)

Select Case mId$(recYBASTAB0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYBASTAB0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YBASTAB0s.bas  ( " & Trim(recYBASTAB0.Obj) & " : " & Trim(recYBASTAB0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYBASTAB0_Monitor(recYBASTAB0 As typeYBASTAB0)
'-----------------------------------------------------

arrYBASTAB0_Suite = False
Select Case mId$(Trim(recYBASTAB0.Method), 1, 4)
    Case "Snap"
              srvYBASTAB0_Monitor = srvYBASTAB0_Snap(recYBASTAB0)
    Case Else
            srvYBASTAB0_Monitor = srvYBASTAB0_Seek(recYBASTAB0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYBASTAB0_GetBuffer(recYBASTAB0 As typeYBASTAB0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYBASTAB0_GetBuffer = Null
recYBASTAB0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYBASTAB0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYBASTAB0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYBASTAB0.Err = Space$(10) Then
    recYBASTAB0.BASTABETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYBASTAB0.BASTABNUM = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYBASTAB0.BASTABARG = mId$(MsgTxt, K + 10, 16)
    recYBASTAB0.BASTABAMJ = CLng(Val(mId$(MsgTxt, K + 26, 8)))
    recYBASTAB0.BASTABVAL = CDbl(Val(mId$(MsgTxt, K + 34, 16)) / 1000000000)
    recYBASTAB0.BASTABDON = mId$(MsgTxt, K + 50, 256)
Else
    srvYBASTAB0_GetBuffer = recYBASTAB0.Err
    recYBASTAB0.BASTABARG = "?"
    recYBASTAB0.BASTABAMJ = 0
    recYBASTAB0.BASTABVAL = 0
End If
MsgTxtIndex = MsgTxtIndex + recYBASTAB0Len

End Function

Public Sub srvYBASTAB0_ElpDisplay(recYBASTAB0 As typeYBASTAB0)
frmElpDisplay.fgData.Rows = 7
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASTABETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASTAB0.BASTABETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASTABNUM    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO TABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASTAB0.BASTABNUM
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASTABARG   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ARGUMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASTAB0.BASTABARG
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASTABAMJ    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "$$$DATE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASTAB0.BASTABAMJ
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASTABVAL   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "$$$VAleur"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASTAB0.BASTABVAL
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASTABDON  256A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONNEES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASTAB0.BASTABDON
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYBASTAB0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YBASTAB0.txt" For Input As #1
Open "C:\Temp\YBASTAB0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "BASTABETA;BASTABNUM;BASTABARG;BASTABAMJ;BASTABVAL;BASTABDON;"
    Print #2, "ETABLISSEMENT;NUMERO TABLE;ARGUMENT;$$$AMJ;$$$Valeur;DONNEES;"
    Print #2, ";;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 4) & ";" _
      & mId$(xIn, 10, 16) & ";" _
      & mId$(xIn, 26, 8) & ";" _
      & mId$(xIn, 34, 16) & ";" _
      & mId$(xIn, 50, 256) & ";"
Loop
Close
End Sub
'---------------------------------------------------------
Public Sub srvYBASTAB0_PutBuffer(recYBASTAB0 As typeYBASTAB0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYBASTAB0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYBASTAB0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYBASTAB0.BASTABETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYBASTAB0.BASTABNUM, "000 ")
    Mid$(MsgTxt, K + 10, 16) = recYBASTAB0.BASTABARG
    Mid$(MsgTxt, K + 26, 8) = Format$(recYBASTAB0.BASTABAMJ, "0000000 ")
    Mid$(MsgTxt, K + 34, 16) = Format$(recYBASTAB0.BASTABVAL * 1000000000, "000000000000000 ")
    If recYBASTAB0.BASTABVAL < 0 Then Mid$(MsgTxt, K + 49, 1) = "-"
    Mid$(MsgTxt, K + 50, 256) = recYBASTAB0.BASTABDON
    
MsgTxtLen = MsgTxtLen + recYBASTAB0Len
End Sub



'---------------------------------------------------------
Private Function srvYBASTAB0_Seek(recYBASTAB0 As typeYBASTAB0)
'---------------------------------------------------------

srvYBASTAB0_Seek = "?"
MsgTxtLen = 0
Call srvYBASTAB0_PutBuffer(recYBASTAB0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYBASTAB0_GetBuffer(recYBASTAB0)) Then
        srvYBASTAB0_Seek = Null
    Else
        Call srvYBASTAB0_Error(recYBASTAB0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYBASTAB0_Snap(recYBASTAB0 As typeYBASTAB0)
'---------------------------------------------------------
srvYBASTAB0_Snap = "?"
MsgTxtLen = 0
Call srvYBASTAB0_PutBuffer(recYBASTAB0)
Call srvYBASTAB0_PutBuffer(arrYBASTAB0(0))
If IsNull(SndRcv()) Then
    srvYBASTAB0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYBASTAB0_GetBuffer(recYBASTAB0)) Then
            Call arrYBASTAB0_AddItem(recYBASTAB0)
            arrYBASTAB0_Suite = True
        Else
            arrYBASTAB0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYBASTAB0_AddItem(recYBASTAB0 As typeYBASTAB0)
'---------------------------------------------------------
          
arrYBASTAB0_NB = arrYBASTAB0_NB + 1
    
If arrYBASTAB0_NB > arrYBASTAB0_NBMax Then
    arrYBASTAB0_NBMax = arrYBASTAB0_NBMax + recYBASTAB0_Block
    ReDim Preserve arrYBASTAB0(arrYBASTAB0_NBMax)
End If
            
arrYBASTAB0(arrYBASTAB0_NB) = recYBASTAB0
End Sub



'---------------------------------------------------------
Public Sub recYBASTAB0_Init(recYBASTAB0 As typeYBASTAB0)
'---------------------------------------------------------
recYBASTAB0.Obj = "ZBASTAB0_S"
recYBASTAB0.Method = ""
recYBASTAB0.Err = ""
recYBASTAB0.BASTABETA = 1
recYBASTAB0.BASTABNUM = 0
recYBASTAB0.BASTABARG = ""
recYBASTAB0.BASTABAMJ = 0
recYBASTAB0.BASTABVAL = 0
recYBASTAB0.BASTABDON = ""

End Sub









