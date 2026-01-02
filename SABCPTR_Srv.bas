Attribute VB_Name = "srvSABCPTR"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recSABCPTRLen = 255  '34 + 221
Public Const recSABCPTR_Block = 50

Type typeSABCPTR
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SCORIG                  As String * 1
    SCCOMPTE                As String * 11
    SCDEVISE                As String * 3
    
    SCSABID                 As String * 20
    SCNATGA                 As String * 1
    SCTDC                   As String * 3
    SCPCEC                  As String * 10
    SCINTITU                As String * 32
    SCDEVISO                As String * 3
    SCOUVAMJ                As String * 8
    SCCLOAMJ                As String * 8
    SCLORO                  As String * 1
    SCSUCCES                As String * 1
    SCSECUR                 As String * 2
    SCSITUAT                As String * 1
    SCCLOMOT                As String * 6
    
    SCTITID                 As String * 15
    SCTITCPT                As String * 1
    SCTITPRN                As String * 1
    SCTITRSP                As String * 1
    
    SCRELCOD                As String * 1
    SCRELADR                As String * 2
    SCRELGES                As String * 1
    SCRELNOR                As String * 6
   
    SCALIASCOD              As String * 2
    SCALIASCPT              As String * 15
    SCFICCOD                As String * 2
    SCFICID                 As String * 15
    
    SCSOLDE                 As Currency
    SCCPTGEN                As String * 8
    
    SCCREAMJ                As String * 8
    SCCREHMS                As String * 6
    SCMODAMJ                As String * 8
    SCMODHMS                As String * 6
    SCUSRNOM                As String * 10
    SCSTATUS                As String * 3
   
End Type
    
Public arrSABCPTR() As typeSABCPTR
Public arrSABCPTR_NB As Integer
Public arrSABCPTR_NBMax As Integer
Public arrSABCPTR_Index As Integer
Public arrSABCPTR_Suite As Boolean

Public xSABCPTR As typeSABCPTR

Public Sub srvSABCPTR_ElpDisplay(recSABCPTR As typeSABCPTR)
frmElpDisplay.fgData.Rows = 37
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.Err
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCORIG"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCORIG
frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCCOMPTE"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCCOMPTE
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCDEVISE"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCDEVISE
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCSABID"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCSABID
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCNATGA"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCNATGA
frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCTDC"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCTDC
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCPCEC"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCPCEC
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCINTITU"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCINTITU
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCDEVISO"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCDEVISO
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCOUVAMJ"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCOUVAMJ
frmElpDisplay.fgData.Row = 14
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCCLOAMJ"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCCLOAMJ
    
frmElpDisplay.fgData.Row = 15
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCLORO"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCLORO
frmElpDisplay.fgData.Row = 16
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCSUCCES"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCSUCCES
frmElpDisplay.fgData.Row = 17
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCSECUR"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCSECUR
    
frmElpDisplay.fgData.Row = 18
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCSITUAT"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCSITUAT
frmElpDisplay.fgData.Row = 19
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCCLOMOT"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCCLOMOT
    
 frmElpDisplay.fgData.Row = 20
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCTITID"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCTITID
frmElpDisplay.fgData.Row = 21
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCTITCPT"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCTITCPT
frmElpDisplay.fgData.Row = 22
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCTITPRN"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCTITPRN
frmElpDisplay.fgData.Row = 23
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCTITRSP"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCTITRSP
   
frmElpDisplay.fgData.Row = 24
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCRELCOD"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCRELCOD
frmElpDisplay.fgData.Row = 25
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCRELADR"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCRELADR
frmElpDisplay.fgData.Row = 26
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCRELGES"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCRELGES
frmElpDisplay.fgData.Row = 27
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCRELNOR"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCRELNOR
   
frmElpDisplay.fgData.Row = 28
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCALIASCOD"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCALIASCOD
frmElpDisplay.fgData.Row = 29
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCALIASCPT"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCALIASCPT
frmElpDisplay.fgData.Row = 30
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCSOLDE"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCSOLDE
    
frmElpDisplay.fgData.Row = 31
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCCREAMJ"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCCREAMJ
frmElpDisplay.fgData.Row = 32
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCCREHMS"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCCREHMS
frmElpDisplay.fgData.Row = 33
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCMODAMJ"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCMODAMJ
frmElpDisplay.fgData.Row = 34
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCMODHMS "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCMODHMS
frmElpDisplay.fgData.Row = 35
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCUSRNOM"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCUSRNOM
frmElpDisplay.fgData.Row = 36
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SCSTATUS"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recSABCPTR.SCSTATUS

frmElpDisplay.Show vbModal

End Sub



'-----------------------------------------------------
Function srvSABCPTR_Update(recSABCPTR As typeSABCPTR)
'-----------------------------------------------------

srvSABCPTR_Update = "?"

MsgTxtLen = 0
Call srvSABCPTR_PutBuffer(recSABCPTR)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvSABCPTR_GetBuffer(recSABCPTR)) Then
        Call srvSABCPTR_Error(recSABCPTR)
        srvSABCPTR_Update = recSABCPTR.Err
        Exit Function
    Else
        srvSABCPTR_Update = Null
    End If
Else
    recSABCPTR.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvSABCPTR_Load(recSABCPTRMin As typeSABCPTR, recSABCPTRMax As typeSABCPTR)
Dim mMethod As String

mMethod = Trim(recSABCPTRMin.Method) & "+"
arrSABCPTR_NBMax = 0
arrSABCPTR_Suite = True: arrSABCPTR_NB = 0
arrSABCPTR_NBMax = recSABCPTR_Block: ReDim arrSABCPTR(arrSABCPTR_NBMax)

arrSABCPTR(0) = recSABCPTRMax
arrSABCPTR_Suite = True
Do Until Not arrSABCPTR_Suite
    srvSABCPTR_Monitor recSABCPTRMin
    recSABCPTRMin = arrSABCPTR(arrSABCPTR_NB)
    recSABCPTRMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Function srvSABCPTR_Dtaq_Put(lFct As String, recSABCPTR As typeSABCPTR)
'-----------------------------------------------------

srvSABCPTR_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvSABCPTR_PutBuffer(recSABCPTR)
                If MsgTxtLen + recSABCPTRLen >= recSABCPTR_Block * recSABCPTRLen Then
                    Call srvSABCPTR_Dtaq_Snd(recSABCPTR): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvSABCPTR_Dtaq_Snd(recSABCPTR)
    Case Else: srvSABCPTR_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvSABCPTR_Dtaq_Snd(recSABCPTR As typeSABCPTR)
'-----------------------------------------------------

srvSABCPTR_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvSABCPTR_GetBuffer(recSABCPTR)) Then
        Call srvSABCPTR_Error(recSABCPTR)
        srvSABCPTR_Dtaq_Snd = recSABCPTR.Err
        Exit Function
    Else
        srvSABCPTR_Dtaq_Snd = Null
    End If
Else
    recSABCPTR.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvSABCPTR_Monitor(recSABCPTR As typeSABCPTR)
'-----------------------------------------------------
blnFR_Convert = False

arrSABCPTR_Suite = False
Select Case mId$(Trim(recSABCPTR.Method), 1, 4)
    Case "Seek", "Comp"
                srvSABCPTR_Monitor = srvSABCPTR_Seek(recSABCPTR)
    Case "Snap"
              srvSABCPTR_Monitor = srvSABCPTR_Snap(recSABCPTR)
    Case Else
                recSABCPTR.Err = recSABCPTR.Method
                Call srvSABCPTR_Error(recSABCPTR)
                srvSABCPTR_Monitor = recSABCPTR.Err
End Select

End Function

'-----------------------------------------------------
Sub srvSABCPTR_Error(recSABCPTR As typeSABCPTR)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "SABCPTR" & Chr$(10) & Chr$(13)

Select Case mId$(recSABCPTR.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recSABCPTR.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recSABCPTR.SCDEVISE & " : " & recSABCPTR.SCCOMPTE & " : " & recSABCPTR.SCORIG _
        , I, "module : SABCPTRs.bas  ( " & Trim(recSABCPTR.obj) & " : " & Trim(recSABCPTR.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvSABCPTR_GetBuffer(recSABCPTR As typeSABCPTR)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvSABCPTR_GetBuffer = Null
recSABCPTR.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recSABCPTR.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recSABCPTR.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recSABCPTR.Err = Space$(10) Then
    recSABCPTR.SCORIG = mId$(MsgTxt, K + 1, 1)
    recSABCPTR.SCCOMPTE = mId$(MsgTxt, K + 2, 11)
    
    recSABCPTR.SCDEVISE = mId$(MsgTxt, K + 13, 3)
    recSABCPTR.SCSABID = mId$(MsgTxt, K + 16, 20)
    recSABCPTR.SCNATGA = mId$(MsgTxt, K + 36, 1)
    
    recSABCPTR.SCTDC = mId$(MsgTxt, K + 37, 3)
    recSABCPTR.SCPCEC = mId$(MsgTxt, K + 40, 10)
  
    recSABCPTR.SCINTITU = mId$(MsgTxt, K + 50, 32)
    recSABCPTR.SCDEVISO = mId$(MsgTxt, K + 82, 3)
    recSABCPTR.SCOUVAMJ = mId$(MsgTxt, K + 85, 8)
    recSABCPTR.SCCLOAMJ = mId$(MsgTxt, K + 93, 8)
    recSABCPTR.SCLORO = mId$(MsgTxt, K + 101, 1)
    recSABCPTR.SCSUCCES = mId$(MsgTxt, K + 102, 1)
    recSABCPTR.SCSECUR = mId$(MsgTxt, K + 103, 2)
    recSABCPTR.SCSITUAT = mId$(MsgTxt, K + 105, 1)
    recSABCPTR.SCCLOMOT = mId$(MsgTxt, K + 106, 6)
    
    recSABCPTR.SCTITID = mId$(MsgTxt, K + 112, 15)
    recSABCPTR.SCTITCPT = mId$(MsgTxt, K + 127, 1)
    recSABCPTR.SCTITPRN = mId$(MsgTxt, K + 128, 1)
    recSABCPTR.SCTITRSP = mId$(MsgTxt, K + 129, 1)
    
    recSABCPTR.SCRELCOD = mId$(MsgTxt, K + 130, 1)
    recSABCPTR.SCRELADR = mId$(MsgTxt, K + 131, 2)
    recSABCPTR.SCRELGES = mId$(MsgTxt, K + 133, 1)
    recSABCPTR.SCRELNOR = mId$(MsgTxt, K + 134, 6)
   
    recSABCPTR.SCALIASCOD = mId$(MsgTxt, K + 140, 2)
    recSABCPTR.SCALIASCPT = mId$(MsgTxt, K + 142, 15)
    
    recSABCPTR.SCSOLDE = CCur(Val(mId$(MsgTxt, K + 157, 16))) / 100
    recSABCPTR.SCCPTGEN = mId$(MsgTxt, K + 173, 8)
   
    recSABCPTR.SCCREAMJ = mId$(MsgTxt, K + 181, 8)
    recSABCPTR.SCCREHMS = mId$(MsgTxt, K + 189, 6)
    recSABCPTR.SCMODAMJ = mId$(MsgTxt, K + 195, 8)
    recSABCPTR.SCMODHMS = mId$(MsgTxt, K + 203, 6)
    recSABCPTR.SCUSRNOM = mId$(MsgTxt, K + 209, 10)
    recSABCPTR.SCSTATUS = mId$(MsgTxt, K + 219, 3)

Else
    srvSABCPTR_GetBuffer = recSABCPTR.Err
End If

MsgTxtIndex = MsgTxtIndex + recSABCPTRLen

End Function

'---------------------------------------------------------
Private Sub srvSABCPTR_PutBuffer(recSABCPTR As typeSABCPTR)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recSABCPTR.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recSABCPTR.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 1) = recSABCPTR.SCORIG
    Mid$(MsgTxt, K + 2, 11) = recSABCPTR.SCCOMPTE
    
    Mid$(MsgTxt, K + 13, 3) = recSABCPTR.SCDEVISE
    Mid$(MsgTxt, K + 16, 20) = recSABCPTR.SCSABID
    Mid$(MsgTxt, K + 36, 1) = recSABCPTR.SCNATGA
    
    Mid$(MsgTxt, K + 37, 3) = recSABCPTR.SCTDC
    Mid$(MsgTxt, K + 40, 10) = recSABCPTR.SCPCEC
    
    Mid$(MsgTxt, K + 50, 32) = recSABCPTR.SCINTITU
    Mid$(MsgTxt, K + 82, 3) = recSABCPTR.SCDEVISO
    Mid$(MsgTxt, K + 85, 8) = recSABCPTR.SCOUVAMJ
    Mid$(MsgTxt, K + 93, 8) = recSABCPTR.SCCLOAMJ
    Mid$(MsgTxt, K + 101, 1) = recSABCPTR.SCLORO
    Mid$(MsgTxt, K + 102, 1) = recSABCPTR.SCSUCCES
    Mid$(MsgTxt, K + 103, 2) = recSABCPTR.SCSECUR
    Mid$(MsgTxt, K + 105, 1) = recSABCPTR.SCSITUAT
    Mid$(MsgTxt, K + 106, 6) = recSABCPTR.SCCLOMOT
    
    Mid$(MsgTxt, K + 112, 15) = recSABCPTR.SCTITID
    Mid$(MsgTxt, K + 127, 1) = recSABCPTR.SCTITCPT
    Mid$(MsgTxt, K + 128, 1) = recSABCPTR.SCTITPRN
    Mid$(MsgTxt, K + 129, 1) = recSABCPTR.SCTITRSP
    
    Mid$(MsgTxt, K + 130, 1) = recSABCPTR.SCRELCOD
    Mid$(MsgTxt, K + 131, 2) = recSABCPTR.SCRELADR
    Mid$(MsgTxt, K + 133, 1) = recSABCPTR.SCRELGES
    Mid$(MsgTxt, K + 134, 6) = recSABCPTR.SCRELNOR
    
    Mid$(MsgTxt, K + 140, 2) = recSABCPTR.SCALIASCOD
    Mid$(MsgTxt, K + 142, 15) = recSABCPTR.SCALIASCPT
    Mid$(MsgTxt, K + 157, 16) = Format$(recSABCPTR.SCSOLDE * 100, "000000000000000-")
    Mid$(MsgTxt, K + 173, 8) = recSABCPTR.SCCPTGEN

    Mid$(MsgTxt, K + 181, 8) = recSABCPTR.SCCREAMJ
    Mid$(MsgTxt, K + 189, 6) = recSABCPTR.SCCREHMS
    Mid$(MsgTxt, K + 195, 8) = recSABCPTR.SCMODAMJ
    Mid$(MsgTxt, K + 203, 6) = recSABCPTR.SCMODHMS
    Mid$(MsgTxt, K + 209, 10) = recSABCPTR.SCUSRNOM
    Mid$(MsgTxt, K + 219, 3) = recSABCPTR.SCSTATUS


MsgTxtLen = MsgTxtLen + recSABCPTRLen


  
End Sub

Public Sub cboSCSTATUS_Init(lcboSCSTATUS As ComboBox)
lcboSCSTATUS.Clear
lcboSCSTATUS.AddItem ""
lcboSCSTATUS.AddItem "Ann"
lcboSCSTATUS.AddItem "Fus"
lcboSCSTATUS.ListIndex = 0

End Sub



'---------------------------------------------------------
Private Function srvSABCPTR_Seek(recSABCPTR As typeSABCPTR)
'---------------------------------------------------------

srvSABCPTR_Seek = "?"
MsgTxtLen = 0
Call srvSABCPTR_PutBuffer(recSABCPTR)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvSABCPTR_GetBuffer(recSABCPTR)) Then
        srvSABCPTR_Seek = Null
    Else
       '' Call srvSABCPTR_Error(recSABCPTR)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvSABCPTR_Snap(recSABCPTR As typeSABCPTR)
'---------------------------------------------------------
srvSABCPTR_Snap = "?"
MsgTxtLen = 0
Call srvSABCPTR_PutBuffer(recSABCPTR)
Call srvSABCPTR_PutBuffer(arrSABCPTR(0))
If IsNull(SndRcv()) Then
    srvSABCPTR_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvSABCPTR_GetBuffer(recSABCPTR)) Then
            Call arrSABCPTR_AddItem(recSABCPTR)
            arrSABCPTR_Suite = True
        Else
            arrSABCPTR_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recSABCPTR_Init(recSABCPTR As typeSABCPTR)
'---------------------------------------------------------
MsgTxt = Space$(recSABCPTRLen)
MsgTxtIndex = 0
Call srvSABCPTR_GetBuffer(recSABCPTR)
recSABCPTR.obj = "SRVSABCPTR"
recSABCPTR.SCSOLDE = 0
recSABCPTR.SCCREAMJ = "00000000"
recSABCPTR.SCCREHMS = "000000"
recSABCPTR.SCMODAMJ = "00000000"
recSABCPTR.SCMODHMS = "000000"
End Sub

'---------------------------------------------------------
Public Sub arrSABCPTR_AddItem(recSABCPTR As typeSABCPTR)
'---------------------------------------------------------
          
arrSABCPTR_NB = arrSABCPTR_NB + 1
    
If arrSABCPTR_NB > arrSABCPTR_NBMax Then
    arrSABCPTR_NBMax = arrSABCPTR_NBMax + recSABCPTR_Block
    ReDim Preserve arrSABCPTR(arrSABCPTR_NBMax)
End If
            
arrSABCPTR(arrSABCPTR_NB) = recSABCPTR
End Sub


