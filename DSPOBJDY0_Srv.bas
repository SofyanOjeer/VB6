Attribute VB_Name = "srvDSPOBJDY0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDSPOBJDY0Len = 648 ' 34 +614
Public Const recDSPOBJDY0_Block = 30
Public Const memoDSPOBJDY0Len = 614
Public Const constDSPOBJDY0 = "DSPOBJDY0  "

Type typeDSPOBJDY0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    ODDCEN          As String * 1                     ' Display century: 0=19xx, 1=20xx
    ODDDAT          As String * 6                     ' Display date (Job date format)
    ODDTIM          As String * 6                     ' Display time (HHMMSS)
    ODLBNM          As String * 10                    ' Library
    ODOBNM          As String * 10                    ' Object
    ODOBTP          As String * 8                     ' Object type
    ODOBAT          As String * 10                    ' Object attribute
    ODOBFR          As String * 1                     ' Storage freed: 0=Not freed, 1=F
    ODOBSZ          As Long                           ' Object size: 9,999,999,999=Use
    ODOBTX          As String * 50                    ' Text description
    ODOBLK          As String * 1                     ' Object locked: 0=Not locked, 1=
    ODOBDM          As String * 1                     ' Object damaged: 0=Not damaged,
    ODCCEN          As String * 1                     ' Creation century: 0=19xx, 1=20x
    ODCDAT          As String * 6                     ' Creation date (MMDDYY)
    ODCTIM          As String * 6                     ' Creation time (HHMMSS)
    ODOBOW          As String * 10                    ' Object owner
    ODSCEN          As String * 1                     ' Save century: 0=19xx, 1=20xx
    ODSDAT          As String * 6                     ' Save date (MMDDYY)
    ODSTIM          As String * 6                     ' Save time (HHMMSS)
    ODSCMD          As String * 10                    ' Save command
    ODSSZE          As Long                           ' Saved size
    ODSSLT          As Long                           ' Starting slot
    ODSDEV          As String * 10                    ' Save device
    ODSV01          As String * 6                     ' Saved volume
    ODSV02          As String * 6                     ' Saved volume
    ODSV03          As String * 6                     ' Saved volume
    ODSV04          As String * 6                     ' Saved volume
    ODSV05          As String * 6                     ' Saved volume
    ODSV06          As String * 6                     ' Saved volume
    ODSV07          As String * 6                     ' Saved volume
    ODSV08          As String * 6                     ' Saved volume
    ODSV09          As String * 6                     ' Saved volume
    ODSV10          As String * 6                     ' Saved volume
    ODSVMR          As String * 1                     ' More volumes: 0=No, 1=Yes, 2=Pa
    ODRCEN          As String * 1                     ' Restore century: 0=19xx, 1=20xx
    ODRDAT          As String * 6                     ' Restore date (MMDDYY)
    ODRTIM          As String * 6                     ' Restore time (HHMMSS)
    ODCPFL          As String * 6                     ' System level
    ODSRCF          As String * 10                    ' Source file name
    ODSRCL          As String * 10                    ' Source file library
    ODSRCM          As String * 10                    ' Source file member
    ODSRCC          As String * 1                     ' Source change century: 0=19xx,
    ODSRCD          As String * 6                     ' Source change date (YYMMDD)
    ODSRCT          As String * 6                     ' Source change time (HHMMSS)
    ODCMNM          As String * 7                     ' Compiler name
    ODCMVR          As String * 6                     ' Compiler level
    ODOBLV          As String * 8                     ' Object level
    ODUMOD          As String * 1                     ' User modified: 0=Not modified,
    ODPPNM          As String * 7                     ' LICPGM name
    ODPPVR          As String * 6                     ' LICPGM level
    ODPCNR          As String * 5                     ' PTF number
    ODAPAR          As String * 6                     ' APAR ID
    ODSSQN          As Long                           ' Sequence number: -5=See ODSSQL
    ODLCEN          As String * 1                     ' Change century: 0=19xx, 1=20xx
    ODLDAT          As String * 6                     ' Change date (MMDDYY)
    ODLTIM          As String * 6                     ' Change time (HHMMSS)
    ODSFIL          As String * 10                    ' Save file
    ODSFLB          As String * 10                    ' Save file library
    ODASP           As Long                           ' ASP number
    ODLBL           As String * 17                    ' File label
    ODPTFN          As String * 7                     ' PTF ID
    ODOBSY          As String * 8                     ' System name
    ODCRTU          As String * 10                    ' Created by user
    ODCRTS          As String * 8                     ' System created on
    ODUUPD          As String * 1                     ' Usage updated: Y=Yes, N=No
    ODUCEN          As String * 1                     ' Last used century: 0=19xx, 1=20
    ODUDAT          As String * 6                     ' Last used date (MMDDYY)
    ODUCNT          As Long                           ' Days used count
    ODTCEN          As String * 1                     ' Reset century: 0=19xx, 1=20xx
    ODTDAT          As String * 6                     ' Reset date (MMDDYY)
    ODODMN          As String * 2                     ' Object domain: *S=System, *U=Us
    ODCPVR          As String * 6                     ' System version
    ODCVRM          As String * 6                     ' Compiler version
    ODPVRM          As String * 6                     ' LICPGM version
    ODCPRS          As String * 1                     ' Compression status
    ODOASP          As String * 1                     ' Overflowed ASP: 0=No, 1=Yes
    ODAAPI          As String * 1                     ' Allow change by API: 0=No, 1=Ye
    ODAPIC          As String * 1                     ' Changed by API: 0=Not changed,1
    ODUATR          As String * 10                    ' User-defined attribute
    ODACEN          As String * 1                     ' Save active century: 0=19xx, 1=
    ODADAT          As String * 6                     ' Save active date (MMDDYY)
    ODATIM          As String * 6                     ' Save active time (HHMMSS)
    ODAUDT          As String * 10                    ' Object auditing value
    ODSIZU          As Long                           ' Object size in units
    ODBPUN          As Long                           ' Bytes per unit
    ODPGP           As String * 10                    ' Primary group
    ODSSQL          As Long                           ' Large sequence number
    ODOSIG          As String * 1                     ' Digitally signed: 0=No, 1=Yes
    ODJRST          As String * 1                     ' Currently journaled: 0=No, 1=Ye
    ODJRNM          As String * 10                    ' Journal name
    ODJRLB          As String * 10                    ' Journal library
    ODJRIM          As String * 1                     ' Journal images: 0=*AFTER, 1=*BO
    ODJREN          As String * 1                     ' Journal entries omitted: 0=*NON
    ODJRCN          As String * 1                     ' Journal century: 0=19xx, 1=20xx
    ODJRDT          As String * 6                     ' Journal date (MMDDYY)
    ODJRTI          As String * 6                     ' Journal time (HHMMSS)
    ODSSZU          As Long                           ' Save size in units
End Type
    
    
Public arrDSPOBJDY0() As typeDSPOBJDY0
Public arrDSPOBJDY0_NB As Integer
Public arrDSPOBJDY0_NBMax As Integer
Public arrDSPOBJDY0_Index As Integer
Public arrDSPOBJDY0_Suite As Boolean

Public Function srvDSPOBJDY0_Import(lnb As Long)
Dim xIn As String, X As String
Dim meMVTP0 As typeMvtP0

On Error GoTo Error_Handle


srvDSPOBJDY0_Import = "?"

Open Trim("C:\Temp\DSPOBJDY0") For Input As #1

lnb = 0

recMvtP0_Init meMVTP0
meMVTP0.Method = constAddNew

mdbMvtP0.tableMvtP0_Open

Do Until EOF(1)
    lnb = lnb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meMVTP0.ID = constDSPOBJDY0 & mId$(xIn, 6, 7)
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
            
    End If
        
Loop


Close
srvDSPOBJDY0_Import = Null
Exit Function

Error_Handle:
 MsgBox "erreur : srvDSPOBJDY0_Import" & xIn, vbCritical, Error
Close

srvDSPOBJDY0_Import = Error
End Function



'-----------------------------------------------------
Function srvDSPOBJDY0_Update(recDSPOBJDY0 As typeDSPOBJDY0)
'-----------------------------------------------------

srvDSPOBJDY0_Update = "?"

MsgTxtLen = 0
Call srvDSPOBJDY0_PutBuffer(recDSPOBJDY0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvDSPOBJDY0_GetBuffer(recDSPOBJDY0)) Then
        Call srvDSPOBJDY0_Error(recDSPOBJDY0)
        srvDSPOBJDY0_Update = recDSPOBJDY0.Err
        Exit Function
    Else
        srvDSPOBJDY0_Update = Null
    End If
Else
    recDSPOBJDY0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvDSPOBJDY0_Error(recDSPOBJDY0 As typeDSPOBJDY0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DSPOBJDY0" & Chr$(10) & Chr$(13)

Select Case mId$(recDSPOBJDY0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDSPOBJDY0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : DSPOBJDY0s.bas  ( " & Trim(recDSPOBJDY0.Obj) & " : " & Trim(recDSPOBJDY0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvDSPOBJDY0_Monitor(recDSPOBJDY0 As typeDSPOBJDY0)
'-----------------------------------------------------

arrDSPOBJDY0_Suite = False
Select Case mId$(Trim(recDSPOBJDY0.Method), 1, 4)
    Case "Snap"
              srvDSPOBJDY0_Monitor = srvDSPOBJDY0_Snap(recDSPOBJDY0)
    Case Else
            srvDSPOBJDY0_Monitor = srvDSPOBJDY0_Seek(recDSPOBJDY0)
End Select

End Function

'---------------------------------------------------------
Public Function srvDSPOBJDY0_GetBuffer(recDSPOBJDY0 As typeDSPOBJDY0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDSPOBJDY0_GetBuffer = Null
recDSPOBJDY0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recDSPOBJDY0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recDSPOBJDY0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDSPOBJDY0.Err = Space$(10) Then
    recDSPOBJDY0.ODDCEN = mId$(MsgTxt, K + 1, 1)
    recDSPOBJDY0.ODDDAT = mId$(MsgTxt, K + 2, 6)
    recDSPOBJDY0.ODDTIM = mId$(MsgTxt, K + 8, 6)
    recDSPOBJDY0.ODLBNM = mId$(MsgTxt, K + 14, 10)
    recDSPOBJDY0.ODOBNM = mId$(MsgTxt, K + 24, 10)
    recDSPOBJDY0.ODOBTP = mId$(MsgTxt, K + 34, 8)
    recDSPOBJDY0.ODOBAT = mId$(MsgTxt, K + 42, 10)
    recDSPOBJDY0.ODOBFR = mId$(MsgTxt, K + 52, 1)
    recDSPOBJDY0.ODOBSZ = CLng(Val(mId$(MsgTxt, K + 53, 11)))
    recDSPOBJDY0.ODOBTX = mId$(MsgTxt, K + 64, 50)
    recDSPOBJDY0.ODOBLK = mId$(MsgTxt, K + 114, 1)
    recDSPOBJDY0.ODOBDM = mId$(MsgTxt, K + 115, 1)
    recDSPOBJDY0.ODCCEN = mId$(MsgTxt, K + 116, 1)
    recDSPOBJDY0.ODCDAT = mId$(MsgTxt, K + 117, 6)
    recDSPOBJDY0.ODCTIM = mId$(MsgTxt, K + 123, 6)
    recDSPOBJDY0.ODOBOW = mId$(MsgTxt, K + 129, 10)
    recDSPOBJDY0.ODSCEN = mId$(MsgTxt, K + 139, 1)
    recDSPOBJDY0.ODSDAT = mId$(MsgTxt, K + 140, 6)
    recDSPOBJDY0.ODSTIM = mId$(MsgTxt, K + 146, 6)
    recDSPOBJDY0.ODSCMD = mId$(MsgTxt, K + 152, 10)
    recDSPOBJDY0.ODSSZE = CLng(Val(mId$(MsgTxt, K + 162, 11)))
    recDSPOBJDY0.ODSSLT = CLng(Val(mId$(MsgTxt, K + 173, 3)))
    recDSPOBJDY0.ODSDEV = mId$(MsgTxt, K + 176, 10)
    recDSPOBJDY0.ODSV01 = mId$(MsgTxt, K + 186, 6)
    recDSPOBJDY0.ODSV02 = mId$(MsgTxt, K + 192, 6)
    recDSPOBJDY0.ODSV03 = mId$(MsgTxt, K + 198, 6)
    recDSPOBJDY0.ODSV04 = mId$(MsgTxt, K + 204, 6)
    recDSPOBJDY0.ODSV05 = mId$(MsgTxt, K + 210, 6)
    recDSPOBJDY0.ODSV06 = mId$(MsgTxt, K + 216, 6)
    recDSPOBJDY0.ODSV07 = mId$(MsgTxt, K + 222, 6)
    recDSPOBJDY0.ODSV08 = mId$(MsgTxt, K + 228, 6)
    recDSPOBJDY0.ODSV09 = mId$(MsgTxt, K + 234, 6)
    recDSPOBJDY0.ODSV10 = mId$(MsgTxt, K + 240, 6)
    recDSPOBJDY0.ODSVMR = mId$(MsgTxt, K + 246, 1)
    recDSPOBJDY0.ODRCEN = mId$(MsgTxt, K + 247, 1)
    recDSPOBJDY0.ODRDAT = mId$(MsgTxt, K + 248, 6)
    recDSPOBJDY0.ODRTIM = mId$(MsgTxt, K + 254, 6)
    recDSPOBJDY0.ODCPFL = mId$(MsgTxt, K + 260, 6)
    recDSPOBJDY0.ODSRCF = mId$(MsgTxt, K + 266, 10)
    recDSPOBJDY0.ODSRCL = mId$(MsgTxt, K + 276, 10)
    recDSPOBJDY0.ODSRCM = mId$(MsgTxt, K + 286, 10)
    recDSPOBJDY0.ODSRCC = mId$(MsgTxt, K + 296, 1)
    recDSPOBJDY0.ODSRCD = mId$(MsgTxt, K + 297, 6)
    recDSPOBJDY0.ODSRCT = mId$(MsgTxt, K + 303, 6)
    recDSPOBJDY0.ODCMNM = mId$(MsgTxt, K + 309, 7)
    recDSPOBJDY0.ODCMVR = mId$(MsgTxt, K + 316, 6)
    recDSPOBJDY0.ODOBLV = mId$(MsgTxt, K + 322, 8)
    recDSPOBJDY0.ODUMOD = mId$(MsgTxt, K + 330, 1)
    recDSPOBJDY0.ODPPNM = mId$(MsgTxt, K + 331, 7)
    recDSPOBJDY0.ODPPVR = mId$(MsgTxt, K + 338, 6)
    recDSPOBJDY0.ODPCNR = mId$(MsgTxt, K + 344, 5)
    recDSPOBJDY0.ODAPAR = mId$(MsgTxt, K + 349, 6)
    recDSPOBJDY0.ODSSQN = CLng(Val(mId$(MsgTxt, K + 355, 5)))
    recDSPOBJDY0.ODLCEN = mId$(MsgTxt, K + 360, 1)
    recDSPOBJDY0.ODLDAT = mId$(MsgTxt, K + 361, 6)
    recDSPOBJDY0.ODLTIM = mId$(MsgTxt, K + 367, 6)
    recDSPOBJDY0.ODSFIL = mId$(MsgTxt, K + 373, 10)
    recDSPOBJDY0.ODSFLB = mId$(MsgTxt, K + 383, 10)
    recDSPOBJDY0.ODASP = CLng(Val(mId$(MsgTxt, K + 393, 3)))
    recDSPOBJDY0.ODLBL = mId$(MsgTxt, K + 396, 17)
    recDSPOBJDY0.ODPTFN = mId$(MsgTxt, K + 413, 7)
    recDSPOBJDY0.ODOBSY = mId$(MsgTxt, K + 420, 8)
    recDSPOBJDY0.ODCRTU = mId$(MsgTxt, K + 428, 10)
    recDSPOBJDY0.ODCRTS = mId$(MsgTxt, K + 438, 8)
    recDSPOBJDY0.ODUUPD = mId$(MsgTxt, K + 446, 1)
    recDSPOBJDY0.ODUCEN = mId$(MsgTxt, K + 447, 1)
    recDSPOBJDY0.ODUDAT = mId$(MsgTxt, K + 448, 6)
    recDSPOBJDY0.ODUCNT = CLng(Val(mId$(MsgTxt, K + 454, 6)))
    recDSPOBJDY0.ODTCEN = mId$(MsgTxt, K + 460, 1)
    recDSPOBJDY0.ODTDAT = mId$(MsgTxt, K + 461, 6)
    recDSPOBJDY0.ODODMN = mId$(MsgTxt, K + 467, 2)
    recDSPOBJDY0.ODCPVR = mId$(MsgTxt, K + 469, 6)
    recDSPOBJDY0.ODCVRM = mId$(MsgTxt, K + 475, 6)
    recDSPOBJDY0.ODPVRM = mId$(MsgTxt, K + 481, 6)
    recDSPOBJDY0.ODCPRS = mId$(MsgTxt, K + 487, 1)
    recDSPOBJDY0.ODOASP = mId$(MsgTxt, K + 488, 1)
    recDSPOBJDY0.ODAAPI = mId$(MsgTxt, K + 489, 1)
    recDSPOBJDY0.ODAPIC = mId$(MsgTxt, K + 490, 1)
    recDSPOBJDY0.ODUATR = mId$(MsgTxt, K + 491, 10)
    recDSPOBJDY0.ODACEN = mId$(MsgTxt, K + 501, 1)
    recDSPOBJDY0.ODADAT = mId$(MsgTxt, K + 502, 6)
    recDSPOBJDY0.ODATIM = mId$(MsgTxt, K + 508, 6)
    recDSPOBJDY0.ODAUDT = mId$(MsgTxt, K + 514, 10)
    recDSPOBJDY0.ODSIZU = CLng(Val(mId$(MsgTxt, K + 524, 11)))
   'recDSPOBJDY0.ODBPUN = CLng(Val(mId$(MsgTxt, K + 535, 11)))
    recDSPOBJDY0.ODPGP = mId$(MsgTxt, K + 546, 10)
    recDSPOBJDY0.ODSSQL = CLng(Val(mId$(MsgTxt, K + 556, 11)))
    recDSPOBJDY0.ODOSIG = mId$(MsgTxt, K + 567, 1)
    recDSPOBJDY0.ODJRST = mId$(MsgTxt, K + 568, 1)
    recDSPOBJDY0.ODJRNM = mId$(MsgTxt, K + 569, 10)
    recDSPOBJDY0.ODJRLB = mId$(MsgTxt, K + 579, 10)
    recDSPOBJDY0.ODJRIM = mId$(MsgTxt, K + 589, 1)
    recDSPOBJDY0.ODJREN = mId$(MsgTxt, K + 590, 1)
    recDSPOBJDY0.ODJRCN = mId$(MsgTxt, K + 591, 1)
    recDSPOBJDY0.ODJRDT = mId$(MsgTxt, K + 592, 6)
    recDSPOBJDY0.ODJRTI = mId$(MsgTxt, K + 598, 6)
    recDSPOBJDY0.ODSSZU = CLng(Val(mId$(MsgTxt, K + 604, 11)))
Else
    srvDSPOBJDY0_GetBuffer = recDSPOBJDY0.Err
End If
MsgTxtIndex = MsgTxtIndex + recDSPOBJDY0Len

End Function

Public Sub srvDSPOBJDY0_ElpDisplay(recDSPOBJDY0 As typeDSPOBJDY0)
frmElpDisplay.fgData.Rows = 98
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODDCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Display century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODDCEN
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODDDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Display date (Job date format)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODDDAT
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODDTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Display time (HHMMSS)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODDTIM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODLBNM   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODLBNM
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBNM   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBNM
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBTP    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object type"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBTP
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBAT   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object attribute"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBAT
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBFR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Storage freed: 0=Not freed, 1=Freed"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBFR
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBSZ   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object size: 9,999,999,999=Use ODSIZU*ODBPUN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBSZ
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBTX   50A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Text description"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBTX
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBLK    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object locked: 0=Not locked, 1=Locked"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBLK
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBDM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object damaged: 0=Not damaged, 1=Full, 2=Partial"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBDM
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Creation century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCCEN
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Creation date (MMDDYY)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCDAT
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Creation time (HHMMSS)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCTIM
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBOW   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object owner"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBOW
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSCEN
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save date (MMDDYY)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSDAT
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save time (HHMMSS)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSTIM
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSCMD   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save command"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSCMD
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSSZE   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved size"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSSZE
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSSLT    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Starting slot"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSSLT
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSDEV   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save device"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSDEV
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV01    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV01
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV02    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV02
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV03    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV03
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV04    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV04
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV05    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV05
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV06    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV06
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV07    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV07
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV08    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV08
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV09    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV09
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSV10    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Saved volume"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSV10
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSVMR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "More volumes: 0=No, 1=Yes, 2=Parallel save format"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSVMR
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODRCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Restore century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODRCEN
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODRDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Restore date (MMDDYY)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODRDAT
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODRTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Restore time (HHMMSS)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODRTIM
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCPFL    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "System level"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCPFL
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSRCF   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Source file name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSRCF
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSRCL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Source file library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSRCL
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSRCM   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Source file member"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSRCM
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSRCC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Source change century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSRCC
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSRCD    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Source change date (YYMMDD)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSRCD
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSRCT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Source change time (HHMMSS)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSRCT
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCMNM    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Compiler name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCMNM
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCMVR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Compiler level"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCMVR
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBLV    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object level"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBLV
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODUMOD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "User modified: 0=Not modified, 1=Modified"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODUMOD
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODPPNM    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LICPGM name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODPPNM
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODPPVR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LICPGM level"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODPPVR
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODPCNR    5A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PTF number"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODPCNR
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODAPAR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "APAR ID"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODAPAR
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSSQN    4S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Sequence number: -5=See ODSSQL field"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSSQN
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODLCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Change century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODLCEN
frmElpDisplay.fgData.Row = 55
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODLDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Change date (MMDDYY)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODLDAT
frmElpDisplay.fgData.Row = 56
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODLTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Change time (HHMMSS)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODLTIM
frmElpDisplay.fgData.Row = 57
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSFIL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save file"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSFIL
frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSFLB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save file library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSFLB
frmElpDisplay.fgData.Row = 59
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODASP    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ASP number"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODASP
frmElpDisplay.fgData.Row = 60
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODLBL   17A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File label"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODLBL
frmElpDisplay.fgData.Row = 61
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODPTFN    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PTF ID"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODPTFN
frmElpDisplay.fgData.Row = 62
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOBSY    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "System name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOBSY
frmElpDisplay.fgData.Row = 63
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCRTU   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Created by user"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCRTU
frmElpDisplay.fgData.Row = 64
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCRTS    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "System created on"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCRTS
frmElpDisplay.fgData.Row = 65
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODUUPD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Usage updated: Y=Yes, N=No"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODUUPD
frmElpDisplay.fgData.Row = 66
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODUCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last used century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODUCEN
frmElpDisplay.fgData.Row = 67
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODUDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last used date (MMDDYY)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODUDAT
frmElpDisplay.fgData.Row = 68
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODUCNT    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Days used count"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODUCNT
frmElpDisplay.fgData.Row = 69
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODTCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reset century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODTCEN
frmElpDisplay.fgData.Row = 70
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODTDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reset date (MMDDYY)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODTDAT
frmElpDisplay.fgData.Row = 71
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODODMN    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object domain: *S=System, *U=User"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODODMN
frmElpDisplay.fgData.Row = 72
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCPVR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "System version"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCPVR
frmElpDisplay.fgData.Row = 73
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCVRM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Compiler version"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCVRM
frmElpDisplay.fgData.Row = 74
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODPVRM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LICPGM version"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODPVRM
frmElpDisplay.fgData.Row = 75
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODCPRS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Compression status"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODCPRS
frmElpDisplay.fgData.Row = 76
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOASP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Overflowed ASP: 0=No, 1=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOASP
frmElpDisplay.fgData.Row = 77
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODAAPI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Allow change by API: 0=No, 1=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODAAPI
frmElpDisplay.fgData.Row = 78
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODAPIC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Changed by API: 0=Not changed,1=Changed"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODAPIC
frmElpDisplay.fgData.Row = 79
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODUATR   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "User-defined attribute"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODUATR
frmElpDisplay.fgData.Row = 80
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODACEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save active century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODACEN
frmElpDisplay.fgData.Row = 81
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODADAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save active date (MMDDYY)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODADAT
frmElpDisplay.fgData.Row = 82
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODATIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save active time (HHMMSS)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODATIM
frmElpDisplay.fgData.Row = 83
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODAUDT   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object auditing value"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODAUDT
frmElpDisplay.fgData.Row = 84
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSIZU   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Object size in units"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSIZU
frmElpDisplay.fgData.Row = 85
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODBPUN   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Bytes per unit"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODBPUN
frmElpDisplay.fgData.Row = 86
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODPGP   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Primary group"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODPGP
frmElpDisplay.fgData.Row = 87
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSSQL   10S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Large sequence number"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSSQL
frmElpDisplay.fgData.Row = 88
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODOSIG    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Digitally signed: 0=No, 1=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODOSIG
frmElpDisplay.fgData.Row = 89
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODJRST    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Currently journaled: 0=No, 1=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODJRST
frmElpDisplay.fgData.Row = 90
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODJRNM   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODJRNM
frmElpDisplay.fgData.Row = 91
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODJRLB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODJRLB
frmElpDisplay.fgData.Row = 92
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODJRIM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal images: 0=*AFTER, 1=*BOTH"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODJRIM
frmElpDisplay.fgData.Row = 93
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODJREN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal entries omitted: 0=*NONE, 1=*OPNCLO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODJREN
frmElpDisplay.fgData.Row = 94
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODJRCN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODJRCN
frmElpDisplay.fgData.Row = 95
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODJRDT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal date (MMDDYY)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODJRDT
frmElpDisplay.fgData.Row = 96
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODJRTI    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal time (HHMMSS)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODJRTI
frmElpDisplay.fgData.Row = 97
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ODSSZU   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Save size in units"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPOBJDY0.ODSSZU
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Public Sub srvDSPOBJDY0_PutBuffer(recDSPOBJDY0 As typeDSPOBJDY0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDSPOBJDY0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDSPOBJDY0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 1) = recDSPOBJDY0.ODDCEN
    Mid$(MsgTxt, K + 2, 6) = recDSPOBJDY0.ODDDAT
    Mid$(MsgTxt, K + 8, 6) = recDSPOBJDY0.ODDTIM
    Mid$(MsgTxt, K + 14, 10) = recDSPOBJDY0.ODLBNM
    Mid$(MsgTxt, K + 24, 10) = recDSPOBJDY0.ODOBNM
    Mid$(MsgTxt, K + 34, 8) = recDSPOBJDY0.ODOBTP
    Mid$(MsgTxt, K + 42, 10) = recDSPOBJDY0.ODOBAT
    Mid$(MsgTxt, K + 52, 1) = recDSPOBJDY0.ODOBFR
    Mid$(MsgTxt, K + 53, 11) = Format$(recDSPOBJDY0.ODOBSZ, "0000000000 ")
    Mid$(MsgTxt, K + 64, 50) = recDSPOBJDY0.ODOBTX
    Mid$(MsgTxt, K + 114, 1) = recDSPOBJDY0.ODOBLK
    Mid$(MsgTxt, K + 115, 1) = recDSPOBJDY0.ODOBDM
    Mid$(MsgTxt, K + 116, 1) = recDSPOBJDY0.ODCCEN
    Mid$(MsgTxt, K + 117, 6) = recDSPOBJDY0.ODCDAT
    Mid$(MsgTxt, K + 123, 6) = recDSPOBJDY0.ODCTIM
    Mid$(MsgTxt, K + 129, 10) = recDSPOBJDY0.ODOBOW
    Mid$(MsgTxt, K + 139, 1) = recDSPOBJDY0.ODSCEN
    Mid$(MsgTxt, K + 140, 6) = recDSPOBJDY0.ODSDAT
    Mid$(MsgTxt, K + 146, 6) = recDSPOBJDY0.ODSTIM
    Mid$(MsgTxt, K + 152, 10) = recDSPOBJDY0.ODSCMD
    Mid$(MsgTxt, K + 162, 11) = Format$(recDSPOBJDY0.ODSSZE, "0000000000 ")
    Mid$(MsgTxt, K + 173, 3) = Format$(recDSPOBJDY0.ODSSLT, "00 ")
    Mid$(MsgTxt, K + 176, 10) = recDSPOBJDY0.ODSDEV
    Mid$(MsgTxt, K + 186, 6) = recDSPOBJDY0.ODSV01
    Mid$(MsgTxt, K + 192, 6) = recDSPOBJDY0.ODSV02
    Mid$(MsgTxt, K + 198, 6) = recDSPOBJDY0.ODSV03
    Mid$(MsgTxt, K + 204, 6) = recDSPOBJDY0.ODSV04
    Mid$(MsgTxt, K + 210, 6) = recDSPOBJDY0.ODSV05
    Mid$(MsgTxt, K + 216, 6) = recDSPOBJDY0.ODSV06
    Mid$(MsgTxt, K + 222, 6) = recDSPOBJDY0.ODSV07
    Mid$(MsgTxt, K + 228, 6) = recDSPOBJDY0.ODSV08
    Mid$(MsgTxt, K + 234, 6) = recDSPOBJDY0.ODSV09
    Mid$(MsgTxt, K + 240, 6) = recDSPOBJDY0.ODSV10
    Mid$(MsgTxt, K + 246, 1) = recDSPOBJDY0.ODSVMR
    Mid$(MsgTxt, K + 247, 1) = recDSPOBJDY0.ODRCEN
    Mid$(MsgTxt, K + 248, 6) = recDSPOBJDY0.ODRDAT
    Mid$(MsgTxt, K + 254, 6) = recDSPOBJDY0.ODRTIM
    Mid$(MsgTxt, K + 260, 6) = recDSPOBJDY0.ODCPFL
    Mid$(MsgTxt, K + 266, 10) = recDSPOBJDY0.ODSRCF
    Mid$(MsgTxt, K + 276, 10) = recDSPOBJDY0.ODSRCL
    Mid$(MsgTxt, K + 286, 10) = recDSPOBJDY0.ODSRCM
    Mid$(MsgTxt, K + 296, 1) = recDSPOBJDY0.ODSRCC
    Mid$(MsgTxt, K + 297, 6) = recDSPOBJDY0.ODSRCD
    Mid$(MsgTxt, K + 303, 6) = recDSPOBJDY0.ODSRCT
    Mid$(MsgTxt, K + 309, 7) = recDSPOBJDY0.ODCMNM
    Mid$(MsgTxt, K + 316, 6) = recDSPOBJDY0.ODCMVR
    Mid$(MsgTxt, K + 322, 8) = recDSPOBJDY0.ODOBLV
    Mid$(MsgTxt, K + 330, 1) = recDSPOBJDY0.ODUMOD
    Mid$(MsgTxt, K + 331, 7) = recDSPOBJDY0.ODPPNM
    Mid$(MsgTxt, K + 338, 6) = recDSPOBJDY0.ODPPVR
    Mid$(MsgTxt, K + 344, 5) = recDSPOBJDY0.ODPCNR
    Mid$(MsgTxt, K + 349, 6) = recDSPOBJDY0.ODAPAR
    Mid$(MsgTxt, K + 355, 5) = Format$(recDSPOBJDY0.ODSSQN, "0000 ")
    Mid$(MsgTxt, K + 360, 1) = recDSPOBJDY0.ODLCEN
    Mid$(MsgTxt, K + 361, 6) = recDSPOBJDY0.ODLDAT
    Mid$(MsgTxt, K + 367, 6) = recDSPOBJDY0.ODLTIM
    Mid$(MsgTxt, K + 373, 10) = recDSPOBJDY0.ODSFIL
    Mid$(MsgTxt, K + 383, 10) = recDSPOBJDY0.ODSFLB
    Mid$(MsgTxt, K + 393, 3) = Format$(recDSPOBJDY0.ODASP, "00 ")
    Mid$(MsgTxt, K + 396, 17) = recDSPOBJDY0.ODLBL
    Mid$(MsgTxt, K + 413, 7) = recDSPOBJDY0.ODPTFN
    Mid$(MsgTxt, K + 420, 8) = recDSPOBJDY0.ODOBSY
    Mid$(MsgTxt, K + 428, 10) = recDSPOBJDY0.ODCRTU
    Mid$(MsgTxt, K + 438, 8) = recDSPOBJDY0.ODCRTS
    Mid$(MsgTxt, K + 446, 1) = recDSPOBJDY0.ODUUPD
    Mid$(MsgTxt, K + 447, 1) = recDSPOBJDY0.ODUCEN
    Mid$(MsgTxt, K + 448, 6) = recDSPOBJDY0.ODUDAT
    Mid$(MsgTxt, K + 454, 6) = Format$(recDSPOBJDY0.ODUCNT, "00000 ")
    Mid$(MsgTxt, K + 460, 1) = recDSPOBJDY0.ODTCEN
    Mid$(MsgTxt, K + 461, 6) = recDSPOBJDY0.ODTDAT
    Mid$(MsgTxt, K + 467, 2) = recDSPOBJDY0.ODODMN
    Mid$(MsgTxt, K + 469, 6) = recDSPOBJDY0.ODCPVR
    Mid$(MsgTxt, K + 475, 6) = recDSPOBJDY0.ODCVRM
    Mid$(MsgTxt, K + 481, 6) = recDSPOBJDY0.ODPVRM
    Mid$(MsgTxt, K + 487, 1) = recDSPOBJDY0.ODCPRS
    Mid$(MsgTxt, K + 488, 1) = recDSPOBJDY0.ODOASP
    Mid$(MsgTxt, K + 489, 1) = recDSPOBJDY0.ODAAPI
    Mid$(MsgTxt, K + 490, 1) = recDSPOBJDY0.ODAPIC
    Mid$(MsgTxt, K + 491, 10) = recDSPOBJDY0.ODUATR
    Mid$(MsgTxt, K + 501, 1) = recDSPOBJDY0.ODACEN
    Mid$(MsgTxt, K + 502, 6) = recDSPOBJDY0.ODADAT
    Mid$(MsgTxt, K + 508, 6) = recDSPOBJDY0.ODATIM
    Mid$(MsgTxt, K + 514, 10) = recDSPOBJDY0.ODAUDT
    Mid$(MsgTxt, K + 524, 11) = Format$(recDSPOBJDY0.ODSIZU, "0000000000 ")
    Mid$(MsgTxt, K + 535, 11) = Format$(recDSPOBJDY0.ODBPUN, "0000000000 ")
    Mid$(MsgTxt, K + 546, 10) = recDSPOBJDY0.ODPGP
    Mid$(MsgTxt, K + 556, 11) = Format$(recDSPOBJDY0.ODSSQL, "0000000000 ")
    Mid$(MsgTxt, K + 567, 1) = recDSPOBJDY0.ODOSIG
    Mid$(MsgTxt, K + 568, 1) = recDSPOBJDY0.ODJRST
    Mid$(MsgTxt, K + 569, 10) = recDSPOBJDY0.ODJRNM
    Mid$(MsgTxt, K + 579, 10) = recDSPOBJDY0.ODJRLB
    Mid$(MsgTxt, K + 589, 1) = recDSPOBJDY0.ODJRIM
    Mid$(MsgTxt, K + 590, 1) = recDSPOBJDY0.ODJREN
    Mid$(MsgTxt, K + 591, 1) = recDSPOBJDY0.ODJRCN
    Mid$(MsgTxt, K + 592, 6) = recDSPOBJDY0.ODJRDT
    Mid$(MsgTxt, K + 598, 6) = recDSPOBJDY0.ODJRTI
    Mid$(MsgTxt, K + 604, 11) = Format$(recDSPOBJDY0.ODSSZU, "0000000000 ")

MsgTxtLen = MsgTxtLen + recDSPOBJDY0Len
End Sub



'---------------------------------------------------------
Private Function srvDSPOBJDY0_Seek(recDSPOBJDY0 As typeDSPOBJDY0)
'---------------------------------------------------------

srvDSPOBJDY0_Seek = "?"
MsgTxtLen = 0
Call srvDSPOBJDY0_PutBuffer(recDSPOBJDY0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDSPOBJDY0_GetBuffer(recDSPOBJDY0)) Then
        srvDSPOBJDY0_Seek = Null
    Else
        Call srvDSPOBJDY0_Error(recDSPOBJDY0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDSPOBJDY0_Snap(recDSPOBJDY0 As typeDSPOBJDY0)
'---------------------------------------------------------
srvDSPOBJDY0_Snap = "?"
MsgTxtLen = 0
Call srvDSPOBJDY0_PutBuffer(recDSPOBJDY0)
Call srvDSPOBJDY0_PutBuffer(arrDSPOBJDY0(0))
If IsNull(SndRcv()) Then
    srvDSPOBJDY0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDSPOBJDY0_GetBuffer(recDSPOBJDY0)) Then
            Call arrDSPOBJDY0_AddItem(recDSPOBJDY0)
            arrDSPOBJDY0_Suite = True
        Else
            arrDSPOBJDY0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrDSPOBJDY0_AddItem(recDSPOBJDY0 As typeDSPOBJDY0)
'---------------------------------------------------------
          
arrDSPOBJDY0_NB = arrDSPOBJDY0_NB + 1
    
If arrDSPOBJDY0_NB > arrDSPOBJDY0_NBMax Then
    arrDSPOBJDY0_NBMax = arrDSPOBJDY0_NBMax + recDSPOBJDY0_Block
    ReDim Preserve arrDSPOBJDY0(arrDSPOBJDY0_NBMax)
End If
            
arrDSPOBJDY0(arrDSPOBJDY0_NB) = recDSPOBJDY0
End Sub



'---------------------------------------------------------
Public Sub recDSPOBJDY0_Init(recDSPOBJDY0 As typeDSPOBJDY0)
'---------------------------------------------------------
recDSPOBJDY0.Obj = "DSPOBJD_S"
recDSPOBJDY0.Method = ""
recDSPOBJDY0.Err = ""

End Sub









