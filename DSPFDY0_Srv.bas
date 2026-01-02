Attribute VB_Name = "srvDSPFDY0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDSPFDY0Len = 246 ' 34 + 212
Public Const recDSPFDY0_Block = 30

Type typeDSPFDY0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    ATRCEN          As String * 1                     ' Retrieval century:  0=19xx, 1=2
    ATRDAT          As String * 6                     ' Retrieval date:  year/month/day
    ATRTIM          As String * 6                     ' Retrieval time:  hour/minute/se
    ATFILE          As String * 10                    ' File
    ATLIB           As String * 10                    ' Library
    ATFTYP          As String * 1                     ' D=Device, P=PF, L=LF, R=DDM PF,
    ATFILA          As String * 4                     ' File attribute
    ATMXDD          As String * 1                     ' Mixed file has display devices:
    ATMXDC          As String * 1                     ' Mixed file has communications d
    ATMXDB          As String * 1                     ' Mixed file has BSC devices:  N=
    ATFATR          As String * 6                     ' File attribute
    ATSYSN          As String * 8                     ' System Name (Source System, if
    ATASP           As Long                           ' Auxiliary storage pool ID:  1=S
    ATRES           As String * 4                     ' Reserved
    ATDTAT          As String * 1                     ' File type:  D=*DATA,  S=*SRC
    ATWAIT          As Long                           ' Maximum file wait time:  -1=*IM
    ATWATR          As Long                           ' Maximum record wait time:  -1=*
    ATSHAR          As String * 1                     ' Share open data path:  N=*NO, Y
    ATLVLC          As String * 1                     ' Record format level check:  N=*
    ATTXT           As String * 50                    ' Text 'description'
    ATNOFM          As Long                           ' Number of record formats
    ATFCCN          As String * 1                     ' Century created:  0=19xx, 1=20x
    ATFCDT          As String * 6                     ' Date created:  year/month/day
    ATFCTM          As String * 6                     ' Time created:  hour/minute/seco
    ATFLS           As String * 1                     ' Externally described file:  N=N
    ATICAP          As String * 1                     ' DBCS capable: N=No, Y=Yes
    ATRES2          As String * 9                     ' Reserved
    ATAQDV          As String * 10                    ' Program device to acquire
    ATMXDV          As Long                           ' Maximum number of devices or pr
    ATSPOL          As String * 1                     ' Spool the data:  N=*NO, Y=*YES
    ATNODV          As Long                           ' Number of devices
    ATUBL           As Long                           ' User buffer length
    ATIDTA          As String * 1                     ' DBCS data: N=*NO, Y=*YES
    ATRES3          As String * 9                     ' Reserved
    ATACCP          As String * 1                     ' Access path: A=Arrival K=Keyed
    ATSELO          As String * 1                     ' Select/omit file: N=No, Y=Yes
    ATCSEQ          As String * 1                     ' Alternative collating sequence:
    ATNOMB          As Long                           ' Number of members
    ATJOIN          As String * 1                     ' Join logical file:  N=No, Y=Yes
    ATSQLT          As String * 1                     ' SQL file type: 0=None, T=TABLE,
    ATRES4          As String * 8                     ' Reserved
End Type
    
    
Public arrDSPFDY0() As typeDSPFDY0
Public arrDSPFDY0_NB As Integer
Public arrDSPFDY0_NBMax As Integer
Public arrDSPFDY0_Index As Integer
Public arrDSPFDY0_Suite As Boolean


Public Sub srvDSPFDY0_ElpDisplay(recDSPFDY0 As typeDSPFDY0)
frmElpDisplay.fgData.Rows = 42
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATRCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval century:  0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATRCEN
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATRDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval date:  year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATRDAT
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATRTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval time:  hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATRTIM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATFILE   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATFILE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATLIB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATLIB
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATFTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "D=Device, P=PF, L=LF, R=DDM PF, S=DDM LF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATFTYP
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATFILA    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File attribute"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATFILA
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATMXDD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mixed file has display devices: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATMXDD
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATMXDC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mixed file has communications devices: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATMXDC
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATMXDB    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mixed file has BSC devices:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATMXDB
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATFATR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File attribute"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATFATR
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATSYSN    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "System Name (Source System, if file is DDM)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATSYSN
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATASP    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Auxiliary storage pool ID:  1=System ASP"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATASP
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATRES    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATRES
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATDTAT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File type:  D=*DATA,  S=*SRC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATDTAT
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATWAIT    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum file wait time:  -1=*IMMED, 0=*CLS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATWAIT
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATWATR    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum record wait time:  -1=*IMMED, -3=*NOMAX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATWATR
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATSHAR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Share open data path:  N=*NO, Y=*YES, ' '=DB file"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATSHAR
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATLVLC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Record format level check:  N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATLVLC
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATTXT   50A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Text 'description'"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATTXT
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATNOFM    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of record formats"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATNOFM
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATFCCN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Century created:  0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATFCCN
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATFCDT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Date created:  year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATFCDT
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATFCTM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Time created:  hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATFCTM
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATFLS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Externally described file:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATFLS
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATICAP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DBCS capable: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATICAP
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATRES2    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATRES2
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATAQDV   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Program device to acquire"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATAQDV
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATMXDV    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum number of devices or program devices"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATMXDV
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATSPOL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Spool the data:  N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATSPOL
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATNODV    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of devices"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATNODV
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATUBL    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "User buffer length"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATUBL
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATIDTA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DBCS data: N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATIDTA
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATRES3    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATRES3
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATACCP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path: A=Arrival K=Keyed E=EVI S=Shared"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATACCP
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATSELO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Select/omit file: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATSELO
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATCSEQ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Alternative collating sequence:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATCSEQ
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATNOMB    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of members"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATNOMB
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATJOIN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Join logical file:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATJOIN
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATSQLT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SQL file type: 0=None, T=TABLE, I=INDEX, V=VIEW"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATSQLT
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ATRES4    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY0.ATRES4
frmElpDisplay.Show vbModal
End Sub
     
'-----------------------------------------------------
Function srvDSPFDY0_Update(recDSPFDY0 As typeDSPFDY0)
'-----------------------------------------------------

srvDSPFDY0_Update = "?"

MsgTxtLen = 0
Call srvDSPFDY0_PutBuffer(recDSPFDY0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvDSPFDY0_GetBuffer(recDSPFDY0)) Then
        Call srvDSPFDY0_Error(recDSPFDY0)
        srvDSPFDY0_Update = recDSPFDY0.Err
        Exit Function
    Else
        srvDSPFDY0_Update = Null
    End If
Else
    recDSPFDY0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvDSPFDY0_Error(recDSPFDY0 As typeDSPFDY0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DSPFDY0" & Chr$(10) & Chr$(13)

Select Case mId$(recDSPFDY0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDSPFDY0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : DSPFDY0s.bas  ( " & Trim(recDSPFDY0.obj) & " : " & Trim(recDSPFDY0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvDSPFDY0_Monitor(recDSPFDY0 As typeDSPFDY0)
'-----------------------------------------------------

arrDSPFDY0_Suite = False
Select Case mId$(Trim(recDSPFDY0.Method), 1, 4)
    Case "Snap"
              srvDSPFDY0_Monitor = srvDSPFDY0_Snap(recDSPFDY0)
    Case Else
            srvDSPFDY0_Monitor = srvDSPFDY0_Seek(recDSPFDY0)
End Select

End Function

'---------------------------------------------------------
Public Function srvDSPFDY0_GetBuffer(recDSPFDY0 As typeDSPFDY0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDSPFDY0_GetBuffer = Null
recDSPFDY0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recDSPFDY0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recDSPFDY0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDSPFDY0.Err = Space$(10) Then
    recDSPFDY0.ATRCEN = mId$(MsgTxt, K + 1, 1)
    recDSPFDY0.ATRDAT = mId$(MsgTxt, K + 2, 6)
    recDSPFDY0.ATRTIM = mId$(MsgTxt, K + 8, 6)
    recDSPFDY0.ATFILE = mId$(MsgTxt, K + 14, 10)
    recDSPFDY0.ATLIB = mId$(MsgTxt, K + 24, 10)
    recDSPFDY0.ATFTYP = mId$(MsgTxt, K + 34, 1)
    recDSPFDY0.ATFILA = mId$(MsgTxt, K + 35, 4)
    recDSPFDY0.ATMXDD = mId$(MsgTxt, K + 39, 1)
    recDSPFDY0.ATMXDC = mId$(MsgTxt, K + 40, 1)
    recDSPFDY0.ATMXDB = mId$(MsgTxt, K + 41, 1)
    recDSPFDY0.ATFATR = mId$(MsgTxt, K + 42, 6)
    recDSPFDY0.ATSYSN = mId$(MsgTxt, K + 48, 8)
    recDSPFDY0.ATASP = CLng(Val(mId$(MsgTxt, K + 56, 4)))
    recDSPFDY0.ATRES = mId$(MsgTxt, K + 60, 4)
    recDSPFDY0.ATDTAT = mId$(MsgTxt, K + 64, 1)
    recDSPFDY0.ATWAIT = CLng(Val(mId$(MsgTxt, K + 65, 6)))
    recDSPFDY0.ATWATR = CLng(Val(mId$(MsgTxt, K + 71, 6)))
    recDSPFDY0.ATSHAR = mId$(MsgTxt, K + 77, 1)
    recDSPFDY0.ATLVLC = mId$(MsgTxt, K + 78, 1)
    recDSPFDY0.ATTXT = mId$(MsgTxt, K + 79, 50)
    recDSPFDY0.ATNOFM = CLng(Val(mId$(MsgTxt, K + 129, 6)))
    recDSPFDY0.ATFCCN = mId$(MsgTxt, K + 135, 1)
    recDSPFDY0.ATFCDT = mId$(MsgTxt, K + 136, 6)
    recDSPFDY0.ATFCTM = mId$(MsgTxt, K + 142, 6)
    recDSPFDY0.ATFLS = mId$(MsgTxt, K + 148, 1)
    recDSPFDY0.ATICAP = mId$(MsgTxt, K + 149, 1)
    recDSPFDY0.ATRES2 = mId$(MsgTxt, K + 150, 9)
    recDSPFDY0.ATAQDV = mId$(MsgTxt, K + 159, 10)
    recDSPFDY0.ATMXDV = CLng(Val(mId$(MsgTxt, K + 169, 4)))
    recDSPFDY0.ATSPOL = mId$(MsgTxt, K + 173, 1)
    recDSPFDY0.ATNODV = CLng(Val(mId$(MsgTxt, K + 174, 4)))
    recDSPFDY0.ATUBL = CLng(Val(mId$(MsgTxt, K + 178, 6)))
    recDSPFDY0.ATIDTA = mId$(MsgTxt, K + 184, 1)
    recDSPFDY0.ATRES3 = mId$(MsgTxt, K + 185, 9)
    recDSPFDY0.ATACCP = mId$(MsgTxt, K + 194, 1)
    recDSPFDY0.ATSELO = mId$(MsgTxt, K + 195, 1)
    recDSPFDY0.ATCSEQ = mId$(MsgTxt, K + 196, 1)
    recDSPFDY0.ATNOMB = CLng(Val(mId$(MsgTxt, K + 197, 6)))
    recDSPFDY0.ATJOIN = mId$(MsgTxt, K + 203, 1)
    recDSPFDY0.ATSQLT = mId$(MsgTxt, K + 204, 1)
    recDSPFDY0.ATRES4 = mId$(MsgTxt, K + 205, 8)

Else
    srvDSPFDY0_GetBuffer = recDSPFDY0.Err
End If

MsgTxtIndex = MsgTxtIndex + recDSPFDY0Len

End Function

'---------------------------------------------------------
Private Sub srvDSPFDY0_PutBuffer(recDSPFDY0 As typeDSPFDY0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDSPFDY0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDSPFDY0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 1) = recDSPFDY0.ATRCEN
    Mid$(MsgTxt, K + 2, 6) = recDSPFDY0.ATRDAT
    Mid$(MsgTxt, K + 8, 6) = recDSPFDY0.ATRTIM
    Mid$(MsgTxt, K + 14, 10) = recDSPFDY0.ATFILE
    Mid$(MsgTxt, K + 24, 10) = recDSPFDY0.ATLIB
    Mid$(MsgTxt, K + 34, 1) = recDSPFDY0.ATFTYP
    Mid$(MsgTxt, K + 35, 4) = recDSPFDY0.ATFILA
    Mid$(MsgTxt, K + 39, 1) = recDSPFDY0.ATMXDD
    Mid$(MsgTxt, K + 40, 1) = recDSPFDY0.ATMXDC
    Mid$(MsgTxt, K + 41, 1) = recDSPFDY0.ATMXDB
    Mid$(MsgTxt, K + 42, 6) = recDSPFDY0.ATFATR
    Mid$(MsgTxt, K + 48, 8) = recDSPFDY0.ATSYSN
    Mid$(MsgTxt, K + 56, 4) = Format$(recDSPFDY0.ATASP, "000 ")
    Mid$(MsgTxt, K + 60, 4) = recDSPFDY0.ATRES
    Mid$(MsgTxt, K + 64, 1) = recDSPFDY0.ATDTAT
    Mid$(MsgTxt, K + 65, 6) = Format$(recDSPFDY0.ATWAIT, "00000 ")
    Mid$(MsgTxt, K + 71, 6) = Format$(recDSPFDY0.ATWATR, "00000 ")
    Mid$(MsgTxt, K + 77, 1) = recDSPFDY0.ATSHAR
    Mid$(MsgTxt, K + 78, 1) = recDSPFDY0.ATLVLC
    Mid$(MsgTxt, K + 79, 50) = recDSPFDY0.ATTXT
    Mid$(MsgTxt, K + 129, 6) = Format$(recDSPFDY0.ATNOFM, "00000 ")
    Mid$(MsgTxt, K + 135, 1) = recDSPFDY0.ATFCCN
    Mid$(MsgTxt, K + 136, 6) = recDSPFDY0.ATFCDT
    Mid$(MsgTxt, K + 142, 6) = recDSPFDY0.ATFCTM
    Mid$(MsgTxt, K + 148, 1) = recDSPFDY0.ATFLS
    Mid$(MsgTxt, K + 149, 1) = recDSPFDY0.ATICAP
    Mid$(MsgTxt, K + 150, 9) = recDSPFDY0.ATRES2
    Mid$(MsgTxt, K + 159, 10) = recDSPFDY0.ATAQDV
    Mid$(MsgTxt, K + 169, 4) = Format$(recDSPFDY0.ATMXDV, "000 ")
    Mid$(MsgTxt, K + 173, 1) = recDSPFDY0.ATSPOL
    Mid$(MsgTxt, K + 174, 4) = Format$(recDSPFDY0.ATNODV, "000 ")
    Mid$(MsgTxt, K + 178, 6) = Format$(recDSPFDY0.ATUBL, "00000 ")
    Mid$(MsgTxt, K + 184, 1) = recDSPFDY0.ATIDTA
    Mid$(MsgTxt, K + 185, 9) = recDSPFDY0.ATRES3
    Mid$(MsgTxt, K + 194, 1) = recDSPFDY0.ATACCP
    Mid$(MsgTxt, K + 195, 1) = recDSPFDY0.ATSELO
    Mid$(MsgTxt, K + 196, 1) = recDSPFDY0.ATCSEQ
    Mid$(MsgTxt, K + 197, 6) = Format$(recDSPFDY0.ATNOMB, "00000 ")
    Mid$(MsgTxt, K + 203, 1) = recDSPFDY0.ATJOIN
    Mid$(MsgTxt, K + 204, 1) = recDSPFDY0.ATSQLT
    Mid$(MsgTxt, K + 205, 8) = recDSPFDY0.ATRES4

    

MsgTxtLen = MsgTxtLen + recDSPFDY0Len
End Sub



'---------------------------------------------------------
Private Function srvDSPFDY0_Seek(recDSPFDY0 As typeDSPFDY0)
'---------------------------------------------------------

srvDSPFDY0_Seek = "?"
MsgTxtLen = 0
Call srvDSPFDY0_PutBuffer(recDSPFDY0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDSPFDY0_GetBuffer(recDSPFDY0)) Then
        srvDSPFDY0_Seek = Null
    Else
        Call srvDSPFDY0_Error(recDSPFDY0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDSPFDY0_Snap(recDSPFDY0 As typeDSPFDY0)
'---------------------------------------------------------
srvDSPFDY0_Snap = "?"
MsgTxtLen = 0
Call srvDSPFDY0_PutBuffer(recDSPFDY0)
Call srvDSPFDY0_PutBuffer(arrDSPFDY0(0))
If IsNull(SndRcv()) Then
    srvDSPFDY0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDSPFDY0_GetBuffer(recDSPFDY0)) Then
            Call arrDSPFDY0_AddItem(recDSPFDY0)
            arrDSPFDY0_Suite = True
        Else
            arrDSPFDY0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrDSPFDY0_AddItem(recDSPFDY0 As typeDSPFDY0)
'---------------------------------------------------------
          
arrDSPFDY0_NB = arrDSPFDY0_NB + 1
    
If arrDSPFDY0_NB > arrDSPFDY0_NBMax Then
    arrDSPFDY0_NBMax = arrDSPFDY0_NBMax + recDSPFDY0_Block
    ReDim Preserve arrDSPFDY0(arrDSPFDY0_NBMax)
End If
            
arrDSPFDY0(arrDSPFDY0_NB) = recDSPFDY0
End Sub



'---------------------------------------------------------
Public Sub recDSPFDY0_Init(recDSPFDY0 As typeDSPFDY0)
'---------------------------------------------------------
recDSPFDY0.obj = "DSPFDY0"
recDSPFDY0.Method = ""
recDSPFDY0.Err = ""

End Sub




