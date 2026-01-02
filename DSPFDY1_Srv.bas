Attribute VB_Name = "srvDSPFDY1"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDSPFDY1Len = 177 ' 34 + 143
Public Const recDSPFDY1_Block = 30

Type typeDSPFDY1
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
     
    APRCEN          As String * 1                     ' Retrieval century:  0=19xx, 1=2
    APRDAT          As String * 6                     ' Retrieval date:  year/month/day
    APRTIM          As String * 6                     ' Retrieval time:  hour/minute/se
    APFILE          As String * 10                    ' File
    APLIB           As String * 10                    ' Library
    APFTYP          As String * 1                     ' P=PF, L=LF, R=DDM PF, S=DDM LF
    APFILA          As String * 4                     ' File attribute:  *PHY or *LGL
    APMXD           As String * 3                     ' Reserved
    APFATR          As String * 6                     ' File attribute:  PF, LF, PF38,
    APSYSN          As String * 8                     ' System Name (Source System, if
    APASP           As Long                           ' Auxiliary storage pool ID:  1=S
    APRES           As String * 4                     ' Reserved
    APMANT          As String * 1                     ' Maintenance:  I=*IMMED, R=*REBL
    APUNIQ          As String * 1                     ' Keys must be unique: N=No, Y=Ye
    APKEYO          As String * 1                     ' L=LIFO, F=FIFO, C=FCFO, N=No sp
    APSELO          As String * 1                     ' Select/omit file:  N=No, Y=Yes
    APACCP          As String * 1                     ' Access path: A=Arrival K=Keyed
    APNSCO          As Long                           ' Number of files accessed by log
    APBOF           As String * 10                    ' Physical file
    APBOL           As String * 10                    ' Library
    APBOLF          As String * 10                    ' Logical file format through whi
    APNKYF          As Long                           ' Number of key fields per format
    APKEYF          As String * 10                    ' Key field name
    APKSEQ          As String * 1                     ' Key sequence: D=Descending, A=A
    APKSIN          As String * 1                     ' Key sign specified: N=UNSIGNED,
    APKZD           As String * 1                     ' Zone/digit specified: N=None, Z
    APKASQ          As String * 1                     ' Alternative collating sequence:
    APKEYN          As Long                           ' Key field number:  1=First key
    APJOIN          As String * 1                     ' Join logical file:  N=No, Y=Yes
    APACPJ          As String * 1                     ' Access path journaled:  N=No, Y
    APRIKY          As String * 1                     ' Constraint Type: P=PRIMARY, U=U
    APUUIV          As Long                           ' Number of unique key values giv
     
End Type
    
    
Public arrDSPFDY1() As typeDSPFDY1
Public arrDSPFDY1_NB As Integer
Public arrDSPFDY1_NBMax As Integer
Public arrDSPFDY1_Index As Integer
Public arrDSPFDY1_Suite As Boolean

'-----------------------------------------------------
Function srvDSPFDY1_Update(recDSPFDY1 As typeDSPFDY1)
'-----------------------------------------------------

srvDSPFDY1_Update = "?"

MsgTxtLen = 0
Call srvDSPFDY1_PutBuffer(recDSPFDY1)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvDSPFDY1_GetBuffer(recDSPFDY1)) Then
        Call srvDSPFDY1_Error(recDSPFDY1)
        srvDSPFDY1_Update = recDSPFDY1.Err
        Exit Function
    Else
        srvDSPFDY1_Update = Null
    End If
Else
    recDSPFDY1.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvDSPFDY1_Error(recDSPFDY1 As typeDSPFDY1)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DSPFDY1" & Chr$(10) & Chr$(13)

Select Case mId$(recDSPFDY1.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDSPFDY1.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : DSPFDY1s.bas  ( " & Trim(recDSPFDY1.obj) & " : " & Trim(recDSPFDY1.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvDSPFDY1_Monitor(recDSPFDY1 As typeDSPFDY1)
'-----------------------------------------------------

arrDSPFDY1_Suite = False
Select Case mId$(Trim(recDSPFDY1.Method), 1, 4)
    Case "Snap"
              srvDSPFDY1_Monitor = srvDSPFDY1_Snap(recDSPFDY1)
    Case Else
            srvDSPFDY1_Monitor = srvDSPFDY1_Seek(recDSPFDY1)
End Select

End Function

'---------------------------------------------------------
Public Function srvDSPFDY1_GetBuffer(recDSPFDY1 As typeDSPFDY1)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDSPFDY1_GetBuffer = Null
recDSPFDY1.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recDSPFDY1.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recDSPFDY1.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDSPFDY1.Err = Space$(10) Then
    recDSPFDY1.APRCEN = mId$(MsgTxt, K + 1, 1)
    recDSPFDY1.APRDAT = mId$(MsgTxt, K + 2, 6)
    recDSPFDY1.APRTIM = mId$(MsgTxt, K + 8, 6)
    recDSPFDY1.APFILE = mId$(MsgTxt, K + 14, 10)
    recDSPFDY1.APLIB = mId$(MsgTxt, K + 24, 10)
    recDSPFDY1.APFTYP = mId$(MsgTxt, K + 34, 1)
    recDSPFDY1.APFILA = mId$(MsgTxt, K + 35, 4)
    recDSPFDY1.APMXD = mId$(MsgTxt, K + 39, 3)
    recDSPFDY1.APFATR = mId$(MsgTxt, K + 42, 6)
    recDSPFDY1.APSYSN = mId$(MsgTxt, K + 48, 8)
    recDSPFDY1.APASP = CLng(Val(mId$(MsgTxt, K + 56, 4)))
    recDSPFDY1.APRES = mId$(MsgTxt, K + 60, 4)
    recDSPFDY1.APMANT = mId$(MsgTxt, K + 64, 1)
    recDSPFDY1.APUNIQ = mId$(MsgTxt, K + 65, 1)
    recDSPFDY1.APKEYO = mId$(MsgTxt, K + 66, 1)
    recDSPFDY1.APSELO = mId$(MsgTxt, K + 67, 1)
    recDSPFDY1.APACCP = mId$(MsgTxt, K + 68, 1)
    recDSPFDY1.APNSCO = CLng(Val(mId$(MsgTxt, K + 69, 4)))
    recDSPFDY1.APBOF = mId$(MsgTxt, K + 73, 10)
    recDSPFDY1.APBOL = mId$(MsgTxt, K + 83, 10)
    recDSPFDY1.APBOLF = mId$(MsgTxt, K + 93, 10)
    recDSPFDY1.APNKYF = CLng(Val(mId$(MsgTxt, K + 103, 4)))
    recDSPFDY1.APKEYF = mId$(MsgTxt, K + 107, 10)
    recDSPFDY1.APKSEQ = mId$(MsgTxt, K + 117, 1)
    recDSPFDY1.APKSIN = mId$(MsgTxt, K + 118, 1)
    recDSPFDY1.APKZD = mId$(MsgTxt, K + 119, 1)
    recDSPFDY1.APKASQ = mId$(MsgTxt, K + 120, 1)
    recDSPFDY1.APKEYN = CLng(Val(mId$(MsgTxt, K + 121, 4)))
    recDSPFDY1.APJOIN = mId$(MsgTxt, K + 125, 1)
    recDSPFDY1.APACPJ = mId$(MsgTxt, K + 126, 1)
    recDSPFDY1.APRIKY = mId$(MsgTxt, K + 127, 1)
    recDSPFDY1.APUUIV = CLng(Val(mId$(MsgTxt, K + 128, 16)))

Else
    srvDSPFDY1_GetBuffer = recDSPFDY1.Err
End If

MsgTxtIndex = MsgTxtIndex + recDSPFDY1Len

End Function

'---------------------------------------------------------
Private Sub srvDSPFDY1_PutBuffer(recDSPFDY1 As typeDSPFDY1)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDSPFDY1.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDSPFDY1.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 1) = recDSPFDY1.APRCEN
    Mid$(MsgTxt, K + 2, 6) = recDSPFDY1.APRDAT
    Mid$(MsgTxt, K + 8, 6) = recDSPFDY1.APRTIM
    Mid$(MsgTxt, K + 14, 10) = recDSPFDY1.APFILE
    Mid$(MsgTxt, K + 24, 10) = recDSPFDY1.APLIB
    Mid$(MsgTxt, K + 34, 1) = recDSPFDY1.APFTYP
    Mid$(MsgTxt, K + 35, 4) = recDSPFDY1.APFILA
    Mid$(MsgTxt, K + 39, 3) = recDSPFDY1.APMXD
    Mid$(MsgTxt, K + 42, 6) = recDSPFDY1.APFATR
    Mid$(MsgTxt, K + 48, 8) = recDSPFDY1.APSYSN
    Mid$(MsgTxt, K + 56, 4) = Format$(recDSPFDY1.APASP, "000 ")
    Mid$(MsgTxt, K + 60, 4) = recDSPFDY1.APRES
    Mid$(MsgTxt, K + 64, 1) = recDSPFDY1.APMANT
    Mid$(MsgTxt, K + 65, 1) = recDSPFDY1.APUNIQ
    Mid$(MsgTxt, K + 66, 1) = recDSPFDY1.APKEYO
    Mid$(MsgTxt, K + 67, 1) = recDSPFDY1.APSELO
    Mid$(MsgTxt, K + 68, 1) = recDSPFDY1.APACCP
    Mid$(MsgTxt, K + 69, 4) = Format$(recDSPFDY1.APNSCO, "000 ")
    Mid$(MsgTxt, K + 73, 10) = recDSPFDY1.APBOF
    Mid$(MsgTxt, K + 83, 10) = recDSPFDY1.APBOL
    Mid$(MsgTxt, K + 93, 10) = recDSPFDY1.APBOLF
    Mid$(MsgTxt, K + 103, 4) = Format$(recDSPFDY1.APNKYF, "000 ")
    Mid$(MsgTxt, K + 107, 10) = recDSPFDY1.APKEYF
    Mid$(MsgTxt, K + 117, 1) = recDSPFDY1.APKSEQ
    Mid$(MsgTxt, K + 118, 1) = recDSPFDY1.APKSIN
    Mid$(MsgTxt, K + 119, 1) = recDSPFDY1.APKZD
    Mid$(MsgTxt, K + 120, 1) = recDSPFDY1.APKASQ
    Mid$(MsgTxt, K + 121, 4) = Format$(recDSPFDY1.APKEYN, "000 ")
    Mid$(MsgTxt, K + 125, 1) = recDSPFDY1.APJOIN
    Mid$(MsgTxt, K + 126, 1) = recDSPFDY1.APACPJ
    Mid$(MsgTxt, K + 127, 1) = recDSPFDY1.APRIKY
    Mid$(MsgTxt, K + 128, 16) = Format$(recDSPFDY1.APUUIV, "000000000000000 ")
    

MsgTxtLen = MsgTxtLen + recDSPFDY1Len
End Sub


Public Sub srvDSPFDY1_ElpDisplay(recDSPFDY1 As typeDSPFDY1)
frmElpDisplay.fgData.Rows = 33
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APRCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval century:  0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APRCEN
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APRDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval date:  year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APRDAT
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APRTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval time:  hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APRTIM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APFILE   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APFILE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APLIB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APLIB
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APFTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "P=PF, L=LF, R=DDM PF, S=DDM LF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APFTYP
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APFILA    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File attribute:  *PHY or *LGL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APFILA
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APMXD    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APMXD
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APFATR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File attribute:  PF, LF, PF38, or LF38"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APFATR
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APSYSN    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "System Name (Source System, if file is DDM)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APSYSN
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APASP    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Auxiliary storage pool ID:  1=System ASP"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APASP
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APRES    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APRES
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APMANT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maintenance:  I=*IMMED, R=*REBLD, D=*DLY"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APMANT
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APUNIQ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Keys must be unique: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APUNIQ
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APKEYO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "L=LIFO, F=FIFO, C=FCFO, N=No specific key order"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APKEYO
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APSELO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Select/omit file:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APSELO
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APACCP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path: A=Arrival K=Keyed E=EVI S=Shared"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APACCP
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APNSCO    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of files accessed by logical file"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APNSCO
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APBOF   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Physical file"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APBOF
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APBOL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APBOL
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APBOLF   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Logical file format through which data is accessed"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APBOLF
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APNKYF    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of key fields per format"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APNKYF
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APKEYF   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Key field name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APKEYF
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APKSEQ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Key sequence: D=Descending, A=Ascending"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APKSEQ
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APKSIN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Key sign specified: N=UNSIGNED, S=SIGNED, A=ABSVAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APKSIN
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APKZD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Zone/digit specified: N=None, Z=ZONE, D=DIGIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APKZD
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APKASQ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Alternative collating sequence:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APKASQ
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APKEYN    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Key field number:  1=First key in format"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APKEYN
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APJOIN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Join logical file:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APJOIN
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APACPJ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path journaled:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APACPJ
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APRIKY    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Constraint Type: P=PRIMARY, U=UNIQUE, N=NONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APRIKY
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "APUUIV   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of unique key values given at file creation"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY1.APUUIV
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvDSPFDY1_Seek(recDSPFDY1 As typeDSPFDY1)
'---------------------------------------------------------

srvDSPFDY1_Seek = "?"
MsgTxtLen = 0
Call srvDSPFDY1_PutBuffer(recDSPFDY1)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDSPFDY1_GetBuffer(recDSPFDY1)) Then
        srvDSPFDY1_Seek = Null
    Else
        Call srvDSPFDY1_Error(recDSPFDY1)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDSPFDY1_Snap(recDSPFDY1 As typeDSPFDY1)
'---------------------------------------------------------
srvDSPFDY1_Snap = "?"
MsgTxtLen = 0
Call srvDSPFDY1_PutBuffer(recDSPFDY1)
Call srvDSPFDY1_PutBuffer(arrDSPFDY1(0))
If IsNull(SndRcv()) Then
    srvDSPFDY1_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDSPFDY1_GetBuffer(recDSPFDY1)) Then
            Call arrDSPFDY1_AddItem(recDSPFDY1)
            arrDSPFDY1_Suite = True
        Else
            arrDSPFDY1_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrDSPFDY1_AddItem(recDSPFDY1 As typeDSPFDY1)
'---------------------------------------------------------
          
arrDSPFDY1_NB = arrDSPFDY1_NB + 1
    
If arrDSPFDY1_NB > arrDSPFDY1_NBMax Then
    arrDSPFDY1_NBMax = arrDSPFDY1_NBMax + recDSPFDY1_Block
    ReDim Preserve arrDSPFDY1(arrDSPFDY1_NBMax)
End If
            
arrDSPFDY1(arrDSPFDY1_NB) = recDSPFDY1
End Sub



'---------------------------------------------------------
Public Sub recDSPFDY1_Init(recDSPFDY1 As typeDSPFDY1)
'---------------------------------------------------------
recDSPFDY1.obj = "DSPFDY1"
recDSPFDY1.Method = ""
recDSPFDY1.Err = ""

End Sub





