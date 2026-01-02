Attribute VB_Name = "srvCptEAR"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCptEARLen = 560 ' 34 + 526
Public Const recCptEAR_Block = 50

Type typeCptEAR
    obj         As String * 12
    Method      As String * 12
    Err         As String * 10
    
    COSOC       As String * 3
    Agence      As String * 3
    Devise      As String * 3
    Compte      As String * 11
    AGEMET      As String * 3
    BIACOP      As String * 4
    SERVIC      As String * 3
    MONDEV      As Currency
    AMJSAI      As String * 8
    AMJOPE      As String * 8
    AMJVAL      As String * 8
    NUMPIE      As String * 7
    NOLIGN      As String * 4
    NUMLOT      As String * 4
    LIBELE      As String * 50
    BDFSTA      As String * 1
    BDFNUM      As String * 3
    EXONER      As String * 1
    SIGENE      As String * 1
    EDAVIS      As String * 1
    NUMCAI      As String * 2
    CPTOPE      As String * 3
    JJCPLT      As String * 1
    SYSGES      As String * 1
    ANCSLD      As Currency
    TRAITAMJ    As String * 8
    INICLI      As String * 1
    NOPROG      As String * 10
    NOMOP       As String * 10
    COVAL       As String * 6
    REFDOS      As String * 5
    SIEXTR     As String * 1
    LIBEL1      As String * 32
    LIBEL2      As String * 32
    LIBEL3      As String * 32
    LIBEL4      As String * 32
    CENPRF      As String * 6
    CODPRO      As String * 3
    REFCON      As String * 16
    RACCPA      As String * 5
    
    
    EARCptOri   As String * 11
    EARCptDes   As String * 11
    EARCptEAR   As String * 11
    EARIdRef    As Long
    EARStatus   As String * 3
    EARNumLot   As String * 4
    EARNumPie   As String * 7
    EARNoLign   As String * 4
    
    EARCptAmj   As String * 8
    EARElpUpd   As Integer
    LogCpteur   As Long
    LogCodErr   As String * 12

    EARMajUsr   As String * 20
    EARMajAmj   As String * 8
    EARMajHms   As String * 6
    EARValUsr   As String * 20
    EARValAmj   As String * 8
    EARValHms   As String * 6

End Type

Public arrCptEAR() As typeCptEAR
Public arrCptEAR_Nb As Integer
Public arrCptEAR_NbMax As Integer
Public arrCptEAR_Index As Integer
Public arrCptEAR_Suite As Boolean


Public Sub srvCptEAR_ElpDisplay(recCptEAR As typeCptEAR)
frmElpDisplay.fgData.Rows = 24
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.Err
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARCPTORi"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARCptOri

frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARCptDes"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARCptDes
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARCptEAR"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARCptEAR
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARIdRef "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARIdRef
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARStatus"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARStatus

frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARNumLot"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARNumLot
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARNumPie"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARNumPie
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARNoLign"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARNoLign
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARCptAmj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARCptAmj
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARElpUpd "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARElpUpd
frmElpDisplay.fgData.Row = 14
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LogCpteur"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.LogCpteur
frmElpDisplay.fgData.Row = 15
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LogCodErr"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.LogCodErr
frmElpDisplay.fgData.Row = 16
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARMajUsr"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARMajUsr
frmElpDisplay.fgData.Row = 17
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARMajAmj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARMajAmj
frmElpDisplay.fgData.Row = 18
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARMajHms"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARMajHms
frmElpDisplay.fgData.Row = 19
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARValUsr"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARValUsr
frmElpDisplay.fgData.Row = 20
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARValAmj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARValAmj
frmElpDisplay.fgData.Row = 21
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EARValHms"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCptEAR.EARValHms

frmElpDisplay.Show vbModal

End Sub

Public Sub srvCptEAR_Load(recCptEARMin As typeCptEAR, recCptEARMax As typeCptEAR)
Dim mMethod As String

mMethod = Trim(recCptEARMin.Method) & "+"
arrCptEAR_NbMax = 0
arrCptEAR_Suite = True: arrCptEAR_Nb = 0
arrCptEAR_NbMax = recCptEAR_Block: ReDim arrCptEAR(arrCptEAR_NbMax)

arrCptEAR(0) = recCptEARMax
arrCptEAR_Suite = True
Do Until Not arrCptEAR_Suite
    srvCptEAR_Monitor recCptEARMin
    recCptEARMin = arrCptEAR(arrCptEAR_Nb)
    recCptEARMin.Method = mMethod
Loop

End Sub


'-----------------------------------------------------
Function srvCptEAR_Update(recCptEAR As typeCptEAR)
'-----------------------------------------------------

srvCptEAR_Update = "?"

MsgTxtLen = 0
Call srvCptEAR_PutBuffer(recCptEAR)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCptEAR_GetBuffer(recCptEAR)) Then
        Call srvCptEAR_Error(recCptEAR)
        srvCptEAR_Update = recCptEAR.Err
        Exit Function
    Else
        srvCptEAR_Update = Null
    End If
Else
    recCptEAR.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Function srvCptEAR_Dtaq_Snd(recCptEAR As typeCptEAR)
'-----------------------------------------------------

srvCptEAR_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCptEAR_GetBuffer(recCptEAR)) Then
        Call srvCptEAR_Error(recCptEAR)
        srvCptEAR_Dtaq_Snd = recCptEAR.Err
        Exit Function
    Else
        srvCptEAR_Dtaq_Snd = Null
    End If
Else
    recCptEAR.Err = "Snd"
End If


'=====================================================
End Function


'-----------------------------------------------------
Function srvCptEAR_Dtaq_Put(lFct As String, recCptEAR As typeCptEAR)
'-----------------------------------------------------

srvCptEAR_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvCptEAR_PutBuffer(recCptEAR)
                If MsgTxtLen + recCptEARLen >= recCptEAR_Block * recCptEARLen Then
                    Call srvCptEAR_Dtaq_Snd(recCptEAR): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvCptEAR_Dtaq_Snd(recCptEAR)
    Case Else: srvCptEAR_Dtaq_Put = lFct
End Select
'=====================================================
End Function



'-----------------------------------------------------
Public Function srvCptEAR_Monitor(recCptEAR As typeCptEAR)
'-----------------------------------------------------

blnFR_Convert = False

arrCptEAR_Suite = False
Select Case mId$(Trim(recCptEAR.Method), 1, 4)
    Case "Seek", "Comp"
                srvCptEAR_Monitor = srvCptEAR_Seek(recCptEAR)
    Case "Snap"
              srvCptEAR_Monitor = srvCptEAR_Snap(recCptEAR)
    Case Else
                recCptEAR.Err = recCptEAR.Method
                Call srvCptEAR_Error(recCptEAR)
                srvCptEAR_Monitor = recCptEAR.Err
End Select
End Function

'-----------------------------------------------------
Sub srvCptEAR_Error(recCptEAR As typeCptEAR)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CptEAR" & Chr$(10) & Chr$(13)

Select Case mId$(recCptEAR.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCptEAR.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : CptEAR_Srv.bas  ( " _
                & Trim(recCptEAR.obj) & " : " & Trim(recCptEAR.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCptEAR_GetBuffer(recCptEAR As typeCptEAR)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCptEAR_GetBuffer = Null
recCptEAR.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCptEAR.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCptEAR.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCptEAR.Err = Space$(10) Then
  

    recCptEAR.COSOC = Format$(Val(mId$(MsgTxt, K + 1, 3)), "000")
    recCptEAR.Agence = Format$(Val(mId$(MsgTxt, K + 4, 3)), "000")
    recCptEAR.Devise = Format$(Val(mId$(MsgTxt, K + 7, 3)), "000")
    recCptEAR.Compte = Format$(Val(mId$(MsgTxt, K + 10, 11)), "00000000000")
    recCptEAR.AGEMET = Format$(Val(mId$(MsgTxt, K + 21, 3)), "000")
    recCptEAR.BIACOP = mId$(MsgTxt, K + 24, 4)
    recCptEAR.SERVIC = Format$(Val(mId$(MsgTxt, K + 28, 3)), "000")
    recCptEAR.MONDEV = CCur(Val(mId$(MsgTxt, K + 31, 19)))
    recCptEAR.AMJSAI = Format$(Val(mId$(MsgTxt, K + 50, 8)), "00000000")
    recCptEAR.AMJOPE = Format$(Val(mId$(MsgTxt, K + 58, 8)), "00000000")
    recCptEAR.AMJVAL = Format$(Val(mId$(MsgTxt, K + 66, 8)), "00000000")
    recCptEAR.NUMPIE = Format$(Val(mId$(MsgTxt, K + 74, 7)), "0000000")
    recCptEAR.NOLIGN = Format$(Val(mId$(MsgTxt, K + 81, 4)), "0000")
    recCptEAR.NUMLOT = Format$(Val(mId$(MsgTxt, K + 85, 4)), "0000")
    recCptEAR.LIBELE = mId$(MsgTxt, K + 89, 50)
    recCptEAR.BDFSTA = mId$(MsgTxt, K + 139, 1)
    recCptEAR.BDFNUM = mId$(MsgTxt, K + 140, 3)
    recCptEAR.EXONER = mId$(MsgTxt, K + 143, 1)
    recCptEAR.SIGENE = mId$(MsgTxt, K + 144, 1)
    recCptEAR.EDAVIS = mId$(MsgTxt, K + 145, 1)
    recCptEAR.NUMCAI = mId$(MsgTxt, K + 146, 2)
    recCptEAR.JJCPLT = mId$(MsgTxt, K + 148, 1)
    recCptEAR.CPTOPE = Format$(Val(mId$(MsgTxt, K + 149, 3)), "000")
    recCptEAR.SYSGES = mId$(MsgTxt, K + 152, 1)
    recCptEAR.ANCSLD = CCur(Val(mId$(MsgTxt, K + 153, 19)))
    recCptEAR.TRAITAMJ = Format$(Val(mId$(MsgTxt, K + 172, 8)), "00000000")
    recCptEAR.INICLI = mId$(MsgTxt, K + 180, 1)
    recCptEAR.NOPROG = mId$(MsgTxt, K + 181, 10)
    recCptEAR.NOMOP = mId$(MsgTxt, K + 191, 10)
    recCptEAR.COVAL = Format$(Val(mId$(MsgTxt, K + 201, 6)), "000000")
    
    recCptEAR.REFDOS = mId$(MsgTxt, K + 207, 5)
    recCptEAR.SIEXTR = mId$(MsgTxt, K + 212, 1)
    recCptEAR.LIBEL1 = mId$(MsgTxt, K + 213, 32)
    recCptEAR.LIBEL2 = mId$(MsgTxt, K + 245, 32)
    recCptEAR.LIBEL3 = mId$(MsgTxt, K + 277, 32)
    recCptEAR.LIBEL4 = mId$(MsgTxt, K + 309, 32)
    recCptEAR.CENPRF = Format$(Val(mId$(MsgTxt, K + 341, 6)), "000000")
    recCptEAR.CODPRO = mId$(MsgTxt, K + 347, 3)
    recCptEAR.REFCON = mId$(MsgTxt, K + 350, 16)
    recCptEAR.RACCPA = Format$(Val(mId$(MsgTxt, K + 366, 5)), "00000")
    
    recCptEAR.EARCptOri = Format$(Val(mId$(MsgTxt, K + 371, 11)), "00000000000")
    recCptEAR.EARCptDes = Format$(Val(mId$(MsgTxt, K + 382, 11)), "00000000000")
    recCptEAR.EARCptEAR = Format$(Val(mId$(MsgTxt, K + 393, 11)), "00000000000")
    recCptEAR.EARIdRef = Format$(Val(mId$(MsgTxt, K + 404, 7)), "0000000")
    recCptEAR.EARStatus = mId$(MsgTxt, K + 411, 3)
    recCptEAR.EARNumLot = Format$(Val(mId$(MsgTxt, K + 414, 4)), "0000")
    recCptEAR.EARNumPie = Format$(Val(mId$(MsgTxt, K + 418, 7)), "0000000")
    recCptEAR.EARNoLign = Format$(Val(mId$(MsgTxt, K + 425, 4)), "0000")
    
    recCptEAR.EARCptAmj = Format$(Val(mId$(MsgTxt, K + 429, 8)), "00000000")
    recCptEAR.EARElpUpd = Format$(Val(mId$(MsgTxt, K + 437, 3)), "000")
    recCptEAR.LogCpteur = Format$(Val(mId$(MsgTxt, K + 440, 7)), "0000000")
    recCptEAR.LogCodErr = mId$(MsgTxt, K + 447, 12)
    
    recCptEAR.EARMajUsr = mId$(MsgTxt, K + 459, 20)
    recCptEAR.EARMajAmj = Format$(Val(mId$(MsgTxt, K + 479, 8)), "00000000")
    recCptEAR.EARMajHms = Format$(Val(mId$(MsgTxt, K + 487, 6)), "000000")
    recCptEAR.EARValUsr = mId$(MsgTxt, K + 493, 20)
    recCptEAR.EARValAmj = Format$(Val(mId$(MsgTxt, K + 513, 8)), "00000000")
    recCptEAR.EARValHms = Format$(Val(mId$(MsgTxt, K + 521, 6)), "000000")

Else
    srvCptEAR_GetBuffer = recCptEAR.Err
End If

MsgTxtIndex = MsgTxtIndex + recCptEARLen

End Function

'---------------------------------------------------------
Private Sub srvCptEAR_PutBuffer(recCptEAR As typeCptEAR)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCptEAR.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCptEAR.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = Format$(Val(recCptEAR.COSOC), "000")
Mid$(MsgTxt, K + 4, 3) = Format$(Val(recCptEAR.Agence), "000")
Mid$(MsgTxt, K + 7, 3) = Format$(Val(recCptEAR.Devise), "000")
Mid$(MsgTxt, K + 10, 11) = Format$(Val(recCptEAR.Compte), "00000000000")
Mid$(MsgTxt, K + 21, 3) = Format$(Val(recCptEAR.AGEMET), "000")
Mid$(MsgTxt, K + 24, 4) = recCptEAR.BIACOP
Mid$(MsgTxt, K + 28, 3) = Format$(Val(recCptEAR.SERVIC), "000")
Mid$(MsgTxt, K + 31, 19) = cur_19P(recCptEAR.MONDEV) '  Format$(recCptEAR.MONDEV * 100, "000000000000000000-")
Mid$(MsgTxt, K + 50, 8) = Format$(Val(recCptEAR.AMJSAI), "00000000")
Mid$(MsgTxt, K + 58, 8) = Format$(Val(recCptEAR.AMJOPE), "00000000")
Mid$(MsgTxt, K + 66, 8) = Format$(Val(recCptEAR.AMJVAL), "00000000")
Mid$(MsgTxt, K + 74, 7) = Format$(Val(recCptEAR.NUMPIE), "0000000")
Mid$(MsgTxt, K + 81, 4) = Format$(Val(recCptEAR.NOLIGN), "0000")
Mid$(MsgTxt, K + 85, 4) = Format$(Val(recCptEAR.NUMLOT), "0000")
Mid$(MsgTxt, K + 89, 50) = recCptEAR.LIBELE
Mid$(MsgTxt, K + 139, 1) = recCptEAR.BDFSTA
Mid$(MsgTxt, K + 140, 3) = recCptEAR.BDFNUM
Mid$(MsgTxt, K + 143, 1) = recCptEAR.EXONER
Mid$(MsgTxt, K + 144, 1) = recCptEAR.SIGENE
Mid$(MsgTxt, K + 145, 1) = recCptEAR.EDAVIS
Mid$(MsgTxt, K + 146, 2) = recCptEAR.NUMCAI
Mid$(MsgTxt, K + 148, 1) = recCptEAR.JJCPLT
Mid$(MsgTxt, K + 149, 3) = Format$(Val(recCptEAR.CPTOPE), "000")
Mid$(MsgTxt, K + 152, 1) = recCptEAR.SYSGES
Mid$(MsgTxt, K + 153, 19) = cur_19P(recCptEAR.ANCSLD) 'Format$(recCptEAR.ANCSLD * 100, "000000000000000000-")
Mid$(MsgTxt, K + 172, 8) = Format$(Val(recCptEAR.TRAITAMJ), "00000000")
Mid$(MsgTxt, K + 180, 1) = recCptEAR.INICLI
Mid$(MsgTxt, K + 181, 10) = recCptEAR.NOPROG
Mid$(MsgTxt, K + 191, 10) = recCptEAR.NOMOP
Mid$(MsgTxt, K + 201, 6) = Format$(Val(recCptEAR.COVAL), "000000")

Mid$(MsgTxt, K + 207, 5) = recCptEAR.REFDOS
Mid$(MsgTxt, K + 212, 1) = recCptEAR.SIEXTR
Mid$(MsgTxt, K + 213, 32) = recCptEAR.LIBEL1
Mid$(MsgTxt, K + 245, 32) = recCptEAR.LIBEL2
Mid$(MsgTxt, K + 277, 32) = recCptEAR.LIBEL3
Mid$(MsgTxt, K + 309, 32) = recCptEAR.LIBEL4
Mid$(MsgTxt, K + 341, 6) = Format$(Val(recCptEAR.CENPRF), "000000")
Mid$(MsgTxt, K + 347, 3) = recCptEAR.CODPRO
Mid$(MsgTxt, K + 350, 16) = recCptEAR.REFCON
Mid$(MsgTxt, K + 366, 5) = Format$(Val(recCptEAR.RACCPA), "00000")

Mid$(MsgTxt, K + 371, 11) = Format$(Val(recCptEAR.EARCptOri), "00000000000")
Mid$(MsgTxt, K + 382, 11) = Format$(Val(recCptEAR.EARCptDes), "00000000000")
Mid$(MsgTxt, K + 393, 11) = Format$(Val(recCptEAR.EARCptEAR), "00000000000")
Mid$(MsgTxt, K + 404, 7) = Format$(Val(recCptEAR.EARIdRef), "0000000")
Mid$(MsgTxt, K + 411, 3) = recCptEAR.EARStatus
Mid$(MsgTxt, K + 414, 4) = Format$(Val(recCptEAR.EARNumLot), "0000")
Mid$(MsgTxt, K + 418, 7) = Format$(Val(recCptEAR.EARNumPie), "0000000")
Mid$(MsgTxt, K + 425, 4) = Format$(Val(recCptEAR.EARNoLign), "0000")

Mid$(MsgTxt, K + 429, 8) = Format$(Val(recCptEAR.EARCptAmj), "00000000")
Mid$(MsgTxt, K + 437, 3) = Format$(Val(recCptEAR.EARElpUpd), "000")
Mid$(MsgTxt, K + 440, 7) = Format$(Val(recCptEAR.LogCpteur), "0000000")
Mid$(MsgTxt, K + 447, 12) = recCptEAR.LogCodErr

Mid$(MsgTxt, K + 459, 20) = recCptEAR.EARMajUsr
Mid$(MsgTxt, K + 479, 8) = Format$(Val(recCptEAR.EARMajAmj), "00000000")
Mid$(MsgTxt, K + 487, 6) = Format$(Val(recCptEAR.EARMajHms), "000000")
Mid$(MsgTxt, K + 493, 20) = recCptEAR.EARValUsr
Mid$(MsgTxt, K + 513, 8) = Format$(Val(recCptEAR.EARValAmj), "00000000")
Mid$(MsgTxt, K + 521, 6) = Format$(Val(recCptEAR.EARValHms), "000000")

MsgTxtLen = MsgTxtLen + recCptEARLen
End Sub



'---------------------------------------------------------
Private Function srvCptEAR_Seek(recCptEAR As typeCptEAR)
'---------------------------------------------------------

srvCptEAR_Seek = "?"
MsgTxtLen = 0
Call srvCptEAR_PutBuffer(recCptEAR)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCptEAR_GetBuffer(recCptEAR)) Then
        srvCptEAR_Seek = Null
    Else
        Call srvCptEAR_Error(recCptEAR)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCptEAR_Snap(recCptEAR As typeCptEAR)
'---------------------------------------------------------
srvCptEAR_Snap = "?"
MsgTxtLen = 0
Call srvCptEAR_PutBuffer(recCptEAR)
Call srvCptEAR_PutBuffer(arrCptEAR(0))
If IsNull(SndRcv()) Then
    srvCptEAR_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCptEAR_GetBuffer(recCptEAR)) Then
            Call arrCptEAR_AddItem(recCptEAR)
            arrCptEAR_Suite = True
        Else
            arrCptEAR_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCptEAR_Init(recCptEAR As typeCptEAR)
'---------------------------------------------------------
MsgTxt = Space$(recCptEARLen)
MsgTxtIndex = 0
Call srvCptEAR_GetBuffer(recCptEAR)
recCptEAR.obj = "SRVCPTEAR"
End Sub

'---------------------------------------------------------
Public Sub arrCptEAR_AddItem(recCptEAR As typeCptEAR)
'---------------------------------------------------------
          
arrCptEAR_Nb = arrCptEAR_Nb + 1
    
If arrCptEAR_Nb > arrCptEAR_NbMax Then
    arrCptEAR_NbMax = arrCptEAR_NbMax + 10
    ReDim Preserve arrCptEAR(arrCptEAR_NbMax)
End If
            
arrCptEAR(arrCptEAR_Nb) = recCptEAR
End Sub
