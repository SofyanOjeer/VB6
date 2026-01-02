Attribute VB_Name = "srvCpj030w0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCpj030W0Len = 404 ' 34 + 370

Type typeCpj030W0
    obj         As String * 12
    Method      As String * 12
    Err         As String * 10
    
    COSOC       As String * 3
    CODENR      As String * 1
    AGEMET      As String * 3
    SERVIC      As String * 3
    BIACOP      As String * 4
    NUMLOT      As String * 4
    NUMPIE      As String * 7
    NOLIGN      As String * 4
    Agence      As String * 3
    Devise      As String * 4
    Compte      As String * 11
    SENECR      As String * 1
    MONDEV      As Currency
    LIBELE      As String * 50
    BDFSTA      As String * 1
    BDFNUM      As String * 3
    EXONER      As String * 1
    SIGENE      As String * 1
    AMJSAI      As String * 8
    AMJVAL      As String * 8
    AMJOPE      As String * 8
    INTERV      As String * 1
    MAJVAL      As String * 1
    OPOCHQ      As String * 1
    JJCPLT      As String * 1
    CODFOR      As String * 1
    FOROPO      As String * 1
    FORVAL      As String * 1
    INICLI      As String * 1
    EDAVIS      As String * 1
    NOMOP       As String * 10
    COVAL       As String * 6
    REFDOS      As String * 5
    LIBEL1      As String * 32
    LIBEL2      As String * 32
    LIBEL3      As String * 32
    LIBEL4      As String * 32
    CENPRF      As String * 6
    CODPRO      As String * 3
    REFCON      As String * 16
    RACCPA      As String * 5
    CTLSTA      As String * 1
    CTLNOM      As String * 10
    CTLAMJ      As String * 8
    CTLHMS      As String * 6
    IMPAMJ      As String * 8
    IMPHMS      As String * 6

End Type
    
Public arrCpj030W0() As typeCpj030W0
Public arrCpj030W0Nb As Integer
Public arrCpj030W0NbMax As Integer
Public arrCpj030W0Index As Integer
Public arrCpj030W0Suite As Boolean

'-----------------------------------------------------
Function srvCpj030W0_Update(recCpj030W0 As typeCpj030W0)
'-----------------------------------------------------

srvCpj030W0_Update = "?"

MsgTxtLen = 0
Call srvCpj030W0_PutBuffer(recCpj030W0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCpj030W0_GetBuffer(recCpj030W0)) Then
        Call srvCpj030W0_Error(recCpj030W0)
        srvCpj030W0_Update = recCpj030W0.Err
        Exit Function
    Else
        srvCpj030W0_Update = Null
    End If
Else
    recCpj030W0.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Function srvCpj030W0_Dtaq_Snd(recCpj030W0 As typeCpj030W0)
'-----------------------------------------------------

srvCpj030W0_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCpj030W0_GetBuffer(recCpj030W0)) Then
        Call srvCpj030W0_Error(recCpj030W0)
        srvCpj030W0_Dtaq_Snd = recCpj030W0.Err
        Exit Function
    Else
        srvCpj030W0_Dtaq_Snd = Null
    End If
Else
    recCpj030W0.Err = "Snd"
End If


'=====================================================
End Function


'-----------------------------------------------------
Function srvCpj030W0_Dtaq_Put(lFct As String, recCpj030W0 As typeCpj030W0)
'-----------------------------------------------------

srvCpj030W0_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvCpj030W0_PutBuffer(recCpj030W0)
                If MsgTxtLen + recCpj030W0Len >= 15 * recCpj030W0Len Then
                    Call srvCpj030W0_Dtaq_Snd(recCpj030W0): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvCpj030W0_Dtaq_Snd(recCpj030W0)
    Case Else: srvCpj030W0_Dtaq_Put = lFct
End Select
'=====================================================
End Function



'-----------------------------------------------------
Public Function srvCpj030W0_Mon(recCpj030W0 As typeCpj030W0)
'-----------------------------------------------------

arrCpj030W0Suite = False
Select Case Trim(recCpj030W0.Method)
    Case "SeekLA"
                srvCpj030W0_Mon = srvCpj030W0_Seek(recCpj030W0)
    Case "SnapJA", "SnapJA+", "PrevJA", "PrevJA+"
              srvCpj030W0_Mon = srvCpj030W0_Snap(recCpj030W0)
    Case "SnapLA", "SnapLA+", "PrevLA", "PrevLA+"
              srvCpj030W0_Mon = srvCpj030W0_Snap(recCpj030W0)
    Case Else
                recCpj030W0.Err = recCpj030W0.Method
                Call srvCpj030W0_Error(recCpj030W0)
                srvCpj030W0_Mon = recCpj030W0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvCpj030W0_Error(recCpj030W0 As typeCpj030W0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Cpj030W0" & Chr$(10) & Chr$(13)

Select Case mId$(recCpj030W0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCpj030W0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : Cpj030W0s.bas  ( " _
                & Trim(recCpj030W0.obj) & " : " & Trim(recCpj030W0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCpj030W0_GetBuffer(recCpj030W0 As typeCpj030W0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCpj030W0_GetBuffer = Null
recCpj030W0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCpj030W0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCpj030W0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCpj030W0.Err = Space$(10) Then
    
    recCpj030W0.COSOC = Format$(Val(mId$(MsgTxt, K + 1, 3)), "000")
    recCpj030W0.CODENR = mId$(MsgTxt, K + 4, 1)
    recCpj030W0.AGEMET = Format$(Val(mId$(MsgTxt, K + 5, 3)), "000")
    recCpj030W0.SERVIC = Format$(Val(mId$(MsgTxt, K + 8, 3)), "000")
    recCpj030W0.BIACOP = mId$(MsgTxt, K + 11, 4)
    recCpj030W0.NUMLOT = Format$(Val(mId$(MsgTxt, K + 15, 4)), "0000")
    recCpj030W0.NUMPIE = Format$(Val(mId$(MsgTxt, K + 19, 7)), "0000000")
    recCpj030W0.NOLIGN = Format$(Val(mId$(MsgTxt, K + 26, 4)), "0000")
    recCpj030W0.Agence = Format$(Val(mId$(MsgTxt, K + 30, 3)), "000")
    recCpj030W0.Devise = Format$(Val(mId$(MsgTxt, K + 33, 4)), "0000")
    recCpj030W0.Compte = Format$(Val(mId$(MsgTxt, K + 37, 11)), "00000000000")
    recCpj030W0.SENECR = mId$(MsgTxt, K + 48, 1)
    recCpj030W0.MONDEV = CCur(Val(mId$(MsgTxt, K + 49, 15)) / 100)
    recCpj030W0.LIBELE = mId$(MsgTxt, K + 64, 50)
    recCpj030W0.BDFSTA = Format$(Val(mId$(MsgTxt, K + 114, 1)), "0")
    recCpj030W0.BDFNUM = Format$(Val(mId$(MsgTxt, K + 115, 3)), "000")
    recCpj030W0.EXONER = Format$(Val(mId$(MsgTxt, K + 118, 1)), "0")
    recCpj030W0.SIGENE = mId$(MsgTxt, K + 119, 1)
    
    recCpj030W0.AMJSAI = Format$(Val(mId$(MsgTxt, K + 124, 4)), "0000") _
                       & Format$(Val(mId$(MsgTxt, K + 122, 2)), "00") _
                       & Format$(Val(mId$(MsgTxt, K + 120, 2)), "00")
    recCpj030W0.AMJVAL = Format$(Val(mId$(MsgTxt, K + 132, 4)), "0000") _
                       & Format$(Val(mId$(MsgTxt, K + 130, 2)), "00") _
                       & Format$(Val(mId$(MsgTxt, K + 128, 2)), "00")
    recCpj030W0.AMJOPE = Format$(Val(mId$(MsgTxt, K + 140, 4)), "0000") _
                       & Format$(Val(mId$(MsgTxt, K + 138, 2)), "00") _
                       & Format$(Val(mId$(MsgTxt, K + 136, 2)), "00")
  
    recCpj030W0.INTERV = mId$(MsgTxt, K + 144, 1)
    recCpj030W0.MAJVAL = mId$(MsgTxt, K + 145, 1)
    recCpj030W0.OPOCHQ = Format$(Val(mId$(MsgTxt, K + 146, 1)), "0")
    recCpj030W0.JJCPLT = mId$(MsgTxt, K + 147, 1)
    recCpj030W0.CODFOR = mId$(MsgTxt, K + 148, 1)
    recCpj030W0.FOROPO = mId$(MsgTxt, K + 149, 1)
    recCpj030W0.FORVAL = mId$(MsgTxt, K + 150, 1)
    recCpj030W0.INICLI = Format$(Val(mId$(MsgTxt, K + 151, 1)), "0")
    recCpj030W0.EDAVIS = Format$(Val(mId$(MsgTxt, K + 152, 1)), "0")
    recCpj030W0.NOMOP = mId$(MsgTxt, K + 153, 10)
    recCpj030W0.COVAL = Format$(Val(mId$(MsgTxt, K + 163, 6)), "000000")
    recCpj030W0.REFDOS = mId$(MsgTxt, K + 169, 5)
    recCpj030W0.LIBEL1 = mId$(MsgTxt, K + 174, 32)
    recCpj030W0.LIBEL2 = mId$(MsgTxt, K + 206, 32)
    recCpj030W0.LIBEL3 = mId$(MsgTxt, K + 238, 32)
    recCpj030W0.LIBEL4 = mId$(MsgTxt, K + 270, 32)
    recCpj030W0.CENPRF = Format$(Val(mId$(MsgTxt, K + 302, 6)), "000000")
    recCpj030W0.CODPRO = mId$(MsgTxt, K + 308, 3)
    recCpj030W0.REFCON = mId$(MsgTxt, K + 311, 16)
    recCpj030W0.RACCPA = Format$(Val(mId$(MsgTxt, K + 327, 5)), "00000")
    recCpj030W0.CTLSTA = Format$(Val(mId$(MsgTxt, K + 332, 1)), "0")
    recCpj030W0.CTLNOM = mId$(MsgTxt, K + 333, 10)
    recCpj030W0.CTLAMJ = Format$(Val(mId$(MsgTxt, K + 343, 8)), "00000000")
    recCpj030W0.CTLHMS = Format$(Val(mId$(MsgTxt, K + 351, 6)), "000000")
    recCpj030W0.IMPAMJ = Format$(Val(mId$(MsgTxt, K + 357, 8)), "00000000")
    recCpj030W0.IMPHMS = Format$(Val(mId$(MsgTxt, K + 365, 6)), "000000")


Else
    srvCpj030W0_GetBuffer = recCpj030W0.Err
End If

MsgTxtIndex = MsgTxtIndex + recCpj030W0Len

End Function

'---------------------------------------------------------
Private Sub srvCpj030W0_PutBuffer(recCpj030W0 As typeCpj030W0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCpj030W0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCpj030W0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34


Mid$(MsgTxt, K + 1, 3) = Format$(Val(recCpj030W0.COSOC), "000")
Mid$(MsgTxt, K + 4, 1) = recCpj030W0.CODENR
Mid$(MsgTxt, K + 5, 3) = Format$(Val(recCpj030W0.AGEMET), "000")
Mid$(MsgTxt, K + 8, 3) = Format$(Val(recCpj030W0.SERVIC), "000")
Mid$(MsgTxt, K + 11, 4) = recCpj030W0.BIACOP
Mid$(MsgTxt, K + 15, 4) = Format$(Val(recCpj030W0.NUMLOT), "0000")
Mid$(MsgTxt, K + 19, 7) = Format$(Val(recCpj030W0.NUMPIE), "0000000")
Mid$(MsgTxt, K + 26, 4) = Format$(Val(recCpj030W0.NOLIGN), "0000")
Mid$(MsgTxt, K + 30, 3) = Format$(Val(recCpj030W0.Agence), "000")
Mid$(MsgTxt, K + 33, 4) = Format$(Val(recCpj030W0.Devise), "0000")
Mid$(MsgTxt, K + 37, 11) = Format$(Val(recCpj030W0.Compte), "00000000000")
Mid$(MsgTxt, K + 48, 1) = recCpj030W0.SENECR
Mid$(MsgTxt, K + 49, 15) = Format$(recCpj030W0.MONDEV * 100, "000000000000000")
Mid$(MsgTxt, K + 64, 50) = recCpj030W0.LIBELE
Mid$(MsgTxt, K + 114, 1) = Format$(Val(recCpj030W0.BDFSTA), "0")
Mid$(MsgTxt, K + 115, 3) = Format$(Val(recCpj030W0.BDFNUM), "000")
Mid$(MsgTxt, K + 118, 1) = Format$(Val(recCpj030W0.EXONER), "0")
Mid$(MsgTxt, K + 119, 1) = recCpj030W0.SIGENE
Mid$(MsgTxt, K + 120, 2) = Format$(mId$(Val(recCpj030W0.AMJSAI), 7, 2), "00")
Mid$(MsgTxt, K + 122, 2) = Format$(mId$(Val(recCpj030W0.AMJSAI), 5, 2), "00")
Mid$(MsgTxt, K + 124, 4) = Format$(mId$(Val(recCpj030W0.AMJSAI), 1, 4), "0000")
Mid$(MsgTxt, K + 128, 2) = Format$(mId$(Val(recCpj030W0.AMJVAL), 7, 2), "00")
Mid$(MsgTxt, K + 130, 2) = Format$(mId$(Val(recCpj030W0.AMJVAL), 5, 2), "00")
Mid$(MsgTxt, K + 132, 4) = Format$(mId$(Val(recCpj030W0.AMJVAL), 1, 4), "0000")
Mid$(MsgTxt, K + 136, 2) = Format$(mId$(Val(recCpj030W0.AMJOPE), 7, 2), "00")
Mid$(MsgTxt, K + 138, 2) = Format$(mId$(Val(recCpj030W0.AMJOPE), 5, 2), "00")
Mid$(MsgTxt, K + 140, 4) = Format$(mId$(Val(recCpj030W0.AMJOPE), 1, 4), "0000")
Mid$(MsgTxt, K + 144, 1) = recCpj030W0.INTERV
Mid$(MsgTxt, K + 145, 1) = recCpj030W0.MAJVAL
Mid$(MsgTxt, K + 146, 1) = Format$(Val(recCpj030W0.OPOCHQ), "0")
Mid$(MsgTxt, K + 147, 1) = recCpj030W0.JJCPLT
Mid$(MsgTxt, K + 148, 1) = recCpj030W0.CODFOR
Mid$(MsgTxt, K + 149, 1) = recCpj030W0.FOROPO
Mid$(MsgTxt, K + 150, 1) = recCpj030W0.FORVAL
Mid$(MsgTxt, K + 151, 1) = Format$(Val(recCpj030W0.INICLI), "0")
Mid$(MsgTxt, K + 152, 1) = Format$(Val(recCpj030W0.EDAVIS), "0")
Mid$(MsgTxt, K + 153, 10) = recCpj030W0.NOMOP
Mid$(MsgTxt, K + 163, 6) = Format$(Val(recCpj030W0.COVAL), "000000")
Mid$(MsgTxt, K + 169, 5) = recCpj030W0.REFDOS
Mid$(MsgTxt, K + 174, 32) = recCpj030W0.LIBEL1
Mid$(MsgTxt, K + 206, 32) = recCpj030W0.LIBEL2
Mid$(MsgTxt, K + 238, 32) = recCpj030W0.LIBEL3
Mid$(MsgTxt, K + 270, 32) = recCpj030W0.LIBEL4
Mid$(MsgTxt, K + 302, 6) = Format$(Val(recCpj030W0.CENPRF), "000000")
Mid$(MsgTxt, K + 308, 3) = recCpj030W0.CODPRO
Mid$(MsgTxt, K + 311, 16) = recCpj030W0.REFCON
Mid$(MsgTxt, K + 327, 5) = Format$(Val(recCpj030W0.RACCPA), "00000")
Mid$(MsgTxt, K + 332, 1) = Format$(Val(recCpj030W0.CTLSTA), "0")
Mid$(MsgTxt, K + 333, 10) = recCpj030W0.CTLNOM
Mid$(MsgTxt, K + 343, 8) = Format$(Val(recCpj030W0.CTLAMJ), "00000000")
Mid$(MsgTxt, K + 351, 6) = Format$(Val(recCpj030W0.CTLHMS), "000000")
Mid$(MsgTxt, K + 357, 8) = Format$(Val(recCpj030W0.IMPAMJ), "00000000")
Mid$(MsgTxt, K + 365, 6) = Format$(Val(recCpj030W0.IMPHMS), "000000")

MsgTxtLen = MsgTxtLen + recCpj030W0Len
End Sub



'---------------------------------------------------------
Private Function srvCpj030W0_Seek(recCpj030W0 As typeCpj030W0)
'---------------------------------------------------------

srvCpj030W0_Seek = "?"
MsgTxtLen = 0
Call srvCpj030W0_PutBuffer(recCpj030W0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCpj030W0_GetBuffer(recCpj030W0)) Then
        srvCpj030W0_Seek = Null
    Else
        Call srvCpj030W0_Error(recCpj030W0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCpj030W0_Snap(recCpj030W0 As typeCpj030W0)
'---------------------------------------------------------
srvCpj030W0_Snap = "?"
MsgTxtLen = 0
Call srvCpj030W0_PutBuffer(recCpj030W0)
Call srvCpj030W0_PutBuffer(arrCpj030W0(0))
If IsNull(SndRcv()) Then
    srvCpj030W0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCpj030W0_GetBuffer(recCpj030W0)) Then
            Call arrCpj030W0_AddItem(recCpj030W0)
            arrCpj030W0Suite = True
        Else
            arrCpj030W0Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCpj030W0_Init(recCpj030W0 As typeCpj030W0)
'---------------------------------------------------------
MsgTxt = Space$(recCpj030W0Len)
MsgTxtIndex = 0
Call srvCpj030W0_GetBuffer(recCpj030W0)
recCpj030W0.obj = "SRVCPJ030W"
End Sub

'---------------------------------------------------------
Public Sub arrCpj030W0_AddItem(recCpj030W0 As typeCpj030W0)
'---------------------------------------------------------
          
arrCpj030W0Nb = arrCpj030W0Nb + 1
    
If arrCpj030W0Nb > arrCpj030W0NbMax Then
    arrCpj030W0NbMax = arrCpj030W0NbMax + 10
    ReDim Preserve arrCpj030W0(arrCpj030W0NbMax)
End If
            
arrCpj030W0(arrCpj030W0Nb) = recCpj030W0
End Sub
