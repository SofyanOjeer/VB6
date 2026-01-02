Attribute VB_Name = "srvDRHMvt"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDRHMvtLen = 130 ' 34 + 96
Public Const memoDRHMvtLen = 96
Type typeDRHMvt
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Matricule               As String * 5
    IdSeq                   As Long
    MvtCode                 As String * 4
    DébutAmj                As String * 8
    DébutAmjK               As String * 1
    RepriseAmj              As String * 8
    RepriseAmjK             As String * 1
    RepriseChk              As String * 1
    Nbj                     As Double
    NbjChk                  As String * 1
    MvtSens                 As String * 1
    MvtCO                   As String * 1
   
    RéfInterne              As String * 12
    NbjOuvré                As Double
    Statut                  As String * 1
    UpdAmj                  As String * 8
    UpdHms                  As String * 6
   
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrDRHMvt() As typeDRHMvt
Public arrDRHMvt_NB As Integer
Public arrDRHMvt_NBMax As Integer
Public arrDRHMvt_Index As Integer
Public arrDRHMvt_Suite As Boolean

Public paramDRHMvt As typeDRHMvt
Public paramDRHMvt_TotalK As Integer

Public arrDRHMvt_Absences_Nb(99) As Double, arrDRHMvt_Droits_Nb(99) As Double, arrDRHMvt_Libellé(99) As String
Public mDRHCalendrier As typeElpTable
Public paramTR_Filename  As String, paramTR_Disquette As String
Public paramTR_Nominal As String * 4, paramTR_PartPatronale As String * 4, paramTR_Id As String
Public Sub arrTotal_Init()
Dim I As Integer
For I = 0 To 99
    arrDRHMvt_Absences_Nb(I) = 0: arrDRHMvt_Droits_Nb(I) = 0
Next I

End Sub

Public Sub arrTotal_Add(lDRHMvt As typeDRHMvt)
Call frmDRH.paramDRHMvt_Init(lDRHMvt.MvtCode)

Select Case lDRHMvt.MvtSens
    Case "-", "P": arrDRHMvt_Absences_Nb(paramDRHMvt_TotalK) = arrDRHMvt_Absences_Nb(paramDRHMvt_TotalK) + lDRHMvt.Nbj
    Case "C": arrDRHMvt_Droits_Nb(paramDRHMvt_TotalK) = arrDRHMvt_Droits_Nb(paramDRHMvt_TotalK) + lDRHMvt.Nbj
    Case "D": arrDRHMvt_Droits_Nb(paramDRHMvt_TotalK) = arrDRHMvt_Droits_Nb(paramDRHMvt_TotalK) - lDRHMvt.Nbj
End Select

End Sub

Public Function Param_DRHCalendrier(lK2 As String, lElpTable As typeElpTable)
Param_DRHCalendrier = Null
If Trim(lElpTable.K2) <> Trim(lK2) Then
    lElpTable.Method = "Seek="
    lElpTable.Id = "DRH"
    lElpTable.K1 = "$Calendrier"
    lElpTable.K2 = ""
    lElpTable.K2 = lK2
    lElpTable.Err = tableElpTable_Read(lElpTable)
    If lElpTable.Err <> 0 Then
        lElpTable.Memo = Space$(62)
        Call MsgBox("srvDRHMvt : Param_DRHCalendrier", vbCritical, "date inconnue : " & lK2)
        Param_DRHCalendrier = "?"
    End If
End If

End Function

Public Function Param_DRHCalendrier_Update(lK2 As String, lElpTable As typeElpTable)
Dim I As Integer, J As Integer, J2 As Integer, K As Integer, M As Integer, wAmj As String * 8
Dim vAmj As Variant
Dim xElpTable As typeElpTable

Param_DRHCalendrier_Update = Null

lElpTable.Id = "DRH"
lElpTable.K1 = "$Calendrier"


For M = 1 To 12
    wAmj = mId$(lK2, 1, 4) & Format$(M, "00") & "01"
    lElpTable.K2 = mId$(wAmj, 1, 6)
    lElpTable.Method = "Seek="
    K = tableElpTable_Read(lElpTable)
    If K = 0 Then
        lElpTable.Method = "Update"
    Else
        lElpTable.Method = "AddNew"
    End If
    
    lElpTable.Name = lElpTable.K2
    lElpTable.Memo = Space$(62)
    wAmj = dateFinDeMois(wAmj)
    J2 = mId$(wAmj, 7, 2)
    
    xElpTable.Id = "DRH"
    xElpTable.K1 = "Férié"
    xElpTable.Method = "Seek="
  
    For J = 1 To J2
        Mid$(wAmj, 7, 2) = Format$(J, "00")
        vAmj = dateImp(wAmj)
        K = Weekday(vAmj)
        If K = 1 Or K = 7 Then
            Mid$(lElpTable.Memo, J * 2 - 1, 2) = "XX"
        Else
            xElpTable.K2 = wAmj
            K = tableElpTable_Read(xElpTable)
            If K = 0 Then
                Mid$(lElpTable.Memo, J * 2 - 1, 2) = mId$(xElpTable.Memo, 1, 2)
            Else
                Mid$(lElpTable.Memo, J * 2 - 1, 2) = "00"
            End If

            
        End If
    Next J
    xElpTable = lElpTable
    K = tableElpTable_Read(xElpTable)
    K = tableElpTable_Update(lElpTable)
    If K <> 0 Then MsgBox "Erreur update : " & lElpTable.K2, vbCritical, "Param_DRHCalendrier_Update"

Next M

End Function


'-----------------------------------------------------
Public Function srvDRHMvt_Monitor(recDRHMvt As typeDRHMvt)
'-----------------------------------------------------

arrDRHMvt_Suite = False
Select Case mId$(Trim(recDRHMvt.Method), 1, 4)
    Case "Seek"
                srvDRHMvt_Monitor = srvDRHMvt_Seek(recDRHMvt)
    Case "Snap"
              srvDRHMvt_Monitor = srvDRHMvt_Snap(recDRHMvt)
    Case Else
                recDRHMvt.Err = recDRHMvt.Method
                Call srvDRHMvt_Error(recDRHMvt)
                srvDRHMvt_Monitor = recDRHMvt.Err
End Select

End Function

'-----------------------------------------------------
Sub srvDRHMvt_Error(recDRHMvt As typeDRHMvt)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DRHMvt" & Chr$(10) & Chr$(13)

Select Case mId$(recDRHMvt.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDRHMvt.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : DRHMvts.bas  ( " _
                & Trim(recDRHMvt.obj) & " : " & Trim(recDRHMvt.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvDRHMvt_GetBuffer(recDRHMvt As typeDRHMvt)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDRHMvt_GetBuffer = Null
recDRHMvt.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recDRHMvt.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recDRHMvt.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDRHMvt.Err = Space$(10) Then
    recDRHMvt.Matricule = mId$(MsgTxt, K + 1, 5)
    recDRHMvt.IdSeq = CLng(Val(mId$(MsgTxt, K + 6, 5)))
    recDRHMvt.MvtCode = mId$(MsgTxt, K + 11, 4)
    recDRHMvt.DébutAmj = mId$(MsgTxt, K + 15, 8)
    recDRHMvt.DébutAmjK = mId$(MsgTxt, K + 23, 1)
    recDRHMvt.RepriseAmj = mId$(MsgTxt, K + 24, 8)
    recDRHMvt.RepriseAmjK = mId$(MsgTxt, K + 32, 1)
    recDRHMvt.RepriseChk = mId$(MsgTxt, K + 33, 1)
    recDRHMvt.Nbj = CDbl(Val(mId$(MsgTxt, K + 34, 4)) / 10)
    recDRHMvt.NbjChk = mId$(MsgTxt, K + 38, 1)
    recDRHMvt.MvtSens = mId$(MsgTxt, K + 39, 1)
    recDRHMvt.MvtCO = mId$(MsgTxt, K + 40, 1)
    
    recDRHMvt.RéfInterne = mId$(MsgTxt, K + 41, 12)
    recDRHMvt.NbjOuvré = CDbl(Val(mId$(MsgTxt, K + 53, 4)) / 10)
    recDRHMvt.Statut = mId$(MsgTxt, K + 57, 1)
    recDRHMvt.UpdAmj = mId$(MsgTxt, K + 58, 8)
    recDRHMvt.UpdHms = mId$(MsgTxt, K + 66, 6)
    recDRHMvt.ElpId = CLng(Val(mId$(MsgTxt, K + 72, 12)))
    recDRHMvt.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 84, 3)))
    recDRHMvt.ElpControl = mId$(MsgTxt, K + 87, 10)

Else
    srvDRHMvt_GetBuffer = recDRHMvt.Err
End If

MsgTxtIndex = MsgTxtIndex + recDRHMvtLen

End Function

'---------------------------------------------------------
Private Sub srvDRHMvt_PutBuffer(recDRHMvt As typeDRHMvt)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDRHMvt.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDRHMvt.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 5) = recDRHMvt.Matricule
Mid$(MsgTxt, K + 6, 5) = Format$(recDRHMvt.IdSeq, "00000")
Mid$(MsgTxt, K + 11, 4) = recDRHMvt.MvtCode
Mid$(MsgTxt, K + 15, 8) = Format$(recDRHMvt.DébutAmj, "00000000")
Mid$(MsgTxt, K + 23, 1) = recDRHMvt.DébutAmjK
Mid$(MsgTxt, K + 24, 8) = Format$(recDRHMvt.RepriseAmj, "00000000")
Mid$(MsgTxt, K + 32, 1) = recDRHMvt.RepriseAmjK
Mid$(MsgTxt, K + 33, 1) = recDRHMvt.RepriseChk
Mid$(MsgTxt, K + 34, 4) = Format$(recDRHMvt.Nbj * 10, "0000")
Mid$(MsgTxt, K + 38, 1) = recDRHMvt.NbjChk
Mid$(MsgTxt, K + 39, 1) = recDRHMvt.MvtSens
Mid$(MsgTxt, K + 40, 1) = recDRHMvt.MvtCO

Mid$(MsgTxt, K + 41, 12) = recDRHMvt.RéfInterne
Mid$(MsgTxt, K + 53, 4) = Format$(recDRHMvt.NbjOuvré * 10, "0000")
Mid$(MsgTxt, K + 57, 1) = recDRHMvt.Statut
Mid$(MsgTxt, K + 58, 8) = Format$(recDRHMvt.UpdAmj, "00000000")
Mid$(MsgTxt, K + 66, 6) = Format$(recDRHMvt.UpdHms, "000000")
Mid$(MsgTxt, K + 72, 12) = Format$(recDRHMvt.ElpId, "000000000000")
Mid$(MsgTxt, K + 84, 3) = Format$(recDRHMvt.ElpUpdate, "000")
Mid$(MsgTxt, K + 87, 10) = recDRHMvt.ElpControl

MsgTxtLen = MsgTxtLen + recDRHMvtLen
End Sub



'---------------------------------------------------------
Private Function srvDRHMvt_Seek(recDRHMvt As typeDRHMvt)
'---------------------------------------------------------

srvDRHMvt_Seek = "?"
MsgTxtLen = 0
Call srvDRHMvt_PutBuffer(recDRHMvt)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDRHMvt_GetBuffer(recDRHMvt)) Then
        srvDRHMvt_Seek = Null
    Else
'        Call srvDRHMvt_Error(recDRHMvt)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDRHMvt_Snap(recDRHMvt As typeDRHMvt)
'---------------------------------------------------------
srvDRHMvt_Snap = "?"
MsgTxtLen = 0
Call srvDRHMvt_PutBuffer(recDRHMvt)
Call srvDRHMvt_PutBuffer(arrDRHMvt(0))
If IsNull(SndRcv()) Then
    srvDRHMvt_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDRHMvt_GetBuffer(recDRHMvt)) Then
            Call arrDRHMvt_AddItem(recDRHMvt)
            arrDRHMvt_Suite = True
        Else
            arrDRHMvt_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'-----------------------------------------------------
Function srvDRHMvt_Update(recDRHMvt As typeDRHMvt)
'-----------------------------------------------------

srvDRHMvt_Update = "?"

MsgTxtLen = 0
Call srvDRHMvt_PutBuffer(recDRHMvt)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvDRHMvt_GetBuffer(recDRHMvt)) Then
        Call srvDRHMvt_Error(recDRHMvt)
        srvDRHMvt_Update = recDRHMvt.Err
        Exit Function
    Else
        srvDRHMvt_Update = Null
    End If
Else
    recDRHMvt.Err = "srv"
End If


'=====================================================
End Function

'---------------------------------------------------------
Public Function srvDRHMvt_ElpBuffer(minDRHMvt As typeDRHMvt, maxDRHMvt As typeDRHMvt, recElpBuffer As typeElpBuffer)
'---------------------------------------------------------
Dim blnDRHMvtSuite As Boolean, mMethod As String

blnDRHMvtSuite = True
srvDRHMvt_ElpBuffer = Null
mMethod = Trim(minDRHMvt.Method) & "+"
recElpBuffer_Init_id recElpBuffer
recElpBuffer.Method = "AddNew"

MsgTxtLen = 0
Call srvDRHMvt_PutBuffer(minDRHMvt)

Do Until Not blnDRHMvtSuite
    blnDRHMvtSuite = False
    Call srvDRHMvt_PutBuffer(maxDRHMvt)
    If Not IsNull(SndRcv()) Then srvDRHMvt_ElpBuffer = "?": Exit Function
    
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If mId$(MsgTxt, MsgTxtIndex + 25, 10) = Space$(10) Then

                recElpBuffer.Seq = recElpBuffer.Seq + 1
                recElpBuffer.Data = mId$(MsgTxt, MsgTxtIndex + 1, recDRHMvtLen)
                MsgTxtIndex = MsgTxtIndex + recDRHMvtLen
                tableElpBuffer_Update recElpBuffer
                blnDRHMvtSuite = True
            Else
                blnDRHMvtSuite = False
                Exit Function
            End If
        Loop
        Mid$(MsgTxt, 1, recDRHMvtLen) = recElpBuffer.Data
        Mid$(MsgTxt, 13, 12) = mMethod
        MsgTxtLen = recDRHMvtLen
'    End If
Loop

End Function




'---------------------------------------------------------
Public Sub recDRHMvt_Init(recDRHMvt As typeDRHMvt)
'---------------------------------------------------------
MsgTxt = Space$(recDRHMvtLen)
MsgTxtIndex = 0
Call srvDRHMvt_GetBuffer(recDRHMvt)
recDRHMvt.DébutAmjK = "0"
recDRHMvt.RepriseAmjK = "0"
recDRHMvt.obj = "SRVDRHMVT"
End Sub

'---------------------------------------------------------
Public Sub arrDRHMvt_AddItem(recDRHMvt As typeDRHMvt)
'---------------------------------------------------------
          
arrDRHMvt_NB = arrDRHMvt_NB + 1
    
If arrDRHMvt_NB > arrDRHMvt_NBMax Then
    arrDRHMvt_NBMax = arrDRHMvt_NBMax + 10
    ReDim Preserve arrDRHMvt(arrDRHMvt_NBMax)
End If
            
arrDRHMvt(arrDRHMvt_NB) = recDRHMvt
End Sub




Public Function srvDRHMvt_RepriseAmj(lDRHMvt As typeDRHMvt, lCalendrier As typeElpTable)
Dim blkOK As Boolean, wNbj As Double, K As Integer, J As Integer, X1 As String * 1
Dim V As Variant, wAmj As String * 8, blnOk As Boolean

blnOk = False
srvDRHMvt_RepriseAmj = Null
wAmj = mId$(lDRHMvt.DébutAmj, 1, 6) & "01"
V = Param_DRHCalendrier(mId$(wAmj, 1, 6), lCalendrier)
If Not IsNull(V) Then srvDRHMvt_RepriseAmj = V: Exit Function
wNbj = lDRHMvt.Nbj
lDRHMvt.NbjOuvré = 0
K = CInt(mId$(lDRHMvt.DébutAmj, 7, 2)) * 2
If lDRHMvt.DébutAmjK = "0" Then K = K - 1

Do
    X1 = mId$(lCalendrier.Memo, K, 1)
    If wNbj <= 0 And X1 = "0" Then
        blnOk = True
    Else
        Select Case X1
            Case "0": wNbj = wNbj - 0.5
                      lDRHMvt.NbjOuvré = lDRHMvt.NbjOuvré + 0.5
            Case "X": If lDRHMvt.MvtCO = "C" Then wNbj = wNbj - 0.5
        End Select
        
        If K < 62 Then
            K = K + 1
        Else
            wAmj = dateElp("MoisAdd", 1, wAmj)
            V = Param_DRHCalendrier(mId$(wAmj, 1, 6), lCalendrier)
            If Not IsNull(V) Then srvDRHMvt_RepriseAmj = V: Exit Function
            K = 1
        End If
        
        
    End If

Loop Until blnOk
If (K Mod 2) = 0 Then
    J = K / 2
    lDRHMvt.RepriseAmjK = "1"
Else
    J = (K + 1) / 2
   lDRHMvt.RepriseAmjK = "0"
End If

lDRHMvt.RepriseAmj = mId$(lCalendrier.K2, 1, 6) & Format$(J, "00")
End Function
