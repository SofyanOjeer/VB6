Attribute VB_Name = "srvDRH"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDRHLen = 209 ' 34 + 175

Type typeDRH
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Matricule               As String * 5
    Nature                  As String * 1
    Nom                     As String * 32
    Prénom                  As String * 24
    Civilité                As String * 1
    EntréeAmj               As String * 8
    SortieAmj               As String * 8
    EnfantNb                As Integer
    
    Compte                  As String * 11
    Service                 As String * 4
    Bureau                  As String * 3
    Téléphone1              As String * 3
    Téléphone2              As String * 3
    Téléphone3              As String * 3
    RéfInterne              As String * 16
    ElpCtlMvt               As String * 10
   
    Statut                  As String * 1
    UpdAmj                  As String * 8
    UpdHms                  As String * 6
   
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrDRH() As typeDRH
Public arrDRH_NB As Integer
Public arrDRH_NBMax As Integer
Public arrDRH_Index As Integer
Public arrDRH_Suite As Boolean
Public xDRH As typeDRH


'-----------------------------------------------------
Public Function srvDRH_Monitor(recDRH As typeDRH)
'-----------------------------------------------------

arrDRH_Suite = False
Select Case mId$(Trim(recDRH.Method), 1, 4)
    Case "Seek"
                srvDRH_Monitor = srvDRH_Seek(recDRH)
    Case "Snap"
              srvDRH_Monitor = srvDRH_Snap(recDRH)
    Case Else
                recDRH.Err = recDRH.Method
                Call srvDRH_Error(recDRH)
                srvDRH_Monitor = recDRH.Err
End Select

End Function

'-----------------------------------------------------
Sub srvDRH_Error(recDRH As typeDRH)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DRH" & Chr$(10) & Chr$(13)

Select Case mId$(recDRH.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDRH.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : DRHs.bas  ( " _
                & Trim(recDRH.obj) & " : " & Trim(recDRH.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvDRH_GetBuffer(recDRH As typeDRH)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDRH_GetBuffer = Null
recDRH.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recDRH.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recDRH.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDRH.Err = Space$(10) Then
    recDRH.Matricule = mId$(MsgTxt, K + 1, 5)
    recDRH.Nature = mId$(MsgTxt, K + 6, 1)
    recDRH.Nom = mId$(MsgTxt, K + 7, 32) ' : FR_ConvertEtoA recDRH.Nom
    recDRH.Prénom = mId$(MsgTxt, K + 39, 24) ' : FR_ConvertEtoA recDRH.Prénom
    recDRH.Civilité = mId$(MsgTxt, K + 63, 1)
    recDRH.EntréeAmj = mId$(MsgTxt, K + 64, 8)
    recDRH.SortieAmj = mId$(MsgTxt, K + 72, 8)
    recDRH.EnfantNb = CInt(Val(mId$(MsgTxt, K + 80, 3)))
    
    recDRH.Compte = mId$(MsgTxt, K + 83, 11)
    recDRH.Service = mId$(MsgTxt, K + 94, 4)
    recDRH.Bureau = mId$(MsgTxt, K + 98, 3)
    recDRH.Téléphone1 = mId$(MsgTxt, K + 101, 3)
    recDRH.Téléphone2 = mId$(MsgTxt, K + 104, 3)
    recDRH.Téléphone3 = mId$(MsgTxt, K + 107, 3)
    
    recDRH.RéfInterne = mId$(MsgTxt, K + 110, 16)
    recDRH.ElpCtlMvt = mId$(MsgTxt, K + 126, 10)
    
    recDRH.Statut = mId$(MsgTxt, K + 136, 1)
    recDRH.UpdAmj = mId$(MsgTxt, K + 137, 8)
    recDRH.UpdHms = mId$(MsgTxt, K + 145, 6)
    
    recDRH.ElpId = CLng(Val(mId$(MsgTxt, K + 151, 12)))
    recDRH.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 163, 3)))
    recDRH.ElpControl = mId$(MsgTxt, K + 166, 10)

Else
    srvDRH_GetBuffer = recDRH.Err
End If

MsgTxtIndex = MsgTxtIndex + recDRHLen

End Function

'---------------------------------------------------------
Private Sub srvDRH_PutBuffer(recDRH As typeDRH)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDRH.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDRH.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 5) = recDRH.Matricule
Mid$(MsgTxt, K + 6, 1) = recDRH.Nature
Mid$(MsgTxt, K + 7, 32) = recDRH.Nom
Mid$(MsgTxt, K + 39, 24) = recDRH.Prénom
Mid$(MsgTxt, K + 63, 1) = recDRH.Civilité
Mid$(MsgTxt, K + 64, 8) = Format$(recDRH.EntréeAmj, "00000000")
Mid$(MsgTxt, K + 72, 8) = Format$(recDRH.SortieAmj, "00000000")
Mid$(MsgTxt, K + 80, 3) = Format$(recDRH.EnfantNb, "000")

Mid$(MsgTxt, K + 83, 11) = Format$(recDRH.Compte, "00000000000")
Mid$(MsgTxt, K + 94, 4) = recDRH.Service
Mid$(MsgTxt, K + 98, 3) = recDRH.Bureau
Mid$(MsgTxt, K + 101, 3) = recDRH.Téléphone1
Mid$(MsgTxt, K + 104, 3) = recDRH.Téléphone2
Mid$(MsgTxt, K + 107, 3) = recDRH.Téléphone3

Mid$(MsgTxt, K + 110, 16) = recDRH.RéfInterne
Mid$(MsgTxt, K + 126, 10) = recDRH.ElpCtlMvt
Mid$(MsgTxt, K + 136, 1) = recDRH.Statut
Mid$(MsgTxt, K + 137, 8) = Format$(recDRH.UpdAmj, "00000000")
Mid$(MsgTxt, K + 145, 6) = Format$(recDRH.UpdHms, "000000")

Mid$(MsgTxt, K + 151, 12) = Format$(recDRH.ElpId, "000000000000")
Mid$(MsgTxt, K + 163, 3) = Format$(recDRH.ElpUpdate, "000")
Mid$(MsgTxt, K + 166, 10) = recDRH.ElpControl
    
MsgTxtLen = MsgTxtLen + recDRHLen
End Sub



'---------------------------------------------------------
Private Function srvDRH_Seek(recDRH As typeDRH)
'---------------------------------------------------------

srvDRH_Seek = "?"
MsgTxtLen = 0
Call srvDRH_PutBuffer(recDRH)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDRH_GetBuffer(recDRH)) Then
        srvDRH_Seek = Null
    Else
        Call srvDRH_Error(recDRH)
    End If
End If

End Function

'---------------------------------------------------------
Public Function srvDRH_SeekX(recDRH As typeDRH)
'---------------------------------------------------------

srvDRH_SeekX = "?"
MsgTxtLen = 0
Call srvDRH_PutBuffer(recDRH)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDRH_GetBuffer(recDRH)) Then
        srvDRH_SeekX = Null
 'x   Else
 'x       Call srvDRH_Error(recDRH)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDRH_Snap(recDRH As typeDRH)
'---------------------------------------------------------
srvDRH_Snap = "?"
MsgTxtLen = 0
Call srvDRH_PutBuffer(recDRH)
Call srvDRH_PutBuffer(arrDRH(0))
If IsNull(SndRcv()) Then
    srvDRH_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDRH_GetBuffer(recDRH)) Then
            Call arrDRH_AddItem(recDRH)
            arrDRH_Suite = True
        Else
            arrDRH_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'-----------------------------------------------------
Function srvDRH_Update(recDRH As typeDRH)
'-----------------------------------------------------

srvDRH_Update = "?"

MsgTxtLen = 0
Call srvDRH_PutBuffer(recDRH)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvDRH_GetBuffer(recDRH)) Then
        Call srvDRH_Error(recDRH)
        srvDRH_Update = recDRH.Err
        Exit Function
    Else
        srvDRH_Update = Null
    End If
Else
    recDRH.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recDRH_Init(recDRH As typeDRH)
'---------------------------------------------------------
MsgTxt = Space$(recDRHLen)
MsgTxtIndex = 0
Call srvDRH_GetBuffer(recDRH)
recDRH.obj = "SRVDRH    "
End Sub

'---------------------------------------------------------
Public Sub arrDRH_AddItem(recDRH As typeDRH)
'---------------------------------------------------------
          
arrDRH_NB = arrDRH_NB + 1
    
If arrDRH_NB > arrDRH_NBMax Then
    arrDRH_NBMax = arrDRH_NBMax + 10
    ReDim Preserve arrDRH(arrDRH_NBMax)
End If
            
arrDRH(arrDRH_NB) = recDRH
End Sub


'---------------------------------------------------------
Public Function arrDRH_Scan(recDRH As typeDRH) As Integer
'---------------------------------------------------------
arrDRH_Scan = 0
For arrDRH_Index = 1 To arrDRH_NB
    If recDRH.Matricule = arrDRH(arrDRH_Index).Matricule Then arrDRH_Scan = arrDRH_Index: Exit Function

Next arrDRH_Index
End Function


Public Function srvDRH_Identité(lDRH As typeDRH)
Dim X As String
Select Case lDRH.Civilité
    Case 1: X = "M.  "
    Case 2: X = "Mme "
    Case 3: X = "Mle "
    Case Else: X = "    "
End Select

srvDRH_Identité = X & Trim(lDRH.Nom) & " " & Trim(lDRH.Prénom)
End Function
