Attribute VB_Name = "srvCptMvt"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const memoCptMvtLen = 191
Public Const recCptMvtLen = 225 ' 34 + 191

Type typeCptMvt
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Société                 As String * 3
    Agence                  As String * 3
    Devise                  As String * 3           ' Num 3
    Compte                  As String * 11
    
    CodeOpération           As String * 4
    Service                 As String * 3
    Mt                      As Currency
    
    AmjSaisie               As String * 8
    AmjOpération            As String * 8
    AmjValeur               As String * 8
    Pièce                   As Long
    Ligne                   As Integer
    Lot                     As Integer
    Libellé                 As String * 50
    BdfK                    As String * 1
    BdfCode                 As String * 3
    Exonéré                 As String * 1
    Automatique             As String * 1
    EditionAvis             As String * 1
    CptComplémentaire       As String * 1
    SoldeVeille             As Currency
    AmjTraitement           As String * 8
    InitiéParClient         As String * 1
    ProgrammeSaisie         As String * 10
    OpérateurSaisie         As String * 10

End Type
    
Public arrCptMvt() As typeCptMvt
Public arrCptMvtNb As Integer
Public arrCptMvtNbMax As Integer
Public arrCptMvtIndex As Integer
Public arrCptMvtSuite As Boolean

'---------------------------------------------------------
Public Sub arrCptMvt_Load(recCompte As typeCompte, valAmjMin As String, valAmjMax As String)
'---------------------------------------------------------
Dim mMethod As String
Dim recCptMvt As typeCptMvt

ReDim arrCptMvt(0): arrCptMvtNbMax = 0
recCptMvtInit recCptMvt
arrCptMvtSuite = True: arrCptMvtNb = 0
recCptMvt.Société = recCompte.Société
recCptMvt.Agence = recCompte.Agence
recCptMvt.Devise = recCompte.Devise
recCptMvt.Compte = recCompte.Numéro
mMethod = "SnapJA+"
recCptMvt.Method = "SnapJA"
arrCptMvt(0) = recCptMvt

recCptMvt.AmjTraitement = valAmjMin
recCptMvt.Pièce = 0
recCptMvt.Ligne = 0
arrCptMvt(0).AmjTraitement = valAmjMax
arrCptMvt(0).Pièce = "999999999"
arrCptMvt(0).Ligne = "9999"

Do Until Not arrCptMvtSuite
    srvCptMvtMon recCptMvt
    recCptMvt = arrCptMvt(arrCptMvtNb)
    recCptMvt.Method = "SnapJA+"
Loop

End Sub


'-----------------------------------------------------
Public Function srvCptMvtMon(recCptMvt As typeCptMvt)
'-----------------------------------------------------
blnFR_Convert = False

arrCptMvtSuite = False
Select Case Trim(recCptMvt.Method)
    Case "SeekLA"
                srvCptMvtMon = srvCptMvtSeek(recCptMvt)
    Case "SnapJA", "SnapJA+", "PrevJA", "PrevJA+"
              srvCptMvtMon = srvCptMvtSnap(recCptMvt)
    Case "SnapLA", "SnapLA+", "PrevLA", "PrevLA+"
              srvCptMvtMon = srvCptMvtSnap(recCptMvt)
    Case Else
                recCptMvt.Err = recCptMvt.Method
                Call srvCptMvterror(recCptMvt)
                srvCptMvtMon = recCptMvt.Err
End Select

End Function

'-----------------------------------------------------
Sub srvCptMvterror(recCptMvt As typeCptMvt)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CptMvt" & Chr$(10) & Chr$(13)

Select Case mId$(recCptMvt.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCptMvt.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : CptMvts.bas  ( " _
                & Trim(recCptMvt.obj) & " : " & Trim(recCptMvt.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCptMvtGetBuffer(recCptMvt As typeCptMvt)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCptMvtGetBuffer = Null
recCptMvt.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCptMvt.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCptMvt.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCptMvt.Err = Space$(10) Then
    recCptMvt.Société = mId$(MsgTxt, K + 1, 3)
    recCptMvt.Agence = mId$(MsgTxt, K + 4, 3)
    recCptMvt.Devise = mId$(MsgTxt, K + 7, 3)
    recCptMvt.Compte = mId$(MsgTxt, K + 10, 11)
    recCptMvt.CodeOpération = mId$(MsgTxt, K + 21, 4)
    recCptMvt.Service = mId$(MsgTxt, K + 25, 3)
    recCptMvt.Mt = CCur(Val(mId$(MsgTxt, K + 28, 19)))
    recCptMvt.AmjSaisie = mId$(MsgTxt, K + 47, 8)
    recCptMvt.AmjOpération = mId$(MsgTxt, K + 55, 8)
    recCptMvt.AmjValeur = mId$(MsgTxt, K + 63, 8)
    recCptMvt.Pièce = Val(mId$(MsgTxt, K + 71, 7))
    recCptMvt.Ligne = Val(mId$(MsgTxt, K + 78, 4))
    recCptMvt.Lot = Val(mId$(MsgTxt, K + 82, 4))
    recCptMvt.Libellé = mId$(MsgTxt, K + 86, 50): '''FR_ConvertEtoA_X recCptMvt.Libellé
    recCptMvt.BdfK = mId$(MsgTxt, K + 136, 1)
    recCptMvt.BdfCode = mId$(MsgTxt, K + 137, 3)
    recCptMvt.Exonéré = mId$(MsgTxt, K + 140, 1)
    recCptMvt.Automatique = mId$(MsgTxt, K + 141, 1)
    recCptMvt.EditionAvis = mId$(MsgTxt, K + 142, 1)
    recCptMvt.CptComplémentaire = mId$(MsgTxt, K + 143, 1)
    recCptMvt.SoldeVeille = CCur(Val(mId$(MsgTxt, K + 144, 19)))
    recCptMvt.AmjTraitement = mId$(MsgTxt, K + 163, 8)
    recCptMvt.InitiéParClient = mId$(MsgTxt, K + 171, 1)
    recCptMvt.ProgrammeSaisie = mId$(MsgTxt, K + 172, 10)
    recCptMvt.OpérateurSaisie = mId$(MsgTxt, K + 182, 10)

Else
    srvCptMvtGetBuffer = recCptMvt.Err
End If

MsgTxtIndex = MsgTxtIndex + recCptMvtLen

End Function

'---------------------------------------------------------
Public Function srvCptMvt_GetX(recCptMvt As typeCptMvt, lMsg As String)
'---------------------------------------------------------
recCptMvt.Société = mId$(lMsg, 1, 3)
recCptMvt.Agence = mId$(lMsg, 4, 3)
recCptMvt.Devise = mId$(lMsg, 7, 3)
recCptMvt.Compte = mId$(lMsg, 10, 11)
recCptMvt.CodeOpération = mId$(lMsg, 21, 4)
recCptMvt.Service = mId$(lMsg, 25, 3)
recCptMvt.Mt = CCur(Val(mId$(lMsg, 28, 19)))
recCptMvt.AmjSaisie = mId$(lMsg, 47, 8)
recCptMvt.AmjOpération = mId$(lMsg, 55, 8)
recCptMvt.AmjValeur = mId$(lMsg, 63, 8)
recCptMvt.Pièce = Val(mId$(lMsg, 71, 7))
recCptMvt.Ligne = Val(mId$(lMsg, 78, 4))
recCptMvt.Lot = Val(mId$(lMsg, 82, 4))
recCptMvt.Libellé = mId$(lMsg, 86, 50)
recCptMvt.BdfK = mId$(lMsg, 136, 1)
recCptMvt.BdfCode = mId$(lMsg, 137, 3)
recCptMvt.Exonéré = mId$(lMsg, 140, 1)
recCptMvt.Automatique = mId$(lMsg, 141, 1)
recCptMvt.EditionAvis = mId$(lMsg, 142, 1)
recCptMvt.CptComplémentaire = mId$(lMsg, 143, 1)
recCptMvt.SoldeVeille = CCur(Val(mId$(lMsg, 144, 19)) / 100)
recCptMvt.AmjTraitement = mId$(lMsg, 163, 8)
recCptMvt.InitiéParClient = mId$(lMsg, 171, 1)
recCptMvt.ProgrammeSaisie = mId$(lMsg, 172, 10)
recCptMvt.OpérateurSaisie = mId$(lMsg, 182, 10)


End Function


'---------------------------------------------------------
Private Sub srvCptMvtPutBuffer(recCptMvt As typeCptMvt)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCptMvt.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCptMvt.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 3) = recCptMvt.Société
Mid$(MsgTxt, K + 4, 3) = recCptMvt.Agence
Mid$(MsgTxt, K + 7, 3) = recCptMvt.Devise
Mid$(MsgTxt, K + 10, 11) = Format$(recCptMvt.Compte, "00000000000")
Mid$(MsgTxt, K + 71, 7) = Format$(recCptMvt.Pièce, "0000000")
Mid$(MsgTxt, K + 78, 4) = Format$(recCptMvt.Ligne, "0000")
Mid$(MsgTxt, K + 82, 4) = Format$(recCptMvt.Lot, "0000")
Mid$(MsgTxt, K + 163, 8) = Format$(recCptMvt.AmjTraitement, "00000000")

MsgTxtLen = MsgTxtLen + recCptMvtLen
End Sub

'---------------------------------------------------------
Public Sub srvCptMvt_PutX(recCptMvt As typeCptMvt, lMsg As String)
'---------------------------------------------------------
Mid$(lMsg, 1, 3) = recCptMvt.Société
Mid$(lMsg, 4, 3) = recCptMvt.Agence
If IsNumeric(recCptMvt.Devise) Then
    Mid$(lMsg, 7, 3) = recCptMvt.Devise
Else
    Mid$(lMsg, 7, 3) = Format$(DevCode(recCptMvt.Devise), "000")
End If
Mid$(lMsg, 10, 11) = Format$(recCptMvt.Compte, "00000000000")
Mid$(lMsg, 21, 4) = recCptMvt.CodeOpération
Mid$(lMsg, 25, 3) = recCptMvt.Service
Mid$(lMsg, 29, 18) = Format$(Abs(recCptMvt.Mt) * 100, "000000000000000000")
If recCptMvt.Mt < 0 Then
    Mid$(lMsg, 28, 1) = "-"
Else
    Mid$(lMsg, 28, 1) = "+"
End If
Mid$(lMsg, 47, 8) = Format$(recCptMvt.AmjSaisie, "00000000")
Mid$(lMsg, 55, 8) = Format$(recCptMvt.AmjOpération, "00000000")
Mid$(lMsg, 63, 8) = Format$(recCptMvt.AmjValeur, "00000000")
Mid$(lMsg, 71, 7) = Format$(recCptMvt.Pièce, "0000000")
Mid$(lMsg, 78, 4) = Format$(recCptMvt.Ligne, "0000")
Mid$(lMsg, 82, 4) = Format$(recCptMvt.Lot, "0000")
Mid$(lMsg, 86, 50) = recCptMvt.Libellé
Mid$(lMsg, 136, 1) = recCptMvt.BdfK
Mid$(lMsg, 137, 3) = recCptMvt.BdfCode
Mid$(lMsg, 140, 1) = recCptMvt.Exonéré
Mid$(lMsg, 141, 1) = recCptMvt.Automatique
Mid$(lMsg, 142, 1) = recCptMvt.EditionAvis
Mid$(lMsg, 143, 1) = recCptMvt.CptComplémentaire
Mid$(lMsg, 145, 18) = Format$(Abs(recCptMvt.SoldeVeille) * 100, "000000000000000000")
If recCptMvt.Mt < 0 Then
    Mid$(lMsg, 144, 1) = "-"
Else
    Mid$(lMsg, 144, 1) = "+"
End If
Mid$(lMsg, 163, 8) = Format$(recCptMvt.AmjTraitement, "00000000")
Mid$(lMsg, 171, 1) = recCptMvt.InitiéParClient
Mid$(lMsg, 172, 10) = recCptMvt.ProgrammeSaisie
Mid$(lMsg, 182, 10) = recCptMvt.OpérateurSaisie


End Sub

'-----------------------------------------------------
Function srvCptMvt_UpdateBuffer(lMsg As String)
'-----------------------------------------------------

srvCptMvt_UpdateBuffer = "?"

MsgTxtLen = 0
''Call srvCptMvt_PutBuffer(recDRHMvt)
Mid$(MsgTxt, MsgTxtLen + 1, recCptMvtLen) = lMsg
MsgTxtLen = MsgTxtLen + recCptMvtLen

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If mId$(MsgTxt, MsgTxtIndex + 25, 10) = Space$(10) Then
        lMsg = Space$(recCptMvtLen)
        lMsg = mId$(MsgTxt, MsgTxtIndex + 1, recCptMvtLen)
        srvCptMvt_UpdateBuffer = Null
    Else
        srvCptMvt_UpdateBuffer = mId$(MsgTxt, MsgTxtIndex + 25, 10)
    End If
Else
    srvCptMvt_UpdateBuffer = "SndRcv"
End If


'=====================================================
End Function


'---------------------------------------------------------
Private Function srvCptMvtSeek(recCptMvt As typeCptMvt)
'---------------------------------------------------------

srvCptMvtSeek = "?"
MsgTxtLen = 0
Call srvCptMvtPutBuffer(recCptMvt)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCptMvtGetBuffer(recCptMvt)) Then
        srvCptMvtSeek = Null
    Else
        Call srvCptMvterror(recCptMvt)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCptMvtSnap(recCptMvt As typeCptMvt)
'---------------------------------------------------------
srvCptMvtSnap = "?"
MsgTxtLen = 0
Call srvCptMvtPutBuffer(recCptMvt)
Call srvCptMvtPutBuffer(arrCptMvt(0))
If IsNull(SndRcv()) Then
    srvCptMvtSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCptMvtGetBuffer(recCptMvt)) Then
            Call arrCptMvtAddItem(recCptMvt)
            arrCptMvtSuite = True
        Else
            arrCptMvtSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Function srvCptMvt_ElpBuffer(minCptMvt As typeCptMvt, maxCptMvt As typeCptMvt, recElpBuffer As typeElpBuffer)
'---------------------------------------------------------
Dim blnCptMvtSuite As Boolean, mMethod As String


blnCptMvtSuite = True
srvCptMvt_ElpBuffer = Null
mMethod = Trim(minCptMvt.Method) & "+"
recElpBuffer_Init_id recElpBuffer
recElpBuffer.Method = "AddNew"

MsgTxtLen = 0
Call srvCptMvtPutBuffer(minCptMvt)

Do Until Not blnCptMvtSuite
    blnCptMvtSuite = False
    Call srvCptMvtPutBuffer(maxCptMvt)
    If Not IsNull(SndRcv()) Then srvCptMvt_ElpBuffer = "?": Exit Function
    
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If mId$(MsgTxt, MsgTxtIndex + 25, 10) = Space$(10) Then

                recElpBuffer.Seq = recElpBuffer.Seq + 1
                recElpBuffer.Data = mId$(MsgTxt, MsgTxtIndex + 1, recCptMvtLen)
                MsgTxtIndex = MsgTxtIndex + recCptMvtLen
                tableElpBuffer_Update recElpBuffer
                blnCptMvtSuite = True
            Else
                blnCptMvtSuite = False
                Exit Function
            End If
        Loop
        Mid$(MsgTxt, 1, recCptMvtLen) = recElpBuffer.Data
        Mid$(MsgTxt, 13, 12) = mMethod
        MsgTxtLen = recCptMvtLen
'    End If
Loop

End Function


'---------------------------------------------------------
Public Sub recCptMvtInit(recCptMvt As typeCptMvt)
'---------------------------------------------------------
recCptMvt.obj = "SRVCPTMVT"
recCptMvt.Method = ""
recCptMvt.Err = ""

recCptMvt.Société = SocId$
recCptMvt.Agence = SocAgence$
recCptMvt.Devise = ""
recCptMvt.Compte = "00000000000"
recCptMvt.CodeOpération = ""
recCptMvt.Service = ""
recCptMvt.Mt = 0
recCptMvt.AmjSaisie = "00000000"
recCptMvt.AmjOpération = "00000000"
recCptMvt.AmjValeur = "00000000"
recCptMvt.Pièce = 0
recCptMvt.Ligne = 0
recCptMvt.Lot = 0
recCptMvt.Libellé = ""
recCptMvt.BdfK = "0"
recCptMvt.BdfCode = "000"
recCptMvt.Exonéré = "0"
recCptMvt.Automatique = "*"
recCptMvt.EditionAvis = "0"
recCptMvt.CptComplémentaire = "0"
recCptMvt.SoldeVeille = 0
recCptMvt.AmjTraitement = "00000000"
recCptMvt.InitiéParClient = "0"
recCptMvt.ProgrammeSaisie = ""
recCptMvt.OpérateurSaisie = usrId
End Sub

'---------------------------------------------------------
Public Sub arrCptMvtAddItem(recCptMvt As typeCptMvt)
'---------------------------------------------------------
          
arrCptMvtNb = arrCptMvtNb + 1
    
If arrCptMvtNb > arrCptMvtNbMax Then
    arrCptMvtNbMax = arrCptMvtNbMax + 10
    ReDim Preserve arrCptMvt(arrCptMvtNbMax)
End If
            
arrCptMvt(arrCptMvtNb) = recCptMvt
End Sub
