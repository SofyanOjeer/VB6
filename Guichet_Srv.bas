Attribute VB_Name = "srvGuichet"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recGuichetLen = 820 '34 + 786

Type typeGuichet
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Référence              As String * 10
    Séquence               As Integer
    CodeOpération          As String * 4
    Journal                As String * 6
    Société                As String * 3
    Agence                 As String * 3
    Devise                 As String * 3
    Compte                 As String * 11
    Montant                As Currency
    Sens                   As String * 1
    AmjOpération           As String * 8
    AmjValeur              As String * 8
    Libellé                As String * 50
    chkCompte              As String * 1
    chkSolde               As String * 1
    chkAmjOpération        As String * 1
    chkAmjValeur           As String * 1
    optAvis                As String * 1
    optVirement            As String * 1
    optSwift               As String * 1
    optAvisLangue          As String * 1

    CptMvtPièce           As Long
    CptMvtLigne           As Long
    CptMvtService         As String * 4
    CptMvtExonéré         As String * 1
    optCours              As String * 1
    CoursChange           As Double
    MontantAjustement     As Currency
    chkChèque             As String * 1
    NoChèque              As String * 10
    chkCoupureChange      As String * 1
    CoupureChange         As String * 88
  
    DeviseEspèces         As String * 3
    CptMvtPièceEspèces    As Long
    CptMvtLigneEspèces    As Long
    MontantEspèces        As Currency
    MontantRendu          As Currency
    CoursChangeEspèces    As Double
    chkCoupureEspèces     As String * 1
    CoupureEspèces        As String * 88
    
    Identité               As String * 50
    Complément1            As String * 50
    Complément2            As String * 50
    Complément3            As String * 50

    ContrepartieCompte     As String * 11
    ContrepartieLibellé    As String * 50
    MontantEuro            As Currency
    Conversion             As String * 1

    SaisieAmj               As String * 8
    SaisieHMS               As String * 6
    SaisieUsr               As String * 10
    ValidationAMJ           As String * 8
    ValidationHMS           As String * 6
    ValidationUsr           As String * 10
    ComptaAMJ               As String * 8
    ComptaHMS               As String * 6
    ComptaUsr               As String * 10
    UpdateSeq               As Integer

End Type
    
Public arrGuichet() As typeGuichet
Public arrGuichetNb As Integer
Public arrGuichetNbMax As Integer
Public arrGuichetIndex As Integer
Public arrGuichetSuite As Boolean


Public Sub recGuichet_CptInfo(recGuichet As typeGuichet, recCptInfo As typeCptInfo)
recCptInfoInit recCptInfo
recCptInfo.Method = "JoinL1      "
recCptInfo.Société = recGuichet.Société
recCptInfo.Agence = recGuichet.Agence
recCptInfo.Devise = Format$(recGuichet.Devise, "000")
recCptInfo.Numéro = recGuichet.Compte
recCptInfo.BiaTyp = "000"
recCptInfo.BiaNum = "00"
recCptInfo.NuméroAncien = "000000"
If Not IsNull(srvCptInfoFind(recCptInfo)) Then
    recCptInfo.Devise = "001"
    srvCptInfoFind recCptInfo
End If
End Sub


Public Sub recGuichet_mdbCptInfo(recGuichet As typeGuichet, recCptInfo As typeCptInfo)
recCptInfoInit recCptInfo
recCptInfo.Method = "JoinL1      "
recCptInfo.Société = recGuichet.Société
recCptInfo.Agence = recGuichet.Agence
recCptInfo.Devise = Format$(recGuichet.Devise, "000")
recCptInfo.Numéro = recGuichet.Compte
recCptInfo.BiaTyp = "000"
recCptInfo.BiaNum = "00"
recCptInfo.NuméroAncien = "000000"
If Not IsNull(mdbCptInfoP0_Find(recCptInfo)) Then
    recCptInfo.Devise = "001"
    srvCptInfoFind recCptInfo
End If
End Sub

'-----------------------------------------------------
Function srvGuichet_Update(recGuichet As typeGuichet)
'-----------------------------------------------------

srvGuichet_Update = "?"

MsgTxtLen = 0
Call srvGuichet_PutBuffer(recGuichet)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGuichet_GetBuffer(recGuichet)) Then
        Call srvGuichet_Error(recGuichet)
        srvGuichet_Update = recGuichet.Err
        Exit Function
    Else
        srvGuichet_Update = Null
    End If
Else
    recGuichet.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Function srvGuichet_UpdateKMax(recGuichet As typeGuichet, recGuichetKmax As typeGuichet)
'-----------------------------------------------------

srvGuichet_UpdateKMax = "?"

MsgTxtLen = 0
Call srvGuichet_PutBuffer(recGuichet)
Call srvGuichet_PutBuffer(recGuichetKmax)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGuichet_GetBuffer(recGuichet)) Then
        Call srvGuichet_Error(recGuichet)
        srvGuichet_UpdateKMax = recGuichet.Err
        Exit Function
    Else
        srvGuichet_UpdateKMax = Null
    End If
Else
    recGuichet.Err = "srv"
End If


'=====================================================
End Function


'-----------------------------------------------------
Public Function srvGuichet_Monitor(recGuichet As typeGuichet)
'-----------------------------------------------------

arrGuichetSuite = False
Select Case mId$(Trim(recGuichet.Method), 1, 4)
    Case "Seek", "Comp"
                srvGuichet_Monitor = srvGuichet_Seek(recGuichet)
    Case "Snap", "Prev"
              srvGuichet_Monitor = srvGuichet_Snap(recGuichet)
    Case Else
    
                recGuichet.Err = recGuichet.Method
                Call srvGuichet_Error(recGuichet)
                srvGuichet_Monitor = recGuichet.Err
End Select

End Function

'-----------------------------------------------------
Sub srvGuichet_Error(recGuichet As typeGuichet)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Opération de transfert (comptabilité): " ' & Chr$(10) & Chr$(13)

Select Case mId$(recGuichet.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recGuichet.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvGuichet_.bas  ( " _
                & Trim(recGuichet.obj) & " : " & Trim(recGuichet.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvGuichet_GetBuffer(recGuichet As typeGuichet)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvGuichet_GetBuffer = Null
recGuichet.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recGuichet.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recGuichet.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recGuichet.Err = Space$(10) Then
    recGuichet.Référence = mId$(MsgTxt, K + 1, 10)
    recGuichet.Séquence = CInt(Val(mId$(MsgTxt, K + 11, 3)))
    recGuichet.CodeOpération = mId$(MsgTxt, K + 14, 4)
    recGuichet.Journal = mId$(MsgTxt, K + 18, 6)
    recGuichet.Société = mId$(MsgTxt, K + 24, 3)
    recGuichet.Agence = mId$(MsgTxt, K + 27, 3)
    recGuichet.Devise = mId$(MsgTxt, K + 30, 3)
    recGuichet.Compte = mId$(MsgTxt, K + 33, 11)
    recGuichet.Montant = CCur(Val(mId$(MsgTxt, K + 44, 17)) / 100)
    recGuichet.Sens = mId$(MsgTxt, K + 61, 1)
    recGuichet.AmjOpération = mId$(MsgTxt, K + 62, 8)
    recGuichet.AmjValeur = mId$(MsgTxt, K + 70, 8)
    recGuichet.Libellé = mId$(MsgTxt, K + 78, 50)
    recGuichet.chkCompte = mId$(MsgTxt, K + 128, 1)
    recGuichet.chkSolde = mId$(MsgTxt, K + 129, 1)
    recGuichet.chkAmjOpération = mId$(MsgTxt, K + 130, 1)
    recGuichet.chkAmjValeur = mId$(MsgTxt, K + 131, 1)
    recGuichet.optAvis = mId$(MsgTxt, K + 132, 1)
    recGuichet.optVirement = mId$(MsgTxt, K + 133, 1)
    recGuichet.optSwift = mId$(MsgTxt, K + 134, 1)
    recGuichet.optAvisLangue = mId$(MsgTxt, K + 135, 1)
    
    recGuichet.CptMvtPièce = CInt(Val(mId$(MsgTxt, K + 136, 7)))
    recGuichet.CptMvtLigne = CInt(Val(mId$(MsgTxt, K + 143, 5)))
    recGuichet.CptMvtService = mId$(MsgTxt, K + 148, 4)
    recGuichet.CptMvtExonéré = mId$(MsgTxt, K + 152, 1)
    recGuichet.optCours = mId$(MsgTxt, K + 153, 1)
    recGuichet.CoursChange = CDbl(Val(mId$(MsgTxt, K + 154, 12)) / 10000000)
    recGuichet.MontantAjustement = CCur(Val(mId$(MsgTxt, K + 166, 16)) / 100)
    If mId$(MsgTxt, K + 182, 1) = "-" Then recGuichet.MontantAjustement = -recGuichet.MontantAjustement
    recGuichet.chkChèque = mId$(MsgTxt, K + 183, 1)
    recGuichet.NoChèque = mId$(MsgTxt, K + 184, 10)
    recGuichet.chkCoupureChange = mId$(MsgTxt, K + 194, 1)
    recGuichet.CoupureChange = mId$(MsgTxt, K + 195, 88)
    
    recGuichet.DeviseEspèces = mId$(MsgTxt, K + 283, 3)
    recGuichet.CptMvtPièceEspèces = CInt(Val(mId$(MsgTxt, K + 286, 7)))
    recGuichet.CptMvtLigneEspèces = CInt(Val(mId$(MsgTxt, K + 293, 5)))
    recGuichet.MontantEspèces = CCur(Val(mId$(MsgTxt, K + 298, 17)) / 100)
    recGuichet.MontantRendu = CCur(Val(mId$(MsgTxt, K + 315, 17)) / 100)
    recGuichet.CoursChangeEspèces = CDbl(Val(mId$(MsgTxt, K + 332, 12)) / 10000000)
    recGuichet.chkCoupureEspèces = mId$(MsgTxt, K + 344, 1)
    recGuichet.CoupureEspèces = mId$(MsgTxt, K + 345, 88)
    
    recGuichet.Identité = mId$(MsgTxt, K + 433, 50)
    recGuichet.Complément1 = mId$(MsgTxt, K + 483, 50)
    recGuichet.Complément2 = mId$(MsgTxt, K + 533, 50)
    recGuichet.Complément3 = mId$(MsgTxt, K + 583, 50)
    
    recGuichet.ContrepartieCompte = mId$(MsgTxt, K + 633, 11)
    recGuichet.ContrepartieLibellé = mId$(MsgTxt, K + 644, 50)
    recGuichet.MontantEuro = CCur(Val(mId$(MsgTxt, K + 694, 17)) / 100)
    recGuichet.Conversion = mId$(MsgTxt, K + 711, 1)
    
    recGuichet.SaisieAmj = mId$(MsgTxt, K + 712, 8)
    recGuichet.SaisieHMS = mId$(MsgTxt, K + 720, 6)
    recGuichet.SaisieUsr = mId$(MsgTxt, K + 726, 10)
    recGuichet.ValidationAMJ = mId$(MsgTxt, K + 736, 8)
    recGuichet.ValidationHMS = mId$(MsgTxt, K + 744, 6)
    recGuichet.ValidationUsr = mId$(MsgTxt, K + 750, 10)
    recGuichet.ComptaAMJ = mId$(MsgTxt, K + 760, 8)
    recGuichet.ComptaHMS = mId$(MsgTxt, K + 768, 6)
    recGuichet.ComptaUsr = mId$(MsgTxt, K + 774, 10)
    recGuichet.UpdateSeq = CInt(Val(mId$(MsgTxt, K + 784, 3)))
Else
    srvGuichet_GetBuffer = recGuichet.Err
End If

MsgTxtIndex = MsgTxtIndex + recGuichetLen

End Function

'---------------------------------------------------------
Public Sub srvGuichet_PutBuffer(recGuichet As typeGuichet)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recGuichet.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recGuichet.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 10) = recGuichet.Référence
Mid$(MsgTxt, K + 11, 3) = Format$(recGuichet.Séquence, "000")
Mid$(MsgTxt, K + 14, 4) = recGuichet.CodeOpération
Mid$(MsgTxt, K + 18, 6) = recGuichet.Journal
Mid$(MsgTxt, K + 24, 3) = recGuichet.Société
Mid$(MsgTxt, K + 27, 3) = recGuichet.Agence
Mid$(MsgTxt, K + 30, 3) = recGuichet.Devise
Mid$(MsgTxt, K + 33, 11) = recGuichet.Compte
Mid$(MsgTxt, K + 44, 17) = Format$(recGuichet.Montant * 100, "00000000000000000")
Mid$(MsgTxt, K + 61, 1) = recGuichet.Sens
Mid$(MsgTxt, K + 62, 8) = recGuichet.AmjOpération
Mid$(MsgTxt, K + 70, 8) = recGuichet.AmjValeur
Mid$(MsgTxt, K + 78, 50) = recGuichet.Libellé
Mid$(MsgTxt, K + 128, 1) = recGuichet.chkCompte
Mid$(MsgTxt, K + 129, 1) = recGuichet.chkSolde
Mid$(MsgTxt, K + 130, 1) = recGuichet.chkAmjOpération
Mid$(MsgTxt, K + 131, 1) = recGuichet.chkAmjValeur
Mid$(MsgTxt, K + 132, 1) = recGuichet.optAvis
Mid$(MsgTxt, K + 133, 1) = recGuichet.optVirement
Mid$(MsgTxt, K + 134, 1) = recGuichet.optSwift
Mid$(MsgTxt, K + 135, 1) = recGuichet.optAvisLangue

Mid$(MsgTxt, K + 136, 7) = Format$(recGuichet.CptMvtPièce, "0000000")
Mid$(MsgTxt, K + 143, 5) = Format$(recGuichet.CptMvtLigne, "00000")
Mid$(MsgTxt, K + 148, 4) = recGuichet.CptMvtService
Mid$(MsgTxt, K + 152, 1) = recGuichet.CptMvtExonéré
Mid$(MsgTxt, K + 153, 1) = recGuichet.optCours
Mid$(MsgTxt, K + 154, 12) = Format$(recGuichet.CoursChange * 10000000, "000000000000")
Mid$(MsgTxt, K + 166, 16) = Format$(Abs(recGuichet.MontantAjustement) * 100, "0000000000000000")
If recGuichet.MontantAjustement < 0 Then
    Mid$(MsgTxt, K + 182, 1) = "-"
Else
    Mid$(MsgTxt, K + 182, 1) = "+"
End If
Mid$(MsgTxt, K + 183, 1) = recGuichet.chkChèque
Mid$(MsgTxt, K + 184, 10) = recGuichet.NoChèque
Mid$(MsgTxt, K + 194, 1) = recGuichet.chkCoupureChange
Mid$(MsgTxt, K + 195, 88) = recGuichet.CoupureChange

Mid$(MsgTxt, K + 283, 3) = recGuichet.DeviseEspèces
Mid$(MsgTxt, K + 286, 7) = Format$(recGuichet.CptMvtPièceEspèces, "0000000")
Mid$(MsgTxt, K + 293, 5) = Format$(recGuichet.CptMvtLigneEspèces, "00000")
Mid$(MsgTxt, K + 298, 17) = Format$(recGuichet.MontantEspèces * 100, "00000000000000000")
Mid$(MsgTxt, K + 315, 17) = Format$(recGuichet.MontantRendu * 100, "00000000000000000")
Mid$(MsgTxt, K + 332, 12) = Format$(recGuichet.CoursChangeEspèces * 10000000, "000000000000")
Mid$(MsgTxt, K + 344, 1) = recGuichet.chkCoupureEspèces
Mid$(MsgTxt, K + 345, 88) = recGuichet.CoupureEspèces

Mid$(MsgTxt, K + 433, 50) = recGuichet.Identité
Mid$(MsgTxt, K + 483, 50) = recGuichet.Complément1
Mid$(MsgTxt, K + 533, 50) = recGuichet.Complément2
Mid$(MsgTxt, K + 583, 50) = recGuichet.Complément3
    
Mid$(MsgTxt, K + 633, 11) = recGuichet.ContrepartieCompte
Mid$(MsgTxt, K + 644, 50) = recGuichet.ContrepartieLibellé
Mid$(MsgTxt, K + 694, 17) = Format$(recGuichet.MontantEuro * 100, "00000000000000000")
Mid$(MsgTxt, K + 711, 1) = recGuichet.Conversion

Mid$(MsgTxt, K + 712, 8) = recGuichet.SaisieAmj
Mid$(MsgTxt, K + 720, 6) = recGuichet.SaisieHMS
Mid$(MsgTxt, K + 726, 10) = recGuichet.SaisieUsr
Mid$(MsgTxt, K + 736, 8) = recGuichet.ValidationAMJ
Mid$(MsgTxt, K + 744, 6) = recGuichet.ValidationHMS
Mid$(MsgTxt, K + 750, 10) = recGuichet.ValidationUsr
Mid$(MsgTxt, K + 760, 8) = recGuichet.ComptaAMJ
Mid$(MsgTxt, K + 768, 6) = recGuichet.ComptaHMS
Mid$(MsgTxt, K + 774, 10) = recGuichet.ComptaUsr
Mid$(MsgTxt, K + 784, 3) = Format$(recGuichet.UpdateSeq, "000")

MsgTxtLen = MsgTxtLen + recGuichetLen
End Sub



'---------------------------------------------------------
Private Function srvGuichet_Seek(recGuichet As typeGuichet)
'---------------------------------------------------------

srvGuichet_Seek = "?"
MsgTxtLen = 0
Call srvGuichet_PutBuffer(recGuichet)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvGuichet_GetBuffer(recGuichet)) Then
        srvGuichet_Seek = Null
    Else
 '       Call srvGuichet_Error(recGuichet)
    End If
End If

End Function

'---------------------------------------------------------
Function srvGuichet_Snd(recGuichet As typeGuichet)
'---------------------------------------------------------

srvGuichet_Snd = "?"
MsgTxtLen = 0
Call srvGuichet_PutBuffer(recGuichet)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    srvGuichet_Snd = Null
End If

End Function

'---------------------------------------------------------
Private Function srvGuichet_Snap(recGuichet As typeGuichet)
'---------------------------------------------------------
Dim I As Integer
srvGuichet_Snap = "?"
MsgTxtLen = 0
Call srvGuichet_PutBuffer(recGuichet)
Call srvGuichet_PutBuffer(arrGuichet(0))
If IsNull(SndRcv()) Then
    srvGuichet_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvGuichet_GetBuffer(recGuichet)) Then
            Call arrGuichet_AddItem(recGuichet)
            arrGuichetSuite = True
        Else
            arrGuichetSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recGuichet_Init(recGuichet As typeGuichet)
'---------------------------------------------------------
MsgTxt = String$(recGuichetLen, "0")
MsgTxtIndex = 0: Mid$(MsgTxt, MsgTxtIndex + 25, 10) = Space$(10)
Call srvGuichet_GetBuffer(recGuichet)
recGuichet.obj = "SRVGUICHET"
recGuichet.Libellé = "": recGuichet.ContrepartieLibellé = ""
recGuichet.Identité = ""
recGuichet.Complément1 = ""
recGuichet.Complément2 = ""
recGuichet.Complément3 = ""
recGuichet.CoupureEspèces = "": recGuichet.CoupureChange = ""
recGuichet.SaisieUsr = "": recGuichet.ValidationUsr = "": recGuichet.ComptaUsr = ""
recGuichet.optCours = " "
End Sub

'---------------------------------------------------------
Public Sub arrGuichet_AddItem(recGuichet As typeGuichet)
'---------------------------------------------------------
          
arrGuichetNb = arrGuichetNb + 1
    
If arrGuichetNb > arrGuichetNbMax Then
    arrGuichetNbMax = arrGuichetNbMax + 10
    ReDim Preserve arrGuichet(arrGuichetNbMax)
End If
            
arrGuichet(arrGuichetNb) = recGuichet
End Sub




