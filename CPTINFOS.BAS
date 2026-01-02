Attribute VB_Name = "srvCptInfo"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCptInfoLen = 845         ' 34 + 811
Public Const memoCptInfoLen = 811         ' 811

Type typeCptInfo
    obj        As String * 12
    Method     As String * 12
    Err        As String * 10
    Société     As String * 3
    Agence     As String * 3
    Devise     As String * 3
    Devisex    As String * 3
    Numéro      As String * 11
    NuméroAncien     As String * 11
    Intitulé     As String * 40
    Intitulé2   As String * 40
    TypeGA     As String * 1
    Situation     As String * 1
    Gestionnaire     As String * 2
    SoldeVeille      As Currency
    SoldeInstantané  As Currency
    MvtceJour    As String * 1
    DécouvertAutorisé    As String * 1
    DécouvertMontant   As Currency
    DécouvertAmj  As String * 8
    LibTyp     As String * 40
    Alpha   As String * 15
    BiaTyp     As String * 3
    BiaNum     As String * 2
   
    NatureTitulaire  As String * 2
    Actionnaire     As String * 1
    Résident     As String * 1
    Apporteur     As String * 2
    CompteGénéral     As String * 11
    Nature     As String * 1
    Sens    As String * 1
    CléRib As String * 2
    Nostro  As String * 1
    GroupeIdentification   As String * 6
    Conditions     As String * 3
    PrélèvementLibératoire  As String * 1
    Courrier As String * 1
    ServiceResponsable As String * 3
    ServiceAutorisé1 As String * 3
    ServiceAutorisé2 As String * 3
    ServiceAutorisé3 As String * 3
    ServiceAutorisé4 As String * 3
    ServiceAutorisé5 As String * 3

    DébitExercice  As Currency
    CréditExercice  As Currency
    DébitFindeMois  As Currency
    CréditFindeMois  As Currency
    SoldeFindeMois  As Currency
    DébitExerciceAntérieur As Currency
    CréditExerciceAntérieur  As Currency
    SoldeEnValeurVeille As Currency
    SoldeEnValeurPostérieur As Currency

   
    Echelle    As String * 1
    EchelleAmj  As String * 8
    EchelleSolde   As Currency
    Extrait    As String * 1
    ExtraitAmj  As String * 8
    ExtraitNuméro   As String * 3
    ExtraitSolde   As Currency
    AmjCréation  As String * 8
    AmjModification  As String * 8
    AmjAnnulation  As String * 8
    AmjRéactivation  As String * 8
    AmjDernierMouvement  As String * 8
    NomOpérateur   As String * 10
    Adresse1   As String * 40
    Adresse2   As String * 32
    Adresse3   As String * 32
    Adresse4   As String * 32
    Adresse5   As String * 32
    AdresseCP   As String * 5
    AdresseBD   As String * 27
    AdressePays As String * 32

End Type
    

Public arrCptInfo() As typeCptInfo
Public arrCptInfoNb As Integer
Public arrCptInfoNbMax As Integer
Public arrCptInfoIndex As Integer
Public arrCptInfosuite As Boolean
'-----------------------------------------------------
Public Function srvCptInfoFind(recCptInfo As typeCptInfo)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrCptInfoNb
    If recCptInfo.Société = arrCptInfo(I).Société _
    And recCptInfo.Agence = arrCptInfo(I).Agence _
    And recCptInfo.Devise = arrCptInfo(I).Devise _
    And recCptInfo.Numéro = arrCptInfo(I).Numéro Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrCptInfoNb + 1
    recCptInfo.Method = "JoinL1"
    arrCptInfo(0) = recCptInfo
    X = srvCptInfoSnap(recCptInfo)
    If I > arrCptInfoNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrCptInfoIndex = I
    recCptInfo = arrCptInfo(I)
End If

srvCptInfoFind = X

End Function



'-----------------------------------------------------
Public Function srvCptInfoMon(recCptInfo As typeCptInfo)
'-----------------------------------------------------

arrCptInfosuite = False
Select Case recCptInfo.Method
    Case "JoinL1      ", "SeekL1      ", "FindFirstL5 "
                srvCptInfoMon = srvCptInfoSeek(recCptInfo)
    Case "SnapL1      ", "SnapL1+     ", "SnapL5      ", "SnapL5+     ", _
         "SnapLZ      ", "SnapLZ+     ", "SnapLX      ", "SnapLX+     ", _
          "SnapLN      ", "SnapLN+     ", "SnapLA      ", "SnapLA+     ", _
          "SnapKG      ", "SnapKG+     ", "SnapPCI     ", "SnapPCI+    ", _
            "SnapKE      ", "SnapKE+     "
          srvCptInfoMon = srvCptInfoSnap(recCptInfo)
    Case Else
                recCptInfo.Err = recCptInfo.Method
                Call srvCptInfoError(recCptInfo)
                srvCptInfoMon = recCptInfo.Err
End Select

End Function

'-----------------------------------------------------
Sub srvCptInfoError(recCptInfo As typeCptInfo)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Compte" & Chr$(10) & Chr$(13)

Select Case mId$(recCptInfo.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCptInfo.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : Cpt.bas  ( " _
                & Trim(recCptInfo.obj) & " : " & Trim(recCptInfo.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCptInfoGetBuffer(recCptInfo As typeCptInfo)
'---------------------------------------------------------
Dim K As Integer
srvCptInfoGetBuffer = Null
recCptInfo.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCptInfo.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCptInfo.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCptInfo.Err = Space$(10) Then
    recCptInfo.Société = mId$(MsgTxt, K + 1, 3)
    recCptInfo.Agence = mId$(MsgTxt, K + 4, 3)
    recCptInfo.Devise = mId$(MsgTxt, K + 7, 3)
    recCptInfo.Devisex = mId$(MsgTxt, K + 10, 3)
    recCptInfo.Numéro = mId$(MsgTxt, K + 13, 11)
    recCptInfo.NuméroAncien = mId$(MsgTxt, K + 24, 11)
    recCptInfo.Intitulé = mId$(MsgTxt, K + 35, 40): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.Intitulé
    recCptInfo.Intitulé2 = mId$(MsgTxt, K + 75, 40): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.Intitulé2
    recCptInfo.TypeGA = mId$(MsgTxt, K + 115, 1)
    recCptInfo.Situation = mId$(MsgTxt, K + 116, 1)
    recCptInfo.Gestionnaire = mId$(MsgTxt, K + 117, 2)
    recCptInfo.SoldeVeille = CCur(Val(mId$(MsgTxt, K + 119, 19)))
    recCptInfo.SoldeInstantané = CCur(Val(mId$(MsgTxt, K + 138, 19)))
    recCptInfo.MvtceJour = mId$(MsgTxt, K + 157, 1)
    recCptInfo.DécouvertAutorisé = mId$(MsgTxt, K + 158, 1)
    recCptInfo.DécouvertMontant = CCur(Val(mId$(MsgTxt, K + 159, 19)))
    recCptInfo.DécouvertAmj = mId$(MsgTxt, K + 178, 8)
    recCptInfo.LibTyp = mId$(MsgTxt, K + 186, 40)
    recCptInfo.Alpha = mId$(MsgTxt, K + 226, 15)
    recCptInfo.BiaTyp = mId$(MsgTxt, K + 241, 3)
    recCptInfo.BiaNum = mId$(MsgTxt, K + 244, 2)

    recCptInfo.NatureTitulaire = mId$(MsgTxt, K + 249, 2)
    recCptInfo.Actionnaire = mId$(MsgTxt, K + 251, 1)
    recCptInfo.Résident = mId$(MsgTxt, K + 252, 1)
    recCptInfo.Apporteur = mId$(MsgTxt, K + 253, 2)
    recCptInfo.CompteGénéral = mId$(MsgTxt, K + 255, 11)
    recCptInfo.Nature = mId$(MsgTxt, K + 266, 1)
    recCptInfo.Sens = mId$(MsgTxt, K + 267, 1)
    recCptInfo.CléRib = mId$(MsgTxt, K + 268, 2)
    recCptInfo.Nostro = mId$(MsgTxt, K + 270, 1)
    recCptInfo.GroupeIdentification = mId$(MsgTxt, K + 271, 6)
    recCptInfo.Conditions = mId$(MsgTxt, K + 277, 3)
    recCptInfo.PrélèvementLibératoire = mId$(MsgTxt, K + 280, 1)
    recCptInfo.Courrier = mId$(MsgTxt, K + 281, 1)
    recCptInfo.ServiceResponsable = mId$(MsgTxt, K + 282, 3)
    recCptInfo.ServiceAutorisé1 = mId$(MsgTxt, K + 285, 3)
    recCptInfo.ServiceAutorisé2 = mId$(MsgTxt, K + 288, 3)
    recCptInfo.ServiceAutorisé3 = mId$(MsgTxt, K + 291, 3)
    recCptInfo.ServiceAutorisé4 = mId$(MsgTxt, K + 294, 3)
    recCptInfo.ServiceAutorisé5 = mId$(MsgTxt, K + 297, 3)
    recCptInfo.DébitExercice = CCur(Val(mId$(MsgTxt, K + 300, 19)))
    recCptInfo.CréditExercice = CCur(Val(mId$(MsgTxt, K + 319, 19)))
    recCptInfo.DébitFindeMois = CCur(Val(mId$(MsgTxt, K + 338, 19)))
    recCptInfo.CréditFindeMois = CCur(Val(mId$(MsgTxt, K + 357, 19)))
    recCptInfo.SoldeFindeMois = CCur(Val(mId$(MsgTxt, K + 376, 19)))
    recCptInfo.DébitExerciceAntérieur = CCur(Val(mId$(MsgTxt, K + 395, 19)))
    recCptInfo.CréditExerciceAntérieur = CCur(Val(mId$(MsgTxt, K + 414, 19)))
    recCptInfo.SoldeEnValeurVeille = CCur(Val(mId$(MsgTxt, K + 433, 19)))
    recCptInfo.SoldeEnValeurPostérieur = CCur(Val(mId$(MsgTxt, K + 452, 19)))
    recCptInfo.Echelle = mId$(MsgTxt, K + 471, 1)
    recCptInfo.EchelleAmj = mId$(MsgTxt, K + 472, 8)
    recCptInfo.EchelleSolde = CCur(Val(mId$(MsgTxt, K + 480, 19)))
    recCptInfo.Extrait = mId$(MsgTxt, K + 499, 1)
    recCptInfo.ExtraitAmj = mId$(MsgTxt, K + 500, 8)
    recCptInfo.ExtraitNuméro = mId$(MsgTxt, K + 508, 3)
    recCptInfo.ExtraitSolde = CCur(Val(mId$(MsgTxt, K + 511, 19)))
    recCptInfo.AmjCréation = mId$(MsgTxt, K + 530, 8)
    recCptInfo.AmjModification = mId$(MsgTxt, K + 538, 8)
    recCptInfo.AmjAnnulation = mId$(MsgTxt, K + 546, 8)
    recCptInfo.AmjRéactivation = mId$(MsgTxt, K + 554, 8)
    recCptInfo.AmjDernierMouvement = mId$(MsgTxt, K + 562, 8)
    recCptInfo.NomOpérateur = mId$(MsgTxt, K + 570, 10)
    recCptInfo.Adresse1 = mId$(MsgTxt, K + 580, 40): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.Adresse1
    recCptInfo.Adresse2 = mId$(MsgTxt, K + 620, 32): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.Adresse2
    recCptInfo.Adresse3 = mId$(MsgTxt, K + 652, 32): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.Adresse3
    recCptInfo.Adresse4 = mId$(MsgTxt, K + 684, 32): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.Adresse4
    recCptInfo.Adresse5 = mId$(MsgTxt, K + 716, 32): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.Adresse5
    recCptInfo.AdresseCP = mId$(MsgTxt, K + 748, 5)
    recCptInfo.AdresseBD = mId$(MsgTxt, K + 753, 27): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.AdresseBD
    recCptInfo.AdressePays = mId$(MsgTxt, K + 780, 32): 'jpl 2001.04.04 FR_ConvertEtoA_X recCptInfo.AdressePays
    If Not ctlGestionnaire_New(recCptInfo.Numéro, recCptInfo.Gestionnaire, recCptInfo.BiaTyp) Then
        recCptInfo.Intitulé2 = "! LE SOLDE DU COMPTE N'EST PAS AFFICHE "
        recCptInfo.SoldeVeille = 0
        recCptInfo.SoldeInstantané = 0
        recCptInfo.DécouvertMontant = 0
        recCptInfo.DébitExercice = 0
        recCptInfo.CréditExercice = 0
        recCptInfo.DébitFindeMois = 0
        recCptInfo.CréditFindeMois = 0
        recCptInfo.DébitExerciceAntérieur = 0
        recCptInfo.CréditExerciceAntérieur = 0
        recCptInfo.SoldeFindeMois = 0
        recCptInfo.SoldeEnValeurVeille = 0
        recCptInfo.SoldeEnValeurPostérieur = 0
        recCptInfo.EchelleSolde = 0
        recCptInfo.ExtraitSolde = 0
    End If

  
Else
    srvCptInfoGetBuffer = recCptInfo.Err
End If

MsgTxtIndex = MsgTxtIndex + recCptInfoLen

End Function

'---------------------------------------------------------
Private Sub srvCptInfoPutBuffer(recCptInfo As typeCptInfo)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCptInfo.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCptInfo.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = recCptInfo.Société
Mid$(MsgTxt, K + 4, 3) = recCptInfo.Agence
Mid$(MsgTxt, K + 7, 3) = recCptInfo.Devise
Mid$(MsgTxt, K + 10, 3) = recCptInfo.Devisex
Mid$(MsgTxt, K + 13, 11) = recCptInfo.Numéro
Mid$(MsgTxt, K + 24, 11) = recCptInfo.NuméroAncien
Mid$(MsgTxt, K + 35, 40) = recCptInfo.Intitulé
Mid$(MsgTxt, K + 75, 40) = recCptInfo.Intitulé2
Mid$(MsgTxt, K + 115, 1) = recCptInfo.TypeGA
Mid$(MsgTxt, K + 116, 1) = recCptInfo.Situation
Mid$(MsgTxt, K + 117, 2) = recCptInfo.Gestionnaire
Mid$(MsgTxt, K + 119, 19) = recCptInfo.SoldeVeille
Mid$(MsgTxt, K + 138, 19) = recCptInfo.SoldeVeille
Mid$(MsgTxt, K + 157, 1) = recCptInfo.MvtceJour
Mid$(MsgTxt, K + 158, 1) = recCptInfo.DécouvertAutorisé
Mid$(MsgTxt, K + 159, 19) = recCptInfo.DécouvertMontant
Mid$(MsgTxt, K + 178, 8) = recCptInfo.DécouvertAmj
Mid$(MsgTxt, K + 186, 40) = recCptInfo.LibTyp
Mid$(MsgTxt, K + 226, 15) = recCptInfo.Alpha
Mid$(MsgTxt, K + 241, 3) = recCptInfo.BiaTyp
Mid$(MsgTxt, K + 244, 2) = recCptInfo.BiaNum


    Mid$(MsgTxt, K + 249, 2) = recCptInfo.NatureTitulaire
    Mid$(MsgTxt, K + 251, 1) = recCptInfo.Actionnaire
    Mid$(MsgTxt, K + 252, 1) = recCptInfo.Résident
    Mid$(MsgTxt, K + 253, 2) = recCptInfo.Apporteur
    Mid$(MsgTxt, K + 255, 11) = recCptInfo.CompteGénéral
    Mid$(MsgTxt, K + 266, 1) = recCptInfo.Nature
    Mid$(MsgTxt, K + 267, 1) = recCptInfo.Sens
    Mid$(MsgTxt, K + 268, 2) = recCptInfo.CléRib
    Mid$(MsgTxt, K + 270, 1) = recCptInfo.Nostro
    Mid$(MsgTxt, K + 271, 6) = recCptInfo.GroupeIdentification
    Mid$(MsgTxt, K + 277, 3) = recCptInfo.Conditions
    Mid$(MsgTxt, K + 280, 1) = recCptInfo.PrélèvementLibératoire
    Mid$(MsgTxt, K + 281, 1) = recCptInfo.Courrier
    Mid$(MsgTxt, K + 282, 3) = recCptInfo.ServiceResponsable
    Mid$(MsgTxt, K + 285, 3) = recCptInfo.ServiceAutorisé1
    Mid$(MsgTxt, K + 288, 3) = recCptInfo.ServiceAutorisé2
    Mid$(MsgTxt, K + 291, 3) = recCptInfo.ServiceAutorisé3
    Mid$(MsgTxt, K + 294, 3) = recCptInfo.ServiceAutorisé4
    Mid$(MsgTxt, K + 297, 3) = recCptInfo.ServiceAutorisé5
    Mid$(MsgTxt, K + 300, 19) = recCptInfo.DébitExercice
    Mid$(MsgTxt, K + 319, 19) = recCptInfo.CréditExercice
    Mid$(MsgTxt, K + 338, 19) = recCptInfo.DébitFindeMois
    Mid$(MsgTxt, K + 357, 19) = recCptInfo.CréditFindeMois
    Mid$(MsgTxt, K + 376, 19) = recCptInfo.SoldeFindeMois
    Mid$(MsgTxt, K + 395, 19) = recCptInfo.DébitExerciceAntérieur
    Mid$(MsgTxt, K + 414, 19) = recCptInfo.CréditExerciceAntérieur
    Mid$(MsgTxt, K + 433, 19) = recCptInfo.SoldeEnValeurVeille
    Mid$(MsgTxt, K + 452, 19) = recCptInfo.SoldeEnValeurPostérieur
    Mid$(MsgTxt, K + 471, 1) = recCptInfo.Echelle
    Mid$(MsgTxt, K + 472, 8) = recCptInfo.EchelleAmj
    Mid$(MsgTxt, K + 480, 19) = recCptInfo.EchelleSolde
    Mid$(MsgTxt, K + 499, 1) = recCptInfo.Extrait
    Mid$(MsgTxt, K + 500, 8) = recCptInfo.ExtraitAmj
    Mid$(MsgTxt, K + 508, 3) = recCptInfo.ExtraitNuméro
    Mid$(MsgTxt, K + 511, 19) = recCptInfo.ExtraitSolde
    Mid$(MsgTxt, K + 530, 8) = recCptInfo.AmjCréation
    Mid$(MsgTxt, K + 538, 8) = recCptInfo.AmjModification
    Mid$(MsgTxt, K + 546, 8) = recCptInfo.AmjAnnulation
    Mid$(MsgTxt, K + 554, 8) = recCptInfo.AmjRéactivation
    Mid$(MsgTxt, K + 562, 8) = recCptInfo.AmjDernierMouvement
    Mid$(MsgTxt, K + 570, 10) = recCptInfo.NomOpérateur
    Mid$(MsgTxt, K + 580, 40) = recCptInfo.Adresse1
    Mid$(MsgTxt, K + 620, 32) = recCptInfo.Adresse2
    Mid$(MsgTxt, K + 652, 32) = recCptInfo.Adresse3
    Mid$(MsgTxt, K + 684, 32) = recCptInfo.Adresse4
    Mid$(MsgTxt, K + 716, 32) = recCptInfo.Adresse5
    Mid$(MsgTxt, K + 748, 5) = recCptInfo.AdresseCP
    Mid$(MsgTxt, K + 753, 27) = recCptInfo.AdresseBD
    Mid$(MsgTxt, K + 780, 32) = recCptInfo.AdressePays

MsgTxtLen = MsgTxtLen + recCptInfoLen
End Sub



'---------------------------------------------------------
Private Function srvCptInfoSeek(recCptInfo As typeCptInfo)
'---------------------------------------------------------

srvCptInfoSeek = "?"
MsgTxtLen = 0
Call srvCptInfoPutBuffer(recCptInfo)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCptInfoGetBuffer(recCptInfo)) Then
        srvCptInfoSeek = Null
    Else
        Call srvCptInfoError(recCptInfo)
    End If
End If

End Function


'---------------------------------------------------------
Private Function srvCptInfoSnap(recCptInfo As typeCptInfo)
'---------------------------------------------------------
srvCptInfoSnap = "?"
MsgTxtLen = 0
Call srvCptInfoPutBuffer(recCptInfo)
Call srvCptInfoPutBuffer(arrCptInfo(0))
If IsNull(SndRcv()) Then
    srvCptInfoSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCptInfoGetBuffer(recCptInfo)) Then
            Call arrCptInfoAddItem(recCptInfo)
            arrCptInfosuite = True
        Else
            arrCptInfosuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCptInfoInit(recCptInfo As typeCptInfo)
'---------------------------------------------------------
MsgTxt = Space$(recCptInfoLen)
MsgTxtIndex = 0
Call srvCptInfoGetBuffer(recCptInfo)
recCptInfo.obj = "SRVCOMPTE   "
recCptInfo.Société = SocId$
recCptInfo.Agence = SocAgence$

End Sub

'---------------------------------------------------------
Public Sub arrCptInfoAddItem(recCptInfo As typeCptInfo)
'---------------------------------------------------------
          
arrCptInfoNb = arrCptInfoNb + 1
    
If arrCptInfoNb > arrCptInfoNbMax Then
    arrCptInfoNbMax = arrCptInfoNbMax + 10
    ReDim Preserve arrCptInfo(arrCptInfoNbMax)
End If
            
arrCptInfo(arrCptInfoNb) = recCptInfo
End Sub

Public Function srvCptInfo_ServiceAutorisé(recCptInfo As typeCptInfo, xService As String) As Boolean
srvCptInfo_ServiceAutorisé = False
If IsNull(srvCptInfoFind(recCptInfo)) Then
    If recCptInfo.ServiceAutorisé1 = "000" And recCptInfo.ServiceAutorisé2 = "000" _
    And recCptInfo.ServiceAutorisé3 = "000" And recCptInfo.ServiceAutorisé4 = "000" _
    And recCptInfo.ServiceAutorisé5 = "000" Then
        srvCptInfo_ServiceAutorisé = True
    Else
        If recCptInfo.ServiceAutorisé1 = xService Or recCptInfo.ServiceAutorisé2 = xService _
        Or recCptInfo.ServiceAutorisé3 = xService Or recCptInfo.ServiceAutorisé4 = xService _
        Or recCptInfo.ServiceAutorisé5 = xService Then
            srvCptInfo_ServiceAutorisé = True
        End If
    End If
End If
End Function
