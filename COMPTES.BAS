Attribute VB_Name = "srvCompte"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCompteLen = 279          ' 34 + 245
Public Const memoCompteLen = 245          ' 34 + 245

Type typeCompte
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
    SoldeInstantané    As Currency
    MvtceJour    As String * 1
    DécouvertAutorisé    As String * 1
    DécouvertMontant   As Currency
    DécouvertAmj  As String * 8
    MvtAmj        As String * 8
    ChéquierInterdit     As String * 1
    chkAnnul             As String * 1
    LibTyp               As String * 30
    Alpha                As String * 15
    BiaTyp               As String * 3
    BiaNum               As String * 2
    SoldeXXX             As Currency
    Groupe               As String * 6
End Type
    

Public arrCompte() As typeCompte
Public arrCompteNb As Integer
Public arrCompteNbMax As Integer
Public arrCompteIndex As Integer
Public arrComptesuite As Boolean

Public selCompte_On As Boolean, selCompte_Suite As Boolean
Public selCompte() As typeCompte
Public selCompte_Nb As Integer
Public selCompte_NbMax As Integer

Dim Xcompte As typeCompte
    

'-----------------------------------------------------
Public Function srvCompteFind(recCompte As typeCompte)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"
If Not IsNumeric(recCompte.Société) Then srvCompteFind = "? recCompte.Société": Exit Function
If Not IsNumeric(recCompte.Agence) Then srvCompteFind = "? recCompte.agence": Exit Function
If Not IsNumeric(recCompte.Devise) Then srvCompteFind = "? recCompte.devise": Exit Function
If Not IsNumeric(recCompte.Numéro) Then srvCompteFind = "? recCompte.compte": Exit Function

For I = 1 To arrCompteNb
    If recCompte.Société = arrCompte(I).Société _
    And recCompte.Agence = arrCompte(I).Agence _
    And recCompte.Devise = arrCompte(I).Devise _
    And recCompte.Numéro = arrCompte(I).Numéro Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrCompteNb + 1
    recCompte.Method = "SeekL1      "
    recCompte.BiaTyp = "000"
    recCompte.BiaNum = "00"
    arrCompte(0) = recCompte
    X = srvCompteSnap(recCompte)
    If I > arrCompteNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrCompteIndex = I
    recCompte = arrCompte(I)
End If

srvCompteFind = X

End Function



'-----------------------------------------------------
Public Function srvCompteMon(recCompte As typeCompte)
'-----------------------------------------------------
blnFR_Convert = False

arrComptesuite = False
Select Case mId$(recCompte.Method, 1, 4)
    Case "Seek", "Find"
                srvCompteMon = srvCompteSeek(recCompte)
    Case "Snap"
          srvCompteMon = srvCompteSnap(recCompte)
    Case Else
                recCompte.Err = recCompte.Method
                Call srvCompteerror(recCompte)
                srvCompteMon = recCompte.Err
End Select
'    Case "SeekL1      ", "FindFirstL5 "

'    Case "SnapL1      ", "SnapL1+     ", "SnapL5      ", "SnapL5+     ", _
         "SnapLZ      ", "SnapLZ+     ", "SnapLX      ", "SnapLX+     ", _
         "SnapLN      ", "SnapLN+     ", "SnapLA      ", "SnapLA+     "

End Function

'-----------------------------------------------------
Sub srvCompteerror(recCompte As typeCompte)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Compte : " & recCompte.Devise & "." & recCompte.Numéro & Chr$(10) & Chr$(13)

Select Case mId$(recCompte.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        If Trim(recCompte.Err) = "FIN" Then
            Msg = Msg & "Inconnu"
        Else
            Msg = Msg & "Error Code : " & recCompte.Err
        End If
        I = vbCritical
End Select

MsgBox Msg, I, "module : Cpt.bas  ( " _
                & Trim(recCompte.obj) & " : " & Trim(recCompte.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCompteGetBuffer(recCompte As typeCompte)
'---------------------------------------------------------
Dim K As Integer
srvCompteGetBuffer = Null
recCompte.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCompte.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCompte.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCompte.Err = Space$(10) Then
    recCompte.Société = mId$(MsgTxt, K + 1, 3)
    recCompte.Agence = mId$(MsgTxt, K + 4, 3)
    recCompte.Devise = mId$(MsgTxt, K + 7, 3)
    recCompte.Devisex = mId$(MsgTxt, K + 10, 3)
    recCompte.Numéro = mId$(MsgTxt, K + 13, 11)
    recCompte.NuméroAncien = mId$(MsgTxt, K + 24, 11)
    recCompte.Intitulé = mId$(MsgTxt, K + 35, 40): '' FR_ConvertEtoA_X recCompte.Intitulé
    recCompte.Intitulé2 = mId$(MsgTxt, K + 75, 40): '' FR_ConvertEtoA_X recCompte.Intitulé2
    recCompte.TypeGA = mId$(MsgTxt, K + 115, 1)
    recCompte.Situation = mId$(MsgTxt, K + 116, 1)
    recCompte.Gestionnaire = mId$(MsgTxt, K + 117, 2)
    recCompte.SoldeVeille = CCur(Val(mId$(MsgTxt, K + 119, 19)))
    recCompte.SoldeInstantané = CCur(Val(mId$(MsgTxt, K + 138, 19)))
    recCompte.MvtceJour = mId$(MsgTxt, K + 157, 1)
    recCompte.DécouvertAutorisé = mId$(MsgTxt, K + 158, 1)
    recCompte.DécouvertMontant = CCur(Val(mId$(MsgTxt, K + 159, 19)))
    recCompte.DécouvertAmj = mId$(MsgTxt, K + 178, 8)
    recCompte.MvtAmj = mId$(MsgTxt, K + 186, 8)
    recCompte.ChéquierInterdit = mId$(MsgTxt, K + 194, 1)
    recCompte.chkAnnul = mId$(MsgTxt, K + 195, 1)
    recCompte.Groupe = mId$(MsgTxt, K + 196, 6)
    recCompte.LibTyp = mId$(MsgTxt, K + 202, 24)
    recCompte.Alpha = mId$(MsgTxt, K + 226, 15)
    recCompte.BiaTyp = mId$(MsgTxt, K + 241, 3)
    recCompte.BiaNum = mId$(MsgTxt, K + 244, 2)
    If Val(recCompte.DécouvertAmj) > DSys Then
        recCompte.SoldeXXX = recCompte.SoldeInstantané + recCompte.DécouvertMontant
    Else
        recCompte.SoldeXXX = recCompte.SoldeInstantané
    End If
    If Not ctlGestionnaire_New(recCompte.Numéro, recCompte.Gestionnaire, recCompte.BiaTyp) Then
        recCompte.Intitulé2 = "! LE SOLDE DU COMPTE N'EST PAS AFFICHE"
        recCompte.SoldeVeille = 0
        recCompte.SoldeInstantané = 0
        recCompte.DécouvertMontant = 0
    End If
Else
    srvCompteGetBuffer = recCompte.Err
End If

MsgTxtIndex = MsgTxtIndex + recCompteLen

End Function

'---------------------------------------------------------
Private Sub srvComptePutBuffer(recCompte As typeCompte)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCompte.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCompte.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = recCompte.Société
Mid$(MsgTxt, K + 4, 3) = recCompte.Agence
Mid$(MsgTxt, K + 7, 3) = Format$(Val(recCompte.Devise), "000")
Mid$(MsgTxt, K + 10, 3) = recCompte.Devisex
Mid$(MsgTxt, K + 13, 11) = Format$(Val(recCompte.Numéro), "00000000000")
Mid$(MsgTxt, K + 24, 11) = recCompte.NuméroAncien
Mid$(MsgTxt, K + 35, 40) = recCompte.Intitulé
Mid$(MsgTxt, K + 75, 40) = recCompte.Intitulé2
Mid$(MsgTxt, K + 115, 1) = recCompte.TypeGA
Mid$(MsgTxt, K + 116, 1) = recCompte.Situation
Mid$(MsgTxt, K + 117, 2) = recCompte.Gestionnaire
Mid$(MsgTxt, K + 119, 1) = IIf(recCompte.SoldeVeille < 0, "-", "+")
Mid$(MsgTxt, K + 120, 18) = Format$(Abs(recCompte.SoldeVeille) * 100, "000000000000000000")
Mid$(MsgTxt, K + 138, 1) = IIf(recCompte.SoldeInstantané < 0, "-", "+")
Mid$(MsgTxt, K + 139, 18) = Format$(Abs(recCompte.SoldeInstantané) * 100, "000000000000000000")
Mid$(MsgTxt, K + 157, 1) = recCompte.MvtceJour
Mid$(MsgTxt, K + 158, 1) = recCompte.DécouvertAutorisé
Mid$(MsgTxt, K + 159, 1) = IIf(recCompte.DécouvertMontant < 0, "-", "+")
Mid$(MsgTxt, K + 160, 18) = Format$(Abs(recCompte.DécouvertMontant) * 100, "000000000000000000")
Mid$(MsgTxt, K + 178, 8) = recCompte.DécouvertAmj
Mid$(MsgTxt, K + 186, 8) = recCompte.MvtAmj
Mid$(MsgTxt, K + 194, 1) = recCompte.ChéquierInterdit
Mid$(MsgTxt, K + 195, 1) = recCompte.chkAnnul
Mid$(MsgTxt, K + 196, 6) = Format$(recCompte.Groupe, "000000")
Mid$(MsgTxt, K + 202, 24) = recCompte.LibTyp
Mid$(MsgTxt, K + 226, 15) = recCompte.Alpha
Mid$(MsgTxt, K + 241, 3) = Format$(Val(recCompte.BiaTyp), "000")
Mid$(MsgTxt, K + 244, 2) = Format$(Val(recCompte.BiaNum), "00")

MsgTxtLen = MsgTxtLen + recCompteLen
End Sub


'-----------------------------------------------------
Function srvCompte_Update(recCompte As typeCompte)
'-----------------------------------------------------

srvCompte_Update = "?"

MsgTxtLen = 0
Call srvComptePutBuffer(recCompte)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCompteGetBuffer(recCompte)) Then
        Call srvCompteerror(recCompte)
        srvCompte_Update = recCompte.Err
        Exit Function
    Else
        srvCompte_Update = Null
    End If
Else
    recCompte.Err = "srv"
End If


'=====================================================
End Function

'---------------------------------------------------------
Private Function srvCompteSeek(recCompte As typeCompte)
'---------------------------------------------------------

srvCompteSeek = "?"
MsgTxtLen = 0
Call srvComptePutBuffer(recCompte)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCompteGetBuffer(recCompte)) Then
        srvCompteSeek = Null
    Else
        Call srvCompteerror(recCompte)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCompteSnap(recCompte As typeCompte)
'---------------------------------------------------------
srvCompteSnap = "?"
MsgTxtLen = 0
Call srvComptePutBuffer(recCompte)
Call srvComptePutBuffer(arrCompte(0))
If IsNull(SndRcv()) Then
    srvCompteSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCompteGetBuffer(recCompte)) Then
            Call arrCompteAddItem(recCompte)
            arrComptesuite = True
        Else
            arrComptesuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Function selCompte_Load(recCompte As typeCompte, recCompteMax As typeCompte, Fct As String)
'---------------------------------------------------------
Dim mMethod As String

selCompte_Load = "?"
If selCompte_On Then
    Select Case Fct
        Case "Add":
        Case "End": selCompte_On = False: Exit Function
        Case Else
            Call MsgBox("sélection de comptes en cours d'utilisation par un autre processus", vbCritical, "SrvCompte.selCompte_Load")
            Exit Function
    End Select
End If

selCompte_On = True
mMethod = Trim(recCompte.Method) & "+"
If Fct <> "Add" Then selCompte_Nb = 0: ReDim selCompte(20): selCompte_NbMax = 20
selCompte_Suite = True
Do Until Not selCompte_Suite

    selCompte_Suite = False
    
    MsgTxtLen = 0
    Call srvComptePutBuffer(recCompte)
    Call srvComptePutBuffer(recCompteMax)
    
    If IsNull(SndRcv()) Then
        selCompte_Load = Null
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvCompteGetBuffer(recCompte)) Then
                If selCompte_Nb >= selCompte_NbMax Then
                    selCompte_NbMax = selCompte_NbMax + 12
                    ReDim Preserve selCompte(selCompte_NbMax)
                End If
                selCompte_Nb = selCompte_Nb + 1
                selCompte(selCompte_Nb) = recCompte
                selCompte_Suite = True
            Else
                selCompte_Suite = False
                Exit Do
            End If
        Loop
    End If
    recCompte = selCompte(selCompte_Nb)
    recCompte.Method = mMethod
Loop

End Function


'---------------------------------------------------------
Public Sub recCompteInit(recCompte As typeCompte)
'---------------------------------------------------------
MsgTxt = Space$(recCompteLen)
MsgTxtIndex = 0
Call srvCompteGetBuffer(recCompte)
recCompte.obj = "SRVCOMPTE   "
recCompte.MvtceJour = "I"
recCompte.MvtAmj = "00000000"
recCompte.chkAnnul = "1"

recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = "001"
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"

End Sub

'---------------------------------------------------------
Public Sub arrCompteAddItem(recCompte As typeCompte)
'---------------------------------------------------------
          
arrCompteNb = arrCompteNb + 1
    
If arrCompteNb > arrCompteNbMax Then
    arrCompteNbMax = arrCompteNbMax + 10
    ReDim Preserve arrCompte(arrCompteNbMax)
End If
            
arrCompte(arrCompteNb) = recCompte
End Sub

'---------------------------------------------------------
Public Function srvCompteCptInfo()
'---------------------------------------------------------
Dim X As String

srvCompteCptInfo = "?"
If arrCompteIndex > 0 And arrCompteIndex <= arrCompteNb Then

    recCptInfoInit arrCptInfo(0)
    arrCptInfo(0).Method = "JoinL1      "
    arrCptInfo(0).Société = arrCompte(arrCompteIndex).Société
    arrCptInfo(0).Agence = arrCompte(arrCompteIndex).Agence
    arrCptInfo(0).Devise = arrCompte(arrCompteIndex).Devise
    arrCptInfo(0).Numéro = arrCompte(arrCompteIndex).Numéro
    arrCptInfo(0).BiaTyp = arrCompte(arrCompteIndex).BiaTyp
    arrCptInfo(0).BiaNum = arrCompte(arrCompteIndex).BiaNum
    arrCptInfo(0).NuméroAncien = arrCompte(arrCompteIndex).NuméroAncien

    srvCompteCptInfo = srvCptInfoMon(arrCptInfo(0))
End If

End Function




Public Function srvCompte_InitFind(recCompte As typeCompte)
Xcompte = recCompte
recCompteInit recCompte
recCompte.Société = Xcompte.Société
recCompte.Agence = Xcompte.Agence
recCompte.Devise = Xcompte.Devise
recCompte.Numéro = Xcompte.Numéro

srvCompte_InitFind = srvCompteFind(recCompte)
End Function
