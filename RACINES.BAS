Attribute VB_Name = "srvRacine"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recRacineLen = 414 ' 34 + 380
Public Const memoRacineLen = 380

Type typeRacine
    obj             As String * 12
    Method          As String * 12
    Err             As String * 10
    
    Numéro            As Long                 ' 5
    Intitulé     As String * 40
    Alpha        As String * 15
    Adresse1        As String * 32
    Adresse2        As String * 32
    Adresse3        As String * 32
    AdresseCP        As String * 5
    Téléphone1       As String * 15
    Téléphone2         As String * 15
    Télex        As String * 16
    TypeBanqueClient         As String * 1
    TypeOuverture      As String * 1 'rubrique 63
    Actionnaire          As String * 1 'rubrique 7
    NatureTitulaire       As String * 2 'rubrique 62
    Résident          As String * 1 'rubrique 23
    RésidentFiscal      As String * 1 'rubrique 7
    RésidentPays      As String * 4 'rubrique 19
    MembreduPersonnel     As String * 1 'rubrique 7
    Succession   As String * 1 'rubrique 7
    Nomination   As String * 1 'rubrique 11
    RégimeMatrimonial    As String * 1 'rubrique 70
    NaissanceAmj As String * 8
    optionTva    As String * 1 'rubrique 7
    GroupeIdentification    As String * 6 'rubrique 65
    Apporteur    As String * 2 'rubrique 58
    Gestionnaire As String * 2 'rubrique 60
    Opposition   As String * 1 'rubrique 16
    NombreTitulaire       As Integer               ' Num 3
    GroupeEconomique       As String * 2 'rubrique 64
    Commentaire  As String * 20
    Courrier     As String * 1 'rubrique 61
    ChèquierInterdit  As String * 1 'rubrique 7
    ChèquierInterditAmj  As String * 8
    
    Fax          As String * 15
    Swift        As String * 11
    Nationalité  As String * 4 'rubrique 19
    APE          As String * 4 'rubrique 65
    SIREN        As String * 9
    BdfIdentification        As String * 13
    SiègeRacine     As Long             'num 5
    
    AmjCréation       As String * 8
    AmjModification       As String * 8
    AmjAnnulation       As String * 8
    AmjRéactivation       As String * 8
    AmjMouvement       As String * 8
    Opérateur        As String * 10
End Type
    
Public arrRacine() As typeRacine
Public arrRacineNb As Integer
Public arrRacineNbMax As Integer
Public arrRacineIndex As Integer
Public arrRacineSuite As Boolean
'-----------------------------------------------------
Public Function srvRacineFind(recRacine As typeRacine)
'-----------------------------------------------------
Dim I As Integer, x As Variant

x = "?"

For I = 1 To arrRacineNb
    If recRacine.Numéro = arrRacine(I).Numéro Then
        x = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(x) Then
    I = arrRacineNb + 1
    recRacine.Method = "SnapL0      "
    arrRacine(0) = recRacine
    x = srvRacineSnap(recRacine)
    If I > arrRacineNb Then x = "inconnu"
End If

If IsNull(x) Then
    arrRacineIndex = I
    recRacine = arrRacine(I)
End If

srvRacineFind = x

End Function



'-----------------------------------------------------
Public Function srvRacineMon(recRacine As typeRacine)
'-----------------------------------------------------

arrRacineSuite = False
Select Case recRacine.Method
    Case "SeekL0      "
                srvRacineMon = srvRacineSeek(recRacine)
    Case "SnapL0      ", "SnapL0+     ", "JoinL0      "
              srvRacineMon = srvRacineSnap(recRacine)
    Case Else
                recRacine.Err = recRacine.Method
                Call srvRacineerror(recRacine)
                srvRacineMon = recRacine.Err
End Select

End Function

'-----------------------------------------------------
Sub srvRacineerror(recRacine As typeRacine)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Racine" & Chr$(10) & Chr$(13)

Select Case mId$(recRacine.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recRacine.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : Racines.bas  ( " _
                & Trim(recRacine.obj) & " : " & Trim(recRacine.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvRacineGetBuffer(recRacine As typeRacine)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvRacineGetBuffer = Null
recRacine.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recRacine.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recRacine.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recRacine.Err = Space$(10) Then
    recRacine.Numéro = Val(mId$(MsgTxt, K + 1, 5))
    recRacine.Intitulé = mId$(MsgTxt, K + 6, 40): 'jpl 2001.04.04 FR_ConvertEtoA_X recRacine.Intitulé
    recRacine.Alpha = mId$(MsgTxt, K + 46, 15): 'jpl 2001.04.04 FR_ConvertEtoA_X recRacine.Alpha
    recRacine.Adresse1 = mId$(MsgTxt, K + 61, 32): 'jpl 2001.04.04 FR_ConvertEtoA_X recRacine.Adresse1
    recRacine.Adresse2 = mId$(MsgTxt, K + 93, 32): 'jpl 2001.04.04 FR_ConvertEtoA_X recRacine.Adresse2
    recRacine.Adresse3 = mId$(MsgTxt, K + 125, 32): 'jpl 2001.04.04 FR_ConvertEtoA_X recRacine.Adresse3
    recRacine.AdresseCP = mId$(MsgTxt, K + 157, 5)
    recRacine.Téléphone1 = mId$(MsgTxt, K + 162, 15)
    recRacine.Téléphone2 = mId$(MsgTxt, K + 177, 15)
    recRacine.Télex = mId$(MsgTxt, K + 192, 16)
    recRacine.TypeBanqueClient = mId$(MsgTxt, K + 208, 1)
    recRacine.TypeOuverture = mId$(MsgTxt, K + 209, 1)
    recRacine.NatureTitulaire = mId$(MsgTxt, K + 210, 2)
    recRacine.Actionnaire = mId$(MsgTxt, K + 212, 1)
    recRacine.Résident = mId$(MsgTxt, K + 213, 1)
    recRacine.RésidentFiscal = mId$(MsgTxt, K + 214, 1)
    recRacine.RésidentPays = mId$(MsgTxt, K + 215, 4)
    recRacine.MembreduPersonnel = mId$(MsgTxt, K + 219, 1)
    recRacine.Succession = mId$(MsgTxt, K + 220, 1)
    recRacine.Nomination = mId$(MsgTxt, K + 221, 1)
    recRacine.RégimeMatrimonial = mId$(MsgTxt, K + 222, 1)
    recRacine.NaissanceAmj = mId$(MsgTxt, K + 227, 4) & _
                               mId$(MsgTxt, K + 225, 2) & _
                               mId$(MsgTxt, K + 223, 2)
    recRacine.optionTva = mId$(MsgTxt, K + 231, 1)
    recRacine.GroupeIdentification = mId$(MsgTxt, K + 232, 6)
    recRacine.Apporteur = mId$(MsgTxt, K + 238, 2)
    recRacine.Gestionnaire = mId$(MsgTxt, K + 240, 2)
    recRacine.Opposition = mId$(MsgTxt, K + 242, 1)
    recRacine.NombreTitulaire = Val(mId$(MsgTxt, K + 243, 3))
    recRacine.GroupeEconomique = mId$(MsgTxt, K + 246, 2)
    recRacine.Commentaire = mId$(MsgTxt, K + 248, 20)
    recRacine.Courrier = mId$(MsgTxt, K + 268, 1)
    recRacine.ChèquierInterdit = mId$(MsgTxt, K + 269, 1)
    recRacine.ChèquierInterditAmj = mId$(MsgTxt, K + 274, 4) & _
                                 mId$(MsgTxt, K + 272, 2) & _
                                 mId$(MsgTxt, K + 270, 2)

    recRacine.Fax = mId$(MsgTxt, K + 278, 15)
    recRacine.Swift = mId$(MsgTxt, K + 293, 11)
    recRacine.Nationalité = mId$(MsgTxt, K + 304, 4)
    recRacine.APE = mId$(MsgTxt, K + 308, 4)
    recRacine.SIREN = mId$(MsgTxt, K + 312, 9)
    recRacine.BdfIdentification = mId$(MsgTxt, K + 321, 13)
    recRacine.SiègeRacine = Val(mId$(MsgTxt, K + 334, 5))
    
    recRacine.AmjCréation = mId$(MsgTxt, K + 343, 4) & _
                         mId$(MsgTxt, K + 341, 2) & _
                         mId$(MsgTxt, K + 339, 2)
    recRacine.AmjModification = mId$(MsgTxt, K + 351, 4) & _
                         mId$(MsgTxt, K + 349, 2) & _
                         mId$(MsgTxt, K + 347, 2)
    recRacine.AmjAnnulation = mId$(MsgTxt, K + 359, 4) & _
                         mId$(MsgTxt, K + 357, 2) & _
                         mId$(MsgTxt, K + 355, 2)
    recRacine.AmjRéactivation = mId$(MsgTxt, K + 367, 4) & _
                         mId$(MsgTxt, K + 365, 2) & _
                         mId$(MsgTxt, K + 363, 2)
    recRacine.Opérateur = mId$(MsgTxt, K + 371, 10)

Else
    srvRacineGetBuffer = recRacine.Err
End If

MsgTxtIndex = MsgTxtIndex + recRacineLen

End Function

'---------------------------------------------------------
Private Sub srvRacinePutBuffer(recRacine As typeRacine)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recRacine.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recRacine.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 5) = Format$(recRacine.Numéro, "00000")
MsgTxtLen = MsgTxtLen + recRacineLen
End Sub



'---------------------------------------------------------
Private Function srvRacineSeek(recRacine As typeRacine)
'---------------------------------------------------------

srvRacineSeek = "?"
MsgTxtLen = 0
Call srvRacinePutBuffer(recRacine)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvRacineGetBuffer(recRacine)) Then
        srvRacineSeek = Null
    Else
        Call srvRacineerror(recRacine)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvRacineSnap(recRacine As typeRacine)
'---------------------------------------------------------
srvRacineSnap = "?"
MsgTxtLen = 0
Call srvRacinePutBuffer(recRacine)
Call srvRacinePutBuffer(arrRacine(0))
If IsNull(SndRcv()) Then
    srvRacineSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvRacineGetBuffer(recRacine)) Then
            Call arrRacineAddItem(recRacine)
            arrRacineSuite = True
        Else
            arrRacineSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recRacineInit(recRacine As typeRacine)
'---------------------------------------------------------
 MsgTxt = Space$(recRacineLen)
 MsgTxtIndex = 0
 Call srvRacineGetBuffer(recRacine)
 recRacine.obj = "SRVRACINE   "
End Sub

'---------------------------------------------------------
Public Sub arrRacineAddItem(recRacine As typeRacine)
'---------------------------------------------------------
          
arrRacineNb = arrRacineNb + 1
    
If arrRacineNb > arrRacineNbMax Then
    arrRacineNbMax = arrRacineNbMax + 10
    ReDim Preserve arrRacine(arrRacineNbMax)
End If
            
arrRacine(arrRacineNb) = recRacine
End Sub
