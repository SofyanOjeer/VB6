Attribute VB_Name = "srvTitulaire"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recTitulaireLen = 662 ' 34 + 628

Type typeTitulaire
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Racine                  As Long                 ' Num 5
    No                      As Integer              ' Num 3
    
    Nationalité             As String * 4           ' Rub 19
    RésidencePays           As String * 4           ' Rub 19
    RésidenceZone           As String * 1           ' Rub 71
    
    NomPatronymique         As String * 64
    Prénom                  As String * 64
    MariNom                 As String * 32
    MariPrénom              As String * 32
    Adresse1                As String * 32
    Adresse2                As String * 32
    Adresse3                As String * 32
    Adresse4                As String * 32
    AdresseCP               As String * 5
    AdresseBD               As String * 26
    
    NaissanceAmj            As String * 8
    NaissanceDépN           As String * 2           ' Rub 51
    NaissanceDépX           As String * 32
    NaissanceCommune        As String * 32
    NaissancePays           As String * 4           ' Rub 19
    NaissanceZone           As String * 1           ' Rub 71
    
    DocIdentitéType         As String * 1           ' Rub 73
    DocIdentitéDélivréPar   As String * 20
    DocIdentitéAmj          As String * 8
    DocIdentitéNo           As String * 15
    
    PmFormeJuridique       As String * 2           ' Rub 3
    PmSiren                As String * 9
    PmApe                  As String * 4           ' Rub 5
    PmBdfId                As String * 13
    PmNationalitéMère      As String * 4           ' Rub 19
       
    PpQualité               As String * 1           ' Rub 69
    PpRégimeMatrimonial     As String * 1           ' Rub 70
    PpNombreEnfant          As String * 1
    PpBdfId                 As String * 11
    PpProfession            As String * 32
    PpCsp                   As String * 1           ' Rub 72
    PpAmjEntrée             As String * 8
    PpAmjSortie             As String * 8
   
    AmjCre                  As String * 8
    AmjMod                  As String * 8
    AmjAnn                  As String * 8
    AmjRea                  As String * 8
    NomOp                   As String * 10
End Type
    
Public arrTitulaire() As typeTitulaire
Public arrTitulaireNb As Integer
Public arrTitulaireNbMax As Integer
Public arrTitulaireIndex As Integer
Public arrTitulaireSuite As Boolean
'-----------------------------------------------------
Public Function srvTitulaireFind(recTitulaire As typeTitulaire)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrTitulaireNb
    If recTitulaire.Racine = arrTitulaire(I).Racine _
    And recTitulaire.No = arrTitulaire(I).No Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrTitulaireNb + 1
    recTitulaire.Method = "SnapL0      "
    arrTitulaire(0) = recTitulaire
    X = srvTitulaireSnap(recTitulaire)
    If I > arrTitulaireNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrTitulaireIndex = I
    recTitulaire = arrTitulaire(I)
End If

srvTitulaireFind = X

End Function



'-----------------------------------------------------
Public Function srvTitulaireMon(recTitulaire As typeTitulaire)
'-----------------------------------------------------

arrTitulaireSuite = False
Select Case recTitulaire.Method
    Case "SeekL0      "
                srvTitulaireMon = srvTitulaireSeek(recTitulaire)
    Case "SnapL0      ", "SnapL0+     "
              srvTitulaireMon = srvTitulaireSnap(recTitulaire)
    Case Else
                recTitulaire.Err = recTitulaire.Method
                Call srvTitulaireerror(recTitulaire)
                srvTitulaireMon = recTitulaire.Err
End Select

End Function

'-----------------------------------------------------
Sub srvTitulaireerror(recTitulaire As typeTitulaire)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Titulaire" & Chr$(10) & Chr$(13)

Select Case Mid$(recTitulaire.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recTitulaire.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : Titulaires.bas  ( " _
                & Trim(recTitulaire.obj) & " : " & Trim(recTitulaire.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvTitulaireGetBuffer(recTitulaire As typeTitulaire)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvTitulaireGetBuffer = Null
recTitulaire.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recTitulaire.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recTitulaire.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recTitulaire.Err = Space$(10) Then
    recTitulaire.Racine = Val(Mid$(MsgTxt, K + 1, 5))
    recTitulaire.No = Val(Mid$(MsgTxt, K + 6, 3))
    recTitulaire.Nationalité = Mid$(MsgTxt, K + 9, 4)
    recTitulaire.RésidencePays = Mid$(MsgTxt, K + 13, 4)
    recTitulaire.RésidenceZone = Mid$(MsgTxt, K + 17, 1)
    recTitulaire.NomPatronymique = Mid$(MsgTxt, K + 18, 64)
    recTitulaire.Prénom = Mid$(MsgTxt, K + 82, 64)
    recTitulaire.MariNom = Mid$(MsgTxt, K + 146, 32)
    recTitulaire.MariPrénom = Mid$(MsgTxt, K + 178, 32)
    recTitulaire.Adresse1 = Mid$(MsgTxt, K + 210, 32)
    recTitulaire.Adresse2 = Mid$(MsgTxt, K + 242, 32)
    recTitulaire.Adresse3 = Mid$(MsgTxt, K + 274, 32)
    recTitulaire.Adresse4 = Mid$(MsgTxt, K + 306, 32)
    recTitulaire.AdresseCP = Mid$(MsgTxt, K + 338, 5)
    recTitulaire.AdresseBD = Mid$(MsgTxt, K + 343, 26)
    recTitulaire.NaissanceAmj = Mid$(MsgTxt, K + 373, 4) & _
                               Mid$(MsgTxt, K + 371, 2) & _
                               Mid$(MsgTxt, K + 369, 2)
    recTitulaire.NaissanceDépN = Mid$(MsgTxt, K + 377, 2)
    recTitulaire.NaissanceDépX = Mid$(MsgTxt, K + 379, 32)
    recTitulaire.NaissanceCommune = Mid$(MsgTxt, K + 411, 32)
    recTitulaire.NaissancePays = Mid$(MsgTxt, K + 443, 4)
    recTitulaire.NaissanceZone = Mid$(MsgTxt, K + 447, 1)
    recTitulaire.DocIdentitéType = Mid$(MsgTxt, K + 448, 1)
    recTitulaire.DocIdentitéDélivréPar = Mid$(MsgTxt, K + 449, 20)
    recTitulaire.DocIdentitéAmj = Mid$(MsgTxt, K + 473, 4) & _
                                 Mid$(MsgTxt, K + 471, 2) & _
                                 Mid$(MsgTxt, K + 469, 2)
recTitulaire.DocIdentitéNo = Mid$(MsgTxt, K + 477, 15)
    recTitulaire.PmFormeJuridique = Mid$(MsgTxt, K + 492, 2)
    recTitulaire.PmSiren = Mid$(MsgTxt, K + 494, 9)
    recTitulaire.PmApe = Mid$(MsgTxt, K + 503, 4)
    recTitulaire.PmBdfId = Mid$(MsgTxt, K + 507, 13)
    recTitulaire.PmNationalitéMère = Mid$(MsgTxt, K + 520, 4)
    recTitulaire.PpQualité = Mid$(MsgTxt, K + 524, 1)
    recTitulaire.PpRégimeMatrimonial = Mid$(MsgTxt, K + 525, 1)
    recTitulaire.PpNombreEnfant = Mid$(MsgTxt, K + 526, 1)
    recTitulaire.PpBdfId = Mid$(MsgTxt, K + 527, 11)
    recTitulaire.PpProfession = Mid$(MsgTxt, K + 538, 32)
    recTitulaire.PpCsp = Mid$(MsgTxt, K + 570, 1)
    recTitulaire.PpAmjSortie = Mid$(MsgTxt, K + 575, 4) & _
                              Mid$(MsgTxt, K + 573, 2) & _
                              Mid$(MsgTxt, K + 571, 2)
    recTitulaire.PpAmjEntrée = Mid$(MsgTxt, K + 583, 4) & _
                              Mid$(MsgTxt, K + 581, 2) & _
                              Mid$(MsgTxt, K + 579, 2)

    recTitulaire.AmjCre = Mid$(MsgTxt, K + 591, 4) & _
                         Mid$(MsgTxt, K + 589, 2) & _
                         Mid$(MsgTxt, K + 587, 2)
    recTitulaire.AmjMod = Mid$(MsgTxt, K + 599, 4) & _
                         Mid$(MsgTxt, K + 597, 2) & _
                         Mid$(MsgTxt, K + 595, 2)
    recTitulaire.AmjAnn = Mid$(MsgTxt, K + 607, 4) & _
                         Mid$(MsgTxt, K + 605, 2) & _
                         Mid$(MsgTxt, K + 603, 2)
    recTitulaire.AmjRea = Mid$(MsgTxt, K + 615, 4) & _
                         Mid$(MsgTxt, K + 613, 2) & _
                         Mid$(MsgTxt, K + 611, 2)
    recTitulaire.NomOp = Mid$(MsgTxt, K + 619, 10)

Else
    srvTitulaireGetBuffer = recTitulaire.Err
End If

MsgTxtIndex = MsgTxtIndex + recTitulaireLen

End Function

'---------------------------------------------------------
Private Sub srvTitulairePutBuffer(recTitulaire As typeTitulaire)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recTitulaire.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recTitulaire.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 5) = Format$(recTitulaire.Racine, "00000")
Mid$(MsgTxt, K + 6, 3) = Format$(recTitulaire.No, "000")
MsgTxtLen = MsgTxtLen + recTitulaireLen
End Sub



'---------------------------------------------------------
Private Function srvTitulaireSeek(recTitulaire As typeTitulaire)
'---------------------------------------------------------

srvTitulaireSeek = "?"
MsgTxtLen = 0
Call srvTitulairePutBuffer(recTitulaire)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvTitulaireGetBuffer(recTitulaire)) Then
        srvTitulaireSeek = Null
    Else
        Call srvTitulaireerror(recTitulaire)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvTitulaireSnap(recTitulaire As typeTitulaire)
'---------------------------------------------------------
srvTitulaireSnap = "?"
MsgTxtLen = 0
Call srvTitulairePutBuffer(recTitulaire)
Call srvTitulairePutBuffer(arrTitulaire(0))
If IsNull(SndRcv()) Then
    srvTitulaireSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvTitulaireGetBuffer(recTitulaire)) Then
            Call arrTitulaireAddItem(recTitulaire)
            arrTitulaireSuite = True
        Else
            arrTitulaireSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recTitulaireInit(recTitulaire As typeTitulaire)
'---------------------------------------------------------
 MsgTxt = Space$(recTitulaireLen)
 MsgTxtIndex = 0
 Call srvTitulaireGetBuffer(recTitulaire)
 recTitulaire.obj = "SRVTITULA"
End Sub

'---------------------------------------------------------
Public Sub arrTitulaireAddItem(recTitulaire As typeTitulaire)
'---------------------------------------------------------
          
arrTitulaireNb = arrTitulaireNb + 1
    
If arrTitulaireNb > arrTitulaireNbMax Then
    arrTitulaireNbMax = arrTitulaireNbMax + 10
    ReDim Preserve arrTitulaire(arrTitulaireNbMax)
End If
            
arrTitulaire(arrTitulaireNb) = recTitulaire
End Sub
