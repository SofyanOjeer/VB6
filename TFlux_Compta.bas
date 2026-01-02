Attribute VB_Name = "TFlux_Compta"
Option Explicit
Public paramTFlux_Application As String * 3
Public paramTFlux_Service As String * 3
Public paramTFlux_Nature As String
Public paramTFlux_BiatypEchéance As String * 3
Public paramTFlux_BiatypEngagement As String * 3
Public paramTFlux_BiatypEngagementCorr As String * 3
Public paramTFlux_CompteIntérêts As String * 11
Public paramTFlux_CompteDéblocageDesFonds As String * 11
Public paramTFlux_CodeOpération_Compta As String * 1
Public paramTFlux_CodeOpération_Avis As String * 1

Dim recNature As typeElpTable
Dim recCodeOpération As typeElpTable

Dim C_arrCV030(7) As typeCpj030W0
Dim C_arrCV030Nb As Integer

Dim C_compte As typeCompte
Public C_Racine As typeRacine

Dim C_TFlux As typeTFlux, C_TOpe As typeTOpe
Dim C_arrTFlux() As typeTFlux
Public C_arrTFlux_Nb As Integer

Dim recGuichet_Compta As typeGuichet_Compta, xGuichet_Compta As typeGuichet_Compta
Dim recCpj030 As typeCpj030W0

Dim mCptMvtLigne As Integer
Dim paramPrêt_Intérêts As String
Dim C_CV1 As typeCV

Dim libRéférence As String, libEchéance As String, libPays As String
Dim V As Variant, X As String
Public Sub MvtCpt_AddNew()
Dim K As Integer


TFlux_Compta.Gen
recGuichet_Compta_Init recGuichet_Compta

For K = 1 To C_arrCV030Nb
    If C_arrCV030(K).MONDEV <> 0 Then
        recGuichet_Compta.Société = C_arrCV030(K).COSOC
        recGuichet_Compta.Agence = C_arrCV030(K).Agence
        recGuichet_Compta.SaisieUsr = C_arrCV030(K).NOMOP
        recGuichet_Compta.Devise = C_arrCV030(K).Devise
        recGuichet_Compta.CodeOpération = C_arrCV030(K).BIACOP
        recGuichet_Compta.CptMvtPièce = C_arrCV030(K).NUMPIE
        recGuichet_Compta.CptMvtLigne = C_arrCV030(K).NOLIGN
            mCptMvtLigne = mCptMvtLigne + 1
            recGuichet_Compta.CptMvtLigne = mCptMvtLigne
        recGuichet_Compta.Service = C_arrCV030(K).SERVIC
        recGuichet_Compta.Compte = C_arrCV030(K).Compte
        recGuichet_Compta.Montant = C_arrCV030(K).MONDEV
        recGuichet_Compta.Sens = C_arrCV030(K).SENECR
        recGuichet_Compta.AmjOpération = C_arrCV030(K).AMJOPE
        recGuichet_Compta.AmjValeur = C_arrCV030(K).AMJVAL
        recGuichet_Compta.Libellé = C_arrCV030(K).LIBELE
        recGuichet_Compta.SaisieAmj = C_arrCV030(K).AMJSAI
'''        recGuichet_Compta.SaisieHMS = "000000"   'C_arrCV030(K).SaisieHMS
        recGuichet_Compta.Référence = C_arrCV030(K).REFCON
        recGuichet_Compta.chkCompte = "0"
        recGuichet_Compta.chkSolde = "0"
        recGuichet_Compta.chkChèque = "0"
        recGuichet_Compta.chkAmjValeur = "0"
       
        recGuichet_Compta.Method = "AddNew"

        dbGuichet_Compta_Update recGuichet_Compta
    End If
Next K

End Sub

Public Sub LotàCompta_Demande(mLot As Long)
Dim I As Integer, Msg As String

ReDim C_arrTFlux(1)

recTFlux_Init C_TFlux
C_TFlux.Method = "SnapLC"
C_TFlux.Statut = constàCompta
C_TFlux.ElpControl = paramTFlux_Service
C_TFlux.CptMvtLot = mLot

C_arrTFlux(0) = C_TFlux
C_arrTFlux(0).CptMvtPièce = 999999999
C_arrTFlux(0).CptMvtLigne = 999999999

Call srvTFlux_Load(C_TFlux, C_arrTFlux(0))
C_arrTFlux_Nb = srvTFlux.arrTFlux_Nb
ReDim C_arrTFlux(arrTFlux_Nb)
For I = 1 To C_arrTFlux_Nb
    C_arrTFlux(I) = srvTFlux.arrTFlux(I)
Next I

Msg = ""
Msg = Format$(mLot, "000000000000")
MvtCpt_Init

For I = 1 To C_arrTFlux_Nb
        C_TFlux = C_arrTFlux(I)
''$JPL20001025        If Trim(C_TFlux.CptMvtUsr) = "" Then
           MvtCpt_AddNew
''$JPL20001025        End If
Next I

prtCompta_Monitor Msg, constDemandeDeValidation, conststrTFlux_Compta
End Sub

Public Sub Gen_Pr01()
C_arrCV030Nb = 2

C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EngagementCompte
C_arrCV030(1).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(1).SENECR = "D"
C_arrCV030(1).LIBELE = "Déblocage des fonds du prêt " & libRéférence

C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = C_TOpe.EngagementCorrCompte
C_arrCV030(2).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(2).SENECR = "C"
C_arrCV030(2).LIBELE = C_arrCV030(1).LIBELE

End Sub
Public Sub Gen_GA01()
C_arrCV030Nb = 2

C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EngagementCompte
C_arrCV030(1).MONDEV = Abs(C_TFlux.Capital)
C_arrCV030(1).SENECR = "D"
C_arrCV030(1).LIBELE = "Emission garantie " & libRéférence

C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = C_TOpe.EngagementCorrCompte
C_arrCV030(2).MONDEV = C_arrCV030(1).MONDEV
C_arrCV030(2).SENECR = "C"
C_arrCV030(2).LIBELE = C_arrCV030(1).LIBELE

End Sub

Public Sub Gen_GA02()
C_arrCV030Nb = 2

C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EngagementCompte
C_arrCV030(1).MONDEV = Abs(C_TFlux.Capital)
C_arrCV030(1).SENECR = "C"
C_arrCV030(1).LIBELE = "Fin validité " & libRéférence

C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = C_TOpe.EngagementCorrCompte
C_arrCV030(2).MONDEV = C_arrCV030(1).MONDEV
C_arrCV030(2).SENECR = "D"
C_arrCV030(2).LIBELE = C_arrCV030(1).LIBELE

End Sub


Public Sub Gen_GA04()
C_arrCV030Nb = 2

C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EngagementCompte
C_arrCV030(1).MONDEV = Abs(C_TFlux.Capital)
C_arrCV030(1).SENECR = "C"
C_arrCV030(1).LIBELE = "Main levée partielle " & libRéférence

C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = C_TOpe.EngagementCorrCompte
C_arrCV030(2).MONDEV = C_arrCV030(1).MONDEV
C_arrCV030(2).SENECR = "D"
C_arrCV030(2).LIBELE = C_arrCV030(1).LIBELE

End Sub

Public Sub Gen_GA11()
C_arrCV030Nb = 2

C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EngagementCompte
C_arrCV030(1).MONDEV = Abs(C_TFlux.Capital)
C_arrCV030(1).SENECR = "D"
C_arrCV030(1).LIBELE = "avenant d'augmentation " & libRéférence

C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = C_TOpe.EngagementCorrCompte
C_arrCV030(2).MONDEV = C_arrCV030(1).MONDEV
C_arrCV030(2).SENECR = "C"
C_arrCV030(2).LIBELE = C_arrCV030(1).LIBELE

End Sub

Public Sub Gen_GA03()
C_arrCV030Nb = 2

C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EngagementCompte
C_arrCV030(1).MONDEV = Abs(C_TFlux.Capital)
C_arrCV030(1).SENECR = "C"
C_arrCV030(1).LIBELE = "Main levée " & libRéférence

C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = C_TOpe.EngagementCorrCompte
C_arrCV030(2).MONDEV = C_arrCV030(1).MONDEV
C_arrCV030(2).SENECR = "D"
C_arrCV030(2).LIBELE = C_arrCV030(1).LIBELE

End Sub


Public Sub Gen_Pr05()
C_arrCV030Nb = 2

C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EngagementCompte
C_arrCV030(1).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(1).SENECR = "C"
C_arrCV030(1).LIBELE = "Remboursement anticipé du prêt " & libRéférence

C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = C_TOpe.EchéanceCompte
C_arrCV030(2).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(2).SENECR = "D"
C_arrCV030(2).LIBELE = C_arrCV030(1).LIBELE

End Sub

Public Sub Gen_Pr02()
C_arrCV030Nb = 3
C_arrCV030(1).Compte = C_TOpe.EchéanceCompte
C_arrCV030(1).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(1).SENECR = "D"
C_arrCV030(1).LIBELE = "Echéance du " & libEchéance & " votre prêt " & libRéférence


C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = C_TOpe.EngagementCompte
C_arrCV030(2).MONDEV = C_TFlux.Capital
C_arrCV030(2).SENECR = "C"
C_arrCV030(2).LIBELE = "Amortissement du " & libEchéance & " votre prêt " & libRéférence


C_arrCV030(3) = C_arrCV030(0)
C_arrCV030(3).NOLIGN = 3
C_arrCV030(3).Compte = paramTFlux_CompteIntérêts
C_arrCV030(3).MONDEV = C_TFlux.Intérêts
C_arrCV030(3).SENECR = "C"
C_arrCV030(3).LIBELE = libPays & "Intérêts du " & libEchéance & " prêt " & libRéférence


End Sub

Public Sub Gen_Pr03()
C_arrCV030Nb = 2

C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EchéanceCompte
C_arrCV030(1).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(1).SENECR = "D"
C_arrCV030(1).LIBELE = "Frais de dossier de votre prêt " & libRéférence

C_arrCV030(2).NOLIGN = 2
C_arrCV030(2).Compte = paramTFlux_CompteIntérêts
C_arrCV030(2).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(2).SENECR = "C"
C_arrCV030(2).LIBELE = libPays & "Frais de dossier du prêt " & libRéférence

End Sub

Public Sub Gen_Pr04()
C_arrCV030Nb = 2
C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EchéanceCompte
C_arrCV030(1).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(1).SENECR = "D"
C_arrCV030(1).LIBELE = "Intérêts intermédiaires de votre prêt " & libRéférence


C_arrCV030(2).NOLIGN = 3
C_arrCV030(2).Compte = paramTFlux_CompteIntérêts
C_arrCV030(2).MONDEV = C_TFlux.Capital + C_TFlux.Intérêts
C_arrCV030(2).SENECR = "C"
C_arrCV030(2).LIBELE = libPays & "Intérêts intermédiaires du prêt " & libRéférence


End Sub


Public Sub Gen_GA51()
C_arrCV030Nb = 2
C_arrCV030(1).NOLIGN = 1
C_arrCV030(1).Compte = C_TOpe.EchéanceCompte
C_arrCV030(1).MONDEV = Abs(C_TFlux.Intérêts)
C_arrCV030(1).SENECR = "D"
C_arrCV030(1).LIBELE = "Commission garantie " & libRéférence


C_arrCV030(2).NOLIGN = 3
C_arrCV030(2).Compte = paramTFlux_CompteIntérêts
C_arrCV030(2).MONDEV = C_arrCV030(1).MONDEV
C_arrCV030(2).SENECR = "C"
C_arrCV030(2).LIBELE = libPays & "Commission garantie " & libRéférence


End Sub


Public Sub LotComptabilisé_Print(xNumlot As String)
'recCptMvtInit minCptMvt
'minCptMvt.obj = "SRVECRITG"
'minCptMvt.Method = "SnapK0"
'minCptMvt.Société = SocId$
'minCptMvt.Agence = SocAgence$
'minCptMvt.Devise = "000"
'minCptMvt.Lot = Val(xNumlot)
'minCptMvt.Pièce = 0
'minCptMvt.Ligne = 0

'maxCptMvt = minCptMvt
'maxCptMvt.Devise = "999"
'maxCptMvt.Pièce = 999999999
'maxCptMvt.Ligne = 9999
'Call srvCptMvt_ElpBuffer(minCptMvt, maxCptMvt, G_ElpBuffer)
'If G_ElpBuffer.Seq > 0 Then prtCompta_Monitor G_ElpBuffer.Id, "Lot N°" & Format$(xNumlot, "### ### ###"), conststrTFlux_Comptabilisé

End Sub


Public Sub LotàCompta_Valider(mTflux As typeTFlux)
Dim I As Integer, Msg As String, trimusrId As String, mNumlot As Long
Dim recTFlux As typeTFlux

MvtCpt_Init
mTflux.Statut = "C"
recTFlux = mTflux
mNumlot = mTflux.CptMvtLot

recTFlux.Method = "SnapLot"
recTFlux.CptMvtPièce = 0
recTFlux.CptMvtLigne = 0

srvTFlux.arrTFlux_NbMax = 35: ReDim srvTFlux.arrTFlux(35)
srvTFlux.arrTFlux_Nb = 0: srvTFlux.arrTFlux_Index = 0

srvTFlux.arrTFlux(0) = recTFlux
srvTFlux.arrTFlux(0).CptMvtPièce = 9999999
srvTFlux.arrTFlux(0).CptMvtLigne = 9999999

srvTFlux.arrTFlux_Suite = True
Do Until Not srvTFlux.arrTFlux_Suite
    srvTFlux_Monitor recTFlux
    recTFlux = srvTFlux.arrTFlux(srvTFlux.arrTFlux_Nb)
    recTFlux.Method = "SnapLot+"
Loop
Msg = "000000000000"

For I = 1 To srvTFlux.arrTFlux_Nb
    C_TFlux = srvTFlux.arrTFlux(I)
    MvtCpt_AddNew
Next I

mCptMvtLigne = 0
recTFlux = mTflux

LotàCompta_Valider_Send recTFlux

recTFlux = mTflux
recTFlux.CptMvtLigne = mCptMvtLigne

recTFlux.Method = "Compta"
If Not IsNull(srvTFlux_Update(recTFlux)) Then Call MsgBox("Erreur informatique : ANNULER LA VALIDATION", vbCritical, "Validation comptable des opérations")
If recTFlux.CptMvtLigne <> 0 Or recTFlux.Capital <> 0 Then Call MsgBox("anomalie en nombre ou en montant : ANNULER LA VALIDATION", vbCritical, "Validation comptable des opérations ")

Msg = "000000000000"
prtCompta_Monitor Msg, "VALIDATION Lot N°" & Format$(mNumlot, "### ### ###"), conststrTFlux_Compta
End Sub
'---------------------------------------------------------
Public Sub LotàCompta_Valider_Send(mTflux As typeTFlux)
'---------------------------------------------------------
Dim iReturn As Integer, mCpj030 As typeCpj030W0

recGuichet_Compta.Method = "MoveFirst"
        recCpj030W0_Init recCpj030
        recCpj030.obj = "SRVCPJ030H"
        recCpj030.Method = "AddNew"
        recCpj030.NUMLOT = mTflux.CptMvtLot
        recCpj030.CTLNOM = usrId
        recCpj030.IMPAMJ = mTflux.CptMvtAMJ
        recCpj030.IMPHMS = mTflux.CptMvtHMS
        recCpj030.CTLAMJ = mTflux.CptMvtAMJ
        recCpj030.CTLHMS = mTflux.CptMvtHMS
        recCpj030.JJCPLT = "0"
        recCpj030.CTLSTA = "9"
        recCpj030.COSOC = recGuichet_Compta.Société
mCpj030 = recCpj030

srvCpj030W0_Dtaq_Put "Init", recCpj030

Do
    iReturn = tableGuichet_Compta_Read(recGuichet_Compta)
    If iReturn = 0 Then
        recCpj030 = mCpj030
        recCpj030.Agence = recGuichet_Compta.Agence
        recCpj030.AGEMET = recGuichet_Compta.Agence
        recCpj030.Devise = recGuichet_Compta.Devise
        recCpj030.BIACOP = recGuichet_Compta.CodeOpération
        recCpj030.NUMPIE = recGuichet_Compta.CptMvtPièce
        mCptMvtLigne = mCptMvtLigne + 1
        recCpj030.NOLIGN = mCptMvtLigne
        recCpj030.NOMOP = recGuichet_Compta.SaisieUsr
        recCpj030.SERVIC = recGuichet_Compta.Service
        recCpj030.Compte = recGuichet_Compta.Compte
        recCpj030.MONDEV = recGuichet_Compta.Montant
        recCpj030.SENECR = recGuichet_Compta.Sens
        recCpj030.AMJOPE = recGuichet_Compta.AmjOpération
        recCpj030.AMJVAL = recGuichet_Compta.AmjValeur
        recCpj030.LIBELE = recGuichet_Compta.Libellé
        recCpj030.AMJSAI = recGuichet_Compta.SaisieAmj
        recCpj030.CODFOR = recGuichet_Compta.chkCompte
        recCpj030.FOROPO = recGuichet_Compta.chkSolde
        recCpj030.OPOCHQ = recGuichet_Compta.chkChèque
        recCpj030.FORVAL = recGuichet_Compta.chkAmjValeur
     
        srvCpj030W0_Dtaq_Put "Add", recCpj030
    End If
    recGuichet_Compta.Method = "MoveNext"
Loop Until iReturn <> 0

srvCpj030W0_Dtaq_Put "Snd", recCpj030

End Sub




Public Sub MvtCpt_Init()
On Error Resume Next

tableGuichet_Compta_Close
MDB.Execute "delete * from Guichet_Compta"
tableGuichet_Compta_Open

recGuichet_Compta_Init recGuichet_Compta
mCptMvtLigne = 0

End Sub

Public Function Param_Init(Msg As String, cbo As ComboBox)
Dim V
Param_Init = Null
recCompteInit C_compte
recRacineInit C_Racine
C_CV1 = CV_Euro

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "TOPE" '"Param"
recElpTable.K1 = Msg
recElpTable.K2 = "Service"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTFlux_Service = mId$(recElpTable.Memo, 1, 3)
If Not IsNumeric(paramTFlux_Service) Then GoTo Num_Error

cbo.Clear

recElpTable_Init recElpTable
recElpTable.Method = "Seek>="
recElpTable.Id = "TOPE" '"Param"
recElpTable.K1 = Msg & "Nature"
recElpTable.Err = 0
recNature = recElpTable

Do
    recElpTable.Err = tableElpTable_Read(recElpTable)
    If recElpTable.Err = 0 Then
        If recElpTable.K1 <> recNature.K1 Then
            recElpTable.Err = 9996
        Else
            cbo.AddItem mId$(recElpTable.K2, 1, 8) & recElpTable.Name

            recElpTable.Method = "Seek>"
       End If
    End If
Loop While recElpTable.Err = 0


recElpTable_Init recCodeOpération
recCodeOpération.Method = "Seek="
recCodeOpération.Id = "TOPE"
recCodeOpération.K1 = Msg & "Opé"
recCodeOpération.Err = 0

Exit Function

Table_Error:
Param_Init = V
Exit Function

Memo_Error:
Param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "TfluxEspèces_Compta_gen"
Exit Function

Num_Error:
Param_Init = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "TfluxEspèces_Param_Init"
End Function
Public Function Param_InitNew(Msg As String, cbo As ComboBox)
Dim V
Param_InitNew = Null
recCompteInit C_compte
recRacineInit C_Racine
C_CV1 = CV_Euro

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = Msg
recElpTable.K1 = "Application"
recElpTable.K2 = "Code"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTFlux_Application = mId$(recElpTable.Memo, 1, 3)

recElpTable.K2 = "Service"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTFlux_Service = mId$(recElpTable.Memo, 1, 3)
If Not IsNumeric(paramTFlux_Service) Then GoTo Num_Error

cbo.Clear

recElpTable_Init recElpTable
recElpTable.Method = "Seek>="
recElpTable.Id = Msg
recElpTable.K1 = "Nature"
recElpTable.Err = 0
recNature = recElpTable

Do
    recElpTable.Err = tableElpTable_Read(recElpTable)
    If recElpTable.Err = 0 Then
        If recElpTable.K1 <> recNature.K1 Then
            recElpTable.Err = 9996
        Else
            cbo.AddItem mId$(recElpTable.K2, 1, 8) & recElpTable.Name

            recElpTable.Method = "Seek>"
       End If
    End If
Loop While recElpTable.Err = 0


recElpTable_Init recCodeOpération
recCodeOpération.Method = "Seek="
recCodeOpération.Id = Msg
recCodeOpération.K1 = "Opération"
recCodeOpération.Err = 0

Exit Function

Table_Error:
Param_InitNew = V
Exit Function

Memo_Error:
Param_InitNew = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "TfluxEspèces_Compta_gen"
Exit Function

Num_Error:
Param_InitNew = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "TfluxEspèces_Param_Init"
End Function









Public Sub Gen()
Dim X11 As String * 11

recTOpe_Init C_TOpe
C_TOpe.Method = "SeekP0"
C_TOpe.IdRéférence = C_TFlux.IdRéférence
C_TOpe.Application = paramTFlux_Service
V = srvTOpe_Monitor(C_TOpe)
If Not IsNull(V) Then
    MsgBox ("erreur tflux_compta gen lectute Tope")
    Exit Sub
End If
Param_Nature C_TOpe.Nature

X = Trim(C_TOpe.RéférenceExterne)
If X = "" Then
    libRéférence = Trim(C_TOpe.RéférenceInterne)
Else
    libRéférence = Trim(C_TOpe.RéférenceInterne) & " / " & X
End If

libEchéance = dateImp(C_TFlux.AmjEchéanceTrt)

recCpj030W0_Init C_arrCV030(0)
C_arrCV030(0).Method = "AddNew"
C_arrCV030(0).COSOC = SocId$
C_arrCV030(0).Agence = SocAgence$
C_arrCV030(0).AGEMET = SocAgence$
C_arrCV030(0).BIACOP = C_TFlux.CodeOpération
C_arrCV030(0).SERVIC = C_TOpe.Application
C_arrCV030(0).AMJSAI = DSys
C_arrCV030(0).AMJVAL = C_TFlux.AmjValeur
C_arrCV030(0).AMJOPE = DSys 'C_TFlux.AmjOpération
C_arrCV030(0).JJCPLT = "0"
C_arrCV030(0).SIGENE = "*"

C_CV1.DeviseIso = C_TOpe.Devise: CV_Attribut C_CV1
C_arrCV030(0).Devise = C_CV1.DeviseN

C_arrCV030(0).NUMLOT = C_TFlux.CptMvtLot
C_arrCV030(0).NUMPIE = C_TFlux.CptMvtPièce
C_arrCV030(0).NOLIGN = C_TFlux.CptMvtLigne

C_arrCV030(0).NOMOP = C_TFlux.CptMvtUsr
C_arrCV030(0).CTLAMJ = C_TFlux.CptMvtAMJ
C_arrCV030(0).REFCON = C_TFlux.IdRéférence

C_arrCV030(1) = C_arrCV030(0)
C_arrCV030(1).NOLIGN = 1
C_arrCV030(2) = C_arrCV030(0)
C_arrCV030(2).NOLIGN = 2

C_Racine.Numéro = CLng(Val(mId$(C_TOpe.EchéanceCompte, 1, 6)))
C_Racine.Method = "SeekL0"
If IsNull(srvRacineFind(C_Racine)) Then
    libPays = mId$(C_Racine.RésidentPays, 2, 3) & " "
Else
    libPays = "999 "
End If

Select Case Trim(C_TFlux.CodeOpération)

    Case "PR01": Gen_Pr01
    Case "PR02": Gen_Pr02
    Case "PR03": Gen_Pr03
    Case "PR04": Gen_Pr04
    Case "PR05": Gen_Pr05

    Case "GA01": Gen_GA01
    Case "GA02": Gen_GA02
    Case "GA03": Gen_GA03
    Case "GA04": Gen_GA04
    Case "GA11": Gen_GA11
    Case "GA51": Gen_GA51
    Case "GA52": Gen_GA51

End Select

End Sub

Public Function Param_Nature(mNature As String)
Dim L As Integer
Param_Nature = Null
If Trim(recNature.K2) <> Trim(mNature) Then
    recNature.Method = "Seek="
    recNature.K2 = mNature
    paramTFlux_Nature = mNature
    recNature.Err = tableElpTable_Read(recNature)
    If recNature.Err <> 0 Then
        recNature.Memo = "000 000 000 00000000000 00000000000"
        Call MsgBox("TFlux_Compta : Param_Nature", vbCritical, "Nature inconnue : " & mNature)
        Param_Nature = "? " & mNature
        recNature.K2 = ""
    End If
    paramTFlux_BiatypEchéance = mId$(recNature.Memo, 1, 3)
    paramTFlux_BiatypEngagement = mId$(recNature.Memo, 5, 3)
    paramTFlux_BiatypEngagementCorr = mId$(recNature.Memo, 9, 3)
    paramTFlux_CompteIntérêts = mId$(recNature.Memo, 13, 11)
    paramTFlux_CompteDéblocageDesFonds = mId$(recNature.Memo, 25, 11)
    L = Len(recNature.Memo) - 37
    If L > 0 Then
        paramTFlux_Nature = mId$(recNature.Memo, 37, L) 'recNature.Name
    Else
        paramTFlux_Nature = recNature.Name
    End If
End If

End Function
Public Function Param_CodeOpération(mCodeOpération As String)
Param_CodeOpération = Null
If Trim(recCodeOpération.K2) <> Trim(mCodeOpération) Then
    recCodeOpération.Method = "Seek="
    recCodeOpération.K2 = mCodeOpération
    recCodeOpération.Err = tableElpTable_Read(recCodeOpération)
    If recCodeOpération.Err <> 0 Then
        Call MsgBox("TFlux_Compta : Param_CodeOpération", vbCritical, "CodeOpération inconnue : " & mCodeOpération)
        Param_CodeOpération = "? " & mCodeOpération
        recCodeOpération.K2 = ""
    End If
End If
Param_CodeOpération = recCodeOpération.Name
paramTFlux_CodeOpération_Compta = mId$(recCodeOpération.Memo, 1, 1)
paramTFlux_CodeOpération_Avis = mId$(recCodeOpération.Memo, 3, 1)

End Function



