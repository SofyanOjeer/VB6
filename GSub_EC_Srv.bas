Attribute VB_Name = "srvGSub_EC"
Option Explicit

Public paramEffetCommerce_TauxMin As Double
Public paramEffetCommerce_TauxMinàConfirmer As Double
Public paramEffetCommerce_TauxMax As Double
Public paramEffetCommerce_TauxMaxàConfirmer As Double
Public paramEffetCommerce_TauxMargeMajoré As Double
Public paramEffetCommerce_TauxMargeNonAccepté As Double

Public paramEffetCommerce_NbjRemMax As String, paramEffetCommerce_AmjRemMax As String * 8
Public paramEffetCommerce_NbjPrésentation As String, paramEffetCommerce_AmjPrésentation As String * 8
Public paramEffetCommerce_NbjEncCpt As String, paramEffetCommerce_AmjEncCpt As String * 8
Public paramEffetCommerce_NbjEncMin As String, paramEffetCommerce_AmjEncMin As String * 8
Public paramEffetCommerce_NbjEscMin As String, paramEffetCommerce_AmjEscMin As String * 8
Public paramEffetCommerce_NbjEscVal As String, paramEffetCommerce_AmjEscVal As String * 8

Public paramEffetCommerce_CompteAgios As String
Public paramEffetCommerce_CompteCompensateur As String
Public paramEffetCommerce_CompteComTaxable As String
Public paramEffetCommerce_ComptePortefeuille As String
Public paramEffetCommerce_CompteRecouvreur As String
Public paramEffetCommerce_CompteTVA As String

Dim meCptMvt As typeCptMvt
Dim wGEch As typeGEch
Public Sub GFlux_Gen(lparam As typeGParam, lGope As typeGOpe, lGFlux_Nb As Integer, lGFlux() As typeGFlux)

ReDim lGFlux(3): lGFlux_Nb = 1

srvGFlux.recGFlux_Init lGFlux(1)

With lGFlux(1)                                   ' Engagement
    .Method = constAddNew
    .IdRéférence = lGope.IdRéférence
    .FluxSéquence = 1
    .Application = lparam.Application
    .Devise1 = lGope.Devise1
    .Montant1 = lGope.Montant1
    .Devise2 = lGope.Devise2
    .Montant2 = -(lGope.Montant2)
    .Taux = lGope.TauxMarge1
    .Nbj = lGope.PériodeNb
    .AmjEchéanceTrt = lGope.AmjEngagement
    .AmjDébut = lGope.AmjDébut
    .AmjFin = lGope.AmjFin
    .AmjOpération = lGope.AmjDébut
    .AmjValeur = lGope.AmjDébut
End With

Select Case Trim(lGope.Nature)
    Case "LCEsN": lGFlux(1).OpérationCode = "ECE1"
    Case "LCEsM": lGFlux(1).OpérationCode = "ECM1"
    Case "LCEnc": lGFlux(1).OpérationCode = "ECC1"
    Case "MCNE": lGFlux(1).OpérationCode = "ECI1"
    Case Else: Call lstErr_AddItem(frmEffetCommerce.lstErr, frmEffetCommerce.cmdContext, "? Nature : " & lGope.Nature)
End Select

End Sub

Public Sub GEch_Gen(lparam As typeGParam, lGope As typeGOpe, lGEch_Nb As Integer, lGEch() As typeGEch)

paramEffetCommerce_AmjPrésentation = DateElp_X(paramEffetCommerce_NbjPrésentation, lGope.AmjFin)
If paramEffetCommerce_AmjPrésentation < DSys Then paramEffetCommerce_AmjPrésentation = DSys

Select Case Trim(lGope.Nature)
    Case "LCEsN": Call GEch_GenLCEsc(lparam, lGope, lGEch_Nb, lGEch())
    Case "LCEsM": Call GEch_GenLCEsc(lparam, lGope, lGEch_Nb, lGEch())
    Case "LCEnc": Call GEch_GenLCEnc(lparam, lGope, lGEch_Nb, lGEch())
    Case "MCNE": Call GEch_GenMCNE(lparam, lGope, lGEch_Nb, lGEch())
    Case Else: Call lstErr_AddItem(frmEffetCommerce.lstErr, frmEffetCommerce.cmdContext, "? Nature : " & lGope.Nature)
End Select

End Sub

Public Sub GEch_GenLCEnc(lparam As typeGParam, lGope As typeGOpe, lGEch_Nb As Integer, lGEch() As typeGEch)
On Error GoTo Error_Handle

wGEch = lGEch(1)

    wGEch.Method = constAddNew
    wGEch.EchAMJ = lGope.AmjEngagement
    wGEch.EchHMS = "000000"
    wGEch.EchUsr = constAuto
    wGEch.FluxSéquence = 1
    
   If lGope.AmjEngagement = DSys Then
        wGEch.EchUsr = ""
        wGEch.Statut = "à"
        wGEch.StatutPlus = "C"
    End If
    
    With wGEch                                   ' Remise
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECRemise
    End With
    lGEch(wGEch.EchSéquence) = wGEch
    
    
    wGEch.EchUsr = constAuto
    wGEch.Statut = ""
    wGEch.StatutPlus = ""
    
   With wGEch                                   ' Présentation
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECPrésentation
        .EchAMJ = paramEffetCommerce_AmjPrésentation
    End With
    lGEch(wGEch.EchSéquence) = wGEch
     
    
    With wGEch                                   ' Echéance
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECEchéance
        .EchAMJ = lGope.AmjEchéance1
    End With
    lGEch(wGEch.EchSéquence) = wGEch
    
    With wGEch                                   ' 'constECRappro
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECRappro
        .EchAMJ = lGope.AmjEchéance1
    End With
    lGEch(wGEch.EchSéquence) = wGEch
    
    lGEch_Nb = wGEch.EchSéquence
    

Exit Sub
'---------------------------------------------------------
Error_Handle:
'---------------------------------------------------------

Call MsgBox("Erreur", vbCritical, "GEch_GenLCesc")
End Sub

Public Sub GEch_GenMCNE(lparam As typeGParam, lGope As typeGOpe, lGEch_Nb As Integer, lGEch() As typeGEch)
On Error GoTo Error_Handle

wGEch = lGEch(1)

    wGEch.Method = constAddNew
    wGEch.EchAMJ = lGope.AmjEngagement
    wGEch.EchHMS = "000000"
    wGEch.EchUsr = constAuto
    wGEch.FluxSéquence = 1
    
   If lGope.AmjEngagement = DSys Then
        wGEch.EchUsr = ""
        wGEch.Statut = "à"
        wGEch.StatutPlus = "C"
    End If
    
    With wGEch                                   ' Remise
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECRemise
    End With
    lGEch(wGEch.EchSéquence) = wGEch
    
    
    wGEch.EchUsr = constAuto
    wGEch.Statut = ""
    wGEch.StatutPlus = ""
    
     
    
    With wGEch                                   ' Echéance
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECEchéance
        .EchAMJ = lGope.AmjEchéance1
    End With
    lGEch(wGEch.EchSéquence) = wGEch
    
       lGEch_Nb = wGEch.EchSéquence
 
'End If


Exit Sub
'---------------------------------------------------------
Error_Handle:
'---------------------------------------------------------

Call MsgBox("Erreur", vbCritical, "gech_GenLCesc")
End Sub
Public Sub GEch_GenLCEsc(lparam As typeGParam, lGope As typeGOpe, lGEch_Nb As Integer, lGEch() As typeGEch)
On Error GoTo Error_Handle

wGEch = lGEch(1)

    wGEch.Method = constAddNew
    wGEch.EchAMJ = lGope.AmjEngagement
    wGEch.EchHMS = "000000"
    wGEch.EchUsr = constAuto
    wGEch.FluxSéquence = 1
    
   If lGope.AmjEngagement = DSys Then
        wGEch.EchUsr = ""
        wGEch.Statut = "à"
        wGEch.StatutPlus = "C"
    End If
    
    With wGEch                                   ' Remise
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECRemise
    End With
    lGEch(wGEch.EchSéquence) = wGEch
    
    
    wGEch.EchUsr = constAuto
    wGEch.Statut = ""
    wGEch.StatutPlus = ""
    
   With wGEch                                   ' Présentation
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECPrésentation
        .EchAMJ = paramEffetCommerce_AmjPrésentation
    End With
    lGEch(wGEch.EchSéquence) = wGEch
     
    
    With wGEch                                   ' Echéance
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECEchéance
        .EchAMJ = lGope.AmjEchéance1
    End With
    lGEch(wGEch.EchSéquence) = wGEch
    
    With wGEch                                   ' 'constECRappro
        .EchSéquence = wGEch.EchSéquence + 1
        .EchFct = constECRappro
        .EchAMJ = lGope.AmjEchéance1
    End With
    lGEch(wGEch.EchSéquence) = wGEch
    
    lGEch_Nb = wGEch.EchSéquence
    
'End If


Exit Sub
'---------------------------------------------------------
Error_Handle:
'---------------------------------------------------------

Call MsgBox("Erreur", vbCritical, "gech_GenLCesc")
End Sub

Public Function param_Init(lparam As typeGParam)
Dim V
param_Init = Null
recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = lparam.TableId
recElpTable.K1 = "Param"

recElpTable.K2 = "TauxMin"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_TauxMin = CDbl(Trim(recElpTable.Memo))
If Not IsNumeric(paramEffetCommerce_TauxMin) Then GoTo Num_Error

recElpTable.K2 = "TauxMinàConf"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_TauxMinàConfirmer = CDbl(Trim(recElpTable.Memo))
If Not IsNumeric(paramEffetCommerce_TauxMinàConfirmer) Then GoTo Num_Error

recElpTable.K2 = "TauxMax"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_TauxMax = CDbl(Trim(recElpTable.Memo))
If Not IsNumeric(paramEffetCommerce_TauxMax) Then GoTo Num_Error

recElpTable.K2 = "TauxMaxàConf"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_TauxMaxàConfirmer = CDbl(Trim(recElpTable.Memo))
If Not IsNumeric(paramEffetCommerce_TauxMaxàConfirmer) Then GoTo Num_Error

recElpTable.K2 = "TauxMargeMaj"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_TauxMargeMajoré = CDbl(Trim(recElpTable.Memo))
If Not IsNumeric(paramEffetCommerce_TauxMargeMajoré) Then GoTo Num_Error

recElpTable.K2 = "TauxMargeNA"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_TauxMargeNonAccepté = CDbl(Trim(recElpTable.Memo))
If Not IsNumeric(paramEffetCommerce_TauxMargeNonAccepté) Then GoTo Num_Error

recElpTable.K2 = "NbjRemMax"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_NbjRemMax = Trim(recElpTable.Memo)

paramEffetCommerce_AmjRemMax = DateElp_X(paramEffetCommerce_NbjRemMax, DSys)

recElpTable.K2 = "NbjPrésentat"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_NbjPrésentation = Trim(recElpTable.Memo)

recElpTable.K2 = "NbjEncMin"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_NbjEncMin = Trim(recElpTable.Memo)

recElpTable.K2 = "NbjEncCpt"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_NbjEncCpt = Trim(recElpTable.Memo)

recElpTable.K2 = "NbjEscMin"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_NbjEscMin = Trim(recElpTable.Memo)

recElpTable.K2 = "NbjEscVal"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_NbjEscVal = Trim(recElpTable.Memo)

recElpTable.K1 = "Compte"

recElpTable.K2 = "Agios"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_CompteAgios = Trim(recElpTable.Memo)
If Not IsNumeric(paramEffetCommerce_CompteAgios) Then GoTo Num_Error

recElpTable.K2 = "Compensateur"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_CompteCompensateur = Trim(recElpTable.Memo)
If Not IsNumeric(paramEffetCommerce_CompteCompensateur) Then GoTo Num_Error

recElpTable.K2 = "ComTaxable"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_CompteComTaxable = Trim(recElpTable.Memo)
If Not IsNumeric(paramEffetCommerce_CompteComTaxable) Then GoTo Num_Error

recElpTable.K2 = "Portefeuille"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_ComptePortefeuille = Trim(recElpTable.Memo)
If Not IsNumeric(paramEffetCommerce_ComptePortefeuille) Then GoTo Num_Error

recElpTable.K2 = "Recouvreur"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_CompteRecouvreur = Trim(recElpTable.Memo)
If Not IsNumeric(paramEffetCommerce_CompteRecouvreur) Then GoTo Num_Error

recElpTable.K2 = "TVA"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEffetCommerce_CompteTVA = Trim(recElpTable.Memo)
If Not IsNumeric(paramEffetCommerce_CompteTVA) Then GoTo Num_Error


Exit Function

Table_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Table", vbCritical, "frmEffetCommerce.Form_Init"
param_Init = V
Exit Function

Memo_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "frmEffetCommerce.Form_Init"
param_Init = V
Exit Function

Num_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "TfluxEspèces_Param_Init"
param_Init = V
End Function

Public Sub GMemo_Gen(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGEch As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
lGmemo_Nb = 0

recGMemo_Init lGmemo(1)
lGmemo(1).IdRéférence = lGEch.IdRéférence
lGmemo(1).MemoSéquencePlus = 0
lGmemo(1).Application = lGEch.Application
lGmemo(1).EchSéquence = lGEch.EchSéquence
lGmemo(1).FluxSéquence = lGEch.FluxSéquence



Select Case Trim(lGEch.EchFct)
    Case constECRemise
        Select Case Trim(lGFlux.OpérationCode)
            Case "ECC1": Call GMemo_Remise_ECC1(lparam, lGope, lGFlux, lGEch, lGmemo_Nb, lGmemo())
            Case "ECE1", "ECM1", "ECI1": Call GMemo_Remise_ECE1(lparam, lGope, lGFlux, lGEch, lGmemo_Nb, lGmemo())
       End Select
    Case constECPrésentation: Call GMemo_Présentation(lparam, lGope, lGFlux, lGEch, lGmemo_Nb, lGmemo())
    Case constECEchéance
        Select Case Trim(lGFlux.OpérationCode)
            Case "ECC1": Call GMemo_Echéance_ECC1(lparam, lGope, lGFlux, lGEch, lGmemo_Nb, lGmemo())
            Case "ECE1", "ECM1": Call GMemo_Echéance_ECE1(lparam, lGope, lGFlux, lGEch, lGmemo_Nb, lGmemo())
            Case "ECI1": Call GMemo_Echéance_ECE1(lparam, lGope, lGFlux, lGEch, lGmemo_Nb, lGmemo())
        End Select
    Case constECRappro: Call GMemo_Rapprochement(lparam, lGope, lGFlux, lGEch, lGmemo_Nb, lGmemo())
        
End Select

End Sub
Public Sub GMemo_Remise_ECC1(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGEch As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim wCur As Currency

lGmemo_Nb = 2

meCptMvt.Devise = lGFlux.Devise1
meCptMvt.Compte = lGope.EngagementCompte
meCptMvt.Mt = lGope.Montant1 - lGope.Montant2
meCptMvt.CodeOpération = mId$(lGFlux.OpérationCode, 1, 3) & "1"
meCptMvt.Service = lparam.Service
meCptMvt.AmjOpération = lGope.AmjEngagement
meCptMvt.AmjValeur = lGope.AmjEngagement
meCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "à": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(meCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
meCptMvt.Compte = paramEffetCommerce_ComptePortefeuille
meCptMvt.Mt = -lGope.Montant1
Call srvCptMvt_PutX(meCptMvt, lGmemo(2).MemoText)

wCur = lGope.Mensualité + lGope.Frais1

If wCur <> 0 Then
    lGmemo_Nb = lGmemo_Nb + 1
    lGmemo(lGmemo_Nb) = lGmemo(1)
    meCptMvt.Compte = paramEffetCommerce_CompteAgios
    meCptMvt.Mt = wCur
    Call srvCptMvt_PutX(meCptMvt, lGmemo(lGmemo_Nb).MemoText)
   
End If

If lGope.Frais2 <> 0 Then
    lGmemo_Nb = lGmemo_Nb + 1
    lGmemo(lGmemo_Nb) = lGmemo(1)
    meCptMvt.Compte = paramEffetCommerce_CompteComTaxable
    meCptMvt.Mt = lGope.Frais2
    Call srvCptMvt_PutX(meCptMvt, lGmemo(lGmemo_Nb).MemoText)
   
End If

If lGope.Frais3 <> 0 Then
    lGmemo_Nb = lGmemo_Nb + 1
    lGmemo(lGmemo_Nb) = lGmemo(1)
    meCptMvt.Compte = paramEffetCommerce_CompteTVA
    meCptMvt.Mt = lGope.Frais3
    Call srvCptMvt_PutX(meCptMvt, lGmemo(lGmemo_Nb).MemoText)
   
End If

End Sub

Public Sub GMemo_Remise_ECE1(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGEch As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim wCur As Currency

lGmemo_Nb = 2

meCptMvt.Devise = lGFlux.Devise1
meCptMvt.Compte = lGope.EchéanceCompte
meCptMvt.Mt = lGope.Montant1 - lGope.Montant2
meCptMvt.CodeOpération = mId$(lGFlux.OpérationCode, 1, 3) & "1"
meCptMvt.Service = lparam.Service
meCptMvt.AmjOpération = lGope.AmjEngagement
meCptMvt.AmjValeur = DateElp_X(paramEffetCommerce_NbjEscVal, lGope.AmjDébut)
meCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "à": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(meCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
meCptMvt.Compte = lGope.EngagementCompte
meCptMvt.Mt = -lGope.Montant1
Call srvCptMvt_PutX(meCptMvt, lGmemo(2).MemoText)

wCur = lGope.Mensualité + lGope.Frais1

If wCur <> 0 Then
    lGmemo_Nb = lGmemo_Nb + 1
    lGmemo(lGmemo_Nb) = lGmemo(1)
    meCptMvt.Compte = paramEffetCommerce_CompteAgios
    meCptMvt.Mt = wCur
    Call srvCptMvt_PutX(meCptMvt, lGmemo(lGmemo_Nb).MemoText)
   
End If

If lGope.Frais2 <> 0 Then
    lGmemo_Nb = lGmemo_Nb + 1
    lGmemo(lGmemo_Nb) = lGmemo(1)
    meCptMvt.Compte = paramEffetCommerce_CompteComTaxable
    meCptMvt.Mt = lGope.Frais2
    Call srvCptMvt_PutX(meCptMvt, lGmemo(lGmemo_Nb).MemoText)
   
End If

If lGope.Frais3 <> 0 Then
    lGmemo_Nb = lGmemo_Nb + 1
    lGmemo(lGmemo_Nb) = lGmemo(1)
    meCptMvt.Compte = paramEffetCommerce_CompteTVA
    meCptMvt.Mt = lGope.Frais3
    Call srvCptMvt_PutX(meCptMvt, lGmemo(lGmemo_Nb).MemoText)
   
End If

End Sub


Public Sub GMemo_Présentation(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGEch As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim wCur As Currency

lGmemo_Nb = 2

meCptMvt.Devise = lGFlux.Devise1
meCptMvt.Compte = paramEffetCommerce_ComptePortefeuille
meCptMvt.CodeOpération = mId$(lGFlux.OpérationCode, 1, 3) & "2"
meCptMvt.Service = lparam.Service
meCptMvt.Mt = lGope.Montant1
meCptMvt.AmjOpération = lGEch.EchAMJ
meCptMvt.AmjValeur = lGEch.EchAMJ
meCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "à": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(meCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
meCptMvt.Compte = lGope.EngagementCorrCompte
meCptMvt.Mt = -lGope.Montant1
Call srvCptMvt_PutX(meCptMvt, lGmemo(2).MemoText)


End Sub
Public Sub GMemo_Echéance_ECC1(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGEch As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim wCur As Currency

lGmemo_Nb = 2

meCptMvt.Devise = lGFlux.Devise1
meCptMvt.Compte = lGope.EchéanceCompte
meCptMvt.CodeOpération = mId$(lGFlux.OpérationCode, 1, 3) & "3"
meCptMvt.Service = lparam.Service
meCptMvt.Mt = lGope.Montant1 - lGope.Montant2
meCptMvt.AmjOpération = lGEch.EchAMJ
meCptMvt.AmjValeur = lGEch.EchAMJ
meCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "à": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(meCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
meCptMvt.Compte = lGope.EngagementCompte
meCptMvt.Mt = -meCptMvt.Mt
Call srvCptMvt_PutX(meCptMvt, lGmemo(2).MemoText)


End Sub

Public Sub GMemo_Echéance_ECE1(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGEch As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim wCur As Currency

lGmemo_Nb = 2

meCptMvt.Devise = lGFlux.Devise1
meCptMvt.Compte = lGope.EngagementCompte
meCptMvt.CodeOpération = mId$(lGFlux.OpérationCode, 1, 3) & "3"
meCptMvt.Service = lparam.Service
meCptMvt.Mt = lGope.Montant1
meCptMvt.AmjOpération = lGEch.EchAMJ
meCptMvt.AmjValeur = lGEch.EchAMJ
meCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "à": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(meCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
meCptMvt.Compte = paramEffetCommerce_ComptePortefeuille
meCptMvt.Mt = -meCptMvt.Mt
Call srvCptMvt_PutX(meCptMvt, lGmemo(2).MemoText)


End Sub

Public Sub GMemo_Rapprochement(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGEch As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim wCur As Currency

lGmemo_Nb = 2

meCptMvt.Devise = lGFlux.Devise1
meCptMvt.Compte = lGope.EngagementCorrCompte
meCptMvt.CodeOpération = mId$(lGFlux.OpérationCode, 1, 3) & "4"
meCptMvt.Service = lparam.Service
meCptMvt.Mt = lGope.Montant1
meCptMvt.AmjOpération = lGEch.EchAMJ
meCptMvt.AmjValeur = lGEch.EchAMJ
meCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "à": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(meCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
meCptMvt.Compte = lGope.EchéanceCorrCompte
meCptMvt.Mt = -meCptMvt.Mt
Call srvCptMvt_PutX(meCptMvt, lGmemo(2).MemoText)


End Sub

Public Function GMemo_Compta_Libellé(lGope As typeGOpe) As String
GMemo_Compta_Libellé = mId$(lGope.EngagementCompte, 1, 5) & " " & lGope.Nature & " " & dateImp10(lGope.AmjFin) & " " & lGope.Devise1 & " " & Trim(Format$(lGope.Montant1, "### ### ### ##0.00"))

End Function


