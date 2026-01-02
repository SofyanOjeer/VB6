Attribute VB_Name = "srvGSub_TC"
Option Explicit

Public Sub GMemo_Compta_CC01(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Call GMemo_ComptaHB_CC01(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())

lGmemo_Nb = 4

GSub_recCptMvt.Devise = lGFlux.Devise1
GSub_recCptMvt.Compte = lGope.EngagementCorrCompte
GSub_recCptMvt.CodeOpération = lGFlux.OpérationCode
GSub_recCptMvt.Service = lparam.Service
GSub_recCptMvt.Mt = -lGFlux.Montant1
GSub_recCptMvt.AmjOpération = lGope.AmjDébut
GSub_recCptMvt.AmjValeur = lGope.AmjDébut
GSub_recCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(3).MemoNature = constCompta
lGmemo(3).Statut = "@": lGmemo(3).StatutPlus = "C"
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(3).MemoText)

lGmemo(4) = lGmemo(3)
GSub_recCptMvt.Compte = paramCompteArbitrage
GSub_recCptMvt.Mt = -GSub_recCptMvt.Mt
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(4).MemoText)

End Sub
Public Function paramTC_Nature(lparam As typeGParam)
paramTC_Nature = Null

GSub_recNature.Method = "Seek="
GSub_recNature.Id = lparam.TableId
GSub_recNature.K1 = "Nature"
GSub_recNature.K2 = lparam.NatureCode
GSub_recNature.Err = tableElpTable_Read(GSub_recNature)
If GSub_recNature.Err <> 0 Then
''    Call MsgBox("GSub_Cpt: param_NatureTC", vbCritical, "Nature inconnue : " & lparam.NatureCode)
    GSub_recNature.K2 = ""
    GSub_recNature.Name = "? " & lparam.NatureCode
    GSub_recNature.Memo = String$(20, "0")
    paramTC_Nature = GSub_recNature.Name
End If

lparam.NatureLib = GSub_recNature.Name
lparam.NatureNbjValeur = Val(mId$(GSub_recNature.Memo, 1, 3))
lparam.NatureSens = mId$(GSub_recNature.Memo, 5, 1)
lparam.NatureDev1 = mId$(GSub_recNature.Memo, 7, 3)
lparam.NatureDev2 = mId$(GSub_recNature.Memo, 11, 3)

End Function
Public Function paramTC_Nature_TypeDeCompte(lparam As typeGParam, lGope As typeGOpe, lCptà As String)
Dim V As Variant, wK2A As String, wK2B As String, wK2C As String


If lCptà = "CptàR" Then
    wK2A = lGope.Devise1 & "___"
    wK2B = "___" & lGope.Devise2
    wK2C = "***___"
Else
    wK2A = lGope.Devise2 & "___"
    wK2B = "___" & lGope.Devise1
    wK2C = "***___"
End If


GSub_recNature.Method = "Seek="
GSub_recNature.Id = lparam.TableId
GSub_recNature.K1 = lCptà & "_" & lGope.Nature
GSub_recNature.K2 = wK2A
GSub_recNature.Err = tableElpTable_Read(GSub_recNature)
If GSub_recNature.Err <> 0 Then
    GSub_recNature.K2 = wK2B
    GSub_recNature.Err = tableElpTable_Read(GSub_recNature)
    If GSub_recNature.Err <> 0 Then
        GSub_recNature.K2 = wK2C
        GSub_recNature.Err = tableElpTable_Read(GSub_recNature)
        If GSub_recNature.Err <> 0 Then V = "? NatureTC_TypeDeCompte"
    End If
End If

If IsNull(V) Then GSub_recNature.Memo = String$(25, "0")

lparam.BiatypEngagement = mId$(GSub_recNature.Memo, 1, 3)
lparam.BiatypEngagementCorr = mId$(GSub_recNature.Memo, 5, 3)
lparam.BiatypEchéance = mId$(GSub_recNature.Memo, 9, 3)
lparam.Contrepartie = mId$(GSub_recNature.Memo, 13, 11)
lparam.BiatypReport = mId$(GSub_recNature.Memo, 25, 3)
paramTC_Nature_TypeDeCompte = V
End Function


Public Sub GMemo_Compta_CT01(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Call GMemo_ComptaHB_CT01(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())

lGmemo_Nb = 4

GSub_recCptMvt.Devise = lGFlux.Devise1
GSub_recCptMvt.Compte = lGope.EngagementCorrCompte
GSub_recCptMvt.CodeOpération = lGFlux.OpérationCode
GSub_recCptMvt.Service = lparam.Service
GSub_recCptMvt.Mt = -lGFlux.Montant1
GSub_recCptMvt.AmjOpération = lGope.AmjDébut
GSub_recCptMvt.AmjValeur = lGope.AmjDébut
GSub_recCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(3).MemoNature = constCompta
lGmemo(3).Statut = "@": lGmemo(3).StatutPlus = "C"
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(3).MemoText)

lGmemo(4) = lGmemo(3)
GSub_recCptMvt.Compte = paramCompteArbitrage
GSub_recCptMvt.Mt = -GSub_recCptMvt.Mt
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(4).MemoText)

End Sub
Public Sub GMemo_SwiftSnd_CC01(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim X As String

lGmemo_Nb = 1

lGmemo(1).MemoNature = constSwiftSnd
lGmemo(1).Statut = "@": lGmemo(1).StatutPlus = "S"
X = "32A: " & lGFlux.AmjValeur & " " & lGFlux.Devise1 & " " & Trim(Format$(lGFlux.Montant1, "############.##"))
lGmemo(1).MemoText = "MT210 : " & lGope.EngagementCorrSwiftN & X & lGope.EchéanceCorrSwiftL


End Sub


Public Sub GMemo_SwiftSnd_CT01(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim X As String

lGmemo_Nb = 1

lGmemo(1).MemoNature = constSwiftSnd
lGmemo(1).Statut = "@": lGmemo(1).StatutPlus = "S"
X = "32A: " & lGFlux.AmjValeur & " " & lGFlux.Devise1 & " " & Trim(Format$(lGFlux.Montant1, "############.##"))
lGmemo(1).MemoText = "MT210 : " & lGope.EngagementCorrSwiftN & X & lGope.EchéanceCorrSwiftL


End Sub


Public Sub GMemo_SwiftSnd_CC51(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim X As String

lGmemo_Nb = 1

lGmemo(1).MemoNature = constSwiftSnd
lGmemo(1).Statut = "@": lGmemo(1).StatutPlus = "S"
X = "32A: " & lGFlux.AmjValeur & " " & lGFlux.Devise1 & " " & Trim(Format$(lGFlux.Montant1, "############.##"))
lGmemo(1).MemoText = "MT210 : " & lGope.EchéanceCorrSwiftN & X & lGope.EchéanceCorrSwiftL

End Sub


Public Sub GMemo_SwiftSnd_CT51(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Dim X As String

lGmemo_Nb = 1

lGmemo(1).MemoNature = constSwiftSnd
lGmemo(1).Statut = "@": lGmemo(1).StatutPlus = "S"
X = "32A: " & lGFlux.AmjValeur & " " & lGFlux.Devise1 & " " & Trim(Format$(lGFlux.Montant1, "############.##"))
lGmemo(1).MemoText = "MT210 : " & lGope.EchéanceCorrSwiftN & X & lGope.EchéanceCorrSwiftL

End Sub



Public Sub GMemo_ComptaHB_CC51(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
lGmemo_Nb = 2
Call srvGSub_TC.paramTC_Nature_TypeDeCompte(lparam, lGope, "CptàL")
        
GSub_recCptMvt.Devise = lGFlux.Devise1
GSub_recCptMvt.Compte = lGope.EchéanceCompte
GSub_recCptMvt.CodeOpération = lGFlux.OpérationCode
GSub_recCptMvt.Service = lparam.Service
If Trim(lGech.EchFct) = constComptaHB Then
    GSub_recCptMvt.Mt = lGFlux.Montant1
    GSub_recCptMvt.AmjOpération = lGope.AmjEngagement
    GSub_recCptMvt.AmjValeur = lGope.AmjEngagement
Else
    GSub_recCptMvt.Mt = -lGFlux.Montant1
    GSub_recCptMvt.AmjOpération = lGFlux.AmjOpération
    GSub_recCptMvt.AmjValeur = lGFlux.AmjValeur
End If
GSub_recCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "@": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
GSub_recCptMvt.Compte = lparam.Contrepartie
GSub_recCptMvt.Mt = -GSub_recCptMvt.Mt
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(2).MemoText)

End Sub

Public Sub GMemo_ComptaHB_CT51(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
lGmemo_Nb = 2
Call srvGSub_TC.paramTC_Nature_TypeDeCompte(lparam, lGope, "CptàL")
        
GSub_recCptMvt.Devise = lGFlux.Devise1
GSub_recCptMvt.Compte = lGope.EchéanceCompte
GSub_recCptMvt.CodeOpération = lGFlux.OpérationCode
GSub_recCptMvt.Service = lparam.Service
If Trim(lGech.EchFct) = constComptaHB Then
    GSub_recCptMvt.Mt = lGFlux.Montant1
    GSub_recCptMvt.AmjOpération = lGope.AmjEngagement
    GSub_recCptMvt.AmjValeur = lGope.AmjEngagement
Else
    GSub_recCptMvt.Mt = -lGFlux.Montant1
    GSub_recCptMvt.AmjOpération = lGFlux.AmjOpération
    GSub_recCptMvt.AmjValeur = lGFlux.AmjValeur
End If
GSub_recCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "@": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
GSub_recCptMvt.Compte = lparam.Contrepartie
GSub_recCptMvt.Mt = -GSub_recCptMvt.Mt
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(2).MemoText)

End Sub

Public Sub GMemo_Compta_CC51(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Call GMemo_ComptaHB_CC51(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
lGmemo_Nb = 4
        
GSub_recCptMvt.Devise = lGFlux.Devise1
GSub_recCptMvt.Compte = lGope.EchéanceCorrCompte
GSub_recCptMvt.CodeOpération = lGFlux.OpérationCode
GSub_recCptMvt.Service = lparam.Service
GSub_recCptMvt.Mt = lGFlux.Montant1
GSub_recCptMvt.AmjOpération = lGope.AmjDébut
GSub_recCptMvt.AmjValeur = lGope.AmjDébut
GSub_recCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(3).MemoNature = constCompta
lGmemo(3).Statut = "@": lGmemo(3).StatutPlus = "C"
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(3).MemoText)

lGmemo(4) = lGmemo(3)
GSub_recCptMvt.Compte = paramCompteArbitrage
GSub_recCptMvt.Mt = -GSub_recCptMvt.Mt
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(4).MemoText)

End Sub
Public Sub GMemo_Compta_CT51(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
Call GMemo_ComptaHB_CT51(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
lGmemo_Nb = 4
        
GSub_recCptMvt.Devise = lGFlux.Devise1
GSub_recCptMvt.Compte = lGope.EchéanceCorrCompte
GSub_recCptMvt.CodeOpération = lGFlux.OpérationCode
GSub_recCptMvt.Service = lparam.Service
GSub_recCptMvt.Mt = lGFlux.Montant1
GSub_recCptMvt.AmjOpération = lGope.AmjDébut
GSub_recCptMvt.AmjValeur = lGope.AmjDébut
GSub_recCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(3).MemoNature = constCompta
lGmemo(3).Statut = "@": lGmemo(3).StatutPlus = "C"
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(3).MemoText)

lGmemo(4) = lGmemo(3)
GSub_recCptMvt.Compte = paramCompteArbitrage
GSub_recCptMvt.Mt = -GSub_recCptMvt.Mt
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(4).MemoText)

End Sub

Public Function GMemo_Compta_Libellé(lGope As typeGOpe) As String
GMemo_Compta_Libellé = lGope.Nature & " " & Trim(lGope.RéférenceInterne) & " " & mId$(lGope.EngagementCompte, 1, 5) & " " & lGope.Devise1 & " / " & lGope.Devise2

End Function



Public Sub GMemo_ComptaHB_CC01(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
lGmemo_Nb = 2
Call srvGSub_TC.paramTC_Nature_TypeDeCompte(lparam, lGope, "CptàR")

GSub_recCptMvt.Devise = lGFlux.Devise1
GSub_recCptMvt.Compte = lGope.EngagementCompte
GSub_recCptMvt.CodeOpération = lGFlux.OpérationCode
GSub_recCptMvt.Service = lparam.Service
If Trim(lGech.EchFct) = constComptaHB Then
    GSub_recCptMvt.Mt = -lGFlux.Montant1
    GSub_recCptMvt.AmjOpération = lGope.AmjEngagement
    GSub_recCptMvt.AmjValeur = lGope.AmjEngagement
Else
    GSub_recCptMvt.Mt = lGFlux.Montant1
    GSub_recCptMvt.AmjOpération = lGFlux.AmjOpération
    GSub_recCptMvt.AmjValeur = lGFlux.AmjValeur
End If
GSub_recCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "@": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
GSub_recCptMvt.Compte = lparam.Contrepartie
GSub_recCptMvt.Mt = -GSub_recCptMvt.Mt
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(2).MemoText)

End Sub
Public Sub GMemo_ComptaHB_CT01(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
lGmemo_Nb = 2
Call srvGSub_TC.paramTC_Nature_TypeDeCompte(lparam, lGope, "CptàR")

GSub_recCptMvt.Devise = lGFlux.Devise1
GSub_recCptMvt.Compte = lGope.EngagementCompte
GSub_recCptMvt.CodeOpération = lGFlux.OpérationCode
GSub_recCptMvt.Service = lparam.Service
If Trim(lGech.EchFct) = constComptaHB Then
    GSub_recCptMvt.Mt = -lGFlux.Montant1
    GSub_recCptMvt.AmjOpération = lGope.AmjEngagement
    GSub_recCptMvt.AmjValeur = lGope.AmjEngagement
Else
    GSub_recCptMvt.Mt = lGFlux.Montant1
    GSub_recCptMvt.AmjOpération = lGFlux.AmjOpération
    GSub_recCptMvt.AmjValeur = lGFlux.AmjValeur
End If
GSub_recCptMvt.Libellé = GMemo_Compta_Libellé(lGope)

lGmemo(1).MemoNature = constCompta
lGmemo(1).Statut = "@": lGmemo(1).StatutPlus = "C"
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(1).MemoText)

lGmemo(2) = lGmemo(1)
GSub_recCptMvt.Compte = lparam.Contrepartie
GSub_recCptMvt.Mt = -GSub_recCptMvt.Mt
Call srvCptMvt_PutX(GSub_recCptMvt, lGmemo(2).MemoText)

End Sub

Public Sub GMemo_Gen(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
lGmemo_Nb = 0

recGMemo_Init lGmemo(1)
lGmemo(1).IdRéférence = lGech.IdRéférence
lGmemo(1).MemoSéquencePlus = 0
lGmemo(1).Application = lGech.Application
lGmemo(1).EchSéquence = lGech.EchSéquence
lGmemo(1).FluxSéquence = lGech.FluxSéquence



Select Case Trim(lGech.EchFct)
    Case constComptaHB
        Select Case Trim(lGFlux.OpérationCode)
            Case "CC01": Call GMemo_ComptaHB_CC01(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CC51": Call GMemo_ComptaHB_CC51(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CT01": Call GMemo_ComptaHB_CT01(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CT51": Call GMemo_ComptaHB_CT51(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
        End Select
     Case constCompta
        Select Case Trim(lGFlux.OpérationCode)
            Case "CC01": Call GMemo_Compta_CC01(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CC51": Call GMemo_Compta_CC51(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CT01": Call GMemo_Compta_CT01(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CT51": Call GMemo_Compta_CT51(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
        End Select
      Case constSwiftSnd
        Select Case Trim(lGFlux.OpérationCode)
            Case "CC01": Call GMemo_SwiftSnd_CC01(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CC51": Call GMemo_SwiftSnd_CC51(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CT01": Call GMemo_SwiftSnd_CT01(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
            Case "CT51": Call GMemo_SwiftSnd_CT51(lparam, lGope, lGFlux, lGech, lGmemo_Nb, lGmemo())
        End Select
         
End Select

End Sub


