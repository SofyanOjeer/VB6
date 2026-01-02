Attribute VB_Name = "Guichet_Compta"
Option Explicit

Public G_CV1 As typeCV, G_CV2 As typeCV, G_CV3 As typeCV

Dim xConversion As String

Public G_strConversion1 As String, G_strConversion2 As String, G_strMontant1 As String, G_strMontant2 As String, G_strMontant3 As String

Public G_recGuichet As typeGuichet

Public G_arrCV030(7) As typeCpj030W0
Public G_arrCV030Nb As Integer

Dim Xcompte As typeCompte
Dim G_curMontant2 As Currency
Public G_CptInfo As typeCptInfo

Dim recGuichet_Compta As typeGuichet_Compta, xGuichet_Compta As typeGuichet_Compta
Dim recOpération_Compta As typeGuichet_Compta
Dim recCpj030 As typeCpj030W0

Public G_Total() As typeGuichet, G_Total_Index As Integer, G_Total_Nb As Integer

Dim mCptMvtLigne As Long

Dim minCptMvt As typeCptMvt, maxCptMvt As typeCptMvt
Dim G_ElpBuffer As typeElpBuffer
Public G_arrOppChq_Numéro() As String * 7, strOppChq_Numéro As String * 7
Public G_arrOppChq_Numéro_Nb As Integer

Public paramGuichetBillets_In As String * 11
Public paramGuichetBillets_Out As String * 11
Public paramGuichetConversion As String * 11
Public paramGuichetArbitrage As String * 11
Public paramGuichetAV_BME As String * 11
Public paramGuichetAjustement As String * 11
Public paramGuichetService As String * 3
Public paramGuichetAMJValeur As String * 8

Public paramGuichetCompensateur As String * 11
Public paramGuichetRecouvreur As String * 11
Public mGuichetAmjEchRecouvreur As String * 8
Public mGuichetAmjValeurCompensateur As String * 8

Public paramGuichetJValeurMin As String
Public paramGuichetJValeurMax As String
Public paramGuichetJValeurSR As String
Public paramGuichetJValeurHR As String
Public paramGuichetJValeurDom As String
Public paramGuichetJValeurPersonnel As String
Public paramGuichetJEchSRRecouvreur As String
Public paramGuichetJEchHRRecouvreur As String
Public paramGuichetJEchDomRecouvreur As String
Public paramGuichetJValeurSRCompensateur As String
Public paramGuichetJValeurHRCompensateur As String
Public paramGuichetJValeurDomCompensateur As String

Public paramGuichetAmjValeurMin As String * 8
Public paramGuichetAmjValeurMax As String * 8
Public paramGuichetAmjValeurSR As String * 8
Public paramGuichetAmjValeurHR As String * 8
Public paramGuichetAmjValeurDom As String * 8
Public paramGuichetAmjValeurPersonnel As String * 8
Public paramGuichetAmjEchSRRecouvreur As String * 8
Public paramGuichetAmjEchHRRecouvreur As String * 8
Public paramGuichetAmjEchDomRecouvreur As String * 8
Public paramGuichetAmjValeurSRCompensateur As String * 8
Public paramGuichetAmjValeurHRCompensateur As String * 8
Public paramGuichetAmjValeurDomCompensateur As String * 8

Public Sub OppChq_Load(Xcompte As String, lstOppChq As ListBox)
Dim I As Integer
ReDim arrOppChq(0): arrOppChqNbMax = (0)
arrOppChqNb = 0: arrOppChqIndex = 0
arrOppChqSuite = True

recOppChq_Init recOppChq
recOppChq.Method = "SnapP0"
recOppChq.Société = SocId$
recOppChq.Agence = SocAgence$
recOppChq.Racine = mId$(Xcompte, 1, 5)
arrOppChq(0) = recOppChq
arrOppChq(0).Numéro = "9999999"

Do Until Not arrOppChqSuite
    srvOppChq_Monitor recOppChq
    recOppChq = arrOppChq(arrOppChqNb)
    recOppChq.Method = "SnapP0+"
Loop

lstOppChq.Clear
G_arrOppChq_Numéro_Nb = arrOppChqNb
ReDim G_arrOppChq_Numéro(arrOppChqNb)

If arrOppChqNb > 0 Then
    For I = 1 To G_arrOppChq_Numéro_Nb
        G_arrOppChq_Numéro(I) = arrOppChq(I).Numéro
        lstOppChq.AddItem G_arrOppChq_Numéro(I)
    Next I
End If

End Sub

Public Function param_Init()
Dim V
param_Init = Null
recElpTable_Init recElpTable
recElpTable.Id = "Param"
recElpTable.K1 = "Guichet"
recElpTable.Method = "Seek="

recElpTable.K2 = "Billets_In"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetBillets_In = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGuichetBillets_In) Then GoTo Num_Error

recElpTable.K2 = "Billets_Out"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetBillets_Out = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGuichetBillets_Out) Then GoTo Num_Error


recElpTable.K2 = "Conversion"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetConversion = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGuichetConversion) Then GoTo Num_Error

recElpTable.K2 = "Arbitrage"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetArbitrage = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGuichetArbitrage) Then GoTo Num_Error

recElpTable.K2 = "A/V_BME"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetAV_BME = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGuichetAV_BME) Then GoTo Num_Error


recElpTable.K2 = "Ajustement"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetAjustement = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGuichetAjustement) Then GoTo Num_Error

recElpTable.K2 = "Service"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetService = mId$(recElpTable.Memo, 1, 3)
If Not IsNumeric(paramGuichetService) Then GoTo Num_Error


recElpTable.K2 = "Compensateur"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetCompensateur = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGuichetCompensateur) Then GoTo Num_Error

recElpTable.K2 = "JValeurSR_C"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurSRCompensateur = Trim(recElpTable.Memo)


recElpTable.K2 = "JValeurHR_C"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurHRCompensateur = Trim(recElpTable.Memo)

recElpTable.K2 = "JValeurDom_C"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurDomCompensateur = Trim(recElpTable.Memo)


recElpTable.K2 = "Recouvreur"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetRecouvreur = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGuichetRecouvreur) Then GoTo Num_Error

recElpTable.K2 = "JEchSR_R"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJEchSRRecouvreur = Trim(recElpTable.Memo)


recElpTable.K2 = "JEchHR_R"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJEchHRRecouvreur = Trim(recElpTable.Memo)

recElpTable.K2 = "JEchDom_R"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJEchDomRecouvreur = Trim(recElpTable.Memo)


recElpTable.K2 = "JValeurSR"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurSR = Trim(recElpTable.Memo)

recElpTable.K2 = "JValeurHR"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurHR = Trim(recElpTable.Memo)


recElpTable.K2 = "JValeurDom"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurDom = Trim(recElpTable.Memo)

recElpTable.K2 = "JValeurMax"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurMax = Trim(recElpTable.Memo)

recElpTable.K2 = "JValeurMin"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurMin = Trim(recElpTable.Memo)

recElpTable.K2 = "JValeurPerso"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGuichetJValeurPersonnel = Trim(recElpTable.Memo)
Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "GuichetEspèces_Compta_gen"
Exit Function

Num_Error:
param_Init = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "GuichetEspèces_Param_Init"
End Function



'---------------------------------------------------------
Public Sub Display(picCompta As PictureBox, lstErr As ListBox)
'---------------------------------------------------------
Dim X As String, I As Integer
Dim mCurrentX1 As Integer, mForeColor1 As Long
Dim mCurrentX2 As Integer, mForeColor2 As Long
Dim curTotal As Currency

DoEvents: picCompta.Cls
picCompta.ForeColor = libUsr.ForeColor
'piccompta.Line (0, 1200)-(9300, 1200)
'piccompta.Line (0, 2400)-(9300, 2400)
picCompta.CurrentY = 50
curTotal = 0
recCompteInit Xcompte
For I = 1 To G_arrCV030Nb
   If G_arrCV030(I).MONDEV <> 0 Then
        If I > 1 And G_arrCV030(I).Devise <> G_arrCV030(I - 1).Devise Then
            picCompta.ForeColor = libUsr.ForeColor
            picCompta.Line (0, picCompta.CurrentY)-(9300, picCompta.CurrentY)
            picCompta.CurrentY = picCompta.CurrentY + 100
        End If
        
        If G_arrCV030(I).SENECR = "D" Then
            curTotal = curTotal - G_arrCV030(I).MONDEV
            mCurrentX1 = 8000: mForeColor1 = errUsr.ForeColor
            mCurrentX2 = 9000: mForeColor2 = libUsr.ForeColor
        Else
            curTotal = curTotal + G_arrCV030(I).MONDEV
            mCurrentX2 = 8000: mForeColor2 = errUsr.ForeColor
            mCurrentX1 = 9000: mForeColor1 = libUsr.ForeColor
        End If
    
        picCompta.FontBold = False
        
        picCompta.ForeColor = libUsr.ForeColor
        
        picCompta.CurrentX = 50: picCompta.Print Format$(G_arrCV030(I).Devise, "000") & "." & Compte_Imp(G_arrCV030(I).Compte);
        
        picCompta.FontBold = False
        If Val(G_arrCV030(I).Compte) = 0 Then
 '           If GuichetAut.Saisir Then Call lstErr_AddItem(lstErr, piccompta, "? compte à préciser")
        Else
            picCompta.CurrentX = 1600
            
            Xcompte.Method = "SeekL1"
            Xcompte.Société = G_arrCV030(I).COSOC
            Xcompte.Agence = G_arrCV030(I).Agence
            Xcompte.Devise = Format$(Val(G_arrCV030(I).Devise), "000")
            Xcompte.Numéro = G_arrCV030(I).Compte
            If IsNull(srvCompteFind(Xcompte)) Then
                 If Xcompte.Situation = "A" Then
                    picCompta.ForeColor = errUsr.ForeColor
                    picCompta.Print Trim(Xcompte.Intitulé) & "Annulé";
                    If Xcompte.Situation = "A" Or Xcompte.Situation = "E" Then Call lstErr_AddItem(lstErr, picCompta, "? compte  annulé")
                Else
                    'If Xcompte.TypeGA = "A" And Xcompte.BiaTyp <> "001" Then
                    '    picCompta.ForeColor = errUsr.ForeColor
                    '    Call lstErr_AddItem(lstErr, picCompta, "? type  = 001 ")
                    'End If
                    If Xcompte.TypeGA = "G" And Xcompte.Numéro >= 90000000 Then
                        picCompta.ForeColor = errUsr.ForeColor
                        Call lstErr_AddItem(lstErr, picCompta, "? compte Hors-Bilan ")
                    End If
                    picCompta.Print Trim(Xcompte.Intitulé);
               End If
            Else
                picCompta.ForeColor = errUsr.ForeColor
                picCompta.Print "???????";
                Call lstErr_AddItem(lstErr, picCompta, "? compte inconnu")
            End If
        End If
        
        If picCompta.CurrentX < 5500 Then picCompta.CurrentX = 5500
        picCompta.ForeColor = warnUsrColor
        picCompta.Print dateImp(G_arrCV030(I).AMJVAL);
        
        picCompta.FontBold = True
        picCompta.ForeColor = mForeColor1
        X = Format$(G_arrCV030(I).MONDEV, "### ### ### ### ##0.00")
        
        picCompta.CurrentX = mCurrentX1 - picCompta.TextWidth(X)
        picCompta.Print X;
        
        picCompta.CurrentY = picCompta.CurrentY + 270
        picCompta.CurrentX = 1600
        picCompta.ForeColor = warnUsrColor
        picCompta.Print G_arrCV030(I).LIBELE;
        
       picCompta.CurrentY = picCompta.CurrentY + 270
    End If
Next I
picCompta.ForeColor = libUsr.ForeColor
picCompta.Line (0, picCompta.CurrentY)-(9300, picCompta.CurrentY)
If curTotal <> 0 Then Call lstErr_AddItem(lstErr, picCompta, "? pièce non équilibrée")
End Sub

Public Sub Init()

If G_CV1.EuroIn And G_CV2.DeviseIso = "EUR" Then G_CV2.EuroIn = True
If G_CV2.EuroIn And G_CV1.DeviseIso = "EUR" Then G_CV1.EuroIn = True

recCpj030W0_Init G_arrCV030(0)
G_arrCV030(0).Method = "AddNew"
G_arrCV030(0).COSOC = G_recGuichet.Société
G_arrCV030(0).Agence = G_recGuichet.Agence
G_arrCV030(0).AGEMET = G_recGuichet.Agence
G_arrCV030(0).BIACOP = G_recGuichet.CodeOpération
G_arrCV030(0).SERVIC = G_recGuichet.CptMvtService
G_arrCV030(0).AMJSAI = G_recGuichet.SaisieAmj
G_arrCV030(0).AMJVAL = G_recGuichet.AmjValeur
G_arrCV030(0).AMJOPE = G_recGuichet.AmjOpération
G_arrCV030(0).NOMOP = G_recGuichet.SaisieUsr
G_arrCV030(0).JJCPLT = "0"

G_curMontant2 = G_recGuichet.Montant + G_recGuichet.MontantAjustement
G_strMontant1 = G_CV1.DeviseIso & " " & Trim(Format$(G_recGuichet.MontantEspèces, "##### ### ##0.00"))
G_strMontant2 = G_CV2.DeviseIso & " " & Trim(Format$(G_curMontant2, "##### ### ##0.00"))
G_strMontant3 = G_CV3.DeviseIso & " " & Trim(Format$(G_recGuichet.MontantEuro, "##### ### ##0.00"))

Select Case G_recGuichet.Conversion
    Case "C":  G_strConversion1 = "Conversion ": G_strConversion2 = "Conversion "
    Case "B":
            If G_CV1.EuroIn Then
                    G_strConversion1 = "Conversion ": G_strConversion2 = "Arbitrage "
            Else
                    G_strConversion2 = "Conversion ": G_strConversion1 = "Arbitrage "
            End If
    Case "A":  G_strConversion1 = "Arbitrage ": G_strConversion2 = "Arbitrage "
End Select

End Sub


Public Sub G001()
G_arrCV030Nb = 2

G_arrCV030(1).LIBELE = G_recGuichet.Libellé

If G_CV2.EuroIn Then
    G_arrCV030(2).Compte = paramGuichetBillets_In
Else
    G_arrCV030(2).Compte = paramGuichetBillets_Out
End If
G_arrCV030(2).Devise = G_recGuichet.DeviseEspèces
G_arrCV030(2).MONDEV = G_recGuichet.MontantEspèces
G_arrCV030(2).LIBELE = G_recGuichet.ContrepartieLibellé

If G_recGuichet.MontantAjustement <> 0 Then
    G_arrCV030Nb = 3
    G_arrCV030(3) = G_arrCV030(1)
    G_arrCV030(3).Compte = paramGuichetAjustement
    G_arrCV030(3).SENECR = IIf(G_recGuichet.MontantAjustement, "C", "D")
    G_arrCV030(3).MONDEV = Abs(G_recGuichet.MontantAjustement)
End If

End Sub

Public Sub G010()
G_arrCV030Nb = 2

G_arrCV030(1).LIBELE = G_recGuichet.Libellé

G_arrCV030(2).Compte = G_recGuichet.ContrepartieCompte
G_arrCV030(2).Devise = G_recGuichet.DeviseEspèces
G_arrCV030(2).MONDEV = G_recGuichet.MontantEspèces
G_arrCV030(2).LIBELE = G_recGuichet.ContrepartieLibellé
G_arrCV030(2).AMJVAL = mId$(G_recGuichet.Complément3, 1, 8)
End Sub

Public Sub G008()
G_arrCV030Nb = 5

G_arrCV030(1).LIBELE = G_recGuichet.Libellé
If G_CV2.EuroIn Then
    G_arrCV030(2).Compte = paramGuichetConversion
Else
    G_arrCV030(2).Compte = paramGuichetArbitrage
End If

G_arrCV030(2).Devise = G_recGuichet.Devise
G_arrCV030(2).MONDEV = G_curMontant2
G_arrCV030(2).LIBELE = G_CV1.DeviseN & " " & G_strConversion1 & G_strMontant1

G_arrCV030(3).MONDEV = 0
If G_recGuichet.MontantAjustement <> 0 Then
    G_arrCV030(3) = G_arrCV030(1)
    G_arrCV030(3).MONDEV = Abs(G_recGuichet.MontantAjustement)
    G_arrCV030(3).Compte = paramGuichetAjustement
    G_arrCV030(3).SENECR = IIf(G_recGuichet.MontantAjustement, "C", "D")
End If

If G_CV1.EuroIn Then
    G_arrCV030(4).Compte = paramGuichetConversion
Else
    G_arrCV030(4).Compte = paramGuichetArbitrage
End If

G_arrCV030(4).Devise = G_recGuichet.DeviseEspèces
G_arrCV030(4).SENECR = G_arrCV030(1).SENECR
G_arrCV030(4).MONDEV = G_recGuichet.MontantEspèces
G_arrCV030(4).LIBELE = G_CV2.DeviseN & " " & G_strConversion2 & G_strMontant2

G_arrCV030(5).Compte = G_recGuichet.ContrepartieCompte
G_arrCV030(5).Devise = G_recGuichet.DeviseEspèces
G_arrCV030(5).SENECR = G_arrCV030(2).SENECR
G_arrCV030(5).MONDEV = G_recGuichet.MontantEspèces
G_arrCV030(5).LIBELE = G_recGuichet.ContrepartieLibellé


If G_recGuichet.Conversion = "B" Then
    G_arrCV030Nb = 7
    G_arrCV030(2).LIBELE = G_CV3.DeviseN & " " & G_strConversion2 & G_strMontant3
    G_arrCV030(4).LIBELE = G_CV3.DeviseN & " " & G_strConversion1 & G_strMontant3
    
    G_arrCV030(6).Devise = G_CV3.DeviseN
    G_arrCV030(6).Compte = G_arrCV030(2).Compte
    G_arrCV030(6).SENECR = G_arrCV030(1).SENECR
    G_arrCV030(6).MONDEV = G_recGuichet.MontantEuro
    G_arrCV030(6).LIBELE = G_CV2.DeviseN & " " & G_strConversion2 & G_strMontant2
    
    G_arrCV030(7).Devise = G_CV3.DeviseN
    G_arrCV030(7).Compte = G_arrCV030(4).Compte
    G_arrCV030(7).SENECR = G_arrCV030(2).SENECR
    G_arrCV030(7).MONDEV = G_recGuichet.MontantEuro
    G_arrCV030(7).LIBELE = G_CV1.DeviseN & " " & G_strConversion1 & G_strMontant1

End If


End Sub
Public Sub G006()

G_arrCV030Nb = 5

G_arrCV030(1).LIBELE = G_recGuichet.Libellé
If G_CV2.EuroIn Then
    G_arrCV030(2).Compte = paramGuichetConversion
Else
    If G_CV1.EuroIn Then
        G_arrCV030(2).Compte = paramGuichetArbitrage
    Else
        G_arrCV030(2).Compte = paramGuichetAV_BME
    End If
End If

G_arrCV030(2).Devise = G_recGuichet.Devise
G_arrCV030(2).MONDEV = G_curMontant2
G_arrCV030(2).LIBELE = G_CV1.DeviseN & " " & G_strConversion1 & G_strMontant1

G_arrCV030(3).MONDEV = 0
If G_recGuichet.MontantAjustement <> 0 Then
    G_arrCV030(3) = G_arrCV030(1)
    G_arrCV030(3).MONDEV = Abs(G_recGuichet.MontantAjustement)
    G_arrCV030(3).Compte = paramGuichetAjustement
    G_arrCV030(3).SENECR = IIf(G_recGuichet.MontantAjustement, "C", "D")
End If

If G_CV1.EuroIn Then
    G_arrCV030(4).Compte = paramGuichetConversion
    G_arrCV030(5).Compte = paramGuichetBillets_In
Else
    G_arrCV030(4).Compte = paramGuichetAV_BME
    G_arrCV030(5).Compte = paramGuichetBillets_Out
End If

G_arrCV030(4).Devise = G_recGuichet.DeviseEspèces
G_arrCV030(4).SENECR = G_arrCV030(1).SENECR
G_arrCV030(4).MONDEV = G_recGuichet.MontantEspèces
G_arrCV030(4).LIBELE = G_CV2.DeviseN & " " & G_strConversion2 & G_strMontant2

G_arrCV030(5).Devise = G_recGuichet.DeviseEspèces
G_arrCV030(5).SENECR = G_arrCV030(2).SENECR
G_arrCV030(5).MONDEV = G_recGuichet.MontantEspèces
G_arrCV030(5).LIBELE = G_recGuichet.ContrepartieLibellé
If G_recGuichet.Conversion = "B" Then
    G_arrCV030Nb = 7
    G_arrCV030(2).LIBELE = G_CV3.DeviseN & " " & G_strConversion2 & G_strMontant3
    G_arrCV030(4).LIBELE = G_CV3.DeviseN & " " & G_strConversion1 & G_strMontant3
    
    G_arrCV030(6).Devise = G_CV3.DeviseN
    G_arrCV030(6).Compte = G_arrCV030(2).Compte
    G_arrCV030(6).SENECR = G_arrCV030(1).SENECR
    G_arrCV030(6).MONDEV = G_recGuichet.MontantEuro
    G_arrCV030(6).LIBELE = G_CV2.DeviseN & " " & G_strConversion2 & G_strMontant2
    
    G_arrCV030(7).Devise = G_CV3.DeviseN
    G_arrCV030(7).Compte = G_arrCV030(4).Compte
    G_arrCV030(7).SENECR = G_arrCV030(2).SENECR
    G_arrCV030(7).MONDEV = G_recGuichet.MontantEuro
    G_arrCV030(7).LIBELE = G_CV1.DeviseN & " " & G_strConversion1 & G_strMontant1

End If

End Sub
Public Function LibelléOptRetrait(X As String) As String
Select Case G_recGuichet.chkChèque
    Case "1": LibelléOptRetrait = X & " chèque " & Format$(G_recGuichet.NoChèque, "0000000 ") & G_recGuichet.Identité
    Case "2": LibelléOptRetrait = X & " chèque guichet " & Format$(G_recGuichet.NoChèque, "0000000 ") & G_recGuichet.Identité
    Case "3": LibelléOptRetrait = X & Trim(G_recGuichet.Identité) & " / " & G_recGuichet.Complément1
    Case "4": LibelléOptRetrait = X & Trim(G_recGuichet.Identité) & " / " & G_recGuichet.Complément1
    Case "5": LibelléOptRetrait = G_recGuichet.Complément1
End Select

End Function
Public Sub Libellé()
Dim X As String
Select Case Trim(G_recGuichet.CodeOpération)
    Case "G001":    G_recGuichet.ContrepartieLibellé = "Versement espèces " & " / " & G_CV2.DeviseIso & " " & Compte_Imp(G_recGuichet.Compte)
                    G_recGuichet.Libellé = "Versement espèces " & G_recGuichet.Identité
    Case "G006":    G_recGuichet.ContrepartieLibellé = "Versement espèces " & G_strMontant2 & " / " & G_CV2.DeviseIso & " " & Compte_Imp(G_recGuichet.Compte)
                    G_recGuichet.Libellé = "Contre-valeur versement " & G_strMontant1 & G_recGuichet.Identité
    Case "G007":    G_recGuichet.ContrepartieLibellé = "Change " & G_strMontant2 & "/ " & G_CV2.DeviseIso & " " & Compte_Imp(G_recGuichet.Compte)
                    G_recGuichet.Libellé = "Change " & G_recGuichet.Identité
    Case "G002": X = "Retrait espèces "
                    G_recGuichet.ContrepartieLibellé = X & " / " & G_CV2.DeviseIso & " " & Compte_Imp(G_recGuichet.Compte)
                    G_recGuichet.Libellé = LibelléOptRetrait(X)
                    
    Case "G005":    G_recGuichet.ContrepartieLibellé = "Délivrance de " & G_strMontant1 & " / " & G_CV2.DeviseIso & " " & Compte_Imp(G_recGuichet.Compte)
                    G_recGuichet.Libellé = LibelléOptRetrait("Délivrance de " & G_strMontant1 & " ")
    Case "G008":    G_recGuichet.ContrepartieLibellé = "Arbitrage " & G_strMontant1 & " / " & G_strMontant2
                    G_recGuichet.Libellé = "Arbitrage " & G_strMontant2 & " / " & G_strMontant1
    Case "G010":    LibelléChèque "BIA": G_recGuichet.ContrepartieLibellé = "chèque " & Format$(G_recGuichet.NoChèque, "0000000 ")
    Case "G011":    LibelléChèque "sur rayon"
    Case "G012":    LibelléChèque "hors rayon"
    Case "G013":    LibelléChèque "DOM TOM"
    Case "G014":    LibelléChèque "sur l'étranger"

End Select
End Sub




Public Sub Gen()
Dim X11 As String * 11

G_arrCV030(0).Devise = G_recGuichet.Devise
G_arrCV030(0).NUMPIE = G_recGuichet.CptMvtPièce
G_arrCV030(0).NOLIGN = G_recGuichet.CptMvtLigne

G_arrCV030(1) = G_arrCV030(0)
G_arrCV030(2) = G_arrCV030(0)
G_arrCV030(3) = G_arrCV030(0)
G_arrCV030(4) = G_arrCV030(0)
G_arrCV030(5) = G_arrCV030(0)
G_arrCV030(6) = G_arrCV030(0)
G_arrCV030(7) = G_arrCV030(0)

G_arrCV030(1).Compte = G_recGuichet.Compte
G_arrCV030(1).MONDEV = G_recGuichet.Montant
G_arrCV030(1).SENECR = G_recGuichet.Sens
G_arrCV030(2).NOLIGN = 1
G_arrCV030(2).SENECR = IIf(G_arrCV030(1).SENECR = "C", "D", "C")
G_arrCV030(4).NUMPIE = G_recGuichet.CptMvtPièceEspèces
G_arrCV030(4).NOLIGN = G_recGuichet.CptMvtLigneEspèces
G_arrCV030(5).NUMPIE = G_recGuichet.CptMvtPièceEspèces
G_arrCV030(5).NOLIGN = G_recGuichet.CptMvtLigneEspèces
G_arrCV030(6).NUMPIE = G_recGuichet.CptMvtPièceEspèces + 9000000
G_arrCV030(7).NUMPIE = G_recGuichet.CptMvtPièceEspèces + 9000000

Select Case Trim(G_recGuichet.CodeOpération)

    Case "G001": G001
    Case "G002": G001
    Case "G005": G006
    Case "G006": G006
    Case "G007": G006
    Case "G008": G008
    Case "G010", "G011", "G012", "G013", "G014":
            If G_recGuichet.Devise = G_recGuichet.DeviseEspèces Then
                G010
            Else
                G010Conversion
            End If
End Select

End Sub


Public Sub CV_Reset(currentAction As String)

CV_Init G_CV1
G_CV1.Normal = "N"

G_CV2 = G_CV1
G_CV2.Normal = "C"
G_CV3 = CV_Euro

Select Case currentAction
    Case constRetrait:      G_CV1.AchatVente = "A": G_CV2.AchatVente = "V"
    Case constArbitrage:    G_CV1.AchatVente = "A": G_CV2.AchatVente = "V"
    Case constVersement:    G_CV1.AchatVente = "V": G_CV2.AchatVente = "A"
    Case constChange:       G_CV1.AchatVente = "V": G_CV2.AchatVente = "A"

End Select

End Sub




Public Sub ValidationDemande(recGuichet As typeGuichet)
Dim I As Integer, Msg As String

MvtCpt_Init
recGuichet.Method = "SnapL3"
recGuichet.CptMvtPièce = 0
recGuichet.CptMvtLigne = 0
arrGuichet(0) = recGuichet
arrGuichet(0).CptMvtPièce = 9999999
arrGuichet(0).CptMvtLigne = 9999999

arrGuichetNb = 0: arrGuichetIndex = 0
arrGuichetSuite = True
Do Until Not arrGuichetSuite
    srvGuichet_Monitor recGuichet
    recGuichet = arrGuichet(arrGuichetNb)
    recGuichet.Method = "SnapL3+"
Loop
Msg = "000000000000"

For I = 1 To arrGuichetNb
        G_recGuichet = arrGuichet(I)
        If Trim(G_recGuichet.ValidationUsr) = "" Then
            MvtCpt_AddNew
            '$JPL 1999-09-14 MvtCpt_Opération
        End If
Next I

prtCompta.mJournal = recGuichet.Journal
prtCompta_Monitor Msg, constDemandeDeValidation, conststrGuichet_Compta
End Sub

Public Sub Validation(lGuichet As typeGuichet)
Dim I As Integer, Msg As String, trimusrId As String, xNumlot As String
Dim recGuichet As typeGuichet

MvtCpt_Init
recGuichet = lGuichet
trimusrId = Trim(usrId)
xNumlot = recGuichet.ComptaUsr

recGuichet.Method = "SnapL3"
recGuichet.CptMvtPièce = 0
recGuichet.CptMvtLigne = 0
arrGuichet(0) = recGuichet
arrGuichet(0).CptMvtPièce = 9999999
arrGuichet(0).CptMvtLigne = 9999999

arrGuichetNb = 0: arrGuichetIndex = 0
arrGuichetSuite = True
Do Until Not arrGuichetSuite
    srvGuichet_Monitor recGuichet
    recGuichet = arrGuichet(arrGuichetNb)
    recGuichet.Method = "SnapL3+"
Loop
Msg = "000000000000"

For I = 1 To arrGuichetNb
    G_recGuichet = arrGuichet(I)
    If Trim(G_recGuichet.ValidationUsr) = trimusrId Then
        MvtCpt_AddNew
    End If
Next I

mCptMvtLigne = 0
recGuichet = lGuichet
Validation_Snd recGuichet

recGuichet = lGuichet
recGuichet.CptMvtLigne = mCptMvtLigne

recGuichet.Method = "Compta"
If Not IsNull(srvGuichet_Monitor(recGuichet)) Then Call MsgBox("Erreur informatique : ANNULER LA VALIDATION", vbCritical, "Validation comptable des opérations de guichet")
If recGuichet.CptMvtLigne <> 0 Or recGuichet.Montant <> 0 Then Call MsgBox("anomalie en nombre ou en montant : ANNULER LA VALIDATION", vbCritical, "Validation comptable des opérations de guichet")

Msg = "000000000000"
prtCompta.mJournal = recGuichet.Journal
prtCompta_Monitor Msg, "VALIDATION Lot N°" & Format$(xNumlot, "### ### ###"), conststrGuichet_Compta
End Sub

'---------------------------------------------------------
Public Sub Validation_Snd(lGuichet As typeGuichet)
'---------------------------------------------------------
Dim iReturn As Integer, mCpj030 As typeCpj030W0

recGuichet_Compta.Method = "MoveFirst"
        recCpj030W0_Init recCpj030
        recCpj030.obj = "SRVCPJ030H"
        recCpj030.Method = "AddNew"
        recCpj030.NUMLOT = Val(mId$(lGuichet.ComptaUsr, 7, 4)) '!!!! NUMLOT : 4 chiffres
        recCpj030.CTLNOM = usrId
        recCpj030.IMPAMJ = lGuichet.ValidationAMJ
        recCpj030.IMPHMS = lGuichet.ValidationHMS
        recCpj030.CTLAMJ = lGuichet.ComptaAMJ
        recCpj030.CTLHMS = lGuichet.ComptaHMS
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
mCptMvtLigne = 100000

recGuichet_Compta_Init recGuichet_Compta
recOpération_Compta = recGuichet_Compta
recOpération_Compta.Société = "000"
recOpération_Compta.Agence = "000"
recOpération_Compta.Method = "AddNew"

recGuichet_Init G_recGuichet
G_recGuichet.Method = "SnapL1"
G_recGuichet.Société = SocId$
G_recGuichet.Agence = SocAgence$

End Sub


Public Sub MvtCpt_AddNew()
Dim K As Integer

G_CV1.DeviseN = G_recGuichet.DeviseEspèces: CV_AttributN G_CV1
G_CV2.DeviseN = G_recGuichet.Devise: CV_AttributN G_CV2

Guichet_Compta.Init
recGuichet_Compta.chkCompte = "0"
recGuichet_Compta.chkSolde = "0"
recGuichet_Compta.chkAmjValeur = "0"

Guichet_Compta.Gen

For K = 1 To G_arrCV030Nb
    If G_arrCV030(K).MONDEV <> 0 Then
        recGuichet_Compta.Société = G_recGuichet.Société
        recGuichet_Compta.Agence = G_recGuichet.Agence
        recGuichet_Compta.SaisieUsr = G_recGuichet.SaisieUsr
        recGuichet_Compta.Devise = G_arrCV030(K).Devise
        recGuichet_Compta.CodeOpération = G_arrCV030(K).BIACOP
        recGuichet_Compta.CptMvtPièce = G_arrCV030(K).NUMPIE
        recGuichet_Compta.CptMvtLigne = G_arrCV030(K).NOLIGN
        If K > 1 Then '2 Then
            mCptMvtLigne = mCptMvtLigne + 100000
            recGuichet_Compta.CptMvtLigne = recGuichet_Compta.CptMvtLigne + mCptMvtLigne
        End If
        recGuichet_Compta.Service = G_arrCV030(K).SERVIC
        recGuichet_Compta.Compte = G_arrCV030(K).Compte
        recGuichet_Compta.Montant = G_arrCV030(K).MONDEV
        recGuichet_Compta.Sens = G_arrCV030(K).SENECR
        recGuichet_Compta.AmjOpération = G_recGuichet.AmjOpération
        recGuichet_Compta.AmjValeur = G_arrCV030(K).AMJVAL
        recGuichet_Compta.Libellé = G_arrCV030(K).LIBELE
        recGuichet_Compta.SaisieAmj = G_recGuichet.SaisieAmj
        recGuichet_Compta.Référence = G_recGuichet.Référence
        recGuichet_Compta.chkChèque = "0"
        
        recGuichet_Compta.Method = "AddNew"
        
        If G_recGuichet.Journal = constCaisse Then
            Select Case recGuichet_Compta.Compte
                Case paramGuichetBillets_In
                                recGuichet_Compta.CptMvtLigne = 999999999
                                MvtCpt_Contrepartie
                Case paramGuichetBillets_Out
                                recGuichet_Compta.CptMvtLigne = 999999997
                                MvtCpt_Contrepartie
            End Select
        Else
             Select Case recGuichet_Compta.Compte
                Case paramGuichetCompensateur
                                recGuichet_Compta.CptMvtLigne = Val(mId$(recGuichet_Compta.AmjValeur, 5, 4)) & "9999"
                                MvtCpt_Contrepartie_Chèque
                Case paramGuichetRecouvreur
                                recGuichet_Compta.CptMvtLigne = Val(mId$(recGuichet_Compta.AmjValeur, 5, 4)) & "8888"
                                MvtCpt_Contrepartie_Chèque
            End Select
       End If
        
        
        recGuichet_Compta.ValidationAMJ = G_recGuichet.ValidationAMJ
        recGuichet_Compta.ValidationHMS = G_recGuichet.ValidationHMS
        recGuichet_Compta.ValidationUsr = G_recGuichet.ValidationUsr
        recGuichet_Compta.ComptaAMJ = G_recGuichet.ComptaAMJ
        recGuichet_Compta.ComptaHMS = G_recGuichet.ComptaHMS
        recGuichet_Compta.ComptaUsr = G_recGuichet.ComptaUsr

        dbGuichet_Compta_Update recGuichet_Compta
    End If
Next K

End Sub

Public Sub LotComptabilisé_Print(xNumlot As String)
recCptMvtInit minCptMvt
minCptMvt.obj = "SRVECRITG"
minCptMvt.Method = "SnapK0"
minCptMvt.Société = SocId$
minCptMvt.Agence = SocAgence$
minCptMvt.Devise = "000"
minCptMvt.Lot = Val(xNumlot)
minCptMvt.Pièce = 0
minCptMvt.Ligne = 0

maxCptMvt = minCptMvt
maxCptMvt.Devise = "999"
maxCptMvt.Pièce = 999999999
maxCptMvt.Ligne = 9999
Call srvCptMvt_ElpBuffer(minCptMvt, maxCptMvt, G_ElpBuffer)
prtCompta.mJournal = ""
If G_ElpBuffer.Seq > 0 Then prtCompta_Monitor G_ElpBuffer.Id, "Lot N°" & Format$(xNumlot, "### ### ###"), conststrGuichet_Comptabilisé

End Sub


Public Sub MvtCpt_Opération()
recOpération_Compta.Agence = G_recGuichet.DeviseEspèces '$$$$$$$$$$$$$$$$$$$$$$$$$$$

recOpération_Compta.SaisieUsr = G_recGuichet.SaisieUsr
recOpération_Compta.Devise = G_recGuichet.Devise
recOpération_Compta.CodeOpération = G_recGuichet.CodeOpération
recOpération_Compta.CptMvtPièce = G_recGuichet.CptMvtPièceEspèces
recOpération_Compta.CptMvtLigne = G_recGuichet.CptMvtLigneEspèces
recOpération_Compta.Service = G_recGuichet.CptMvtService
recOpération_Compta.Compte = G_recGuichet.Compte
recOpération_Compta.Montant = G_recGuichet.MontantEspèces
recOpération_Compta.Sens = G_recGuichet.Sens
recOpération_Compta.AmjOpération = G_recGuichet.AmjOpération
recOpération_Compta.AmjValeur = G_recGuichet.AmjValeur
recOpération_Compta.Libellé = G_recGuichet.Libellé
recOpération_Compta.SaisieAmj = G_recGuichet.SaisieAmj
recOpération_Compta.chkCompte = G_recGuichet.chkCompte
recOpération_Compta.chkSolde = G_recGuichet.chkSolde
recOpération_Compta.chkAmjValeur = G_recGuichet.chkAmjValeur

recOpération_Compta.Référence = G_recGuichet.Référence
recOpération_Compta.Devise2 = G_CV1.DeviseIso
recOpération_Compta.Montant2 = G_recGuichet.Montant
recOpération_Compta.Devise3 = G_CV3.DeviseIso
recOpération_Compta.Montant3 = G_recGuichet.MontantEuro
dbGuichet_Compta_Update recOpération_Compta

End Sub

Public Sub MvtCpt_Contrepartie()
Dim X As String, Nb As Integer, I As Integer
If recGuichet_Compta.Sens = "D" Then
    recGuichet_Compta.CptMvtLigne = recGuichet_Compta.CptMvtLigne - 1
    X = " versements du "
Else
    X = " retraits du "
End If

xGuichet_Compta = recGuichet_Compta
xGuichet_Compta.Method = "Seek="
If tableGuichet_Compta_Read(xGuichet_Compta) = 0 Then
    recGuichet_Compta.Method = "Update"
    recGuichet_Compta.Montant = recGuichet_Compta.Montant + xGuichet_Compta.Montant
    I = InStr(recGuichet_Compta.Libellé, " ")
    Nb = Val(mId$(xGuichet_Compta.Libellé, 1, I - 1)) + 1
Else
    Nb = 1
    If recGuichet_Compta.Sens = "D" Then
        X = " versement du "
    Else
        X = " retrait du "
    End If
End If

recGuichet_Compta.Libellé = Trim(Format$(Nb, "#### ###")) & X & dateImp(DSys)
           
End Sub

Public Sub MvtCpt_Contrepartie_Chèque()
Dim X As String, Nb As Integer, I As Integer

xGuichet_Compta = recGuichet_Compta
xGuichet_Compta.Method = "Seek="
If tableGuichet_Compta_Read(xGuichet_Compta) = 0 Then
    recGuichet_Compta.Method = "Update"
    recGuichet_Compta.Montant = recGuichet_Compta.Montant + xGuichet_Compta.Montant
    I = InStr(recGuichet_Compta.Libellé, " ")
    Nb = Val(mId$(xGuichet_Compta.Libellé, 1, I - 1)) + 1
Else
    Nb = 1
End If

If Nb < 2 Then
    recGuichet_Compta.Libellé = Trim(Format$(Nb, "#### ###")) & " Remise " & dateImp(DSys)
Else
    recGuichet_Compta.Libellé = Trim(Format$(Nb, "#### ###")) & " Remises " & dateImp(DSys)
End If

End Sub

Public Sub OpérationsNonvalidées_Print(Msg As String)
Dim I As Integer

recGuichet_Init G_recGuichet
G_recGuichet.Method = "SnapL1"
G_recGuichet.Société = SocId$
G_recGuichet.Agence = SocAgence$
G_recGuichet.SaisieUsr = usrId
G_recGuichet.Devise = "000"
G_recGuichet.CptMvtPièce = 0
G_recGuichet.CptMvtLigne = 0

arrGuichet(0) = G_recGuichet
arrGuichet(0).Devise = "999"
arrGuichet(0).CodeOpération = "9999"
arrGuichet(0).CptMvtPièce = 9999999
arrGuichet(0).CptMvtLigne = 9999999

arrGuichetNb = 0: arrGuichetIndex = 0
arrGuichetSuite = True
Do Until Not arrGuichetSuite
    srvGuichet_Monitor G_recGuichet
    G_recGuichet = arrGuichet(arrGuichetNb)
    G_recGuichet.Method = "SnapL1+"
Loop

If arrGuichetNb = 0 Then
    Call MsgBox("Pas d'opération en attente", vbInformation, "Guichet")
    Exit Sub
End If
MvtCpt_Init

For I = 1 To arrGuichetNb
        G_recGuichet = arrGuichet(I)
        If Trim(G_recGuichet.ValidationUsr) = "" And Trim(G_recGuichet.Journal) = Msg Then
            MvtCpt_AddNew
        End If
Next I
prtCompta.mJournal = ""
prtCompta_Monitor "000000000000", "Liste des opérations en cours", conststrGuichet_Compta
End Sub

Public Sub OpérationsTC_Print(lAmj As String)
Dim I As Integer, Msg As String

recGuichet_Init G_recGuichet
G_recGuichet.Method = "SnapL4"
G_recGuichet.Société = SocId$
G_recGuichet.Agence = SocAgence$
G_recGuichet.SaisieAmj = lAmj 'DSys
G_recGuichet.CodeOpération = ""

arrGuichet(0) = G_recGuichet
arrGuichet(0).CodeOpération = "9999"
arrGuichet(0).Référence = "9999999999"

arrGuichetNb = 0: arrGuichetIndex = 0
arrGuichetSuite = True
Do Until Not arrGuichetSuite
    srvGuichet_Monitor G_recGuichet
    G_recGuichet = arrGuichet(arrGuichetNb)
    G_recGuichet.Method = "SnapL4+"
Loop

If arrGuichetNb = 0 Then
    Call MsgBox("Pas d'opération en attente", vbInformation, "Guichet")
    Exit Sub
End If
Msg = Space$(20)
Mid$(Msg, 1, 6) = "000001"
Mid$(Msg, 7, 6) = Format$(arrGuichetNb, "000000")
prtGuichetListX Msg

End Sub

Public Sub G010Conversion()
G_arrCV030Nb = 5

G_arrCV030(1).LIBELE = G_recGuichet.Libellé
If G_CV2.EuroIn Then
    G_arrCV030(2).Compte = paramGuichetConversion
Else
    G_arrCV030(2).Compte = paramGuichetArbitrage
End If

G_arrCV030(2).Devise = G_recGuichet.Devise
G_arrCV030(2).MONDEV = G_curMontant2
G_arrCV030(2).LIBELE = G_recGuichet.ContrepartieLibellé
G_arrCV030(2).AMJVAL = mId$(G_recGuichet.Complément3, 1, 8)

G_arrCV030(3).MONDEV = 0

If G_CV1.EuroIn Then
    G_arrCV030(4).Compte = paramGuichetConversion
Else
    G_arrCV030(4).Compte = paramGuichetArbitrage
End If

G_arrCV030(4).Devise = G_recGuichet.DeviseEspèces
G_arrCV030(4).SENECR = G_arrCV030(1).SENECR
G_arrCV030(4).MONDEV = G_recGuichet.MontantEspèces
G_arrCV030(4).LIBELE = G_CV2.DeviseN & " " & G_strConversion2 & G_strMontant2
G_arrCV030(4).AMJVAL = G_arrCV030(2).AMJVAL

G_arrCV030(5).Compte = G_recGuichet.ContrepartieCompte
G_arrCV030(5).Devise = G_recGuichet.DeviseEspèces
G_arrCV030(5).SENECR = G_arrCV030(2).SENECR
G_arrCV030(5).MONDEV = G_recGuichet.MontantEspèces
G_arrCV030(5).LIBELE = G_recGuichet.ContrepartieLibellé
G_arrCV030(5).AMJVAL = G_arrCV030(2).AMJVAL


If G_recGuichet.Conversion = "B" Then
    G_arrCV030Nb = 7
    G_arrCV030(2).LIBELE = G_CV3.DeviseN & " " & G_strConversion2 & G_strMontant3
    G_arrCV030(4).LIBELE = G_CV3.DeviseN & " " & G_strConversion1 & G_strMontant3
    
    G_arrCV030(6).Devise = G_CV3.DeviseN
    G_arrCV030(6).Compte = G_arrCV030(2).Compte
    G_arrCV030(6).SENECR = G_arrCV030(1).SENECR
    G_arrCV030(6).MONDEV = G_recGuichet.MontantEuro
    G_arrCV030(6).LIBELE = G_CV2.DeviseN & " " & G_strConversion2 & G_strMontant2
    G_arrCV030(6).AMJVAL = G_arrCV030(2).AMJVAL

    G_arrCV030(7).Devise = G_CV3.DeviseN
    G_arrCV030(7).Compte = G_arrCV030(4).Compte
    G_arrCV030(7).SENECR = G_arrCV030(2).SENECR
    G_arrCV030(7).MONDEV = G_recGuichet.MontantEuro
    G_arrCV030(7).LIBELE = G_CV1.DeviseN & " " & G_strConversion1 & G_strMontant1
    G_arrCV030(7).AMJVAL = G_arrCV030(2).AMJVAL

End If

End Sub

Public Sub LibelléChèque(Msg As String)
Dim Nb As Integer, leurRéf As String

leurRéf = ""
If Trim(G_recGuichet.Complément1) <> "" Then
    leurRéf = Trim(G_recGuichet.Complément1)
    If Trim(G_recGuichet.Complément2) <> "" Then leurRéf = leurRéf & " / " & Trim(G_recGuichet.Complément2)
End If

Nb = CInt(mId$(G_recGuichet.Complément3, 10, 4))
If Nb = 1 Then
    G_recGuichet.ContrepartieLibellé = leurRéf & " Remise 1 chèque " & Msg
Else
    G_recGuichet.ContrepartieLibellé = leurRéf & " Remise " & Trim(Format$(Nb, "###0")) & " chèques " & Msg
End If
G_recGuichet.Libellé = G_recGuichet.ContrepartieLibellé ''''& "sauf bonne fin"

End Sub
