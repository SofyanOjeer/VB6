Attribute VB_Name = "prtSAB_Echelles"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer
Dim X_1ERE_FOIS As String
Dim N_Page As Integer

Dim blnPage As Boolean

Dim Nb As Long
Dim xAdresse As String
Dim xTotal As String

Dim arrYECHREL0(365) As typeYECHREL0, arrYECHREL0_Nb As Integer
Dim mYECHIMP0 As typeYECHIMP0
Dim mCurrenty_Top As Long

Dim blnNewPage As Boolean, blnAvis As Boolean
Dim wZADRESS0 As typeZADRESS0
Public Function isBanque(lCompte As String) As Boolean
Dim xSQL As String
Dim rsDenis As ADODB.Recordset

    isBanque = False
    xSQL = "select clienacat from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & lCompte & " ' and clienacat in('BQE','BQG')"
    Set rsDenis = cnsab.Execute(xSQL)
    If Not rsDenis.EOF Then
        If rsDenis(0) = "BQE" Or rsDenis(0) = "BQG" Then
            isBanque = True
        End If
    End If
    If rsDenis.State = adStateOpen Then
        rsDenis.Close
    End If
    Set rsDenis = Nothing
    
End Function

Public Sub prtSAB_Echelles_FTP(lFile As String, lMe As Form)
'Dim meYBIAMON0 As typeYBIAMON0
Dim xMax As String
Dim wIBM_ECH_File As String, wNT_ECH_File As String
Dim xIn As String
Dim idFile As Integer
On Error GoTo Error_Handle

Call lstErr_Clear(lMe.lstErr, lMe.cmdContext, "cmdYECHEDI01P2 : ECHELLE")

'idFile = FreeFile
'Open lFile For Input As #idFile
'Line Input #idFile, xIn
'Close idFile

'wIBM_ECH_File = "YECHEDIW"
'wNT_ECH_File = paramYBase_DataF & wIBM_ECH_File & paramYBase_Data_ExtensionP

'   recYBIAMON0_Init meYBIAMON0
'    meYBIAMON0.MONAPP = "ECHELLES"
'    meYBIAMON0.MONFLUX = ""
    
'    meYBIAMON0.Method = "OUT_FTP"
'    meYBIAMON0.MONJOB = Mid$(xIn, 71, 10)  'nom du job
'    meYBIAMON0.MONUSR = Mid$(xIn, 34, 10)  'usr
'    meYBIAMON0.MONAMJ = Mid$(xIn, 13, 6)  'No job
'    meYBIAMON0.MONHMS = Mid$(xIn, 19, 5)  'séquence
'    If IsNull(srvYBIAMON0_Monitor(meYBIAMON0)) Then
        
'        If Trim(meYBIAMON0.MONSTATUS) = "OUT_FTP" Then
            
'            Call Shell_FTP(wNT_ECH_File, paramIBM_Library_SABSPE, wIBM_ECH_File, True, False)
            
'            meYBIAMON0.Method = "OK"
'            meYBIAMON0.MONUSR = usrId
'            meYBIAMON0.MONAMJ = DSys
'            meYBIAMON0.MONHMS = time_Hms
'            Call srvYBIAMON0_Monitor(meYBIAMON0)
        
'        End If
'    End If
prtSAB_Echelles_Monitor wNT_ECH_File

Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "cmdYECHEDI01P2 : Fin*")
Exit Sub

Error_Handle:
 MsgBox "prtSAB_Echelles_FTP" & Error, vbCritical, Error
Close

End Sub

Public Sub prtSAB_Echelles_ECHAVI02P1(lFile As String, lECHIMPJOBS As Integer)
Dim V, X As String
Dim xIn As String
Dim idFile As Integer
Dim zYECHIMP0 As typeYECHIMP0, xYECHIMP0 As typeYECHIMP0
Dim blnInsert As Boolean, blnECHIMPCPT As Boolean
Dim K As Integer, X8 As String, wAAAA As Long
Dim wMontant As Currency, wSens As String, wValeur As Long
Dim kECHIMPAD As Integer
'On Error GoTo Error_Handle

rsYECHIMP0_Init zYECHIMP0
blnInsert = False

idFile = FreeFile
Open lFile For Input As #idFile

Do Until EOF(1)
    Line Input #idFile, xIn
    
    Select Case Mid$(xIn, 1, 3)
        Case "$  "
                    If Mid$(xIn, 24, 10) <> "ECHAVI02P1" Then V = "Ce n'est pas un état ECHAVI02P1": GoTo Error_MsgBox
                    
                    zYECHIMP0.ECHIMPDTRT = Val(Mid$(xIn, 5, 8))
                    zYECHIMP0.ECHIMPJOB = Val(Mid$(xIn, 13, 6))
                    zYECHIMP0.ECHIMPJOBS = Val(Mid$(xIn, 19, 5))
                    lECHIMPJOBS = zYECHIMP0.ECHIMPJOBS
                    xIn = "delete from " & paramIBM_Library_SABSPE & ".YECHIMP0 where echimpjob=" & zYECHIMP0.ECHIMPJOB & " and echimpjobs=" & zYECHIMP0.ECHIMPJOBS
                    Call FEU_ROUGE
                    Set rsSab = cnsab.Execute(xIn)
                    Call FEU_VERT
       Case "010"
                   If blnInsert Then
                        V = sqlYECHIMP0_Insert(xYECHIMP0)
                        If Not IsNull(V) Then GoTo Error_MsgBox
                    End If
                   
                   blnInsert = True
                   blnECHIMPCPT = False
                   zYECHIMP0.ECHIMPSEQ = zYECHIMP0.ECHIMPSEQ + 1
                   xYECHIMP0 = zYECHIMP0
        Case "$$ "
                   If blnInsert Then
                        V = sqlYECHIMP0_Insert(xYECHIMP0)
                        If Not IsNull(V) Then GoTo Error_MsgBox
                    End If
                        
        Case Else
            If Not blnECHIMPCPT Then
            '________________________________________________________________________________
                K = InStr(xIn, "N/REF")
                If K > 0 Then
                    xYECHIMP0.ECHIMPNREF = Mid$(xIn, 25, 6)
                    kECHIMPAD = 0
                Else
                    K = InStr(xIn, "Date d'opération :")
                    If K > 0 Then
                        xYECHIMP0.ECHIMPDOPE = dateX8_N8(Mid$(xIn, 26, 8))
                    Else
                        K = InStr(xIn, "Arrété du")
                        If K > 0 Then
                            xYECHIMP0.ECHIMPDDEB = dateX8_N8(Mid$(xIn, 17, 8))
                            wAAAA = Int(xYECHIMP0.ECHIMPDDEB / 10000) * 10000
                            xYECHIMP0.ECHIMPDFIN = dateX8_N8(Mid$(xIn, 29, 8))
                        Else
                            K = InStr(xIn, "COMPTE :")
                            If K > 0 Then
                                blnECHIMPCPT = True
                                xYECHIMP0.ECHIMPCPT = Mid$(xIn, 16, 20)
                                xYECHIMP0.ECHIMPDEV = Mid$(xIn, 37, 3)
                            End If
                        End If
                    End If
                End If
                kECHIMPAD = kECHIMPAD + 1
                Select Case kECHIMPAD
                    Case 1: xYECHIMP0.ECHIMPAD1 = Mid$(xIn, 52, 32)
                    Case 2: xYECHIMP0.ECHIMPAD2 = Mid$(xIn, 52, 32)
                    Case 3: xYECHIMP0.ECHIMPAD3 = Mid$(xIn, 52, 32)
                    Case 4: xYECHIMP0.ECHIMPAD4 = Mid$(xIn, 52, 32)
                    Case 5: xYECHIMP0.ECHIMPAD5 = Mid$(xIn, 52, 32)
                    Case 6: xYECHIMP0.ECHIMPAD6 = Mid$(xIn, 52, 32)
                    Case 7: xYECHIMP0.ECHIMPAD7 = Mid$(xIn, 52, 32)
                End Select
        Else
            '________________________________________________________________________________
            If Mid$(xIn, 7, 3) = "ECH" Then
                X = Replace(Trim(Mid$(xIn, 22, 20)), ".", "")
                wMontant = CCur(X)
                wSens = Mid$(xIn, 44, 1)
                wValeur = wAAAA + Val(Mid$(xIn, 50, 2)) * 100 + Val(Mid$(xIn, 47, 2))
                Select Case Trim(Mid$(xIn, 54, 32))
                    Case "Intérêts créditeurs"
                            xYECHIMP0.ECHIMPICRM = wMontant
                            xYECHIMP0.ECHIMPICRS = wSens
                            xYECHIMP0.ECHIMPICRV = wValeur
                            Line Input #idFile, xIn
                            xYECHIMP0.ECHIMPICRT = CDbl(Trim(Mid$(xIn, 66, 11)))
                    Case "Intérêts débiteurs"
                            xYECHIMP0.ECHIMPIDEM = wMontant
                            xYECHIMP0.ECHIMPIDES = wSens
                            xYECHIMP0.ECHIMPIDEV = wValeur
                            Line Input #idFile, xIn
                            xYECHIMP0.ECHIMPIDET = CDbl(Trim(Mid$(xIn, 59, 8)))
                    Case "Com.de mouvements", "Com. de mouvements"
                           xYECHIMP0.ECHIMPIDEV = wValeur
                           xYECHIMP0.ECHIMPCMVT = wMontant
                    Case "Com. de plus fort découvert"
                            xYECHIMP0.ECHIMPCPFD = wMontant
                            xYECHIMP0.ECHIMPIDEV = wValeur
                    Case "Com. de compte", "Com. de tenue de compte", "Frais de tenue de compte"
                            xYECHIMP0.ECHIMPCCPT = wMontant
                            xYECHIMP0.ECHIMPIDEV = wValeur
                            '===============================
                            'TODO Prélèvement libératoire
                            '===============================
                  Case Else
                        MsgBox "ECHAVI02P1 : non traité :" & Error, vbCritical, Mid$(xIn, 54, 32)
                      
               End Select
            Else
                If Mid$(xIn, 7, 6) = " Total" Then
                    X = Replace(Trim(Mid$(xIn, 22, 20)), ".", "")
                    xYECHIMP0.ECHIMPMON = CCur(X)
                    xYECHIMP0.ECHIMPMONS = Mid$(xIn, 44, 1)
                End If
             End If
       End If

            
   End Select
Loop

Close idFile

Exit Sub

Error_Handle:
V = Error
Error_MsgBox:
MsgBox "prtSAB_Echelles_ECHAVI02P1" & Error, vbCritical, V
Close

End Sub

Public Sub prtSAB_Echelles_ECHEDI01P2(lFile As String, lECHIMPJOBS As Integer, blnNostro As Boolean, blnSoldeZ As Boolean, blnCompteAvis As Boolean, selCompte As String)
Dim V, X As String, I As Integer
Dim xIn As String, xIn_Report As String
Dim idFile As Integer
Dim zYECHIMP0 As typeYECHIMP0, xYECHIMP0 As typeYECHIMP0
Dim blnHeader As Boolean, blnEnd As Boolean
Dim K As Integer, X8 As String, wAAAA As Long
Dim wMontant As Currency, wSens As String, wValeur As Long
Dim iAdresse As Integer
Dim blnOk As Boolean
Dim newName As String
Dim uneSeuleArchivePDF As Boolean
Dim tmpPDFname As String

On Error GoTo Error_Handle

tmpPDFname = paramIMP_PDF_Path_Temp & "\Releve_.pdf"

uneSeuleArchivePDF = True

traitementPDF:

Height8_6 = frmElpPrt.prtHeightDelta(9, 7)

rsYECHIMP0_Init zYECHIMP0
blnHeader = False: blnEnd = False
arrYECHREL0_Nb = 0

idFile = FreeFile
Open lFile For Input As #idFile
'____________________________________________________________________________________________
Line Input #idFile, xIn
If Mid$(xIn, 24, 10) <> "ECHEDI01P2" Then V = "Ce n'est pas un état ECHEDI01P2": GoTo Error_MsgBox

zYECHIMP0.ECHIMPDTRT = Val(Mid$(xIn, 5, 8))
zYECHIMP0.ECHIMPJOB = Val(Mid$(xIn, 13, 6))
zYECHIMP0.ECHIMPJOBS = lECHIMPJOBS '''Val(Mid$(xIn, 19, 5)) + 2  ''!!! à vérifier
If uneSeuleArchivePDF Then
    prtSAB_Echelles_ECHEDI01P2_Open
End If
'____________________________________________________________________________________________
Do Until EOF(idFile)
    Line Input #idFile, xIn
    If Mid$(xIn, 1, 3) = "064" Or Mid$(xIn, 13, 10) = "Nbre Débit" Then
                 arrYECHREL0_Nb = arrYECHREL0_Nb + 1
                 arrYECHREL0(arrYECHREL0_Nb).ECHRELSD = 0
                 arrYECHREL0(arrYECHREL0_Nb).ECHRELSDS = " "
                 arrYECHREL0(arrYECHREL0_Nb).ECHRELDVAL = 0
                 
                 X = Trim(Mid$(xIn, 26, 14))
                 If X = "" Then
                    arrYECHREL0(arrYECHREL0_Nb).ECHRELMDB = 0
                 Else
                    arrYECHREL0(arrYECHREL0_Nb).ECHRELMDB = CCur(X)     '!!!!!! nbr DB
                 End If
                 X = Trim(Mid$(xIn, 63, 14))
                 If X = "" Then
                    arrYECHREL0(arrYECHREL0_Nb).ECHRELMCR = 0
                 Else
                    arrYECHREL0(arrYECHREL0_Nb).ECHRELMCR = CCur(X)         '!!!!!! nbr CR
                 End If
                 
                Line Input #idFile, xIn
                Line Input #idFile, xIn
                arrYECHREL0(arrYECHREL0_Nb).ECHRELDVAL = dateX8_N8(Mid$(xIn, 46, 8))
                X = Trim(Mid$(xIn, 57, 14))
                If X = "" Then
                   arrYECHREL0(arrYECHREL0_Nb).ECHRELSD = 0
                Else
                   arrYECHREL0(arrYECHREL0_Nb).ECHRELSD = CCur(X)     '!!!!!! solde final
                End If
                arrYECHREL0(arrYECHREL0_Nb).ECHRELSDS = Mid$(xIn, 71, 1)

                If blnHeader Then
'______________________________________________________________________________________________
                    blnOk = True
                    If selCompte <> Mid$(xYECHIMP0.ECHIMPCPT, 1, Len(selCompte)) Then blnOk = False
                    If Mid$(xYECHIMP0.ECHIMPCPT, 1, 1) = "N" And Not blnNostro Then blnOk = False
                    If Not blnSoldeZ And arrYECHREL0_Nb = 2 And arrYECHREL0(0).ECHRELSD = 0 And arrYECHREL0(arrYECHREL0_Nb).ECHRELSD = 0 Then blnOk = False
                    If Not uneSeuleArchivePDF Then
                            If InStr(UCase(Trim(Printer.Devicename)), "PDF") >= 1 And InStr(Trim(xYECHIMP0.ECHIMPCPT), "11441") > 0 Then
                                'DR 20/11/2019 On force l'impression des échelles du Client 11441 HYPROC SHIPPING, qui n'est pas une Banque
                                'Si blnOk est à True alors cette échelle sera imprimée
                            ElseIf InStr(UCase(Trim(Printer.Devicename)), "PDF") >= 1 And Not isBanque(Trim(xYECHIMP0.ECHIMPCPT)) Then
                                blnOk = False
                            End If
                    End If
                    If Not uneSeuleArchivePDF Then
                        If blnOk Then
                            If prtSAB_Echelles_ECHEDI01P2_Relevé(xYECHIMP0, blnCompteAvis, uneSeuleArchivePDF) Then
                                Call prtSAB_Echelles_Close(1000)
                                If Dir(tmpPDFname) <> "" Then
                                    newName = Retourne_Num_Client(Trim(xYECHIMP0.ECHIMPCPT)) & "_" & Trim(xYECHIMP0.ECHIMPCPT) & "_" & Left(DSYS_Time, Len(DSYS_Time) - 1) & ".pdf"
                                    FileCopy tmpPDFname, prtPgmName & "\" & newName
                                    Kill tmpPDFname
                                End If
                            End If
                        End If
                    Else
                        If blnOk Then
                            prtSAB_Echelles_ECHEDI01P2_Relevé xYECHIMP0, blnCompteAvis, uneSeuleArchivePDF
                        End If
                    End If
 '______________________________________________________________________________________________
                   blnHeader = False
                End If
                               
                Line Input #idFile, xIn
                If Mid$(xIn, 1, 3) = "$$ " Then Exit Do
                K = InStr(xIn, "ECHELLES")
                If K <= 0 Then Line Input #idFile, xIn
                If Mid$(xIn, 1, 3) = "$$ " Then Exit Do
    Else
        Select Case Mid$(xIn, 1, 3)
           Case "013"
                     If Not blnHeader Then
                       xYECHIMP0 = zYECHIMP0
                       xYECHIMP0.ECHIMPAD1 = Mid$(xIn, 47, 32)
                       iAdresse = 2
                       Do
                            Line Input #idFile, xIn
                            If Mid$(xIn, 1, 3) <> "022" Then
                                Select Case iAdresse
                                    Case 2: xYECHIMP0.ECHIMPAD2 = Mid$(xIn, 47, 32)
                                    Case 3: xYECHIMP0.ECHIMPAD3 = Mid$(xIn, 47, 32)
                                    Case 4: xYECHIMP0.ECHIMPAD4 = Mid$(xIn, 47, 32)
                                    Case 5: xYECHIMP0.ECHIMPAD5 = Mid$(xIn, 47, 32)
                                    Case 6: xYECHIMP0.ECHIMPAD6 = Mid$(xIn, 47, 32)
                                    Case 7: xYECHIMP0.ECHIMPAD7 = Mid$(xIn, 47, 32)
                                End Select
                                iAdresse = iAdresse + 1
                            Else
                            
           'Case "021"
                                Line Input #idFile, xIn
                                X = Trim(Mid$(xIn, 19, 20))
                                
                                 If blnHeader Then
                                    If X <> Trim(xYECHIMP0.ECHIMPCPT) Then V = "Erreur rupture page compte " & X: GoTo Error_MsgBox
                                 Else
                                   xYECHIMP0.ECHIMPCPT = X
                                   xYECHIMP0.ECHIMPDEV = Mid$(xIn, 13, 3)
                                End If
                                Exit Do
                            End If
                        Loop
                    End If
           Case "028"
                    Line Input #idFile, xIn
                    If Mid$(xIn, 13, 13) <> "Report valeur" Then
                        V = "Erreur Report valeur, compte " & xYECHIMP0.ECHIMPCPT
                        GoTo Error_MsgBox
                    End If
                    Line Input #idFile, xIn
                    If Not blnHeader Then
                        blnHeader = True
                        xYECHIMP0.ECHIMPDDEB = dateX6_N8(Mid$(xIn, 34, 6))
                        arrYECHREL0_Nb = 0
                        arrYECHREL0(0).ECHRELDVAL = xYECHIMP0.ECHIMPDDEB
                        arrYECHREL0(0).ECHRELSD = CCur(Trim(Mid$(xIn, 45, 14)))
                        arrYECHREL0(0).ECHRELSDS = Mid$(xIn, 59, 1)
                    End If
                    Line Input #idFile, xIn
           Case "062"
                    Line Input #idFile, xIn_Report
                    Do
                        Line Input #idFile, xIn
                    Loop Until Mid$(xIn, 1, 3) = "028"
                    Line Input #idFile, xIn
                    Line Input #idFile, xIn
                    Line Input #idFile, xIn
           Case Else
                    
                    If blnHeader Then
                        If Trim(Mid$(xIn, 5, 29)) <> "" Then        ' rupture mois dans certains cas
                             arrYECHREL0_Nb = arrYECHREL0_Nb + 1
                             X = Trim(Mid$(xIn, 13, 15))
                             If X = "" Then
                                arrYECHREL0(arrYECHREL0_Nb).ECHRELMDB = 0
                             Else
                                arrYECHREL0(arrYECHREL0_Nb).ECHRELMDB = CCur(X)
                             End If
                             
                            
                             Line Input #idFile, xIn
                             
                             X = Trim(Mid$(xIn, 17, 16))
                             If X = "" Then
                                arrYECHREL0(arrYECHREL0_Nb).ECHRELMCR = 0
                              Else
                                arrYECHREL0(arrYECHREL0_Nb).ECHRELMCR = CCur(X)
                             End If
                             X = Trim(Mid$(xIn, 45, 14))
                             If X = "" Then
                               arrYECHREL0(arrYECHREL0_Nb).ECHRELSD = 0
                              Else
                               arrYECHREL0(arrYECHREL0_Nb).ECHRELSD = CCur(X)
                             End If
                               
                             arrYECHREL0(arrYECHREL0_Nb).ECHRELSDS = Mid$(xIn, 59, 1)
                             arrYECHREL0(arrYECHREL0_Nb).ECHRELDVAL = dateX6_N8(Mid$(xIn, 34, 6))
                             
                             X = Trim(Mid$(xIn, 41, 3))
                             If X = "" Then
                                arrYECHREL0(arrYECHREL0_Nb).ECHRELNBJ = 0
                             Else
                                arrYECHREL0(arrYECHREL0_Nb).ECHRELNBJ = CInt(X)
                             End If
                             X = Trim(Mid$(xIn, 61, 14))
                             If X = "" Then
                                 arrYECHREL0(arrYECHREL0_Nb).ECHRELNBR = 0
                             Else
                                  arrYECHREL0(arrYECHREL0_Nb).ECHRELNBR = CCur(X)
                            End If
                             X = Trim(Mid$(xIn, 75, 10))
                            If X = "" Then
                                 arrYECHREL0(arrYECHREL0_Nb).ECHRELTAUX = 0
                             Else
                                If InStr(X, "%") > 0 Then
                                    X = Replace(Trim(Mid$(xIn, 75, 10)), "%", ",")
                                    arrYECHREL0(arrYECHREL0_Nb).ECHRELTAUX = CDbl(X)
                                Else
                                     arrYECHREL0(arrYECHREL0_Nb).ECHRELTAUX = CDbl(X) / 100000
                               End If
                            End If
                        Else
                             X = Trim(Mid$(xIn, 75, 10))
                             If arrYECHREL0(arrYECHREL0_Nb).ECHRELTAUX = 0 And X <> "" Then
                                If InStr(X, "%") > 0 Then
                                    X = Replace(Trim(Mid$(xIn, 75, 10)), "%", ",")
                                    arrYECHREL0(arrYECHREL0_Nb).ECHRELTAUX = CDbl(X)
                                Else
                                     arrYECHREL0(arrYECHREL0_Nb).ECHRELTAUX = CDbl(X) / 100000
                               End If
                            End If
                       
                        End If
                    End If
        End Select
    End If
Loop

Close idFile
If uneSeuleArchivePDF Then
    Call prtSAB_Echelles_Close(60000)
    '------------------------------------------------------------------------------------------
    prtPgmName = paramServer("\\Facturation\") & "Echelles\IMP_PDF_Echelles_" & Left(DSYS_Time, 4) & "_" & Mid(DSYS_Time, 5, 2) & "_" & Mid(DSYS_Time, 7, 2)
    If Dir(prtPgmName, vbDirectory) = "" Then
        MkDir prtPgmName
    End If
    uneSeuleArchivePDF = False
    blnHeader = False: blnEnd = False
    GoTo traitementPDF
End If
If Not uneSeuleArchivePDF Then
    If Dir(tmpPDFname) <> "" Then
        newName = Retourne_Num_Client(Trim(xYECHIMP0.ECHIMPCPT)) & "_" & Trim(xYECHIMP0.ECHIMPCPT) & "_" & Left(DSYS_Time, Len(DSYS_Time) - 1) & ".pdf"
        FileCopy tmpPDFname, prtPgmName & "\" & newName
        Call pause_with_events(Retourne_WAIT_PDF)
        Kill tmpPDFname
    End If
End If

MsgBox "Fin du traitement Echelles..."
Exit Sub

Error_Handle:
V = Error
Error_MsgBox:

MsgBox "prtSAB_Echelles_ECHAVI02P1" & Error, vbCritical, V
Close

End Sub
Public Sub prtSAB_Echelles_Monitor(lNT_ECH_File As String)
Dim xIn As String

prtTitleText = "SAB : Edition Echelles"
prtFontName = prtFontName_Arial
prtSAB_Echelles_Open
prtHeaderHeight = 300

XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

Open lNT_ECH_File For Input As #1

Nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    Select Case Mid$(xIn, 1, 1)
        Case 1:
            N_Page = 0
            Nb = Nb + 1
            '   If Nb = 3 Then Exit Do
            xAdresse = xIn
            '   If mId$(xAdresse, 175, 20) = "11084978001         " Then
            X_1ERE_FOIS = "O"
            If Nb <> 1 Then prtSAB_Echelles_Colonne: frmElpPrt.prtNewPage
            prtSAB_Echelles_Form
            '   End If
         Case 2:
            '   If mId$(xAdresse, 175, 20) = "11084978001         " Then
            prtSAB_Echelles_Line xIn
            '   End If
         Case 3:
            '   If mId$(xAdresse, 175, 20) = "11084978001         " Then
            prtSAB_Echelles_Piedpage xIn
            '   End If
    End Select
Loop
Close
prtSAB_Echelles_Colonne
Call prtSAB_Echelles_Close(1000)

End Sub

'---------------------------------------------------------
Public Sub prtSAB_Echelles_Form()
'---------------------------------------------------------
Dim X As String
Dim curX As Currency

XPrt.FontSize = 9: XPrt.FontBold = False
XPrt.CurrentY = prtMinY
XPrt.CurrentX = prtMedX - 800: XPrt.Print Mid$(xAdresse, 152, 20);
XPrt.FontBold = True
XPrt.CurrentX = prtMedX + 3000: XPrt.Print Mid$(xAdresse, 2, 30);
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMedX + 3000: XPrt.Print Mid$(xAdresse, 32, 30);
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMedX + 3000: XPrt.Print Mid$(xAdresse, 62, 30);
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMedX - 800: XPrt.Print "Devise   : ";
XPrt.FontBold = True
XPrt.CurrentX = prtMedX + 100: XPrt.Print Mid$(xAdresse, 172, 3);
XPrt.FontBold = False
XPrt.CurrentX = prtMedX + 3000: XPrt.Print Mid$(xAdresse, 92, 30);
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMedX - 800: XPrt.Print "Compte : ";
XPrt.FontBold = True
XPrt.CurrentX = prtMedX + 100: XPrt.Print Mid$(xAdresse, 175, 20);
XPrt.FontBold = False
XPrt.CurrentX = prtMedX + 3000: XPrt.Print Mid$(xAdresse, 122, 30);

' Titre de l'édition
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontSize = 11: XPrt.FontBold = True     'Caractère GRAS
frmElpPrt.prtCentré prtMedX, "RELEVE  ECHELLE  D'INTERETS"
XPrt.FontBold = False

N_Page = N_Page + 1
If N_Page > 1 Then
    XPrt.FontSize = 7
    XPrt.CurrentX = prtMaxX - 1500: XPrt.Print "Page : " & N_Page;
End If

' Entête de colonne
XPrt.DrawWidth = 1
XPrt.FontSize = 9: XPrt.FontBold = True     'Caractère GRAS

Call frmElpPrt.prtTrame(prtMinX, prtMinY + 2200, prtMaxX, prtMinY + 2200 + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 2200 + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print "Date valeur";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Mvts débiteurs";
XPrt.CurrentX = prtMinX + 4400: XPrt.Print "Mvts créditeurs";
XPrt.CurrentX = prtMinX + 6600: XPrt.Print "Solde débiteur";
XPrt.CurrentX = prtMinX + 8500: XPrt.Print "Solde créditeur";
XPrt.CurrentX = prtMinX + 10400: XPrt.Print "Nb jours";
XPrt.CurrentX = prtMinX + 11500: XPrt.Print "Nombres débiteurs";
XPrt.CurrentX = prtMinX + 14000: XPrt.Print "Nombres créditeurs";

' Report de solde à chaque nouveau compte : lecture code enreg = 1
If X_1ERE_FOIS = "O" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    X = Mid$(xAdresse, 195, 2) & " - " & Mid$(xAdresse, 197, 2) & " - 20" & Mid$(xAdresse, 199, 2)
    XPrt.CurrentX = prtMinX + 100: XPrt.Print X;
    XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Report solde en valeur";
    curX = CCur(Mid$(xAdresse, 201, 15))
    X = Format$(curX, "### ### ### ##0.00")
    If Mid$(xAdresse, 216, 2) = "CR" Then
            XPrt.CurrentX = prtMinX + 9800 - XPrt.TextWidth(X): XPrt.Print X;
    Else
            XPrt.CurrentX = prtMinX + 7900 - XPrt.TextWidth(X): XPrt.Print X;
    End If
Else
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

X_1ERE_FOIS = "N"
XPrt.FontBold = False   'Plus de caractère GRAS

End Sub


'---------------------------------------------------------
Public Sub prtSAB_Echelles_ECHEDI01P2_Form()
'---------------------------------------------------------
Dim X As String
Dim curX As Currency
blnNewPage = True
XPrt.FontSize = 9: XPrt.FontBold = False
XPrt.CurrentY = prtMinY + prtlineHeight * 4
XPrt.CurrentX = prtMinX + 7800: XPrt.Print "Paris,le " & dateImp10(mYECHIMP0.ECHIMPDTRT);
XPrt.FontBold = True
XPrt.CurrentY = 2400
XPrt.CurrentX = 5700: XPrt.Print mYECHIMP0.ECHIMPAD1;
XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 100: XPrt.Print "Devise   : ";
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 1500: XPrt.Print mYECHIMP0.ECHIMPDEV;
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 100: XPrt.Print "Compte : ";
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 1500: XPrt.Print mYECHIMP0.ECHIMPCPT;
XPrt.FontBold = False
XPrt.CurrentX = 5700: XPrt.Print mYECHIMP0.ECHIMPAD2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5700: XPrt.Print mYECHIMP0.ECHIMPAD3;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5700: XPrt.Print mYECHIMP0.ECHIMPAD4;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5700: XPrt.Print mYECHIMP0.ECHIMPAD5;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5700: XPrt.Print mYECHIMP0.ECHIMPAD6;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5700: XPrt.Print mYECHIMP0.ECHIMPAD7;



' Titre de l'édition
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontSize = 11: XPrt.FontBold = True     'Caractère GRAS
If mYECHIMP0.ECHIMPDOPE <= mYECHIMP0.ECHIMPDFIN Then
    frmElpPrt.prtCentré prtMedX, "RELEVE  D'ECHELLES  D'INTERETS"
Else
    XPrt.ForeColor = vbMagenta
    frmElpPrt.prtCentré prtMedX, "RELEVE RECTIFICATIF D'ECHELLES  D'INTERETS"
    XPrt.ForeColor = prtForeColor
End If

XPrt.FontBold = False

N_Page = N_Page + 1
If N_Page > 1 Then
    XPrt.FontSize = 7
    XPrt.CurrentX = prtMaxX - 1500: XPrt.Print "Page : " & N_Page;
End If

' Entête de colonne
XPrt.DrawWidth = 1
XPrt.FontSize = 8: XPrt.FontBold = True     'Caractère GRAS
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
mCurrenty_Top = XPrt.CurrentY

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.CurrentX = prtMinX + 80: XPrt.Print "Date valeur";
XPrt.CurrentX = prtMinX + 1200: XPrt.Print "Mvts débiteurs";
XPrt.CurrentX = prtMinX + 2700: XPrt.Print "Mvts créditeurs";
XPrt.CurrentX = prtMinX + 4150: XPrt.Print "Solde débiteur";
XPrt.CurrentX = prtMinX + 5550: XPrt.Print "Solde créditeur";
XPrt.CurrentX = prtMinX + 7130: XPrt.Print "jours";
XPrt.CurrentX = prtMinX + 7700: XPrt.Print "Nombres débiteurs/créditeurs";
XPrt.CurrentX = prtMinX + 10400: XPrt.Print "Taux";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 100

XPrt.FontBold = False
XPrt.FontSize = 8

End Sub

'---------------------------------------------------------
Public Sub prtSAB_Echelles_ECHEDI01P2_Avis(lYECHIMP0 As typeYECHIMP0)
'---------------------------------------------------------
Dim X As String, xLib As String
Dim blnNormal As Boolean

    If XPrt.CurrentY + prtlineHeight * 10 >= prtMaxY Then
        prtSAB_Echelles_ECHEDI01P2_Colonne
        frmElpPrt.prtNewPage
        prtSAB_Echelles_ECHEDI01P2_Form
    End If
' Titre de l'édition
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
    XPrt.FontSize = 13: XPrt.FontBold = True
    If mYECHIMP0.ECHIMPDOPE <= mYECHIMP0.ECHIMPDFIN Then
        blnNormal = True
        XPrt.CurrentX = prtMinX + 1600
        XPrt.Print "Avis d'opération n° " & Trim(lYECHIMP0.ECHIMPNREF) & ", arrêté du " & dateImp10(lYECHIMP0.ECHIMPDDEB) & " au " & dateImp10(lYECHIMP0.ECHIMPDFIN);
    Else
        blnNormal = False
        XPrt.CurrentX = prtMinX + 600
        XPrt.ForeColor = vbMagenta
        XPrt.Print "Avis rectificatif d'opération n° " & Trim(lYECHIMP0.ECHIMPNREF) & ", arrêté du " & dateImp10(lYECHIMP0.ECHIMPDDEB) & " au " & dateImp10(lYECHIMP0.ECHIMPDFIN);
        XPrt.ForeColor = prtForeColor
    End If
    XPrt.FontBold = False
    XPrt.DrawWidth = 1
    XPrt.FontSize = 9: XPrt.FontBold = True     'Caractère GRAS
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    mCurrenty_Top = XPrt.CurrentY
'---------------------------------------------------------
    'Call frmElpPrt.prtTrame(prtMinX + 11500, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight * 4, " ", 240)
    'XPrt.FontSize = 13: XPrt.FontBold = True
    'XPrt.CurrentX = prtMinX + 12500: XPrt.Print "Avis d'opération n° " & Trim(lYECHIMP0.ECHIMPNREF);
    'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    'XPrt.CurrentX = prtMinX + 12000: XPrt.Print "arrêté du " & dateImp10(lYECHIMP0.ECHIMPDDEB) & " au " & dateImp10(lYECHIMP0.ECHIMPDFIN);
    'XPrt.FontBold = False
'---------------------------------------------------------
    prtFillColor = RGB(0, 123, 141)
    XPrt.ForeColor = vbWhite
    XPrt.CurrentY = mCurrenty_Top
    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMinX + 11100, XPrt.CurrentY + prtlineHeight + 50, "B")
'---------------------------------------------------------
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.FontSize = 10
    XPrt.CurrentX = prtMinX + 200: XPrt.Print "Date valeur";
    XPrt.CurrentX = prtMinX + 2800: XPrt.Print "débit";
    XPrt.CurrentX = prtMinX + 5180: XPrt.Print "crédit";
    XPrt.CurrentX = prtMinX + 6500: XPrt.Print "Libellé";
    XPrt.ForeColor = vbBlack
    XPrt.CurrentY = XPrt.CurrentY + 50
    If lYECHIMP0.ECHIMPIDEM <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
        XPrt.CurrentX = prtMinX + 200: XPrt.Print dateImp10(lYECHIMP0.ECHIMPIDEV);
        X = Format$(lYECHIMP0.ECHIMPIDEM, "### ### ### ##0.00")
        If lYECHIMP0.ECHIMPIDES = "D" Then
            XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X)
        Else
            XPrt.CurrentX = prtMinX + 5700 - XPrt.TextWidth(X)
       End If
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 6500
        XPrt.Print "Intérêts débiteurs, TEG : " & Format$(lYECHIMP0.ECHIMPIDET, "#0.00") & " %";
    End If
    If lYECHIMP0.ECHIMPICRM <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
        XPrt.CurrentX = prtMinX + 200: XPrt.Print dateImp10(lYECHIMP0.ECHIMPICRV);
        X = Format$(lYECHIMP0.ECHIMPICRM, "### ### ### ##0.00")
        If lYECHIMP0.ECHIMPICRS = "D" Then
            XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X)
        Else
            XPrt.CurrentX = prtMinX + 5700 - XPrt.TextWidth(X)
       End If
       XPrt.Print X;
       XPrt.CurrentX = prtMinX + 6500
       If blnNormal Then
           XPrt.Print "Intérêts créditeurs, taux moyen : " & Format$(lYECHIMP0.ECHIMPICRT, "#0.000000") & " %";
       Else
           XPrt.Print "Intérêts créditeurs";
       End If
    End If
    If lYECHIMP0.ECHIMPCPFD <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
        XPrt.CurrentX = prtMinX + 200: XPrt.Print dateImp10(lYECHIMP0.ECHIMPIDEV);
        X = Format$(lYECHIMP0.ECHIMPCPFD, "### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 6500
        XPrt.Print "Commission plus fort découvert";
    End If
    If lYECHIMP0.ECHIMPCMVT <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
        XPrt.CurrentX = prtMinX + 200: XPrt.Print dateImp10(lYECHIMP0.ECHIMPIDEV);
        X = Format$(lYECHIMP0.ECHIMPCMVT, "### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 6500
        XPrt.Print "Commission de mouvements";
    End If
    If lYECHIMP0.ECHIMPCCPT <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
        XPrt.CurrentX = prtMinX + 200: XPrt.Print dateImp10(lYECHIMP0.ECHIMPIDEV);
        X = Format$(lYECHIMP0.ECHIMPCCPT, "### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 6500
        XPrt.Print "Commission de compte";
    End If
'_______________________________________________________________________________________
    prtFillColor = RGB(250, 255, 255) ' RGB(0, 123, 141)
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMinX + 11100, XPrt.CurrentY + prtlineHeight + 50, "B")
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 200
    XPrt.Print "Total";
    X = Format$(lYECHIMP0.ECHIMPMON, "### ### ### ##0.00")
    If lYECHIMP0.ECHIMPMONS = "D" Then
        XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X)
        xLib = "    au débit du compte : "
    Else
        XPrt.CurrentX = prtMinX + 5700 - XPrt.TextWidth(X)
        xLib = "   au crédit du compte : "
    End If
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 6500
    XPrt.Print lYECHIMP0.ECHIMPDEV & xLib;
    XPrt.ForeColor = vbBlue
    XPrt.Print lYECHIMP0.ECHIMPCPT;
    XPrt.ForeColor = vbBlack
    XPrt.FontBold = False
    XPrt.Line (prtMinX, mCurrenty_Top)-(prtMinX, XPrt.CurrentY + prtlineHeight), prtLineColor
    XPrt.Line (prtMinX + 1600, mCurrenty_Top)-(prtMinX + 1600, XPrt.CurrentY), prtLineColor
    XPrt.Line (prtMinX + 6200, mCurrenty_Top)-(prtMinX + 6200, XPrt.CurrentY), prtLineColor
    XPrt.Line (prtMinX + 11100, mCurrenty_Top)-(prtMinX + 11100, XPrt.CurrentY), prtLineColor
    prtFillColor = prtFillColor_Standard

End Sub


Public Sub prtSAB_Echelles_Close(dureePose As Long)
On Error GoTo prtError


Call frmElpPrt.prtEndDoc(dureePose)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtSAB_Echelles_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape
prtPgmName = "prtSAB_Echelles"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100


prtFormType = ""
prtSocInit
blnNewPage = False
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
Public Sub prtSAB_Echelles_ECHEDI01P2_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait
'$JPL 20141203 archivage automatique prtPgmName = "prtSAB_Echelles"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100


prtFormType = ""
prtSocInit
prtMaxX = prtMinX + 11000
blnNewPage = False
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub





Public Sub prtSAB_Echelles_Line(lX As String)
Dim curX As Currency
Dim X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 500 > prtMaxY Then
    prtSAB_Echelles_Colonne
    frmElpPrt.prtNewPage
    prtSAB_Echelles_Form
End If

X = Mid$(lX, 36, 2) & " - " & Mid$(lX, 38, 2) & " - 20" & Mid$(lX, 40, 2)
XPrt.CurrentX = prtMinX + 100: XPrt.Print X;

curX = CCur(Mid$(lX, 2, 17))
X = Format$(curX, "### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X): XPrt.Print X;

curX = CCur(Mid$(lX, 19, 17))
X = Format$(curX, "### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 5700 - XPrt.TextWidth(X): XPrt.Print X;

curX = CCur(Mid$(lX, 42, 3))
X = Format$(curX, "###")
XPrt.CurrentX = prtMinX + 10900 - XPrt.TextWidth(X): XPrt.Print X;

curX = CCur(Mid$(lX, 45, 15))
X = Format$(curX, "### ### ### ##0.00")
If Mid$(lX, 60, 2) = "CR" Then
        XPrt.CurrentX = prtMinX + 9800 - XPrt.TextWidth(X): XPrt.Print X;
        If Not IsNumeric(Mid$(lX, 62, 14)) Then
            curX = 0
        Else
            curX = CCur(Mid$(lX, 62, 14))
        End If
        X = Format$(curX, "## ### ### ### ###")
        XPrt.CurrentX = prtMinX + 15700 - XPrt.TextWidth(X): XPrt.Print X;
Else
        XPrt.CurrentX = prtMinX + 7900 - XPrt.TextWidth(X): XPrt.Print X;
        If Not IsNumeric(Mid$(lX, 62, 14)) Then
            curX = 0
        Else
            curX = CCur(Mid$(lX, 62, 14))
        End If
        X = Format$(curX, "## ### ### ### ###")
        XPrt.CurrentX = prtMinX + 13120 - XPrt.TextWidth(X): XPrt.Print X;
End If

End Sub

Public Sub prtSAB_Echelles_ECHEDI01P2_Line(lYECHREL0 As typeYECHREL0)
Dim X As String, X2 As String
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + prtlineHeight >= prtMaxY Then
    prtSAB_Echelles_ECHEDI01P2_Colonne
    frmElpPrt.prtNewPage
    prtSAB_Echelles_ECHEDI01P2_Form
End If

XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 100: XPrt.Print dateImp10(lYECHREL0.ECHRELDVAL);
X = Format$(lYECHREL0.ECHRELMDB, "### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 2400 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lYECHREL0.ECHRELMCR, "### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 3900 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lYECHREL0.ECHRELNBJ, "##0")
XPrt.CurrentX = prtMinX + 7500 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lYECHREL0.ECHRELSD, "### ### ### ##0.00")
X2 = Format$(lYECHREL0.ECHRELNBR, "### ### ### ###")

XPrt.FontSize = 7
XPrt.CurrentY = XPrt.CurrentY + Height8_6
If lYECHREL0.ECHRELSDS <> "C" Then
    XPrt.CurrentX = prtMinX + 5400 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X2)
    XPrt.ForeColor = vbMagenta
    XPrt.Print X2;
Else
    XPrt.CurrentX = prtMinX + 6900 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 9900 - XPrt.TextWidth(X2)
    XPrt.Print X2;
    XPrt.ForeColor = vbBlue
End If
If lYECHREL0.ECHRELTAUX <> 0 Then
    XPrt.FontItalic = True
    X2 = Format$(lYECHREL0.ECHRELTAUX, " #0.000000") & " %"
    XPrt.CurrentX = prtMinX + 10900 - XPrt.TextWidth(X2)
    XPrt.Print X2;
    XPrt.FontItalic = False
End If
XPrt.ForeColor = vbBlack
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
End Sub


Public Sub prtSAB_Echelles_Piedpage(lX As String)
Dim curX As Currency
Dim X As String

' Saut de 2 lignes, si dépassement, saut de page...
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 500 > prtMaxY Then
    prtSAB_Echelles_Colonne
    frmElpPrt.prtNewPage
    prtSAB_Echelles_Form
    'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

XPrt.DrawWidth = 1
XPrt.FontSize = 9: XPrt.FontBold = True      'Caractère GRAS

Call frmElpPrt.prtTrame(prtMinX, prtMinY + XPrt.CurrentY, prtMaxX, prtMinY + XPrt.CurrentY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtMinY + 50

' Solde en valeur en pied de page
X = Mid$(lX, 64, 2) & " - " & Mid$(lX, 67, 2) & " - 20" & Mid$(lX, 70, 2)
XPrt.CurrentX = prtMinX + 100: XPrt.Print X;
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Solde en valeur";
If Not IsNumeric(Mid$(lX, 72, 15)) Then
    curX = 0
Else
    curX = CCur(Mid$(lX, 72, 15))
End If
X = Format$(curX, "### ### ### ##0.00")
If Mid$(lX, 87, 2) = "CR" Then
        XPrt.CurrentX = prtMinX + 9800 - XPrt.TextWidth(X): XPrt.Print X;
Else
        XPrt.CurrentX = prtMinX + 7900 - XPrt.TextWidth(X): XPrt.Print X;
End If

' Cumul des nombres débiteurs
curX = CCur(Mid$(lX, 2, 14))
X = Format$(curX, "## ### ### ### ###")
XPrt.CurrentX = prtMinX + 13120 - XPrt.TextWidth(X): XPrt.Print X;

' Cumul des nombres créditeurs
curX = CCur(Mid$(lX, 16, 14))
X = Format$(curX, "## ### ### ### ###")
XPrt.CurrentX = prtMinX + 15700 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.FontBold = False     'Caractère non GRAS
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub


Public Sub prtSAB_Echelles_Colonne()

XPrt.Line (prtMinX, prtMinY + 2200)-(prtMinX, XPrt.CurrentY + prtlineHeight), prtLineColor
XPrt.Line (prtMinX + 1600, prtMinY + 2200)-(prtMinX + 1600, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMinX + 6200, prtMinY + 2200)-(prtMinX + 6200, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMinX + 10250, prtMinY + 2200)-(prtMinX + 10250, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMinX + 11300, prtMinY + 2200)-(prtMinX + 11300, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMaxX, prtMinY + 2200)-(prtMaxX, XPrt.CurrentY), prtLineColor

End Sub

Public Function prtSAB_Echelles_ECHEDI01P2_Relevé(lYECHIMP0 As typeYECHIMP0, blnCompteAvis As Boolean, uneSeuleArchivePDF As Boolean) As Boolean
Dim V, I As Integer
Dim X As String, xSQL As String, Nb As Long
On Error GoTo Error_Handle
    
    prtSAB_Echelles_ECHEDI01P2_Relevé = True
    N_Page = 0
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YECHIMP0" _
   & " where ECHIMPJOB = " & lYECHIMP0.ECHIMPJOB _
     & " and ECHIMPJOBS = " & lYECHIMP0.ECHIMPJOBS _
     & " and ECHIMPCPT = '" & Trim(lYECHIMP0.ECHIMPCPT) & "'" _
     & " and ECHIMPDDEB = " & lYECHIMP0.ECHIMPDDEB _
     & " and ECHIMPDTRT = " & lYECHIMP0.ECHIMPDTRT
    Set rsSab = cnsab.Execute(xSQL, Nb)
    If rsSab.EOF Then
        mYECHIMP0 = lYECHIMP0
        blnAvis = False
        If blnCompteAvis Then
            prtSAB_Echelles_ECHEDI01P2_Relevé = False
            Exit Function
        End If
    Else
        V = rsYECHIMP0_GetBuffer(rsSab, mYECHIMP0)
        If Not IsNull(V) Then GoTo Error_Handle
        blnAvis = True
    End If
    If Not uneSeuleArchivePDF Then
        prtSAB_Echelles_ECHEDI01P2_Open
        prtSAB_Echelles_ECHEDI01P2_Form
    Else
        If blnNewPage Then
            frmElpPrt.prtNewPage
        End If
        prtSAB_Echelles_ECHEDI01P2_Form
    End If
'_______________________________________________________________________________
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 1100
    XPrt.Print "Report solde en valeur";
    XPrt.CurrentX = prtMinX + 100: XPrt.Print dateImp10(arrYECHREL0(0).ECHRELDVAL);
    X = Format$(arrYECHREL0(0).ECHRELSD, "### ### ### ##0.00")
    If arrYECHREL0(0).ECHRELSDS <> "C" Then
        XPrt.CurrentX = prtMinX + 5400 - XPrt.TextWidth(X)
    Else
        XPrt.CurrentX = prtMinX + 6900 - XPrt.TextWidth(X)
   End If
    XPrt.Print X;
'_______________________________________________________________________________

    For I = 1 To arrYECHREL0_Nb - 1
        prtSAB_Echelles_ECHEDI01P2_Line arrYECHREL0(I)
    Next I
'_______________________________________________________________________________
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, "B", 240)
'---------------------------------------------------------
    XPrt.CurrentY = XPrt.CurrentY + 50
        XPrt.FontBold = True
        XPrt.CurrentX = prtMinX + 1100
        XPrt.Print "Solde en valeur";
        XPrt.CurrentX = prtMinX + 100: XPrt.Print dateImp10(arrYECHREL0(arrYECHREL0_Nb).ECHRELDVAL);
        X = Format$(arrYECHREL0(arrYECHREL0_Nb).ECHRELSD, "### ### ### ##0.00")
        If arrYECHREL0(arrYECHREL0_Nb).ECHRELSDS <> "C" Then
            XPrt.CurrentX = prtMinX + 5400 - XPrt.TextWidth(X)
        Else
            XPrt.CurrentX = prtMinX + 6900 - XPrt.TextWidth(X)
       End If
        XPrt.Print X;
    XPrt.FontSize = 7
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
    X = Format$(arrYECHREL0(arrYECHREL0_Nb).ECHRELMDB, "### ### ### ###")
    XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(arrYECHREL0(arrYECHREL0_Nb).ECHRELMCR, "### ### ### ###")
    XPrt.CurrentX = prtMinX + 9900 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.FontSize = 8
    XPrt.CurrentY = XPrt.CurrentY - Height8_6
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    prtSAB_Echelles_ECHEDI01P2_Colonne
    If blnAvis Then
        Call prtSAB_Echelles_ECHEDI01P2_Avis(mYECHIMP0)
    End If
'_______________________________________________________________________________

    Exit Function

Error_Handle:
V = Error
Error_MsgBox:
MsgBox "prtSAB_Echelles_ECHEDI01P2_Relevé" & Error, vbCritical, V


End Function

Public Sub prtSAB_Echelles_ECHEDI01P2_Colonne()

XPrt.Line (prtMinX, mCurrenty_Top)-(prtMinX, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMinX + 1000, mCurrenty_Top)-(prtMinX + 1000, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMinX + 4000, mCurrenty_Top)-(prtMinX + 4000, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMinX + 7000, mCurrenty_Top)-(prtMinX + 7000, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMinX + 7600, mCurrenty_Top)-(prtMinX + 7600, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMinX + 10000, mCurrenty_Top)-(prtMinX + 10000, XPrt.CurrentY), prtLineColor
XPrt.Line (prtMaxX, mCurrenty_Top)-(prtMaxX, XPrt.CurrentY), prtLineColor

End Sub
