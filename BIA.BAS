Attribute VB_Name = "BIA"
Option Explicit

'''2003.12.15 compatibilité BIAS820I

Public frmRTF_Form_K2 As String, frmRTF_blnA5 As Boolean
Public frmRTF_FileName As String
Public frmRTF_Référence As String
Public frmRTF_UsrId_Origine As String
Public frmRTF_Caller As String
Public frmRTF_Buffer_Name As String

Public frmRTF_blnOK As Boolean
Public frmRTF_blnCourrier As Boolean
Public frmRTF_prtOrientation As Integer
'''2003.12.15 compatibilité BIAS820I



Public AccAutId As String * 10
Public paramBiaPgm  As String * 12
Public paramBiaPgmAut As String * 12
Public arrBiaPgm_Name() As String

Public Const constàAvis = "à Avis"
Public Const constàAvisDb = "à Avis Db"
Public Const constàAvisCr = "à Avis Cr"
Public Const constàMT100 = "à MT100"
Public Const constàMT202 = "à MT202"
Public Const constàVirXX = "à Vir XX"
Public Const constàVirAvis = "à Vir Avis"
Public Const constàVirCRI = "à Vir CRI"
Public Const constàRapprocher = "à Rappro"
Public Const constVersement = "Versement"
Public Const constRetrait = "Retrait"
Public Const constArbitrage = "Arbitrage"
Public Const constChange = "Change"
Public Const constRemiseChèques = "RemiseChèques"
Public Const constCaisse = "Caisse"
Public Const constChèque = "Chèque"
Public Const constJournalGuichet = "Guiche"
Public Const constRemboursementAnticipé = "Remboursement Anticipé"
Public Const constDemandeDeValidation = "Demande de validation"

Public Const constECRemise = "ECRemise"
Public Const constECPrésentation = "ECPrésent"
Public Const constECEchéance = "ECEchéance"
Public Const constECRappro = "ECRappro"

Public Const NostroBicVir = "SOGEFRPP   "
Public Const NostroBicADB = "SOGEFRPPADB"
Public Const NostroBicCRI = "SOGEFRPPTGV"
Public constMontantCRI As Currency
Public Const constCriCompteSG = "20000550011"
Public Const constCriCompteSG2 = "20000550029"
Public Const constCriCompteBDF = "20002550017"
Public Const constCriCompteUBAF = "20001550014"


Public constDeviseChange_MargeNormal As Double
Public constDeviseChange_MargeEnCompte As Double
Public constDeviseChange_MargePrivilégié As Double

Public AmjCptVeille As String * 8
Public DsysValueMin As String * 8
Public DsysValueMax As String * 8
Public DValPrevious As String * 8
Public DValNext As String * 8, DValNext2 As String * 8
Public DsysMinus2 As String * 8
Public IbanE As String
Public arrDevise(2) As typeDevise

'Public arrDevise(2) As typeDevise

Type typeAuthorization
    Consulter  As Boolean
    Saisir  As Boolean
    Valider As Boolean
    Comptabiliser As Boolean
    Rapprocher  As Boolean
    Swift  As Boolean
    Virement  As Boolean
    Avis  As Boolean
    Xspécial  As Boolean
End Type

Public CompteAut As typeAuthorization

Public CV_Euro As typeCV, CV_X1 As typeCV, CV_X2 As typeCV, CV_X3 As typeCV

Public Const constFTP_Dir = "\\Fr80826417\disk D\FTP\"



Public Const conststrGuichet_Compta = "Guichet_Comptabilisation"
Public Const conststrGuichet_Comptabilisé = "Guichet_Comptabilisé"

Public Const conststrTFlux_Compta = "Echéances_Comptabilisation"
Public Const conststrTFlux_Comptabilisé = "Echéances Comptabilisés"

Public paramComptaSld_Cpt_Import As String
Public paramComptaSld_Mvt_Import As String
Public paramComptaSld_Cpt_Export As String
Public paramComptaSld_Rac_Import As String
Public paramComptaExt_Cpt_Import As String
Public paramComptaExt_Mvt_Import As String
Public paramComptaExt_Cpt_Export As String

Public paramAmjOpérationMin As String * 8

Public paramCompteArbitrage  As String
Public paramCompteConversion  As String
'================================================================
Public Const Amj20011231 = "20011231"


Public paramIBM_BIA_INFO_Password As String
Public paramIBM_BIA_AUTO_Password As String
Public paramIBM_BIA_ODBC_Password As String
Public paramIBM_BIA_DWH_Password As String
Public paramIBM_BO_DWH_Password As String

Public Function param_AmjEchéanceS(lAmjEchéanceS As String) As String
Select Case lAmjEchéanceS
    Case "A"
        param_AmjEchéanceS = "Anniversaire"
    Case Else
        param_AmjEchéanceS = "Fin de Mois"
End Select

End Function


Public Function param_Statut(lStatut As String) As String
Dim iReturn As Integer
xElpTable.ID = "Param"
xElpTable.K1 = "Statut"
xElpTable.K2 = lStatut
xElpTable.Method = "Seek="
iReturn = tableElpTable_Read(xElpTable)
If iReturn <> 0 Then xElpTable.Name = lStatut
param_Statut = xElpTable.Name
End Function


Public Sub prtAdresse(mCurrenty As Integer, recCptInfo As typeCptInfo)
'-----------------------encadrement petit tirets---------
Dim Y2300 As Integer, Y2400 As Integer, Y4100 As Integer, Y4200 As Integer
Y2300 = mCurrenty + 2300
Y2400 = mCurrenty + 2400
Y4100 = mCurrenty + 4100
Y4200 = mCurrenty + 4200

XPrt.Line (5600, Y2300)-(5700, Y2300)
XPrt.Line (5600, Y2300)-(5600, Y2400)

XPrt.Line (10900, Y2300)-(11000, Y2300)
XPrt.Line (11000, Y2300)-(11000, Y2400)

XPrt.Line (5600, Y4200)-(5700, Y4200)
XPrt.Line (5600, Y4200)-(5600, Y4100)

XPrt.Line (10900, Y4200)-(11000, Y4200)
XPrt.Line (11000, Y4200)-(11000, Y4100)
XPrt.CurrentY = mCurrenty + 2400
XPrt.CurrentX = 5700
XPrt.FontBold = True
XPrt.Print recCptInfo.Intitulé;

'-----------------------------------------------------
XPrt.FontBold = False
If Trim(recCptInfo.Adresse2) <> "" Then
   XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = 5700
    XPrt.Print recCptInfo.Adresse2;
End If
'-----------------------------------3---------------
If Trim(recCptInfo.Adresse3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = 5700
    XPrt.Print recCptInfo.Adresse3;
End If
'----------------------------------4-------------------
If Trim(recCptInfo.Adresse4) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = 5700
    XPrt.Print recCptInfo.Adresse4;
End If

'-----------------------------------5------------------
If Trim(recCptInfo.Adresse5) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = 5700
    XPrt.Print recCptInfo.Adresse5;
End If
'------------------------------------6------------------
If Trim(recCptInfo.AdresseCP) <> "" _
Or Trim(recCptInfo.AdresseBD) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = 5700
    If Trim(recCptInfo.AdresseCP) <> "" Then XPrt.Print recCptInfo.AdresseCP & "  ";
    XPrt.Print recCptInfo.AdresseBD;
End If
'------------------------------------8------------------
If Trim(recCptInfo.AdressePays) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = 5700
   XPrt.Print recCptInfo.AdressePays;

End If
'------------------------------------------

End Sub

'---------------------------------------------------------
Public Sub prtSocInit()
'---------------------------------------------------------
prtFormType = "SOC"

frmElpPrt.prtInit
prtSoc

End Sub

'---------------------------------------------------------
Public Sub prtSoc()
'---------------------------------------------------------
Dim X As String, I As Integer
I = frmElpPrt.imgSocLogo.Width * 0.5
XPrt.PaintPicture frmElpPrt.imgSocLogo.Picture _
                , (prtMinX + prtMaxX - I) / 2, 0 _
                , I _
                , frmElpPrt.imgSocLogo.Height * 0.5
XPrt.PaintPicture frmElpPrt.imgSocSigle.Picture _
                , 7000, 300 _
                , frmElpPrt.imgSocSigle.Width * 0.5 _
                , frmElpPrt.imgSocSigle.Height * 0.5
                
XPrt.CurrentY = prtMinY
XPrt.FontBold = True
XPrt.FontSize = 11
frmElpPrt.prtCentré 2500, "BANQUE INTERCONTINENTALE ARABE"
'-----------------------------------------------------
XPrt.FontSize = 9
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + 250
frmElpPrt.prtCentré 2500, "67, avenue Franklin D. Roosevelt"
'------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 250
frmElpPrt.prtCentré 2500, "75008 PARIS"
'--------------------------------------------------------
XPrt.CurrentY = prtMaxY
XPrt.FontSize = 6
frmElpPrt.prtCentré (prtMaxX - prtMinX) / 2, "S.A. au capital de 158 100 000 Euros - R.C.Paris B 302590070 - L.B.E. 116 Tél: 01 53 76 62 62 - Téléfax: 01 42 89 09 59 - Télex: 644 030 BIAPA - Swift: BIARFRPP "
End Sub

Public Sub prtSocMini(mCurrenty As Integer, Amj As String)
XPrt.PaintPicture frmElpPrt.imgSocLogo.Picture _
                , 4850, mCurrenty + prtMinY _
                , frmElpPrt.imgSocLogo.Width * 0.5 _
                , frmElpPrt.imgSocLogo.Height * 0.5
XPrt.PaintPicture frmElpPrt.imgSocSigle.Picture _
                , 7500, mCurrenty + prtMinY + 200 _
                , frmElpPrt.imgSocSigle.Width * 0.3 _
              , frmElpPrt.imgSocSigle.Height * 0.3

'----------------------------------------------------
XPrt.CurrentY = mCurrenty + prtMinY
XPrt.FontBold = True
XPrt.FontSize = 8
frmElpPrt.prtCentré 2000, "BANQUE INTERCONTINENTALE ARABE"
'-----------------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + 250
frmElpPrt.prtCentré 2000, "67, avenue Franklin D. Roosevelt"
'------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 250
frmElpPrt.prtCentré 2000, "75008 PARIS"
'--------------------------------------------------------

XPrt.CurrentY = XPrt.CurrentY + 250 * 2
XPrt.CurrentX = 7500
XPrt.Print "Paris, le " & dateImp_jjMoisAAAA(Amj);

End Sub

'---------------------------------------------------------
Public Function PcIdUsrId()
'---------------------------------------------------------
Dim I As Integer, Nb As Integer

PcIdUsrId = "?"

Mid$(MsgTxt, 1, 12) = "ELPDTAQUSR  "
Mid$(MsgTxt, 25, 10) = Space$(10)
MsgTxtLen = 45
If pcIdUsrIdCtl Then
    Mid$(MsgTxt, 13, 12) = "PCID        "
    Mid$(MsgTxt, 35, 10) = UCase$(Elp.pcId)
Else
    Mid$(MsgTxt, 13, 12) = "USRID       "
    Mid$(MsgTxt, 35, 10) = UCase$(Elp.usrId)
End If

If IsNull(SndRcv()) Then

    SocId$ = mId$(MsgTxt, 45, 3)
    SocAgence$ = mId$(MsgTxt, 48, 3)
    strSocBdfE = mId$(MsgTxt, 51, 5)
    SocBdfE = Val(SocBdfE)
    strSocBdfG = mId$(MsgTxt, 56, 5)
    SocBdfG = Val(SocBdfG)
    socName = "Banque Intercontinentale Arabe (Paris)" 'Trim(Mid$(MsgTxt, 61, 40))
    Nb = Val(mId$(MsgTxt, 101, 2))

   XListBox.Clear
    MsgTxtIndex = 102
    For I = 1 To Nb
        XListBox.AddItem mId$(MsgTxt, MsgTxtIndex + 1, 10) _
            & mId$(MsgTxt, MsgTxtIndex + 11, 3) & mId$(MsgTxt, MsgTxtIndex + 14, 2) _
            & Chr$(9) & Trim(mId$(MsgTxt, MsgTxtIndex + 16, 4)) _
            & " " & Trim(mId$(MsgTxt, MsgTxtIndex + 20, 15)) _
            & " " & Trim(mId$(MsgTxt, MsgTxtIndex + 35, 15))
        MsgTxtIndex = MsgTxtIndex + 49
    Next I

'   If Not pcIdUsrIdCtl Then
'       PcIdUsrId = Null
'       XListBox.AddItem Elp.usrId & Chr$(9) & Elp.usrId
'   Else
        Select Case Nb
            Case Is = 0
                MsgBox Elp.usrId & ", Vous n'êtes pas branché !", vbCritical, "Liaison Serveur AS400"
                mainEnd
            Case Else
                PcIdUsrId = Null
            End Select
   End If
'End If
End Function

'---------------------------------------------------------
Public Function BiaPgm_Init()
'---------------------------------------------------------
Dim Nb As Integer, X As String

ReDim arrBiaPgm_Name(20)
BiaPgm_Init = "?"

usrSituationCompte_Forçage = False
usrService_DisplayAll = False
            

Nb = -1
XListBox.Clear
XListBox.Visible = True
XListBox.Height = 200
XLabel.Visible = True
XLabel.Caption = "Menu"

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.ID = paramBiaPgmAut
recElpTable.K1 = Elp.usrId
recElpTable.Method = "Seek>="
recElpTable.Err = 0
Do
    recElpTable.Err = tableElpTable_Read(recElpTable)
    If recElpTable.Err = 0 Then
        If paramBiaPgmAut <> recElpTable.ID Or Trim(Elp.usrId) <> Trim(recElpTable.K1) Then
            recElpTable.Err = 9996
        Else
            X = Trim(recElpTable.K2)
            If X = "$usr_Forçage" Then usrSituationCompte_Forçage = True
            If X = "$usr_Service" Then usrService_DisplayAll = True
            
           xElpTable.Method = "Seek="
           xElpTable.ID = paramBiaPgm
            xElpTable.K1 = recElpTable.K2
            xElpTable.K2 = ""
            xElpTable.Err = tableElpTable_Read(xElpTable)
            If xElpTable.Err = 0 And frmElp_Caption = Trim(mId$(xElpTable.Memo, 21, 20)) Then
                XListBox.AddItem recElpTable.K2 & Chr$(9) & Chr$(9) & xElpTable.Name
                Nb = Nb + 1
                If Nb >= UBound(arrBiaPgm_Name) Then ReDim arrBiaPgm_Name(Nb + 10)
                arrBiaPgm_Name(Nb) = xElpTable.K1 & " : " & xElpTable.Name
            End If
            recElpTable.Method = "Seek>"
        End If
    End If
Loop While recElpTable.Err = 0

XListBox.AddItem "X_Reset     " & Chr$(9) & Chr$(9) & "réplication BiaSrv"
Nb = Nb + 1
If Nb >= UBound(arrBiaPgm_Name) Then ReDim arrBiaPgm_Name(Nb + 10)
arrBiaPgm_Name(Nb) = "X_Reset     " & " : " & "réplication BiaSrv"

ReDim arrBiaPgm_Name(Nb + 1)
Elp_ResizeControl XListBox
End Function

'---------------------------------------------------------
Public Sub srvDevise()
'---------------------------------------------------------
Dim K As Integer, r As Integer, Dmax As String * 8
Dim X, Y
Dim recdictio As typeDictio

recDictioInit recdictio

recdictio.Method = "Seek>=      "
recdictio.DicRub = 888
recdictio.DicCode = ""
Dmax = "00000000"

X = dbDictioRead(recdictio)
recdictio.Method = "MoveNext    "
Do While Trim(recdictio.Err) = "0" _
     And recdictio.DicRub = "888"

    If recdictio.DicAmj > Dmax Then Dmax = recdictio.DicAmj
    
    X = dbDictioRead(recdictio)

Loop

recdictio.obj = "SRVDEVISE   "
recdictio.Method = "DevCours    "
recdictio.DicCode = "000"
recdictio.DicAmj = Dmax

arrDictio(0) = recdictio
arrDictioSuite = True

Do While arrDictioSuite
    srvDictioMon recdictio
    arrDictio(0) = arrDictio(arrDictioNb)
Loop

For K = 1 To arrDictioNb
    arrDictio(K).Method = constUpdate
    arrDictio(K).DicAmj = Format$(DSys, "00000000")
'!!!Mid$(arrDictio(K).DicTxt, 1, 3) = StrConv(Mid$(arrDictio(K).DicTxt, 1, 3), vbProperCase)
    arrDictio(K).DicLib = StrConv(arrDictio(K).DicLib, vbProperCase)
    arrDictio(0) = arrDictio(K)
    r = tableDictioRead(arrDictio(0))
    
    If r = 9923 Then
        r = 0
        arrDictio(K).Method = constAddNew
    End If
    
    If r = 0 Then
        Call dbDictioUpdate(arrDictio(K))
        
        arrDictio(K).Method = constUpdate
        arrDictio(K).DicRub = 889
        X = arrDictio(K).DicCode
        arrDictio(K).DicCode = mId$(arrDictio(K).DicTxt, 1, 3)
        Mid$(arrDictio(K).DicTxt, 1, 3) = X
        arrDictio(0) = arrDictio(K)
        r = tableDictioRead(arrDictio(0))
        If r = 9923 Then
            r = 0
            arrDictio(K).Method = constAddNew
        End If
    
        If r = 0 Then
            Call dbDictioUpdate(arrDictio(K))
        End If
    End If
Next K

End Sub

'---------------------------------------------------------
Public Sub srvDevise_1999()
 '---------------------------------------------------------
Dim K As Integer, r As Integer
Dim X, Y
Dim recdictio As typeDictio, recdictio889 As typeDictio
Dim CV1 As typeCV, CV2 As typeCV, CV3 As typeCV
Dim dblX As Double, Conversion As String

CV_Init CV1
CV1.OpéAmj = DSys
CV2 = CV1
CV3 = CV_Euro

CV2.DeviseIso = "FRF"
CV1.AchatVente = " "
CV2.AchatVente = " "
CV1.Normal = "P"
CV2.Normal = "P"

recDictioInit recdictio

recdictio.Method = "Seek>=      "
recdictio.DicRub = 888
recdictio.DicCode = ""

X = dbDictioRead(recdictio)
recdictio.Method = "MoveNext    "
Do While Trim(recdictio.Err) = "0" _
     And recdictio.DicRub = "888"
    
    CV1.DeviseIso = mId$(recdictio.DicTxt, 1, 3)
    CV1.Montant = 1000000
    
    If Not IsNull(CV_Transitoire(CV1, CV2, CV3, Conversion)) Then
        dblX = 0
    Else
        dblX = CV2.Montant / CV1.Montant
    End If
    
    Mid$(recdictio.DicTxt, 4, 13) = Format$(dblX, "00000.0000000")
    Mid$(recdictio.DicTxt, 9, 1) = "."
    Mid$(recdictio.DicTxt, 17, 8) = DSys
    
    recdictio.Method = constUpdate
    arrDictio(0) = recdictio
    r = tableDictioRead(arrDictio(0))
    If r = 0 Then
        Call tableDictioUpdate(recdictio)
    
        recdictio889 = recdictio
        recdictio889.Method = constUpdate
        recdictio889.DicRub = 889
        X = recdictio889.DicCode
        recdictio889.DicCode = mId$(recdictio889.DicTxt, 1, 3)
        Mid$(recdictio889.DicTxt, 1, 3) = X
        arrDictio(0) = recdictio889
        r = tableDictioRead(arrDictio(0))
        If r = 9923 Then
            r = 0
            recdictio889.Method = constAddNew
        End If
    
        If r = 0 Then
            Call dbDictioUpdate(recdictio889)
        End If
     End If
   
    recdictio.Method = "Seek>    "
    X = dbDictioRead(recdictio)

Loop


End Sub


'---------------------------------------------------------
Public Sub mainSoc()
'---------------------------------------------------------
Dim V As String, X As String


arrDictioNbMax = 1
arrDictioNb = 0
ReDim arrDictio(1)

mainSocExe

recElpTable_Init recElpTable
recElpTable.ID = "Param"
recElpTable.Method = "Seek>="
Do
    recElpTable.Err = tableElpTable_Read(recElpTable)
    If recElpTable.Err = 0 Then
        If "Param       " <> recElpTable.ID Then
            recElpTable.Err = 9996
        Else
            Select Case Trim(recElpTable.K1)
                Case "BiaPgm"
                    Select Case Trim(recElpTable.K2)
                        Case "Programmes": paramBiaPgm = Trim(recElpTable.Memo)
                        Case "Autorisation": paramBiaPgmAut = Trim(recElpTable.Memo)
                    End Select
                 Case "Compte"
                    Select Case Trim(recElpTable.K2)
                        Case "Arbitrage": paramCompteArbitrage = Trim(recElpTable.Memo)
                        Case "Conversion": paramCompteConversion = Trim(recElpTable.Memo)
                    End Select
                Case "XUsrId"
                    If Trim(mId$(recElpTable.K2, 1, 10)) = Trim(Elp.usrId) Then
                        usrId = UCase$(Trim(recElpTable.Memo))
                        Elp.usrId = usrId
                        Xcom_UsrId usrId
                    End If
                    

            End Select
                        
            recElpTable.Method = "Seek>"
        End If
    End If
    
Loop While recElpTable.Err = 0


constMontantCRI = 5000000
constDeviseChange_MargeNormal = 0.03
constDeviseChange_MargePrivilégié = 0.006
constDeviseChange_MargeEnCompte = 0.01

V = DateAdd("d", -7, Now)
DsysValueMin = Year(V)
Mid$(DsysValueMin, 5, 2) = Format$(Month(V), "00")
Mid$(DsysValueMin, 7, 2) = Format$(Day(V), "00")
V = DateAdd("d", 7, Now)
DsysValueMax = Year(V)
Mid$(DsysValueMax, 5, 2) = Format$(Month(V), "00")
Mid$(DsysValueMax, 7, 2) = Format$(Day(V), "00")
DsysMinus2 = dateElp("Ouvré", -2, DSys)
AmjCptVeille = dateElp("Ouvré", -1, DSys)

DValPrevious = dateElp("Ouvré", -1, DSys)
DValNext = dateElp("Ouvré", 1, DSys)
DValNext2 = dateElp("Ouvré", 2, DSys)

paramAmjOpérationMin = mId$(DSys, 1, 6) & "01"
If mId$(DSys, 7, 2) = "01" Then paramAmjOpérationMin = dateElp("Ouvré", -1, DSys)

socName = "Banque Intercontinentale Arabe"
SocBicId = "BIARFRPP"
SocBicIdNostro = "BIARFRPPNOS"
SocRibDom = "B INTERCONT ARABE PARIS"
socTéléphone = "(33) 01 53 76 62 62"
Set XListBox = frmElp.lstMain
Set XLabel = frmElp.lblMain


CV_Euro.CoursAmjMin = "19990101"
CV_Init CV_Euro: CV_X1 = CV_Euro: CV_X2 = CV_Euro: CV_X3 = CV_Euro

CV_Euro.DeviseIso = "EUR"
CV_Euro.DeviseN = "978"
CV_Euro.DeviseLibellé = "Euro"

tableDictio_Open
tableDeviseChange_Open
tableDeviseCompta_Open

XListBox.Visible = False

If elpSrvXcom = "" Then
    XListBox.Clear
    XListBox.AddItem Elp.usrId
    SocId$ = "001"
    SocAgence$ = "001"

Else
    If Not IsNull(PcIdUsrId()) Then
        End
    Else
        srvDevise
        dbDeviseCompta_Replication
        dbDeviseChange_Replication  '19990101
        srvDevise_1999
    End If
End If

If XListBox.ListCount = 1 Then
    XListBox.ListIndex = 0
    X = Space$(100)
    X = Trim(XListBox.Text)
    Elp.usrId = mId$(X, 1, 10)
    usrService = mId$(X, 11, 3)
    usrGestionnaire = mId$(X, 14, 2)
    usrName = mId$(X, 17, 34)
    BiaPgm_Init
    usrDRH usrIdNT
    
Else

    XLabel.Caption = "Qui êtes-vous ?"
    XLabel.ForeColor = errUsr.ForeColor
    XLabel.Visible = True
    XListBox.Visible = True
End If

recPériodicité_Init
recStatut_Init

If mId$(Elp.SrvDtaqIn, 1, 1) <> "P" Then
    socName = Elp.SrvDTaqOut
    focusUsr.BackColor = vbYellow
    MouseMoveUsr.ForeColor = vbYellow
    MouseMoveUsr.BackColor = vbYellow
End If
'''ElpChat_Init

End Sub

'---------------------------------------------------------
'Public Function ctlGestionnaire(strGestionnaire As String, strBiaTyp As String) As Boolean
'---------------------------------------------------------

'ctlGestionnaire = True
'If strGestionnaire = "60" Then              ' gestionnaire personnel
'    Select Case strBiaTyp
'        Case "050", "080", "081", "006"           ' prêts personnels, immobiliers
'            If usrGestionnaire <> "00" _
'            And usrGestionnaire <> "60" _
'            And usrGestionnaire <> "70" Then
'                ctlGestionnaire = False
'            End If
'        Case Else
'            If usrGestionnaire <> "00" _
'            And usrGestionnaire <> "22" _
'            And usrGestionnaire <> "60" _
'            And usrGestionnaire <> "70" Then
'                ctlGestionnaire = False
'            End If
'    End Select
'End If
'End Function

'---------------------------------------------------------
Public Function ctlGestionnaire_New(lNuméro As String, lGestionnaire As String, llstrBiaTyp As String) As Boolean
'---------------------------------------------------------

ctlGestionnaire_New = True
If mId$(lNuméro, 1, 5) = usrRacine Then

Else
    If usrGestionnaire = "98" Then
        ctlGestionnaire_New = False
    Else
        If lGestionnaire = "60" Then              ' gestionnaire personnel
            Select Case llstrBiaTyp
                Case "050", "080", "081", "006"           ' prêts personnels, immobiliers
                    If usrGestionnaire <> "00" _
                    And usrGestionnaire <> "60" _
                    And usrGestionnaire <> "70" Then
                        ctlGestionnaire_New = False
                    End If
                Case Else
                    If usrGestionnaire <> "00" _
                    And usrGestionnaire <> "22" _
                    And usrGestionnaire <> "60" _
                    And usrGestionnaire <> "70" Then
                        ctlGestionnaire_New = False
                    End If
            End Select
        End If
    End If
End If
End Function

Public Function Compte_Display(ByVal X As String) As String
Select Case Val(X)
    Case Is <= 99999: Compte_Display = Format$(X, "### ##")
    Case Is <= 99999999: Compte_Display = Format$(X, "### ### ##")
    Case Else: Compte_Display = Format$(X, "##### ### ## #")
End Select
End Function
Public Function CompteSAB_Imp(ByVal X As String) As String

CompteSAB_Imp = X
'If Len(Trim(X)) = 12 Then
'    CompteSAB_Imp = mId$(X, 1, 6) & "." & mId$(X, 7, 3) & "." & mId$(X, 10, 3)
'Else
'    CompteSAB_Imp = mId$(X, 1, 5) & "." & mId$(X, 6, 3) & "." & mId$(X, 9, 3) & " " & mId$(X, 12, 9)
'End If

End Function

Public Function Compte_Imp(ByVal X As String) As String
If Val(X) < 99999999 Then
    If Val(X) = 0 Then
        Compte_Imp = ""
    Else
        Compte_Imp = Format$(Val(X), "@@@ @@@.@@")
    End If
Else
    Compte_Imp = Format$(Val(X), "@@@@@.@@@.@@.@")
End If
End Function


Public Sub BiaPgmAut_Init(X As String, recAut As typeAuthorization)

recElpTable_Init recElpTable
recElpTable.ID = paramBiaPgmAut
recElpTable.K1 = Elp.usrId
recElpTable.K2 = X
recElpTable.Method = "Seek="
recElpTable.Err = tableElpTable_Read(recElpTable)
If recElpTable.Err <> 0 Then recElpTable.Memo = Space$(20)
If recElpTable.Err <> 0 And blnJPL Then recElpTable.Memo = "XXXXXXXXXX": MsgBox "Autorisation", vbExclamation, "BIA.bas"

recAut.Consulter = IIf(mId$(recElpTable.Memo, 1, 1) = "X", True, False)
recAut.Saisir = IIf(mId$(recElpTable.Memo, 2, 1) = "X", True, False)
recAut.Valider = IIf(mId$(recElpTable.Memo, 3, 1) = "X", True, False)
recAut.Comptabiliser = IIf(mId$(recElpTable.Memo, 4, 1) = "X", True, False)
recAut.Rapprocher = IIf(mId$(recElpTable.Memo, 5, 1) = "X", True, False)
recAut.Swift = IIf(mId$(recElpTable.Memo, 6, 1) = "X", True, False)
recAut.Virement = IIf(mId$(recElpTable.Memo, 7, 1) = "X", True, False)
recAut.Avis = IIf(mId$(recElpTable.Memo, 8, 1) = "X", True, False)
recAut.Xspécial = IIf(mId$(recElpTable.Memo, 9, 1) = "X", True, False)

End Sub

Public Sub prtSocMiniFin()
XPrt.FontSize = 6
frmElpPrt.prtCentré (prtMaxX - prtMinX) / 2, "S.A. au capital de 158 100 000 Euros - R.C.Paris B 302590070 - L.B.E. 116 Tél: 01 53 76 62 62 - Téléfax: 01 42 89 09 59 - Télex: 644 030 BIAPA - Swift: BIARFRPP "

End Sub


Public Sub mainSoc_Environment()

paramFolder_Master = "\\FR11024427\.BiaSrv" ' "\\FR11024427\Bia_Data\Appli\BiaSrv"
paramFolder_Local = "C:\BiaSrv"

If Dir(paramFolder_Local & "\BiaSigno.bmp") = "" Then
    paramFolder_Local = "D:\BiaSrv"
    If Dir(paramFolder_Local & "\BiaSigno.bmp") = "" Then
        MsgBox "MANQUE :" & paramFolder_Local, vbCritical, "BIA_mainSoc_Environment"
        ''End
    End If
End If

DataBase_Open = ""
DataBase_Local = UCase$(paramFolder_Local & "\Bia.mdb")
DataBase_Master = UCase$(paramFolder_Master & "\Bia.mdb")

Elp.SrvObj = "ELPDTAQ"
Elp.pcId = "FR"
Elp.SrvType = "AS400"
'20031215 jpl Elp.SrvId = "AS400BIA"
Elp.SrvId = "SABA"

Elp.SrvDtaqLib = "BIADTAQ"

Select Case App_EXEName
    Case "BIA", "BIACPT", "BIAODBC", "BIACD", "BIA_C", "BIACPT_C":
            Elp.SrvDtaqIn = "PB000001"
            Elp.SrvDTaqOut = "PB000000"
    Case Else
            Elp.SrvDtaqIn = "PC000001"
            Elp.SrvDTaqOut = "PC000000"
End Select

pcIdUsrIdCtl = False
strSocSignon = paramFolder_Local & "\BiaSigno.bmp"
imgSocLogo = paramFolder_Local & "\BiaLogo.bmp"
imgSocSigle = paramFolder_Local & "\BiaSigle.bmp"
imgGuichet = paramFolder_Local & "\BiaGuichet.bmp"
prtFontName = "arial"
elpSrvXcom = "CAV4"
'strSocSignon = "\\Fr80826417\disk D\BiaSrv\BiaSigno.bmp"
'imgSocLogo = "\\Fr80826417\disk D\BiaSrv\BiaLogo.bmp"
'imgSocSigle = "\\Fr80826417\disk D\BiaSrv\BiaSigle.bmp"
'imgGuichet = "\\Fr80826417\disk D\BiaSrv\BiaGuichet.bmp"
'prtFontName = "arial"
'DataBaseName = "\\Fr80826417\disk D\BiaSrv\Bia.mdb"

End Sub

Public Sub Compte_BiaTyp(Xcompte As String, mBiatyp As String)
Mid$(Xcompte, 6, 3) = mId$(mBiatyp, 1, 3)
Compte_BiaClé Xcompte
End Sub

Public Sub Compte_BiaClé(Xcompte As String)
Dim N As Integer
N = CInt(mId$(Xcompte, 1, 1)) * 11 _
  + CInt(mId$(Xcompte, 2, 1)) * 10 _
  + CInt(mId$(Xcompte, 3, 1)) * 9 _
  + CInt(mId$(Xcompte, 4, 1)) * 8 _
  + CInt(mId$(Xcompte, 5, 1)) * 7 _
  + CInt(mId$(Xcompte, 6, 1)) * 6 _
  + CInt(mId$(Xcompte, 7, 1)) * 5 _
  + CInt(mId$(Xcompte, 8, 1)) * 4 _
  + CInt(mId$(Xcompte, 9, 1)) * 3 _
  + CInt(mId$(Xcompte, 10, 1)) * 2
  
N = 10 - (N Mod 10)
If N = 10 Then
    Mid$(Xcompte, 11, 1) = "0"
Else
    Mid$(Xcompte, 11, 1) = CStr(N)
End If

End Sub



Public Sub usrDRH(lName)
recDRH_Init xDRH
xDRH.Method = "SeekLRI"
xDRH.RéfInterne = UCase$(lName)
If IsNull(srvDRH_SeekX(xDRH)) Then
    usrCompte = xDRH.Compte
    usrRacine = mId$(usrCompte, 1, 5)
End If

End Sub

Public Function CompteSituation_Saisie(lCompte As typeCompte, blnCompteSituation_Saisie As Boolean)
Dim X As String, blnMsgBox As Boolean, K As Integer

CompteSituation_Saisie = Null
blnCompteSituation_Saisie = True

K = vbExclamation
blnMsgBox = True
Select Case lCompte.Situation
    Case " ": blnMsgBox = False
    Case "E": X = " : Conversion EURO":
    Case "F": X = " : Autorisé Devise IN"
    Case "G": X = " : Réservé comptaG"
    Case "A": X = " : Annulé": K = vbCritical: blnCompteSituation_Saisie = False
    Case "B": X = " : Bloqué"
    Case "D": X = " : Décédé"
    Case "R": X = " : Regroupement": K = vbCritical: blnCompteSituation_Saisie = False
    Case Else: X = " : " & lCompte.Situation: K = vbCritical: blnCompteSituation_Saisie = False
End Select

If blnMsgBox Then
    X = "Compte : " & lCompte.Devise & "_" & Compte_Imp(lCompte.Numéro) & X
    Call MsgBox(lCompte.Intitulé & Chr$(13) & Chr$(13) & X, K, "Situation du compte (Saisie)")
End If

If K = vbCritical Then CompteSituation_Saisie = X


End Function
Public Function CompteSituation_Validation(lCompte As typeCompte, blnCompteSituation_Validation As Boolean, blnCompteSituation_Forçage As Boolean)
Dim X As String, xR As String, blnMsgBox As Boolean, K As Integer

CompteSituation_Validation = Null
blnCompteSituation_Validation = True
blnCompteSituation_Forçage = False

K = vbQuestion + vbYesNo
blnMsgBox = True

Select Case lCompte.Situation
    Case " ": blnMsgBox = False
    Case "E": X = " : Conversion EURO":
    Case "F": X = " : Autorisé Devise IN"
    Case "G": X = " : Réservé comptaG"
    Case "A": X = " : Annulé": K = vbCritical
    Case "B": X = " : Bloqué"
                        If usrSituationCompte_Forçage Then
                            blnCompteSituation_Forçage = True
                        Else
                            K = vbCritical
                        End If
                        
    Case "D": X = " : Décédé"
                         If usrSituationCompte_Forçage Then
                            blnCompteSituation_Forçage = True
                        Else
                            K = vbCritical
                        End If
   Case "R": X = " : Regroupement": K = vbCritical
    Case Else: X = " : " & lCompte.Situation: K = vbCritical
End Select

If blnMsgBox Then
    X = "Compte : " & lCompte.Devise & "_" & Compte_Imp(lCompte.Numéro) & X
    
    If K = vbCritical Then
        xR = MsgBox(lCompte.Intitulé & Chr$(13) & Chr$(13) & X, K, "Situation du compte (Validation)")
    Else
        xR = MsgBox(lCompte.Intitulé & Chr$(13) & Chr$(13) & X & Chr$(13) & Chr$(13) & "Confirmez-vous ?", K, "Situation du compte (Validation)")
        If xR = vbNo Then
           K = vbCritical
        Else
            blnCompteSituation_Validation = True
        End If
    End If
    
End If

If K = vbCritical Then CompteSituation_Validation = X: blnCompteSituation_Validation = False: blnCompteSituation_Forçage = False


End Function


Public Sub CDDossier_Référence_Libellé(lLibellé As String, lCDDossier_Référence As String)
Dim I1 As Integer, I2 As Integer, lMax As Integer, xLibellé As String, K1 As Integer, K2 As Integer
Dim blnOk As Boolean, blnNum As Boolean

lCDDossier_Référence = "000000"

xLibellé = Trim(lLibellé)
lMax = Len(xLibellé)
I1 = InStr(1, xLibellé, "CD")
If I1 > 0 Then
    blnOk = False
    K2 = lMax
    For I2 = I1 + 2 To lMax
        blnNum = IIf(IsNumeric(mId$(xLibellé, I2, 1)), True, False)
        If Not blnOk Then
            If blnNum Then K1 = I2: blnOk = True
        Else
            If Not blnNum Then K2 = I2 - 1: Exit For
        End If
    Next I2
    lCDDossier_Référence = Format$(Val(mId$(xLibellé, K1, K2)), "000000")

End If
End Sub

Public Sub CV_Calc_XXXXXX(CV1 As typeCV, CV2 As typeCV)
Dim errorV

CV2.Montant = CV1.Montant

End Sub
