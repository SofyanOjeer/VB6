Attribute VB_Name = "BIA"
Option Explicit
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
Public Const constRemboursementAnticipé = "Remboursement Anticipé"

Public Const NostroBicVir = "SOGEFRPP   "
Public Const NostroBicADB = "SOGEFRPPADB"
Public Const NostroBicCRI = "SOGEFRPPTGV"
Public constMontantCRI As Currency
Public Const constCriCompteSG = "20000550011"
Public Const constCriCompteBDF = "20002550017"
Public Const constCriCompteUBAF = "20001550014"


Public constDeviseChange_MargeNormal As Double
Public constDeviseChange_MargeEnCompte As Double
Public constDeviseChange_MargePrivilégié As Double

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

Public CV_Euro As typeCV, CV_X1 As typeCV, CV_X2 As typeCV, CV_X3 As typeCV

Public Const constFTP_Dir = "S:\FTP\"

Public Const constLucaRisques_AS400Trf = "LRCDRTRF"
Public Const constLucaRisques_AS400Ext = "LRCDRCL"
Public Const constLucaRisques_AS400BdfTrf = "LRCDRBDFTR"
Public Const constLucaRisques_LrCdrAller = "LRCDRALLER"
Public Const constLucaRisques_SendFilename = "S:\Pelint\Data\Send\RS01\CDR"
Public Const constLucaRisques_ReceiveFilename = "S:\Pelint\Data\Receive\RS02\CDR"

'Public Const constLucaRisques_Directory = "R:\Luca_Risques7\Engine\Files\"
Public Const constLucaRisques_LrRisque = "R:\Luca_Risques7\Engine\Files\CRRisque01.IDD"
Public Const constLucaRisques_LrRetris = "R:\Luca_Risques7\Engine\Files\CRRetris02.dat"
Public Const constLucaRisques_LrSgnBnf = "R:\Luca_Risques7\Engine\Files\CRSgnBnf01.dat"
Public Const constLucaRisques_Start = "R:\Luca_Risques7\Bia\LrCdr_Start.bat "
Public Const constLucaRisques_End = "R:\Luca_Risques7\Bia\LrCdr_End.bat "
Public Const constLucaRisques_LrCdr_FileName = "R:\Luca_Risques7\Engine\Files\CRENTSTD01.dat"
Public Const constLucaRisques_LrTiers_FileName = "R:\Luca_Risques7\Engine\Files\CRCLICLI"
Public Const constLucaRisques_Msg_FileName = "R:\Luca_Risques7\Bia\LrCdr_Msg.txt "
Public Const constLucaRisques_LrCdrBdf_FileName = "R:\Luca_Risques7\Engine\Files\CRRETBDF01.dat"
Public Const constLucaRisques_LrCdrAller_FileName = "R:\Luca_Risques7\Engine\Files\CRDECBDF03.tmp"
Public Const constLucaRisques_PrintSopra220 = "R:\Luca_Risques7\Engine\Files\CRCRB22001.lis"
Public Const constLucaRisques_PrintSopra400 = "R:\Luca_Risques7\Engine\Files\CRCRB40001.lis"
Public Const constLucaRisques_PrintSopra470 = "R:\Luca_Risques7\Engine\Files\CRCRB47001.lis"
Public Const constLucaRisques_PrintSopra490 = "R:\Luca_Risques7\Engine\Files\CRCRB49001.lis"
Public Const constLucaRisques_PrintSopra870 = "R:\Luca_Risques7\Engine\Files\CRCRB87001.lis"
Public Const constLucaRisques_PrintSopra880 = "R:\Luca_Risques7\Engine\Files\CRCRB88001.lis"

Public Const constLrBafi_AS400Trf = "LRBAFITRF"
Public Const constLrBafi_AS400Ext = "LRBAFICL"
Public Const constLrBafi_Engine_Start = "R:\LucaReport\Bia\lr97_Engine_Start.bat "
Public Const constLrBafi_Engine_End = "R:\LucaReport\Bia\lr97_Engine_End.bat "
Public Const constLrBafi_LrEstd_FileName = "R:\LucaReport\Bia\EstdNoEc.S"
Public Const constLrBafi_LrSolde_FileName = "R:\LucaReport\Bia\SolCgeEp.S"
Public Const constLrBafi_LrBafiMsg_FileName = "R:\LucaReport\Bia\LrBafiMsg.S"
Public Const constLrBafi_PilFab_FileName = "R:\LucaReport\Server\Send\Data\PilFab"
Public Const constLrBafi_Descri_FileName = "R:\LucaReport\Engine\Pc_Out\Des"
Public Const constLrBafi_Archive = "R:\LucaReport\Bia_Archive\"
Public Const constLrBafi_àTransmettre = "S:\Pelint\Data\Send\CB04\" ''''"R:\LucaReport\Bia_àTransmettre\"
Public Const constLrBafi_Bia = "R:\LucaReport\Bia\"
Public Const constLrBafi_Bia_Filename = "BalPa*.*;Cad*.*;IME*.*;Inter*.*;Sit*.*"

Public Const conststrGuichet_Compta = "Guichet_Comptabilisation"
Public Const conststrGuichet_Comptabilisé = "Guichet_Comptabilisé"

Public Const conststrPrêt_Compta = "Prêts_Comptabilisation"
Public Const conststrPrêt_Comptabilisé = "Prêts_Comptabilisés"


Type typeXcom
   SrvObj       As String * 12
   SrvMethod    As String * 12
   SrvErr       As String * 10
   usrId       As String * 10
   pcId        As String * 10
   SrvType     As String * 10
   SrvId       As String * 10
   SrvDtaqLib  As String * 10
   SrvDtaqIn   As String * 10
   SrvDTaqOut  As String * 10
   SrvDTaqLen  As String * 5
   jplFree        As String * 5
End Type

    
Type typeCV
    DeviseIso     As String * 3
    DeviseN       As String * 3
    DeviseLibellé As String * 20
    Cours         As Double
    CoursAmj      As String * 8
    maxD          As String * 1
    EuroIn        As Boolean
    CotationCertain As Boolean
    
    Montant       As Currency
    OpéAmj        As String * 8
    CoursAmjMin   As String * 8
    AchatVente    As String * 1
    Normal        As String * 1
    CoursCompta   As String * 1

End Type

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

Public Function param_Statut(lStatut As String) As String

xElpTable.Id = "Param"
xElpTable.K1 = "Statut"
xElpTable.K2 = lStatut
xElpTable.Method = "Seek="
iReturn = tableElpTable_Read(xElpTable)
If iReturn <> 0 Then xElpTable.Name = lStatut
param_Statut = xElpTable.Name
End Function

Public Function param_AmjEchéanceS(lAmjEchéanceS As String) As String
Select Case lAmjEchéanceS
    Case "A"
        param_AmjEchéanceS = "Anniversaire"
    Case Else
        param_AmjEchéanceS = "Fin de Mois"
End Select

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
frmElpPrt.prtCentré (prtMaxX - prtMinX) / 2, "S.A. au capital de 90 000 000 Euros - R.C.Paris B 302590070 - L.B.E. 116 Tél: 01 53 76 62 62 - Téléfax: 01 42 89 09 59 - Télex: 644 030 BIAPA - Swift: BIARFRPP "
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
Dim Nb As Integer
ReDim arrBiaPgm_Name(20)
BiaPgm_Init = "?"
Nb = -1
XListBox.Clear
XListBox.Visible = True
XListBox.Height = 200
XLabel.Visible = True
XLabel.Caption = "Menu"

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Id = paramBiaPgmAut
recElpTable.K1 = Elp.usrId
recElpTable.Method = "Seek>="
recElpTable.Err = 0
Do
    recElpTable.Err = tableElpTable_Read(recElpTable)
    If recElpTable.Err = 0 Then
        If paramBiaPgmAut <> recElpTable.Id Or Trim(Elp.usrId) <> Trim(recElpTable.K1) Then
            recElpTable.Err = 9996
        Else
           xElpTable.Method = "Seek="
           xElpTable.Id = paramBiaPgm
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
recElpTable.Id = "Param"
recElpTable.Method = "Seek>="
Do
    recElpTable.Err = tableElpTable_Read(recElpTable)
    If recElpTable.Err = 0 Then
        If "Param       " <> recElpTable.Id Then
            recElpTable.Err = 9996
        Else
            Select Case Trim(recElpTable.K1)
                Case "BiaPgm"
                    Select Case Trim(recElpTable.K2)
                        Case "Programmes": paramBiaPgm = Trim(recElpTable.Memo)
                        Case "Autorisation": paramBiaPgmAut = Trim(recElpTable.Memo)
                    End Select
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

DValPrevious = dateElp("Ouvré", -1, DSys)
DValNext = dateElp("Ouvré", 1, DSys)
DValNext2 = dateElp("Ouvré", 2, DSys)

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

'If IsNull(PcIdUsrId()) Then
XListBox.Clear
XListBox.AddItem Elp.usrId
    If XListBox.ListCount = 1 Then
        XListBox.ListIndex = 0
        X = Space$(100)
        X = Trim(XListBox.Text)
        Elp.usrId = mId$(X, 1, 10)
        usrService = mId$(X, 11, 3)
        usrGestionnaire = mId$(X, 14, 2)
        usrName = mId$(X, 17, 34)
        BiaPgm_Init

    Else
        XLabel.Caption = "Qui êtes-vous ?"
        XLabel.ForeColor = errUsr.ForeColor
        XLabel.Visible = True
        XListBox.Visible = True
    End If
'End If

'srvDevise

'dbDeviseCompta_Replication
'dbDeviseChange_Replication  '19990101
'srvDevise_1999
recPériodicité_Init
recStatut_Init

End Sub

'---------------------------------------------------------
Public Function ctlGestionnaire(strGestionnaire As String, strBiaTyp As String) As Boolean
'---------------------------------------------------------

ctlGestionnaire = True
If strGestionnaire = "60" Then              ' gestionnaire personnel
    Select Case strBiaTyp
        Case "050", "080", "081"            ' prêts personnels, immobiliers
            If usrGestionnaire <> "00" _
            And usrGestionnaire <> "60" _
            And usrGestionnaire <> "70" Then
                ctlGestionnaire = False
            End If
        Case Else
            If usrGestionnaire <> "00" _
            And usrGestionnaire <> "22" _
            And usrGestionnaire <> "60" _
            And usrGestionnaire <> "70" Then
                ctlGestionnaire = False
            End If
    End Select
End If
End Function

Public Function Compte_Display(ByVal X As String) As String
Select Case Val(X)
    Case Is <= 99999: Compte_Display = Format$(X, "### ##")
    Case Is <= 99999999: Compte_Display = Format$(X, "### ### ##")
    Case Else: Compte_Display = Format$(X, "##### ### ## #")
End Select
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
recElpTable.Id = paramBiaPgmAut
recElpTable.K1 = Elp.usrId
recElpTable.K2 = X
recElpTable.Method = "Seek="
recElpTable.Err = tableElpTable_Read(recElpTable)
If recElpTable.Err <> 0 Then recElpTable.Memo = Space$(20)

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
frmElpPrt.prtCentré (prtMaxX - prtMinX) / 2, "S.A. au capital de 90 000 000 Euros - R.C.Paris B 302590070 - L.B.E. 116 Tél: 01 53 76 62 62 - Téléfax: 01 42 89 09 59 - Télex: 644 030 BIAPA - Swift: BIARFRPP "

End Sub


Public Sub mainSoc_Environment()
Elp.SrvObj = "ELPDTAQ"
Elp.pcId = "FR"
Elp.SrvType = "AS400"
Elp.SrvId = "S44H1212"
Elp.SrvDtaqLib = "BIADTAQ"
Elp.SrvDtaqIn = "PC000001"
Elp.SrvDTaqOut = "PC000000"
pcIdUsrIdCtl = False
strSocSignon = "BiaSigno.bmp"
imgSocLogo = "BiaLogo.bmp"
imgSocSigle = "BiaSigle.bmp"
imgGuichet = "BiaGuichet.bmp"
prtFontName = "arial"
DataBaseName = "Bia.mdb"
elpSrvXcom = "CAV4"
'strSocSignon = "S:\BiaSrv\BiaSigno.bmp"
'imgSocLogo = "S:\BiaSrv\BiaLogo.bmp"
'imgSocSigle = "S:\BiaSrv\BiaSigle.bmp"
'imgGuichet = "S:\BiaSrv\BiaGuichet.bmp"
'prtFontName = "arial"
'DataBaseName = "S:\BiaSrv\Bia.mdb"

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


