Attribute VB_Name = "srvGSub"
Option Explicit

Type typeGParam
    TableId             As String * 12
    Application         As String * 5
    Service             As String * 3

    NatureCode          As String
    NatureLib           As String
    NatureSens          As String
    NatureNbjValeur     As Integer
    NatureDev1          As String * 3
    NatureDev2          As String * 3

    BiatypEchéance      As String * 3
    BiatypEngagement    As String * 3
    BiatypEngagementCorr As String * 3
    Commission          As String * 11
    Contrepartie        As String * 11
    BiatypReport        As String * 3

    OpérationCode       As String
    OpérationLib        As String

    GMemoSys            As Long
    ComptaLot           As Long
End Type


Public GSub_recNature As typeElpTable
Public GSub_recOpération As typeElpTable
Public GSub_CV1 As typeCV, GSub_CV2 As typeCV, GSub_CV3 As typeCV
Public GSub_recCompte As typeCompte, GSub_recRacine As typeRacine

Public GSub_arrCorrespondant() As typeCompte
Public GSub_arrCorrespondant_Nb As Integer
Public GSub_arrCorrespondant_NbMax As Integer
Public GSub_arrCorrespondant_Index As Integer
Public GSub_arrCorrespondant_Suite As Boolean

Public GSub_recCptMvt As typeCptMvt
Dim col1 As Integer, col2 As Integer, col3 As Integer, Col4 As Integer, Col5 As Integer, Col6  As Integer, Col7  As Integer
Dim wResize As Double

Public Function fctGOpe_Intérêts(mGOpe As typeGOpe, mCV As typeCV, xIntérêts As Currency, xNbj As Long)
Dim wNbjBase As Double, curX As Currency, V1 As Variant, V2 As Variant
'
' calcul des intérêts entre AmjDébut <> mGOpe.AmjFin

fctGOpe_Intérêts = "? Erreur"

xIntérêts = 0: xNbj = 0
If Trim(mGOpe.TauxRéférence1) <> "" Then fctGOpe_Intérêts = "? Taux indéxé non pgm": Exit Function
If mGOpe.AmjDébut > mGOpe.AmjFin Then fctGOpe_Intérêts = "? Date Début > fin": Exit Function

Select Case mGOpe.NbjBase
    Case "0": wNbjBase = 36000
    Case "5": wNbjBase = 36500
    Case Else: fctGOpe_Intérêts = "? Base": Exit Function
End Select
V1 = CDate(dateImp_jjMoisAAAA(mGOpe.AmjDébut))
V2 = CDate(dateImp_jjMoisAAAA(mGOpe.AmjFin))
xNbj = DateDiff("d", V1, V2)

curX = mGOpe.Montant1 * mGOpe.TauxMarge1 * xNbj / wNbjBase

xIntérêts = Round(curX, mCV.maxD)
fctGOpe_Intérêts = Null

End Function

Public Function param_Opération(lparam As typeGParam)
param_Opération = Null

GSub_recOpération.Method = "Seek="
GSub_recOpération.Id = lparam.TableId
GSub_recOpération.K1 = "Opération"
GSub_recOpération.K2 = lparam.OpérationCode
GSub_recOpération.Err = tableElpTable_Read(GSub_recOpération)
If GSub_recOpération.Err <> 0 Then
    Call MsgBox("GSub_Cpt: GSub_paramOpération", vbCritical, "Opération inconnue : " & lparam.OpérationCode)
    GSub_recOpération.K2 = ""
    GSub_recOpération.Name = "? " & lparam.OpérationCode
    param_Opération = GSub_recOpération.Name
End If

lparam.OpérationLib = GSub_recOpération.Name

End Function


Public Sub Correspondant_LoadProduction()
Dim I As Integer, Xcompte As typeCompte, V As Variant

recCompteInit GSub_recCompte
GSub_recCompte.Method = "SnapLA"
GSub_recCompte.Société = SocId$
GSub_recCompte.Agence = SocAgence$
GSub_recCompte.BiaTyp = "550"
GSub_recCompte.Numéro = "00000000000"
GSub_recCompte.BiaNum = "00"
GSub_recCompte.Devise = "000"

Xcompte = GSub_recCompte
Xcompte.Numéro = "99999999999"

V = selCompte_Load(GSub_recCompte, Xcompte, "Init")
If IsNull(V) Then
    GSub_arrCorrespondant_Nb = selCompte_Nb
    ReDim GSub_arrCorrespondant(GSub_arrCorrespondant_Nb + 1)
    For I = 1 To GSub_arrCorrespondant_Nb
        GSub_arrCorrespondant(I) = selCompte(I)
    Next I
    Call selCompte_Load(recCompte, Xcompte, "End")
End If
End Sub

Public Sub Correspondant_cbo(cbo As ComboBox, lparam As typeGParam, lDeviseN As String, lDeviseISO As String, lCompte As String)
Dim I As Integer

cbo.Clear

cbo.AddItem Compte_Imp(lCompte) & "   LORO       " & "     [" & lCompte & "]"
For I = 1 To GSub_arrCorrespondant_Nb
    If GSub_arrCorrespondant(I).Devise = lDeviseN Then
        cbo.AddItem Compte_Imp(GSub_arrCorrespondant(I).Numéro) & " _" & GSub_arrCorrespondant(I).Alpha & "     [" & GSub_arrCorrespondant(I).Numéro & "]"
    End If
Next I
        
If cbo.ListCount = 2 Then
    cbo.ListIndex = 1
Else
    cbo.ListIndex = 0
End If

    
recElpTable.Method = "Seek="
recElpTable.Id = lparam.TableId
recElpTable.K1 = "Nostro"
recElpTable.K2 = lDeviseISO
recElpTable.Err = tableElpTable_Read(recElpTable)
If recElpTable.Err = 0 Then
    Call cbo_Scan(Compte_Imp(mId$(recElpTable.Memo, 1, 11)), cbo)
End If

End Sub


Public Function param_Init(lparam As typeGParam, cbo As ComboBox)
Dim V
'Dim GSub_recNature As typeElpTable, recCodeOpération As typeElpTable

param_Init = Null
recCompteInit GSub_recCompte

recRacineInit GSub_recRacine

recCptMvtInit GSub_recCptMvt
GSub_recCptMvt.Agence = SocAgence$
GSub_recCptMvt.Société = SocId$

GSub_CV1 = CV_Euro: GSub_CV2 = CV_Euro: GSub_CV3 = CV_Euro

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = lparam.TableId
recElpTable.K1 = "Application"
recElpTable.K2 = "Code"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
lparam.Application = mId$(recElpTable.Memo, 1, 3)

recElpTable.K2 = "Service"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
lparam.Service = mId$(recElpTable.Memo, 1, 3)
If Not IsNumeric(lparam.Service) Then GoTo Num_Error

recElpTable.K2 = "GMemoSys"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
lparam.GMemoSys = mId$(recElpTable.Memo, 1, Len(recElpTable.Memo))
If Not IsNumeric(lparam.GMemoSys) Then GoTo Num_Error

recGMemo_Init xGMemo
xGMemo.Method = "NUMLOT"

xGMemo.IdRéférence = lparam.GMemoSys
xGMemo.MemoSéquence = 1
xGMemo.Application = lparam.Application
xGMemo.MemoNature = "$Sys"
xGMemo.Statut = "$"
xGMemo.MemoText = DSys
If IsNull(srvGMemo_Monitor(xGMemo)) Then
    lparam.ComptaLot = xGMemo.MemoLien1
Else
    param_Init = "GMemoSys"
    MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : GMémoSys absent", vbCritical, "srvGSub.Param_Init"
End If

recElpTable_Init GSub_recNature
GSub_recNature.Method = "Seek>="
GSub_recNature.Id = lparam.TableId
GSub_recNature.K1 = "Nature"
Call cbo_Load(GSub_recNature, cbo, 5)

recElpTable_Init GSub_recOpération
GSub_recOpération.Method = "Seek="
GSub_recOpération.Id = lparam.TableId
GSub_recOpération.K1 = "Opération"

If blnJPL Then param_Init = Null

Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "GFluxEspèces_Compta_gen"
Exit Function

Num_Error:
param_Init = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "GFluxEspèces_Param_Init"
End Function


Public Function param_Nature(lparam As typeGParam)
param_Nature = Null

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
    param_Nature = GSub_recNature.Name
End If

lparam.NatureLib = GSub_recNature.Name
lparam.BiatypEchéance = mId$(GSub_recNature.Memo, 1, 3)
lparam.BiatypEngagement = mId$(GSub_recNature.Memo, 5, 3)
lparam.BiatypEngagementCorr = mId$(GSub_recNature.Memo, 9, 3)
lparam.Commission = mId$(GSub_recNature.Memo, 13, 11)
lparam.Contrepartie = mId$(GSub_recNature.Memo, 25, 11)

End Function

Public Sub Correspondant_LoadRéplication()
Dim V As Variant
GSub_arrCorrespondant_Nb = 0
    ReDim GSub_arrCorrespondant(300)

mdbCptP0.tableCptP0_Open
reccptp0.Method = "MoveFirst"
Mid$(MsgTxt, 1, 34) = Space$(34)

V = dbCptP0_ReadE(reccptp0)

Do While reccptp0.Err = 0
    ' compte correspondant actif
    If mId$(reccptp0.Id, 15, 3) = "550" And mId$(reccptp0.Id, 15, 3) <> "000" And mId$(reccptp0.Text, 116, 1) = " " Then
        MsgTxtIndex = 0
        MsgTxt = Space$(recCompteLen)
        Mid$(MsgTxt, 35, memoCompteLen) = mId$(reccptp0.Text, 1, memoCompteLen)
        If IsNull(srvCompteGetBuffer(recCompte)) Then
            GSub_arrCorrespondant_Nb = GSub_arrCorrespondant_Nb + 1
            GSub_arrCorrespondant(GSub_arrCorrespondant_Nb) = recCompte
    
        End If
    End If
    
    reccptp0.Method = "MoveNext    "
    reccptp0.Err = tableCptP0_Read(reccptp0)
Loop
mdbCptP0.tableCptP0_Close

ReDim Preserve GSub_arrCorrespondant(GSub_arrCorrespondant_Nb + 1)


End Sub


Public Sub GMemo_Display(picX As PictureBox, lGMemo_Nb As Integer, lGMemo() As typegMemo)
Dim X As String, I As Integer
Dim mCurrentX1 As Integer, mForeColor1 As Long
Dim mCurrentX2 As Integer, mForeColor2 As Long
Dim curTotal As Currency

DoEvents: picX.Cls
Call pic_Resize(picX, lGMemo_Nb)
picX.ForeColor = libUsr.ForeColor
''picX.Line (0, 600)-(9300, 600)
'picX.Line (0, 1200)-(9300, 1200)
picX.CurrentY = 20
curTotal = 0

For I = 1 To lGMemo_Nb
    picX.FontBold = False
    If lGMemo(I).MemoLien1 = 0 Then
        picX.ForeColor = warnUsrColor
        mForeColor1 = libUsr.ForeColor
    Else
        picX.ForeColor = libUsr.ForeColor
        mForeColor1 = warnUsrColor
   End If
    
    If Trim(lGMemo(I).MemoNature) <> constCompta Then
        picX.CurrentX = col1: picX.Print lGMemo(I).MemoText
    Else
    
        Call srvCptMvt_GetX(GSub_recCptMvt, lGMemo(I).MemoText)
        If GSub_recCptMvt.Mt <> 0 Then
            curTotal = curTotal + GSub_recCptMvt.Mt

            picX.CurrentX = col1: picX.Print GSub_recCptMvt.Devise & "." & Compte_Imp(GSub_recCptMvt.Compte);
            
                picX.CurrentX = col2
                GSub_recCompte.Devise = GSub_recCptMvt.Devise
                GSub_recCompte.Numéro = GSub_recCptMvt.Compte
                If IsNull(mdbCptP0_Find(GSub_recCompte)) Then
                        picX.Print Trim(GSub_recCompte.Intitulé);
                Else
                    picX.ForeColor = errUsr.ForeColor
                    picX.Print "???????";
                End If
           
            If picX.CurrentX < col3 Then picX.CurrentX = col3
            picX.ForeColor = warnUsrColor
            picX.Print dateImp(GSub_recCptMvt.AmjValeur);
            
            picX.FontBold = True
            If GSub_recCptMvt.Mt < 0 Then
                picX.CurrentX = Col4: picX.ForeColor = errUsr.ForeColor
            Else
                picX.CurrentX = Col5: picX.ForeColor = libUsr.ForeColor
            End If

            X = Format$(Abs(GSub_recCptMvt.Mt), "### ### ### ### ##0.00")
            picX.CurrentX = picX.CurrentX - picX.TextWidth(X)
            picX.Print X;
        End If
    End If
    picX.CurrentY = picX.CurrentY + 350
Next I
If curTotal <> 0 Then Call MsgBox("pièce non équilibrée", vbCritical, "srvGSub.GMemo_Display")


End Sub
Public Sub GEch_Display(picX As PictureBox, lGEch_Nb As Integer, lGech() As typeGEch)
Dim X As String, I As Integer, wColor As Long

DoEvents: picX.Cls
Call pic_Resize(picX, lGEch_Nb)
picX.CurrentY = 20
col2 = 500 * wResize: col3 = 2000 * wResize: Col4 = 4000 * wResize: Col5 = 5000 * wResize: Col6 = 6000 * wResize: Col7 = 8000 * wResize
For I = 1 To lGEch_Nb
    If Trim(lGech(I).Statut) <> "" Then
        picX.ForeColor = libUsr.ForeColor
    Else
        picX.ForeColor = warnUsrColor
    End If
    picX.CurrentX = col1: picX.Print lGech(I).EchSéquence;
    picX.CurrentX = col2: picX.Print lGech(I).EchFct;
    picX.CurrentX = col3: picX.Print dateImp(lGech(I).EchAMJ) & " - " & timeImp(lGech(I).EchHMS);
    picX.CurrentX = Col4: picX.Print lGech(I).EchUsr;
    picX.CurrentX = Col5: picX.Print lGech(I).ActionFct;
    picX.CurrentX = Col6: picX.Print dateImp(lGech(I).ActionAmj) & " - " & timeImp(lGech(I).ActionHms);
    picX.CurrentX = Col7: picX.Print lGech(I).ActionUsr;
    picX.CurrentY = picX.CurrentY + 350
Next I



End Sub


Public Sub GEch_New(lparam As typeGParam, lGope As typeGOpe, lGEch_Nb As Integer, lGech() As typeGEch)

ReDim lGech(10)

srvGEch.recGEch_Init lGech(1)
lGech(1).Method = constAddNew

With lGech(1)                                   ' Saisie
    .IdRéférence = lGope.IdRéférence
    .EchSéquence = 1
    .Application = lparam.Application
    .EchFct = constSaisie
    .EchAMJ = DSys
    .EchHMS = time_Hms
    .EchUsr = usrId
    .Statut = "à"
End With
lGEch_Nb = 1

End Sub
Public Sub Gech_Update(lGech As typeGEch)

If Trim(lGech.Method) = "" Then lGech.Method = constUpdate

With lGech                                ' Saisie
    .EchAMJ = DSys
    .EchHMS = time_Hms
    .EchUsr = usrId
End With

End Sub

Public Sub GMemo_New(lparam As typeGParam, lGope As typeGOpe, lGMemo_Nb As Integer, lGMemo() As typegMemo)

ReDim lGMemo(10)

srvGMemo.recGMemo_Init lGMemo(1)

With lGMemo(1)                      ' Saisie
    .Method = constAddNew
    .IdRéférence = 0
    .MemoSéquence = 1
    .MemoSéquencePlus = 0
    .Application = lparam.Application
    .MemoNature = "Info"
End With
lGMemo_Nb = 1
End Sub

Public Sub pic_Resize(picX As PictureBox, lNb As Integer)
Dim I As Integer
I = lNb * 350
If I < 700 Then
    picX.Height = 700
Else
    If I < picX.Container.Height Then
        picX.Height = I
    Else
        picX.Height = picX.Container.Height - 50
    End If
End If
picX.Top = picX.Container.Height - picX.Height - 50


If picX.Width < 10000 Then
    picX.FontSize = 8
    wResize = 1
Else
    picX.FontSize = 12
    wResize = picX.Width / 9350
End If
col1 = 50: col2 = 1600 * wResize: col3 = 5500 * wResize: Col4 = 8000 * wResize: Col5 = 9000 * wResize

End Sub


Public Function GOpération_Load(lparam As typeGParam, lGope As typeGOpe, lGEch_Nb As Integer, lGech() As typeGEch, lGFlux_Nb As Integer, lGFlux() As typeGFlux, lGMemo_Nb As Integer, lGMemo() As typegMemo)
GOpération_Load = "?"
lGope.obj = "SRVGSUB     "
lGope.Method = "Seek*"

MsgTxtLen = 0
lGEch_Nb = 0: lGFlux_Nb = 0: lGMemo_Nb = 0

Call srvGOpe_PutBuffer(lGope)
If IsNull(SndRcv()) Then
    GOpération_Load = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        Select Case Trim(mId$(MsgTxt, MsgTxtIndex + 1, 12))
            Case "SRVGOPE":
                            If Not IsNull(srvGOpe_GetBuffer(lGope)) Then Exit Do
                            
            Case "SRVGECH":
                            If lGEch_Nb = UBound(lGech) Then ReDim Preserve lGech(lGEch_Nb + 10)
                            lGEch_Nb = lGEch_Nb + 1
                            If Not IsNull(srvGEch_GetBuffer(lGech(lGEch_Nb))) Then Exit Do
                            
            Case "SRVGFLUX":
                            If lGFlux_Nb = UBound(lGFlux) Then ReDim Preserve lGFlux(lGFlux_Nb + 10)
                            lGFlux_Nb = lGFlux_Nb + 1
                            If Not IsNull(srvGFlux_GetBuffer(lGFlux(lGFlux_Nb))) Then Exit Do
                            
            Case "SRVGMEMO":
                            If lGMemo_Nb = UBound(lGMemo) Then ReDim Preserve lGMemo(lGMemo_Nb + 10)
                            lGMemo_Nb = lGMemo_Nb + 1
                            If Not IsNull(srvGMemo_GetBuffer(lGMemo(lGMemo_Nb))) Then Exit Do
      End Select
    Loop
End If


End Function
