Attribute VB_Name = "srvSwift_MT950"
Option Explicit
'------------------------------------------------------------------------------
'YBIAMVT0.MOUVEMMON .SD db >0 et cr < 0
'YBIARELV.BIARELSD0 SD1  db < 0 et cr < 0
'------------------------------------------------------------------------------

Dim blnError_Exit As Boolean

Dim paramMT950_Loro As String, paramMT950_Nostro As String
Dim paramMT950F_Loro As String, paramMT950F_Nostro As String, paramMT950_YBIARELV As String
Dim mYBIAMVT0 As typeYBIAMVT0, xYBIAMVT0 As typeYBIAMVT0
Dim xZRELEVE0 As typeZRELEVE0
Dim newYBIARELV As typeYBIARELV, oldYBIARELV As typeYBIARELV
Dim newYBIARELV_Method As String

Dim blnRelevéA4W_Loro As Boolean, blnRelevéA4W_Nostro As Boolean, blnRelevéA4W_Update  As Boolean
Dim blnMT950_Loro As Boolean, blnMT950_Open As Boolean

Dim mLine28c_Sequence As Integer, mLine28c_Number As Integer

Dim xMT950 As String

Dim xLine20 As String, xLine25 As String, xLine28C As String, xLine60 As String, xLine61 As String

Dim mLine60_Amj As String, mSolde As Currency
Dim xSwift As String, mFile1_Seq As Long, mFile2_Seq As Long

Dim mRacine_BIC As String * 11, mRacine_SAB As String * 7

Dim IbmAmjMin As String, IbmAmjMax As String

Dim wAMJHMS As String

Dim blnTransaction As Boolean

Dim blnMT950_64 As Boolean, xMT950_64 As String
Dim curMt950_64 As Currency, curMTD_MOUVEMDVA As Currency
Public Function GetDevise(z As String) As String
Dim rsDev As ADODB.Recordset
Dim xSQL As String

    GetDevise = ""
    xSQL = "select COMPTEDEV from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM ='" & z & "'"
    Set rsDev = cnsab.Execute(xSQL)

    Do Until rsDev.EOF
        GetDevise = rsDev(0)
        Exit Do
    Loop
    rsDev.Close
    
End Function

Public Sub Swift_MT950_Block_SANS_MOUVEMENTS(l60FM As String)
Dim lDev As String
blnMT950_Open = True

mLine28c_Number = mLine28c_Number + 1
xLine28C = ":28C:" & mLine28c_Sequence & "/" & mLine28c_Number & Asc13 & Asc10
If mSolde <= 0 Then
    xLine60 = ":60" & l60FM & ":C"
Else
    xLine60 = ":60" & l60FM & ":D"
End If

'on va chercher la devise du compte car inconnue si pas de mouvements
lDev = GetDevise(xZRELEVE0.RELEVECOM)
xLine60 = xLine60 & mLine60_Amj & lDev & cur_AbsV_Dev(mSolde, lDev) & Asc13 & Asc10

xSwift = Asc01 & "{1:F01" & paramBic8 & "AXXX0000000000}{2:I950" & mRacine_BIC & "XN}{4:" & Asc13 & Asc10 & xLine20 & xLine25 & xLine28C & xLine60

End Sub

Public Sub Swift_MT950_Close_SANS_MOUVEMENTS(l61FM As String)
Dim X As String, X6 As String
Dim V
Dim lDev As String

lDev = GetDevise(xZRELEVE0.RELEVECOM)

If mSolde <= 0 Then
    X = ":62" & l61FM & ":C"
Else
    X = ":62" & l61FM & ":D"
End If

newYBIARELV.BIARELSD1 = -mSolde
If l61FM = "F" Then
    newYBIARELV.BIARELD1 = IbmAmjMax + 19000000
    X6 = Mid$(newYBIARELV.BIARELD1, 3, 6)
    
    If blnError_Exit Then
       newYBIARELV_Method = "$Exit"
    Else
        If Not blnRelevéA4W_Update Then
            newYBIARELV_Method = "sans màj"
        Else
'maj BIARELID = 0 => BIARELNUM dernier extrait
'-------------------------------------------------
            If newYBIARELV_Method = constAddNew Then
                V = sqlYBIARELV_Insert(newYBIARELV)
            Else
                V = sqlYBIARELV_Update(newYBIARELV, oldYBIARELV)
            End If
            
            If Not IsNull(V) Then
                blnError_Exit = True
                Print #3, "$ MT950 :ERREUR màj YBIARELV :  " & newYBIARELV.BIARELCOM & "/" & newYBIARELV.BIARELID & " " & V
                Shell_MsgBox "# MT950 :ERREUR Reprise Solde : compte : " & newYBIARELV.BIARELCOM, vbCritical, "srvSwift_MT950", False
            Else
'maj historique BIARELNUM dernier extrait
'-------------------------------------------------
                newYBIARELV.BIARELID = newYBIARELV.BIARELNUM
                V = sqlYBIARELV_Insert(newYBIARELV)
                If Not IsNull(V) Then
                   blnError_Exit = True
                   Print #3, "$ MT950 :ERREUR màj SAB073 /YBIARELV :  " & newYBIARELV.BIARELCOM & " " & V
                   Shell_MsgBox "# MT950 :ERREUR Reprise Solde : compte : " & newYBIARELV.BIARELCOM, vbCritical, "srvSwift_MT950", False
                End If
           End If
    
        End If
    End If
    
    Print #3, newYBIARELV_Method & " " & newYBIARELV.BIARELCOM & "/" & newYBIARELV.BIARELID _
         , newYBIARELV.BIARELNUM, newYBIARELV.BIARELSD0, newYBIARELV.BIARELD0 _
         , newYBIARELV.BIARELSD1, newYBIARELV.BIARELD1
        
Else
    If blnMT950_Loro Then
        'newYBIARELV.BIARELD1 = mYBIAMVT0.MOUVEMDTR + 19000000
        'X6 = mLine60_Amj
        X6 = Mid$(CStr(IbmAmjMax + 19000000), 3, 6)
    Else
        
        X6 = Mid$(CStr(IbmAmjMax + 19000000), 3, 6)
   End If
End If

    
    
X = X & X6 & lDev & cur_AbsV_Dev(mSolde, lDev) & Asc13 & Asc10
'___________________________________________________________________________________________
'$JPL 20120615 Champ 64

'If blnMT950_64 Then
'    curMt950_64 = mSolde - curMTD_MOUVEMDVA
'    If curMt950_64 <= 0 Then
'        xMT950_64 = ":64:C"
'    Else
'        xMT950_64 = ":64:D"
'    End If
'    X = X & xMT950_64 & X6 & mYBIAMVT0.COMPTEDEV & cur_AbsV_Dev(curMt950_64, mYBIAMVT0.COMPTEDEV) & Asc13 & Asc10
'End If

'___________________________________________________________________________________________

xSwift = xSwift & X & "-}" & Asc03
Swift_Mt950_Write


blnMT950_Open = False

End Sub

Public Function Swift_MT950_Init_SANS_MOUVEMENTS() As Boolean
Dim xId As String
Dim V, xSQL As String
Dim lDev As String

Swift_MT950_Init_SANS_MOUVEMENTS = True
If Mid$(mYBIAMVT0.MOUVEMCOM, 1, 1) = "N" Then
    blnMT950_Loro = False
Else
    blnMT950_Loro = True
End If

lDev = GetDevise(xZRELEVE0.RELEVECOM)

'Rechercher le BIC du destinaire, 11 caractères et doit se terminer par 'XXX'
'----------------------------------------------------------------------------
V = rsZADRESS0_BIC_Compte(xZRELEVE0.RELEVECOM, mRacine_BIC)
If Mid$(mRacine_BIC, 9, 3) <> "XXX" Then Mid$(mRacine_BIC, 9, 3) = "XXX"

If Len(Trim(mRacine_BIC)) <> 11 Then
    Swift_MT950_Init_SANS_MOUVEMENTS = False
    Shell_MsgBox "# MT950 : BIC erroné, compte : " & xZRELEVE0.RELEVECOM, vbInformation, "srvSwift_MT950", False
Else

'Historique des envois MT950
'----------------------------------------------------------------------------
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIARELV" _
     & " where BIARELCOM = '" & xZRELEVE0.RELEVECOM & "'" _
     & " and   BIARELREL = 'W'" _
     & " and   BIARELID  = 0 "
     
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        newYBIARELV_Method = "MT950_OK"
        V = rsYBIARELV_GetBuffer(rsSab, oldYBIARELV)
        If Not IsNull(V) Then
            Swift_MT950_Init_SANS_MOUVEMENTS = False
            blnError_Exit = True
        End If
        
    Else
        rsYBIARELV_Init oldYBIARELV
        oldYBIARELV.BIARELCOM = xZRELEVE0.RELEVECOM
        oldYBIARELV.BIARELREL = "W"
        oldYBIARELV.BIARELD1 = YBIATAB0_DATE_CPT_JP1
        oldYBIARELV.BIARELSD1 = oldYBIARELV.BIARELSD1 * (-1)
        newYBIARELV_Method = constAddNew
    End If
             
        xLine20 = ":20:COBKMT950 " & Mid$(DSys, 3, 6) & Asc13 & Asc10
        
        If Trim(oldYBIARELV.BIAOLDCOM) <> "" Then
            xLine25 = ":25:" & Format$(oldYBIARELV.BIAOLDCOM, "@@@@@.@@@.@@.@") & "." & lDev & Asc13 & Asc10
        Else
            xLine25 = ":25:" & xZRELEVE0.RELEVECOM & Asc13 & Asc10
        End If
       
        newYBIARELV = oldYBIARELV
        newYBIARELV.BIARELNUM = oldYBIARELV.BIARELNUM + 1
        newYBIARELV.BIARELD0 = oldYBIARELV.BIARELD1
        newYBIARELV.BIARELSD0 = -oldYBIARELV.BIARELSD1
        
        mLine28c_Sequence = newYBIARELV.BIARELNUM
        mLine28c_Number = 0
        mLine60_Amj = Mid$(newYBIARELV.BIARELD0, 3, 6)
        mSolde = newYBIARELV.BIARELSD0
        Swift_MT950_Block_SANS_MOUVEMENTS "F"
        
        If oldYBIARELV.BIARELSD1 <> -mSolde Then
            Swift_MT950_Init_SANS_MOUVEMENTS = False
            blnError_Exit = True
            Print #3, "$ MT950 :ERREUR Reprise Solde : compte : " & xZRELEVE0.RELEVECOM;
            Print #3, " solde comptable veille (YBIAMVT0) : " & -mSolde;
            Print #3, " solde MT950 (YBIARELV) : " & oldYBIARELV.BIARELSD1
            Shell_MsgBox "# MT950 :ERREUR Reprise Solde : compte : " & xZRELEVE0.RELEVECOM, vbCritical, "srvSwift_MT950", False
        End If
'   End If
    
End If
End Function

Public Sub Swift_MT950_Monitor(lAMJMin As String, lAMJMax As String, lRelevéA4W_Loro As Boolean, lRelevéA4W_Nostro As Boolean, lRelevéA4W_Update As Boolean, lRelevéA4W_Confirmation As Boolean, cnAdo As ADODB.Connection)
Dim xFileName As String, X As String
Dim rsMT950 As ADODB.Recordset
Dim rsYBIAMVT0 As ADODB.Recordset
Dim xSQL As String, V
Dim blnOk As Boolean
Dim meYBIAMON0 As typeYBIAMON0
On Error GoTo Error_Handler

blnError_Exit = False
blnRelevéA4W_Loro = lRelevéA4W_Loro
blnRelevéA4W_Nostro = lRelevéA4W_Nostro
blnRelevéA4W_Update = lRelevéA4W_Update

wAMJHMS = DSys & "_" & time_Hms & "_"

'====================================================================================================
If blnRelevéA4W_Update Then
    meYBIAMON0.MONAPP = "COMPTA"
    meYBIAMON0.MONFLUX = "MT950"
    meYBIAMON0.MONSTATUS = ""
    
    V = fctExploitation_Auto_Control(meYBIAMON0)
    If Not IsNull(V) Then Exit Sub
    'If Not blnAuto_Exploitation_Ok("DATE_CPT_J", "@MT950") Then
    '    Call MsgBox("MT950 : Traitement déjà effectué au " & YBIATAB0_DATE_CPT_J, vbCritical, paramYBase_DataF & "\@MT950_Exploitation_ok.txt")
    '    Exit Sub
    'End If
    '-----------------------------------------------------------------------------------------------
    App_Debug = " -Swift_MT950_Monitor  "
    '-----------------------------------------------------------------------------------------------
    '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    'V = V = cnSAB_Transaction("BeginTrans")
    'If Not IsNull(V) Then GoTo Error_Handler
    blnTransaction = True

End If

paramSAA_Init
Set rsMT950 = Nothing
Set rsYBIAMVT0 = Nothing

IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)

paramMT950_Loro = "MT950_Loro_" & lAMJMax

paramMT950F_Loro = paramYBase_DataF & paramMT950_Loro & paramYBase_Data_ExtensionP
If Dir(paramMT950F_Loro) <> "" Then Kill paramMT950F_Loro
Call FEU_ROUGE
Open paramMT950F_Loro For Binary Access Write As #1

paramMT950_Nostro = "MT950_Nostro_" & lAMJMax
paramMT950F_Nostro = paramYBase_DataF & paramMT950_Nostro & paramYBase_Data_ExtensionP
If Dir(paramMT950F_Nostro) <> "" Then Kill paramMT950F_Nostro
Open paramMT950F_Nostro For Binary Access Write As #2

mFile1_Seq = 1: mFile2_Seq = 1

paramMT950_YBIARELV = paramYBase_DataF & "log\" & wAMJHMS & "MT950_YBIARELV_" & lAMJMax & paramYBase_Data_ExtensionP
Open paramMT950_YBIARELV For Output As #3
Print #3, Time & " : " & paramMT950_YBIARELV
Print #3, Time & " :________________________________________________________________ "

'********************************* AJOUT KOKOU 18/11/2024 ********************************

'Comptes pour lesquels on doit envoyer les MT950 même lorsqu'il n' y a pas de mouvement

Dim rsSansMVT As ADODB.Recordset
Dim xSQLSMVT As String
Dim arrCompteSansMVT()
Dim compteSansMVT_Nb As Integer
Dim ss As Integer
Dim blnSansMVT As Boolean

xSQLSMVT = "SELECT MTCOMPTE FROM SAB073SPE.MT950SSMV0"
Set rsSansMVT = cnsab.Execute(xSQLSMVT)
compteSansMVT_Nb = 0
Do Until rsSansMVT.EOF
  compteSansMVT_Nb = compteSansMVT_Nb + 1
  ReDim Preserve arrCompteSansMVT(compteSansMVT_Nb)
  arrCompteSansMVT(compteSansMVT_Nb) = rsSansMVT("MTCOMPTE")
  rsSansMVT.MoveNext
Loop

'************************************** FIN AJOUT KOKOU ************************************

xSwift = ""

'============================================================
xSQL = "select * from " & paramIBM_Library_SAB & ".ZRELEVE0  " _
     & " where RELEVEREL = 'W'" _
     & " and   RELEVEETA = " & currentZMNURUT0.MNURUTETB _
     
Set rsMT950 = cnsab.Execute(xSQL)

Do Until rsMT950.EOF
    V = rsZRELEVE0_GetBuffer(rsMT950, xZRELEVE0)
    blnOk = False
    
    '******************************** AJOUT KOKOU 18/11/2024 ******************************
    
    'Comptes pour lesquels on doit envoyer les MT950 même lorsqu'il n' y a pas de mouvement
    
    blnSansMVT = False
    For ss = 1 To compteSansMVT_Nb
        If Trim(xZRELEVE0.RELEVECOM) = Trim(arrCompteSansMVT(ss)) Then
            blnSansMVT = True
            Exit For
        End If
    Next ss
    
    '************************************** FIN AJOUT KOKOU ********************************
    
'___________________________________________________________________________________________
'$JPL 20120615 Champ 64

blnMT950_64 = False: curMTD_MOUVEMDVA = 0: xMT950_64 = ""
If Mid$(xZRELEVE0.RELEVECOM, 1, 5) = "11001" Then
    blnMT950_64 = True
    Call SWIFT_MT950_MOUVEMDVA(xZRELEVE0.RELEVECOM, IbmAmjMax)
End If
    
'___________________________________________________________________________________________
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
     & " where MOUVEMCOM = '" & xZRELEVE0.RELEVECOM & "'" _
     & " and MOUVEMDTR >= " & IbmAmjMin _
     & " and MOUVEMDTR <= " & IbmAmjMax _
     & " order by MOUVEMDTR, MOUVEMPIE, MOUVEMECR"
     
    Set rsYBIAMVT0 = cnsab.Execute(xSQL)
    If Not rsYBIAMVT0.EOF Then

        V = rsYBIAMVT0_GetBuffer(rsYBIAMVT0, xYBIAMVT0)
        If Not IsNull(V) Then
            MsgBox "Swift_MT950_Monitor " & V
            blnError_Exit = True
            Exit Do
        End If
        mYBIAMVT0 = xYBIAMVT0
        If lRelevéA4W_Confirmation Then
            X = MsgBox("Voulez-vous générer un swift MT950 ?", vbYesNo + vbQuestion + vbDefaultButton2, "COMPTE : " & mYBIAMVT0.MOUVEMCOM)
            If X = vbYes Then blnOk = Swift_MT950_Init
        Else
            blnOk = Swift_MT950_Init
        End If
        If blnOk Then
            Do Until rsYBIAMVT0.EOF

                Swift_MT950_Line
                rsYBIAMVT0.MoveNext
                Call rsYBIAMVT0_GetBuffer(rsYBIAMVT0, xYBIAMVT0)
           Loop
                
        End If
    Else
        If blnSansMVT Then ' AJOUT KOKOU 18/11/2024
        'If InStr(xZRELEVE0.RELEVECOM, "50533") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51374978001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51374400001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51352978001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51352400001") > 0 Then    ' Commentaire KOKOU Sur demande des Métiers j'ai ajouté le 30/09/2024 les comptes 51352400001,51374978001,51374400001,51352978001
            blnOk = Swift_MT950_Init_SANS_MOUVEMENTS
        End If
    End If
    If blnSansMVT Then ' AJOUT KOKOU 18/11/2024
    'If InStr(xZRELEVE0.RELEVECOM, "50533") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51374978001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51374400001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51352978001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51352400001") > 0 Then    ' Commentaire KOKOU Sur demande des Métiers j'ai ajouté le 30/09/2024 les comptes 51352400001,51374978001,51374400001,51352978001
        If blnMT950_Open Then Swift_MT950_Close_SANS_MOUVEMENTS "F"
    Else
        If blnMT950_Open Then Swift_MT950_Close "F"
    End If

'___________________________________________________________________________________________

    rsMT950.MoveNext
Loop
'======================================================
'Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "affichage : " & Nb)

If blnSansMVT Then ' AJOUT KOKOU 18/11/2024
'If InStr(xZRELEVE0.RELEVECOM, "50533") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51374978001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51374400001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51352978001") > 0 Or InStr(xZRELEVE0.RELEVECOM, "51352400001") > 0 Then    ' Commentaire KOKOU Sur demande des Métiers j'ai ajouté le 30/09/2024 les comptes 51352400001,51374978001,51374400001,51352978001
    If blnMT950_Open Then Swift_MT950_Close_SANS_MOUVEMENTS "F"
Else
    If blnMT950_Open Then Swift_MT950_Close "F"
End If
Close #1
Close #2


Print #3, Time & " :________________________________________________________________ "
If blnError_Exit Then

    Print #3, " $$$$$$$$$$$$$$ ERREUR GRAVE : FICHIERS LORO ET NOSTRO NON TRANSFERES $$$$$$$$$$$$"
    If blnRelevéA4W_Update Then V = cnSAB_Transaction("Rollback")
Else
    Swift_MT950_Corona_from_SAB
    Swift_MT950_SAA_from_SAB
    '====================================================================================================
    'If blnRelevéA4W_Update Then Call blnAuto_Exploitation_Ok("Update", "@MT950")
    If blnRelevéA4W_Update Then V = fctExploitation_Auto_End(meYBIAMON0)

End If

Print #3, Time & " :================================================================="
Close
Call FEU_VERT

frmElpPrt.Shell_Print paramMT950_YBIARELV

Exit Sub

Error_Handler:

Close
Shell_MsgBox "Swift_MT950_Monitor " & Error, vbCritical, "srvSwift_MT950", True

End Sub
Public Function Swift_MT950_Init() As Boolean
Dim xId As String
Dim V, xSQL As String
Swift_MT950_Init = True
If Mid$(mYBIAMVT0.MOUVEMCOM, 1, 1) = "N" Then
    blnMT950_Loro = False
Else
    blnMT950_Loro = True
End If

'Rechercher le BIC du destinaire, 11 caractères et doit se terminer par 'XXX'
'----------------------------------------------------------------------------
V = rsZADRESS0_BIC_Compte(mYBIAMVT0.MOUVEMCOM, mRacine_BIC)
If Mid$(mRacine_BIC, 9, 3) <> "XXX" Then Mid$(mRacine_BIC, 9, 3) = "XXX"

If Len(Trim(mRacine_BIC)) <> 11 Then
    Swift_MT950_Init = False
    Shell_MsgBox "# MT950 : BIC erroné, compte : " & mYBIAMVT0.MOUVEMCOM, vbInformation, "srvSwift_MT950", False
Else

'Historique des envois MT950
'----------------------------------------------------------------------------
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIARELV" _
     & " where BIARELCOM = '" & mYBIAMVT0.MOUVEMCOM & "'" _
     & " and   BIARELREL = 'W'" _
     & " and   BIARELID  = 0 "
     
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        newYBIARELV_Method = "MT950_OK"
        V = rsYBIARELV_GetBuffer(rsSab, oldYBIARELV)
        If Not IsNull(V) Then
            Swift_MT950_Init = False
            blnError_Exit = True
        End If
        
    Else
        rsYBIARELV_Init oldYBIARELV
        oldYBIARELV.BIARELCOM = mYBIAMVT0.MOUVEMCOM
        oldYBIARELV.BIARELREL = "W"
        oldYBIARELV.BIARELD1 = YBIATAB0_DATE_CPT_JP1
        oldYBIARELV.BIARELSD1 = -mYBIAMVT0.BIAMVTSD0
        newYBIARELV_Method = constAddNew
    End If
             
        xLine20 = ":20:COBKMT950 " & Mid$(DSys, 3, 6) & Asc13 & Asc10
        
        If Trim(oldYBIARELV.BIAOLDCOM) <> "" Then
            xLine25 = ":25:" & Format$(oldYBIARELV.BIAOLDCOM, "@@@@@.@@@.@@.@") & "." & oldYBIARELV.BIAOLDDEV & Asc13 & Asc10
        Else
            xLine25 = ":25:" & mYBIAMVT0.MOUVEMCOM & Asc13 & Asc10
        End If
       
        newYBIARELV = oldYBIARELV
        newYBIARELV.BIARELNUM = oldYBIARELV.BIARELNUM + 1
        newYBIARELV.BIARELD0 = oldYBIARELV.BIARELD1
        newYBIARELV.BIARELSD0 = mYBIAMVT0.BIAMVTSD0
        'newYBIARELV.BIAMVTID0 = mYBIAMVT0.BIAMVTID
        
        mLine28c_Sequence = newYBIARELV.BIARELNUM
        mLine28c_Number = 0
        mLine60_Amj = Mid$(newYBIARELV.BIARELD0, 3, 6)
        mSolde = mYBIAMVT0.BIAMVTSD0
        Swift_MT950_Block "F"
        
        If oldYBIARELV.BIARELSD1 <> -mSolde Then
            Swift_MT950_Init = False
            blnError_Exit = True
            Print #3, "$ MT950 :ERREUR Reprise Solde : compte : " & mYBIAMVT0.MOUVEMCOM;
            Print #3, " solde comptable veille (YBIAMVT0) : " & -mSolde;
            Print #3, " solde MT950 (YBIARELV) : " & oldYBIARELV.BIARELSD1
            Shell_MsgBox "# MT950 :ERREUR Reprise Solde : compte : " & mYBIAMVT0.MOUVEMCOM, vbCritical, "srvSwift_MT950", False
        End If
'   End If
    
End If

End Function
Public Sub Swift_MT950_Block(l60FM As String)
blnMT950_Open = True

mLine28c_Number = mLine28c_Number + 1
xLine28C = ":28C:" & mLine28c_Sequence & "/" & mLine28c_Number & Asc13 & Asc10
If mSolde <= 0 Then
    xLine60 = ":60" & l60FM & ":C"
Else
    xLine60 = ":60" & l60FM & ":D"
End If

xLine60 = xLine60 & mLine60_Amj & mYBIAMVT0.COMPTEDEV & cur_AbsV_Dev(mSolde, mYBIAMVT0.COMPTEDEV) & Asc13 & Asc10

xSwift = Asc01 & "{1:F01" & paramBic8 & "AXXX0000000000}{2:I950" & mRacine_BIC & "XN}{4:" & Asc13 & Asc10 & xLine20 & xLine25 & xLine28C & xLine60
End Sub

Public Sub Swift_MT950_Close(l61FM As String)
Dim X As String, X6 As String
Dim V

If mSolde <= 0 Then
    X = ":62" & l61FM & ":C"
Else
    X = ":62" & l61FM & ":D"
End If

newYBIARELV.BIARELSD1 = -mSolde
If l61FM = "F" Then
    newYBIARELV.BIARELD1 = IbmAmjMax + 19000000
    X6 = Mid$(newYBIARELV.BIARELD1, 3, 6)
    
    If blnError_Exit Then
       newYBIARELV_Method = "$Exit"
    Else
        If Not blnRelevéA4W_Update Then
            newYBIARELV_Method = "sans màj"
        Else
'maj BIARELID = 0 => BIARELNUM dernier extrait
'-------------------------------------------------
            If newYBIARELV_Method = constAddNew Then
                V = sqlYBIARELV_Insert(newYBIARELV)
            Else
                V = sqlYBIARELV_Update(newYBIARELV, oldYBIARELV)
            End If
            
            If Not IsNull(V) Then
                blnError_Exit = True
                Print #3, "$ MT950 :ERREUR màj YBIARELV :  " & newYBIARELV.BIARELCOM & "/" & newYBIARELV.BIARELID & " " & V
                Shell_MsgBox "# MT950 :ERREUR Reprise Solde : compte : " & newYBIARELV.BIARELCOM, vbCritical, "srvSwift_MT950", False
            Else
'maj historique BIARELNUM dernier extrait
'-------------------------------------------------
                newYBIARELV.BIARELID = newYBIARELV.BIARELNUM
                V = sqlYBIARELV_Insert(newYBIARELV)
                If Not IsNull(V) Then
                   blnError_Exit = True
                   Print #3, "$ MT950 :ERREUR màj SAB073 /YBIARELV :  " & newYBIARELV.BIARELCOM & " " & V
                   Shell_MsgBox "# MT950 :ERREUR Reprise Solde : compte : " & newYBIARELV.BIARELCOM, vbCritical, "srvSwift_MT950", False
                End If
           End If
    
        End If
    End If
    
    Print #3, newYBIARELV_Method & " " & newYBIARELV.BIARELCOM & "/" & newYBIARELV.BIARELID _
         , newYBIARELV.BIARELNUM, newYBIARELV.BIARELSD0, newYBIARELV.BIARELD0 _
         , newYBIARELV.BIARELSD1, newYBIARELV.BIARELD1
        
Else
    If blnMT950_Loro Then
        newYBIARELV.BIARELD1 = mYBIAMVT0.MOUVEMDTR + 19000000
        X6 = mLine60_Amj
    Else
        
        X6 = Mid$(CStr(IbmAmjMax + 19000000), 3, 6)
   End If
End If

    
    
X = X & X6 & mYBIAMVT0.COMPTEDEV & cur_AbsV_Dev(mSolde, mYBIAMVT0.COMPTEDEV) & Asc13 & Asc10
'___________________________________________________________________________________________
'$JPL 20120615 Champ 64

If blnMT950_64 Then
    curMt950_64 = mSolde - curMTD_MOUVEMDVA
    If curMt950_64 <= 0 Then
        xMT950_64 = ":64:C"
    Else
        xMT950_64 = ":64:D"
    End If
    X = X & xMT950_64 & X6 & mYBIAMVT0.COMPTEDEV & cur_AbsV_Dev(curMt950_64, mYBIAMVT0.COMPTEDEV) & Asc13 & Asc10
End If

'___________________________________________________________________________________________

xSwift = xSwift & X & "-}" & Asc03
Swift_Mt950_Write


blnMT950_Open = False
End Sub

Public Sub Swift_MT950_Line()
Dim X As String, xSens As String, X1 As String
Dim I As Integer, lenX As Integer
Dim xRef As String

If Len(xSwift) > 1700 Then
    Swift_MT950_Close "M"
    Swift_MT950_Block "M"
End If

If xYBIAMVT0.MOUVEMMON <= 0 Then
    xSens = "C"
Else
    xSens = "D"
End If

For I = 1 To 3
    Select Case Mid$(xYBIAMVT0.MOUVEMOPE, I, 1)
        Case "A" To "Z":
        Case "0" To "9"
       Case Else: Mid$(xYBIAMVT0.MOUVEMOPE, I, 1) = "X"
    End Select
Next I

If blnMT950_Loro Then
    xRef = "NTRFNONREF //" & Table_Ope_Unit(xYBIAMVT0.MOUVEMSER & xYBIAMVT0.MOUVEMSSE & xYBIAMVT0.MOUVEMOPE) & xYBIAMVT0.MOUVEMOPE & Format$(xYBIAMVT0.MOUVEMNUM, "000000000")
Else
    xRef = Table_Ope_Unit(xYBIAMVT0.MOUVEMSER & xYBIAMVT0.MOUVEMSSE & xYBIAMVT0.MOUVEMOPE) & xYBIAMVT0.MOUVEMOPE & Format$(xYBIAMVT0.MOUVEMNUM, "000000000")
End If

xLine61 = ":61:" & Mid$(xYBIAMVT0.MOUVEMDVA, 2, 6) & Mid$(xYBIAMVT0.MOUVEMDTR, 4, 4) & xSens & cur_AbsV_Dev(xYBIAMVT0.MOUVEMMON, mYBIAMVT0.COMPTEDEV) & xRef

X = UCase$(Trim(xYBIAMVT0.LIBELLIB1) & " " & Trim(xYBIAMVT0.LIBELLIB2))

'''' 2004.11.29 à remplcer par call SAA_Text_Control(lX As String, lenMax
'========================================================================
lenX = Len(X)
If lenX > 34 Then X = Mid$(X, 1, 34): lenX = 34
For I = 1 To lenX
    Select Case Mid$(X, I, 1)
        Case "A" To "Z":
        Case "0" To "9"
        Case ".", "-":
        Case Chr$(200), Chr$(201), Chr$(202): Mid$(X, I, 1) = "E"
       Case Else: Mid$(X, I, 1) = " "
    End Select
Next I

xSwift = xSwift & xLine61 & Asc13 & Asc10 & X & Asc13 & Asc10

mSolde = mSolde + xYBIAMVT0.MOUVEMMON

End Sub

Public Sub Swift_Mt950_Write()
Dim K As Integer, lenX As Long

lenX = Len(xSwift)
xSwift = xSwift & Space$(513)
For K = 1 To lenX Step 512
    If blnMT950_Loro Then
        Put #1, mFile1_Seq, Mid$(xSwift, K, 512)
        mFile1_Seq = mFile1_Seq + 512
    Else
        Put #2, mFile2_Seq, Mid$(xSwift, K, 512)
        mFile2_Seq = mFile2_Seq + 512
    End If
    
Next K

End Sub

Public Sub Swift_MT950_SAA_from_SAB()
Dim xFileName As String, xDest As String, X As String

Call FEU_ROUGE
If blnRelevéA4W_Loro Then

    xFileName = wAMJHMS & paramMT950_Loro & paramSAA_Data_from_SAB_ExtensionP_sav
    xDest = paramSAA_DataF_from_MT950 & wAMJHMS & paramMT950_Loro & paramSAA_Data_from_SAB_ExtensionP_pcc
    
        msFileSystem.MoveFile paramMT950F_Loro, paramSAA_DataF_from_MT950 & xFileName
        
        msFileSystem.CopyFile paramSAA_DataF_from_MT950 & xFileName, paramSAA_DataF_Archive & "\SAA_from_MT950_" & xFileName
        msFileSystem.MoveFile paramSAA_DataF_from_MT950 & xFileName, xDest
        Print #3, paramMT950F_Loro
        Print #3, "  (archivé et ) TRANSFERE vers " & xDest

Else
        Print #3, paramMT950F_Loro & " NON TRANSFERE vers SAA"

End If
Call FEU_VERT

End Sub
Public Sub Swift_MT950_Corona_from_SAB()
Dim xFileName As String, xDest As String, X As String

Call FEU_ROUGE
If blnRelevéA4W_Nostro Then
    xFileName = wAMJHMS & paramMT950_Nostro & paramSAA_Data_from_SAB_ExtensionP_sav
    xDest = paramCorona_DataF_Swift_In & wAMJHMS & paramMT950_Nostro & paramSAA_Data_from_SAB_ExtensionP_pcc

    msFileSystem.CopyFile paramMT950F_Nostro, paramSAA_DataF_Archive & "\SAB_to_CORONA" & xFileName
    msFileSystem.MoveFile paramMT950F_Nostro, xDest
    Print #3, paramMT950F_Nostro
    Print #3, "  (archivé et ) TRANSFERE vers " & xDest

Else
    Print #3, paramMT950F_Nostro & " NON TRANSFERE vers CORONA"

End If
Call FEU_VERT
End Sub


Public Sub SWIFT_MT950_MOUVEMDVA(lRELEVECOM As String, lMOUVEMDTR_Max As String)
Dim rsYBIAMVT0 As ADODB.Recordset
Dim xSQL As String, V
'___________________________________________________________________________________________
'$JPL 20120615 Champ 64

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
 & " where MOUVEMCOM = '" & lRELEVECOM & "'" _
 & " and MOUVEMDVA > " & IbmAmjMax _
 & " and MOUVEMDTR <= " & IbmAmjMax _
 & " order by MOUVEMDTR, MOUVEMPIE, MOUVEMECR"
 
Set rsYBIAMVT0 = cnsab.Execute(xSQL)
Do Until rsYBIAMVT0.EOF
    curMTD_MOUVEMDVA = curMTD_MOUVEMDVA + rsYBIAMVT0("MOUVEMMON")
    rsYBIAMVT0.MoveNext
Loop
End Sub
