Attribute VB_Name = "srvSAA"
Option Explicit

'???????????????????????????????????????
Public paramSwiftSABSAA_SAA_In As String
Public paramSwift_File_AMJHMS As String
'???????????????????????????????????????


Dim xIn As String, xOut As String, xIn1 As String

Dim mField As String, blnCreation As Boolean, blnUnit As Boolean

Public arrMsgFile_Printer(20) As String, arrMsgFile_Printer_Nb As Integer, arrMsgFile_Printer_NbMax As Integer, arrMsgFile_Printer_Index As Integer
Public arrMsgFile_Seq(20, 1000) As Integer, arrMsgFile_Seq_Nb As Integer, arrMsgFile_Seq_Index As Integer, arrMsgFile_Seq_AMJ(20, 1000) As String


Public meZSWIRAL0 As typeZSWIRAL0
Public arrZSWIRAL0() As typeZSWIRAL0, arrZSWIRAL0_NbMax As Integer, arrZSWIRAL0_Nb As Integer

Dim blnTransaction As Boolean

Dim oldYBIAMON0 As typeYBIAMON0, meYBIAMON0 As typeYBIAMON0
Public Sub SAA_to_Corona_Put()
Dim wAMJHMS As String
' vérifier si le fichier paramSwiftSABSAA_SAA_In n'existe pas
' rename les fichiers avant lecture
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Dim xFileName As String, I As Integer

On Error GoTo Error_Handle
xIn = "": xOut = ""
Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "SAA => Corona : début ...")
wAMJHMS = DSys & "_" & time_Hms & "_"

For I = 1 To frmSAA.fgSAA_to_Corona.Rows - 1

    frmSAA.fgSAA_to_Corona.Col = 0
    frmSAA.fgSAA_to_Corona.Row = I
    frmSAA.fgSAA_to_Corona.CellForeColor = warnUsrColor
    xFileName = frmSAA.fgSAA_to_Corona.Text
    Call lstErr_ChangeLastItem(frmSAA.lstErr, frmSAA.cmdContext, "TI > Swift : " & xFileName): DoEvents
    
    msFileSystem.CopyFile paramSAA_DataF_to_Corona & xFileName, paramSAA_DataF_Archive & "\SAA_to_Corona_" & wAMJHMS & xFileName

    msFileSystem.MoveFile paramSAA_DataF_to_Corona & xFileName, paramCorona_DataF_Swift_In & wAMJHMS & xFileName


    frmSAA.fgSAA_to_Corona.CellForeColor = vbGreen
Next I

Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "OK : "): DoEvents

Exit Sub
Error_Handle:
 Close
 MsgBox Error, vbCritical, "srvSAA.SAA_Corona_Put"
End Sub
Public Sub SAA_from_SAB()
Dim V
Dim xFileName As String, I As Integer, K As Integer, X As String
Dim arrFileName_Sab() As String
Dim pccFileName As String, savFileName As String
Dim wFolder As String, wName As String, wExtension As String
Dim wAMJHMS As String

On Error GoTo Error_Handle

wAMJHMS = DSys & "_" & time_Hms & "_"
xIn = "": xOut = ""
Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "SAB => SAA : début ...")

ReDim arrFileName_Sab(frmSAA.fgSAA_from_SAB.Rows)

For I = 1 To frmSAA.fgSAA_from_SAB.Rows - 1
    frmSAA.fgSAA_from_SAB.Col = 0
    frmSAA.fgSAA_from_SAB.Row = I
    frmSAA.fgSAA_from_SAB.CellForeColor = vbYellow 'warnUsrColor
    X = frmSAA.fgSAA_from_SAB.Text
    xFileName = X
    K = InStr(1, X, paramSAA_Data_from_SAB_ExtensionP_sab)
    arrFileName_Sab(I) = wAMJHMS & Mid$(X, 1, K - 1) & paramSAA_Data_from_SAB_ExtensionP_sav
    msFileSystem.MoveFile paramSAA_DataF_from_SAB & xFileName, paramSAA_DataF_from_SAB & arrFileName_Sab(I)
    
Next I

For I = 1 To frmSAA.fgSAA_from_SAB.Rows - 1
    frmSAA.fgSAA_from_SAB.Col = 0
    frmSAA.fgSAA_from_SAB.Row = I
    frmSAA.fgSAA_from_SAB.CellForeColor = warnUsrColor
    savFileName = paramSAA_DataF_from_SAB & arrFileName_Sab(I)
    
    Call fileName_Split(savFileName, wFolder, wName, wExtension)
    pccFileName = wFolder & wName & paramSAA_Data_from_SAB_ExtensionP_pcc
    Call lstErr_ChangeLastItem(frmSAA.lstErr, frmSAA.cmdContext, "SAB > SAA : " & pccFileName): DoEvents

    V = SAA_from_SAB_ZSWIALL0(savFileName, pccFileName)
    If Not IsNull(V) Then
        Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "? " & V)
    Else
        msFileSystem.MoveFile savFileName, paramSAA_DataF_Archive & "\SAA_from_SAB_" & arrFileName_Sab(I)
    End If
Next I

Close
    
    Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "OK : "): DoEvents
Exit Sub
Error_Handle:
 Close
 MsgBox Error, vbCritical, "srvSAA.SAA_from_SAB"

End Sub

Public Sub SAA_to_SAB(lblnAut_Swift As Boolean)
Dim V
Dim xFileName As String, I As Integer, K As Integer, X As String
Dim arrFileName_Sab() As String
Dim wAMJHMS As String
Dim wFile_Export As String

On Error GoTo Error_Handler
'------------------------------------------------------------------------------------
Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "SAA => SAB : YBIAMON7>SWIFT>ZSWIRAL0")
frmSAA.MousePointer = vbHourglass

meYBIAMON0.MONAPP = "SWIFT"
meYBIAMON0.MONFLUX = "ZSWIRAL0"
V = rsYBIAMON0_Read(meYBIAMON0)
If Not IsNull(V) Then GoTo Error_MsgBox
If Trim(meYBIAMON0.MONSTATUS) <> "" Then
    V = "Action précédente en cours : " & meYBIAMON0.MONAPP & "_" & meYBIAMON0.MONFLUX & " > " & meYBIAMON0.MONSTATUS
    GoTo Error_MsgBox
End If
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "SAA => SAB : INIT")
frmSAA.MousePointer = vbHourglass
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True
oldYBIAMON0 = meYBIAMON0
meYBIAMON0.MONSTATUS = "SAA_PUT"   ''"INIT"
meYBIAMON0.MONNUM = meYBIAMON0.MONNUM + 1
V = sqlYBIAMON0_Update(meYBIAMON0, oldYBIAMON0, True)
If Not IsNull(V) Then GoTo Error_MsgBox
oldYBIAMON0 = meYBIAMON0
'------------------------------------------------------------------------------------

wFile_Export = paramSAA_DataF_to_SAB & paramSAA_Data_to_SAB_YFile
'x = Dir(wFile_Export)
'If x <> "" Then
'    Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "? " & wFile_Export & " existe déjà")
'    Exit Sub
'End If
wAMJHMS = DSys & "_" & time_Hms & "_"
xIn = "": xOut = ""
ReDim arrZSWIRAL0(100): arrZSWIRAL0_NbMax = 100: arrZSWIRAL0_Nb = 0

ReDim arrFileName_Sab(frmSAA.fgSAA_to_SAb.Rows)

For I = 1 To frmSAA.fgSAA_to_SAb.Rows - 1
    frmSAA.fgSAA_to_SAb.Col = 0
    frmSAA.fgSAA_to_SAb.Row = I
    frmSAA.fgSAA_to_SAb.CellForeColor = vbYellow 'warnUsrColor
    X = frmSAA.fgSAA_to_SAb.Text
    xFileName = X
    K = InStr(1, X, paramSAA_Data_to_SAB_ExtensionP_out)
    arrFileName_Sab(I) = wAMJHMS & Mid$(X, 1, K - 1) & paramSAA_Data_to_SAB_ExtensionP_sav
    msFileSystem.MoveFile paramSAA_DataF_to_SAB & xFileName, paramSAA_DataF_to_SAB & arrFileName_Sab(I)
    
Next I

For I = 1 To frmSAA.fgSAA_to_SAb.Rows - 1
    frmSAA.fgSAA_to_SAb.Col = 0
    frmSAA.fgSAA_to_SAb.Row = I
    frmSAA.fgSAA_to_SAb.CellForeColor = warnUsrColor
    xFileName = paramSAA_DataF_to_SAB & arrFileName_Sab(I)
    Call lstErr_ChangeLastItem(frmSAA.lstErr, frmSAA.cmdContext, "SAB > Swift : " & xFileName): DoEvents

    V = SAA_to_SAB_ZSWIRAL0(xFileName, wFile_Export)
    If Not IsNull(V) Then Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "? " & V)
    
Next I

Close
'------------------------------------------------------------------------------------
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "SAA => SAB : ZSWIRAL0")
frmSAA.MousePointer = vbHourglass
meYBIAMON0.MONSTATUS = "ZSWIRAL0"
V = sqlYBIAMON0_Update(meYBIAMON0, oldYBIAMON0, False)
If Not IsNull(V) Then GoTo Error_MsgBox
oldYBIAMON0 = meYBIAMON0
'------------------------------------------------------------------------------------
Call incrementeSAV
Call ZSWIRAL0_Est_Vide
For I = 1 To arrZSWIRAL0_Nb
    DoEvents
    V = sqlZSWIRAL0_Insert(arrZSWIRAL0(I))
    If Not IsNull(V) Then GoTo Transaction_End
Next I
'------------------------------------------------------------------------------------
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "SAA => SAB : Archive")
frmSAA.MousePointer = vbHourglass
meYBIAMON0.MONSTATUS = "Archive"
V = sqlYBIAMON0_Update(meYBIAMON0, oldYBIAMON0, False)
If Not IsNull(V) Then GoTo Error_MsgBox
oldYBIAMON0 = meYBIAMON0
'------------------------------------------------------------------------------------

For I = 1 To frmSAA.fgSAA_to_SAb.Rows - 1
    msFileSystem.MoveFile paramSAA_DataF_to_SAB & arrFileName_Sab(I), paramSAA_DataF_Archive & "\SAA_to_SAB_" & arrFileName_Sab(I)
    
Next I
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "OK : " & wFile_Export): DoEvents

Transaction_End:
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'------------------------------------------------------------------------------------
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "SAA => SAB : FIN")
frmSAA.MousePointer = vbHourglass
meYBIAMON0.MONSTATUS = "SAA_END" ' ""
V = sqlYBIAMON0_Update(meYBIAMON0, oldYBIAMON0, False)
If Not IsNull(V) Then GoTo Error_MsgBox
'------------------------------------------------------------------------------------

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        meYBIAMON0.MONSTATUS = "Rollback"
    Else
        V = cnSAB_Transaction("Commit")
        meYBIAMON0.MONSTATUS = ""
    End If
End If

Exit Sub


Error_Handler:
V = Error
Error_MsgBox:
 If blnTransaction Then Call cnSAB_Transaction("Rollback")
 Close
 If Not lblnAut_Swift Then MsgBox V, vbCritical, "srvSAA.SAA_to_SAB"
End Sub

Public Function SAA_to_SAB_ZSWIRAL0(lFile_Import As String, lFile_Export As String)

Dim xIn As String, xIn2 As String, xSwift As String
Dim xK As Integer, lenX As Long, lenSwift As Long
Dim wFile_Import As String, wFile_Export As String
Dim I As Long
On Error GoTo Error_Handle

SAA_to_SAB_ZSWIRAL0 = "?"
rsZSWIRAL0_Init meZSWIRAL0
 
wFile_Import = Trim(lFile_Import)
If Dir(wFile_Import) = "" Then SAA_to_SAB_ZSWIRAL0 = "! pas de fichier : " & wFile_Import: GoTo Error_Handle
wFile_Export = Trim(lFile_Export)

Open wFile_Import For Input As #1
''Open wFile_Export For Append As #2

xSwift = ""
Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    xIn2 = Trim(xIn): lenX = Len(xIn2)
    xK = InStr(1, xIn2, Asc03)
    If xK <= 0 Then
        xSwift = xSwift & xIn2 & Chr$(13) & Chr$(10) 'Chr$(13) & Chr$(13) ''
    Else
        xSwift = xSwift & Mid$(xIn2, 1, xK)
        lenSwift = Len(Trim(xSwift))
        xSwift = xSwift & Space$(512)
        For I = 1 To lenSwift Step 512
            arrZSWIRAL0_Nb = arrZSWIRAL0_Nb + 1
             If arrZSWIRAL0_Nb > arrZSWIRAL0_NbMax Then
                 arrZSWIRAL0_NbMax = arrZSWIRAL0_NbMax + 100
                 ReDim Preserve arrZSWIRAL0(arrZSWIRAL0_NbMax)
             End If
            arrZSWIRAL0(arrZSWIRAL0_Nb).SWIRALDON = Mid$(xSwift, I, 512)
            arrZSWIRAL0(arrZSWIRAL0_Nb).SWIRALETA = currentZMNURUT0.MNURUTETB                  '
            arrZSWIRAL0(arrZSWIRAL0_Nb).SWIRALMES = ""               '

           
        Next I
        If xK < lenX Then
            xSwift = Trim(Mid$(xIn2, xK + 1, lenX - xK)) & Chr$(13) & Chr$(10) ''& Chr$(13) & Chr$(13)  ''
        Else
            xSwift = ""
        End If
   End If
Loop

'vérifier si les champs 44E et 44F existent dans le message MODIFICATIONS APPORTEES PAR KOKOU AYAWLI LE 20/12/2023 SUITE A LA RELEASE SWIFT DE NOV  2023
'vérifier si les champs 44E et 44F existent dans le message MODIFICATIONS APPORTEES PAR KOKOU 19/12/2023 SUITE A LA RELEASE SWIFT DE NOV  2023

Dim taille44E
Dim taille44F
Dim balise44E
Dim balise44F
Dim contenuFichier
Dim positionFin44E
Dim positionFin44F
Dim positionDebut44E
Dim positionDebut44F
Dim z As Integer

contenuFichier = ""

For z = 1 To arrZSWIRAL0_Nb
contenuFichier = contenuFichier & arrZSWIRAL0(z).SWIRALDON
Next


Dim tableauSWIFT() As String

tableauSWIFT = Split(contenuFichier, "{1:")

' chaque élément du tableau
Dim T As Integer
For T = LBound(tableauSWIFT) To UBound(tableauSWIFT)
    taille44E = 0
    taille44F = 0
    positionDebut44E = 0
    positionDebut44F = 0
    positionFin44E = 0
    positionFin44F = 0
    balise44E = ""
    balise44F = ""
'    Debug.Print ("___________________________________________________________________")
'    Debug.Print tableauSWIFT(t)
'    Debug.Print t
    
    positionDebut44E = InStr(tableauSWIFT(T), Chr$(13) & Chr$(10) & ":44E:")
positionDebut44F = InStr(tableauSWIFT(T), Chr$(13) & Chr$(10) & ":44F:")
If positionDebut44E > 0 Then

    positionFin44E = InStr(positionDebut44E + 1, tableauSWIFT(T), Chr$(13) & Chr$(10) & ":")
    If (positionFin44E = 0) Then
    positionFin44E = InStr(positionDebut44E + 1, tableauSWIFT(T), "-}")
    End If
    
    If positionFin44E > positionDebut44E Then
        balise44E = Mid(tableauSWIFT(T), positionDebut44E, positionFin44E - positionDebut44E)
        balise44E = Trim(balise44E)
        taille44E = Len(balise44E)
    End If
    
End If

If positionDebut44F > 0 Then

    positionFin44F = InStr(positionDebut44F + 1, tableauSWIFT(T), Chr$(13) & Chr$(10) & ":")
    If (positionFin44F = 0) Then
    positionFin44F = InStr(positionDebut44F + 1, tableauSWIFT(T), "-}")
    End If
    
    If positionFin44F > positionDebut44F Then
        balise44F = Mid(tableauSWIFT(T), positionDebut44F, positionFin44F - positionDebut44F)
        balise44F = Trim(balise44F)
        taille44F = Len(balise44F)
    End If
    
End If



If taille44E > 65 Or taille44F > 65 Then
    Dim type_message As String
    Dim bic As String
    Dim position_debut As Integer
    Dim position_fin As Integer
    
    type_message = ""
    bic = ""
    position_debut = 0
    position_fin = 0
    
    position_debut = InStr(tableauSWIFT(T), "{2:O")
    If position_debut > 0 Then
        type_message = Mid(tableauSWIFT(T), position_debut + 4, 3)
        bic = Mid(tableauSWIFT(T), position_debut + 17, 12)
    End If
    
    
    
    Dim objMessage As Object
    Set objMessage = CreateObject("CDO.Message")
    
    contenuFichier = "{1:" & Replace(tableauSWIFT(T), Chr$(13) & Chr$(10), "</br>")
    
    ' Paramètres de configuration du serveur SMTP
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "exg2016a"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 ' cdoBasic
    'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = ""
    'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""
    objMessage.Configuration.Fields.Update
    
    ' Destinataire, expéditeur, objet et corps du message
    objMessage.To = "CREDOC@bia-paris.fr"
    objMessage.CC = "BENAISSA.L@bia-paris.fr;AIDI.M@bia-paris.fr; ayawli.k@bia-paris.fr;"
    objMessage.From = "bia_info@bia-paris.fr"
    objMessage.Subject = type_message & "_" & bic & "_Message entrant tronqué dans SAB"
    'objMessage.TextBody = "CECI EST UN TEST"
    'objMessage.HTMLBody = "<html><body style=background-color: red><p>Champ 44E ou 44F tronqué.</br>Veuillez trouver le message complet dans ce mail</br></br><table><tr style=background-color: white;>" & contenuFichier & "</tr></table> Cordialement,</br></p></body></html>"
    objMessage.HTMLBody = "<html><body style='background-color: red;'><p>Message entrant dont le champ 44E (" & taille44E & " caractères) ou 44F (" & taille44F & " caractères) tronqué dans SAB.</br></br><table><tr style='background-color: yellow;'>" & balise44E & "</tr><tr style='background-color: yellow;'>" & balise44F & "</tr></table></br>Veuillez trouver le message complet dans ce mail : </br></br><table><tr style='background-color: white;'>" & contenuFichier & "</tr></table> Cordialement,</br></p></body></html>"

    
    ' Envoyer l'e-mail
    objMessage.Send
    
    ' Libérer l'objet
    Set objMessage = Nothing

End If
'    Debug.Print ("************")
'    Debug.Print balise44E
'    Debug.Print balise44F
'    Debug.Print taille44E
'    Debug.Print taille44F
'    Debug.Print type_message
'    Debug.Print bic
'    Debug.Print ("___________________________________________________________________")
Next T





' FIN AJOUT KOKOU


' FIN AJOUT  KOKOU

Close

SAA_to_SAB_ZSWIRAL0 = Null

GoTo fin

Error_Handle:

xIn = lFile_Import & ":" & Error
MsgBox xIn, vbCritical, "SAA_ZSWIRAL0"
SAA_to_SAB_ZSWIRAL0 = xIn

fin:

Close

End Function


Public Function SAA_from_SAB_ZSWIALL0(lsabFile As String, lpccFile As String)

Dim xIn As String, xIn2 As String, xSwift As String
Dim Seq As Long, xK As Integer, lenX As Integer, lenSwift As Integer
Dim paramImport As String, paramExport As String
Dim I As Integer, K As Integer, K1 As Integer, K2 As Integer, K3 As Integer
Dim Seq1 As Long, Seq2 As Long
Dim x512_CRLF As String
Dim blnCRLF_Skip As Boolean

Dim MT_Nb As Integer
On Error GoTo Error_Handle


SAA_from_SAB_ZSWIALL0 = "?"
 
paramImport = Trim(lsabFile)
If Dir(paramImport) = "" Then SAA_from_SAB_ZSWIALL0 = "! pas de fichier : " & paramImport: GoTo Error_Handle
paramExport = Trim(lpccFile)

Open paramImport For Input As #1
Open paramExport For Binary Access Write As #2 'Len = 512
Seq1 = 1: Seq2 = 1
xSwift = "": x512_CRLF = ""
MT_Nb = 0
blnCRLF_Skip = False

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Len(xIn) > 0 Then
        xSwift = xSwift & xIn
    
        K3 = InStr(1, xIn, Asc03)
    
        If K3 > 0 Then Call SAA_from_SAB_DOS_PCC(xSwift, Seq2): xSwift = "": MT_Nb = MT_Nb + 1: blnCRLF_Skip = False
    End If
    
Loop

Call SAA_from_SAB_DOS_PCC(xSwift, Seq2)

Close

SAA_from_SAB_ZSWIALL0 = Null

GoTo fin

Error_Handle:

SAA_from_SAB_ZSWIALL0 = xIn
Shell_MsgBox "# SAA_from_SAB_ZSWIALL0 # ", vbCritical, lsabFile & ":" & Error, True

fin:

Close



End Function





'
Public Sub Nostro_Put()

Dim xFileName As String, I As Integer

On Error GoTo Error_Handle
paramSwift_File_AMJHMS = "\Nostro_Snd_" & DSys & "_" & time_Hms
Open paramSwiftNostro_Corona_Wait & paramSwift_File_AMJHMS For Output As #2
xIn = "": xOut = ""
Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "Nostro => Corona : début ...")

Open paramSwiftNostro_MT950_File For Input As #1
Do Until EOF(1)
     DoEvents
     Line Input #1, xIn
     xOut = Trim(xIn)
     If xOut <> "" Then Print #2, xOut
Loop


Close


Kill paramSwiftNostro_MT950_File
msFileSystem.MoveFile paramSwiftNostro_Corona_Wait & paramSwift_File_AMJHMS, paramSwiftNostro_Corona_In & paramSwift_File_AMJHMS
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "OK : " & paramSwiftNostro_Corona_In & paramSwift_File_AMJHMS): DoEvents

Exit Sub
Error_Handle:
Close
MsgBox Error, vbCritical, "srvSAA.Nostro_Put"
Kill paramSwiftNostro_Corona_Wait & paramSwift_File_AMJHMS
End Sub

Public Sub Loro_Put()

Dim xFileName As String, I As Integer

On Error GoTo Error_Handle
paramSwift_File_AMJHMS = "\Loro_Snd_" & DSys & "_" & time_Hms
Open paramSwiftLoro_SAA_Wait & paramSwift_File_AMJHMS For Output As #2
xIn = "": xOut = ""
Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "Loro => SAA : début ...")

Open paramSwiftLoro_MT950_File For Input As #1
Do Until EOF(1)
     DoEvents
     Line Input #1, xIn
     xOut = Trim(xIn)
     If xOut <> "" Then Print #2, xOut
Loop


Close


Kill paramSwiftLoro_MT950_File
msFileSystem.MoveFile paramSwiftLoro_SAA_Wait & paramSwift_File_AMJHMS, paramSwiftLoro_SAA_In & paramSwift_File_AMJHMS
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "OK : " & paramSwiftSABSAA_SAA_In & paramSwift_File_AMJHMS): DoEvents

Exit Sub
Error_Handle:
Close
MsgBox Error, vbCritical, "srvSAA.Loro_Put"
Kill paramSwiftLoro_SAA_Wait & paramSwift_File_AMJHMS
End Sub





Public Sub ImportBIC_Load(lFile As String)
Dim xFileName As String, I As Integer, lenX As Integer
Dim X As String, xIn As String, xIn2 As String
Dim wSWIBIC As String
Dim Nb As Long
On Error GoTo Error_Handle
paramSwift_BIC_YFile = "C:\TEMP\YSWIBIC0.txt"
X = Dir(paramSwift_BIC_YFile)
If X <> "" Then
    Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "? " & paramSwift_BIC_YFile & " existe déjà")
''''    Exit Sub
End If
paramSwift_File_AMJHMS = DSys & "_" & time_Hms & "_"
xIn = "": xOut = ""
Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "BIC => SAB : début ...")

Open lFile For Input As #1
Open paramSwift_BIC_YFile For Output As #2
Nb = 0
Do Until EOF(1)
    DoEvents
    xIn = Input(856, #1)
    If Mid$(xIn, 1, 2) = "FI" Then
        Nb = Nb + 1
' SWIBICBIC & SWIBICINT & SWIBICVIL & SWIBICCOM
'-----------------------------------------------
        
        wSWIBIC = Mid$(xIn, 4, 11) _
                & Mid$(xIn, 15, 105) _
                & Mid$(xIn, 190, 35) _
                & Mid$(xIn, 120, 70)
            Print #2, wSWIBIC
    Else
        Debug.Print xIn
            
   End If
Loop
Close

Call Shell_FTP(paramSwift_BIC_YFile, paramIBM_Library_SABSPE, "YSWIBIC0", False, False)
MsgBox "Temporisation FTP", vbInformation

X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YSWIBIC0 "
Set rsSab = cnsab.Execute(X)
If rsSab("Tally") <> Nb Then
    MsgBox " : Lu : " & Nb & " nb FTP " & rsSab("Tally"), vbCritical, paramIBM_AS400_ID & "FTP => YSWIBIC0"
Else
    Kill paramSwift_BIC_YFile
        
    Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "OK : " & paramSwift_BIC_YFile): DoEvents
    
    MsgBox paramIBM_AS400_ID & " : faire CPYF SAB073SPE / YSWIBIC0  => SAB073 / ZSWIBIC0 *Replace", vbInformation, paramIBM_AS400_ID & "Import BIC, nb item  :  " & Nb
End If

Exit Sub
Error_Handle:
 Close
 MsgBox Error, vbCritical, "srvSAA.SAA_to_SAB"

End Sub

Public Sub ImportMsgFile_Load(lFile As String)
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Integer
Dim mPrinter As String, mPrinter_len As Integer
Dim K1 As Integer, I1 As Integer, I As Integer

Seq = 0: arrMsgFile_Seq_Nb = 0

arrMsgFile_Printer_NbMax = 20
arrMsgFile_Printer_Nb = 0
For I = 1 To arrMsgFile_Printer_NbMax
    arrMsgFile_Printer(I) = ""
    arrMsgFile_Seq(I, 0) = 0
Next I

mPrinter = " Sent to APPLI " & Chr$(34) & "Printer"   '& lPrinter & "Rcv"
mPrinter_len = Len(mPrinter)

Open lFile For Input As #1
Seq = 0: arrMsgFile_Seq_Nb = 0
Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "import : " & lFile)
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "Lecture : ")

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If Mid$(xIn, 1, 13) = "U-UMID      =" Then
            Seq = Seq + 1
            Call lstErr_ChangeLastItem(frmSAA.lstErr, frmSAA.cmdContext, Mid$(xIn, 15, 11) & Seq)
        End If
         If Mid$(xIn, 1, mPrinter_len) = mPrinter Then
            X = Mid$(xIn, 1 + mPrinter_len, 4)
            arrMsgFile_Printer_Index = ImportMsgFile_Load_Printer(X)
            arrMsgFile_Seq_Nb = arrMsgFile_Seq(arrMsgFile_Printer_Index, 0)
           If arrMsgFile_Seq(arrMsgFile_Printer_Index, arrMsgFile_Seq_Nb) < Seq Then
                arrMsgFile_Seq_Nb = arrMsgFile_Seq_Nb + 1
                arrMsgFile_Seq(arrMsgFile_Printer_Index, 0) = arrMsgFile_Seq_Nb
                arrMsgFile_Seq(arrMsgFile_Printer_Index, arrMsgFile_Seq_Nb) = Seq
            End If
            
            I1 = InStr(mPrinter_len, xIn, "on ")
            K1 = Len(xIn) - I1 - 2
            If I1 > 0 Then arrMsgFile_Seq_AMJ(arrMsgFile_Printer_Index, arrMsgFile_Seq_Nb) = Mid$(xIn, I1 + 3, K1)
'2001.06.20 jpl
'            arrMsgFile_Seq_Nb = arrMsgFile_Seq_Nb + 1
'            arrMsgFile_Seq(arrMsgFile_Seq_Nb) = Seq
'            K1 = Len(xIn) - mPrinter_len - 2
'            arrMsgFile_Seq_AMJ(arrMsgFile_Seq_Nb) = mId$(xIn, mPrinter_len + 5, K1)
        End If
    End If
 '   Debug.Print Xin
Loop


Close
Call lstErr_ChangeLastItem(frmSAA.lstErr, frmSAA.cmdContext, "Sélectionner un printer / " & Seq)


Exit Sub

Error_Handle:
 MsgBox "erreur : srvSwift:ImportMsgFile_Load" & xIn, vbCritical, Error
Close


End Sub
Public Sub ImportMsgFile_Print(lFile, lPrinter As String)
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Integer
Dim mPrinter As String, mPrinter_len As Integer
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer
Dim xRouting As String, blnRouting As Boolean, xRoutingK As String
Dim blnSession As Boolean, xSession As String

arrMsgFile_Printer_Index = 0

For I = 1 To arrMsgFile_Printer_Nb
    If arrMsgFile_Printer(I) = lPrinter Then arrMsgFile_Printer_Index = I
Next I

Open lFile For Input As #1
Seq = 0:
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "Impression : " & lPrinter)
blnOk = False: blnPrint = False
arrMsgFile_Seq_Index = 1

prtSwift_Open lPrinter & " : Swift messages  "
Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If Mid$(xIn, 1, 13) = "U-UMID      =" Then
            If blnOk Then
               prtSwiftMsgFile_Line xSession, 2
               prtSwiftMsgFile_Line xRouting, 4
           End If
            Seq = Seq + 1
            If Seq = arrMsgFile_Seq(arrMsgFile_Printer_Index, arrMsgFile_Seq_Index) Then
                blnOk = True: blnPrint = False: blnRouting = False
                Call lstErr_ChangeLastItem(frmSAA.lstErr, frmSAA.cmdContext, "impression : " & Mid$(xIn, 15, 11))
                xOut = "MT" & Mid$(xIn, 27, 3) & "  :   " & arrMsgFile_Seq_AMJ(arrMsgFile_Printer_Index, arrMsgFile_Seq_Index) & "   " & Mid$(xIn, 16, 11)
                ''''prtSwiftMsgFile_Line x, 1
                If IsNumeric(Mid$(xIn, 27, 3)) Then
                    blnSwift = True
                Else
                    blnSwift = False
                End If
                xRouting = ""
                blnSession = False: xSession = ""
                arrMsgFile_Seq_Index = arrMsgFile_Seq_Index + 1
                ''Exit Do
            Else
                blnOk = False
            End If
        End If
        If blnOk Then
            Select Case Mid$(xIn, 1, 13)
                Case "Sender       ": ImportMsgFile_Load_Sender
                Case "Receiver     ": blnPrint = False
                Case "Transaction r": prtSwiftMsgFile_Line xIn, 20   'ImportMsgFile_Load_Amount
                Case "Amount      =": prtSwiftMsgFile_Line xIn, 2

                Case "Text         ":
                    If blnSwift Then
                        ImportMsgFile_Load_text
                    Else
                        blnPrint = True: kPrint = 0: Line Input #1, xIn: Line Input #1, xIn
                    End If
                    
                Case "Block 5:": blnPrint = False
                Case "Message Histo": blnPrint = False: blnRouting = True: xRoutingK = "?"
                Case " Sent to APPL":
                        xRouting = xRouting & xRoutingK & ImportMsgFile_Load_Routing(xIn)
                        If InStr(xIn, lPrinter) > 0 Then
                            blnSession = True
                            xSession = xSession & xIn
                        Else
                            blnSession = False
                        End If
            End Select
            If blnRouting Then
                Select Case Mid$(xIn, 1, 5)
                    Case "*Orig": xRoutingK = ">>>"
                    Case "*Copy": xRoutingK = " _ "
                End Select
            End If
            If blnSession Then
                If InStr(xIn, "Session") > 0 Then
                            blnSession = False
                            xSession = xSession & xIn
                End If
            End If
            If blnPrint Then
                prtSwiftMsgFile_Line xIn, kPrint
            End If
        End If
   End If
Loop

Close

If blnOk Then
    prtSwiftMsgFile_Line xSession, 2
    prtSwiftMsgFile_Line xRouting, 4
End If
prtSwift_Close

Exit Sub

Error_Handle:
 MsgBox "erreur : srvSwift:ImportMsgFile_Load" & xIn, vbCritical, Error
Close


End Sub


Public Sub ImportMsgFile_Load_Amount()
Line Input #1, xIn1

prtSwiftMsgFile_Line xIn1 & "  " & xIn, 2

End Sub
Public Sub ImportMsgFile_Load_Sender()
Dim I As Integer
Line Input #1, xIn
Line Input #1, xIn

For I = 1 To 6
    Line Input #1, xIn
    xOut = xOut & "  " & Trim(xIn)
Next I

prtSwiftMsgFile_Line xOut, 1
End Sub

Public Sub ImportMsgFile_Load_text()
Dim I As Integer
Line Input #1, xIn
Line Input #1, xIn
xOut = xIn

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Mid$(xIn, 1, 8) = "Block 5:" Then Exit Do
    
    If Mid$(xIn, 1, 1) = ":" Then
        prtSwiftMsgFile_Line xOut, 3
        xOut = ""
    End If
    xOut = xOut & Trim(xIn)

Loop

prtSwiftMsgFile_Line xOut, 3

End Sub


Public Function ImportMsgFile_Load_Routing(lX As String) As String
Dim I1 As Integer, I2 As Integer

ImportMsgFile_Load_Routing = ""
I1 = InStr(13, lX, Asc34)
If I1 > 0 Then
    I2 = InStr(I1 + 1, lX, Asc34)
    If I2 > I1 Then ImportMsgFile_Load_Routing = Mid$(lX, I1 + 1, I2 - I1 - 1)
End If

End Function

Public Function ImportMsgFile_Load_Printer(lX As String) As Integer
Dim I As Integer

ImportMsgFile_Load_Printer = 0

For I = 1 To arrMsgFile_Printer_Nb
    If lX = arrMsgFile_Printer(I) Then ImportMsgFile_Load_Printer = I: Exit Function
Next I
If I >= arrMsgFile_Printer_NbMax Then
    MsgBox "ImportMsgFile_Load_Printer : trop de printer", vbCritical
Else
    arrMsgFile_Printer_Nb = arrMsgFile_Printer_Nb + 1
    arrMsgFile_Printer(arrMsgFile_Printer_Nb) = lX
    ImportMsgFile_Load_Printer = arrMsgFile_Printer_Nb
End If
End Function

Public Sub SAA_from_SAB_DOS_PCC(lSwift As String, lSeq As Long)
Dim K As Long, K2 As Long, K3 As Integer, K4 As Integer, lenX As Long
Dim TRN_Unit As String * 4, Dossier As String
Dim xTRN_SAB As String
Dim xTRN_BIA As String
Dim X As String


Dim wBlock1 As String, wBlock2 As String, wBlock3 As String

On Error GoTo Error_Handle

lenX = Len(lSwift)
If lenX > 0 Then

    For K = 1 To lenX
        Select Case Mid$(lSwift, K, 1)
            Case Chr$(&HC): Mid$(lSwift, K, 1) = Asc13
            Case Chr$(&HB): Mid$(lSwift, K, 1) = Asc10
            Case Chr$(&HE9): Mid$(lSwift, K, 1) = Chr$(&H7B)
            Case Chr$(&HE8): Mid$(lSwift, K, 1) = Chr$(&H7D)
            Case Asc34: Mid$(lSwift, K, 1) = " "
     End Select
    Next K
    
    K2 = InStr(4, lSwift, "{2:")
    K3 = InStr(4, lSwift, "{3:")
    K4 = InStr(4, lSwift, "{4:")
    wBlock1 = "{1:F01" & paramBic8 & "AXXX0000000000}"
    If K3 > 0 Then
        wBlock2 = Mid$(lSwift, K2, K3 - K2)
        wBlock3 = Mid$(lSwift, K3, K4 - K3)
    Else
        wBlock2 = Mid$(lSwift, K2, K4 - K2)
        wBlock3 = ""
    End If
    Call SAA_from_SAB_Block2(wBlock2)
    lSwift = Asc01 & wBlock1 & wBlock2 & wBlock3 & Mid$(lSwift, K4, lenX - K4 + 1)
    
    lenX = Len(lSwift)

'****************spécial BIARFRPP
    K = InStr(4, lSwift, ":20:")
    If K > 0 Then
        K = K + 4
        K2 = InStr(K, lSwift, ":")
        If K2 > 0 Then
            xTRN_SAB = Mid$(lSwift, K, K2 - K - 2)
            MsgBox "à faire xTRN_BIA = SAA_from_SAB_TRN(xTRN_SAB)"
            lSwift = Mid$(lSwift, 1, K - 1) & xTRN_BIA & Asc13 & Asc10 & Mid$(lSwift, K2, lenX - K2 + 2)
        End If
    End If
    lenX = Len(lSwift)
'****************spécial BIARFRPP
  
    lSwift = lSwift & Space$(534)
    For K = 1 To lenX Step 512
        If Trim(Mid$(lSwift, K, 512)) <> "" Then
            Put #2, lSeq, Mid$(lSwift, K, 512)
            lSeq = lSeq + 512
        End If
        
   Next K
End If

lSwift = ""
Exit Sub

Error_Handle:
MsgBox xIn, vbCritical, "SAA_YSWIALI0"

End Sub

Public Sub SAA_from_SAB_Block2(lBlock2 As String)


Select Case Mid$(lBlock2, 1, 19)
    Case "{2:I754BEXADZALXXXX": Mid$(lBlock2, 1, 19) = "{2:I754BEXADZALXDOE"
    Case "{2:I202CRESCHZZXXXX": Mid$(lBlock2, 1, 19) = "{2:I202CRESCHZZX80A"
    Case "{2:I103CRESCHZZXXXX": Mid$(lBlock2, 1, 19) = "{2:I103CRESCHZZX80A"
End Select


End Sub

