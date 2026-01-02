Attribute VB_Name = "srvSwift"
Option Explicit
Dim xIn As String, xOut As String, xIn1 As String

Public paramSwift_File_AMJHMS As String
Public paramSwiftTiSaa_SAA_Wait As String
Public paramSwiftTiSaa_SAA_In As String
Public paramSwiftTiSaa_TI_Out As String
Public paramSwiftTiSaa_TI_Pattern As String
Public paramSwiftTiSaa_TI_Archive As String

Public paramSwiftSaaTi_SAA_Out As String
Public paramSwiftSaaTi_TI_Wait As String
Public paramSwiftSaaTi_TI_File As String

Public paramSwiftSaaCorona_SAA_Out As String
Public paramSwiftSaaCorona_Corona_Wait As String
Public paramSwiftSaaCorona_Corona_In As String

Public paramSwiftLoro_SAA_In As String
Public paramSwiftLoro_SAA_Wait As String
Public paramSwiftLoro_MT950_File As String

Public paramSwiftNostro_Corona_In As String
Public paramSwiftNostro_Corona_Wait As String
Public paramSwiftNostro_MT950_File As String

Public paramSwiftHisto_Input As String
Dim mField As String, blnCreation As Boolean, blnUnit As Boolean

Public arrMsgFile_Printer(20) As String, arrMsgFile_Printer_Nb As Integer, arrMsgFile_Printer_NbMax As Integer, arrMsgFile_Printer_Index As Integer
Public arrMsgFile_Seq(20, 1000) As Integer, arrMsgFile_Seq_Nb As Integer, arrMsgFile_Seq_Index As Integer, arrMsgFile_Seq_AMJ(20, 1000) As String

Public Function param_Init()
Dim V
param_Init = Null

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "Swift"
Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "BIA.mdb : table : " & recElpTable.Id)

recElpTable.K1 = "TI_SAA"

recElpTable.K2 = "TI_Out"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftTiSaa_TI_Out = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "TI_Pattern"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftTiSaa_TI_Pattern = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "TI_Archive"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftTiSaa_TI_Archive = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "SAA_Wait"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftTiSaa_SAA_Wait = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "SAA_In"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftTiSaa_SAA_In = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))


recElpTable.K1 = "SAA_TI"

recElpTable.K2 = "SAA_Out"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftSaaTi_SAA_Out = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "TI_Wait"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftSaaTi_TI_Wait = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "TI_File"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftSaaTi_TI_File = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))
recElpTable.K1 = "Loro"

recElpTable.K2 = "SAA_In"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftLoro_SAA_In = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "SAA_Wait"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftLoro_SAA_Wait = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "MT950_File"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftLoro_MT950_File = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Nostro"

recElpTable.K2 = "MT950_File"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftNostro_MT950_File = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Corona_In"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftNostro_Corona_In = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Corona_Wait"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftNostro_Corona_Wait = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))


recElpTable.K1 = "SAA_Corona"

recElpTable.K2 = "SAA_Out"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftSaaCorona_SAA_Out = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Corona_Wait"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftSaaCorona_Corona_Wait = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))
V = dbElpTable_ReadE(recElpTable)

recElpTable.K2 = "Corona_In"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSwiftSaaCorona_Corona_In = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "BIA.mdb : table : " & recElpTable.Id & ": ok ")


Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "srvTI.Param_Init"
Exit Function

Num_Error:
param_Init = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "srvTI.Param_Init"
End Function

Public Sub SAA_Corona_Put()

' vérifier si le fichier paramSwiftTiSaa_SAA_In n'existe pas
' rename les fichiers avant lecture
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Dim xFileName As String, I As Integer

On Error GoTo Error_Handle
xIn = "": xOut = ""
Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "SAA => Corona : début ...")

For I = 1 To frmSwift.fgSAA_Corona.Rows - 1

    frmSwift.fgSAA_Corona.Col = 0
    frmSwift.fgSAA_Corona.Row = I
    frmSwift.fgSAA_Corona.CellForeColor = warnUsrColor
    xFileName = "\" & frmSwift.fgSAA_Corona.Text
    Call lstErr_ChangeLastItem(frmSwift.lstErr, frmSwift.cmdContext, "TI > Swift : " & xFileName): DoEvents
    msFileSystem.MoveFile paramSwiftSaaCorona_SAA_Out & xFileName, paramSwiftSaaCorona_Corona_Wait & xFileName
    msFileSystem.MoveFile paramSwiftSaaCorona_Corona_Wait & xFileName, paramSwiftSaaCorona_Corona_In & xFileName
    frmSwift.fgSAA_Corona.CellForeColor = vbGreen
Next I

Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, "OK : " & paramSwiftTiSaa_SAA_In & paramSwift_File_AMJHMS): DoEvents

Exit Sub
Error_Handle:
 Close
 MsgBox Error, vbCritical, "srvSwift.SAA_Corona_Put"
End Sub
Public Sub TI_SAA_Put()

' vérifier si le fichier paramSwiftTiSaa_SAA_In n'existe pas
' rename les fichiers avant lecture
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Dim xFileName As String, I As Integer, blnNext As Boolean
Dim kMT730 As Integer, blnMT730 As Boolean

On Error GoTo Error_Handle
paramSwift_File_AMJHMS = "\TI_Snd_" & DSys & "_" & time_Hms
Open paramSwiftTiSaa_SAA_Wait & paramSwift_File_AMJHMS For Output As #2
blnNext = False
xIn = "": xOut = ""
Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "TI > Swift : début ...")

For I = 1 To frmSwift.fgTI_SAA.Rows - 1

    frmSwift.fgTI_SAA.Col = 0
    frmSwift.fgTI_SAA.Row = I
    frmSwift.fgTI_SAA.CellForeColor = warnUsrColor
    xFileName = paramSwiftTiSaa_TI_Out & "\" & frmSwift.fgTI_SAA.Text
    Call lstErr_ChangeLastItem(frmSwift.lstErr, frmSwift.cmdContext, "TI > Swift : " & xFileName): DoEvents
    Open xFileName For Input As #1
    Do Until EOF(1)
        DoEvents
        Line Input #1, xIn
        
        If mId$(xIn, 1, 18) = "{1:F01BIARFRPP1XXX" Then
            Mid$(xIn, 1, 18) = "{1:F01BIARFRPPAXXX"
            kMT730 = InStr(18, xIn, "{2:I730")
            If kMT730 > 0 Then blnMT730 = True
        End If
        
        If Trim(xIn) = "*** SWIFT MESSAGE ***" Then
             blnMT730 = False
            If xOut <> "" Then blnNext = True
        Else
            If blnMT730 Then
                If blnNext Then
                    blnNext = False
                    xOut = xOut & "$" & xIn
                Else
                    If xOut <> "" Then Print #2, xOut
                    xOut = xIn
                End If
            End If
        End If
    Loop
    
    Close #1
Next I

If xOut <> "" Then Print #2, xOut

Close


For I = 1 To frmSwift.fgTI_SAA.Rows - 1
    frmSwift.fgTI_SAA.Col = 0
    frmSwift.fgTI_SAA.Row = I
    frmSwift.fgTI_SAA.CellForeColor = vbGreen 'warnUsrColor
    xFileName = paramSwiftTiSaa_TI_Out & "\" & frmSwift.fgTI_SAA.Text
        msFileSystem.MoveFile xFileName, paramSwiftTiSaa_TI_Archive & paramSwift_File_AMJHMS & "_" & frmSwift.fgTI_SAA.Text
Next I
'''Name paramSwiftTiSaa_SAA_In As paramSwiftTiSaa_SAA_In & ".ok"
msFileSystem.MoveFile paramSwiftTiSaa_SAA_Wait & paramSwift_File_AMJHMS, paramSwiftTiSaa_SAA_In & paramSwift_File_AMJHMS
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, "OK : " & paramSwiftTiSaa_SAA_In & paramSwift_File_AMJHMS): DoEvents

Exit Sub
Error_Handle:
 Close
 MsgBox Error, vbCritical, "srvSwift.TI_SAA_Put"
''Kill paramSwiftTiSaa_SAA_Wait & paramSwift_File_AMJHMS
End Sub

Public Sub SAA_TI_Put()

Dim xFileName As String, I As Integer, x As String

On Error GoTo Error_Handle

x = Dir(paramSwiftSaaTi_TI_File)
If x <> "" Then
    Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "srvSwift.SAA_TI_Put" & paramSwiftSaaTi_TI_File & " existe déjà")
    'Call MsgBox(paramSwiftSaaTi_TI_File & " existe déjà", vbCritical, "srvSwift.SAA_TI_Put")
    Exit Sub
End If
paramSwift_File_AMJHMS = "\SAA_Snd_" & DSys & "_" & time_Hms
Open paramSwiftSaaTi_TI_Wait & paramSwift_File_AMJHMS For Output As #2
xIn = "": xOut = ""
Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "SAA => TI : début ...")

'For I = 1 To frmSwift.fgSAA_TI.Rows - 1
I = 1
    frmSwift.fgSAA_TI.Col = 0
    frmSwift.fgSAA_TI.Row = I
    frmSwift.fgSAA_TI.CellForeColor = warnUsrColor
    xFileName = paramSwiftSaaTi_SAA_Out & "\" & frmSwift.fgSAA_TI.Text
    Call lstErr_ChangeLastItem(frmSwift.lstErr, frmSwift.cmdContext, "TI > Swift : " & xFileName): DoEvents
    Open xFileName For Input As #1
    Do Until EOF(1)
        DoEvents
        Line Input #1, xIn
        xOut = Trim(xIn)
        If xOut <> "" Then Print #2, xOut
    Loop
    
    Close #1
''Next I


Close


'''For I = 1 To frmSwift.fgSAA_TI.Rows - 1
    frmSwift.fgSAA_TI.Col = 0
    frmSwift.fgSAA_TI.Row = I
    frmSwift.fgSAA_TI.CellForeColor = vbGreen 'warnUsrColor
    xFileName = paramSwiftSaaTi_SAA_Out & "\" & frmSwift.fgSAA_TI.Text
        Kill xFileName
'''next i
'''Name paramSwiftTiSaa_SAA_In As paramSwiftTiSaa_SAA_In & ".ok"
msFileSystem.MoveFile paramSwiftSaaTi_TI_Wait & paramSwift_File_AMJHMS, paramSwiftSaaTi_TI_File
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, "OK : " & paramSwiftSaaTi_TI_File): DoEvents

Exit Sub
Error_Handle:
 Close
 MsgBox Error, vbCritical, "srvSwift.SAA_TI_Put"
''Kill paramSwiftTiSaa_TI_Wait & paramSwift_File_AMJHMS
End Sub

'
Public Sub Nostro_Put()

Dim xFileName As String, I As Integer

On Error GoTo Error_Handle
paramSwift_File_AMJHMS = "\Nostro_Snd_" & DSys & "_" & time_Hms
Open paramSwiftNostro_Corona_Wait & paramSwift_File_AMJHMS For Output As #2
xIn = "": xOut = ""
Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "Nostro => Corona : début ...")

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
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, "OK : " & paramSwiftNostro_Corona_In & paramSwift_File_AMJHMS): DoEvents

Exit Sub
Error_Handle:
Close
MsgBox Error, vbCritical, "srvSwift.Nostro_Put"
Kill paramSwiftNostro_Corona_Wait & paramSwift_File_AMJHMS
End Sub

Public Sub Loro_Put()

Dim xFileName As String, I As Integer

On Error GoTo Error_Handle
paramSwift_File_AMJHMS = "\Loro_Snd_" & DSys & "_" & time_Hms
Open paramSwiftLoro_SAA_Wait & paramSwift_File_AMJHMS For Output As #2
xIn = "": xOut = ""
Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "Loro => saa : début ...")

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
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, "OK : " & paramSwiftTiSaa_SAA_In & paramSwift_File_AMJHMS): DoEvents

Exit Sub
Error_Handle:
Close
MsgBox Error, vbCritical, "srvSwift.Loro_Put"
Kill paramSwiftLoro_SAA_Wait & paramSwift_File_AMJHMS
End Sub





Public Sub Param_Init_Test()
paramSwiftTiSaa_SAA_Wait = "C:\temp\SAA_TI_Wait"
paramSwiftTiSaa_SAA_In = "C:\temp\SAA_TI_In"
paramSwiftTiSaa_TI_Out = "C:\temp\TI_Posting_Out"
paramSwiftTiSaa_TI_Pattern = "SWO*.txt"
paramSwiftTiSaa_TI_Archive = "C:\temp\TI_Posting_Archive"

paramSwiftSaaTi_SAA_Out = "C:\temp\SAA_TI_Out"
paramSwiftSaaTi_TI_Wait = "C:\temp\TI_Posting_Wait"
paramSwiftSaaTi_TI_File = "C:\temp\TI_Posting_In\SWIFTIN.dta"

paramSwiftLoro_SAA_In = "C:\temp\SAA_MT950_In"
paramSwiftLoro_SAA_Wait = "C:\temp\SAA_MT950_Wait"
paramSwiftLoro_MT950_File = "C:\temp\FTP\Loro.txt"

paramSwiftNostro_Corona_In = "C:\temp\Corona_In"
paramSwiftNostro_Corona_Wait = "C:\temp\Corona_Wait"
paramSwiftNostro_MT950_File = "C:\temp\FTP\Nostro.txt"

paramSwiftSaaCorona_Corona_In = "C:\temp\Corona_In"
paramSwiftSaaCorona_Corona_Wait = "C:\temp\Corona_Wait"
paramSwiftSaaCorona_SAA_Out = "C:\temp\SAA_MT950_Out"

End Sub

Public Sub ImportHisto_Load()
Dim blnUpdate As Boolean
Dim wSwiftHisto As typeSwiftHisto, zSwiftHisto As typeSwiftHisto
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle

Open paramSwiftHisto_Input For Input As #1

MDB.Execute "delete * from SwiftHisto"
mdbSwiftHisto.tableSwiftHisto_Open
recSwiftHisto_Init zSwiftHisto
zSwiftHisto.Method = "AddNew"

blnUpdate = False
Seq = -1

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If Seq < 0 Then
            Seq = ImportHisto_Load_Init(xIn, zSwiftHisto)
        Else
            If mId$(xIn, 2, 3) = "---" Then
                kIn = ImportHisto_Load_Rupture(xIn, wSwiftHisto)
                If kIn = 10 Then
                    If blnUpdate Then dbSwiftHisto_Update wSwiftHisto
                    blnUpdate = True
                    Seq = Seq + 1
                    wSwiftHisto = zSwiftHisto
                    Mid$(wSwiftHisto.Id, 11, 6) = Format$(Seq, "000000")
                Else
                    wSwiftHisto.Text = wSwiftHisto.Text & Chr$(13) & Trim(xIn)
                End If
            Else
                wSwiftHisto.Text = wSwiftHisto.Text & Chr$(13) & Trim(xIn)
                
                Select Case kIn
                    Case 20: Call ImportHisto_Load_MessageHeaderE(xIn, wSwiftHisto)
                    Case 21: Call ImportHisto_Load_MessageHeaderR(xIn, wSwiftHisto)
                    Case 30: Call ImportHisto_Load_MessageText(xIn, wSwiftHisto)
                    Case 50: Call ImportHisto_Load_Interventions(xIn, wSwiftHisto)
                End Select
            End If
       End If
    End If
Loop


If blnUpdate Then dbSwiftHisto_Update wSwiftHisto
Close

mdbSwiftHisto.tableSwiftHisto_Close
Exit Sub

Error_Handle:
 MsgBox "erreur : srvSwift:ImportHisto_Load" & xIn, vbCritical, Error
Close


End Sub

Public Sub ImportBIC_Load(lFile As String)
Dim blnUpdate As Boolean
Dim wdictio As typeDictio
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim xIn As String

Open lFile For Input As #1
mdbDictio.tableDictio_Open
recDictioInit wdictio
wdictio.Method = "AddNew"
wdictio.DicRub = 9000
wdictio.DicAmj = DSys
Seq = 0

Do Until EOF(1)
    Seq = Seq + 1
    If Seq Mod 1000 = 0 Then Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "import BIC : " & Seq)
    DoEvents
    xIn = Input(856, #1)
 '   Debug.Print Xin
    wdictio.DicCode = mId$(xIn, 4, 11)
    wdictio.DicLib = mId$(xIn, 15, 40)
    wdictio.DicTxt = mId$(xIn, 190, 235)
    dbDictioUpdate wdictio
Loop


Close

mdbDictio.tableDictio_Close
Exit Sub

Error_Handle:
 MsgBox "erreur : srvSwift:ImportBIC_Load" & xIn, vbCritical, Error
Close


End Sub

Public Sub ImportMsgFile_Load(lFile As String)
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim x As String
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
Call lstErr_Clear(frmSwift.lstErr, frmSwift.cmdContext, "import : " & lFile)
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, "Lecture : ")

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If mId$(xIn, 1, 13) = "U-UMID      =" Then
            Seq = Seq + 1
            Call lstErr_ChangeLastItem(frmSwift.lstErr, frmSwift.cmdContext, mId$(xIn, 15, 11) & Seq)
        End If
         If mId$(xIn, 1, mPrinter_len) = mPrinter Then
            x = mId$(xIn, 1 + mPrinter_len, 4)
            arrMsgFile_Printer_Index = ImportMsgFile_Load_Printer(x)
            arrMsgFile_Seq_Nb = arrMsgFile_Seq(arrMsgFile_Printer_Index, 0)
           If arrMsgFile_Seq(arrMsgFile_Printer_Index, arrMsgFile_Seq_Nb) < Seq Then
                arrMsgFile_Seq_Nb = arrMsgFile_Seq_Nb + 1
                arrMsgFile_Seq(arrMsgFile_Printer_Index, 0) = arrMsgFile_Seq_Nb
                arrMsgFile_Seq(arrMsgFile_Printer_Index, arrMsgFile_Seq_Nb) = Seq
            End If
            
            I1 = InStr(mPrinter_len, xIn, "on ")
            K1 = Len(xIn) - I1 - 2
            If I1 > 0 Then arrMsgFile_Seq_AMJ(arrMsgFile_Printer_Index, arrMsgFile_Seq_Nb) = mId$(xIn, I1 + 3, K1)
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
Call lstErr_ChangeLastItem(frmSwift.lstErr, frmSwift.cmdContext, "Sélectionner un printer / " & Seq)


Exit Sub

Error_Handle:
 MsgBox "erreur : srvSwift:ImportMsgFile_Load" & xIn, vbCritical, Error
Close


End Sub
Public Sub ImportMsgFile_Print(lFile, lPrinter As String)
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim x As String
Dim mSeq As Integer
Dim mPrinter As String, mPrinter_len As Integer
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer
Dim xRouting As String, blnRouting As Boolean, xRoutingK As String

arrMsgFile_Printer_Index = 0

For I = 1 To arrMsgFile_Printer_Nb
    If arrMsgFile_Printer(I) = lPrinter Then arrMsgFile_Printer_Index = I
Next I

Open lFile For Input As #1
Seq = 0:
Call lstErr_AddItem(frmSwift.lstErr, frmSwift.cmdContext, "Impression : " & lPrinter)
blnOk = False: blnPrint = False
arrMsgFile_Seq_Index = 1

prtSwift_Open lPrinter & " : Swift messages  "
Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If mId$(xIn, 1, 13) = "U-UMID      =" Then
            If blnOk Then prtSwiftMsgFile_Line xRouting, 4
            Seq = Seq + 1
            If Seq = arrMsgFile_Seq(arrMsgFile_Printer_Index, arrMsgFile_Seq_Index) Then
                blnOk = True: blnPrint = False: blnRouting = False
                Call lstErr_ChangeLastItem(frmSwift.lstErr, frmSwift.cmdContext, "impression : " & mId$(xIn, 15, 11))
                xOut = arrMsgFile_Seq_AMJ(arrMsgFile_Printer_Index, arrMsgFile_Seq_Index) & "   " & mId$(xIn, 16, 11) & "   MT" & mId$(xIn, 27, 3) & "  : "
                ''''prtSwiftMsgFile_Line x, 1
                If IsNumeric(mId$(xIn, 27, 3)) Then
                    blnSwift = True
                Else
                    blnSwift = False
                End If
                xRouting = ""
                arrMsgFile_Seq_Index = arrMsgFile_Seq_Index + 1
                ''Exit Do
            Else
                blnOk = False
            End If
        End If
        If blnOk Then
            Select Case mId$(xIn, 1, 13)
                Case "Sender      =": ImportMsgFile_Load_Sender
                Case "Receiver    =": blnPrint = False
                Case "Transaction r": ImportMsgFile_Load_Amount
                Case "Amount      =": prtSwiftMsgFile_Line xIn, 2
                Case "Text         ":
                    If blnSwift Then
                        ImportMsgFile_Load_text
                    Else
                        blnPrint = True: kPrint = 0: Line Input #1, xIn: Line Input #1, xIn
                    End If
                    
                Case "Block 5:": blnPrint = False
                Case "Message Histo": blnPrint = False: blnRouting = True: xRoutingK = "?"
                Case " Sent to APPL": xRouting = xRouting & xRoutingK & ImportMsgFile_Load_Routing(xIn)
            End Select
            If blnRouting Then
                Select Case mId$(xIn, 1, 5)
                    Case "*Orig": xRoutingK = ">>>"
                    Case "*Copy": xRoutingK = " _ "
                End Select
            End If
            If blnPrint Then
                prtSwiftMsgFile_Line xIn, kPrint
            End If
        End If
   End If
Loop

Close

If blnOk Then prtSwiftMsgFile_Line xRouting, 4
prtSwift_Close

Exit Sub

Error_Handle:
 MsgBox "erreur : srvSwift:ImportMsgFile_Load" & xIn, vbCritical, Error
Close


End Sub


Public Function ImportHisto_Load_Init(lIn As String, lSwiftHisto As typeSwiftHisto)
Dim I1 As Integer, I2 As Integer, x As String

ImportHisto_Load_Init = -1
lSwiftHisto.AMJ = "20" & mId(lIn, 7, 2) & mId(lIn, 4, 2) & mId(lIn, 1, 2)
I1 = InStr(1, lIn, "Histo")
If I1 > 0 Then
   ImportHisto_Load_Init = 0
   I2 = InStr(I1, lIn, "-") + 1
    lSwiftHisto.Id = mId$(lIn, I2, 4) & mId$(lIn, I2 + 5, 6) & "000000"
   Select Case mId$(lIn, I1 + 5, 3)
        Case "Emi": lSwiftHisto.RcvSnd = "S"
        Case "Rec": lSwiftHisto.RcvSnd = "R"
        Case Else: lSwiftHisto.RcvSnd = "?"
    End Select
Else
    MsgBox "manque 'Histo'" & Trim(xIn), vbInformation, "ImportHisto_Load_Init"
End If

End Function
Public Function ImportHisto_Load_Rupture(lIn As String, lSwiftHisto As typeSwiftHisto)
ImportHisto_Load_Rupture = -1
Dim I1 As Integer, I2 As Integer

I1 = InStr(1, lIn, "- ") + 2
I2 = InStr(I1, lIn, " -")
If I1 > 0 And I2 > I1 Then
   
   Select Case Trim(mId$(lIn, I1, I2 - I1))
        Case "Instance Type and Transmission": ImportHisto_Load_Rupture = 10
        Case "Message Header":
                        If lSwiftHisto.RcvSnd = "S" Then
                            ImportHisto_Load_Rupture = 20
                        Else
                            ImportHisto_Load_Rupture = 21
                        End If
        Case "Message Text": ImportHisto_Load_Rupture = 30: mField = ""
        Case "Message Trailer": ImportHisto_Load_Rupture = 40
        Case "Interventions": ImportHisto_Load_Rupture = 50: blnCreation = False: blnUnit = False
        Case Else:     MsgBox "nature inconnue" & Trim(xIn), vbInformation, "ImportHisto_Load_Rupture"
    End Select
Else
    MsgBox "manque nature" & Trim(xIn), vbInformation, "ImportHisto_Load_Rupture"
End If

End Function

Public Sub ImportHisto_Load_MessageHeaderE(lIn As String, lSwiftHisto As typeSwiftHisto)
Dim I1 As Integer, I2 As Integer
I1 = InStr(1, lIn, ":")
If I1 > 0 Then
   
   Select Case Trim(mId$(lIn, 1, I1 - 1))
        Case "Swift Input": lSwiftHisto.MT = mId$(lIn, I1 + 6, 3)
        Case "Receiver": I2 = Len(lIn): lSwiftHisto.BIC = Trim(mId$(lIn, I1 + 1, I2 - I1))
 
    End Select
End If

End Sub

Public Sub ImportHisto_Load_MessageHeaderR(lIn As String, lSwiftHisto As typeSwiftHisto)
Dim I1 As Integer, I2 As Integer
I1 = InStr(1, lIn, ":")
If I1 > 0 Then
   
   Select Case Trim(mId$(lIn, 1, I1 - 1))
        Case "Swift Output": lSwiftHisto.MT = mId$(lIn, I1 + 6, 3)
        Case "Sender": I2 = Len(lIn): lSwiftHisto.BIC = Trim(mId$(lIn, I1 + 1, I2 - I1))
 
    End Select
End If

End Sub

Public Sub ImportHisto_Load_Currency(lIn As String, lSwiftHisto As typeSwiftHisto)
Dim I1 As Integer, I2 As Integer
I1 = InStr(1, lIn, ":")
lSwiftHisto.F32C = mId$(lIn, I1 + 2, 3)
End Sub

Public Sub ImportHisto_Load_Amount(lIn As String, lSwiftHisto As typeSwiftHisto)
Dim I1 As Integer, I2 As Integer, x
I1 = InStr(1, lIn, "#") + 1
I2 = InStr(I1, lIn, "#") - 1
x = mId$(lIn, I1, I2)
x = Replace(x, ",", " ")
lSwiftHisto.F32A = CCur(Val(x))
End Sub

Public Sub ImportHisto_Load_MessageText(lIn As String, lSwiftHisto As typeSwiftHisto)
Dim I1 As Integer, I2 As Integer, wField As String
I1 = InStr(1, lIn, ":")
If I1 > 0 Then
   wField = Trim(mId$(lIn, 1, I1 - 1))
   If IsNumeric(mId$(wField, 1, 2)) Then mField = mId$(wField, 1, 2)
   Select Case wField
        Case "20":
                Line Input #1, xIn
                lSwiftHisto.F20 = Trim(xIn)
        Case "21":
                Line Input #1, xIn
                lSwiftHisto.F21 = Trim(xIn)
        Case "Currency": If mField = "32" Then Call ImportHisto_Load_Currency(lIn, lSwiftHisto)
        Case "Amount": If mField = "32" Then Call ImportHisto_Load_Amount(lIn, lSwiftHisto)
        'Case "Date": Call ImportHisto_Load_Date(lIn, lSwiftHisto)
        
    End Select
End If

End Sub

Public Sub ImportHisto_Load_Interventions(lIn As String, lSwiftHisto As typeSwiftHisto)
Dim I1 As Integer, I2 As Integer, wField As String
I1 = InStr(1, lIn, ":")
If I1 > 0 Then
   Select Case Trim(mId$(lIn, 1, I1 - 1))
        Case "Creation Time"
                    If Not blnCreation Then
                         blnCreation = True
                         I1 = I1 + 2
                         lSwiftHisto.AMJ = "20" & mId(lIn, I1 + 6, 2) & mId(lIn, I1 + 3, 2) & mId(lIn, I1, 2)
                         lSwiftHisto.HMS = mId(lIn, I1 + 9, 2) & mId(lIn, I1 + 12, 2) & mId(lIn, I1 + 15, 2)
                    End If
    End Select
    Else
        I1 = InStr(1, lIn, "assigned to unit [") + 18
        If I1 > 18 Then
            I2 = InStr(I1, lIn, "]")
            lSwiftHisto.Unit = mId$(lIn, I1, I2 - I1)
        End If
        

End If

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
    If mId$(xIn, 1, 8) = "Block 5:" Then Exit Do
    
    If mId$(xIn, 1, 1) = ":" Then
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
    If I2 > I1 Then ImportMsgFile_Load_Routing = mId$(lX, I1 + 1, I2 - I1 - 1)
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
