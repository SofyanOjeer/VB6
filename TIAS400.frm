VERSION 5.00
Begin VB.Form frmTIAS400 
   Caption         =   "Load TI vers AS400"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   Icon            =   "TIAS400.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "TIAS400.frx":0442
   ScaleHeight     =   6750
   ScaleWidth      =   11565
   Begin VB.CommandButton CmdFullPosting 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DB2 'FULLPOSTING'"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSABPty 
      Caption         =   "Load 'PARTYDTLS' pour reprise SAB"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton CmdRepriseSAB 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Reprise SAB"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6000
      TabIndex        =   13
      Top             =   0
      Width           =   4695
   End
   Begin VB.CommandButton cmdCHARGE 
      Caption         =   "Load 'BASECHARGE'  'EVENTCHG'"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0FFC0&
      Caption         =   "load *All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   2160
   End
   Begin VB.CommandButton cmdMasterExport 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DB2 'MASTER'/'POSTING'"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdFTP 
      BackColor       =   &H00C0FFC0&
      Caption         =   "FTP : TICD_GET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   2160
   End
   Begin VB.CommandButton CmdLCAMEND 
      Caption         =   "Load 'LCAMEND'"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton CmdBASEEVENT 
      Caption         =   "Load 'BASEEVENT'"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton CmdSTEPHIST 
      Caption         =   "Load 'STEPHIST'"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton CmdEXEMPL30 
      Caption         =   "Load 'EXEMPL30'"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton CmdRELITEM 
      Caption         =   "Load 'RELITEM'"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdSCPF 
      Caption         =   "Load 'SCPF'"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdGFPF 
      Caption         =   "Load 'GFPF'"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdPARTYDTLS 
      Caption         =   "Load 'PARTYDTLS'"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox lstMsg 
      Height          =   3765
      Left            =   6120
      TabIndex        =   0
      Top             =   960
      Width           =   5295
   End
End
Attribute VB_Name = "frmTIAS400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsADO As New ADODB.Recordset
Dim X As String

Dim DateAMJ As String
Dim HeureHMS As String

' BASECHARGE et EVENTCHARGE = Commssions à l'ouverture des CREDOC EXPORT

Private Sub cmdCHARGE_Click()

lstMsg.AddItem "BASECHARGE et EVENTCHG : début": DoEvents

Set rsADO = Nothing
' rsADO.Open "Select Basecharge.Key97, Basecharge.event_key, Basecharge.Master_key, Basecharge.Chg_for, Basecharge.status, Basecharge.action, Eventchg.chg_sch, Eventchg.tchg_sch from Basecharge, Eventchg, Baseevent, Master where basecharge.master_key = master.key97 and Basecharge.event_key = baseevent.key97 and Basecharge.key97 = Eventchg.key97 and Basecharge.fromchgkey = 0 and Baseevent.Refno_pfix = 'ADE' and Master.Refno_pfix ='CDE' and Master.Status ='LIV'", "DSN=TIT1TS1"
rsADO.Open "Select Basecharge.Key97, Basecharge.event_key, Basecharge.Master_key, Basecharge.Chg_for, Basecharge.status, Basecharge.action, Eventchg.chg_sch, Eventchg.tchg_sch from Basecharge, Eventchg where Basecharge.event_ref = 'ADE' and Basecharge.serial_no =1 and Basecharge.action='Y' and Basecharge.key97 = Eventchg.key97", "DSN=TIT1TS1"

Open paramAS400IN & "CDBCHW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(81)
    Mid$(X, 1, 12) = Format$(rsADO("KEY97"), "000000000000")
    Mid$(X, 13, 12) = Format$(rsADO("event_key"), "000000000000")
    Mid$(X, 25, 3) = "   "
    Mid$(X, 28, 6) = "000000"
    Mid$(X, 34, 12) = Format$(rsADO("Master_key"), "000000000000")
    Mid$(X, 46, 3) = "   "
    Mid$(X, 49, 6) = "000000"
    Mid$(X, 55, 1) = Format$(rsADO("Chg_for"), " ")
    Mid$(X, 56, 1) = Format$(rsADO("status"), " ")
    Mid$(X, 57, 1) = Format$(rsADO("action"), " ")
    Mid$(X, 58, 12) = Format$(rsADO("chg_sch"), "000000000000")
    If IsNull(rsADO("tchg_sch")) Then
        Mid$(X, 70, 12) = "000000000000"
    Else
        Mid$(X, 70, 12) = Format$(rsADO("tchg_sch"), "000000000000")
    End If
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "BASECHARGE et EVENTCHARGE : Terminé"

End Sub


Private Sub cmdAll_Click()

CmdBASEEVENT_Click
CmdEXEMPL30_Click
cmdGFPF_Click
CmdLCAMEND_Click
cmdPARTYDTLS_Click
CmdRELITEM_Click
cmdSCPF_Click
CmdSTEPHIST_Click
cmdCHARGE_Click

cmdMasterExport_Click

End Sub

Private Sub cmdContext_Click()
Unload Me
End Sub

Private Sub CmdFullPosting_Click()

lstMsg.AddItem "FullPosting pour reprise SAB :début": DoEvents

srvTIAS400.TIDB2_FullPosting lstMsg

lstMsg.AddItem "FullPosting pour reprise SAB : Terminé"

End Sub

Private Sub cmdMasterExport_Click()
lstMsg.AddItem "Master, Posting, PayDiff, CommEncais, RegTrans :début": DoEvents

srvTIAS400.TIDB2_Master lstMsg
srvTIAS400.TIDB2_Posting lstMsg
srvTIAS400.TIDB2_PayDiff lstMsg
srvTIAS400.TIDB2_ComEnc lstMsg
srvTIAS400.TIDB2_RegTrans lstMsg

lstMsg.AddItem "Master, Posting, PayDiff, CommEncais, RegTrans : Terminés"

End Sub

' BASEEVENT = Informations gloales des évènements

Private Sub CmdBASEEVENT_Click()

lstMsg.AddItem "BASEEVENT : début": DoEvents

Set rsADO = Nothing
rsADO.Open "select KEY97, REFNO_PFIX, REFNO_SERL, MASTER_KEY, EXEMPLAR, EV_INDEX, STATUS, STATUS_EV, STEP, THEIR_REF, START_DATE from BASEEVENT order by MASTER_KEY", "DSN=TIT1TS1"

Open paramAS400IN & "CDEVTW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(157)
    Mid$(X, 1, 12) = Format$(rsADO("KEY97"), "000000000000")
    Mid$(X, 13, 3) = Format$(rsADO("REFNO_PFIX"), "   ")
    Mid$(X, 16, 6) = Format$(rsADO("REFNO_SERL"), "000000")
    Mid$(X, 22, 12) = Format$(rsADO("MASTER_KEY"), "000000000000")
    Mid$(X, 34, 3) = "   "
    Mid$(X, 37, 6) = "000000"
    Mid$(X, 43, 12) = Format$(rsADO("EXEMPLAR"), "000000000000")
    If IsNull(rsADO("EV_INDEX")) Then
        Mid$(X, 55, 3) = "000"
    Else
        Mid$(X, 55, 3) = Format$(rsADO("EV_INDEX"), "000")
    End If
    Mid$(X, 58, 1) = rsADO("STATUS")
    Mid$(X, 59, 20) = rsADO("STATUS_EV")
    If rsADO("STEP") = "- " Then
        Mid$(X, 79, 2) = "  "
    Else
        Mid$(X, 79, 2) = rsADO("STEP")
    End If
    If IsNull(rsADO("THEIR_REF")) Then
        Mid$(X, 81, 20) = "                    "
    Else
        Mid$(X, 81, 20) = rsADO("THEIR_REF")
    End If

    '  Mid$(X, 81, 20) = rsADO("THEIR_REF")
    
    Mid$(X, 101, 1) = " "
    dateJma08_Amj08 rsADO("START_DATE"), DateAMJ
    Mid$(X, 102, 8) = DateAMJ
    Mid$(X, 110, 8) = "00000000"
    Mid$(X, 118, 20) = "                    "
    Mid$(X, 138, 20) = "                    "
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "BASEEVENT : Terminé"

End Sub

' EXEMPL30 = Libellés des évènements

Private Sub CmdEXEMPL30_Click()

lstMsg.AddItem "EXEMPL30 : début": DoEvents

Set rsADO = Nothing
rsADO.Open "select KEY97, SHORTN13, LONGNA85 from EXEMPL30", "DSN=TIT1TS1"

Open paramAS400IN & "CDELBW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(82)
    Mid$(X, 1, 12) = Format$(rsADO("KEY97"), "000000000000")
    Mid$(X, 13, 10) = rsADO("SHORTN13")
    Mid$(X, 23, 60) = rsADO("LONGNA85")
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "EXEMPL30 : Terminé"

End Sub

' GFPF = Customers Trade Innovation

Private Sub cmdGFPF_Click()

lstMsg.AddItem "GFPF : début": DoEvents

Set rsADO = Nothing
rsADO.Open "select GFCUS1, GFCPNC, GFCNAR, GFCNAL from GFPF", "DSN=TIT1TS1"

Open paramAS400IN & "CDCUSW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(30)
    Mid$(X, 1, 20) = rsADO("GFCUS1")
    Mid$(X, 21, 6) = rsADO("GFCPNC")
    Mid$(X, 27, 2) = rsADO("GFCNAR")
    Mid$(X, 29, 2) = rsADO("GFCNAL")
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "GFPF : Terminé"

End Sub

' LCAMEND = LC Amend / Récupérer seulement le code de réactivation

Private Sub CmdLCAMEND_Click()

lstMsg.AddItem "LCAMEND : début": DoEvents

Set rsADO = Nothing
rsADO.Open "select KEY97, REINSTATE from LCAMEND where REINSTATE = 'Y'", "DSN=TIT1TS1"

Open paramAS400IN & "CDLCAW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(13)
    Mid$(X, 1, 12) = Format$(rsADO("KEY97"), "000000000000")
    Mid$(X, 13, 1) = rsADO("REINSTATE")
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "LCAMEND : Terminé"

End Sub

' PARTYDTLS = Différentes intervenants d un DOSSIER

Private Sub cmdPARTYDTLS_Click()

lstMsg.AddItem "PARTYDTLS : début": DoEvents

Set rsADO = Nothing
'  rsADO.Open "select KEY97, ADDRESS1,CUS_MNM from PARTYDTLS where CUS_MNM <> '     ' and CUS_MNM <> '- '", "DSN=TIT1TS1"
rsADO.Open "select KEY97, ADDRESS1, CUS_MNM from PARTYDTLS", "DSN=TIT1TS1"

Open paramAS400IN & "CDPTYW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(67)
    Mid$(X, 1, 12) = Format$(rsADO("KEY97"), "000000000000")
    If IsNull(rsADO("ADDRESS1")) Then
        Mid$(X, 13, 35) = "                                   "
    Else
        Mid$(X, 13, 35) = rsADO("ADDRESS1")
    End If
    If IsNull(rsADO("CUS_MNM")) Then
        Mid$(X, 48, 20) = "                    "
    Else
        Mid$(X, 48, 20) = rsADO("CUS_MNM")
    End If
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "PARTYDTLS : Terminé"

End Sub

Private Sub cmdFTP_Click()
lstMsg.AddItem "SBMJOB 'TICD_GET'": DoEvents

srvAs400Cmd.SBMJOB "TICD_GET"

End Sub

' RELITEM = Liens Evènement/Postings

Private Sub CmdRELITEM_Click()

lstMsg.AddItem "RELITEM : début": DoEvents

Set rsADO = Nothing
rsADO.Open "select EVENT_KEY, KEY97 from RELITEM", "DSN=TIT1TS1"

Open paramAS400IN & "CDEPSW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(24)
    Mid$(X, 1, 12) = Format$(rsADO("EVENT_KEY"), "000000000000")
    Mid$(X, 13, 12) = Format$(rsADO("KEY97"), "000000000000")
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "RELITEM : Terminé"

End Sub

Private Sub CmdRepriseSAB_Click()

lstMsg.AddItem "CDOUTI, CDOESC, CDOFRS, etc... : début": DoEvents

srvTIAS400.TIDB2_CDOUTI lstMsg
srvTIAS400.TIDB2_CDOESC lstMsg
'  srvTIAS400.TIDB2_CDOFRS lstMsg

lstMsg.AddItem "CDOUTI, CDOESC, CDOFRS, etc... : Terminés"

End Sub

Private Sub CmdSABPty_Click()

' Load PARTYDTLS pour repise TIERS CREDOC SAB

lstMsg.AddItem "PARTYDTLS reprise Tiers SAB : début": DoEvents

Set rsADO = Nothing
'  rsADO.Open "select KEY97, ADDRESS1,CUS_MNM from PARTYDTLS where CUS_MNM <> '     ' and CUS_MNM <> '- '", "DSN=TIT1TS1"
rsADO.Open "select KEY97, CUS_MNM, ADDRESS1, ADDRESS2, ADDRESS3, ADDRESS4, ADDRESS5, COUNTRY from PARTYDTLS", "DSN=TIT1TS1"

Open paramAS400IN & "RCDPTYW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(209)
    Mid$(X, 1, 12) = Format$(rsADO("KEY97"), "000000000000")
    If IsNull(rsADO("CUS_MNM")) Then
        Mid$(X, 13, 20) = "                    "
    Else
        Mid$(X, 13, 20) = rsADO("CUS_MNM")
    End If
    If IsNull(rsADO("ADDRESS1")) Then
        Mid$(X, 33, 35) = "                                   "
    Else
        Mid$(X, 33, 35) = rsADO("ADDRESS1")
    End If
    If IsNull(rsADO("ADDRESS2")) Then
        Mid$(X, 68, 35) = "                                   "
    Else
        Mid$(X, 68, 35) = rsADO("ADDRESS2")
    End If
    If IsNull(rsADO("ADDRESS3")) Then
        Mid$(X, 103, 35) = "                                   "
    Else
        Mid$(X, 103, 35) = rsADO("ADDRESS3")
    End If
    If IsNull(rsADO("ADDRESS4")) Then
        Mid$(X, 138, 35) = "                                   "
    Else
        Mid$(X, 138, 35) = rsADO("ADDRESS4")
    End If
    If IsNull(rsADO("ADDRESS5")) Then
        Mid$(X, 173, 35) = "                                   "
    Else
        Mid$(X, 173, 35) = rsADO("ADDRESS5")
    End If
    If IsNull(rsADO("COUNTRY")) Then
        Mid$(X, 208, 2) = "  "
    Else
        Mid$(X, 208, 2) = rsADO("COUNTRY")
    End If
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "PARTYDTLS reprise Tiers SAB : Terminé"

End Sub

' SCPF = Comptes Trade Innovation

Private Sub cmdSCPF_Click()

lstMsg.AddItem "SCPF : début": DoEvents

Set rsADO = Nothing
rsADO.Open "select SCAB, SCAN, SCAS, SCACT, SCCCY, SCEAN1 from SCPF", "DSN=TIT1TS1"

Open paramAS400IN & "CDACCW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(43)
    Mid$(X, 1, 4) = rsADO("SCAB")
    Mid$(X, 5, 6) = rsADO("SCAN")
    Mid$(X, 11, 3) = rsADO("SCAS")
    Mid$(X, 14, 2) = rsADO("SCACT")
    Mid$(X, 16, 3) = rsADO("SCCCY")
    Mid$(X, 19, 25) = rsADO("SCEAN1")
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "SCPF : Terminé"

End Sub


' STEPHIST = Historique des étapes des évènements

Private Sub CmdSTEPHIST_Click()

lstMsg.AddItem "STEPHIST : début": DoEvents

Set rsADO = Nothing
rsADO.Open "select KEY97, EVENT_KEY, STATUS, TYPE, DATE_START, START_TIME, USERID from STEPHIST order by EVENT_KEY, YEAR(DATE_START), MONTH(DATE_START), DAY(DATE_START), START_TIME, KEY97", "DSN=TIT1TS1"

Open paramAS400IN & "CDEEVW0" For Output As #2

Do While Not rsADO.EOF
    X = Space$(91)
    Mid$(X, 1, 12) = Format$(rsADO("EVENT_KEY"), "000000000000")
    Mid$(X, 13, 3) = "   "
    Mid$(X, 16, 6) = "000000"
    Mid$(X, 22, 12) = "000000000000"
    Mid$(X, 34, 3) = "   "
    Mid$(X, 37, 6) = "000000"
    dateJma08_Amj08 rsADO("DATE_START"), DateAMJ
    Mid$(X, 43, 8) = DateAMJ
    HeureHMS = mId$(rsADO("START_TIME"), 1, 2) & mId$(rsADO("START_TIME"), 4, 2) & "00"
    Mid$(X, 51, 6) = HeureHMS
    Mid$(X, 57, 12) = Format$(rsADO("KEY97"), "000000000000")
    Mid$(X, 69, 2) = rsADO("TYPE")
    Mid$(X, 71, 1) = rsADO("STATUS")
    Mid$(X, 72, 20) = rsADO("USERID")
    Print #2, X
    rsADO.MoveNext
Loop

Close

lstMsg.AddItem "STEPHIST : Terminé"

End Sub

Private Sub Form_Load()

'paramAS400IN = paramServer("\\AS400_IN\"): lstMsg.AddItem paramAS400IN
'paramTIDB2_Master = paramServer("\\TI\TI.DAT\Log\TIExtract\TI_Master.txt"): lstMsg.AddItem paramTIDB2_Master
'paramTIDB2_Posting = paramServer("\\TI\TI.DAT\Log\TIExtract\TI_Posting.txt"): lstMsg.AddItem paramTIDB2_Posting
'paramTIDB2_PayDiff = paramServer("\\TI\TI.DAT\Log\TIExtract\TI_PayDiff.txt"): lstMsg.AddItem paramTIDB2_PayDiff
'paramTIDB2_ComEnc = paramServer("\\TI\TI.DAT\Log\TIExtract\TI_ComEnc.txt"): lstMsg.AddItem paramTIDB2_ComEnc
'p 'aramTIDB2_OK = paramServer("\\TI\TI.DAT\Log\TIExtract\TI_Extract_OK")
'paramTIDB2_TIAS400 = paramServer("\\TI\TI.DAT\Log\TIExtract\TI_Extract_TIAS400")

End Sub


Public Sub Msg_Rcv(Msg As String)

srvTIAS400.param_Init

If UCase$(Trim(mId$(Msg, 1, 12))) = "@AUTO_TIAS40" Then
    Auto_TIAS400
    If blnTimer_Enabled Then
        End
    Else
        Unload Me
    End If
End If

End Sub

Public Sub Auto_TIAS400()
Dim X As String
On Error GoTo Exit_Sub

X = Dir(paramTIDB2_TIAS400)
If X <> "" Then Kill paramTIDB2_TIAS400

X = Dir(paramTIDB2_OK)
If X = "" Then GoTo Exit_Sub
Name paramTIDB2_OK As paramTIDB2_TIAS400

cmdAll_Click
cmdFTP_Click

Kill paramTIDB2_TIAS400
Exit_Sub:

End Sub
