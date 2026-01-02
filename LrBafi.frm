VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLrBafi 
   Caption         =   "Lr_Bafi : Inferface AS400 => Luca Report"
   ClientHeight    =   6345
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9330
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
      Max             =   15000
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5160
      Top             =   0
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   0
      Width           =   3585
   End
   Begin VB.CommandButton cmdActualiser 
      Caption         =   "Act&ualiser l'affichage"
      Height          =   300
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1845
   End
   Begin VB.Frame fraLrBia 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   9300
      Begin TabDlg.SSTab SSTab1 
         Height          =   5700
         Left            =   45
         TabIndex        =   4
         Top             =   195
         Width           =   9210
         _ExtentX        =   16245
         _ExtentY        =   10054
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Options"
         TabPicture(0)   =   "LrBafi.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraTransfert"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraLrBafi"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Interface AS400"
         TabPicture(1)   =   "LrBafi.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraStatut"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Interface Sopra"
         TabPicture(2)   =   "LrBafi.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "filDoc"
         Tab(2).Control(1)=   "fraLrEngineEnd"
         Tab(2).Control(2)=   "fraLrEngineStart"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Archivage / Transmission"
         TabPicture(3)   =   "LrBafi.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraFolder"
         Tab(3).Control(1)=   "fraArchive"
         Tab(3).ControlCount=   2
         Begin VB.Frame fraFolder 
            Height          =   2655
            Left            =   -74880
            TabIndex        =   47
            Top             =   2880
            Width           =   8895
            Begin VB.OptionButton optFolderMoveFile 
               Caption         =   "Afficher le dossier R:\....\Bia_Transmettre"
               Height          =   615
               Left            =   5280
               TabIndex        =   50
               Top             =   1560
               Value           =   -1  'True
               Width           =   3375
            End
            Begin VB.OptionButton optFolderCopyFile 
               Caption         =   "Afficher le dossier R:\....\Bia_Archive"
               Height          =   615
               Left            =   5280
               TabIndex        =   49
               Top             =   720
               Width           =   3375
            End
            Begin MSFlexGridLib.MSFlexGrid fgFolder 
               Height          =   2250
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   5025
               _ExtentX        =   8864
               _ExtentY        =   3969
               _Version        =   393216
               Rows            =   1
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   14737632
               ForeColor       =   12582912
               BackColorFixed  =   12632256
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   14737632
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   0
               GridLinesFixed  =   1
               FormatString    =   "<Nom du fichier                                           |<Date dernière modif         "
            End
         End
         Begin VB.Frame fraArchive 
            Height          =   2055
            Left            =   -74880
            TabIndex        =   42
            Top             =   600
            Width           =   8895
            Begin VB.CheckBox chkArchiveCopyFile 
               Caption         =   "Copie vers R:\....\Bia_Archive"
               Height          =   495
               Left            =   5280
               TabIndex        =   44
               Top             =   720
               Value           =   1  'Checked
               Width           =   3495
            End
            Begin VB.CheckBox chkArchiveMoveFile 
               Caption         =   "Déplacer vers R:\....\Bia_Transmettre"
               Height          =   495
               Left            =   5280
               TabIndex        =   43
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3495
            End
            Begin MSFlexGridLib.MSFlexGrid fgArchive 
               Height          =   1650
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   5025
               _ExtentX        =   8864
               _ExtentY        =   2910
               _Version        =   393216
               Rows            =   1
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   14737632
               ForeColor       =   12582912
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   14737632
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   0
               GridLinesFixed  =   1
               FormatString    =   "<Nom du fichier à archiver / transmettre      |<Date dernière modif       "
            End
            Begin VB.Label lblArchive 
               Caption         =   "<= Cliquer sur une ligne pour archiver / transmettre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5160
               TabIndex        =   46
               Top             =   360
               Width           =   3615
            End
         End
         Begin VB.FileListBox filDoc 
            Height          =   2040
            Left            =   -70560
            TabIndex        =   41
            Top             =   2160
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Frame fraLrEngineEnd 
            Height          =   4815
            Left            =   -68760
            TabIndex        =   38
            Top             =   720
            Width           =   2775
            Begin VB.CommandButton cmdLrEngineEnd 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Après Fabrication Sopra"
               Height          =   720
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   3960
               Width           =   2000
            End
            Begin MSFlexGridLib.MSFlexGrid fgLrEngineEnd 
               Height          =   3570
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   6297
               _Version        =   393216
               Rows            =   1
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   14737632
               ForeColor       =   12582912
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   14737632
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   0
               GridLinesFixed  =   1
               FormatString    =   "<Session |<Descri                          "
            End
         End
         Begin VB.Frame fraLrEngineStart 
            Height          =   4935
            Left            =   -74880
            TabIndex        =   35
            Top             =   720
            Width           =   5900
            Begin VB.CommandButton cmdLrEngineStart 
               BackColor       =   &H00C0FFC0&
               Caption         =   " Avant Fabrication Sopra"
               Height          =   720
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   4080
               Width           =   2000
            End
            Begin MSFlexGridLib.MSFlexGrid fgLrEngineStart 
               Height          =   3570
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   5655
               _ExtentX        =   9975
               _ExtentY        =   6297
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   14737632
               ForeColor       =   12582912
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   14737632
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   0
               GridLinesFixed  =   1
               FormatString    =   "<Session |<Pilfab                          |< Estd                           |<Solde                       "
            End
         End
         Begin VB.Frame fraStatut 
            Caption         =   "Statut du traitement"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4500
            Left            =   -74880
            TabIndex        =   17
            Top             =   600
            Width           =   8895
            Begin VB.CommandButton cmdExtractionTransfert 
               BackColor       =   &H00C0FFC0&
               Caption         =   "&Extraction  + Transfert"
               Height          =   960
               Left            =   6480
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   1680
               Width           =   2000
            End
            Begin VB.Label lblStatus00 
               Caption         =   "00 - Extraction demandée"
               Height          =   375
               Left            =   120
               TabIndex        =   32
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label lblStatus01 
               Caption         =   "01 - Extraction en cours"
               Height          =   375
               Left            =   120
               TabIndex        =   31
               Top             =   1080
               Width           =   1935
            End
            Begin VB.Label lblStatus02 
               Caption         =   "02 - Extraction terminée"
               Height          =   375
               Left            =   120
               TabIndex        =   30
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblStatus03 
               Caption         =   "03 - Transfert demandé"
               Height          =   375
               Left            =   120
               TabIndex        =   29
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label lblStatus04 
               Caption         =   "04 - Transfert en cours"
               Height          =   375
               Left            =   120
               TabIndex        =   28
               Top             =   3120
               Width           =   1935
            End
            Begin VB.Label lblStatus05 
               Caption         =   "05 - Transfert terminé"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   3800
               Width           =   2415
            End
            Begin VB.Label libLrBafi04 
               Caption         =   "(Estd : )"
               Height          =   255
               Left            =   2280
               TabIndex        =   26
               Top             =   3120
               Width           =   1215
            End
            Begin VB.Label liblRSolde04 
               Caption         =   "(Solde : )"
               Height          =   255
               Left            =   4440
               TabIndex        =   25
               Top             =   3120
               Width           =   1215
            End
            Begin VB.Label libLrBafi02 
               Caption         =   "(Estd : )"
               Height          =   255
               Left            =   2280
               TabIndex        =   24
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label liblRSolde02 
               Caption         =   "(Solde : )"
               Height          =   255
               Left            =   4320
               TabIndex        =   23
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label libLrBafiH 
               Caption         =   "-"
               Height          =   255
               Left            =   4320
               TabIndex        =   22
               Top             =   405
               Width           =   975
            End
            Begin VB.Label libLrBafiD 
               Caption         =   "-"
               Height          =   255
               Left            =   2280
               TabIndex        =   21
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label libLrTransfertD 
               Caption         =   "-"
               Height          =   255
               Left            =   2280
               TabIndex        =   20
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label libLrTransfertH 
               Caption         =   "-"
               Height          =   255
               Left            =   4500
               TabIndex        =   19
               Top             =   1900
               Width           =   1335
            End
         End
         Begin VB.Frame fraLrBafi 
            Caption         =   "Ecritures standardisées"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   8655
            Begin VB.TextBox txtSociété 
               Height          =   285
               Left            =   7560
               MaxLength       =   2
               TabIndex        =   34
               Text            =   "03"
               Top             =   480
               Width           =   375
            End
            Begin VB.OptionButton optSoldeFindAnnée 
               Alignment       =   1  'Right Justify
               Caption         =   "Solde fin d'année"
               Height          =   375
               Left            =   240
               TabIndex        =   16
               Top             =   840
               Width           =   1695
            End
            Begin VB.CheckBox chkDrac 
               Alignment       =   1  'Right Justify
               Caption         =   "générer DRAC"
               Height          =   255
               Left            =   2760
               TabIndex        =   15
               Top             =   480
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CommandButton cmdLrBafi 
               Caption         =   "Nouvelle extraction"
               Height          =   500
               Left            =   6360
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   1080
               Width           =   2000
            End
            Begin VB.OptionButton optSoldeVeille 
               Alignment       =   1  'Right Justify
               Caption         =   "Solde veille"
               Height          =   375
               Left            =   240
               TabIndex        =   13
               Top             =   1320
               Width           =   1695
            End
            Begin VB.OptionButton optSoldeFinDeMois 
               Alignment       =   1  'Right Justify
               Caption         =   "Solde Fin de mois"
               Height          =   375
               Left            =   240
               TabIndex        =   12
               Top             =   360
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.Label lblSociété 
               Caption         =   "N° Société"
               Height          =   255
               Left            =   6360
               TabIndex        =   33
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.Frame fraTransfert 
            Caption         =   "Transfert des fichiers AS400 => Serveur"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   240
            TabIndex        =   5
            Top             =   3120
            Width           =   8775
            Begin VB.CommandButton cmdTransfert 
               Caption         =   "Demande de transfert"
               Height          =   500
               Left            =   6480
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   1200
               Width           =   2000
            End
            Begin VB.TextBox txtSession 
               Height          =   285
               Left            =   7560
               MaxLength       =   2
               TabIndex        =   6
               Text            =   "01"
               Top             =   600
               Width           =   375
            End
            Begin VB.Label lblTransfertA1 
               Caption         =   "LrEstdP0 , LrSoldeP0 => EstdNoEc.S** , SolCgeEp.S**"
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   1320
               Width           =   4095
            End
            Begin VB.Label Label1 
               Caption         =   "Biafil                   => S:\LucaReport\BiaLr97"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   9
               Top             =   720
               Width           =   3975
            End
            Begin VB.Label lblSession 
               Caption         =   "N° session"
               Height          =   255
               Left            =   6480
               TabIndex        =   8
               Top             =   600
               Width           =   855
            End
         End
      End
   End
   Begin VB.Menu mnuActualiser 
      Caption         =   "mnuActualiser"
      Visible         =   0   'False
      Begin VB.Menu mnuActualiserManuel 
         Caption         =   "Manuel"
      End
      Begin VB.Menu mnuActualiserAuto 
         Caption         =   "Automatique"
      End
      Begin VB.Menu mnuActualiserStop 
         Caption         =   "Stop"
      End
   End
End
Attribute VB_Name = "frmLrBafi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim x As String, X1 As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String

Dim recAccAut As typeAccAut
Dim IdShell
Dim recAs400Cmd As typeAs400Cmd
Dim recLrBafi As typeLrBafi
Dim recLrSolde As typeLrSolde

Dim blnActualiserAuto As Boolean, blnExtraireTransfert As Boolean
Dim strStatusMax As String * 2

Dim LrBafiNb As Integer, LrSoldeNb As Integer

Dim recLrBafi_Run As typeAccAut
Dim recLrBafiMsg As typeLrBafiMsg

Public Sub AccAut_Load()

srvAccAut.Init recLrBafi_Run
recLrBafi_Run.Method = "SeekP0"
recLrBafi_Run.AccAutId = "SRVBIALR"
recLrBafi_Run.AccAutK1 = "AUTO"
recLrBafi_Run.AccAutK2 = "LRBAFI_RUN"
If Not IsNull(srvAccAut.Monitor(recLrBafi_Run)) Then Unload Me

fraLrBia.Enabled = False

If Trim(recLrBafi_Run.AccAutTxt) <> "" Then
    x = MsgBox("Le module LrBafi est en cours d'utilisation par : " & Trim(recLrBafi_Run.AccAutTxt) & Chr$(10) & "Voulez-vous continuer ?", vbYesNo + vbQuestion + vbDefaultButton2, "Autorisation : AccAut ( SRVBIALR / AUTO / LRBAFI_RUN)")
Else
    x = vbYes
End If
If x = vbYes Then
    fraLrBia.Enabled = True
    recLrBafi_Run.AccAutTxt = usrId
    recLrBafi_Run.AccAutDD = DSys
    recLrBafi_Run.AccAutHD = time_Hms
    AccAut_Update
End If
End Sub
Public Sub AccAut_Unload()
If fraLrBia.Enabled Then
    recLrBafi_Run.AccAutTxt = ""
    recLrBafi_Run.AccAutDF = DSys
    recLrBafi_Run.AccAutHF = time_Hms
    AccAut_Update
End If

End Sub

Public Sub AccAut_Update()
recLrBafi_Run.Method = constUpdate
If Not IsNull(srvAccAut.Update(recLrBafi_Run)) Then
    Call lstErr_AddItem(lstErr, cmdActualiser, "AccAut : mise à jour non effectuée")
End If

End Sub




Public Sub Msg_Rcv(x As String)
'---------------------------------------------------------

End Sub


Private Sub chkArchiveCopyFile_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkArchiveCopyFile
End Sub


Private Sub chkArchiveMoveFile_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkArchiveMoveFile
End Sub


Private Sub chkDrac_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkDrac

End Sub


Private Sub cmdActualiser_Click()
Select Case SSTab1.Tab
    Case 2: fgLrEngine_Reset
    Case 3: fgArchive_Reset
    Case Else
            mnuActualiserStop.Enabled = blnActualiserAuto
            PopupMenu mnuActualiser

End Select
End Sub
Public Sub cmdContext_Quit()
If blnActualiserAuto Then
    cmdActualiser_Stop
Else

    If blnMsgBox_Quit Then
       x = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
       x = vbYes
    End If
    If x = vbYes Then Unload Me
End If

End Sub


Public Sub cmdContext_Return()

SendKeys "{TAB}"

End Sub

Private Sub cmdActualiser_Start()
Timer.Enabled = True
prgBar.Value = 0
prgBar.Visible = True
blnActualiserAuto = True
End Sub

Private Sub cmdActualiser_Stop()
Timer.Enabled = False
prgBar.Visible = False
blnActualiserAuto = False

End Sub

Private Sub cmdActualiser_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdActualiser
End Sub

Private Sub cmdExtractionTransfert_Click()
blnExtraireTransfert = True
cmdLrBafi_Click
End Sub

Private Sub cmdExtractionTransfert_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdExtractionTransfert
End Sub


Private Sub cmdLrBafi_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdLrBafi
End Sub

Private Sub cmdLrEngineEnd_Click()
IdShell = Shell(paramErBafi_Engine_End & Format$(txtSession, "00"), 1)

End Sub

Private Sub cmdLrEngineEnd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdLrEngineEnd
End Sub


Private Sub cmdLrEngineStart_Click()
x = paramErBafi_Engine_Start & Format$(txtSession, " 00 ") & paramErBafi_Engine_Folder & " " & paramErBafi_Estd_FileName & " " & paramErBafi_Solde_FileName
IdShell = Shell(x, 1)
End Sub

Private Sub cmdLrBafi_Click()
txtSociété_Control
x = MsgBox("Voulez-vous réellement lancer une nouvelle extraction ?", vbYesNo + vbQuestion + vbDefaultButton2, "Interface AS400 => Luca Report")
If x <> vbYes Then Exit Sub
fraEnabled False
recAccAut.AccAutTxt = "00" & Format$(txtSession, "00")

Mid$(recAccAut.AccAutTxt, 5, 1) = "M"
If optSoldeFindAnnée = "1" Then Mid$(recAccAut.AccAutTxt, 5, 1) = "A"
If optSoldeVeille = "1" Then Mid$(recAccAut.AccAutTxt, 5, 1) = "V"

Mid$(recAccAut.AccAutTxt, 6, 1) = IIf(chkDrac = "1", "D", " ")
Mid$(recAccAut.AccAutTxt, 7, 2) = Format$(txtSociété, "00")

recAccAut.AccAutDD = "00000000"
recAccAut.AccAutHD = "000000"
recAccAut.AccAutDF = "00000000"
recAccAut.AccAutHF = "000000"

cmdUpdate

srvAs400Cmd.Init recAs400Cmd
recAs400Cmd.Method = "SBMJOB"
x = "SBMJOB CMD(CALL PGM(" & paramErBafi_AS400Ext & ") PARM('" & mId$(recAccAut.AccAutTxt, 5, 1) & "'))"
recAs400Cmd.Text = x & " JOB(" & paramErBafi_AS400Ext & ") USER(" & Trim(usrId) & ") JOBQ(QINTER)"
srvAs400Cmd.Update recAs400Cmd
strStatusMax = "02"
prgBar.Max = 15000
cmdActualiser_Start

End Sub

Private Sub cmdLrEngineStart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdLrEngineStart
End Sub

Private Sub cmdTransfert_Click()
Dim xFileName As String, Iter As Integer, x As String



'PGM
                                                                                
'/*********************************************************************/
'/*                                                                   */
'/* LUCA REPORT : TRANFERT DES FICHIERS 'LRBAFIP0' ET 'LRSOLDEP0'     */
'/*               SUR LE SERVEUR                                      */
'/*                                                                   */
'/*********************************************************************/
'
'             OVRDBF     FILE(INPUT) TOFILE(BIASRC/QFTPSRC) MBR(LRBAFI)
'
'             FTP RMTSYS(SERVEUR2)
'
'             DLTOVR     FILE(*ALL)
'
'             ENDPGM
'/*********************************************************************/
'/* SOURCE DE 'BIASRC / QFTPSRC / LRBAFI.TXT                          */
'/*********************************************************************/
'anonymous X
'PUT BIAFIL/LRBAFIP0   /DFTP/LRBAFI
'PUT BIAFIL/LRSOLDEP0  /DFTP/LRSOLDE
'PUT BIAFIL/LRBAFIMSG  /DFTP/LRBAFIMSG
'quit

On Error GoTo cmdTransfert_Error
lstErr.Clear
txtSession_Control
If lstErr.ListCount > 0 Then Call lstErr_AddItem(lstErr, cmdTransfert, "Corrigez les anomalies"): Exit Sub

SSTab1.Tab = 1
Timer.Enabled = False
fraEnabled False
cmdActualiser_Display

Mid$(recAccAut.AccAutTxt, 1, 4) = "04" & Format$(txtSession, "00")

xFileName = paramErBafi_FTP_LrBafi   ''''constFTP_Dir & "LrBafi"
Call lstErr_AddItem(lstErr, cmdTransfert, "Suppression " & xFileName)
x = Dir(xFileName)
If x <> "" Then Kill xFileName


xFileName = paramErBafi_FTP_LrSolde    ''''constFTP_Dir & "LrSolde"
Call lstErr_AddItem(lstErr, cmdTransfert, "Suppression " & xFileName)
x = Dir(xFileName)
If x <> "" Then Kill xFileName

xFileName = paramErBafi_FTP_LrBafiMsg   ''''constFTP_Dir & "LrBafiMsg"
Call lstErr_AddItem(lstErr, cmdTransfert, "Suppression " & xFileName)
x = Dir(xFileName)
If x <> "" Then Kill xFileName

cmdUpdate

prgBar.Visible = True
prgBar.Value = 0
libLrBafi04.Visible = True
libLrBafi04.ForeColor = warnUsrColor

Call lstErr_AddItem(lstErr, cmdTransfert, "Transfert AS400 => " & paramErBafi_FTP_LrBafi)
srvAs400Cmd.Init recAs400Cmd
recAs400Cmd.Method = "SBMJOB"
x = "SBMJOB CMD(CALL PGM(" & paramErBafi_AS400Trf & ")) "
recAs400Cmd.Text = x & " JOB(" & paramErBafi_AS400Trf & ") USER(" & Trim(usrId) & ") JOBQ(QINTER)"
srvAs400Cmd.Update recAs400Cmd

Iter = 0
If LrBafiNb > 0 Then prgBar.Max = 31000
Do
    DoEvents
    x = Dir(xFileName)
    Iter = Iter + 1
    prgBar.Value = Iter
    If Iter > 90000 Then
        x = MsgBox("Voulez-vous réessayer ?", vbQuestion, "LrBafi : cmdTransfert : FTP en cours ")
        If x = vbYes Then
            Iter = 0
        Else
            Err = 9999: GoTo cmdTransfert_Error
        End If
    End If
Loop While x = ""
'''''GoTo solde

xFileName = paramErBafi_FTP_LrBafiMsg   ''constFTP_Dir & "LrBafimsg"
Open xFileName For Input As #2

arrLrBafiMsgNb = 0
xFileName = paramErBafi_Msg_FileName & txtSession
Call lstErr_AddItem(lstErr, cmdTransfert, "Copie vers " & xFileName)
Open xFileName For Output As #1

Do Until EOF(2)
    Line Input #2, x
    arrLrBafiMsgNb = arrLrBafiMsgNb + 1
    Print #1, x
Loop
Close #1
Close #2
prtLrBafiMsgX xFileName


xFileName = paramErBafi_FTP_LrBafi   '''constFTP_Dir & "LrBafi"
Open xFileName For Input As #2

arrLrBafiNb = 0
xFileName = paramErBafi_Estd_FileName & txtSession
Call lstErr_AddItem(lstErr, cmdTransfert, "Copie vers " & xFileName)

Open xFileName For Output As #1
If LrBafiNb > 0 Then prgBar.Max = LrBafiNb

Do Until EOF(2)
    DoEvents
    Line Input #2, x
    arrLrBafiNb = arrLrBafiNb + 1
    Print #1, x
    libLrBafi04 = "( Estd :" & Format$(arrLrBafiNb, "####0") & " ) "
    If LrBafiNb > arrLrBafiNb Then prgBar.Value = arrLrBafiNb
Loop
Close #2
Close #1
libLrBafi04.ForeColor = libUsr.ForeColor

solde:
If LrSoldeNb > 0 Then prgBar.Max = LrSoldeNb
prgBar.Value = 0
liblRSolde04.Visible = True
liblRSolde04.ForeColor = warnUsrColor

xFileName = paramErBafi_FTP_LrSolde     '''constFTP_Dir & "LrSolde"
Open xFileName For Input As #2


arrLrSoldeNb = 0
xFileName = paramErBafi_Solde_FileName & txtSession
Open xFileName For Output As #1
Do Until EOF(2)
    DoEvents
    Line Input #2, x
    arrLrSoldeNb = arrLrSoldeNb + 1
    Print #1, x
    liblRSolde04 = "( solde :" & Format$(arrLrSoldeNb, "####0") & " ) "
    If LrSoldeNb > arrLrSoldeNb Then prgBar.Value = arrLrSoldeNb
Loop
Close #2
Close #1
prgBar.Visible = False
liblRSolde04.ForeColor = libUsr.ForeColor

strStatusMax = "05"
prgBar.Max = 5000
cmdActualiser_Start
Mid$(recAccAut.AccAutTxt, 1, 4) = "05" & Format$(txtSession, "00")
cmdUpdate
fraEnabled True
If LrBafiNb <> arrLrBafiNb Then
    x = MsgBox("Estd générées :" & Format$(LrBafiNb, "####0") & Chr$(10) & "Estd transférées :" & Format$(arrLrBafiNb, "####0"), vbCritical, "Interface AS400 => Luca Report : transfert")
End If
If LrSoldeNb <> arrLrSoldeNb Then
    x = MsgBox("Soldes générés :" & Format$(LrSoldeNb, "####0") & Chr$(10) & "Soldes transférés :" & Format$(arrLrSoldeNb, "####0"), vbCritical, "Interface AS400 => Luca Report : transfert")
End If

lstErr.Clear: lstErr.Height = 200

Exit Sub

'---------------------------------------------------------
cmdTransfert_Error:
'---------------------------------------------------------

MsgBox "erreur : " & Err & " " & Error$(Err), vbCritical, "LrBafi : cmdTransfert : " & xFileName
Resume cmdTransfert_End

cmdTransfert_End:

End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub cmdTransfert_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdTransfert
End Sub

Private Sub fgArchive_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set fgArchive
End Sub

Private Sub Form_Activate()
cmdActualiser_Click
If Not fraLrBia.Enabled Then Unload Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  cmdContext_Return
    Case Is = 27:  cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

srvLrBafi.param_Init

lblArchive.ForeColor = warnUsrColor
blnExtraireTransfert = False
libLrBafi04.Visible = False: liblRSolde04.Visible = False
cmdActualiser_Display
cmdActualiser_Stop
AccAut_Load
End Sub



Public Sub cmdUpdate()
recAccAut.Method = constUpdate
If Not IsNull(srvAccAut.Update(recAccAut)) Then
    Call lstErr_AddItem(lstErr, cmdActualiser, "Mise à jour non effectuée")
End If
cmdActualiser_Display

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_Unload(Cancel As Integer)
AccAut_Unload
End Sub

Private Sub fgArchive_Click()
Dim xSrc As String, xDest As String, I As Integer
On Error GoTo Error_Msg
lstErr.Clear
I = fgArchive.Row * fgArchive.Cols
If fgArchive.Row > 0 Then
    xSrc = Trim(fgArchive.TextArray(0 + I))
   xDest = DSys & "_" & time_Hms & "_" & xSrc
    If chkArchiveCopyFile Then msFileSystem.CopyFile paramErBafi_Out_Folder & xSrc, paramErBafi_Archive & xDest
    If chkArchiveMoveFile Then
        msFileSystem.MoveFile paramErBafi_Out_Folder & xSrc, paramErBafi_Emission & xDest
        Set msFile = msFileSystem.GetFile(paramErBafi_Archive & xDest)
        If msFile.Attributes And 1 Then msFile.Attributes = msFile.Attributes - 1
        msFile.Attributes = msFile.Attributes + 1
    End If
End If
fgArchive_Reset
Exit Sub
'---------------------------------------------------------
Error_Msg:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "frmLrBafi : fraArchive : " & xSrc & " / " & xDest)

fgArchive_Reset

End Sub

Private Sub fraArchive_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraLrBafi_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraLrBia_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraLrEngineEnd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraLrEngineStart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraStatut_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraTransfert_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnuActualiserAuto_Click()
strStatusMax = "99"
cmdActualiser_Display
cmdActualiser_Start
End Sub

Private Sub mnuActualiserManuel_Click()
cmdActualiser_Display
End Sub


Private Sub mnuActualiserStop_Click()
cmdActualiser_Stop
End Sub

Private Sub optFolderCopyFile_Click()
fgArchive_Reset
End Sub

Private Sub optFolderCopyFile_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optFolderCopyFile
End Sub


Private Sub optFolderMoveFile_Click()
fgArchive_Reset
End Sub

Private Sub optFolderMoveFile_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optFolderMoveFile
End Sub


Private Sub optSoldeFindAnnée_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optSoldeFindAnnée
End Sub


Private Sub optSoldeFinDeMois_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optSoldeFinDeMois
End Sub


Private Sub optSoldeVeille_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optSoldeVeille
End Sub


Private Sub prgBar_Click()
cmdActualiser_Stop
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
    Case 2: fgLrEngine_Reset
    Case 3: fgArchive_Reset
End Select
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1
End Sub


Private Sub Timer_Timer()
If prgBar.Value + Timer.Interval >= prgBar.Max Then
    prgBar.Value = 0
    cmdActualiser_Display
    If mId$(recAccAut.AccAutTxt, 1, 2) = strStatusMax Then
        fraEnabled True
        cmdActualiser_Stop
        If blnExtraireTransfert Then cmdTransfert_Click: blnExtraireTransfert = False
    End If
Else
    prgBar.Value = prgBar.Value + Timer.Interval
End If

End Sub

Private Sub txtSession_GotFocus()
txt_GotFocus txtSession
txtSession_Control
End Sub


Private Sub txtSession_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtSession_LostFocus()
txt_LostFocus txtSession

End Sub


Public Sub txtSociété_Control()
Dim x As String

I = Val(txtSociété)
If I < 1 Then Call lstErr_AddItem(lstErr, txtSociété, "Préciser n° Société"): Exit Sub
If I > 99 Then Call lstErr_AddItem(lstErr, txtSociété, "N° Société < 99"): Exit Sub
    
x = vbYes

If optSoldeFindAnnée = "1" And I <> 2 Then
    x = MsgBox("Le N° société devrait être égal à '02' pour une extraction de fin d'année" & Chr$(13) & "Confirmez-vous ce numéro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Contrôle N° société ")
End If

If optSoldeFindAnnée <> "1" And I = 2 Then
    x = MsgBox("Le N° société ne devrait pas être égal à '02' pour une extraction de fin d'année" & Chr$(13) & "Confirmez-vous ce numéro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Contrôle N° société ")
End If

If x <> vbYes Then Call lstErr_AddItem(lstErr, txtSociété, " ? N° Société ")

End Sub


Public Sub txtSession_Control()
I = Val(txtSession)
If I < 1 Then Call lstErr_AddItem(lstErr, txtSession, "Préciser n° session"): Exit Sub
If I > 10 Then Call lstErr_AddItem(lstErr, txtSession, "N° session < 10"): Exit Sub
End Sub

Public Sub cmdActualiser_Display()

Dim X2 As String * 2

srvAccAut.Init recAccAut
recAccAut.Method = "SeekP0"
recAccAut.AccAutId = "SRVBIALR"
recAccAut.AccAutK1 = "AUTO"
recAccAut.AccAutK2 = "LRBAFI"
If Not IsNull(srvAccAut.Monitor(recAccAut)) Then Unload Me

libLrBafiD = dateImp(recAccAut.AccAutDD)
libLrBafiH = timeImp(recAccAut.AccAutHD)

libLrTransfertD = dateImp(recAccAut.AccAutDF)
libLrTransfertH = timeImp(recAccAut.AccAutHF)
txtSession = mId$(recAccAut.AccAutTxt, 3, 2)

Select Case mId$(recAccAut.AccAutTxt, 5, 1)
    Case Is = "V": optSoldeVeille = "1"
    Case Is = "A": optSoldeFindAnnée = "1"
    Case Else: optSoldeFinDeMois = "1"
End Select
chkDrac.Value = IIf(mId$(recAccAut.AccAutTxt, 6, 1) = "D", 1, 0)
txtSociété = mId$(recAccAut.AccAutTxt, 7, 2)

X2 = mId$(recAccAut.AccAutTxt, 1, 2)
lblStatus00.ForeColor = IIf(X2 = "00", warnUsrColor, lblUsr.ForeColor)
lblStatus01.ForeColor = IIf(X2 = "01", warnUsrColor, lblUsr.ForeColor)
lblStatus02.ForeColor = IIf(X2 = "02", warnUsrColor, lblUsr.ForeColor)
lblStatus03.ForeColor = IIf(X2 = "03", warnUsrColor, lblUsr.ForeColor)
lblStatus04.ForeColor = IIf(X2 = "04", warnUsrColor, lblUsr.ForeColor)
lblStatus05.ForeColor = IIf(X2 = "05", warnUsrColor, lblUsr.ForeColor)

LrBafiNb = Val(mId$(recAccAut.AccAutTxt, 11, 5))
libLrBafi02 = "( Estd :" & Format$(LrBafiNb, "####0") & " ) "
LrSoldeNb = Val(mId$(recAccAut.AccAutTxt, 16, 5))
liblRSolde02 = "( Solde :" & Format$(LrSoldeNb, "####0") & " ) "

'If X2 = "02" Or X2 = "05" Then
    cmdTransfert.Enabled = True
'Else
'    cmdTransfert.Enabled = False
'End If

If X2 = "05" Then
    cmdLrEngineStart.Enabled = True
    cmdLrEngineEnd.Enabled = True
Else
    cmdLrEngineStart.Enabled = False
    cmdLrEngineEnd.Enabled = False
End If

End Sub

Public Sub fraEnabled(bln As Boolean)
'fraStatut.Enabled = bln
'fraLrBafi.Enabled = bln
'fraTransfert.Enabled = bln
frmLrBafi.Enabled = bln
End Sub

Private Sub txtSociété_GotFocus()
txt_GotFocus txtSociété
txtSociété_Control

End Sub


Private Sub txtSociété_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSociété_LostFocus()
txt_LostFocus txtSociété

End Sub



Public Sub cmdTransfert_Dtaq()

'  ancienne Version remplacée par FTP
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'srvLrBafiMsg.Init recLrBafiMsg
'recLrBafiMsg.Method = "SnapP0"
'arrLrBafiMsgSuite = True
'arrLrBafiMsgNb = 0
'xFileName = paramErBafi_Msg_FileName & txtSession
'Open xFileName For Output As #1
'Do Until Not arrLrBafiMsgSuite
'    srvLrBafiMsg.Monitor recLrBafiMsg
'    blnSnd = Not arrLrBafiMsgSuite
'Loop
'Close #1
'prtLrBafiMsgX xFileName'

'srvLrBafi.Init recLrBafi
'recLrBafi.Method = "SnapP0"
'arrLrBafiSuite = True
'arrLrBafiNb = 0
'xFileName = paramErBafi_Estd_FileName & txtSession


'Open xFileName For Output As #1

'Do Until Not arrLrBafiSuite
'    srvLrBafi.Monitor recLrBafi
'    libLrBafi04 = "( Estd :" & Format$(arrLrBafiNb, "####0") & " ) "
'    If LrBafiNb > arrLrBafiNb Then prgBar.Value = arrLrBafiNb
'    blnSnd = Not arrLrBafiSuite
'Loop
'Close #1
'libLrBafi04.ForeColor = libUsr.ForeColor

'solde:
'If LrSoldeNb > 0 Then prgBar.Max = LrSoldeNb
'prgBar.Value = 0
'liblRSolde04.Visible = True
'liblRSolde04.ForeColor = warnUsrColor

'srvLrSolde.Init recLrSolde
'recLrSolde.Method = "SnapP0"
'arrLrSoldeSuite = True
'arrLrSoldeNb = 0
'xFileName = paramErBafi_Solde_FileName & txtSession
'Open xFileName For Output As #1
'Do Until Not arrLrSoldeSuite
'    srvLrSolde.Monitor recLrSolde
'    liblRSolde04 = "( solde :" & Format$(arrLrSoldeNb, "####0") & " ) "
'    If LrSoldeNb > arrLrSoldeNb Then prgBar.Value = arrLrSoldeNb
'    blnSnd = Not arrLrSoldeSuite
'Loop
'Close #1
'prgBar.Visible = False
'liblRSolde04.ForeColor = libUsr.ForeColor

End Sub

Public Sub fgLrEngine_Reset()
Dim I As Integer, K As Integer, x As String, L As Integer, iSession As Integer
Dim arrSession() As Integer, arrD1(), arrD2(), arrD3()

ReDim arrSession(99), arrD1(99), arrD2(99), arrD3(99)

cmdLrEngineStart.Enabled = False: cmdLrEngineStart.BackColor = lblUsr.BackColor
cmdLrEngineStart.Caption = "Session : " & Trim(txtSession) & Chr$(13) & "Avant fabrication Sopra"
cmdLrEngineEnd.Enabled = False: cmdLrEngineEnd.BackColor = lblUsr.BackColor
cmdLrEngineEnd.Caption = "Session : " & Trim(txtSession) & Chr$(13) & "Après fabrication Sopra"
iSession = CInt(txtSession)

Call filDoc_Pattern(filDoc, paramErBafi_PilFab_FileName)

For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    x = Trim(filDoc.Filename)
    L = InStr(x, ".")
    x = mId$(x, L - 2, 2)
    If x >= "01" And x <= "99" Then
        K = CInt(x)
        arrSession(K) = 1
        arrD1(K) = msFile.DateLastModified
    End If
Next I
Call filDoc_Pattern(filDoc, paramErBafi_Estd_FileName)

For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    x = Trim(filDoc.Filename)
    L = Len(x)
    x = mId$(x, L - 1, 2)
    If x >= "01" And x <= "99" Then
        K = CInt(x)
        arrSession(K) = 1
        arrD2(K) = msFile.DateLastModified
    End If
Next I


Call filDoc_Pattern(filDoc, paramErBafi_Estd_FileName)

For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    x = Trim(filDoc.Filename)
    L = Len(x)
    x = mId$(x, L - 1, 2)
    If x >= "01" And x <= "99" Then
        K = CInt(x)
        arrSession(K) = 1
        arrD3(K) = msFile.DateLastModified
    End If
Next I


fgLrEngineStart.Redraw = False
'fgLrEngineStart.Clear
fgLrEngineStart.Rows = 1
fgLrEngineStart.Enabled = True
For I = 0 To 99
    If arrSession(I) = 1 Then
        fgLrEngineStart.Rows = fgLrEngineStart.Rows + 1
        fgLrEngineStart.Row = fgLrEngineStart.Rows - 1
        K = (fgLrEngineStart.Row) * fgLrEngineStart.Cols
        fgLrEngineStart.TextArray(0 + K) = Format$(I, "00")
        fgLrEngineStart.TextArray(1 + K) = arrD1(I)
        fgLrEngineStart.TextArray(2 + K) = arrD2(I)
        fgLrEngineStart.TextArray(3 + K) = arrD3(I)
    End If
Next I
fgLrEngineStart.Redraw = True

If Not IsEmpty(arrD1(iSession)) And Not IsEmpty(arrD2(iSession)) And Not IsEmpty(arrD3(iSession)) Then cmdLrEngineStart.Enabled = True: cmdLrEngineStart.BackColor = &HC0FFC0

'==================================================================================
ReDim arrSession(99), arrD1(99)
Call filDoc_Pattern(filDoc, paramErBafi_Descri_FileName)

For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    x = Trim(filDoc.Filename)
    L = InStr(x, ".")
    x = mId$(x, L - 2, 2)
    If x >= "01" And x <= "99" Then
        K = CInt(x)
        arrSession(K) = 1
        arrD1(K) = msFile.DateLastModified
    End If
Next I


fgLrEngineEnd.Redraw = False
'fgLrEngineend.Clear
fgLrEngineEnd.Rows = 1
fgLrEngineEnd.Enabled = True
For I = 0 To 99
    If arrSession(I) = 1 Then
        fgLrEngineEnd.Rows = fgLrEngineEnd.Rows + 1
        fgLrEngineEnd.Row = fgLrEngineEnd.Rows - 1
        K = (fgLrEngineEnd.Row) * fgLrEngineEnd.Cols
        fgLrEngineEnd.TextArray(0 + K) = Format$(I, "00")
        fgLrEngineEnd.TextArray(1 + K) = arrD1(I)
    End If
Next I
fgLrEngineEnd.Redraw = True
If Not IsEmpty(arrD1(iSession)) Then cmdLrEngineEnd.Enabled = True: cmdLrEngineEnd.BackColor = &HC0FFC0

End Sub
Public Sub fgArchive_Reset()
Dim I As Integer, K As Integer, x As String, L As Integer, iSession As Integer

chkArchiveCopyFile.Caption = "Copie vers " & vbCrLf & paramErBafi_Archive
optFolderCopyFile.Caption = "Afficher " & vbCrLf & paramErBafi_Archive
chkArchiveMoveFile.Caption = "Déplacer vers " & vbCrLf & paramErBafi_Emission
optFolderMoveFile.Caption = "Afficher " & vbCrLf & paramErBafi_Emission

filDoc.Path = paramErBafi_Out_Folder
filDoc.Pattern = "*.*" ' constLrBafi_Bia_Filename

fgArchive.Redraw = False
fgArchive.Rows = 1
fgArchive.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    fgArchive.Rows = fgArchive.Rows + 1
    fgArchive.Row = fgArchive.Rows - 1
    K = (fgArchive.Row) * fgArchive.Cols
    fgArchive.TextArray(0 + K) = Trim(filDoc.Filename)
    fgArchive.TextArray(1 + K) = msFile.DateLastModified
Next I
fgArchive.Redraw = True


'==================================================================================
If optFolderCopyFile Then
    filDoc.Path = paramErBafi_Archive
Else
    filDoc.Path = paramErBafi_Emission
End If
filDoc.Pattern = "*.*"

fgFolder.Redraw = False
fgFolder.Rows = 1
fgFolder.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    fgFolder.Rows = fgFolder.Rows + 1
    fgFolder.Row = fgFolder.Rows - 1
    K = (fgFolder.Row) * fgFolder.Cols
    fgFolder.TextArray(0 + K) = Trim(filDoc.Filename)
    fgFolder.TextArray(1 + K) = msFile.DateLastModified
Next I
fgFolder.Redraw = True

End Sub

Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Or TypeOf xobj Is MSFlexGrid Then
           xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Or TypeOf C Is MSFlexGrid Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
             C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub


