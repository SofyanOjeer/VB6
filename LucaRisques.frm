VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLucaRisques 
   Caption         =   "Luca Risques"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9180
   Begin VB.Frame fraLrBia 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9135
      Begin VB.TextBox txtDTCENTEnCours 
         Height          =   285
         Left            =   3000
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.Timer Timer 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   8400
         Top             =   480
      End
      Begin VB.ListBox lstErr 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   23
         Top             =   120
         Width           =   4185
      End
      Begin VB.CommandButton cmdActualiser 
         Caption         =   "Act&ualiser l'affichage"
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1845
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4815
         Left            =   0
         TabIndex        =   2
         Top             =   1080
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8493
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Extraction AS400"
         TabPicture(0)   =   "LucaRisques.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraStatut"
         Tab(0).Control(1)=   "fraOption1"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Interface Luca Risques"
         TabPicture(1)   =   "LucaRisques.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraLrTiers"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame1"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame2"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Impression BIA"
         TabPicture(2)   =   "LucaRisques.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraImport"
         Tab(2).Control(1)=   "fraPrintChk"
         Tab(2).ControlCount=   2
         Begin VB.Frame Frame2 
            Caption         =   "mise en page et impression"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   3120
            TabIndex        =   49
            Top             =   480
            Width           =   2700
            Begin VB.CheckBox chkPrintSopraBmRetour 
               Caption         =   "états Retour BDF"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   2500
               Width           =   2295
            End
            Begin VB.CheckBox chkPrintSopraBmAller 
               Caption         =   "détail bande Aller BDF"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   2100
               Width           =   2295
            End
            Begin VB.CheckBox chkPrintSopra490 
               Caption         =   "bénéficiares déclarés"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   1700
               Width           =   2295
            End
            Begin VB.CheckBox chkPrintSopra470 
               Caption         =   "bordereau de remise"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   1300
               Width           =   2295
            End
            Begin VB.CheckBox chkPrintSopra400 
               Caption         =   "fiches signalétiques erronées"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   900
               Width           =   2535
            End
            Begin VB.CheckBox chkPrintSopra220 
               Caption         =   "bénéficaires à déclarer"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   500
               Width           =   2295
            End
            Begin VB.CommandButton cmdPrintSopra 
               Caption         =   "Impression des états Sopra"
               Height          =   615
               Left            =   120
               TabIndex        =   50
               Top             =   3120
               Width           =   2415
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "transfert des fichiers"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   6000
            TabIndex        =   44
            Top             =   480
            Width           =   2700
            Begin VB.CommandButton cmdExport_LrCdrAller 
               Caption         =   "Export Aller BDF"
               Height          =   700
               Left            =   480
               TabIndex        =   48
               Top             =   480
               Width           =   1800
            End
            Begin VB.CommandButton cmdLrEngineEnd 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Archivage"
               Height          =   700
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   3120
               Width           =   1800
            End
            Begin VB.CommandButton cmdImport_LrCdrBDF 
               Caption         =   "Import Retour BDF"
               Height          =   700
               Left            =   480
               TabIndex        =   45
               Top             =   1800
               Width           =   1800
            End
         End
         Begin VB.Frame fraLrTiers 
            Caption         =   "fiches signalétiques"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   240
            TabIndex        =   40
            Top             =   480
            Width           =   2700
            Begin VB.CommandButton cmdLrTiers_Update 
               Caption         =   "Modification fiches "
               Height          =   700
               Left            =   240
               TabIndex        =   43
               Top             =   1800
               Width           =   1800
            End
            Begin VB.CommandButton cmdExport_LrTiersX 
               Caption         =   "Mise à jour (CliCli)"
               Height          =   700
               Left            =   240
               TabIndex        =   42
               Top             =   3000
               Width           =   1800
            End
            Begin VB.CommandButton cmdImport_LrTiersX 
               Caption         =   "Chargement (CliCli)"
               Height          =   700
               Left            =   240
               TabIndex        =   41
               Top             =   600
               Width           =   1800
            End
         End
         Begin VB.Frame fraOption1 
            Caption         =   "Options (bénéficiaires déclarés)"
            Height          =   1095
            Left            =   -74880
            TabIndex        =   36
            Top             =   3480
            Width           =   8655
            Begin VB.CheckBox chkLrTiersInit 
               Caption         =   "Initialisation des fiches signalétiquespour la déclaration en cours (annule et remplace) "
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   720
               Value           =   1  'Checked
               Width           =   8175
            End
            Begin VB.CheckBox chkLrTiersAS400 
               Caption         =   "Etat des fiches signalétiques modifiées depuis la dernière déclaration"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   360
               Value           =   1  'Checked
               Width           =   7695
            End
         End
         Begin VB.Frame fraImport 
            Caption         =   "Import"
            Height          =   1815
            Left            =   -74880
            TabIndex        =   32
            Top             =   480
            Width           =   8775
            Begin VB.OptionButton optEUR 
               Caption         =   "EUR"
               Height          =   255
               Left            =   7680
               TabIndex        =   59
               Top             =   1440
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton optFRF 
               Caption         =   "FRF"
               Height          =   255
               Left            =   6720
               TabIndex        =   58
               Top             =   1440
               Width           =   735
            End
            Begin VB.CheckBox chkCDCPCO_Add 
               Caption         =   "rattachement des comptes collectifs"
               Height          =   495
               Left            =   6840
               TabIndex        =   39
               Top             =   800
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CommandButton cmdImport 
               Caption         =   "Import LR => Bia.mdb"
               Height          =   495
               Left            =   6600
               TabIndex        =   33
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label libImport_FileName 
               Caption         =   "FileName"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   6255
            End
            Begin VB.Label libImport_Dta 
               Caption         =   "Dta"
               Height          =   960
               Left            =   120
               TabIndex        =   34
               Top             =   720
               Width           =   6375
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame fraPrintChk 
            Caption         =   "Sélection"
            Height          =   2415
            Left            =   -74880
            TabIndex        =   24
            Top             =   2280
            Width           =   4575
            Begin VB.CheckBox chkPrintLstBNFPM 
               Alignment       =   1  'Right Justify
               Caption         =   "Liste Bénéficaires PM"
               Height          =   255
               Left            =   2400
               TabIndex        =   61
               Top             =   1200
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox chkExport 
               Alignment       =   1  'Right Justify
               Caption         =   "Export Service Compta"
               Height          =   255
               Left            =   240
               TabIndex        =   60
               Top             =   2040
               Width           =   1935
            End
            Begin VB.CheckBox chkPrintRFBENFAll 
               Alignment       =   1  'Right Justify
               Caption         =   "tous les bénéficiaires"
               Height          =   255
               Left            =   240
               TabIndex        =   57
               Top             =   600
               Width           =   1935
            End
            Begin VB.CommandButton cmdPrint 
               BackColor       =   &H00E0E0E0&
               Height          =   400
               Left            =   3840
               Picture         =   "LucaRisques.frx":0054
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   1800
               Width           =   500
            End
            Begin VB.TextBox txtRFBENF 
               Height          =   285
               Left            =   2760
               TabIndex        =   30
               Top             =   300
               Width           =   1575
            End
            Begin VB.CheckBox chkPrintLstBNF 
               Alignment       =   1  'Right Justify
               Caption         =   "Liste Bénéficaires"
               Height          =   255
               Left            =   240
               TabIndex        =   29
               Top             =   1200
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox chkPrintTOTBDF 
               Alignment       =   1  'Right Justify
               Caption         =   "Ventilation Cot BdF"
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   1500
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox chkPrintCOTBDF 
               Alignment       =   1  'Right Justify
               Caption         =   "Cotation BdF"
               Height          =   255
               Left            =   240
               TabIndex        =   27
               Top             =   1800
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox chkPrintTotal 
               Alignment       =   1  'Right Justify
               Caption         =   "Total général"
               Height          =   255
               Left            =   240
               TabIndex        =   26
               Top             =   900
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox chkPrintRFBENF 
               Alignment       =   1  'Right Justify
               Caption         =   "choix un Bénéficiaire"
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   300
               Width           =   1935
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
            Height          =   2775
            Left            =   -74880
            TabIndex        =   3
            Top             =   480
            Width           =   8655
            Begin VB.CommandButton cmdTransfert 
               Caption         =   "Demande de transfert"
               Height          =   500
               Left            =   6240
               TabIndex        =   6
               Top             =   2160
               Width           =   2000
            End
            Begin VB.CommandButton cmdLrCdr 
               Caption         =   "Nouvelle extraction"
               Height          =   500
               Left            =   6240
               TabIndex        =   5
               Top             =   1320
               Width           =   2000
            End
            Begin VB.CommandButton cmdExtractionTransfert 
               BackColor       =   &H00C0FFC0&
               Caption         =   "&Extraction  + Transfert"
               Height          =   600
               Left            =   6240
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   360
               Width           =   2000
            End
            Begin VB.Label libLrTransfertH 
               Caption         =   "-"
               Height          =   255
               Left            =   4000
               TabIndex        =   20
               Top             =   1100
               Width           =   1215
            End
            Begin VB.Label libLrTransfertD 
               Caption         =   "-"
               Height          =   255
               Left            =   2400
               TabIndex        =   19
               Top             =   1100
               Width           =   1095
            End
            Begin VB.Label libLrCdrD 
               Caption         =   "-"
               Height          =   250
               Left            =   2400
               TabIndex        =   18
               Top             =   300
               Width           =   1335
            End
            Begin VB.Label libLrCdrH 
               Caption         =   "-"
               Height          =   250
               Left            =   4000
               TabIndex        =   17
               Top             =   300
               Width           =   975
            End
            Begin VB.Label libLrCdr02 
               Caption         =   "(Estd : )"
               Height          =   250
               Left            =   2400
               TabIndex        =   16
               Top             =   700
               Width           =   1215
            End
            Begin VB.Label libLrCdr04 
               Caption         =   "(Estd : )"
               Height          =   250
               Left            =   2400
               TabIndex        =   15
               Top             =   1900
               Width           =   1215
            End
            Begin VB.Label lblStatus05 
               Caption         =   "05 - Transfert terminé"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   2300
               Width           =   1695
            End
            Begin VB.Label lblStatus04 
               Caption         =   "04 - Transfert en cours"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   1905
               Width           =   1935
            End
            Begin VB.Label lblStatus03 
               Caption         =   "03 - Transfert demandé"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   1500
               Width           =   1935
            End
            Begin VB.Label lblStatus02 
               Caption         =   "02 - Extraction terminée"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   1100
               Width           =   1935
            End
            Begin VB.Label lblStatus01 
               Caption         =   "01 - Extraction en cours"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   700
               Width           =   1815
            End
            Begin VB.Label lblStatus00 
               Caption         =   "00 - Extraction demandée"
               Height          =   250
               Left            =   120
               TabIndex        =   9
               Top             =   300
               Width           =   2055
            End
            Begin VB.Label libLrTiers04 
               Caption         =   "(Tiers :)"
               Height          =   250
               Left            =   4000
               TabIndex        =   8
               Top             =   1900
               Width           =   1215
            End
            Begin VB.Label libLrTiers02 
               Caption         =   "(Tiers :)"
               Height          =   250
               Left            =   4000
               TabIndex        =   7
               Top             =   700
               Width           =   1215
            End
         End
      End
      Begin ComctlLib.ProgressBar prgBar 
         Height          =   300
         Left            =   2040
         TabIndex        =   22
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
         Max             =   15000
      End
      Begin VB.Label lblDTCENTEnCours 
         Caption         =   "Traitement en cours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   2175
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
Attribute VB_Name = "frmLucaRisques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String

Dim recAccAut As typeAccAut

Dim IdShell
Dim recAs400Cmd As typeAs400Cmd
Dim recLrCdr As typeLrCdr
Dim recLrTiers As typeLrTiers

Dim blnActualiserAuto As Boolean, blnExtraireTransfert As Boolean
Dim strStatusMax As String * 2

Dim LrCdrNb As Integer, LrTiersNb As Integer

Dim recLrCdr_Run As typeAccAut
Dim recLrCdrMsg As typeLrCdrMsg

Dim totalBiaLrRisque As typeLrRisque
Dim totalBiaLrRetris As typeLrRetris
Dim totalBdfLrRisque As typeLrRisque
Dim totalBdfLrRetris As typeLrRetris

Dim wLrRisque As typeLrRisque
Dim wLrRetris As typeLrRetris

'---------------------------------------------------------
Public Sub frmLrTiersD_Show()
'---------------------------------------------------------
Dim X As String

frmLrTiers.Show vbModeless
frmLrTiers.WindowState = vbNormal
frmLrTiers.Visible = True
'frmLrTiers.frmLrTiersDInit
X = frmLrTiers.Caption
AppActivate X

End Sub


Public Sub AccAut_Load()

srvAccAut.Init recLrCdr_Run
recLrCdr_Run.Method = "SeekP0"
recLrCdr_Run.AccAutId = "SRVBIALR"
recLrCdr_Run.AccAutK1 = "AUTO"
recLrCdr_Run.AccAutK2 = "LRCDR_RUN"
If Not IsNull(srvAccAut.Monitor(recLrCdr_Run)) Then Unload Me

fraLrBia.Enabled = False

If Trim(recLrCdr_Run.AccAutTxt) <> "" Then
    X = MsgBox("Le module LrCdr est en cours d'utilisation par : " & Trim(recLrCdr_Run.AccAutTxt) & Chr$(10) & "Voulez-vous continuer ?", vbYesNo + vbQuestion + vbDefaultButton2, "Autorisation : AccAut ( SRVBIALR / AUTO / LRCDR_RUN)")
Else
    X = vbYes
End If
If X = vbYes Then
    fraLrBia.Enabled = True
    recLrCdr_Run.AccAutTxt = usrId
    recLrCdr_Run.AccAutDD = DSys
    recLrCdr_Run.AccAutHD = time_Hms
    AccAut_Update
End If
End Sub
Public Sub AccAut_Unload()
If fraLrBia.Enabled Then
    recLrCdr_Run.AccAutTxt = ""
    recLrCdr_Run.AccAutDF = DSys
    recLrCdr_Run.AccAutHF = time_Hms
    AccAut_Update
End If

End Sub


Public Sub AccAut_Update()
recLrCdr_Run.Method = constUpdate
If Not IsNull(srvAccAut.Update(recLrCdr_Run)) Then
    Call lstErr_AddItem(lstErr, cmdActualiser, "AccAut : mise à jour non effectuée")
End If

End Sub



Public Sub cmdActualiser_Display()

Dim X2 As String * 2

srvAccAut.Init recAccAut
recAccAut.Method = "SeekP0"
recAccAut.AccAutId = "SRVBIALR"
recAccAut.AccAutK1 = "AUTO"
recAccAut.AccAutK2 = "LRCDR"
If Not IsNull(srvAccAut.Monitor(recAccAut)) Then Unload Me

libLrCdrD = dateImp(recAccAut.AccAutDD)
libLrCdrH = timeImp(recAccAut.AccAutHD)

libLrTransfertD = dateImp(recAccAut.AccAutDF)
libLrTransfertH = timeImp(recAccAut.AccAutHF)

'Select Case Mid$(recAccAut.AccAutTxt, 5, 1)
'    Case Is = "V": optSoldeVeille = "1"
'    Case Else: optSoldeFinDeMois = "1"
'End Select
X2 = mId$(recAccAut.AccAutTxt, 1, 2)
lblStatus00.ForeColor = IIf(X2 = "00", warnUsrColor, lblUsr.ForeColor)
lblStatus01.ForeColor = IIf(X2 = "01", warnUsrColor, lblUsr.ForeColor)
lblStatus02.ForeColor = IIf(X2 = "02", warnUsrColor, lblUsr.ForeColor)
lblStatus03.ForeColor = IIf(X2 = "03", warnUsrColor, lblUsr.ForeColor)
lblStatus04.ForeColor = IIf(X2 = "04", warnUsrColor, lblUsr.ForeColor)
lblStatus05.ForeColor = IIf(X2 = "05", warnUsrColor, lblUsr.ForeColor)

LrCdrNb = Val(mId$(recAccAut.AccAutTxt, 11, 5))
libLrCdr02 = "( Estd :" & Format$(LrCdrNb, "####0") & " ) "
LrTiersNb = Val(mId$(recAccAut.AccAutTxt, 16, 5))
libLrTiers02 = "( Tiers :" & Format$(LrTiersNb, "####0") & " ) "

'If X2 = "02" Or X2 = "05" Then
    cmdTransfert.Enabled = True
'Else
'    cmdTransfert.Enabled = False
'End If

If X2 = "05" Then
'    cmdLrEngineStart.Enabled = True
    cmdLrEngineEnd.Enabled = True
Else
'    cmdLrEngineStart.Enabled = False
    cmdLrEngineEnd.Enabled = False
End If

End Sub

Private Sub cmdActualiser_Click()
mnuActualiserStop.Enabled = blnActualiserAuto
PopupMenu mnuActualiser
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
Public Sub cmdContext_Quit()
If blnActualiserAuto Then
    cmdActualiser_Stop
Else

    If blnMsgBox_Quit Then
       X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
       X = vbYes
    End If
    If X = vbYes Then Unload Me
End If

End Sub

Public Sub cmdContext_Return()

SendKeys "{TAB}"

End Sub



Public Sub cmdUpdate()
recAccAut.Method = constUpdate
If Not IsNull(srvAccAut.Update(recAccAut)) Then
    Call lstErr_AddItem(lstErr, cmdActualiser, "Mise à jour non effectuée")
End If
cmdActualiser_Display

End Sub


Public Sub fraEnabled(bln As Boolean)
frmLucaRisques.Enabled = bln
End Sub

Private Sub cmdExport_LrCdrAller_Click()
Dim xFileName As String
Dim X As String, X240 As String * 240
Dim I As Integer

xFileName = paramLrCdr_BdfSend_Filename
Call lstErr_AddItem(lstErr, cmdTransfert, "Transfert Aller BDF vers" & xFileName)
X = Dir(xFileName)
If X <> "" Then Kill xFileName

I = 0
Open paramLrCdr_LrBdfAller_FileName For Input As #1
Open xFileName For Output As #2

Do Until EOF(1)
    Line Input #1, X
    I = I + 1
    X240 = X
    Print #2, X240
Loop

Close
Call lstErr_AddItem(lstErr, cmdTransfert, I & " enregistrements exportés")


'========================================Version bande sur AS400 supprimé le 14-10-1999
'FileCopy paramLrCdr_lrBdfAller, xFileName

'srvAs400Cmd.Init recAs400Cmd
'recAs400Cmd.Method = "SBMJOB"

'X = "SBMJOB CMD(CALL PGM(" & constLucaRisques_LrCdrAller & "))"
'recAs400Cmd.Text = X & " JOB(" & constLucaRisques_LrCdrAller & ") USER(" & Trim(usrId) & ") JOBQ(QINTER)"
'srvAs400Cmd.Update recAs400Cmd
'===================================================================
End Sub

Private Sub cmdExport_LrTiersX_Click()
cmdExport_LrTiers

End Sub

Private Sub cmdExtractionTransfert_Click()
blnExtraireTransfert = True
cmdLrCdr_Click
End Sub

Private Sub cmdImport_LrCdrBDF_Click()
Dim xFileName As String, Iter As Integer, X As String, intLrCdrBdfNb As Integer
Dim DTCENT_Estd As String

On Error GoTo cmdTransfert_Error

If lstErr.ListCount > 0 Then Call lstErr_AddItem(lstErr, cmdTransfert, "Corrigez les anomalies"): Exit Sub

fraEnabled False

'xFileName = constFTP_Dir & "LrCdrBdf"
'Call lstErr_AddItem(lstErr, cmdImport_LrCdrBDF, "Suppression " & xFileName)
'X = Dir(xFileName)
'If X <> "" Then Kill xFileName

'Call lstErr_AddItem(lstErr, cmdImport_LrCdrBDF, "Transfert AS400 => " & constFTP_Dir)
'srvAs400Cmd.Init recAs400Cmd
'recAs400Cmd.Method = "SBMJOB"
'X = "SBMJOB CMD(CALL PGM(" & constLucaRisques_AS400BdfTrf & ")) "
'recAs400Cmd.Text = X & " JOB(" & constLucaRisques_AS400BdfTrf & ") USER(" & Trim(usrId) & ") JOBQ(QINTER)"
'srvAs400Cmd.Update recAs400Cmd


'Iter = 0
'If LrCdrNb > 0 Then prgBar.Max = 31000
'Do
'    DoEvents
'    X = Dir(xFileName)
'    Iter = Iter + 1
'    prgBar.Value = Iter
'    If Iter > 30000 Then
'        X = MsgBox("Voulez-vous réessayer ?", vbQuestion, "LucaRisques : cmdTransfert : FTP en cours ")
'        If X = vbYes Then
'            Iter = 0
'        Else
'            Err = 9999: GoTo cmdTransfert_Error
'        End If
'    End If
'Loop While X = ""


'xFileName = constFTP_Dir & "LrCdrBdf"

Call lstErr_AddItem(lstErr, cmdTransfert, "Transfert retour BDF : " & paramLrCdr_BdfReceive_Filename)

X = Dir(paramLrCdr_BdfReceive_Filename)
If X = "" Then
    Call lstErr_AddItem(lstErr, cmdTransfert, "fichier n'existe pas ")
    GoTo cmdTransfert_Error
End If

xFileName = paramLrCdr_BdfReceive_Filename & "_" & DTCENTenCours
X = Dir(xFileName)
If X <> "" Then
    Call lstErr_AddItem(lstErr, cmdTransfert, "fichier existe déjà :" & xFileName)
    GoTo cmdTransfert_Error
End If


Name paramLrCdr_BdfReceive_Filename As xFileName
Open xFileName For Input As #2


intLrCdrBdfNb = 0
xFileName = paramLrCdr_LrBdfRetour_FileName
Open xFileName For Output As #1
Do Until EOF(2)
    DoEvents
    Line Input #2, X
    intLrCdrBdfNb = intLrCdrBdfNb + 1
    Print #1, X
Loop
Close #2
Close #1
DTCENT_Estd = mId$(X, 7, 6)


fraEnabled True
Call lstErr_AddItem(lstErr, cmdImport_LrCdrBDF, Format(intLrCdrBdfNb, "#####0") & " : " & paramLrCdr_LrBdfRetour_FileName)

'lstErr.Clear: lstErr.Height = 200
If DTCENT_Estd <> DTCENTenCours Then
    X = MsgBox("Date centralisation des ESTD :" & DTCENT_Estd, vbCritical, "Interface AS400 => Luca Report : transfert")
End If
Exit Sub

'---------------------------------------------------------
cmdTransfert_Error:
'---------------------------------------------------------

MsgBox "erreur : " & Err & " " & Error$(Err), vbCritical, "LucaRisques : cmdTransfert : " & xFileName
Resume cmdTransfert_End

cmdTransfert_End:

End Sub

Private Sub cmdImport_LrTiersX_Click()
cmdImport_LrTiers ""
End Sub

Private Sub cmdLrCdr_Click()
X = MsgBox("Voulez-vous réellement lancer une nouvelle extraction ?", vbYesNo + vbQuestion + vbDefaultButton2, "Interface AS400 => Luca Report")
If X <> vbYes Then Exit Sub
fraEnabled False
recAccAut.AccAutTxt = "00"
Mid$(recAccAut.AccAutTxt, 5, 1) = "M"
''''If optSoldeVeille = "1" Then Mid$(recAccAut.AccAutTxt, 5, 1) = "V"

recAccAut.AccAutDD = "00000000"
recAccAut.AccAutHD = "000000"
recAccAut.AccAutDF = "00000000"
recAccAut.AccAutHF = "000000"

cmdUpdate

srvAs400Cmd.Init recAs400Cmd
recAs400Cmd.Method = "SBMJOB"
'!!!!X = "SBMJOB CMD(CALL PGM(paramLrCdr_AS400_Ext) PARM('" & Mid$(recAccAut.AccAutTxt, 5, 1) & "'))"
X = "SBMJOB CMD(CALL PGM(" & paramLrCdr_AS400_Ext & "))"
recAs400Cmd.Text = X & " JOB(" & paramLrCdr_AS400_Ext & ") USER(" & Trim(usrId) & ") JOBQ(QINTER)"
srvAs400Cmd.Update recAs400Cmd
strStatusMax = "02"
prgBar.Max = 15000
cmdActualiser_Start

End Sub

Private Sub cmdLrEngineEnd_Click()
IdShell = Shell(paramLrCdr_Archive_Proc & Format$(DTCENTenCours, " 0000"), 1)

End Sub


Private Sub cmdLrTiers_Update_Click()
frmLrTiersD_Show
End Sub

Private Sub cmdPrintSopra_Click()
If chkPrintSopra220 = "1" Then prtLucaRisques_SopraX paramLrCdr_PrintSopra_220
If chkPrintSopra400 = "1" Then prtLucaRisques_SopraX paramLrCdr_PrintSopra_400
If chkPrintSopra470 = "1" Then prtLucaRisques_SopraX paramLrCdr_PrintSopra_470
If chkPrintSopra490 = "1" Then prtLucaRisques_SopraX paramLrCdr_PrintSopra_490
If chkPrintSopraBmAller = "1" Then prtLucaRisques_SopraX paramLrCdr_LrBdfAller_FileName
If chkPrintSopraBmRetour = "1" Then
    prtLucaRisques_SopraX paramLrCdr_PrintSopra_870
    prtLucaRisques_SopraX paramLrCdr_PrintSopra_880
End If
End Sub

Private Sub cmdTransfert_Click()
Dim xFileName As String, Iter As Integer, X As String, intLrTiersNb As Integer
Dim DTCENT_Estd As String

On Error GoTo cmdTransfert_Error

If lstErr.ListCount > 0 Then Call lstErr_AddItem(lstErr, cmdTransfert, "Corrigez les anomalies"): Exit Sub

Timer.Enabled = False
fraEnabled False
cmdActualiser_Display

Mid$(recAccAut.AccAutTxt, 1, 4) = "04"

xFileName = paramLrCdr_AS400_LrEstd ' constFTP_Dir & "LrCdr"
Call lstErr_AddItem(lstErr, cmdTransfert, "Suppression " & xFileName)
X = Dir(xFileName)
If X <> "" Then Kill xFileName


xFileName = paramLrCdr_AS400_LrCliCli  'constFTP_Dir & "LrTiers"
Call lstErr_AddItem(lstErr, cmdTransfert, "Suppression " & xFileName)
X = Dir(xFileName)
If X <> "" Then Kill xFileName


cmdUpdate

prgBar.Visible = True
If LrCdrNb > 0 Then prgBar.Max = LrCdrNb
prgBar.Value = 0
libLrCdr04.Visible = True
libLrCdr04.ForeColor = warnUsrColor


Call lstErr_AddItem(lstErr, cmdTransfert, "Transfert AS400 => " & paramLrCdr_AS400_LrEstd)
srvAs400Cmd.Init recAs400Cmd
recAs400Cmd.Method = "SBMJOB"
X = "SBMJOB CMD(CALL PGM(" & paramLrCdr_AS400_Trf & ")) "
recAs400Cmd.Text = X & " JOB(" & paramLrCdr_AS400_Trf & ") USER(" & Trim(usrId) & ") JOBQ(QINTER)"
srvAs400Cmd.Update recAs400Cmd

Iter = 0
If LrCdrNb > 0 Then prgBar.Max = 31000
Do
    DoEvents
    X = Dir(xFileName)
    Iter = Iter + 1
    prgBar.Value = Iter
    If Iter > 30000 Then
        X = MsgBox("Voulez-vous réessayer ?", vbQuestion, "LucaRisques : cmdTransfert : FTP en cours ")
        If X = vbYes Then
            Iter = 0
        Else
            Err = 9999: GoTo cmdTransfert_Error
        End If
    End If
Loop While X = ""


xFileName = paramLrCdr_AS400_LrEstd  'constFTP_Dir & "LrCdr"
Open xFileName For Input As #2

arrLrCdrNb = 0
xFileName = paramLrCdr_LrEstd_FileName
Call lstErr_AddItem(lstErr, cmdTransfert, "Copie vers " & xFileName)

Open xFileName For Output As #1
If LrCdrNb > 0 Then prgBar.Max = LrCdrNb

Do Until EOF(2)
    DoEvents
    Line Input #2, X
    arrLrCdrNb = arrLrCdrNb + 1
    Print #1, X
    libLrCdr04 = "( Estd :" & Format$(arrLrCdrNb, "####0") & " ) "
    If LrCdrNb > arrLrCdrNb Then prgBar.Value = arrLrCdrNb
Loop
DTCENT_Estd = mId$(X, 46, 6)

Close #2
Close #1
libLrCdr04.ForeColor = libUsr.ForeColor

solde:
If LrTiersNb > 0 Then prgBar.Max = LrTiersNb
prgBar.Value = 0
libLrTiers04.Visible = True
libLrTiers04.ForeColor = warnUsrColor

xFileName = paramLrCdr_AS400_LrCliCli 'constFTP_Dir & "lrTiers"
Open xFileName For Input As #2


intLrTiersNb = 0
xFileName = paramLrCdr_LrClicli_FileName & "_AS400"
Open xFileName For Output As #1
Do Until EOF(2)
    DoEvents
    Line Input #2, X
    intLrTiersNb = intLrTiersNb + 1
    Print #1, X
    libLrTiers04 = "( Tiers :" & Format$(intLrTiersNb, "####0") & " ) "
    If LrTiersNb > intLrTiersNb Then prgBar.Value = intLrTiersNb
Loop
Close #2
Close #1
prgBar.Visible = False
libLrTiers04.ForeColor = libUsr.ForeColor


strStatusMax = "05"
prgBar.Max = 5000
cmdActualiser_Start
Mid$(recAccAut.AccAutTxt, 1, 2) = "05"
cmdUpdate
fraEnabled True
If LrCdrNb <> arrLrCdrNb Then
    X = MsgBox("Estd générées :" & Format$(LrCdrNb, "####0") & Chr$(10) & "Estd transférées :" & Format$(arrLrCdrNb, "####0"), vbCritical, "Interface AS400 => Luca Report : transfert")
End If
If LrTiersNb <> intLrTiersNb Then
    X = MsgBox("Tiers générés :" & Format$(LrTiersNb, "####0") & Chr$(10) & "Tiers transférés :" & Format$(intLrTiersNb, "####0"), vbCritical, "Interface AS400 => Luca Report : transfert")
End If

If chkLrTiersInit = "1" Then SGNBNF_Init

lstErr.Clear: lstErr.Height = 200
If DTCENT_Estd <> DTCENTenCours Then
    X = MsgBox("Date centralisation des ESTD :" & DTCENT_Estd, vbCritical, "Interface AS400 => Luca Report : transfert")
End If
Exit Sub

'---------------------------------------------------------
cmdTransfert_Error:
'---------------------------------------------------------

MsgBox "erreur : " & Err & " " & Error$(Err), vbCritical, "LucaRisques : cmdTransfert : " & xFileName
Resume cmdTransfert_End

cmdTransfert_End:

End Sub

Private Sub Form_Activate()
If Not fraLrBia.Enabled Then Unload Me

End Sub


Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------

End Sub

Private Sub cmdPrint_Click()
Dim Msg As String
lstErr.Clear

txtDTCENTenCours_Control

If lstErr.ListCount = 0 Then

    Msg = "000000000000"
    mdbLrRisque.tableLrRisque_Open
    mdbLrRETRIS.tableLrRetris_Open
    mdbLrSgnBnf.tableLrSgnBnf_Open
    mdbLrSort.tableLrSort_Open
    
    If chkPrintRFBENF = "1" Then prtLucaRisquesX "RFBENF" & Trim(txtRFBENF)
    If chkPrintRFBENFAll = "1" Then prtLucaRisquesX "RFBALL"
    If chkPrintTotal = "1" Then prtLucaRisquesX "TOTAL "
    If chkPrintLstBNF = "1" Then prtLucaRisquesX "LSTBNF"
    If chkPrintLstBNFPM = "1" Then prtLucaRisquesX "LSTBPM"
    If chkPrintTOTBDF = "1" Then prtLucaRisquesX "TOTBDF"
    If chkPrintCOTBDF = "1" Then prtLucaRisquesX "COTBDF"
    If chkExport = "1" Then prtLucaRisquesX "EXPORT" & Trim(txtRFBENF)
    
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
AccAut_Unload
End Sub


Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

srvLrCdr.param_Init

libImport_Dta.ForeColor = warnUsrColor
X = dateElp("FinDeMoisP", -1, DSys)
If Not IsNumeric(X) Then
    Call MsgBox("Erreur calcul fin de mois précédent", vbCritical, "Luca Risques")
    X = DSys
End If
txtDTCENTEnCours = mId$(X, 1, 6): DTCENTenCours = txtDTCENTEnCours
'$$$$$$$$$$$$$$$$$
lstErr.Clear
blnExtraireTransfert = False
libLrCdr04.Visible = False
cmdActualiser_Display
cmdActualiser_Stop
AccAut_Load
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

Private Sub prgBar_Click()
cmdActualiser_Stop
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


Private Sub cmdImport_LrSgnBnf()
Dim X As String, xOut As String * 500
On Error Resume Next

Dim I As Integer

MDB.Execute "delete * from LrSgnBnf"
mdbLrSgnBnf.tableLrSgnBnf_Open

X = paramLrCdr_LrSgnBnf_Filename
libImport_FileName = X

I = 0
Open X For Input As #1

Do Until EOF(1)
    Line Input #1, X
    I = I + 1
    xOut = X
    libImport_Dta = xOut: DoEvents
    Import_LrSgnBnf xOut, recLrSgnBnf
    If recLrSgnBnf.CDBANQ = strSocBdfE Then
        recLrSgnBnf.Method = constAddNew
        dbLrSgnBnf_Update recLrSgnBnf
    Else
        MsgBox xOut, vbCritical, "cmdImport_LrSgnBnf : " & I
    End If
Loop
libImport_Dta = "Enregistrements lus : " & I
Close #1
mdbLrSgnBnf.tableLrSgnBnf_Close

End Sub


Private Sub cmdImport_LrTiers(strDest As String)
Dim X As String, xOut As String * 500
On Error Resume Next

Dim I As Integer

MDB.Execute "delete * from LrTiers"
mdbLrTiers.tableLrTiers_Open

X = paramLrCdr_LrClicli_FileName & strDest
libImport_FileName = X

I = 0
Open X For Input As #1

Do Until EOF(1)
    Line Input #1, X
    I = I + 1
    xOut = X
    libImport_Dta = xOut: DoEvents
    Import_LrTiers xOut, recLrTiers
    If recLrTiers.CDDECL = strSocBdfE Then
        recLrTiers.Method = constAddNew
        recLrTiers.FILL02 = DTCENTenCours
        dbLrTiers_Update recLrTiers
    Else
        MsgBox xOut, vbCritical, "cmdImport_LrTiers : " & I
    End If
Loop
libImport_Dta = "Enregistrements lus : " & I
Close #1
mdbLrTiers.tableLrTiers_Close

End Sub

Private Sub cmdExport_LrTiers()
Dim X As String, xOut As String * 500, intReturn As Integer
On Error Resume Next

Dim I As Integer
fraLrBia.Enabled = False

mdbLrTiers.tableLrTiers_Open

X = paramLrCdr_LrClicli_FileName
libImport_FileName = X

I = 0
Open X For Output As #1

recLrTiers.Method = "MoveFirst"
xOut = Space(500)
Do
    intReturn = tableLrTiers_Read(recLrTiers)
    If intReturn = 0 Then
        If Trim(recLrTiers.FILL02) = DTCENTenCours Then
            recLrTiers.Method = "MoveNext"
            recLrTiers.FILL02 = ""
            Export_LrTiers xOut, recLrTiers
            Print #1, Trim(xOut)
            I = I + 1
        End If
    End If
Loop While intReturn = 0

Call lstErr_Clear(lstErr, cmdPrint, "Bénéficiaires exportés : " & I)
Close #1
mdbLrTiers.tableLrTiers_Close
IdShell = Shell(paramLrCdr_LrClicli_Proc, 1)
fraLrBia.Enabled = True

End Sub


Private Sub cmdImport_LrRetris()
Dim X As String, xOut As String * 480
On Error Resume Next

Dim I As Integer

MDB.Execute "delete * from LrRetris"
mdbLrRETRIS.tableLrRetris_Open

X = paramLrCdr_LrRetris_Filename
libImport_FileName = X

I = 0
Open X For Input As #1
Open "C:\BiaSrv\JPLSGNBNF.wri" For Output As #2

Do Until EOF(1)
    Line Input #1, X
    I = I + 1
    xOut = X
    libImport_Dta = xOut: DoEvents
    If mId$(xOut, 40, 2) = "H2" Then
        Import_LrRetris xOut, recLrRetris, optFRF
        recLrRetris.Method = constAddNew
'       dbLrRetris_Update recLrRetris
        tableLrRetris_Update recLrRetris   '!!!!!!!!!! Ne gère pas les erreurs (3022 ....)
'    Else
    '        Print #2, Mid$(xOut, 1, 100)
    End If
    
Loop
libImport_Dta = "Enregistrements lus : " & I
Close #2

Close #1
mdbLrRETRIS.tableLrRetris_Close

End Sub


Private Sub cmdImport_LrRisque()
Dim X As String, xOut As String * 450, X2 As String
Dim I As Integer

On Error Resume Next

MDB.Execute "delete * from LrRisque"
mdbLrRisque.tableLrRisque_Open

I = 0
X = paramLrCdr_LrRisque_Filename
libImport_FileName = X
Open X For Input As #1
''X = Input(209, #1)     '20020103 V8 avant => X = Input(43, #1)


Do Until EOF(1)
 '   X = Input(452, #1)
  '   X = Input(439, #1): X2 = Input(13, #1) ' !!! dernier enregistrement du fichier SOPRA  tronqué
      Line Input #1, X
      I = I + 1
    xOut = X
        
    libImport_Dta = xOut: DoEvents
    Import_LrRisque xOut, recLrRisque, optFRF
    recLrRisque.Method = constAddNew
    tableLrRisque_Update recLrRisque
Loop
libImport_Dta = "Enregistrements lus : " & I
Close #1

mdbLrRisque.tableLrRisque_Close

End Sub



Private Sub cmdImport_Click()
Dim X1 As String

fraLrBia.Enabled = False
cmdImport.Visible = False

CV_Init LrCdr_CV1
LrCdr_CV1.OpéAmj = DSys
LrCdr_CV2 = LrCdr_CV1
LrCdr_CV3 = CV_Euro

LrCdr_CV2.DeviseIso = "FRF"
LrCdr_CV1.AchatVente = " "
LrCdr_CV2.AchatVente = " "
LrCdr_CV1.Normal = "P"
LrCdr_CV2.Normal = "P"
If optFRF Then
    LrCdr_CV1.DeviseIso = "EUR"
    LrCdr_CV2.DeviseIso = "FRF"
Else
    LrCdr_CV1.DeviseIso = "FRF"
    LrCdr_CV2.DeviseIso = "EUR"
End If
LrCdr_CV1.Montant = 100
Call CV_Transitoire(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3, X1)

cmdImport_LrRisque
cmdImport_LrRetris
cmdImport_LrSgnBnf
If chkCDCPCO_Add = "1" Then cmdImport_CDCPCO_Add

cmdImport_Totalisation

fraLrBia.Enabled = True
cmdImport.Visible = True

End Sub


Private Sub txtDTCENTenCours_GotFocus()
txt_GotFocus txtDTCENTEnCours

End Sub


Private Sub txtDTCENTenCours_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtDTCENTenCours_LostFocus()
txt_LostFocus txtDTCENTEnCours
txtDTCENTenCours_Control

End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub

Public Sub txtDTCENTenCours_Control()
Dim D As Double, Dj As Double

D = Val(txtDTCENTEnCours)
Dj = Val(mId$(DSys, 1, 6))
If D < Dj - 100 Then Call lstErr_AddItem(lstErr, txtDTCENTEnCours, "Préciser AAAAMM _enCours >:" & Dj): Exit Sub
If D > Dj Then Call lstErr_AddItem(lstErr, txtDTCENTEnCours, "AAAAMM _enCours < " & Dj): Exit Sub
DTCENTenCours = Format$(D, "000000")
End Sub


Public Sub cmdImport_Totalisation()
Dim intReturn As Integer, intReturn2 As Integer
Dim arrCotBdf(100) As String * 4, arrCotBdf_Nb As Integer, arrCotBdf_Index As Integer, arrCotBdf_Bln As Boolean
Dim xCOTBDF As String * 4
Dim X As String

arrCotBdf_Nb = 0
libImport_FileName = "Totalisaton"

mdbLrRisque.tableLrRisque_Open
mdbLrRETRIS.tableLrRetris_Open

MDB.Execute "delete * from LrSort"
mdbLrSort.tableLrSort_Open
cmdImport_TotalisationLrSort "Init"

recLrRisque.Method = "MoveFirst"

Do
    intReturn = tableLrRisque_Read(recLrRisque)
    If intReturn = 0 Then
        libImport_Dta = recLrRisque.RFBENF & "  " & recLrRisque.DTCENT1: DoEvents
        If recLrRisque.CDBANQ = strSocBdfE Then
        
            If recLrSort.RFBENF <> recLrRisque.RFBENF _
            Or recLrSort.CDCPCO <> recLrRisque.CDCPCO Then cmdImport_TotalisationLrSort constUpdate
            recLrSort.RFBENF = recLrRisque.RFBENF
            If recLrSort.DTCENT1 < recLrRisque.DTCENT1 Then
                recLrSort.DTCENT1 = recLrRisque.DTCENT1
                recLrSort.MTTOTAL = recLrRisque.MTTOTAL
                recLrSort.CDCPCO = recLrRisque.CDCPCO
            End If

            recLrRetris.RFBENF = recLrRisque.RFBENF
            recLrRetris.DTCENT1 = recLrRisque.DTCENT1
            recLrRetris.CDCPCO = recLrRisque.CDCPCO
            recLrRetris.Method = "Seek="
            
            Call dbLrRetris_Read(recLrRetris)
            If intReturn2 = 0 Then
                If Trim(recLrRetris.COTBDF) = "" Then recLrRetris.COTBDF = "0000"
                xCOTBDF = mId$(recLrRetris.COTBDF, 2, 3)
                
                arrCotBdf_Bln = False
                For arrCotBdf_Index = 0 To arrCotBdf_Nb
                    If xCOTBDF = arrCotBdf(arrCotBdf_Index) Then arrCotBdf_Bln = True: Exit For
                Next arrCotBdf_Index
                If Not arrCotBdf_Bln Then
                    arrCotBdf_Nb = arrCotBdf_Nb + 1
                    arrCotBdf(arrCotBdf_Nb) = xCOTBDF
                End If
                
                cmdImport_TotalisationLrRisque "TOTAL"
                cmdImport_TotalisationLrRetris "TOTAL"
                cmdImport_TotalisationLrRisque "TOTAL" & xCOTBDF
                cmdImport_TotalisationLrRetris "TOTAL" & xCOTBDF
            End If
        End If
    End If
    recLrRisque.Method = "Seek>"
Loop While intReturn = 0

cmdImport_TotalisationLrSort constUpdate

mdbLrSgnBnf.tableLrSgnBnf_Open

Msg = Space$(600)
mdbLrSgnBnf.Import_LrSgnBnf Msg, recLrSgnBnf
recLrSgnBnf.RFBENF = "TOTAL"
If optFRF Then
    X = "de FRF)"
Else
    X = "d'EUROS)"
End If
recLrSgnBnf.NOMBNF = "Centralisation des risques bancaires (en Milliers " & X

recLrSgnBnf.Method = constAddNew
dbLrSgnBnf_Update recLrSgnBnf

For arrCotBdf_Index = 0 To arrCotBdf_Nb
    recLrSgnBnf.RFBENF = "TOTAL" & arrCotBdf(arrCotBdf_Index)
    recLrSgnBnf.NOMBNF = "Cotation Banque de France " & arrCotBdf(arrCotBdf_Index)
    recLrSgnBnf.Method = constAddNew
    dbLrSgnBnf_Update recLrSgnBnf
Next arrCotBdf_Index


mdbLrRisque.tableLrRisque_Close
mdbLrRETRIS.tableLrRetris_Close
mdbLrSgnBnf.tableLrSgnBnf_Close
mdbLrSort.tableLrSort_Close

End Sub

Public Sub cmdImport_TotalisationLrRetris(strTotal As String)

totalBiaLrRetris.RFBENF = strTotal
totalBiaLrRetris.DTCENT1 = recLrRisque.DTCENT1
totalBiaLrRetris.CDCPCO = "1"
totalBiaLrRetris.Method = "Seek="
If (tableLrRetris_Read(totalBiaLrRetris)) = 0 Then
    totalBiaLrRetris.Method = constUpdate
Else
    Msg = Space$(600)
    mdbLrRETRIS.Import_LrRetris Msg, totalBiaLrRetris, optFRF
    totalBiaLrRetris.RFBENF = strTotal
    totalBiaLrRetris.DTCENT1 = recLrRisque.DTCENT1
    totalBiaLrRetris.CDCPCO = "1"
    totalBiaLrRetris.Method = constAddNew
End If

totalBiaLrRetris.MT01 = totalBiaLrRetris.MT01 + recLrRetris.MT01
totalBiaLrRetris.MT02 = totalBiaLrRetris.MT02 + recLrRetris.MT02
totalBiaLrRetris.MT03 = totalBiaLrRetris.MT03 + recLrRetris.MT03
totalBiaLrRetris.MT04 = totalBiaLrRetris.MT04 + recLrRetris.MT04
totalBiaLrRetris.MT05 = totalBiaLrRetris.MT05 + recLrRetris.MT05
totalBiaLrRetris.MT06 = totalBiaLrRetris.MT06 + recLrRetris.MT06
totalBiaLrRetris.MT07 = totalBiaLrRetris.MT07 + recLrRetris.MT07
totalBiaLrRetris.MT08 = totalBiaLrRetris.MT08 + recLrRetris.MT08
totalBiaLrRetris.MT09 = totalBiaLrRetris.MT09 + recLrRetris.MT09
totalBiaLrRetris.MT10 = totalBiaLrRetris.MT10 + recLrRetris.MT10
totalBiaLrRetris.MT11 = totalBiaLrRetris.MT11 + recLrRetris.MT11
totalBiaLrRetris.MT12 = totalBiaLrRetris.MT12 + recLrRetris.MT12
totalBiaLrRetris.MT13 = totalBiaLrRetris.MT13 + recLrRetris.MT13
totalBiaLrRetris.MT14 = totalBiaLrRetris.MT14 + recLrRetris.MT14
totalBiaLrRetris.MT15 = totalBiaLrRetris.MT15 + recLrRetris.MT15
totalBiaLrRetris.MT16 = totalBiaLrRetris.MT16 + recLrRetris.MT16
totalBiaLrRetris.MT17 = totalBiaLrRetris.MT17 + recLrRetris.MT17
totalBiaLrRetris.MT18 = totalBiaLrRetris.MT18 + recLrRetris.MT18
totalBiaLrRetris.MT19 = totalBiaLrRetris.MT19 + recLrRetris.MT19
totalBiaLrRetris.MT20 = totalBiaLrRetris.MT20 + recLrRetris.MT20
totalBiaLrRetris.MT21 = totalBiaLrRetris.MT21 + recLrRetris.MT21
totalBiaLrRetris.MT22 = totalBiaLrRetris.MT22 + recLrRetris.MT22
totalBiaLrRetris.MT23 = totalBiaLrRetris.MT23 + recLrRetris.MT23
totalBiaLrRetris.MT24 = totalBiaLrRetris.MT24 + recLrRetris.MT24
totalBiaLrRetris.MT25 = totalBiaLrRetris.MT25 + recLrRetris.MT25
totalBiaLrRetris.MTTOTAL = totalBiaLrRetris.MTTOTAL + recLrRetris.MTTOTAL

dbLrRetris_Update totalBiaLrRetris
End Sub
Public Sub cmdImport_TotalisationLrRisque(strTotal As String)

totalBiaLrRisque.RFBENF = strTotal
totalBiaLrRisque.DTCENT1 = recLrRisque.DTCENT1
totalBiaLrRisque.CDCPCO = "1"
totalBiaLrRisque.Method = "Seek="
If (tableLrRisque_Read(totalBiaLrRisque)) = 0 Then
    totalBiaLrRisque.Method = constUpdate
Else
    Msg = Space$(600)
    mdbLrRisque.Import_LrRisque Msg, totalBiaLrRisque, optFRF
    totalBiaLrRisque.RFBENF = strTotal
    totalBiaLrRisque.DTCENT1 = recLrRisque.DTCENT1
    totalBiaLrRisque.CDCPCO = "1"
    totalBiaLrRisque.Method = constAddNew
End If

totalBiaLrRisque.MT01 = totalBiaLrRisque.MT01 + recLrRisque.MT01
totalBiaLrRisque.MT02 = totalBiaLrRisque.MT02 + recLrRisque.MT02
totalBiaLrRisque.MT03 = totalBiaLrRisque.MT03 + recLrRisque.MT03
totalBiaLrRisque.MT04 = totalBiaLrRisque.MT04 + recLrRisque.MT04
totalBiaLrRisque.MT05 = totalBiaLrRisque.MT05 + recLrRisque.MT05
totalBiaLrRisque.MT06 = totalBiaLrRisque.MT06 + recLrRisque.MT06
totalBiaLrRisque.MT07 = totalBiaLrRisque.MT07 + recLrRisque.MT07
totalBiaLrRisque.MT08 = totalBiaLrRisque.MT08 + recLrRisque.MT08
totalBiaLrRisque.MT09 = totalBiaLrRisque.MT09 + recLrRisque.MT09
totalBiaLrRisque.MT10 = totalBiaLrRisque.MT10 + recLrRisque.MT10
totalBiaLrRisque.MT11 = totalBiaLrRisque.MT11 + recLrRisque.MT11
totalBiaLrRisque.MT12 = totalBiaLrRisque.MT12 + recLrRisque.MT12
totalBiaLrRisque.MT13 = totalBiaLrRisque.MT13 + recLrRisque.MT13
totalBiaLrRisque.MT14 = totalBiaLrRisque.MT14 + recLrRisque.MT14
totalBiaLrRisque.MT15 = totalBiaLrRisque.MT15 + recLrRisque.MT15
totalBiaLrRisque.MT16 = totalBiaLrRisque.MT16 + recLrRisque.MT16
totalBiaLrRisque.MT17 = totalBiaLrRisque.MT17 + recLrRisque.MT17
totalBiaLrRisque.MT18 = totalBiaLrRisque.MT18 + recLrRisque.MT18
totalBiaLrRisque.MT19 = totalBiaLrRisque.MT19 + recLrRisque.MT19
totalBiaLrRisque.MT20 = totalBiaLrRisque.MT20 + recLrRisque.MT20
totalBiaLrRisque.MTTOTAL = totalBiaLrRisque.MTTOTAL + recLrRisque.MTTOTAL

dbLrRisque_Update totalBiaLrRisque
End Sub


Public Sub cmdImport_TotalisationLrSort(Msg As String)

If Msg = constUpdate And Trim(recLrSort.RFBENF) <> "" Then dbLrSort_Update recLrSort

recLrSort.MTTOTAL = 0
recLrSort.RFBENF = ""
recLrSort.DTCENT1 = ""
recLrSort.CDCPCO = ""
recLrSort.Method = constAddNew

End Sub

Public Sub SGNBNF_Init()
Dim intReturn As Integer, I As Integer, X As String

X = MsgBox("Annule et remplace,  voulez-vous continuer ?", vbQuestion + vbYesNo, "Fiches signalétiques des bénéficiaires")
If X <> vbYes Then Exit Sub

cmdImport_LrSgnBnf
cmdImport_LrTiers "_AS400"

mdbLrTiers.tableLrTiers_Open
mdbLrSgnBnf.tableLrSgnBnf_Open

recLrSgnBnf.Method = "MoveFirst"

Do
    intReturn = tableLrSgnBnf_Read(recLrSgnBnf)
    If intReturn = 0 Then
        recLrTiers.RFBENF = recLrSgnBnf.RFBENF
        recLrTiers.FILL02 = DTCENTenCours
        recLrTiers.Method = "Seek="
        I = tableLrTiers_Read(recLrTiers)
        If I = 0 Then
            recLrTiers.NSIREN = recLrSgnBnf.NSIREN
            recLrTiers.NOMBNF = recLrSgnBnf.NOMBNF
            recLrTiers.PRENOM = recLrSgnBnf.PRENOM
            recLrTiers.CDSEXE = recLrSgnBnf.CDSEXE
            recLrTiers.DTNAIS = recLrSgnBnf.JMA3
            recLrTiers.CDPAYS1 = recLrSgnBnf.CDPAYS1
            recLrTiers.CDDEPT1 = recLrSgnBnf.CDDEPT1
            recLrTiers.CDCOMM1 = recLrSgnBnf.CDCOMM1
            recLrTiers.LBCOMM1 = recLrSgnBnf.LBCOMM1
            recLrTiers.NOMCJT = recLrSgnBnf.NOMCJT
            recLrTiers.CDACCO = recLrSgnBnf.CDACCO
            recLrTiers.CTJURI = recLrSgnBnf.CTJURI
            recLrTiers.CDRESI = recLrSgnBnf.CDRESI
            recLrTiers.NOVOIE = recLrSgnBnf.NOVOIE
            recLrTiers.CDPOST = recLrSgnBnf.CDPOST
            recLrTiers.LBCOMM2 = recLrSgnBnf.LBCOMM2
            recLrTiers.CDDEPT2 = recLrSgnBnf.CDDEPT2
            recLrTiers.CDPAYS2 = recLrSgnBnf.CDPAYS2
            recLrTiers.CDTRI1 = recLrSgnBnf.CDTRI1
            recLrTiers.CDTRI2 = recLrSgnBnf.CDTRI2
            recLrTiers.CDAGCO = recLrSgnBnf.CDAGCO

            recLrTiers.Method = constUpdate
            
            Vx = dbLrTiers_Update(recLrTiers)
        End If
        recLrSgnBnf.Method = "MoveNext"
    End If
Loop While intReturn = 0

mdbLrTiers.tableLrTiers_Close
mdbLrSgnBnf.tableLrSgnBnf_Close

cmdExport_LrTiers

End Sub

Public Sub cmdImport_CDCPCO_Add()
mdbLrRisque.tableLrRisque_Open
mdbLrRETRIS.tableLrRetris_Open

recLrRisque.Method = "MoveFirst"

Do
    intReturn = tableLrRisque_Read(recLrRisque)
    If intReturn = 0 Then
        If recLrRisque.CDBANQ = strSocBdfE Then
        
            If recLrRisque.CDCPCO <> "1" Then
                wLrRisque.RFBENF = recLrRisque.RFBENF
                wLrRisque.DTCENT1 = recLrRisque.DTCENT1
                wLrRisque.CDCPCO = "1"
                wLrRisque.Method = "Seek="
                intReturn2 = tableLrRisque_Read(wLrRisque)
                If intReturn2 = 0 Then
                    wLrRisque.Method = constUpdate
                    wLrRisque.MT01 = wLrRisque.MT01 + recLrRisque.MT01
                    wLrRisque.MT02 = wLrRisque.MT02 + recLrRisque.MT02
                    wLrRisque.MT03 = wLrRisque.MT03 + recLrRisque.MT03
                    wLrRisque.MT04 = wLrRisque.MT04 + recLrRisque.MT04
                    wLrRisque.MT05 = wLrRisque.MT05 + recLrRisque.MT05
                    wLrRisque.MT06 = wLrRisque.MT06 + recLrRisque.MT06
                    wLrRisque.MT07 = wLrRisque.MT07 + recLrRisque.MT07
                    wLrRisque.MT08 = wLrRisque.MT08 + recLrRisque.MT08
                    wLrRisque.MT09 = wLrRisque.MT09 + recLrRisque.MT09
                    wLrRisque.MT10 = wLrRisque.MT10 + recLrRisque.MT10
                    wLrRisque.MT11 = wLrRisque.MT11 + recLrRisque.MT11
                    wLrRisque.MT12 = wLrRisque.MT12 + recLrRisque.MT12
                    wLrRisque.MT13 = wLrRisque.MT13 + recLrRisque.MT13
                    wLrRisque.MT14 = wLrRisque.MT14 + recLrRisque.MT14
                    wLrRisque.MT15 = wLrRisque.MT15 + recLrRisque.MT15
                    wLrRisque.MT16 = wLrRisque.MT16 + recLrRisque.MT16
                    wLrRisque.MT17 = wLrRisque.MT17 + recLrRisque.MT17
                    wLrRisque.MT18 = wLrRisque.MT18 + recLrRisque.MT18
                    wLrRisque.MT19 = wLrRisque.MT19 + recLrRisque.MT19
                    wLrRisque.MT20 = wLrRisque.MT20 + recLrRisque.MT20
                    wLrRisque.MTTOTAL = wLrRisque.MTTOTAL + recLrRisque.MTTOTAL
                Else
                    wLrRisque = recLrRisque
                    wLrRisque.CDCPCO = "1"
    
                    wLrRisque.Method = constAddNew
                End If
                
                dbLrRisque_Update wLrRisque
                recLrRisque.Method = constDelete
                If tableLrRisque_Read(recLrRisque) = 0 Then dbLrRisque_Update recLrRisque

               
                recLrRetris.RFBENF = recLrRisque.RFBENF
                recLrRetris.DTCENT1 = recLrRisque.DTCENT1
                recLrRetris.CDCPCO = recLrRisque.CDCPCO
                recLrRetris.Method = "Seek="
                intReturn3 = tableLrRetris_Read(recLrRetris)
                If intReturn3 = 0 Then
                    wLrRetris.RFBENF = recLrRisque.RFBENF
                    wLrRetris.DTCENT1 = recLrRisque.DTCENT1
                    wLrRetris.CDCPCO = "1"
                    wLrRetris.Method = "Seek="
                    intReturn2 = tableLrRetris_Read(wLrRetris)
                    If intReturn2 = 0 Then
                        wLrRetris.Method = constUpdate
                        wLrRetris.MT01 = wLrRetris.MT01 + recLrRetris.MT01
                        wLrRetris.MT02 = wLrRetris.MT02 + recLrRetris.MT02
                        wLrRetris.MT03 = wLrRetris.MT03 + recLrRetris.MT03
                        wLrRetris.MT04 = wLrRetris.MT04 + recLrRetris.MT04
                        wLrRetris.MT05 = wLrRetris.MT05 + recLrRetris.MT05
                        wLrRetris.MT06 = wLrRetris.MT06 + recLrRetris.MT06
                        wLrRetris.MT07 = wLrRetris.MT07 + recLrRetris.MT07
                        wLrRetris.MT08 = wLrRetris.MT08 + recLrRetris.MT08
                        wLrRetris.MT09 = wLrRetris.MT09 + recLrRetris.MT09
                        wLrRetris.MT10 = wLrRetris.MT10 + recLrRetris.MT10
                        wLrRetris.MT11 = wLrRetris.MT11 + recLrRetris.MT11
                        wLrRetris.MT12 = wLrRetris.MT12 + recLrRetris.MT12
                        wLrRetris.MT13 = wLrRetris.MT13 + recLrRetris.MT13
                        wLrRetris.MT14 = wLrRetris.MT14 + recLrRetris.MT14
                        wLrRetris.MT15 = wLrRetris.MT15 + recLrRetris.MT15
                        wLrRetris.MT16 = wLrRetris.MT16 + recLrRetris.MT16
                        wLrRetris.MT17 = wLrRetris.MT17 + recLrRetris.MT17
                        wLrRetris.MT18 = wLrRetris.MT18 + recLrRetris.MT18
                        wLrRetris.MT19 = wLrRetris.MT19 + recLrRetris.MT19
                        wLrRetris.MT20 = wLrRetris.MT20 + recLrRetris.MT20
                        wLrRetris.MT21 = wLrRetris.MT21 + recLrRetris.MT21
                        wLrRetris.MT22 = wLrRetris.MT22 + recLrRetris.MT22
                        wLrRetris.MT23 = wLrRetris.MT23 + recLrRetris.MT23
                        wLrRetris.MT24 = wLrRetris.MT24 + recLrRetris.MT24
                        wLrRetris.MT25 = wLrRetris.MT25 + recLrRetris.MT25
                        wLrRetris.MTTOTAL = wLrRetris.MTTOTAL + recLrRetris.MTTOTAL
                Else
                    wLrRetris = recLrRetris
                    wLrRetris.CDCPCO = "1"
    
                    wLrRetris.Method = constAddNew
                End If


                    dbLrRetris_Update wLrRetris
                    recLrRetris.Method = constDelete
                    If tableLrRetris_Read(recLrRetris) = 0 Then dbLrRetris_Update recLrRetris
                End If

            End If
        End If
    End If
    recLrRisque.Method = "Seek>"
Loop While intReturn = 0
mdbLrRisque.tableLrRisque_Close
mdbLrRETRIS.tableLrRetris_Close

End Sub
