VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCompteSolde 
   Caption         =   "Etat des soldes"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9180
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   48
      Top             =   0
      Width           =   3585
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exporter"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   0
      Width           =   1140
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10821
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Choix d'un état"
      TabPicture(0)   =   "CompteSolde.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraBalance"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sélection des comptes"
      TabPicture(1)   =   "CompteSolde.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOptions"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Présentation de l'état"
      TabPicture(2)   =   "CompteSolde.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraPrésentationEtat"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraPrésentationEtat 
         Height          =   5535
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   8775
         Begin VB.Frame fraEnTete 
            Height          =   2055
            Left            =   240
            TabIndex        =   38
            Top             =   3360
            Width           =   8295
            Begin VB.TextBox txtExport_Filename 
               Height          =   285
               Left            =   2280
               TabIndex        =   47
               Top             =   1560
               Width           =   4935
            End
            Begin VB.TextBox txtEnTete 
               Height          =   285
               Left            =   2280
               TabIndex        =   42
               Text            =   "Etat des soldes"
               Top             =   480
               Width           =   4935
            End
            Begin VB.TextBox txtDestinataire 
               Height          =   285
               Left            =   2280
               TabIndex        =   39
               Top             =   1080
               Width           =   4935
            End
            Begin VB.Label lblExport_Filename 
               Caption         =   "Fichier d'exportation"
               Height          =   255
               Left            =   240
               TabIndex        =   46
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label lblEnTete 
               Caption         =   "En Tête de l'état"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblDestinataire 
               Caption         =   "Destinataire de l'état"
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   1080
               Width           =   1575
            End
         End
         Begin VB.Frame fraPrint 
            Caption         =   "Impression"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   4320
            TabIndex        =   24
            Top             =   360
            Width           =   4215
            Begin VB.CheckBox chkPrintRuptureRacine 
               Caption         =   "Ligne total racine "
               Height          =   375
               Left            =   120
               TabIndex        =   82
               Top             =   1100
               Width           =   3255
            End
            Begin VB.CheckBox chkPrintReliure 
               Caption         =   "marge pour reliure"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   2300
               Width           =   3015
            End
            Begin VB.CheckBox chkPrintTotal 
               Caption         =   "Récapitulatif (type / devise)"
               Height          =   375
               Left            =   120
               TabIndex        =   28
               Top             =   1500
               Width           =   3255
            End
            Begin VB.CheckBox chkPrintRupture 
               Caption         =   "Ligne total racine / Type"
               Height          =   375
               Left            =   120
               TabIndex        =   27
               Top             =   700
               Width           =   3255
            End
            Begin VB.CheckBox chkPrintLine 
               Caption         =   "Ligne détail des comptes auxilaires"
               Height          =   375
               Left            =   120
               TabIndex        =   26
               Top             =   300
               Value           =   1  'Checked
               Width           =   3255
            End
            Begin VB.CheckBox chkPrintSoldé 
               Caption         =   "Imprimer les comptes soldés"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   1900
               Width           =   2295
            End
         End
         Begin VB.Frame fraSort 
            Caption         =   "Tri"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   3855
            Begin VB.OptionButton optSort6 
               Caption         =   " Racine (B, HB)"
               Height          =   375
               Left            =   120
               TabIndex        =   79
               Top             =   2520
               Width           =   3495
            End
            Begin VB.OptionButton optSort5 
               Caption         =   "Type / Racine / Devise"
               Height          =   375
               Left            =   120
               TabIndex        =   59
               Top             =   2160
               Width           =   3495
            End
            Begin VB.OptionButton optSort4 
               Caption         =   "Pays Résidence / Type / Racine / Devise"
               Height          =   375
               Left            =   120
               TabIndex        =   51
               Top             =   1800
               Width           =   3495
            End
            Begin VB.OptionButton optSort3 
               Caption         =   "Devise / Racine / Type / N° ordre "
               Height          =   375
               Left            =   120
               TabIndex        =   23
               Top             =   1320
               Width           =   3495
            End
            Begin VB.OptionButton optSort2 
               Caption         =   "Racine / N° ordre / Devise / Type "
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   840
               Width           =   3375
            End
            Begin VB.OptionButton optSort1 
               Caption         =   "Racine / Type / N° ordre / Devise"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   360
               Value           =   -1  'True
               Width           =   3255
            End
         End
      End
      Begin VB.Frame fraOptions 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   8775
         Begin VB.Frame fraSelect 
            Caption         =   "Sélection"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5175
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   8535
            Begin VB.CheckBox chkNostroNo 
               Caption         =   "exclure comptes Nostros"
               Height          =   255
               Left            =   5400
               TabIndex        =   81
               Top             =   4320
               Width           =   2655
            End
            Begin VB.CheckBox chkNostro 
               Caption         =   "uniquement comptes Nostros"
               Height          =   255
               Left            =   5400
               TabIndex        =   80
               Top             =   3960
               Width           =   2655
            End
            Begin VB.CheckBox chkComptePP 
               Caption         =   "uniquement comptes 'P P'"
               Height          =   255
               Left            =   5400
               TabIndex        =   58
               Top             =   3600
               Width           =   2655
            End
            Begin VB.CheckBox chkComptePM 
               Caption         =   "uniquement comptes 'P M'"
               Height          =   255
               Left            =   5400
               TabIndex        =   57
               Top             =   3240
               Width           =   2655
            End
            Begin VB.CheckBox chkCompteClient 
               Caption         =   "uniquement comptes 'Client'"
               Height          =   255
               Left            =   5400
               TabIndex        =   56
               Top             =   2880
               Width           =   2655
            End
            Begin VB.CheckBox chkCompteBanque 
               Caption         =   "uniquement comptes 'Banque'"
               Height          =   255
               Left            =   5400
               TabIndex        =   55
               Top             =   2520
               Width           =   2655
            End
            Begin VB.ListBox lstDevise 
               Height          =   2010
               Left            =   5400
               TabIndex        =   54
               Top             =   240
               Width           =   3015
            End
            Begin VB.CheckBox chkPays 
               Caption         =   "sélectionner le pays (code BdF)"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   4400
               Width           =   2535
            End
            Begin VB.TextBox txtPays 
               Height          =   285
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   52
               Top             =   4400
               Width           =   495
            End
            Begin VB.TextBox txtDeviseCV 
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   49
               Text            =   "EUR"
               Top             =   400
               Width           =   495
            End
            Begin VB.CheckBox chkDeviseIn 
               Caption         =   "Uniquement devises In et Euro"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   800
               Width           =   2775
            End
            Begin VB.TextBox txtBiaTyp 
               Height          =   285
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   18
               Top             =   4000
               Width           =   495
            End
            Begin VB.TextBox txtGestionnaire 
               Height          =   285
               Left            =   3240
               MaxLength       =   2
               TabIndex        =   17
               Top             =   3600
               Width           =   495
            End
            Begin VB.TextBox txtService 
               Height          =   285
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   16
               Top             =   3240
               Width           =   495
            End
            Begin VB.CheckBox chkBiaTyp 
               Caption         =   "sélectionner le type de compte"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   4000
               Width           =   2535
            End
            Begin VB.CheckBox chkGestionnaire 
               Caption         =   "sélectionner le gestionnaire"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   3600
               Width           =   2535
            End
            Begin VB.CheckBox chkService 
               Caption         =   "sélectionner le service gestionnaire"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   3240
               Width           =   2895
            End
            Begin VB.CheckBox chkCompteMinMax 
               Caption         =   "sélectionner les comptes de"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   2400
               Width           =   2535
            End
            Begin VB.CheckBox chkCompteHorsBilan 
               Caption         =   "sélectionner les comptes de hors-bilan"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   2000
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.CheckBox chkCompteBilan 
               Caption         =   "sélectionner les comptes de bilan"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   1600
               Value           =   1  'Checked
               Width           =   2895
            End
            Begin VB.TextBox txtDevise 
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   9
               Top             =   1200
               Width           =   495
            End
            Begin VB.CheckBox chkDevise 
               Caption         =   "Sélectionner la devise"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   1200
               Width           =   2055
            End
            Begin VB.TextBox txtCompteMax 
               Height          =   285
               Left            =   3000
               MaxLength       =   11
               TabIndex        =   6
               Top             =   2800
               Width           =   1575
            End
            Begin VB.TextBox txtCompteMin 
               Height          =   285
               Left            =   3000
               MaxLength       =   11
               TabIndex        =   5
               Top             =   2400
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "devise de contre-valeur"
               Height          =   255
               Left            =   360
               TabIndex        =   50
               Top             =   400
               Width           =   2415
            End
            Begin VB.Label lblMax 
               Caption         =   "à"
               Height          =   255
               Left            =   2280
               TabIndex        =   7
               Top             =   2800
               Width           =   255
            End
         End
      End
      Begin VB.Frame fraBalance 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   8775
         Begin VB.Frame fraScript 
            Caption         =   "Script"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3735
            Left            =   240
            TabIndex        =   60
            Top             =   240
            Width           =   8415
            Begin VB.OptionButton optEtatDafi 
               Caption         =   "Etat DAFI (racine /N° ordre)"
               Height          =   255
               Left            =   2520
               TabIndex        =   78
               Top             =   2760
               Width           =   2535
            End
            Begin VB.OptionButton optEtatCrédocPays 
               Caption         =   "CréDoc (pays / type/racine)"
               Height          =   255
               Left            =   2520
               TabIndex        =   77
               Top             =   1320
               Width           =   2295
            End
            Begin VB.OptionButton optEtatManuel 
               Caption         =   "Manuel"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   360
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optEtatTCDeviseIn 
               Caption         =   "Etat TC Devise In"
               Height          =   255
               Left            =   5640
               TabIndex        =   75
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optEtatTCCAD 
               Caption         =   "Etat TC CAD"
               Height          =   255
               Left            =   5640
               TabIndex        =   74
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optEtatTCGBP 
               Caption         =   "Etat TC GBP"
               Height          =   255
               Left            =   5640
               TabIndex        =   73
               Top             =   1320
               Width           =   1335
            End
            Begin VB.OptionButton optEtatTCUSD 
               Caption         =   "Etat TC USD"
               Height          =   255
               Left            =   5640
               TabIndex        =   72
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton optEtatDafiPays 
               Caption         =   "Etat DAFI (pays / type/racine)"
               Height          =   255
               Left            =   2520
               TabIndex        =   71
               Top             =   3240
               Width           =   2535
            End
            Begin VB.OptionButton optEtatBOTCBanque 
               Caption         =   "Etat BOTC (banques)"
               Height          =   255
               Left            =   5640
               TabIndex        =   70
               Top             =   2400
               Width           =   1935
            End
            Begin VB.OptionButton optEtatBOTCPM 
               Caption         =   "Etat BOTC (PM)"
               Height          =   255
               Left            =   5640
               TabIndex        =   69
               Top             =   2040
               Width           =   1575
            End
            Begin VB.OptionButton optEtatFOTCBanque 
               Caption         =   "Etat FOTC (banques)"
               Height          =   255
               Left            =   5640
               TabIndex        =   68
               Top             =   3240
               Width           =   1935
            End
            Begin VB.OptionButton optEtatFOTCPM 
               Caption         =   "Etat FOTC (PM)"
               Height          =   255
               Left            =   5640
               TabIndex        =   67
               Top             =   2880
               Width           =   1575
            End
            Begin VB.OptionButton optEtatDG 
               Caption         =   "Etat DG"
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   2040
               Width           =   975
            End
            Begin VB.OptionButton optEtatDGA 
               Caption         =   "Etat DGA"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   2400
               Width           =   1695
            End
            Begin VB.OptionButton optEtatInspection 
               Caption         =   "Etat Inspection"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   2760
               Width           =   1455
            End
            Begin VB.OptionButton optEtatInspection019 
               Caption         =   "Etat Inspection 019"
               Height          =   255
               Left            =   120
               TabIndex        =   63
               Top             =   3240
               Width           =   1935
            End
            Begin VB.OptionButton optEtatCaisse 
               Caption         =   "Etat Caisse"
               Height          =   255
               Left            =   2520
               TabIndex        =   62
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton optEtatCrédocTdC 
               Caption         =   "Etat CréDoc (type/racine)"
               Height          =   255
               Left            =   2520
               TabIndex        =   61
               Top             =   1800
               Width           =   2175
            End
         End
         Begin VB.Frame fraEtat 
            Caption         =   "Etat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1400
            Left            =   240
            TabIndex        =   35
            Top             =   4080
            Width           =   3855
            Begin VB.OptionButton optEtatCptGen 
               Caption         =   "Etat des soldes des comptes généraux"
               Height          =   375
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Width           =   3135
            End
            Begin VB.OptionButton optEtatCptAux 
               Caption         =   "Etat des soldes des comptes auxiliaires"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   300
               Value           =   -1  'True
               Width           =   3135
            End
         End
         Begin VB.Frame fraSolde 
            Caption         =   "des soldes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1400
            Left            =   5160
            TabIndex        =   30
            Top             =   4080
            Width           =   3375
            Begin VB.OptionButton optFinDeMoisOpération 
               Caption         =   "Fin de mois (date d'opération)"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   750
               Width           =   2415
            End
            Begin VB.OptionButton optFinDAnnée 
               Caption         =   "Fin d'année"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   1000
               Width           =   2415
            End
            Begin VB.OptionButton optFinDeMois 
               Caption         =   "Fin de mois (date de traitement)"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   500
               Width           =   2655
            End
            Begin VB.OptionButton optVeille 
               Caption         =   "Veille"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   250
               Value           =   -1  'True
               Width           =   2415
            End
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   8640
      Picture         =   "CompteSolde.frx":0054
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
End
Attribute VB_Name = "frmCompteSolde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean, blnSetfocus As Boolean
Dim CompteSoldeAut As typeAuthorization
Dim X As String, X1 As String, I As Long
'Dim Msg As String, valX As String, V As Variant
Dim reccptp0 As typeCptP0
Dim recCompte As typeCompte, recRacine As typeRacine

Dim optEtat As String * 1, optSolde As String * 1, optAmj As String * 8, SrvCptP0_Amj As String * 8
Dim blnCompteMinMax As Boolean, selCompteMin As String * 11, selCompteMax As String * 11
Dim blnDevise As Boolean, selDeviseN As String * 3, blnDeviseIn As Boolean, selDeviseCV As String * 3
Dim blnService As Boolean, selService As String * 3
Dim blnGestionnaire As Boolean, selGestionnaire As String * 2
Dim blnBiaTyp As Boolean, selBiaTyp As String * 3
Dim blnPays As Boolean, selPays As String * 3
Dim optSortK As String * 1
Dim blnCompteBilan As Boolean, blnCompteHorsBilan As Boolean
Dim blnCompteBanque As Boolean, blnCompteClient As Boolean, blnComptePP As Boolean, blnComptePM As Boolean
Dim blnNostroNo As Boolean, blnNostro As Boolean
Dim optEtatSortK As String * 2
Dim mDestinataire As String, mEnTete As String
Dim PrintRupture_Len As Integer

Dim blnExport As Boolean, X200 As String * 200
Dim X1000 As String * 1000
Dim cmdImport_Select_Nb As Long, cmdImport_Nb As Long

Dim blnService_Enabled As Boolean
Dim wL As Long, wPAys As String * 4, wX As String
Dim recdictio As typeDictio
Dim optEtatCrédocBanque As Boolean, optEtatNostro  As Boolean
Dim optEtatPersonnel  As Boolean, optEtatRisques As Boolean, optEtatCompte_010_3 As Boolean
Dim optEtatCompteConversion As Boolean
Dim mEtatPass As Integer

Dim recElpBuffer As typeElpBuffer

Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Then
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
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Then
            
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
       Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
blnSetfocus = False
optEtatCrédocBanque = False
optEtatNostro = False
optEtatPersonnel = False
optEtatCompteConversion = False
optEtatRisques = False
optEtatCompte_010_3 = False
mEtatPass = 0

SrvCptP0_Amj = "00000000"
If UCase$(Trim(mId$(Msg, 1, 12))) = "COMPTE_SLD" Then
    blnService_Enabled = True 'False
    Call BiaPgmAut_Init("Compte_Sld", CompteSoldeAut)
Else
    blnService_Enabled = True
    Call BiaPgmAut_Init("Compte_Sld+", CompteSoldeAut)
End If

If Not IsNull(param_Init) Then cmdPrint.Visible = False: cmdExport.Visible = False
cmdReset
'''''test chkCompteMinMax = "1": txtCompteMin = "10002000000": txtCompteMax = "30000999999"
Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "COMPTE_001": optEtatCaisse = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_010": optEtatCrédocTdC = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_010_2": optEtatCrédocBanque = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_010_3": optEtatCompte_010_3 = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_000": optEtatDG = True: optEtat_Script: cmdPrint_Click
                    optEtatNostro = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_000_2": optEtatDGA = True: optEtat_Script: cmdPrint_Click
                    optEtatNostro = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_000_3": optEtatInspection = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_000_4": optEtatInspection019 = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_041": optEtatDafi = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_050": optEtatBOTCBanque = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_050_2": optEtatBOTCPM = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_060": optEtatPersonnel = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_060_2": optEtatCompteConversion = True: optEtat_Script: cmdPrint_Click: Unload Me
    Case "COMPTE_070": optEtatRisques = True: optEtat_Script: cmdPrint_Click
                         mEtatPass = 2: optEtat_Script: cmdPrint_Click: Unload Me
    Case Else
            blnSetfocus = True

End Select


End Sub


Private Sub cmdImport_CptP0()
Dim xInput As String, blnOk As Boolean
Dim vReturn As Variant
On Error Resume Next

Dim I As Integer

blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
cmdControl
If lstErr.ListCount <> 0 Then Exit Sub

optSolde = " "
If optVeille Then optSolde = "V": optAmj = dateElp("Ouvré", -1, DSys) 'SrvCptP0_Amj
If optFinDeMois Then optSolde = "M": optAmj = dateElp("FinDeMoisP", 0, DSys)
If optFinDeMoisOpération Then optSolde = "O": optAmj = dateElp("FinDeMoisP", 0, DSys)
If optFinDAnnée Then optSolde = "A": optAmj = dateElp("FinDAnnéeP", 0, DSys)

X = Dir(paramComptaSld_Cpt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier des comptes n'existe pas"): Exit Sub

Call lstErr_Clear(lstErr, cmdPrint, "Chargement des comptes, tri ...")
frmCompteSolde.MousePointer = vbHourglass
frmCompteSolde.Enabled = False
Call prtCompteSolde_CV_Init(optAmj, selDeviseCV)

MDB.Execute "delete * from CptP0"
mdbCptP0.tableCptP0_Open

Open paramComptaSld_Cpt_Import For Input As #1
recCptP0_Init reccptp0
reccptp0.Method = "AddNew"

If blnExport Then Open paramComptaSld_Cpt_Export For Output As #2

Do Until EOF(1)
    Line Input #1, xInput
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        SrvCptP0_Amj = mId$(xInput, 35, 8)
        I = Val(mId$(xInput, 43, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteSolde : cmdImport_Cptp0 :SrvCptP0 ")
            Exit Do
        End If
    End If
    cmdImport_Nb = cmdImport_Nb + 1
    vReturn = cmdImport_Select(xInput)
    If vReturn <> "" Then
        reccptp0.Id = vReturn & Format$(cmdImport_Nb, "000000")
        reccptp0.Text = xInput
        If blnExport Then
            cmdExport_Write
        Else
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            dbCptP0_Update reccptp0
        End If
    End If
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Sélection : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
 
Loop

Close
mdbCptP0.tableCptP0_Close
frmCompteSolde.MousePointer = 0

If Not blnOk Then
    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteSolde : cmdImport_Cptp0 :SrvCptP0 ")
End If

End Sub

Public Function param_Init()
Dim V
param_Init = Null
recElpTable_Init recElpTable
recElpTable.Id = "Param"
recElpTable.K1 = "ComptaSld"
recElpTable.Method = "Seek="

recElpTable.K2 = "Cpt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramComptaSld_Cpt_Import = paramServer(recElpTable.Memo)
'''Call lstErr_Clear(lstErr, cmdContext, "Fichier :" & paramComptaSld_Cpt_Import)

recElpTable.K2 = "Cpt_Export"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramComptaSld_Cpt_Export = paramServer(recElpTable.Memo)

recElpTable.K2 = "Mvt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramComptaSld_Mvt_Import = paramServer(recElpTable.Memo)

recElpTable.K2 = "Rac_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramComptaSld_Rac_Import = paramServer(recElpTable.Memo)

Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "CompteSolde_Param_Init"
Exit Function

End Function


Private Sub chkBiaTyp_Click()
If chkBiaTyp = "1" Then
    txtBiaTyp.Visible = True: If blnSetfocus Then txtBiaTyp.SetFocus
Else
    txtBiaTyp.Visible = False
End If
End Sub

Private Sub chkBiaTyp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkBiaTyp
End Sub


Private Sub chkCompteBanque_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCompteBanque
End Sub


Private Sub chkCompteBilan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCompteBilan
End Sub


Private Sub chkCompteClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCompteClient
End Sub


Private Sub chkCompteHorsBilan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCompteHorsBilan
End Sub


Private Sub chkCompteMinMax_Click()
If chkCompteMinMax = "1" Then
    txtCompteMin.Visible = True: txtCompteMax.Visible = True
    If blnSetfocus Then txtCompteMin.SetFocus
Else
    txtCompteMin.Visible = False: txtCompteMax.Visible = False
End If

End Sub

Private Sub chkCompteMinMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCompteMinMax
End Sub


Private Sub chkComptePM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkComptePM
End Sub


Private Sub chkComptePP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkComptePP
End Sub


Private Sub chkDevise_Click()
If chkDevise = "1" Then
    txtDevise.Visible = True: If blnSetfocus Then txtDevise.SetFocus
Else
    txtDevise.Visible = False
End If

End Sub

Private Sub chkDevise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDevise
End Sub


Private Sub chkDeviseIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDeviseIn
End Sub


Private Sub chkGestionnaire_Click()
If chkGestionnaire = "1" Then
    txtGestionnaire.Visible = True: If blnSetfocus Then txtGestionnaire.SetFocus
Else
    txtGestionnaire.Visible = False
End If

End Sub

Private Sub chkGestionnaire_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkGestionnaire
End Sub


Private Sub chkNostroNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkNostroNo
End Sub


Private Sub chkNostro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkNostro
End Sub


Private Sub chkPays_Click()
If chkPays = "1" Then
    txtPays.Visible = True: If blnSetfocus Then txtPays.SetFocus
Else
    txtPays.Visible = False
End If

End Sub

Private Sub chkPays_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPays

End Sub


Private Sub chkPrintLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPrintLine
End Sub


Private Sub chkPrintReliure_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPrintReliure
End Sub


Private Sub chkPrintRupture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPrintRupture
End Sub


Private Sub chkPrintRuptureRacine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPrintRuptureRacine
End Sub


Private Sub chkPrintSoldé_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPrintSoldé
End Sub


Private Sub chkPrintTotal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPrintTotal
End Sub


Private Sub chkService_Click()
On Error GoTo Exit_Sub
If chkService = "1" Then
    txtService.Visible = True: If blnSetfocus Then txtService.SetFocus
Else
    txtService.Visible = False
End If
Exit_Sub:
End Sub

Private Sub chkService_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkService
End Sub


Private Sub cmdExport_Click()
If Trim(txtExport_Filename) = "" Then
    Call lstErr_Clear(lstErr, cmdPrint, "? préciser le nom du fichier d'export")
Else
    blnExport = True
    cmdImport_CptP0
    blnExport = False
    Call lstErr_AddItem(lstErr, cmdPrint, "Export terminé : " & cmdImport_Select_Nb)
    frmCompteSolde.Enabled = True
    AppActivate frmCompteSolde.Caption
End If
End Sub

Private Sub cmdPrint_Click()
Dim X, Nb As Integer, curX As Currency, IdKey As String, mIdKey As String
Dim Msg As String

cmdImport_CptP0
cmdImport_Racine

If cmdImport_Select_Nb = 0 Then
    Call lstErr_AddItem(lstErr, cmdPrint, "Aucun compte sélectionné !")
    GoTo cmdPrint_End
End If

Msg = "000000000000" & Space$(50)
Mid$(Msg, 14, 3) = selDeviseCV
Mid$(Msg, 17, 1) = "B"
Mid$(Msg, 18, 1) = optSolde
Mid$(Msg, 19, 8) = optAmj
Mid$(Msg, 27, 1) = IIf(chkPrintSoldé = "1", "S", " ")
Mid$(Msg, 28, 1) = IIf(chkPrintReliure = "1", ">", "=")
Mid$(Msg, 29, 1) = IIf(chkPrintLine = "1", "L", "-")
Mid$(Msg, 30, 1) = IIf(chkPrintRupture = "1", "R", "-")
Mid$(Msg, 31, 1) = IIf(chkPrintTotal = "1", "T", "-")
Mid$(Msg, 32, 2) = optEtatSortK
Mid$(Msg, 34, 1) = IIf(chkPrintRuptureRacine = "1", "R", "-")

prtCompteSolde_Open Msg, mEnTete, mDestinataire
Call lstErr_AddItem(lstErr, cmdPrint, "Impression : début")
mdbCptP0.tableCptP0_Open
Mid$(MsgTxt, 1, 34) = Space$(34)
reccptp0.Method = "MoveFirst"

X = dbCptP0_ReadE(reccptp0)
IdKey = mId$(reccptp0.Id, 1, PrintRupture_Len): arrCompteNb = 0

Do While reccptp0.Err = 0
    
    If IdKey <> mId$(reccptp0.Id, 1, PrintRupture_Len) Then
        Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Impression " & IdKey & " : " & arrCompteNb): DoEvents

        Call cmdPrint_Call(IdKey, Msg)
        IdKey = mId$(reccptp0.Id, 1, PrintRupture_Len): arrCompteNb = 0
        '''Exit Do
    End If
    
        MsgTxtIndex = 0
        Mid$(MsgTxt, 35, memoCptInfoLen) = mId$(reccptp0.Text, 1, memoCptInfoLen)
        If IsNull(srvCompteGetBuffer(recCompte)) Then
            recCompte.LibTyp = reccptp0.Id
            
            Call arrCompteAddItem(recCompte)
        End If
        
        arrCompte(arrCompteNb).SoldeInstantané = CptSolde_Get(optSolde, MsgTxt)
    
    reccptp0.Method = "MoveNext    "
    reccptp0.Err = tableCptP0_Read(reccptp0)
Loop

Call cmdPrint_Call(IdKey, Msg)
mdbCptP0.tableCptP0_Close
prtCompteSolde_Close
Call lstErr_AddItem(lstErr, cmdPrint, "Impression terminé : " & cmdImport_Select_Nb)

cmdPrint_End:
frmCompteSolde.Enabled = True
AppActivate frmCompteSolde.Caption

End Sub



Public Sub cmdControl()
lstErr.Clear
optEtat = "A"
If optEtatCptGen Then optEtat = "G"

blnCompteBilan = IIf(chkCompteBilan = "1", True, False)
blnCompteHorsBilan = IIf(chkCompteHorsBilan = "1", True, False)
blnCompteBanque = IIf(chkCompteBanque = "1", True, False)
blnCompteClient = IIf(chkCompteClient = "1", True, False)
blnComptePP = IIf(chkComptePP = "1", True, False)
blnComptePM = IIf(chkComptePM = "1", True, False)
blnNostroNo = IIf(chkNostroNo = "1", True, False)
blnNostro = IIf(chkNostro = "1", True, False)

blnCompteMinMax = IIf(chkCompteMinMax = "1", True, False)

selCompteMin = Format$(Val(Trim(txtCompteMin)), "00000000000")
selCompteMax = Format$(Val(Trim(txtCompteMax)), "00000000000")

If blnCompteMinMax Then
    If selCompteMin = "00000000000" Then
        Call lstErr_AddItem(lstErr, cmdContext, "? préciser le compte min")
    Else
        If selCompteMax = "00000000000" Then selCompteMax = selCompteMin
    End If
    If selCompteMin > selCompteMax Then Call lstErr_AddItem(lstErr, cmdContext, "? compte min > compte max")

End If

selDeviseCV = Trim(txtDeviseCV)
Call CV_AttributS(selDeviseCV, CV_X2)
selDeviseCV = CV_X2.DeviseIso

blnDeviseIn = IIf(chkDeviseIn = "1", True, False)
blnDevise = IIf(chkDevise = "1", True, False)
selDeviseN = Trim(txtDevise)
'If IsNumeric(selDeviseN) Then
'    CV_X1.DeviseN = Format$(selDeviseN, "000")
'    Call CV_AttributN(CV_X1)
    Call CV_AttributS(selDeviseN, CV_X1)
    selDeviseN = CV_X1.DeviseIso
'End If
If blnDeviseIn And blnDevise Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser 1 devise ou devise In")
If blnDevise Then
    If Trim(txtDevise) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser la devise")
End If


blnService = IIf(chkService = "1", True, False)
selService = Format$(Trim(txtService), "000")
If blnService Then
    If Trim(txtService) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le service")
End If

blnGestionnaire = IIf(chkGestionnaire = "1", True, False)
selGestionnaire = Format$(Trim(txtGestionnaire), "00")
If blnGestionnaire Then
    If Trim(txtGestionnaire) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le Gestionnaire")
End If

blnBiaTyp = IIf(chkBiaTyp = "1", True, False)
selBiaTyp = Format$(Trim(txtBiaTyp), "000")
If blnBiaTyp Then
    If Trim(txtBiaTyp) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le Type")
End If

blnPays = IIf(chkPays = "1", True, False)
selPays = Format$(Trim(txtPays), "000")
If blnPays Then
    If Trim(txtPays) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le Type")
End If

optSortK = 1
If optSort2 Then optSortK = 2
If optSort3 Then optSortK = 3
If optSort4 Then optSortK = 4
If optSort5 Then optSortK = 5
If optSort6 Then optSortK = 6

optEtatSortK = optEtat & optSortK
Select Case optEtatSortK
    Case "A1": PrintRupture_Len = 5
    Case "A2": PrintRupture_Len = 5
    Case "A3": PrintRupture_Len = 8
    Case "A4": PrintRupture_Len = 7
    Case "A5": PrintRupture_Len = 3
    Case "A6": PrintRupture_Len = 5
    Case "G1": PrintRupture_Len = 11
    Case "G2": PrintRupture_Len = 11
    Case "G3": PrintRupture_Len = 14
    Case "G4": PrintRupture_Len = 15

End Select

mDestinataire = Trim(txtDestinataire)
mEnTete = Trim(txtEnTete)
paramComptaSld_Cpt_Export = Trim(txtExport_Filename)

End Sub
Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub



Public Sub cmdContext_Return()

End Sub

Public Sub cmdContext_Quit()
Unload Me
End Sub

Private Function cmdImport_Select(Msg As String) As String
Dim wCompteGénéral As String * 11, wNuméro As String * 11, wDeviseN As String * 3, wBilan As String * 1
Dim X2 As String * 2

cmdImport_Select = ""

If optEtat = "A" Then
    If mId$(Msg, 115, 1) <> "A" Then Exit Function
End If

wDeviseN = mId$(Msg, 7, 3)
wNuméro = mId$(Msg, 13, 11)
wCompteGénéral = Format$(Val(mId$(Msg, 255, 11)), "00000000000")

If CV_X1.DeviseN <> wDeviseN Then
    CV_X1.DeviseN = wDeviseN
    Call CV_AttributN(CV_X1)
End If
wBilan = IIf(mId$(wCompteGénéral, 4, 1) = "9", "H", "B")
Mid$(Msg, 10, 3) = CV_X1.DeviseIso

If blnDeviseIn Then
    If Not CV_X1.EuroIn And CV_X1.DeviseIso <> "EUR" Then Exit Function
End If

If blnCompteMinMax Then
    If optEtat = "A" Then
        If wNuméro < selCompteMin Or wNuméro > selCompteMax Then Exit Function
    Else
        If wCompteGénéral < selCompteMin Or wCompteGénéral > selCompteMax Then Exit Function
    End If
End If
If Not blnCompteBilan Then
     If wBilan = "B" Then Exit Function
 End If

 If Not blnCompteHorsBilan Then
     If wBilan = "H" Then Exit Function
 End If
 
 If blnDevise Then
     If CV_X1.DeviseIso <> selDeviseN Then Exit Function
 End If
 
 If blnCompteBanque Then
     If wNuméro > "30000000000" Then Exit Function
 End If
  
 If blnCompteClient Then
     If wNuméro < "30000000000" Then Exit Function
 End If
 
 If blnComptePP Then
     If wNuméro < "30000000000" Then Exit Function
     X2 = mId$(Msg, 249, 2)
     If X2 <> "01" And X2 <> "02" Then Exit Function
 End If
 
 If blnComptePM Then
     If wNuméro < "30000000000" Then Exit Function
     X2 = mId$(Msg, 249, 2)
     If X2 = "01" Or X2 = "02" Then Exit Function
 End If
 
 If blnNostro Then
     If mId$(Msg, 270, 1) <> "N" Then Exit Function
 End If
 
 If blnNostroNo Then
     If mId$(Msg, 270, 1) = "N" Then Exit Function
 End If


If blnService Then
     If mId$(Msg, 282, 3) <> selService Then Exit Function
 End If
 
 If blnGestionnaire Then
     If mId$(Msg, 117, 2) <> selGestionnaire Then Exit Function
 End If

 If mId$(Msg, 115, 1) = "A" Then
     If blnBiaTyp Then
         If mId$(Msg, 241, 3) <> selBiaTyp Then Exit Function
     End If
     If chkPrintLine = "1" Or blnCompteMinMax Then
        If Not ctlGestionnaire_New(mId$(Msg, 13, 11), mId$(Msg, 117, 2), mId$(Msg, 241, 3)) Then Exit Function
    End If
End If

'If optEtatSortK = "A4" Then
If blnPays Then
    wL = Val(mId$(wNuméro, 1, 5))
    If recRacine.Numéro <> wL Then
        recRacine.Method = "SeekL0"
        recRacine.Numéro = wL
        Racine_Load
        ''''If Not IsNull(srvRacineMon(recRacine)) Then Call MsgBox("Erreur lecture racine", , "frmCompteSolde : cmdImport_Select")
        recdictio.Method = "Seek=       "
        recdictio.DicRub = "19"
        recdictio.DicCode = recRacine.RésidentPays
        If IsNull(dbDictioRead(recdictio)) Then wPAys = mId$(recdictio.DicTxt, 7, 2) & "  "
    End If
    If blnPays Then
     If mId$(recRacine.RésidentPays, 2, 3) <> selPays Then Exit Function
    End If
End If

Select Case optEtatSortK
    Case "A1": X = wNuméro & CV_X1.DeviseIso & wCompteGénéral
    Case "A2": X = mId$(wNuméro, 1, 5) & mId$(wNuméro, 9, 2) & "0" & CV_X1.DeviseIso & mId$(wNuméro, 6, 3) & wCompteGénéral
    Case "A3": X = CV_X1.DeviseIso & wNuméro & wCompteGénéral
    Case "A4": X = wPAys & mId$(wNuméro, 6, 3) & mId$(wNuméro, 1, 5) & mId$(wNuméro, 9, 2) & "0" & CV_X1.DeviseIso
    Case "A5": X = mId$(wNuméro, 6, 3) & mId$(wNuméro, 1, 5) & mId$(wNuméro, 9, 2) & CV_X1.DeviseIso
    Case "A6": X = mId$(wNuméro, 1, 5)
    Case "G1": X = wCompteGénéral & CV_X1.DeviseIso & wNuméro & CV_X1.DeviseIso
    Case "G2": X = wCompteGénéral & wNuméro & CV_X1.DeviseIso
    Case "G3": X = CV_X1.DeviseIso & wCompteGénéral & wNuméro
    Case "G4": X = wPAys & wCompteGénéral & wNuméro & CV_X1.DeviseIso
End Select
cmdImport_Select = X & wBilan
End Function

Private Sub fraCompteSolde_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraEtat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraPrésentationEtat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSolde_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

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


Private Sub lstDevise_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case currentActiveControl_Name
    Case "txtDevise": txtDevise = mId$(lstDevise.Text, 1, 3): If blnSetfocus Then txtDevise.SetFocus
    Case "txtDeviseCV": txtDeviseCV = mId$(lstDevise.Text, 1, 3): If txtDeviseCV.Enabled Then txtDeviseCV.SetFocus
End Select

End Sub


Private Sub optEtatBOTCBanque_Click()
optEtat_Script

End Sub

Private Sub optEtatBOTCPM_Click()
optEtat_Script
End Sub


Private Sub optEtatCaisse_Click()
optEtat_Script
End Sub

Private Sub optEtatCptAux_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatCptAux
End Sub


Private Sub optEtatCptGen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatCptGen
End Sub


Private Sub optEtatCrédocPays_Click()
optEtat_Script
End Sub

Private Sub optEtatCrédocPays_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatCrédocPays

End Sub

Private Sub optEtatCrédocTdC_Click()
optEtat_Script
End Sub

Private Sub optEtatDafi_Click()
optEtat_Script
End Sub

Private Sub optEtatDafi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatDafi
End Sub


Private Sub optEtatDafiPays_Click()
optEtat_Script

End Sub

Private Sub optEtatDafiPays_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatDafiPays

End Sub


Private Sub optEtatDG_Click()
optEtat_Script
End Sub

Private Sub optEtatDGA_Click()
optEtat_Script
End Sub


Private Sub optEtatFOTCBanque_Click()
optEtat_Script
End Sub

Private Sub optEtatFOTCPM_Click()
optEtat_Script
End Sub

Private Sub optEtatInspection_Click()
optEtat_Script
End Sub

Private Sub optEtatInspection019_Click()
optEtat_Script
End Sub


Private Sub optEtatManuel_Click()
optEtat_Script
End Sub

Private Sub optEtatManuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatManuel
End Sub


Private Sub optEtatTCCAD_Click()
optEtat_Script
End Sub

Private Sub optEtatTCCAD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatTCCAD
End Sub


Private Sub optEtatTCDeviseIn_Click()
optEtat_Script
End Sub

Private Sub optEtatTCDeviseIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatTCDeviseIn
End Sub


Private Sub optEtatTCGBP_Click()
optEtat_Script
End Sub

Private Sub optEtatTCGBP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatTCGBP
End Sub


Private Sub optEtatTCUSD_Click()
optEtat_Script
End Sub

Private Sub optEtatTCUSD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatTCUSD
End Sub


Private Sub optFinDAnnée_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optFinDAnnée
End Sub


Private Sub optFinDeMois_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optFinDeMois
End Sub


Private Sub optFinDeMoisOpération_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optFinDeMoisOpération
End Sub


Private Sub optSort1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort1
End Sub


Private Sub optSort2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort2
End Sub


Private Sub optSort3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort3
End Sub


Private Sub optSort4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort4
End Sub


Private Sub optSort5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort5
End Sub


Private Sub optSort6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort6
End Sub


Private Sub optVeille_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optVeille
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

If optEtatCptAux Then
    optSort1.Caption = "Racine / Type / N° ordre / Devise"
    optSort2.Caption = "Racine / N° ordre / Devise / Type "
    optSort3.Caption = "Devise / Racine / Type / N° ordre "
    optSort4.Caption = "Pays Résidence / Type / Racine / Devise"
    optSort5.Caption = "Type / Racine / Devise"
    chkPrintRupture.Caption = "Ligne total racine / Type"
    chkPrintRuptureRacine.Caption = "Ligne total racine ": chkPrintRuptureRacine.Enabled = True
Else
    optSort1.Caption = "PCI / Devise / Compte"
    optSort2.Caption = "PCI / Compte  / Devise "
    optSort3.Caption = "Devise / PCI / Compte  "
    optSort4.Caption = "Pays Résidence / PCI / Compte / Devise"
    chkPrintRupture.Caption = "Ligne total PCI"
    chkPrintRuptureRacine.Enabled = False
End If


End Sub

Private Sub txtBiaTyp_GotFocus()
txt_GotFocus txtBiaTyp

End Sub


Private Sub txtBiaTyp_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtBiaTyp_LostFocus()
txt_LostFocus txtBiaTyp

End Sub


Private Sub txtCompteMax_GotFocus()
txt_GotFocus txtCompteMax
End Sub


Private Sub txtCompteMax_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtCompteMax_LostFocus()
txt_LostFocus txtCompteMax
End Sub


Private Sub txtCompteMin_GotFocus()
txt_GotFocus txtCompteMin
End Sub


Private Sub txtCompteMin_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtCompteMin_LostFocus()
txt_LostFocus txtCompteMin
End Sub


Private Sub txtDestinataire_GotFocus()
txt_GotFocus txtDestinataire
End Sub


Private Sub txtDestinataire_LostFocus()
txt_LostFocus txtDestinataire
End Sub


Private Sub txtDevise_GotFocus()
txt_GotFocus txtDevise
lstDevise.Visible = True

End Sub


Private Sub txtDevise_KeyPress(KeyAscii As Integer)
'Call num_KeyAscii(KeyAscii)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtDevise_LostFocus()
txt_LostFocus txtDevise
lstDevise.Visible = False
End Sub


Private Sub txtDeviseCV_GotFocus()
txt_GotFocus txtDeviseCV
lstDevise.Visible = True

End Sub


Private Sub txtDeviseCV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtDeviseCV_LostFocus()
txt_LostFocus txtDeviseCV
lstDevise.Visible = False

End Sub

Private Sub txtEnTete_GotFocus()
txt_GotFocus txtEnTete
End Sub


Private Sub txtEnTete_LostFocus()
txt_LostFocus txtEnTete
End Sub


Private Sub txtGestionnaire_GotFocus()
txt_GotFocus txtGestionnaire
End Sub


Private Sub txtGestionnaire_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtGestionnaire_LostFocus()
txt_LostFocus txtGestionnaire
End Sub


Private Sub txtPays_GotFocus()
txt_GotFocus txtPays

End Sub


Private Sub txtPays_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtPays_LostFocus()
txt_LostFocus txtPays

End Sub

Private Sub txtService_GotFocus()
txt_GotFocus txtService
End Sub


Private Sub txtService_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub

Private Sub txtService_LostFocus()
txt_LostFocus txtService
End Sub



Public Sub cmdReset()
recRacineInit recRacine
cmdExport.Visible = CompteSoldeAut.Comptabiliser
SSTab1.Enabled = CompteSoldeAut.Saisir
'SSTab1.Tabs = 0
blnExport = False
optVeille.Value = True

chkCompteBilan.Value = "1"
chkCompteHorsBilan.Value = "1"
chkCompteMinMax.Value = "0": txtCompteMin = "": txtCompteMax = ""
txtCompteMin.Visible = False:: txtCompteMax.Visible = False
chkCompteBanque.Value = "0": chkCompteClient.Value = "0": chkComptePP.Value = "0": chkComptePM.Value = "0"
chkNostroNo.Value = "0": chkNostro.Value = "0"
lstDevise.Visible = False
Call LstDictio(889, lstDevise)
chkDeviseIn.Value = "0"
chkDevise = "0": txtDevise = "": txtDevise.Visible = False
chkGestionnaire.Value = "0": txtGestionnaire = "": txtGestionnaire.Visible = False
chkService.Value = "0": txtService = "": txtService.Visible = False
chkBiaTyp.Value = "0": txtBiaTyp = "": txtBiaTyp.Visible = False
chkPays.Value = "0": txtPays = "": txtPays.Visible = False

chkPrintLine.Value = "1"
chkPrintSoldé.Value = "0"
chkPrintReliure.Value = "0"
chkPrintRuptureRacine.Value = "0"
txtEnTete = "Etat des soldes"
txtDestinataire = ""
txtExport_Filename = paramComptaSld_Cpt_Export
CV_X1 = CV_Euro
X1000 = ""
If CompteSoldeAut.Comptabiliser Then
    optEtatCptGen.Value = True
    optSort2.Value = True
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
Else
    optEtatCptAux.Value = True
    optSort1.Value = True
    chkPrintRupture.Value = "0"
    chkPrintTotal.Value = "0"

End If
If Not blnService_Enabled Then
    txtService = usrService: txtService.Visible = True
    chkService.Value = "1"
    chkService.Enabled = False: txtService.Enabled = False
End If

End Sub

Public Sub cmdPrint_Call(IdKey As String, Msg As String)
If arrCompteNb > 0 Then
    Call lstErr_Clear(lstErr, cmdPrint, "Impression : " & IdKey & " (" & arrCompteNb & ")")

    Mid$(Msg, 1, 12) = Format$(1, "000000") & Format$(arrCompteNb, "000000")
    prtCompteSolde_Print Msg
End If

End Sub

Public Sub optEtat_Script()
cmdReset
If optEtatCrédocBanque Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkBiaTyp.Value = "1": txtBiaTyp = "001"
    chkCompteBanque.Value = "1"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "0"
    chkPrintTotal.Value = "0"
    txtEnTete = "Etat des soldes des comptes ordinaires 'Banque' "
    txtDestinataire = "Service des crédits documentaires"
    Exit Sub
End If
If optEtatCrédocPays Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort4.Value = True
    If blnService_Enabled Then chkService.Value = "1": txtService = "010"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes CREDOC (Pays / type de compte / racine) "
    txtDestinataire = "Service des crédits documentaires"
    Exit Sub
End If

If optEtatCrédocTdC Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort5.Value = True
    If blnService_Enabled Then chkService.Value = "1": txtService = "010"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "0"
    txtEnTete = "Etat des soldes des comptes CREDOC (type de compte / racine) "
    txtDestinataire = "Service des crédits documentaires"
    Exit Sub
End If

If optEtatCompte_010_3 Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort1.Value = True
    If blnService_Enabled Then chkService.Value = "1": txtService = "010"
    chkPrintLine.Value = "0"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat récapitulatif des soldes des comptes CREDOC "
    txtDestinataire = "Service des crédits documentaires"
    Exit Sub
End If

If optEtatDafi Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort2.Value = True
    If blnService_Enabled Then chkService.Value = "1": txtService = "041"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "0"
    chkPrintTotal.Value = "0"
    txtEnTete = "Etat des soldes des comptes DAFI (Racine / N° d'ordre) "
    txtDestinataire = "DAFI"
    Exit Sub
End If

If optEtatDafiPays Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort4.Value = True
    If blnService_Enabled Then chkService.Value = "1": txtService = "041"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes DAFI (Pays / type de compte / racine) "
    txtDestinataire = "DAFI"
    Exit Sub
End If


If optEtatDG Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort6.Value = True
    chkPrintLine.Value = "0"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtDestinataire = "Mr Shallouf"
    If optEtatNostro Then
        chkNostro = "1"
        txtEnTete = "Etat des soldes des comptes NOSTROS"
    Else
        chkNostroNo = "1"
        txtEnTete = "Etat des soldes"
    End If
    Exit Sub
End If

If optEtatDGA Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort6.Value = True
    chkPrintLine.Value = "0"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtDestinataire = "Mr Younsi"
    If optEtatNostro Then
        chkNostro = "1"
        txtEnTete = "Etat des soldes des comptes NOSTROS"
    Else
        chkNostroNo = "1"
        txtEnTete = "Etat des soldes"
    End If
    Exit Sub
End If

If optEtatInspection Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort6.Value = True
    chkPrintLine.Value = "0"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes auxilaires "
    txtDestinataire = "Mr Neffati"
    Exit Sub
End If

If optEtatInspection019 Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    chkBiaTyp.Value = "1": txtBiaTyp = "019"
    optSort1.Value = True
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes 019 "
    txtDestinataire = "Mr Neffati"
    Exit Sub
End If

If optEtatBOTCBanque Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkCompteBanque.Value = "1"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes 'Banque' "
    txtDestinataire = "B O T C"
    Exit Sub
End If

If optEtatBOTCPM Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkComptePM.Value = "1"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "0"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes 'Personnes morales' "
    txtDestinataire = "B O T C"
    Exit Sub
End If


If optEtatFOTCBanque Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkBiaTyp.Value = "1": txtBiaTyp = "001"
    chkCompteBanque.Value = "1"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes ordinaires 'Banque' "
    txtDestinataire = "F O T C"
    Exit Sub
End If

If optEtatFOTCPM Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkBiaTyp.Value = "1": txtBiaTyp = "001"
    chkComptePM.Value = "1"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "0"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes ordinaires 'Personnes morales' "
    txtDestinataire = "F O T C"
    Exit Sub
End If

If optEtatCaisse Then
    optEtatCptGen.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkCompteMinMax.Value = "1": txtCompteMin = "10100000":: txtCompteMax = "10199999"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "0"
    txtEnTete = "Etat des soldes des comptes de la caisse 'espèces' "
    txtDestinataire = "Caisse"
    Exit Sub
End If

If optEtatTCDeviseIn Then
    optEtatCptGen.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkDeviseIn.Value = "1"
    chkPrintLine.Value = "0"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes en Devises In + Euro "
    txtDestinataire = "T C"
    Exit Sub
End If


If optEtatTCCAD Then
    optEtatCptGen.Value = True
    optVeille.Value = True
    optSort3.Value = True
    chkDevise.Value = "1": txtDevise = "CAD"
    chkPrintLine.Value = "0"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes CAD "
    txtDestinataire = "T C"
    Exit Sub
End If

If optEtatTCGBP Then
    optEtatCptGen.Value = True
    optVeille.Value = True
    optSort3.Value = True
    chkDevise.Value = "1": txtDevise = "GBP"
    chkPrintLine.Value = "0"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes GBP "
    txtDestinataire = "T C"
    Exit Sub
End If

If optEtatTCUSD Then
    optEtatCptGen.Value = True
    optVeille.Value = True
    optSort3.Value = True
    chkDevise.Value = "1": txtDevise = "USD"
    chkPrintLine.Value = "0"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes USD "
    txtDestinataire = "T C"
    Exit Sub
End If

If optEtatPersonnel Then
    optEtatCptAux.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkCompteMinMax.Value = "1": txtCompteMin = "60000000000":: txtCompteMax = "69999999999"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "0"
    chkPrintRuptureRacine.Value = "1"
    chkPrintTotal.Value = "1"
    txtEnTete = "Etat des soldes des comptes du personnel "
    txtDestinataire = "Comptabilité"
    Exit Sub
End If


If optEtatCompteConversion Then
    optEtatCptGen.Value = True
    optVeille.Value = True
    optSort1.Value = True
    chkCompteMinMax.Value = "1": txtCompteMin = "38212005":: txtCompteMax = "38212005"
    chkPrintLine.Value = "1"
    chkPrintRupture.Value = "1"
    chkPrintTotal.Value = "0"
    txtEnTete = "Etat des soldes des comptes de conversion "
    txtDestinataire = "Comptabilité"
    Exit Sub
End If

If optEtatRisques Then
    Select Case mEtatPass
        Case 0
            optEtatCptAux.Value = True
            optFinDeMoisOpération.Value = True
            optSort1.Value = True
            chkComptePM.Value = "1"
            chkPrintLine.Value = "1"
            chkPrintRupture.Value = "0"
            chkPrintRuptureRacine.Value = "0" '"1"
            chkPrintTotal.Value = "0"
            txtEnTete = "Etat des soldes des comptes 'Personne morale' "
            txtDestinataire = "Risques (CDR)"
            Exit Sub
    
    Case 2
            optEtatCptAux.Value = True
            optFinDeMoisOpération.Value = True
            optSort1.Value = True
            chkCompteBanque.Value = "1"
            chkNostroNo.Value = "1"
            chkPrintLine.Value = "1"
            chkPrintRupture.Value = "0"
            chkPrintRuptureRacine.Value = "0" ' "1"
            chkPrintTotal.Value = "0"
            txtEnTete = "Etat des soldes des comptes 'Banque' "
            txtDestinataire = "Risques (CDR)"
            Exit Sub
    End Select

End If

End Sub

Public Sub cmdExport_Write()
Dim curX As Currency, X16D As String * 16, X16C As String * 16, XSigne As String * 1, Xcompte As String * 11
Dim Dev16D As String * 16, Dev16C As String * 16

Mid$(X1000, 35, memoCptInfoLen) = mId$(reccptp0.Text, 1, memoCptInfoLen)
curX = CptSolde_Get(optSolde, X1000)
If curX <> 0 Then
    CV_X1.Montant = curX
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X1)
    'X1 = IIf(CV_X2.Montant < 0, "-", "+")
    If CV_X2.Montant < 0 Then
        Dev16D = "-" & Format$(Abs(CV_X1.Montant), "000000000000.00")
        Dev16C = "                "
        X16D = "-" & Format$(Abs(CV_X2.Montant), "000000000000.00")
        X16C = "                "
    Else
        Dev16C = "+" & Format$(Abs(CV_X1.Montant), "000000000000.00")
        Dev16D = "                "
        X16C = "+" & Format$(Abs(CV_X2.Montant), "000000000000.00")
        X16D = "                "
    End If
    If mId$(X1000, 34 + 115, 1) = "A" Then
        Xcompte = mId$(X1000, 34 + 13, 11)
   Else
        Xcompte = mId$(X1000, 34 + 16, 5) & "000000"
    End If
    
    X200 = mId$(X1000, 34 + 7, 3) & ";" & mId$(X1000, 34 + 255, 11) & ";" & Xcompte & ";" & Dev16D & ";" & Dev16C & ";" & X16D & ";" & X16C & ";" & mId$(X1000, 34 + 35, 40)
    cmdImport_Select_Nb = cmdImport_Select_Nb + 1
    Print #2, X200
End If
End Sub

Public Function CptSolde_Get(optSolde As String, Msg As String) As Currency
CptSolde_Get = 0
Select Case optSolde
    Case "V": CptSolde_Get = CCur(Val(mId$(Msg, 34 + 300, 19))) + CCur(Val(mId$(Msg, 34 + 319, 19)))
    Case "O": CptSolde_Get = CCur(Val(mId$(Msg, 34 + 338, 19))) + CCur(Val(mId$(Msg, 34 + 357, 19)))
    Case "M": CptSolde_Get = CCur(Val(mId$(Msg, 34 + 376, 19)))
    Case "A": CptSolde_Get = CCur(Val(mId$(Msg, 34 + 414, 19))) + CCur(Val(mId$(Msg, 34 + 395, 19)))
End Select

End Function

Public Sub cmdImport_Racine()
'***** RACINE
Dim cmdImport_Nb  As Long, I As Long
Dim xInput As String, r As Integer

Open paramComptaSld_Rac_Import For Input As #1
recElpBuffer_Init recElpBuffer
recElpBuffer.Id = "Racine"

cmdImport_Nb = 0: I = 0

Do Until EOF(1)
    Line Input #1, xInput
    If mId$(xInput, 1, 3) = "$$$" Then
        I = Val(mId$(xInput, 5, 18))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "me : cmdImport_ElpBuffer :SrvRacine ")
        End If
        Exit Do
    End If
    cmdImport_Nb = cmdImport_Nb + 1
    recElpBuffer.Seq = Val(mId$(xInput, 1, 5))
    recElpBuffer.Method = constAddNew
    r = tableElpBuffer_Read(recElpBuffer)
    If r = 9998 Then
        r = 0
    Else
        recElpBuffer.Method = constUpdate
    End If

    recElpBuffer.Data = xInput
    dbElpBuffer_Update recElpBuffer

   '' If cmdImport_Nb Mod 1000 = 0 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Import : " & cmdImport_Nb): DoEvents
 
Loop

Close

End Sub

Public Sub Racine_Load()
Dim r As Integer
recRacine.RésidentPays = "0000"

recElpBuffer.Method = "Seek="
recElpBuffer.Id = "Racine"
recElpBuffer.Seq = recRacine.Numéro
r = tableElpBuffer_Read(recElpBuffer)

If r = 0 Then
        MsgTxtIndex = 0
        Mid$(MsgTxt, 35, memoRacineLen) = mId$(recElpBuffer.Data, 1, memoRacineLen)
        If Not IsNull(srvRacineGetBuffer(recRacine)) Then Call MsgBox("Erreur lecture racine", , "frmCompteSolde : Racine_Load")

Else
    If Not IsNull(srvRacineMon(recRacine)) Then Call MsgBox("Erreur lecture racine", , "frmCompteSolde : Racine_Load")
End If

End Sub
