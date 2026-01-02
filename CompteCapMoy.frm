VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompteCapMoy 
   Caption         =   "Etat des capitaux moyens"
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
      TabIndex        =   35
      Top             =   0
      Width           =   3585
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
      TabIndex        =   34
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Choix d'un état"
      TabPicture(0)   =   "CompteCapMoy.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBalance"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sélection des comptes"
      TabPicture(1)   =   "CompteCapMoy.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOptions"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Présentation de l'état"
      TabPicture(2)   =   "CompteCapMoy.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPrésentationEtat"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraPrésentationEtat 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   8775
         Begin VB.Frame fraEnTete 
            Height          =   1575
            Left            =   240
            TabIndex        =   28
            Top             =   2160
            Width           =   8295
            Begin VB.TextBox txtEnTete 
               Height          =   285
               Left            =   2280
               TabIndex        =   32
               Text            =   "Etat des capitaux moyens"
               Top             =   480
               Width           =   4935
            End
            Begin VB.TextBox txtDestinataire 
               Height          =   285
               Left            =   2280
               TabIndex        =   29
               Top             =   1080
               Width           =   4935
            End
            Begin VB.Label lblEnTete 
               Caption         =   "En Tête de l'état"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblDestinataire 
               Caption         =   "Destinataire de l'état"
               Height          =   255
               Left            =   240
               TabIndex        =   30
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
            Height          =   1455
            Left            =   4320
            TabIndex        =   22
            Top             =   360
            Width           =   4215
            Begin VB.CheckBox chkPrintLine 
               Caption         =   "Ligne détail des comptes auxilaires"
               Height          =   375
               Left            =   120
               TabIndex        =   24
               Top             =   300
               Value           =   1  'Checked
               Width           =   3255
            End
            Begin VB.CheckBox chkPrintSoldé 
               Caption         =   "Imprimer les comptes soldés"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   720
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
            Height          =   1455
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   3855
            Begin VB.OptionButton optSort2 
               Caption         =   "Racine / N° ordre / Devise / Type "
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   840
               Width           =   3375
            End
            Begin VB.OptionButton optSort1 
               Caption         =   "Racine / Type / N° ordre / Devise"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   360
               Value           =   -1  'True
               Width           =   3255
            End
         End
      End
      Begin VB.Frame fraOptions 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   2
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
            TabIndex        =   3
            Top             =   240
            Width           =   8535
            Begin VB.CheckBox chkCrédoc 
               Caption         =   "Spécial Crédoc"
               Height          =   255
               Left            =   5400
               TabIndex        =   59
               Top             =   4800
               Width           =   2175
            End
            Begin VB.CheckBox chkNostroNo 
               Caption         =   "exclure comptes Nostros"
               Height          =   255
               Left            =   5400
               TabIndex        =   48
               Top             =   4320
               Width           =   2655
            End
            Begin VB.CheckBox chkNostro 
               Caption         =   "uniquement comptes Nostros"
               Height          =   255
               Left            =   5400
               TabIndex        =   47
               Top             =   3960
               Width           =   2655
            End
            Begin VB.CheckBox chkComptePP 
               Caption         =   "uniquement comptes 'P P'"
               Height          =   255
               Left            =   5400
               TabIndex        =   44
               Top             =   3600
               Width           =   2655
            End
            Begin VB.CheckBox chkComptePM 
               Caption         =   "uniquement comptes 'P M'"
               Height          =   255
               Left            =   5400
               TabIndex        =   43
               Top             =   3240
               Width           =   2655
            End
            Begin VB.CheckBox chkCompteClient 
               Caption         =   "uniquement comptes 'Client'"
               Height          =   255
               Left            =   5400
               TabIndex        =   42
               Top             =   2880
               Width           =   2655
            End
            Begin VB.CheckBox chkCompteBanque 
               Caption         =   "uniquement comptes 'Banque'"
               Height          =   255
               Left            =   5400
               TabIndex        =   41
               Top             =   2520
               Width           =   2655
            End
            Begin VB.ListBox lstDevise 
               Height          =   2010
               Left            =   5400
               TabIndex        =   40
               Top             =   240
               Width           =   3015
            End
            Begin VB.CheckBox chkPays 
               Caption         =   "sélectionner le pays (code BdF)"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   4400
               Width           =   2535
            End
            Begin VB.TextBox txtPays 
               Height          =   285
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   38
               Top             =   4400
               Width           =   495
            End
            Begin VB.TextBox txtDeviseCV 
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   36
               Text            =   "EUR"
               Top             =   400
               Width           =   495
            End
            Begin VB.CheckBox chkDeviseIn 
               Caption         =   "Uniquement devises In et Euro"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   800
               Width           =   2775
            End
            Begin VB.TextBox txtBiaTyp 
               Height          =   285
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   17
               Top             =   4000
               Width           =   495
            End
            Begin VB.TextBox txtGestionnaire 
               Height          =   285
               Left            =   3240
               MaxLength       =   2
               TabIndex        =   16
               Top             =   3600
               Width           =   495
            End
            Begin VB.TextBox txtService 
               Height          =   285
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   15
               Top             =   3240
               Width           =   495
            End
            Begin VB.CheckBox chkBiaTyp 
               Caption         =   "sélectionner le type de compte"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   4000
               Width           =   2535
            End
            Begin VB.CheckBox chkGestionnaire 
               Caption         =   "sélectionner le gestionnaire"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   3600
               Width           =   2535
            End
            Begin VB.CheckBox chkService 
               Caption         =   "sélectionner le service gestionnaire"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   3240
               Width           =   2895
            End
            Begin VB.CheckBox chkCompteMinMax 
               Caption         =   "sélectionner les comptes de"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   2400
               Width           =   2535
            End
            Begin VB.CheckBox chkCompteHorsBilan 
               Caption         =   "sélectionner les comptes de hors-bilan"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   2000
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.CheckBox chkCompteBilan 
               Caption         =   "sélectionner les comptes de bilan"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   1600
               Value           =   1  'Checked
               Width           =   2895
            End
            Begin VB.TextBox txtDevise 
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   8
               Top             =   1200
               Width           =   495
            End
            Begin VB.CheckBox chkDevise 
               Caption         =   "Sélectionner la devise"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   1200
               Width           =   2055
            End
            Begin VB.TextBox txtCompteMax 
               Height          =   285
               Left            =   3000
               MaxLength       =   11
               TabIndex        =   5
               Top             =   2800
               Width           =   1575
            End
            Begin VB.TextBox txtCompteMin 
               Height          =   285
               Left            =   3000
               MaxLength       =   11
               TabIndex        =   4
               Top             =   2400
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "devise de contre-valeur"
               Height          =   255
               Left            =   360
               TabIndex        =   37
               Top             =   400
               Width           =   2415
            End
            Begin VB.Label lblMax 
               Caption         =   "à"
               Height          =   255
               Left            =   2280
               TabIndex        =   6
               Top             =   2800
               Width           =   255
            End
         End
      End
      Begin VB.Frame fraBalance 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8775
         Begin VB.Frame fraPériode 
            Height          =   3375
            Left            =   240
            TabIndex        =   49
            Top             =   2040
            Width           =   8295
            Begin VB.CheckBox chkPrint 
               Caption         =   "Impression des capitaux moyens"
               Height          =   375
               Left            =   1920
               TabIndex        =   58
               Top             =   2640
               Value           =   1  'Checked
               Width           =   3495
            End
            Begin VB.CheckBox chkExport 
               Caption         =   "Calcul des capitaux moyens"
               Height          =   375
               Left            =   1920
               TabIndex        =   57
               Top             =   2040
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.CommandButton cmdOk 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ok"
               Height          =   885
               Left            =   6360
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   2160
               Width           =   1455
            End
            Begin VB.TextBox txtExport_Filename 
               Height          =   285
               Left            =   1920
               TabIndex        =   55
               Top             =   1320
               Width           =   5895
            End
            Begin MSComCtl2.DTPicker txtAmjMin 
               Height          =   300
               Left            =   1920
               TabIndex        =   50
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   65404931
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtAmjMax 
               Height          =   300
               Left            =   3840
               TabIndex        =   53
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   65404931
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   2
            End
            Begin VB.Label lblExport_Filename 
               Caption         =   "Fichier d'exportation"
               Height          =   255
               Left            =   240
               TabIndex        =   54
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label lblAmjMax 
               Caption         =   "au"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3360
               TabIndex        =   52
               Top             =   720
               Width           =   315
            End
            Begin VB.Label lblAmjMin 
               Caption         =   "du"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   51
               Top             =   720
               Width           =   315
            End
         End
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
            Height          =   1695
            Left            =   240
            TabIndex        =   45
            Top             =   240
            Width           =   3135
            Begin VB.OptionButton optEtatManuel 
               Caption         =   "Manuel"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Value           =   -1  'True
               Width           =   1095
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
            Height          =   1635
            Left            =   3720
            TabIndex        =   25
            Top             =   360
            Width           =   4815
            Begin VB.OptionButton optEtatFlux 
               Caption         =   "Balance des comptes généraux en flux en CV Euros"
               Height          =   375
               Left            =   120
               TabIndex        =   60
               Top             =   960
               Width           =   4455
            End
            Begin VB.OptionButton optEtatCptGen 
               Caption         =   "Etat des soldes des comptes généraux"
               Height          =   375
               Left            =   120
               TabIndex        =   27
               Top             =   600
               Width           =   4335
            End
            Begin VB.OptionButton optEtatCptAux 
               Caption         =   "Etat des soldes des comptes auxiliaires"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   300
               Value           =   -1  'True
               Width           =   4455
            End
         End
      End
   End
End
Attribute VB_Name = "frmCompteCapMoy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' à faire : filtre sélection des comptes
' compte sans mvt
' enlever bricolage prtCompteCapMoy_Monitor
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean, blnSetfocus As Boolean
Dim CompteCapMoyAut As typeAuthorization
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

Dim blnExport As Boolean, X137 As String * 137
Dim X1000 As String * 1000
Dim cmdImport_Select_Nb As Long, cmdImport_Nb As Long

Dim blnService_Enabled As Boolean
Dim wL As Long, wPAys As String * 4, wX As String
Dim recdictio As typeDictio

Dim wAmjMinTxt As String * 8
Dim wAmjMin As String * 8, wAmjMax As String * 8, wAmj As String * 8
Dim xAmjMin As String, xAmjMax As String, xAMJ As String
Dim vReturn As Variant
Dim arrMtCV() As Currency, arrMt() As Currency, arrMt_Index As Long, arrMt_NbMax As Long
Dim mDbMt As Currency, mDbNb As Long
Dim mCrMt As Currency, mCrNb As Long
Dim mID14 As String * 14, wMt As Currency, wMtCV As Currency
Dim blnSoldeInitial As Boolean

Dim mIntitulé As String

Dim wDB1 As Currency, wCR1 As Currency, wDB2 As Currency, wCR2 As Currency, wVR4 As Currency
Dim sSD1 As Currency, sCR As Currency, sDB As Currency, sSD2 As Currency
Dim tSD1 As Currency, tCR As Currency, tDB As Currency, tSD2 As Currency
Dim sDev As String * 3, tCompte As String * 11, tIntitulé As String

Private Sub cmdImport_MvtP0()
Dim xInput As String, blnOk As Boolean, blnMvtAdd As Boolean, blnCptOk As Boolean
Dim wCodeOpération As String
Dim mRupture As String

On Error Resume Next

Dim I As Long
blnOk = False: blnCptOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
vReturn = DTPicker_Control(txtAmjMin, wAmjMinTxt)
If Not IsNull(vReturn) Then Call lstErr_AddItem(lstErr, txtAmjMin, vReturn): Exit Sub
vReturn = DTPicker_Control(txtAmjMax, wAmjMax)
If Not IsNull(vReturn) Then Call lstErr_AddItem(lstErr, txtAmjMax, vReturn): Exit Sub

wAmjMin = dateElp("Jour", -1, wAmjMinTxt)

X = Dir(paramCompteCapMoy_Mvt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdOk, "? Le fichier des mouvements n'existe pas"): Exit Sub
xAmjMin = dateImp(wAmjMin)
xAmjMax = dateImp(wAmjMax)
arrMt_NbMax = DateDiff("d", xAmjMin, xAmjMax)
ReDim arrMt(arrMt_NbMax)
ReDim arrMtCV(arrMt_NbMax)

cmdImport_MvtP0_Reset
mID14 = ""

Call lstErr_AddItem(lstErr, cmdOk, "Chargement des mouvements ...")
CV_X2 = CV_Euro
CV_X3 = CV_Euro
CV_X1.OpéAmj = DSys: CV_X1.CoursCompta = "C"
CV_X2.OpéAmj = DSys: CV_X2.CoursCompta = "C"
CV_X3.OpéAmj = DSys: CV_X3.CoursCompta = "C"


Open paramCompteCapMoy_Mvt_Import For Input As #1
Open paramCompteCapMoy_Cpt_Export For Output As #2
tableCptP0_Open
I = 0

Do Until EOF(1)
    Line Input #1, xInput
      
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        ''SrvMvtP0_Amj = mId$(xInput, 86, 8)
        I = Val(mId$(xInput, 94, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteCapMoy : cmdImport_Cptp0 :SrvMvtP0 ")
        End If
        Exit Do
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    I = I + 1
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdOk, "Sélection des comptes : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
       
    If mID14 <> mId$(xInput, 7, 14) Then
         If blnCptOk Then cmdImport_Mvtp0_Export
         cmdImport_MvtP0_Reset
         mID14 = mId$(xInput, 7, 14)
         reccptp0.Id = mID14
         reccptp0.Method = "Seek="
         If tableCptP0_Read(reccptp0) = 0 Then
            blnCptOk = True
            mIntitulé = mId$(reccptp0.Text, 35, 40)
            CV_X1.Montant = CCur(Val(mId$(reccptp0.Text, 119, 19)))
            CV_X1.OpéAmj = DSys
            CV_X2.OpéAmj = DSys
            CV_X3.OpéAmj = DSys
            CV_X1.DeviseN = mId$(xInput, 7, 3): CV_X1.DeviseIso = ""
            Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X1)
            arrMtCV(0) = arrMtCV(0) + CV_X2.Montant
            arrMt(0) = arrMt(0) + CV_X1.Montant
            blnSoldeInitial = True
            Call lstErr_ChangeLastItem(lstErr, cmdContext, mID14 & " Mvt : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
         Else
            blnCptOk = False
         End If
    End If
    
    If blnCptOk Then
    
        wAmj = mId$(xInput, 63, 8)
        
        If wAmj > wAmjMin Then
        
            xAMJ = dateImp(wAmj)
            CV_X1.Montant = CCur(Val(mId$(xInput, 28, 19)))
            CV_X1.OpéAmj = wAmj
            CV_X2.OpéAmj = wAmj
            CV_X3.OpéAmj = wAmj
            Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X1)
            arrMtCV(0) = arrMtCV(0) - CV_X2.Montant
            arrMt(0) = arrMt(0) - CV_X1.Montant

            If wAmj <= wAmjMax Then

                cmdImport_Select_Nb = cmdImport_Select_Nb + 1
                arrMt_Index = DateDiff("d", xAmjMin, xAMJ)
                arrMtCV(arrMt_Index) = arrMtCV(arrMt_Index) + CV_X2.Montant
                arrMt(arrMt_Index) = arrMt(arrMt_Index) + CV_X1.Montant
   
                blnMvtAdd = True
                If mId$(mID14, 9, 3) = "001" Then
                    wCodeOpération = mId$(xInput, 21, 4)
                    If wCodeOpération = "G051" Or wCodeOpération = "G052" Then blnMvtAdd = False
                End If
                
                If blnMvtAdd Then
                    If chkCrédoc = "1" Then
                        cmdImport_MvtP0_Crédoc mId$(xInput, 86, 50)
                    Else
                        cmdImport_MvtP0_Cumul
                   End If
                End If
            End If
        End If
    End If
       
 
Loop

If blnCptOk Then cmdImport_Mvtp0_Export

Close
tableCptP0_Close

If Not blnOk Then
'    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteCapMoy : cmdImport_Mptp0 :SrvMvtP0 ")
End If

End Sub




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

SrvCptP0_Amj = "00000000"
blnService_Enabled = True
Call BiaPgmAut_Init("Compte_CapMo", CompteCapMoyAut)

If Not IsNull(param_Init) Then cmdOk.Visible = False
cmdReset

blnSetfocus = True
End Sub


Public Function param_Init()
Dim V
param_Init = Null
recElpTable_Init recElpTable
recElpTable.Id = "Param"
recElpTable.K1 = "CompteCapMoy"
recElpTable.Method = "Seek="

recElpTable.K2 = "Cpt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramCompteCapMoy_Cpt_Import = paramServer(recElpTable.Memo)
'''Call lstErr_Clear(lstErr, cmdContext, "Fichier :" & paramCompteCapMoy_Cpt_Import)

recElpTable.K2 = "Cpt_Export"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramCompteCapMoy_Cpt_Export = Trim(recElpTable.Memo)


recElpTable.K2 = "Mvt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramCompteCapMoy_Mvt_Import = paramServer(recElpTable.Memo)

Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "CompteCapMoy_Param_Init"
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


Private Sub chkPrintSoldé_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPrintSoldé
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


Private Sub cmdCompteCapMoy()
If Trim(txtExport_Filename) = "" Then
    Call lstErr_Clear(lstErr, cmdOk, "? préciser le nom du fichier de calcul ")
Else
    blnExport = True
    cmdImport_CptP0
    cmdImport_MvtP0
    blnExport = False
    Call lstErr_AddItem(lstErr, cmdOk, "Calcul terminé : " & cmdImport_Select_Nb)
    frmCompteCapMoy.Enabled = True
    AppActivate frmCompteCapMoy.Caption
End If
End Sub

Private Sub cmdImport_CptP0()
Dim xInput As String, blnOk As Boolean
Dim vReturn As Variant
On Error Resume Next

Dim I As Long
blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
X = Dir(paramCompteCapMoy_Cpt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdOk, "? Le fichier des comptes n'existe pas"): Exit Sub

Call lstErr_AddItem(lstErr, cmdOk, "Chargement des comptes, tri ...")
Me.MousePointer = vbHourglass
Me.Enabled = False

MDB.Execute "delete * from CptP0"
mdbCptP0.tableCptP0_Open

Open paramCompteCapMoy_Cpt_Import For Input As #1
recCptP0_Init reccptp0
reccptp0.Method = "AddNew"


Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        SrvCptP0_Amj = mId$(xInput, 35, 8)
        I = Val(mId$(xInput, 43, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteCapMoy : cmdImport_Cptp0 :SrvCptP0 ")
        End If
        Exit Do
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    vReturn = cmdImport_CptP0_Select(xInput)
    If vReturn <> "" Then
        reccptp0.Id = vReturn
        reccptp0.Text = xInput
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            dbCptP0_Update reccptp0
    End If
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdOk, "Sélection des comptes : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
 
Loop

Close
mdbCptP0.tableCptP0_Close
Me.MousePointer = 0
If Not blnOk Then
    'cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteCapMoy : cmdImport_Cptp0 :SrvCptP0 ")
End If

End Sub


Private Sub cmdPrint()
Dim X, Nb As Integer, curX As Currency, IdKey As String, mIdKey As String
Dim Msg As String

'cmdImport_CptP0

'If cmdImport_Select_Nb = 0 Then
'    Call lstErr_AddItem(lstErr, cmdok , "Aucun compte sélectionné !")
'    GoTo cmdPrint_End
'End If

Msg = "000000000000" & Space$(50)
Mid$(Msg, 14, 3) = selDeviseCV
Mid$(Msg, 17, 1) = "B"
Mid$(Msg, 18, 1) = optSolde
Mid$(Msg, 19, 8) = optAmj
Mid$(Msg, 27, 1) = IIf(chkPrintSoldé = "1", "S", " ")
'Mid$(Msg, 28, 1) = IIf(chkPrintReliure = "1", ">", "=")
Mid$(Msg, 29, 1) = IIf(chkPrintLine = "1", "L", "-")
'Mid$(Msg, 30, 1) = IIf(chkPrintRupture = "1", "R", "-")
'Mid$(Msg, 31, 1) = IIf(chkPrintTotal = "1", "T", "-")
Mid$(Msg, 32, 2) = optEtatSortK
'Mid$(Msg, 34, 1) = IIf(chkPrintRuptureRacine = "1", "R", "-")

prtCompteCapMoy_Monitor Msg
Call lstErr_AddItem(lstErr, cmdOk, "Impression terminée ")

cmdPrint_End:
frmCompteCapMoy.Enabled = True
AppActivate frmCompteCapMoy.Caption

End Sub



Public Sub cmdControl()
lstErr.Clear
optEtat = "A"
If optEtatCptGen Then optEtat = "G"
If optEtatFlux Then
    optEtat = "F"
    chkExport.Caption = "Exclure les écritures 'virement à résultat'"
    chkPrint.Enabled = False
Else
    chkExport.Caption = "Calcul des capitaux moyens"
    chkPrint.Enabled = True
End If

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

optEtatSortK = optEtat & optSortK
Select Case optEtatSortK
    Case "A1": PrintRupture_Len = 5
    Case "A2": PrintRupture_Len = 5
    Case "G1": PrintRupture_Len = 11
    Case "G2": PrintRupture_Len = 11

End Select

mDestinataire = Trim(txtDestinataire)
mEnTete = Trim(txtEnTete)
paramCompteCapMoy_Cpt_Export = Trim(txtExport_Filename)

End Sub
Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOk_Click()
Call lstErr_Clear(lstErr, cmdOk, "Début du traitement")
cmdControl
If lstErr.ListCount <> 0 Then Exit Sub
If optEtatFlux Then
    cmdFlux
Else

    If chkExport Then cmdCompteCapMoy
    If chkPrint Then cmdPrint
End If

End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Form_Init
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

Private Function cmdImport_CptP0_Select(Msg As String) As String
Dim wCompteGénéral As String * 11, wNuméro As String * 11, wDeviseN As String * 3, wBilan As String * 1
Dim X2 As String * 2

cmdImport_CptP0_Select = ""

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
    If optEtat = "G" Then
        If wCompteGénéral < selCompteMin Or wCompteGénéral > selCompteMax Then Exit Function
    Else
        If wNuméro < selCompteMin Or wNuméro > selCompteMax Then Exit Function
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
        If Not IsNull(srvRacineMon(recRacine)) Then Call MsgBox("Erreur lecture racine", , "frmCompteCapMoy : cmdImport_Select")
        recdictio.Method = "Seek=       "
        recdictio.DicRub = "19"
        recdictio.DicCode = recRacine.RésidentPays
        If IsNull(dbDictioRead(recdictio)) Then wPAys = mId$(recdictio.DicTxt, 7, 2) & "  "
    End If
    If blnPays Then
     If mId$(recRacine.RésidentPays, 2, 3) <> selPays Then Exit Function
    End If
End If

cmdImport_CptP0_Select = mId$(Msg, 7, 3) & mId$(Msg, 13, 11)
End Function

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


Private Sub optEtatCptAux_Click()
cmdControl
End Sub

Private Sub optEtatCptAux_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatCptAux
End Sub


Private Sub optEtatCptGen_Click()
cmdControl
End Sub

Private Sub optEtatCptGen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatCptGen
End Sub


Private Sub optEtatFlux_Click()
cmdControl
End Sub

Private Sub optEtatFlux_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatFlux

End Sub


Private Sub optEtatManuel_Click()
optEtat_Script
End Sub

Private Sub optEtatManuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatManuel
End Sub


Private Sub optSort1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort1
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)

If optEtatCptAux Then
    optSort1.Caption = "Racine / Type / N° ordre / Devise"
    optSort2.Caption = "Type / Racine "
Else
    optSort1.Caption = "PCI / Devise / Compte"
    optSort2.Caption = "PCI / Compte  / Devise "
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
SSTab1.Enabled = True 'CompteCapMoyAut.Saisir
'SSTab1.Tabs = 0
blnExport = False

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
txtEnTete = "Etat des soldes"
txtDestinataire = ""
txtExport_Filename = paramCompteCapMoy_Cpt_Export
CV_X1 = CV_Euro
X1000 = ""
    optEtatCptAux.Value = True
    optSort1.Value = True

If Not blnService_Enabled Then
    txtService = usrService: txtService.Visible = True
    chkService.Value = "1"
    chkService.Enabled = False: txtService.Enabled = False
End If
recCptP0_Init reccptp0
End Sub

Public Sub cmdPrint_Call(IdKey As String, Msg As String)
If arrCompteNb > 0 Then
    Call lstErr_Clear(lstErr, cmdOk, "Impression : " & IdKey & " (" & arrCompteNb & ")")

    Mid$(Msg, 1, 12) = Format$(1, "000000") & Format$(arrCompteNb, "000000")
   ' prtCompteCapMoy_Print Msg
End If

End Sub

Public Sub optEtat_Script()
cmdReset

End Sub

Public Sub cmdImport_MvtP0_Reset()
Dim I As Integer
For I = 0 To arrMt_NbMax: arrMt(I) = 0: arrMtCV(I) = 0: Next I
mDbMt = 0: mDbNb = 0
mCrMt = 0: mCrNb = 0
blnSoldeInitial = False
End Sub

Public Sub cmdImport_Mvtp0_Export()
Dim X250 As String * 250, wDec, wDecCV
Dim blnSdZero As Boolean, wNbj As Long

blnSdZero = True: wNbj = 0
wDec = CDec(0): wDecCV = CDec(0)

wMt = arrMt(0)
wMtCV = arrMtCV(0)

'If mID14 = "00160248001012" Then
'    wMt = 0
'End If

wNbj = arrMt_NbMax

For arrMt_Index = 1 To arrMt_NbMax
    wMtCV = wMtCV + arrMtCV(arrMt_Index)
    wMt = wMt + arrMt(arrMt_Index)
'    If blnSdZero Then
'        If wMt <> 0 Then
'            blnSdZero = False
'            If arrMt_Index > 0 Then wNbj = 1
'        End If
'    Else
'        wNbj = wNbj + 1
'    End If
    
    wDec = wDec + wMt
    wDecCV = wDecCV + wMtCV
    
Next arrMt_Index

X250 = ""
Mid$(X250, 1, 3) = mId$(mID14, 1, 3)
Mid$(X250, 4, 1) = ";"
Mid$(X250, 5, 11) = mId$(mID14, 4, 11)
Mid$(X250, 16, 1) = ";"
Mid$(X250, 17, 10) = Format$(wAmjMinTxt, "####/##/##")
Mid$(X250, 27, 1) = ";"
Mid$(X250, 28, 10) = Format$(wAmjMax, "####/##/##")
Mid$(X250, 38, 1) = ";"
Mid$(X250, 39, 10) = Format$(wNbj, "0000000000")
Mid$(X250, 49, 1) = ";"
Mid$(X250, 50, 1) = IIf(wDec < 0, "-", "+")
Mid$(X250, 51, 29) = Format$(Abs(wDecCV), "00000000000000000000000000.00")
Mid$(X250, 80, 1) = ";"
Mid$(X250, 81, 1) = IIf(wDec < 0, "-", "+")
Mid$(X250, 82, 29) = Format$(Abs(wDec), "00000000000000000000000000.00")
Mid$(X250, 111, 1) = ";"

Mid$(X250, 112, 10) = Format$(mDbNb, "0000000000")
Mid$(X250, 122, 1) = ";"
Mid$(X250, 123, 19) = Format$(Abs(mDbMt), "0000000000000000.00")
Mid$(X250, 142, 1) = ";"
Mid$(X250, 143, 10) = Format$(mCrNb, "0000000000")
Mid$(X250, 153, 1) = ";"
Mid$(X250, 154, 19) = Format$(Abs(mCrMt), "0000000000000000.00")
Mid$(X250, 173, 1) = ";"
Mid$(X250, 174, 40) = mIntitulé

Print #2, X250

End Sub

Public Sub Form_Init()

fraEtat.Enabled = False
fraScript.Enabled = False
fraPrésentationEtat.Enabled = False
fraEtat.Enabled = True 'False

wAmj = dateElp("FinDeMoisP", 0, DSys)
Call DTPicker_Set(txtAmjMax, wAmj)
Mid$(wAmj, 7, 2) = "01"
Call DTPicker_Set(txtAmjMin, wAmj)

End Sub

Public Sub cmdImport_MvtP0_Cumul()
If CV_X2.Montant < 0 Then
    mDbNb = mDbNb + 1
    mDbMt = mDbMt + CV_X2.Montant 'CV
Else
    mCrNb = mCrNb + 1
    mCrMt = mCrMt + CV_X2.Montant 'CV
End If

End Sub

Public Sub cmdImport_MvtP0_Crédoc(lX As String)
Dim iDos As Long

If mId$(lX, 1, 6) = "CDE000" And mId$(lX, 12, 6) = "ADE001" Then
    iDos = mId$(lX, 7, 5) '
'     If iDos >= 61335 And iDos <= 63071 Then
     If iDos > 63071 Then
        '''''If Abs(CV_X2.Montant) >= 150000 Then
        cmdImport_MvtP0_Cumul
    End If
End If

End Sub

Public Sub cmdFlux()
Call lstErr_Clear(lstErr, cmdOk, "Flux : Début du traitement")
vReturn = DTPicker_Control(txtAmjMin, wAmjMin)
If Not IsNull(vReturn) Then Call lstErr_AddItem(lstErr, txtAmjMin, vReturn): Exit Sub
vReturn = DTPicker_Control(txtAmjMax, wAmjMax)
If Not IsNull(vReturn) Then Call lstErr_AddItem(lstErr, txtAmjMax, vReturn): Exit Sub

X = Dir(paramCompteCapMoy_Mvt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdOk, "? Le fichier des mouvements n'existe pas"): Exit Sub
xAmjMin = dateImp(wAmjMin)
xAmjMax = dateImp(wAmjMax)

Me.MousePointer = vbHourglass
Me.Enabled = False

cmdFlux_Mvt

cmdFlux_Cpt


cmdFlux_Export

Me.MousePointer = 0
Me.Enabled = True



End Sub

Public Sub cmdFlux_Mvt()
Dim xInput As String, blnOk As Boolean, blnMvtAdd As Boolean, blnCptOk As Boolean
Dim wCodeOpération As String
Dim mRupture As String

On Error Resume Next

Dim I As Long
Call lstErr_AddItem(lstErr, cmdOk, "Chargement des mouvements ...")

blnOk = False: blnCptOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
mID14 = ""

MDB.Execute "delete * from MvtP0"
mdbMvtP0.tableMvtP0_Open

Open paramCompteCapMoy_Mvt_Import For Input As #1
recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"


Open paramCompteCapMoy_Mvt_Import For Input As #1
I = 0
wDB1 = 0: wCR1 = 0: wDB2 = 0: wCR2 = 0: wVR4 = 0

Do Until EOF(1)
    Line Input #1, xInput
      
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        ''SrvMvtP0_Amj = mId$(xInput, 86, 8)
        I = Val(mId$(xInput, 94, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteCapMoy : cmdflux_Cptp0 :SrvMvtP0 ")
        End If
        Exit Do
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    I = I + 1
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdOk, "Sélection des mouvements : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
       
    If mID14 <> mId$(xInput, 7, 14) Then
         If blnCptOk Then cmdFlux_Mvt_Update
        
        mID14 = mId$(xInput, 7, 14)
        recMvtp0.Id = mID14
        blnCptOk = True
        wDB1 = 0: wCR1 = 0: wDB2 = 0: wCR2 = 0: wVR4 = 0

    End If

        wAmj = mId$(xInput, 55, 8)   ' date opération
        blnOk = True
        
        '  exclure une pièce comptable anormale du 14.02.2000 2552/0000077 (cf 22.02.200 2634/0000017 )
        If wAmj = "20000214" Then
            If mId$(xInput, 82, 4) = "2552" And mId$(xInput, 71, 7) = "0000077" Then blnOk = False
        End If
         
        '  exclure les annulations de virement à résultat du 31.12.2001 cours de réévaluation erroné (Tréso du 31.12 au lieu compta du 30.12)
        If wAmj = "20011231" Then
            If mId$(xInput, 172, 10) = "CPA052    " Or mId$(xInput, 172, 10) = "REP-CPA052" Then blnOk = False
        End If

'$JPL        If mID14 = "97800060121002" Then
'$JPL            blnOk = True
'$JPL        Else
'$JPL            blnOk = False
'$JPL       End If
        
        
        If blnOk And wAmj >= wAmjMin Then
            wMt = CCur(Val(mId$(xInput, 28, 19)))
              

            If wAmj <= wAmjMax Then
                cmdImport_Select_Nb = cmdImport_Select_Nb + 1
                
        '  exclure les écritures automatiques Virement à résultat (JJCPLT = 4)
               If chkExport = "1" And mId$(xInput, 143, 1) = "4" Then
                    wVR4 = wVR4 + wMt
                Else

                    If wMt < 0 Then
                        wDB1 = wDB1 + wMt
                    Else
                        wCR1 = wCR1 + wMt
                    End If
                End If
            Else
                 If wMt < 0 Then
                    wDB2 = wDB2 + wMt
                Else
                    wCR2 = wCR2 + wMt
                End If
           
        End If
    End If
       
 
Loop

If blnCptOk Then cmdFlux_Mvt_Update


Close
tableMvtP0_Close

If Not blnOk Then
'    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteCapMoy : cmdflux_Mptp0 :SrvMvtP0 ")
End If

Call lstErr_AddItem(lstErr, cmdOk, "fin des mouvements  : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents

End Sub

Public Sub cmdFlux_Cpt()
Dim xInput As String, blnOk As Boolean
Dim vReturn As Variant
On Error Resume Next

Dim I As Long
blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
X = Dir(paramCompteCapMoy_Cpt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdOk, "? Le fichier des comptes n'existe pas"): Exit Sub

Call lstErr_AddItem(lstErr, cmdOk, "Chargement des comptes, tri ..."): DoEvents

MDB.Execute "delete * from CptP0"
mdbCptP0.tableCptP0_Open

mdbMvtP0.tableMvtP0_Open
recMvtP0_Init recMvtp0
recMvtp0.Method = "Seek="

Open paramCompteCapMoy_Cpt_Import For Input As #1
recCptP0_Init reccptp0
reccptp0.Method = "AddNew"

I = 0

Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        SrvCptP0_Amj = mId$(xInput, 35, 8)
        I = Val(mId$(xInput, 43, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteCapMoy : cmdImport_Cptp0 :SrvCptP0 ")
        End If
        Exit Do
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    I = I + 1
    vReturn = cmdImport_CptP0_Select(xInput)
    If vReturn <> "" Then
        reccptp0.Text = xInput
        reccptp0.Id = mId$(reccptp0.Text, 255, 11) & mId$(reccptp0.Text, 7, 3) & mId$(reccptp0.Text, 13, 11) ' CPTGEN DEVISE COMPTE
        recMvtp0.Id = mId$(reccptp0.Text, 7, 3) & mId$(reccptp0.Text, 13, 11)
        If tableMvtP0_Read(recMvtp0) = 0 Then
            Mid$(reccptp0.Text, 300, 95) = mId$(recMvtp0.Text, 1, 95)   ' $$$$$$$$
        Else
            Mid$(reccptp0.Text, 300, 95) = "-000000000000000.00+000000000000000.00-000000000000000.00+000000000000000.00-000000000000000.00"   ' $$$$$$$$
        End If
        
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            dbCptP0_Update reccptp0
    End If
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdOk, "Sélection des comptes : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
 
Loop

Close
mdbCptP0.tableCptP0_Close
mdbMvtP0.tableMvtP0_Close

If Not blnOk Then
    'cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteCapMoy : cmdImport_Cptp0 :SrvCptP0 ")
End If
Call lstErr_ChangeLastItem(lstErr, cmdOk, "Fin des comptes : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
End Sub

Public Sub cmdFlux_Mvt_Update()
Dim X95 As String * 95
Mid$(X95, 1, 19) = cur_19P(wDB1)
Mid$(X95, 20, 19) = cur_19P(wCR1)
Mid$(X95, 39, 19) = cur_19P(wDB2)
Mid$(X95, 58, 19) = cur_19P(wCR2)
Mid$(X95, 77, 19) = cur_19P(wVR4)
recMvtp0.Text = X95
mdbMvtP0.dbMvtP0_Update recMvtp0

End Sub

Public Sub cmdFlux_Export()
Dim X250 As String * 250, wDec, wDecCV
Dim blnSdZero As Boolean, wNbj As Long
Dim wSD As Currency
Call lstErr_AddItem(lstErr, cmdOk, "Export : début"): DoEvents

Open paramCompteCapMoy_Cpt_Export For Output As #2

CV_X2 = CV_Euro
CV_X3 = CV_Euro
CV_X1.OpéAmj = wAmjMax: CV_X1.CoursCompta = "C"
CV_X2.OpéAmj = wAmjMax: CV_X2.CoursCompta = "C"
CV_X3.OpéAmj = wAmjMax: CV_X3.CoursCompta = "C"

blnSdZero = True: wNbj = 0
wDec = CDec(0): wDecCV = CDec(0)
sCR = 0: sDB = 0: sSD1 = 0: sSD2 = 0
tCR = 0: tDB = 0: tSD1 = 0: tSD2 = 0

sDev = "000"
tCompte = "00000000000"

mdbCptP0.tableCptP0_Open
recCptP0_Init reccptp0
reccptp0.Method = "MoveFirst"

Do
    intReturn = tableCptP0_Read(reccptp0)

    If intReturn = 0 Then

        If tCompte <> mId$(reccptp0.Text, 255, 11) Then
            cmdFlux_Export_S
            cmdFlux_Export_T
            tCompte = mId$(reccptp0.Text, 255, 11)
            tIntitulé = mId$(reccptp0.Text, 35, 40)
            sDev = mId$(reccptp0.Text, 7, 3)
        End If
        If sDev <> mId$(reccptp0.Text, 7, 3) Then
            cmdFlux_Export_S
            sDev = mId$(reccptp0.Text, 7, 3)
        End If
              
        wMt = CCur(Val(mId$(reccptp0.Text, 119, 19)))
        wDB1 = CCur(Val(mId$(reccptp0.Text, 300, 19)))
        wCR1 = CCur(Val(mId$(reccptp0.Text, 319, 19)))
        wDB2 = CCur(Val(mId$(reccptp0.Text, 338, 19)))
        wCR2 = CCur(Val(mId$(reccptp0.Text, 357, 19)))
        wVR4 = CCur(Val(mId$(reccptp0.Text, 376, 19)))
        
        
        wSD = wMt - wDB1 - wCR1 - wDB2 - wCR2 - wVR4
        sSD1 = sSD1 + wSD
        sCR = sCR + wCR1
        sDB = sDB + wDB1
        wSD = wMt - wDB2 - wCR2 - wVR4
        sSD2 = sSD2 + wSD

   End If
    reccptp0.Method = "MoveNext"
Loop While intReturn = 0

cmdFlux_Export_S
cmdFlux_Export_T

Close
mdbCptP0.tableCptP0_Close
mdbMvtP0.tableMvtP0_Close
Call lstErr_AddItem(lstErr, cmdOk, "Export : fin"): DoEvents
End Sub

Public Sub cmdFlux_Export_T()
Dim X250 As String * 250

If tSD1 <> 0 Or tSD2 <> 0 Or tCR <> 0 Or tDB <> 0 Then
    X250 = ""
    Mid$(X250, 1, 11) = Format(tCompte, "00000000000")
    Mid$(X250, 12, 1) = ";"
    Mid$(X250, 13, 40) = tIntitulé
    Mid$(X250, 53, 1) = ";"
           
    Mid$(X250, 54, 19) = cur_19V(tSD1)
    Mid$(X250, 73, 1) = ";"
    Mid$(X250, 75, 19) = cur_19V(tDB)
    Mid$(X250, 94, 1) = ";"
    Mid$(X250, 95, 19) = cur_19V(tCR)
    Mid$(X250, 114, 1) = ";"
    Mid$(X250, 115, 19) = cur_19V(tSD2)
    Print #2, X250
End If

tSD1 = 0: tSD2 = 0: tCR = 0: tDB = 0

End Sub

Public Sub cmdFlux_Export_S()
If sDev <> "978" Then
    CV_X1.DeviseN = sDev: CV_X1.DeviseIso = ""
    CV_X1.Montant = sSD1
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X1)
    sSD1 = CV_X2.Montant
    
    CV_X1.Montant = sDB
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X1)
    sDB = CV_X2.Montant
    
    CV_X1.Montant = sCR
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X1)
    sCR = CV_X2.Montant
    
    CV_X1.Montant = sSD2
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X1)
    sSD2 = CV_X2.Montant
End If

tSD1 = tSD1 + sSD1: sSD1 = 0
tSD2 = tSD2 + sSD2: sSD2 = 0
tCR = tCR + sCR: sCR = 0
tDB = tDB + sDB: sDB = 0


End Sub
