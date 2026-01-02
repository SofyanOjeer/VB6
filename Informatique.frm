VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInformatique 
   Caption         =   "Informatique"
   ClientHeight    =   6390
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   10380
   Icon            =   "Informatique.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10380
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   0
      Width           =   3825
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Divers"
      TabPicture(0)   =   "Informatique.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Commission Bancaire"
      TabPicture(1)   =   "Informatique.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblCBFile"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdCB_ClientExport"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtCBFile"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lstCB"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdCB_CompteNR"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdCB_COBIA"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Euro"
      TabPicture(2)   =   "Informatique.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCompteAnnulation"
      Tab(2).Control(1)=   "txtEuroBascule"
      Tab(2).Control(2)=   "cmdEuroBascule"
      Tab(2).Control(3)=   "cmdEuroEXport"
      Tab(2).Control(4)=   "cmdEuroPrint"
      Tab(2).Control(5)=   "cmdEuroImport"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "FICDAT NUMERO"
      TabPicture(3)   =   "Informatique.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdMailing"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame4"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.CommandButton cmdMailing 
         Caption         =   "Mailing Libye"
         Height          =   735
         Left            =   -68040
         TabIndex        =   42
         Top             =   840
         Width           =   2415
      End
      Begin VB.Frame Frame5 
         Caption         =   "NUMERO"
         Height          =   1575
         Left            =   -74640
         TabIndex        =   36
         Top             =   2880
         Width           =   5535
         Begin VB.CommandButton cmdNumeroAdd 
            Caption         =   "Add +1"
            Height          =   495
            Left            =   3960
            TabIndex        =   40
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdNumeroRead 
            Caption         =   "Read AS400"
            Height          =   435
            Left            =   3960
            TabIndex        =   39
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtNumero 
            Height          =   375
            Left            =   1320
            TabIndex        =   38
            Top             =   360
            Width           =   735
         End
         Begin VB.Label libNumero 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   2280
            TabIndex        =   41
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblNumero 
            Caption         =   "Compteur"
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "FICDAT"
         Height          =   2175
         Left            =   -74640
         TabIndex        =   31
         Top             =   600
         Width           =   5655
         Begin VB.CommandButton cmddateBIa 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Test dateBIA JPL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1200
            Width           =   2400
         End
         Begin VB.CommandButton cmdFicDatP1Import 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Import D:\Temp\FicDatP1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   480
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker txtAmjMin 
            Height          =   300
            Left            =   2880
            TabIndex        =   34
            Top             =   1440
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
            Format          =   19595267
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtAmjMax 
            Height          =   300
            Left            =   4320
            TabIndex        =   35
            Top             =   1440
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
            Format          =   19595267
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin VB.CommandButton cmdCompteAnnulation 
         BackColor       =   &H000000FF&
         Caption         =   "Annulation des comptes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -67800
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3240
         Width           =   2520
      End
      Begin VB.TextBox txtEuroBascule 
         Height          =   375
         Left            =   -71160
         TabIndex        =   29
         Text            =   "D:\Temp\Euro_Test.txt"
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CommandButton cmdEuroBascule 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Bascule des comptes IN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74400
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3240
         Width           =   3000
      End
      Begin VB.CommandButton cmdEuroEXport 
         Caption         =   "Export comptes en dev IN"
         Height          =   735
         Left            =   -67680
         TabIndex        =   27
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton cmdEuroPrint 
         Caption         =   "impression contrôle du fichier"
         Height          =   975
         Left            =   -74520
         TabIndex        =   26
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CommandButton cmdEuroImport 
         Caption         =   "import comptes en dev IN ( Serv= ""000"")"
         Height          =   735
         Left            =   -74520
         TabIndex        =   25
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdCB_COBIA 
         Caption         =   "Client C/O BIA"
         Height          =   855
         Left            =   480
         TabIndex        =   24
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdCB_CompteNR 
         Caption         =   "Comptes NR 001 Date Der Mvt"
         Height          =   855
         Left            =   360
         TabIndex        =   23
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ListBox lstCB 
         Height          =   5130
         Left            =   5640
         TabIndex        =   22
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtCBFile 
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Text            =   "D:\Temp\Audit_Client_Courrier_"
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton cmdCB_ClientExport 
         Caption         =   "Export Client : 1 compte ordinaire (001)  toutes devises ***"
         Height          =   1095
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4815
         Left            =   -74640
         TabIndex        =   4
         Top             =   840
         Width           =   9615
         Begin VB.CommandButton cmdExportCptMvt 
            Caption         =   "Export_CptMvt"
            Height          =   495
            Left            =   600
            TabIndex        =   17
            Top             =   3720
            Width           =   1695
         End
         Begin VB.ListBox lst 
            Height          =   3180
            Left            =   4320
            TabIndex        =   16
            Top             =   1440
            Width           =   4815
         End
         Begin VB.Frame Frame2 
            Height          =   855
            Left            =   4200
            TabIndex        =   11
            Top             =   360
            Width           =   4935
            Begin VB.TextBox txtEnd 
               Height          =   285
               Left            =   3000
               TabIndex        =   13
               Text            =   "99999999"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtStart 
               Height          =   285
               Left            =   960
               TabIndex        =   12
               Text            =   "0"
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblEnd 
               Caption         =   "à"
               Height          =   255
               Left            =   2400
               TabIndex        =   15
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lblStart 
               Caption         =   "de"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            Height          =   2655
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   3615
            Begin VB.OptionButton optAdresse 
               Caption         =   "Liste des adresses"
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   1200
               Value           =   -1  'True
               Width           =   2295
            End
            Begin VB.OptionButton optLr97Pcec 
               Caption         =   "LR97 : Chargement Pcec"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   2280
               Width           =   2295
            End
            Begin VB.OptionButton optLR97PCI 
               Caption         =   "LR97 : préparation PCI"
               Height          =   255
               Left            =   240
               TabIndex        =   8
               Top             =   1800
               Width           =   2295
            End
            Begin VB.OptionButton optCptGen 
               Caption         =   "Liste des Comptes Généraux"
               Height          =   255
               Left            =   240
               TabIndex        =   7
               Top             =   720
               Width           =   3015
            End
            Begin VB.OptionButton optRacine 
               Caption         =   "Liste des racines"
               Height          =   255
               Left            =   240
               TabIndex        =   6
               Top             =   240
               Width           =   2295
            End
         End
      End
      Begin VB.Label lblCBFile 
         Caption         =   "fichier d'exportation"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.TextBox txtRecherche 
      Height          =   300
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H000000FF&
      Caption         =   "&Recherche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   9960
      Picture         =   "Informatique.frx":04B2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   400
   End
End
Attribute VB_Name = "frmInformatique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim recCptInfo As typeCptInfo
Dim recdictio As typeDictio

Dim recRacine As typeRacine
Dim blnPrint As Boolean

Dim FileNumber As Integer
Dim RecLength As Long
Public Position As Long

Dim cmdImport_Select_Nb As Long, cmdImport_Nb As Long

Dim E0(3) As Long, E1(3) As Long, E1A(3) As Long, EBHB(3) As Long, EX(3) As Long
Dim C0(3) As Long, C1(3) As Long, CA(3) As Long, CB(3) As Long, CHB(3) As Long

Dim wSituation As String * 1, wBiaTyp As String * 3, wNatureTitulaire As String * 2

Dim kCB As Integer
Dim meCV1 As typeCV, meCV2 As typeCV, meCV3 As typeCV

Dim paramEuroBascule As String, meCompte As typeCompte, meCptinfo As typeCptInfo
Dim meNumero As typeNumeroP0

Public Sub optRacine_Sql()
Dim I As Integer, Msg As String
Dim x54 As String * 54

ReDim arrRacine(1): arrRacineNbMax = 1

recRacineInit recRacine
recRacine.Method = "SnapL0"
recRacine.Numéro = CLng(Val(txtStart))
arrRacine(0) = recRacine
arrRacine(0).Numéro = CLng(Val(txtEnd))

lst.Clear
arrRacineSuite = True

Open "C:\Temp\RacineP" For Output As #2

Do Until Not arrRacineSuite
    arrRacineNb = 0: arrRacineIndex = 0
    
    srvRacineMon recRacine
    recRacine = arrRacine(arrRacineNb)
    recRacine.Method = "SnapL0+"
    
    lst.Clear
    For I = 1 To arrRacineNb
        lst.AddItem arrRacine(I).Numéro & Chr$(9) & arrRacine(I).Intitulé
        x54 = arrRacine(I).Numéro & " " & arrRacine(I).Intitulé & " " & mId$(arrRacine(I).RésidentPays, 2, 3) & " " & mId$(arrRacine(I).Nationalité, 2, 3)
        Print #2, x54
    Next I
    'If blnPrint And arrRacineNb > 0 Then
    '     Msg = Format$(1, "000000") & Format$(arrRacineNb, "000000")
    '     prtRacineX Msg
    ' End If
Loop



Close

End Sub


Public Sub Msg_Rcv(txtMsg As String)
'---------------------------------------------------------
param_Init
meCV1 = CV_Euro
meCV1.CoursCompta = "C"
meCV1.OpéAmj = DSys
meCV1.Normal = "P"
meCV1.AchatVente = " "
meCV2 = meCV1: meCV3 = meCV1

End Sub



Private Sub cmdCB_ClientExport_Click()
Dim I As Integer
lstCB.Clear

cmdCB_ClientExport_Load
cmdCB_ClientExport_Write
cmdCB_ClientExport_Ex
For I = 1 To 3
    C0(0) = C0(0) + C0(I)
    C1(0) = C1(0) + C1(I)
    CB(0) = CB(0) + CB(I)
    CHB(0) = CHB(0) + CHB(I)
    CA(0) = CA(0) + CA(I)
    
    E0(0) = E0(0) + E0(I)
    E1(0) = E1(0) + E1(I)
    E1A(0) = E1A(0) + E1A(I)
    EBHB(0) = EBHB(0) + EBHB(I)
    EX(0) = EX(0) + EX(I)
Next I

lstCB.AddItem "Entité     : " & Chr$(9) & E0(1) & Chr$(9) & E0(2) & Chr$(9) & E0(3) & Chr$(9) & E0(0)
lstCB.AddItem "Entité 001 : " & Chr$(9) & E1(1) & Chr$(9) & E1(2) & Chr$(9) & E1(3) & Chr$(9) & E1(0)
lstCB.AddItem "Entité 001A: " & Chr$(9) & E1A(1) & Chr$(9) & E1A(2) & Chr$(9) & E1A(3) & Chr$(9) & E1A(0)
lstCB.AddItem "Entité B_HB: " & Chr$(9) & EBHB(1) & Chr$(9) & EBHB(2) & Chr$(9) & EBHB(3) & Chr$(9) & EBHB(0)
lstCB.AddItem "Entité ??? : " & Chr$(9) & EX(1) & Chr$(9) & EX(2) & Chr$(9) & EX(3) & Chr$(9) & EX(0)

lstCB.AddItem "  "
lstCB.AddItem "Compte     : " & Chr$(9) & C0(1) & Chr$(9) & C0(2) & Chr$(9) & C0(3) & Chr$(9) & C0(0)
lstCB.AddItem "Compte 001: " & Chr$(9) & C1(1) & Chr$(9) & C1(2) & Chr$(9) & C1(3) & Chr$(9) & C1(0)
lstCB.AddItem "Compte B  : " & Chr$(9) & CB(1) & Chr$(9) & CB(2) & Chr$(9) & CB(3) & Chr$(9) & CB(0)
lstCB.AddItem "Compte HB : " & Chr$(9) & CHB(1) & Chr$(9) & CHB(2) & Chr$(9) & CHB(3) & Chr$(9) & CHB(0)
lstCB.AddItem "Compte Ann: " & Chr$(9) & CA(1) & Chr$(9) & CA(2) & Chr$(9) & CA(3) & Chr$(9) & CA(0)

End Sub

Private Sub cmdCB_COBIA_Click()
cmdCB_ClientExport_Load
cmdCB_COBIA_Write

End Sub


Private Sub cmdCB_CompteNR_Click()
cmdCB_CompteNR_Load
End Sub

Private Sub cmdCompteAnnulation_Click()
paramEuroBascule = Trim(txtEuroBascule)
cmdCompteAnnulation_DB
End Sub

Private Sub cmdContext_Click()
If SSTab1.Tab = 0 Then
    blnPrint = False
    If optRacine Then optRacine_Sql
    If optCptGen Then optCptGen_Sql
    If optAdresse Then optAdresse_Sql
    If optLR97PCI Then LR97Pci_Sql
    If optLr97Pcec Then LR97Pcec_Write
'optCptGen_Sql
End If

End Sub

Private Sub cmddateBIa_Click()
Dim X8 As String
Call DTPicker_Control(txtAmjMin, X8)
'X8 = dateBIA("Jour", -5, X8)
X8 = dateBIA("Ouvré", 0, X8)
Call DTPicker_Set(txtAmjMax, X8)
End Sub

Private Sub cmdEuroBascule_Click()
paramEuroBascule = Trim(txtEuroBascule)
cmdEuroBascule_DB

End Sub

Private Sub cmdEuroEXport_Click()
paramEuroBascule = Trim(txtEuroBascule)
cmdEuroExport_List "cmdEuroEXport"
End Sub

Private Sub cmdEuroImport_Click()
cmdEuroImport_Load
End Sub

Private Sub cmdEuroPrint_Click()
paramEuroBascule = Trim(txtEuroBascule)
cmdEuroPrint_List "cmdEuroPrint"
End Sub

Private Sub cmdExportCptMvt_Click()
Dim curX As Currency, X16D As String * 16, X16C As String * 16, X147 As String * 147, xSens As String * 1
Dim I As Integer, X3 As String, Nb As Integer
Dim recCompte As typeCompte
Dim X As String
Dim xInput As String
CV_X2 = CV_Euro
CV_X1.OpéAmj = DSys: CV_X1.CoursCompta = "C"
CV_X2.OpéAmj = DSys: CV_X2.CoursCompta = "C"
CV_X3.OpéAmj = DSys: CV_X3.CoursCompta = "C"
''Call CV_AttributS(XDevise, CV_X2)

Open "C:\BiaSrv\CptMvt" For Output As #2

recCompteInit recCompte

'Open "S:\FTP\SrvCptP0" For Input As #1
Call lstErr_Clear(lstErr, cmdPrint, "Chargement des comptes, ...")
Nb = 0
'Do Until EOF(1)
 '   Line Input #1, xInput
'    If mId$(xInput, 1, 3) = "$$$" Then
'            Exit Do
'    End If
    Nb = Nb + 1
'    MsgTxtIndex = 0
'    MsgTxt = Space$(recCptInfoLen)
'    Mid$(MsgTxt, 35, memoCptInfoLen) = mId$(xInput, 1, memoCptInfoLen)
'    If IsNull(srvCompteGetBuffer(recCompte)) Then
'    If Nb = 1000 Then Nb = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Compte : " & recCompte.Numéro): DoEvents
    recCompte.Société = "001"
    recCompte.Agence = "001"
    recCompte.Devise = "001"
    recCompte.Numéro = "00038150009"

'        If recCompte.TypeGA = "A" And mId$(xInput, 282, 3) = "041" Then
'            X3 = mId$(recCompte.Numéro, 6, 3)
'           If X3 = "901" Or X3 = "902" Or X3 = "943" Then
              Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Mouvement: " & recCompte.Numéro): DoEvents
              arrCptMvtSuite = True: arrCptMvtNb = 0
                Call arrCptMvt_Load(recCompte, "20010701", DSys)
                For I = 1 To arrCptMvtNb
 '                   CV_X1.DeviseN = Format$(arrCptMvt(I).Devise, "000")
  '                  CV_X1.DeviseIso = ""
   '                 CV_X1.Montant = arrCptMvt(I).Mt
    '                Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
                    If arrCptMvt(I).MT < 0 Then
                        xSens = "D"
                        X16D = "0" & Format$(Abs(arrCptMvt(I).MT), "000000000000.00")
                        X16C = "                "
  '$$$ cv euro dans la zone crédit
    '                    X16C = "0" & Format$(Abs(CV_X2.Montant), "000000000000.00")
                    Else
                        xSens = "C"
                        X16C = "0" & Format$(Abs(arrCptMvt(I).MT), "000000000000.00")
                        X16D = "                "
                    End If
                    X147 = arrCptMvt(I).Devise & " " & arrCptMvt(I).Compte & " " _
                        & arrCptMvt(I).AmjTraitement & " " & X16D & X16C & " " & xSens & " " & arrCptMvt(I).Libellé
                    Print #2, X147
                    
                   ' End If
                    
                Next I
           ' End If
       ' End If
   ' End If
'Loop

Close

End Sub

Private Sub cmdFicDatP1Import_Click()
Dim X As String, Nb As Long
X = "C:\Temp\FICDATP1"

Call dbFicDatP1_Import(X, Nb)

Call lstErr_Clear(lstErr, cmdPrint, "Chargement FICDATP1 : " & Nb)

End Sub

Private Sub cmdMailing_Click()
Dim xInput As String, blnOk As Boolean, wAmjCréation As String * 8, wAmjAnnulation As String * 8
Dim vReturn As Variant, X As String, SrvCptP0_Amj As String * 8

On Error GoTo Error_Handler


Call lstErr_Clear(lstErr, cmdPrint, "Mailing ....."): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False

mdbCptP0.tableCptP0_Open

Open "C:\Temp\SOBF_Compte_002.txt" For Input As #1
Open "C:\Temp\SOBF_Adresse_002.txt" For Output As #2

recCptP0_Init reccptp0
reccptp0.Method = "Seek="


Do Until EOF(1)
    Line Input #1, xInput
    
    If Trim(xInput) <> "" Then
            reccptp0.Id = "001001978" & mId$(Trim(xInput), 1, 11)
            reccptp0.Method = "Seek="
            If tableCptP0_Read(reccptp0) = 0 Then
            
                MsgTxtIndex = 0
                MsgTxt = Space$(recCptInfoLen)
                Mid$(MsgTxt, 35, memoCptInfoLen) = mId$(reccptp0.Text, 1, memoCptInfoLen)
                Call srvCptInfoGetBuffer(recCptInfo)
                X = recCptInfo.Numéro & ";" & Trim(recCptInfo.Intitulé) & ";" & Trim(recCptInfo.Adresse2) & ";" & Trim(recCptInfo.Adresse3) _
                    & ";" & Trim(recCptInfo.Adresse4) & ";" & Trim(recCptInfo.Adresse5) & ";" & Trim(recCptInfo.AdresseCP) & " " & Trim(recCptInfo.AdresseBD) & ";" & Trim(recCptInfo.AdressePays)
                Print #2, X

            Else
                MsgBox "? : " & xInput, vbCritical, "Mailing"
                X = recCptInfo.Numéro & ";" & ";" & ";" _
                    & ";" & ";" & ";" & " " & ";"
                Print #2, X
        
            End If
            
        End If
    
 
Loop


mdbCptP0.tableCptP0_Close
Close
Me.MousePointer = 0

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdCB_ClientExport_Load")
Me.Enabled = True

End Sub

Private Sub cmdNumeroAdd_Click()
recNumeroP0_Init meNumero
meNumero.NOCTR = CInt(Val(txtNumero))
meNumero.Method = "Add"
Call srvNumeroP0_Update(meNumero)
libNumero = meNumero.CTEUR

End Sub

Private Sub cmdNumeroRead_Click()
recNumeroP0_Init meNumero
meNumero.NOCTR = CInt(Val(txtNumero))
meNumero.Method = "SeekP0"
Call srvNumeroP0_Monitor(meNumero)
libNumero = meNumero.CTEUR
End Sub

Private Sub cmdPrint_Click()
blnPrint = True
If optRacine Then optRacine_Sql
If optCptGen Then optCptGen_Sql
If optLR97PCI Then LR97Pci_Sql

End Sub



Public Sub optCptGen_Sql()
Dim I As Integer, Msg As String

ReDim arrCptInfo(1): arrCptInfoNbMax = 1

prtCompteAttribut.mNuméro = ""

recCptInfoInit recCptInfo
recCptInfo.Method = "SnapKG"

recCptInfo.Société = SocId$
recCptInfo.Agence = SocAgence$
recCptInfo.Devise = Format$(0, "000")
recCptInfo.CompteGénéral = Format$(CLng(Val(txtEnd)), "00000000000")
recCptInfo.Numéro = Format$(0, "00000000000")
recCptInfo.BiaTyp = Format$(0, "000")
recCptInfo.BiaNum = Format$(0, "00")

arrCptInfo(0) = recCptInfo
recCptInfo.CompteGénéral = Format$(CLng(Val(txtStart)), "00000000000")

prtCompteAttribut_Open

lst.Clear
arrCptInfosuite = True
Do Until Not arrCptInfosuite
    arrCptInfoNb = 0: arrCptInfoIndex = 0
    
    srvCptInfoMon recCptInfo
    recCptInfo = arrCptInfo(arrCptInfoNb)
    recCptInfo.Method = "SnapKG+"
    
    lst.Clear
    For I = 1 To arrCptInfoNb
        lst.AddItem arrCptInfo(I).Numéro & Chr$(9) & arrCptInfo(I).Intitulé
    Next I
    If blnPrint And arrCptInfoNb > 0 Then
         Msg = Format$(1, "000000") & Format$(arrCptInfoNb, "000000")
         prtCompteAttributX Msg
     End If
Loop

prtCompteAttributTrait

frmElpPrt.prtEndDoc
frmElpPrt.Hide


End Sub

Public Sub LR97Pci_Sql()
Dim I As Integer, X As String, Nb As Integer, strNull As String
strNull = ""
ReDim arrCptInfo(1): arrCptInfoNbMax = 1

prtCompteAttribut.mNuméro = ""

recCptInfoInit recCptInfo
recCptInfo.Method = "SnapPCI"

recCptInfo.Société = SocId$
recCptInfo.Agence = SocAgence$
recCptInfo.Devise = Format$(0, "000")
recCptInfo.CompteGénéral = Format$(CLng(Val(txtEnd)), "00000000000")
recCptInfo.Numéro = Format$(0, "00000000000")
recCptInfo.BiaTyp = Format$(0, "000")
recCptInfo.BiaNum = Format$(0, "00")

arrCptInfo(0) = recCptInfo
recCptInfo.CompteGénéral = Format$(CLng(Val(txtStart)), "00000000000")
recDictioInit recdictio

lst.Clear
If blnPrint Then
    prtCptGenLst_Open
Else
    FileNumber = FreeFile
    Open "c:\Bialr97\Bia_Pci.txt" For Output As FileNumber
End If

Nb = 0
arrCptInfosuite = True
Do Until Not arrCptInfosuite
    arrCptInfoNb = 0: arrCptInfoIndex = 0
    
    srvCptInfoMon recCptInfo
    recCptInfo = arrCptInfo(arrCptInfoNb)
    recCptInfo.Method = "SnapPCI+"
    
    lst.Clear
    For arrCptInfoIndex = 1 To arrCptInfoNb
        lst.AddItem arrCptInfo(arrCptInfoIndex).Numéro & Chr$(9) & arrCptInfo(arrCptInfoIndex).Intitulé
        If blnPrint Then
            prtCptGenLst_Line
        Else
            LR97Pci_Write
        End If
        Nb = Nb + 1
    Next arrCptInfoIndex
Loop
If blnPrint Then
    frmElpPrt.prtEndDoc
    frmElpPrt.Hide
Else
    X = Format$(Nb, "0000")
    Write #FileNumber, "000", "origine:BIA ", "version:BAFI", Format$(Nb * 2 + 2, "0000"), "000", "1", "500", "1", "530", X, "590", X
    Write #FileNumber, "500", "00002", "", "0C", "BIAJPL", strNull, strNull
    Close #FileNumber
 End If
End Sub

Public Function LR97Pci_Write()
Dim X As String, I As Integer, strSens As String, strPcec As String

On Error GoTo Error_Handler

strSens = "F"
strPcec = "09999"

For I = 5 To 2 Step -1
    X = mId$(arrCptInfo(arrCptInfoIndex).Numéro, 4, I)

    recdictio.Method = "Seek=       "
    recdictio.DicRub = "892"
    recdictio.DicCode = Trim(X)
    dbDictioRead recdictio
    If recdictio.Err = 0 Then
        strPcec = Trim(X)
        strSens = Trim(recdictio.DicLib)
        Exit For
    End If
Next I

Write #FileNumber, "530", "8", mId$(arrCptInfo(arrCptInfoIndex).Numéro, 4, 8), "0C", Trim(arrCptInfo(arrCptInfoIndex).Intitulé), "O", "", ""
Write #FileNumber, "590", mId$(arrCptInfo(arrCptInfoIndex).Numéro, 4, 8), strPcec, "", "", "", "O", "N", strSens, "O"
LR97Pci_Write = Null

GoTo End_Function

Error_Handler:
     MsgBox "erreur " & Error
   LR97Pci_Write = Err
End_Function:
End Function




Public Sub LR97Pcec_Write()
Dim FileNumber As Integer, X As String
On Error GoTo Error_Handler
FileNumber = FreeFile
Open "c:\BiaLr97\Lr_Pcec.txt" For Input As FileNumber
    lst.Clear
recDictioInit recdictio
Do Until EOF(FileNumber)
    Input #FileNumber, X
    lst.AddItem Trim(X)
    recdictio.Method = constAddNew
    recdictio.DicRub = "892"
    recdictio.DicCode = mId$(X, 2, 5)
    recdictio.DicLib = mId$(X, 1, 1)
    recdictio.DicAmj = DSys
    Call dbDictioUpdate(recdictio)

Loop
Close #FileNumber
GoTo End_Function

Error_Handler:
     MsgBox "erreur " & Error
End_Function:

End Sub

Public Sub optAdresse_Sql()
Dim I As Integer, Msg As String

ReDim arrRacine(1): arrRacineNbMax = 1

recRacineInit recRacine
recRacine.Method = "SnapL0"
recRacine.Numéro = CLng(Val(txtStart))
arrRacine(0) = recRacine
arrRacine(0).Numéro = CLng(Val(txtEnd))
blnPrint = True
lst.Clear
If blnPrint Then prtAdresse_Open
arrRacineSuite = True
Do Until Not arrRacineSuite
    arrRacineNb = 0: arrRacineIndex = 0
    
    srvRacineMon recRacine
    recRacine = arrRacine(arrRacineNb)
    recRacine.Method = "SnapL0+"
    
    lst.Clear
    For I = 1 To arrRacineNb
        lst.AddItem arrRacine(I).Numéro & Chr$(9) & arrRacine(I).Intitulé
    Next I
    If blnPrint And arrRacineNb > 0 Then
         Msg = Format$(1, "000000") & Format$(arrRacineNb, "000000")
         prtAdresseX Msg
     End If
Loop
If blnPrint Then prtAdresse_Close

End Sub

Private Sub cmdCB_ClientExport_Load()
Dim xInput As String, blnOk As Boolean, wAmjCréation As String * 8, wAmjAnnulation As String * 8
Dim vReturn As Variant, X As String, SrvCptP0_Amj As String * 8

On Error GoTo Error_Handler

Dim I As Integer

For I = 0 To 3
    C0(I) = 0: C1(I) = 0: CB(I) = 0: CHB(I) = 0: CA(I) = 0:
    E0(I) = 0: E1(I) = 0: EBHB(I) = 0: E1A(I) = 0: EX(I) = 0
Next I

blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0

'''paramComptaExt_Cpt_Import = "c:\Biasrv\SrvCptP0"

X = Dir(paramComptaExt_Cpt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier des comptes n'existe pas"): Exit Sub

Call lstErr_Clear(lstErr, cmdPrint, "Chargement des comptes, tri ..."): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False

MDB.Execute "delete * from CptP0"
mdbCptP0.tableCptP0_Open

Open paramComptaExt_Cpt_Import For Input As #1
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
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvCptP0 ")
            Exit Do
        End If
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    
    If mId$(xInput, 115, 1) = "A" Then   ' compte auxilaire  kCb = 1(banque), 2 (pers morales), 3 (pers physiques)
    
        wSituation = mId$(xInput, 116, 1)
        wBiaTyp = mId$(xInput, 18, 3)
        
        If mId$(xInput, 13, 5) < "30000" Then
            kCB = 1
        Else
            If mId$(xInput, 249, 2) = "01" Or mId$(xInput, 249, 2) = "02" Then
                kCB = 3
            Else
                kCB = 2
            End If
        End If
        C0(kCB) = C0(kCB) + 1
      
        If wSituation = "A" Then
            CA(kCB) = CA(kCB) + 1
        Else
            Select Case wBiaTyp
                Case "001": C1(kCB) = C1(kCB) + 1
                Case Is < "900": CB(kCB) = CB(kCB) + 1
                Case Else: CHB(kCB) = CHB(kCB) + 1
           End Select
        End If
'$$$CB
''        If wBiaTyp = "001" Then
''        If wBiaTyp = "550" Then
''$$$ AUDIT
        If wBiaTyp = "001" And mId$(xInput, 281, 1) <> "0" Then   ' compte ordinaire et conservation du courrier
            reccptp0.Id = mId$(xInput, 13, 5)
            If wSituation <> "A" Then Mid$(xInput, 546, 8) = "00000000"
    
            reccptp0.Method = "Seek="
            If tableCptP0_Read(reccptp0) <> 0 Then
                reccptp0.Method = "AddNew"
                reccptp0.Text = xInput
                cmdImport_Select_Nb = cmdImport_Select_Nb + 1
                dbCptP0_Update reccptp0
            Else
                
                wAmjCréation = mId$(xInput, 530, 8)
                If mId$(reccptp0.Text, 530, 8) > wAmjCréation Then Mid$(reccptp0.Text, 530, 8) = wAmjCréation
               
                 If wSituation = "A" Then
                    If mId$(reccptp0.Text, 546, 8) <> "00000000" Then
                        wAmjAnnulation = mId$(xInput, 546, 8)
                        If mId$(reccptp0.Text, 546, 8) < wAmjAnnulation Then Mid$(reccptp0.Text, 546, 8) = wAmjAnnulation
                   End If
               Else
                    Mid$(reccptp0.Text, 546, 8) = "00000000"
                End If
                reccptp0.Method = "Update"
                dbCptP0_Update reccptp0
            
            End If
            
        End If
    End If
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Sélection des comptes : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
 
Loop

Close

 ''Open "c:\Biasrv\SrvRacine" For Input As #1

Open "\\FR11024427\AS400_OUT\SrvRacine" For Input As #1
Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
            Exit Do
        End If
        
        If mId$(xInput, 1, 5) < "30000" Then
            kCB = 1
        Else
            If mId$(xInput, 210, 2) = "01" Or mId$(xInput, 210, 2) = "02" Then
                kCB = 3
            Else
                kCB = 2
            End If
        End If
        E0(kCB) = E0(kCB) + 1
        
            reccptp0.Id = mId$(xInput, 1, 5)
            
            reccptp0.Method = "Seek="
            If tableCptP0_Read(reccptp0) = 0 Then
                    Mid$(reccptp0.Text, 186, 3) = mId$(xInput, 216, 3)   ' pays résidence
                    reccptp0.Method = "Update"
                    dbCptP0_Update reccptp0

            End If
Loop
mdbCptP0.tableCptP0_Close
Close
Me.MousePointer = 0
If Not blnOk Then
    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvCptP0 ")
End If

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdCB_ClientExport_Load")
Me.Enabled = True

End Sub
Private Sub cmdCB_CompteNR_Load()
Dim xInput As String, blnOk As Boolean, wAmjCréation As String * 8, wAmjAnnulation As String * 8
Dim vReturn As Variant, X As String, SrvCptP0_Amj As String * 8, X2 As String
Dim blnSelect As Boolean
Dim curX As Currency
On Error GoTo Error_Handler

Dim I As Integer
blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0

'jpl   paramComptaExt_Cpt_Import = "c:\Biasrv\SrvCptP0"

X = Dir(paramComptaExt_Cpt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier des comptes n'existe pas"): Exit Sub

Call lstErr_Clear(lstErr, cmdPrint, "Chargement des comptes, tri ..."): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False


Open paramComptaExt_Cpt_Import For Input As #1
Open txtCBFile For Output As #2

Print #2, "Sélection ;Particulier non résident compte ordinaire actif; ; "
Print #2, "Compte(11);Intitulé(40);Date dernier mouvement (8 AMJ); Devise (3)"

Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        SrvCptP0_Amj = mId$(xInput, 35, 8)
        I = Val(mId$(xInput, 43, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvCptP0 ")
            Exit Do
        End If
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    
    If mId$(xInput, 115, 1) = "A" Then   ' compte auxilaire  kCb = 1(banque), 2 (pers morales), 3 (pers physiques)
        blnSelect = True
        
        If mId$(xInput, 116, 1) = "A" Then blnSelect = False   ' ? compte annulé
        If mId$(xInput, 252, 1) = "1" Then blnSelect = False    ' ? Résident
        If mId$(xInput, 562, 8) = "00000000" Then blnSelect = False    ' ? sans mouvement
        
        If mId$(xInput, 13, 5) < "30000" Then                   ' particulier
            blnSelect = False
        Else
            If mId$(xInput, 249, 2) = "01" Or mId$(xInput, 249, 2) = "02" Then
            Else
                blnSelect = False
        End If
         wBiaTyp = mId$(xInput, 18, 3)
       
        If wBiaTyp = "001" Then
        If blnSelect Then
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            curX = CCur(Val(mId$(xInput, 119, 19)))
               X2 = Format$(curX, "############0.00")
            If curX < 0 Then
                X2 = X2 & "DB"
            Else
                X2 = X2 & "CR"
            End If
            
            X = mId$(xInput, 13, 11) & ";" & mId$(xInput, 35, 40) & ";" & mId$(xInput, 562, 8) & ";" & mId$(xInput, 7, 3) & ";" & X2
            Print #2, X
        End If
        
            End If
            
        End If
    End If
 
Loop

Close


Me.MousePointer = 0
If Not blnOk Then
    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvCptP0 ")
End If

Call lstErr_AddItem(lstErr, cmdPrint, "Terminé : " & cmdImport_Select_Nb): DoEvents

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdCB_ClientExport_Load")
Me.Enabled = True

End Sub

Private Sub cmdEuroImport_Load()
Dim xInput As String, blnOk As Boolean, wAmjCréation As String * 8, wAmjAnnulation As String * 8
Dim vReturn As Variant, X As String, SrvCptP0_Amj As String * 8, X2 As String
Dim blnSelect As Boolean
Dim curX As Currency
Dim wService As String

On Error GoTo Error_Handler

Dim I As Integer
blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
meCV1 = CV_Euro

If blnJPL Then paramComptaExt_Cpt_Import = "c:\Temp\AS400_Out\SrvCptP0"

X = Dir(paramComptaExt_Cpt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier des comptes n'existe pas"): Exit Sub

Call lstErr_Clear(lstErr, cmdPrint, "Chargement des comptes, tri ..."): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False

MDB.Execute "delete * from CptP0"
mdbCptP0.tableCptP0_Open
recCptP0_Init reccptp0
reccptp0.Method = "AddNew"

MDB.Execute "delete * from mvtP0"
mdbMvtP0.tableMvtP0_Open
recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"

Open paramComptaExt_Cpt_Import For Input As #1


Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        SrvCptP0_Amj = mId$(xInput, 35, 8)
        I = Val(mId$(xInput, 43, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "cmdEuroImport_Load")
            Exit Do
        End If
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    wSituation = mId$(xInput, 116, 1)
    wBiaTyp = mId$(xInput, 18, 3)
    If meCV1.DeviseN <> mId$(xInput, 7, 3) Then
        meCV1.DeviseN = mId$(xInput, 7, 3)
        CV_AttributN meCV1
    End If
            
     If "978" = mId$(xInput, 7, 3) Then    ' compte en EUROS / ancien N°
        recMvtp0.Id = mId$(xInput, 24, 11)
        recMvtp0.Method = "AddNew"
        recMvtp0.Text = xInput
        dbMvtP0_Update recMvtp0
    End If

    blnSelect = meCV1.EuroIn
    If mId$(xInput, 115, 1) = "R" Then blnSelect = False
    If mId$(xInput, 115, 1) = "A" Then   ' compte auxilaire  kCb = 1(banque), 2 (pers morales), 3 (pers physiques)
        wService = mId$(xInput, 282, 3)
        If wBiaTyp = "001" Then wService = "000"
    Else
        wService = "999": wBiaTyp = "000"
        '''blnSelect = False
    End If
    
   '' blnSelect = True
    
    If blnSelect Then
        wService = "000" ' mId$(xInput, 18, 3) ' 20021218 BIATYP
        reccptp0.Id = wService & mId$(xInput, 13, 11) & mId$(xInput, 7, 3)  ' meCV1.DeviseIso
        reccptp0.Method = "AddNew"
        reccptp0.Text = xInput
        cmdImport_Select_Nb = cmdImport_Select_Nb + 1
        dbCptP0_Update reccptp0
    End If
            
    
 
Loop

Close
mdbCptP0.tableCptP0_Close
mdbMvtP0.tableMvtP0_Close


Me.MousePointer = 0
If Not blnOk Then
    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "cmdEuroImport_Load")
End If

Call lstErr_AddItem(lstErr, cmdPrint, "cmdEuroImport_Load,Terminé : " & cmdImport_Select_Nb): DoEvents

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdEuroImport_Load")
Me.Enabled = True

End Sub

Private Sub cmdEuroPrint_List(lFct As String)
Dim X As String, blnSelect As Boolean, blnPrint As Boolean
Dim mService As String, mBiatyp As String
Dim blnPrint_Init As Boolean, blnPrint_Service As Boolean, blnPrint_Biatyp As Boolean
Dim wSituation As String
Dim wCompte As String, wX8 As String

On Error GoTo Error_Handler

Dim I As Integer
If lFct = "cmdEuroPrint" Then
    blnPrint = True
Else
    blnPrint = False
    Open paramEuroBascule For Output As #2
End If

cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
meCV1 = CV_Euro

Call lstErr_Clear(lstErr, cmdPrint, "cmdEuro_Print"): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False


mdbMvtP0.tableMvtP0_Open
recMvtP0_Init recMvtp0

mdbCptP0.tableCptP0_Open
recCptP0_Init reccptp0
reccptp0.Method = "MoveFirst"

If blnPrint Then prtInformatique_Open

blnPrint_Init = False
blnPrint_Service = False
blnPrint_Biatyp = False

'$jpl20011025$$$$$$$$$$$$$$$$$$$$$$$$$
Do
intReturn = tableCptP0_Read(reccptp0)

'$$$$$$$$$$$$$$$$$$$$$$$$$

'X = Dir(paramEuroBascule)
'If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier EURO_BASCULE n'existe pas"): Exit Sub
'recCompteInit meCompte
'recCptInfoInit meCptinfo
'Open paramEuroBascule For Input As #1
'reccptp0.Method = "Seek="
'Do Until EOF(1)
'    Line Input #1, X
'    If Trim(X) <> "" Then
'        Call lstErr_ChangeLastItem(lstErr, cmdPrint, X)
'        reccptp0.Id = "000" & mId$(X, 5, 11) & mId$(X, 1, 3)
'        intReturn = tableCptP0_Read(reccptp0)
'        If intReturn <> 0 Then MsgBox Trim(X), vbCritical, "Compte inconnu"

'$$$$$$$$$$$$$$$$$$$$$$$$$

    If intReturn = 0 Then
        blnSelect = False
        cmdImport_Select_Nb = cmdImport_Select_Nb + 1
        recMvtp0.Id = mId$(reccptp0.Text, 24, 8) & "978"
        recMvtp0.Method = "Seek="
        If tableMvtP0_Read(recMvtp0) <> 0 Then
            X = "+"
        Else
            If mId$(recMvtp0.Text, 13, 11) = mId$(reccptp0.Text, 13, 11) Then
                X = ""
                If mId$(recMvtp0.Text, 116, 1) <> " " Then
                  ''  blnSelect = True
                    X = "!eur= " & mId$(recMvtp0.Text, 116, 1)
                End If
            Else
               '' blnSelect = True
                X = "!!! " & Compte_Imp(mId$(recMvtp0.Text, 13, 11)) & "_" & mId$(reccptp0.Id, 1, 3)
            End If
        End If
        
      '  blnSelect = False
     '  If CCur(Val(mId$(reccptp0.Text, 119, 19))) <> 0 Then blnSelect = True
       blnSelect = True
         wSituation = mId$(reccptp0.Text, 116, 1)
         If wSituation = "E" Then blnSelect = False
         If wSituation = "A" Then blnSelect = False
         If wSituation = " " Then blnSelect = False
         If wSituation = "F" Then blnSelect = False
        
      '  blnSelect = False
      '      wCompte = mId$(reccptp0.Text, 13, 11)
      '      wX8 = mId$(reccptp0.Text, 13, 7)
      '      If wCompte = "20000013019" Then blnSelect = True
      '      If wCompte = "25066010011" Then blnSelect = True
      '
      '      If wCompte = "00010100006" Then blnSelect = True
      '      If wCompte = "00010101002" Then blnSelect = True
      '      If wCompte = "00016217005" Then blnSelect = True
      '      If wCompte = "00020111009" Then blnSelect = True
      '      If wCompte = "00020211036" Then blnSelect = True
       '     If wCompte = "00026211001" Then blnSelect = True
        '    If wCompte = "00026212007" Then blnSelect = True
         '   If wCompte = "00038121000" Then blnSelect = True
          '  If wCompte = "00038212005" Then blnSelect = True
           ' If wCompte = "00038150009" Then blnSelect = True
            'If wCompte = "00093600001" Then blnSelect = True
            
     '       wX8 = mId$(reccptp0.Text, 13, 8)
      '      If wX8 = "00038890" Then blnSelect = True
       '     If wX8 = "00096890" Then blnSelect = True
  
        If blnSelect Then
            If mService <> mId$(reccptp0.Id, 1, 3) Then
              mService = mId$(reccptp0.Id, 1, 3)
              If Not blnPrint_Init Then
                    blnPrint_Init = True
                Else
                    prtTitleText = "Liste des comptes en devise IN : " & mService

                      If blnPrint Then XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
                 ' XPrt.CurrentY = prtMaxY + 10000  'saut de page
                End If
            Else
              '   If mBiatyp <> mId$(reccptp0.Id, 4, 3) Then
              '      If blnPrint Then XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
              ' End If
            End If
            
            mBiatyp = mId$(reccptp0.Id, 4, 3)
           
            MsgTxtIndex = 0
            MsgTxt = Space$(recCptInfoLen)
            Mid$(MsgTxt, 35, memoCptInfoLen) = mId$(reccptp0.Text, 1, memoCptInfoLen)
            Call srvCptInfoGetBuffer(recCptInfo)
            If blnPrint Then
                Call prtInformatique_Line(X, recCptInfo, mId$(reccptp0.Id, 18, 3))
        
            Else
                X = recCptInfo.Devise & " " & recCptInfo.Numéro & " " & recCptInfo.Intitulé
                Print #2, X
            End If
        End If
        End If
 '   End If
'$$$$$$$$$$$$$$$$$$$$$$$$$
    reccptp0.Method = "MoveNext" '$jpl20011025
Loop While intReturn = 0  '$jpl20011025
'Loop
'$$$$$$$$$$$$$$$$$$$$$$$$$

If blnPrint Then prtInformatique_Close
Close
mdbCptP0.tableCptP0_Close
mdbMvtP0.tableMvtP0_Close


Me.MousePointer = 0
Call lstErr_AddItem(lstErr, cmdPrint, "cmdEuro_Print Terminé : " & cmdImport_Select_Nb): DoEvents

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "ccmdEuro_Print")
Me.Enabled = True

End Sub

Private Sub cmdEuroExport_List(lFct As String)
Dim X As String, blnSelect As Boolean, blnPrint As Boolean
Dim mService As String, mBiatyp As String
Dim blnPrint_Init As Boolean, blnPrint_Service As Boolean, blnPrint_Biatyp As Boolean

On Error GoTo Error_Handler

Dim I As Integer
If lFct = "cmdEuroPrint" Then
    blnPrint = True
Else
    blnPrint = False
    Open paramEuroBascule For Output As #2
End If

cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
meCV1 = CV_Euro

Call lstErr_Clear(lstErr, cmdPrint, "cmdEuro_Print"): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False


mdbMvtP0.tableMvtP0_Open
recMvtP0_Init recMvtp0

mdbCptP0.tableCptP0_Open
recCptP0_Init reccptp0
reccptp0.Method = "MoveFirst"

If blnPrint Then prtInformatique_Open

blnPrint_Init = False
blnPrint_Service = False
blnPrint_Biatyp = False

Do
intReturn = tableCptP0_Read(reccptp0)



    If intReturn = 0 Then
        blnSelect = False
        cmdImport_Select_Nb = cmdImport_Select_Nb + 1
        recMvtp0.Id = mId$(reccptp0.Text, 24, 8) & "978"
        recMvtp0.Method = "Seek="
        If tableMvtP0_Read(recMvtp0) <> 0 Then
            X = "+"
        Else
            If mId$(recMvtp0.Text, 13, 11) = mId$(reccptp0.Text, 13, 11) Then
                X = ""
                If mId$(recMvtp0.Text, 116, 1) <> " " Then
                  ''  blnSelect = True
                    X = "!eur= " & mId$(recMvtp0.Text, 116, 1)
                End If
            Else
               '' blnSelect = True
                X = "!!! " & Compte_Imp(mId$(recMvtp0.Text, 13, 11)) & "_" & mId$(reccptp0.Id, 1, 3)
            End If
        End If
        
       blnSelect = False
       'blnSelect = True
       If CCur(Val(mId$(reccptp0.Text, 119, 19))) = 0 Then blnSelect = True
       
       'If CCur(Val(mId$(reccptp0.Text, 119, 19))) = 0 Then
       '     If mId$(reccptp0.Text, 18, 3) = "032" _
       '     Or mId$(reccptp0.Text, 18, 3) = "100" _
       '     Or mId$(reccptp0.Text, 18, 3) = "910" _
       '     Or mId$(reccptp0.Text, 18, 3) = "913" _
       '     Or mId$(reccptp0.Text, 18, 3) = "914" _
       '     Or mId$(reccptp0.Text, 18, 3) = "916" _
       '     Or mId$(reccptp0.Text, 18, 3) = "918" _
       '     Or mId$(reccptp0.Text, 18, 3) = "944" Then

       '                 blnSelect = True
       '     End If
       'End If
       ' If mId$(reccptp0.Text, 116, 1) = " " And mId$(reccptp0.Text, 18, 3) = "976" And CCur(Val(mId$(reccptp0.Text, 119, 19))) = 0 Then
       '     If mId$(reccptp0.Text, 13, 5) > "30000" Then blnSelect = True
       'End If
       ' If blnSelect Then
       '     If CCur(Val(mId$(reccptp0.Text, 119, 19))) = 0 Then blnSelect = False
       ' End If
       
       wSituation = mId$(reccptp0.Text, 116, 1)
    If wSituation = "A" Or wSituation = "E" Or wSituation = "R" Then blnSelect = False
        If blnSelect Then
            If mService <> mId$(reccptp0.Id, 1, 3) Then
              mService = mId$(reccptp0.Id, 1, 3)
              If Not blnPrint_Init Then
                    blnPrint_Init = True
                Else
                    prtTitleText = "Liste des comptes en devise IN : " & mService

                    If blnPrint Then XPrt.CurrentY = prtMaxY + 10000  'saut de page
                End If
            Else
                 If mBiatyp <> mId$(reccptp0.Id, 4, 3) Then
                    If blnPrint Then XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
               End If
            End If
            
            mBiatyp = mId$(reccptp0.Id, 4, 3)
           
            MsgTxtIndex = 0
            MsgTxt = Space$(recCptInfoLen)
            Mid$(MsgTxt, 35, memoCptInfoLen) = mId$(reccptp0.Text, 1, memoCptInfoLen)
            Call srvCptInfoGetBuffer(recCptInfo)
            If blnPrint Then
                Call prtInformatique_Line(X, recCptInfo, mId$(reccptp0.Id, 18, 3))
        
            Else
                X = recCptInfo.Devise & " " & recCptInfo.Numéro & " " & recCptInfo.Intitulé
                Print #2, X
            End If
        End If
        End If
    reccptp0.Method = "MoveNext" '$jpl20011025
Loop While intReturn = 0  '$jpl20011025

If blnPrint Then prtInformatique_Close
Close
mdbCptP0.tableCptP0_Close
mdbMvtP0.tableMvtP0_Close


Me.MousePointer = 0
Call lstErr_AddItem(lstErr, cmdPrint, "cmdEuro_Print Terminé : " & cmdImport_Select_Nb): DoEvents

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "ccmdEuro_Print")
Me.Enabled = True

End Sub


Public Sub cmdEuroBascule_DB()
Dim X As String, V
On Error GoTo Error_Handler

X = Dir(paramEuroBascule)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier EURO_BASCULE n'existe pas"): Exit Sub

recCompteInit meCompte

recCptInfoInit meCptinfo
Call lstErr_Clear(lstErr, cmdPrint, "EURO BASCULE")

Me.MousePointer = vbHourglass
Me.Enabled = False

Open paramEuroBascule For Input As #1
cmdImport_Select_Nb = 0

Do Until EOF(1)
    Line Input #1, X
    If Trim(X) <> "" Then
        Call lstErr_ChangeLastItem(lstErr, cmdPrint, X)
        meCompte.Société = SocId$
        meCompte.Agence = SocAgence$
        meCompte.Devise = mId$(X, 1, 3)
        meCompte.Numéro = mId$(X, 5, 11)
        
        meCompte.obj = "SRVCOMPTE"
        meCompte.Method = "SeekL1"
        V = srvCompte_InitFind(meCompte)
        
        If Not IsNull(V) Then
            Call MsgBox(X, vbCritical, "? compte inconnu : ")
        Else
        
            meCptinfo.Devise = meCompte.Devise
            meCptinfo.Numéro = meCompte.Numéro
            V = srvCptInfoFind(meCptinfo)
            
             If Not IsNull(V) Then
                Call MsgBox(X, vbCritical, "? CptINfo inconnu : ")
            Else
           
                meCV1.DeviseN = meCompte.Devise
                meCV1.DeviseIso = ""
                meCV1.Montant = meCompte.DécouvertMontant
                Call CV_Transitoire(meCV1, meCV2, meCV3, X)
                meCompte.DécouvertMontant = Round(meCV3.Montant, 0)
                
                meCV1.Montant = meCptinfo.EchelleSolde
                Call CV_Transitoire(meCV1, meCV2, meCV3, X)
                meCompte.SoldeVeille = meCV3.Montant
                
                meCompte.obj = "SRVEURO     "
                meCompte.Method = "Bascule"
                cmdImport_Select_Nb = cmdImport_Select_Nb + 1
                If Not IsNull(srvCompte_Update(meCompte)) Then Call MsgBox("Erreur Maj", vbCritical, "Compte_Euro")
            
            End If
        End If
    End If
Loop

Close

Me.MousePointer = 0

Call lstErr_AddItem(lstErr, cmdPrint, "cmdEuroUpdate_DB,Terminé : " & cmdImport_Select_Nb): DoEvents

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdEuroUpdate_DB")
Me.Enabled = True


End Sub

Public Sub cmdCompteAnnulation_DB()
Dim X As String, V
On Error GoTo Error_Handler

X = Dir(paramEuroBascule)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier  n'existe pas : " & paramEuroBascule): Exit Sub

recCompteInit meCompte

recCptInfoInit meCptinfo
Call lstErr_Clear(lstErr, cmdPrint, "EURO BASCULE")

Me.MousePointer = vbHourglass
Me.Enabled = False

Open paramEuroBascule For Input As #1
cmdImport_Select_Nb = 0

Do Until EOF(1)
    Line Input #1, X
    If Trim(X) <> "" Then
        Call lstErr_ChangeLastItem(lstErr, cmdPrint, X)
        meCompte.Société = SocId$
        meCompte.Agence = SocAgence$
        meCompte.Devise = mId$(X, 1, 3)
        meCompte.Numéro = mId$(X, 5, 11)
        
        meCompte.obj = "SRVCOMPTE"
        meCompte.Method = "SeekL1"
        V = srvCompte_InitFind(meCompte)
        
        If Not IsNull(V) Then
            Call MsgBox(X, vbCritical, "? compte inconnu : ")
        Else
        
            If meCompte.SoldeVeille = 0 And meCompte.SoldeInstantané = 0 Then
                
                meCompte.obj = "SRVEURO     "
                meCompte.Method = "Annul"
                cmdImport_Select_Nb = cmdImport_Select_Nb + 1
                If Not IsNull(srvCompte_Update(meCompte)) Then Call MsgBox("Erreur Maj", vbCritical, "Compte_Euro")
            
            End If
        End If
    End If
Loop

Close

Me.MousePointer = 0

Call lstErr_AddItem(lstErr, cmdPrint, "cmdEuroUpdate_DB,Terminé : " & cmdImport_Select_Nb): DoEvents

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdEuroUpdate_DB")
Me.Enabled = True


End Sub

Private Sub cmdCB_ClientExport_Ex()
Dim xInput As String, blnOk As Boolean, wAmjCréation As String * 8, wAmjAnnulation As String * 8
Dim vReturn As Variant, X As String, SrvCptP0_Amj As String * 8

On Error GoTo Error_Handler


Call lstErr_Clear(lstErr, cmdPrint, "Chargement des comptes, tri ..."): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False

mdbCptP0.tableCptP0_Open

Open paramComptaExt_Cpt_Import For Input As #1
Open txtCBFile & "_Divers" For Output As #2

recCptP0_Init reccptp0
reccptp0.Method = "AddNew"


Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
           Exit Do
    End If
    
    If mId$(xInput, 115, 1) = "A" Then   ' compte auxilaire  kCb = 1(banque), 2 (pers morales), 3 (pers physiques)
    
        wBiaTyp = mId$(xInput, 18, 3)
        
              
        If wBiaTyp <> "001" Then
            reccptp0.Id = mId$(xInput, 13, 5)
            reccptp0.Method = "Seek="
            If tableCptP0_Read(reccptp0) <> 0 Then
            
                 If mId$(xInput, 13, 5) < "30000" Then
                    kCB = 1
                Else
                    If mId$(xInput, 249, 2) = "01" Or mId$(xInput, 249, 2) = "02" Then
                        kCB = 3
                    Else
                        kCB = 2
                    End If
                End If
               EBHB(kCB) = EBHB(kCB) + 1
               
                reccptp0.Method = "AddNew"
                
                reccptp0.Text = xInput
                dbCptP0_Update reccptp0
                           
                Print #2, mId$(reccptp0.Text, 13, 11) & ";" & mId$(reccptp0.Text, 35, 40)

           
            End If
            
        End If
    End If
 
Loop

Close #1
'''jpl Open "c:\Biasrv\SrvRacine" For Input As #1

Open "\\FR11024427\AS400_OUT\SrvRacine" For Input As #1
Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
            Exit Do
        End If
        
        
            reccptp0.Id = mId$(xInput, 1, 5)
            
            reccptp0.Method = "Seek="
            If tableCptP0_Read(reccptp0) = 0 Then
                If mId$(xInput, 210, 2) <> mId$(reccptp0.Text, 249, 2) Then Debug.Print mId$(xInput, 1, 60)

            Else
            
                If mId$(xInput, 1, 5) < "30000" Then
                    kCB = 1
                Else
                    If mId$(xInput, 210, 2) = "01" Or mId$(xInput, 210, 2) = "02" Then
                        kCB = 3
                    Else
                        kCB = 2
                    End If
                End If
                EX(kCB) = EX(kCB) + 1

                Print #2, mId$(xInput, 1, 5) & ";" & mId$(xInput, 6, 40)

            End If
Loop

mdbCptP0.tableCptP0_Close
Close
Me.MousePointer = 0

Me.Enabled = True
Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdCB_ClientExport_Load")
Me.Enabled = True

End Sub

Public Sub cmdCB_ClientExport_Write()
Dim X250 As String * 250, wDec, wDecCV

On Error GoTo Error_Handler


Call lstErr_AddItem(lstErr, cmdContext, "Export : début"): DoEvents

Open txtCBFile For Output As #2

cmdImport_Nb = 0

mdbCptP0.tableCptP0_Open
recCptP0_Init reccptp0
reccptp0.Method = "MoveFirst"

Do
    intReturn = tableCptP0_Read(reccptp0)

    If intReturn = 0 Then
        If mId$(reccptp0.Text, 13, 5) < "30000" Then
            kCB = 1
        Else
            If mId$(reccptp0.Text, 249, 2) = "01" Or mId$(reccptp0.Text, 249, 2) = "02" Then
                kCB = 3
            Else
                kCB = 2
            End If
        End If
        
        If mId$(reccptp0.Text, 546, 8) = "00000000" Then
            E1(kCB) = E1(kCB) + 1
        Else
            E1A(kCB) = E1A(kCB) + 1
        End If
        

        X250 = mId$(reccptp0.Text, 13, 11) & ";" & mId$(reccptp0.Text, 35, 40) & ";" & mId$(reccptp0.Text, 530, 8) & ";" & mId$(reccptp0.Text, 546, 8) & ";" & mId$(reccptp0.Text, 186, 3)
        Print #2, X250
        cmdImport_Nb = cmdImport_Nb + 1
   End If
    reccptp0.Method = "MoveNext"
Loop While intReturn = 0


Close
mdbCptP0.tableCptP0_Close
Call lstErr_AddItem(lstErr, cmdContext, "Export : fin" & cmdImport_Nb): DoEvents


Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdCB_ClientExport_write")


End Sub

Public Sub cmdCB_COBIA_Write()
Dim X As String, X2 As String, K As Integer, blnSelect As Boolean

On Error GoTo Error_Handler


Call lstErr_AddItem(lstErr, cmdContext, "Export : début"): DoEvents

Open txtCBFile For Output As #2

Print #2, "Sélection ;Client ayant un compte ordinaire avec envoi de courrier; ; "
Print #2, "Compte(11);Intitulé(40);Date création (8 AMJ);Date annulation (8 AMJ);Date dernier mouvement (8 AMJ); adresse "

cmdImport_Nb = 0

mdbCptP0.tableCptP0_Open
recCptP0_Init reccptp0
reccptp0.Method = "MoveFirst"

Do
    intReturn = tableCptP0_Read(reccptp0)

    If intReturn = 0 Then
        MsgTxtIndex = 0
        MsgTxt = Space$(recCptInfoLen)
        Mid$(MsgTxt, 35, memoCptInfoLen) = mId$(reccptp0.Text, 1, memoCptInfoLen)
        Call srvCptInfoGetBuffer(recCptInfo)
        X2 = Trim(recCptInfo.Adresse1) & " " & Trim(recCptInfo.Adresse2) & " " & Trim(recCptInfo.Adresse3) & " " & Trim(recCptInfo.Adresse4) & " " & Trim(recCptInfo.AdresseCP) & " " & Trim(recCptInfo.AdresseBD) & " " & Trim(recCptInfo.AdressePays)
        blnSelect = True
        
'        blnSelect = False
'        K = InStr(1, X2, "ROOSEV")
'        If K > 0 Then blnSelect = True
'        K = InStr(1, X2, "C/O")
'        If K > 0 Then
'            K = InStr(1, X2, "BIA")
'            If K > 0 Then blnSelect = True
'        End If
'        If Trim(recCptInfo.AdresseBD) & Trim(recCptInfo.AdressePays) = "" Then blnSelect = True
        
        If recCptInfo.AmjAnnulation <> "00000000" Then blnSelect = False
        
        If blnSelect Then
            X = recCptInfo.Numéro & ";" & recCptInfo.Intitulé & ";" & recCptInfo.AmjCréation & ";" & recCptInfo.AmjAnnulation & ";" & recCptInfo.AmjDernierMouvement _
               & ";" & X2
            Print #2, X
            cmdImport_Nb = cmdImport_Nb + 1
        End If
   End If
    reccptp0.Method = "MoveNext"
Loop While intReturn = 0


Close
mdbCptP0.tableCptP0_Close
Call lstErr_AddItem(lstErr, cmdContext, "Export : fin" & cmdImport_Nb): DoEvents


Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdCB_ClientExport_write")


End Sub




Private Function cmdImport_Select(Msg As String) As String
Dim wCompteGénéral As String * 11, wDevise As String * 3, wNuméro As String * 11, wExtraitPériodicité As String * 1, wTypeGa As String * 1
Dim xAMJ As String * 8, xDébitFindeMois As Currency, xCréditFindeMois As Currency

cmdImport_Select = ""
wDevise = mId$(Msg, 7, 3)
wNuméro = mId$(Msg, 13, 11)
wTypeGa = mId$(Msg, 115, 1)



End Function


Public Function param_Init()
Dim V
param_Init = Null

recElpTable_Init recElpTable
recElpTable.Id = "Param"
recElpTable.K1 = "ComptaExt"
recElpTable.Method = "Seek="

recElpTable.K2 = "Cpt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramComptaExt_Cpt_Import = paramServer(recElpTable.Memo)
Call lstErr_Clear(lstErr, cmdContext, "Cpt_Import:" & paramComptaExt_Cpt_Import)


Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "Param_Init"
Exit Function

End Function

Private Sub Form_Load()
Call DTPicker_Now(txtAmjMin)
Call DTPicker_Now(txtAmjMax)
End Sub


