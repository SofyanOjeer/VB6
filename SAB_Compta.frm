VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_Compta 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Comptabilité : Interfaces"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SAB_Compta.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8400
      TabIndex        =   4
      Top             =   0
      Width           =   4875
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8925
      Left            =   0
      TabIndex        =   2
      Top             =   495
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   15743
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Mouvements Comptables"
      TabPicture(0)   =   "SAB_Compta.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Impression des relevés"
      TabPicture(1)   =   "SAB_Compta.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRelevéA4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Exploitation JOUR"
      TabPicture(2)   =   "SAB_Compta.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraInfo_Jour"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "EXploitation Mensuelle"
      TabPicture(3)   =   "SAB_Compta.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fraInfo"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraInfo 
         Height          =   8415
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   13545
         Begin VB.ListBox lstSort 
            Height          =   6885
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   67
            Top             =   720
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.Frame fraRelevéA4M 
            Caption         =   "Exploitation mensuelle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7980
            Left            =   6360
            TabIndex        =   37
            Top             =   360
            Width           =   7080
            Begin VB.CommandButton cmdImpRepertoire 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Impression d'un répertoire PDF"
               Height          =   375
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   73
               Top             =   5340
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.CommandButton cmdNIC_FGDR 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Insertion NIC FGDR (annuel)"
               Height          =   435
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   7320
               Width           =   2535
            End
            Begin VB.CommandButton cmdFraisBancairesPDF 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Info.préalable frais bancaires PDF"
               Height          =   375
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   71
               Top             =   2880
               Width           =   2565
            End
            Begin VB.CommandButton cmdRelevéA4MPDF 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Extraits Mensuels PDF"
               Height          =   375
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   70
               Top             =   1200
               Width           =   2565
            End
            Begin VB.CommandButton cmdRelevéA4DPDF 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Extraits Décadaires PDF"
               Height          =   375
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   69
               Top             =   2040
               Width           =   2565
            End
            Begin VB.CommandButton cmdRelevé_Annuel_FraisPDF 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Relevé des frais (annuel) PDF"
               Height          =   435
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   68
               Top             =   6780
               Width           =   2535
            End
            Begin VB.CommandButton cmdExtractAdresses 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Extraire les adresses relevés"
               Height          =   680
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   66
               Top             =   4320
               Width           =   2565
            End
            Begin VB.CommandButton cmdFraisBancaires 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Info.préalable frais bancaires"
               Height          =   375
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   65
               Top             =   2520
               Width           =   2565
            End
            Begin VB.CheckBox chkListCLIENACLI 
               Caption         =   "Utiliser la liste commerciale (YBIATAB0)"
               Height          =   315
               Left            =   240
               TabIndex        =   64
               Top             =   840
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.CommandButton cmdRelevé_Annuel_Frais 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Relevé des frais (annuel)"
               Height          =   435
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   6240
               Width           =   2535
            End
            Begin VB.CheckBox chkExtrait_AmjMin 
               Caption         =   "Imprimer extraits pour les comptes n'ayant pas mouvementé après le"
               Height          =   315
               Left            =   240
               TabIndex        =   61
               Top             =   360
               Width           =   5175
            End
            Begin VB.TextBox txtRelevéA4M_Compte 
               Height          =   255
               Left            =   4200
               TabIndex        =   51
               Top             =   3480
               Width           =   2385
            End
            Begin VB.TextBox txtRelevéA4M_Service 
               Height          =   285
               Left            =   3480
               TabIndex        =   50
               Top             =   4200
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CommandButton cmdETAFI_Ok 
               BackColor       =   &H00C0FFFF&
               Caption         =   "ETAFI "
               Height          =   435
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   5700
               Width           =   2535
            End
            Begin VB.Frame fraECH 
               Caption         =   "Echelles"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   240
               TabIndex        =   45
               Top             =   6000
               Width           =   2775
               Begin VB.CheckBox chkECH_Print 
                  Caption         =   "Impression"
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   48
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   1095
               End
               Begin VB.CheckBox chkECH_FTP 
                  Caption         =   "FTP"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   47
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   855
               End
               Begin VB.CommandButton cmdECH_Ok 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Traitement ECH"
                  Height          =   765
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   46
                  Top             =   720
                  Width           =   2055
               End
            End
            Begin VB.CommandButton cmdRelevéA4D 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Impression extraits Décadaires"
               Height          =   375
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   1680
               Width           =   2565
            End
            Begin VB.CommandButton cmdCOSI 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Déclaratif COSI"
               Height          =   680
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   4320
               Width           =   2565
            End
            Begin VB.CheckBox chkRelevéA4M_Rib 
               Caption         =   "Imprimer Relevé ET Rib"
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   2760
               Width           =   2250
            End
            Begin VB.CommandButton cmdRelevéA4M 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Impression extraits Mensuels"
               Height          =   375
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   840
               Width           =   2565
            End
            Begin VB.CheckBox chkRelevéA4M_CptOrdinaire 
               Caption         =   "uniquement Comptes Ordinaires"
               CausesValidation=   0   'False
               Height          =   285
               Left            =   240
               TabIndex        =   40
               Top             =   1680
               Value           =   1  'Checked
               Width           =   2640
            End
            Begin VB.CheckBox chkRelevéA4M_Responsable 
               Caption         =   "tri par Responsable"
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   2160
               Value           =   1  'Checked
               Width           =   2250
            End
            Begin VB.TextBox txtRelevéA4M_PCEC 
               Height          =   255
               Left            =   1080
               TabIndex        =   38
               Top             =   1320
               Width           =   1425
            End
            Begin MSComCtl2.DTPicker txtExtrait_AmjMin 
               Height          =   300
               Left            =   5400
               TabIndex        =   60
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   39190531
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Line Line6 
               X1              =   0
               X2              =   7080
               Y1              =   5280
               Y2              =   5280
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   7080
               Y1              =   4080
               Y2              =   4080
            End
            Begin VB.Label libRelevéA4M_Compte 
               Caption         =   "Reprise impression à partir du compte"
               Height          =   255
               Left            =   240
               TabIndex        =   54
               Top             =   3480
               Width           =   2715
            End
            Begin VB.Label lblRelevéA4M_Service 
               Caption         =   "Service G*           (blanc à faire JPL)"
               Height          =   495
               Left            =   2520
               TabIndex        =   53
               Top             =   4680
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label lblRelevéA4M_PCEC 
               Caption         =   "PCEC"
               Height          =   255
               Left            =   195
               TabIndex        =   52
               Top             =   1350
               Width           =   540
            End
         End
         Begin VB.ListBox lstW 
            Height          =   7860
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   3525
         End
         Begin VB.Line Line5 
            X1              =   6480
            X2              =   13320
            Y1              =   5880
            Y2              =   5880
         End
      End
      Begin VB.Frame fraInfo_Jour 
         Height          =   8385
         Left            =   -74895
         TabIndex        =   7
         Top             =   360
         Width           =   13500
         Begin VB.Frame fraRelevéA4J 
            Caption         =   "Exploitation JOUR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7620
            Left            =   8160
            TabIndex        =   27
            Top             =   480
            Width           =   3615
            Begin VB.CommandButton cmdJournal_Solde 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Impression contrôle des soldes"
               Height          =   1095
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   4200
               Width           =   2520
            End
            Begin VB.CommandButton cmdRelevéA4J 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Impression extraits J"
               Height          =   1095
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   6120
               Width           =   2520
            End
            Begin VB.CommandButton cmdJournal_Devises 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Impression journal Devises"
               Height          =   1095
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   2280
               Width           =   2520
            End
            Begin VB.CommandButton cmdJournal_Unit 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Impression journal Services"
               Height          =   1095
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   360
               Width           =   2520
            End
         End
         Begin VB.Frame fraRelevéA4W 
            Caption         =   "MT950"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7455
            Left            =   1080
            TabIndex        =   19
            Top             =   480
            Width           =   5355
            Begin VB.CommandButton cmdMT900 
               BackColor       =   &H00FF80FF&
               Caption         =   "émission MT900 -  Liste 'en dur ' dans Swift_MT900_Monitor"
               Height          =   1380
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   5520
               Width           =   2475
            End
            Begin VB.CheckBox chkRelevéA4W_Confirmation 
               Caption         =   "Confirmer la préparation MT950 compte par compte"
               Height          =   495
               Left            =   840
               TabIndex        =   26
               Top             =   1560
               Width           =   2550
            End
            Begin VB.CheckBox chkRelevéA4W_Update 
               Caption         =   "màj SAB073_SPE / YBIARELH"
               Height          =   270
               Left            =   840
               TabIndex        =   23
               Top             =   1200
               Width           =   2565
            End
            Begin VB.CheckBox chkRelevéA4W_Nostro 
               Caption         =   "Nostro ==> CORONA"
               Height          =   270
               Left            =   840
               TabIndex        =   22
               Top             =   780
               Width           =   2415
            End
            Begin VB.CheckBox chkRelevéA4W_Loro 
               Caption         =   "Loro   ==> SAA"
               Height          =   270
               Left            =   840
               TabIndex        =   21
               Top             =   405
               Width           =   2190
            End
            Begin VB.CommandButton cmdRelevéA4W 
               BackColor       =   &H000000FF&
               Caption         =   "MT950_Extraction"
               Height          =   1380
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   3600
               Width           =   2475
            End
            Begin MSComCtl2.DTPicker txtRelevéA4W_Amj 
               Height          =   300
               Left            =   3840
               TabIndex        =   33
               Top             =   1680
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
               Format          =   39190531
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblRelevéW 
               BackColor       =   &H0000FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Boutons désactivés ; voir @AUTO_COMPTA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   57
               Top             =   2640
               Width           =   4935
            End
         End
      End
      Begin VB.Frame fraRelevéA4 
         Height          =   8505
         Left            =   -74925
         TabIndex        =   6
         Top             =   360
         Width           =   13620
         Begin VB.TextBox txtRelevéA4_Compte 
            Height          =   285
            Left            =   3120
            TabIndex        =   24
            Top             =   8040
            Width           =   2565
         End
         Begin MSFlexGridLib.MSFlexGrid fgRelevéA4 
            Height          =   7785
            Left            =   60
            TabIndex        =   16
            Top             =   180
            Width           =   12150
            _ExtentX        =   21431
            _ExtentY        =   13732
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   12582912
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_Compta.frx":037A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker txtRelevéA4_AmjMin 
            Height          =   300
            Left            =   10680
            TabIndex        =   17
            Top             =   8040
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
            Format          =   39190531
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtRelevéA4_AmjMax 
            Height          =   300
            Left            =   12240
            TabIndex        =   18
            Top             =   8040
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
            Format          =   39190531
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblRelevéA4_Compte 
            Caption         =   "Compte"
            Height          =   255
            Left            =   165
            TabIndex        =   25
            Top             =   8040
            Width           =   1245
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   8355
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   13530
         Begin VB.Frame fraContextOptions 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6375
            Left            =   8880
            TabIndex        =   9
            Top             =   1200
            Width           =   4395
            Begin VB.OptionButton optOptions_SelectJC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFFF&
               Caption         =   "Selection JC exercice précédent"
               Height          =   210
               Left            =   720
               TabIndex        =   62
               Top             =   3480
               Width           =   2900
            End
            Begin VB.ComboBox cboDevise 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2400
               TabIndex        =   58
               Text            =   "devise"
               Top             =   3960
               Width           =   1335
            End
            Begin VB.CommandButton cmdContextOptions_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Exécuter la requête"
               Height          =   885
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   5040
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker txtSelectAmj_Max 
               Height          =   300
               Left            =   2280
               TabIndex        =   34
               Top             =   360
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   39190531
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.OptionButton optOptions_SelectOD 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFFF&
               Caption         =   "Selection OD de la période"
               Height          =   210
               Left            =   720
               TabIndex        =   32
               Top             =   3000
               Width           =   2900
            End
            Begin VB.ComboBox cboOptions_Unit 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   2280
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   960
               Width           =   1350
            End
            Begin VB.OptionButton optOptions_SortUnit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFFF&
               Caption         =   "tri  : Service / Opération"
               Height          =   240
               Left            =   720
               TabIndex        =   12
               Top             =   1560
               Value           =   -1  'True
               Width           =   2900
            End
            Begin VB.OptionButton optOptions_SortDevise 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFFF&
               Caption         =   "tri : Devise / Opération"
               Height          =   240
               Left            =   720
               TabIndex        =   11
               Top             =   2040
               Width           =   2900
            End
            Begin VB.OptionButton optOptions_SortCompte 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFFF&
               Caption         =   "tri : Compte / Devise "
               Height          =   240
               Left            =   720
               TabIndex        =   10
               Top             =   2520
               Width           =   2900
            End
            Begin MSComCtl2.DTPicker txtSelectAmj 
               Height          =   300
               Left            =   720
               TabIndex        =   14
               Top             =   360
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   39190531
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblDevise 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Devise"
               Height          =   255
               Left            =   960
               TabIndex        =   59
               Top             =   4080
               Width           =   1215
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   4440
               Y1              =   4680
               Y2              =   4680
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   4320
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label lblOptions_Unit 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Service"
               Height          =   240
               Left            =   840
               TabIndex        =   15
               Top             =   960
               Width           =   645
            End
         End
         Begin VB.CommandButton cmdOptions 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Options"
            Height          =   645
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   1350
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7425
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   13275
            _ExtentX        =   23416
            _ExtentY        =   13097
            _Version        =   393216
            Rows            =   1
            Cols            =   10
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   12582912
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_Compta.frx":0442
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SAB_Compta.frx":04ED
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print 
         Caption         =   "Imprimer le journal"
      End
      Begin VB.Menu mnuSelect_Print_Recap 
         Caption         =   "Imprimer les totaux"
      End
   End
   Begin VB.Menu mnuPrint1 
      Caption         =   "mnuPrint1"
      Visible         =   0   'False
      Begin VB.Menu mnuRelevé_Print 
         Caption         =   "Imprimer Relevé"
      End
      Begin VB.Menu mnuRIB_Print 
         Caption         =   "Imprimer RIB"
      End
      Begin VB.Menu mnuRelevéRIB_Print 
         Caption         =   "Imprimer Relevé +RIB"
      End
   End
End
Attribute VB_Name = "frmSAB_Compta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim SAB_Compta_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean
 
Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgRelevéA4_FormatString As String, fgRelevéA4_K As Integer
Dim fgRelevéA4_RowDisplay As Integer, fgRelevéA4_RowClick As Integer, fgRelevéA4_ColClick As Integer
Dim fgRelevéA4_ColorClick As Long, fgRelevéA4_ColorDisplay As Long
Dim fgRelevéA4_Sort1 As Integer, fgRelevéA4_Sort2 As Integer
Dim fgRelevéA4_SortAD As Integer, fgRelevéA4_Sort1_Old As Integer
Dim fgRelevéA4_arrIndex As Integer
Dim blnfgRelevéA4_DisplayLine As Boolean


Dim meYBIACPT0 As typeYBIACPT0, xYBIACPT0 As typeYBIACPT0
Dim meYBIAMVT0 As typeYBIAMVT0, xYBIAMVT0 As typeYBIAMVT0
Dim arrYBIAMVT0() As typeYBIAMVT0, arrYBIAMVT0_Nb As Long, arrYBIAMVT0_Max As Long

Dim mUnit As String, previousUnit As String
Dim blnTotal As Boolean
Dim blnNewPage As Boolean

Dim wSelectAmj As String * 8, xSelectAmj_IBM As String * 7, Nb As Long
Dim xSelectAmj_Max_IBM As String * 7

Dim mMOUVEMCOM As String
Dim meZADRESS0 As typeZADRESS0
Dim xZRELEVE0 As typeZRELEVE0
Dim meUnit As typeUnit
Dim xZCLIENA0 As typeZCLIENA0


Dim marrZCOMREF0() As typeZCOMREF0, marrZCOMREF0_Nb As Long, marrZCOMREF0_Index As Long

Dim specialMVT() As String, specialMVT_nb As Long, specialMVT_Max As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel

Dim PDF_WAIT As Long
Private Sub cmdInfoFrais_PrintPDF(lRELEVEREL As String, lMin As String, lMax As String)
Dim iW As Long
Dim blnOk As Boolean
Dim s() As String
Dim wRéférence As String
Dim wRELEVENUM As String

'//////////////////////////////////////////////////////////////////////////
Dim numCli As String
Dim increment As Long
Dim repPDF As String
Dim tmpPDFname As String
'//////////////////////////////////////////////////////////////////////////

If nomDuServeur = paramServerSplf Then
    Call MsgBox("Pas d'impression PDF sur ce serveur !!!")
    Exit Sub
End If

    repPDF = "\\docSRV2013\_GROUPS\PUBLIC\_DOSSIERS PARTAGES\RELEVES_CLIENTS\Frais\"
    tmpPDFname = paramIMP_PDF_Path_Temp & "\Releve_.pdf"
    numCli = ""
    increment = 0
    lstSort.Clear
    wRéférence = ""
    For iW = 0 To lstW.ListCount - 1
        If iW Mod 10 = 0 Then
            Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Information frais, sélection en cours: " & iW & " / " & lstW.ListCount)
        End If
        lstW.ListIndex = iW
        s = Split(lstW.Text, vbTab)
        If UBound(s) > 0 Then
            wRéférence = s(0)
            wRELEVENUM = s(1)
            numCli = Retourne_Num_Client(wRELEVENUM)
            blnOk = False
            If Trim(numCli) <> "" Then
                lstSort.AddItem numCli
                increment = comptage_releve_par_client(numCli)
                blnOk = prtZBAGFAC0_A4_Extrait_FraisPDF(wRELEVENUM, lMin, lMax, False, lstErr, lRELEVEREL, wRéférence, blnNewPage)
            End If
            If blnOk Then
                If Dir(paramIMP_PDF_Path_Temp & "\" & numCli & "_" & CStr(increment) & ".pdf", vbNormal) <> "" Then
                    Kill paramIMP_PDF_Path_Temp & "\" & numCli & "_" & CStr(increment) & ".pdf"
                End If
                XPrt.EndDoc
                Call pause_with_events(PDF_WAIT)
                Name tmpPDFname As paramIMP_PDF_Path_Temp & "\" & numCli & "_" & CStr(increment) & ".pdf"
                '//////////////////////////// range le pdf sur le serveur DOCSRV20XX //////////////////////////////////////////////////////////////////////////////////////////////
                If Dir(repPDF & CStr(Year(Now)) & "_" & Mid(CStr(100 + Month(Now)), 2), vbDirectory) = "" Then
                    Call MkDir(repPDF & CStr(Year(Now)) & "_" & Mid(CStr(100 + Month(Now)), 2))
                End If
                FileCopy paramIMP_PDF_Path_Temp & "\" & numCli & "_" & CStr(increment) & ".pdf", repPDF & CStr(Year(Now)) & "_" & Mid(CStr(100 + Month(Now)), 2) & "\" & numCli & "_" & CStr(increment) & ".pdf"
                '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Kill paramIMP_PDF_Path_Temp & "\" & numCli & "_" & CStr(increment) & ".pdf"
            End If
        End If
    Next iW
    lstW.Visible = False
    Call MsgBox("Fin de l'édition PDF 'Information préalable des frais bancaires'...")

End Sub

Private Sub cmdJournal_Devises_Click_xlsManual()
Dim ii As Long
Dim jj As Long
Dim ar() As String
Dim nbSheetRows As Long

'                                               '
Call init_xlsManual
'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
FileCopy paramFolder_Local & "\Modeles\modele_JOURNAL_D.xlsx", paramIMP_PDF_Path_Temp & "\modele_JOURNAL_D.xlsx"
'on charge CE classeur dans Excel
Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_JOURNAL_D.xlsx")
Set wbExcel = appExcelPublic.ActiveWorkbook
With wbExcel
    .Title = "AAJOURNAL_D"
    .Subject = "AAJOURNAL_D"
End With
'                                               '
Me.Enabled = False: Me.MousePointer = vbHourglass
optOptions_SortDevise = True
cbo_Scan " ", cboOptions_Unit
fraContextOptions_Exit
If fgSelect.Rows > 1 Then
    Call mnuSelect_Print_Click_xlsManual(wbExcel)
    ReDim Preserve ar(1 To wbExcel.Sheets.Count - 1)
    jj = 0
    For ii = 1 To wbExcel.Sheets.Count
        If wbExcel.Sheets(ii).Name <> "AAJOURNAL_D" Then
            jj = jj + 1
            ar(jj) = wbExcel.Sheets(ii).Name
            wbExcel.Sheets(ii).Activate
            nbSheetRows = retourne_fin_de_sheet(wbExcel.Sheets(ii))
            Call zoneImpression_xlsManual(wbExcel.Sheets(ii).Name, nbSheetRows, wbExcel.Sheets(ii))
        End If
    Next ii
    wbExcel.Sheets(ar).Select
    Call ActiveSheet.ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
    'sauvegarde du fichier
    Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
    Call wbExcel.Close(True)
    Set wbExcel = Nothing
    Kill paramIMP_PDF_Path_Temp & "\modele_JOURNAL_D.xlsx"
End If
Me.Show
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdJournal_Solde_Click_xlsManual()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdJournal_Solde  : ")
Call cmdJournal_Solde_Print_xlsManual
Call lstErr_AddItem(lstErr, cmdContext, "cmdJournal_Solde : ")
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdJournal_Solde_Print_xlsManual()
Dim xSQL As String, V
Dim wYBIACPT_C_Nb As Long, wYBIACPT_C_NbMax As Long
Dim K As Long, K0 As Long
Dim xCOMPTECOM As String, curX As Currency, curM As Currency
Dim blnOk As Boolean
Dim xText As String
Dim Erreur_Nb As Long
Dim I As Long
Dim currentRow As Long
Dim currentrow2 As Long
Dim currentrow3 As Long
Dim wbExcel As Excel.Workbook
Dim wsExcel As Excel.Worksheet
Dim wsexcel2 As Excel.Worksheet
Dim wsexcel3 As Excel.Worksheet

'                                               '
Call init_xlsManual
'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
FileCopy paramFolder_Local & "\Modeles\modele_SOLDEJ.xlsx", paramIMP_PDF_Path_Temp & "\modele_SOLDEJ.xlsx"
'on charge CE classeur dans Excel
Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_SOLDEJ.xlsx")
Set wbExcel = appExcelPublic.ActiveWorkbook
With wbExcel
    .Title = "SOLDEJ"
    .Subject = "SOLDEJ"
End With
'                                               '
Set wsExcel = wbExcel.Sheets("SOLDEJ")
Set wsexcel2 = wbExcel.Sheets("CONTROLE")
Set wsexcel3 = wbExcel.Sheets("SOLDES")
wsexcel2.Activate
currentRow = 3
currentrow2 = 5
Set rsSab = Nothing
ReDim prtBIA_Compta_Control.arrYBIACPT_C(10001)
wYBIACPT_C_NbMax = 10000: wYBIACPT_C_Nb = 0
xSQL = "select COMPTECOM,COMPTEOBL,COMPTEINT,COMPTEDEV,SOLDEDMO,SOLDECEN from " _
     & paramIBM_Library_SABSPE & ".YBIACPT0 order by COMPTECOM"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    wYBIACPT_C_Nb = wYBIACPT_C_Nb + 1
    If wYBIACPT_C_Nb >= wYBIACPT_C_NbMax Then ReDim Preserve prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb + 1000): wYBIACPT_C_NbMax = wYBIACPT_C_NbMax + 1000
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).COMPTECOM = rsSab("COMPTECOM")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).COMPTEOBL = rsSab("COMPTEOBL")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).COMPTEINT = rsSab("COMPTEINT")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).COMPTEDEV = rsSab("COMPTEDEV")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).SOLDEDMO = rsSab("SOLDEDMO")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).SOLDECEN = rsSab("SOLDECEN") / 1000
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).SOLDEJ_2 = 0
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).MOUVEMMON_DB = 0
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).MOUVEMMON_CR = 0
    rsSab.MoveNext
Loop
prtBIA_Compta_Control.arrYBIACPT_C_Nb = wYBIACPT_C_Nb
prtBIA_Compta_Control.arrYBIACPT_C_NbMax = wYBIACPT_C_NbMax
prtTitleText = " Contrôle des soldes du " & dateImp(YBIATAB0_DATE_CPT_JP1) & " au " & dateImp(YBIATAB0_DATE_CPT_J)
wsExcel.Activate
wsExcel.Cells(1, 4) = prtTitleText
'                                   '
wsexcel2.Activate
xText = "- Contrôle des soldes JOUR = Veille + cumul de mouvements du jour"
wsexcel2.Cells(1, 1) = xText
Erreur_Nb = 0
K0 = 1
xSQL = "select SOLDECOM,SOLDECEN from " _
     & paramIBM_Library_SABSPE & ".ZSOLDE0J_2 order by SOLDECOM"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    xCOMPTECOM = rsSab("SOLDECOM")
    blnOk = False
    For K = K0 To wYBIACPT_C_Nb
        If xCOMPTECOM = prtBIA_Compta_Control.arrYBIACPT_C(K).COMPTECOM Then
            prtBIA_Compta_Control.arrYBIACPT_C(K).SOLDEJ_2 = rsSab("SOLDECEN")
            K0 = K + 1
            blnOk = True
            Exit For
        End If
    Next K
    If Not blnOk Then
        Erreur_Nb = Erreur_Nb + 1
        xText = "? ZSOLDE0J_2 inconnu " & xCOMPTECOM
        Call prtBIA_Compta_Control_Anomalie_xlsManual(xText, currentrow2, wsexcel2)
    End If
    rsSab.MoveNext
Loop
xSQL = "select MOUVEMCOM,MOUVEMMON from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMDTR = " & YBIATAB0_DIBM_CPT_J _
     & " order by MOUVEMCOM"
Set rsSab = cnsab.Execute(xSQL)
K0 = 1
Do While Not rsSab.EOF
    xCOMPTECOM = rsSab("MOUVEMCOM")
    curX = rsSab("MOUVEMMON")
    blnOk = False
    For K = K0 To wYBIACPT_C_Nb
        If xCOMPTECOM = prtBIA_Compta_Control.arrYBIACPT_C(K).COMPTECOM Then
            If curX > 0 Then
                prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_DB = prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_DB + curX
            Else
                prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_CR = prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_CR + curX
            End If
            K0 = K
            blnOk = True
            Exit For
        End If
    Next K
    If Not blnOk Then
        xText = "? YBIAMVTH inconnu " & xCOMPTECOM: prtBIA_Compta_Control_Anomalie xText, False
        Erreur_Nb = Erreur_Nb + 1
    End If
    rsSab.MoveNext
Loop
For K = 1 To wYBIACPT_C_Nb
    curM = prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_DB + prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_CR
    curX = prtBIA_Compta_Control.arrYBIACPT_C(K).SOLDEJ_2
    If curX + curM <> prtBIA_Compta_Control.arrYBIACPT_C(K).SOLDECEN Then
        xText = "? Ecart " & prtBIA_Compta_Control.arrYBIACPT_C(K).COMPTECOM & curX & vbTab & curM & vbTab & prtBIA_Compta_Control.arrYBIACPT_C(K).SOLDECEN
        Call prtBIA_Compta_Control_Anomalie_xlsManual(xText, currentrow2, wsexcel2)
        Erreur_Nb = Erreur_Nb + 1
    End If
Next K
If Erreur_Nb = 0 Then
    xText = wYBIACPT_C_Nb & " comptes contrôlés : aucune erreur de solde détectée"
    currentrow2 = currentrow2 + 1
    Range("A3:I3").Select
    Selection.Copy
    Range("A" & CStr(currentrow2)).Select
    ActiveSheet.Paste
    wsexcel2.Cells(currentrow2, 1) = xText
Else
    xText = wYBIACPT_C_Nb & " comptes contrôlés : !!!!!!! " & Erreur_Nb & " erreur(s) détectée(s) !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    currentrow2 = currentrow2 + 1
    Range("A4:I4").Select
    Selection.Copy
    Range("A" & CStr(currentrow2)).Select
    ActiveSheet.Paste
    wsexcel2.Cells(currentrow2, 1) = xText
End If
prtBIA_Compta_Control.arrDev_Nb = cboDevise.ListCount - 1
ReDim arrDev_B(prtBIA_Compta_Control.arrDev_Nb + 1)
ReDim arrDev_HB(prtBIA_Compta_Control.arrDev_Nb + 1)
For I = 0 To prtBIA_Compta_Control.arrDev_Nb
    cboDevise.ListIndex = I
    prtBIA_Compta_Control.arrDev_B(I + 1).COMPTEDEV = Trim(cboDevise.Text)
    prtBIA_Compta_Control.arrDev_HB(I + 1).COMPTEDEV = Trim(cboDevise.Text)
Next I
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'On supprimer les lignes modèles de la feuille CONTROLE
Rows("2:4").Select
Selection.Delete
currentrow2 = currentrow2 - 3
'On copie les lignes dans la feuille SOLDEJ
wsexcel2.Activate
Range("A1:I" & CStr(currentrow2)).Select
Selection.Copy
wsExcel.Activate
Range("A3:A3").Select
ActiveSheet.Paste
currentRow = currentRow + currentrow2 + 1
wsexcel3.Activate
currentrow3 = 5
Call prtBIA_Compta_Control_Cumul_xlsManual(currentrow3, wsexcel3)
'On copie les lignes dans la feuille SOLDEJ
wsexcel3.Activate
Range("A1:I" & CStr(currentrow3)).Select
Selection.Copy
wsExcel.Activate
Range("A" & CStr(currentRow) & ":A" & CStr(currentRow)).Select
ActiveSheet.Paste
currentRow = currentRow + currentrow3 + 1
Call zoneImpression_xlsManual("SOLDEJ", currentRow, wbExcel.Sheets("SOLDEJ"))
Call ActiveSheet.ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path_Temp & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
Set wsExcel = Nothing
Set wsexcel2 = Nothing
Set wsexcel3 = Nothing
Call wbExcel.Close(True)
Set wbExcel = Nothing
Kill paramIMP_PDF_Path_Temp & "\modele_SOLDEJ.xlsx"
End Sub

Private Sub cmdJournal_Unit_Click_xlsManual()
Dim ii As Long
Dim I As Integer, X As String
Dim currentSheet As Long

Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdJournal_Devises  : ")
For I = 0 To cboOptions_Unit.ListCount - 1
    cboOptions_Unit.ListIndex = I
    If Trim(cboOptions_Unit.Text) <> "" And Trim(cboOptions_Unit.Text) <> "DG" Then 'DR 05/02/2015
        Call lstErr_AddItem(lstErr, cmdContext, cboOptions_Unit.Text & "  : " & fgSelect.Rows - 1)
        optOptions_SortUnit = True
        fraContextOptions_Exit
    End If
Next I
PDF:
Call init_xlsManual
'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
FileCopy paramFolder_Local & "\Modeles\modele_JOURNAL_S.xlsx", paramIMP_PDF_Path_Temp & "\modele_JOURNAL_S.xlsx"
'on charge CE classeur dans Excel
Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_JOURNAL_S.xlsx")
Set wbExcel = appExcelPublic.ActiveWorkbook
With wbExcel
    .Title = "JOURNAL_S"
    .Subject = "JOURNAL_S"
End With
For I = 0 To cboOptions_Unit.ListCount - 1
    cboOptions_Unit.ListIndex = I
    If Trim(cboOptions_Unit.Text) <> "" And Trim(cboOptions_Unit.Text) <> "DG" Then 'DR 05/02/2015
        Call lstErr_AddItem(lstErr, cmdContext, cboOptions_Unit.Text & "  : " & fgSelect.Rows - 1)
        optOptions_SortUnit = True
        fraContextOptions_Exit
        If fgSelect.Rows > 1 Then
            'on écrit systématiquement sur Feuil1 car JOURNAL_S est notre feuille modèle
            wbExcel.Sheets.Add
            currentSheet = indice_nouvelle_feuille(wbExcel)
            'on recopie les 5 premières lignes de Feuil1 vers Feuil2
            wbExcel.Sheets("JOURNAL_S").Select
            Range("1:7").Select
            Selection.Copy
            wbExcel.Sheets(currentSheet).Activate
            Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
            ActiveSheet.Paste
            Range("A8").Select
            X = Table_Unit_SSI("", Trim(cboOptions_Unit.Text))
            If X = "S00" Then X = "S60" 'comptabilité par défaut
            Call frmElpPrt.prtIMP_PDF_NoPaper_Init(X, "BIA-CPT-JAL-SRV", "Archive")
            Call mnuSelect_Print_Click_xlsManual(wbExcel)
        End If
    End If
Next I
Call wbExcel.Close(True)
Set wbExcel = Nothing
Kill paramIMP_PDF_Path_Temp & "\modele_JOURNAL_S.xlsx"
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdPrint_Journal_xlsManual(ByRef wbExcel As Excel.Workbook)
Dim currentSheet As Long

Me.Enabled = False

Msg = Space$(50)
Select Case fgSelect_Sort1
    Case 0:
            currentSheet = indice_nouvelle_feuille(wbExcel)
            Call prtSAB_Compta.prtSAB_Compta_Unit_xlsManual(arrYBIAMVT0(), fgSelect, fgSelect_arrIndex, Me, blnTotal, wbExcel.Sheets(currentSheet))
    Case 5: Call prtSAB_Compta.prtSAB_Compta_Devise_xlsManual(arrYBIAMVT0(), fgSelect, fgSelect_arrIndex, Me, blnTotal, wbExcel)
    Case 6: Call prtSAB_Compta.prtSAB_Compta_Compte_xlsManual(arrYBIAMVT0(), fgSelect, fgSelect_arrIndex, Me, blnTotal, wbExcel)
End Select
Me.Show
Me.Enabled = True

End Sub

Private Sub cmdRelevé_Annuel_Frais_Print_PDF(lMin As String, lMax As String)
Dim iRow As Long, xId As String, X As String
Dim iW As Long
Dim blnOk As Boolean
Dim blnRIB As Boolean, blnTest As Boolean
Dim K As Integer, K2 As Integer
Dim wRéférence As String
Dim wRELEVENUM As String
Dim blnRelevéA4M_Compte As Boolean, mRelevéA4M_Compte As String
'//////////////////////////////////////////////////////////////////////////
Dim numCli As String
Dim increment As Long
Dim repPDF As String
Dim tmpPDFname As String
'//////////////////////////////////////////////////////////////////////////

If nomDuServeur = paramServerSplf Then
    Call MsgBox("Pas d'impression PDF sur ce serveur !!!")
    Exit Sub
End If
    
    repPDF = "\\docSRV2013\_GROUPS\PUBLIC\_DOSSIERS PARTAGES\RELEVES_CLIENTS\Annuel_"
    tmpPDFname = paramIMP_PDF_Path_Temp & "\Releve_.pdf"
    numCli = ""
    increment = 0
    lstSort.Clear
    mRelevéA4M_Compte = Trim(txtRelevéA4M_Compte)
    If Trim(mRelevéA4M_Compte) = "" Then
        blnRelevéA4M_Compte = False
    Else
        blnRelevéA4M_Compte = True
    End If
    'prtBIA_Relevé_Annuel_Frais_OpenX_ResetPDF
    blnOk = True
    wRéférence = ""
    For iW = 0 To lstW.ListCount - 1
        If iW Mod 10 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection en cours: " & iW & " / " & lstW.ListCount)
        lstW.ListIndex = iW
        X = lstW.Text
        K = InStr(1, X, vbTab)
        K2 = Len(X) 'InStr(K + 1, X, vbTab)
        If K > 0 Then
            wRéférence = Mid$(X, 1, 3)
            wRELEVENUM = Trim(Mid$(X, K + 1, K2 - K - 1))
                If blnRelevéA4M_Compte Then
                    blnOk = False
                    If wRELEVENUM = mRelevéA4M_Compte Then
                        blnOk = True
                        blnRelevéA4M_Compte = False
                    End If
                End If
                If blnOk Then
                    numCli = Retourne_Num_Client(wRELEVENUM)
                    If Trim(numCli) <> "" Then
                        lstSort.AddItem numCli
                        increment = comptage_releve_par_client(numCli)
                        Call prtBIA_Relevé_Annuel_Frais_OpenX_ResetPDF
                        prtBIA_Relevé_Annuel_Frais_Extrait wRELEVENUM, lMin, lMax, False, lstErr, "M", wRéférence, blnNewPage
                        XPrt.EndDoc
                        Call pause_with_events(PDF_WAIT)
                        If Dir(paramIMP_PDF_Path_Temp & "\" & wRéférence & "_" & wRELEVENUM & "_" & numCli & "_" & CStr(increment) & ".pdf", vbNormal) <> "" Then
                            Kill paramIMP_PDF_Path_Temp & "\" & wRéférence & "_" & wRELEVENUM & "_" & numCli & "_" & CStr(increment) & ".pdf"
                        End If
                    End If
                    Call pause_with_events(PDF_WAIT)
                    Name tmpPDFname As paramIMP_PDF_Path_Temp & "\" & wRéférence & "_" & wRELEVENUM & "_" & numCli & "_" & CStr(increment) & ".pdf"
                    '//////////////////////////// range le pdf sur le serveur DOCSRV20XX //////////////////////////////////////////////////////////////////////////////////////////////
                    If Dir(repPDF & Left(lMax, 4), vbDirectory) = "" Then
                        Call MkDir(repPDF & Left(lMax, 4))
                    End If
                    If Dir(repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2), vbDirectory) = "" Then
                        Call MkDir(repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2))
                    End If
                    FileCopy paramIMP_PDF_Path_Temp & "\" & wRéférence & "_" & wRELEVENUM & "_" & numCli & "_" & CStr(increment) & ".pdf", repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2) & "\" & wRéférence & "_" & wRELEVENUM & "_" & numCli & "_" & CStr(increment) & ".pdf"
                    Kill paramIMP_PDF_Path_Temp & "\" & wRéférence & "_" & wRELEVENUM & "_" & numCli & "_" & CStr(increment) & ".pdf"
                    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                End If
            End If
    Next iW
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection : " & lstW.ListCount & " / " & fgRelevéA4.Rows - 1)
    Call MsgBox("Fin de l'édition PDF des 'Relevés annuels des frais'...")
    Me.Show
End Sub

Private Sub cmdRelevéA4M_PrintPDF(lRELEVEREL As String, lMin As String, lMax As String)
Dim iRow As Long, xId As String, X As String
Dim iW As Long
Dim blnOk As Boolean
Dim blnRIB As Boolean, blnTest As Boolean
Dim K As Integer, K2 As Integer
Dim wRéférence As String
Dim wRELEVENUM As String
Dim blnRelevéA4M_Compte As Boolean, mRelevéA4M_Compte As String * 20
Dim numResponsable As String
Static nbEssais As Long

If nomDuServeur = paramServerSplf Then
    Call MsgBox("Pas d'impression PDF sur ce serveur !!!")
    Exit Sub
End If

arrECHTAB_K = 0: arrECHTAB_Nb = 0

mRelevéA4M_Compte = Trim(txtRelevéA4M_Compte)
If Trim(mRelevéA4M_Compte) = "" Then
    blnRelevéA4M_Compte = False
Else
    blnRelevéA4M_Compte = True
End If

blnOk = True
wRéférence = ""

'//////////////////////////////////////////////////////////////////////////
Dim numCpt As String
Dim numCli As String
Dim increment As Long
Dim repPDF As String
Dim tmpPDFname As String
'///////////////////////////// impression PDF ensuite /////////////////////////////////////////////////////////////////////////////////////////////////////////
    repPDF = "\\docSRV2013\_GROUPS\PUBLIC\_DOSSIERS PARTAGES\RELEVES_CLIENTS\"
    If lRELEVEREL = "D" Then
        repPDF = "\\docSRV2013\_GROUPS\PUBLIC\_DOSSIERS PARTAGES\RELEVES_CLIENTS\Decadaires\"
    End If
    tmpPDFname = paramIMP_PDF_Path_Temp & "\Releve_.pdf"
    numCli = ""
    increment = 0
    lstSort.Clear
    '//////////////////////////////////////////////////////////////////////////
    'Les numéro client sont stockés dans une liste triée, elle sert à incrémenter les relevés par client
    For iW = 0 To lstW.ListCount - 1
        nbEssais = 0
        If iW Mod 10 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection en cours: " & iW & " / " & lstW.ListCount)
        lstW.ListIndex = iW
        X = lstW.Text
        K = InStr(1, X, vbTab)
        K2 = InStr(K + 1, X, vbTab)
        If K > 0 Then
            wRéférence = Mid$(X, 1, 3)
            wRELEVENUM = Mid$(X, K + 1, K2 - K - 1)
                If blnRelevéA4M_Compte Then
                    blnOk = False
                    If wRELEVENUM = mRelevéA4M_Compte Then
                        blnOk = True
                        blnRelevéA4M_Compte = False
                    End If
                End If
                If blnOk Then
                    numCli = Retourne_Num_Client(wRELEVENUM)
                    If Trim(numCli) <> "" Then
                        lstSort.AddItem numCli
                        increment = comptage_releve_par_client(numCli)
                        numResponsable = Trim(wRéférence)
                        Call prtYBIAMVT0_A4_OpenX_ResetPDF
                        prtYBIAMVT0_A4_Extrait wRELEVENUM, lMin, lMax, False, lstErr, lRELEVEREL, wRéférence, blnNewPage
                        If Dir(paramIMP_PDF_Path_Temp & "\" & numResponsable & "_" & numCli & "_" & Trim(wRELEVENUM) & "_" & CStr(increment) & ".pdf", vbNormal) <> "" Then
                            Kill paramIMP_PDF_Path_Temp & "\" & numResponsable & "_" & numCli & "_" & Trim(wRELEVENUM) & "_" & CStr(increment) & ".pdf"
                        End If
                    End If
                    XPrt.EndDoc
retry:
                    nbEssais = nbEssais + 1
                    If XPrt.Page > 5 Then
                        Call pause_with_events(PDF_WAIT * 2)
                    Else
                        Call pause_with_events(PDF_WAIT)
                    End If
                    On Error Resume Next
                    Name tmpPDFname As paramIMP_PDF_Path_Temp & "\" & numResponsable & "_" & numCli & "_" & Trim(wRELEVENUM) & "_" & CStr(increment) & ".pdf"
                    If Err.Number > 0 Then
                        If Err.Number = 53 And nbEssais < 5 Then
                            GoTo retry
                        End If
                    End If
                    '//////////////////////////// range le pdf sur le serveur DOCSRV20XX //////////////////////////////////////////////////////////////////////////////////////////////
                    If Dir(repPDF & Left(lMax, 4), vbDirectory) = "" Then
                        Call MkDir(repPDF & Left(lMax, 4))
                    End If
                    If lRELEVEREL = "D" Then
                        If Dir(repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2) & "_" & Mid(lMin, 7, 2), vbDirectory) = "" Then
                            Call MkDir(repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2) & "_" & Mid(lMin, 7, 2))
                        End If
                    Else
                        If Dir(repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2), vbDirectory) = "" Then
                            Call MkDir(repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2))
                        End If
                    End If
                    If lRELEVEREL = "D" Then
                        FileCopy paramIMP_PDF_Path_Temp & "\" & numResponsable & "_" & numCli & "_" & Trim(wRELEVENUM) & "_" & CStr(increment) & ".pdf", repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2) & "_" & Mid(lMin, 7, 2) & "\" & numResponsable & "_" & numCli & "_" & Trim(wRELEVENUM) & "_" & CStr(increment) & ".pdf"
                    Else
                        FileCopy paramIMP_PDF_Path_Temp & "\" & numResponsable & "_" & numCli & "_" & Trim(wRELEVENUM) & "_" & CStr(increment) & ".pdf", repPDF & Left(lMax, 4) & "\" & Mid(lMax, 5, 2) & "\" & numResponsable & "_" & numCli & "_" & Trim(wRELEVENUM) & "_" & CStr(increment) & ".pdf"
                    End If
                    Kill paramIMP_PDF_Path_Temp & "\" & numResponsable & "_" & numCli & "_" & Trim(wRELEVENUM) & "_" & CStr(increment) & ".pdf"
                    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                End If
        End If
    Next iW
    If lRELEVEREL = "D" Then
        Call MsgBox("Fin de la conversion des extraits de comptes décadaires en PDF...")
    Else
        Call MsgBox("Fin de la conversion des extraits de comptes mensuels en PDF...")
    End If
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection : " & lstW.ListCount & " / " & fgRelevéA4.Rows - 1)


Me.Show

End Sub

Function comptage_releve_par_client(client As String) As Long
Dim I As Long
Dim nombre As Long

    nombre = 0
    For I = 0 To lstSort.ListCount - 1
        If lstSort.List(I) > client Then
            Exit For
        End If
        If lstSort.List(I) = client Then
            nombre = nombre + 1
        End If
    Next I
    comptage_releve_par_client = nombre
    
End Function

Private Function isCompteOrdinaire(lPCEC As String) As Boolean

    Select Case Mid$(Trim(lPCEC), 1, 5)
      Case "11120", "12120", "12121", "12122", "25111", "25112", "25113", "25114", "25115", "25116", "25117":
        isCompteOrdinaire = True
      Case Else:
        isCompteOrdinaire = False
    End Select

End Function


Private Sub cmdInfoFrais_Print(lRELEVEREL As String, lMin As String, lMax As String)
Dim iW As Long
Dim blnOk As Boolean
Dim s() As String
Dim wRéférence As String
Dim wRELEVENUM As String

    blnOk = True
    wRéférence = ""
    For iW = 0 To lstW.ListCount - 1
        If iW Mod 10 = 0 Then
            Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Information frais, sélection en cours: " & iW & " / " & lstW.ListCount)
        End If
        lstW.ListIndex = iW
        s = Split(lstW.Text, vbTab)
        If UBound(s) > 0 Then
            wRéférence = s(0)
            wRELEVENUM = s(1)
            Call prtZBAGFAC0_A4_Extrait_Frais(wRELEVENUM, lMin, lMax, False, lstErr, lRELEVEREL, wRéférence, blnNewPage)
        End If
    Next iW
    lstW.Visible = False
    Call MsgBox("Fin de l'édition 'Information préalable des frais bancaires'...")
    
End Sub
Private Sub cmdInfoFrais_Select(lRELEVEREL As String, lAMJMin As String)
Dim blnOk As Boolean
Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnTest As Boolean, blnIban As Boolean
Dim xResponsable As String * 3
Dim xRib As String
Dim xSQL As String

    lstW.Clear
    lstW.Visible = False
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZRELEVE0 R , " _
                        & paramIBM_Library_SABSPE & ".YBIACPT0 C" _
     & " where RELEVEREL = '" & lRELEVEREL & "'" _
     & " and   RELEVECOM = C.COMPTECOM" _
     & " and   RELEVEETA = " & currentZMNURUT0.MNURUTETB
    Set rsSab = cnsab.Execute(xSQL)
    Do Until rsSab.EOF
        V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
        V = rsZRELEVE0_GetBuffer(rsSab, xZRELEVE0)
        blnOk = True
        Call fctPCEC_Atribut(xYBIACPT0.COMPTEOBL, xYBIACPT0.COMPTEDEV, blnCptOrdinaire, blnRIB, blnTest, blnIban)
        If Not blnCptOrdinaire Then
            blnOk = False
        End If
        xResponsable = "   "
        If blnOk Then
            If blnRIB Then
                xRib = "RIB"
            Else
                xRib = "RIX"
            End If
            lstW.AddItem xResponsable & vbTab & xZRELEVE0.RELEVECOM & vbTab & xYBIACPT0.COMPTEDEV & xZRELEVE0.RELEVECOM & xRib
        End If
        rsSab.MoveNext
    Loop
    lstW.Visible = True
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Frais bancaires, sélection : " & lstW.ListCount)

End Sub


Private Sub mnuSelect_Print_Click_xlsManual(ByRef wbExcel As Excel.Workbook)
    
    blnTotal = True
    Call cmdPrint_Journal_xlsManual(wbExcel)

End Sub

Private Sub remplit_CLIENACLI(llstCLIENACLI() As Long)
'DR 24/09/2013
Dim sql As String
Dim Nb As Long

    ReDim llstCLIENACLI(1)
    llstCLIENACLI(0) = 0 'la dimension 0 contient le nombre de postes de la table

    sql = "select * from YBIATAB0 where BIATABID = 'SAB_COMPTA' and BIATABK1 = 'REL_SANS_MVT'"
    Set rsMDB = cnMDB.Execute(sql)
    
    Do Until rsMDB.EOF
        Nb = llstCLIENACLI(0) + 1
        ReDim Preserve llstCLIENACLI(Nb)
        If Val(rsMDB("BIATABK2")) > 0 Then
            llstCLIENACLI(Nb) = CLng(rsMDB("BIATABK2"))
            llstCLIENACLI(0) = Nb
        End If
        rsMDB.MoveNext
    Loop
    rsMDB.Close
    Set rsMDB = Nothing
    
End Sub


Public Sub fgRelevéA4_Reset()
fgRelevéA4.Clear
fgRelevéA4_Sort1 = 0: fgRelevéA4_Sort2 = 0
fgRelevéA4_Sort1_Old = -1
fgRelevéA4_RowDisplay = 0: fgRelevéA4_RowClick = 0
fgRelevéA4_arrIndex = 6
blnfgRelevéA4_DisplayLine = False
End Sub


Public Sub fgRelevéA4_Sort()
If fgRelevéA4.Rows > 1 Then
    fgRelevéA4.Row = 1
    fgRelevéA4.RowSel = fgRelevéA4.Rows - 1
    
    If fgRelevéA4_Sort1_Old = fgRelevéA4_Sort1 Then
        If fgRelevéA4_SortAD = 5 Then
            fgRelevéA4_SortAD = 6
        Else
            fgRelevéA4_SortAD = 5
        End If
    Else
        fgRelevéA4_SortAD = 5
    End If
    fgRelevéA4_Sort1_Old = fgRelevéA4_Sort1
    
    fgRelevéA4.Col = fgRelevéA4_Sort1
    fgRelevéA4.ColSel = fgRelevéA4_Sort2
    fgRelevéA4.Sort = fgRelevéA4_SortAD
End If

End Sub

Public Sub fgRelevéA4_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgRelevéA4.Rows - 1
    fgRelevéA4.Row = I
    fgRelevéA4.Col = lK
    X = Format$(Val(fgRelevéA4.Text), "0000000")
    fgRelevéA4.Col = fgRelevéA4_arrIndex - 1
    Select Case lK
        Case 1, 2: fgRelevéA4.Text = X
    End Select
Next I


fgRelevéA4_Sort1 = fgRelevéA4_arrIndex - 1: fgRelevéA4_Sort2 = fgRelevéA4_arrIndex - 1
fgRelevéA4_Sort
End Sub



Public Sub fgRelevéA4_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgRelevéA4.Row

If lRow > 0 And lRow < fgRelevéA4.Rows Then
    fgRelevéA4.Row = lRow
    For I = 0 To fgRelevéA4_arrIndex
        fgRelevéA4.Col = I: fgRelevéA4.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgRelevéA4.Row = mRow
    If fgRelevéA4.Row > 0 Then
        lRow = fgRelevéA4.Row
        lColor_Old = fgRelevéA4.CellBackColor
        For I = 0 To fgRelevéA4_arrIndex
          fgRelevéA4.Col = I: fgRelevéA4.CellBackColor = lColor
        Next I
        fgRelevéA4.Col = 0
    End If
End If

End Sub



Private Sub fgSelect_Display()
Dim Nb As Long, wUnit As String
Dim xSQL As String, xWhere As String
Dim wIndex As Long
SSTab1.Tab = 0

fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False


For wIndex = 1 To arrYBIAMVT0_Nb
    xYBIAMVT0 = arrYBIAMVT0(wIndex)
    If fgSelect_DisplaySelect(wUnit) Then fgSelect_DisplayLine wIndex, wUnit
    
    Nb = Nb + 1
    If Nb Mod 500 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "affichage : " & Nb)
Next wIndex

Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "affichage : " & Nb)

fgSelect.Visible = True
fgSelect_Sort1 = -1
fgSelect_Sort_Options

End Sub

Private Sub cmdRelevé_Annuel_Frais_Select(lAMJMin As String, lAMJMax As String)
Dim xMin7 As String, xMax7 As String
xMax7 = lAMJMax - 19000000: xMin7 = lAMJMin - 19000000

Dim iRow As Long, xId As String
Dim blnOk As Boolean
Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnTest As Boolean, blnIban As Boolean
Dim blnPCEC As Boolean, mPCEC As String, lenPCEC As Integer
Dim xResponsable As String * 3
Dim xRib As String
Dim xSQL As String
Dim wDate As String, xWhere As String
Dim okSuite As Boolean

iRow = 0
lstW.Clear
lstW.Visible = False

'=========================================================================
xSQL = "select distinct mouvemcom ,clienares, clienasrn, clienaeta, C.COMPTEOBL from " & paramIBM_Library_SABSPE & ".ybiamvthf M , " _
                        & paramIBM_Library_SABSPE & ".YBIACPT0 C, " _
                        & paramIBM_Library_SAB & ".ZBASTAB0 Z" _
     & " where mouvemdtr >= " & xMin7 _
     & " and   mouvemdtr <= " & xMax7 _
     & " and   mouvemcom = C.COMPTECOM" _
     & " and C.PLANCOPRO in('CAV','SUC','CBO','DTX','DTT')" _
     & " and clienasrn=''" _
     & " and ((z.bastabnum=5 and CLIENAETA = substr(z.bastabarg, 4, 4) and substr(z.bastabdon, 49, 3) not in ('004','005','006','007'))" _
     & " or clienaeta = 'ASSO')" _
     & " order by CLIENARES, mouvemcom"
     
Set rsSab = cnsab.Execute(xSQL)
Do Until rsSab.EOF
    lstW.AddItem rsSab("CLIENARES") & vbTab & rsSab("MOUVEMCOM")
    rsSab.MoveNext
Loop

Fin:

lstW.Visible = True
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection : " & lstW.ListCount)


End Sub
Private Sub cmdRelevéA4M_Select(lRELEVEREL As String, lAMJMin As String)
Dim iRow As Long, xId As String
Dim blnOk As Boolean
Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnTest As Boolean, blnIban As Boolean
Dim blnPCEC As Boolean, mPCEC As String, lenPCEC As Integer
Dim xResponsable As String * 3
Dim xRib As String
Dim xSQL As String
Dim wDate As String, xWhere As String
'DR 24/09/2013 --------------------
Dim lstCLIENACLI() As Long
Dim blnCLIENACLI As Boolean

If chkListCLIENACLI.Value = 1 Then
    blnCLIENACLI = True
Else
    blnCLIENACLI = False
End If
'DR ------------------------------

mPCEC = Trim(txtRelevéA4M_PCEC)
If mPCEC = "" Then
    blnPCEC = False
Else
    blnPCEC = True
    lenPCEC = Len(mPCEC)
End If

iRow = 0
lstW.Clear
lstW.Visible = False

If chkExtrait_AmjMin = "1" Then
    Call DTPicker_Control(txtExtrait_AmjMin, wDate)
'    xWhere = " And SOLDEDMO <= " & wDate - 19000000 & " And SOLDEDMO > " & YBIATAB0_DATE_CPT_AP1 - 19000000 & " And COMPTEFON <> '4' "
    xWhere = " And SOLDEDMO <= " & wDate - 19000000 & " And COMPTEFON <> '4' "
Else
    xWhere = " And SOLDEDMO >= " & lAMJMin - 19000000
End If

xSQL = "select * from " & paramIBM_Library_SAB & ".ZRELEVE0 R , " _
                        & paramIBM_Library_SABSPE & ".YBIACPT0 C" _
     & " where RELEVEREL = '" & lRELEVEREL & "'" _
     & " and   RELEVECOM = C.COMPTECOM" _
     & xWhere _
     & " and   RELEVEETA = " & currentZMNURUT0.MNURUTETB   '20081021 CAC & " and PLANCOPRO like 'L%'"
     
Set rsSab = cnsab.Execute(xSQL)

'DR 24/09/2013
Call remplit_CLIENACLI(lstCLIENACLI())

Do Until rsSab.EOF
    V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    V = rsZRELEVE0_GetBuffer(rsSab, xZRELEVE0)
    blnOk = True
    
    If blnCLIENACLI = False Then
        If chkExtrait_AmjMin = "1" And xYBIACPT0.SOLDECEN = 0 Then blnOk = False
    End If
    
    If blnPCEC Then
        If mPCEC <> Mid$(xYBIACPT0.COMPTEOBL, 1, lenPCEC) Then blnOk = False
    End If
    
    If chkRelevéA4M_CptOrdinaire = "1" Then
        Call fctPCEC_Atribut(xYBIACPT0.COMPTEOBL, xYBIACPT0.COMPTEDEV, blnCptOrdinaire, blnRIB, blnTest, blnIban)
        If Not blnCptOrdinaire Then blnOk = False
        'DR 24/09/2013 ----------------------------------------------------------------
        'If blnCLIENACLI = True And cmdRelevéA4M.Tag = "CLIC" And chkExtrait_AmjMin = "1" Then
        '    If Not fctCLIENACLI(xYBIACPT0.CLIENACLI, lstCLIENACLI) Then blnOk = False
        'End If
        'DR ---------------------------------------------------------------------------
    End If


                
    If chkRelevéA4M_Responsable = "1" Then
        xResponsable = xYBIACPT0.CLIENARES
    Else
        xResponsable = "   "
    End If

    If blnOk Then
        If blnRIB Then
            xRib = "RIB"
        Else
            xRib = "RIX"
        End If
        lstW.AddItem xResponsable & vbTab & xZRELEVE0.RELEVECOM & vbTab & xYBIACPT0.COMPTEDEV & xZRELEVE0.RELEVECOM & xRib
    End If
    rsSab.MoveNext
Loop

        
lstW.Visible = True
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection : " & lstW.ListCount)


End Sub

Private Sub cmdRelevéA4M_Select_Control()
Dim iRow As Long, xId As String
Dim blnOk As Boolean
Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnTest As Boolean, blnIban As Boolean
Dim blnPCEC As Boolean, mPCEC As String, lenPCEC As Integer
Dim xResponsable As String * 3
Dim xRib As String
Dim xSQL As String
Dim wDate As String, xWhere As String
iRow = 0
lstW.Clear
lstW.Visible = False
Dim rsSabX As New ADODB.Recordset


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 C" _
     & " where COMPTEFON <> '4'" _
     & " order by C.COMPTECOM"
     
Set rsSab = cnsab.Execute(xSQL)


Do Until rsSab.EOF
    V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    blnOk = True
    
        Call fctPCEC_Atribut(xYBIACPT0.COMPTEOBL, xYBIACPT0.COMPTEDEV, blnCptOrdinaire, blnRIB, blnTest, blnIban)
        If Not blnCptOrdinaire Then
            blnOk = False
        Else
            xSQL = "select RELEVEREL from " & paramIBM_Library_SAB & ".ZRELEVE0" _
                 & " where RELEVECOM = '" & xYBIACPT0.COMPTECOM & "'"
                 
            Set rsSabX = cnsab.Execute(xSQL)
        
            Do Until rsSabX.EOF
                If rsSabX("RELEVEREL") = "M" Or rsSabX("RELEVEREL") = "W" Or rsSabX("RELEVEREL") = "D" _
                Or rsSabX("RELEVEREL") = "A" Then
                    blnOk = False
                    Exit Do
                End If
                rsSabX.MoveNext
            Loop
            
        End If
                

    If blnOk Then
        lstW.AddItem xYBIACPT0.COMPTECOM & vbTab & xYBIACPT0.COMPTEDEV & " " & xZRELEVE0.RELEVEREL
        Debug.Print xResponsable & vbTab & xZRELEVE0.RELEVECOM & vbTab & xYBIACPT0.COMPTEDEV & xZRELEVE0.RELEVECOM & xZRELEVE0.RELEVEREL
    End If
    rsSab.MoveNext
Loop

        
lstW.Visible = True
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection : " & lstW.ListCount)

Call MsgBox("Pour info : liste des comptes non D A M W", vbInformation, "Extraits mensuels")
End Sub


Private Sub cmdRelevéA4G_Select()

End Sub

Private Sub cmdRelevéA4M_Print(lRELEVEREL As String, lMin As String, lMax As String)
Dim iRow As Long, xId As String, X As String
Dim iW As Long
Dim blnOk As Boolean
Dim blnRIB As Boolean, blnTest As Boolean
Dim K As Integer, K2 As Integer
Dim wRéférence As String
Dim wRELEVENUM As String
Dim blnRelevéA4M_Compte As Boolean, mRelevéA4M_Compte As String * 20

arrECHTAB_K = 0: arrECHTAB_Nb = 0

mRelevéA4M_Compte = Trim(txtRelevéA4M_Compte)
If Trim(mRelevéA4M_Compte) = "" Then
    blnRelevéA4M_Compte = False
Else
    blnRelevéA4M_Compte = True
End If

blnOk = True
wRéférence = ""

    For iW = 0 To lstW.ListCount - 1
        If iW Mod 10 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection en cours: " & iW & " / " & lstW.ListCount)
        lstW.ListIndex = iW
        X = lstW.Text
        K = InStr(1, X, vbTab)
        K2 = InStr(K + 1, X, vbTab)
        If K > 0 Then
            wRéférence = Mid$(X, 1, 3)
            wRELEVENUM = Mid$(X, K + 1, K2 - K - 1)
                If blnRelevéA4M_Compte Then
                    blnOk = False
                    If wRELEVENUM = mRelevéA4M_Compte Then
                        blnOk = True
                        blnRelevéA4M_Compte = False
                    End If
                End If
                If blnOk Then
                    '//////////////////////////// impression papier //////////////////////////////////////////////////////////////////////////////////////////////
                    Call prtYBIAMVT0_A4_OpenX_Reset
                    prtYBIAMVT0_A4_Extrait wRELEVENUM, lMin, lMax, False, lstErr, lRELEVEREL, wRéférence, blnNewPage
                    Call prtYBIAMVT0_A4_Close
                    If chkRelevéA4M_Rib Then
                        K = InStr(1, X, "RIB")
                        If K > 0 Then
                                prtRIB_A4 wRELEVENUM
                                XPrt.NewPage
                                prtYBIAMVT0_A4_OpenX_Reset
                            End If
                        End If
                   End If
                End If
    Next iW
If lRELEVEREL = "D" Then
    Call MsgBox("Fin de l'édition des extraits de comptes décadaires...")
Else
    Call MsgBox("Fin de l'édition des extraits de comptes mensuels...")
End If
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection : " & lstW.ListCount & " / " & fgRelevéA4.Rows - 1)

prtYBIAMVT0_A4_Close
Me.Show
End Sub
Private Sub cmdRelevé_Annuel_Frais_Print(lMin As String, lMax As String)
Dim iRow As Long, xId As String, X As String
Dim iW As Long
Dim blnOk As Boolean
Dim blnRIB As Boolean, blnTest As Boolean
Dim K As Integer, K2 As Integer
Dim wRéférence As String
Dim wRELEVENUM As String
Dim blnRelevéA4M_Compte As Boolean, mRelevéA4M_Compte As String


mRelevéA4M_Compte = Trim(txtRelevéA4M_Compte)
If Trim(mRelevéA4M_Compte) = "" Then
    blnRelevéA4M_Compte = False
Else
    blnRelevéA4M_Compte = True
End If

'prtBIA_Relevé_Annuel_Frais_OpenX
prtBIA_Relevé_Annuel_Frais_OpenX_Reset
blnOk = True
wRéférence = ""
For iW = 0 To lstW.ListCount - 1

    If iW Mod 10 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection en cours: " & iW & " / " & lstW.ListCount)
    
    lstW.ListIndex = iW
    X = lstW.Text
    K = InStr(1, X, vbTab)
    K2 = Len(X) 'InStr(K + 1, X, vbTab)
    If K > 0 Then
        wRéférence = Mid$(X, 1, 3)
        wRELEVENUM = Trim(Mid$(X, K + 1, K2 - K - 1))
            If blnRelevéA4M_Compte Then
                blnOk = False
                If wRELEVENUM = mRelevéA4M_Compte Then
                    blnOk = True
                    blnRelevéA4M_Compte = False
                End If
            End If
            
                    
            If blnOk Then
                
                prtBIA_Relevé_Annuel_Frais_Extrait wRELEVENUM, lMin, lMax, False, lstErr, "M", wRéférence, blnNewPage
    
            End If
        End If
   
Next iW
        
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Relevé, selection : " & lstW.ListCount & " / " & fgRelevéA4.Rows - 1)


prtBIA_Relevé_Annuel_Frais_Close
Call MsgBox("Fin de l'édition 'Relevé annuel des frais '...")
Me.Show
End Sub

Private Sub cmdRelevéA4_Print(lMin As String, lMax As String)
arrECHTAB_K = 0: arrECHTAB_Nb = 0

prtYBIAMVT0_A4_OpenX
prtYBIAMVT0_A4_Extrait meYBIAMVT0.MOUVEMCOM, lMin, lMax, False, lstErr, "*", "", blnNewPage
prtYBIAMVT0_A4_Close

Me.Show

End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long, lUnit As String)
Dim X As String
On Error Resume Next

fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = dateIBM10(xYBIAMVT0.MOUVEMDTR, False)
fgSelect.Col = 1: fgSelect.Text = lUnit
fgSelect.Col = 2: fgSelect.Text = dateIBM10(xYBIAMVT0.MOUVEMDCO, False)
fgSelect.Col = 3: fgSelect.Text = xYBIAMVT0.MOUVEMOPE & " " & Format$(xYBIAMVT0.MOUVEMNUM, "000000000")
fgSelect.Col = 4: fgSelect.Text = Format$(xYBIAMVT0.MOUVEMPIE, "000000000") & " " & Format$(xYBIAMVT0.MOUVEMECR, "0000000")
fgSelect.Col = 5: fgSelect.Text = xYBIAMVT0.COMPTEDEV
fgSelect.Col = 6: fgSelect.Text = xYBIAMVT0.MOUVEMCOM
fgSelect.Col = 7: fgSelect.Text = xYBIAMVT0.COMPTEINT
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub


Public Function fgSelect_DisplaySelect(lUnit As String) As Boolean
Dim xService As String
Dim xMOUVEMOPE As String, xMOUVEMNUM As Long
On Error Resume Next

fgSelect_DisplaySelect = False

xService = xYBIAMVT0.MOUVEMSER & xYBIAMVT0.MOUVEMSSE

If optOptions_SelectOD Then
    If Mid$(xYBIAMVT0.MOUVEMOPE, 1, 1) = "*" Then
        lUnit = Table_Ope_Unit(xService & xYBIAMVT0.MOUVEMOPE)
        If mUnit = "" Or mUnit = lUnit Then fgSelect_DisplaySelect = True
    End If
Else

       lUnit = "CPXX"
       If xService <> "0000" Then lUnit = Table_Ope_Unit(xService) ' par défaut
       
        ' Cas particulier RDE / RDI => SOBI  ou SOBF
        xMOUVEMOPE = xYBIAMVT0.MOUVEMOPE
        If xMOUVEMOPE = "RDE" Or xMOUVEMOPE = "RDI" Then
            xMOUVEMNUM = xYBIAMVT0.MOUVEMNUM
            lUnit = Table_Ope_Unit_RDE(xMOUVEMOPE, xMOUVEMNUM, cnsab, rsSab)
        End If
       
       If lUnit = "CPXX" Or lUnit = xService Then
            lUnit = Table_Ope_Unit(xService & xYBIAMVT0.MOUVEMOPE)
        End If
        
       If lUnit = xService Then lUnit = "CPT"

       If mUnit = "" Or mUnit = lUnit Then fgSelect_DisplaySelect = True
End If

End Function



Public Sub fgSelect_Sort()
fgSelect.Visible = False
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    If fgSelect_Sort1_Old = fgSelect_Sort1 Then
        If fgSelect_SortAD = 5 Then
            fgSelect_SortAD = 6
        Else
            fgSelect_SortAD = 5
        End If
    Else
        fgSelect_SortAD = 5
    End If
    fgSelect_Sort1_Old = fgSelect_Sort1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = fgSelect_SortAD
End If
fgSelect.Visible = True
End Sub
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
Dim mK As Integer
mK = lK
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    Select Case lK
        Case 5:
            fgSelect.Col = 0: X = fgSelect.Text
            fgSelect.Col = 2: X = X & fgSelect.Text
            fgSelect.Col = 5: X = X & fgSelect.Text
            fgSelect.Col = 3: X = X & fgSelect.Text
            fgSelect.Col = 4: X = X & fgSelect.Text
         Case 6
            fgSelect.Col = 6: X = fgSelect.Text
            fgSelect.Col = 5: X = X & fgSelect.Text
            fgSelect.Col = 3: X = X & fgSelect.Text
            fgSelect.Col = 4: X = X & fgSelect.Text
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
fgSelect_Sort1 = mK: fgSelect_Sort2 = mK
End Sub



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init("SAB_Compta", SAB_Compta_Aut)

'blnSetfocus = True
Form_Init

PDF_WAIT = Retourne_WAIT_PDF

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "TEST":
                Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-CPT-CTL-BHB", "Archive")
                blnAuto = True
                SSTab1.Tab = 2
                cmdJournal_Solde_Click
                'Call frmElpPrt.prtIMP_PDF_NoPaper_Print("INFO")

    Case "@SOLDEJ":
                Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-CPT-CTL-BHB", "Archive")
                blnAuto = True
                SSTab1.Tab = 2
                If xlsManual Then
                    Call cmdJournal_Solde_Click_xlsManual
                Else
                    cmdJournal_Solde_Click
                End If
    Case "@JOURNAL_D":
                Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-CPT-JAL-DEV", "Archive")
                blnAuto = True
                SSTab1.Tab = 2
                If xlsManual Then
                    Call cmdJournal_Devises_Click_xlsManual
                Else
                    cmdJournal_Devises_Click
                End If
    Case "@JOURNAL_S":
                blnAuto = True
                SSTab1.Tab = 2
                If xlsManual Then
                    Call cmdJournal_Unit_Click_xlsManual
                Else
                    cmdJournal_Unit_Click
                End If
    Case "@MT950":
                 If paramEnvironnement = constProduction Then
                    meUnit.Id = "INFO"
                    Table_Unit meUnit
                    Printer_Set meUnit.Printer
                End If
                blnAuto = True
                SSTab1.Tab = 2
                chkRelevéA4W_Loro = "1"
                chkRelevéA4W_Nostro = "1"
                chkRelevéA4W_Update = "1"
                cmdRelevéA4W_Click
                'Call MsgBox("Fin du traitement...")
    Case "@MT900":
                blnAuto = True
                SSTab1.Tab = 2
                cmdMT900_Click
    Case Else: blnAuto = False
End Select
If blnAuto Then
    Unload Me
End If

End Sub


Public Sub Form_Init()
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True
If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmSAB_YBIAMVT0.param_init"
    fraTab0.Enabled = False
End If

    blnControl = False
    fgSelect_FormatString = fgSelect.FormatString
    fgSelect.Enabled = True
  '  fraTAb1.Visible = SAB_Compta_Aut.Xspécial
    cmdReset
Me.Enabled = True

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim X As String
On Error Resume Next

blnControl = False
blnError = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""

fraInfo.Visible = SAB_Compta_Aut.Xspécial
fraInfo_Jour.Visible = SAB_Compta_Aut.Xspécial
'''fraRelevéA4W.Enabled = False


cmdReset_Date

txtRelevéA4_Compte = ""
txtRelevéA4M_PCEC = ""
'fraInfo_Jour.Enabled = SAB_Compta_Aut.Xspécial

ReDim specialMVT(100): specialMVT_nb = 0: specialMVT_Max = 100
Call DTPicker_Set(txtExtrait_AmjMin, YBIATAB0_DATE_CPT_MP1)
txtExtrait_AmjMin.Enabled = SAB_Compta_Aut.Xspécial
chkExtrait_AmjMin.Value = "0"
chkExtrait_AmjMin.Enabled = SAB_Compta_Aut.Xspécial

SSTab1.Tab = 0
cmdSelect_Click

blnControl = True
End Sub



Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String

Dim V
param_Init = Null
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "BIA.mdb : table : Unit")

Call cbo_LoadId_K2("Unit", "", cboOptions_Unit)
cboOptions_Unit.AddItem ""

Call rsYBIATAB0_cboK2("DEVISE", "ISO", cboDevise)

End Function
Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 0 To fgSelect_arrIndex
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Public Sub zoneImpression_xlsManual(lFct As String, nbRows As Long, wsheet As Excel.Worksheet)

    Call init_TypePagesetup
    If nbRows > 0 Then
        If lFct = "SOLDEJ" Then
            wsheet.Activate
            wsheet.Range("A1:I" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$I$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSAB_Compta_Contrôle   &D &T  BIA_INFO"
            zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
            zoneImpressionPagesetup.Orientation = xlLandscape
            zoneImpressionPagesetup.Zoom = 85
        Else
            wsheet.Activate
            wsheet.Range("A1:K" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$K$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSAB_Compta   &D &T  BIA_INFO"
            zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
            zoneImpressionPagesetup.Orientation = xlLandscape
            zoneImpressionPagesetup.Zoom = 85
        End If
    End If
    Call SetTypePageSetup(wsheet)
    
End Sub

Private Sub cboOptions_Unit_Click()
mUnit = Trim(cboOptions_Unit.Text)
End Sub


Private Sub chkRelevéA4W_Confirmation_Click()
If chkRelevéA4W_Confirmation = "1" Then
    txtRelevéA4W_Amj.Enabled = True
Else
    txtRelevéA4W_Amj.Enabled = False
End If
End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdContextOptions_Ok_Click()
On Error Resume Next
fraContextOptions.Visible = False
cmdSelect_Click

End Sub

Private Sub cmdCOSI_Click()
Dim appli As String
Dim xName As String
Dim repertoirSortie As String
Dim fic As Long
Dim ligIn As String
Dim s() As String
Dim s2() As String
Dim ok As Boolean
Dim okFic As Boolean
Dim ret As Long
Dim repIni As String

    ok = False
    V = rsElpTable_Read("Server", "Application", "COSI_INI", xName, repIni)
    If IsNull(V) Then
        repIni = "\\" & repIni
        If IsNull(V) Then
            V = rsElpTable_Read("Server", "Application", "COSI", xName, appli)
            If IsNull(V) Then
                If Trim(appli) <> "" Then
                    appli = "\\" & appli
                    'Trouver le répertoire de production des fichiers xml
                    fic = FreeFile
                    Open repIni For Input As #fic
                    Do While Not EOF(fic)
                        Line Input #fic, ligIn
                        If Trim(ligIn) <> "" Then
                            If InStr(UCase(ligIn), "KEY=""REPERTOIRSORTIE""") > 0 Then
                                s = Split(ligIn, "=")
                                If UBound(s) = 2 Then
                                    repertoirSortie = Trim(s(2))
                                    s2 = Split(repertoirSortie, Chr(34))
                                    repertoirSortie = s2(1)
                                    ok = True
                                    Exit Do
                                End If
                            End If
                        End If
                    Loop
                    Close #fic
                End If
                If ok Then
                    'Voir si un traitement n'a pas déjà été lancé ce mois-ci
                    okFic = False
                    xName = CStr(Year(Now)) & "_" & Mid(CStr(Month(Now) + 100), 2) & "_" & Mid(CStr(Day(Now) + 100), 2) & "_"
                    ligIn = Dir(repertoirSortie & "\*.*", vbNormal)
                    Do While ligIn <> ""
                        If InStr(1, ligIn, xName) = 1 And (InStr(ligIn, "ListeRetraits") > 0 Or InStr(ligIn, "ListeVersements") > 0) Then
                            okFic = True
                            Exit Do
                        End If
                        ligIn = Dir
                    Loop
                    If okFic Then
                        ret = MsgBox("Un fichier a déjà été produit pour aujourd'hui, voulez-vous réellement produire un autre fichier ?", vbYesNo)
                        If ret = vbYes Then
                            Call Shell_Exe(appli)
                            MsgBox "L'application COSI a été lancée !"
                        End If
                    Else
                        Call Shell_Exe(appli)
                        MsgBox "L'application COSI a été lancée !"
                    End If
                End If
            End If
        End If
    End If
    If Not ok Then
        MsgBox "Impossible d'atteindre l'application COSI !"
    End If


End Sub

Private Sub cmdECH_Ok_Click()

If chkECH_FTP = "1" Then Call Shell_FTP(paramYBase_DataF & "YECHEDIW.txt", paramIBM_Library_SABSPE, "YECHEDIW", True, False)
If chkECH_Print = "1" Then
    Call MsgBox("Transfert terminé ? .... ", vbInformation, "Impression ECH ")
    prtSAB_Echelles_Monitor paramYBase_DataF & "YECHEDIW" & paramYBase_Data_ExtensionP
End If
End Sub

Private Sub cmdETAFI_Ok_Click()
Dim wFileName As String

X = MsgBox("Après traitement I5A7 : YETAFI0_B," & Asc10_13 & "création du fichier C:\Temp\ETAFI.txt", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
If X = vbNo Then
    Exit Sub
End If

Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "ETAFI : début .... ")

Call lstErr_Clear(lstErr, cmdContext, "C:\TEMP\ETAFI.txt")
Me.Enabled = True: Me.MousePointer = 0

prtBIA_ETAFI_Monitor wFileName

Call lstErr_AddItem(lstErr, cmdContext, "ETAFI : fin ")
Me.Show


End Sub

Private Sub cmdExtract_Click()
Dim borneInferieure As Long
Dim xSQL As String
Dim rsSabnew As ADODB.Recordset
Dim ndate As Date
Dim aCli As String
Dim fic As Long
Dim ligOut As String
Dim indice As Long
Dim s() As String

    Call MsgBox("Les adresses seront extraites dans le fichier c:\temp\AdressesClients.csv...")
    MousePointer = vbHourglass
    ndate = DateAdd("m", -2, Now) 'moins 2 mois
    borneInferieure = CLng(Year(ndate) & Mid(100 + Month(ndate), 2) & Mid(100 + Day(ndate), 2))
    lstW.Clear
    lstW.Visible = False
    xSQL = "select distinct clienacli, clienara1, relevenum, relevetyp, releveadr, compteclo from " & paramIBM_Library_SAB & ".ZRELEVE0 R, " & paramIBM_Library_SABSPE & ".YBIACPT0 C"
    xSQL = xSQL & " where   RELEVECOM = C.COMPTECOM"
    xSQL = xSQL & " and   RELEVEETA = " & currentZMNURUT0.MNURUTETB
    xSQL = xSQL & " and substr(clienares,1,1) <>'X' order by clienara1"
    Set rsSabnew = cnsab.Execute(xSQL)
    aCli = ""
    Do Until rsSabnew.EOF
        If (19000000 + CLng(rsSabnew("compteclo"))) < borneInferieure Or (19000000 + CLng(rsSabnew("compteclo"))) = 0 Then
            If (Trim(rsSabnew("clienara1")) & ";" & Trim(rsSabnew("clienacli")) & ";" & Trim(rsSabnew("relevenum")) & ";" & Trim(rsSabnew("relevetyp")) & ";" & Trim(rsSabnew("releveadr"))) <> aCli Then
                lstW.AddItem Trim(rsSabnew("clienara1")) & ";" & Trim(rsSabnew("clienacli")) & ";" & Trim(rsSabnew("relevenum")) & ";" & Trim(rsSabnew("relevetyp")) & ";" & Trim(rsSabnew("releveadr"))
                aCli = Trim(rsSabnew("clienara1")) & ";" & Trim(rsSabnew("clienacli")) & ";" & Trim(rsSabnew("relevenum")) & ";" & Trim(rsSabnew("relevetyp")) & ";" & Trim(rsSabnew("releveadr"))
            End If
        End If
        rsSabnew.MoveNext
    Loop
    lstW.Visible = True
    
    If Dir("c:\temp\AdressesClients.csv") <> "" Then
        Kill "c:\temp\AdressesClients.csv"
    End If
    fic = FreeFile
    Open "c:\temp\AdressesClients.csv" For Output As #fic
    ligOut = "client;code;adresse1;adresse2;adresse3;code postal;ville;pays"
    Print #fic, ligOut
    aCli = ""
    lstSort.Clear
    For indice = 0 To lstW.ListCount - 1
        s = Split(lstW.List(indice), ";")
        If UBound(s) >= 4 Then
            Call rsZADRESS0_Init(meZADRESS0)
            meZADRESS0.ADRESSNUM = s(2)
            meZADRESS0.ADRESSTYP = s(3)
            meZADRESS0.ADRESSCOA = s(4)
            If meZADRESS0.ADRESSTYP = "1" Then
                Call rsZADRESS0_Client(meZADRESS0)
            Else
                Call rsZADRESS0_Compte(meZADRESS0)
            End If
            If Trim(meZADRESS0.ADRESSRA1) <> "" Then
                ligOut = s(1) & ";" & Trim(meZADRESS0.ADRESSRA1) & " " & Trim(meZADRESS0.ADRESSRA2) & ";" & Trim(meZADRESS0.ADRESSAD1) & ";"
                ligOut = ligOut & Trim(meZADRESS0.ADRESSAD2) & ";" & Trim(meZADRESS0.ADRESSAD3) & ";" & meZADRESS0.ADRESSCOP & ";" & meZADRESS0.ADRESSVIL & ";" & meZADRESS0.ADRESSPAY
                If ligOut <> aCli Then
                    lstSort.AddItem (ligOut)
                    aCli = ligOut
                End If
            End If
        End If
    Next indice
    aCli = ""
    For indice = 0 To lstSort.ListCount - 1
        If lstSort.List(indice) <> aCli Then
            Print #fic, lstSort.List(indice)
            aCli = lstSort.List(indice)
        End If
    Next indice
    Close #fic
    lstW.Clear
    lstW.Visible = False
    lstSort.Clear
    MousePointer = vbDefault
    Call MsgBox("Fin de l'extraction...")

End Sub

Private Sub cmdExtractAdresses_Click()
Dim borneInferieure As Long
Dim xSQL As String
Dim rsSabnew As ADODB.Recordset
Dim ndate As Date
Dim aCli As String
Dim fic As Long
Dim ligOut As String
Dim indice As Long
Dim s() As String

    Call MsgBox("Les adresses seront extraites dans le fichier c:\temp\AdressesClients.csv...")
    MousePointer = vbHourglass
    ndate = DateAdd("m", -2, Now) 'moins 2 mois
    borneInferieure = CLng(Year(ndate) & Mid(100 + Month(ndate), 2) & Mid(100 + Day(ndate), 2))
    lstW.Clear
    lstW.Visible = False
    xSQL = "select distinct C.COMPTECOM, C.COMPTEOBL, clienacli, clienara1, relevenum, relevetyp, releveadr, compteclo from " & paramIBM_Library_SAB & ".ZRELEVE0 R, " & paramIBM_Library_SABSPE & ".YBIACPT0 C"
    xSQL = xSQL & " where   RELEVECOM = C.COMPTECOM"
    xSQL = xSQL & " and   RELEVEETA = " & currentZMNURUT0.MNURUTETB
    xSQL = xSQL & " and substr(clienares,1,1) <>'X' order by clienara1"
    Set rsSabnew = cnsab.Execute(xSQL)
    aCli = ""
    Do Until rsSabnew.EOF
        If (19000000 + CLng(rsSabnew("compteclo"))) < borneInferieure Or (19000000 + CLng(rsSabnew("compteclo"))) = 0 Then
            If isCompteOrdinaire(Trim(rsSabnew("COMPTEOBL"))) Then
                If (Trim(rsSabnew("COMPTECOM")) & ";" & Trim(rsSabnew("clienara1")) & ";" & Trim(rsSabnew("clienacli")) & ";" & Trim(rsSabnew("relevenum")) & ";" & Trim(rsSabnew("relevetyp")) & ";" & Trim(rsSabnew("releveadr"))) <> aCli Then
                    lstW.AddItem Trim(rsSabnew("COMPTECOM")) & ";" & Trim(rsSabnew("clienara1")) & ";" & Trim(rsSabnew("clienacli")) & ";" & Trim(rsSabnew("relevenum")) & ";" & Trim(rsSabnew("relevetyp")) & ";" & Trim(rsSabnew("releveadr"))
                    aCli = Trim(rsSabnew("COMPTECOM")) & ";" & Trim(rsSabnew("clienara1")) & ";" & Trim(rsSabnew("clienacli")) & ";" & Trim(rsSabnew("relevenum")) & ";" & Trim(rsSabnew("relevetyp")) & ";" & Trim(rsSabnew("releveadr"))
                End If
            End If
        End If
        rsSabnew.MoveNext
    Loop
    lstW.Visible = True
    
    If Dir("c:\temp\AdressesClients.csv") <> "" Then
        Kill "c:\temp\AdressesClients.csv"
    End If
    fic = FreeFile
    Open "c:\temp\AdressesClients.csv" For Output As #fic
    ligOut = "code;compte;client;adresse1;adresse2;adresse3;code postal;ville;pays"
    Print #fic, ligOut
    aCli = ""
    lstSort.Clear
    For indice = 0 To lstW.ListCount - 1
        s = Split(lstW.List(indice), ";")
        If UBound(s) >= 5 Then
            Call rsZADRESS0_Init(meZADRESS0)
            meZADRESS0.ADRESSNUM = s(3)
            meZADRESS0.ADRESSTYP = s(4)
            meZADRESS0.ADRESSCOA = s(5)
            If meZADRESS0.ADRESSTYP = "1" Then
                Call rsZADRESS0_Client(meZADRESS0)
            Else
                Call rsZADRESS0_Compte(meZADRESS0)
            End If
            If Trim(meZADRESS0.ADRESSRA1) <> "" Then
                ligOut = s(2) & ";" & s(0) & ";" & Trim(meZADRESS0.ADRESSRA1) & " " & Trim(meZADRESS0.ADRESSRA2) & ";" & Trim(meZADRESS0.ADRESSAD1) & ";"
                ligOut = ligOut & Trim(meZADRESS0.ADRESSAD2) & ";" & Trim(meZADRESS0.ADRESSAD3) & ";" & meZADRESS0.ADRESSCOP & ";" & meZADRESS0.ADRESSVIL & ";" & meZADRESS0.ADRESSPAY
                If ligOut <> aCli Then
                    lstSort.AddItem (ligOut)
                    aCli = ligOut
                End If
            End If
        End If
    Next indice
    aCli = ""
    For indice = 0 To lstSort.ListCount - 1
        If lstSort.List(indice) <> aCli Then
            Print #fic, lstSort.List(indice)
            aCli = lstSort.List(indice)
        End If
    Next indice
    Close #fic
    lstW.Clear
    lstW.Visible = False
    lstSort.Clear
    MousePointer = vbDefault
    Call MsgBox("Fin de l'extraction...")
    
End Sub


Private Sub cmdFraisBancaires_Click()
Dim xMin As String, xMax As String

    Me.Enabled = False: Me.MousePointer = vbHourglass
    Call lstErr_Clear(lstErr, cmdContext, "cmdFraisBancaires  : début")
    Call DTPicker_Control(txtRelevéA4_AmjMin, xMin)
    xMax = dateFinDeMois(xMin)
    'X = MsgBox("Factures validées du " & dateImp10(xMin) & " au " & dateImp10(xMax) & ". Voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Information préalable - Frais bancaires :")
    X = MsgBox("Factures validées à partir du " & dateImp10(DSys) & ". Voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Information préalable - Frais bancaires :")
    If X = vbNo Then GoTo Exit_sub
    Call cmdInfoFrais_Select("M", xMin)
    If lstW.ListCount > 0 Then
        Call cmdInfoFrais_Print("M", xMin, xMax)
    End If
    Call lstErr_AddItem(lstErr, cmdContext, "cmdFraisBancaires : Fin ")
Exit_sub:
    Me.Show
    Me.Enabled = True: Me.MousePointer = 0
    
End Sub

Private Sub cmdFraisBancairesPDF_Click()
Dim xMin As String, xMax As String

    Me.Enabled = False: Me.MousePointer = vbHourglass
    Call lstErr_Clear(lstErr, cmdContext, "cmdFraisBancaires  : début")
    Call DTPicker_Control(txtRelevéA4_AmjMin, xMin)
    xMax = dateFinDeMois(xMin)
    'X = MsgBox("Factures validées du " & dateImp10(xMin) & " au " & dateImp10(xMax) & ". Voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Information préalable - Frais bancaires :")
    X = MsgBox("Factures validées à partir du " & dateImp10(DSys) & "." & vbCrLf & "Avez-vous sélectionné l'imprimante PDF et voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Information préalable - Frais bancaires :")
    If X = vbNo Then GoTo Exit_sub
    Call cmdInfoFrais_Select("M", xMin)
    If lstW.ListCount > 0 Then
        Call cmdInfoFrais_PrintPDF("M", xMin, xMax)
    End If
    Call lstErr_AddItem(lstErr, cmdContext, "cmdFraisBancaires : Fin ")
Exit_sub:
    Me.Show
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdImpRepertoire_Click()
Dim nicSource As String
Dim nicCommandes As String
Dim nici As Long
Dim nicShellProgram As String
Dim resultShell As Long

    nicShellProgram = "c:\program files\adobe\reader 11.0\reader\acroRd32.exe"
    
    'nicSource = "\\DOCSRV2013\_GROUPS\PUBLIC\_DOSSIERS PARTAGES\RELEVES_CLIENTS\Annuel_2017\12\"
    Load dlgRep
    dlgRep.Tag = ""
    Call dlgRep.Show(vbModal)
    If dlgRep.Tag = "" Then
        Exit Sub
    End If
    nicSource = dlgRep.Tag
    If Right(nicSource, 1) <> "\" Then
        nicSource = nicSource & "\"
    End If
    'Le fichier de commandes se trouve dans c:\BIASRV ou  d:\BIASRV
    nicCommandes = "c:\biasrv\impressionPDF_Annuel.cmd"
    If Dir(nicCommandes) = "" Then
        nicCommandes = "d:\biasrv\impressionPDF_Annuel.cmd"
    End If
    lstW.Clear
    frmEdition.filDoc.PATH = nicSource
    frmEdition.filDoc.Pattern = "*.pdf"
    For nici = 0 To frmEdition.filDoc.ListCount - 1
        lstW.AddItem frmEdition.filDoc.List(nici)
    Next nici
    DoEvents
    If lstW.ListCount > 0 Then
        For nici = 0 To lstW.ListCount - 1
            lstW.ListIndex = nici
            resultShell = Shell(nicShellProgram & " /n /s /h /t " & nicSource & lstW.List(nici))
            DoEvents
        Next nici
    End If
    
    Call MsgBox("Fin de l'impression du répertoire...")
    
End Sub

Private Sub cmdJournal_Devises_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdJournal_Devises  : ")

optOptions_SortDevise = True
cbo_Scan " ", cboOptions_Unit
fraContextOptions_Exit

If fgSelect.Rows > 1 Then
    'DR 08/10/2019
    If InStr(Printer.Devicename, "IMP_INFO") > 0 Then
        'Pas d'impression sur l'imprimante de l'informatique
    ElseIf InStr(Printer.Devicename, "MFP_4ET_HP") > 0 Then
        'Pas d'impression sur l'imprimante de l'informatique
    ElseIf InStr(Printer.Devicename, "MFP_4ET") > 0 Then
        'Pas d'impression sur l'imprimante de l'informatique
    Else
        mnuSelect_Print_Click
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "cmdJournal_Devises : ")
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub
Private Sub cmdJournal_Solde_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdJournal_Solde  : ")
'DR 08/10/2019
If InStr(Printer.Devicename, "IMP_INFO") > 0 Then
    'Pas d'impression sur l'imprimante de l'informatique
ElseIf InStr(Printer.Devicename, "MFP_4ET_HP") > 0 Then
    'Pas d'impression sur l'imprimante de l'informatique
ElseIf InStr(Printer.Devicename, "MFP_4ET") > 0 Then
    'Pas d'impression sur l'imprimante de l'informatique
Else
    cmdJournal_Solde_Print
End If
Call lstErr_AddItem(lstErr, cmdContext, "cmdJournal_Solde : ")
Me.Show

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdJournal_Unit_Click()
Dim I As Integer, X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdJournal_Devises  : ")

'''GoTo PDF

For I = 0 To cboOptions_Unit.ListCount - 1
    cboOptions_Unit.ListIndex = I
    'If Trim(cboOptions_Unit.Text) <> "" Then
    If Trim(cboOptions_Unit.Text) <> "" And Trim(cboOptions_Unit.Text) <> "DG" Then 'DR 05/02/2015
        Call lstErr_AddItem(lstErr, cmdContext, cboOptions_Unit.Text & "  : " & fgSelect.Rows - 1)

        optOptions_SortUnit = True
        fraContextOptions_Exit
        
        If fgSelect.Rows > 1 Then
            If paramEnvironnement = constProduction Then
                meUnit.Id = cboOptions_Unit.Text
                Table_Unit meUnit
                Printer_Set meUnit.Printer
            End If
            'DR 08/10/2019
            If InStr(Printer.Devicename, "IMP_INFO") > 0 Then
                'Pas d'impression sur l'imprimante de l'informatique
            ElseIf InStr(Printer.Devicename, "MFP_4ET_HP") > 0 Then
                'Pas d'impression sur l'imprimante de l'informatique
            ElseIf InStr(Printer.Devicename, "MFP_4ET") > 0 Then
                'Pas d'impression sur l'imprimante de l'informatique
            Else
                mnuSelect_Print_Click
            End If
        End If
    End If
Next I

PDF:
For I = 0 To cboOptions_Unit.ListCount - 1
    cboOptions_Unit.ListIndex = I
    'If Trim(cboOptions_Unit.Text) <> "" Then
    If Trim(cboOptions_Unit.Text) <> "" And Trim(cboOptions_Unit.Text) <> "DG" Then 'DR 05/02/2015
        Call lstErr_AddItem(lstErr, cmdContext, cboOptions_Unit.Text & "  : " & fgSelect.Rows - 1)

        optOptions_SortUnit = True
        fraContextOptions_Exit
        
        If fgSelect.Rows > 1 Then
            X = Table_Unit_SSI("", Trim(cboOptions_Unit.Text))
            If X = "S00" Then X = "S60" 'comptabilité par défaut
            Call frmElpPrt.prtIMP_PDF_NoPaper_Init(X, "BIA-CPT-JAL-SRV", "Archive")

            mnuSelect_Print_Click
            '''Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "")
        End If
    End If
Next I

Call lstErr_AddItem(lstErr, cmdContext, "cmdJournal_Unit : ")
Me.Show

Me.Enabled = True: Me.MousePointer = 0
End Sub


Private Sub cmdMT900_Click()
Dim xMin As String, xMax As String
On Error Resume Next
Call DTPicker_Control(txtRelevéA4W_Amj, xMin)

Call lstErr_Clear(lstErr, cmdContext, "cmdMT900  : " & xMin)

Call Swift_MT900_Monitor(xMin, xMin)
Call lstErr_AddItem(lstErr, cmdContext, "cmdMT900  : terminé")

End Sub

Private Sub cmdNIC_FGDR_Click()
Dim madb As ADODB.Connection
Dim nicSourceFGDR As String
Dim nicSource As String
Dim nicClient As String
Dim nicCpt As String
Dim nici As Long
Dim nicZ As String
Dim xMemo As String
Dim xName As String
Dim V As Variant
Dim s() As String

    V = rsElpTable_Read("SIDE", "PasswordX", "SIDE_READ", xName, xMemo)
    Set madb = New ADODB.Connection
    madb.ConnectionString = "Provider= SQLOLEDB;Data Source=COMPTA2015;User ID=connexBIA;Password=" & xMemo & ";"
    madb.Open
    'on pointe vers un répertoire en dur pour l'instant !
    'nicSource = "\\DOCSRV2013\_GROUPS\PUBLIC\_DOSSIERS PARTAGES\RELEVES_CLIENTS\2017\12\"
    'nicSource = "\\DOCSRV2013\_GROUPS\PUBLIC\_DOSSIERS PARTAGES\RELEVES_CLIENTS\Decadaires\2017\12_21\"
    Load dlgRep
    dlgRep.Tag = ""
    Call dlgRep.Show(vbModal)
    If dlgRep.Tag = "" Then
        Exit Sub
    End If
    nicSource = dlgRep.Tag
    If Right(nicSource, 1) <> "\" Then
        nicSource = nicSource & "\"
    End If
    nicSourceFGDR = "c:\biasrv\modeles\Note Information Garantie des dépôts FGDR 2017.pdf"
    If Dir(nicSourceFGDR) = "" Then
        nicSourceFGDR = "d:\biasrv\modeles\Note Information Garantie des dépôts FGDR 2017.pdf"
    End If
    lstW.Clear
    frmEdition.filDoc.PATH = nicSource
    frmEdition.filDoc.Pattern = "*.pdf"
    For nici = 0 To frmEdition.filDoc.ListCount - 1
        lstW.AddItem frmEdition.filDoc.List(nici)
    Next nici
    DoEvents
    If lstW.ListCount > 0 Then
        For nici = 0 To lstW.ListCount - 1
            lstW.ListIndex = nici
            s = Split(lstW.List(nici), "_")
            nicClient = Trim(s(1))
            nicCpt = Trim(s(2))
            If nicClient <> "" Then
                If isEligbleFGDR(madb, nicClient, nicCpt) Then
                    'insertion de la page
                    nicZ = Replace(lstW.List(nici), ".pdf", "_nic.pdf")
                    Call FileCopy(nicSourceFGDR, nicSource & nicZ)
                End If
            End If
        Next nici
        Call MsgBox("Fin du traitement")
    End If
    If madb.State = adStateOpen Then
        madb.Close
    End If
    Set madb = Nothing
    
    Call MsgBox("Fin de l'insertion de la page FGDR...")
    
End Sub

Private Sub cmdOptions_Click()
If fraContextOptions.Visible Then
    fraContextOptions_Exit
Else
    mnuContextOptions_Click
End If
End Sub

Private Sub cmdPrint_Click()
Msg = Space$(50)

Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then
                fgSelect_Sort_Options
                Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
           End If
    Case 1:
            If fgRelevéA4.Rows > 1 Then
                Me.PopupMenu mnuPrint1, vbPopupMenuLeftButton
           End If
End Select

End Sub

Private Sub cmdRelevé_Annuel_Frais_Click()
Dim xMin As String, xMax As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdRelevé_Annuel_Frais  : début")
Call DTPicker_Control(txtRelevéA4_AmjMax, xMax)
xMax = Format(Val(Mid$(DSys, 1, 4)) - 1, "0000") & "1231"
xMin = Mid$(xMax, 1, 4) & "0101"
X = MsgBox("Période du " & dateImp10(xMin) & " au " & dateImp10(xMax) & ". Voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des relevés annuels de frais :")

If X = vbNo Then GoTo Exit_sub


Call MsgBox("Configurer RECTO/VERSO, 1 page par feuille", vbInformation, "Impression Extraits")

cmdRelevé_Annuel_Frais_Select xMin, xMax
If lstW.ListCount > 0 Then cmdRelevé_Annuel_Frais_Print xMin, xMax

Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevé_Annuel_Frais : Fin ")
Exit_sub:
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdRelevé_Annuel_FraisPDF_Click()
Dim xMin As String, xMax As String

    Me.Enabled = False: Me.MousePointer = vbHourglass
    Call lstErr_Clear(lstErr, cmdContext, "cmdRelevé_Annuel_Frais  : début")
    Call DTPicker_Control(txtRelevéA4_AmjMax, xMax)
    xMax = Format(Val(Mid$(DSys, 1, 4)) - 1, "0000") & "1231"
    xMin = Mid$(xMax, 1, 4) & "0101"
    X = MsgBox("Période du " & dateImp10(xMin) & " au " & dateImp10(xMax) & "." & vbCrLf & "Avez-vous sélectionné l'imprimante PDF et voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Impression PDF des relevés annuels de frais :")
    If X = vbNo Then GoTo Exit_sub
    cmdRelevé_Annuel_Frais_Select xMin, xMax
    If lstW.ListCount > 0 Then cmdRelevé_Annuel_Frais_Print_PDF xMin, xMax
    Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevé_Annuel_Frais : Fin ")
Exit_sub:
    Me.Show
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdRelevéA4D_Click()
Dim xMin As String, xMax As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdRelevéA4M  : début")





Select Case Val(Mid$(YBIATAB0_DATE_CPT_J, 7, 2))
    Case 1 To 7
                xMin = Mid$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "21"
                xMax = dateFinDeMois(YBIATAB0_DATE_CPT_MP1)
     Case 8 To 17:
                xMin = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "01"
                xMax = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "10"
    Case 18 To 27:
                xMin = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "11"
                xMax = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "20"
    Case Else
                xMin = Mid$(YBIATAB0_DATE_CPT_M, 1, 6) & "21"
                xMax = dateFinDeMois(YBIATAB0_DATE_CPT_M)

End Select
X = MsgBox("Période du " & dateImp10(xMin) & " au " & dateImp10(xMax) & ". Voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des extraits décadaires :")

If X = vbNo Then GoTo Exit_sub

Call MsgBox("Configurer RECTO/VERSO, 1 page par feuille ", vbInformation, "Impression Extraits")

If chkRelevéA4M_CptOrdinaire = "0" And chkRelevéA4M_Responsable = "0" Then
    MsgBox "à faire JPL ;Call prtYBIAMVT0_A4_Select" '("D", xMin, xMax)
Else
    cmdRelevéA4M_Select "D", xMin
    If lstW.ListCount > 0 Then cmdRelevéA4M_Print "D", xMin, xMax
End If

Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevéA4M : Fin ")
Exit_sub:

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdRelevéA4G_Click()
Dim xMin As String, xMax As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdRelevéA4G : début")
Call DTPicker_Control(txtRelevéA4_AmjMin, xMin)
Call DTPicker_Control(txtRelevéA4_AmjMax, xMax)


    cmdRelevéA4G_Select
'' MsgBox "TEST :  impression désactivée  "
   If lstW.ListCount > 0 Then cmdRelevéA4M_Print "M", xMin, xMax

Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevéA4M : Fin ")
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdRelevéA4DPDF_Click()
Dim xMin As String, xMax As String
    
    Me.Enabled = False: Me.MousePointer = vbHourglass
    Call lstErr_Clear(lstErr, cmdContext, "cmdRelevéA4M  : début")
    Select Case Val(Mid$(YBIATAB0_DATE_CPT_J, 7, 2))
        Case 1 To 7
                    xMin = Mid$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "21"
                    xMax = dateFinDeMois(YBIATAB0_DATE_CPT_MP1)
         Case 8 To 17:
                    xMin = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "01"
                    xMax = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "10"
        Case 18 To 27:
                    xMin = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "11"
                    xMax = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "20"
        Case Else
                    xMin = Mid$(YBIATAB0_DATE_CPT_M, 1, 6) & "21"
                    xMax = dateFinDeMois(YBIATAB0_DATE_CPT_M)
    End Select
    X = MsgBox("Période du " & dateImp10(xMin) & " au " & dateImp10(xMax) & "." & vbCrLf & "Avez-vous sélectionné l'imprimante PDF et voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des extraits décadaires :")
    If X = vbNo Then GoTo Exit_sub
    If chkRelevéA4M_CptOrdinaire = "0" And chkRelevéA4M_Responsable = "0" Then
        MsgBox "à faire JPL ;Call prtYBIAMVT0_A4_Select" '("D", xMin, xMax)
    Else
        cmdRelevéA4M_Select "D", xMin
        If lstW.ListCount > 0 Then cmdRelevéA4M_PrintPDF "D", xMin, xMax
    End If
Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevéA4M : Fin ")
Exit_sub:
    Me.Show
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdRelevéA4J_Click()
Dim xMax As String

Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdRelevéA4J  : ")

Call DTPicker_Control(txtRelevéA4_AmjMax, xMax)


MsgBox "$jpl à faire :Call prtYBIAMVT0_A4_Select" '("J", xMax, xMax)

Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevéA4J : ")
Me.Show

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdRelevéA4M_Click()
Dim xMin As String, xMax As String
Me.Enabled = False: Me.MousePointer = vbHourglass

'cmdRelevéA4M_Select_Control

Call lstErr_Clear(lstErr, cmdContext, "cmdRelevéA4M  : début")
Call DTPicker_Control(txtRelevéA4_AmjMin, xMin)
Call DTPicker_Control(txtRelevéA4_AmjMax, xMax)
X = vbCrLf & "sinon précisez la période dans les champs de l'onglet 'impression des relevés'"
X = MsgBox("Période du " & dateImp10(xMin) & " au " & dateImp10(xMax) & ". Voulez-vous continuer?" & X, vbYesNo + vbQuestion + vbDefaultButton2, "Impression des extraits mensuels :")

If X = vbNo Then GoTo Exit_sub


Call MsgBox("Configurer RECTO/VERSO, 1 page par feuille", vbInformation, "Impression Extraits")

If chkRelevéA4M_CptOrdinaire = "0" And chkRelevéA4M_Responsable = "0" Then
    MsgBox "$jpl à faire :Call prtYBIAMVT0_A4_Select" '("M", xMin, xMax)
Else
    cmdRelevéA4M.Tag = "CLIC"
    cmdRelevéA4M_Select "M", xMin
    If lstW.ListCount > 0 Then cmdRelevéA4M_Print "M", xMin, xMax
End If

Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevéA4M : Fin ")
cmdRelevéA4M.Tag = ""
Exit_sub:
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdRelevéA4MPDF_Click()
Dim xMin As String, xMax As String

    Me.Enabled = False: Me.MousePointer = vbHourglass
    Call lstErr_Clear(lstErr, cmdContext, "cmdRelevéA4M  : début")
    Call DTPicker_Control(txtRelevéA4_AmjMin, xMin)
    Call DTPicker_Control(txtRelevéA4_AmjMax, xMax)
    X = vbCrLf & "sinon précisez la période dans les champs de l'onglet 'impression des relevés'"
    X = MsgBox("Période du " & dateImp10(xMin) & " au " & dateImp10(xMax) & "." & vbCrLf & "Avez-vous sélectionné l'imprimante PDF et voulez-vous continuer?" & X, vbYesNo + vbQuestion + vbDefaultButton2, "Impression des extraits mensuels :")
    If X = vbNo Then GoTo Exit_sub
    If chkRelevéA4M_CptOrdinaire = "0" And chkRelevéA4M_Responsable = "0" Then
        MsgBox "$jpl à faire :Call prtYBIAMVT0_A4_Select" '("M", xMin, xMax)
    Else
        cmdRelevéA4M.Tag = "CLIC"
        cmdRelevéA4M_Select "M", xMin
        If lstW.ListCount > 0 Then cmdRelevéA4M_PrintPDF "M", xMin, xMax
    End If
    Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevéA4M : Fin ")
    cmdRelevéA4M.Tag = ""
Exit_sub:
    Me.Show
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdRelevéA4W_Click()
Dim xMax As String, X8 As String * 8
Dim blnRelevéA4W_Loro As Boolean, blnRelevéA4W_Nostro As Boolean, blnRelevéA4W_Update  As Boolean
Dim blnRelevéA4W_Confirmation As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass


Call lstErr_Clear(lstErr, cmdContext, "cmdRelevéA4W  : début")

Call DTPicker_Control(txtRelevéA4W_Amj, X8)


If chkRelevéA4W_Loro = "1" Then
    blnRelevéA4W_Loro = True
Else
    blnRelevéA4W_Loro = False
End If
If chkRelevéA4W_Nostro = "1" Then
    blnRelevéA4W_Nostro = True
Else
    blnRelevéA4W_Nostro = False
End If
If chkRelevéA4W_Update = "1" Then
    blnRelevéA4W_Update = True
Else
    blnRelevéA4W_Update = False
End If

If chkRelevéA4W_Confirmation = "1" Then
    blnRelevéA4W_Confirmation = True
Else
    blnRelevéA4W_Confirmation = False
End If

Swift_MT950_Monitor X8, X8, blnRelevéA4W_Loro, blnRelevéA4W_Nostro, blnRelevéA4W_Update, blnRelevéA4W_Confirmation, cnsab

Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevéA4W : fin")

Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub fgRelevéA4_Click()
fgRelevéA4.LeftCol = 0

End Sub

Private Sub fgRelevéA4_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnTest As Boolean, blnIban As Boolean
Dim wIndex As Long
On Error Resume Next
If y <= fgRelevéA4.RowHeightMin Then
    Select Case fgRelevéA4.Col
        Case 0: fgRelevéA4_Sort1 = 0: fgRelevéA4_Sort2 = 1: fgRelevéA4_Sort
        Case 1:  fgRelevéA4_Sort1 = 1: fgRelevéA4_Sort2 = 1: fgRelevéA4_Sort
        Case 2: fgRelevéA4_Sort1 = 2: fgRelevéA4_Sort2 = 2: fgRelevéA4_Sort
        Case 3: fgRelevéA4_Sort1 = 3: fgRelevéA4_Sort2 = 3: fgRelevéA4_Sort
    End Select
Else
    If fgRelevéA4.Rows > 1 Then
        Call fgRelevéA4_Color(fgRelevéA4_RowClick, MouseMoveUsr.BackColor, fgRelevéA4_ColorClick)
        fgRelevéA4.Col = fgRelevéA4_arrIndex
        wIndex = Val(fgRelevéA4.Text)
        meYBIAMVT0 = arrYBIAMVT0(wIndex)
        
        Call fctPCEC_Atribut(meYBIAMVT0.COMPTEOBL, meYBIAMVT0.COMPTEDEV, blnCptOrdinaire, blnRIB, blnTest, blnIban)
        mnuRIB_Print.Enabled = blnRIB
        Me.PopupMenu mnuPrint1, vbPopupMenuLeftButton
    End If
   End If

End Sub



Private Sub mnuRelevé_Print_Click()
Dim xMin As String, xMax As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdRelevéA4M  : ")
Call DTPicker_Control(txtRelevéA4_AmjMin, xMin)
Call DTPicker_Control(txtRelevéA4_AmjMax, xMax)

cmdRelevéA4_Print xMin, xMax
Call lstErr_AddItem(lstErr, cmdContext, "cmdRelevéA4M : ")

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuRelevéRIB_Print_Click()
mnuRelevé_Print_Click
mnuRIB_Print_Click
End Sub

Private Sub mnuRIB_Print_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

prtRIB_Monitor meYBIAMVT0.MOUVEMCOM
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Click()
blnTotal = True
cmdPrint_Journal

End Sub

Private Sub cmdSelect_Click()
On Error Resume Next

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAb_Compta_Import"): DoEvents
fgSelect.Visible = False

cmdYBIAMVT0_Import


Call lstErr_AddItem(lstErr, cmdContext, "! préparation affichage ...... "): DoEvents

fgSelect_Display

Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAb_Compta_Import"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub

Private Sub fgSelect_LeaveCell()
On Error Resume Next
'fgSelect.CellBackColor = &HE0E0E0
End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub


Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear: lstErr.Height = 200

If fraContextOptions.Visible Then fraContextOptions_Exit: Exit Sub
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If


If currentAction = "" Then
   
Else
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If

End Sub
Private Sub mnuContextOptions_Click()
fraContextOptions.Visible = True
previousUnit = mUnit

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
fgRelevéA4_FormatString = fgRelevéA4.FormatString
fgRelevéA4.Clear: fgRelevéA4.Row = 0


End Sub





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
Dim V

On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX fgSelect_Sort1
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_SortX fgSelect_Sort1
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex
    End If
   End If
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 4
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 9
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            If TypeOf C Is ListBox Then
                Elp_ResizeControl C
            Else
                MouseMoveActiveControl.ForeColor = C.ForeColor
                C.ForeColor = MouseMoveUsr.ForeColor
            End If
        End If
    End If
End If

End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
        If TypeOf xobj Is CommandButton Then
            xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            If TypeOf xobj Is ListBox Then
                xobj.Height = 200
            Else
                xobj.ForeColor = MouseMoveActiveControl.ForeColor
            End If
        End If
        Exit For
    End If
Next xobj
End Sub





Public Sub cmdYBIAMVT0_Import()
Dim xSQL As String, xWhere As String
Dim xIn As String, X As String
Dim seq As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

ReDim arrYBIAMVT0(5000): arrYBIAMVT0_Nb = 0: arrYBIAMVT0_Max = 5000
ReDim specialMVT(500): specialMVT_nb = 0: specialMVT_Max = 500


Call lstErr_AddItem(lstErr, cmdContext, "Import YBIAMVTH"): DoEvents
seq = 0
mMOUVEMCOM = ""
txtRelevéA4_Compte = ""
fgRelevéA4_Reset
fgRelevéA4.Rows = 1
fgRelevéA4.FormatString = fgRelevéA4_FormatString
fgRelevéA4.Visible = False

Call DTPicker_Control(txtSelectAmj, wSelectAmj)
xSelectAmj_IBM = dateIBM(wSelectAmj)
Call DTPicker_Control(txtSelectAmj_Max, wSelectAmj)
xSelectAmj_Max_IBM = dateIBM(wSelectAmj)

If optOptions_SelectJC Then
    xWhere = " where MOUVEMDCO <= " & dateIBM(YBIATAB0_DATE_CAL_AP1) & " and "
Else
    If optOptions_SelectOD Then
        xWhere = " where MOUVEMOPE like '*%'  and"
    Else
        xWhere = " where "
    End If
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & xWhere & " MOUVEMDTR >= " & xSelectAmj_IBM _
     & " and MOUVEMDTR <= " & xSelectAmj_Max_IBM _
     & " order by MOUVEMCOM,MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    seq = seq + 1
    
    If seq Mod 1000 = 0 Then
        Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, constYBIAMVT0 & " : " & seq)
      End If
      
    DoEvents
    If fctUser_Classe_Aut(rsSab("COMPTECLA")) Then    ' ? autorisé à voir ce mouvement
        Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)

        arrYBIAMVT0_Nb = arrYBIAMVT0_Nb + 1
        If arrYBIAMVT0_Nb > arrYBIAMVT0_Max Then
            arrYBIAMVT0_Max = arrYBIAMVT0_Max + 1000
            ReDim Preserve arrYBIAMVT0(arrYBIAMVT0_Max + 1000)
        End If
        arrYBIAMVT0(arrYBIAMVT0_Nb) = xYBIAMVT0
        'DR 04/03/2014
        If xYBIAMVT0.MOUVEMOPE = "ENG" And xYBIAMVT0.MOUVEMSER = "00" And xYBIAMVT0.MOUVEMSSE = "CD" Then
            xYBIAMVT0.MOUVEMSSE = "G4"
        ElseIf xYBIAMVT0.MOUVEMOPE = "ENG" Then
            Call cmdYBIAMVT0_Import_Special(arrYBIAMVT0_Nb)  '=> xYBIAMVT0
        Else
            If Mid$(xYBIAMVT0.COMPTEOBL, 1, 5) = "90319" And InStr(xYBIAMVT0.MOUVEMCOM, "HB1") > 0 Then Call cmdYBIAMVT0_Import_Special(arrYBIAMVT0_Nb)
        End If
        If mMOUVEMCOM <> xYBIAMVT0.MOUVEMCOM Then
            mMOUVEMCOM = xYBIAMVT0.MOUVEMCOM
            fgRelevéA4.Rows = fgRelevéA4.Rows + 1
            fgRelevéA4.Row = fgRelevéA4.Rows - 1
            fgRelevéA4.Col = 0: fgRelevéA4.Text = mMOUVEMCOM
            fgRelevéA4.Col = 1: fgRelevéA4.Text = xYBIAMVT0.COMPTEDEV
            fgRelevéA4.Col = 2: fgRelevéA4.Text = xYBIAMVT0.COMPTEINT
            fgRelevéA4.Col = 3: fgRelevéA4.Text = xYBIAMVT0.BIAMVTID
            fgRelevéA4.Col = fgRelevéA4_arrIndex: fgRelevéA4.Text = arrYBIAMVT0_Nb
        
        End If
        
    End If
                       
    rsSab.MoveNext
Loop

Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Nb mouvements : " & seq)

fgRelevéA4_Sort1 = 0: fgRelevéA4_Sort2 = 1: fgRelevéA4_Sort
fgRelevéA4.Visible = True

cmdYBIAMVT0_Import_Special_MOUVEMSSE ' réaffectation  des sous-services
Exit Sub

Error_Handler:

blnError = True
Shell_MsgBox "me.cmdYBIAMVT0_Import# " & Error, vbCritical, Me.Caption, False
Close

End Sub


Private Sub mnuSelect_Print_Recap_Click()
blnTotal = False
cmdPrint_Journal

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 1 Then txtRelevéA4_Compte.SetFocus

End Sub

Private Sub txtRelevéA4_AmjMax_GotFocus()
DTPicker_GotFocus txtRelevéA4_AmjMax

End Sub


Private Sub txtRelevéA4_AmjMax_LostFocus()
DTPicker_LostFocus txtRelevéA4_AmjMax

End Sub


Private Sub txtRelevéA4_AmjMin_GotFocus()
DTPicker_GotFocus txtRelevéA4_AmjMin

End Sub


Private Sub txtRelevéA4_AmjMin_LostFocus()
DTPicker_LostFocus txtRelevéA4_AmjMin

End Sub


Private Sub txtRelevéA4_Compte_Change()
Dim I As Long, X As String, lenX As Integer
On Error Resume Next
X = Trim(txtRelevéA4_Compte)
lenX = Len(X)
fgRelevéA4.Col = 0
For I = 1 To fgRelevéA4.Rows - 1
    fgRelevéA4.Row = I
    
    If X <= Mid$(fgRelevéA4.Text, 1, lenX) Then
        fgRelevéA4.TopRow = I
        Exit Sub
    End If
Next I

End Sub

Private Sub txtRelevéA4_Compte_GotFocus()
Call txt_GotFocus(txtRelevéA4_Compte)

End Sub


Private Sub txtRelevéA4_Compte_KeyPress(KeyAscii As Integer)
'num_KeyAscii KeyAscii
End Sub

Private Sub txtRelevéA4_Compte_LostFocus()
Call txt_LostFocus(txtRelevéA4_Compte)

End Sub


Private Sub txtRelevéA4M_PCEC_GotFocus()
Call txt_GotFocus(txtRelevéA4M_PCEC)

End Sub


Private Sub txtRelevéA4M_PCEC_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtRelevéA4M_PCEC_LostFocus()
Call txt_LostFocus(txtRelevéA4M_PCEC)

End Sub


Private Sub txtRelevéA4M_Service_GotFocus()
Call txt_GotFocus(txtRelevéA4M_Service)

End Sub


Private Sub txtRelevéA4M_Service_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtRelevéA4M_Service_LostFocus()
Call txt_LostFocus(txtRelevéA4M_Service)

End Sub







Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub


Public Sub fgSelect_Sort_Options()

fgSelect_Sort1_Old = -1
fgSelect_SortAD = 0
If optOptions_SortUnit Or optOptions_SelectOD Then
    If fgSelect_Sort1 <> 0 Then fgSelect_Sort1 = 0: fgSelect_Sort2 = 4: fgSelect_Sort
End If
If optOptions_SortDevise Or optOptions_SelectJC Then
    If fgSelect_Sort1 <> 5 Then fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX fgSelect_Sort1
End If
If optOptions_SortCompte Then
    If fgSelect_Sort1 <> 6 Then fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_SortX fgSelect_Sort1
End If

End Sub

Public Sub cmdPrint_Journal()
Me.Enabled = False

Msg = Space$(50)
Select Case fgSelect_Sort1
    Case 0: prtSAB_Compta.prtSAB_Compta_Unit arrYBIAMVT0(), fgSelect, fgSelect_arrIndex, Me, blnTotal
    Case 5: prtSAB_Compta.prtSAB_Compta_Devise arrYBIAMVT0(), fgSelect, fgSelect_arrIndex, Me, blnTotal
    Case 6: prtSAB_Compta.prtSAB_Compta_Compte arrYBIAMVT0(), fgSelect, fgSelect_arrIndex, Me, blnTotal
End Select
Me.Show
Me.Enabled = True

End Sub

Public Sub fraContextOptions_Exit()
fraContextOptions.Visible = False
fgSelect_Sort_Options
If mUnit <> previousUnit Then
    Call lstErr_Clear(lstErr, cmdContext, "! préparation affichage ...... "): DoEvents
    
    fgSelect_Display
    
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAb_Compta_Import"): DoEvents

End If

End Sub

Private Sub txtSelectAmj_GotFocus()
DTPicker_GotFocus txtSelectAmj

End Sub


Private Sub txtSelectAmj_LostFocus()
DTPicker_LostFocus txtSelectAmj

End Sub



Public Sub cmdReset_Date()
Call DTPicker_Set(txtSelectAmj, YBIATAB0_DATE_CPT_JP0)
Call DTPicker_Set(txtSelectAmj_Max, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtRelevéA4_AmjMax, YBIATAB0_DATE_CPT_J)
X = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "01"
Call DTPicker_Set(txtRelevéA4_AmjMin, X)

mnuContextOptions_Click
previousUnit = "$$$$$"
fraContextOptions_Exit
''cmdRelevéA4W.Caption = "Mvt du " & txtSelectAmj
txtRelevéA4W_Amj.Enabled = False
Call DTPicker_Set(txtRelevéA4W_Amj, YBIATAB0_DATE_CPT_JP0)
If paramEnvironnement = constProduction Then
    chkRelevéA4W_Nostro.Enabled = True
    chkRelevéA4W_Nostro.Value = "1"
Else
    chkRelevéA4W_Nostro.Enabled = False
    chkRelevéA4W_Nostro.Value = "0"
End If
' ATTENTION PAR DEFAUT NON
chkRelevéA4W_Nostro.Value = "0"
chkRelevéA4W_Loro.Value = "0"
chkRelevéA4W_Update.Value = "0"
chkRelevéA4W_Confirmation = "0"
End Sub




Public Sub cmdRIB_Print_A4()
'Dim X20 As String * 20

'X20 = " 00" & mId$(meYBIAMVT0.MOUVEMCOM, 1, 5)
'xMvtP0.Id = constZADRESS0 & "1" & X20 & " "
'xMvtP0.Method = "Seek="
'intReturn = tableMvtP0_Read(xMvtP0)

'If intReturn = 0 Then
'    MsgTxt = Space$(34) & xMvtP0.Text
'    MsgTxtIndex = 0
    
'    srvZADRESS0_GetBuffer meZADRESS0

'    prtRIB_Monitor meYBIAMVT0.MOUVEMCOM, meZADRESS0
    
'End If

End Sub

Private Sub txtSelectAmj_Max_GotFocus()
DTPicker_GotFocus txtSelectAmj_Max

End Sub


Private Sub txtSelectAmj_Max_LostFocus()
DTPicker_LostFocus txtSelectAmj_Max
End Sub



Public Sub cmdYBIAMVT0_Import_Special(lIndex As Long)
Dim X As String, wMOUVEMSSE As String

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' affectation des mouvements MOUVEMOPE = "ENG"
'   si au moins 1 mvt sur le compte '...12376DEC'  => DRH  : XX G9
'   si au moins 1 mvt sur un compte '........DEC'  => AUTO : XX G0
'   si liste service DER                           => DER  : XX GA
'   sinon                                          => DAFI : XX G3
' mise à jour de la table BIAS820I.mdb / SAB_Param / Ope_Unit
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

X = arrYBIAMVT0(lIndex).MOUVEMCOM  'Compte

wMOUVEMSSE = "G3" 'DAFI credits par défaut

If InStr(X, "12376DEC") > 0 Then
    wMOUVEMSSE = "G9" 'DGA
Else
'    If InStr(X, "DEC") > 0 Then wMOUVEMSSE = "G0" 'sans affectation : écritures automatiques
    If arrYBIAMVT0(lIndex).MOUVEMEVE = "AMS" Then wMOUVEMSSE = "G0" 'sans affectation : écritures automatiques
                                                         'voir cmdYBIAMVT0_Import_Special_MOUVEMSSE
End If
Select Case arrYBIAMVT0(lIndex).MOUVEMCOM
    Case "R903191EUR11225MLC", "R903191EUR50078DEC", "R903199EUR50062MLC", "R903191EUR50239MLC", "R903196EUR11477MLC": wMOUVEMSSE = "GA"
End Select

'20060310 Demande A DElalande
If Mid$(arrYBIAMVT0(lIndex).COMPTEOBL, 1, 5) = "90319" And InStr(X, "HB1") > 0 Then wMOUVEMSSE = "GA"  'DER

'----------------------------------------------------------------
arrYBIAMVT0(lIndex).MOUVEMSER = "XX"
arrYBIAMVT0(lIndex).MOUVEMSSE = wMOUVEMSSE
'----------------------------------------------------------------

specialMVT_nb = specialMVT_nb + 1
If specialMVT_nb > specialMVT_Max Then
    specialMVT_Max = specialMVT_Max + 50
    ReDim Preserve specialMVT(specialMVT_Max + 50)
End If
specialMVT(specialMVT_nb) = lIndex

End Sub
Public Sub cmdYBIAMVT0_Import_Special_MOUVEMSSE()
Dim I As Integer, X As String, wMOUVEMSSE As String
Dim wIndex As Long

'Priorité Compta / Risques / crédits
'-------------------------------------
For I = 1 To specialMVT_nb
    wIndex = specialMVT(I)
    If arrYBIAMVT0(wIndex).MOUVEMSSE = "G9" Then cmdYBIAMVT0_Import_Special_MOUVEMSSE_Pièce wIndex
Next I

For I = 1 To specialMVT_nb
    wIndex = specialMVT(I)
    If arrYBIAMVT0(wIndex).MOUVEMSSE = "G0" Then cmdYBIAMVT0_Import_Special_MOUVEMSSE_Pièce wIndex
Next I

For I = 1 To specialMVT_nb
    wIndex = specialMVT(I)
    If arrYBIAMVT0(wIndex).MOUVEMSSE = "GA" Then cmdYBIAMVT0_Import_Special_MOUVEMSSE_Pièce wIndex
Next I

End Sub

Public Sub cmdYBIAMVT0_Import_Special_MOUVEMSSE_Pièce(lIndex As Long)
Dim I As Integer, X As String, wMOUVEMSSE As String, wMOUVEMPIE As String
Dim majIndex As Long
wMOUVEMSSE = arrYBIAMVT0(lIndex).MOUVEMSSE
wMOUVEMPIE = arrYBIAMVT0(lIndex).MOUVEMPIE
For I = 1 To specialMVT_nb
    majIndex = specialMVT(I)
    If arrYBIAMVT0(majIndex).MOUVEMPIE = wMOUVEMPIE Then arrYBIAMVT0(majIndex).MOUVEMSSE = wMOUVEMSSE
Next I

End Sub


Public Sub cmdJournal_Solde_Print()
Dim xSQL As String, V
Dim wYBIACPT_C_Nb As Long, wYBIACPT_C_NbMax As Long
Dim K As Long, K0 As Long
Dim xCOMPTECOM As String, curX As Currency, curM As Currency
Dim blnOk As Boolean
Dim xText As String
Dim Erreur_Nb As Long
Dim I As Long

Set rsSab = Nothing
ReDim prtBIA_Compta_Control.arrYBIACPT_C(10001)
wYBIACPT_C_NbMax = 10000: wYBIACPT_C_Nb = 0

xSQL = "select COMPTECOM,COMPTEOBL,COMPTEINT,COMPTEDEV,SOLDEDMO,SOLDECEN from " _
     & paramIBM_Library_SABSPE & ".YBIACPT0 order by COMPTECOM"
 
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    wYBIACPT_C_Nb = wYBIACPT_C_Nb + 1
    If wYBIACPT_C_Nb >= wYBIACPT_C_NbMax Then ReDim Preserve prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb + 1000): wYBIACPT_C_NbMax = wYBIACPT_C_NbMax + 1000
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).COMPTECOM = rsSab("COMPTECOM")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).COMPTEOBL = rsSab("COMPTEOBL")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).COMPTEINT = rsSab("COMPTEINT")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).COMPTEDEV = rsSab("COMPTEDEV")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).SOLDEDMO = rsSab("SOLDEDMO")
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).SOLDECEN = rsSab("SOLDECEN") / 1000
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).SOLDEJ_2 = 0
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).MOUVEMMON_DB = 0
    prtBIA_Compta_Control.arrYBIACPT_C(wYBIACPT_C_Nb).MOUVEMMON_CR = 0
    rsSab.MoveNext
Loop

prtBIA_Compta_Control.arrYBIACPT_C_Nb = wYBIACPT_C_Nb
prtBIA_Compta_Control.arrYBIACPT_C_NbMax = wYBIACPT_C_NbMax

prtTitleText = " Contrôle des soldes du " & dateImp(YBIATAB0_DATE_CPT_JP1) & " au " & dateImp(YBIATAB0_DATE_CPT_J)
prtBIA_Compta_Control_Open
XPrt.FontSize = 12

xText = "- Contrôle des soldes JOUR = Veille + cumul de mouvements du jour"
XPrt.FontUnderline = True
prtBIA_Compta_Control_Anomalie xText, True
XPrt.FontUnderline = False
prtBIA_Compta_Control_Anomalie "", False
Erreur_Nb = 0
K0 = 1
xSQL = "select SOLDECOM,SOLDECEN from " _
     & paramIBM_Library_SABSPE & ".ZSOLDE0J_2 order by SOLDECOM"
 
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    xCOMPTECOM = rsSab("SOLDECOM")
    blnOk = False
    For K = K0 To wYBIACPT_C_Nb
        If xCOMPTECOM = prtBIA_Compta_Control.arrYBIACPT_C(K).COMPTECOM Then
            prtBIA_Compta_Control.arrYBIACPT_C(K).SOLDEJ_2 = rsSab("SOLDECEN")
            K0 = K + 1
            blnOk = True
            Exit For
        End If
       
    Next K
    If Not blnOk Then
        Erreur_Nb = Erreur_Nb + 1
        xText = "? ZSOLDE0J_2 inconnu " & xCOMPTECOM
        prtBIA_Compta_Control_Anomalie xText, False
    End If
    rsSab.MoveNext
Loop

xSQL = "select MOUVEMCOM,MOUVEMMON from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMDTR = " & YBIATAB0_DIBM_CPT_J _
     & " order by MOUVEMCOM"

Set rsSab = cnsab.Execute(xSQL)
K0 = 1
Do While Not rsSab.EOF
    xCOMPTECOM = rsSab("MOUVEMCOM")
    curX = rsSab("MOUVEMMON")
    blnOk = False
    For K = K0 To wYBIACPT_C_Nb
        If xCOMPTECOM = prtBIA_Compta_Control.arrYBIACPT_C(K).COMPTECOM Then
            If curX > 0 Then
                prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_DB = prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_DB + curX
            Else
                prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_CR = prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_CR + curX
            End If
            K0 = K
            blnOk = True
            Exit For
        End If
       
    Next K
    If Not blnOk Then
        xText = "? YBIAMVTH inconnu " & xCOMPTECOM: prtBIA_Compta_Control_Anomalie xText, False
        Erreur_Nb = Erreur_Nb + 1
    End If
    rsSab.MoveNext
Loop

For K = 1 To wYBIACPT_C_Nb
    curM = prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_DB + prtBIA_Compta_Control.arrYBIACPT_C(K).MOUVEMMON_CR
    curX = prtBIA_Compta_Control.arrYBIACPT_C(K).SOLDEJ_2
    If curX + curM <> prtBIA_Compta_Control.arrYBIACPT_C(K).SOLDECEN Then
        xText = "? Ecart " & prtBIA_Compta_Control.arrYBIACPT_C(K).COMPTECOM & curX & vbTab & curM & vbTab & prtBIA_Compta_Control.arrYBIACPT_C(K).SOLDECEN
        prtBIA_Compta_Control_Anomalie xText, False
        Erreur_Nb = Erreur_Nb + 1
    End If
   
Next K
If Erreur_Nb = 0 Then
    xText = " aucune erreur de solde détectée"
Else
    XPrt.FontSize = 14
    xText = "!!!!!!! " & Erreur_Nb & " erreur(s) détectée(s) !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
End If
xText = wYBIACPT_C_Nb & " comptes contrôlés : " & xText
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
prtBIA_Compta_Control_Anomalie xText, True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor


prtBIA_Compta_Control.arrDev_Nb = cboDevise.ListCount - 1
ReDim arrDev_B(prtBIA_Compta_Control.arrDev_Nb + 1)
ReDim arrDev_HB(prtBIA_Compta_Control.arrDev_Nb + 1)
For I = 0 To prtBIA_Compta_Control.arrDev_Nb
    cboDevise.ListIndex = I
    prtBIA_Compta_Control.arrDev_B(I + 1).COMPTEDEV = Trim(cboDevise.Text)
    prtBIA_Compta_Control.arrDev_HB(I + 1).COMPTEDEV = Trim(cboDevise.Text)
Next I
prtBIA_Compta_Control_Cumul

prtBIA_Compta_Control_Close
End Sub
