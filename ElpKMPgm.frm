VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmElpKMPgm 
   AutoRedraw      =   -1  'True
   Caption         =   "ElpKMPgm : Gestion des sources"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "ElpKMPgm.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   9090
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5100
      TabIndex        =   11
      Top             =   -15
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   9
      Top             =   495
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "DSPFFD"
      TabPicture(0)   =   "ElpKMPgm.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraElpKMPgm0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstElpKMPgm_W"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "DSPFFD Indexation"
      TabPicture(1)   =   "ElpKMPgm.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraElpKMPgm1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CLP TEST <=> PROD"
      TabPicture(2)   =   "ElpKMPgm.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSABPF"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "DSPOBJD"
      TabPicture(3)   =   "ElpKMPgm.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraDSPOBJD"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "DSPFDY2"
      TabPicture(4)   =   "ElpKMPgm.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraDSPFDY2"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraDSPFDY2 
         Height          =   4965
         Left            =   -74910
         TabIndex        =   54
         Top             =   420
         Width           =   8880
         Begin VB.Frame fraDSPFDY2_CLP 
            Caption         =   "AS400 : librairie DESTINATION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1485
            Left            =   5535
            TabIndex        =   62
            Top             =   3270
            Width           =   3225
            Begin VB.CommandButton cmdDSPFDY2_CLP 
               BackColor       =   &H0000FF00&
               Caption         =   "Générer CLP"
               Height          =   420
               Left            =   1425
               Style           =   1  'Graphical
               TabIndex        =   66
               Top             =   300
               Width           =   1650
            End
            Begin VB.OptionButton optDSPFDY2_SAB073 
               Caption         =   "SAB073"
               Height          =   240
               Left            =   240
               TabIndex        =   65
               Top             =   240
               Value           =   -1  'True
               Width           =   1170
            End
            Begin VB.OptionButton optDSPFDY2_SAB073T 
               Caption         =   "SAB073T"
               Height          =   240
               Left            =   210
               TabIndex        =   64
               Top             =   555
               Width           =   1125
            End
            Begin VB.TextBox txtDSPFDY2_CLP 
               Height          =   300
               Left            =   300
               TabIndex        =   63
               Top             =   990
               Width           =   2730
            End
         End
         Begin VB.CommandButton cmdDSPFDY2_Print 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Imprimer "
            Height          =   360
            Left            =   6255
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   2340
            Width           =   2040
         End
         Begin VB.TextBox txtDSPFDY2_Src 
            Height          =   300
            Left            =   5610
            TabIndex        =   60
            Text            =   "D:\Temp\FTP\DSPFDY2.txt"
            Top             =   1425
            Width           =   3135
         End
         Begin VB.CommandButton cmdDSPFDY2_Import 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Importer"
            Height          =   360
            Left            =   6225
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   1860
            Width           =   2040
         End
         Begin VB.CommandButton cmdDSPFDY2_Clear 
            BackColor       =   &H000000FF&
            Caption         =   "Effacer référentiel 23***"
            Height          =   420
            Left            =   6225
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   900
            Width           =   1800
         End
         Begin VB.TextBox txtDSPFDY2_Select 
            Height          =   285
            Left            =   7545
            TabIndex        =   57
            Top             =   315
            Width           =   1155
         End
         Begin VB.CommandButton cmdDSPFDY2_Select 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher fichier  (*nom)"
            Height          =   375
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   240
            Width           =   2055
         End
         Begin VB.ListBox lstDSPFDY2 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4470
            Left            =   75
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   255
            Width           =   5265
         End
      End
      Begin VB.Frame fraDSPOBJD 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   33
         Top             =   405
         Width           =   8925
         Begin VB.OptionButton optDSPOBJD_22200 
            Caption         =   "affichage 2 / 2"
            Height          =   225
            Left            =   5730
            TabIndex        =   44
            Top             =   1125
            Width           =   2730
         End
         Begin VB.OptionButton optDSPOBJD_22100 
            Caption         =   "affichage 1 / 2"
            Height          =   195
            Left            =   5715
            TabIndex        =   43
            Top             =   795
            Value           =   -1  'True
            Width           =   2910
         End
         Begin VB.CommandButton cmdDSPOBJD_Compare_Print 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Imprimer  (comparaison)"
            Height          =   360
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   4170
            Width           =   2040
         End
         Begin VB.CommandButton cmdDSPOBJD_Compare 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Comparer"
            Height          =   360
            Left            =   6570
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   3660
            Width           =   2040
         End
         Begin VB.TextBox txtDSPOBJD_Select 
            Height          =   285
            Left            =   7605
            TabIndex        =   40
            Top             =   300
            Width           =   1155
         End
         Begin VB.CommandButton cmdDSPOBJD_Select 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher fichier  (*nom)"
            Height          =   375
            Left            =   5505
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   270
            Width           =   2055
         End
         Begin VB.CommandButton cmdDSPOBJD_Import 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Importer"
            Height          =   360
            Left            =   6540
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   3120
            Width           =   2040
         End
         Begin VB.ListBox lstDSPOBJD 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4470
            Left            =   75
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   180
            Width           =   5265
         End
         Begin VB.TextBox txtDSPOBJD_Src2 
            Height          =   285
            Left            =   5670
            TabIndex        =   36
            Text            =   "C:\Temp\FTP\DSPOBJDY0_SAB073T"
            Top             =   2745
            Width           =   3090
         End
         Begin VB.TextBox txtDSPOBJD_Src1 
            Height          =   300
            Left            =   5670
            TabIndex        =   35
            Text            =   "C:\Temp\FTP\DSPOBJDY0_SAB073"
            Top             =   2295
            Width           =   3135
         End
         Begin VB.CommandButton cmdDSPOBJD_Clear 
            BackColor       =   &H000000FF&
            Caption         =   "Effacer référentiel 22***"
            Height          =   420
            Left            =   6570
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1785
            Width           =   1800
         End
      End
      Begin VB.Frame fraSABPF 
         Height          =   4905
         Left            =   -74940
         TabIndex        =   23
         Top             =   405
         Width           =   8940
         Begin VB.CommandButton cmdSABPF_Gen 
            BackColor       =   &H000000FF&
            Caption         =   "Générer référentiel"
            Height          =   420
            Left            =   7215
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   2880
            Width           =   1515
         End
         Begin VB.CommandButton cmdSABPF_Clear 
            BackColor       =   &H000000FF&
            Caption         =   "Effacer référentiel"
            Height          =   420
            Left            =   5625
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2880
            Width           =   1530
         End
         Begin VB.CommandButton cmdSABPF_Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer les mises à jour"
            Height          =   390
            Left            =   6105
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   2400
            Width           =   2145
         End
         Begin VB.Frame fraSABPF_CLP 
            Caption         =   "AS400 : librairie DESTINATION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1485
            Left            =   5580
            TabIndex        =   45
            Top             =   3360
            Width           =   3225
            Begin VB.TextBox txtSABPF_CLP 
               Height          =   300
               Left            =   300
               TabIndex        =   49
               Top             =   990
               Width           =   2730
            End
            Begin VB.OptionButton optSABPF_SAB073T 
               Caption         =   "SAB073T"
               Height          =   240
               Left            =   210
               TabIndex        =   48
               Top             =   555
               Width           =   1125
            End
            Begin VB.OptionButton optSABPF_SAB073 
               Caption         =   "SAB073"
               Height          =   240
               Left            =   240
               TabIndex        =   47
               Top             =   255
               Value           =   -1  'True
               Width           =   1170
            End
            Begin VB.CommandButton cmdSABPF_CLP 
               BackColor       =   &H0000FF00&
               Caption         =   "Générer CLP"
               Height          =   420
               Left            =   1425
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   300
               Width           =   1650
            End
         End
         Begin VB.Frame fraSABPF_Select 
            Caption         =   "Référentiel : Bia.mdb / INFO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   5520
            TabIndex        =   27
            Top             =   240
            Width           =   3240
            Begin VB.OptionButton optSABPF_QRY 
               Caption         =   "QRY : Query "
               Height          =   330
               Left            =   120
               TabIndex        =   53
               Top             =   1560
               Width           =   2370
            End
            Begin VB.OptionButton optSABPF_PAR 
               Caption         =   "CPY : copier les paramètres"
               Height          =   330
               Left            =   135
               TabIndex        =   32
               Top             =   1290
               Width           =   2370
            End
            Begin VB.OptionButton optSABPF_CLR 
               Caption         =   "CLR : effacer les opérations"
               Height          =   225
               Left            =   150
               TabIndex        =   31
               Top             =   1065
               Value           =   -1  'True
               Width           =   2400
            End
            Begin VB.OptionButton optSABPF_CPY 
               Caption         =   "CPY : copier environnement"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   825
               Width           =   2355
            End
            Begin VB.CommandButton cmdSABPF_Select 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Rechercher fichier >"
               Height          =   420
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   300
               Width           =   1600
            End
            Begin VB.TextBox txtSABPF_Select 
               Height          =   285
               Left            =   1965
               TabIndex        =   28
               Top             =   345
               Width           =   1155
            End
         End
         Begin VB.CommandButton cmdSABPF_False 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Tout déséléctionner"
            Height          =   420
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   4425
            Width           =   2085
         End
         Begin VB.CommandButton cmdSABPF_True 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Tout séléctionner"
            Height          =   420
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   4425
            Width           =   2070
         End
         Begin VB.ListBox lstSABPF 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4140
            ItemData        =   "ElpKMPgm.frx":0396
            Left            =   75
            List            =   "ElpKMPgm.frx":0398
            Style           =   1  'Checkbox
            TabIndex        =   24
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraElpKMPgm1 
         Height          =   4785
         Left            =   -74880
         TabIndex        =   16
         Top             =   510
         Width           =   8880
         Begin VB.Frame fraElpKM_mdb 
            Caption         =   "Import / Export"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1365
            Left            =   105
            TabIndex        =   67
            Top             =   1665
            Width           =   8655
            Begin VB.Frame fraElpKM_mdb_Table 
               Height          =   1080
               Left            =   2505
               TabIndex        =   73
               Top             =   150
               Width           =   1575
               Begin VB.OptionButton optElpKM_mdb_Index 
                  Caption         =   "ElpKMIndex"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   75
                  Top             =   705
                  Width           =   1320
               End
               Begin VB.OptionButton optElpKM_mdb_Info 
                  Caption         =   "ElpKMInfo"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   74
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   1320
               End
            End
            Begin VB.OptionButton optElpKM_mdb_Import 
               Caption         =   "Importer"
               Height          =   195
               Left            =   180
               TabIndex        =   71
               Top             =   1020
               Value           =   -1  'True
               Width           =   1230
            End
            Begin VB.OptionButton optElpKM_mdb_Clear 
               Caption         =   "Effacer BIAS820I.mdb"
               Height          =   240
               Left            =   180
               TabIndex        =   70
               Top             =   660
               Width           =   1920
            End
            Begin VB.OptionButton optElpKM_mdb_Export 
               Caption         =   "Exporter"
               Height          =   225
               Left            =   165
               TabIndex        =   69
               Top             =   300
               Width           =   1770
            End
            Begin VB.CommandButton cmdElpKM_mdb_Ok 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Ok"
               Height          =   720
               Left            =   5550
               Style           =   1  'Graphical
               TabIndex        =   68
               Top             =   390
               Width           =   2175
            End
         End
         Begin VB.Frame fraElpKMPgm0_1 
            Caption         =   "Indexation  mots clés"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   120
            TabIndex        =   17
            Top             =   255
            Width           =   8715
            Begin VB.CommandButton cmdElpKMPgm_Clear 
               BackColor       =   &H000000FF&
               Caption         =   "SUPPRESSION  TOUS LES INDEX et INFOS"
               Height          =   705
               Left            =   4305
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   315
               Width           =   4170
            End
            Begin VB.OptionButton optElpKMPgm_Index_12000 
               Caption         =   "Desciption champs"
               Height          =   315
               Left            =   135
               TabIndex        =   20
               Top             =   705
               Width           =   1785
            End
            Begin VB.OptionButton optElpKMPgm_Index_11000 
               Caption         =   "Descrition fichiers"
               Height          =   225
               Left            =   135
               TabIndex        =   19
               Top             =   345
               Width           =   1770
            End
            Begin VB.CommandButton cmdElpKMPgm_Index 
               BackColor       =   &H00C0FFC0&
               Caption         =   "lancer le traitement"
               Height          =   660
               Left            =   2070
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   330
               Width           =   1950
            End
         End
      End
      Begin VB.ListBox lstElpKMPgm_W 
         Height          =   1425
         Left            =   3135
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   2190
         Width           =   1080
      End
      Begin VB.Frame fraElpKMPgm0 
         Height          =   4995
         Left            =   75
         TabIndex        =   10
         Top             =   345
         Width           =   8850
         Begin VB.CommandButton cmdElpKMPgm_File_TableDef 
            BackColor       =   &H000000FF&
            Caption         =   "Création TableDef => .mdb "
            Height          =   480
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   1920
            Width           =   2655
         End
         Begin VB.CommandButton cmdElpKMPgm_JRN_VB 
            BackColor       =   &H0080C0FF&
            Caption         =   "Génération VB <= D:\Temp\Sab\ZMNUFIC0_VB"
            Height          =   480
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   1440
            Width           =   2655
         End
         Begin VB.CommandButton cmdElpKMPgm_JRN_CL 
            BackColor       =   &H0080C0FF&
            Caption         =   "Génération IBM  CL <= D:\Temp\Sab\ZMNUFIC0_JRN"
            Height          =   480
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton cmdElpKMPgm_JRN 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Génération IBM  J***.pf    <= D:\Temp\Sab\ZMNUFIC0_JRN"
            Height          =   480
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txtElpKMPgm_Export 
            Height          =   285
            Left            =   5880
            TabIndex        =   1
            Text            =   "D:\Temp\ .txt"
            Top             =   2880
            Width           =   2820
         End
         Begin VB.TextBox txtElpKMPgm_Select 
            Height          =   285
            Left            =   7575
            TabIndex        =   0
            Top             =   120
            Width           =   1155
         End
         Begin VB.CommandButton cmdElpKMPgm_Select 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher fichier >"
            Height          =   375
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   1665
         End
         Begin VB.ListBox lstElpKMPgm_File 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4470
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   225
            Width           =   5730
         End
         Begin VB.Frame fraElpKMPgm_Import 
            Caption         =   "Import"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Left            =   5880
            TabIndex        =   12
            Top             =   3240
            Width           =   2910
            Begin VB.CommandButton cmdElpKMPgm_Import_Auto 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Import auto"
               Height          =   555
               Left            =   1425
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   840
               Width           =   1350
            End
            Begin VB.OptionButton optElpKMPgm_Import_DSPFDY1 
               Caption         =   "DSPFDY1"
               Height          =   285
               Left            =   90
               TabIndex        =   5
               Top             =   960
               Width           =   1275
            End
            Begin VB.OptionButton optElpKMPgm_Import_DSPFDY0 
               Caption         =   "DSPFDY0"
               Height          =   195
               Left            =   75
               TabIndex        =   4
               Top             =   675
               Width           =   1170
            End
            Begin VB.OptionButton optElpKMPgm_Import_DSPFFDY0 
               Caption         =   "DSPFFDY0"
               Height          =   195
               Left            =   75
               TabIndex        =   3
               Top             =   330
               Width           =   1275
            End
            Begin VB.CommandButton cmdElpKMPgm_Import 
               Caption         =   "Import manuel"
               Height          =   555
               Left            =   1425
               TabIndex        =   13
               Top             =   180
               Width           =   1350
            End
            Begin VB.TextBox txtElpKMPgm_Import 
               Height          =   285
               Left            =   120
               TabIndex        =   2
               Text            =   "D:\Temp\ .txt"
               Top             =   1320
               Width           =   2715
            End
         End
         Begin VB.Label lblElpKMPgm_Export 
            Caption         =   "Export DDS, CBL,VB"
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
            Left            =   6000
            TabIndex        =   21
            Top             =   2640
            Width           =   1995
         End
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8600
      Picture         =   "ElpKMPgm.frx":039A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuElpKMPgm_File 
      Caption         =   "mnuElpKMPgm_File"
      Visible         =   0   'False
      Begin VB.Menu mnuElpKMPgm_File_Afficher 
         Caption         =   "Afficher la description"
      End
      Begin VB.Menu mnuElpKMPgm_File_Afficher_Alpha 
         Caption         =   "Afficher la description (# alpha)"
      End
      Begin VB.Menu mnuElpKMPgm_File_X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuElpKMPgm_File_Imprimer 
         Caption         =   "Imprimer la description"
      End
      Begin VB.Menu mnuElpKMPgm_File_Imprimer_Alpha 
         Caption         =   "Imprimer la description (# alpha)"
      End
      Begin VB.Menu mnuElpKMPgm_File_X2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuElpKMPgm_File_Exporter 
         Caption         =   "Exporter"
      End
   End
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_SelectTrue 
         Caption         =   "Imprimer la liste des fichiers sélectionnés"
      End
      Begin VB.Menu mnuPrint2_SelectFalse 
         Caption         =   "Imprimer la liste des fichiers NON sélectionnés"
      End
   End
End
Attribute VB_Name = "frmElpKMPgm"
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
Dim ElpKMPgmAut As typeAuthorization

Dim blnError As Boolean
Dim blnAuto_ElpKMPgm As Boolean
Dim Nb As Long, Nb_Rec As Long

Dim meElpKMInfo As typeElpKMInfo, xElpKMInfo As typeElpKMInfo, pflfElpKMInfo As typeElpKMInfo
Dim mElpKMSrc_Id As Long, mElpKMPgm_Library As Integer, blnLibrary As Boolean
Dim mElpKMSrc_WHLIB As String * 10, mElpKMSrc_WHFILE As String * 10
Dim mElpKMSrc_WHFILE_Z As String, mElpKMSrc_WHFILE_Y As String, mElpKMSrc_WHFILE_J As String
Dim X80 As String * 80, X80B As String * 80

Dim meDSPFDY0 As typeDSPFDY0
Dim meDSPFDY1 As typeDSPFDY1
Dim meDSPFFDY0 As typeDSPFFDY0
Dim meElpKMIndex As typeElpKMIndex, xElpKMIndex As typeElpKMIndex

Dim arrSABPF(4000) As typeElpKMInfo, arrSABPF_Nb As Integer

Dim paramSABPF_K As Integer, paramSABPF_opt As String * 3
Dim paramFromLib As String, paramToLib As String

Dim meDSPOBJDY0 As typeDSPOBJDY0, xDSPOBJDY0 As typeDSPOBJDY0
Dim lstDSPOBJD_ElpKMSrc_Id As Long

Dim arrExport_Nb As Integer
Dim arrExport_WHFLDE(500) As String
Dim arrExport_WHCHD1(500) As String
Dim arrExport_WHCHD2(500) As String
Dim arrExport_Pos(500) As Integer
Dim arrExport_Len(500) As Integer

Dim arrExport_JRN_VB_Nb As Integer
Dim arrExport_JRN_VB() As String


'-----------------------------------------------------------
Dim mdbDataBase As Database
Dim mdbTable As TableDef
Dim mdbField As Field
Dim mdbIndex As Index

Public Sub cmdDSPOBJD_Import_Src(lElpKMSrc_Id As Long)
Dim xIn As String, X As String, X5 As String * 5

recElpKMInfo_Init meElpKMInfo
meElpKMInfo.Method = constAddNew
meElpKMInfo.Pass = 1000
meElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id

xElpKMInfo = meElpKMInfo


Do Until EOF(1)
    Line Input #1, xIn
  '  MsgTxt = Space$(34) & xIn
  '  MsgTxtIndex = 0
  '  srvDSPFDY0_GetBuffer meDSPFDY0
    Nb_Rec = Nb_Rec + 1
    If Nb_Rec Mod 100 = 0 Then Call lstErr_Clear(lstErr, cmdContext, "cmdDSPOBJD_Import_Src : " & Nb_Rec)
    

    meElpKMInfo.ID = mId$(xIn, 24, 18)
    
    meElpKMInfo.Description = mId$(xIn, 64, 40)
    meElpKMInfo.Memo = Trim(xIn)
    xElpKMInfo.ID = meElpKMInfo.ID
    
    dbElpKMInfo_Update meElpKMInfo


    DoEvents

Loop

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

Call BiaPgmAut_Init(mId$(Msg, 1, 12), ElpKMPgmAut)

'blnSetfocus = True
Form_Init

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "@AUTO_SPLF":     blnAuto_ElpKMPgm = True: Auto_ElpKMPgm
    Case Else: blnAuto_ElpKMPgm = False
End Select

End Sub


Public Sub Form_Init()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistent", vbCritical, "frmElpKMPgm.param_init"
    Unload Me
End If

blnControl = False
cmdReset
Me.Enabled = True: Me.MousePointer = 0

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""
lstElpKMPgm_W.Visible = False
mdbElpKMInfo.tableElpKMInfo_Open
SSTab1.Tab = 0
mElpKMSrc_Id = 12000
recDSPFFDY0_Init meDSPFFDY0
recElpKMInfo_Init meElpKMInfo

optElpKMPgm_Import_DSPFFDY0 = True
txtElpKMPgm_Import = "C:\Temp\SAB\DSPFFDY0" 'paramElpKMPgmFtp_File
txtElpKMPgm_Export = "C:\Temp\SAB\ElpKMPgm.txt" 'paramElpKMPgmSplf_Folder

''lstElpKMPgm_11000_Load
lstDSPOBJD_ElpKMSrc_Id = 22100
cmdDSPOBJD_Compare_Print.Enabled = False

fraSABPF_CLP.Enabled = False
fraDSPFDY2_CLP.Enabled = False
blnControl = True

End Sub


Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String

Dim V
param_Init = Null
'If Not IsNull(paramElpKMPgm_Init(Me.lstErr)) Then Exit Function



End Function






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

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub


Private Sub cmdElpKMPgm_Export()
On Error Resume Next

Me.Enabled = False: Me.MousePointer = vbHourglass
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Export: " & Time)
DoEvents

'lstElpKMPgm_File_Load

mElpKMPgm_Library = srvDSPFFDY0_Library(mElpKMSrc_WHLIB)
mElpKMSrc_Id = 12000 + mElpKMPgm_Library
Call srvDSPFFDY0_lstAddItem(lstElpKMPgm_W, mElpKMSrc_Id, mElpKMSrc_WHFILE, False)

mElpKMSrc_WHFILE_Z = Trim(mElpKMSrc_WHFILE)
mElpKMSrc_WHFILE_Y = Trim(mElpKMSrc_WHFILE)
Mid$(mElpKMSrc_WHFILE_Y, 1, 1) = "Y"
mElpKMSrc_WHFILE_J = mElpKMSrc_WHFILE_Y
Mid$(mElpKMSrc_WHFILE_J, 1, 1) = "J"

Open Trim(txtElpKMPgm_Export) For Output As #2

'Call lstErr_AddItem(lstErr, cmdContext, "TableDef"): cmdElpKMPgm_Export_TableDef

mElpKMSrc_Id = 12000 + mElpKMPgm_Library
Call lstErr_AddItem(lstErr, cmdContext, "VB"): cmdElpKMPgm_Export_VB
Call lstErr_AddItem(lstErr, cmdContext, "init"): cmdElpKMPgm_Export_VB_Init
Call lstErr_AddItem(lstErr, cmdContext, "GetBuffer_Rs"): cmdElpKMPgm_Export_VB_GetBuffer_Rs
Call lstErr_AddItem(lstErr, cmdContext, "PutBuffer_Rs"): cmdElpKMPgm_Export_VB_PutBuffer_Rs

Call lstErr_AddItem(lstErr, cmdContext, "GetBuffer"): cmdElpKMPgm_Export_VB_GetBuffer_ODBC
Call lstErr_AddItem(lstErr, cmdContext, "Display"): cmdElpKMPgm_Export_VB_frmElpDisplay
Call lstErr_AddItem(lstErr, cmdContext, "CSV"): cmdElpKMPgm_Export_VB_CSV

Call lstErr_AddItem(lstErr, cmdContext, "Get"): cmdElpKMPgm_Export_VB_GetBuffer
Call lstErr_AddItem(lstErr, cmdContext, "Put"): cmdElpKMPgm_Export_VB_PutBuffer

Call lstErr_AddItem(lstErr, cmdContext, "DDS"): cmdElpKMPgm_Export_DDS
Call lstErr_AddItem(lstErr, cmdContext, "CBL"): cmdElpKMPgm_Export_CBL_ZY
Call lstErr_AddItem(lstErr, cmdContext, "CBL"): cmdElpKMPgm_Export_CBL_YZ
Close
Call lstErr_AddItem(lstErr, cmdContext, "cmdElpKMPgm_Export :  Fin")
Me.MousePointer = vbDefault
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDSPFDY2_Clear_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdDSPFDY2_Clear  : " & Time)
mdbElpKMInfo.tableElpKMInfo_Open
X = "delete * from ElpKMinfo where ElpKMSrc_Id >= 23000 and ElpKMSrc_Id <= 23999"
MDB.Execute X
Call lstErr_AddItem(lstErr, cmdContext, "cmdDSPFDY2_Clear : terminé")

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDSPFDY2_CLP_Click()
On Error Resume Next

Me.Enabled = False: Me.MousePointer = vbHourglass
Me.MousePointer = vbHourglass
lstErr.Visible = True
lstDSPFDY2.Visible = False

Call lstErr_Clear(lstErr, cmdContext, "cmdDSPFDY2_CLP: " & Time)
DoEvents
If optDSPFDY2_SAB073 Then
    paramFromLib = "SAB073T": paramToLib = "SAB073"
Else
    paramFromLib = "SAB073": paramToLib = "SAB073T"
End If

Open Trim(txtDSPFDY2_CLP) For Output As #2

X80 = Space(13) & "PGM": Print #2, X80
X80 = Space(13) & "DCL        VAR(&FROMLIB) TYPE(*CHAR) LEN(10) VALUE(" & paramFromLib & ")": Print #2, X80
X80 = Space(13) & "DCL        VAR(&TOLIB) TYPE(*CHAR) LEN(10) VALUE(" & paramToLib & ")": Print #2, X80
X80 = "": Print #2, X80

cmdDSPFDY2_CLP_QRY

X80 = "": Print #2, X80
X80 = Space(13) & "ENDPGM": Print #2, X80

Close
lstDSPFDY2.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "cmdDSPFDY2_CLP :  Fin")
Me.MousePointer = vbDefault
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDSPFDY2_Import_Click()
Dim wFileName_FTP As String
Dim paramDSPFDY2_Import As String

Me.Enabled = False: Me.MousePointer = vbHourglass
blnError = True
Nb_Rec = 0
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdDSPFDY2_Import : " & Time)

paramDSPFDY2_Import = Trim(txtDSPFDY2_Src)
wFileName_FTP = Dir(paramDSPFDY2_Import)
If wFileName_FTP = "" Then Call lstErr_AddItem(lstErr, cmdContext, "! pas de fichier : " & paramDSPFDY2_Import): GoTo Error_Handle
Open paramDSPFDY2_Import For Input As #1
cmdDSPFDY2_Import_Src 23000
Close
Call lstErr_AddItem(lstErr, cmdContext, paramDSPFDY2_Import & " : " & Nb_Rec)


blnError = False

Call lstErr_AddItem(lstErr, cmdContext, paramDSPFDY2_Import & " : " & Nb_Rec)
GoTo fin

Error_Handle:
If Not blnAuto_ElpKMPgm Then MsgBox wFileName_FTP & ":" & Error, vbCritical, "cmdDSPFDY2_Import_Click"
fin:
Close
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDSPFDY2_Print_Click()
lstDSPFDY2.Visible = False
MsgBox "2005.02.11 à revoir"
'Call prtElpKm_CLP(lstDSPFDY2, True, Trim(txtDSPFDY2_Select))
lstDSPFDY2.Visible = True

End Sub

Private Sub cmdDSPFDY2_Select_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDSPFDY2_Print.Enabled = True

lstDSPFDY2_Load 23000
fraDSPFDY2_CLP.Enabled = True
txtDSPFDY2_CLP = "C:\Temp\SABPF_QRY.txt"

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDSPOBJD_Clear_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdDSPOBJD_Clear  : " & Time)
mdbElpKMInfo.tableElpKMInfo_Open
X = "delete * from ElpKMinfo where ElpKMSrc_Id >= 22000 and ElpKMSrc_Id <= 22999"
MDB.Execute X
Call lstErr_AddItem(lstErr, cmdContext, "cmdDSPOBJD_Clear : terminé")

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDSPOBJD_Compare_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "cmdDSPOBJD_Compare : " & Time)
Nb_Rec = 0

lstDSPOBJD.Visible = False
lstDSPOBJD.Clear
Call lstErr_Clear(lstErr, cmdContext, "cmdDSPOBJD_Compare 1/2 : " & Time)

cmdDSPOBJD_Compare_Memo 22100, 22200, True
Call lstErr_AddItem(lstErr, cmdContext, "cmdDSPOBJD_Compare 1/2 : " & Nb_Rec)
Call lstErr_Clear(lstErr, cmdContext, "cmdDSPOBJD_Compare 2/1 : " & Time)

cmdDSPOBJD_Compare_Memo 22200, 22100, False

Call lstErr_AddItem(lstErr, cmdContext, "cmdDSPOBJD_Compare 2/1: " & Nb_Rec)
lstDSPOBJD.Visible = True
lstDSPOBJD.ListIndex = 0
cmdDSPOBJD_Compare_Print.Enabled = True
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDSPOBJD_Compare_Print_Click()
MsgBox "2005.02.11 à revoir"
'prtDSPOBJD lstDSPOBJD
End Sub

Private Sub cmdDSPOBJD_Import_Click()
Dim wFileName_FTP As String
Dim paramDSPOBJD_Import As String

Me.Enabled = False: Me.MousePointer = vbHourglass
blnError = True
Nb_Rec = 0
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdDSPOBJD_Import : " & Time)

paramDSPOBJD_Import = Trim(txtDSPOBJD_Src1)
wFileName_FTP = Dir(paramDSPOBJD_Import)
If wFileName_FTP = "" Then Call lstErr_AddItem(lstErr, cmdContext, "! pas de fichier : " & paramDSPOBJD_Import): GoTo Error_Handle
Open paramDSPOBJD_Import For Input As #1
cmdDSPOBJD_Import_Src 22100
Close
Call lstErr_AddItem(lstErr, cmdContext, paramDSPOBJD_Import & " : " & Nb_Rec)

paramDSPOBJD_Import = Trim(txtDSPOBJD_Src2)
wFileName_FTP = Dir(paramDSPOBJD_Import)
If wFileName_FTP = "" Then Call lstErr_AddItem(lstErr, cmdContext, "! pas de fichier : " & paramDSPOBJD_Import): GoTo Error_Handle
Open paramDSPOBJD_Import For Input As #1
cmdDSPOBJD_Import_Src 22200

blnError = False

Call lstErr_AddItem(lstErr, cmdContext, paramDSPOBJD_Import & " : " & Nb_Rec)
GoTo fin

Error_Handle:
If Not blnAuto_ElpKMPgm Then MsgBox wFileName_FTP & ":" & Error, vbCritical, "cmdDSPOBJD_Import_Click"
fin:
Close
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDSPOBJD_Select_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDSPOBJD_Compare_Print.Enabled = False

If optDSPOBJD_22100 Then
    lstDSPOBJD_ElpKMSrc_Id = 22100
Else
    lstDSPOBJD_ElpKMSrc_Id = 22200
End If

lstDSPOBJD_Load lstDSPOBJD_ElpKMSrc_Id
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdElpKM_mdb_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If optElpKM_mdb_Export And optElpKM_mdb_Info Then cmdElpKm_mdb_Info_Export
If optElpKM_mdb_Clear And optElpKM_mdb_Info Then cmdElpKM_mdb_Info_Clear
If optElpKM_mdb_Import And optElpKM_mdb_Info Then cmdElpKm_mdb_Info_Import

If optElpKM_mdb_Export And optElpKM_mdb_Index Then cmdElpKm_mdb_Index_Export
If optElpKM_mdb_Clear And optElpKM_mdb_Index Then cmdElpKM_mdb_Index_Clear
If optElpKM_mdb_Import And optElpKM_mdb_Index Then cmdElpKm_mdb_Index_Import

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdElpKMPgm_Clear_Click()
X = MsgBox("Voulez-vous vraiment SUPPRIMER les INFOS PGM & les INDEX )   ?", vbQuestion + vbYesNo, Me.Caption)
If X = vbNo Then Exit Sub

mdbElpKMIndex.tableElpKMIndex_Open
X = "delete * from ElpKMIndex where Classe >= 11000 and Classe <= 12000"
MDB.Execute X

mdbElpKMInfo.tableElpKMInfo_Open
X = "delete * from ElpKMInfo where ElpKMSrc_Id >= 11000 and ElpKMSrc_Id <= 15999"
MDB.Execute X
End Sub

Private Sub cmdElpKMPgm_File_TableDef_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdElpKMPgm_File_TableDef_Create
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdElpKMPgm_Import_Auto_Click()
Dim K As Integer, lenX As Integer
Dim X As String, Xc As String
X = Trim(txtElpKMPgm_Import)
K = InStr(1, X, "DSPFFDY0")
If K < 0 Then
    MsgBox "Le nom du 1er fichier doit être ...'DSPFFDY0'.....", vbCritical
    Exit Sub
End If
If Dir(X) = "" Then
    MsgBox "Le fichier n'existe pas", vbCritical
    Exit Sub

End If

Xc = ""
lenX = Len(X)
If lenX > K + 7 Then Xc = mId$(X, K + 8, lenX - K - 7)
optElpKMPgm_Import_DSPFFDY0 = True
cmdElpKMPgm_Import_Click

X = mId$(X, 1, K - 1) & "DSPFDY0" & Xc
txtElpKMPgm_Import = X
optElpKMPgm_Import_DSPFDY0 = True
cmdElpKMPgm_Import_Click

Mid(X, K, 7) = "DSPFDY1"
txtElpKMPgm_Import = X
optElpKMPgm_Import_DSPFDY1 = True
cmdElpKMPgm_Import_Click

End Sub

Private Sub cmdElpKMPgm_Import_Click()
Dim wFileName_FTP As String
Dim paramElpKMPgm_Import As String

Me.Enabled = False: Me.MousePointer = vbHourglass
blnError = True
Nb_Rec = 0
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Import_Click : " & Time)

paramElpKMPgm_Import = Trim(txtElpKMPgm_Import)
wFileName_FTP = Dir(paramElpKMPgm_Import)
If wFileName_FTP = "" Then Call lstErr_AddItem(lstErr, cmdContext, "! pas de fichier : " & paramElpKMPgm_Import): GoTo Error_Handle

Open paramElpKMPgm_Import For Input As #1
If optElpKMPgm_Import_DSPFFDY0 Then cmdElpKMPgm_Import_DSPFFDY0
If optElpKMPgm_Import_DSPFDY0 Then cmdElpKMPgm_Import_DSPFDY0
If optElpKMPgm_Import_DSPFDY1 Then cmdElpKMPgm_Import_DSPFDY1

blnError = False

Call lstErr_AddItem(lstErr, cmdContext, "cmdSPLF_Click fin : " & Nb_Rec)
GoTo fin

Error_Handle:
If Not blnAuto_ElpKMPgm Then MsgBox wFileName_FTP & ":" & Error, vbCritical, "cmdElpKMPgm_Import_Click"
fin:
Close
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdElpKMPgm_Index_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
blnError = True
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Index début : " & Time)

If optElpKMPgm_Index_11000 Then Call cmdElpKMPgm_Index_Reset(11000, 11000, 11000)
If optElpKMPgm_Index_12000 Then Call cmdElpKMPgm_Index_Reset(12000, 12000, 12999)
blnError = False

Call lstErr_AddItem(lstErr, cmdContext, "cmdElpKMPgm_Index fin  : " & Time)
GoTo fin

Error_Handle:
If Not blnAuto_ElpKMPgm Then MsgBox ":" & Error, vbCritical, "cmdElpKMPgm_Index_Click"
fin:
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdElpKMPgm_JRN_CL_Click()
Dim xIn As String, dirJRN_F As String, wJRN_File As String
Dim Jfile As String
Dim xUpdate As String

Dim K As Integer
On Error GoTo Error_Handler

dirJRN_F = paramTemp_Folder & "\JRN\"
Me.Enabled = False: Me.MousePointer = vbHourglass
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Export: " & Time)
DoEvents

Open dirJRN_F & "CL.txt" For Output As #2


Open dirJRN_F & "ZMNUFIC0_JRN.txt" For Input As #3
Do Until EOF(3)
    Line Input #3, xIn
    K = InStr(1, xIn, " ")          ' nom fichier + libellé
    If K = 0 Then K = Len(xIn)      ' nom fichier seul
    If K > 2 Then

        mElpKMSrc_WHLIB = "SAB073"
        mElpKMSrc_WHFILE = Trim(mId$(xIn, 1, K))
        
        Print #2, Space$(25) & "CLRPFM SAB073JRN/" & mElpKMSrc_WHFILE
    End If
Loop

GoTo Exit_Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdElpKMPgm_JRN")
'---------------------------------------------------------
Exit_Sub:
'---------------------------------------------------------


Close

Me.MousePointer = vbDefault
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdElpKMPgm_JRN_Click()
Dim xIn As String, dirJRN_F As String, wJRN_File As String
Dim Jfile As String
Dim xUpdate As String

Dim K As Integer
On Error GoTo Error_Handler

dirJRN_F = paramTemp_Folder & "\JRN\"
Me.Enabled = False: Me.MousePointer = vbHourglass
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Export: " & Time)
DoEvents

Open dirJRN_F & "QDDSSRC_ftp.txt" For Output As #12
Print #12, "cd /d " & dirJRN_F
Print #12, "FTP I5A7"
Print #12, "I_LOULERGU"
Print #12, "I_LOULERGU"
Print #12, "CD BIAJRNSRC"

Open dirJRN_F & "QLBLSRC_ftp.txt" For Output As #14
Print #14, "cd /d " & dirJRN_F
Print #14, "FTP I5A7"
Print #14, "I_LOULERGU"
Print #14, "I_LOULERGU"
Print #14, "CD BIAJRNSRC"


Open dirJRN_F & "ZMNUFIC0_JRN.txt" For Input As #3
Do Until EOF(3)
    Line Input #3, xIn
    K = InStr(1, xIn, " ")          ' nom fichier + libellé
    If K = 0 Then K = Len(xIn)      ' nom fichier seul
    If K > 2 Then

        mElpKMSrc_WHLIB = "SAB073"
        mElpKMSrc_WHFILE = Trim(mId$(xIn, 1, K))
        
        mElpKMPgm_Library = srvDSPFFDY0_Library(mElpKMSrc_WHLIB)
        
        mElpKMSrc_Id = 12000 + mElpKMPgm_Library
        Call srvDSPFFDY0_lstAddItem(lstElpKMPgm_W, mElpKMSrc_Id, mElpKMSrc_WHFILE, False)
        
        mElpKMSrc_WHFILE_Y = Trim(mElpKMSrc_WHFILE)
        Mid$(mElpKMSrc_WHFILE_Y, 1, 1) = "Y"
        mElpKMSrc_WHFILE_J = mElpKMSrc_WHFILE_Y
        Mid$(mElpKMSrc_WHFILE_J, 1, 1) = "J"
        Jfile = Trim(mElpKMSrc_WHFILE_J)
        
        wJRN_File = dirJRN_F & "QDDSSRC\" & Jfile & ".txt"
        Open wJRN_File For Output As #2

               
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "DDS_JRN : " & Jfile): cmdElpKMPgm_Export_DDS_JRN
        Print #12, "put " & wJRN_File & " QDDSSRC." & Jfile
        Close 2
     
        Open dirJRN_F & "JXXXXXX0_U.txt" For Input As #5
        If Len(Jfile) < 9 Then
            xUpdate = "_U"
        Else
            xUpdate = "U"
        End If
        
        wJRN_File = dirJRN_F & "QLBLSRC\" & Jfile & xUpdate & ".txt"
        Open wJRN_File For Output As #4

        Do Until EOF(5)
            Line Input #5, xIn
            K = InStr(1, xIn, "JXXXXXX0")
            If K <> 0 Then
                Mid$(xIn, K, 8) = Jfile
                K = InStr(K, xIn, "JXXXXXX0")
                If K <> 0 Then Mid$(xIn, K, 8) = Jfile
            End If
            Print #4, xIn
            
        Loop
        
        Print #14, "put " & wJRN_File & " QLBLSRC." & Jfile & xUpdate
        Close 5
        Close 4
    End If
Loop

GoTo Exit_Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdElpKMPgm_JRN")
'---------------------------------------------------------
Exit_Sub:
'---------------------------------------------------------

Print #12, "quit"
Print #14, "quit"

Close

Me.MousePointer = vbDefault
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdElpKMPgm_JRN_VB_Click()
Dim xIn As String, dirJRN_F As String, wJRN_File As String
Dim Jfile As String
Dim xUpdate As String
Dim wApp As String
Dim K As Integer, I  As Integer
On Error GoTo Error_Handler

dirJRN_F = paramTemp_Folder & "\JRN\"
Me.Enabled = False: Me.MousePointer = vbHourglass
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Export: " & Time)
DoEvents
''wApp = "ADR"
wApp = UCase$(InputBox("Code application *** ", "ADR,AUT,... SWI)", wApp))
ReDim arrExport_JRN_VB(2000)
arrExport_JRN_VB_Nb = 0

Open dirJRN_F & "ZMNUFIC0_JRN.txt" For Input As #3
Do Until EOF(3)
    Line Input #3, xIn
    K = InStr(1, xIn, " ")          ' nom fichier + libellé
    If K = 0 Then K = Len(xIn)      ' nom fichier seul
    If K > 2 Then
    
        If mId$(xIn, 2, 3) = wApp Then
            arrExport_JRN_VB_Nb = arrExport_JRN_VB_Nb + 1
            arrExport_JRN_VB(arrExport_JRN_VB_Nb) = Trim(mId$(xIn, 1, K))
        End If
    End If
Loop

wJRN_File = dirJRN_F & "VB\J" & wApp & "_Srv.bas"
Open wJRN_File For Output As #2

'=========================================================================
Print #2, "'    Case " & Asc34 & "Z" & wApp & Asc34 & ": V = srvJ" & wApp & "_Sql(meJRNENT0, rsADO_X, fgSelect_D)"

Print #2, "'---------------------------------------------------------"
Print #2, "Option Explicit"
Print #2, "'---------------------------------------------------------"


For I = 1 To arrExport_JRN_VB_Nb
        
        mElpKMSrc_WHLIB = "SAB073"
        mElpKMSrc_WHFILE = arrExport_JRN_VB(I)
        
        mElpKMPgm_Library = srvDSPFFDY0_Library(mElpKMSrc_WHLIB)
        
        mElpKMSrc_Id = 12000 + mElpKMPgm_Library
        Call srvDSPFFDY0_lstAddItem(lstElpKMPgm_W, mElpKMSrc_Id, mElpKMSrc_WHFILE, False)
        
        mElpKMSrc_WHFILE_Y = Trim(mElpKMSrc_WHFILE)
        Mid$(mElpKMSrc_WHFILE_Y, 1, 1) = "Y"
        mElpKMSrc_WHFILE_J = mElpKMSrc_WHFILE_Y
        Mid$(mElpKMSrc_WHFILE_J, 1, 1) = "J"
        Jfile = Trim(mElpKMSrc_WHFILE_J)
        
        
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "JRN_Déclaration : " & Jfile)
        cmdElpKMPgm_Export_JRN_Déclaration
Next I
'=========================================================================
cmdElpKMPgm_Export_JRN_Sql wApp
'=========================================================================
For I = 1 To arrExport_JRN_VB_Nb
        
        mElpKMSrc_WHLIB = "SAB073"
        mElpKMSrc_WHFILE = arrExport_JRN_VB(I)
        
        mElpKMPgm_Library = srvDSPFFDY0_Library(mElpKMSrc_WHLIB)
        
        mElpKMSrc_Id = 12000 + mElpKMPgm_Library
        Call srvDSPFFDY0_lstAddItem(lstElpKMPgm_W, mElpKMSrc_Id, mElpKMSrc_WHFILE, False)
        
        mElpKMSrc_WHFILE_Y = Trim(mElpKMSrc_WHFILE)
        Mid$(mElpKMSrc_WHFILE_Y, 1, 1) = "Y"
        mElpKMSrc_WHFILE_J = mElpKMSrc_WHFILE_Y
        Mid$(mElpKMSrc_WHFILE_J, 1, 1) = "J"
        Jfile = Trim(mElpKMSrc_WHFILE_J)
        
        
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "JRN_ODBC : " & Jfile)
        cmdElpKMPgm_Export_JRN_GetBuffer_ODBC
Next I
'=========================================================================
For I = 1 To arrExport_JRN_VB_Nb
        
        mElpKMSrc_WHLIB = "SAB073"
        mElpKMSrc_WHFILE = arrExport_JRN_VB(I)
        
        mElpKMPgm_Library = srvDSPFFDY0_Library(mElpKMSrc_WHLIB)
        
        mElpKMSrc_Id = 12000 + mElpKMPgm_Library
        Call srvDSPFFDY0_lstAddItem(lstElpKMPgm_W, mElpKMSrc_Id, mElpKMSrc_WHFILE, False)
        
        mElpKMSrc_WHFILE_Y = Trim(mElpKMSrc_WHFILE)
        Mid$(mElpKMSrc_WHFILE_Y, 1, 1) = "Y"
        mElpKMSrc_WHFILE_J = mElpKMSrc_WHFILE_Y
        Mid$(mElpKMSrc_WHFILE_J, 1, 1) = "J"
        Jfile = Trim(mElpKMSrc_WHFILE_J)
        
        
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "JRN_fgDisplay : " & Jfile)
        cmdElpKMPgm_Export_JRN_fgDisplay
Next I

'=========================================================================
Call lstErr_AddItem(lstErr, cmdContext, "cmdElpKMPgm_Export: fin " & Time)
'=========================================================================
GoTo Exit_Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdElpKMPgm_JRN")
'---------------------------------------------------------
Exit_Sub:
'---------------------------------------------------------

Close

Me.MousePointer = vbDefault
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdElpKMPgm_Select_Click()
lstElpKMPgm_11000_Load
End Sub


Private Sub cmdPrint_Click()
Select Case SSTab1.Tab
    Case 0:
    Case 2: Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
    Case 3:
End Select

End Sub

Private Sub cmdSABPF_Clear_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Import : " & Time)
mdbElpKMInfo.tableElpKMInfo_Open
X = "delete * from ElpKMinfo where ElpKMSrc_Id = 21000"
MDB.Execute X
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Import : terminé")

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSABPF_CLP_Click()
On Error Resume Next

Me.Enabled = False: Me.MousePointer = vbHourglass
Me.MousePointer = vbHourglass
lstErr.Visible = True
lstSABPF.Visible = False

Call lstErr_Clear(lstErr, cmdContext, "cmdSABPF_CLP: " & Time)
DoEvents
If optSABPF_SAB073 Then
    paramFromLib = "SAB073T": paramToLib = "SAB073"
Else
    paramFromLib = "SAB073": paramToLib = "SAB073T"
End If

Open Trim(txtSABPF_CLP) For Output As #2

X80 = Space(13) & "PGM": Print #2, X80
X80 = Space(13) & "DCL        VAR(&FROMLIB) TYPE(*CHAR) LEN(10) VALUE(" & paramFromLib & ")": Print #2, X80
X80 = Space(13) & "DCL        VAR(&TOLIB) TYPE(*CHAR) LEN(10) VALUE(" & paramToLib & ")": Print #2, X80
X80 = "": Print #2, X80

If optSABPF_CPY Then cmdSABPF_CLP_CPY
If optSABPF_CLR Then cmdSABPF_CLP_CLR
If optSABPF_PAR Then cmdSABPF_CLP_CPY
If optSABPF_QRY Then cmdSABPF_CLP_QRY

X80 = "": Print #2, X80
X80 = Space(13) & "ENDPGM": Print #2, X80

Close
lstSABPF.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "cmdSABPF_CLP :  Fin")
Me.MousePointer = vbDefault
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdSABPF_False_Click()
Dim I As Integer, wIndex As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

On Error Resume Next
Call lstErr_Clear(lstErr, cmdContext, "Tout désélectionner > " & Time)

For I = 1 To arrSABPF_Nb
    If mId$(arrSABPF(I).Memo, paramSABPF_K, 3) = paramSABPF_opt Then
        Mid$(arrSABPF(I).Memo, paramSABPF_K, 3) = "xxx"
        arrSABPF(I).Method = constUpdate
    End If
Next I
arrSABPF_Display

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSABPF_Gen_Click()
Dim xlen As Integer, X As String, wFile As String
Dim blnOk As Boolean
Dim K1 As Integer

On Error Resume Next

Me.Enabled = False: Me.MousePointer = vbHourglass
blnError = True
Nb_Rec = 0
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdSABPF_Gen : " & Time)

meElpKMInfo.Method = "AddNew"
meElpKMInfo.ElpKMSrc_Id = 21000
meElpKMInfo.Pass = 1000

xElpKMInfo.Method = "Seek>="
xElpKMInfo.ElpKMSrc_Id = 11000
xElpKMInfo.ID = ""
intReturn = tableElpKMInfo_Read(xElpKMInfo)
xElpKMInfo.Method = "Seek>"

Do
    If intReturn = 0 Then
        If xElpKMInfo.ElpKMSrc_Id <> 11000 Then
            intReturn = -1
        Else
            MsgTxt = Space$(34) & xElpKMInfo.Memo
            MsgTxtIndex = 0
            srvDSPFDY0_GetBuffer meDSPFDY0
            If Trim(meDSPFDY0.ATFATR) = "PF" Then

                Nb_Rec = Nb_Rec + 1
                meElpKMInfo.ID = mId$(xElpKMInfo.ID, 11, 10)
                meElpKMInfo.Description = xElpKMInfo.Description
                If mId$(meElpKMInfo.ID, 5, 3) = "TAB" Then
                    meElpKMInfo.Memo = "CPYxxxPAR"
                Else
                    meElpKMInfo.Memo = "CPYCLRxxx"
                End If
                intReturn = tableElpKMInfo_Update(meElpKMInfo)
            End If
            intReturn = tableElpKMInfo_Read(xElpKMInfo)
            
        End If
    End If
    
  
Loop While intReturn = 0
Call lstErr_AddItem(lstErr, cmdContext, "cmdSABPF_Gen, fin : " & Nb_Rec)
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSABPF_Select_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSABPF_Select.Enabled = False
If optSABPF_CPY Then paramSABPF_K = 1: paramSABPF_opt = "CPY"
If optSABPF_CLR Then paramSABPF_K = 4: paramSABPF_opt = "CLR"
If optSABPF_PAR Then paramSABPF_K = 7: paramSABPF_opt = "PAR"
If optSABPF_QRY Then paramSABPF_K = 1: paramSABPF_opt = "QRY"

arrSABPF_Load
arrSABPF_Display

fraSABPF_CLP.Enabled = True
txtSABPF_CLP = "C:\Temp\SABPF_" & paramSABPF_opt & ".txt"
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSABPF_True_Click()
Dim I As Integer, wIndex As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

On Error Resume Next
Call lstErr_Clear(lstErr, cmdContext, "Tout sélectionner > " & Time)

For I = 1 To arrSABPF_Nb
    If mId$(arrSABPF(I).Memo, paramSABPF_K, 3) <> paramSABPF_opt Then
        Mid$(arrSABPF(I).Memo, paramSABPF_K, 3) = paramSABPF_opt
        arrSABPF(I).Method = constUpdate
    End If
Next I
arrSABPF_Display

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdSABPF_Update_Click()
For I = 1 To arrSABPF_Nb
    If arrSABPF(I).Method = constUpdate Then
        xElpKMInfo = arrSABPF(I)
        xElpKMInfo.Method = "Seek="
        If tableElpKMInfo_Read(xElpKMInfo) = 0 Then
            arrSABPF(I).Method = constUpdate
            intReturn = tableElpKMInfo_Update(arrSABPF(I))
        End If

    End If
Next I

End Sub

Private Sub lstDSPOBJD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xElpKMInfo.Method = "Seek="
If optDSPOBJD_22100 Then
    xElpKMInfo.ElpKMSrc_Id = 22100
Else
    xElpKMInfo.ElpKMSrc_Id = 22200
End If

xElpKMInfo.ID = mId$(lstDSPOBJD, 1, 18)
If tableElpKMInfo_Read(xElpKMInfo) = 0 Then
    MsgTxt = Space$(34) & xElpKMInfo.Memo
    MsgTxtIndex = 0
    srvDSPOBJDY0_GetBuffer meDSPOBJDY0
    srvDSPOBJDY0_ElpDisplay meDSPOBJDY0
Else
    MsgBox "Erreur lecture " & Error, vbCritical
End If


End Sub


Private Sub lstElpKMPgm_File_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.Enabled Then
    If lstElpKMPgm_File.ListIndex > -1 Then
        If Button = vbLeftButton Then
            mElpKMSrc_WHLIB = mId$(lstElpKMPgm_File, 1, 10)
            mElpKMSrc_WHFILE = mId$(lstElpKMPgm_File, 12, 10)
            txtElpKMPgm_Export = paramTemp_Folder & "\SAB\" & Trim(mElpKMSrc_WHFILE) & ".txt"
            Me.PopupMenu mnuElpKMPgm_File, vbPopupMenuLeftButton
        End If
    End If
End If

End Sub

Private Sub lstSABPF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Enabled = False: Me.MousePointer = vbHourglass
arrSABPF_Select
Me.Enabled = True: Me.MousePointer = 0

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
If SSTab1.Tab = 2 Then
    If Not fraSABPF_Select.Enabled Then fraSABPF_Select.Enabled = True: lstSABPF.Clear: Exit Sub
Else
    If currentAction = "" Then
       
    Else
        X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
        If X = vbYes Then
            currentAction = ""
        Else
            Exit Sub
        End If
    End If
End If
End Sub

Public Sub cmdContext_Return()
Select Case SSTab1.Tab
    Case 0: cmdElpKMPgm_Select_Click
    Case 2: cmdSABPF_Select_Click
    Case 3: cmdDSPOBJD_Select_Click
    Case Else:    SendKeys "{TAB}"
End Select
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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


Public Sub txt_X()
'Call txt_GotFocus(txt)
'KeyAscii = convUCase(KeyAscii)
'Call txt_LostFocus(txt)

'Call txt_GotFocus(txt)
'If XopDevise(2).maxD = 0 Then
'    Call num_KeyAscii(KeyAscii)
'Else
'    Call num_KeyAsciiD(KeyAscii, txt)
'End If
'Call txt_LostFocus(txt)
'Call txt_LostFocus(txt)

End Sub


Public Sub Auto_ElpKMPgm()
'cmdElpKMPgm_AS400_Click
Unload Me
End Sub




Public Sub lstElpKMPgm_11000_Load()
Dim xlen As Integer, X As String, wFile As String
Dim blnOk As Boolean
Dim blnGenX As Boolean, K As Integer

On Error Resume Next

X = Trim(txtElpKMPgm_Select)
If mId$(X, 1, 1) <> "*" Then
    blnGenX = False
'    xElpKMInfo.Id = X
Else
    blnGenX = True
    Mid$(X, 1, 1) = " "
    X = Trim(X)
'    xElpKMInfo.Id = ""
End If
xlen = Len(X)

lstElpKMPgm_File.Clear

xElpKMInfo.Method = "Seek>="
xElpKMInfo.ElpKMSrc_Id = 11000
xElpKMInfo.ID = ""
intReturn = tableElpKMInfo_Read(xElpKMInfo)
xElpKMInfo.Method = "MoveNext"

Do
    If intReturn = 0 Then
        If xElpKMInfo.ElpKMSrc_Id <> 11000 Then
            intReturn = -1
        Else
            wFile = mId$(xElpKMInfo.ID, 11, 10)
            blnOk = True
            If xlen > 0 Then
                If blnGenX Then
                    K = InStr(1, wFile, X)
                    If K = 0 Then blnOk = False
                Else
                    If X <> mId$(wFile, 1, xlen) Then blnOk = False
                End If
                
            End If
            'If blnOK Then lstElpKMPgm_File.AddItem wFile & Chr$(9) & mId$(xElpKMInfo.Id, 1, 10) & " " & xElpKMInfo.description
            If blnOk Then lstElpKMPgm_File.AddItem mId$(xElpKMInfo.ID, 1, 10) & " " & xElpKMInfo.Description
            intReturn = tableElpKMInfo_Read(xElpKMInfo)
            
        End If
    End If
    
  
Loop While intReturn = 0
If lstElpKMPgm_File.ListCount > 0 Then lstElpKMPgm_File.ListIndex = 0
txtElpKMPgm_Select.SetFocus
End Sub

Public Sub arrSABPF_Load()

Dim xlen As Integer, X As String, wFile As String
Dim blnOk As Boolean
Dim blnGenX As Boolean, K As Integer

On Error Resume Next

X = Trim(txtSABPF_Select)
If mId$(X, 1, 1) <> "*" Then
    blnGenX = False
    xElpKMInfo.ID = X
Else
    blnGenX = True
    Mid$(X, 1, 1) = " "
    X = Trim(X)
    xElpKMInfo.ID = ""
End If
xlen = Len(X)

lstSABPF.Clear
arrSABPF_Nb = 0

xElpKMInfo.Method = "Seek>="
xElpKMInfo.ElpKMSrc_Id = 21000
intReturn = tableElpKMInfo_Read(xElpKMInfo)
xElpKMInfo.Method = "MoveNext"

Do
    If intReturn = 0 Then
        If xElpKMInfo.ElpKMSrc_Id <> 21000 Then
            intReturn = -1
        Else
            blnOk = True
            If xlen > 0 Then
                If blnGenX Then
                    K = InStr(1, xElpKMInfo.ID, X)
                    If K = 0 Then blnOk = False
                Else
                    If X <> mId$(xElpKMInfo.ID, 1, xlen) Then blnOk = False
                End If
                
            End If
            
            If blnOk Then
                arrSABPF_Nb = arrSABPF_Nb + 1
                arrSABPF(arrSABPF_Nb) = xElpKMInfo
                arrSABPF(arrSABPF_Nb).Method = ""
            End If
            intReturn = tableElpKMInfo_Read(xElpKMInfo)
            
        End If
    End If
    
Loop While intReturn = 0
End Sub

Public Sub lstDSPOBJD_Load(lElpKMSrc_Id As Long)
Dim xlen As Integer, X As String, wFile As String
Dim blnOk As Boolean, K As Integer
Dim blnGenX As Boolean

On Error Resume Next

lstDSPOBJD.Visible = False
lstDSPOBJD.Clear
xElpKMInfo.Method = "Seek>="
xElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id

X = Trim(txtDSPOBJD_Select)
If mId$(X, 1, 1) <> "*" Then
    blnGenX = False
    xElpKMInfo.ID = X
Else
    blnGenX = True
    Mid$(X, 1, 1) = " "
    X = Trim(X)
    xElpKMInfo.ID = ""
End If
xlen = Len(X)

intReturn = tableElpKMInfo_Read(xElpKMInfo)
xElpKMInfo.Method = "MoveNext"

Do
    If intReturn = 0 Then
        If xElpKMInfo.ElpKMSrc_Id <> lElpKMSrc_Id Then
            intReturn = -1
        Else
            blnOk = True
            If xlen > 0 Then
                If blnGenX Then
                    K = InStr(1, xElpKMInfo.ID, X)
                    If K = 0 Then blnOk = False
                Else
                    If X <> mId$(xElpKMInfo.ID, 1, xlen) Then blnOk = False
                End If
                
            End If
            
            If blnOk Then lstDSPOBJD.AddItem xElpKMInfo.ID & xElpKMInfo.Description

            intReturn = tableElpKMInfo_Read(xElpKMInfo)
            
        End If
    End If
    
Loop While intReturn = 0
lstDSPOBJD.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "Affichage DSPOBJD (22***): " & Time)
lstDSPOBJD.ListIndex = 0

End Sub

Public Sub lstDSPFDY2_Load(lElpKMSrc_Id As Long)
Dim xlen As Integer, X As String, wFile As String
Dim blnOk As Boolean, K As Integer
Dim blnGenX As Boolean

On Error Resume Next

lstDSPFDY2.Visible = False
lstDSPFDY2.Clear
xElpKMInfo.Method = "Seek>="
xElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id

X = Trim(txtDSPFDY2_Select)
If mId$(X, 1, 1) <> "*" Then
    blnGenX = False
    xElpKMInfo.ID = X
Else
    blnGenX = True
    Mid$(X, 1, 1) = " "
    X = Trim(X)
    xElpKMInfo.ID = ""
End If
xlen = Len(X)

intReturn = tableElpKMInfo_Read(xElpKMInfo)
xElpKMInfo.Method = "MoveNext"

Do
    If intReturn = 0 Then
        If xElpKMInfo.ElpKMSrc_Id <> lElpKMSrc_Id Then
            intReturn = -1
        Else
            blnOk = True
            If xlen > 0 Then
                If blnGenX Then
                    K = InStr(1, xElpKMInfo.ID, X)
                    If K = 0 Then blnOk = False
                Else
                    If X <> mId$(xElpKMInfo.ID, 1, xlen) Then blnOk = False
                End If
                
            End If
            
            If blnOk Then lstDSPFDY2.AddItem xElpKMInfo.ID & xElpKMInfo.Description

            intReturn = tableElpKMInfo_Read(xElpKMInfo)
            
        End If
    End If
    
Loop While intReturn = 0
lstDSPFDY2.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "Affichage DSPFDY2 (23***): " & Time)
lstDSPFDY2.ListIndex = 0

End Sub


Public Sub cmdDSPOBJD_Compare_Memo(lElpKMSrc_Id1 As Long, lElpKMSrc_Id2 As Long, blnMemo_Compare As Boolean)
Dim xlen As Integer, X As String, wFile As String
Dim blnOk As Boolean, K As Integer
Dim blnGenX As Boolean
Dim xErr As String * 12
On Error Resume Next

meElpKMInfo.Method = "Seek>="
meElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id1
meElpKMInfo.ID = ""
intReturn = tableElpKMInfo_Read(meElpKMInfo)
meElpKMInfo.Method = "Seek>"

xElpKMInfo.Method = "Seek="
xElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id2
xElpKMInfo.Memo = Space$(memoDSPOBJDY0Len)
Do
    If intReturn = 0 Then
        If meElpKMInfo.ElpKMSrc_Id <> lElpKMSrc_Id1 Then
            intReturn = -1
        Else
            blnOk = True
            xElpKMInfo.ID = meElpKMInfo.ID
            Call lstErr_ChangeLastItem(lstErr, cmdContext, meElpKMInfo.ID)
            DoEvents
            If tableElpKMInfo_Read(xElpKMInfo) <> 0 Then
                blnOk = False: xErr = "? " & mId$(xElpKMInfo.Memo, 14, 10)
            Else
               If blnMemo_Compare Then
                  '  MsgTxt = Space$(34) & meElpKMInfo.Memo
                   ' MsgTxtIndex = 0
                   ' srvDSPOBJDY0_GetBuffer meDSPOBJDY0
                    
                   ' MsgTxt = Space$(34) & xElpKMInfo.Memo
                   ' MsgTxtIndex = 0
                   ' srvDSPOBJDY0_GetBuffer xDSPOBJDY0
                   ' If blnOk And meDSPOBJDY0.ODOBAT <> xDSPOBJDY0.ODOBAT Then blnOk = False: xErr = "ODOBAT"
                    If blnOk And mId$(meElpKMInfo.Memo, 24, 29) <> mId$(xElpKMInfo.Memo, 24, 29) Then blnOk = False: xErr = "? TYP"
                    If blnOk And mId$(meElpKMInfo.Memo, 266, 43) <> mId$(xElpKMInfo.Memo, 266, 43) Then blnOk = False: xErr = "? SRC"
                    If blnOk Then
                        If mId$(meElpKMInfo.Memo, 34, 8) = "*FILE   " Then
                            If mId$(meElpKMInfo.Memo, 64, 50) <> mId$(xElpKMInfo.Memo, 64, 50) Then blnOk = False: xErr = "? Nom"
                        Else
                            If mId$(meElpKMInfo.Memo, 53, 61) <> mId$(xElpKMInfo.Memo, 53, 61) Then blnOk = False: xErr = "? Size,Nom"
                        End If
                    End If
                End If
            End If

            If Not blnOk Then
                Nb_Rec = Nb_Rec + 1
                lstDSPOBJD.AddItem meElpKMInfo.ID & xErr & " " & meElpKMInfo.Description
            End If
            
            intReturn = tableElpKMInfo_Read(meElpKMInfo)
            
        End If
    End If
    
Loop While intReturn = 0

End Sub

Public Sub arrSABPF_Display()
Dim I As Integer, wIndex As Integer

On Error Resume Next
Call lstErr_Clear(lstErr, cmdContext, "Affichage SABPF (21000) > " & Time)
lstSABPF.Visible = False
lstSABPF.Clear
For I = 1 To arrSABPF_Nb
    xElpKMInfo = arrSABPF(I)
    
    wIndex = lstSABPF.ListCount
    lstSABPF.AddItem xElpKMInfo.Description
    If mId$(xElpKMInfo.Memo, paramSABPF_K, 3) = paramSABPF_opt Then lstSABPF.Selected(wIndex) = True
    
Next I

lstSABPF.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "Affichage SABPF (21000): " & Time)
lstSABPF.ListIndex = 0
End Sub

Private Sub cmdElpKMPgm_Index_Reset(lClasse As Long, lElpKMSrc_Id As Long, lElpKMSrc_Max As Long)

Dim blnUpdate As Boolean, bln11000 As Boolean
Dim kIn As Integer, seq As Long
Dim xIn2 As String, X As String


On Error GoTo Error_Handle
X = MsgBox("Voulez-vous vraiment mettre à jour les INDEX ) (delete/AddNew)  ?", vbQuestion + vbYesNo, Me.Caption)
If X = vbNo Then Exit Sub

Call lstErr_AddItem(lstErr, cmdContext, "cmdElpKMPgm_Index_Reset: début"): DoEvents

If lElpKMSrc_Id = 11000 Then
    bln11000 = True
Else
    bln11000 = False
End If

mdbElpKMIndex.tableElpKMIndex_Open
X = "delete * from ElpKMIndex where Classe = " & lClasse
MDB.Execute X


xElpKMInfo.Method = "Seek>="
xElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id
xElpKMInfo.ID = ""
intReturn = tableElpKMInfo_Read(xElpKMInfo)
xElpKMInfo.Method = "MoveNext"


recElpKMIndex_Init meElpKMIndex
meElpKMIndex.Method = constAddNew
meElpKMIndex.Classe = lClasse

seq = 0
''Exit Sub
Do
    If intReturn = 0 Then
        If xElpKMInfo.ElpKMSrc_Id > lElpKMSrc_Max Then
            intReturn = -1
        Else
            seq = seq + 1
            If seq Mod 100 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, cmdContext, xElpKMInfo.ID): DoEvents
            
            MsgTxt = Space$(34) & xElpKMInfo.Memo
            MsgTxtIndex = 0
            If bln11000 Then
                srvDSPFDY0_GetBuffer meDSPFDY0
                xIn2 = Text_LCase(meDSPFDY0.ATFILE & " " & meDSPFDY0.ATTXT)
            Else
                srvDSPFFDY0_GetBuffer meDSPFFDY0
                xIn2 = Text_LCase(meDSPFFDY0.WHFILE & " " & meDSPFFDY0.WHFTXT)
            End If
            
            If xIn2 <> "" Then
                    kIn = 0
                    
                    meElpKMIndex.ElpKMSrc_Id = xElpKMInfo.ElpKMSrc_Id
        
                    blnUpdate = True
                    
                    Do
                        X = Text_KeyWord(xIn2, kIn, False)
                    
                        If X <> "" Then
                            meElpKMIndex.ID = X
        
                            meElpKMIndex.Method = "Seek="
                            If tableElpKMIndex_Read(meElpKMIndex) = 0 Then
                                meElpKMIndex.Method = constUpdate
                                meElpKMIndex.Memo = meElpKMIndex.Memo & xElpKMInfo.ID
                            Else
                                meElpKMIndex.Method = constAddNew
                                meElpKMIndex.Memo = xElpKMInfo.ID
                            End If
                            
                            dbElpKMIndex_Update meElpKMIndex
        
                           ' Print #2, X, lMNUOPTCOD, xIn
                        Else
                            blnUpdate = False
                        End If
                        
                    Loop While blnUpdate
                End If
                
            End If
            intReturn = tableElpKMInfo_Read(xElpKMInfo)
        End If
    'End If
    
  
Loop While intReturn = 0

Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Index_Reset, fin : " & seq)


Exit Sub

Error_Handle:
 MsgBox "erreur : cmdElpKMPgm_Index_Reset" & xElpKMInfo.ID, vbCritical, Error


End Sub

Public Sub cmdElpKMPgm_Export_DDS()
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id
X80 = Space(5) & "A*" & String(64, "-"): Print #2, X80
X80 = Space(5) & "A* " & mElpKMSrc_WHFILE_Y & Space(41) & dateImp(DSys): Print #2, X80
X80 = Space(5) & "A*" & String(64, "-"): Print #2, X80
X80 = Space(5) & "A" & Space(10) & "R F" & mId$(mElpKMSrc_WHFILE_Y, 2, 7)
    Mid$(X80, 45, 6) = "TEXT('"
    Mid$(X80, 51, 17) = mElpKMSrc_WHFILE_Y
    Mid$(X80, 70, 2) = "')"
Print #2, X80

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        X80 = ""
        Mid$(X80, 6, 1) = "A"
        Mid$(X80, 19, 10) = meDSPFFDY0.WHFLDE
        Mid$(X80, 35, 1) = "A"   'meDSPFFDY0.WHFLDT
        If meDSPFFDY0.WHFLDT = "A" Then
            Call strMoveR(meDSPFFDY0.WHFLDB, X80, 30, 5)
        Else
            Call strMoveR(meDSPFFDY0.WHFLDD + 1, X80, 30, 5)
        End If
        Mid$(X80, 45, 8) = "COLHDG('"
        Mid$(X80, 53, 17) = meDSPFFDY0.WHCHD1
        Mid$(X80, 70, 2) = "')"
        Print #2, X80
    End If
    
    DoEvents

Next I

End Sub
Public Sub cmdElpKMPgm_Export_DDS_JRN()
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id
X80 = Space(5) & "A*" & String(64, "-"): Print #2, X80
X80 = Space(5) & "A* " & mElpKMSrc_WHFILE_J & Space(41) & dateImp(DSys): Print #2, X80
X80 = Space(5) & "A*" & String(64, "-"): Print #2, X80
X80 = Space(5) & "A"
    Mid$(X80, 45, 6) = "UNIQUE"
Print #2, X80
X80 = Space(5) & "A" & Space(10) & "R F" & mId$(mElpKMSrc_WHFILE_J, 2, 7)
    Mid$(X80, 45, 6) = "TEXT('"
    Mid$(X80, 51, 17) = mElpKMSrc_WHFILE_J
    Mid$(X80, 70, 2) = "')"
Print #2, X80

X80 = Space(5) & "A            JORCV          7  0       TEXT('RECEVEUR     ')": Print #2, X80
X80 = Space(5) & "A            JOSEQN        10  0       TEXT('SEQUENCE     ')": Print #2, X80
X80 = Space(5) & "A            JRNBIATRN     10  0       TEXT('REF TRANSACTION')": Print #2, X80

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        X80 = ""
        Mid$(X80, 6, 1) = "A"
        Mid$(X80, 19, 10) = meDSPFFDY0.WHFLDE
        Mid$(X80, 35, 1) = meDSPFFDY0.WHFLDT
        If meDSPFFDY0.WHFLDT = "A" Then
            Call strMoveR(meDSPFFDY0.WHFLDB, X80, 30, 5)
        Else
            Call strMoveR(meDSPFFDY0.WHFLDD, X80, 30, 5)
            Call strMoveR(meDSPFFDY0.WHFLDP, X80, 36, 2)
        End If
        Mid$(X80, 45, 8) = "COLHDG('"
        Mid$(X80, 53, 17) = meDSPFFDY0.WHCHD1
        Mid$(X80, 70, 2) = "')"
        Print #2, X80
    End If
    
    DoEvents

Next I

X80 = Space(5) & "A*": Print #2, X80

X80 = Space(5) & "A          K JORCV": Print #2, X80
X80 = Space(5) & "A          K JOSEQN": Print #2, X80


End Sub

Public Sub cmdElpKMPgm_Export_CBL_ZY()
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

X80 = "      * " & meDSPFFDY0.WHFILE & " => " & mElpKMSrc_WHFILE_Y
Print #2, X80
X80 = "      *" & String(63, "-")
Print #2, X80

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        X80 = ""
        Mid$(X80, 12, 4) = "MOVE"
        Mid$(X80, 17, 10) = meDSPFFDY0.WHFLDE
        Mid$(X80, 28, 5) = "OF L-"
        Mid$(X80, 33, 10) = meDSPFFDY0.WHFILE
        Mid$(X80, 44, 2) = "TO"
        Mid$(X80, 47, 10) = meDSPFFDY0.WHFLDE
        Mid$(X80, 58, 5) = "OF L-"
        Mid$(X80, 63, 10) = mElpKMSrc_WHFILE_Y
        Mid$(X80, 71, 1) = "."
        
        If meDSPFFDY0.WHFLDT <> "A" Then
            X80B = X80
            Mid$(X80B, 47, 27) = Space$(27)
            Mid$(X80B, 47, 27) = "NV" & meDSPFFDY0.WHFLDP & "S"
            Print #2, X80B
            Mid$(X80, 17, 26) = Space$(27)
            Mid$(X80, 17, 26) = "NX" & meDSPFFDY0.WHFLDD & "S"
        End If
        
        Print #2, X80
    End If
    
    DoEvents

Next I

End Sub
Public Sub cmdElpKMPgm_Export_CBL_YZ()
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

X80 = "      * " & mElpKMSrc_WHFILE_Y & " => " & meDSPFFDY0.WHFILE
Print #2, X80
X80 = "      *" & String(63, "-")
Print #2, X80
X80 = "      * !!!!!!! vérifier numérique avec décimales & NUM SIGNE"
Print #2, X80
X80 = "      *" & String(63, "-")
Print #2, X80

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        X80 = ""
        Mid$(X80, 12, 4) = "MOVE"
        Mid$(X80, 17, 10) = meDSPFFDY0.WHFLDE
        Mid$(X80, 28, 5) = "OF L-"
        Mid$(X80, 33, 10) = mElpKMSrc_WHFILE_Y
        Mid$(X80, 44, 2) = "TO"
        Mid$(X80, 47, 10) = meDSPFFDY0.WHFLDE
        Mid$(X80, 58, 5) = "OF L-"
        Mid$(X80, 63, 10) = meDSPFFDY0.WHFILE
        Mid$(X80, 71, 1) = "."
        
        If meDSPFFDY0.WHFLDT <> "A" Then
            X80B = X80
            Mid$(X80B, 47, 27) = Space$(27)
            Mid$(X80B, 47, 27) = "XN" & Format$(meDSPFFDY0.WHFLDD, "00") & "S"
            Print #2, X80B
            Mid$(X80, 17, 26) = Space$(27)
            Mid$(X80, 17, 26) = "XN" & Format$(meDSPFFDY0.WHFLDD, "00")
        End If
        
        Print #2, X80
    End If
    
    DoEvents

Next I

End Sub

Public Sub cmdElpKMPgm_Export_VB()

Print #2, "'---------------------------------------------------------"
Print #2, "Option Explicit"
Print #2, "'---------------------------------------------------------"
'Print #2, "Public Const const" & mElpKMSrc_WHFILE_Y & " = " & Chr$(34) & mElpKMSrc_WHFILE_Y & Chr$(34)

Print #2, "Type type" & mElpKMSrc_WHFILE_Y
'Print #2, "    Obj                     As String * 12"
'Print #2, "    Method                  As String * 12"
'Print #2, "    Err                     As String * 10"
    
cmdElpKMPgm_Export_VB_Déclaration


End Sub


Public Sub cmdElpKMPgm_Export_JRN_Déclaration()

Print #2, " "
Print #2, "Type type" & mElpKMSrc_WHFILE_J
Print #2, "    JORCV                   As long"
Print #2, "    JOSEQN                  As long"
Print #2, "    JRNBIATRN               As long"
 Print #2, " "
   
cmdElpKMPgm_Export_VB_Déclaration

Print #2, "Public x" & mElpKMSrc_WHFILE_J & " as type" & mElpKMSrc_WHFILE_J


End Sub

Public Sub cmdElpKMPgm_Export_JRN_Sql(wApp As String)
Dim I As Integer, Jfile As String

Print #2, "Public Function srvJ" & wApp & "_Sql(lJRNENT0 As typeJRNENT0, rsADO As ADODB.Recordset, fgDisplay As MSFlexGrid)"
Print #2, "Dim V"
Print #2, "Dim xSql As String"
Print #2, ""
Print #2, "V = Null"
Print #2, "Do While Not rsADO.EOF"

Print #2, "    Select Case Trim(lJRNENT0.JOOBJ)"
For I = 1 To arrExport_JRN_VB_Nb
        
        mElpKMSrc_WHLIB = "SAB073"
        mElpKMSrc_WHFILE = arrExport_JRN_VB(I)
        
        ''mElpKMPgm_Library = srvDSPFFDY0_Library(mElpKMSrc_WHLIB)
        
        ''mElpKMSrc_Id = 12000 + mElpKMPgm_Library
        ''Call srvDSPFFDY0_lstAddItem(lstElpKMPgm_W, mElpKMSrc_Id, mElpKMSrc_WHFILE, False)
        
        mElpKMSrc_WHFILE_Y = Trim(mElpKMSrc_WHFILE)
        ''''Mid$(mElpKMSrc_WHFILE_Y, 1, 1) = "Y"
        mElpKMSrc_WHFILE_J = mElpKMSrc_WHFILE_Y
        Mid$(mElpKMSrc_WHFILE_J, 1, 1) = "J"
        Jfile = Trim(mElpKMSrc_WHFILE_J)
        Print #2, "        Case " & Asc34 & mElpKMSrc_WHFILE_Y & Asc34 & ": Call srv" & Jfile & "_GetBuffer_ODBC(rsADO, x" & Jfile & "):  srv" & Jfile & "_fgDisplay x" & Jfile & ", fgDisplay"
        
        
Next I
    
Print #2, "  End Select"
Print #2, "    rsADO.MoveNext"
Print #2, "Loop"
Print #2, "srvJ" & wApp & "_Sql = V"
Print #2, ""
Print #2, "End Function"

End Sub

Public Sub cmdElpKMPgm_Export_VB_GetBuffer()
Dim X As String, Xa As String, xB As String, Xc As String, Xd As String
Dim wPos As Integer, wLen As Integer
arrExport_Nb = 0
wPos = 1
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0
        If meDSPFFDY0.WHFLDT = "A" Then
            wLen = meDSPFFDY0.WHFLDB
        Else
            wLen = meDSPFFDY0.WHFLDD + 1
        End If

        Xa = Chr$(9) & "rec" & mElpKMSrc_WHFILE_Y & "." & Trim(meDSPFFDY0.WHFLDE) & " = "
        ''Mid$(Xa, 5, 1) = "X"
        Xc = "mid$(MsgTxt , k + " & wPos & " , " & wLen & ")"
        
        arrExport_Nb = arrExport_Nb + 1
        arrExport_WHFLDE(arrExport_Nb) = Trim(meDSPFFDY0.WHFLDE)
        arrExport_WHCHD1(arrExport_Nb) = Trim(meDSPFFDY0.WHCHD1)
        arrExport_WHCHD2(arrExport_Nb) = Trim(meDSPFFDY0.WHCHD2)
        arrExport_Pos(arrExport_Nb) = wPos
        arrExport_Len(arrExport_Nb) = wLen
        
        Select Case meDSPFFDY0.WHFLDT
            Case "A": xB = "": Xd = ""
            Case "B": xB = "Cint(Val(": Xd = "))"
            Case Else:
                Select Case meDSPFFDY0.WHFLDP
                    Case 0: xB = "Clng(Val(": Xd = "))"
                    Case 2: xB = "Ccur(Val(": Xd = "))/ 100"
                    Case Else: xB = "Cdbl(Val(": Xd = "))/ 1" & String$(meDSPFFDY0.WHFLDP, "0")
                End Select
        End Select
        
        X = Xa & xB & Xc & Xd
        Print #2, X
        wPos = wPos + wLen
    DoEvents


    End If
    

Next I

End Sub

Public Sub cmdElpKMPgm_Export_VB_Init()
Dim wPos As Integer, wLen As Integer
Dim recName As String
Dim xB As String

recName = "rs" & mElpKMSrc_WHFILE_Y & "."
Print #2, "Public Sub srv" & mElpKMSrc_WHFILE_Y & "_Init(rs" & mElpKMSrc_WHFILE_Y & " As type" & mElpKMSrc_WHFILE_Y & ")"
'Print #2, recName & "Obj = " & Chr$(34) & mElpKMSrc_WHFILE_Y & Chr$(34)
'Print #2, recName & "Method = " & Chr$(34) & Chr$(34)
'Print #2, recName & "Err = " & Chr$(34) & Chr$(34)

meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        Select Case meDSPFFDY0.WHFLDT
            Case "A":   xB = "= " & Chr$(34) & Chr$(34)
            Case Else:  xB = "= 0"
        End Select
        
        Print #2, "rs" & mElpKMSrc_WHFILE_Y & "." & Trim(meDSPFFDY0.WHFLDE) & xB

    DoEvents

    End If
    

Next I

Print #2, "end sub"

End Sub

Public Sub cmdElpKMPgm_Export_VB_GetBuffer_ODBC()
Dim wPos As Integer, wLen As Integer
Dim recName As String
Dim xB As String

recName = "rec" & mElpKMSrc_WHFILE_Y & "."
Print #2, "Public Function srv" & mElpKMSrc_WHFILE_Y & "_GetBuffer_ODBC(rsADO as ADODB.Recordset,rec" & mElpKMSrc_WHFILE_Y & " As type" & mElpKMSrc_WHFILE_Y & ")"

Print #2, "On Error GoTo Error_Handler"
Print #2, "srv" & mElpKMSrc_WHFILE_Y & "_GetBuffer_ODBC = Null"

meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        xB = Trim(meDSPFFDY0.WHFLDE)
        Print #2, "rec" & mElpKMSrc_WHFILE_Y & "." & xB & " = rsADO(" & Chr$(34) & xB & Chr$(34) & ")"

    DoEvents

    End If
    

Next I
Print #2, "Exit Function"

Print #2, "Error_Handler:"
Print #2, "srv" & mElpKMSrc_WHFILE_Y & "_GetBuffer_ODBC = Error"

Print #2, "End Function"

End Sub

Public Sub cmdElpKMPgm_Export_VB_GetBuffer_Rs()
Dim wPos As Integer, wLen As Integer
Dim recName As String
Dim xB As String

recName = "rs" & mElpKMSrc_WHFILE_Z & "."
Print #2, "Public Function rs" & mElpKMSrc_WHFILE_Z & "_GetBuffer(rsADO as ADODB.Recordset, rs" & mElpKMSrc_WHFILE_Z & " As type" & mElpKMSrc_WHFILE_Z & ")"

Print #2, "On Error GoTo Error_Handler"
Print #2, "rs" & mElpKMSrc_WHFILE_Z & "_GetBuffer = Null"

meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        xB = Trim(meDSPFFDY0.WHFLDE)
        Print #2, "rs" & mElpKMSrc_WHFILE_Z & "." & xB & " = rsADO(" & Chr$(34) & xB & Chr$(34) & ")"

    DoEvents

    End If
    

Next I
Print #2, "Exit Function"

Print #2, "Error_Handler:"
Print #2, "rs" & mElpKMSrc_WHFILE_Z & "_GetBuffer = Error"

Print #2, "End Function"

End Sub

Public Sub cmdElpKMPgm_Export_VB_PutBuffer_Rs()
Dim wPos As Integer, wLen As Integer
Dim recName As String
Dim xB As String

recName = "rs" & mElpKMSrc_WHFILE_Z & "."
Print #2, "Public Function rs" & mElpKMSrc_WHFILE_Z & "_PutBuffer(rsADO as ADODB.Recordset, rs" & mElpKMSrc_WHFILE_Z & " As type" & mElpKMSrc_WHFILE_Z & ")"

Print #2, "On Error GoTo Error_Handler"
Print #2, "mdb" & mElpKMSrc_WHFILE_Z & "_PutBuffer_Rs = Null"

meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        xB = Trim(meDSPFFDY0.WHFLDE)
        Print #2, "rsADO(" & Chr$(34) & xB & Chr$(34) & ") = " & "rs" & mElpKMSrc_WHFILE_Z & "." & xB
    DoEvents

    End If
    

Next I
Print #2, "Exit Function"

Print #2, "Error_Handler:"
Print #2, "rs" & mElpKMSrc_WHFILE_Z & "_PutBuffer = Error"

Print #2, "End Function"

End Sub


Public Sub cmdElpKMPgm_Export_JRN_GetBuffer_ODBC()
Dim wPos As Integer, wLen As Integer
Dim recName As String
Dim xB As String

recName = "rec" & mElpKMSrc_WHFILE_J & "."
Print #2, "Public Function srv" & mElpKMSrc_WHFILE_J & "_GetBuffer_ODBC(rsADO as ADODB.Recordset,l" & mElpKMSrc_WHFILE_J & " As type" & mElpKMSrc_WHFILE_J & ")"

Print #2, "On Error GoTo Error_Handler"
Print #2, "srv" & mElpKMSrc_WHFILE_J & "_GetBuffer_ODBC = Null"
Print #2, "l" & mElpKMSrc_WHFILE_J & ".JORCV = rsADO(" & Chr$(34) & "JORCV" & Chr$(34) & ")"
Print #2, "l" & mElpKMSrc_WHFILE_J & ".JOSEQN = rsADO(" & Chr$(34) & "JOSEQN" & Chr$(34) & ")"
Print #2, "l" & mElpKMSrc_WHFILE_J & ".JRNBIATRN = rsADO(" & Chr$(34) & "JRNBIATRN" & Chr$(34) & ")"

meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        xB = Trim(meDSPFFDY0.WHFLDE)
        Print #2, "l" & mElpKMSrc_WHFILE_J & "." & xB & " = rsADO(" & Chr$(34) & xB & Chr$(34) & ")"

    DoEvents

    End If
    

Next I
Print #2, "Exit Function"

Print #2, "Error_Handler:"
Print #2, "srv" & mElpKMSrc_WHFILE_J & "_GetBuffer_ODBC = Error"

Print #2, "End Function"

End Sub


Public Sub cmdElpKMPgm_Export_VB_frmElpDisplay()
Dim X As String, Xa As String, xB As String, Xc As String, Xd As String
Dim wPos As Integer, wLen As Integer

Print #2, "Public Sub srv" & mElpKMSrc_WHFILE_Y & "_ElpDisplay(rec" & mElpKMSrc_WHFILE_Y & " As type" & mElpKMSrc_WHFILE_Y & ")"
Print #2, "frmElpDisplay.fgData.Rows = " & lstElpKMPgm_W.ListCount + 1

wPos = 1
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0
        Print #2, "frmElpDisplay.fgData.Row = " & I + 1
        Print #2, "frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData =" & Chr$(34) & Trim(meDSPFFDY0.WHFLDE) & mId$(meElpKMInfo.Description, 30, 6) & Chr$(34)
        Print #2, "frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData =" & Chr$(34) & Trim(meDSPFFDY0.WHFTXT) & Chr$(34)
        Print #2, "frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = rec" & mElpKMSrc_WHFILE_Y & "." & meDSPFFDY0.WHFLDE

    End If
    
    DoEvents

Next I
Print #2, "frmElpDisplay.Show vbModal"
Print #2, "end sub"

End Sub

Public Sub cmdElpKMPgm_Export_JRN_fgDisplay()
Dim X As String, Xa As String, xB As String, Xc As String, Xd As String
Dim wPos As Integer, wLen As Integer

Print #2, "Public Sub srv" & mElpKMSrc_WHFILE_J & "_fgDisplay(l" & mElpKMSrc_WHFILE_J & " As type" & mElpKMSrc_WHFILE_J & ", fgDisplay As MSFlexGrid)"
Print #2, "fgDisplay.Rows = " & lstElpKMPgm_W.ListCount + 1

wPos = 1
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0
        Print #2, "fgdisplay.Row = " & I + 1
        Print #2, "fgdisplay.Col = 0: fgdisplay =" & Chr$(34) & Trim(meDSPFFDY0.WHFLDE) & mId$(meElpKMInfo.Description, 30, 6) & Chr$(34)
        Print #2, "fgdisplay.Col = 1: fgdisplay =" & Chr$(34) & Trim(meDSPFFDY0.WHFTXT) & Chr$(34)
        Print #2, "fgdisplay.Col = 2: fgdisplay = l" & mElpKMSrc_WHFILE_J & "." & meDSPFFDY0.WHFLDE

    End If
    
    DoEvents

Next I
Print #2, "end sub"

End Sub


Public Sub cmdElpKMPgm_Export_VB_CSV()
Dim X As String, Xa As String, xB As String, Xc As String, Xd As String
Dim wPos As Integer, wLen As Integer

Print #2, "Public Sub srv" & mElpKMSrc_WHFILE_Y & "_Export_CSV()"
Print #2, "Dim xIn as string"
Print #2, "Open " & Asc34 & "C:\Temp\" & mElpKMSrc_WHFILE_Y & ".txt" & Asc34 & " For input As #1"
Print #2, "Open " & Asc34 & "C:\Temp\" & mElpKMSrc_WHFILE_Y & ".csv" & Asc34 & " For Output As #2"
Xa = ""
xB = ""
Xc = ""
For I = 1 To arrExport_Nb
        Xa = Xa & arrExport_WHFLDE(I) & ";"
        xB = xB & arrExport_WHCHD1(I) & ";"
        Xc = Xc & arrExport_WHCHD2(I) & ";"
Next I
Print #2, "Print #2, " & Asc34 & Xa & Asc34
Print #2, "Print #2, " & Asc34 & xB & Asc34
Print #2, "Print #2, " & Asc34 & Xc & Asc34
Print #2, "Do Until EOF(1)"

Print #2, "      Line Input #1, xIn"

Xa = "      Print #2, "
For I = 1 To arrExport_Nb
        Xc = "mid$(xin, " & arrExport_Pos(I) & " , " & arrExport_Len(I) & ")" & " & " & Asc34 & ";" & Asc34
        Select Case I
            Case 1: Print #2, "      Print #2, " & Xc & " _"
            Case arrExport_Nb: Print #2, "      & " & Xc
            Case Else: Print #2, "      & " & Xc & " _"
        End Select
Next I
''Print #2, "      Print #2, " & Xc & " & " & Asc34 & ";" & Asc34

Print #2, "Loop"

Print #2, "Close "
Print #2, "end sub"

End Sub



Public Sub cmdElpKMPgm_Export_VB_PutBuffer()
Dim X As String, Xa As String, xB As String, Xc As String, Xd As String
Dim wPos As Integer, wLen As Integer

wPos = 1
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0
        If meDSPFFDY0.WHFLDT = "A" Then
            wLen = meDSPFFDY0.WHFLDB
        Else
            wLen = meDSPFFDY0.WHFLDD + 1
        End If

        Xa = Chr$(9) & "mid$(MsgTxt , k + " & wPos & " , " & wLen & ") = "
        Xc = "rec" & mElpKMSrc_WHFILE_Y & "." & Trim(meDSPFFDY0.WHFLDE)
        ''Mid$(Xc, 4, 1) = "X"
        Select Case meDSPFFDY0.WHFLDT
            Case "A":   xB = "": Xd = ""
            Case Else:
                        xB = "Format$(": Xd = ", " & Asc34 & String$(wLen - 1, "0") & " " & Asc34 & ")"
                        If meDSPFFDY0.WHFLDP > 0 Then Xc = Xc & " * 1" & String$(meDSPFFDY0.WHFLDP, "0")
        End Select
        
        X = Xa & xB & Xc & Xd
        Print #2, X
        wPos = wPos + wLen

    DoEvents

    End If
    

Next I

End Sub





Public Sub cmdElpKMPgm_Import_DSPFFDY0_Init(lWHLIB As String, lWHFILE As String)

lWHLIB = meDSPFFDY0.WHLIB
lWHFILE = meDSPFFDY0.WHFILE

recElpKMInfo_Init meElpKMInfo
meElpKMInfo.Method = "Seek="
meElpKMInfo.Pass = 1000
meElpKMInfo.ElpKMSrc_Id = 11000
meElpKMInfo.ID = meDSPFFDY0.WHLIB & meDSPFFDY0.WHFILE
If tableElpKMInfo_Read(meElpKMInfo) <> 0 Then
    meElpKMInfo.Method = constAddNew
    dbElpKMInfo_Update meElpKMInfo
End If
    
mElpKMSrc_Id = 12000
mElpKMPgm_Library = srvDSPFFDY0_Library(meDSPFFDY0.WHLIB)

meElpKMInfo.Method = constAddNew
meElpKMInfo.Pass = 1000
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id + mElpKMPgm_Library
'mdbElpKMInfo.tableElpKMInfo_Close
X = "delete * from ElpKMInfo where [ElpKMSrc_Id] = " & meElpKMInfo.ElpKMSrc_Id & " AND [Id] BETWEEN '" & meDSPFFDY0.WHFILE & "'  AND   '" & meDSPFFDY0.WHFILE & Chr$(255) & "'"
MDB.Execute X
'mdbElpKMInfo.tableElpKMInfo_Open

End Sub

Public Sub cmdElpKMPgm_Import_DSPFDY1_Init(lAPLIB As String, lAPFILE As String)
Dim wId As Long
    
mElpKMSrc_Id = 13000
mElpKMPgm_Library = srvDSPFFDY0_Library(lAPLIB)

meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id + mElpKMPgm_Library
X = "delete * from ElpKMInfo where [ElpKMSrc_Id] = " & meElpKMInfo.ElpKMSrc_Id & " AND [Id] BETWEEN '" & meDSPFDY1.APFILE & "'  AND   '" & meDSPFDY1.APFILE & Chr$(255) & "'"
MDB.Execute X

If Trim(meDSPFDY1.APBOF) <> "" Then
    pflfElpKMInfo = meElpKMInfo
    pflfElpKMInfo.ElpKMSrc_Id = 14000 + mElpKMPgm_Library
    pflfElpKMInfo.ID = meDSPFDY1.APBOF & meDSPFDY1.APFILE
    pflfElpKMInfo.Description = meDSPFDY1.APLIB & meDSPFDY1.APBOF & meDSPFDY1.APFILE
    xElpKMInfo = pflfElpKMInfo
    If tableElpKMInfo_Read(xElpKMInfo) = 0 Then
        pflfElpKMInfo.Method = constUpdate
    Else
        pflfElpKMInfo.Method = constAddNew
    End If

    dbElpKMInfo_Update pflfElpKMInfo
    
    pflfElpKMInfo.ElpKMSrc_Id = 15000 + mElpKMPgm_Library
    pflfElpKMInfo.ID = meDSPFDY1.APFILE
    xElpKMInfo = pflfElpKMInfo
    If tableElpKMInfo_Read(xElpKMInfo) = 0 Then
        pflfElpKMInfo.Method = constUpdate
    Else
        pflfElpKMInfo.Method = constAddNew
    End If

    dbElpKMInfo_Update pflfElpKMInfo

End If

End Sub

Private Sub mnuElpKMPgm_File_Afficher_Alpha_Click()
mnuElpKMPgm_File_Afficher_exe True, False 'blnalpha,blnPrint
End Sub

Private Sub mnuElpKMPgm_file_Afficher_Click()
mnuElpKMPgm_File_Afficher_exe False, False 'blnalpha,blnPrint
End Sub


Private Sub mnuElpKMPgm_file_Exporter_Click()
cmdElpKMPgm_Export
End Sub



Public Sub cmdElpKMPgm_Import_DSPFFDY0()
Dim mWHFILE As String, mWHLIB As String
Dim xIn As String, X As String, X5 As String * 5

Line Input #1, xIn
MsgTxt = Space$(34) & xIn
MsgTxtIndex = 0
srvDSPFFDY0_GetBuffer meDSPFFDY0

Call cmdElpKMPgm_Import_DSPFFDY0_Init(mWHLIB, mWHFILE)


Do Until EOF(1)
    Nb_Rec = Nb_Rec + 1
 '   If Nb_Rec Mod 100 = 0 Then Call lstErr_Clear(lstErr, cmdContext, "cmdSPLF_Click : " & Nb_Rec)
    

    meElpKMInfo.ID = meDSPFFDY0.WHFILE & meDSPFFDY0.WHFLDE
    If meDSPFFDY0.WHFLDT = "A" Then
        X = Format$(meDSPFFDY0.WHFLDB, "    0")
    Else
        If meDSPFFDY0.WHFLDP = 0 Then
            X = Format$(meDSPFFDY0.WHFLDD, "    0")
        Else
            X = Format$(meDSPFFDY0.WHFLDD, "  0") & "." & meDSPFFDY0.WHFLDP
        End If
        
    End If
    
    meElpKMInfo.Description = meDSPFFDY0.WHCHD1
    Mid$(meElpKMInfo.Description, 20, 10) = meDSPFFDY0.WHFLDE
    Call strMoveR(meDSPFFDY0.WHFOBO, meElpKMInfo.Description, 36, 5)
    Mid$(meElpKMInfo.Description, 35, 1) = meDSPFFDY0.WHFLDT
    Call strMoveR(X, meElpKMInfo.Description, 30, 5)
   
   ' meElpKMInfo.description = meDSPFFDY0.WHFLDE
   ' Mid$(meElpKMInfo.description, 18, 23) = meDSPFFDY0.WHCHD1
   ' Mid$(meElpKMInfo.description, 16, 1) = meDSPFFDY0.WHFLDT
   ' Call strMoveR(X, meElpKMInfo.description, 11, 5)
 
    
    meElpKMInfo.Memo = Trim(xIn)
    
    If meDSPFFDY0.WHFTYP = "P" Then dbElpKMInfo_Update meElpKMInfo

    Line Input #1, xIn
    MsgTxt = Space$(34) & xIn
    MsgTxtIndex = 0
    srvDSPFFDY0_GetBuffer meDSPFFDY0
    If meDSPFFDY0.WHLIB <> mWHLIB Or meDSPFFDY0.WHFILE <> mWHFILE Then Call cmdElpKMPgm_Import_DSPFFDY0_Init(mWHLIB, mWHFILE)

    DoEvents

Loop
lstElpKMPgm_11000_Load

End Sub
Public Sub cmdElpKMPgm_Import_DSPFDY0()
Dim xIn As String, X As String, X5 As String * 5

recElpKMInfo_Init meElpKMInfo
meElpKMInfo.Method = "Seek="
meElpKMInfo.Pass = 1000
meElpKMInfo.ElpKMSrc_Id = 11000

xElpKMInfo = meElpKMInfo


Do Until EOF(1)
    Line Input #1, xIn
    MsgTxt = Space$(34) & xIn
    MsgTxtIndex = 0
    srvDSPFDY0_GetBuffer meDSPFDY0
    Nb_Rec = Nb_Rec + 1
    If Nb_Rec Mod 100 = 0 Then Call lstErr_Clear(lstErr, cmdContext, "cmdSPLF_Click : " & Nb_Rec)
    

    meElpKMInfo.ID = meDSPFDY0.ATLIB & meDSPFDY0.ATFILE
    
    meElpKMInfo.Description = meDSPFDY0.ATFILE & " " & meDSPFDY0.ATTXT
    meElpKMInfo.Memo = Trim(xIn)
    xElpKMInfo.ID = meElpKMInfo.ID
    If tableElpKMInfo_Read(xElpKMInfo) = 0 Then
        meElpKMInfo.Method = constUpdate
    Else
        meElpKMInfo.Method = constAddNew
    End If
    
    
    dbElpKMInfo_Update meElpKMInfo


    DoEvents

Loop

End Sub

Public Sub cmdDSPFDY2_Import_Src(lElpKMSrc_Id As Long)
Dim xIn As String, X As String, X5 As String * 5

recElpKMInfo_Init meElpKMInfo
meElpKMInfo.Method = constAddNew
meElpKMInfo.Pass = 1000
meElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id

xElpKMInfo = meElpKMInfo


Do Until EOF(1)
    Line Input #1, xIn
  '  MsgTxt = Space$(34) & xIn
  '  MsgTxtIndex = 0
  '  srvDSPFDY0_GetBuffer meDSPFDY0
    Nb_Rec = Nb_Rec + 1
    If Nb_Rec Mod 100 = 0 Then Call lstErr_Clear(lstErr, cmdContext, "cmdDSPFDY2_Import_Src : " & Nb_Rec)
    
    If mId$(xIn, 42, 6) = "PF    " Then
        If Val(mId$(xIn, 387, 11)) > 0 Then
            meElpKMInfo.ID = mId$(xIn, 14, 10)
             
             meElpKMInfo.Description = mId$(xIn, 79, 40)
             meElpKMInfo.Memo = Trim(xIn)
             xElpKMInfo.ID = meElpKMInfo.ID
    
             dbElpKMInfo_Update meElpKMInfo
        End If
    End If
    
    DoEvents
    
Loop

End Sub



Public Sub cmdElpKMPgm_Import_DSPFDY1()
Dim xIn As String, X As String
Dim mAPLIB As String, mAPFILE As String

mAPLIB = "": mAPFILE = ""
recElpKMInfo_Init meElpKMInfo
meElpKMInfo.Method = "Seek="
meElpKMInfo.Pass = 1000
meElpKMInfo.ElpKMSrc_Id = 11000

xElpKMInfo = meElpKMInfo


Do Until EOF(1)
    Line Input #1, xIn
    MsgTxt = Space$(34) & xIn
    MsgTxtIndex = 0
    srvDSPFDY1_GetBuffer meDSPFDY1
    Nb_Rec = Nb_Rec + 1
    If Nb_Rec Mod 100 = 0 Then Call lstErr_Clear(lstErr, cmdContext, "cmdSPLF_Click : " & Nb_Rec)
    
    meElpKMInfo.Method = constAddNew
    meElpKMInfo.ID = meDSPFDY1.APFILE & meDSPFDY1.APKEYN
    meElpKMInfo.Description = meDSPFDY1.APKEYF
    meElpKMInfo.Memo = Trim(xIn)
    
    If mAPLIB <> meDSPFDY1.APLIB Or mAPFILE <> meDSPFDY1.APFILE Then
        mAPLIB = meDSPFDY1.APLIB
        mAPFILE = meDSPFDY1.APFILE
        Call cmdElpKMPgm_Import_DSPFDY1_Init(mAPLIB, mAPFILE)
    End If
       
    
    
    dbElpKMInfo_Update meElpKMInfo
    
    DoEvents

Loop

End Sub


Private Sub mnuElpKMPgm_File_Imprimer_Alpha_Click()
mnuElpKMPgm_File_Afficher_exe True, True 'blnalpha,blnPrint

End Sub

Private Sub mnuElpKMPgm_File_Imprimer_Click()
mnuElpKMPgm_File_Afficher_exe False, True 'blnalpha,blnPrint

End Sub

Private Sub mnuPrint2_SelectFalse_Click()
lstSABPF.Visible = False
MsgBox "2005.02.11 à revoir"
'Call prtElpKm_CLP(lstSABPF, False, Trim(txtSABPF_CLP))
lstSABPF.Visible = True
End Sub

Private Sub mnuPrint2_SelectTrue_Click()
lstSABPF.Visible = False
MsgBox "2005.02.11 à revoir"
'Call prtElpKm_CLP(lstSABPF, True, Trim(txtSABPF_CLP))
lstSABPF.Visible = True
End Sub

Private Sub txtDSPFDY2_Select_GotFocus()
txt_GotFocus txtDSPFDY2_Select

End Sub


Private Sub txtDSPFDY2_Select_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtDSPFDY2_Select_LostFocus()
txt_LostFocus txtDSPFDY2_Select

End Sub


Private Sub txtDSPOBJD_Select_GotFocus()
txt_GotFocus txtDSPOBJD_Select

End Sub


Private Sub txtDSPOBJD_Select_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtDSPOBJD_Select_LostFocus()
txt_LostFocus txtDSPOBJD_Select

End Sub


Private Sub txtDSPOBJD_Src1_GotFocus()
txt_GotFocus txtDSPOBJD_Src1

End Sub


Private Sub txtDSPOBJD_Src1_LostFocus()
txt_LostFocus txtDSPOBJD_Src1

End Sub


Private Sub txtDSPOBJD_Src2_GotFocus()
txt_GotFocus txtDSPOBJD_Src2

End Sub


Private Sub txtDSPOBJD_Src2_LostFocus()
txt_LostFocus txtDSPOBJD_Src2

End Sub


Private Sub txtElpKMPgm_Export_GotFocus()
txt_GotFocus txtElpKMPgm_Export

End Sub


Private Sub txtElpKMPgm_Export_LostFocus()
txt_LostFocus txtElpKMPgm_Export

End Sub


Private Sub txtElpKMPgm_Import_GotFocus()
txt_GotFocus txtElpKMPgm_Import

End Sub


Private Sub txtElpKMPgm_Import_LostFocus()
txt_LostFocus txtElpKMPgm_Import

End Sub


Private Sub txtElpKMPgm_Select_GotFocus()
txt_GotFocus txtElpKMPgm_Select

End Sub

Private Sub txtElpKMPgm_Select_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub



Public Sub mnuElpKMPgm_File_Afficher_exe(blnAlpha As Boolean, blnPrint As Boolean)
Me.Enabled = False: Me.MousePointer = vbHourglass
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Display: " & Time)
DoEvents

mElpKMPgm_Library = srvDSPFFDY0_Library(mElpKMSrc_WHLIB)
mElpKMSrc_Id = 12000 + mElpKMPgm_Library
''meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

Call srvDSPFFDY0_lstAddItem(lstElpKMPgm_W, mElpKMSrc_Id, mElpKMSrc_WHFILE, blnAlpha)

Call srvDSPFFDY0_frmRTF(lstElpKMPgm_W, mElpKMSrc_Id, mElpKMSrc_WHLIB, mElpKMSrc_WHFILE, blnPrint)

Call lstErr_AddItem(lstErr, cmdContext, "cmdElpKMPgm_Display :  Fin")
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub txtElpKMPgm_Select_LostFocus()
txt_LostFocus txtElpKMPgm_Select


End Sub


Private Sub txtSABPF_Select_GotFocus()
txt_GotFocus txtSABPF_Select


End Sub


Private Sub txtSABPF_Select_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSABPF_Select_LostFocus()
txt_LostFocus txtSABPF_Select

End Sub



Public Sub arrSABPF_Select()
Dim I As Integer, blnSelect As Boolean

blnSelect = lstSABPF.Selected(lstSABPF.ListIndex)
I = lstSABPF.ListIndex + 1
If blnSelect Then
    If mId$(arrSABPF(I).Memo, paramSABPF_K, 3) <> paramSABPF_opt Then
        Mid$(arrSABPF(I).Memo, paramSABPF_K, 3) = paramSABPF_opt
        arrSABPF(I).Method = constUpdate
    End If
Else
    If mId$(arrSABPF(I).Memo, paramSABPF_K, 3) = paramSABPF_opt Then
        Mid$(arrSABPF(I).Memo, paramSABPF_K, 3) = "xxx"
        arrSABPF(I).Method = constUpdate
    End If

End If
If mId$(arrSABPF(I).Memo, paramSABPF_K, 3) = paramSABPF_opt Then
    lstSABPF.Selected(lstSABPF.ListIndex) = True
Else
    lstSABPF.Selected(lstSABPF.ListIndex) = False
End If

End Sub

Public Sub cmdSABPF_CLP_CLR()
For I = 0 To lstSABPF.ListCount - 1
    lstSABPF.ListIndex = I
    If lstSABPF.Selected(I) Then
        X = lstSABPF
        X80 = Space(13) & "CLRPFM     FILE(&TOLIB/" & mId$(X, 1, 10) & ")"
        Print #2, X80
    End If
    DoEvents

Next I

End Sub
Public Sub cmdSABPF_CLP_CPY()

X80 = Space(13) & "MONMSG     MSGID(CPF2817) /* fichier vide*/"
Print #2, X80

For I = 0 To lstSABPF.ListCount - 1
    lstSABPF.ListIndex = I
    If lstSABPF.Selected(I) Then
        X = lstSABPF
        X80 = Space(13) & "CPYF       FROMFILE(&FROMLIB/" & mId$(X, 1, 10) & ") +"
        Print #2, X80
        X80 = Space(13) & "               TOFILE(&TOLIB/" & mId$(X, 1, 10) & ") MBROPT(*REPLACE)"
        Print #2, X80
   End If
    DoEvents

Next I

End Sub

Public Sub cmdSABPF_CLP_PAR()

X80 = Space(13) & "MONMSG     MSGID(CPF2817) /* fichier vide*/"
Print #2, X80

For I = 0 To lstSABPF.ListCount - 1
    lstSABPF.ListIndex = I
    If lstSABPF.Selected(I) Then
        X = lstSABPF
        X80 = Space(13) & "CPYF       FROMFILE(&FROMLIB/" & mId$(X, 1, 10) & ") +"
        Print #2, X80
        X80 = Space(13) & "               TOFILE(&TOLIB/" & mId$(X, 1, 10) & ") MBROPT(*REPLACE)"
        Print #2, X80
   End If
    DoEvents
Next I

End Sub

Public Sub cmdSABPF_CLP_QRY()

For I = 0 To lstSABPF.ListCount - 1
    lstSABPF.ListIndex = I
    If lstSABPF.Selected(I) Then
        X = lstSABPF
        X80 = Space(13) & "RUNQRY     QRY(*NONE) QRYFILE((&TOLIB/" & mId$(X, 1, 10) & ")) +"
        Print #2, X80
        X80 = Space(13) & "               OUTTYPE(*PRINTER)"
        Print #2, X80
   End If
    DoEvents

Next I


End Sub

Public Sub cmdDSPFDY2_CLP_QRY()

For I = 0 To lstDSPFDY2.ListCount - 1
    lstDSPFDY2.ListIndex = I
    If lstDSPFDY2.Selected(I) Then
        X = lstDSPFDY2
        X80 = Space(13) & "RUNQRY     QRY(*NONE) QRYFILE((&TOLIB/" & mId$(X, 1, 10) & ")) +"
        Print #2, X80
        X80 = Space(13) & "               OUTTYPE(*PRINTER)"
        Print #2, X80
   End If
    DoEvents

Next I


End Sub


Public Sub cmdElpKm_mdb_Info_Export()
Dim X As String, K As Integer, Nb As Long
On Error GoTo Error_Handler

Dim txtImportExport As String
txtImportExport = paramTemp_Folder & "FTP\ElmKMInfo_Export.txt"

Open Trim(txtImportExport) For Output As #2

Call lstErr_Clear(lstErr, cmdPrint, "Export " & Time): DoEvents

Nb = 0
mdbElpKMInfo.tableElpKMInfo_Open
recElpKMInfo_Init xElpKMInfo
xElpKMInfo.Method = "Seek>="
xElpKMInfo.ElpKMSrc_Id = 10000

Do
    intReturn = tableElpKMInfo_Read(xElpKMInfo)

    If intReturn = 0 Then
    
        If xElpKMInfo.ElpKMSrc_Id >= 30000 Then Exit Do
        
        If IsNull(xElpKMInfo.Memo) Then
            K = 0
        Else
            K = Len(xElpKMInfo.Memo)
        End If
        X = Space$(79 + K)
        Mid$(X, 1, 9) = Format(xElpKMInfo.ElpKMSrc_Id, "########0")
        Mid$(X, 10, 20) = xElpKMInfo.ID
        Mid$(X, 30, 40) = xElpKMInfo.Description
        Mid$(X, 70, 9) = Format(xElpKMInfo.Pass, "########0")
        If K > 0 Then Mid$(X, 79, K) = xElpKMInfo.Memo
        Nb = Nb + 1
        Print #2, Trim(X)
   End If
    xElpKMInfo.Method = "MoveNext"
Loop While intReturn = 0


Close
Call lstErr_AddItem(lstErr, cmdPrint, "fin : " & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "Export")
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdElpKm_mdb_Index_Export()
Dim X As String, K As Integer, Nb As Long
On Error GoTo Error_Handler

Dim txtImportExport As String
txtImportExport = "C:\Temp\ElmKMIndex_Export.txt"

Open Trim(txtImportExport) For Output As #2

Call lstErr_Clear(lstErr, cmdPrint, "Export " & Time): DoEvents

Nb = 0
mdbElpKMIndex.tableElpKMIndex_Open
recElpKMIndex_Init xElpKMIndex
xElpKMIndex.Method = "MoveFirst"

Do
    intReturn = tableElpKMIndex_Read(xElpKMIndex)

    If intReturn = 0 Then
        If xElpKMIndex.Classe >= 10000 And xElpKMIndex.Classe < 30000 Then
        
            If IsNull(xElpKMIndex.Memo) Then
                K = 0
            Else
                K = Len(xElpKMIndex.Memo)
            End If
            X = Space$(35 + K)
            Mid$(X, 1, 16) = xElpKMIndex.ID
            Mid$(X, 17, 9) = Format(xElpKMIndex.Classe, "########0")
            Mid$(X, 26, 9) = Format(xElpKMIndex.ElpKMSrc_Id, "########0")
            If K > 0 Then Mid$(X, 35, K) = xElpKMIndex.Memo
            Nb = Nb + 1
            Print #2, Trim(X)
        End If
   End If
    xElpKMIndex.Method = "MoveNext"
Loop While intReturn = 0


Close
Call lstErr_AddItem(lstErr, cmdPrint, "fin : " & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "Export")
Me.Enabled = True: Me.MousePointer = 0

End Sub


Public Sub cmdElpKm_mdb_Info_Import()
Dim X As String, K As Integer, Nb As Long
Dim xIn As String
On Error GoTo Error_Handler

Dim txtImportExport As String
txtImportExport = paramTemp_Folder & "FTP\ElmKMInfo_Export.txt"

Open Trim(txtImportExport) For Input As #1

Call lstErr_Clear(lstErr, cmdPrint, "Import " & Time): DoEvents
Call lstErr_AddItem(lstErr, cmdPrint, "  ")

Nb = 0
mdbElpKMInfo.tableElpKMInfo_Open
recElpKMInfo_Init xElpKMInfo
xElpKMInfo.Method = "AddNew"
xElpKMInfo.ElpKMSrc_Id = 10000

Do Until EOF(1)
    Line Input #1, xIn
    Nb = Nb + 1
    If Nb Mod 100 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdPrint, mId$(xIn, 1, 29))

    xElpKMInfo.ElpKMSrc_Id = CLng(Val(mId$(xIn, 1, 9)))
    xElpKMInfo.ID = mId$(xIn, 10, 20)
    xElpKMInfo.Description = mId$(xIn, 30, 40)
    xElpKMInfo.Pass = CLng(Val(mId$(xIn, 70, 9)))
    K = Len(xIn)
    If K > 79 Then
        xElpKMInfo.Memo = mId$(xIn, 79, K - 78)
    Else
        xElpKMInfo.Memo = ""
    End If
    dbElpKMInfo_Update xElpKMInfo


    DoEvents

Loop

Close
Call lstErr_AddItem(lstErr, cmdPrint, "fin : " & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "Export")
Me.Enabled = True: Me.MousePointer = 0

End Sub


Public Sub cmdElpKm_mdb_Index_Import()
Dim X As String, K As Integer, Nb As Long
Dim xIn As String
On Error GoTo Error_Handler

Dim txtImportExport As String
txtImportExport = paramTemp_Folder & "\FTP\ElmKMIndex_Export.txt"

Open Trim(txtImportExport) For Input As #1

Call lstErr_Clear(lstErr, cmdPrint, "Import " & Time): DoEvents
Call lstErr_AddItem(lstErr, cmdPrint, "  ")

Nb = 0
mdbElpKMIndex.tableElpKMIndex_Open
recElpKMIndex_Init xElpKMIndex
xElpKMIndex.Method = "AddNew"
xElpKMIndex.ElpKMSrc_Id = 10000

Do Until EOF(1)
    Line Input #1, xIn
    Nb = Nb + 1
    If Nb Mod 100 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdPrint, mId$(xIn, 1, 29))

    xElpKMIndex.ID = mId$(xIn, 1, 16)
    xElpKMIndex.Classe = CLng(Val(mId$(xIn, 17, 9)))
    xElpKMIndex.ElpKMSrc_Id = CLng(Val(mId$(xIn, 26, 9)))
    K = Len(xIn)
    If K > 35 Then
        xElpKMIndex.Memo = mId$(xIn, 35, K - 34)
    Else
        xElpKMIndex.Memo = ""
    End If
    dbElpKMIndex_Update xElpKMIndex


    DoEvents

Loop

Close
Call lstErr_AddItem(lstErr, cmdPrint, "fin : " & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "Export")
Me.Enabled = True: Me.MousePointer = 0

End Sub



Public Sub cmdElpKM_mdb_Info_Clear()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdElpKM_mdb_Clear  : " & Time)
mdbElpKMInfo.tableElpKMInfo_Open
X = "delete * from ElpKMinfo where ElpKMSrc_Id >= 10000 and ElpKMSrc_Id <= 29999"
MDB.Execute X
Call lstErr_AddItem(lstErr, cmdContext, "cmdElpKM_mdb_Clear : terminé")

Me.Enabled = True: Me.MousePointer = 0

End Sub
Public Sub cmdElpKM_mdb_Index_Clear()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdElpKM_mdb_Index_Clear  : " & Time)
mdbElpKMIndex.tableElpKMIndex_Open
X = "delete * from ElpKMIndex where Classe >= 10000 and Classe <= 29999"
MDB.Execute X
Call lstErr_AddItem(lstErr, cmdContext, "cmdElpKM_mdb_Index_Clear : terminé")

Me.Enabled = True: Me.MousePointer = 0

End Sub


Public Sub cmdElpKMPgm_Export_VB_Déclaration()
meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        X80 = ""
        Mid$(X80, 1, 1) = Chr$(9)
        Mid$(X80, 2, 10) = meDSPFFDY0.WHFLDE
        Mid$(X80, 12, 2) = Chr$(9) & Chr$(9)
       
        Select Case meDSPFFDY0.WHFLDT
            Case "A": Mid$(X80, 14, 20) = "As String * " & meDSPFFDY0.WHFLDB
            Case "B": Mid$(X80, 14, 20) = "As Integer"
            Case Else:
                Select Case meDSPFFDY0.WHFLDP
                    Case 0: Mid$(X80, 14, 20) = "As Long"
                    Case 2: Mid$(X80, 14, 20) = "As Currency"
                    Case Else: Mid$(X80, 14, 20) = "As Double"
                End Select
      End Select
        Mid$(X80, 48, 52) = "' " & meDSPFFDY0.WHFTXT
        Print #2, X80
    End If
    
    DoEvents

Next I

Print #2, ""
Print #2, "End Type"

End Sub
Public Sub cmdElpKMPgm_File_TableDef_Create()
Dim xIn As String
Dim blnCreate As Boolean
Dim K As Integer

On Error Resume Next 'GoTo Error_Handler
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdElpKMPgm_Display: " & Time)
DoEvents

'>TableDef
Set mdbDataBase = OpenDatabase(DataBase_Data, False, False, paramDataBase_Password)
blnCreate = False
Open Trim(txtElpKMPgm_Export) For Input As #1

Do Until EOF(1)
    Line Input #1, xIn
    
    If mId$(xIn, 1, 10) = "'>TableDef" Then blnCreate = True
    
    If mId$(xIn, 1, 10) = "'<TableDef" Then
        If blnCreate Then
            ''mdbDataBase.TableDef.Delete mdbTable
            mdbDataBase.TableDefs.Append mdbTable
        End If
    End If
    
    K = 2
    Select Case mId$(xIn, 1, 1)
        Case "T": Set mdbTable = New TableDef
                  mdbTable.Name = CSV_Scan(xIn, K)

        Case "F": Set mdbField = New Field
                  mdbField.Name = Trim(CSV_Scan(xIn, K))
                  mdbField.Type = CInt(CSV_Scan(xIn, K))
                  mdbField.Size = CInt(CSV_Scan(xIn, K))
                  mdbTable.Fields.Append mdbField

        Case "I": Set mdbIndex = New Index
                  mdbIndex.Name = CSV_Scan(xIn, K)
                  mdbIndex.Fields = mId$(xIn, K + 1, Len(xIn) - K + 1)
                  mdbIndex.Primary = True
                  mdbIndex.Unique = True
                  mdbTable.Indexes.Append mdbIndex
    End Select
    
Loop



GoTo Exit_Sub

Error_Handler:
MsgBox Error, vbCritical, Me.Caption & " : " & txtElpKMPgm_Export
Exit_Sub:
Close

mdbDataBase.Close

End Sub

Public Sub cmdElpKMPgm_Export_TableDef()
Dim wName As String, wType As String, wSize As Integer

Print #2, "'>TableDef-------------------------------------------------------"
Print #2, "T;" & Trim(mElpKMSrc_WHFILE)

meElpKMInfo.Method = "Seek="
meElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id
'------------------------------------------------------------------------------------------

For I = 0 To lstElpKMPgm_W.ListCount - 1
    lstElpKMPgm_W.ListIndex = I
    X = lstElpKMPgm_W
    meElpKMInfo.ID = mId$(X, 6, Len(X) - 5)
    intReturn = tableElpKMInfo_Read(meElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer meDSPFFDY0

        wName = meDSPFFDY0.WHFLDE
        wSize = 0
        Select Case meDSPFFDY0.WHFLDT
            Case "A": wType = dbText
                      wSize = meDSPFFDY0.WHFLDB
            Case "B": wType = dbInteger
            Case Else:
                Select Case meDSPFFDY0.WHFLDP
                    Case 0: wType = dbLong
                    Case 2: wType = dbCurrency
                    Case Else: wType = dbDouble
                End Select
      End Select
    End If
    Print #2, "F;" & wName & ";" & wType & ";"; wSize

    DoEvents

Next I


'-----------------------------------------------
'Dim mElpKMSrc_Id As Long
Dim wElpKMInfo As typeElpKMInfo ', X As String
Dim wDSPFDY1 As typeDSPFDY1
Dim wField As typeElpKMInfo
Dim wDSPFFDY0 As typeDSPFFDY0
Dim xDes As String

On Error Resume Next

wField.Method = "Seek="
wField.ElpKMSrc_Id = mElpKMSrc_Id


wElpKMInfo.Method = "Seek>="
mElpKMSrc_Id = mElpKMSrc_Id + 1000
wElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id '13***
wElpKMInfo.ID = mElpKMSrc_WHFILE
intReturn = tableElpKMInfo_Read(wElpKMInfo)
wElpKMInfo.Method = "Seek>"
wName = ""
Do
    If intReturn = 0 Then
        MsgTxt = Space$(34) & wElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFDY1_GetBuffer wDSPFDY1
        If wElpKMInfo.ElpKMSrc_Id <> mElpKMSrc_Id Or wDSPFDY1.APFILE <> mElpKMSrc_WHFILE Then
            intReturn = -1
        Else
            wName = wName & Trim(wDSPFDY1.APKEYF) & ";"
            intReturn = tableElpKMInfo_Read(wElpKMInfo)
    
        End If
    End If
      
Loop While intReturn = 0

Print #2, "I;" & "PrimaryKey" & ";" & wName
   
Print #2, "'<TableDef-------------------------------------------------------"


End Sub


