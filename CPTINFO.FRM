VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCptInfo 
   Caption         =   "COMPTE : Fiche"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9600
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   9600
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   9240
      Picture         =   "CPTINFO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdRacine 
      Caption         =   "&Racine"
      Height          =   300
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6030
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10636
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   16711680
      TabCaption(0)   =   "Compte :"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Conditions :"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraExtrait"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FraEchelle"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Solde :"
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraSoldeInstantané"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraSoldeComptable"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Adresse :"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Divers"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   147
         Top             =   480
         Width           =   9000
         Begin VB.TextBox txtCourrier 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   152
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtGestionnaire 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   151
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtApporteur 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   150
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtServiceResponsable 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   149
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtGroupeIdentification 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   148
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblCourrier 
            Caption         =   "Courrier"
            Height          =   225
            Left            =   50
            TabIndex        =   162
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblGestionnaire 
            Caption         =   "Gestionnaire"
            Height          =   255
            Left            =   50
            TabIndex        =   161
            Top             =   720
            Width           =   2130
         End
         Begin VB.Label lblApporteur 
            Caption         =   "Apporteur"
            Height          =   255
            Left            =   50
            TabIndex        =   160
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblServiceResponsable 
            Caption         =   "Service responsable "
            Height          =   225
            Left            =   50
            TabIndex        =   159
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblGroupeIdentification 
            Caption         =   "Compte de Regroupement"
            Height          =   255
            Left            =   50
            TabIndex        =   158
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label libCourrier 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   157
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label libGestionnaire 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   156
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label libApporteur 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   155
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label libServiceResponsable 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   154
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label libGroupeIdentification 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   153
            Top             =   1680
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5450
         Left            =   -74760
         TabIndex        =   105
         Top             =   480
         Width           =   9000
         Begin VB.TextBox txtAlpha 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6600
            TabIndex        =   122
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txtSituation 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   121
            Top             =   4920
            Width           =   255
         End
         Begin VB.TextBox txtNostro 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   120
            Top             =   4200
            Width           =   495
         End
         Begin VB.TextBox txtTypeGA 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   119
            Top             =   3120
            Width           =   255
         End
         Begin VB.TextBox txtAgence 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   118
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtDevise 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   117
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtNuméro 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   116
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtIntitulé 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   115
            Top             =   1320
            Width           =   4935
         End
         Begin VB.TextBox txtIntitulé2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   114
            Top             =   1680
            Width           =   4935
         End
         Begin VB.TextBox txtRésident 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   113
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txtActionnaire 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   112
            Top             =   2400
            Width           =   255
         End
         Begin VB.TextBox txtNatureTitulaire 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   111
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox txtCompteGénéral 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   110
            Top             =   3480
            Width           =   1095
         End
         Begin VB.TextBox txtSens 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   109
            Top             =   3840
            Width           =   495
         End
         Begin VB.TextBox txtNature 
            Height          =   285
            Left            =   1680
            TabIndex        =   108
            Top             =   4560
            Width           =   375
         End
         Begin VB.TextBox txtNuméroAncien 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6600
            TabIndex        =   107
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox txtClérib 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6630
            TabIndex        =   106
            Top             =   3315
            Width           =   375
         End
         Begin VB.Label lblCléRib 
            Caption         =   "Clé rib "
            Height          =   225
            Left            =   5265
            TabIndex        =   146
            Top             =   3375
            Width           =   855
         End
         Begin VB.Label libNature 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   145
            Top             =   4560
            Width           =   2535
         End
         Begin VB.Label libNostro 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   144
            Top             =   4200
            Width           =   2535
         End
         Begin VB.Label libActionnaire 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   143
            Top             =   2400
            Width           =   2535
         End
         Begin VB.Label lblAgence 
            Caption         =   "Code Agence"
            Height          =   225
            Left            =   50
            TabIndex        =   142
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label lblDevise 
            Caption         =   "Code Devise"
            Height          =   225
            Left            =   50
            TabIndex        =   141
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label lblNuméro 
            Caption         =   "Numéro de compte "
            Height          =   225
            Left            =   50
            TabIndex        =   140
            Top             =   1080
            Width           =   1530
         End
         Begin VB.Label lblIntitulé 
            Caption         =   "Intitulé du compte"
            Height          =   255
            Left            =   45
            TabIndex        =   139
            Top             =   1440
            Width           =   1530
         End
         Begin VB.Label lblRésident 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Résident BDF"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   50
            TabIndex        =   138
            Top             =   2160
            Width           =   1410
         End
         Begin VB.Label libRésident 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   137
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label lblActionnaire 
            Caption         =   "Actionnaire "
            Height          =   225
            Left            =   50
            TabIndex        =   136
            Top             =   2520
            Width           =   1290
         End
         Begin VB.Label lblNatureTitulaire 
            Caption         =   "Nature du Titulaire"
            Height          =   255
            Left            =   50
            TabIndex        =   135
            Top             =   2880
            Width           =   1530
         End
         Begin VB.Label lblTypeOuverture 
            Caption         =   "Type de Compte"
            Height          =   255
            Left            =   50
            TabIndex        =   134
            Top             =   3240
            Width           =   1410
         End
         Begin VB.Label lblCompteGénéral 
            Caption         =   "Compte général"
            Height          =   225
            Left            =   50
            TabIndex        =   133
            Top             =   3600
            Width           =   1530
         End
         Begin VB.Label lblSens 
            Caption         =   "Sens"
            Height          =   255
            Left            =   50
            TabIndex        =   132
            Top             =   3960
            Width           =   1170
         End
         Begin VB.Label lblNostro 
            Caption         =   "Nostro/Loro"
            Height          =   225
            Left            =   50
            TabIndex        =   131
            Top             =   4320
            Width           =   1410
         End
         Begin VB.Label lblNature 
            Caption         =   "Nature de Compte"
            Height          =   255
            Left            =   50
            TabIndex        =   130
            Top             =   4680
            Width           =   1530
         End
         Begin VB.Label lblSituation 
            Caption         =   "Situation  "
            Height          =   225
            Left            =   50
            TabIndex        =   129
            Top             =   5040
            Width           =   1530
         End
         Begin VB.Label lblAlpha 
            Caption         =   "Recherche Alpha"
            Height          =   255
            Left            =   5160
            TabIndex        =   128
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblNuméroAncien 
            Caption         =   "Ancien compte"
            Height          =   225
            Left            =   5160
            TabIndex        =   127
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label libNatureTitulaire 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   126
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label libSens 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   125
            Top             =   3840
            Width           =   2535
         End
         Begin VB.Label libSituation 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   124
            Top             =   4920
            Width           =   2535
         End
         Begin VB.Label libTypeGA 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   123
            Top             =   3120
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   88
         Top             =   1200
         Width           =   9000
         Begin VB.TextBox txtAdressePays 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   104
            Top             =   2400
            Width           =   4335
         End
         Begin VB.TextBox txtAdresseBd 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3120
            TabIndex        =   102
            Top             =   2040
            Width           =   3375
         End
         Begin VB.TextBox txtAdresseCp 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   101
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox txtAdresse5 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   98
            Top             =   1680
            Width           =   4335
         End
         Begin VB.TextBox txtAdresse4 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   96
            Top             =   1320
            Width           =   4335
         End
         Begin VB.TextBox txtAdresse3 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   94
            Top             =   960
            Width           =   4335
         End
         Begin VB.TextBox txtAdresse2 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   92
            Top             =   600
            Width           =   4335
         End
         Begin VB.TextBox txtAdresse1 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   90
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label lblAdressePays 
            Caption         =   "Pays"
            Height          =   255
            Left            =   50
            TabIndex        =   103
            Top             =   2400
            Width           =   1530
         End
         Begin VB.Label lblAdresseBd 
            Caption         =   "Ville"
            Height          =   255
            Left            =   1200
            TabIndex        =   100
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label lblAdresseCp 
            Caption         =   "Code Postal"
            Height          =   255
            Left            =   50
            TabIndex        =   99
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label lblAdresse5 
            Caption         =   "Commune"
            Height          =   255
            Left            =   50
            TabIndex        =   97
            Top             =   1680
            Width           =   1530
         End
         Begin VB.Label lblAdresse4 
            Caption         =   "N° de rue"
            Height          =   225
            Left            =   50
            TabIndex        =   95
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblAdresse3 
            Caption         =   "Batiment"
            Height          =   225
            Left            =   50
            TabIndex        =   93
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblAdresse2 
            Caption         =   "Complt. Adresse"
            Height          =   225
            Left            =   50
            TabIndex        =   91
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblAdresse1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Intitulé"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   50
            TabIndex        =   89
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Height          =   4395
         Left            =   -74760
         TabIndex        =   61
         Top             =   600
         Width           =   9000
         Begin VB.TextBox txtServiceAutorisé5 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   71
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtServiceAutorisé4 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   70
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtServiceAutorisé3 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   69
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtServiceAutorisé2 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   68
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtServiceAutorisé1 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1680
            TabIndex        =   67
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtAmjCréation 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2280
            TabIndex        =   66
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtAmjModification 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2280
            TabIndex        =   65
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtAmjAnnulation 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2280
            TabIndex        =   64
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox txtAmjRéactivation 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2280
            TabIndex        =   63
            Top             =   3240
            Width           =   1095
         End
         Begin VB.TextBox txtNomOpérateur 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2280
            TabIndex        =   62
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label lblAmjAnnulation 
            Caption         =   "Date Annulation"
            Height          =   255
            Left            =   50
            TabIndex        =   86
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label libServiceAutorisé5 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2520
            TabIndex        =   85
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label libServiceAutorisé4 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2520
            TabIndex        =   84
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label libServiceAutorisé3 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2520
            TabIndex        =   83
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label libServiceAutorisé1 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2520
            TabIndex        =   82
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label libServiceAutorisé2 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2520
            TabIndex        =   81
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblServiceAutorisé5 
            Caption         =   "Service autorisé 5"
            Height          =   255
            Left            =   50
            TabIndex        =   80
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lblServiceAutorisé3 
            Caption         =   "Service autorisé 3"
            Height          =   255
            Left            =   50
            TabIndex        =   79
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblServiceAutorisé4 
            Caption         =   "Service autorisé 4"
            Height          =   255
            Left            =   50
            TabIndex        =   78
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblServiceAutorisé2 
            Caption         =   "Service autorisé 2"
            Height          =   255
            Left            =   50
            TabIndex        =   77
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblServiceAutorisé1 
            Caption         =   "Service autorisé 1"
            Height          =   225
            Left            =   50
            TabIndex        =   76
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblAmjCréation 
            Caption         =   "Date création du Compte"
            Height          =   255
            Left            =   50
            TabIndex        =   75
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblAmjModification 
            Caption         =   "Date Dernière Modification"
            Height          =   255
            Left            =   50
            TabIndex        =   74
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label lblAmjRéactivation 
            Caption         =   "Date de Réactivation"
            Height          =   255
            Left            =   50
            TabIndex        =   73
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label lblNomOpérateur 
            Caption         =   "Nom de l'opérateur"
            Height          =   255
            Left            =   50
            TabIndex        =   72
            Top             =   3720
            Width           =   1935
         End
      End
      Begin VB.TextBox txtSociété 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1665
         TabIndex        =   53
         Top             =   -592
         Width           =   495
      End
      Begin VB.TextBox txtConditions 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2640
         TabIndex        =   52
         Top             =   -1228
         Width           =   495
      End
      Begin VB.TextBox txtPrélèvementLibératoire 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   51
         Top             =   -442
         Width           =   255
      End
      Begin VB.TextBox txtServiceAutorisét2 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -71640
         TabIndex        =   50
         Top             =   1800
         Width           =   495
      End
      Begin VB.Frame FraExtrait 
         Caption         =   "Extrait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   40
         Top             =   2880
         Width           =   9000
         Begin VB.TextBox txtExtrait 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   44
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtExtraitAmj 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   43
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtExtraitSolde 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6600
            TabIndex        =   42
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtExtraitNuméro 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6600
            TabIndex        =   41
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblExtrait 
            Caption         =   "Code Extrait"
            Height          =   225
            Left            =   50
            TabIndex        =   49
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblExtraitAmj 
            Caption         =   "Date  Dernier extrait"
            Height          =   225
            Left            =   50
            TabIndex        =   48
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblExtraitSolde 
            Caption         =   "Solde Dernier extrait"
            Height          =   225
            Left            =   4560
            TabIndex        =   47
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblExtraitNuméro 
            Caption         =   "Numéro dernier extrait"
            Height          =   225
            Left            =   4560
            TabIndex        =   46
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label libExtrait 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3120
            TabIndex        =   45
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame FraEchelle 
         Caption         =   "Echelles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   32
         Top             =   4200
         Width           =   9000
         Begin VB.TextBox txtEchelle 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   35
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtEchelleAmj 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   34
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtEchelleSolde 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6585
            TabIndex        =   33
            Top             =   540
            Width           =   1935
         End
         Begin VB.Label lblEchelles 
            Caption         =   "Code Echelles"
            Height          =   225
            Left            =   50
            TabIndex        =   39
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label lblEchelleAmj 
            Caption         =   "Date Arrêté"
            Height          =   225
            Left            =   50
            TabIndex        =   38
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label lblEchelleSolde 
            Caption         =   "Solde en valeur dernier arrêté"
            Height          =   225
            Left            =   4320
            TabIndex        =   37
            Top             =   720
            Width           =   2250
         End
         Begin VB.Label libEchelle 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3120
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraSoldeComptable 
         Caption         =   "SOLDES COMPTABLES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   9000
         Begin VB.Frame Frame3 
            Height          =   1455
            Left            =   1200
            TabIndex        =   18
            Top             =   555
            Width           =   5310
            Begin VB.TextBox txtDébitExercice 
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   24
               Top             =   195
               Width           =   2000
            End
            Begin VB.TextBox txtDébitFindeMois 
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   150
               TabIndex        =   23
               Top             =   585
               Width           =   2000
            End
            Begin VB.TextBox txtDébitExerciceAntérieur 
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   22
               Top             =   1020
               Width           =   2000
            End
            Begin VB.TextBox txtCréditExercice 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3120
               TabIndex        =   21
               Top             =   210
               Width           =   2000
            End
            Begin VB.TextBox txtCréditFindeMois 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3120
               TabIndex        =   20
               Top             =   600
               Width           =   2000
            End
            Begin VB.TextBox txtCréditExerciceAntérieur 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3120
               TabIndex        =   19
               Top             =   960
               Width           =   2000
            End
         End
         Begin VB.Frame Frame4 
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
            Left            =   6585
            TabIndex        =   14
            Top             =   555
            Width           =   2295
            Begin VB.TextBox txtSoldeExercice 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   2000
            End
            Begin VB.TextBox txtSoldeFindeMois 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   120
               TabIndex        =   16
               Top             =   555
               Width           =   2000
            End
            Begin VB.TextBox txtSoldeExerciceAntérieur 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   120
               TabIndex        =   15
               Top             =   960
               Width           =   2000
            End
         End
         Begin VB.TextBox txtAmjDernierMouvement 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   13
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblDéb 
            Caption         =   "Débit"
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
            Left            =   2040
            TabIndex        =   31
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblCré 
            Caption         =   "Crédit"
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
            Left            =   5040
            TabIndex        =   30
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblDébitExercice 
            Caption         =   "Veille          :"
            Height          =   255
            Left            =   50
            TabIndex        =   29
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label lblDébitFindeMois 
            Caption         =   "Fin de Mois :"
            Height          =   255
            Left            =   50
            TabIndex        =   28
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label lblDébitExerciceAntérieur 
            Caption         =   "Fin d'année :"
            Height          =   255
            Left            =   50
            TabIndex        =   27
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label lblSolde 
            Caption         =   "Solde"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   26
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblAmjDernierMouvement 
            Caption         =   "Dernier Mouvement :"
            Height          =   255
            Left            =   50
            TabIndex        =   25
            Top             =   2280
            Width           =   1935
         End
      End
      Begin VB.Frame fraSoldeInstantané 
         Caption         =   "SOLDES INSTANTANES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   9000
         Begin VB.TextBox txtSoldeInstantané 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1600
            TabIndex        =   164
            Top             =   1150
            Width           =   2000
         End
         Begin VB.TextBox txtSoldeEnValeurVeille 
            Height          =   300
            Left            =   1600
            TabIndex        =   163
            Top             =   1650
            Width           =   2000
         End
         Begin VB.TextBox txtDécouvertMontant 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1600
            TabIndex        =   6
            Top             =   520
            Width           =   2000
         End
         Begin VB.TextBox txtDécouvertAmj 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   5
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtSoldeEnValeurPostérieur 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   4
            Top             =   1545
            Width           =   1575
         End
         Begin VB.Label lblSoldeIntantané 
            Caption         =   "Solde Instantané"
            Height          =   255
            Left            =   50
            TabIndex        =   11
            Top             =   1200
            Width           =   1410
         End
         Begin VB.Label lblSoldeEnValeurVeille 
            Caption         =   "Solde en valeur"
            Height          =   255
            Left            =   50
            TabIndex        =   10
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label lblDécouvertMontant 
            Caption         =   "Découvert Autorisé"
            Height          =   255
            Left            =   50
            TabIndex        =   9
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblDécouvertAmj 
            Caption         =   "Jusqu'au :"
            Height          =   255
            Left            =   4560
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblSoldeEnValeurPostérieure 
            Caption         =   "Val > jour :"
            Height          =   255
            Left            =   4560
            TabIndex        =   7
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15255
         Left            =   -240
         TabIndex        =   2
         Top             =   6720
         Width           =   9975
      End
      Begin VB.Label lblCosOc 
         Caption         =   "Code Societé"
         Height          =   225
         Left            =   50
         TabIndex        =   60
         Top             =   -272
         Width           =   1935
      End
      Begin VB.Label lblConditions 
         Caption         =   "Conditions"
         Height          =   225
         Left            =   50
         TabIndex        =   59
         Top             =   -1108
         Width           =   2295
      End
      Begin VB.Label lblPrélèvementLibératoire 
         Caption         =   " Prélèvement  libératoire "
         Height          =   225
         Left            =   60
         TabIndex        =   58
         Top             =   -397
         Width           =   2490
      End
      Begin VB.Label lblSociété 
         Caption         =   "Code Société"
         Height          =   255
         Left            =   50
         TabIndex        =   57
         Top             =   -487
         Width           =   1410
      End
      Begin VB.Label libPrélèvementLibératoire 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3435
         TabIndex        =   56
         Top             =   -382
         Width           =   2655
      End
      Begin VB.Label libConditions 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3600
         TabIndex        =   55
         Top             =   -276
         Width           =   2655
      End
      Begin VB.Label libDeviseX 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2625
         TabIndex        =   54
         Top             =   -382
         Width           =   855
      End
   End
   Begin VB.Label lblErrMsg 
      Caption         =   "lblErrMsg"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6135
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2805
   End
End
Attribute VB_Name = "frmCptInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean

Private recCptInfo As typeCptInfo
Private Mt As Currency





Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim Msg As String

Msg = Format$(0, "000000") & Format$(0, "000000")

prtCptInfoX Msg

End Sub



'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub



'---------------------------------------------------------
Private Sub cmdRacine_Click()
'---------------------------------------------------------
Dim x As String

If recCptInfo.TypeGA <> "A" Then
    lblErrMsg = "Ce n'est pas un compte auxilaire"
    lblErrMsgStd
Else
    
    arrRacineNbMax = 1
    arrRacineNb = 0
    ReDim arrRacine(1)
    Call recRacineInit(arrRacine(0))
    arrRacine(0).Method = "SeekL0"
    arrRacine(0).Numéro = Val(mId$(recCptInfo.Numéro, 1, 5))
    If IsNull(srvRacineMon(arrRacine(0))) Then

        frmRacine.Show vbModeless
        frmRacine.Visible = True
        frmRacine.dspRacineInit
        x = frmRacine.Caption
        AppActivate x
    End If
End If

End Sub
'---------------------------------------------------------
Public Sub lblErrMsgStd()
'---------------------------------------------------------

Beep
lblErrMsg.Visible = True
lblErrMsg.FontBold = True
lblErrMsg.FontSize = 8
lblErrMsg.Top = 0
lblErrMsg.Left = 6000

End Sub



Private Sub Form_Activate()
Set XForm = Me

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select

lblErrMsg.Visible = False

End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb, "Locked")
ReDim arrTag(arrTagNb + 1)

End Sub



'---------------------------------------------------------
Public Sub dspCompte()
'---------------------------------------------------------
Dim Mt As Currency

txtSociété = recCptInfo.Société
txtAgence = recCptInfo.Agence

txtDevise = recCptInfo.Devise
libDeviseX = recCptInfo.Devise

txtCourrier = recCptInfo.Courrier
libCourrier = DicLib(61, recCptInfo.Courrier)

txtNuméro = Format$(Val(recCptInfo.Numéro), "@@@@@.@@@.@@.@")

txtNuméroAncien = Format$(Val(recCptInfo.NuméroAncien), "@@@@@@.@@.@@@")

txtIntitulé = recCptInfo.Intitulé

txtIntitulé2 = recCptInfo.Intitulé2

txtTypeGA = recCptInfo.TypeGA
Select Case recCptInfo.TypeGA
    Case "A": libTypeGA = "Auxiliaire"
    Case "G": libTypeGA = "General"
    Case "R": libTypeGA = "Regroupement"
End Select

txtSituation = recCptInfo.Situation
Select Case recCptInfo.Situation
    Case "A": libSituation = "Annulé"
    Case "B": libSituation = "Bloqué"
    Case Else: libSituation = recCptInfo.Situation
End Select

libGestionnaire = DicLib(60, recCptInfo.Gestionnaire)
txtGestionnaire = recCptInfo.Gestionnaire

txtSoldeInstantané.ForeColor = IIf(recCptInfo.SoldeInstantané < 0, vbRed, vbBlue)
txtSoldeInstantané.Text = Format$(recCptInfo.SoldeInstantané, "#### ### ### ### ##0.00")

txtDécouvertMontant.ForeColor = IIf(recCptInfo.DécouvertMontant < 0, vbRed, vbBlue)
txtDécouvertMontant = Format$(recCptInfo.DécouvertMontant, "#### ### ### ### ##0.00")

txtDécouvertAmj = dateImp(recCptInfo.DécouvertAmj)
txtAlpha = recCptInfo.Alpha
txtNatureTitulaire = recCptInfo.NatureTitulaire

libActionnaire = DicLib(7, recCptInfo.Actionnaire)
txtActionnaire = recCptInfo.Actionnaire

libRésident = DicLib(23, recCptInfo.Résident)
txtRésident = recCptInfo.Résident

libApporteur = DicLib(58, recCptInfo.Apporteur)
txtApporteur = recCptInfo.Apporteur

txtCompteGénéral = recCptInfo.CompteGénéral

txtSens = recCptInfo.Sens
libSens = DicLib(2, recCptInfo.Sens)
txtClérib = recCptInfo.CléRib

txtNostro = recCptInfo.Nostro
Select Case recCptInfo.Nostro
    Case "L": libNostro = "Loro"
    Case "N": libNostro = "Nostro"
End Select



libConditions = DicLib(522, recCptInfo.Conditions)
txtConditions = recCptInfo.Conditions

libPrélèvementLibératoire = DicLib(7, recCptInfo.PrélèvementLibératoire)
txtPrélèvementLibératoire = recCptInfo.PrélèvementLibératoire

libServiceResponsable = DicLib(4, recCptInfo.ServiceResponsable)
txtServiceResponsable = recCptInfo.ServiceResponsable

libServiceAutorisé1 = DicLib(4, recCptInfo.ServiceAutorisé1)
txtServiceAutorisé1 = recCptInfo.ServiceAutorisé1

libServiceAutorisé2 = DicLib(4, recCptInfo.ServiceAutorisé2)
txtServiceAutorisé2 = recCptInfo.ServiceAutorisé2

libServiceAutorisé3 = DicLib(4, recCptInfo.ServiceAutorisé3)
txtServiceAutorisé3 = recCptInfo.ServiceAutorisé3

libServiceAutorisé4 = DicLib(4, recCptInfo.ServiceAutorisé4)
txtServiceAutorisé4 = recCptInfo.ServiceAutorisé4

libServiceAutorisé5 = DicLib(4, recCptInfo.ServiceAutorisé5)
txtServiceAutorisé5 = recCptInfo.ServiceAutorisé5

txtDébitExercice.ForeColor = IIf(recCptInfo.DébitExercice < 0, vbRed, vbBlue)
txtDébitExercice = Format$(recCptInfo.DébitExercice, "#### ### ### ### ##0.00")

txtCréditExercice.ForeColor = IIf(recCptInfo.CréditExercice < 0, vbRed, vbBlue)
txtCréditExercice = Format$(recCptInfo.CréditExercice, "#### ### ### ### ##0.00")

Mt = recCptInfo.DébitExercice + recCptInfo.CréditExercice
txtSoldeExercice.ForeColor = IIf(Mt < 0, vbRed, vbBlue)
txtSoldeExercice = Format$(Mt, "#### ### ### ### ##0.00")

txtDébitFindeMois.ForeColor = IIf(recCptInfo.DébitFindeMois < 0, vbRed, vbBlue)
txtDébitFindeMois = Format$(recCptInfo.DébitFindeMois, "#### ### ### ### ##0.00")

txtCréditFindeMois.ForeColor = IIf((recCptInfo.CréditFindeMois) < 0, vbRed, vbBlue)
txtCréditFindeMois = Format$(recCptInfo.CréditFindeMois, "#### ### ### ### ##0.00")

Mt = recCptInfo.DébitFindeMois + recCptInfo.CréditFindeMois
txtSoldeFindeMois.ForeColor = IIf(Mt < 0, vbRed, vbBlue)
txtSoldeFindeMois = Format$(Mt, "#### ### ### ### ##0.00")



txtDébitExerciceAntérieur.ForeColor = IIf(recCptInfo.DébitExerciceAntérieur < 0, vbRed, vbBlue)
txtDébitExerciceAntérieur = Format$(recCptInfo.DébitExerciceAntérieur, "#### ### ### ### ##0.00")

txtCréditExerciceAntérieur.ForeColor = IIf(recCptInfo.CréditExerciceAntérieur < 0, vbRed, vbBlue)
txtCréditExerciceAntérieur = Format$(recCptInfo.CréditExerciceAntérieur, "#### ### ### ### ##0.00")

Mt = recCptInfo.DébitExerciceAntérieur + recCptInfo.CréditExerciceAntérieur
txtSoldeExerciceAntérieur.ForeColor = IIf(Mt < 0, vbRed, vbBlue)
txtSoldeExerciceAntérieur = Format$(Mt, "#### ### ### ### ##0.00")

txtSoldeEnValeurVeille.ForeColor = IIf(recCptInfo.SoldeEnValeurVeille < 0, vbRed, vbBlue)
txtSoldeEnValeurVeille.Text = Format$(recCptInfo.SoldeEnValeurVeille, "#### ### ### ### ##0.00")

txtSoldeEnValeurPostérieur.ForeColor = IIf(recCptInfo.SoldeEnValeurPostérieur < 0, vbRed, vbBlue)
txtSoldeEnValeurPostérieur = Format$(recCptInfo.SoldeEnValeurPostérieur, "#### ### ### ### ##0.00")


libEchelle = DicLib(7, recCptInfo.Echelle)
txtEchelle = recCptInfo.Echelle
txtEchelleSolde.ForeColor = IIf(recCptInfo.EchelleSolde < 0, vbRed, vbBlue)
txtEchelleSolde = Format$(recCptInfo.EchelleSolde, "#### ### ### ### ##0.00")
txtEchelleAmj = dateImp(recCptInfo.EchelleAmj)
'------------------------------1)

libExtrait = DicLib(6, recCptInfo.Extrait)
txtExtrait = recCptInfo.Extrait
txtExtraitSolde.ForeColor = IIf(recCptInfo.ExtraitSolde < 0, vbRed, vbBlue)
txtExtraitSolde = Format$(recCptInfo.ExtraitSolde, "#### ### ### ### ##0.00")
txtExtraitAmj = dateImp(recCptInfo.ExtraitAmj)
txtExtraitNuméro = recCptInfo.ExtraitNuméro

txtAmjCréation = dateImp(recCptInfo.AmjCréation)
txtAmjModification = dateImp(recCptInfo.AmjModification)
txtAmjAnnulation = dateImp(recCptInfo.AmjAnnulation)
txtAmjRéactivation = dateImp(recCptInfo.AmjRéactivation)
txtAmjDernierMouvement = dateImp(recCptInfo.AmjDernierMouvement)
txtNomOpérateur = recCptInfo.NomOpérateur
txtAdresse1 = recCptInfo.Adresse1
txtAdresse2 = recCptInfo.Adresse2
txtAdresse3 = recCptInfo.Adresse3
txtAdresse4 = recCptInfo.Adresse4
txtAdresse5 = recCptInfo.Adresse5
txtAdresseCp = recCptInfo.AdresseCP
txtAdresseBd = recCptInfo.AdresseBD
txtAdressePays = recCptInfo.AdressePays
txtAmjDernierMouvement = dateImp(recCptInfo.AmjDernierMouvement)
txtAmjCréation = dateImp(recCptInfo.AmjCréation)
txtAmjModification = dateImp(recCptInfo.AmjModification)
txtAmjAnnulation = dateImp(recCptInfo.AmjAnnulation)
txtAmjRéactivation = dateImp(recCptInfo.AmjRéactivation)
txtEchelleAmj = dateImp(recCptInfo.EchelleAmj)
txtDécouvertAmj = dateImp(recCptInfo.DécouvertAmj)
txtExtraitAmj = dateImp(recCptInfo.ExtraitAmj)

End Sub


'---------------------------------------------------------
Public Sub dspCompteInit()
'---------------------------------------------------------

recCptInfo = arrCptInfo(0)
dspCompte


End Sub



Private Sub libCourrier1_Click()

End Sub


Private Sub lblcptAgc_Click()

End Sub




