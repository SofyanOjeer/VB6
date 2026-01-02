VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRacine 
   Caption         =   "Compte : Racine"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6915
   ScaleWidth      =   9660
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   9240
      Picture         =   "RACINE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   246
      Top             =   0
      Width           =   400
   End
   Begin VB.TextBox txtTitulaireNo 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3720
      TabIndex        =   19
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtRacine 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Top             =   0
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7320
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12912
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   10
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   16711680
      TabCaption(0)   =   "Racine 1/3"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Racine 2/3"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Racine 3/3"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Titulaire 1/4"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11"
      Tab(3).Control(1)=   "Frame12"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Titulaire 2/4"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(1)=   "Frame13"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Titulaire 3/4"
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame15"
      Tab(5).Control(1)=   "Frame14"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Titulaire 4/4"
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame16"
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame14 
         Caption         =   "Personne Morale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74950
         TabIndex        =   219
         Top             =   480
         Width           =   9400
         Begin VB.TextBox txtPmFormeJuridique 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   224
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtPmSiren 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   223
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtPmApe 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            TabIndex        =   222
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtPmBdfId 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   221
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtPmNationalitéMère 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   220
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label libPmNationalitéMère 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   6720
            TabIndex        =   241
            Top             =   1130
            Width           =   2295
         End
         Begin VB.Label libPmApe 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   240
            Top             =   1130
            Width           =   1215
         End
         Begin VB.Label libPmFormeJuridique 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   239
            Top             =   380
            Width           =   3015
         End
         Begin VB.Label lblPmSiren 
            Caption         =   "Numéro de Siren"
            Height          =   255
            Left            =   50
            TabIndex        =   229
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblPmApe 
            Caption         =   "Code Ape"
            Height          =   255
            Left            =   50
            TabIndex        =   228
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblPmBdfId 
            Caption         =   "Identification Bénéficiaire"
            Height          =   255
            Left            =   4080
            TabIndex        =   227
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label lblPmFormeJuridique 
            Caption         =   "Forme Juridique"
            Height          =   255
            Left            =   50
            TabIndex        =   226
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblPmNationalitéMère 
            Caption         =   "Société Mère"
            Height          =   255
            Left            =   4080
            TabIndex        =   225
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.Frame Frame16 
         Height          =   2295
         Left            =   -74950
         TabIndex        =   208
         Top             =   480
         Width           =   9400
         Begin VB.TextBox txtAmjCre 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   213
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtAmjMod 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   212
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtAmjAnn 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   211
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtAmjRea 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   210
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtNomOp 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   209
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblAmjCre 
            Caption         =   "Date de Création"
            Height          =   255
            Left            =   50
            TabIndex        =   218
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblAmjMod 
            Caption         =   "Date de Modification"
            Height          =   255
            Left            =   50
            TabIndex        =   217
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblAmjAnn 
            Caption         =   "Date Annulation"
            Height          =   255
            Left            =   50
            TabIndex        =   216
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblAmjRea 
            Caption         =   "Date de Réactivation"
            Height          =   255
            Left            =   50
            TabIndex        =   215
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblNomOp 
            Caption         =   "Nom Opérateur"
            Height          =   255
            Left            =   50
            TabIndex        =   214
            Top             =   1680
            Width           =   2175
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Personne Physique"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74950
         TabIndex        =   191
         Top             =   2640
         Width           =   9400
         Begin VB.TextBox txtPpCsp 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   207
            Top             =   2280
            Width           =   255
         End
         Begin VB.TextBox txtPpAmjSortie 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   5520
            TabIndex        =   205
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox txtPpAmjEntrée 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   203
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox txtPpProfession 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   201
            Top             =   1920
            Width           =   3975
         End
         Begin VB.TextBox txtPpBdfId 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   199
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtPpNombreEnfant 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   197
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox txtPpRégimeMatrimonial 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   195
            Top             =   840
            Width           =   255
         End
         Begin VB.TextBox txtPpQualité 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   193
            Top             =   480
            Width           =   255
         End
         Begin VB.Label libPpCsp 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2910
            TabIndex        =   244
            Top             =   2330
            Width           =   1935
         End
         Begin VB.Label libPpRégimeMatrimonial 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2910
            TabIndex        =   243
            Top             =   910
            Width           =   3135
         End
         Begin VB.Label libPpQualité 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2910
            TabIndex        =   242
            Top             =   550
            Width           =   1455
         End
         Begin VB.Label lblPpCsp 
            Caption         =   "Catégorie Socio-Profes."
            Height          =   255
            Left            =   45
            TabIndex        =   206
            Top             =   2400
            Width           =   1770
         End
         Begin VB.Label lblPpAmjSortie 
            Caption         =   "Date sortie de France"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   204
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label lblPpAmjEntrée 
            Caption         =   "Date entrée en France"
            Height          =   255
            Left            =   50
            TabIndex        =   202
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label lblPpProfession 
            Caption         =   "Profession"
            Height          =   255
            Index           =   0
            Left            =   50
            TabIndex        =   200
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label lblPpBdfId 
            Caption         =   "Domiciliation-Clé bdf"
            Height          =   255
            Left            =   50
            TabIndex        =   198
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lblPpNombreEnfant 
            Caption         =   "Nombre enfants"
            Height          =   255
            Index           =   0
            Left            =   45
            TabIndex        =   196
            Top             =   1320
            Width           =   1650
         End
         Begin VB.Label lblPpRégimeMatrimonial 
            Caption         =   "Régime Matrimonial"
            Height          =   255
            Index           =   0
            Left            =   50
            TabIndex        =   194
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblPpQualité 
            Caption         =   "Qualité Titulaire"
            Height          =   255
            Left            =   50
            TabIndex        =   192
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Document Identité"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74950
         TabIndex        =   182
         Top             =   3480
         Width           =   9400
         Begin VB.TextBox txtDocIdentitéNo 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   190
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtDocIdentitéAmj 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   188
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtDocIdentitéDélivréPar 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   186
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtDocIdentitéType 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   184
            Top             =   240
            Width           =   255
         End
         Begin VB.Label libDocIdentitéType 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3600
            TabIndex        =   238
            Top             =   290
            Width           =   3255
         End
         Begin VB.Label lblDocIdentitéNo 
            Caption         =   "Numéro "
            Height          =   255
            Left            =   50
            TabIndex        =   189
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblDocIdentitéAmj 
            Caption         =   "Délivré le"
            Height          =   255
            Left            =   50
            TabIndex        =   187
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblDocIdentitéDélivréPar 
            Caption         =   "Délivré Par"
            Height          =   255
            Index           =   0
            Left            =   50
            TabIndex        =   185
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblDocIdentitéType 
            Caption         =   "Nat.Doc.Identité"
            Height          =   255
            Left            =   50
            TabIndex        =   183
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame12 
         Height          =   4095
         Left            =   -74950
         TabIndex        =   165
         Top             =   1920
         Width           =   9400
         Begin VB.TextBox txtNomPatronymique 
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   1920
            TabIndex        =   232
            Top             =   240
            Width           =   7215
         End
         Begin VB.TextBox txtPrénom 
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   1920
            TabIndex        =   231
            Top             =   840
            Width           =   7215
         End
         Begin VB.TextBox txtAdresseBD 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2760
            TabIndex        =   181
            Top             =   3600
            Width           =   3255
         End
         Begin VB.TextBox txtAdresseCP 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   180
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtAdresse4 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   179
            Top             =   3240
            Width           =   3975
         End
         Begin VB.TextBox txtAdresse3 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   178
            Top             =   2880
            Width           =   3975
         End
         Begin VB.TextBox txtAdresse2 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   177
            Top             =   2520
            Width           =   3975
         End
         Begin VB.TextBox txtAdresse1 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   176
            Top             =   2160
            Width           =   3975
         End
         Begin VB.TextBox txtMariPrénom 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   170
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox txtMariNom 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   169
            Top             =   1440
            Width           =   3975
         End
         Begin VB.Label lblMariPrénom 
            Caption         =   "Prénom du Mari"
            Height          =   255
            Left            =   50
            TabIndex        =   230
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblAdresseCP 
            Caption         =   "Cd Posal Fiscale Client 5"
            Height          =   255
            Left            =   50
            TabIndex        =   175
            Top             =   3720
            Width           =   1890
         End
         Begin VB.Label lblAdresse4 
            Caption         =   "Adresse Fiscale Client 4"
            Height          =   255
            Left            =   50
            TabIndex        =   174
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label lblAdresse3 
            Caption         =   "Adresse Fiscale Client 3"
            Height          =   255
            Left            =   50
            TabIndex        =   173
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Label lblAdresse2 
            Caption         =   "Adresse Fiscale Client 2"
            Height          =   255
            Left            =   50
            TabIndex        =   172
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label lblAdresse1 
            Caption         =   "Adresse Fiscale Client 1"
            Height          =   255
            Left            =   50
            TabIndex        =   171
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label lblMariNom 
            Caption         =   "Nom du Mari"
            Height          =   255
            Left            =   50
            TabIndex        =   168
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblPrénom 
            Caption         =   "Prénoms"
            Height          =   255
            Index           =   0
            Left            =   50
            TabIndex        =   167
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblNomPatronimique 
            Caption         =   "Nom Patronymique"
            Height          =   255
            Left            =   50
            TabIndex        =   166
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1455
         Left            =   -74950
         TabIndex        =   158
         Top             =   360
         Width           =   9400
         Begin VB.TextBox txtRésidenceZone 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2400
            TabIndex        =   164
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox txtRésidencePays 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2400
            TabIndex        =   163
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox TxtNationalité 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2400
            TabIndex        =   162
            Top             =   240
            Width           =   615
         End
         Begin VB.Label libRésidenceZone 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3360
            TabIndex        =   235
            Top             =   1010
            Width           =   3015
         End
         Begin VB.Label libRésidencePays 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3360
            TabIndex        =   234
            Top             =   650
            Width           =   3015
         End
         Begin VB.Label libNationalité 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3360
            TabIndex        =   233
            Top             =   290
            Width           =   2895
         End
         Begin VB.Label lblRésidenceZone 
            Caption         =   "Zone Géographique Fiscale"
            Height          =   255
            Left            =   50
            TabIndex        =   161
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblRésidencepays 
            Caption         =   "Pays de Résidence"
            Height          =   255
            Left            =   50
            TabIndex        =   160
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label lblNationalité 
            Caption         =   "Nationalité"
            Height          =   255
            Left            =   50
            TabIndex        =   159
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame10 
         Height          =   2175
         Left            =   -74950
         TabIndex        =   106
         Top             =   3960
         Width           =   9400
         Begin VB.TextBox txtSuccession 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   127
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox txtRésident 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   125
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtRégimeMatrimonial 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   124
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox txtBdfIdentification 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   122
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtRacNationalité 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   121
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtOptionTva 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   120
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtAmjNaissance 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6000
            TabIndex        =   119
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtRésidentPays 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   118
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtRésidentFiscal 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   117
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox txtTypeOuverture 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   116
            Top             =   240
            Width           =   255
         End
         Begin VB.Label libRacNationalité 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7200
            TabIndex        =   147
            Top             =   1010
            Width           =   1935
         End
         Begin VB.Label libOptionTva 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7200
            TabIndex        =   139
            Top             =   650
            Width           =   1935
         End
         Begin VB.Label libAmjNaissance 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7200
            TabIndex        =   138
            Top             =   290
            Width           =   2055
         End
         Begin VB.Label libRégimeMatrimonial 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7200
            TabIndex        =   137
            Top             =   1730
            Width           =   1935
         End
         Begin VB.Label libSuccession 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   135
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label libRésidentPays 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   133
            Top             =   1370
            Width           =   1935
         End
         Begin VB.Label libRésidentFiscal 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   132
            Top             =   1010
            Width           =   1935
         End
         Begin VB.Label libRésident 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   131
            Top             =   650
            Width           =   1935
         End
         Begin VB.Label libTypeOuverture 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   128
            Top             =   290
            Width           =   1935
         End
         Begin VB.Label lblRégimeMatrimonial 
            Caption         =   "Régime Matrimonial"
            Height          =   255
            Left            =   4440
            TabIndex        =   123
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblBdfIdentification 
            Caption         =   "Identification Bdf"
            Height          =   255
            Left            =   4440
            TabIndex        =   115
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lblRacNationalité 
            Caption         =   "Nationalité"
            Height          =   255
            Left            =   4440
            TabIndex        =   114
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblOptionTva 
            Caption         =   "Option TVA"
            Height          =   255
            Left            =   4440
            TabIndex        =   113
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblAmjNaissance 
            Caption         =   "Date de Naissance"
            Height          =   255
            Left            =   4440
            TabIndex        =   112
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblSuccession 
            Caption         =   "Succession"
            Height          =   255
            Left            =   50
            TabIndex        =   111
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblRésidentPays 
            Caption         =   "Pays de Résidence"
            Height          =   255
            Left            =   50
            TabIndex        =   110
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblRésidentFiscal 
            Caption         =   "Résident Fiscal"
            Height          =   255
            Left            =   50
            TabIndex        =   109
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblRésident 
            Caption         =   "Résident Bdf"
            Height          =   255
            Left            =   50
            TabIndex        =   108
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblTypeOuverture 
            Caption         =   "Type de Compte"
            Height          =   255
            Left            =   50
            TabIndex        =   107
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2415
         Left            =   -74950
         TabIndex        =   80
         Top             =   3120
         Width           =   9400
         Begin VB.TextBox txtOpérateur 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   90
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox txtAmjRéactivation 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   88
            Top             =   1320
            Width           =   1150
         End
         Begin VB.TextBox txtAmjAnnulation 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   86
            Top             =   960
            Width           =   1150
         End
         Begin VB.TextBox txtAmjModification 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   84
            Top             =   600
            Width           =   1150
         End
         Begin VB.TextBox txtAmjCréation 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   82
            Top             =   240
            Width           =   1150
         End
         Begin VB.Label lblOpérateur 
            Caption         =   "Opérateur"
            Height          =   255
            Left            =   50
            TabIndex        =   89
            Top             =   1800
            Width           =   1530
         End
         Begin VB.Label lblAmjRéactivation 
            Caption         =   "Date de Réactivation"
            Height          =   255
            Left            =   50
            TabIndex        =   87
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label lblAmjAnnulation 
            Caption         =   "Date Annulation"
            Height          =   255
            Left            =   45
            TabIndex        =   85
            Top             =   1080
            Width           =   1650
         End
         Begin VB.Label lblAmjModification 
            Caption         =   "Date de Modification"
            Height          =   255
            Left            =   50
            TabIndex        =   83
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblAmjCréation 
            Caption         =   "Date de Création"
            Height          =   255
            Left            =   45
            TabIndex        =   81
            Top             =   360
            Width           =   1650
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2655
         Left            =   -74955
         TabIndex        =   60
         Top             =   360
         Width           =   9400
         Begin VB.TextBox txtSiègeRacine 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6840
            TabIndex        =   152
            Top             =   840
            Width           =   255
         End
         Begin VB.TextBox txtApporteur 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   126
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtSiren 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6840
            TabIndex        =   79
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtAPE 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6840
            TabIndex        =   77
            Top             =   1920
            Width           =   375
         End
         Begin VB.TextBox txtChèquierInterditAmj 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6840
            TabIndex        =   75
            Top             =   1560
            Width           =   1150
         End
         Begin VB.TextBox txtChèquierInterdit 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6840
            TabIndex        =   73
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox txtCommentaire 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   71
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox txtCourrier 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   69
            Top             =   1560
            Width           =   255
         End
         Begin VB.TextBox txtGroupeEconomique 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   68
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtGestionnaire 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   67
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtGroupeIdentification 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   66
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label libSiègeRacine 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7680
            TabIndex        =   153
            Top             =   890
            Width           =   1455
         End
         Begin VB.Label lblSiègeRacine 
            Caption         =   "Racine Siège"
            Height          =   255
            Left            =   5400
            TabIndex        =   151
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label libAPE 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7680
            TabIndex        =   148
            Top             =   1970
            Width           =   1935
         End
         Begin VB.Label libChèquierInterdit 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7680
            TabIndex        =   146
            Top             =   1250
            Width           =   1935
         End
         Begin VB.Label libCourrier 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3000
            TabIndex        =   145
            Top             =   1610
            Width           =   2175
         End
         Begin VB.Label libGroupeEconomique 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3000
            TabIndex        =   144
            Top             =   1220
            Width           =   2175
         End
         Begin VB.Label libGestionnaire 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3000
            TabIndex        =   142
            Top             =   890
            Width           =   2175
         End
         Begin VB.Label libApporteur 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3000
            TabIndex        =   141
            Top             =   530
            Width           =   2175
         End
         Begin VB.Label libGroupeIdentification 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3000
            TabIndex        =   140
            Top             =   170
            Width           =   2175
         End
         Begin VB.Label lblSiren 
            Caption         =   "N° de Siren"
            Height          =   255
            Left            =   5400
            TabIndex        =   78
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblAPE 
            Caption         =   "Code APE"
            Height          =   255
            Left            =   5400
            TabIndex        =   76
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblChèquierInterditAmj 
            Caption         =   "Date Interdiction"
            Height          =   255
            Left            =   5400
            TabIndex        =   74
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblChèquierInterdit 
            Caption         =   "Interdiction Chq"
            Height          =   255
            Left            =   5400
            TabIndex        =   72
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblCommentaire 
            Caption         =   "Commentaires"
            Height          =   255
            Left            =   50
            TabIndex        =   70
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblCourrier 
            Caption         =   "Courrier"
            Height          =   255
            Left            =   50
            TabIndex        =   65
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblGroupeEconomique 
            Caption         =   "Groupe Economique"
            Height          =   255
            Left            =   45
            TabIndex        =   64
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblGestionnaire 
            Caption         =   "Gestionnaire"
            Height          =   255
            Left            =   50
            TabIndex        =   63
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblApporteur 
            Caption         =   "Apporteur"
            Height          =   255
            Left            =   50
            TabIndex        =   62
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblGroupeIdentification 
            Caption         =   "Regroupement Client"
            Height          =   255
            Left            =   50
            TabIndex        =   61
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Naissance"
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
         Left            =   -74955
         TabIndex        =   47
         Top             =   480
         Width           =   9400
         Begin VB.TextBox txtNaissanceAmj 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   53
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtNaissanceZone 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   52
            Top             =   2040
            Width           =   255
         End
         Begin VB.TextBox txtNaissanceDépX 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   51
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox txtNaissanceCommune 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   50
            Top             =   1320
            Width           =   3975
         End
         Begin VB.TextBox TXTNaissanceDépN 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   49
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtNaissancePays 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2520
            TabIndex        =   48
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label libNaissanceZone 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3720
            TabIndex        =   245
            Top             =   2090
            Width           =   2775
         End
         Begin VB.Label libNaissancePays 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3720
            TabIndex        =   237
            Top             =   1730
            Width           =   2655
         End
         Begin VB.Label libNaissanceDépN 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3720
            TabIndex        =   236
            Top             =   650
            Width           =   2775
         End
         Begin VB.Label lblNaissanceAmj 
            Caption         =   "Date          "
            Height          =   255
            Left            =   45
            TabIndex        =   59
            Top             =   360
            Width           =   1770
         End
         Begin VB.Label lblNaissanceZone 
            Caption         =   "Zone Géographique "
            Height          =   255
            Left            =   45
            TabIndex        =   58
            Top             =   2160
            Width           =   1650
         End
         Begin VB.Label lblNaissanceDépX 
            Caption         =   "Département  ou Pays     "
            Height          =   255
            Left            =   50
            TabIndex        =   57
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label lblNaissanceCommune 
            Caption         =   "Commune "
            Height          =   255
            Left            =   50
            TabIndex        =   56
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblNaissanceDépN 
            Caption         =   "Département "
            Height          =   255
            Left            =   50
            TabIndex        =   55
            Top             =   720
            Width           =   1890
         End
         Begin VB.Label lblNaissancePays 
            Caption         =   "Pays "
            Height          =   255
            Left            =   45
            TabIndex        =   54
            Top             =   1800
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1575
         Left            =   -74950
         TabIndex        =   36
         Top             =   2280
         Width           =   9400
         Begin VB.Frame Frame9 
            Height          =   3615
            Left            =   -120
            TabIndex        =   91
            Top             =   3720
            Width           =   9015
            Begin VB.TextBox Text12 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   6000
               TabIndex        =   105
               Text            =   "Text5"
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox Text6 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1560
               TabIndex        =   103
               Text            =   "Text5"
               Top             =   480
               Width           =   1935
            End
            Begin VB.TextBox Text13 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   6000
               TabIndex        =   101
               Text            =   "Text5"
               Top             =   1200
               Width           =   1935
            End
            Begin VB.TextBox Text7 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1560
               TabIndex        =   99
               Text            =   "Text5"
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox Text14 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   6000
               TabIndex        =   97
               Text            =   "Text5"
               Top             =   1560
               Width           =   1935
            End
            Begin VB.TextBox Text8 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1560
               TabIndex        =   95
               Text            =   "Text5"
               Top             =   1200
               Width           =   1935
            End
            Begin VB.TextBox Text9 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1560
               TabIndex        =   93
               Text            =   "Text5"
               Top             =   1560
               Width           =   1935
            End
            Begin VB.Label Label8 
               Caption         =   "Option TVA"
               Height          =   255
               Left            =   4440
               TabIndex        =   104
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label2 
               Caption         =   "Résident BDF"
               Height          =   255
               Left            =   0
               TabIndex        =   102
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label9 
               Caption         =   "Nationalité"
               Height          =   255
               Left            =   4440
               TabIndex        =   100
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label3 
               Caption         =   "Résident Fiscal"
               Height          =   255
               Left            =   0
               TabIndex        =   98
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label10 
               Caption         =   "Identification Bdf"
               Height          =   255
               Left            =   4440
               TabIndex        =   96
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label Label4 
               Caption         =   "Pays de Résidence"
               Height          =   255
               Left            =   0
               TabIndex        =   94
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label5 
               Caption         =   "Succession"
               Height          =   255
               Left            =   0
               TabIndex        =   92
               Top             =   1680
               Width           =   1695
            End
         End
         Begin VB.TextBox txtTélex 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6480
            TabIndex        =   46
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtSwift 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6480
            TabIndex        =   45
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtFax 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6480
            TabIndex        =   44
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtTéléphone2 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   43
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtTéléphone1 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1800
            TabIndex        =   42
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblSwift 
            Caption         =   "Numéro Swift"
            Height          =   255
            Left            =   4560
            TabIndex        =   41
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblFax 
            Caption         =   "Numéro de Fax"
            Height          =   255
            Left            =   4560
            TabIndex        =   40
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblTélex 
            Caption         =   "Numéro de Télex"
            Height          =   255
            Left            =   4560
            TabIndex        =   39
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblTéléphone2 
            Caption         =   "N° de Téléphone 2"
            Height          =   255
            Left            =   50
            TabIndex        =   38
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblTéléphone1 
            Caption         =   "N° de Téléphone 1"
            Height          =   255
            Left            =   50
            TabIndex        =   37
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   -74950
         TabIndex        =   27
         Top             =   480
         Width           =   9400
         Begin VB.TextBox txtRacAdrCp 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1560
            TabIndex        =   35
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtRacAdr03 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1560
            TabIndex        =   34
            Top             =   960
            Width           =   5415
         End
         Begin VB.TextBox txtRacAdr02 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1560
            TabIndex        =   33
            Top             =   600
            Width           =   5415
         End
         Begin VB.TextBox txtRacAdr01 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1560
            TabIndex        =   32
            Top             =   240
            Width           =   5415
         End
         Begin VB.Label lblRacAdrCp 
            Caption         =   "Code Postal"
            Height          =   255
            Left            =   45
            TabIndex        =   31
            Top             =   1440
            Width           =   1410
         End
         Begin VB.Label lblRacAdr03 
            Caption         =   "Adresse 03"
            Height          =   255
            Left            =   45
            TabIndex        =   30
            Top             =   1080
            Width           =   1410
         End
         Begin VB.Label lblRacAdr02 
            Caption         =   "Adresse 02"
            Height          =   255
            Left            =   45
            TabIndex        =   29
            Top             =   720
            Width           =   1410
         End
         Begin VB.Label lblRacAdr01 
            Caption         =   "Adresse 01"
            Height          =   255
            Left            =   45
            TabIndex        =   28
            Top             =   360
            Width           =   1410
         End
      End
      Begin VB.Frame Frame8 
         Height          =   4575
         Left            =   50
         TabIndex        =   5
         Top             =   720
         Width           =   9400
         Begin VB.TextBox txtNombreTitulaire 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   157
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox txtNatureTitulaire 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   156
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txtOpposition 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   26
            Top             =   3480
            Width           =   255
         End
         Begin VB.TextBox txtNomination 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   24
            Top             =   3120
            Width           =   255
         End
         Begin VB.TextBox txtMembreduPersonnel 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   23
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox txtActionnaire 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   22
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox txtTypeBanqueClient 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   21
            Top             =   1320
            Width           =   255
         End
         Begin VB.TextBox txtAlpha 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   20
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtNuméro 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtIntitulé 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   555
            Width           =   6735
         End
         Begin VB.Label libNombreTitulaire 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   155
            Top             =   2450
            Width           =   2175
         End
         Begin VB.Label lblNombreTitulaire 
            Caption         =   "Nombre de Titulaire"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   154
            Top             =   2505
            Width           =   1695
         End
         Begin VB.Label libTypeBanqueClient 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   150
            Top             =   1420
            Width           =   3015
         End
         Begin VB.Label libOpposition 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   143
            Top             =   3530
            Width           =   4215
         End
         Begin VB.Label libNomination 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   136
            Top             =   3170
            Width           =   4215
         End
         Begin VB.Label libMembreduPersonnel 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   134
            Top             =   2810
            Width           =   4215
         End
         Begin VB.Label libNatureTitulaire 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   130
            Top             =   2090
            Width           =   4215
         End
         Begin VB.Label libActionnaire 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   129
            Top             =   1730
            Width           =   4215
         End
         Begin VB.Label lblOpposition 
            Caption         =   "Opposition"
            Height          =   255
            Left            =   50
            TabIndex        =   25
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Label lblNuméro 
            Caption         =   "Identification"
            Height          =   255
            Left            =   50
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblIntitulé 
            Caption         =   "Intitulé"
            Height          =   255
            Left            =   50
            TabIndex        =   14
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblAlpha 
            Caption         =   "Recherche Alpha"
            Height          =   255
            Left            =   50
            TabIndex        =   13
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblTypeBanqueClient 
            Caption         =   "Type"
            Height          =   255
            Left            =   45
            TabIndex        =   12
            Top             =   1425
            Width           =   1215
         End
         Begin VB.Label lblActionnaire 
            Caption         =   "Actionnaire"
            Height          =   255
            Left            =   50
            TabIndex        =   11
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lblMembreduPersonnel 
            Caption         =   "Membre du Personnel"
            Height          =   255
            Left            =   30
            TabIndex        =   10
            Top             =   2865
            Width           =   1695
         End
         Begin VB.Label lblNatureTitulaire 
            Caption         =   "Nature Titulaire"
            Height          =   255
            Left            =   50
            TabIndex        =   9
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label lblNomination 
            Caption         =   "Nomination"
            Height          =   255
            Left            =   50
            TabIndex        =   8
            Top             =   3240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtSrvAut2 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -71640
         TabIndex        =   3
         Top             =   1800
         Width           =   495
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
         Height          =   17535
         Left            =   -76320
         TabIndex        =   2
         Top             =   6240
         Width           =   9975
      End
      Begin VB.Label lblCosOc 
         Caption         =   "Code Societé"
         Height          =   225
         Left            =   50
         TabIndex        =   4
         Top             =   -272
         Width           =   1935
      End
   End
   Begin VB.Label libTitulaireNb 
      Caption         =   "/"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4380
      TabIndex        =   149
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblNbTitu 
      Caption         =   "Numéro Titulaire"
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
      Left            =   2160
      TabIndex        =   18
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblRacine 
      Caption         =   "Racine "
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
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblErrMsg 
      Caption         =   "lblErrMsg"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   7320
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   800
   End
End
Attribute VB_Name = "frmRacine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean

Private recRacine As typeRacine
Private recTitulaire As typeTitulaire
Private Mt As Currency





Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim Msg As String

Msg = Format$(0, "000000") & Format$(0, "000000")

prtRacineX Msg

End Sub



'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub



'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
If txtTitulaireNo.Enabled Then txtTitulaireNo.SetFocus

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
KeyPreview = True
Set XForm = Me
Call MeInit(arrTagNb, "Locked")
ReDim arrTag(arrTagNb + 1)

txtTitulaireNo.Locked = False
txtTitulaireNo.Enabled = True

dspRacineInit

End Sub



'---------------------------------------------------------
Public Sub dspRacine()
'---------------------------------------------------------

txtNuméro = recRacine.Numéro
txtIntitulé = recRacine.Intitulé
txtAlpha = recRacine.Alpha
txtRacAdr01 = recRacine.Adresse1
txtRacAdr02 = recRacine.Adresse2
txtRacAdr03 = recRacine.Adresse3
txtRacAdrCp = recRacine.AdresseCP

txtTéléphone2 = recRacine.Téléphone2
txtTéléphone1 = recRacine.Téléphone1
txtTélex = recRacine.Télex
txtTypeBanqueClient = recRacine.TypeBanqueClient
libTypeBanqueClient = DicLib(63, recRacine.TypeBanqueClient)
Select Case recRacine.TypeBanqueClient
  Case "B": libTypeBanqueClient = "Banque"
  Case "C": libTypeBanqueClient = "Client"
End Select

txtActionnaire = recRacine.Actionnaire
libActionnaire = DicLib(7, recRacine.Actionnaire)

txtNatureTitulaire = recRacine.NatureTitulaire
libNatureTitulaire = DicLib(62, recRacine.NatureTitulaire)

txtRésident = recRacine.Résident
libRésident = DicLib(23, recRacine.Résident)

txtRésidentFiscal = recRacine.RésidentFiscal
libRésidentFiscal = DicLib(7, recRacine.RésidentFiscal)

txtRésidentPays = recRacine.RésidentPays
libRésidentPays = DicLib(19, recRacine.RésidentPays)

txtMembreduPersonnel = recRacine.MembreduPersonnel
libMembreduPersonnel = DicLib(7, recRacine.MembreduPersonnel)

txtTypeOuverture = recRacine.TypeOuverture
libTypeOuverture = DicLib(63, recRacine.TypeOuverture)

txtSuccession = recRacine.Succession
libSuccession = DicLib(7, recRacine.Succession)

txtNomination = recRacine.Nomination
libNomination = DicLib(11, recRacine.Nomination)

txtRégimeMatrimonial = recRacine.RégimeMatrimonial
libRégimeMatrimonial = DicLib(70, recRacine.RégimeMatrimonial)

txtNaissanceAmj = dateImp(recRacine.NaissanceAmj)

txtOptionTva = recRacine.optionTva
libOptionTva = DicLib(7, recRacine.optionTva)

txtGroupeIdentification = recRacine.GroupeIdentification
libGroupeIdentification = DicLib(65, recRacine.GroupeIdentification)

libApporteur = DicLib(58, recRacine.Apporteur)
txtApporteur = recRacine.Apporteur

libGestionnaire = DicLib(60, recRacine.Gestionnaire)
txtGestionnaire = recRacine.Gestionnaire

libOpposition = DicLib(16, recRacine.Opposition)
txtOpposition = recRacine.Opposition

txtNombreTitulaire = recRacine.NombreTitulaire
libGroupeEconomique = DicLib(64, recRacine.GroupeEconomique)

txtGroupeEconomique = recRacine.GroupeEconomique
txtCommentaire = recRacine.Commentaire
libCourrier = DicLib(61, recRacine.Courrier)
txtCourrier = recRacine.Courrier

txtChèquierInterdit = recRacine.ChèquierInterdit
libChèquierInterdit = DicLib(7, recRacine.ChèquierInterdit)

txtChèquierInterditAmj = dateImp(recRacine.ChèquierInterditAmj)

txtFax = recRacine.Fax
txtSwift = recRacine.Swift

libRacNationalité = DicLib(19, recRacine.Nationalité)
txtRacNationalité = recRacine.Nationalité

libAPE = DicLib(65, recRacine.APE)
txtAPE = recRacine.APE
txtSiren = recRacine.SIREN
txtBdfIdentification = recRacine.BdfIdentification

txtSiègeRacine = recRacine.SiègeRacine

txtAmjCréation = dateImp(recRacine.AmjCréation)

txtAmjModification = dateImp(recRacine.AmjModification)

txtAmjAnnulation = dateImp(recRacine.AmjAnnulation)

txtAmjRéactivation = dateImp(recRacine.AmjRéactivation)

txtOpérateur = recRacine.Opérateur

End Sub


'---------------------------------------------------------
Public Sub dspRacineInit()
'---------------------------------------------------------

recRacine = arrRacine(0)
txtRacine = recRacine.Numéro
txtTitulaireNo = IIf(recRacine.NombreTitulaire = 0, 0, 1)
libTitulaireNb = "/ " & recRacine.NombreTitulaire

dspTitulaireInit

dspRacine

End Sub




'---------------------------------------------------------
Public Sub dspTitulaire()
'---------------------------------------------------------

txtNomPatronymique = recTitulaire.NomPatronymique

TxtNationalité = recTitulaire.Nationalité
libNationalité = DicLib(19, recTitulaire.Nationalité)

txtRésidencePays = recTitulaire.RésidencePays
libRésidencePays = DicLib(19, recTitulaire.RésidencePays)

txtRésidenceZone = recTitulaire.RésidenceZone
libRésidenceZone = DicLib(71, recTitulaire.RésidenceZone)
txtPrénom = recTitulaire.Prénom
txtMariNom = recTitulaire.MariNom
txtMariPrénom = recTitulaire.MariPrénom
txtAdresse1 = recTitulaire.Adresse1
txtAdresse2 = recTitulaire.Adresse2
txtAdresse3 = recTitulaire.Adresse3
txtAdresse4 = recTitulaire.Adresse4
txtAdresseCP = recTitulaire.AdresseCP
txtAdresseBD = recTitulaire.AdresseBD

txtNaissanceAmj = dateImp(recTitulaire.NaissanceAmj)

TXTNaissanceDépN = recTitulaire.NaissanceDépN
libNaissanceDépN = DicLib(51, recTitulaire.NaissanceDépN)

txtNaissanceDépX = recTitulaire.NaissanceDépX

txtNaissanceCommune = recTitulaire.NaissanceCommune
txtNaissancePays = recTitulaire.NaissancePays
libNaissancePays = DicLib(19, recTitulaire.NaissancePays)
txtNaissanceZone = recTitulaire.NaissanceZone
libNaissanceZone = DicLib(71, recTitulaire.NaissanceZone)
txtDocIdentitéType = recTitulaire.DocIdentitéType
libDocIdentitéType = DicLib(73, recTitulaire.DocIdentitéType)
txtDocIdentitéDélivréPar = recTitulaire.DocIdentitéDélivréPar
txtDocIdentitéAmj = recTitulaire.DocIdentitéAmj
txtDocIdentitéNo = recTitulaire.DocIdentitéNo

txtPmFormeJuridique = recTitulaire.PmFormeJuridique
libPmFormeJuridique = DicLib(3, recTitulaire.PmFormeJuridique)

txtPmSiren = recTitulaire.PmSiren

txtPmApe = recTitulaire.PmApe
libPmApe = DicLib(5, recTitulaire.PmApe)

txtPmBdfId = recTitulaire.PmBdfId

txtPmNationalitéMère = recTitulaire.PmNationalitéMère
libPmNationalitéMère = DicLib(19, recTitulaire.PmNationalitéMère)

txtPpQualité = recTitulaire.PpQualité
libPpQualité = DicLib(69, recTitulaire.PpQualité)
txtPpRégimeMatrimonial = recTitulaire.PpRégimeMatrimonial
libPpRégimeMatrimonial = DicLib(70, recTitulaire.PpRégimeMatrimonial)
txtPpNombreEnfant = recTitulaire.PpNombreEnfant
txtPpBdfId = recTitulaire.PpBdfId
txtPpProfession = recTitulaire.PpProfession
txtPpCsp = recTitulaire.PpCsp
libPpCsp = DicLib(72, recTitulaire.PpCsp)

txtPpAmjEntrée = dateImp(recTitulaire.PpAmjEntrée)

txtPpAmjSortie = dateImp(recTitulaire.PpAmjSortie)
txtAmjCre = dateImp(recTitulaire.AmjCre)
txtAmjMod = dateImp(recTitulaire.AmjMod)
txtAmjAnn = dateImp(recTitulaire.AmjAnn)
txtAmjRea = dateImp(recTitulaire.AmjRea)
txtNomOp = recTitulaire.NOMOP

End Sub


'---------------------------------------------------------
Public Sub dspTitulaireInit()
'---------------------------------------------------------

arrTitulaireNbMax = 1
arrTitulaireNb = 0
ReDim arrTitulaire(1)
Call recTitulaireInit(recTitulaire)
If recRacine.NombreTitulaire > 0 Then
    recTitulaire.Method = "SnapL0"
    recTitulaire.Racine = recRacine.Numéro
    arrTitulaire(0) = recTitulaire
    Call srvTitulaireMon(recTitulaire)
    txtTitulaireNo = 1
    Call txtTitulaireNo_LostFocus
End If


End Sub


Private Sub Label18_Click()

End Sub

'-------------------------------------------------'
Private Sub txtTitulaireNo_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtTitulaireNo)

End Sub


'-------------------------------------------------'
Private Sub txtTitulaireNo_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call elpNum(txtTitulaireNo, KeyAscii, 3, 0)


End Sub


'-------------------------------------------------'
Private Sub txtTitulaireNo_LostFocus()
'-------------------------------------------------'

txtTitulaireNo.BackColor = txtUsr.BackColor
If Val(txtTitulaireNo) > arrTitulaireNb Then txtTitulaireNo = arrTitulaireNb
If Val(txtTitulaireNo) > 0 Then
    recTitulaire = arrTitulaire(Val(txtTitulaireNo))
    dspTitulaire
End If

End Sub




