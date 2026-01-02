VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_TVAFAC 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_TVA : facturation"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "BIA_TVAFAC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   16325
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Factures"
      TabPicture(0)   =   "BIA_TVAFAC.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Commissions"
      TabPicture(1)   =   "BIA_TVAFAC.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDétail"
      Tab(1).Control(1)=   "lstW"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "N° TVA Intracommunautaire "
      TabPicture(2)   =   "BIA_TVAFAC.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraNIF"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Paramétrage"
      TabPicture(3)   =   "BIA_TVAFAC.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraParam"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraParam 
         Height          =   8445
         Left            =   -74880
         TabIndex        =   164
         Top             =   360
         Width           =   13545
         Begin VB.Frame fraParam_Update 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7095
            Left            =   8280
            TabIndex        =   168
            Top             =   1320
            Visible         =   0   'False
            Width           =   5175
            Begin VB.CommandButton cmdParam_Update_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   182
               Top             =   6360
               Width           =   1455
            End
            Begin VB.Frame fraParam_Update_B 
               BackColor       =   &H00D0D0D0&
               Height          =   3135
               Left            =   120
               TabIndex        =   171
               Top             =   3840
               Width           =   4935
               Begin VB.Frame fraParam_Update_B1 
                  BackColor       =   &H00D0D0D0&
                  Height          =   1095
                  Left            =   240
                  TabIndex        =   186
                  Top             =   720
                  Width           =   4455
                  Begin VB.TextBox txtParamUpdate_CLIENACLI 
                     Height          =   285
                     Left            =   360
                     TabIndex        =   188
                     Top             =   600
                     Width           =   1095
                  End
                  Begin VB.TextBox txtParamUpdate_CLIENARES 
                     Height          =   285
                     Left            =   2760
                     TabIndex        =   187
                     Top             =   600
                     Width           =   1095
                  End
                  Begin VB.Label lblParamUpdate_CLIENACLI 
                     BackColor       =   &H00D0D0D0&
                     Caption         =   "Racine client"
                     Height          =   255
                     Left            =   360
                     TabIndex        =   190
                     Top             =   360
                     Width           =   1095
                  End
                  Begin VB.Label lblParamUpdate_CLIENARES 
                     BackColor       =   &H00D0D0D0&
                     Caption         =   "code routage"
                     Height          =   255
                     Left            =   2760
                     TabIndex        =   189
                     Top             =   360
                     Width           =   1095
                  End
               End
               Begin VB.CheckBox chkParamUpdate_Insert 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Création d'un enregistrement"
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   184
                  Top             =   1920
                  Width           =   2415
               End
               Begin VB.Frame fraParam_Update_B2 
                  BackColor       =   &H00D0D0D0&
                  Height          =   1455
                  Left            =   240
                  TabIndex        =   175
                  Top             =   360
                  Width           =   4335
                  Begin VB.Frame fraParam_Update_TVACOMSTA 
                     BackColor       =   &H00D0D0D0&
                     Caption         =   "état"
                     Height          =   1095
                     Left            =   3000
                     TabIndex        =   179
                     Top             =   240
                     Width           =   1215
                     Begin VB.OptionButton optParam_Update_TVACOMSTA_V 
                        BackColor       =   &H00D0D0D0&
                        Caption         =   "Valider"
                        Height          =   195
                        Left            =   120
                        TabIndex        =   181
                        Top             =   720
                        Width           =   855
                     End
                     Begin VB.OptionButton optParam_Update_TVACOMSTA_I 
                        BackColor       =   &H00D0D0D0&
                        Caption         =   "Ignorer"
                        Height          =   195
                        Left            =   120
                        TabIndex        =   180
                        Top             =   360
                        Value           =   -1  'True
                        Width           =   855
                     End
                  End
                  Begin VB.Frame fraParam_Update_TVACOMOPE 
                     BackColor       =   &H00D0D0D0&
                     Caption         =   "Opération"
                     Height          =   1095
                     Left            =   120
                     TabIndex        =   176
                     Top             =   240
                     Width           =   2895
                     Begin VB.TextBox txtParam_Update_TVACOMOPE 
                        Height          =   285
                        Left            =   1320
                        TabIndex        =   183
                        Top             =   480
                        Width           =   1095
                     End
                     Begin VB.OptionButton optParam_Update_TVACOMOPE_CRE 
                        BackColor       =   &H00D0D0D0&
                        Caption         =   "CRE"
                        Height          =   195
                        Left            =   120
                        TabIndex        =   178
                        Top             =   360
                        Value           =   -1  'True
                        Width           =   855
                     End
                     Begin VB.OptionButton optParam_Update_TVACOMOPE_ENG 
                        BackColor       =   &H00D0D0D0&
                        Caption         =   "ENG"
                        Height          =   195
                        Left            =   120
                        TabIndex        =   177
                        Top             =   720
                        Width           =   855
                     End
                  End
               End
               Begin VB.CommandButton cmdParam_Update_Annuler 
                  BackColor       =   &H00FF00FF&
                  Caption         =   "Supprimer"
                  Height          =   525
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   173
                  Top             =   2520
                  Width           =   1455
               End
               Begin VB.CommandButton cmdParam_Update_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   525
                  Left            =   3360
                  Style           =   1  'Graphical
                  TabIndex        =   172
                  Top             =   2520
                  Width           =   1455
               End
            End
            Begin VB.Frame fraParam_Update_A 
               BackColor       =   &H00E0FFFF&
               ForeColor       =   &H00000000&
               Height          =   3015
               Left            =   120
               TabIndex        =   170
               Top             =   120
               Width           =   4935
               Begin VB.Label Label1 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "!! fichier SAB073SPE/YBIATAB0 non journalisé"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   360
                  TabIndex        =   185
                  Top             =   2040
                  Width           =   4095
               End
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   169
               Top             =   6360
               Width           =   1455
            End
         End
         Begin VB.ComboBox cboParam_SQL 
            Height          =   315
            Left            =   9960
            Sorted          =   -1  'True
            TabIndex        =   167
            Top             =   240
            Width           =   3015
         End
         Begin VB.CommandButton cmdParam_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   10560
            Style           =   1  'Graphical
            TabIndex        =   166
            Top             =   600
            Width           =   1815
         End
         Begin VB.Frame fraParam_Options_1 
            BackColor       =   &H00E0FFFF&
            Height          =   1125
            Left            =   120
            TabIndex        =   165
            Top             =   120
            Width           =   8355
         End
         Begin MSFlexGridLib.MSFlexGrid fgParam 
            Height          =   7065
            Left            =   0
            TabIndex        =   174
            Top             =   1200
            Visible         =   0   'False
            Width           =   8160
            _ExtentX        =   14393
            _ExtentY        =   12462
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "<Id           |<K1               |<K2                 |<Text                                                             ||"
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
      Begin VB.Frame fraDétail 
         Height          =   8655
         Left            =   -74880
         TabIndex        =   79
         Top             =   360
         Width           =   13455
         Begin VB.Frame fraDétail_Update 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7215
            Left            =   2880
            TabIndex        =   92
            Top             =   1440
            Visible         =   0   'False
            Width           =   10335
            Begin VB.CommandButton cmdDétail_Update_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   9000
               Style           =   1  'Graphical
               TabIndex        =   155
               Top             =   5880
               Width           =   1095
            End
            Begin VB.Frame fraDétail_Update_A 
               BackColor       =   &H00E0FFFF&
               Height          =   3615
               Left            =   120
               TabIndex        =   113
               Top             =   120
               Width           =   10095
               Begin VB.TextBox txtUpdate_TVACOMQTE 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   8280
                  TabIndex        =   162
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.ComboBox txtUpdate_TVACOMSTA 
                  Height          =   315
                  Left            =   1320
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   132
                  Top             =   3120
                  Width           =   2655
               End
               Begin VB.TextBox txtUpdate_TVACOMUSR 
                  Height          =   285
                  Left            =   8400
                  TabIndex        =   131
                  Top             =   3120
                  Width           =   1335
               End
               Begin VB.TextBox txtUpdate_TVACOMMTVE 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4560
                  TabIndex        =   130
                  Top             =   2040
                  Width           =   1695
               End
               Begin VB.ComboBox txtUpdate_TVACOMCOMB 
                  Height          =   315
                  Left            =   4560
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   129
                  Top             =   1320
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVACOMCPT 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   128
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVACOMMTVA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   127
                  Top             =   2040
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVACOMX 
                  Height          =   285
                  Left            =   4560
                  TabIndex        =   126
                  Top             =   2760
                  Width           =   5175
               End
               Begin VB.TextBox txtUpdate_TVACOMDOS 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   125
                  Top             =   1320
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVACOMEVE 
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   124
                  Top             =   960
                  Width           =   495
               End
               Begin VB.TextBox txtUpdate_TVACOMNAT 
                  Height          =   285
                  Left            =   4560
                  TabIndex        =   123
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.TextBox txtUpdate_TVACOMECR 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   8280
                  TabIndex        =   122
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtUpdate_TVACOMPLA 
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   121
                  Top             =   240
                  Width           =   495
               End
               Begin VB.TextBox txtUpdate_TVACOMPIE 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4560
                  TabIndex        =   120
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox txtUpdate_TVACOMMONE 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4560
                  TabIndex        =   119
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVACOMFACN 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   118
                  Top             =   2400
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVACOMMON 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   117
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVACOMDEV 
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   116
                  Top             =   1680
                  Width           =   495
               End
               Begin VB.TextBox txtUpdate_TVACOMOPE 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   115
                  Top             =   960
                  Width           =   735
               End
               Begin VB.TextBox txtUpdate_TVACOMETA 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   114
                  Top             =   240
                  Width           =   855
               End
               Begin MSComCtl2.DTPicker txtUpdate_TVACOMDVA 
                  Height          =   300
                  Left            =   4560
                  TabIndex        =   133
                  Top             =   600
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   16711680
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   76349443
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin MSComCtl2.DTPicker txtUpdate_TVACOMDTR 
                  Height          =   300
                  Left            =   8280
                  TabIndex        =   134
                  Top             =   600
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
                  Format          =   76349443
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label libUpdate_TVACOMECR 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Height          =   1335
                  Left            =   6600
                  TabIndex        =   154
                  Top             =   1320
                  Width           =   3255
               End
               Begin VB.Label lblUpdate_TVACOMSTA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Etat"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   153
                  Top             =   3120
                  Width           =   855
               End
               Begin VB.Label lblUpdate_TVACOMUSR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Utilisateur"
                  Height          =   255
                  Left            =   6720
                  TabIndex        =   152
                  Top             =   3240
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMMONB 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Bén-DO-Sha"
                  Height          =   255
                  Left            =   3240
                  TabIndex        =   151
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMCPT 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Compte"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   150
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMFACN 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "N°Facture"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   149
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMMTVA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Montant TVA"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   148
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMMONE 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "€"
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   147
                  Top             =   1680
                  Width           =   255
               End
               Begin VB.Label lblUpdate_TVACOMMON 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Montant"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   146
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMX 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "NUR-UTI-EVE-SEQ-SPE * ECRX * GTYP - GORD"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   145
                  Top             =   2760
                  Width           =   3615
               End
               Begin VB.Label lblUpdate_TVACOMDOS 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Dossier"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   144
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMQTE 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Quantité"
                  Height          =   255
                  Left            =   6600
                  TabIndex        =   143
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMNAT 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Nature"
                  Height          =   255
                  Left            =   3240
                  TabIndex        =   142
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMOPE 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Opé/Evé"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   141
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMDVA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "D.Opération"
                  Height          =   255
                  Left            =   3240
                  TabIndex        =   140
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMTR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "D.Traitement"
                  Height          =   255
                  Left            =   6600
                  TabIndex        =   139
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMECR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Ecriture"
                  Height          =   255
                  Left            =   6600
                  TabIndex        =   138
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMPIE 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Pièce"
                  Height          =   255
                  Left            =   3240
                  TabIndex        =   137
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMPLA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Plan"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   136
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Label lblUpdate_TVACOMETA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Etablissement"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   135
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.Frame fraDétail_Update_B 
               BackColor       =   &H00D0D0D0&
               Height          =   3255
               Left            =   120
               TabIndex        =   93
               Top             =   3840
               Width           =   10095
               Begin VB.Frame fraUpdate_TVACOMFACL 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Annulation : indiquer le numéro de facture de com d'origine"
                  ForeColor       =   &H000000FF&
                  Height          =   855
                  Left            =   4320
                  TabIndex        =   192
                  Top             =   2280
                  Width           =   4455
                  Begin VB.TextBox txtUpdate_TVACOMFACL 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   1440
                     TabIndex        =   193
                     Top             =   360
                     Width           =   1695
                  End
                  Begin VB.Label lblUpdate_TVACOMFACL 
                     BackColor       =   &H00C0FFFF&
                     Caption         =   "N°FAC liée"
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Left            =   120
                     TabIndex        =   194
                     Top             =   480
                     Width           =   975
                  End
               End
               Begin VB.ComboBox txtUpdate_TVACOMSRVR 
                  Height          =   315
                  Left            =   2520
                  Sorted          =   -1  'True
                  TabIndex        =   160
                  Text            =   "SRVR"
                  Top             =   2880
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVACOMECRX 
                  Height          =   285
                  Left            =   4560
                  TabIndex        =   103
                  Top             =   240
                  Width           =   735
               End
               Begin VB.ComboBox txtUpdate_TVACOMCOMC 
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   5640
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   102
                  Top             =   720
                  Width           =   4335
               End
               Begin VB.ComboBox txtUpdate_TVACOMCOME 
                  Height          =   315
                  Left            =   240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   101
                  Top             =   2880
                  Width           =   1575
               End
               Begin VB.ComboBox txtUpdate_TVACOMTVAC 
                  Height          =   315
                  Left            =   5640
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   100
                  Top             =   1920
                  Width           =   1695
               End
               Begin VB.ComboBox txtUpdate_TVACOMCLIC 
                  Height          =   315
                  Left            =   840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   99
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.ComboBox txtUpdate_TVACOMCLIP 
                  Height          =   315
                  Left            =   5640
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   98
                  Top             =   1560
                  Width           =   2535
               End
               Begin VB.ComboBox txtUpdate_TVACOMCOMT 
                  Height          =   315
                  Left            =   5640
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   97
                  Top             =   1080
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdate_TVACOMCLI 
                  Height          =   285
                  Left            =   2400
                  TabIndex        =   96
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.CommandButton cmdDétail_Update_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   525
                  Left            =   8880
                  Style           =   1  'Graphical
                  TabIndex        =   95
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.CommandButton cmdDétail_Update_Annuler 
                  BackColor       =   &H00FF00FF&
                  Caption         =   "Ann/Reprise"
                  Height          =   525
                  Left            =   8880
                  Style           =   1  'Graphical
                  TabIndex        =   94
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.Label lblUpdate_TVACOMSRVR 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Service "
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   159
                  Top             =   2950
                  Width           =   735
               End
               Begin VB.Label lblUpdate_TVACOMECRX 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "=ECRX"
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   112
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label libUpdate_TVACOMECRX 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Height          =   375
                  Left            =   5640
                  TabIndex        =   111
                  Top             =   240
                  Width           =   4335
               End
               Begin VB.Label libUpdate_TVACOMCLI 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Height          =   2175
                  Left            =   240
                  TabIndex        =   110
                  Top             =   720
                  Width           =   3975
               End
               Begin VB.Label lblUpdate_TVACOMCOME 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "C.Edition"
                  Height          =   255
                  Left            =   7680
                  TabIndex        =   109
                  Top             =   1200
                  Width           =   735
               End
               Begin VB.Label lblUpdate_TVACOMCOMC 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Commission"
                  Height          =   255
                  Left            =   4560
                  TabIndex        =   108
                  Top             =   720
                  Width           =   855
               End
               Begin VB.Label lblUpdate_TVACOMCLIP 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Pays Rés"
                  Height          =   255
                  Left            =   4560
                  TabIndex        =   107
                  Top             =   1560
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMCLIC 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Tiers"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   106
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label lblUpdate_TVACOMCOMT 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Com Taxable"
                  Height          =   255
                  Left            =   4560
                  TabIndex        =   105
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVACOMTVAC 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "C. TVA"
                  Height          =   255
                  Left            =   4560
                  TabIndex        =   104
                  Top             =   2040
                  Width           =   735
               End
            End
         End
         Begin VB.ComboBox cboDétail_SQL 
            Height          =   315
            Left            =   10200
            Sorted          =   -1  'True
            TabIndex        =   91
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   3015
         End
         Begin VB.Frame fraDétail_Options_1 
            BackColor       =   &H00E0FFFF&
            Height          =   1125
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   9915
            Begin VB.CheckBox chkDétail_TVACOMSTA 
               BackColor       =   &H00E0FFFF&
               Caption         =   "exclure 'A I F X'"
               Height          =   195
               Left            =   2400
               TabIndex        =   163
               Top             =   240
               Width           =   1695
            End
            Begin VB.ComboBox txtDétail_TVACOMSRVR 
               Height          =   315
               Left            =   6000
               Sorted          =   -1  'True
               TabIndex        =   157
               Text            =   "SRVR"
               Top             =   720
               Width           =   1455
            End
            Begin VB.ComboBox txtDétail_TVACOMCLIC 
               Height          =   315
               Left            =   8160
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   240
               Width           =   1455
            End
            Begin VB.ComboBox txtDétail_TVACOMOPE 
               Height          =   315
               Left            =   4200
               TabIndex        =   85
               Text            =   "opé"
               Top             =   720
               Width           =   1455
            End
            Begin VB.ComboBox txtDétail_TVACOMSTA 
               Height          =   315
               Left            =   1920
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   720
               Width           =   2175
            End
            Begin VB.TextBox txtDétail_TVACOMCLI 
               Height          =   285
               Left            =   8160
               TabIndex        =   83
               Top             =   720
               Width           =   1455
            End
            Begin VB.CheckBox chkDétail_TVACOMDTR 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Période de création"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   120
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker txtDétail_TVACOMDTR 
               Height          =   300
               Left            =   480
               TabIndex        =   87
               Top             =   360
               Width           =   1332
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   76349443
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtDétail_TVACOMDTR_Max 
               Height          =   300
               Left            =   480
               TabIndex        =   88
               Top             =   720
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
               Format          =   76349443
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblDétail_TVACOMSRVR 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Service "
               Height          =   255
               Left            =   6240
               TabIndex        =   158
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblDétail_TVACOMOPE 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Opé"
               Height          =   255
               Left            =   4560
               TabIndex        =   90
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblDétail_TVACOMSTA 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Etat"
               Height          =   255
               Left            =   1920
               TabIndex        =   89
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.CommandButton cmdDétail_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   11040
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   720
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid fgDétail 
            Height          =   7185
            Left            =   0
            TabIndex        =   156
            Top             =   1440
            Visible         =   0   'False
            Width           =   13440
            _ExtentX        =   23707
            _ExtentY        =   12674
            _Version        =   393216
            Rows            =   1
            Cols            =   17
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"BIA_TVAFAC.frx":007C
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
      Begin VB.Frame fraNIF 
         Height          =   8445
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   13545
         Begin VB.Frame fraNIF_Options_1 
            BackColor       =   &H00E0FFFF&
            Height          =   1125
            Left            =   120
            TabIndex        =   67
            Top             =   120
            Width           =   8355
            Begin VB.CheckBox chkNIF_TVANIFCLIT 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Uniquement Tiers/NIF à compléter"
               Height          =   255
               Left            =   3000
               TabIndex        =   161
               Top             =   240
               Width           =   3135
            End
            Begin VB.TextBox txtNIF_TVANIFCLI 
               Height          =   285
               Left            =   6000
               TabIndex        =   70
               Top             =   720
               Width           =   2055
            End
            Begin VB.ComboBox txtNIF_TVANIFSTA 
               Height          =   315
               Left            =   480
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   720
               Width           =   2295
            End
            Begin VB.ComboBox txtNIF_TVANIFCLIC 
               Height          =   315
               Left            =   6240
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblNIF_TVANIFCLI 
               BackColor       =   &H00E0FFFF&
               Caption         =   "N° du Tiers ou nom partiel"
               Height          =   255
               Left            =   3960
               TabIndex        =   77
               Top             =   720
               Width           =   1935
            End
            Begin VB.Label lblNIF_TVANIFSTA 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Etat"
               Height          =   255
               Left            =   1320
               TabIndex        =   71
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdNIF_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   10560
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cboNIF_SQL 
            Height          =   315
            Left            =   9960
            Sorted          =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   3015
         End
         Begin VB.Frame fraNIF_Update 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7095
            Left            =   8280
            TabIndex        =   50
            Top             =   1320
            Visible         =   0   'False
            Width           =   5175
            Begin VB.CommandButton cmdNIF_Update_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   6360
               Width           =   1455
            End
            Begin VB.Frame fraNIF_Update_A 
               BackColor       =   &H00E0FFFF&
               Height          =   4575
               Left            =   120
               TabIndex        =   55
               Top             =   120
               Width           =   4935
               Begin VB.ComboBox txtUpdate_TVANIFCLIC 
                  Height          =   315
                  Left            =   960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   59
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdate_TVANIFCLI 
                  Height          =   285
                  Left            =   2760
                  TabIndex        =   58
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.ComboBox txtUpdate_TVANIFSTA 
                  Height          =   315
                  Left            =   1440
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   57
                  Top             =   3720
                  Width           =   2655
               End
               Begin VB.TextBox txtUpdate_TVANIFUSR 
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   56
                  Top             =   4080
                  Width           =   1335
               End
               Begin VB.Label libUpdate_TVANIFCLI 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Height          =   2775
                  Left            =   240
                  TabIndex        =   63
                  Top             =   720
                  Width           =   4455
               End
               Begin VB.Label lblUpdate_TVANIFCLIC 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Tiers"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   62
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label lblUpdate_TVANIFSTA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Etat"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   61
                  Top             =   3720
                  Width           =   615
               End
               Begin VB.Label lblUpdate_TVANIFUSR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Utilisateur"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   60
                  Top             =   4080
                  Width           =   855
               End
            End
            Begin VB.Frame fraNIF_Update_B 
               BackColor       =   &H00D0D0D0&
               Height          =   2295
               Left            =   120
               TabIndex        =   51
               Top             =   4680
               Width           =   4935
               Begin VB.OptionButton optUpdate_TVANIFCLIF_Idem 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Idem Tiers"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   75
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.OptionButton optUpdate_TVANIFCLIF 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "N I F"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   74
                  Top             =   960
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.CommandButton cmdNIF_Update_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   525
                  Left            =   3360
                  Style           =   1  'Graphical
                  TabIndex        =   54
                  Top             =   1680
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNIF_Update_Annuler 
                  BackColor       =   &H00FF00FF&
                  Caption         =   "Supprimer"
                  Height          =   525
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   53
                  Top             =   1680
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdate_TVANIFCLIT 
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   52
                  Top             =   1080
                  Width           =   3255
               End
               Begin VB.Label libUpdate_TVANIFCLIT_URL 
                  Alignment       =   2  'Center
                  Caption         =   " http://ec.europa.eu/taxation_customs/vies/fr/vieshome.htm"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   78
                  Top             =   600
                  Width           =   4695
               End
               Begin VB.Label libUpdate_TVANIFCLIT 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Le site suivant permet de vérifier la validité du n° de TVA :"
                  ForeColor       =   &H00FF00FF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   76
                  Top             =   240
                  Width           =   4575
               End
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgNIF 
            Height          =   7065
            Left            =   0
            TabIndex        =   64
            Top             =   1200
            Visible         =   0   'False
            Width           =   8160
            _ExtentX        =   14393
            _ExtentY        =   12462
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "<Tiers    |<Intitulé                        |<Pays|<Tva intracommunautaire|<Etat | ||"
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
      Begin VB.ListBox lstW 
         Height          =   255
         Left            =   -67800
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5310
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   73
            Top             =   1680
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Frame fraSelect_Update 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7095
            Left            =   4200
            TabIndex        =   15
            Top             =   1320
            Visible         =   0   'False
            Width           =   9015
            Begin VB.CommandButton cmdSelect_Update_Détail_Display 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Afficher le détail des commissions"
               Height          =   525
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   6240
               Width           =   1575
            End
            Begin VB.CommandButton cmdSelect_Update_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   5280
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   6240
               Width           =   1575
            End
            Begin VB.Frame fraSelect_Update_B 
               BackColor       =   &H00D0D0D0&
               Height          =   2535
               Left            =   120
               TabIndex        =   17
               Top             =   4440
               Width           =   8775
               Begin VB.TextBox txtUpdate_TVAFACCLIT 
                  Height          =   285
                  Left            =   3960
                  TabIndex        =   40
                  Top             =   720
                  Width           =   3375
               End
               Begin VB.CommandButton cmdSelect_Update_Annuler 
                  BackColor       =   &H000000FF&
                  Caption         =   "Annuler définitivement"
                  Height          =   525
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   19
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.CommandButton cmdSelect_Update_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   525
                  Left            =   6840
                  Style           =   1  'Graphical
                  TabIndex        =   18
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.Label lblUpdate_TVAFACLIT 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Code TVA intracommunautaire"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   41
                  Top             =   840
                  Width           =   2415
               End
            End
            Begin VB.Frame fraSelect_Update_A 
               BackColor       =   &H00E0FFFF&
               Height          =   4215
               Left            =   120
               TabIndex        =   16
               Top             =   120
               Width           =   8775
               Begin VB.TextBox txtUpdate_TVAFACETA 
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   42
                  Top             =   240
                  Width           =   495
               End
               Begin VB.ComboBox txtUpdate_TVAFACCLIP 
                  Height          =   315
                  Left            =   5880
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   600
                  Width           =   2535
               End
               Begin VB.TextBox txtUpdate_TVAFACMEXO 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   34
                  Top             =   2040
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVAFACMTVA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   32
                  Top             =   1560
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVAFACMTTC 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   29
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVAFACFACN 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   28
                  Top             =   2520
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_TVAFACUSR 
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   27
                  Top             =   3360
                  Width           =   1335
               End
               Begin VB.ComboBox txtUpdate_TVAFACSTA 
                  Height          =   315
                  Left            =   5880
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   24
                  Top             =   3720
                  Width           =   2655
               End
               Begin VB.TextBox txtUpdate_TVAFACCLI 
                  Height          =   285
                  Left            =   2760
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.ComboBox txtUpdate_TVAFACCLIC 
                  Height          =   315
                  Left            =   960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   240
                  Width           =   1455
               End
               Begin MSComCtl2.DTPicker txtUpdate_TVAFACDTR 
                  Height          =   300
                  Left            =   5880
                  TabIndex        =   36
                  Top             =   2880
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
                  Format          =   76349443
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblUpdate_TVAFACETA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Etablissement"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   43
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVAFACCLIP 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Pays Rés"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   39
                  Top             =   600
                  Width           =   735
               End
               Begin VB.Label lblUpdate_TVAFACDTR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "D.Traitement"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   37
                  Top             =   2880
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVAFACMEXO 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Mt Exo"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   35
                  Top             =   2040
                  Width           =   735
               End
               Begin VB.Label lblUpdate_TVAFACMTVA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Mt TVA"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   33
                  Top             =   1560
                  Width           =   735
               End
               Begin VB.Label lblUpdate_TVAFACMTTC 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Mt TTC"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   31
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.Label lbtUpdate_TVAFACFACN 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "N°Facture"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   30
                  Top             =   2520
                  Width           =   855
               End
               Begin VB.Label lblUpdate_TVAFACUSR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Utilisateur"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   26
                  Top             =   3360
                  Width           =   975
               End
               Begin VB.Label lblUpdate_TVAFACSTA 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Etat"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   25
                  Top             =   3720
                  Width           =   615
               End
               Begin VB.Label lblUpdate_TVAFACCLI 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Tiers"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   23
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label libUpdate_TVAFACCLI 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Height          =   3375
                  Left            =   240
                  TabIndex        =   22
                  Top             =   720
                  Width           =   4455
               End
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7065
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Visible         =   0   'False
            Width           =   13440
            _ExtentX        =   23707
            _ExtentY        =   12462
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"BIA_TVAFAC.frx":0106
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
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   9720
            Sorted          =   -1  'True
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   10560
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   600
            Width           =   1815
         End
         Begin VB.Frame fraSelect_Options_1 
            Height          =   1125
            Left            =   1200
            TabIndex        =   6
            Top             =   120
            Width           =   8115
            Begin VB.CheckBox chkSelect_TVAFACSTA 
               Caption         =   "exclure 'A I F X'"
               Height          =   195
               Left            =   3240
               TabIndex        =   191
               Top             =   240
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.ComboBox txtSelect_TVAFACCLIC 
               Height          =   315
               Left            =   5880
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   240
               Width           =   1455
            End
            Begin VB.ComboBox txtSelect_TVAFACSTA 
               Height          =   315
               Left            =   2280
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   720
               Width           =   2775
            End
            Begin VB.CheckBox chkSelect_TVAFACDTR 
               Caption         =   "Période de création"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   1815
            End
            Begin VB.TextBox txtSelect_TVAFACCLI 
               Height          =   285
               Left            =   5880
               TabIndex        =   11
               Top             =   720
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker txtSelect_TVAFACDTR 
               Height          =   300
               Left            =   480
               TabIndex        =   10
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
               Format          =   76349443
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_TVAFACDTR_Max 
               Height          =   300
               Left            =   480
               TabIndex        =   14
               Top             =   720
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
               Format          =   76349443
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_TVAFACSTA 
               Caption         =   "Etat"
               Height          =   255
               Left            =   2400
               TabIndex        =   47
               Top             =   240
               Width           =   615
            End
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "BIA_TVAFAC.frx":0193
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContext_x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint1 
      Caption         =   "mnuPrint1"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint1_Liste1 
         Caption         =   "Imprimer la liste des commissions"
      End
      Begin VB.Menu mnuPrint1_Liste2 
         Caption         =   "Imprimer la liste des commissions + libellé comptable"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_Facture 
         Caption         =   "Imprimer la facture"
      End
      Begin VB.Menu mnuPrint0_Liste1 
         Caption         =   "Imprimer la liste des factures"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_Liste1 
         Caption         =   "NIF : Imprimer la liste des tiers"
      End
   End
End
Attribute VB_Name = "frmBIA_TVAFAC"
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
Dim BIA_TVAFAC_Aut As typeAuthorization
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim wAmjMin7 As Long, wAmjMax7 As Long
Dim rsSabX As New ADODB.Recordset


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean
Dim cmdSelect_Ok_Caption As String
Dim cmdSelect_SQL_K As String
Dim xYTVAFAC0 As typeYTVAFAC0, meYTVAFAC0 As typeYTVAFAC0
Dim newYTVAFAC0 As typeYTVAFAC0, oldYTVAFAC0 As typeYTVAFAC0
Dim arrYTVAFAC0() As typeYTVAFAC0, arrYTVAFAC0_Nb As Long, arrYTVAFAC0_Max As Long, arrYTVAFAC0_Index As Long
Dim selYTVAFAC0() As typeYTVAFAC0, selYTVAFAC0_Nb As Long, selYTVAFAC0_Max As Long, selYTVAFAC0_Index As Long

'______________________________________________________________________

Dim fgDétail_FormatString As String, fgDétail_K As Integer
Dim fgDétail_RowDisplay As Integer, fgDétail_RowClick As Integer, fgDétail_ColClick As Integer
Dim fgDétail_ColorClick As Long, fgDétail_ColorDisplay As Long
Dim fgDétail_Sort1 As Integer, fgDétail_Sort2 As Integer
Dim fgDétail_SortAD As Integer, fgDétail_Sort1_Old As Integer
Dim fgDétail_arrIndex As Integer
Dim blnfgDétail_DisplayLine As Boolean
Dim cmdDétail_Ok_Caption As String
Dim cmdDétail_SQL_K As String
Dim xYTVACOM0 As typeYTVACOM0, meYTVACOM0 As typeYTVACOM0
Dim newYTVACOM0 As typeYTVACOM0, oldYTVACOM0 As typeYTVACOM0
Dim arrYTVACOM0() As typeYTVACOM0, arrYTVACOM0_Nb As Long, arrYTVACOM0_Max As Long, arrYTVACOM0_Index As Long

'______________________________________________________________________

Dim fgNIF_FormatString As String, fgNIF_K As Integer
Dim fgNIF_RowDisplay As Integer, fgNIF_RowClick As Integer, fgNIF_ColClick As Integer
Dim fgNIF_ColorClick As Long, fgNIF_ColorDisplay As Long
Dim fgNIF_Sort1 As Integer, fgNIF_Sort2 As Integer
Dim fgNIF_SortAD As Integer, fgNIF_Sort1_Old As Integer
Dim fgNIF_arrIndex As Integer
Dim blnfgNIF_DisplayLine As Boolean
Dim cmdNIF_Ok_Caption As String
Dim cmdNIF_SQL_K As String
Dim xYTVANIF0 As typeYTVANIF0, meYTVANIF0 As typeYTVANIF0
Dim newYTVANIF0 As typeYTVANIF0, oldYTVANIF0 As typeYTVANIF0
Dim arrYTVANIF0() As typeYTVANIF0, arrYTVANIF0_Nb As Long, arrYTVANIF0_Max As Long, arrYTVANIF0_Index As Long
'______________________________________________________________________

Dim BIA_TVAPARAM_Aut As typeAuthorization
Dim fgParam_FormatString As String, fgParam_K As Integer
Dim fgParam_RowDisplay As Integer, fgParam_RowClick As Integer, fgParam_ColClick As Integer
Dim fgParam_ColorClick As Long, fgParam_ColorDisplay As Long
Dim fgParam_Sort1 As Integer, fgParam_Sort2 As Integer
Dim fgParam_SortAD As Integer, fgParam_Sort1_Old As Integer
Dim fgParam_arrIndex As Integer
Dim blnfgParam_DisplayLine As Boolean
Dim cmdParam_Ok_Caption As String
Dim cmdParam_SQL_K As String
Dim xYBIATAB0 As typeYBIATAB0, meYBIATAB0 As typeYBIATAB0
Dim newYBIATAB0 As typeYBIATAB0, oldYBIATAB0 As typeYBIATAB0
Dim arrYBIATAB0() As typeYBIATAB0, arrYBIATAB0_Nb As Long, arrYBIATAB0_Max As Long, arrYBIATAB0_Index As Long
'______________________________________________________________________

Dim meCV1 As typeCV, meCV2 As typeCV
Dim xZCLIENA0 As typeZCLIENA0
Dim meYBIAMVT0 As typeYBIAMVT0
Dim curDB As Currency, curCR As Currency

Dim xZADRESS0 As typeZADRESS0, xZCDOTIE0 As typeZCDOTIE0

Dim arrCommission() As String, arrCommission_K As Integer, arrCommission_Max As Integer

Dim mTVAFACCLIP As String, mTVAFACCLIP_NIF As Boolean, mTVAFACCLIP_Code_Fiscal As String

Dim mnuPrint1_Liste As String
'__________________________________________________________________________

Dim BIA_TVACOM_Aut As typeAuthorization
Dim BIA_TVANIF_Aut As typeAuthorization
Dim BIA_TVASRVR_Aut As typeAuthorization

Dim paramFacturation_Path As String, paramFacturation_Path_AAAA As String
'______________________________________________________________________
Dim arrPays() As typePays, arrPays_NB As Integer
Dim mTVANIFCLIT_Pays As Boolean
'______________________________________________________________________
Dim arrStat_Db1(500) As typeYTVACOM0, arrStat_Db2(500) As typeYTVACOM0, arrStat_Db3(500) As typeYTVACOM0
Dim arrStat_Cr1(500) As typeYTVACOM0, arrStat_Cr2(500) As typeYTVACOM0, arrStat_Cr3(500) As typeYTVACOM0
Dim arrStat_K As Integer

Dim blnPrint_Reprise_Ok As Boolean, Print_Reprise_TVAFACFACN As Long

Dim mTVAFACMUE_EXO As Currency, mTVAFACMUE_EXO_Total As Currency, mTVAFACMUE_CLIT_Total As Currency
Dim mTVAFACMUE_HT As Currency, mTVAFACMUE_HT_Total As Currency
Dim mTVAFACMUE_TVA As Currency, mTVAFACMUE_TVA_Total As Currency

Dim mTVA_DES_File As String
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
        For I = fgSelect_arrIndex To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.LeftCol = 0
    End If
End If

End Sub
Public Sub fgDétail_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgDétail.Row

If lRow > 0 And lRow < fgDétail.Rows Then
    fgDétail.Row = lRow
    For I = 0 To fgDétail_arrIndex
        fgDétail.Col = I: fgDétail.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDétail.Row = mRow
    If fgDétail.Row > 0 Then
        lRow = fgDétail.Row
        lColor_Old = fgDétail.CellBackColor
        For I = fgDétail_arrIndex To 0 Step -1
          fgDétail.Col = I: fgDétail.CellBackColor = lColor
        Next I
        fgDétail.LeftCol = 0
    End If
End If

End Sub

Private Sub fgSelect_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To arrYTVAFAC0_Nb

        xYTVAFAC0 = arrYTVAFAC0(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub fgSelect_Display_Stat()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = "Srv  |Opé   |Nat   " _
                      & "| > nb Db (1)|>     Mt Db (1)|> nb Cr (1)|>     Mt Cr (1)" _
                      & "| > nb Db (2)|>     Mt Db (2)|> nb Cr (2)|>     Mt Cr (2)" _
                      & "| > nb Db (3)|>     Mt Db (3)|> nb Cr (3)|>     Mt Cr (3)"
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To arrStat_K
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 0: fgSelect.Text = arrStat_Db1(I).TVACOMSER
    fgSelect.Col = 1: fgSelect.Text = arrStat_Db1(I).TVACOMOPE
    fgSelect.Col = 2: fgSelect.Text = arrStat_Db1(I).TVACOMNAT
    
    fgSelect.Col = 3: fgSelect.Text = arrStat_Db1(I).TVACOMUPDS
    fgSelect.Col = 4: fgSelect.Text = arrStat_Db1(I).TVACOMMONE
    fgSelect.Col = 5: fgSelect.Text = arrStat_Cr1(I).TVACOMUPDS
    fgSelect.Col = 6: fgSelect.Text = arrStat_Cr1(I).TVACOMMONE
    
    fgSelect.Col = 7: fgSelect.Text = arrStat_Db2(I).TVACOMUPDS
    fgSelect.Col = 8: fgSelect.Text = arrStat_Db2(I).TVACOMMONE
    fgSelect.Col = 9: fgSelect.Text = arrStat_Cr2(I).TVACOMUPDS
    fgSelect.Col = 10: fgSelect.Text = arrStat_Cr2(I).TVACOMMONE
    
    fgSelect.Col = 11: fgSelect.Text = arrStat_Db3(I).TVACOMUPDS
    fgSelect.Col = 12: fgSelect.Text = arrStat_Db3(I).TVACOMMONE
    fgSelect.Col = 13: fgSelect.Text = arrStat_Cr3(I).TVACOMUPDS
    fgSelect.Col = 14: fgSelect.Text = arrStat_Cr3(I).TVACOMMONE
        
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgDétail_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 1
fgDétail.Visible = False
fgDétail_Reset
fgDétail.Rows = 1
fgDétail.FormatString = fgDétail_FormatString
cmdPrint.Enabled = False
currentAction = "fgDétail_Display"

For I = 1 To arrYTVACOM0_Nb

     xYTVACOM0 = arrYTVACOM0(I)

        fgDétail.Rows = fgDétail.Rows + 1
        fgDétail.Row = fgDétail.Rows - 1
        fgDétail_DisplayLine I
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgDétail.Rows - 1): DoEvents
If fgDétail.Rows > 1 Then
'    fgDétail_Sort1 = 0: fgDétail_Sort2 = 2: fgDétail_Sort
    cmdPrint.Enabled = True
End If
fgDétail.LeftCol = 0

fgDétail.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub lstSelect_Load_1()
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_TVAFACDTR.Enabled = True
chkSelect_TVAFACDTR = "0"
txtSelect_TVAFACSTA.Enabled = True
txtSelect_TVAFACCLI.Enabled = True
chkSelect_TVAFACSTA = "1"
chkSelect_TVAFACSTA.Visible = True
chkSelect_TVAFACSTA.Enabled = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Private Sub lstDétail_Load_1()

On Error GoTo Error_Handler
fgDétail.Visible = False
cmdPrint.Enabled = False
currentAction = "lstDétail_Load_1"
cmdDétail_Ok_Caption = "Lancer la requête"
cmdDétail_Ok.Caption = cmdDétail_Ok_Caption
cmdDétail_Ok.Visible = True
txtDétail_TVACOMSTA.Enabled = True
fraDétail_Options_1.Visible = True
fraDétail_Options_1.Enabled = True
chkDétail_TVACOMDTR.Enabled = True
chkDétail_TVACOMDTR = "1"
chkDétail_TVACOMDTR.Enabled = False
Call DTPicker_Set(txtDétail_TVACOMDTR, Mid$(YBIATAB0_DATE_CPT_M, 1, 6) & "01")
Call DTPicker_Set(txtDétail_TVACOMDTR_Max, YBIATAB0_DATE_CPT_M)
chkDétail_TVACOMSTA.Value = "1"
chkDétail_TVACOMSTA.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub lstDétail_Load_2()

On Error GoTo Error_Handler
fgDétail.Visible = False
cmdPrint.Enabled = False
currentAction = "lstDétail_Load_2"
cmdDétail_Ok_Caption = "Lancer la requête"
cmdDétail_Ok.Caption = cmdDétail_Ok_Caption
cmdDétail_Ok.Visible = True
txtDétail_TVACOMSTA.Enabled = False
fraDétail_Options_1.Visible = True
fraDétail_Options_1.Enabled = True
chkDétail_TVACOMSTA.Value = "0"
chkDétail_TVACOMSTA.Visible = False
chkDétail_TVACOMDTR = "0"
chkDétail_TVACOMDTR.Enabled = False

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub lstDétail_Load_3()

On Error GoTo Error_Handler
fgDétail.Visible = False
cmdPrint.Enabled = False
currentAction = "lstDétail_Load_3"
cmdDétail_Ok_Caption = "Lancer la requête"
cmdDétail_Ok.Caption = cmdDétail_Ok_Caption
cmdDétail_Ok.Visible = True
txtDétail_TVACOMSTA.ListIndex = 0
txtDétail_TVACOMSTA.Enabled = False
txtDétail_TVACOMSRVR.ListIndex = 0
txtDétail_TVACOMSRVR.Enabled = False
fraDétail_Options_1.Visible = True
fraDétail_Options_1.Enabled = True
chkDétail_TVACOMSTA.Value = "0"
chkDétail_TVACOMSTA.Visible = False
chkDétail_TVACOMDTR = "0"
chkDétail_TVACOMDTR.Enabled = False

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub lstSelect_Load_7()
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_7"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_TVAFACDTR.Enabled = False
chkSelect_TVAFACDTR = "1"
txtSelect_TVAFACDTR.Visible = True
txtSelect_TVAFACDTR_Max.Visible = False
Call DTPicker_Set(txtSelect_TVAFACDTR, YBIATAB0_DATE_CPT_MP1)
txtSelect_TVAFACSTA.Enabled = False
txtSelect_TVAFACCLI.Enabled = True
chkSelect_TVAFACSTA.Visible = False

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub lstSelect_Load_8()
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_8"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_TVAFACDTR.Enabled = False
chkSelect_TVAFACDTR = "1"
txtSelect_TVAFACDTR.Visible = True
txtSelect_TVAFACDTR_Max.Visible = False
Call DTPicker_Set(txtSelect_TVAFACDTR, YBIATAB0_DATE_CPT_MP1)
txtSelect_TVAFACSTA.Enabled = False
txtSelect_TVAFACCLI.Enabled = True
chkSelect_TVAFACSTA.Visible = False

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub lstSelect_Load_S()
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = True
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_S"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_TVAFACDTR.Enabled = False
chkSelect_TVAFACDTR = "1"
txtSelect_TVAFACDTR.Visible = True
txtSelect_TVAFACDTR_Max.Visible = True
Call DTPicker_Set(txtSelect_TVAFACDTR, YBIATAB0_DATE_CAL_AP1)
Call DTPicker_Set(txtSelect_TVAFACDTR, Mid$(YBIATAB0_DATE_CAL_AP1, 1, 4) & "0101")
txtSelect_TVAFACSTA.Enabled = False
txtSelect_TVAFACCLI.Enabled = False
chkSelect_TVAFACSTA.Visible = False

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long

On Error Resume Next
     Select Case xYTVAFAC0.TVAFACSTA
        Case Is = "V": wColor = vbBlue
        Case Is = "F": wColor = &HC000&
        Case Is = "A": wColor = &HC0C0C0
        Case Is = "I": wColor = &HC0C000
        Case Is = "0", "9": wColor = vbMagenta
        Case Else: wColor = vbBlack
    End Select
fgSelect.Col = 0: fgSelect.Text = xYTVAFAC0.TVAFACCLIC & " " & xYTVAFAC0.TVAFACCLI
fgSelect.CellForeColor = wColor
fgSelect.Col = 1: fgSelect.Text = xYTVAFAC0.TVAFACCLIP
fgSelect.CellForeColor = wColor
fgSelect.Col = 2: fgSelect.Text = xYTVAFAC0.TVAFACCLIT
fgSelect.CellForeColor = wColor
fgSelect.Col = 3: fgSelect.Text = xYTVAFAC0.TVAFACSTA
fgSelect.CellForeColor = wColor
fgSelect.Col = 4: fgSelect.Text = Trim(Format$(xYTVAFAC0.TVAFACFACN, "### ### ###"))
fgSelect.CellForeColor = wColor
fgSelect.Col = 5: fgSelect.Text = dateIBM10(xYTVAFAC0.TVAFACDTR, True)
fgSelect.CellForeColor = wColor

X = Trim(Format$(Abs(xYTVAFAC0.TVAFACMTTC), "### ### ### ###.00"))
fgSelect.Col = 6: fgSelect.Text = X
fgSelect.CellForeColor = wColor
X = Trim(Format$(Abs(xYTVAFAC0.TVAFACMTVA), "### ### ### ###.00"))
fgSelect.Col = 7: fgSelect.Text = X
fgSelect.CellForeColor = wColor
X = Trim(Format$(Abs(xYTVAFAC0.TVAFACMEXO), "### ### ### ###.00"))
fgSelect.Col = 8: fgSelect.Text = X
fgSelect.CellForeColor = wColor

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub

Public Sub fgDétail_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long

On Error Resume Next
     Select Case xYTVACOM0.TVACOMSTA
        Case Is = "V": wColor = vbBlue
        Case Is = "F": wColor = &HC000&
        Case Is = "A": wColor = &HC0C0C0
        Case Is = "I": wColor = &HC0C000
        Case Is = "0", "9": wColor = vbMagenta
        Case Else: wColor = vbBlack
    End Select

fgDétail.Col = 0: fgDétail.Text = xYTVACOM0.TVACOMSRVR & " " & xYTVACOM0.TVACOMOPE & " " & xYTVACOM0.TVACOMNAT & " " & xYTVACOM0.TVACOMEVE
fgDétail.CellForeColor = wColor
fgDétail.Col = 1: fgDétail.Text = Trim(Format$(xYTVACOM0.TVACOMDOS, "### ### ###"))
fgDétail.CellForeColor = wColor
X = Trim(Format$(Abs(xYTVACOM0.TVACOMMON), "### ### ### ###.00"))
fgDétail.Col = 2: fgDétail.Text = X
If xYTVACOM0.TVACOMMON > 0 Then
    fgDétail.CellForeColor = vbRed
Else
    fgDétail.CellForeColor = wColor
End If
fgDétail.Col = 3: fgDétail.Text = xYTVACOM0.TVACOMDEV
fgDétail.CellForeColor = wColor
fgDétail.Col = 4: fgDétail.Text = dateIBM10(xYTVACOM0.TVACOMDTR, True)
fgDétail.CellForeColor = wColor
fgDétail.Col = 5: fgDétail.Text = xYTVACOM0.TVACOMCLIC & " " & xYTVACOM0.TVACOMCLI
fgDétail.CellForeColor = wColor
fgDétail.Col = 6: fgDétail.Text = xYTVACOM0.TVACOMCLIP
fgDétail.CellForeColor = wColor
fgDétail.Col = 7: fgDétail.Text = xYTVACOM0.TVACOMCOMC
fgDétail.CellForeColor = wColor
fgDétail.Col = 8: fgDétail.Text = xYTVACOM0.TVACOMCOME & " " & xYTVACOM0.TVACOMCOMT
fgDétail.CellForeColor = wColor
fgDétail.Col = 9: fgDétail.Text = xYTVACOM0.TVACOMTVAC
fgDétail.CellForeColor = wColor
fgDétail.Col = 10: fgDétail.Text = xYTVACOM0.TVACOMSTA
fgDétail.CellForeColor = wColor
fgDétail.Col = 11: fgDétail.Text = Trim(Format$(xYTVACOM0.TVACOMFACN, "### ### ###"))
fgDétail.CellForeColor = wColor
fgDétail.Col = 12: fgDétail.Text = Trim(Format$(xYTVACOM0.TVACOMFACL, "### ### ###"))
fgDétail.CellForeColor = wColor

fgDétail.Col = fgDétail_arrIndex: fgDétail.Text = lIndex

End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

End Sub


Public Sub fgDétail_Reset()
fgDétail.Clear
fgDétail_Sort1 = 0: fgDétail_Sort2 = 0
fgDétail_Sort1_Old = -1
fgDétail_RowDisplay = 0: fgDétail_RowClick = 0
fgDétail_arrIndex = fgDétail.Cols - 1
blnfgDétail_DisplayLine = False
fgDétail_SortAD = 6
fgDétail.LeftCol = 0

End Sub

Public Sub fgSelect_Sort()
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

End Sub

Public Sub fgDétail_Sort()
If fgDétail.Rows > 1 Then
    fgDétail.Row = 1
    fgDétail.RowSel = fgDétail.Rows - 1
    
    If fgDétail_Sort1_Old = fgDétail_Sort1 Then
        If fgDétail_SortAD = 5 Then
            fgDétail_SortAD = 6
        Else
            fgDétail_SortAD = 5
        End If
    Else
        fgDétail_SortAD = 5
    End If
    fgDétail_Sort1_Old = fgDétail_Sort1
    
    fgDétail.Col = fgDétail_Sort1
    fgDétail.ColSel = fgDétail_Sort2
    fgDétail.Sort = fgDétail_SortAD
End If

End Sub

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
        Case 4: X = Format$(arrYTVAFAC0(wIndex).TVAFACFACN, "000000000")
        Case 5: X = arrYTVAFAC0(wIndex).TVAFACDTR
        Case 6: X = Format$(arrYTVAFAC0(wIndex).TVAFACMTTC, "000000000000000.00")
        Case 7: X = Format$(arrYTVAFAC0(wIndex).TVAFACMTVA, "000000000000000.00")
        Case 8: X = Format$(arrYTVAFAC0(wIndex).TVAFACMEXO, "000000000000000.00")
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub
Public Sub fgDétail_SortX(lK As Integer)
Dim I As Integer, X As String
Dim wIndex As Integer
For I = 1 To fgDétail.Rows - 1
    fgDétail.Row = I
    fgDétail.Col = fgDétail_arrIndex
    wIndex = Val(fgDétail.Text)
    Select Case lK
        Case 0: X = arrYTVACOM0(wIndex).TVACOMOPE & Format$(arrYTVACOM0(wIndex).TVACOMDOS, "000000000")
        Case 1: X = Format$(arrYTVACOM0(wIndex).TVACOMDOS, "000000000")
        Case 2: X = Format$(arrYTVACOM0(wIndex).TVACOMMON, "000000000000000.00")
        Case 3: X = arrYTVACOM0(wIndex).TVACOMDEV & Format$(arrYTVACOM0(wIndex).TVACOMMON, "000000000000000.00")
        Case 4: X = arrYTVACOM0(wIndex).TVACOMDTR
        Case 11: X = Format$(arrYTVACOM0(wIndex).TVACOMFACN, "00000000000")
        Case 12: X = Format$(arrYTVACOM0(wIndex).TVACOMFACL, "00000000000")
    End Select
    fgDétail.Col = fgDétail_arrIndex - 1
    fgDétail.Text = X
Next I


fgDétail_Sort1 = fgDétail_arrIndex - 1: fgDétail_Sort2 = fgDétail_arrIndex - 1
fgDétail_Sort
End Sub


Public Sub fgNIF_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgNIF.Row

If lRow > 0 And lRow < fgNIF.Rows Then
    fgNIF.Row = lRow
    For I = 0 To fgNIF_arrIndex
        fgNIF.Col = I: fgNIF.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgNIF.Row = mRow
    If fgNIF.Row > 0 Then
        lRow = fgNIF.Row
        lColor_Old = fgNIF.CellBackColor
        For I = fgNIF_arrIndex To 0 Step -1
          fgNIF.Col = I: fgNIF.CellBackColor = lColor
        Next I
        fgNIF.LeftCol = 0
    End If
End If

End Sub

Private Sub fgNIF_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 2
fgNIF.Visible = False
fgNIF_Reset
fgNIF.Rows = 1
fgNIF.FormatString = fgNIF_FormatString
cmdPrint.Enabled = False
currentAction = "fgNIF_Display"

For I = 1 To arrYTVANIF0_Nb

     xYTVANIF0 = arrYTVANIF0(I)

        fgNIF.Rows = fgNIF.Rows + 1
        fgNIF.Row = fgNIF.Rows - 1
        fgNIF_DisplayLine I
Next I

Call lstErr_AddItem(lstErr, cmdContext, "Nb enregistrements : " & fgNIF.Rows - 1): DoEvents
If fgNIF.Rows > 1 Then
'    fgNIF_Sort1 = 0: fgNIF_Sort2 = 2: fgNIF_Sort
    cmdPrint.Enabled = True
End If
fgNIF.LeftCol = 0

fgNIF.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Public Sub fgNIF_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long
Dim V
On Error Resume Next
     Select Case xYTVANIF0.TVANIFSTA
        Case Is = "V": wColor = vbBlue
        Case Is = "F": wColor = &HC000&
        Case Is = "A": wColor = &HC0C0C0
        Case Is = "I": wColor = &HC0C000
        Case Is = " ": wColor = vbMagenta
        Case Else: wColor = vbBlack
    End Select

fgNIF.Col = 0: fgNIF.Text = xYTVANIF0.TVANIFCLIC & " " & xYTVANIF0.TVANIFCLI
fgNIF.CellForeColor = wColor
fgNIF.Col = 1: fgNIF.Text = xYTVANIF0.TVANIFRS
fgNIF.CellForeColor = wColor
fgNIF.Col = 2: fgNIF.Text = xYTVANIF0.TVANIFCLIP
fgNIF.CellForeColor = wColor
fgNIF.Col = 3: fgNIF.Text = TVANIFCLIT_Format(xYTVANIF0.TVANIFCLIT)
fgNIF.CellForeColor = wColor
If Trim(xYTVANIF0.TVANIFCLIT) <> "" Then
    V = TVANIFCLIT_Control(xYTVANIF0.TVANIFCLIT)
    If xYTVANIF0.TVANIFCLIP <> Mid$(xYTVANIF0.TVANIFCLIT, 1, 2) Then
        If xYTVANIF0.TVANIFCLIP = "GR" And Mid$(xYTVANIF0.TVANIFCLIT, 1, 2) = "EL" Then
        Else
            V = "Erreur pays"
        End If
    End If
    If Not IsNull(V) Then
        fgNIF.CellForeColor = vbRed
        Call lstErr_AddItem(lstErr, cmdContext, xYTVANIF0.TVANIFCLI & " : " & V): DoEvents
    End If
End If
fgNIF.Col = 4: fgNIF.Text = xYTVANIF0.TVANIFSTA
fgNIF.CellForeColor = wColor

fgNIF.Col = fgNIF_arrIndex: fgNIF.Text = lIndex

End Sub

Public Sub fgNIF_Reset()
fgNIF.Clear
fgNIF_Sort1 = 0: fgNIF_Sort2 = 0
fgNIF_Sort1_Old = -1
fgNIF_RowDisplay = 0: fgNIF_RowClick = 0
fgNIF_arrIndex = fgNIF.Cols - 1
blnfgNIF_DisplayLine = False
fgNIF_SortAD = 6
fgNIF.LeftCol = 0

End Sub

Public Sub fgNIF_Sort()
If fgNIF.Rows > 1 Then
    fgNIF.Row = 1
    fgNIF.RowSel = fgNIF.Rows - 1
    
    If fgNIF_Sort1_Old = fgNIF_Sort1 Then
        If fgNIF_SortAD = 5 Then
            fgNIF_SortAD = 6
        Else
            fgNIF_SortAD = 5
        End If
    Else
        fgNIF_SortAD = 5
    End If
    fgNIF_Sort1_Old = fgNIF_Sort1
    
    fgNIF.Col = fgNIF_Sort1
    fgNIF.ColSel = fgNIF_Sort2
    fgNIF.Sort = fgNIF_SortAD
End If

End Sub

Public Sub fgNIF_SortX(lK As Integer)
Dim I As Integer, X As String
Dim wIndex As Integer
For I = 1 To fgNIF.Rows - 1
    fgNIF.Row = I
    fgNIF.Col = fgNIF_arrIndex
    wIndex = Val(fgNIF.Text)
    Select Case lK
'        Case 0: X = arrYTVANIF0(wIndex).TVANIFOPE & Format$(arrYTVANIF0(wIndex).TVANIFDOS, "000000000")
    End Select
    fgNIF.Col = fgNIF_arrIndex - 1
    fgNIF.Text = X
Next I


fgNIF_Sort1 = fgNIF_arrIndex - 1: fgNIF_Sort2 = fgNIF_arrIndex - 1
fgNIF_Sort
End Sub




Public Sub fgParam_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgParam.Row

If lRow > 0 And lRow < fgParam.Rows Then
    fgParam.Row = lRow
    For I = 0 To fgParam_arrIndex
        fgParam.Col = I: fgParam.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgParam.Row = mRow
    If fgParam.Row > 0 Then
        lRow = fgParam.Row
        lColor_Old = fgParam.CellBackColor
        For I = fgParam_arrIndex To 0 Step -1
          fgParam.Col = I: fgParam.CellBackColor = lColor
        Next I
        fgParam.LeftCol = 0
    End If
End If

End Sub
Private Sub fgParam_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 3
fgParam.Visible = False
fgParam_Reset
fgParam.Rows = 1
fgParam.FormatString = fgParam_FormatString
cmdPrint.Enabled = False
currentAction = "fgParam_Display"

For I = 1 To arrYBIATAB0_Nb

     xYBIATAB0 = arrYBIATAB0(I)

        fgParam.Rows = fgParam.Rows + 1
        fgParam.Row = fgParam.Rows - 1
        fgParam_DisplayLine I
Next I

Call lstErr_AddItem(lstErr, cmdContext, "Nb enregistrements : " & fgParam.Rows - 1): DoEvents
If fgParam.Rows > 1 Then
'    fgParam_Sort1 = 0: fgParam_Sort2 = 2: fgParam_Sort
    cmdPrint.Enabled = True
End If
fgParam.LeftCol = 0

fgParam.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Public Sub fgParam_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long
Dim V
On Error Resume Next
If Trim(xYBIATAB0.BIATABID) = "TVAFACSTA" Then
    wColor = vbBlue
Else
     Select Case Mid$(xYBIATAB0.BIATABTXT, 1, 1)
        Case Is = "V": wColor = vbBlue
        Case Is = "F": wColor = &HC000&
        Case Is = "A": wColor = &HC0C0C0
        Case Is = "I": wColor = &HC0C000
        Case Is = " ": wColor = vbMagenta
        Case Else: wColor = vbBlack
    End Select
End If
fgParam.Col = 0: fgParam.Text = xYBIATAB0.BIATABID
fgParam.CellForeColor = wColor
fgParam.Col = 1: fgParam.Text = xYBIATAB0.BIATABK1
fgParam.CellForeColor = wColor
fgParam.Col = 2: fgParam.Text = xYBIATAB0.BIATABK2
fgParam.CellForeColor = wColor
fgParam.Col = 3: fgParam.Text = xYBIATAB0.BIATABTXT
fgParam.CellForeColor = wColor
fgParam.Col = fgParam_arrIndex: fgParam.Text = lIndex

End Sub

Public Sub fgParam_Reset()
fgParam.Clear
fgParam_Sort1 = 0: fgParam_Sort2 = 0
fgParam_Sort1_Old = -1
fgParam_RowDisplay = 0: fgParam_RowClick = 0
fgParam_arrIndex = fgParam.Cols - 1
blnfgParam_DisplayLine = False
fgParam_SortAD = 6
fgParam.LeftCol = 0

End Sub

Public Sub fgParam_Sort()
If fgParam.Rows > 1 Then
    fgParam.Row = 1
    fgParam.RowSel = fgParam.Rows - 1
    
    If fgParam_Sort1_Old = fgParam_Sort1 Then
        If fgParam_SortAD = 5 Then
            fgParam_SortAD = 6
        Else
            fgParam_SortAD = 5
        End If
    Else
        fgParam_SortAD = 5
    End If
    fgParam_Sort1_Old = fgParam_Sort1
    
    fgParam.Col = fgParam_Sort1
    fgParam.ColSel = fgParam_Sort2
    fgParam.Sort = fgParam_SortAD
End If

End Sub

Public Sub fgParam_SortX(lK As Integer)
Dim I As Integer, X As String
Dim wIndex As Integer
For I = 1 To fgParam.Rows - 1
    fgParam.Row = I
    fgParam.Col = fgParam_arrIndex
    wIndex = Val(fgParam.Text)
    Select Case lK
'        Case 0: X = arrYBIATAB0(wIndex).TVAParamOPE & Format$(arrYBIATAB0(wIndex).TVAParamDOS, "000000000")
    End Select
    fgParam.Col = fgParam_arrIndex - 1
    fgParam.Text = X
Next I


fgParam_Sort1 = fgParam_arrIndex - 1: fgParam_Sort2 = fgParam_arrIndex - 1
fgParam_Sort
End Sub







'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
Select Case SSTab1.Tab
    Case Is = 1
        If fraDétail_Update.Visible Then fraDétail_Update.Visible = False: Exit Sub
        If fgDétail.Visible Then fgDétail.Visible = False: cmdDétail_Ok.Caption = "Extraire les commissions": Exit Sub
        SSTab1.Tab = 0: Exit Sub
    Case 2
        If fraNIF_Update.Visible Then fraNIF_Update.Visible = False: Exit Sub
        If fgNIF.Visible Then fgNIF.Visible = False: cmdNIF_Ok.Caption = "Extraire les NIF": Exit Sub
        SSTab1.Tab = 1: Exit Sub
    Case 3
        If fraParam_Update.Visible Then fraParam_Update.Visible = False: Exit Sub
        If fgParam.Visible Then fgParam.Visible = False: cmdParam_Ok.Caption = "Extraire les Param": Exit Sub
        SSTab1.Tab = 1: Exit Sub
    Case Else
    
        If fraSelect_Update.Visible Then fraSelect_Update.Visible = False: Exit Sub
        If fgSelect.Visible Then fgSelect.Visible = False: cmdSelect_Ok.Caption = "Extraire les factures": Exit Sub
        Unload Me
End Select
End Sub




Private Sub cboDétail_SQL_Click()
cmdDétail_SQL_K = Mid$(cboDétail_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    fraDétail_Update.Visible = False
    fgDétail.Visible = False
    Select Case cmdDétail_SQL_K
        Case "1": lstDétail_Load_1
        Case "2": lstDétail_Load_2
        Case "3": lstDétail_Load_3
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If

End Sub


Private Sub cboNIF_SQL_Click()
cmdNIF_SQL_K = Mid$(cboNIF_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    fraNIF_Update.Visible = False
    fgNIF.Visible = False
    Select Case cmdNIF_SQL_K
        Case "1": lstNIF_Load_1
        Case "2": lstNIF_Load_2
        Case "3": lstNIF_Load_3
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If

End Sub


Private Sub cboParam_SQL_Click()
cmdParam_SQL_K = Mid$(cboParam_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    fraParam_Update.Visible = False
    fgParam.Visible = False
    Select Case cmdParam_SQL_K
        Case "1": lstParam_Load_1
        Case "2": lstParam_Load_2
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = Mid$(cboSelect_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    fraSelect_Options_1.Visible = False
    fraSelect_Update.Visible = False
    fgDétail.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1": lstSelect_Load_1
        Case "2": lstSelect_Load_2
        Case "3": lstSelect_Load_3
        Case "6": lstSelect_Load_6
        Case "7": lstSelect_Load_7
        Case "8": lstSelect_Load_8
        Case "9": lstSelect_Load_8
        Case "S": lstSelect_Load_S
   End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub chkDétail_TVACOMDTR_Click()
If chkDétail_TVACOMDTR = "1" Then
    txtDétail_TVACOMDTR.Visible = True
    txtDétail_TVACOMDTR_Max.Visible = True
Else
    txtDétail_TVACOMDTR.Visible = False
    txtDétail_TVACOMDTR_Max.Visible = False
End If

End Sub

Private Sub chkParamUpdate_Insert_Click()
chkParamUpdate_Insert.Enabled = Not chkParamUpdate_Insert.Enabled
If chkParamUpdate_Insert = "1" Then fraParam_Update_TVACOMOPE.Enabled = True: txtParamUpdate_CLIENACLI.Enabled = True

End Sub

Private Sub chkSelect_TVAFACDTR_Click()
Select Case chkSelect_TVAFACDTR
    Case Is = "1"
        txtSelect_TVAFACDTR.Visible = True
        txtSelect_TVAFACDTR_Max.Visible = True
    Case Is = "7"
        txtSelect_TVAFACDTR.Visible = True
        txtSelect_TVAFACDTR_Max.Visible = False
   Case Else
    txtSelect_TVAFACDTR.Visible = False
    txtSelect_TVAFACDTR_Max.Visible = False
End Select


End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdDétail_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
blnOk = Not fgDétail.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmDétail_Ok ........"): DoEvents
cmdDétail_Ok.Visible = False
fraDétail_Update.Visible = False
fraDétail_Options_1.Enabled = False
fgDétail.Clear
DoEvents
If blnOk Then
    cmdDétail_Ok.Caption = "Modifier les options"
    cmdDétail_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraDétail_Options_1.BackColor = &H8000000F
    Call usrColor_Container(fraDétail_Options_1, fraDétail_Options_1.BackColor)
    Select Case cmdDétail_SQL_K
        Case "1": cmdDétail_SQL
        Case "2": cmdDétail_SQL
        Case "3": cmdDétail_SQL_3: cboDétail_SQL.ListIndex = 0
        Case "6": cmdDétail_SQL_6: cboDétail_SQL.ListIndex = 0
    End Select

    fgDétail.Enabled = True
Else
    cmdDétail_Ok.Caption = cmdDétail_Ok_Caption
    cmdDétail_Ok.BackColor = &HE0FFFF
    fraDétail_Options_1.BackColor = &HE0FFFF
    Call usrColor_Container(fraDétail_Options_1, fraDétail_Options_1.BackColor)
    fgDétail.Visible = False
    fgDétail.Enabled = False
    fraDétail_Options_1.Enabled = True

End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmDétail_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdDétail_Ok.Visible = True


End Sub

Private Sub cmdDétail_Update_Annuler_Click()
Dim V
App_Debug = "cmdDétail_Update_Annuler"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

newYTVACOM0 = oldYTVACOM0
If newYTVACOM0.TVACOMSTA = "A" Then
    If xYTVACOM0.TVACOMMON > 0 Then
        newYTVACOM0.TVACOMSTA = "8"
    Else
        newYTVACOM0.TVACOMSTA = "9"
    End If
Else
    newYTVACOM0.TVACOMSTA = "A"
End If


    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données " & App_Debug): DoEvents
    'V = cmdSelect_Update_Annuler_Transaction
    V = cmdDétail_Update_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYTVACOM0(arrYTVACOM0_Index) = newYTVACOM0
        xYTVACOM0 = newYTVACOM0
        fgDétail_DisplayLine arrYTVACOM0_Index
        fraDétail_Update.Visible = False

    Else
        MsgBox V, vbCritical, Me.Name & " : cmdSelect_Update_Ok" & App_Debug
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdDétail_Update_Ok_Click()

Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement cmdDétail_Update_Ok"): DoEvents

If IsNull(fraDétail_Update_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdDétail_Update_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYTVACOM0(arrYTVACOM0_Index) = newYTVACOM0
        xYTVACOM0 = newYTVACOM0
        fgDétail_DisplayLine arrYTVACOM0_Index
        fraDétail_Update.Visible = False
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdSelect_Update_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement  cmdDétail_Update_Ok"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDétail_Update_Quit_Click()
fraDétail_Update.Visible = False

End Sub



Private Sub cmdNIF_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
blnOk = Not fgNIF.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdNIF_Ok ........"): DoEvents
cmdNIF_Ok.Visible = False
fraNIF_Update.Visible = False
fraNIF_Options_1.Enabled = False
fgNIF.Clear
DoEvents
If blnOk Then
    cmdNIF_Ok.Caption = "Modifier les options"
    cmdNIF_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraNIF_Options_1.BackColor = &H8000000F
    Call usrColor_Container(fraNIF_Options_1, fraNIF_Options_1.BackColor)
    Select Case cmdNIF_SQL_K
        Case "1": cmdNIF_SQL_1
        Case "2": cmdNIF_SQL_2
        Case "3": cmdNIF_SQL_3
    End Select

    fgNIF.Enabled = True
Else
    cmdNIF_Ok.Caption = cmdNIF_Ok_Caption
    cmdNIF_Ok.BackColor = &HE0FFFF
    fraNIF_Options_1.BackColor = &HE0FFFF    '&HD0D0D0
    Call usrColor_Container(fraNIF_Options_1, fraNIF_Options_1.BackColor)
    fgNIF.Visible = False
    fgNIF.Enabled = False
    fraNIF_Options_1.Enabled = True

End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdNIF_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdNIF_Ok.Visible = True

End Sub

Private Sub cmdNIF_Update_Annuler_Click()
Dim V
App_Debug = "cmdNIF_Update_Annuler"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

newYTVANIF0 = oldYTVANIF0
If newYTVANIF0.TVANIFSTA = "A" Then
    newYTVANIF0.TVANIFSTA = "V"
Else
    newYTVANIF0.TVANIFSTA = "A"
End If


    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données " & App_Debug): DoEvents
'    V = cmdNIF_Update_Ok_Transaction("Update")
    V = cmdNIF_Update_Ok_Transaction("Delete")
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        Select Case cmdNIF_SQL_K
            Case "1": cmdNIF_SQL_1
            Case "2": cmdNIF_SQL_2
            Case "3": cmdNIF_SQL_3
        End Select
        fraNIF_Update.Visible = False

    Else
        MsgBox V, vbCritical, Me.Name & " : cmdNIF_Update_Ok" & App_Debug
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdNIF_Update_Ok_Click()
Dim V
App_Debug = "cmdNIF_Update_Ok"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

If IsNull(fraNIF_Update_Control) Then

    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données " & App_Debug): DoEvents
    If oldYTVANIF0.TVANIFSTA = " " Then
        V = cmdNIF_Update_Ok_Transaction("Insert")
    Else
        V = cmdNIF_Update_Ok_Transaction("Update")
    End If
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYTVANIF0(arrYTVANIF0_Index) = newYTVANIF0
        xYTVANIF0 = newYTVANIF0
        fgNIF_DisplayLine arrYTVANIF0_Index
        fraNIF_Update.Visible = False

    Else
        MsgBox V, vbCritical, Me.Name & " : cmdNIF_Update_Ok" & App_Debug
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdNIF_Update_Quit_Click()
fraNIF_Update.Visible = False

End Sub

Private Sub cmdParam_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
blnOk = Not fgParam.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdParam_Ok ........"): DoEvents
cmdParam_Ok.Visible = False
fraParam_Update.Visible = False
fraParam_Options_1.Enabled = False
fgParam.Clear
DoEvents
If blnOk Then
    cmdParam_Ok.Caption = "Modifier les options"
    cmdParam_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraParam_Options_1.BackColor = &H8000000F
    Call usrColor_Container(fraParam_Options_1, fraParam_Options_1.BackColor)
    Select Case cmdParam_SQL_K
        Case "1": cmdParam_SQL
        Case "2": cmdParam_SQL
    End Select

    fgParam.Enabled = True
Else
    cmdParam_Ok.Caption = cmdParam_Ok_Caption
    cmdParam_Ok.BackColor = &HE0FFFF
    fraParam_Options_1.BackColor = &HE0FFFF    '&HD0D0D0
    Call usrColor_Container(fraParam_Options_1, fraParam_Options_1.BackColor)
    fgParam.Visible = False
    fgParam.Enabled = False
    fraParam_Options_1.Enabled = True

End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdParam_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdParam_Ok.Visible = True

End Sub

Private Sub cmdParam_Update_Annuler_Click()
Dim V
App_Debug = "cmdParam_Update_Annuler"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents


    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données " & App_Debug): DoEvents
    V = cmdParam_Update_Ok_Transaction("Delete")
    If IsNull(V) Then
        cmdParam_SQL
        fraParam_Update.Visible = False
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdParam_Update_Ok" & App_Debug
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Update_Ok_Click()
Dim V
App_Debug = "cmdParam_Update_Ok"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

If IsNull(fraParam_Update_Control) Then

    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données " & App_Debug): DoEvents
    If chkParamUpdate_Insert = "1" Then
        V = cmdParam_Update_Ok_Transaction("Insert")
        
    Else
        V = cmdParam_Update_Ok_Transaction("Update")
    End If
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        If chkParamUpdate_Insert = "1" Then
            cmdParam_SQL
            fraParam_Update.Visible = False
        Else
            arrYBIATAB0(arrYBIATAB0_Index) = newYBIATAB0
            xYBIATAB0 = newYBIATAB0
            fgParam_DisplayLine arrYBIATAB0_Index
            fraParam_Update.Visible = False
        End If
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdParam_Update_Ok" & App_Debug
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Update_Quit_Click()
fraParam_Update.Visible = False

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False
fraSelect_Options_1.Visible = False
cmdSelect_Ok.Caption = "Extraire les mouvements"

libRéférenceInterne = ""
If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0
lstSelect_Load_1
Call DTPicker_Set(txtSelect_TVAFACDTR, YBIATAB0_DATE_CPT_JP0)
Call DTPicker_Set(txtSelect_TVAFACDTR_Max, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtDétail_TVACOMDTR, YBIATAB0_DATE_CPT_JP0)
Call DTPicker_Set(txtDétail_TVACOMDTR_Max, YBIATAB0_DATE_CPT_J)
fraSelect_Update.Visible = False
fgSelect.Visible = False  'True
'cmdSelect_Ok_Click

chkDétail_TVACOMDTR = "0"
fraDétail_Update.Visible = False
If cboDétail_SQL.ListCount > 0 Then cboDétail_SQL.ListIndex = 0
lstDétail_Load_1

If cboNIF_SQL.ListCount > 0 Then cboNIF_SQL.ListIndex = 0
lstNIF_Load_1
fraNIF_Update.Visible = False
fgNIF.Visible = False  'True
'cmdNIF_Ok_Click

If cboParam_SQL.ListCount > 0 Then cboParam_SQL.ListIndex = 0
lstParam_Load_1
fraParam_Update.Visible = False
fgParam.Visible = False  'True

If BIA_TVAFAC_Aut.Consulter Then
    SSTab1.Tab = 0
Else
    If BIA_TVACOM_Aut.Consulter Then
        SSTab1.Tab = 1
    Else
        SSTab1.Tab = 2
    End If
End If


blnControl = True
End Sub
Public Sub Form_Init()
Dim xZBASTAB0 As typeZBASTAB0
Dim K As Integer

Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents


blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fgDétail_FormatString = fgDétail.FormatString
fraSelect.Enabled = BIA_TVAFAC_Aut.Consulter
cmdSelect_Ok.Visible = False
fraSelect_Options_1.Visible = False
txtSelect_TVAFACDTR.Visible = False
txtSelect_TVAFACDTR_Max.Visible = False
cboSelect_SQL.Clear
If BIA_TVAFAC_Aut.Consulter Then
    cboSelect_SQL.AddItem "1 - Liste des factures"
    cboSelect_SQL.AddItem "2 - Factures incomplètes(n° TVA)"
    cboSelect_SQL.AddItem "3 - Factures à valider"
    cboSelect_SQL.AddItem "S - Statistiques commissions(BDF/CMP) "
End If
If BIA_TVAFAC_Aut.Xspécial Then
    cboSelect_SQL.AddItem "6 - Regroupement des commissions => facture"
    cboSelect_SQL.AddItem "7 - Emission des factures(archivage)"
    cboSelect_SQL.AddItem "8 - Impression des factures(papier)"
End If
If BIA_TVAFAC_Aut.Comptabiliser Then cboSelect_SQL.AddItem "9 - Déclaration TVA aux douanes (.pdf et .xml)"

'_____________________________________________________________________________TVAFAC

cboDétail_SQL.Clear
If BIA_TVACOM_Aut.Consulter Then
    cboDétail_SQL.AddItem "1 - Liste des commsissions"
    cboDétail_SQL.AddItem "2 - Liste des commsissions à réviser"
    cboDétail_SQL.AddItem "3 - Etat/Service des commsissions à réviser"
    cboDétail_SQL.AddItem "6 - Ignorer les commissions / Client hors UE (SAB >5)"
End If
txtUpdate_TVAFACCLIC.Clear
txtUpdate_TVAFACCLIC.AddItem "  client BIA"
txtUpdate_TVAFACCLIC.AddItem "G tiers change"
txtUpdate_TVAFACCLIC.AddItem "D tiers crédoc"
txtUpdate_TVAFACCLIC.AddItem "R tiers remdoc"

txtUpdate_TVAFACCLIP.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENAPAY", txtUpdate_TVAFACCLIP
txtUpdate_TVAFACCLIP.AddItem "  ???"

txtSelect_TVAFACCLIC.Clear
txtSelect_TVAFACCLIC.AddItem "* "
txtSelect_TVAFACCLIC.AddItem "  client BIA"
txtSelect_TVAFACCLIC.AddItem "G tiers change"
txtSelect_TVAFACCLIC.AddItem "D tiers crédoc"
txtSelect_TVAFACCLIC.AddItem "R tiers remdoc"
txtSelect_TVAFACCLIC.ListIndex = 0

txtSelect_TVAFACSTA.Clear
txtSelect_TVAFACSTA.AddItem "  ???"
txtSelect_TVAFACSTA.AddItem "0 - manque code TVA intracommunautaire"
txtSelect_TVAFACSTA.AddItem "1 - OK code TVA intracommunautaire connu"
txtSelect_TVAFACSTA.AddItem "2 - OK non soumis à la TVA intracommunautaire"
txtSelect_TVAFACSTA.AddItem "9 - en anomalie"
txtSelect_TVAFACSTA.AddItem "A --annulé"
txtSelect_TVAFACSTA.AddItem "F - facture émise"
txtSelect_TVAFACSTA.AddItem "I - à ignorer"
txtSelect_TVAFACSTA.AddItem "V - validé pour impression"

txtUpdate_TVAFACSTA.Clear
txtUpdate_TVAFACSTA.AddItem "  ???"
txtUpdate_TVAFACSTA.AddItem "0 - manque code TVA intracommunautaire"
txtUpdate_TVAFACSTA.AddItem "1 - OK code TVA intracommunautaire connu"
txtUpdate_TVAFACSTA.AddItem "2 - OK non soumis à la TVA intracommunautaire"
txtUpdate_TVAFACSTA.AddItem "9 - en anomalie"
txtUpdate_TVAFACSTA.AddItem "A - annulé"
txtUpdate_TVAFACSTA.AddItem "F - facture émise"
txtUpdate_TVAFACSTA.AddItem "I - à ignorer"
txtUpdate_TVAFACSTA.AddItem "V - validé pour impression"

'_____________________________________________________________________________TVACOM
fraDétail.Enabled = BIA_TVACOM_Aut.Consulter
txtDétail_TVACOMDTR.Visible = False
txtDétail_TVACOMDTR_Max.Visible = False


txtUpdate_TVACOMCOMB.Clear
txtUpdate_TVACOMCOMB.AddItem "  ???"
txtUpdate_TVACOMCOMB.AddItem "Bénéficiaire"
txtUpdate_TVACOMCOMB.AddItem "Donneur d'ordre"
txtUpdate_TVACOMCOMB.AddItem "Share"

txtUpdate_TVACOMCOMT.Clear
txtUpdate_TVACOMCOMT.AddItem "  ???"
txtUpdate_TVACOMCOMT.AddItem "E exonéré"
txtUpdate_TVACOMCOMT.AddItem "N normal"
txtUpdate_TVACOMCOMT.AddItem "R Réduit"

txtUpdate_TVACOMCOME.Clear
txtUpdate_TVACOMCOME.AddItem "  facture papier"
txtUpdate_TVACOMCOME.AddItem "W swift (=facture)"
txtUpdate_TVACOMCOME.AddItem "A avis (=facture)"

txtUpdate_TVACOMTVAC.Clear
txtUpdate_TVACOMTVAC.AddItem "  ???"
txtUpdate_TVACOMTVAC.AddItem "E exonéré"
txtUpdate_TVACOMTVAC.AddItem "N normal"
txtUpdate_TVACOMTVAC.AddItem "R Réduit"
txtUpdate_TVACOMTVAC.AddItem "* non applicable"
'txtUpdate_TVACOMTVAC.AddItem "T tva"

txtUpdate_TVACOMCLIC.Clear
txtUpdate_TVACOMCLIC.AddItem "  client BIA"
txtUpdate_TVACOMCLIC.AddItem "G tiers change"
txtUpdate_TVACOMCLIC.AddItem "D tiers crédoc"
txtUpdate_TVACOMCLIC.AddItem "R tiers remdoc"

txtUpdate_TVACOMCLIP.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENAPAY", txtUpdate_TVACOMCLIP
txtUpdate_TVACOMCLIP.AddItem "  ???"

txtUpdate_TVACOMSTA.Clear
txtUpdate_TVACOMSTA.AddItem "  ???"
txtUpdate_TVACOMSTA.AddItem "0 import mvt"
txtUpdate_TVACOMSTA.AddItem "8 avoir/annulation com"
txtUpdate_TVACOMSTA.AddItem "9 en anomalie"
txtUpdate_TVACOMSTA.AddItem "A annulé"
txtUpdate_TVACOMSTA.AddItem "F facture"
txtUpdate_TVACOMSTA.AddItem "I à ignorer"
txtUpdate_TVACOMSTA.AddItem "V validé pour facturation"
txtUpdate_TVACOMSTA.AddItem "X traité par appli spécifique"


txtDétail_TVACOMSTA.Clear
txtDétail_TVACOMSTA.AddItem "  ???"
txtDétail_TVACOMSTA.AddItem "0 import mvt"
txtDétail_TVACOMSTA.AddItem "8 avoir/annulation com"
txtDétail_TVACOMSTA.AddItem "9 en anomalie"
txtDétail_TVACOMSTA.AddItem "A annulé"
txtDétail_TVACOMSTA.AddItem "F facture"
txtDétail_TVACOMSTA.AddItem "I à ignorer"
txtDétail_TVACOMSTA.AddItem "V validé pour facturation"
txtDétail_TVACOMSTA.AddItem "X traité par appli spécifique"

Call rsZBASTAB0_cboK2(44, txtUpdate_TVACOMCOMC, "")
txtUpdate_TVACOMCOMC.AddItem "=ECRX  *   : libellé écriture comptable"

X = "select count(*) as Tally from " & paramIBM_Library_SAB & ".ZECHTAB0 " _
    & " where ECHTABNUM = 13 "
Set rsSab = cnsab.Execute(X)
K = rsSab("Tally")

ReDim arrCommission(txtUpdate_TVACOMCOMC.ListCount + K, 2)
X = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where BASTABNUM = 44 order by BASTABARG"
    
Set rsSab = cnsab.Execute(X)
arrCommission_Max = 0

Do While Not rsSab.EOF
        arrCommission_Max = arrCommission_Max + 1
        arrCommission(arrCommission_Max, 1) = Trim(Mid$(rsSab("BASTABARG"), 1, 6))
        arrCommission(arrCommission_Max, 2) = Mid$(rsSab("BASTABDON"), 1, 30)
    rsSab.MoveNext
Loop
X = "select * from " & paramIBM_Library_SAB & ".ZECHTAB0 " _
    & " where ECHTABNUM = 13 order by ECHTABARG"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
        X = "=" & Mid$(rsSab("ECHTABARG"), 4, 5) '
        arrCommission_Max = arrCommission_Max + 1
        arrCommission(arrCommission_Max, 1) = Trim(X)
        arrCommission(arrCommission_Max, 2) = Trim(Mid$(rsSab("ECHTABDON"), 1, 30))
        txtUpdate_TVACOMCOMC.AddItem X & " *   : " & arrCommission(arrCommission_Max, 2)
    rsSab.MoveNext
Loop

txtDétail_TVACOMCLIC.Clear
txtDétail_TVACOMCLIC.AddItem "* "
txtDétail_TVACOMCLIC.AddItem "  client BIA"
txtDétail_TVACOMCLIC.AddItem "G tiers change"
txtDétail_TVACOMCLIC.AddItem "D tiers crédoc"
txtDétail_TVACOMCLIC.AddItem "R tiers remdoc"
txtDétail_TVACOMCLIC.ListIndex = 0

txtDétail_TVACOMOPE.Clear
txtDétail_TVACOMOPE.AddItem "  "
txtDétail_TVACOMOPE.AddItem "*"
txtDétail_TVACOMOPE.AddItem "AP"
txtDétail_TVACOMOPE.AddItem "AV"
txtDétail_TVACOMOPE.AddItem "CDE"
txtDétail_TVACOMOPE.AddItem "CDI"
txtDétail_TVACOMOPE.AddItem "CPT"
txtDétail_TVACOMOPE.AddItem "CRE"
txtDétail_TVACOMOPE.AddItem "ECH"
txtDétail_TVACOMOPE.AddItem "ENG"
txtDétail_TVACOMOPE.AddItem "FRS"
txtDétail_TVACOMOPE.AddItem "PRE"
txtDétail_TVACOMOPE.AddItem "RDE"
txtDétail_TVACOMOPE.AddItem "RDI"
txtDétail_TVACOMOPE.AddItem "REM"
txtDétail_TVACOMOPE.AddItem "RPC"
txtDétail_TVACOMOPE.AddItem "TRF"

cbo_Load_Unit txtUpdate_TVACOMSRVR
cbo_Load_Unit txtDétail_TVACOMSRVR
txtDétail_TVACOMSRVR.AddItem "  "
cbo_Scan currentUser.Unit, txtDétail_TVACOMSRVR

'_____________________________________________________________________________TVANIF
fraNIF.Enabled = BIA_TVANIF_Aut.Consulter
libUpdate_TVANIFCLIT.ForeColor = vbMagenta

fgNIF_FormatString = fgNIF.FormatString
cboNIF_SQL.Clear
If BIA_TVANIF_Aut.Consulter Then
    cboNIF_SQL.AddItem "1 - Liste des NIF"
    cboNIF_SQL.AddItem "2 - Liste des tiers"
    cboNIF_SQL.AddItem "3 - Liste des tiers à facturer - NIF manaquant"
End If
txtUpdate_TVANIFCLIC.Clear
txtUpdate_TVANIFCLIC.AddItem "  client BIA"
txtUpdate_TVANIFCLIC.AddItem "D tiers crédoc"
txtUpdate_TVANIFCLIC.AddItem "G tiers change"
txtUpdate_TVANIFCLIC.AddItem "B BIC "
txtUpdate_TVANIFCLIC.AddItem "R tiers remdoc"

txtNIF_TVANIFCLIC.Clear
txtNIF_TVANIFCLIC.AddItem "  client BIA"
txtNIF_TVANIFCLIC.AddItem "B BIC"
txtNIF_TVANIFCLIC.AddItem "D tiers crédoc"
txtNIF_TVANIFCLIC.AddItem "R tiers remdoc"
txtNIF_TVANIFCLIC.AddItem "G tiers change"
Select Case currentUser.Unit
    Case "GSOP": txtNIF_TVANIFCLIC.ListIndex = 0
    Case "SOBI": txtNIF_TVANIFCLIC.ListIndex = 2
    Case "GDMP": txtNIF_TVANIFCLIC.ListIndex = 3
    Case Else: txtNIF_TVANIFCLIC.ListIndex = 0
End Select

txtNIF_TVANIFSTA.Clear
txtNIF_TVANIFSTA.AddItem " "
txtNIF_TVANIFSTA.AddItem "Annulé"
txtNIF_TVANIFSTA.AddItem "Validé"


'_____________________________________________________________________________Param
fraParam.Enabled = BIA_TVAPARAM_Aut.Consulter

fgParam_FormatString = fgParam.FormatString
cboParam_SQL.Clear
If BIA_TVAPARAM_Aut.Consulter Then
    cboParam_SQL.AddItem "1 - Liste routage : TVAFACSTA"
    cboParam_SQL.AddItem "2 - Liste CRE/ENG : TVACOMSTA"
End If
'___________________________________________________________________________

paramFacturation_Path = paramServer("\\Facturation\")
If paramEnvironnement = constTest Then paramFacturation_Path = Replace(paramFacturation_Path, "Facturation", "Test\Facturation")

Call rsZBASTAB0_Pays(arrPays(), arrPays_NB)
cmdReset


SSTab1.Visible = True



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

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
        If cmdSelect_SQL_K = "S" Then
            cmdPrint_SQL_S
        Else
            mnuPrint0_Facture = fraSelect_Update.Visible And BIA_TVAFAC_Aut.Xspécial
            Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
        End If
    Case 1:
        Me.PopupMenu mnuPrint1, vbPopupMenuLeftButton
    Case 2:
        Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
End Select
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
blnOk = Not fgSelect.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents
cmdSelect_Ok.Visible = False
fraSelect_Update.Visible = False
fraSelect_Options_1.Enabled = False
fgSelect.Clear
DoEvents
If blnOk Then
    cmdSelect_Ok.Caption = "Modifier les options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options_1.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options_1, fraSelect_Options_1.BackColor)
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_SQL
        Case "2": cmdSelect_SQL_2
        Case "3": cmdSelect_SQL_3
        Case "6": cmdSelect_SQL_6
        Case "7": cmdSelect_SQL_7: cboSelect_SQL.ListIndex = 0
        Case "8": cmdSelect_SQL_8: cboSelect_SQL.ListIndex = 0
        Case "9": cmdSelect_SQL_9: 'cboSelect_SQL.ListIndex = 0
        Case "S": cmdSelect_SQL_S

    End Select

    fgSelect.Enabled = True
Else
    cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
    cmdSelect_Ok.BackColor = &HE0FFFF
    fraSelect_Options_1.BackColor = &HE0FFFF  ' &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options_1, fraSelect_Options_1.BackColor)
    fgSelect.Visible = False
    fgSelect.Enabled = False
    fraSelect_Options_1.Enabled = True

End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok.Visible = True

End Sub


Private Sub cmdSelect_SQL()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
xWhere = " where TVAFACETA > 0 "


If chkSelect_TVAFACSTA.Value = "1" Then
    xWhere = xWhere & " and TVAFACSTA not in ('A','I','F','X')"
End If

X = Trim(Mid$(txtSelect_TVAFACSTA, 1, 1))
If X <> "" Then xWhere = xWhere & " and  TVAFACSTA = '" & X & "'"

Set rsSab = Nothing
Call DTPicker_Control(txtSelect_TVAFACDTR, wAmjMin)
Call DTPicker_Control(txtSelect_TVAFACDTR_Max, wAmjMax)

If chkSelect_TVAFACDTR = "1" Then
    xWhere = xWhere & " and TVAFACDTR >= " & wAmjMin - 19000000 _
                    & " and TVAFACDTR <= " & wAmjMax - 19000000
End If

X = Mid$(txtSelect_TVAFACCLIC, 1, 1)
If X <> "*" Then xWhere = xWhere & " and  TVAFACCLIC = '" & X & "'"
X = Trim(txtSelect_TVAFACCLI)
If X <> "" Then xWhere = xWhere & " and TVAFACCLI like '%" & X & "%'"

arrYTVAFAC0_SQL xWhere & " order by TVAFACCLIC,TVAFACCLI"
    
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_S()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim zStat As typeYTVACOM0
Dim curX As Currency

On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_S"): DoEvents

currentAction = "cmdSelect_SQL_S"
Call rsYTVACOM0_Init(zStat)
arrStat_K = 0
arrStat_Db1(0) = zStat: arrStat_Db2(0) = zStat: arrStat_Db3(0) = zStat
arrStat_Cr1(0) = zStat: arrStat_Cr2(0) = zStat: arrStat_Cr3(0) = zStat

Call DTPicker_Control(txtSelect_TVAFACDTR, wAmjMin)
Call DTPicker_Control(txtSelect_TVAFACDTR_Max, wAmjMax)

xWhere = " where TVACOMSTA <> 'A'" _
                & " and TVACOMDTR >= " & wAmjMin - 19000000 _
                & " and TVACOMDTR <= " & wAmjMax - 19000000
    
xSql = "select TVACOMSER,TVACOMOPE,TVACOMNAT,TVACOMCLIC,TVACOMMONE from " & paramIBM_Library_SABSPE & ".YTVACOM0 " _
    & xWhere & " order by TVACOMSER,TVACOMOPE,TVACOMNAT,TVACOMCLIC"
    
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    zStat.TVACOMSER = rsSab("TVACOMSER")
    zStat.TVACOMOPE = rsSab("TVACOMOPE")
    zStat.TVACOMNAT = rsSab("TVACOMNAT")
    curX = -CCur(rsSab("TVACOMMONE"))
    
    If zStat.TVACOMSER = arrStat_Db1(arrStat_K).TVACOMSER _
    And zStat.TVACOMOPE = arrStat_Db1(arrStat_K).TVACOMOPE _
    And zStat.TVACOMNAT = arrStat_Db1(arrStat_K).TVACOMNAT Then
    
    Else
        arrStat_K = arrStat_K + 1
        arrStat_Db1(arrStat_K) = zStat: arrStat_Db2(arrStat_K) = zStat: arrStat_Db3(arrStat_K) = zStat
        arrStat_Cr1(arrStat_K) = zStat: arrStat_Cr2(arrStat_K) = zStat: arrStat_Cr3(arrStat_K) = zStat
    End If
    If curX < 0 Then
        Select Case rsSab("TVACOMCLIC")
            Case "G": arrStat_Db3(arrStat_K).TVACOMMONE = arrStat_Db3(arrStat_K).TVACOMMONE + curX
                      arrStat_Db3(arrStat_K).TVACOMUPDS = arrStat_Db3(arrStat_K).TVACOMUPDS + 1
            Case "D": arrStat_Db2(arrStat_K).TVACOMMONE = arrStat_Db2(arrStat_K).TVACOMMONE + curX
                      arrStat_Db2(arrStat_K).TVACOMUPDS = arrStat_Db2(arrStat_K).TVACOMUPDS + 1
            Case Else: arrStat_Db1(arrStat_K).TVACOMMONE = arrStat_Db1(arrStat_K).TVACOMMONE + curX
                      arrStat_Db1(arrStat_K).TVACOMUPDS = arrStat_Db1(arrStat_K).TVACOMUPDS + 1
        End Select
    Else
        Select Case rsSab("TVACOMCLIC")
            Case "G": arrStat_Cr3(arrStat_K).TVACOMMONE = arrStat_Cr3(arrStat_K).TVACOMMONE + curX
                      arrStat_Cr3(arrStat_K).TVACOMUPDS = arrStat_Cr3(arrStat_K).TVACOMUPDS + 1
            Case "D": arrStat_Cr2(arrStat_K).TVACOMMONE = arrStat_Cr2(arrStat_K).TVACOMMONE + curX
                      arrStat_Cr2(arrStat_K).TVACOMUPDS = arrStat_Cr2(arrStat_K).TVACOMUPDS + 1
            Case Else: arrStat_Cr1(arrStat_K).TVACOMMONE = arrStat_Cr1(arrStat_K).TVACOMMONE + curX
                      arrStat_Cr1(arrStat_K).TVACOMUPDS = arrStat_Cr1(arrStat_K).TVACOMUPDS + 1
        End Select
End If
    
    rsSab.MoveNext

Loop

arrStat_K = arrStat_K + 1
Call rsYTVACOM0_Init(zStat)
arrStat_Db1(arrStat_K) = zStat: arrStat_Db2(arrStat_K) = zStat: arrStat_Db3(arrStat_K) = zStat
arrStat_Cr1(arrStat_K) = zStat: arrStat_Cr2(arrStat_K) = zStat: arrStat_Cr3(arrStat_K) = zStat

For K = 1 To arrStat_K - 1
    arrStat_Db1(arrStat_K).TVACOMUPDS = arrStat_Db1(arrStat_K).TVACOMUPDS + arrStat_Db1(K).TVACOMUPDS
    arrStat_Cr1(arrStat_K).TVACOMUPDS = arrStat_Cr1(arrStat_K).TVACOMUPDS + arrStat_Cr1(K).TVACOMUPDS
    arrStat_Db2(arrStat_K).TVACOMUPDS = arrStat_Db2(arrStat_K).TVACOMUPDS + arrStat_Db2(K).TVACOMUPDS
    arrStat_Cr2(arrStat_K).TVACOMUPDS = arrStat_Cr2(arrStat_K).TVACOMUPDS + arrStat_Cr2(K).TVACOMUPDS
    arrStat_Db3(arrStat_K).TVACOMUPDS = arrStat_Db3(arrStat_K).TVACOMUPDS + arrStat_Db3(K).TVACOMUPDS
    arrStat_Cr3(arrStat_K).TVACOMUPDS = arrStat_Cr3(arrStat_K).TVACOMUPDS + arrStat_Cr3(K).TVACOMUPDS
    
    arrStat_Db1(arrStat_K).TVACOMMONE = arrStat_Db1(arrStat_K).TVACOMMONE + arrStat_Db1(K).TVACOMMONE
    arrStat_Cr1(arrStat_K).TVACOMMONE = arrStat_Cr1(arrStat_K).TVACOMMONE + arrStat_Cr1(K).TVACOMMONE
    arrStat_Db2(arrStat_K).TVACOMMONE = arrStat_Db2(arrStat_K).TVACOMMONE + arrStat_Db2(K).TVACOMMONE
    arrStat_Cr2(arrStat_K).TVACOMMONE = arrStat_Cr2(arrStat_K).TVACOMMONE + arrStat_Cr2(K).TVACOMMONE
    arrStat_Db3(arrStat_K).TVACOMMONE = arrStat_Db3(arrStat_K).TVACOMMONE + arrStat_Db3(K).TVACOMMONE
    arrStat_Cr3(arrStat_K).TVACOMMONE = arrStat_Cr3(arrStat_K).TVACOMMONE + arrStat_Cr3(K).TVACOMMONE
Next K
fgSelect_Display_Stat
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_3()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL_3"
xWhere = " where TVAFACETA > 0 and (TVAFACSTA = '1' or TVAFACSTA = '2')"


Set rsSab = Nothing

X = Mid$(txtSelect_TVAFACCLIC, 1, 1)
If X <> "*" Then xWhere = xWhere & " and  TVAFACCLIC = '" & X & "'"
X = Trim(txtSelect_TVAFACCLI)
If X <> "" Then xWhere = xWhere & " and TVAFACCLI like '%" & X & "%'"

arrYTVAFAC0_SQL xWhere & " order by TVAFACCLIC,TVAFACCLI"
    
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdDétail_SQL()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdcmDétail_SQL"): DoEvents

currentAction = "cmdcmDétail_SQL"
xWhere = " where TVACOMETA > 0 "

If chkDétail_TVACOMSTA.Value = "1" Then
    xWhere = xWhere & " and TVACOMSTA not in ('A','I','F','X')"
End If

Set rsSab = Nothing
Call DTPicker_Control(txtDétail_TVACOMDTR, wAmjMin)
Call DTPicker_Control(txtDétail_TVACOMDTR_Max, wAmjMax)

If chkDétail_TVACOMDTR = "1" Then
    xWhere = xWhere & " and TVACOMDTR >= " & wAmjMin - 19000000 _
                    & " and TVACOMDTR <= " & wAmjMax - 19000000
End If

X = Trim(txtDétail_TVACOMSRVR)
If X <> "" Then xWhere = xWhere & " and TVACOMSRVR = '" & X & "'"

X = Mid$(txtDétail_TVACOMCLIC, 1, 1)
If X <> "*" Then xWhere = xWhere & " and TVACOMCLIC = '" & X & "'"
X = Trim(txtDétail_TVACOMCLI)
If X <> "" Then xWhere = xWhere & " and TVACOMCLI like '%" & X & "%'"

If cmdDétail_SQL_K <> 2 Then
    X = Trim(Mid$(txtDétail_TVACOMSTA, 1, 1))
    If X <> "" Then xWhere = xWhere & " and TVACOMSTA = '" & X & "'"
Else
    xWhere = xWhere & " and TVACOMSTA >= '0' and TVACOMSTA <= '9'"
End If

X = Trim(txtDétail_TVACOMOPE)
If X <> "" Then xWhere = xWhere & " and TVACOMOPE like '" & X & "%'"


arrYTVACOM0_SQL xWhere & " order by TVACOMSRVR,TVACOMCLIC,TVACOMCLI,TVACOMOPE,TVACOMDOS"
    
fgDétail_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdDétail_SQL_3()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdcmDétail_SQL_3"): DoEvents

currentAction = "cmdcmDétail_SQL_3"
xWhere = " where TVACOMSTA >= '0' and TVACOMSTA not in ('A','I','F','V','X')"

Set rsSab = Nothing
Call DTPicker_Control(txtDétail_TVACOMDTR, wAmjMin)
Call DTPicker_Control(txtDétail_TVACOMDTR_Max, wAmjMax)



arrYTVACOM0_SQL xWhere & " order by TVACOMSRVR,TVACOMCLIC,TVACOMCLI,TVACOMOPE,TVACOMDOS"
    
'fgDétail_Display
cmdPrintDétail_Liste3 "2"

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdDétail_SQL_6()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String

Dim Nb_Lu As Long, Nb_maj As Long
On Error GoTo Error_Handler

App_Debug = "cmdDétail_SQL_6"
Nb_Lu = 0: Nb_maj = 0
Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, App_Debug): DoEvents

currentAction = "cmdcmDétail_SQL_6"
xWhere = " where  TVACOMSTA not in ('A','I','F','X')"

Set rsSab = Nothing
xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVACOM0 " & xWhere & " order by TVACOMCLIP,TVACOMCLI"
Set rsSab = cnsab.Execute(xSql)
Call lstErr_AddItem(lstErr, cmdContext, "........."): DoEvents

Do While Not rsSab.EOF
    Nb_Lu = Nb_Lu + 1
    oldYTVACOM0.TVACOMCLIP = rsSab("TVACOMCLIP")
    If Trim(oldYTVACOM0.TVACOMCLIP) <> "" Then
        mTVANIFCLIT_Pays = TVANIFCLIT_Pays(oldYTVACOM0.TVACOMCLIP)
        If Not mTVANIFCLIT_Pays Then
            V = rsYTVACOM0_GetBuffer(rsSab, oldYTVACOM0)
    
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "frmBIA_TVACOM.cmdDétail_SQL_6"
            Else
                newYTVACOM0 = oldYTVACOM0
                newYTVACOM0.TVACOMSTA = "I"
                Call lstErr_ChangeLastItem(lstErr, cmdContext, oldYTVACOM0.TVACOMCLIP & oldYTVACOM0.TVACOMPIE): DoEvents
                V = cmdDétail_Update_Ok_Transaction
                Nb_maj = Nb_maj + 1
                Debug.Print oldYTVACOM0.TVACOMCLIP & oldYTVACOM0.TVACOMPIE
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                If Not IsNull(V) Then
                    MsgBox V, vbCritical, "Pièce : " & oldYTVACOM0.TVACOMPIE & " : " & oldYTVACOM0.TVACOMECR
                    Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
                End If
            End If
        End If
    End If
    rsSab.MoveNext

Loop
    
Call lstErr_AddItem(lstErr, cmdContext, "< NB maj / lu : " & Nb_maj & " / " & Nb_Lu): DoEvents


'____________________________________________________________________________________
' ignorer les commissions sur le NOSTRO SG 0011015
'_____________________________________________________________________________________


App_Debug = "cmdDétail_SQL_6_0011015"
Nb_Lu = 0: Nb_maj = 0
Set rsSab = Nothing

Call lstErr_AddItem(lstErr, cmdContext, App_Debug): DoEvents

currentAction = "cmdcmDétail_SQL_6_0011015"
xWhere = " where  TVACOMCLI = '0011015' and TVACOMCLIc = ' ' and TVACOMSTA not in ('A','I','F','X')"

Set rsSab = Nothing
xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVACOM0 " & xWhere
Set rsSab = cnsab.Execute(xSql)
Call lstErr_AddItem(lstErr, cmdContext, "........."): DoEvents

Do While Not rsSab.EOF
    Nb_Lu = Nb_Lu + 1
            V = rsYTVACOM0_GetBuffer(rsSab, oldYTVACOM0)
    
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "frmBIA_TVACOM.cmdDétail_SQL_6"
            Else
                newYTVACOM0 = oldYTVACOM0
                newYTVACOM0.TVACOMSTA = "I"
                Call lstErr_ChangeLastItem(lstErr, cmdContext, oldYTVACOM0.TVACOMCLIP & oldYTVACOM0.TVACOMPIE): DoEvents
                V = cmdDétail_Update_Ok_Transaction
                Nb_maj = Nb_maj + 1
                Debug.Print oldYTVACOM0.TVACOMCLIP & oldYTVACOM0.TVACOMPIE
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                If Not IsNull(V) Then
                    MsgBox V, vbCritical, "Pièce : " & oldYTVACOM0.TVACOMPIE & " : " & oldYTVACOM0.TVACOMECR
                    Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
                End If
            End If
    rsSab.MoveNext

Loop
    
Call lstErr_AddItem(lstErr, cmdContext, "< NB maj / lu : " & Nb_maj & " / " & Nb_Lu): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_6()
Dim V
Dim X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim lDate As Long
On Error GoTo Error_Handler

fgDétail.Visible = False
Set rsSab = Nothing
currentAction = "cmdSelect_SQL_6"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents

currentAction = "cmdSelect_SQL_6"
xAnd = ""
Call DTPicker_Control(txtSelect_TVAFACDTR, wAmjMin)
Call DTPicker_Control(txtSelect_TVAFACDTR_Max, wAmjMax)

If chkSelect_TVAFACDTR = "1" Then
    xAnd = xAnd & " and TVACOMDTR >=" & wAmjMin - 19000000 _
                & " and TVACOMDTR <=" & wAmjMax - 19000000
End If

xWhere = "where TVACOMSTA not in ('V','X','F','I','A') " & xAnd

xSql = "select count(*) as Tally from " & paramIBM_Library_SABSPE & ".YTVACOM0 " & xWhere
Set rsSab = cnsab.Execute(xSql)
K = rsSab("Tally")
If K <> 0 Then
    Call MsgBox("Il y a encore " & K & " opérations non validées pour cette période", vbCritical, "Regroupement des commissions")
    Exit Sub
End If


X = Mid$(txtSelect_TVAFACCLIC, 1, 1)
If X <> "*" Then xAnd = xAnd & " and  TVACOMCLIC = '" & X & "'"

X = Trim(txtSelect_TVAFACCLI)
If X <> "" Then xAnd = xAnd & " and TVACOMCLI like '%" & X & "%'"

xWhere = "where TVACOMSTA = 'V' " & xAnd


arrYTVACOM0_SQL xWhere & " order by TVACOMETA, TVACOMCLIC, TVACOMCLI, TVACOMCLIP"
If arrYTVACOM0_Nb = 0 Then
    V = "aucune commission valide à traiter"
    GoTo Error_MsgBox
Else
    cmdSelect_SQL_6_Regroupement
End If

    
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub cmdSelect_SQL_7()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim mTVAFACDTR As Long
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_7"): DoEvents
currentAction = "cmdSelect_SQL_7"
X = UCase$(Trim(Printer.Devicename))
If InStr(1, X, "ADOBE PDF") <= 0 Then
    Call MsgBox("Choisir l'imprimante 'Adobe PDF'", vbCritical, "Archivage des factures émises")
    Exit Sub
End If
Call DTPicker_Control(txtSelect_TVAFACDTR, wAmjMin)
X = MsgBox("confirmez-vous la sélection code état 0,1,2,V : ", vbQuestion & vbYesNo, "Facturation définitive")
If X <> vbYes Then Exit Sub

X = MsgBox("confirmez-vous la date de facturation : " & dateImp(wAmjMin), vbQuestion & vbYesNo, "Facturation définitive")
If X <> vbYes Then Exit Sub

mTVAFACDTR = Val(wAmjMin) - 19000000
X = paramFacturation_Path & Mid$(wAmjMin, 1, 4)
If Not msFileSystem.FolderExists(X) Then MkDir X
paramFacturation_Path_AAAA = paramFacturation_Path & Mid$(wAmjMin, 1, 4) & "\"


'xWhere = " where TVAFACETA > 0 and TVAFACSTA = 'V'"
xWhere = " where TVAFACETA > 0 and TVAFACSTA in ('0','1','2','V')"

X = Trim(txtSelect_TVAFACCLI)
If X <> "" Then xWhere = xWhere & " and TVAFACCLI like '%" & X & "%'"

arrYTVAFAC0_SQL xWhere & " order by TVAFACFACN"

fgSelect_Display

X = paramFacturation_Path & "log\BIA_TVAFAC_" & DSys & "_" & time_Hms & ".log"

Open X For Append As #1
'=============================================================
Print #1, Time, arrYTVAFAC0_Nb & " factures à émettre au " & mTVAFACDTR
Print #1, "------------------------------------------------------------"

For I = 1 To arrYTVAFAC0_Nb
    oldYTVAFAC0 = arrYTVAFAC0(I)
'____________________________________________________________________
     Print #1, Time, oldYTVAFAC0.TVAFACFACN & " > mise à jour TVAFACSTA = 'F'"
    '----------------------------------------------
    newYTVAFAC0 = oldYTVAFAC0
    newYTVAFAC0.TVAFACSTA = "F"
    newYTVAFAC0.TVAFACDTR = mTVAFACDTR
    V = cmdSelect_Update_Ok_Transaction
    If Not IsNull(V) Then GoTo Error_MsgBox
     Print #1, Time, oldYTVAFAC0.TVAFACFACN & " < mise à jour terminée"
     Print #1, Time, oldYTVAFAC0.TVAFACFACN & " > Impression"
    '----------------------------------------------
    oldYTVAFAC0 = newYTVAFAC0
    V = cmdPrint_Facture
    If Not IsNull(V) Then GoTo Error_MsgBox
    '----------------------------------------------
'____________________________________________________________________
Next I
Print #1, Time, " < Facturation : traitement terminé"
Print #1, "==============================================="
Close 1
'=============================================================
Set rsSab = Nothing
Call MsgBox("Traitement terminé : " & arrYTVAFAC0_Nb & " factures", vbInformation, "Archivage des factures émises")

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
     Print #1, Time, oldYTVAFAC0.TVAFACFACN & " ? Facturation erreur : " & V
    '----------------------------------------------
   Close 1

End Sub

Private Sub cmdSelect_SQL_8()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim mTVAFACDTR As Long
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_7"): DoEvents
currentAction = "cmdSelect_SQL_7"
X = UCase$(Trim(Printer.Devicename))
If InStr(1, X, "ADOBE PDF") > 0 Then
    Call MsgBox("Choisir une imprimante non 'Adobe PDF'", vbCritical, "Archivage des factures émises")
    Exit Sub
'    Call MsgBox("TEST désactivé", vbCritical, "Archivage des factures émises")
End If
Call DTPicker_Control(txtSelect_TVAFACDTR, wAmjMin)

X = MsgBox("confirmez-vous la date de facturation : " & dateImp(wAmjMin), vbQuestion & vbYesNo, "Facturation définitive")
If X <> vbYes Then Exit Sub


X = InputBox("N° de la dernière facture imprimée (sinon 0)", "Reprise de 'l'impression des factures", "0")
Print_Reprise_TVAFACFACN = Val(X)
If Print_Reprise_TVAFACFACN = 0 Then
    blnPrint_Reprise_Ok = True
Else
    blnPrint_Reprise_Ok = False
End If

mTVAFACDTR = Val(wAmjMin) - 19000000
X = paramFacturation_Path & Mid$(wAmjMin, 1, 4)
If Not msFileSystem.FolderExists(X) Then MkDir X
paramFacturation_Path_AAAA = paramFacturation_Path & Mid$(wAmjMin, 1, 4) & "\"


xWhere = " where TVAFACETA > 0 and TVAFACSTA = 'F' and TVAFACDTR = " & mTVAFACDTR

X = Trim(txtSelect_TVAFACCLI)
If X <> "" Then
    xAnd = xWhere & " and TVAFACCLI like '%" & X & "%'"
    arrYTVAFAC0_SQL xAnd & " order by TVAFACFACN"
    If arrYTVAFAC0_Nb > 0 Then cmdSelect_SQL_8_Print
Else
'Routage différent du code RESPONSABLE
    xAnd = xWhere & " and TVAFACCLIC = ' '"
    cmdSelect_SQL_8_ZCLIENA0_Routage xAnd
    If arrYTVAFAC0_Nb > 0 Then cmdSelect_SQL_8_Print
'Routage code RESPONSABLE
    xAnd = xWhere & " and TVAFACCLIC = ' '"
    cmdSelect_SQL_8_ZCLIENA0 xAnd
    If arrYTVAFAC0_Nb > 0 Then cmdSelect_SQL_8_Print
'Routage TIERS
    xAnd = xWhere & " and TVAFACCLIC <> ' '"
    arrYTVAFAC0_SQL xAnd & " order by TVAFACFACN"
    If arrYTVAFAC0_Nb > 0 Then cmdSelect_SQL_8_Print
End If

'=============================================================
Set rsSab = Nothing
Call MsgBox("Traitement terminé  ", vbInformation, "Impression des factures émises")

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
     Print #1, Time, oldYTVAFAC0.TVAFACFACN & " ? Facturation erreur : " & V
    '----------------------------------------------
   Close 1

End Sub
Private Sub cmdSelect_SQL_9()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim mTVAFACDTR As Long
Dim Nb As Long, Nb_HT As Long, Nb_CLIT As Long
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_9"): DoEvents
currentAction = "cmdSelect_SQL_9"
If Not IsEmpty(Xprt_Previous) Then Set Xprt_Previous = XPrt
'If Not IsEmpty(XPrt) Then Set Xprt_Previous = XPrt
'Printer_PDF
X = UCase$(Trim(Printer.Devicename))
If InStr(1, X, "ADOBE PDF") = 0 Then
    X = MsgBox("Attention l'imprimante n'est pas'Adobe PDF', voulez-vous continuer ?", vbQuestion & vbYesNo, "Etat de déclaration de la TVA intracommunautaire aux douanes")
    If X <> vbYes Then Exit Sub

    'Call MsgBox("Choisir une imprimante 'Adobe PDF'", vbCritical, "Etat de déclaration de la TVA intracommunautaire aux douanes")
    'Exit Sub
'    Call MsgBox("TEST désactivé", vbCritical, "Archivage des factures émises")
End If
Call DTPicker_Control(txtSelect_TVAFACDTR, wAmjMin)

X = MsgBox("confirmez-vous la date de facturation : " & dateImp(wAmjMin), vbQuestion & vbYesNo, "Etat de déclaration de la TVA intracommunautaire aux douanes")
If X <> vbYes Then Exit Sub


mTVA_DES_File = paramFacturation_Path & "TVA_DES\TVA_DES_" & wAmjMin & " en date du " & DSys & "_" & time_Hms & ".xml"

Open mTVA_DES_File For Output As #1
'=============================================================
Print #1, "<?xml version=" & Asc34 & "1.0" & Asc34 & " encoding=" & Asc34 & "UTF-8" & Asc34 & "?>"
Print #1, "<fichier_des>"
Print #1, "        <declaration_des>"
Print #1, "               <num_des>" & "00001" & "</num_des>"
X = Replace(paramSOC_TVA_Intracommunautaire, " ", "")
Print #1, "               <num_tvaFr>" & X & "</num_tvaFr>"
Print #1, "               <mois_des>" & Mid$(wAmjMin, 5, 2) & "</mois_des>"
Print #1, "               <an_des>" & Mid$(wAmjMin, 1, 4) & "</an_des>"
'_________________________________________________________________________

prtBIA_TVAFAC_Open 9, "Etat de déclaration de la TVA intracommunautaire aux douanes - (relevé du " & dateImp10(wAmjMin) & " )"



mTVAFACDTR = Val(wAmjMin) - 19000000
mTVAFACMUE_EXO_Total = 0: Nb = 0: Nb_HT = 0: Nb_CLIT = 0
mTVAFACMUE_HT_Total = 0
mTVAFACMUE_TVA_Total = 0
mTVAFACMUE_CLIT_Total = 0

xWhere = " where TVAFACETA > 0 and TVAFACSTA = 'F' and TVAFACCLIP <> 'FR' and TVAFACDTR = " & mTVAFACDTR

arrYTVAFAC0_SQL xWhere & " order by TVAFACCLIP , TVAFACFACN"

For K = 1 To arrYTVAFAC0_Nb
    oldYTVAFAC0 = arrYTVAFAC0(K)
    mTVANIFCLIT_Pays = TVANIFCLIT_Pays(oldYTVAFAC0.TVAFACCLIP)
    If mTVAFACCLIP_Code_Fiscal = "4" Or mTVAFACCLIP_Code_Fiscal = "5" Then
    
        Call cmdSelect_SQL_9_Detail
 '_________________________________________________________________________________________
        If mTVAFACMUE_EXO <> 0 Then
            prtBIA_TVAFAC_NewLine 9
            
            If Trim(oldYTVAFAC0.TVAFACCLIT) = "" Then
                newYTVAFAC0 = oldYTVAFAC0
                cmdSelect_SQL_6_Regroupement_NIF
                oldYTVAFAC0.TVAFACCLIT = newYTVAFAC0.TVAFACCLIT
            End If
            If Trim(oldYTVAFAC0.TVAFACCLIT) = "" Then
               XPrt.ForeColor = vbMagenta
                 Nb_CLIT = Nb_CLIT + 1
                mTVAFACMUE_CLIT_Total = mTVAFACMUE_CLIT_Total + mTVAFACMUE_EXO
           Else
                XPrt.ForeColor = vbBlue
                Nb = Nb + 1
                mTVAFACMUE_EXO_Total = mTVAFACMUE_EXO_Total + mTVAFACMUE_EXO
            
                X = Format$(Nb, "### ###")
                XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
                XPrt.Print X;
                Print #1, "               <ligne_des>"
                Print #1, "                     <numlin_des>" & Format(Nb, "000000") & "</numlin_des>"
                Print #1, "                     <valeur>" & Trim(Format(-mTVAFACMUE_EXO, "############0")) & "</valeur>"
                Print #1, "                     <partner_des>" & Trim(oldYTVAFAC0.TVAFACCLIT) & "</partner_des>"
                Print #1, "               </ligne_des>"
            End If
            
            XPrt.CurrentX = prtMinX + 50: XPrt.Print oldYTVAFAC0.TVAFACCLIC & " " & oldYTVAFAC0.TVAFACCLI;
            V = sqlTVACOMCLI(oldYTVAFAC0.TVAFACCLIC, oldYTVAFAC0.TVAFACCLI, xZCLIENA0, xZADRESS0)
            XPrt.CurrentX = prtMinX + 900: XPrt.Print Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2);

            XPrt.CurrentX = prtMinX + 7450: XPrt.Print dateIBM10(oldYTVAFAC0.TVAFACDTR, True);
            XPrt.CurrentX = prtMinX + 9800: XPrt.Print oldYTVAFAC0.TVAFACCLIP;
            XPrt.CurrentX = prtMinX + 10300: XPrt.Print oldYTVAFAC0.TVAFACCLIT;
            X = Format$((oldYTVAFAC0.TVAFACFACN), "### ### ### ###")
            XPrt.CurrentX = prtMinX + 9650 - XPrt.TextWidth(X)
            XPrt.Print X;
   
            X = Format$(Abs(mTVAFACMUE_EXO), "### ### ### ###.00")
            If mTVAFACMUE_EXO > 0 Then XPrt.ForeColor = vbRed
            XPrt.CurrentX = prtMinX + 13800 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
 '_________________________________________________________________________________________

       If mTVAFACMUE_HT <> 0 Then
            prtBIA_TVAFAC_NewLine 9
            XPrt.ForeColor = RGB(64, 128, 64)
            Nb_HT = Nb_HT + 1
            mTVAFACMUE_HT_Total = mTVAFACMUE_HT_Total + mTVAFACMUE_HT
            mTVAFACMUE_TVA_Total = mTVAFACMUE_TVA_Total + mTVAFACMUE_TVA
            
            XPrt.CurrentX = prtMinX + 50: XPrt.Print oldYTVAFAC0.TVAFACCLIC & " " & oldYTVAFAC0.TVAFACCLI;
            V = sqlTVACOMCLI(oldYTVAFAC0.TVAFACCLIC, oldYTVAFAC0.TVAFACCLI, xZCLIENA0, xZADRESS0)
            XPrt.CurrentX = prtMinX + 900: XPrt.Print Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2);

            XPrt.CurrentX = prtMinX + 7450: XPrt.Print dateIBM10(oldYTVAFAC0.TVAFACDTR, True);
            XPrt.CurrentX = prtMinX + 9800: XPrt.Print oldYTVAFAC0.TVAFACCLIP;
            XPrt.CurrentX = prtMinX + 10300: XPrt.Print oldYTVAFAC0.TVAFACCLIT;
            X = Format$((oldYTVAFAC0.TVAFACFACN), "### ### ### ###")
            XPrt.CurrentX = prtMinX + 9650 - XPrt.TextWidth(X)
            XPrt.Print X;
   
            X = Format$(Abs(mTVAFACMUE_HT), "### ### ### ###.00")
            If mTVAFACMUE_HT > 0 Then XPrt.ForeColor = vbRed
            XPrt.CurrentX = prtMinX + 13800 - XPrt.TextWidth(X)
            XPrt.Print X;
            X = Format$(Abs(mTVAFACMUE_TVA), "### ### ### ###.00")
            XPrt.CurrentX = prtMinX + 15100 - XPrt.TextWidth(X)
            XPrt.Print X;
       End If
    End If
 
 '_________________________________________________________________________________________
Next K

'=============================================================

XPrt.DrawWidth = 5
prtBIA_TVAFAC_NewLine 9
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50
If mTVAFACMUE_EXO_Total > 0 Then
   XPrt.ForeColor = vbRed
Else
    XPrt.ForeColor = vbBlue
End If
XPrt.CurrentX = prtMinX + 50: XPrt.Print "prestations à déclarer au service des douanes, (" & Nb & " lignes)";
X = Format$(Abs(mTVAFACMUE_EXO_Total), "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 13800 - XPrt.TextWidth(X)
XPrt.Print X;
prtBIA_TVAFAC_NewLine 9
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

If Nb_CLIT > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.ForeColor = vbMagenta
    
    XPrt.CurrentX = prtMinX + 50: XPrt.Print "prestations à déclarer au service des douanes, mais l'identifiant TVA est manquant, (" & Nb_CLIT & " lignes)";
    X = Format$(Abs(mTVAFACMUE_CLIT_Total), "### ### ### ###.00")
    If mTVAFACMUE_HT_Total > 0 Then XPrt.ForeColor = vbRed
    XPrt.CurrentX = prtMinX + 13800 - XPrt.TextWidth(X)
    XPrt.Print X;
    prtBIA_TVAFAC_NewLine 9
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
End If

If Nb_HT > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.ForeColor = RGB(64, 128, 64)
    
    XPrt.CurrentX = prtMinX + 50: XPrt.Print "prestations déjà taxées à la TVA française, (" & Nb_HT & " lignes)";
    X = Format$(Abs(mTVAFACMUE_HT_Total), "### ### ### ###.00")
    If mTVAFACMUE_HT_Total > 0 Then XPrt.ForeColor = vbRed
    XPrt.CurrentX = prtMinX + 13800 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(Abs(mTVAFACMUE_TVA_Total), "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 15100 - XPrt.TextWidth(X)
    XPrt.Print X;
    prtBIA_TVAFAC_NewLine 9
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
End If
prtBIA_TVAFAC_Form_9_Col

prtBIA_TVAFAC_Close 9
'_________________________________________________________________________
Print #1, "        </declaration_des>"
Print #1, "</fichier_des>"
Close #1

cmdSendMail_Douanes mTVA_DES_File & ";" & prtIMP_PDF_FileName

If Not IsEmpty(Xprt_Previous) Then Set XPrt = Xprt_Previous

'_________________________________________________________________________
Set rsSab = Nothing
Call MsgBox("Traitement terminé  ", vbInformation, "Etat de déclaration de la TVA intracommunautaire aux douanes")

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
'_______________________________________________________________________________________

End Sub
Private Sub cmdSelect_SQL_9_Detail()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim mTVAFACDTR As Long, xCur As Currency
On Error GoTo Error_Handler

Call lstErr_ChangeLastItem(lstErr, cmdContext, "Facture : " & oldYTVAFAC0.TVAFACFACN): DoEvents
currentAction = "cmdSelect_SQL_9_Detail"

mTVAFACMUE_EXO = 0
mTVAFACMUE_HT = 0
mTVAFACMUE_TVA = 0

xWhere = " where TVACOMFACN = " & oldYTVAFAC0.TVAFACFACN & " and TVACOMCOMT in ('N' , 'R')" '  and TVACOMMTVE =0"

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVACOM0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    xCur = rsSab("TVACOMMTVE")
    If xCur = 0 Then
        mTVAFACMUE_EXO = mTVAFACMUE_EXO + rsSab("TVACOMMONE")
    Else
        mTVAFACMUE_HT = mTVAFACMUE_HT + rsSab("TVACOMMONE")
        mTVAFACMUE_TVA = mTVAFACMUE_TVA + xCur
   End If
    rsSab.MoveNext

Loop


'=============================================================
Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub




Private Sub cmdSelect_SQL_8_Print()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim mTVAFACDTR As Long
On Error GoTo Error_Handler


X = paramFacturation_Path & "log\BIA_TVAFAC_" & DSys & "_" & time_Hms & ".log"

Open X For Append As #1
'=============================================================
Print #1, Time, arrYTVAFAC0_Nb & " impression papier des factures du " & wAmjMin
Print #1, "----------------------------------------------"

For I = 1 To arrYTVAFAC0_Nb
    oldYTVAFAC0 = arrYTVAFAC0(I)
    If Not blnPrint_Reprise_Ok Then
        If arrYTVAFAC0(I).TVAFACFACN = Print_Reprise_TVAFACFACN Then blnPrint_Reprise_Ok = True
    End If
    If blnPrint_Reprise_Ok Then
        V = cmdPrint_Facture
        Print #1, Time, xZCLIENA0.CLIENARES; oldYTVAFAC0.TVAFACCLIC; oldYTVAFAC0.TVAFACCLI; oldYTVAFAC0.TVAFACFACN & " = Impression"
    '----------------------------------------------
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If
'____________________________________________________________________
'____________________________________________________________________
Next I
Print #1, Time, " < impression papier des factures : traitement terminé"
Print #1, "==============================================="
Close 1
'=============================================================

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
     Print #1, Time, oldYTVAFAC0.TVAFACFACN & " ? Facturation erreur : " & V
    '----------------------------------------------
   Close 1

End Sub



Private Sub cmdSelect_SQL_2()
Dim V
Dim xWhere As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_2"): DoEvents

currentAction = "cmdSelect_SQL_2"
    
xWhere = " where TVAFACETA > 0 and  TVAFACSTA = '0'"


arrYTVAFAC0_SQL xWhere & " order by TVAFACCLIC,TVAFACCLI"
    
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub arrYTVACOM0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYTVACOM0(101)
arrYTVACOM0_Max = 100: arrYTVACOM0_Nb = 0
fraDétail_Update.Visible = False
Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVACOM0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYTVACOM0_GetBuffer(rsSab, xYTVACOM0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmBIA_TVACOM.fgselect_Display"
        '' Exit Sub
     Else
         arrYTVACOM0_Nb = arrYTVACOM0_Nb + 1
         If arrYTVACOM0_Nb > arrYTVACOM0_Max Then
             arrYTVACOM0_Max = arrYTVACOM0_Max + 50
             ReDim Preserve arrYTVACOM0(arrYTVACOM0_Max)
         End If
         
         arrYTVACOM0(arrYTVACOM0_Nb) = xYTVACOM0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub



Private Sub arrYTVAFAC0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYTVAFAC0(101)
arrYTVAFAC0_Max = 100: arrYTVAFAC0_Nb = 0
fraSelect_Update.Visible = False
Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVAFAC0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYTVAFAC0_GetBuffer(rsSab, xYTVAFAC0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmBIA_TVAFAC.fgselect_Display"
        '' Exit Sub
     Else
         arrYTVAFAC0_Nb = arrYTVAFAC0_Nb + 1
         If arrYTVAFAC0_Nb > arrYTVAFAC0_Max Then
             arrYTVAFAC0_Max = arrYTVAFAC0_Max + 50
             ReDim Preserve arrYTVAFAC0(arrYTVAFAC0_Max)
         End If
         
         arrYTVAFAC0(arrYTVAFAC0_Nb) = xYTVAFAC0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_8_ZCLIENA0(xWhere As String)
Dim V, K As Integer, blnOk As Boolean
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYTVAFAC0(101)
arrYTVAFAC0_Max = 100: arrYTVAFAC0_Nb = 0
fraSelect_Update.Visible = False
Set rsSab = Nothing
xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVAFAC0 inner join " & paramIBM_Library_SAB & ".ZCLIENA0 on  CLIENACLI = TVAFACCLI" & xWhere & "order by CLIENARES,CLIENACLI"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYTVAFAC0_GetBuffer(rsSab, xYTVAFAC0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmBIA_TVAFAC.fgselect_Display"
        '' Exit Sub
     Else
        blnOk = True
        For K = 1 To selYTVAFAC0_Nb
            If xYTVAFAC0.TVAFACCLI = selYTVAFAC0(K).TVAFACCLI Then
                blnOk = False
            End If
        Next K
        If blnOk Then
            arrYTVAFAC0_Nb = arrYTVAFAC0_Nb + 1
            If arrYTVAFAC0_Nb > arrYTVAFAC0_Max Then
                arrYTVAFAC0_Max = arrYTVAFAC0_Max + 50
                ReDim Preserve arrYTVAFAC0(arrYTVAFAC0_Max)
            End If
         
            arrYTVAFAC0(arrYTVAFAC0_Nb) = xYTVAFAC0
        End If
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_8_ZCLIENA0_Routage(xWhere As String)
Dim V, K As Integer
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYTVAFAC0(101)
arrYTVAFAC0_Max = 100: arrYTVAFAC0_Nb = 0
fraSelect_Update.Visible = False
Set rsSab = Nothing
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 inner join " & paramIBM_Library_SABSPE & ".YTVAFAC0 on  BIATABK2 = TVAFACCLI" & xWhere & "  and BIATABID = 'TVAFACSTA' and BIATABK1 = 'CLIENARES' order by BIATABTXT,TVAFACCLI"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYTVAFAC0_GetBuffer(rsSab, xYTVAFAC0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmBIA_TVAFAC.fgselect_Display"
        '' Exit Sub
     Else
         arrYTVAFAC0_Nb = arrYTVAFAC0_Nb + 1
         If arrYTVAFAC0_Nb > arrYTVAFAC0_Max Then
             arrYTVAFAC0_Max = arrYTVAFAC0_Max + 50
             ReDim Preserve arrYTVAFAC0(arrYTVAFAC0_Max)
         End If
         
         arrYTVAFAC0(arrYTVAFAC0_Nb) = xYTVAFAC0
    End If
    rsSab.MoveNext

Loop

ReDim selYTVAFAC0(arrYTVAFAC0_Max)
selYTVAFAC0_Nb = arrYTVAFAC0_Nb
For K = 1 To selYTVAFAC0_Nb
    selYTVAFAC0(K) = arrYTVAFAC0(K)
Next K
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYTVANIF0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVANIF0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYTVANIF0_GetBuffer(rsSab, xYTVANIF0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmBIA_TVANIF.fgselect_Display"
        '' Exit Sub
     Else
         arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
         If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
             arrYTVANIF0_Max = arrYTVANIF0_Max + 50
             ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
         End If
         
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYBIATAB0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmBIA_TVAParam.fgselect_Display"
        '' Exit Sub
     Else
         arrYBIATAB0_Nb = arrYBIATAB0_Nb + 1
         If arrYBIATAB0_Nb > arrYBIATAB0_Max Then
             arrYBIATAB0_Max = arrYBIATAB0_Max + 50
             ReDim Preserve arrYBIATAB0(arrYBIATAB0_Max)
         End If
         
         arrYBIATAB0(arrYBIATAB0_Nb) = xYBIATAB0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_Update_Annuler_Click()
Dim V, X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Annulation de la Facture :" & oldYTVAFAC0.TVAFACFACN): DoEvents

arrYTVACOM0_SQL " where TVACOMFACN = " & oldYTVAFAC0.TVAFACFACN & " order by TVACOMFACN"
    
fgDétail_Display

X = MsgBox("Confirmer l'annulation de cette facture et la restauration des  " & arrYTVACOM0_Nb & " commissions attachées?", vbYesNo + vbQuestion + vbDefaultButton2, "Annulation de la Facture :" & oldYTVAFAC0.TVAFACFACN)
If X = vbNo Then SSTab1.Tab = 0: GoTo Exit_sub



    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdSelect_Update_Annuler_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    If IsNull(V) Then
        arrYTVAFAC0(arrYTVAFAC0_Index) = newYTVAFAC0
        xYTVAFAC0 = newYTVAFAC0
        fgSelect_DisplayLine arrYTVAFAC0_Index
        fraSelect_Update.Visible = False

    Else
        MsgBox V, vbCritical, Me.Name & " : cmdSelect_Update_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Update_Détail_Display_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Facture :" & oldYTVAFAC0.TVAFACFACN): DoEvents

arrYTVACOM0_SQL " where TVACOMFACN = " & oldYTVAFAC0.TVAFACFACN & " order by TVACOMFACN"
    
fgDétail_Display


Me.Enabled = True: Me.MousePointer = 0



End Sub


Private Sub cmdSelect_Update_Ok_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(fraSelect_Update_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdSelect_Update_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYTVAFAC0(arrYTVAFAC0_Index) = newYTVAFAC0
        xYTVAFAC0 = newYTVAFAC0
        fgSelect_DisplayLine arrYTVAFAC0_Index
        fraSelect_Update.Visible = False
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdSelect_Update_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Update_Quit_Click()
fraSelect_Update.Visible = False
End Sub

Private Sub fgDétail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If Y <= fgDétail.RowHeightMin Then
        Select Case fgDétail.Col
            Case 0: fgDétail_SortX 0
            Case 1: fgDétail_SortX 1
            Case 2: fgDétail_SortX 2
            Case 3: fgDétail_SortX 3
            Case 4: fgDétail_SortX 4
            Case 5: fgDétail_Sort1 = 5: fgDétail_Sort2 = 5: fgDétail_Sort
            Case 6: fgDétail_Sort1 = 6: fgDétail_Sort2 = 6: fgDétail_Sort
            Case 7: fgDétail_Sort1 = 7: fgDétail_Sort2 = 7: fgDétail_Sort
            Case 8: fgDétail_Sort1 = 8: fgDétail_Sort2 = 8: fgDétail_Sort
            Case 9: fgDétail_Sort1 = 9: fgDétail_Sort2 = 9: fgDétail_Sort
            Case 10: fgDétail_Sort1 = 10: fgDétail_Sort2 = 10: fgDétail_Sort
            Case 11:  fgDétail_SortX 11
            Case 12: fgDétail_SortX 12
           Case fgDétail_arrIndex:  fgDétail_SortX fgDétail_arrIndex
        End Select
Else
    If fgDétail.Rows > 1 Then
        fgDétail.Col = fgDétail_arrIndex:  arrYTVACOM0_Index = CLng(fgDétail.Text)
        Call fgDétail_Color(fgDétail_RowClick, MouseMoveUsr.BackColor, fgDétail_ColorClick)
        xYTVACOM0 = arrYTVACOM0(arrYTVACOM0_Index)
        oldYTVACOM0 = xYTVACOM0
        fraDétail_Display
   End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgNIF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If Y <= fgNIF.RowHeightMin Then
        Select Case fgNIF.Col
            Case 0: fgNIF_Sort1 = 0: fgNIF_Sort2 = 0: fgNIF_Sort
            Case 1: fgNIF_Sort1 = 1: fgNIF_Sort2 = 1: fgNIF_Sort
            Case 2: fgNIF_Sort1 = 2: fgNIF_Sort2 = 2: fgNIF_Sort
            Case 3: fgNIF_Sort1 = 3: fgNIF_Sort2 = 3: fgNIF_Sort
        End Select
Else
    If fgNIF.Rows > 1 Then
        fgNIF.Col = fgNIF_arrIndex:  arrYTVANIF0_Index = CLng(fgNIF.Text)
        Call fgNIF_Color(fgNIF_RowClick, MouseMoveUsr.BackColor, fgNIF_ColorClick)
        xYTVANIF0 = arrYTVANIF0(arrYTVANIF0_Index)
        oldYTVANIF0 = xYTVANIF0
        fraNIF_Display
   End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgParam_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If Y <= fgParam.RowHeightMin Then
        Select Case fgParam.Col
            Case 0: fgParam_Sort1 = 0: fgParam_Sort2 = 0: fgParam_Sort
            Case 1: fgParam_Sort1 = 1: fgParam_Sort2 = 1: fgParam_Sort
            Case 2: fgParam_Sort1 = 2: fgParam_Sort2 = 2: fgParam_Sort
            Case 3: fgParam_Sort1 = 3: fgParam_Sort2 = 3: fgParam_Sort
        End Select
Else
    If fgParam.Rows > 1 Then
        fgParam.Col = fgParam_arrIndex:  arrYBIATAB0_Index = CLng(fgParam.Text)
        Call fgParam_Color(fgParam_RowClick, MouseMoveUsr.BackColor, fgParam_ColorClick)
        xYBIATAB0 = arrYBIATAB0(arrYBIATAB0_Index)
        oldYBIATAB0 = xYBIATAB0
        If Trim(oldYBIATAB0.BIATABID) <> "$Supprimé" Then fraParam_Display
   End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
            Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
            Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_SortX 4
            Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX 5
            Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_SortX 6
            Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_SortX 7
            Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_SortX 8
           Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
        End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex:  arrYTVAFAC0_Index = CLng(fgSelect.Text)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        xYTVAFAC0 = arrYTVAFAC0(arrYTVAFAC0_Index)
        oldYTVAFAC0 = xYTVAFAC0
        fraSelect_Display
   End If
End If
Me.Enabled = True: Me.MousePointer = 0
End Sub


Public Function fraSelect_Update_Control()
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
newYTVAFAC0 = oldYTVAFAC0
If xYTVAFAC0.TVAFACSTA = "0" Or xYTVAFAC0.TVAFACSTA = "1" Then

    X = Trim(txtUpdate_TVAFACCLIT)
    If Len(X) < 3 Then
            blnUpdate_Control = False
            txtUpdate_TVAFACCLIT.BackColor = errUsr.BackColor
            Call lstErr_AddItem(lstErr, cmdContext, "?_________préciser le code TVA intracommunautaire")
    Else
        If Mid$(X, 1, 2) <> oldYTVAFAC0.TVAFACCLIP Then
            blnUpdate_Control = False
            txtUpdate_TVAFACCLIT.BackColor = errUsr.BackColor
            Call lstErr_AddItem(lstErr, cmdContext, "?_________code pays incompatible")
    End If
    End If
    newYTVAFAC0.TVAFACCLIT = X
End If
If blnUpdate_Control Then
    fraSelect_Update_Control = Null
    newYTVAFAC0.TVAFACSTA = "V"
Else
    fraSelect_Update_Control = "?_________fraSelect_Update_Control"
End If
End Function

Public Function fraNIF_Update_Control()
Dim V
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
newYTVANIF0 = oldYTVANIF0

X = Replace(txtUpdate_TVANIFCLIT, " ", "")
If oldYTVANIF0.TVANIFCLIP <> "" Then
    If Mid$(X, 1, 2) <> oldYTVANIF0.TVANIFCLIP Then
        If oldYTVANIF0.TVANIFCLIP = "GR" And Mid$(X, 1, 2) = "EL" Then
        Else
            blnUpdate_Control = False
            txtUpdate_TVANIFCLIT.BackColor = errUsr.BackColor
            Call lstErr_AddItem(lstErr, cmdContext, "?_________code pays incompatible")
        End If
    End If
End If
V = TVANIFCLIT_Control(X)
If Not IsNull(V) Then
    blnUpdate_Control = False
    txtUpdate_TVANIFCLIT.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, V)
End If

newYTVANIF0.TVANIFCLIT = X

If blnUpdate_Control Then
    fraNIF_Update_Control = Null
    newYTVANIF0.TVANIFSTA = "V"
Else
    fraNIF_Update_Control = "<_________Fin du contrôle des données "
End If
End Function

Public Function fraParam_Update_Control()
Dim V
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
newYBIATAB0 = oldYBIATAB0
Select Case cmdParam_SQL_K
    Case 1: V = fraParam_Update_Control_1
    Case 2: V = fraParam_Update_Control_2
End Select
If Not IsNull(V) Then blnUpdate_Control = False

If blnUpdate_Control Then
    fraParam_Update_Control = Null
Else
    fraParam_Update_Control = "<_________Fin du contrôle des données "
End If
End Function

Public Function fraParam_Update_Control_2()
Dim V
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents

If optParam_Update_TVACOMOPE_CRE Then
    newYBIATAB0.BIATABK1 = "CRE"
Else
    newYBIATAB0.BIATABK1 = "ENG"
End If

If optParam_Update_TVACOMSTA_V Then
    newYBIATAB0.BIATABTXT = "V"
Else
    newYBIATAB0.BIATABTXT = "I"
End If

X = Trim(txtParam_Update_TVACOMOPE)
If IsNumeric(X) Then
    newYBIATAB0.BIATABK2 = Format$(Val(X), "000000000000")
Else
    newYBIATAB0.BIATABK2 = X
End If


If blnUpdate_Control Then
    fraParam_Update_Control_2 = Null
Else
    fraParam_Update_Control_2 = "<_________Fin du contrôle des données "
End If
End Function
Public Function fraParam_Update_Control_1()
Dim V, xSql As String
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents

newYBIATAB0.BIATABID = "TVAFACSTA"
newYBIATAB0.BIATABK1 = "CLIENARES"

X = Trim(txtParamUpdate_CLIENACLI)
If IsNumeric(X) Then
    newYBIATAB0.BIATABK2 = Format$(Val(X), "0000000")
    xSql = "select count(*) as Tally from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & newYBIATAB0.BIATABK2 & "'"
    Set rsSab = cnsab.Execute(xSql)
    If rsSab("Tally") = 0 Then
        blnUpdate_Control = False
        txtParamUpdate_CLIENACLI.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code client")
    End If

Else
    blnUpdate_Control = False
    txtParamUpdate_CLIENACLI.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code client")
End If

X = Trim(txtParamUpdate_CLIENARES)
If X = "" Then
    blnUpdate_Control = False
    txtParamUpdate_CLIENARES.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code routage")
End If
newYBIATAB0.BIATABTXT = X





If blnUpdate_Control Then
    fraParam_Update_Control_1 = Null
Else
    fraParam_Update_Control_1 = "<_________Fin du contrôle des données "
End If
End Function



Public Function fraDétail_Update_Control()
Dim blnUpdate_Control As Boolean
Dim blnUpdate_TVACOMCLIP As Boolean, blnUpdate_TVACOMCOMT As Boolean, blnUpdate_TVACOMTVAC As Boolean
Dim X As String, X1 As String
Dim V
Dim K As Integer
blnUpdate_Control = True
blnUpdate_TVACOMCLIP = True
blnUpdate_TVACOMCOMT = True
blnUpdate_TVACOMTVAC = True

Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données TVACOM"): DoEvents
newYTVACOM0 = oldYTVACOM0

newYTVACOM0.TVACOMSRVR = Trim(txtUpdate_TVACOMSRVR)

newYTVACOM0.TVACOMCLIC = Mid$(txtUpdate_TVACOMCLIC, 1, 1)

X = Trim(txtUpdate_TVACOMCLI)
If X = "" Then
    blnUpdate_Control = False
    txtUpdate_TVACOMCLI.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code client")
Else
    txtUpdate_TVACOMCLI.BackColor = txtUsr.BackColor
End If
newYTVACOM0.TVACOMCLI = Format$(X, "0000000")

If newYTVACOM0.TVACOMCLI <> oldYTVACOM0.TVACOMCLI _
Or newYTVACOM0.TVACOMCLIC <> oldYTVACOM0.TVACOMCLIC _
Or newYTVACOM0.TVACOMCLI <> meYTVACOM0.TVACOMCLI _
Or newYTVACOM0.TVACOMCLIC <> meYTVACOM0.TVACOMCLIC Then
    V = sqlTVACOMCLI(newYTVACOM0.TVACOMCLIC, newYTVACOM0.TVACOMCLI, xZCLIENA0, xZADRESS0)
    fraDétail_Display_TVACOMCLI
    If IsNull(V) Then
        If newYTVACOM0.TVACOMCLI <> meYTVACOM0.TVACOMCLI _
        Or newYTVACOM0.TVACOMCLIC <> meYTVACOM0.TVACOMCLIC Then
            meYTVACOM0.TVACOMCLI = newYTVACOM0.TVACOMCLI
            meYTVACOM0.TVACOMCLIC = newYTVACOM0.TVACOMCLIC
            X1 = Trim(xZCLIENA0.CLIENARSD)
            If X1 <> Mid$(txtUpdate_TVACOMCLIP, 1, 2) Then
                cbo_Scan X1, txtUpdate_TVACOMCLIP
                blnUpdate_Control = False
                blnUpdate_TVACOMCLIP = False
                txtUpdate_TVACOMCLIP.BackColor = vbYellow
                txtUpdate_TVACOMTVAC.BackColor = vbYellow
                Call lstErr_AddItem(lstErr, cmdContext, "!_________Nouveau pays de résidence")
            End If
        End If
    End If
End If


If blnUpdate_TVACOMCLIP Then
    X = Mid$(txtUpdate_TVACOMCLIP, 1, 2)
    If X = "  " Then
        blnUpdate_Control = False
        txtUpdate_TVACOMCLIP.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le pays de résidence")
    Else
        txtUpdate_TVACOMCLIP.BackColor = txtUsr.BackColor
    End If
    newYTVACOM0.TVACOMCLIP = X
End If

If xYTVACOM0.TVACOMTVAC = "T" Then GoTo Fin
'___________________________________________________________________ non mvt TVA
X = Trim(txtUpdate_TVACOMCOMC)
If X = "" Then
    blnUpdate_Control = False
    txtUpdate_TVACOMCOMC.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code commission")
    GoTo Fin
Else
    txtUpdate_TVACOMCOMC.BackColor = txtUsr.BackColor
End If
newYTVACOM0.TVACOMCOMC = X

If newYTVACOM0.TVACOMCOMC <> oldYTVACOM0.TVACOMCOMC Then
    If newYTVACOM0.TVACOMCOMC = "=ECRX " Then
        If newYTVACOM0.TVACOMMTVA = 0 Then
            X1 = "E"
        Else
            X1 = "N"
        End If
    Else
        K = InStr(7, X, "*")
        Select Case Mid$(X, K, 3)
            Case "* T": X1 = "N"
            Case "*  ": X1 = "E"
            Case Else: X1 = " "
        End Select
    End If
    If X1 <> Mid$(txtUpdate_TVACOMCOMT, 1, 1) Then
        cbo_Scan X1, txtUpdate_TVACOMCOMT
        blnUpdate_Control = False
        blnUpdate_TVACOMCOMT = False
        txtUpdate_TVACOMCOMT.BackColor = vbYellow
        txtUpdate_TVACOMTVAC.BackColor = vbYellow
        Call lstErr_AddItem(lstErr, cmdContext, "!_________Nouveau code TVA")
    End If
End If

If blnUpdate_TVACOMCOMT Then
    X = Mid$(txtUpdate_TVACOMCOMT, 1, 1)
    If X = " " Then
        blnUpdate_Control = False
        txtUpdate_TVACOMCOMT.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser commission taxable")
    Else
        txtUpdate_TVACOMCOMT.BackColor = txtUsr.BackColor
    End If
    newYTVACOM0.TVACOMCOMT = X
End If


If txtUpdate_TVACOMTVAC.Visible Then
    X = Mid$(txtUpdate_TVACOMTVAC, 1, 1)
    If X = " " Or X = "T" Then
        blnUpdate_Control = False
        txtUpdate_TVACOMTVAC.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code TVA")
    Else
        If txtUpdate_TVACOMTVAC.BackColor <> vbYellow Then txtUpdate_TVACOMTVAC.BackColor = txtUsr.BackColor
    End If
    newYTVACOM0.TVACOMTVAC = X
End If

If newYTVACOM0.TVACOMTVAC <> newYTVACOM0.TVACOMCOMT Then
    If txtUpdate_TVACOMTVAC.BackColor <> vbYellow Then
        blnUpdate_Control = False
        txtUpdate_TVACOMTVAC.BackColor = vbYellow
        Call lstErr_AddItem(lstErr, cmdContext, "!_________Vérifier la cohérence des codes TVA")
    End If
    
End If

newYTVACOM0.TVACOMCOME = Mid$(txtUpdate_TVACOMCOME, 1, 1)


If newYTVACOM0.TVACOMCOMC = "=ECRX " Then
    newYTVACOM0.TVACOMECRX = Val(txtUpdate_TVACOMECRX)
    If newYTVACOM0.TVACOMECRX = 0 Then
        meYTVACOM0.TVACOMECRX = 0
        libUpdate_TVACOMECRX = ""
        blnUpdate_Control = False
        txtUpdate_TVACOMECRX.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le n° d'écriture")
    Else
        If newYTVACOM0.TVACOMECRX <> meYTVACOM0.TVACOMECRX Then
            V = sqlYBIAMVTHP(newYTVACOM0.TVACOMETA, newYTVACOM0.TVACOMPIE, newYTVACOM0.TVACOMECRX, meYBIAMVT0)
    
             If IsNull(V) Then
                 libUpdate_TVACOMECRX = meYBIAMVT0.LIBELLIB1 & meYBIAMVT0.LIBELLIB2 & meYBIAMVT0.LIBELLIB3 & meYBIAMVT0.LIBELLIB4
                 txtUpdate_TVACOMECRX.BackColor = txtUsr.BackColor
                 meYTVACOM0.TVACOMECRX = newYTVACOM0.TVACOMECRX
                 blnUpdate_Control = False
             Else
                meYTVACOM0.TVACOMECRX = 0
                libUpdate_TVACOMECRX = ""
                txtUpdate_TVACOMECRX.BackColor = errUsr.BackColor
                Call lstErr_AddItem(lstErr, cmdContext, "?_________" & V)
             End If
        End If
    End If
End If
If newYTVACOM0.TVACOMMTVA = 0 Then
    If newYTVACOM0.TVACOMTVAC = "N" Or newYTVACOM0.TVACOMTVAC = "R" Then blnUpdate_Control = False: Call lstErr_AddItem(lstErr, cmdContext, "?_________incompatibilté TVA = 0")
Else
    If newYTVACOM0.TVACOMTVAC = "E" Or newYTVACOM0.TVACOMTVAC = "*" Then blnUpdate_Control = False: Call lstErr_AddItem(lstErr, cmdContext, "?_________incompatibilté TVA <> 0")
End If


If fraUpdate_TVACOMFACL.Visible Then blnUpdate_Control = fraDétail_Update_Control_TVACOMFACL

If blnUpdate_Control Then
    If Trim(newYTVACOM0.TVACOMCOMC) = "" Then
        blnUpdate_Control = False
        txtUpdate_TVACOMCOMT.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser commission taxable")
    End If
    If Trim(newYTVACOM0.TVACOMCLIP) = "" Then
        blnUpdate_Control = False
        txtUpdate_TVACOMCLIP.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le pays")
    End If
    If Trim(newYTVACOM0.TVACOMTVAC) = "" Then
        blnUpdate_Control = False
        txtUpdate_TVACOMTVAC.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code TVA")
    End If
End If

Fin:
If blnUpdate_Control Then
    newYTVACOM0.TVACOMSTA = "V"
    fraDétail_Update_Control = Null
Else
    If newYTVACOM0.TVACOMSRVR <> oldYTVACOM0.TVACOMSRVR Then
        newYTVACOM0 = oldYTVACOM0
        newYTVACOM0.TVACOMSRVR = Trim(txtUpdate_TVACOMSRVR)
        fraDétail_Update_Control = Null
    Else
        fraDétail_Update_Control = "?_________fraDétail_Update_Control TVACOM"
    End If
End If
End Function


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


Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim meUnit As typeUnit, X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_TVAFAC_Aut)
Call BiaPgmAut_Init("BIA_TVACOM", BIA_TVACOM_Aut)
Call BiaPgmAut_Init("BIA_TVANIF", BIA_TVANIF_Aut)
Call BiaPgmAut_Init("BIA_TVASRVR", BIA_TVASRVR_Aut)
Call BiaPgmAut_Init("BIA_TVAPARAM", BIA_TVAPARAM_Aut)

blnSetfocus = True
Form_Init
blnAuto = False

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@AUTO_TVAFAC": blnAuto = True
                        If Not IsEmpty(XPrt) Then Set Xprt_Previous = XPrt
                        Printer_PDF
'màj ignorer les pays non UE
                         Call cbo_Scan("6 -", cboDétail_SQL)
                         cmdDétail_Ok_Click
'état des commissions à réviser
                         Call cbo_Scan("3 -", cboDétail_SQL)
                         cmdDétail_Ok_Click
'état des NIF manquants
                         Call cbo_Scan("3 -", cboNIF_SQL)
                         
                         fgNIF.Visible = False
                         Call cbo_Scan(" ", txtNIF_TVANIFCLIC)
                         cmdNIF_Ok_Click
                         If arrYTVANIF0_Nb > 0 Then mnuPrint2_Liste1_Click
                         fgNIF.Visible = False
                         Call cbo_Scan("D", txtNIF_TVANIFCLIC)
                         cmdNIF_Ok_Click
                         If arrYTVANIF0_Nb > 0 Then mnuPrint2_Liste1_Click
                         fgNIF.Visible = False
                         Call cbo_Scan("G", txtNIF_TVANIFCLIC)
                         cmdNIF_Ok_Click
                         If arrYTVANIF0_Nb > 0 Then mnuPrint2_Liste1_Click
                         If Not IsEmpty(Xprt_Previous) Then Set XPrt = Xprt_Previous

                         Unload Me

    Case Else: blnAuto = False
End Select


End Sub


Public Sub cmdContext_Return()
Select Case SSTab1.Tab
    Case Is = 0
        If fraSelect_Update.Visible _
        And fraSelect_Update_B.Enabled _
        And cmdSelect_Update_Ok.Enabled Then cmdSelect_Update_Ok_Click: Exit Sub
    Case 1
        If fraDétail_Update.Visible _
        And fraDétail_Update_B.Enabled _
        And cmdDétail_Update_Ok.Enabled Then cmdDétail_Update_Ok_Click: Exit Sub
    Case 2
        If fraNIF_Update.Visible _
        And fraNIF_Update_B.Enabled _
        And cmdNIF_Update_Ok.Enabled Then cmdNIF_Update_Ok_Click: Exit Sub
        If Not fraNIF_Update.Visible Then cmdNIF_Ok_Click
    Case 3
        If fraParam_Update.Visible _
        And fraParam_Update_B.Enabled _
        And cmdParam_Update_Ok.Enabled Then cmdParam_Update_Ok_Click: Exit Sub
        If Not fraParam_Update.Visible Then cmdParam_Ok_Click
    Case Else
        If currentAction = "" Then
            If SSTab1.Tab > 0 Then
                SSTab1.Tab = 0
            Else
               'SendKeys "{TAB}"
               ' cmdSelect_Click
            End If
        End If
End Select
End Sub









Private Sub mnuPrint0_All_Click()
Dim I As Long, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
    
For I = 1 To arrYTVAFAC0_Nb
    fgSelect.Row = I
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xYTVAFAC0 = arrYTVAFAC0(K)
    'prtSAB_CDR_Monitor xYTVAFAC0
Next I

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub




Public Function cmdPrint_Facture()
Dim V
Dim K As Long, X As String, X2 As String, xSql As String
Dim curHT_N As Currency, curTVA_N As Currency, nbHT_N As Integer
Dim curHT_R As Currency, curTVA_R As Currency, nbHT_R As Integer
Dim curHT_E As Currency, nbHT_E As Integer
Dim curHT_X As Currency, nbHT_X As Integer
Dim curHT_NE As Currency, nbHT_NE As Integer

Dim curX As Currency
Dim kCom As Integer, blnCOM As Boolean
Dim mForeColor As Long
Dim Height8_7 As Integer
Dim nbL As Integer, nbL_T As Integer
Dim blnAnnulation_Légende As Boolean
Dim mCurrentX As Long, mCurrenty As Long

On Error GoTo Error_Handler

cmdPrint_Facture = Null
curHT_N = 0: curTVA_N = 0: nbHT_N = 0
curHT_R = 0: curTVA_R = 0: nbHT_R = 0
curHT_E = 0: nbHT_E = 0
curHT_X = 0: nbHT_X = 0
curHT_NE = 0: nbHT_NE = 0
nbL = 1: nbL_T = 1
blnAnnulation_Légende = False
Call lstErr_Clear(lstErr, cmdContext, "> Impression de la Facture :" & oldYTVAFAC0.TVAFACFACN): DoEvents

arrYTVACOM0_SQL " where TVACOMFACN = " & oldYTVAFAC0.TVAFACFACN & " order by TVACOMDTR,TVACOMOPE,TVACOMDOS"
V = sqlTVACOMCLI(oldYTVAFAC0.TVAFACCLIC, oldYTVAFAC0.TVAFACCLI, xZCLIENA0, xZADRESS0)

prtBIA_TVAFAC_Open 1, "BIA_TVAFAC"
        '_____________________________________________________________________________
'If mYTVAFAC0.TVAFACSTA = "F" Then
    frmElpPrt.prtFiligrane paramEditionFiligrane_Folder & "PAYE.jpg"
'_______________________________________________________________________________

Height8_7 = frmElpPrt.prtHeightDelta(8, 7)

mTVANIFCLIT_Pays = TVANIFCLIT_Pays(oldYTVAFAC0.TVAFACCLIP)
If Trim(oldYTVAFAC0.TVAFACCLIT) = "" Then
    newYTVAFAC0 = oldYTVAFAC0
    cmdSelect_SQL_6_Regroupement_NIF
    oldYTVAFAC0.TVAFACCLIT = newYTVAFAC0.TVAFACCLIT
End If

prtBIA_TVAFAC_Init_1 oldYTVAFAC0, xZADRESS0, xZCLIENA0, mTVANIFCLIT_Pays
prtBIA_TVAFAC_Form_1
mForeColor = XPrt.ForeColor

For K = 1 To arrYTVACOM0_Nb
        xYTVACOM0 = arrYTVACOM0(K)
        

        prtBIA_TVAFAC_NewLine 1
        XPrt.CurrentX = prtMinX + 50: XPrt.Print dateIBM10(xYTVACOM0.TVACOMDTR, True);
        
        XPrt.CurrentX = prtMinX + 1100: XPrt.Print xYTVACOM0.TVACOMOPE;
        X = Format$(xYTVACOM0.TVACOMDOS, "### ##0")
        XPrt.CurrentX = prtMinX + 2350 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 2450
        blnCOM = False
        X = Trim(xYTVACOM0.TVACOMCOMC)
        If X = "=ECRX" Then
            V = sqlYBIAMVTHP(xYTVACOM0.TVACOMETA, xYTVACOM0.TVACOMPIE, xYTVACOM0.TVACOMECRX, meYBIAMVT0)
            If IsNull(V) Then X = meYBIAMVT0.LIBELLIB1

        Else
            For arrCommission_K = 1 To arrCommission_Max
                If X = arrCommission(arrCommission_K, 1) Then
                    XPrt.Print arrCommission(arrCommission_K, 2); '& xYTVACOM0.TVAREFCLI
                    blnCOM = True
                    Exit For
                End If
            Next arrCommission_K
        End If
        
        If Not blnCOM Then XPrt.Print X;
        XPrt.ForeColor = vbBlue
        
 
        'If (Trim(xYTVAFAC0.TVAFACCLIT) = "" And xYTVAFAC0.TVAFACSTA = "V") _
        'Or xYTVAFAC0.TVAFACSTA = "2" Then
        If Not mTVANIFCLIT_Pays Then
           xYTVACOM0.TVACOMTVAC = "*"
        End If
        If mTVAFACCLIP_Code_Fiscal = "4" Or mTVAFACCLIP_Code_Fiscal = "5" Then
            If xYTVACOM0.TVACOMCOMT = "N" Or xYTVACOM0.TVACOMCOMT = "R" Then
                If xYTVACOM0.TVACOMMTVE = 0 Then
                    XPrt.ForeColor = vbMagenta
                    xYTVACOM0.TVACOMTVAC = "#"
                End If
            End If
        End If
        XPrt.CurrentX = prtMinX + 8350: XPrt.Print xYTVACOM0.TVACOMTVAC;
        
        XPrt.ForeColor = mForeColor
       If xYTVACOM0.TVACOMMON > 0 Then
            XPrt.ForeColor = vbRed
            blnAnnulation_Légende = True
            XPrt.CurrentX = prtMinX + 10000
            XPrt.Print "A";
        End If
        If xYTVACOM0.TVACOMQTE <= 1 Then
            xYTVACOM0.TVACOMQTE = 1
            curX = Abs(xYTVACOM0.TVACOMMON)
        Else
            curX = Round(Abs(xYTVACOM0.TVACOMMON) / xYTVACOM0.TVACOMQTE, 2)
        End If
        
        X = Format$(xYTVACOM0.TVACOMQTE, "####")
        XPrt.CurrentX = prtMinX + 6300 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(curX, "### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 7700 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 7800: XPrt.Print xYTVACOM0.TVACOMDEV;
        X = Format$(Abs(xYTVACOM0.TVACOMMONE), "### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 9900 - XPrt.TextWidth(X)
        XPrt.Print X;
        If xYTVACOM0.TVACOMMTVE <> 0 Then
            X = Format$(Abs(xYTVACOM0.TVACOMMTVE), "### ### ### ##0.00")
        Else
            X = "0"
        End If
        XPrt.CurrentX = prtMaxX - 50 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.ForeColor = mForeColor
        
        Select Case xYTVACOM0.TVACOMTVAC
            Case "N": nbHT_N = nbHT_N + 1
                      curHT_N = curHT_N + xYTVACOM0.TVACOMMONE
                      curTVA_N = curTVA_N + xYTVACOM0.TVACOMMTVE
            Case "E": nbHT_E = nbHT_E + 1
                      curHT_E = curHT_E + xYTVACOM0.TVACOMMONE
            Case "#": nbHT_NE = nbHT_NE + 1
                      curHT_NE = curHT_NE + xYTVACOM0.TVACOMMONE
            Case "R": nbHT_R = nbHT_R + 1
                      curHT_R = curHT_R + xYTVACOM0.TVACOMMONE
                      curTVA_R = curTVA_R + xYTVACOM0.TVACOMMTVE
            Case "*": nbHT_X = nbHT_X + 1
                      curHT_X = curHT_X + xYTVACOM0.TVACOMMONE
            Case Else: Call MsgBox("TVACOMTVAC non géré : " & xYTVACOM0.TVACOMTVAC, vbCritical, "Impression des factures TVA")
        End Select
'______________________________________________________annulation
        If xYTVACOM0.TVACOMMON > 0 And xYTVACOM0.TVACOMFACL > 0 Then
            XPrt.FontSize = 7
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
            'prtBIA_TVAFAC_NewLine 1
            XPrt.CurrentX = prtMinX + 2450
            XPrt.Print "# annulation de commission du justificatif N° " & xYTVACOM0.TVACOMFACL;
            XPrt.FontSize = 8
            XPrt.CurrentY = XPrt.CurrentY + 100
        End If
Next K

XPrt.DrawWidth = 5
prtBIA_TVAFAC_NewLine 1

If nbHT_N > 0 Then nbL = nbL + 2: nbL_T = nbL_T + 1
If nbHT_NE > 0 Then nbL = nbL + 2: nbL_T = nbL_T + 1
If nbHT_R > 0 Then nbL = nbL + 2: nbL_T = nbL_T + 1
If nbHT_E > 0 Then nbL = nbL + 2: nbL_T = nbL_T + 1
If nbHT_X > 0 Then nbL = nbL + 2: nbL_T = nbL_T + 1
If blnAnnulation_Légende Then nbL_T = nbL_T + 1

If XPrt.CurrentY + nbL * prtlineHeight > prtMaxY Then
    prtBIA_TVAFAC_Form_1_Col ("C")
    XPrt.CurrentX = prtMinX + 10300: XPrt.Print "---/---";
    frmElpPrt.prtNewPage
    prtBIA_TVAFAC_Form_1
    prtBIA_TVAFAC_NewLine 1
Else
    prtBIA_TVAFAC_Form_1_Col ("C")
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), RGB(0, 123, 141)
End If
XPrt.ForeColor = vbBlue

XPrt.FontBold = True
If nbHT_N > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + 50
    Call frmElpPrt.prtTrame(prtMinX + 8250, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, " ", 240)
    XPrt.ForeColor = vbBlue
    XPrt.CurrentX = prtMinX + 8350: XPrt.Print "N";
    XPrt.ForeColor = RGB(0, 123, 141)
    X = "Totaux HT et TVA des prestations soumises au taux de TVA à  19.60 %"
    XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.ForeColor = prtForeColor

    If curHT_N > 0 Then XPrt.ForeColor = vbRed
    X = Format$(Abs(curHT_N), "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9900 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(Abs(curTVA_N), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 50 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.ForeColor = prtForeColor
End If
If nbHT_R > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + 50
     Call frmElpPrt.prtTrame(prtMinX + 8250, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, " ", 255)
    XPrt.ForeColor = vbBlue
    XPrt.CurrentX = prtMinX + 8350: XPrt.Print "R";
    XPrt.ForeColor = RGB(0, 123, 141)
    X = "Totaux HT et TVA des prestations soumises au taux de TVA à 5.5 %"
    XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.ForeColor = prtForeColor
    If curHT_R > 0 Then XPrt.ForeColor = vbRed

    X = Format$(Abs(curHT_R), "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9900 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(Abs(curTVA_R), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 50 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.ForeColor = prtForeColor
End If
If nbHT_NE > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + 50
    Call frmElpPrt.prtTrame(prtMinX + 8250, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, " ", 240)
    XPrt.ForeColor = vbMagenta
    XPrt.CurrentX = prtMinX + 8350: XPrt.Print "#";
    XPrt.ForeColor = RGB(0, 123, 141)
    X = "Total HT des prestations soumises au régime d'autoliquidation"
    XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.ForeColor = prtForeColor
    If curHT_NE > 0 Then XPrt.ForeColor = vbRed
    X = Format$(Abs(curHT_NE), "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9900 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMaxX - 50 - XPrt.TextWidth("0")
    XPrt.Print "0";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.ForeColor = prtForeColor
End If

If nbHT_E > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + 50
     Call frmElpPrt.prtTrame(prtMinX + 8250, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, " ", 255)
   XPrt.ForeColor = vbBlue
    XPrt.CurrentX = prtMinX + 8350: XPrt.Print "E";
    XPrt.ForeColor = RGB(0, 123, 141)
    X = "Total des prestations exonérées de TVA"
    XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.ForeColor = prtForeColor
    If curHT_E > 0 Then XPrt.ForeColor = vbRed
    X = Format$(Abs(curHT_E), "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9900 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMaxX - 50 - XPrt.TextWidth("0")
    XPrt.Print "0";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.ForeColor = prtForeColor
End If
If nbHT_X > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + 50
     Call frmElpPrt.prtTrame(prtMinX + 8250, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, " ", 255)
    XPrt.ForeColor = RGB(0, 123, 141)
    X = "Total des prestations non soumises à la TVA française"
    XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.ForeColor = vbBlue
    XPrt.CurrentX = prtMinX + 8350: XPrt.Print "*";
    XPrt.ForeColor = prtForeColor
    If curHT_X > 0 Then XPrt.ForeColor = vbRed
    X = Format$(Abs(curHT_X), "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9900 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMaxX - 50 - XPrt.TextWidth("0")
    XPrt.Print "0";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.ForeColor = prtForeColor
'____________________________________________________________________
End If

prtBIA_TVAFAC_Form_1_Col " "
XPrt.FontSize = 10
curX = curHT_N + curHT_R + curHT_E + curHT_X + curTVA_N + curTVA_R + curHT_NE
Call frmElpPrt.prtTrame(prtMinX + 8250, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, "B", 240)
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.ForeColor = RGB(0, 123, 141)
X = "Total TTC du présent relevé"
XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.ForeColor = prtForeColor

If curX > 0 Then XPrt.ForeColor = vbRed
X = Format$(Abs(curX), "### ### ### ##0.00") & " €"
XPrt.CurrentX = prtMaxX - 500 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.ForeColor = prtForeColor


If curX <> oldYTVAFAC0.TVAFACMTTC Then
    XPrt.FontSize = 12: XPrt.FontBold = True
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 6450
    XPrt.ForeColor = vbRed
    XPrt.Print "ERREUR TOTAL TTC " & oldYTVAFAC0.TVAFACMTTC
    XPrt.ForeColor = prtForeColor
    MsgBox "ERREUR TOTAL TTC " & oldYTVAFAC0.TVAFACFACN, vbCritical, "BIA_TVAFAC"
End If
'____________________________________________________________________

XPrt.CurrentY = prtMaxY - nbL_T * prtlineHeight
XPrt.FontSize = 8
If blnAnnulation_Légende Then
    XPrt.ForeColor = vbRed
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print "A  : prestation annulée";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If
XPrt.ForeColor = vbBlue
XPrt.FontBold = True
XPrt.FontUnderline = True
XPrt.CurrentX = prtMinX + 50
XPrt.Print "Tx: Régime de TVA appliqué";
XPrt.FontBold = False
XPrt.FontUnderline = False
If nbHT_N > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print "N : TVA à 19.60 % (taux normal)";
End If
If nbHT_R > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print "R : TVA à 5.5 % (taux réduit)";
End If
If nbHT_E > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print "E : Exonération prévue par l'article 261 C-1° du CGI";
End If
If nbHT_X > 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print "* : Prestation non soumise à la TVA française en application des règles de territorialité";
End If
If nbHT_NE > 0 Then
    XPrt.ForeColor = vbMagenta
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print "# : prestations soumises au régime d'autoliquidation (article 196 de la directive 2006/112/CE modifiée)";
End If


XPrt.ForeColor = prtForeColor

'____________________________________________________________________

'nom du fichier .pdf
If oldYTVAFAC0.TVAFACCLIC = " " Then
    X = "C" & oldYTVAFAC0.TVAFACCLI
Else
    X = oldYTVAFAC0.TVAFACCLIC & oldYTVAFAC0.TVAFACCLI
End If
prtPgmName = "F" & Format(oldYTVAFAC0.TVAFACFACN, "0000000") & "_" & X & ".pdf"

'Archivage de la facturation

If cmdSelect_SQL_K = 7 Then
    Print #1, Time, oldYTVAFAC0.TVAFACFACN & " > Archivage : " & prtPgmName
    '----------------------------------------------
    prtPgmName = paramFacturation_Path_AAAA & prtPgmName
    prtBIA_TVAFAC_Close 1
    Print #1, Time, oldYTVAFAC0.TVAFACFACN & " < Archivage terminé"
    '----------------------------------------------
Else
    prtBIA_TVAFAC_Close 1
End If

fgSelect.Visible = True

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    cmdPrint_Facture = V
    If cmdSelect_SQL_K = 7 Then
        Print #1, Time, oldYTVAFAC0.TVAFACFACN & " ? Impression erreur : " & V
    End If
Exit_sub:

End Function
Public Sub cmdSendMail_Douanes(lAttachment As String)
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
Dim xAlerte As String, xSql As String

On Error Resume Next

'____________________________________________________________________________________________


wSendMail.FromDisplayName = "TVA_DES"
wSendMail.RecipientDisplayName = "TVA"

wSendMail.Subject = "Traitement TVA_DES : " & dateImp10(wAmjMin) & " (cf. pièce jointe)"
wSendMail.Attachment = lAttachment
wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<span style='font-size:12.0pt;font-family:Arial Unicode MS'>" & "<Font color = #0000FF>" _
                    & "Bonjour," _
                    & "<BR> Veuillez trouver ci-joint l'état de déclaration de la TVA intracommunautaire aux douanes" _
                    & "<BR> et le fichier DES au format XML à envoyer via l'adresse https://pro.douane.gouv.fr ." _
                    & "<BR> Bonne réception." _

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub



Public Sub cmdPrintDétail_Liste(lK As String)
Dim V, K As Long, X As String
Dim wIndex As Integer
Dim mForeColor As Long
rsYTVACOM0_Init meYTVACOM0
meYTVACOM0.TVACOMCLIC = "?"

fgDétail.Visible = False
prtBIA_TVAFAC_Open 2, "Facturation : Liste des commissions"

For K = 1 To fgDétail.Rows - 1
    fgDétail.Row = K
    fgDétail.Col = fgDétail_arrIndex
    wIndex = Val(fgDétail.Text)
    xYTVACOM0 = arrYTVACOM0(wIndex)
    
    cmdPrintDétail_Liste_Détail lK
Next K

cmdPrintDétail_Liste_Close lK

fgDétail.Visible = True
End Sub
Public Sub cmdPrintNIF_Liste(lK As String)
Dim V, K As Long, X As String
Dim wSubject As String, wRecipient As String
Dim wIndex As Integer
Dim mForeColor As Long
rsYTVANIF0_Init meYTVANIF0
meYTVANIF0.TVANIFCLIC = "?"

fgNIF.Visible = False
Select Case Mid$(txtNIF_TVANIFCLIC, 1, 1)
    Case " ": wRecipient = "GSOP"
    Case "D", "R": wRecipient = "SOBI"
    Case "G": wRecipient = "GDMP"
    Case Else: wRecipient = "INFO"
End Select

wSubject = "Facturation TVA : Liste des numéros de TVA intracommunautaire - " & wRecipient
If cmdNIF_SQL_K = 3 Then wSubject = "Facturation TVA : Liste des numéros de TVA intracommunautaire MANQUANTS - " & wRecipient
prtBIA_TVAFAC_Open 4, wSubject

For K = 1 To fgNIF.Rows - 1
    fgNIF.Row = K
    fgNIF.Col = fgNIF_arrIndex
    wIndex = Val(fgNIF.Text)
    xYTVANIF0 = arrYTVANIF0(wIndex)
    
    cmdPrintNIF_Liste_Détail lK
Next K

XPrt.DrawWidth = 5
prtBIA_TVAFAC_NewLine 4
prtBIA_TVAFAC_Form_4_Col

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
prtBIA_TVAFAC_Close 4

If blnAuto Then cmdSendMail prtIMP_PDF_FileName, "BIA_TVANIF", wRecipient, wSubject

fgNIF.Visible = True
End Sub

Public Sub cmdPrintDétail_Liste_Close(lK As String)

XPrt.DrawWidth = 5
prtBIA_TVAFAC_NewLine 2
prtBIA_TVACOM_Form_2_Col

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
prtBIA_TVAFAC_Close 2

If blnAuto Then

    cmdSendMail prtIMP_PDF_FileName, "BIA_TVACOM", Trim(meYTVACOM0.TVACOMSRVR), " Facturation TVA : Liste des commissions à réviser - " & meYTVACOM0.TVACOMSRVR

End If

End Sub

Public Sub cmdSendMail(lFileName As String, lFrom As String, lRecipient As String, lSubject As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
wSendMail.FromDisplayName = lFrom
wSendMail.RecipientDisplayName = lRecipient

bgColor = "MAGENTA"
wSendMail.Subject = lSubject
wSendMail.Attachment = lFileName
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Arial & Asc34 & ">" _
                    & htmlFontColor("BLUE") & "<B><CENTER>" & "voir pièce jointe" _
                    & "<BR>"

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail
End Sub

Public Sub cmdPrintDétail_Liste3(lK As String)
Dim V, K As Long, X As String
Dim wIndex As Integer
Dim mForeColor As Long
rsYTVACOM0_Init meYTVACOM0
meYTVACOM0.TVACOMCLIC = "?"
Dim blnXprt_Open As Boolean

fgDétail.Visible = False
blnXprt_Open = False
For K = 1 To arrYTVACOM0_Nb
    xYTVACOM0 = arrYTVACOM0(K)
    If xYTVACOM0.TVACOMSRVR <> meYTVACOM0.TVACOMSRVR Then
        If blnXprt_Open Then cmdPrintDétail_Liste_Close lK
        meYTVACOM0 = xYTVACOM0
        meYTVACOM0.TVACOMCLIC = "?"
        prtBIA_TVAFAC_Open 2, "Facturation TVA : Liste des commissions à réviser - " & meYTVACOM0.TVACOMSRVR
        blnXprt_Open = True
    End If
    cmdPrintDétail_Liste_Détail lK
Next K
If blnXprt_Open Then cmdPrintDétail_Liste_Close lK

fgDétail.Visible = True



End Sub




Public Sub cmdPrintDétail_Liste_Détail(lK As String)
Dim V, K As Long, X As String
Dim mForeColor As Long

    
    prtBIA_TVAFAC_NewLine 2
    
     Select Case xYTVACOM0.TVACOMSTA
        Case Is = "V": XPrt.ForeColor = vbBlue
        Case Is = "F": XPrt.ForeColor = &HC000&
        Case Is = "A": XPrt.ForeColor = &HC0C0C0
        Case Is = "I": XPrt.ForeColor = &HC0C000
        Case Is = "0", "9": XPrt.ForeColor = vbMagenta
        Case Else: XPrt.ForeColor = vbBlack
    End Select
    mForeColor = XPrt.ForeColor
   If meYTVACOM0.TVACOMCLIC <> xYTVACOM0.TVACOMCLIC Or meYTVACOM0.TVACOMCLI <> xYTVACOM0.TVACOMCLI Then
        meYTVACOM0 = xYTVACOM0
        XPrt.CurrentX = prtMinX + 50: XPrt.Print xYTVACOM0.TVACOMCLIC & " " & xYTVACOM0.TVACOMCLI;
        V = sqlTVACOMCLI(xYTVACOM0.TVACOMCLIC, xYTVACOM0.TVACOMCLI, xZCLIENA0, xZADRESS0)
        XPrt.CurrentX = prtMinX + 900: XPrt.Print Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2);
        If lK = "2" Then prtBIA_TVAFAC_NewLine 2
   End If
    
    If lK = "2" Then
        If xYTVACOM0.TVACOMECRX <> 0 Then
            V = sqlYBIAMVTHP(xYTVACOM0.TVACOMETA, xYTVACOM0.TVACOMPIE, xYTVACOM0.TVACOMECRX, meYBIAMVT0)
        Else
            V = sqlYBIAMVTHP(xYTVACOM0.TVACOMETA, xYTVACOM0.TVACOMPIE, xYTVACOM0.TVACOMECR, meYBIAMVT0)
        End If
        
        XPrt.CurrentX = prtMinX + 900: XPrt.Print Trim(meYBIAMVT0.LIBELLIB1) & Trim(meYBIAMVT0.LIBELLIB2);
    End If
    
   XPrt.CurrentX = prtMinX + 6200: XPrt.Print xYTVACOM0.TVACOMOPE;
   XPrt.CurrentX = prtMinX + 6600: XPrt.Print xYTVACOM0.TVACOMNAT;
   XPrt.CurrentX = prtMinX + 7400: XPrt.Print xYTVACOM0.TVACOMEVE;
    X = Format$(Abs(xYTVACOM0.TVACOMDOS), "### ### ### ###")
    XPrt.CurrentX = prtMinX + 8400 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 8500: XPrt.Print dateIBM10(xYTVACOM0.TVACOMDTR, True);
    
    If xYTVACOM0.TVACOMMON > 0 Then XPrt.ForeColor = vbRed
    X = Format$(Abs(xYTVACOM0.TVACOMMON), "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 10400 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(Abs(xYTVACOM0.TVACOMMTVA), "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 11500: XPrt.Print xYTVACOM0.TVACOMDEV;
    
     XPrt.ForeColor = mForeColor

    XPrt.CurrentX = prtMinX + 12000: XPrt.Print xYTVACOM0.TVACOMCOMC;
    XPrt.CurrentX = prtMinX + 12700: XPrt.Print xYTVACOM0.TVACOMCOMT;
    XPrt.CurrentX = prtMinX + 13000: XPrt.Print xYTVACOM0.TVACOMCLIP;
    XPrt.CurrentX = prtMinX + 13500: XPrt.Print xYTVACOM0.TVACOMTVAC & xYTVACOM0.TVACOMCOME;
    X = Format$(Abs(xYTVACOM0.TVACOMFACN), "### ### ### ###")
    XPrt.CurrentX = prtMinX + 14500 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(Abs(xYTVACOM0.TVACOMFACL), "### ### ### ###")
    XPrt.CurrentX = prtMinX + 15300 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMaxX - 200: XPrt.Print xYTVACOM0.TVACOMSTA;
    XPrt.ForeColor = prtForeColor



End Sub

Public Sub cmdPrintNIF_Liste_Détail(lK As String)
Dim V, K As Long, X As String
Dim mForeColor As Long

    
    prtBIA_TVAFAC_NewLine 4
    
     Select Case xYTVANIF0.TVANIFSTA
        Case Is = "V": XPrt.ForeColor = vbBlue
        Case Is = "F": XPrt.ForeColor = &HC000&
        Case Is = "A": XPrt.ForeColor = &HC0C0C0
        Case Is = "I": XPrt.ForeColor = &HC0C000
        Case Is = "0", "9", " ": XPrt.ForeColor = vbMagenta
        Case Else: XPrt.ForeColor = vbBlack
    End Select
    mForeColor = XPrt.ForeColor

    XPrt.CurrentX = prtMinX + 50: XPrt.Print xYTVANIF0.TVANIFCLIC & " " & xYTVANIF0.TVANIFCLI;
    XPrt.CurrentX = prtMinX + 1000: XPrt.Print xYTVANIF0.TVANIFRS;
    XPrt.CurrentX = prtMinX + 7000: XPrt.Print xYTVANIF0.TVANIFCLIP;
    XPrt.CurrentX = prtMinX + 7500: XPrt.Print TVANIFCLIT_Format(xYTVANIF0.TVANIFCLIT);
    XPrt.ForeColor = prtForeColor

End Sub

Public Sub cmdPrintSelect_Liste()
Dim V, K As Long, X As String, curX As Currency

Dim wIndex As Integer
Dim mForeColor As Long
rsYTVAFAC0_Init meYTVAFAC0
meYTVAFAC0.TVAFACCLIC = "?"

fgSelect.Visible = False
prtBIA_TVAFAC_Open 3, "Facturation : Liste des factures"
For K = 1 To fgSelect.Rows - 1
    fgSelect.Row = K
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    xYTVAFAC0 = arrYTVAFAC0(wIndex)
    
    prtBIA_TVAFAC_NewLine 3
     Select Case xYTVAFAC0.TVAFACSTA
        Case Is = "V": XPrt.ForeColor = vbBlue
        Case Is = "F": XPrt.ForeColor = &HC000&
        Case Is = "A": XPrt.ForeColor = &HC0C0C0
        Case Is = "I": XPrt.ForeColor = &HC0C000
        Case Is = "0", "1": XPrt.ForeColor = vbMagenta
        Case Else: XPrt.ForeColor = vbBlack
    End Select
    mForeColor = XPrt.ForeColor
    If meYTVAFAC0.TVAFACCLIC <> xYTVAFAC0.TVAFACCLIC Or meYTVAFAC0.TVAFACCLI <> xYTVAFAC0.TVAFACCLI Then
        meYTVAFAC0 = xYTVAFAC0
        XPrt.CurrentX = prtMinX + 50: XPrt.Print xYTVAFAC0.TVAFACCLIC & " " & xYTVAFAC0.TVAFACCLI;
        V = sqlTVACOMCLI(xYTVAFAC0.TVAFACCLIC, xYTVAFAC0.TVAFACCLI, xZCLIENA0, xZADRESS0)
        XPrt.CurrentX = prtMinX + 900: XPrt.Print Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2);
   End If
    XPrt.CurrentX = prtMinX + 7500: XPrt.Print xYTVAFAC0.TVAFACCLIP;
    XPrt.CurrentX = prtMinX + 7900: XPrt.Print xYTVAFAC0.TVAFACCLIT;
    XPrt.CurrentX = prtMinX + 9500: XPrt.Print dateIBM10(xYTVAFAC0.TVAFACDTR, True);
     X = Format$((xYTVAFAC0.TVAFACFACN), "### ### ### ###")
    XPrt.CurrentX = prtMinX + 10400 - XPrt.TextWidth(X)
    XPrt.Print X;
   
    If xYTVAFAC0.TVAFACMEXO = 0 Then
        X = "-"
        XPrt.ForeColor = mForeColor
    Else
        X = Format$(Abs(xYTVAFAC0.TVAFACMEXO), "### ### ### ###.00")
        If xYTVAFAC0.TVAFACMEXO = 0 Then
           XPrt.ForeColor = vbRed
        Else
            XPrt.ForeColor = mForeColor
        End If
    End If
    XPrt.CurrentX = prtMinX + 11500 - XPrt.TextWidth(X)
    XPrt.Print X;
    
    curX = xYTVAFAC0.TVAFACMTTC - xYTVAFAC0.TVAFACMEXO - xYTVAFAC0.TVAFACMTVA
    If curX = 0 Then
        X = "-"
        XPrt.ForeColor = mForeColor
    Else
        X = Format$(Abs(curX), "### ### ### ###.00")
         If curX > 0 Then
             XPrt.ForeColor = vbRed
         Else
              XPrt.ForeColor = mForeColor
        End If
    End If
    XPrt.CurrentX = prtMinX + 13000 - XPrt.TextWidth(X)
    XPrt.Print X;
     
    If xYTVAFAC0.TVAFACMTVA = 0 Then
        X = "-"
        XPrt.ForeColor = mForeColor
    Else
        X = Format$(Abs(xYTVAFAC0.TVAFACMTVA), "### ### ### ###.00")
         If xYTVAFAC0.TVAFACMTVA > 0 Then
             XPrt.ForeColor = vbRed
         Else
              XPrt.ForeColor = mForeColor
        End If
    End If
    XPrt.CurrentX = prtMinX + 14000 - XPrt.TextWidth(X)
    XPrt.Print X;
     
    If xYTVAFAC0.TVAFACMTTC = 0 Then
        X = "-"
        XPrt.ForeColor = mForeColor
    Else
        X = Format$(Abs(xYTVAFAC0.TVAFACMTTC), "### ### ### ###.00")
         If xYTVAFAC0.TVAFACMTTC > 0 Then
             XPrt.ForeColor = vbRed
         Else
              XPrt.ForeColor = mForeColor
        End If
    End If
    XPrt.CurrentX = prtMinX + 15400 - XPrt.TextWidth(X)
    XPrt.Print X;
     XPrt.ForeColor = mForeColor

    XPrt.CurrentX = prtMaxX - 200: XPrt.Print xYTVAFAC0.TVAFACSTA;
    XPrt.ForeColor = prtForeColor

Next K
XPrt.DrawWidth = 5
prtBIA_TVAFAC_NewLine 3
prtBIA_TVAFAC_Form_3_Col

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
prtBIA_TVAFAC_Close 3

fgSelect.Visible = True


End Sub

Public Function cmdSelect_Update_Ok_Transaction()
Dim V, X As String, xSql As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_Update_Ok_Transaction"
'-------------------------------------------------------
cmdSelect_Update_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYTVAFAC0_Update(newYTVAFAC0, oldYTVAFAC0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdSelect_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Public Function cmdSelect_Update_Annuler_Transaction()
Dim V, X As String, xSql As String
Dim Nb As Long, K As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_Update_Annuler_Transaction"
'-------------------------------------------------------
cmdSelect_Update_Annuler_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
newYTVAFAC0 = oldYTVAFAC0
newYTVAFAC0.TVAFACSTA = "A"
newYTVAFAC0.TVAFACDTR = DSys - 19000000
V = sqlYTVAFAC0_Update(newYTVAFAC0, oldYTVAFAC0)
If Not IsNull(V) Then GoTo Error_MsgBox
For K = 1 To arrYTVACOM0_Nb
    newYTVACOM0 = arrYTVACOM0(K)
    newYTVACOM0.TVACOMFACN = 0
    newYTVACOM0.TVACOMSTA = "V"
    V = sqlYTVACOM0_Update(newYTVACOM0, arrYTVACOM0(K))
    If Not IsNull(V) Then GoTo Error_MsgBox

Next K
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdSelect_Update_Annuler_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Public Function cmdDétail_Update_Ok_Transaction()
Dim V, X As String, xSql As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdDétail_Update_Ok_Transaction"
'-------------------------------------------------------
cmdDétail_Update_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYTVACOM0_Update(newYTVACOM0, oldYTVACOM0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdDétail_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function



Public Sub lstSelect_Load_2()

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_2"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = False
fraSelect_Options_1.Enabled = False
chkSelect_TVAFACDTR.Enabled = True
chkSelect_TVAFACDTR = "0"
txtSelect_TVAFACSTA.Enabled = True
txtSelect_TVAFACCLI.Enabled = True
chkSelect_TVAFACSTA.Visible = False

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Public Sub lstSelect_Load_3()

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_3"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_TVAFACDTR.Enabled = False
chkSelect_TVAFACDTR = "0"
txtSelect_TVAFACSTA.Enabled = False
txtSelect_TVAFACCLI.Enabled = True
chkSelect_TVAFACSTA.Visible = False

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Public Sub lstSelect_Load_6()
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_6"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_TVAFACDTR.Enabled = False
chkSelect_TVAFACDTR = "1"
chkSelect_TVAFACDTR.Enabled = False
Call DTPicker_Set(txtSelect_TVAFACDTR, Mid$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "01")
Call DTPicker_Set(txtSelect_TVAFACDTR_Max, YBIATAB0_DATE_CPT_MP1)
txtSelect_TVAFACSTA.Enabled = False
txtSelect_TVAFACCLI.Enabled = True
chkSelect_TVAFACSTA.Visible = False

End Sub

Public Sub fraSelect_Display()
Dim V
Dim X As String, X1 As String
fraSelect_Update.Visible = True
fraSelect_Update_A.Enabled = False
fraSelect_Update_B.Enabled = True
Select Case xYTVAFAC0.TVAFACSTA
    Case "F": fraSelect_Update_B.Enabled = BIA_TVAFAC_Aut.Xspécial 'False
    Case "0": fraSelect_Update_B.Enabled = BIA_TVAFAC_Aut.Saisir
    Case "1", "2": fraSelect_Update_B.Enabled = BIA_TVAFAC_Aut.Valider
    Case Else: fraSelect_Update_B.Enabled = BIA_TVAFAC_Aut.Xspécial
End Select
txtUpdate_TVAFACCLIT.Enabled = BIA_TVAFAC_Aut.Xspécial

txtUpdate_TVAFACETA = xYTVAFAC0.TVAFACETA
cbo_Scan xYTVAFAC0.TVAFACCLIC, txtUpdate_TVAFACCLIC: txtUpdate_TVAFACCLIC.BackColor = txtUsr.BackColor
txtUpdate_TVAFACCLI = xYTVAFAC0.TVAFACCLI: txtUpdate_TVAFACCLI.BackColor = txtUsr.BackColor
V = sqlTVACOMCLI(xYTVAFAC0.TVAFACCLIC, xYTVAFAC0.TVAFACCLI, xZCLIENA0, xZADRESS0)
fraSelect_Display_TVACOMCLI
cbo_Scan xYTVAFAC0.TVAFACCLIP, txtUpdate_TVAFACCLIP: txtUpdate_TVAFACCLIP.BackColor = txtUsr.BackColor

txtUpdate_TVAFACMTTC = Format$(Abs(xYTVAFAC0.TVAFACMTTC), "### ### ### ###.00")
If xYTVAFAC0.TVAFACMTTC > 0 Then
    txtUpdate_TVAFACMTTC.ForeColor = vbRed
Else
    txtUpdate_TVAFACMTTC.ForeColor = vbBlue
End If

txtUpdate_TVAFACMTVA = Format$(Abs(xYTVAFAC0.TVAFACMTVA), "### ### ### ###.00")
If xYTVAFAC0.TVAFACMTVA > 0 Then
    txtUpdate_TVAFACMTVA.ForeColor = vbRed
Else
    txtUpdate_TVAFACMTVA.ForeColor = vbBlue
End If

txtUpdate_TVAFACMEXO = Format$(Abs(xYTVAFAC0.TVAFACMEXO), "### ### ### ###.00")
If xYTVAFAC0.TVAFACMEXO > 0 Then
    txtUpdate_TVAFACMEXO.ForeColor = vbRed
Else
    txtUpdate_TVAFACMEXO.ForeColor = vbBlue
End If

Call DTPicker_Set(txtUpdate_TVAFACDTR, xYTVAFAC0.TVAFACDTR + 19000000)
txtUpdate_TVAFACFACN = xYTVAFAC0.TVAFACFACN


cbo_Scan xYTVAFAC0.TVAFACSTA, txtUpdate_TVAFACSTA
txtUpdate_TVAFACUSR = xYTVAFAC0.TVAFACUSR
If xYTVAFAC0.TVAFACSTA = "0" Or xYTVAFAC0.TVAFACSTA = "1" Then
    txtUpdate_TVAFACCLIT.Enabled = True
Else
    txtUpdate_TVAFACCLIT.Enabled = False
End If
If xYTVAFAC0.TVAFACSTA = "0" And Trim(xYTVAFAC0.TVAFACCLIT) = "" Then
    newYTVAFAC0 = xYTVAFAC0
    cmdSelect_SQL_6_Regroupement_NIF
    xYTVAFAC0.TVAFACCLIT = newYTVAFAC0.TVAFACCLIT
End If

txtUpdate_TVAFACCLIT = xYTVAFAC0.TVAFACCLIT
txtUpdate_TVAFACCLIT.BackColor = txtUsr.BackColor

Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail d'une facture"): DoEvents
End Sub

Public Sub fraDétail_Display()
Dim V
Dim X As String, X1 As String
Dim blnSaisir As Boolean, blnValider As Boolean
'______________________________________________________________________________
fraDétail_Update_A.Enabled = False
fraDétail_Update_B.Enabled = True
meYTVACOM0 = xYTVACOM0
X = Trim(xYTVACOM0.TVACOMSRVR)
If X = Trim(currentUser.Unit) Or BIA_TVASRVR_Aut.Saisir Then
    blnSaisir = BIA_TVACOM_Aut.Saisir
    blnValider = BIA_TVACOM_Aut.Valider
Else
    blnSaisir = False
    blnValider = False
End If

Select Case xYTVACOM0.TVACOMSTA
    Case "F": fraDétail_Update_B.Enabled = False
    Case "9": fraDétail_Update_B.Enabled = blnSaisir
    Case Else: fraDétail_Update_B.Enabled = blnValider
End Select
txtUpdate_TVACOMSRVR.Enabled = BIA_TVASRVR_Aut.Saisir
txtUpdate_TVACOMCOME.Enabled = BIA_TVACOM_Aut.Xspécial
cmdDétail_Update_Annuler.Enabled = blnValider  'BIA_TVACOM_Aut.Xspécial
'______________________________________________________________________________
fraDétail_Update.Visible = True
txtUpdate_TVACOMETA = xYTVACOM0.TVACOMETA & "-" & xYTVACOM0.TVACOMSER & "/" & xYTVACOM0.TVACOMSER
txtUpdate_TVACOMPLA = xYTVACOM0.TVACOMPLA
txtUpdate_TVACOMPIE = xYTVACOM0.TVACOMPIE
txtUpdate_TVACOMECR = xYTVACOM0.TVACOMECR
txtUpdate_TVACOMCPT = xYTVACOM0.TVACOMCPT
Call DTPicker_Set(txtUpdate_TVACOMDVA, xYTVACOM0.TVACOMDVA + 19000000)
Call DTPicker_Set(txtUpdate_TVACOMDTR, xYTVACOM0.TVACOMDTR + 19000000)
txtUpdate_TVACOMOPE = xYTVACOM0.TVACOMOPE
txtUpdate_TVACOMNAT = xYTVACOM0.TVACOMNAT
txtUpdate_TVACOMQTE = xYTVACOM0.TVACOMQTE
txtUpdate_TVACOMEVE = xYTVACOM0.TVACOMEVE
txtUpdate_TVACOMDOS = xYTVACOM0.TVACOMDOS
txtUpdate_TVACOMDEV = xYTVACOM0.TVACOMDEV
txtUpdate_TVACOMMON = Format$(Abs(xYTVACOM0.TVACOMMON), "### ### ### ###.00")
txtUpdate_TVACOMMONE = Format$(Abs(xYTVACOM0.TVACOMMONE), "### ### ### ###.00")
xYTVACOM0.TVACOMAVOIR = " "
If xYTVACOM0.TVACOMMON > 0 Then
    txtUpdate_TVACOMMON.ForeColor = vbRed
    txtUpdate_TVACOMMONE.ForeColor = vbRed
    fraUpdate_TVACOMFACL.Visible = True
    txtUpdate_TVACOMFACL = xYTVACOM0.TVACOMFACL
    txtUpdate_TVACOMFACL.ForeColor = vbRed
    txtUpdate_TVACOMFACL.BackColor = txtUsr.BackColor
    If xYTVACOM0.TVACOMFACL = 0 And xYTVACOM0.TVACOMSTA = "8" Then
        xYTVACOM0.TVACOMAVOIR = fraDétail_Display_TVACOMAVOIR
    End If
Else
    txtUpdate_TVACOMMON.ForeColor = vbBlue
    txtUpdate_TVACOMMONE.ForeColor = vbBlue
    fraUpdate_TVACOMFACL.Visible = False
End If
txtUpdate_TVACOMMTVA = Format$(Abs(xYTVACOM0.TVACOMMTVA), "### ### ### ###.00")
txtUpdate_TVACOMMTVE = Format$(Abs(xYTVACOM0.TVACOMMTVE), "### ### ### ###.00")
If xYTVACOM0.TVACOMMTVA > 0 Then
    txtUpdate_TVACOMMTVA.ForeColor = vbRed
    txtUpdate_TVACOMMTVE.ForeColor = vbRed
Else
    txtUpdate_TVACOMMTVA.ForeColor = vbBlue
    txtUpdate_TVACOMMTVE.ForeColor = vbBlue
End If

txtUpdate_TVACOMFACN = xYTVACOM0.TVACOMFACN
txtUpdate_TVACOMUSR = xYTVACOM0.TVACOMUSR
txtUpdate_TVACOMX = xYTVACOM0.TVACOMXNUR & " - " & xYTVACOM0.TVACOMXUTI & " - " & xYTVACOM0.TVACOMXEVE & " - " & xYTVACOM0.TVACOMXSEQ & " - " & " - " & xYTVACOM0.TVACOMXSPE _
                  & " * " & xYTVACOM0.TVACOMECRX _
                  & " * " & xYTVACOM0.TVACOMGTYP & " - " & xYTVACOM0.TVACOMGORD
cbo_Scan xYTVACOM0.TVACOMCOMB, txtUpdate_TVACOMCOMB


cbo_Scan xYTVACOM0.TVACOMCOME, txtUpdate_TVACOMCOME: txtUpdate_TVACOMCOME.BackColor = txtUsr.BackColor

cbo_Scan xYTVACOM0.TVACOMCLIC, txtUpdate_TVACOMCLIC: txtUpdate_TVACOMCLIC.BackColor = txtUsr.BackColor
txtUpdate_TVACOMCLI = xYTVACOM0.TVACOMCLI: txtUpdate_TVACOMCLI.BackColor = txtUsr.BackColor

V = sqlTVACOMCLI(xYTVACOM0.TVACOMCLIC, xYTVACOM0.TVACOMCLI, xZCLIENA0, xZADRESS0)
fraDétail_Display_TVACOMCLI

cbo_Scan xYTVACOM0.TVACOMCLIP, txtUpdate_TVACOMCLIP: txtUpdate_TVACOMCLIP.BackColor = txtUsr.BackColor
If xYTVACOM0.TVACOMTVAC = "T" Then
    txtUpdate_TVACOMCOMT.Visible = False
    txtUpdate_TVACOMCOMC.Visible = False
    txtUpdate_TVACOMTVAC.Visible = False
Else
    cbo_Scan xYTVACOM0.TVACOMCOMC, txtUpdate_TVACOMCOMC: txtUpdate_TVACOMCOMC.BackColor = txtUsr.BackColor
    txtUpdate_TVACOMCOMC.Visible = True
    cbo_Scan xYTVACOM0.TVACOMCOMT, txtUpdate_TVACOMCOMT: txtUpdate_TVACOMCOMT.BackColor = txtUsr.BackColor
    txtUpdate_TVACOMCOMT.Visible = True
    cbo_Scan xYTVACOM0.TVACOMTVAC, txtUpdate_TVACOMTVAC: txtUpdate_TVACOMTVAC.BackColor = txtUsr.BackColor
    txtUpdate_TVACOMTVAC.Visible = True
End If


cbo_Scan xYTVACOM0.TVACOMSRVR, txtUpdate_TVACOMSRVR
cbo_Scan xYTVACOM0.TVACOMSTA, txtUpdate_TVACOMSTA

V = sqlYBIAMVTHP(xYTVACOM0.TVACOMETA, xYTVACOM0.TVACOMPIE, xYTVACOM0.TVACOMECR, meYBIAMVT0)

If IsNull(V) Then
    libUpdate_TVACOMECR = meYBIAMVT0.LIBELLIB1 & meYBIAMVT0.LIBELLIB2 & meYBIAMVT0.LIBELLIB3 & meYBIAMVT0.LIBELLIB4
Else
    MsgBox V, vbCritical, "BIA_TVA: Affichage du détail d'une commission"
End If

txtUpdate_TVACOMECRX = xYTVACOM0.TVACOMECRX
If xYTVACOM0.TVACOMECRX = 0 Then
    libUpdate_TVACOMECRX = ""
Else
    V = sqlYBIAMVTHP(xYTVACOM0.TVACOMETA, xYTVACOM0.TVACOMPIE, xYTVACOM0.TVACOMECRX, meYBIAMVT0)
    
    If IsNull(V) Then
        libUpdate_TVACOMECRX = meYBIAMVT0.LIBELLIB1 & meYBIAMVT0.LIBELLIB2 & meYBIAMVT0.LIBELLIB3 & meYBIAMVT0.LIBELLIB4
    Else
        MsgBox V, vbCritical, "BIA_TVA: Affichage du détail d'une commission"
    End If
End If
Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail d'une commission"): DoEvents
End Sub






Public Sub lstTVAFACLIEN_Load()
Dim X As String, K As Integer, blnOk As Boolean
Dim xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub mnuPrint0_Facture_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Facture

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint0_Liste1_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrintSelect_Liste

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuPrint1_Liste1_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrintDétail_Liste "1"

Me.Show

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuPrint1_Liste2_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrintDétail_Liste "2"

Me.Show

Me.Enabled = True: Me.MousePointer = 0
End Sub


Private Sub mnuPrint2_Liste1_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrintNIF_Liste "1"

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub txtDétail_TVACOMCLI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtDétail_TVACOMOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtNIF_TVANIFCLI_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtParam_Update_TVACOMOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtParamUpdate_CLIENACLI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtParamUpdate_CLIENARES_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_TVAFACCLI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtUpdate_TVACOMCLI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub

Private Sub txtUpdate_TVACOMCOMC_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub



Public Sub cmdSelect_SQL_6_Regroupement()
Dim blnOk As Boolean
Dim K0 As Integer, K As Integer

Call rsYTVACOM0_Init(arrYTVACOM0(0))
Call rsYTVAFAC0_Init(newYTVAFAC0)
blnOk = False
K0 = 1
For K = 1 To arrYTVACOM0_Nb
    If newYTVAFAC0.TVAFACETA <> arrYTVACOM0(K).TVACOMETA _
    Or newYTVAFAC0.TVAFACCLIC <> arrYTVACOM0(K).TVACOMCLIC _
    Or newYTVAFAC0.TVAFACCLI <> arrYTVACOM0(K).TVACOMCLI _
    Or newYTVAFAC0.TVAFACCLIP <> arrYTVACOM0(K).TVACOMCLIP Then
        If blnOk Then Call cmdSelect_SQL_6_Regroupement_Insert(K0, K - 1)
        K0 = K
        blnOk = True
        Call rsYTVAFAC0_Init(newYTVAFAC0)
        newYTVAFAC0.TVAFACETA = arrYTVACOM0(K).TVACOMETA
        newYTVAFAC0.TVAFACCLIC = arrYTVACOM0(K).TVACOMCLIC
        newYTVAFAC0.TVAFACCLI = arrYTVACOM0(K).TVACOMCLI
        newYTVAFAC0.TVAFACCLIP = arrYTVACOM0(K).TVACOMCLIP
        cmdSelect_SQL_6_Regroupement_NIF
    End If
    newYTVAFAC0.TVAFACMTTC = newYTVAFAC0.TVAFACMTTC + arrYTVACOM0(K).TVACOMMONE + arrYTVACOM0(K).TVACOMMTVE
    newYTVAFAC0.TVAFACMTVA = newYTVAFAC0.TVAFACMTVA + arrYTVACOM0(K).TVACOMMTVE
    If arrYTVACOM0(K).TVACOMMTVE = 0 Then
        newYTVAFAC0.TVAFACMEXO = newYTVAFAC0.TVAFACMEXO + arrYTVACOM0(K).TVACOMMONE
    End If
    
    
Next K
If blnOk Then Call cmdSelect_SQL_6_Regroupement_Insert(K0, K - 1)

End Sub

Public Sub cmdSelect_SQL_6_Regroupement_Insert(lK0 As Integer, lK As Integer)
Dim V, X As String, xSql As String
Dim Nb As Long, I As Integer
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_SQL_6_Regroupement_Insert"
'-------------------------------------------------------

mMsgBox = newYTVAFAC0.TVAFACCLIC & newYTVAFAC0.TVAFACCLI & " / " & lK0 & "-" & lK

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Ajout MAD : " & mMsgBox): DoEvents
'________________________________________________________________________________
V = sqlYTVAFAC0_Init(newYTVAFAC0)
If Not IsNull(V) Then GoTo Error_MsgBox

V = sqlYTVAFAC0_Insert(newYTVAFAC0)
If Not IsNull(V) Then GoTo Error_MsgBox

For I = lK0 To lK
    newYTVACOM0 = arrYTVACOM0(I)
    newYTVACOM0.TVACOMFACN = newYTVAFAC0.TVAFACFACN
    newYTVACOM0.TVACOMSTA = "F"
    V = sqlYTVACOM0_Update(newYTVACOM0, arrYTVACOM0(I))
    If Not IsNull(V) Then GoTo Error_MsgBox
Next I
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    

End Sub

Private Sub txtUpdate_TVACOMECRX_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtUpdate_TVAFACCLIT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub



Public Sub fraDétail_Display_TVACOMCLI()
libUpdate_TVACOMCLI = Trim(xZADRESS0.ADRESSRA1) & Trim(xZADRESS0.ADRESSRA2) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD1) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD2) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD3) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSCOP) _
                    & " " & Trim(xZADRESS0.ADRESSVIL) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSPAY) _
                    & vbCrLf & "_______________________________________________" _
                    & vbCrLf & "agent économique : " & xZCLIENA0.CLIENAECO _
                    & vbCrLf & "catégorie client      : " & xZCLIENA0.CLIENACAT _
                    & vbCrLf & "pays résidence       : " & xZCLIENA0.CLIENARSD

End Sub
Public Sub fraNIF_Display_TVANIFCLI()
libUpdate_TVANIFCLI = Trim(xZADRESS0.ADRESSRA1) & Trim(xZADRESS0.ADRESSRA2) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD1) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD2) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD3) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSCOP) _
                    & " " & Trim(xZADRESS0.ADRESSVIL) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSPAY) _
                    & vbCrLf & "_______________________________________________" _
                    & vbCrLf & "agent économique : " & xZCLIENA0.CLIENAECO _
                    & vbCrLf & "catégorie client      : " & xZCLIENA0.CLIENACAT _
                    & vbCrLf & "pays résidence       : " & xZCLIENA0.CLIENARSD

End Sub

Public Sub fraSelect_Display_TVACOMCLI()
libUpdate_TVAFACCLI = Trim(xZADRESS0.ADRESSRA1) & Trim(xZADRESS0.ADRESSRA2) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD1) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD2) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSAD3) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSCOP) _
                    & " " & Trim(xZADRESS0.ADRESSVIL) _
                    & vbCrLf & Trim(xZADRESS0.ADRESSPAY) _
                    & vbCrLf & "_______________________________________________" _
                    & vbCrLf & "agent économique : " & xZCLIENA0.CLIENAECO _
                    & vbCrLf & "catégorie client      : " & xZCLIENA0.CLIENACAT _
                    & vbCrLf & "pays résidence       : " & xZCLIENA0.CLIENARSD

End Sub



Public Sub cmdSelect_SQL_6_Regroupement_NIF()
Dim xSql As String

newYTVAFAC0.TVAFACSTA = "0"
If Not TVANIFCLIT_Pays(newYTVAFAC0.TVAFACCLIP) Then newYTVAFAC0.TVAFACSTA = "2": Exit Sub
         

If newYTVAFAC0.TVAFACCLIC = " " Then
    xSql = "select CLIFISNIF from " & paramIBM_Library_SAB & ".ZCLIFIS0 " _
         & " where CLIFISETA =" & newYTVAFAC0.TVAFACETA _
         & " and   CLIFISCLI = '" & newYTVAFAC0.TVAFACCLI & "'"
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        newYTVAFAC0.TVAFACCLIT = rsSab("CLIFISNIF")
        newYTVAFAC0.TVAFACSTA = "1"
    End If
Else
    xSql = "select TVANIFCLIT from " & paramIBM_Library_SABSPE & ".YTVANIF0 " _
         & " where TVANIFCLIC ='" & newYTVAFAC0.TVAFACCLIC & "'" _
         & " and   TVANIFCLI = '" & newYTVAFAC0.TVAFACCLI & "'"
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        newYTVAFAC0.TVAFACCLIT = rsSab("TVANIFCLIT")
        newYTVAFAC0.TVAFACSTA = "1"
    End If

End If

If Trim(newYTVAFAC0.TVAFACCLIT) <> "" Then
    V = TVANIFCLIT_Control(newYTVAFAC0.TVAFACCLIT)
    If Not IsNull(V) Then
        newYTVAFAC0.TVAFACCLIT = ""
        newYTVAFAC0.TVAFACSTA = "9"
    End If
End If

End Sub

Public Sub lstNIF_Load_1()
On Error GoTo Error_Handler
fgNIF.Visible = False
cmdPrint.Enabled = False
currentAction = "lstNIF_Load_1"
cmdNIF_Ok_Caption = "Lancer la requête"
cmdNIF_Ok.Caption = cmdNIF_Ok_Caption
cmdNIF_Ok.Visible = True
txtNIF_TVANIFSTA.Enabled = True
txtNIF_TVANIFCLI.Enabled = True
chkNIF_TVANIFCLIT.Visible = False
fraNIF_Options_1.Visible = True
fraNIF_Options_1.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Public Sub lstParam_Load_1()
On Error GoTo Error_Handler
fgParam.Visible = False
cmdPrint.Enabled = False
currentAction = "lstParam_Load_1"
cmdParam_Ok_Caption = "Lancer la requête"
cmdParam_Ok.Caption = cmdParam_Ok_Caption
cmdParam_Ok.Visible = True
fraParam_Options_1.Visible = True
fraParam_Options_1.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Public Sub lstParam_Load_2()
On Error GoTo Error_Handler
fgParam.Visible = False
cmdPrint.Enabled = False
currentAction = "lstParam_Load_2"
cmdParam_Ok_Caption = "Lancer la requête"
cmdParam_Ok.Caption = cmdParam_Ok_Caption
cmdParam_Ok.Visible = True
fraParam_Options_1.Visible = True
fraParam_Options_1.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub lstNIF_Load_2()
On Error GoTo Error_Handler
fgNIF.Visible = False
cmdPrint.Enabled = False
currentAction = "lstNIF_Load_2"
cmdNIF_Ok_Caption = "Lancer la requête"
cmdNIF_Ok.Caption = cmdNIF_Ok_Caption
cmdNIF_Ok.Visible = True
txtNIF_TVANIFSTA.Enabled = True
txtNIF_TVANIFCLI.Enabled = True
chkNIF_TVANIFCLIT.Visible = False
fraNIF_Options_1.Visible = True
fraNIF_Options_1.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub lstNIF_Load_3()
On Error GoTo Error_Handler
fgNIF.Visible = False
cmdPrint.Enabled = False
currentAction = "lstNIF_Load_3"
cmdNIF_Ok_Caption = "Lancer la requête"
cmdNIF_Ok.Caption = cmdNIF_Ok_Caption
cmdNIF_Ok.Visible = True
txtNIF_TVANIFSTA.Enabled = False
txtNIF_TVANIFCLI.Enabled = False
fraNIF_Options_1.Visible = True
fraNIF_Options_1.Enabled = True
chkNIF_TVANIFCLIT.Visible = True
chkNIF_TVANIFCLIT.Value = "1"

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_1()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

ReDim arrYTVANIF0(101)
arrYTVANIF0_Max = 100: arrYTVANIF0_Nb = 0
Call lstErr_Clear(lstErr, cmdContext, "cmdNIF_SQL"): DoEvents

currentAction = "cmdNIF_SQL"
xWhere = ""

X = Trim(Mid$(txtNIF_TVANIFSTA, 1, 1))
If X <> "" Then xWhere = xWhere & " and  TVANIFSTA = '" & X & "'"

Set rsSab = Nothing

X = Mid$(txtNIF_TVANIFCLIC, 1, 1)
Select Case X
    Case " ": cmdNIF_SQL_1_ZCLIENA0
    Case "D": cmdNIF_SQL_1_ZCDOTIE0
    Case "G": cmdNIF_SQL_1_ZCHGPAS0
    Case "R": cmdNIF_SQL_1_ZENCTIE0
End Select
fgNIF_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub cmdParam_SQL()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

ReDim arrYBIATAB0(101)
arrYBIATAB0_Max = 100: arrYBIATAB0_Nb = 0
Call lstErr_Clear(lstErr, cmdContext, "cmdParam_SQL"): DoEvents

currentAction = "cmdParam_SQL"
Select Case cmdParam_SQL_K
    Case 1: xWhere = " where BIATABID = 'TVAFACSTA'"
    Case 2: xWhere = " where BIATABID = 'TVACOMSTA'"
End Select
Set rsSab = Nothing

arrYBIATAB0_SQL xWhere & " order by BIATABK1,BIATABK2"

fgParam_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 3
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub cmdNIF_SQL_2()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Call lstErr_Clear(lstErr, cmdContext, "cmdNIF_SQL"): DoEvents
ReDim arrYTVANIF0(101)
arrYTVANIF0_Max = 100: arrYTVANIF0_Nb = 0
fraNIF_Update.Visible = False

currentAction = "cmdNIF_SQL_2"
xWhere = ""

fgNIF_Reset

X = Mid$(txtNIF_TVANIFCLIC, 1, 1)
Select Case Mid$(txtNIF_TVANIFCLIC, 1, 1)
    Case "D": cmdNIF_SQL_2_ZCDOTIE0
    Case "G": cmdNIF_SQL_2_ZCHGPAS0
    Case "R": cmdNIF_SQL_2_ZENCTIE0
    Case " ": cmdNIF_SQL_2_ZCLIENA0
    Case Else: Call lstErr_AddItem(lstErr, cmdContext, "Préciser l'origine du tiers B D G"): DoEvents
End Select

fgNIF_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_3()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler
currentAction = "cmdNIF_SQL_3"

Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
ReDim arrYTVANIF0(101)
arrYTVANIF0_Max = 100: arrYTVANIF0_Nb = 0
fraNIF_Update.Visible = False

xWhere = ""

fgNIF_Reset

X = Mid$(txtNIF_TVANIFCLIC, 1, 1)
Select Case Mid$(txtNIF_TVANIFCLIC, 1, 1)
    Case "D": cmdNIF_SQL_3_ZCDOTIE0
    Case "G": cmdNIF_SQL_3_ZCHGPAS0
    Case "R": cmdNIF_SQL_3_ZENCTIE0
    Case " ": cmdNIF_SQL_3_ZCLIENA0
    Case Else: Call lstErr_AddItem(lstErr, cmdContext, "Préciser l'origine du tiers B D G"): DoEvents
End Select

fgNIF_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_2_ZCDOTIE0()
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

X = Trim(txtNIF_TVANIFCLI)
If Len(X) < 4 Then
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser les critères de recherche")
    Exit Sub
End If
If IsNumeric(X) Then
    X = " where CDOTIETIE = '" & Format(X, "0000000") & "'"
Else
    X = " where CDOTIERA1 like '%" & X & "%'"
End If

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOTIE0 left outer join " & paramIBM_Library_SABSPE & ".YTVANIF0 on TVANIFCLI = CDOTIETIE " & X
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    If IsNull(rsSab("TVANIFCLI")) Then
        rsYTVANIF0_Init xYTVANIF0
        xYTVANIF0.TVANIFCLIC = "D"
        xYTVANIF0.TVANIFCLI = Format(rsSab("CDOTIETIE"), "0000000")
        xYTVANIF0.TVANIFCLIP = Trim(rsSab("CDOTIEPAR"))
        xYTVANIF0.TVANIFRS = Trim(rsSab("CDOTIERA1")) & Trim(rsSab("CDOTIERA2"))
        V = Null
    Else
        V = rsYTVANIF0_GetBuffer(rsSab, xYTVANIF0)
    End If
    
     If Not IsNull(V) Then
         MsgBox V, vbCritical, "cmdNIF_SQL_2_ZCDOTIE0"
        '' Exit Sub
     Else
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         xYTVANIF0.TVANIFCLIP = Trim(rsSab("CDOTIEPAR"))
         xYTVANIF0.TVANIFRS = Trim(rsSab("CDOTIERA1")) & Trim(rsSab("CDOTIERA2"))
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
         
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_2_ZENCTIE0()
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

X = Trim(txtNIF_TVANIFCLI)
If Len(X) < 4 Then
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser les critères de recherche")
    Exit Sub
End If
If IsNumeric(X) Then
    X = " where ENCTIETIE = '" & Format(X, "0000000") & "'"
Else
    X = " where ENCTIERA1 like '%" & X & "%'"
End If

xSql = "select * from " & paramIBM_Library_SAB & ".ZENCTIE0 left outer join " & paramIBM_Library_SABSPE & ".YTVANIF0 on TVANIFCLI = ENCTIETIE " & X
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    If IsNull(rsSab("TVANIFCLI")) Then
        rsYTVANIF0_Init xYTVANIF0
        xYTVANIF0.TVANIFCLIC = "R"
        xYTVANIF0.TVANIFCLI = Format(rsSab("ENCTIETIE"), "0000000")
        xYTVANIF0.TVANIFCLIP = Trim(rsSab("ENCTIEPAR"))
        xYTVANIF0.TVANIFRS = Trim(rsSab("ENCTIERA1")) & Trim(rsSab("ENCTIERA2"))
        V = Null
    Else
        V = rsYTVANIF0_GetBuffer(rsSab, xYTVANIF0)
    End If
    
     If Not IsNull(V) Then
         MsgBox V, vbCritical, "cmdNIF_SQL_2_ZENCTIE0"
        '' Exit Sub
     Else
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         xYTVANIF0.TVANIFCLIP = Trim(rsSab("ENCTIEPAR"))
         xYTVANIF0.TVANIFRS = Trim(rsSab("ENCTIERA1")) & Trim(rsSab("ENCTIERA2"))
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
         
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_2_ZCHGPAS0()
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

X = Trim(txtNIF_TVANIFCLI)
If Len(X) < 3 Then
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser les critères de recherche")
    Exit Sub
End If
If IsNumeric(X) Then
    X = " where CHGPASNU = '" & Format(X, "0000000") & "'"
Else
    X = " where CHGPASN1 like '%" & X & "%'"
End If

xSql = "select * from " & paramIBM_Library_SAB & ".ZCHGPAS0 left outer join " & paramIBM_Library_SABSPE & ".YTVANIF0 on TVANIFCLI = CHGPASNU " & X
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    If IsNull(rsSab("TVANIFCLI")) Then
        rsYTVANIF0_Init xYTVANIF0
        xYTVANIF0.TVANIFCLIC = "G"
        xYTVANIF0.TVANIFCLI = Format(rsSab("CHGPASNU"), "0000000")
        xYTVANIF0.TVANIFCLIP = Trim(rsSab("CHGPASRE"))
        xYTVANIF0.TVANIFRS = Trim(rsSab("CHGPASN1")) & Trim(rsSab("CHGPASN2"))
        V = Null
    Else
        V = rsYTVANIF0_GetBuffer(rsSab, xYTVANIF0)
    End If
    
     If Not IsNull(V) Then
         MsgBox V, vbCritical, "cmdNIF_SQL_2_ZCHGPAS0"
        '' Exit Sub
     Else
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         xYTVANIF0.TVANIFCLIP = Trim(rsSab("CHGPASRE"))
         xYTVANIF0.TVANIFRS = Trim(rsSab("CHGPASN1")) & Trim(rsSab("CHGPASN2"))
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
         
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_2_ZCLIENA0()
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

X = Trim(txtNIF_TVANIFCLI)
If IsNumeric(X) Then
    X = " where CLIENACLI = '" & Format(X, "0000000") & "'"
Else
    X = " where CLIENARA1 like '%" & X & "%'"
End If

xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 left outer join " & paramIBM_Library_SAB & ".ZCLIFIS0 on CLIFISCLI = CLIENACLI " & X
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = " "
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLI = Format(rsSab("CLIENACLI"), "0000000")
    xYTVANIF0.TVANIFCLIP = Trim(rsSab("CLIENARSD"))
    xYTVANIF0.TVANIFRS = Trim(rsSab("CLIENARA1")) & Trim(rsSab("CLIENARA2"))
    If Not IsNull(rsSab("CLIFISNIF")) Then xYTVANIF0.TVANIFCLIT = Trim(rsSab("CLIFISNIF"))
    
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
         
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_3_ZCLIENA0()
Dim V, blnTVANIFCLIT As Boolean, blnOk As Boolean
Dim X As String, xSql As String
On Error GoTo Error_Handler

xSql = "select distinct TVACOMCLI,TVACOMCLIP,CLIENARA1,CLIENARA2,CLIFISNIF from " _
    & paramIBM_Library_SABSPE & ".YTVACOM0" _
    & " left join " & paramIBM_Library_SAB & ".ZCLIENA0 on CLIENACLI = TVACOMCLI " _
    & " left join " & paramIBM_Library_SAB & ".ZCLIFIS0 on CLIFISCLI = TVACOMCLI " _
    & " where TVACOMSTA not in ('F','I','X','A') and TVACOMCLIC = ' '"
    
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = " "
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLI = rsSab("TVACOMCLI")
    xYTVANIF0.TVANIFCLIP = rsSab("TVACOMCLIP")
    If IsNull(rsSab("CLIENARA1")) Then
        xYTVANIF0.TVANIFSTA = "1"
    Else
        xYTVANIF0.TVANIFRS = Trim(rsSab("CLIENARA1")) & Trim(rsSab("CLIENARA2"))
    End If
    If IsNull(rsSab("CLIFISNIF")) Then
        xYTVANIF0.TVANIFSTA = " "
    Else
        xYTVANIF0.TVANIFCLIT = Trim(rsSab("CLIFISNIF"))
    End If
    
    blnTVANIFCLIT = TVANIFCLIT_Pays(xYTVANIF0.TVANIFCLIP)
    If Not blnTVANIFCLIT Then xYTVANIF0.TVANIFSTA = "I"
    If chkNIF_TVANIFCLIT.Value = 0 Then
        blnOk = True
    Else
        If xYTVANIF0.TVANIFSTA = " " Then
            blnOk = True
        Else
            blnOk = False
        End If
        
    End If
    
    If blnOk Then
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_3_ZCDOTIE0()
Dim V, blnTVANIFCLIT As Boolean, blnOk As Boolean
Dim X As String, xSql As String
On Error GoTo Error_Handler

xSql = "select distinct (TVACOMCLI),TVACOMCLIP,CDOTIERA1,CDOTIERA2,TVANIFCLIT,TVACOMCLIC from " _
    & paramIBM_Library_SABSPE & ".YTVACOM0" _
    & " left join " & paramIBM_Library_SAB & ".ZCDOTIE0 on CDOTIETIE = TVACOMCLI " _
    & " left join " & paramIBM_Library_SABSPE & ".YTVANIF0 on TVANIFCLI = TVACOMCLI and TVANIFCLIC = 'D'" _
    & " where TVACOMSTA not in ('F','I','X','A') and TVACOMCLIC = 'D'"
    
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = "D"
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLIC = rsSab("TVACOMCLIC")
    xYTVANIF0.TVANIFCLI = rsSab("TVACOMCLI")
    xYTVANIF0.TVANIFCLIP = rsSab("TVACOMCLIP")
    If IsNull(rsSab("CDOTIERA1")) Then
        xYTVANIF0.TVANIFSTA = "1"
    Else
        xYTVANIF0.TVANIFRS = Trim(rsSab("CDOTIERA1")) & Trim(rsSab("CDOTIERA2"))
    End If
    If IsNull(rsSab("TVANIFCLIT")) Then
        xYTVANIF0.TVANIFSTA = " "
    Else
        xYTVANIF0.TVANIFCLIT = Trim(rsSab("TVANIFCLIT"))
    End If
    
    blnTVANIFCLIT = TVANIFCLIT_Pays(xYTVANIF0.TVANIFCLIP)
    If Not blnTVANIFCLIT Then xYTVANIF0.TVANIFSTA = "I"
    If chkNIF_TVANIFCLIT.Value = 0 Then
        blnOk = True
    Else
        If xYTVANIF0.TVANIFSTA = " " Then
            blnOk = True
        Else
            blnOk = False
        End If
        
    End If
    
    If blnOk Then
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_3_ZENCTIE0()
Dim V, blnTVANIFCLIT As Boolean, blnOk As Boolean
Dim X As String, xSql As String
On Error GoTo Error_Handler

xSql = "select distinct (TVACOMCLI),TVACOMCLIP,ENCTIERA1,ENCTIERA2,TVANIFCLIT,TVACOMCLIC from " _
    & paramIBM_Library_SABSPE & ".YTVACOM0" _
    & " left join " & paramIBM_Library_SAB & ".ZENCTIE0 on ENCTIETIE = TVACOMCLI " _
    & " left join " & paramIBM_Library_SABSPE & ".YTVANIF0 on TVANIFCLI = TVACOMCLI and TVANIFCLIC = 'R'" _
    & " where TVACOMSTA not in ('F','I','X','A') and TVACOMCLIC = 'R'"
    
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = "R"
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLIC = rsSab("TVACOMCLIC")
    xYTVANIF0.TVANIFCLI = rsSab("TVACOMCLI")
    xYTVANIF0.TVANIFCLIP = rsSab("TVACOMCLIP")
    If IsNull(rsSab("ENCTIERA1")) Then
        xYTVANIF0.TVANIFSTA = "1"
    Else
        xYTVANIF0.TVANIFRS = Trim(rsSab("ENCTIERA1")) & Trim(rsSab("ENCTIERA2"))
    End If
    If IsNull(rsSab("TVANIFCLIT")) Then
        xYTVANIF0.TVANIFSTA = " "
    Else
        xYTVANIF0.TVANIFCLIT = Trim(rsSab("TVANIFCLIT"))
    End If
    
    blnTVANIFCLIT = TVANIFCLIT_Pays(xYTVANIF0.TVANIFCLIP)
    If Not blnTVANIFCLIT Then xYTVANIF0.TVANIFSTA = "I"
    If chkNIF_TVANIFCLIT.Value = 0 Then
        blnOk = True
    Else
        If xYTVANIF0.TVANIFSTA = " " Then
            blnOk = True
        Else
            blnOk = False
        End If
        
    End If
    
    If blnOk Then
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_3_ZCHGPAS0()
Dim V, blnTVANIFCLIT As Boolean, blnOk As Boolean
Dim X As String, xSql As String
On Error GoTo Error_Handler

xSql = "select distinct (TVACOMCLI),TVACOMCLIP,CHGPASN1,CHGPASN2,TVANIFCLIT,TVACOMCLIC from " _
    & paramIBM_Library_SABSPE & ".YTVACOM0" _
    & " left join " & paramIBM_Library_SAB & ".ZCHGPAS0 on CHGPASNU = TVACOMCLI " _
    & " left join " & paramIBM_Library_SABSPE & ".YTVANIF0 on TVANIFCLI = TVACOMCLI and TVANIFCLIC = 'G'" _
    & " where TVACOMSTA not in ('F','I','X','A') and TVACOMCLIC = 'G'"
    
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = "G"
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLIC = rsSab("TVACOMCLIC")
    xYTVANIF0.TVANIFCLI = rsSab("TVACOMCLI")
    xYTVANIF0.TVANIFCLIP = rsSab("TVACOMCLIP")
    If IsNull(rsSab("CHGPASN1")) Then
        xYTVANIF0.TVANIFSTA = "1"
    Else
        xYTVANIF0.TVANIFRS = Trim(rsSab("CHGPASN1")) & Trim(rsSab("CHGPASN2"))
    End If
    If IsNull(rsSab("TVANIFCLIT")) Then
        xYTVANIF0.TVANIFSTA = " "
    Else
        xYTVANIF0.TVANIFCLIT = Trim(rsSab("TVANIFCLIT"))
    End If
    
    blnTVANIFCLIT = TVANIFCLIT_Pays(xYTVANIF0.TVANIFCLIP)
    If Not blnTVANIFCLIT Then xYTVANIF0.TVANIFSTA = "I"
    If chkNIF_TVANIFCLIT.Value = 0 Then
        blnOk = True
    Else
        If xYTVANIF0.TVANIFSTA = " " Then
            blnOk = True
        Else
            blnOk = False
        End If
        
    End If
    
    If blnOk Then
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_1_ZCLIENA0()
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

X = Trim(txtNIF_TVANIFCLI)
If IsNumeric(X) Then
    X = " where CLIENACLI = '" & Format(X, "0000000") & "'"
Else
    X = " where CLIENARA1 like '%" & X & "%'"
End If

xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 inner join " & paramIBM_Library_SAB & ".ZCLIFIS0 on CLIFISCLI = CLIENACLI " & X
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = " "
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLI = Format(rsSab("CLIENACLI"), "0000000")
    xYTVANIF0.TVANIFCLIP = Trim(rsSab("CLIENARSD"))
    xYTVANIF0.TVANIFRS = Trim(rsSab("CLIENARA1")) & Trim(rsSab("CLIENARA2"))
    If Not IsNull(rsSab("CLIFISNIF")) Then xYTVANIF0.TVANIFCLIT = Trim(rsSab("CLIFISNIF"))
    
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
         
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub cmdNIF_SQL_1_ZCDOTIE0()
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

X = Trim(txtNIF_TVANIFCLI)
If IsNumeric(X) Then
    X = " where  TVANIFCLIC = 'D' and CDOTIETIE = '" & Format(X, "0000000") & "'"
Else
    X = " where TVANIFCLIC = 'D' and CDOTIERA1 like '%" & X & "%'"
End If

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVANIF0 inner join " & paramIBM_Library_SAB & ".ZCDOTIE0 on TVANIFCLI = CDOTIETIE " & X
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = "D"
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLI = Format(rsSab("CDOTIETIE"), "0000000")
    xYTVANIF0.TVANIFCLIP = Trim(rsSab("CDOTIEPAR"))
    xYTVANIF0.TVANIFRS = Trim(rsSab("CDOTIERA1")) & Trim(rsSab("CDOTIERA2"))
    xYTVANIF0.TVANIFCLIT = Trim(rsSab("TVANIFCLIT"))
    
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
         
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_1_ZENCTIE0()
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

X = Trim(txtNIF_TVANIFCLI)
If IsNumeric(X) Then
    X = " where  TVANIFCLIC = 'R' and ENCTIETIE = '" & Format(X, "0000000") & "'"
Else
    X = " where TVANIFCLIC = 'R' and ENCTIERA1 like '%" & X & "%'"
End If

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVANIF0 inner join " & paramIBM_Library_SAB & ".ZENCTIE0 on TVANIFCLI = ENCTIETIE " & X
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = "R"
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLI = Format(rsSab("ENCTIETIE"), "0000000")
    xYTVANIF0.TVANIFCLIP = Trim(rsSab("ENCTIEPAR"))
    xYTVANIF0.TVANIFRS = Trim(rsSab("ENCTIERA1")) & Trim(rsSab("ENCTIERA2"))
    xYTVANIF0.TVANIFCLIT = Trim(rsSab("TVANIFCLIT"))
    
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
         
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdNIF_SQL_1_ZCHGPAS0()
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

X = Trim(txtNIF_TVANIFCLI)
If IsNumeric(X) Then
    X = " where TVANIFCLIC = 'G' and CHGPASNU = " & Format(X, "0000000")
Else
    X = " where TVANIFCLIC = 'G' and CHGPASn1 like '%" & X & "%'"
End If

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVANIF0 inner join " & paramIBM_Library_SAB & ".ZCHGPAS0 on TVANIFCLI = CHGPASNU " & X
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    rsYTVANIF0_Init xYTVANIF0
    xYTVANIF0.TVANIFCLIC = "G"
    xYTVANIF0.TVANIFSTA = "V"
    xYTVANIF0.TVANIFCLI = Format(rsSab("CHGPASNU"), "0000000")
    xYTVANIF0.TVANIFCLIP = Trim(rsSab("CHGPASRE"))
    xYTVANIF0.TVANIFRS = Trim(rsSab("CHGPASN1")) & Trim(rsSab("CHGPASN2"))
    xYTVANIF0.TVANIFCLIT = Trim(rsSab("TVANIFCLIT"))
    
        arrYTVANIF0_Nb = arrYTVANIF0_Nb + 1
        If arrYTVANIF0_Nb > arrYTVANIF0_Max Then
            arrYTVANIF0_Max = arrYTVANIF0_Max + 50
            ReDim Preserve arrYTVANIF0(arrYTVANIF0_Max)
        End If
         arrYTVANIF0(arrYTVANIF0_Nb) = xYTVANIF0
         
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 2
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fraNIF_Display()
Dim V
Dim X As String, X1 As String
fraNIF_Update.Visible = True
fraNIF_Update_A.Enabled = False
fraNIF_Update_B.Enabled = True

If xYTVANIF0.TVANIFCLIC = " " Then
    fraNIF_Update_B.Enabled = False
Else
    If Trim(xYTVANIF0.TVANIFCLIT) = "" Then
        fraNIF_Update_B.Enabled = BIA_TVANIF_Aut.Saisir
    Else
        fraNIF_Update_B.Enabled = BIA_TVANIF_Aut.Valider
    End If
    
End If
cbo_Scan xYTVANIF0.TVANIFCLIC, txtUpdate_TVANIFCLIC: txtUpdate_TVANIFCLIC.BackColor = txtUsr.BackColor
txtUpdate_TVANIFCLI = xYTVANIF0.TVANIFCLI: txtUpdate_TVANIFCLI.BackColor = txtUsr.BackColor
V = sqlTVACOMCLI(xYTVANIF0.TVANIFCLIC, xYTVANIF0.TVANIFCLI, xZCLIENA0, xZADRESS0)
fraNIF_Display_TVANIFCLI
If xYTVANIF0.TVANIFCLIF = " " Then
    optUpdate_TVANIFCLIF.Value = True
Else
    optUpdate_TVANIFCLIF_Idem.Value = True
End If
optUpdate_TVANIFCLIF_Idem.Enabled = False

txtUpdate_TVANIFCLIT = Trim(TVANIFCLIT_Format(xYTVANIF0.TVANIFCLIT))
If txtUpdate_TVANIFCLIT = "" Then txtUpdate_TVANIFCLIT = xYTVANIF0.TVANIFCLIP
txtUpdate_TVANIFCLIT.BackColor = txtUsr.BackColor

txtUpdate_TVANIFCLIT.BackColor = txtUsr.BackColor

cbo_Scan xYTVANIF0.TVANIFSTA, txtUpdate_TVANIFSTA
txtUpdate_TVANIFUSR = xYTVANIF0.TVANIFUSR
Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail d'une facture"): DoEvents
End Sub

Public Sub fraParam_Display()
Dim V
Dim X As String, X1 As String
fraParam_Update.Visible = True
fraParam_Update_A.Enabled = False
fraParam_Update_B1.Visible = False
fraParam_Update_B2.Visible = False
chkParamUpdate_Insert.Enabled = True
chkParamUpdate_Insert = "0"
Select Case cmdParam_SQL_K
    Case 1: fraParam_Display_1
    Case 2: fraParam_Display_2
End Select
    
Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail PARAM"): DoEvents
End Sub


Public Function cmdNIF_Update_Ok_Transaction(lFct As String)
Dim V, X As String, xSql As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdNIF_Update_Ok_Transaction"
'-------------------------------------------------------
cmdNIF_Update_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lFct
    Case "Update": V = sqlYTVANIF0_Update(newYTVANIF0, oldYTVANIF0)
    Case "Insert": V = sqlYTVANIF0_Insert(newYTVANIF0)
    Case "Delete": V = sqlYTVANIF0_Delete(oldYTVANIF0)
    Case Else: V = "? fct non traitée : " & lFct
End Select
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdNIF_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Public Function cmdParam_Update_Ok_Transaction(lFct As String)
Dim V, X As String, xSql As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdParam_Update_Ok_Transaction"
'-------------------------------------------------------
cmdParam_Update_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'V = cnSAB_Transaction("BeginTrans")
'If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lFct
    Case "Update": V = sqlYBIATAB0_Update(newYBIATAB0, oldYBIATAB0)
    Case "Insert": V = sqlYBIATAB0_Insert(newYBIATAB0)
    Case "Delete": V = sqlYBIATAB0_Delete(oldYBIATAB0)
    Case Else: V = "? fct non traitée : " & lFct
End Select
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
'    If Not IsNull(V) Then
'        V = cnSAB_Transaction("Rollback")
'    Else
'        V = cnSAB_Transaction("Commit")
'    End If
    
    cmdParam_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function


Private Sub txtUpdate_TVANIFCLIT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Public Function TVANIFCLIT_Pays(lPays As String) As Boolean
Static mTVAFACCLIP As String, blnFiscal As Boolean
Dim K As Integer
If mTVAFACCLIP = lPays Then TVANIFCLIT_Pays = blnFiscal: Exit Function

mTVAFACCLIP = lPays
For K = 0 To arrPays_NB
    If lPays = arrPays(K).Id Then
        mTVAFACCLIP_Code_Fiscal = arrPays(K).Fiscal
        If arrPays(K).Fiscal <= "5" Then
            TVANIFCLIT_Pays = True: blnFiscal = True
        Else
            TVANIFCLIT_Pays = False: blnFiscal = False
        End If
        Exit Function
    End If
Next K

TVANIFCLIT_Pays = False
blnFiscal = False
End Function

Public Sub fraParam_Display_1()
fraParam_Update_B1.Visible = True
txtParamUpdate_CLIENACLI.Enabled = False
txtParamUpdate_CLIENACLI.BackColor = txtUsr.BackColor
txtParamUpdate_CLIENACLI = Trim(xYBIATAB0.BIATABK2)
txtParamUpdate_CLIENARES.BackColor = txtUsr.BackColor
txtParamUpdate_CLIENARES = Trim(xYBIATAB0.BIATABTXT)
fraParam_Update_B.Enabled = BIA_TVAPARAM_Aut.Valider

End Sub
Public Sub fraParam_Display_2()
fraParam_Update_B2.Visible = True
fraParam_Update_TVACOMOPE.Enabled = False

Select Case Trim(xYBIATAB0.BIATABK1)
    Case "CRE": optParam_Update_TVACOMOPE_CRE = True
    Case Else: optParam_Update_TVACOMOPE_ENG = True
End Select
Select Case Mid$(xYBIATAB0.BIATABTXT, 1, 1)
    Case "V": optParam_Update_TVACOMSTA_V = True
    Case Else: optParam_Update_TVACOMSTA_I = True
End Select
txtParam_Update_TVACOMOPE = Trim(xYBIATAB0.BIATABK2)
fraParam_Update_B.Enabled = BIA_TVAPARAM_Aut.Saisir
End Sub


Public Function fraDétail_Display_TVACOMAVOIR() As String
Dim xWhere As String, xSql As String
Dim wL As Long, wTVACOMFACL As Long
Dim blnOk As Boolean, blnTVAFACN As Boolean, blnDoublon As Boolean

fraDétail_Display_TVACOMAVOIR = "?"
If Mid$(xYTVACOM0.TVACOMOPE, 1, 1) = "*" Then Exit Function
wTVACOMFACL = 0
blnOk = False
blnTVAFACN = False
blnDoublon = False

'recherche même : opé,client code com
xWhere = "where TVACOMOPE = '" & xYTVACOM0.TVACOMOPE & "' and TVACOMDOS = " & xYTVACOM0.TVACOMDOS _
      & " and TVACOMCLIC = '" & xYTVACOM0.TVACOMCLIC & "' and TVACOMCLI = '" & xYTVACOM0.TVACOMCLI & "'" _
      & " and TVACOMCOMC = '" & xYTVACOM0.TVACOMCOMC & "'"

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVACOM0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    If xYTVACOM0.TVACOMPIE = rsSab("TVACOMPIE") And xYTVACOM0.TVACOMECR = rsSab("TVACOMECR") Then
    Else
        wL = rsSab("TVACOMFACN")
        If blnOk Then
            blnDoublon = True
        Else
            blnOk = True
            wTVACOMFACL = wL
        End If
    End If
    rsSab.MoveNext

Loop
txtUpdate_TVACOMFACL = wTVACOMFACL
If Not blnOk Then
    Call MsgBox("aucun dossier YTVACOM0 :" & vbCrLf & xWhere, vbCritical, "TVA : annulation de commission")
    fraDétail_Display_TVACOMAVOIR = "?"
End If
If blnDoublon Then
    Call MsgBox("plusieurs commissions trouvées :" & vbCrLf & xWhere, vbInformation, "TVA : annulation de commission")
    fraDétail_Display_TVACOMAVOIR = "?"
End If
If blnOk Then
    fraDétail_Display_TVACOMAVOIR = " "
    If wTVACOMFACL = 0 Then
        Call MsgBox("1 commission trouvée non facturée", vbInformation, "TVA : annulation de commission")
    End If
End If
End Function

Public Function fraDétail_Update_Control_TVACOMFACL() As Boolean
Dim xWhere As String, xSql As String

newYTVACOM0.TVACOMFACL = Val(Trim(txtUpdate_TVACOMFACL))

If newYTVACOM0.TVACOMFACL <> 0 Then

    xWhere = "where TVAFACFACN = '" & newYTVACOM0.TVACOMFACL & "'" _
          & " and TVAFACCLIC = '" & newYTVACOM0.TVACOMCLIC & "' and TVAFACCLI = '" & newYTVACOM0.TVACOMCLI & "'"
    
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YTVAFAC0 " & xWhere
    Set rsSab = cnsab.Execute(xSql)
    If rsSab.EOF Then
        Call MsgBox("Cette facture ne concerne pas le client : " & newYTVACOM0.TVACOMCLI, vbCritical, "Contrôle de la facture " & newYTVACOM0.TVACOMFACL)
        fraDétail_Update_Control_TVACOMFACL = False
    Else
        fraDétail_Update_Control_TVACOMFACL = True
    End If
Else
    X = MsgBox("Confirmez-vous que cette annulation concerne une commission qui n'a pas fait l'objet d'une facture ?", vbQuestion + vbYesNo, "Validation d'une annulation de commission sans antécédent")
    If X <> vbYes Then
        fraDétail_Update_Control_TVACOMFACL = False
        txtUpdate_TVACOMFACL.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________N° facture liée")
    Else
        fraDétail_Update_Control_TVACOMFACL = True
    End If
End If

End Function

Public Sub cmdPrint_SQL_S()
Dim V, K As Long, X As String, curX As Currency

Dim wIndex As Integer
Dim mForeColor As Long

fgSelect.Visible = False
prtBIA_TVAFAC_Open 10, "Statistiques commissions, frais, intérêts ...."
For I = 1 To arrStat_K
    prtBIA_TVAFAC_NewLine 10
    If I > 0 Then
       XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
       XPrt.CurrentY = XPrt.CurrentY + 30
       If I = arrStat_K Then
           Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, "B", 240)
           XPrt.CurrentX = prtMinX + 50: XPrt.Print "TOTAL";
       Else
            XPrt.CurrentX = prtMinX + 50: XPrt.Print arrStat_Db1(I).TVACOMSER;
            XPrt.CurrentX = prtMinX + 550: XPrt.Print arrStat_Db1(I).TVACOMOPE;
            XPrt.CurrentX = prtMinX + 1050: XPrt.Print arrStat_Db1(I).TVACOMNAT;
       End If
       
    End If
        
    If arrStat_Db1(I).TVACOMUPDS > 0 Then
        XPrt.ForeColor = vbRed
        X = Format$(arrStat_Db1(I).TVACOMUPDS, "### ### ###")
        XPrt.CurrentX = prtMinX + 2500 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrStat_Db1(I).TVACOMMONE, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 4000 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If arrStat_Cr1(I).TVACOMUPDS > 0 Then
        XPrt.ForeColor = mForeColor
        X = Format$(arrStat_Cr1(I).TVACOMUPDS, "### ### ###")
        XPrt.CurrentX = prtMinX + 4800 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrStat_Cr1(I).TVACOMMONE, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 6300 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
        
    If arrStat_Db2(I).TVACOMUPDS > 0 Then
        XPrt.ForeColor = vbRed
        X = Format$(arrStat_Db2(I).TVACOMUPDS, "### ### ###")
        XPrt.CurrentX = prtMinX + 7100 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrStat_Db2(I).TVACOMMONE, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 8600 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If arrStat_Cr2(I).TVACOMUPDS > 0 Then
        XPrt.ForeColor = mForeColor
        X = Format$(arrStat_Cr2(I).TVACOMUPDS, "### ### ###")
        XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrStat_Cr2(I).TVACOMMONE, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 10900 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
        
    If arrStat_Db3(I).TVACOMUPDS > 0 Then
        XPrt.ForeColor = vbRed
        X = Format$(arrStat_Db3(I).TVACOMUPDS, "### ### ###")
        XPrt.CurrentX = prtMinX + 11700 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrStat_Db3(I).TVACOMMONE, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 13200 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If arrStat_Cr3(I).TVACOMUPDS > 0 Then
        XPrt.ForeColor = mForeColor
        X = Format$(arrStat_Cr3(I).TVACOMUPDS, "### ### ###")
        XPrt.CurrentX = prtMinX + 14000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrStat_Cr3(I).TVACOMMONE, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 15500 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
      
Next I
XPrt.DrawWidth = 5
prtBIA_TVAFAC_NewLine 10
prtBIA_TVAFAC_Form_10_Col

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
prtBIA_TVAFAC_Close 10

fgSelect.Visible = True

End Sub
