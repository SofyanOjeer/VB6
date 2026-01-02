VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSwift_Messages 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_Swift"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "Swift_Messages.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Suivi des Messages"
      TabPicture(0)   =   "Swift_Messages.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SWIFT ALLIANCE"
      TabPicture(1)   =   "Swift_Messages.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSaa_CB"
      Tab(1).Control(1)=   "fraSAA_Options"
      Tab(1).Control(2)=   "fraSAA"
      Tab(1).Control(3)=   "cmdSAA_Ok"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "SAB : Swifts à émettre / émis "
      TabPicture(2)   =   "Swift_Messages.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdZSWIFTA0_Update"
      Tab(2).Control(1)=   "cmdZSWIFTA0_Ok"
      Tab(2).Control(2)=   "fraZSWIFTA0_Options"
      Tab(2).Control(3)=   "fraZSWIFTA0"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Informatique : SAA instances"
      TabPicture(3)   =   "Swift_Messages.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraStatus"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Services"
      TabPicture(4)   =   "Swift_Messages.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraTab4"
      Tab(4).ControlCount=   1
      Begin VB.CommandButton cmdSaa_CB 
         BackColor       =   &H000000FF&
         Caption         =   "Export => CB"
         Height          =   525
         Left            =   -63240
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Frame fraTab4 
         Height          =   8055
         Left            =   -74880
         TabIndex        =   64
         Top             =   600
         Width           =   13095
         Begin VB.TextBox txtImport_File_Out 
            Height          =   375
            Left            =   5280
            TabIndex        =   68
            Text            =   "C:\temp\SAA\YYY.txt"
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtImport_File_In 
            Height          =   375
            Left            =   2280
            TabIndex        =   66
            Text            =   "C:\temp\SAA\XXX.txt"
            Top             =   480
            Width           =   2415
         End
         Begin VB.CommandButton cmdImport_File 
            Caption         =   "Lancer le traitement"
            Height          =   615
            Left            =   9240
            TabIndex        =   65
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label lblImport_File 
            Caption         =   "Fichiers à traiter In / Out"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.Frame fraStatus 
         Height          =   8205
         Left            =   -74880
         TabIndex        =   54
         Top             =   480
         Width           =   13290
         Begin VB.CommandButton cmdStatus_Update 
            BackColor       =   &H008080FF&
            Caption         =   "Mise à jour globale"
            Height          =   525
            Left            =   11895
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdStatus_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   405
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   720
            Width           =   1095
         End
         Begin VB.Frame fraStatus_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   55
            Top             =   120
            Width           =   11355
            Begin VB.TextBox txtStatus_Hms 
               Height          =   285
               Left            =   8880
               TabIndex        =   59
               Top             =   360
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker txtStatus_Amj 
               Height          =   300
               Left            =   7200
               TabIndex        =   56
               Top             =   360
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
               Format          =   100663299
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgStatus 
            Height          =   6825
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12039
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
            FormatString    =   ">AId       |>Umidh             |> umidl          |<appe date mod                         ||"
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
      Begin VB.CommandButton cmdZSWIFTA0_Update 
         BackColor       =   &H008080FF&
         Caption         =   "Emettre les SWIFT"
         Height          =   525
         Left            =   -63120
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame fraSAA_Options 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   31
         Top             =   480
         Width           =   11295
         Begin VB.TextBox txtSAA_CB 
            Height          =   375
            Left            =   1320
            TabIndex        =   70
            Text            =   "C:\temp\SAA\out_200501"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtSelect_Utilisateur 
            Height          =   285
            Left            =   7800
            TabIndex        =   63
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox cboSelect_Unit 
            Height          =   315
            Left            =   10200
            TabIndex        =   42
            Text            =   "cboSelect_SAA"
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cboSelect_IO 
            Height          =   315
            Left            =   10200
            TabIndex        =   41
            Text            =   "cboSelect_SAA"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtSelect_MT_type 
            Height          =   285
            Left            =   8760
            TabIndex        =   39
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox cboSelect_SAA 
            Height          =   315
            Left            =   1320
            TabIndex        =   32
            Text            =   "cboSelect_SAA"
            Top             =   240
            Width           =   2895
         End
         Begin MSComCtl2.DTPicker txtSelect_from_crea_date_time 
            Height          =   300
            Left            =   5160
            TabIndex        =   34
            Top             =   240
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
            Format          =   100663299
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtSelect_to_crea_date_time 
            Height          =   300
            Left            =   6840
            TabIndex        =   35
            Top             =   240
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
            Format          =   100663299
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblSelect_Utilisateur 
            Caption         =   "Utilisateur"
            Height          =   255
            Left            =   6840
            TabIndex        =   62
            Top             =   720
            Width           =   855
         End
         Begin VB.Label txtSelect_Unit 
            Caption         =   "Unit"
            Height          =   255
            Left            =   9600
            TabIndex        =   43
            Top             =   720
            Width           =   495
         End
         Begin VB.Label txtSelect_swift_IO 
            Caption         =   "Swift"
            Height          =   255
            Left            =   9600
            TabIndex        =   40
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblSelect_MT 
            Caption         =   "MT"
            Height          =   255
            Left            =   8400
            TabIndex        =   38
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblSelect_A 
            Caption         =   "A"
            Height          =   255
            Left            =   6600
            TabIndex        =   37
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblSelect_De 
            Caption         =   "De"
            Height          =   255
            Left            =   4800
            TabIndex        =   36
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblSelect_SAA 
            Caption         =   "Requête"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraSAA 
         Height          =   6615
         Left            =   -74760
         TabIndex        =   30
         Top             =   1800
         Width           =   13095
         Begin MSFlexGridLib.MSFlexGrid fgrTextField 
            Height          =   6135
            Left            =   6360
            TabIndex        =   44
            Top             =   120
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   10821
            _Version        =   393216
            Cols            =   4
            BackColorFixed  =   8438015
            BackColorBkg    =   12640511
            GridColor       =   255
            FormatString    =   $"Swift_Messages.frx":04CE
         End
         Begin MSFlexGridLib.MSFlexGrid fgSAA 
            Height          =   6525
            Left            =   120
            TabIndex        =   45
            Top             =   120
            Width           =   12960
            _ExtentX        =   22860
            _ExtentY        =   11509
            _Version        =   393216
            Rows            =   1
            Cols            =   19
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   15269886
            ForeColor       =   8388608
            BackColorFixed  =   12648447
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   15269886
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"Swift_Messages.frx":0581
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
      Begin VB.CommandButton cmdSAA_Ok 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Exécuter la requête"
         Height          =   645
         Left            =   -63240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   13290
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   11355
            Begin VB.CheckBox chkSelect_SWIMONSTA 
               Caption         =   "uniquement messages 'LIVE' , Sinon flux du : "
               Height          =   255
               Left            =   360
               TabIndex        =   61
               Top             =   360
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.CheckBox chkSelect_SAAAID 
               Caption         =   "uniquement messages à rapprocher"
               Height          =   255
               Left            =   8040
               TabIndex        =   53
               Top             =   600
               Width           =   3135
            End
            Begin VB.ComboBox cboSelect_SWIMONX32D 
               Height          =   315
               Left            =   8280
               Sorted          =   -1  'True
               TabIndex        =   24
               Text            =   "DEV"
               Top             =   240
               Width           =   1300
            End
            Begin MSComCtl2.DTPicker txtSelect_SWIMONFLUD 
               Height          =   300
               Left            =   4080
               TabIndex        =   25
               Top             =   360
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
               Format          =   100663299
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSelect_SWIMONFLUD_Max 
               Height          =   300
               Left            =   9840
               TabIndex        =   26
               Top             =   240
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
               Format          =   100663299
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_SWIMONX32D 
               Caption         =   "Devise"
               Height          =   255
               Left            =   7320
               TabIndex        =   27
               Top             =   240
               Width           =   600
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   14
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
            FormatString    =   $"Swift_Messages.frx":06B7
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
      Begin VB.CommandButton cmdZSWIFTA0_Ok 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Exécuter la requête"
         Height          =   525
         Left            =   -63120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.Frame fraZSWIFTA0_Options 
         Height          =   1095
         Left            =   -74640
         TabIndex        =   6
         Top             =   360
         Width           =   11295
         Begin VB.ComboBox cboSelect_SWIFTASER 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   52
            Top             =   720
            Width           =   700
         End
         Begin VB.ComboBox cboSelect_SWIFTADE1 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   12
            Text            =   "DEV"
            Top             =   720
            Width           =   1300
         End
         Begin VB.TextBox txtSelect_SWIFTAREF 
            Height          =   285
            Left            =   1080
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtSelect_SWIFTAMES 
            Height          =   285
            Left            =   5040
            TabIndex        =   10
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSelect_SWIFTADES 
            Height          =   285
            Left            =   4320
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboSelect_SWIFTAXXX 
            Height          =   315
            Left            =   7440
            TabIndex        =   8
            Text            =   "SWIFTAXXX"
            Top             =   240
            Width           =   2895
         End
         Begin VB.CheckBox chkSelect_ZSWIHIA0 
            Alignment       =   1  'Right Justify
            Caption         =   "Rechercher historique"
            Height          =   255
            Left            =   8760
            TabIndex        =   7
            Top             =   720
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker txtSelect_SWIFTADVA 
            Height          =   300
            Left            =   7440
            TabIndex        =   13
            Top             =   720
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
            Format          =   100663299
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblSelect_SWIFTASER 
            Caption         =   "Service"
            Height          =   255
            Left            =   2880
            TabIndex        =   51
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblSelect_SWIFTADE1 
            Caption         =   "Devise"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   600
         End
         Begin VB.Label lblSelect_SWIFTAREF 
            Caption         =   "Référence"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblSelect_SWIFTAMES 
            Caption         =   "MT"
            Height          =   255
            Left            =   4560
            TabIndex        =   17
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblSelect_SWIFTADES 
            Caption         =   "BIC Destinataire"
            Height          =   255
            Left            =   2880
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblSelect_SWIFTAXXX 
            Caption         =   "Indicateurs d'état"
            Height          =   255
            Left            =   5880
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblSelect_SWIFTADVA 
            Caption         =   "Date valeur <="
            Height          =   255
            Left            =   5880
            TabIndex        =   14
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame fraZSWIFTA0 
         Height          =   6945
         Left            =   -74640
         TabIndex        =   5
         Top             =   1560
         Width           =   13170
         Begin VB.ListBox lstW 
            Height          =   4545
            Left            =   5400
            TabIndex        =   49
            Top             =   1560
            Visible         =   0   'False
            Width           =   7335
         End
         Begin MSFlexGridLib.MSFlexGrid fgZSWI_D 
            Height          =   6015
            Left            =   5400
            TabIndex        =   46
            Top             =   600
            Visible         =   0   'False
            Width           =   7400
            _ExtentX        =   13044
            _ExtentY        =   10610
            _Version        =   393216
            Cols            =   4
            BackColorFixed  =   8438015
            BackColorBkg    =   12640511
            GridColor       =   255
            GridLines       =   3
            AllowUserResizing=   3
            FormatString    =   $"Swift_Messages.frx":075E
         End
         Begin MSFlexGridLib.MSFlexGrid fgZSWIFTA0 
            Height          =   3645
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   12960
            _ExtentX        =   22860
            _ExtentY        =   6429
            _Version        =   393216
            Rows            =   1
            Cols            =   14
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   15269886
            ForeColor       =   8388608
            BackColorFixed  =   12648447
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   15269886
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"Swift_Messages.frx":0846
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
         Begin MSFlexGridLib.MSFlexGrid fgZSWIHIA0 
            Height          =   2925
            Left            =   120
            TabIndex        =   48
            Top             =   3960
            Width           =   12960
            _ExtentX        =   22860
            _ExtentY        =   5159
            _Version        =   393216
            Rows            =   1
            Cols            =   14
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   15268863
            ForeColor       =   8388608
            BackColorFixed  =   12640511
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14153215
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"Swift_Messages.frx":0911
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
      Left            =   13080
      Picture         =   "Swift_Messages.frx":099B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Label libSelect 
      BackColor       =   &H00FFFED9&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4905
      WordWrap        =   -1  'True
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
      Begin VB.Menu mnux1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZSWIALI0_Update 
         Caption         =   "Emission ZSWIALI0 => SAA"
      End
      Begin VB.Menu mnux1b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZSWIFTA0_Update_BackOffice 
         Caption         =   "Emission SWIFT : Back Office"
      End
      Begin VB.Menu mnuZSWIFTA0_Update_BOTC_Jour 
         Caption         =   "Emission SWIFT : BOTC valeur jour"
      End
      Begin VB.Menu mnuZSWIFTA0_Update_BOTC_MT3 
         Caption         =   "Emission SWIFT : BOTC MT3**"
      End
      Begin VB.Menu mnux2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZSWIFTA0_Update_Manuel_48H 
         Caption         =   "Emission SWIFT : manuel <= 48 heures"
      End
      Begin VB.Menu mnuZSWIFTA0_Update_Manuel 
         Caption         =   "Emission SWIFT : manuel <= 31.12.2999"
      End
      Begin VB.Menu mnuX3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZSWIHIA0_Display 
         Caption         =   "Affichage Historique  JJ.MM.AAA"
      End
      Begin VB.Menu mnuZSWIHIA0_Reprise_YSWIMON0 
         Caption         =   "$$$ Reprise Automatique YSWIMON0 $$$$"
      End
      Begin VB.Menu mnuX4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAuto_Status 
         Caption         =   "mnuAuto_Status"
      End
      Begin VB.Menu mnuAuto_Status_Complément 
         Caption         =   "mnuAuto_Status_Complément"
      End
      Begin VB.Menu mnuAuto_Status_S200 
         Caption         =   "mnuAuto_Status_S200"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer liste"
      End
      Begin VB.Menu mnuSelect_Print_Détail 
         Caption         =   "Imprimer liste détaillée"
      End
   End
   Begin VB.Menu mnuSAA_Queue 
      Caption         =   "mnuSAA_Queue"
      Visible         =   0   'False
      Begin VB.Menu mnuSAA_Queue_Modification 
         Caption         =   "SAA_Modification"
      End
      Begin VB.Menu mnuSAA_Queue_Autorisation 
         Caption         =   "SAA_Autorisation"
      End
      Begin VB.Menu mnuSAA_Queue_SWIFT 
         Caption         =   "SAA_SWIFT"
      End
   End
   Begin VB.Menu mnuReprise 
      Caption         =   "mnuReprise"
      Visible         =   0   'False
      Begin VB.Menu mnuReprise_Restauration 
         Caption         =   "Restaurer dans ZSWIFTA0 ....confirmation ?"
      End
      Begin VB.Menu mnuReprise_H 
         Caption         =   "Ajouter dans YSWIMON0"
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_S999 
         Caption         =   "Annulation : S999"
      End
   End
End
Attribute VB_Name = "frmSwift_Messages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim SWI_MESSAGES_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean, blnDebug As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim blnTransaction As Boolean

Dim fgZSWIFTA0_FormatString As String, fgZSWIFTA0_K As Integer
Dim fgZSWIFTA0_RowDisplay As Integer, fgZSWIFTA0_RowClick As Integer, fgZSWIFTA0_ColClick As Integer
Dim fgZSWIFTA0_ColorClick As Long, fgZSWIFTA0_ColorDisplay As Long
Dim fgZSWIFTA0_Sort1 As Integer, fgZSWIFTA0_Sort2 As Integer
Dim fgZSWIFTA0_SortAD As Integer, fgZSWIFTA0_Sort1_Old As Integer
Dim fgZSWIFTA0_arrIndex As Integer
Dim blnfgZSWIFTA0_DisplayLine As Boolean
Dim fgZSWIFTA0_Height As Integer

Dim meZSWIFTA0 As typeZSWIFTA0, xZSWIFTA0 As typeZSWIFTA0
Dim arrZSWIFTA0() As typeZSWIFTA0, arrZSWIFTA0_Nb As Long, arrZSWIFTA0_Max As Long, arrZSWIFTA0_Index As Long
Dim arrZSWIFTB0() As typeZSWIFTB0, arrZSWIFTB0_Nb As Long, arrZSWIFTB0_Max As Long
Dim arrSWIFTCSIG() As String
Dim arrZSWIFTC0() As typeZSWIFTC0, arrZSWIFTC0_Nb As Long, arrZSWIFTC0_Max As Long
Dim arrZSWITEM0() As typeZSWITEM0, arrZSWITEM0_Nb As Long, arrZSWITEM0_Max As Long
Dim arrZSWIFTA0_SAA_Queue() As String

Dim meZSWIHIA0 As typeZSWIHIA0, xZSWIHIA0 As typeZSWIHIA0
Dim arrZSWIHIA0() As typeZSWIHIA0, arrZSWIHIA0_Nb As Long, arrZSWIHIA0_Max As Long
Dim arrZSWIHIB0() As typeZSWIHIB0, arrZSWIHIB0_Nb As Long, arrZSWIHIB0_Max As Long
Dim arrSWIHICSIG() As String
Dim arrZSWIHIC0() As typeZSWIHIC0, arrZSWIHIC0_Nb As Long, arrZSWIHIC0_Max As Long
Dim arrZSWIHIT0() As typeZSWIHIT0, arrZSWIHIT0_Nb As Long, arrZSWIHIT0_Max As Long

Dim fgZSWIHIA0_FormatString As String, fgZSWIHIA0_K As Integer
Dim fgZSWIHIA0_RowDisplay As Integer, fgZSWIHIA0_RowClick As Integer, fgZSWIHIA0_ColClick As Integer
Dim fgZSWIHIA0_ColorClick As Long, fgZSWIHIA0_ColorDisplay As Long
Dim fgZSWIHIA0_Sort1 As Integer, fgZSWIHIA0_Sort2 As Integer
Dim fgZSWIHIA0_SortAD As Integer, fgZSWIHIA0_Sort1_Old As Integer
Dim fgZSWIHIA0_arrIndex As Integer
Dim blnfgZSWIHIA0_DisplayLine As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xYSWIMON0 As typeYSWIMON0, newYSWIMON0 As typeYSWIMON0, oldYSWIMON0 As typeYSWIMON0
Dim arrYSWIMON0() As typeYSWIMON0, arrYSWIMON0_Nb As Long, arrYSWIMON0_Max As Long, arrYSWIMON0_Index As Long

'______________________________________________________________________

Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset

Dim merAppe As typerAppe, xrAppe As typerAppe, xrAppe_E As typerAppe, xrAppe_R As typerAppe
Dim merIntv As typerIntv, xrIntv As typerIntv
Dim merInst As typerInst, xrInst As typerInst
Dim merJrnl As typerJrnl, xrJrnl As typerJrnl
Dim merMesg As typerMesg, xrMesg As typerMesg
Dim merTextField As typerTextField, xrTextField As typerTextField

Dim fgSAA_FormatString As String, fgSAA_K As Integer
Dim fgSAA_RowDisplay As Integer, fgSAA_RowClick As Integer, fgSAA_ColClick As Integer
Dim fgSAA_ColorClick As Long, fgSAA_ColorDisplay As Long
Dim fgSAA_Sort1 As Integer, fgSAA_Sort2 As Integer
Dim fgSAA_SortAD As Integer, fgSAA_Sort1_Old As Integer
Dim fgSAA_arrIndex As Integer
Dim blnfgSAA_DisplayLine As Boolean

Dim arrrAppe() As typerAppe, arrrAppe_Nb As Long, arrrAppe_Max As Long
Dim arrrAppe_E() As typerAppe
Dim arrrAppe_R() As typerAppe
Dim arrrInst() As typerInst, arrrInst_Nb As Long, arrrInst_Max As Long
Dim arrrIntv() As typerIntv, arrrIntv_Nb As Long, arrrIntv_Max As Long
Dim arrrJrnl() As typerJrnl, arrrJrnl_Nb As Long, arrrJrnl_Max As Long
Dim arrrMesg() As typerMesg, arrrMesg_Nb As Long, arrrMesg_Max As Long
Dim arrrTextField() As typerTextField, arrrTextField_Nb As Long, arrrTextField_Max As Long

Dim fgrTextField_FormatString As String, fgrTextField_K As Integer
Dim fgrTextField_RowDisplay As Integer, fgrTextField_RowClick As Integer, fgrTextField_ColClick As Integer
Dim fgrTextField_ColorClick As Long, fgrTextField_ColorDisplay As Long
Dim fgrTextField_Sort1 As Integer, fgrTextField_Sort2 As Integer
Dim fgrTextField_SortAD As Integer, fgrTextField_Sort1_Old As Integer
Dim fgrTextField_arrIndex As Integer
Dim blnfgrTextField_DisplayLine As Boolean

Dim fgStatus_FormatString As String, fgStatus_K As Integer
Dim fgStatus_RowDisplay As Integer, fgStatus_RowClick As Integer, fgStatus_ColClick As Integer
Dim fgStatus_ColorClick As Long, fgStatus_ColorDisplay As Long
Dim fgStatus_Sort1 As Integer, fgStatus_Sort2 As Integer
Dim fgStatus_SortAD As Integer, fgStatus_Sort1_Old As Integer
Dim fgStatus_arrIndex As Integer
Dim blnfgStatus_DisplayLine As Boolean
Dim arrrInst_Status() As typerInst, arrrInst_Status_Nb As Long, arrrInst_Status_Max As Long, arrrInst_Status_Index As Long
Dim arrrIntv_Status() As typerIntv, arrrIntv_Status_Nb As Long, arrrIntv_Status_Max As Long, arrrIntv_Status_Index As Long
Dim arrrIntv_Flag() As typerIntv, arrrIntv_Flag_Nb As Long, arrrIntv_Flag_Max As Long, arrrIntv_Flag_Index As Long
Dim merIntv_Flag As typerIntv
Dim res_date_time As Date
Dim merInst_Status As typerInst, merMesg_Status As typerMesg
Dim merAppe_Status As typerAppe, merIntv_Status As typerIntv, resIntv_Status As typerIntv

Dim meYSWIMON0_Status As typeYSWIMON0, oldYSWIMON0_Status As typeYSWIMON0

Dim rsSab_Update As New ADODB.Recordset


'______________________________________________________________________

Dim meZSWIALI0 As typeZSWIALI0, xZSWIALI0 As typeZSWIALI0
Dim arrZSWIALI0() As typeZSWIALI0, arrZSWIALI0_Nb As Long, arrZSWIALI0_Max As Long, arrZSWIALI0_Index As Long

Dim wSWISABCOP As String, wSWISABDOS As Long, wSWISABUnit As String

Dim arrMt_Unit(1000, 11, 2) As Integer


Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim ts_DSys_7 As String
Private Sub cmdPrint_Click_xlsManual()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 1:
            Select Case Mid$(cboSelect_SAA, 1, 2)
                Case "6 ": Call cmdPrint_List6_Ok_xlsManual
                Case "7 ": Call cmdPrint_List6_Ok_xlsManual
                Case "8 ": cmdPrint_List8_Ok
                
               ' Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
            End Select
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdPrint_List6_Ok_xlsManual()
Dim iRow As Integer, K As Integer, I As Integer, X As String
Dim blnOk As Boolean, blnOpen As Boolean
Dim xCOMPTEINT As String
Dim m_unit_name As String
Dim mNrequest As String
Dim wbExcel As Excel.Workbook
Dim currentrow As Long
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long
Dim nbSheetRows As Long

mNrequest = Mid$(cboSelect_SAA, 1, 2)
fgSAA.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSAA.Rows - 1)
If fgSAA.Rows > 1 Then
    fgSAA_Sort1_Old = -1
    fgSAA_SortX 13
    '                                               '
    Call init_xlsManual
    'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
    FileCopy paramFolder_Local & "\Modeles\modele_SWI_MESSAGES.xlsx", paramIMP_PDF_Path_Temp & "\modele_SWI_MESSAGES.xlsx"
    'on charge CE classeur dans Excel
    Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_SWI_MESSAGES.xlsx")
    Set wbExcel = appExcelPublic.ActiveWorkbook
    With wbExcel
        .Title = "JRN_MNU"
        .Subject = "JRN_MNU"
    End With
    wbExcel.Worksheets(1).Activate
    '                                               '
    currentrow = 7
    comptageRows = currentrow
    maxRows = 38
    maxRowsPlus = 4
End If
blnOpen = False
m_unit_name = ""
For iRow = 1 To fgSAA.Rows - 1
    fgSAA.Row = iRow
    fgSAA.Col = fgSAA_arrIndex:  K = CLng(fgSAA.Text)
    xrMesg = arrrMesg(K)
    xrAppe_E = arrrAppe_E(K)
    xrAppe_R = arrrAppe_R(K)
    xrInst = arrrInst(K)
' Rupture Service : ligne totale de l'ancien code de service
    If m_unit_name <> xrMesg.x_inst0_unit_name Then
        If blnOpen Then
            Call prtSWI_Messages_List6_Close_xlsManual(m_unit_name, mNrequest, currentrow, wbExcel.Sheets(1), comptageRows, maxRows, maxRowsPlus)
            If blnAuto Then
                Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "")
            End If
        End If
        m_unit_name = xrMesg.x_inst0_unit_name
        X = Table_Unit_SSI("", m_unit_name)
        If blnAuto Then
            Call frmElpPrt.prtIMP_PDF_NoPaper_Init(X, "BIA-SAA-CTL", "Archive")
        End If
        Call prtSWI_Messages_List6_Open_xlsManual(Mid$(cboSelect_SAA, 1, 2), txtSelect_from_crea_date_time, txtSelect_to_crea_date_time, wbExcel.Sheets(1))
        blnOpen = True
   End If
    Call prtSWI_Messages_List6_Line_xlsManual(Mid$(cboSelect_SAA, 1, 2), xrMesg, xrAppe_E, xrAppe_R, xrInst, currentrow, wbExcel.Sheets(1), comptageRows, maxRows, maxRowsPlus)
Next iRow
If blnOpen Then
    Call prtSWI_Messages_List6_Close_xlsManual(m_unit_name, mNrequest, currentrow, wbExcel.Sheets(1), comptageRows, maxRows, maxRowsPlus)
    wbExcel.Worksheets(1).Cells(currentrow + 1, 1) = "END_OF_SHEET"
    'on supprime les 4 lignes du modèle
    Rows("4:7").Select
    Selection.Delete
    nbSheetRows = retourne_fin_de_sheet(wbExcel.Worksheets(1))
    Call zoneImpression_xlsManual(wbExcel.Sheets(1).Name, nbSheetRows, wbExcel.Sheets(1))
    Call ActiveSheet.ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path_Temp & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
    Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
    Call wbExcel.Close(True)
    Set wbExcel = Nothing
    If Dir(paramIMP_PDF_Path_Temp & "\modele_SWI_MESSAGES.xlsx") <> "" Then
        Kill paramIMP_PDF_Path_Temp & "\modele_SWI_MESSAGES.xlsx"
    End If
   '' prtSWI_Messages_List6_Rupture m_unit_name, mNrequest
    If blnAuto Then
        Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "")
    End If
End If
fgSAA.Visible = True
Me.Show
End Sub

Public Sub zoneImpression_xlsManual(lFct As String, nbRows As Long, wsheet As Excel.Worksheet)

    Call init_TypePagesetup
    If nbRows > 0 Then
        wsheet.Activate
        zoneImpressionPagesetup.Zoom = 75
        wsheet.Range("A1:J" & CStr(nbRows)).Select
        zoneImpressionPagesetup.PrintArea = "$A$1:$J$" & CStr(nbRows)
        zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSWI_Messages   &D &T  BIA_INFO"
    End If
    wsheet.Activate
    zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
    zoneImpressionPagesetup.Orientation = xlLandscape
    Call SetTypePageSetup(wsheet)
    
End Sub

Public Sub SAA_Statistiques_Export_Recap()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim mBackColor As Long, mForeColor As Long
'______________________________________________

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
End With

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 10
For wCol = 2 To fgSAA.Cols - 1
    wsExcel.Columns(wCol).ColumnWidth = 10: wsExcel.Columns(wCol).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
Next wCol
fgSAA.Row = 0: fgSAA.Col = 0
mBackColor = colorHex_RGB(fgSAA.BackColorFixed)
'mForeColor = colorHex_RGB(fgSAA.ForeColorFixed)
For wRow = 0 To fgSAA.Rows - 1
    fgSAA.Row = wRow
    For wCol = 0 To fgSAA.Cols - 1
        fgSAA.Col = wCol
        If wCol < 1 Or wRow < fgSAA.FixedRows Then
            wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = mBackColor
            wsExcel.Cells(wRow + 1, wCol + 1).Font.Color = mForeColor
            wsExcel.Cells(wRow + 1, wCol + 1) = fgSAA.Text

       Else
            wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSAA.BackColor)
            wsExcel.Cells(wRow + 1, wCol + 1).Font.Color = colorHex_RGB(fgSAA.CellForeColor)
            wsExcel.Cells(wRow + 1, wCol + 1) = Val(fgSAA.Text)
        End If
    Next wCol
Next wRow

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub
Public Sub SAA_Statistiques_Export()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFilex As String, wFile As String, xSQL As String
Dim wAmjMin As String, wAmjMax As String
Dim X As String, K As Long
'______________________________________________

cmdSAA_SQL_08

Call DTPicker_Control(txtSelect_from_crea_date_time, wAmjMin)
Call DTPicker_Control(txtSelect_to_crea_date_time, wAmjMax)

wFile = Trim("C:\Temp\Swift Statistiques " & wAmjMin & "-" & wAmjMax & ".xlsx")
'______________________________________________
X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "Swift Statistiques: nom du fichier d'exportation", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
End If
'_________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SWIFT stat"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet

wsExcel.Name = wAmjMin & "-" & wAmjMax

SAA_Statistiques_Export_Recap

'__________________________________________________________________________________
'__________________________________________________________________________________
Exit_sub:
'__________________________________________________________________________________
Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub


Public Sub fgrTextField_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgrTextField.Row

If lRow > 0 And lRow < fgrTextField.Rows Then
    fgrTextField.Row = lRow
    For I = 0 To fgrTextField_arrIndex
        fgrTextField.Col = I: fgrTextField.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgrTextField.Row = mRow
    If fgrTextField.Row > 0 Then
        lRow = fgrTextField.Row
        lColor_Old = fgrTextField.CellBackColor
        For I = 0 To fgrTextField_arrIndex
          fgrTextField.Col = I: fgrTextField.CellBackColor = lColor
        Next I
        fgrTextField.Col = 0
    End If
End If

End Sub


Public Sub fgSAA_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSAA.Row

If lRow > 0 And lRow < fgSAA.Rows Then
    fgSAA.Row = lRow
    For I = 0 To fgSAA_arrIndex
        fgSAA.Col = I: fgSAA.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSAA.Row = mRow
    If fgSAA.Row > 0 Then
        lRow = fgSAA.Row
        lColor_Old = fgSAA.CellBackColor
        For I = 0 To fgSAA_arrIndex
          fgSAA.Col = I: fgSAA.CellBackColor = lColor
        Next I
        fgSAA.Col = 0
    End If
End If

End Sub

Private Sub fgSAA_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgSAA.Visible = False
fgSAA_Reset
cmdPrint.Enabled = False

txtSelect_Utilisateur = ""
lblSelect_Utilisateur.Visible = False
txtSelect_Utilisateur.Visible = False

fgSAA.Rows = 1
fgSAA.FormatString = fgSAA_FormatString
currentAction = "fgSAA_Display"
    
For I = 1 To arrrIntv_Nb
         
    xrIntv = arrrIntv(I)
    xrMesg = arrrMesg(I)
    xrInst = arrrInst(I)
    xrAppe = arrrAppe(I)
    xrAppe_E = arrrAppe_E(I)
    xrAppe_R = arrrAppe_R(I)
    
    fgSAA_DisplayLine (I)
Next I

fgSAA.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & fgSAA.Rows - 1): DoEvents
If fgSAA.Rows > 1 Then
    fgSAA_Sort1 = 0: fgSAA_Sort2 = 1: fgSAA_Sort
    cmdPrint.Enabled = True
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
fgSAA.Visible = True

End Sub


Private Sub fgSAA_Display_99()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgSAA.Visible = False
fgSAA_Reset
cmdPrint.Enabled = False

lblSelect_Utilisateur.Visible = True
txtSelect_Utilisateur.Visible = True

fgSAA.Rows = 1
fgSAA.FormatString = "<Date interventions |<Interventions                        | I/O |Correspondant| MT |< Notre Référence|Date valeur|> Montant           |Format/Status| ACK/NAK | File d'attente  |<Créé le...         | Expéditeur / Destinataire |Service| Créé par... | Validé par...   |||"
currentAction = "fgSAA_Display_99"
    
For I = 1 To arrrIntv_Nb
         
    xrIntv = arrrIntv(I)
    xrMesg = arrrMesg(I)
    xrInst = arrrInst(I)
    xrAppe = arrrAppe(I)
    xrAppe_E = arrrAppe_E(I)
    xrAppe_R = arrrAppe_R(I)
    
    fgSAA_DisplayLine_99 (I)
    
Next I

fgSAA.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & fgSAA.Rows - 1): DoEvents
'If fgSAA.Rows > 1 Then
'    fgSAA_Sort1 = 0: fgSAA_Sort2 = 1: fgSAA_Sort
'    cmdPrint.Enabled = True
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
fgSAA.Visible = True

End Sub

Private Sub fgSAA_Display_08()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim K1 As Integer, K2 As Integer, K3 As Integer, xUnit As String

On Error GoTo Error_Handler
fgSAA.Visible = False
fgSAA_Reset

lblSelect_Utilisateur.Visible = True
txtSelect_Utilisateur.Visible = True

fgSAA.Rows = 1
fgSAA.FormatString = "<Message   |   BOTC   |   CSOP   |   DAFI   |  DCOM    |   ORPA   |   SCLE   |  SOBF    |    SOBI  |  AUTRES  |   NONE   |"
currentAction = "fgSAA_Display_99"
    
For K1 = 0 To 1000
    If arrMt_Unit(K1, 0, 0) <> 0 Then
        fgSAA.Rows = fgSAA.Rows + 1
        fgSAA.Row = fgSAA.Rows - 1
        fgSAA.Col = 0: fgSAA.Text = Format(K1, "000") & " Emis"
        fgSAA.CellForeColor = vbRed
        For K2 = 1 To 10
            fgSAA.Col = K2: fgSAA.Text = Format(arrMt_Unit(K1, K2, 1), "#####")
            fgSAA.CellForeColor = vbRed
        Next K2
        fgSAA.Rows = fgSAA.Rows + 1
        fgSAA.Row = fgSAA.Rows - 1
        fgSAA.CellForeColor = vbBlue
        fgSAA.Col = 0: fgSAA.Text = Format(K1, "000") & " Reçus"
        For K2 = 1 To 10
            fgSAA.Col = K2: fgSAA.Text = Format(arrMt_Unit(K1, K2, 2), "#####")
            fgSAA.CellForeColor = vbBlue
        Next K2
    End If
    
Next K1

fgSAA.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & fgSAA.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
fgSAA.Visible = True

End Sub

Public Sub fgrTextField_DisplayLine(lIndexe As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
On Error Resume Next

 fgrTextField.Rows = fgrTextField.Rows + 1
 fgrTextField.Row = fgrTextField.Rows - 1
 
 fgrTextField.Col = 0: fgrTextField.Text = xrTextField.field_code
 fgrTextField.Col = 1: fgrTextField.Text = xrTextField.field_option
 iLen = Len(xrTextField.value)
 K = 1
 Do
    iAsc13 = InStr(K, xrTextField.value, Asc13)
    If iAsc13 > 0 Then
        fgrTextField.Col = 2: fgrTextField.Text = Mid$(xrTextField.value, K, iAsc13 - K)
        K = iAsc13 + 2
        fgrTextField.Rows = fgrTextField.Rows + 1
        fgrTextField.Row = fgrTextField.Rows - 1
    End If
 Loop Until iAsc13 = 0
      
 fgrTextField.Col = 2: fgrTextField.Text = Mid$(xrTextField.value, K, iLen - K + 1)

 ' fgrTextField.Col = fgrTextField_arrIndex: fgrTextField.Text = lIndexe

End Sub

Public Sub fgSAA_DisplayLine(lIndex As Long)
Dim K As Integer, xDev As String, X As String, xCur As Currency

On Error Resume Next
' Tout le paragraphe d'affichage ne s'exécute que si le format de message est SWIFT

If Mid$(cboSelect_IO, 1, 1) <> " " And Mid$(xrMesg.mesg_sub_format, 1, 1) <> Mid$(cboSelect_IO, 1, 1) Then Exit Sub

If Trim(txtSelect_MT_type) <> "" And xrMesg.mesg_type <> Trim(txtSelect_MT_type) Then Exit Sub

If Mid$(cboSelect_Unit, 1, 1) <> " " And Mid$(xrMesg.x_inst0_unit_name, 1, 4) <> Mid$(cboSelect_Unit, 1, 4) Then Exit Sub

If xrMesg.mesg_frmt_name = "Swift" Then

    If Mid$(cboSelect_SAA, 1, 2) <> "7 " Or (Mid$(cboSelect_SAA, 1, 2) = "7 " And xrMesg.mesg_is_text_modified = 1) Then

        fgSAA.Rows = fgSAA.Rows + 1
        fgSAA.Row = fgSAA.Rows - 1
        
        fgSAA.Col = 0: fgSAA.Text = xrMesg.mesg_sub_format
        
        If xrMesg.mesg_sub_format Like "INPUT%" Then
            fgSAA.Col = 1: fgSAA.Text = xrMesg.mesg_receiver_swift_address
        Else
            fgSAA.Col = 1: fgSAA.Text = xrMesg.mesg_sender_swift_address
        End If
        
        fgSAA.Col = 2: fgSAA.Text = xrMesg.mesg_type
        fgSAA.Col = 3: fgSAA.Text = xrMesg.mesg_trn_ref
        fgSAA.Col = 4: fgSAA.Text = xrMesg.mesg_fin_value_date
        ''fgSAA.Col = 5: fgSAA.Text = xrMesg.mesg_fin_ccy_amount
        K = 0
        X = xrMesg.mesg_fin_ccy_amount
        xDev = Space_Scan(X, K)
        xCur = num_CDec_USA(Space_Scan(X, K))
        If xCur <> 0 Then fgSAA.Col = 5: fgSAA.Text = Format$(xCur, "### ### ### ##0.00") & " " & xDev
    
        fgSAA.Col = 6: fgSAA.Text = xrMesg.mesg_frmt_name & " " & xrMesg.mesg_status
        
        fgSAA.Col = 7: fgSAA.Text = xrAppe_E.appe_network_delivery_status
        If fgSAA.Text = "" Then
            fgSAA.Col = 7: fgSAA.Text = xrAppe_R.appe_network_delivery_status
        End If
        
        fgSAA.Col = 8: fgSAA.Text = xrInst.inst_rp_name
        fgSAA.Col = 9: fgSAA.Text = xrMesg.mesg_crea_date_time
        
        ' -Input- commence par "APPE_RECEPTION%" puis "APPE_EMISSION%"
        fgSAA.Col = 10: fgSAA.Text = xrAppe_E.appe_iapp_name & " " & xrAppe_E.appe_session_nbr & " " & xrAppe_E.appe_sequence_nbr
        fgSAA.Col = 11: fgSAA.Text = xrAppe_R.appe_iapp_name & " " & xrAppe_R.appe_session_nbr & " " & xrAppe_R.appe_sequence_nbr
        
        fgSAA.Col = 12: fgSAA.Text = xrMesg.mesg_sender_swift_address & " " & xrMesg.mesg_receiver_swift_address
        fgSAA.Col = 13: fgSAA.Text = xrMesg.x_inst0_unit_name
        fgSAA.Col = 14: fgSAA.Text = xrMesg.mesg_crea_oper_nickname
        fgSAA.Col = 15: fgSAA.Text = xrInst.inst_auth_oper_nickname
        
        fgSAA.Col = fgSAA_arrIndex: fgSAA.Text = lIndex
        
    End If

End If

End Sub


Public Sub fgSAA_DisplayLine_99(lIndex As Long)
Dim K As Integer, xDev As String, X As String, xCur As Currency

On Error Resume Next
' Tout le paragraphe d'affichage ne s'exécute que si le format de message est SWIFT

If Mid$(cboSelect_IO, 1, 1) <> " " And Mid$(xrMesg.mesg_sub_format, 1, 1) <> Mid$(cboSelect_IO, 1, 1) Then Exit Sub

If Trim(txtSelect_MT_type) <> "" And xrMesg.mesg_type <> Trim(txtSelect_MT_type) Then Exit Sub

If Mid$(cboSelect_Unit, 1, 1) <> " " And Mid$(xrMesg.x_inst0_unit_name, 1, 4) <> Mid$(cboSelect_Unit, 1, 4) Then Exit Sub

If xrMesg.mesg_frmt_name = "Swift" Then

    fgSAA.Rows = fgSAA.Rows + 1
    fgSAA.Row = fgSAA.Rows - 1
    
    fgSAA.Col = 0: fgSAA.Text = xrIntv.intv_date_time
    fgSAA.Col = 1: fgSAA.Text = xrIntv.intv_text

    fgSAA.Col = 2: fgSAA.Text = xrMesg.mesg_sub_format
    
    If xrMesg.mesg_sub_format Like "INPUT%" Then
        fgSAA.Col = 3: fgSAA.Text = xrMesg.mesg_receiver_swift_address
    Else
        fgSAA.Col = 3: fgSAA.Text = xrMesg.mesg_sender_swift_address
    End If
    
    fgSAA.Col = 4: fgSAA.Text = xrMesg.mesg_type
    fgSAA.Col = 5: fgSAA.Text = xrMesg.mesg_trn_ref
    fgSAA.Col = 6: fgSAA.Text = xrMesg.mesg_fin_value_date
    
    K = 0
    X = xrMesg.mesg_fin_ccy_amount
    xDev = Space_Scan(X, K)
    xCur = num_CDec_USA(Space_Scan(X, K))
    If xCur <> 0 Then fgSAA.Col = 7: fgSAA.Text = Format$(xCur, "### ### ### ##0.00") & " " & xDev

    fgSAA.Col = 8: fgSAA.Text = xrMesg.mesg_frmt_name & " " & xrMesg.mesg_status
    
    fgSAA.Col = 9: fgSAA.Text = xrAppe_E.appe_network_delivery_status
    If fgSAA.Text = "" Then
        fgSAA.Col = 9: fgSAA.Text = xrAppe_R.appe_network_delivery_status
    End If
    
    fgSAA.Col = 10: fgSAA.Text = xrInst.inst_rp_name
    fgSAA.Col = 11: fgSAA.Text = xrMesg.mesg_crea_date_time
    
    ' -Input- commence par "APPE_RECEPTION%" puis "APPE_EMISSION%"
    'fgSAA.Col = 10: fgSAA.Text = xrAppe_E.appe_iapp_name & " " & xrAppe_E.appe_session_nbr & " " & xrAppe_E.appe_sequence_nbr
    'fgSAA.Col = 11: fgSAA.Text = xrAppe_R.appe_iapp_name & " " & xrAppe_R.appe_session_nbr & " " & xrAppe_R.appe_sequence_nbr
    
    fgSAA.Col = 12: fgSAA.Text = xrMesg.mesg_sender_swift_address & " " & xrMesg.mesg_receiver_swift_address
    fgSAA.Col = 13: fgSAA.Text = xrMesg.x_inst0_unit_name
    fgSAA.Col = 14: fgSAA.Text = xrMesg.mesg_crea_oper_nickname
    fgSAA.Col = 15: fgSAA.Text = xrMesg.mesg_verf_oper_nickname
    
    fgSAA.Col = fgSAA_arrIndex: fgSAA.Text = lIndex

End If

End Sub


Public Sub fgSAA_Reset()
fgSAA.Clear
fgSAA_Sort1 = 0: fgSAA_Sort2 = 0
fgSAA_Sort1_Old = -1
fgSAA_RowDisplay = 0: fgSAA_RowClick = 0
fgSAA_arrIndex = fgSAA.Cols - 1
blnfgSAA_DisplayLine = False
End Sub

Public Sub fgrTextField_Reset()
fgrTextField.Clear
fgrTextField_Sort1 = 0: fgrTextField_Sort2 = 0
fgrTextField_Sort1_Old = -1
fgrTextField_RowDisplay = 0: fgrTextField_RowClick = 0
fgrTextField_arrIndex = fgrTextField.Cols - 1
blnfgrTextField_DisplayLine = False
End Sub

Public Sub fgSAA_Sort()
If fgSAA.Rows > 1 Then
    fgSAA.Row = 1
    fgSAA.RowSel = fgSAA.Rows - 1
    
    If fgSAA_Sort1_Old = fgSAA_Sort1 Then
        If fgSAA_SortAD = 5 Then
            fgSAA_SortAD = 6
        Else
            fgSAA_SortAD = 5
        End If
    Else
        fgSAA_SortAD = 5
    End If
    fgSAA_Sort1_Old = fgSAA_Sort1
    
    fgSAA.Col = fgSAA_Sort1
    fgSAA.ColSel = fgSAA_Sort2
    fgSAA.Sort = fgSAA_SortAD
End If
'cboDevise_Reset
End Sub


Public Sub fgrTextField_Sort()
If fgrTextField.Rows > 1 Then
    fgrTextField.Row = 1
    fgrTextField.RowSel = fgrTextField.Rows - 1
    
    If fgrTextField_Sort1_Old = fgrTextField_Sort1 Then
        If fgrTextField_SortAD = 5 Then
            fgrTextField_SortAD = 6
        Else
            fgrTextField_SortAD = 5
        End If
    Else
        fgrTextField_SortAD = 5
    End If
    fgrTextField_Sort1_Old = fgrTextField_Sort1
    
    fgrTextField.Col = fgrTextField_Sort1
    fgrTextField.ColSel = fgrTextField_Sort2
    fgrTextField.Sort = fgrTextField_SortAD
End If
End Sub

Public Sub fgSAA_SortX(lK As Integer)
Dim I As Integer, K  As Integer, X As String
Dim xCur As Currency

For I = 1 To fgSAA.Rows - 1
    fgSAA.Row = I
    fgSAA.Col = lK
    Select Case lK
        Case 13: fgSAA.Col = fgSAA_arrIndex: K = CLng(fgSAA.Text)
                 xrMesg = arrrMesg(K)
                 X = xrMesg.x_inst0_unit_name & xrMesg.mesg_type & xrMesg.mesg_crea_date_time

        Case 5:
            xCur = Val(fgSAA.Text)
            X = Format$(xCur, "000000000000000.00")
   '     Case 9: X = Format$(Val(fgSAA.Text), "000000000000000")
    End Select
    fgSAA.Col = fgSAA_arrIndex - 1
    fgSAA.Text = X
Next I


fgSAA_Sort1 = fgSAA_arrIndex - 1: fgSAA_Sort2 = fgSAA_arrIndex - 1
fgSAA_Sort
End Sub



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
Public Sub fgZSWIHIA0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgZSWIHIA0.Row

If lRow > 0 And lRow < fgZSWIHIA0.Rows Then
    fgZSWIHIA0.Row = lRow
    For I = 0 To fgZSWIHIA0_arrIndex
        fgZSWIHIA0.Col = I: fgZSWIHIA0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgZSWIHIA0.Row = mRow
    If fgZSWIHIA0.Row > 0 Then
        lRow = fgZSWIHIA0.Row
        lColor_Old = fgZSWIHIA0.CellBackColor
        For I = 0 To fgZSWIHIA0_arrIndex
          fgZSWIHIA0.Col = I: fgZSWIHIA0.CellBackColor = lColor
        Next I
        fgZSWIHIA0.Col = 0
    End If
End If

End Sub

Private Sub fgSelect_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
cmdPrint.Enabled = False

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
For I = 1 To arrYSWIMON0_Nb
         
    xYSWIMON0 = arrYSWIMON0(I)
    If xYSWIMON0.SWIMONID > 0 Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
    End If
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrYSWIMON0_Nb): DoEvents
If fgSelect.Rows > 1 Then
    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    'cmdPrint.Enabled = True
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgrTextField_Display(lAid As Integer, ls_umidl As Long, ls_umidh As Long)
Dim V
Dim X As String, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Long

On Error GoTo Error_Handler
fgrTextField.Visible = False
fgrTextField_Reset
cmdPrint.Enabled = False

fgrTextField.Rows = 1
fgrTextField.FormatString = fgrTextField_FormatString
currentAction = "fgrTextField_Display"

ReDim arrrTextField(101)
arrrTextField_Max = 100: arrrTextField_Nb = 0
 
blnOk = False
Set rsSIDE_DB = Nothing

xSQL = "select * from rTextField " _
    & "where Aid = " & lAid _
    & " and text_s_umidl = " & ls_umidl _
    & " and text_s_umidh  = " & ls_umidh

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrTextField_GetBuffer_ODBC(rsSIDE_DB, xrTextField)

     If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
        Exit Sub
     Else
         arrrTextField_Nb = arrrTextField_Nb + 1
         If arrrTextField_Nb = arrrTextField_Max Then   '>
            If arrrTextField_Max >= 10000 Then
                MsgBox "10000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
                arrrTextField_Max = arrrTextField_Max + 50
                ReDim Preserve arrrTextField(arrrTextField_Max + 1)
            End If
         End If
         
         arrrTextField(arrrTextField_Nb) = xrTextField
    End If
    rsSIDE_DB.MoveNext

Loop

For I = 1 To arrrTextField_Nb

    xrTextField = arrrTextField(I)
    fgrTextField_DisplayLine (I)

Next I

fgrTextField.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & fgrTextField.Rows - 1): DoEvents
'If fgSAA.Rows > 1 Then
'    fgrTextField_Sort1 = 0: fgrTextField_Sort2 = 1: fgrTextField_Sort
'    cmdPrint.Enabled = True
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub fgZSWIHIA0_Display()
Dim I As Integer
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgZSWIHIA0.Visible = False
fgZSWIHIA0_Reset
cmdPrint.Enabled = False

fgZSWIHIA0.Rows = 1
fgZSWIHIA0.FormatString = fgZSWIHIA0_FormatString
currentAction = "fgZSWIHIA0_Display"
    
For I = 1 To arrZSWIHIA0_Nb
         
    xZSWIHIA0 = arrZSWIHIA0(I)
    fgZSWIHIA0_DisplayLine I
Next I

fgZSWIHIA0.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrZSWIHIA0_Nb): DoEvents
If fgZSWIHIA0.Rows > 1 Then
    fgZSWIHIA0_Sort1 = 10: fgZSWIHIA0_Sort2 = 10: fgZSWIHIA0_SortX 10
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub arrZSWIHIA0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim K As Integer

On Error GoTo Error_Handler
ReDim arrZSWIHIA0(101)
arrZSWIHIA0_Max = 100: arrZSWIHIA0_Nb = 0

Set rsSab = Nothing
K = 1
Do
    K = InStr(K, xWhere, "SWIFTA")
    If K > 0 Then Mid$(xWhere, K, 6) = "SWIHIA"
Loop Until K <= 0

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIHIA0 " & xWhere & " order by SWIHIADES , SWIHIAMES"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = srvZSWIHIA0_GetBuffer_ODBC(rsSab, xZSWIHIA0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWIHIA0_Display"
        '' Exit Sub
     Else
         arrZSWIHIA0_Nb = arrZSWIHIA0_Nb + 1
         If arrZSWIHIA0_Nb > arrZSWIHIA0_Max Then
            If arrZSWIHIA0_Max > 1000 Then
                MsgBox "1000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Sub
            Else
                arrZSWIHIA0_Max = arrZSWIHIA0_Max + 50
                ReDim Preserve arrZSWIHIA0(arrZSWIHIA0_Max)
            End If
         End If
         
         arrZSWIHIA0(arrZSWIHIA0_Nb) = xZSWIHIA0
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


Public Sub fgZSWIFTA0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgZSWIFTA0.Row

If lRow > 0 And lRow < fgZSWIFTA0.Rows Then
    fgZSWIFTA0.Row = lRow
    For I = 0 To fgZSWIFTA0_arrIndex
        fgZSWIFTA0.Col = I: fgZSWIFTA0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgZSWIFTA0.Row = mRow
    If fgZSWIFTA0.Row > 0 Then
        lRow = fgZSWIFTA0.Row
        lColor_Old = fgZSWIFTA0.CellBackColor
        For I = 0 To fgZSWIFTA0_arrIndex
          fgZSWIFTA0.Col = I: fgZSWIFTA0.CellBackColor = lColor
        Next I
        
    End If
End If
fgZSWIFTA0.LeftCol = 0
End Sub

Private Sub fgZSWIFTA0_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgZSWIFTA0.Visible = False
fgZSWIFTA0_Reset
cmdPrint.Enabled = False

fgZSWIFTA0.Rows = 1
fgZSWIFTA0.FormatString = fgZSWIFTA0_FormatString
currentAction = "fgZSWIFTA0_Display"
    
For I = 1 To arrZSWIFTA0_Nb
         
    xZSWIFTA0 = arrZSWIFTA0(I)
    fgZSWIFTA0_DisplayLine I
Next I

fgZSWIFTA0.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrZSWIFTA0_Nb): DoEvents
If fgZSWIFTA0.Rows > 1 Then
    fgZSWIFTA0_Sort1 = 1: fgZSWIFTA0_Sort2 = 2: fgZSWIFTA0_Sort
'    cmdPrint.Enabled = True
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub arrZSWIFTA0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler
ReDim arrZSWIFTA0(101)
arrZSWIFTA0_Max = 100: arrZSWIFTA0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIFTA0 " & xWhere & " order by SWIFTADES , SWIFTAMES"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = srvZSWIFTA0_GetBuffer_ODBC(rsSab, xZSWIFTA0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWIFTA0_Display"
        '' Exit Sub
     Else
         arrZSWIFTA0_Nb = arrZSWIFTA0_Nb + 1
         If arrZSWIFTA0_Nb > arrZSWIFTA0_Max Then
             arrZSWIFTA0_Max = arrZSWIFTA0_Max + 50
             ReDim Preserve arrZSWIFTA0(arrZSWIFTA0_Max)
         End If
         
         arrZSWIFTA0(arrZSWIFTA0_Nb) = xZSWIFTA0
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
Private Sub arrZSWIALI0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim I As Integer, K As Integer, wX As String, lenX As String
On Error GoTo Error_Handler
ReDim arrZSWIALI0(101)
arrZSWIALI0_Max = 100: arrZSWIALI0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIALI0 " & xWhere _
     & " order by SWIALIETA , SWIALIAGE , SWIALISER , SWIALISSE , SWIALIMES , SWIALINUM , SWIALINEN , SWIALINLI"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    V = srvZSWIALI0_GetBuffer_ODBC(rsSab, xZSWIALI0)
'----------------------------------------------
     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWIALI0_Display"
        '' Exit Sub
     Else
         arrZSWIALI0_Nb = arrZSWIALI0_Nb + 1
         If arrZSWIALI0_Nb > arrZSWIALI0_Max Then
             arrZSWIALI0_Max = arrZSWIALI0_Max + 50
             ReDim Preserve arrZSWIALI0(arrZSWIALI0_Max)
         End If

         arrZSWIALI0(arrZSWIALI0_Nb) = xZSWIALI0
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

Private Sub arrZSWITEM0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler
ReDim arrZSWITEM0(101)
arrZSWITEM0_Max = 100: arrZSWITEM0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWITEM0 " & xWhere & " order by SWITEMNUM , SWITEMCON"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = srvZSWITEM0_GetBuffer_ODBC(rsSab, xZSWITEM0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWITEM0_Display"
        '' Exit Sub
     Else
         arrZSWITEM0_Nb = arrZSWITEM0_Nb + 1
         If arrZSWITEM0_Nb > arrZSWITEM0_Max Then
             arrZSWITEM0_Max = arrZSWITEM0_Max + 50
             ReDim Preserve arrZSWITEM0(arrZSWITEM0_Max)
         End If
         
         arrZSWITEM0(arrZSWITEM0_Nb) = xZSWITEM0
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
Private Sub arrZSWIHIT0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler
ReDim arrZSWIHIT0(101)
arrZSWIHIT0_Max = 100: arrZSWIHIT0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIHIT0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = srvZSWIHIT0_GetBuffer_ODBC(rsSab, xZSWIHIT0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWIHIT0_Display"
        '' Exit Sub
     Else
         arrZSWIHIT0_Nb = arrZSWIHIT0_Nb + 1
         If arrZSWIHIT0_Nb > arrZSWIHIT0_Max Then
             arrZSWIHIT0_Max = arrZSWIHIT0_Max + 50
             ReDim Preserve arrZSWIHIT0(arrZSWIHIT0_Max)
         End If
         
         arrZSWIHIT0(arrZSWIHIT0_Nb) = xZSWIHIT0
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

Private Sub arrZSWIHIC0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler
ReDim arrZSWIHIC0(101)
arrZSWIHIC0_Max = 100: arrZSWIHIC0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIHIC0 " & xWhere & " order by  SWIHICNLI"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = srvZSWIHIC0_GetBuffer_ODBC(rsSab, xZSWIHIC0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWIHIC0_Display"
        '' Exit Sub
     Else
         arrZSWIHIC0_Nb = arrZSWIHIC0_Nb + 1
         If arrZSWIHIC0_Nb > arrZSWIHIC0_Max Then
             arrZSWIHIC0_Max = arrZSWIHIC0_Max + 50
             ReDim Preserve arrZSWIHIC0(arrZSWIHIC0_Max)
         End If
         
         arrZSWIHIC0(arrZSWIHIC0_Nb) = xZSWIHIC0
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


Private Sub arrYSWIMON0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYSWIMON0(101)
arrYSWIMON0_Max = 100: arrYSWIMON0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0 " & xWhere & " order by SWIMONID"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = srvYSWIMON0_GetBuffer_ODBC(rsSab, xYSWIMON0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgselect_Display"
        '' Exit Sub
     Else
         arrYSWIMON0_Nb = arrYSWIMON0_Nb + 1
         If arrYSWIMON0_Nb > arrYSWIMON0_Max Then
             arrYSWIMON0_Max = arrYSWIMON0_Max + 50
             ReDim Preserve arrYSWIMON0(arrYSWIMON0_Max)
         End If
         
         arrYSWIMON0(arrYSWIMON0_Nb) = xYSWIMON0
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
Private Sub arrrInst_Status_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler



'>>  Lecture des interventions à partir de la date stockée dans l'entête ZSWIMON0

ReDim arrrIntv_Status(101)
arrrIntv_Status_Max = 100: arrrIntv_Status_Nb = 0
ReDim arrrIntv_Flag(101)
arrrIntv_Flag_Max = 100: arrrIntv_Flag_Nb = 0
ReDim arrrInst_Status(101)
arrrInst_Status_Max = 100: arrrInst_Status_Nb = 0
res_date_time = 0

Set rsSIDE_DB = Nothing

'xSQL = "select * from rIntv " & xWhere & " order by aid, intv_s_umidl, intv_s_umidh, intv_date_time,intv_seq_nbr"
xSQL = "select * from rIntv , rMesg " & xWhere _
     & " and rIntv.AID = rMesg.AId and intv_s_umidl = mesg_s_umidl and intv_s_umidh = mesg_s_umidh" _
     & " and mesg_crea_date_time >= " & ts_DSys_7 _
     & " order by rIntv.AID, intv_s_umidl, intv_s_umidh, intv_date_time,intv_seq_nbr"
Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrIntv_GetBuffer_ODBC(rsSIDE_DB, xrIntv)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.arrrInst_Status_SQL"
        '' Exit Sub
     Else
         arrrIntv_Status_Nb = arrrIntv_Status_Nb + 1
         If arrrIntv_Status_Nb > arrrIntv_Status_Max Then
             arrrIntv_Status_Max = arrrIntv_Status_Max + 50
             ReDim Preserve arrrIntv_Status(arrrIntv_Status_Max)
         End If
         
         arrrIntv_Status(arrrIntv_Status_Nb) = xrIntv
         ' Garder l'heure (et date) la plus lointaine
         merIntv_Status = arrrIntv_Status(arrrIntv_Status_Nb)
         If merIntv_Status.intv_date_time > res_date_time Then
            res_date_time = merIntv_Status.intv_date_time
         End If
         
         ' Mémorisation des interventions OFAC et Modification, update YSWIMON0 après rapprochement
         Select Case Trim(merIntv_Status.intv_appl_serv_name)
            Case "Mesg Modification", "OFCA_Interface":
                         arrrIntv_Flag_Nb = arrrIntv_Flag_Nb + 1
                        If arrrIntv_Flag_Nb > arrrIntv_Flag_Max Then
                            arrrIntv_Flag_Max = arrrIntv_Flag_Max + 50
                            ReDim Preserve arrrIntv_Flag(arrrIntv_Flag_Max)
                        End If
                        arrrIntv_Flag(arrrIntv_Flag_Nb) = xrIntv

         End Select
    End If
    rsSIDE_DB.MoveNext

Loop

'>>  Le tableau - arrrIntv_Status - est trié par no message SWIFT puis intv_date_time
'>>  Lecture des instances à partir de la dernière date lue d'un message du - arrrIntv_Status -

srvrIntv_Init resIntv_Status

For arrrIntv_Status_Index = arrrIntv_Status_Nb To 1 Step -1

    merIntv_Status = arrrIntv_Status(arrrIntv_Status_Index)


    If merIntv_Status.intv_s_umidh <> resIntv_Status.intv_s_umidh _
    Or merIntv_Status.intv_s_umidl <> resIntv_Status.intv_s_umidl Then
       
        Set rsSIDE_DB = Nothing
        
        xSQL = "select * from rInst where Aid = " & merIntv_Status.Aid _
                & " and inst_s_umidl = " & merIntv_Status.intv_s_umidl _
                & " and inst_s_umidh = " & merIntv_Status.intv_s_umidh _
                & " and inst_num = 0"
                
        '2004.10.26 jpl         & " and inst_crea_rp_name = '_AI_from_APPLI' " _
        ' Il est possible que rIntv ne génère aucun enreg lu dans rInst à cause de inst_crea_rp_name = '_AI_from_APPLI'
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        
        Do While Not rsSIDE_DB.EOF
            V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
        
             If Not IsNull(V) Then
                 MsgBox V, vbCritical, "frmSwift_Messages.arrrInst_Status_SQL : rInst"
                 Exit Sub ''20070530 JPL
             Else
                 arrrInst_Status_Nb = arrrInst_Status_Nb + 1
                 If arrrInst_Status_Nb > arrrInst_Status_Max Then
                     arrrInst_Status_Max = arrrInst_Status_Max + 50
                     ReDim Preserve arrrInst_Status(arrrInst_Status_Max)
                 End If
                 
                 arrrInst_Status(arrrInst_Status_Nb) = xrInst
            End If
            rsSIDE_DB.MoveNext
        
        Loop
        
        resIntv_Status = arrrIntv_Status(arrrIntv_Status_Index)
        
    End If

Next arrrIntv_Status_Index

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 3
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub arrZSWIFTB0_SQL(xWhere As String, blnZSWIFTC0_Load As Boolean)
Dim V
Dim X As String, xSQL As String
Dim K As Integer

On Error GoTo Error_Handler
ReDim arrZSWIFTB0(51)
arrZSWIFTB0_Max = 50: arrZSWIFTB0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIFTB0 " & xWhere & " order by SWIFTBNLI"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = srvZSWIFTB0_GetBuffer_ODBC(rsSab, xZSWIFTB0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWIFTB0_Display"
        '' Exit Sub
     Else
         arrZSWIFTB0_Nb = arrZSWIFTB0_Nb + 1
         If arrZSWIFTB0_Nb > arrZSWIFTB0_Max Then
             arrZSWIFTB0_Max = arrZSWIFTB0_Max + 50
             ReDim Preserve arrZSWIFTB0(arrZSWIFTB0_Max)
         End If
         
         arrZSWIFTB0(arrZSWIFTB0_Nb) = xZSWIFTB0
    End If
    rsSab.MoveNext

Loop


If blnZSWIFTC0_Load Then
    arrZSWIFTC0_Max = arrZSWIFTB0_Max
    arrZSWIFTC0_Nb = 0
    ReDim arrZSWIFTC0(arrZSWIFTC0_Max)
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIFTC0 " & xWhere & " order by SWIFTCNLI"
Else
    xSQL = "select SWIFTCNLI,SWIFTCSIG from " & paramIBM_Library_SAB & ".ZSWIFTC0 " & xWhere & " order by SWIFTCNLI"
End If

ReDim arrSWIFTCSIG(arrZSWIFTB0_Max)
    
Set rsSab = Nothing

K = 1
Do
    K = InStr(K, xSQL, "SWIFTB")
    If K > 0 Then Mid$(xSQL, K, 6) = "SWIFTC"
Loop Until K = 0

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF

'Indicateur pour affichage du message
'=====================================================
    K = rsSab("SWIFTCNLI")
    If K > 0 And K < arrZSWIFTB0_Nb Then arrSWIFTCSIG(K) = rsSab("SWIFTCSIG")
    
'Chargement ZSWIFTC0 pour maj après transfert vers SAA
'=====================================================
    If blnZSWIFTC0_Load Then
        V = srvZSWIFTC0_GetBuffer_ODBC(rsSab, xZSWIFTC0)
        If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWIFTC0_Display"
        '' Exit Sub
     Else
         arrZSWIFTC0_Nb = arrZSWIFTC0_Nb + 1
         If arrZSWIFTC0_Nb > arrZSWIFTC0_Max Then
             arrZSWIFTC0_Max = arrZSWIFTC0_Max + 50
             ReDim Preserve arrZSWIFTC0(arrZSWIFTC0_Max)
         End If
         
         arrZSWIFTC0(arrZSWIFTC0_Nb) = xZSWIFTC0
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

Private Sub arrZSWIHIB0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim K As Integer
On Error GoTo Error_Handler
ReDim arrZSWIHIB0(101)
arrZSWIHIB0_Max = 100: arrZSWIHIB0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIHIB0 " & xWhere & " order by SWIHIBNLI"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = srvZSWIHIB0_GetBuffer_ODBC(rsSab, xZSWIHIB0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgZSWIHIB0_Display"
        '' Exit Sub
     Else
         arrZSWIHIB0_Nb = arrZSWIHIB0_Nb + 1
         If arrZSWIHIB0_Nb > arrZSWIHIB0_Max Then
             arrZSWIHIB0_Max = arrZSWIHIB0_Max + 50
             ReDim Preserve arrZSWIHIB0(arrZSWIHIB0_Max)
         End If
         
         arrZSWIHIB0(arrZSWIHIB0_Nb) = xZSWIHIB0
    End If
    rsSab.MoveNext

Loop

ReDim arrSWIHICSIG(arrZSWIHIB0_Max)
Set rsSab = Nothing


xSQL = "select SWIHICNLI,SWIHICSIG from " & paramIBM_Library_SAB & ".ZSWIHIC0 " & xWhere & " order by SWIHICNLI"
K = 1
Do
    K = InStr(K, xSQL, "SWIHIB")
    If K > 0 Then Mid$(xSQL, K, 6) = "SWIHIC"
Loop Until K = 0

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    K = rsSab("SWIHICNLI")
    If K > 0 And K < arrZSWIHIB0_Nb Then arrSWIHICSIG(K) = rsSab("SWIHICSIG")
    
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrZSWIFTB0_lstW()
Dim V, X As String, X1 As String, X2 As String
On Error GoTo Error_Handler
currentAction = "arrZSWIFTB0_fgDisplay"
lstW.Clear
lstW.Visible = True
For I = 1 To arrZSWIFTB0_Nb
    If arrSWIFTCSIG(I) = ">" Then
        X = arrZSWIFTB0(I).SWIFTBDET
        If Mid$(X, 4, 1) = ":" Then
            X1 = Mid$(X, 1, 4): X2 = Mid$(X, 5, Len(X) - 4)
        Else
        
            If Mid$(X, 5, 1) = ":" Then
                X1 = Mid$(X, 1, 5): X2 = Mid$(X, 6, Len(X) - 5)
            Else
                X1 = "": X2 = X
            End If
        End If
          
             
        lstW.AddItem X1 & Chr$(9) & X2
    End If
Next I

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub



Private Sub arrZSWIHIB0_lstW()
Dim V, X As String, X1 As String, X2 As String
On Error GoTo Error_Handler
currentAction = "arrZSWIHIB0_fgDisplay"
lstW.Clear
lstW.Visible = True
For I = 1 To arrZSWIHIB0_Nb
    If arrSWIHICSIG(I) = ">" Then
        X = arrZSWIHIB0(I).SWIHIBDET
        If Mid$(X, 4, 1) = ":" Then
            X1 = Mid$(X, 1, 4): X2 = Mid$(X, 5, Len(X) - 4)
        Else
        
            If Mid$(X, 5, 1) = ":" Then
                X1 = Mid$(X, 1, 5): X2 = Mid$(X, 6, Len(X) - 5)
            Else
                X1 = "": X2 = X
            End If
        End If
          
             
        lstW.AddItem X1 & Chr$(9) & X2
    End If
Next I

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdZSWIFTA0_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdZSWIFTA0_SQL"
Call DTPicker_Amj7(txtSelect_SWIFTADVA, wAmj7)
xWhere = " where SWIFTADVA <= " & wAmj7 & " "


X = Mid$(cboSelect_SWIFTAXXX, 1, 3)
If X <> "   " Then
    If Mid$(X, 1, 1) <> "_" Then xWhere = xWhere & " and " & "SWIFTAVAL = '" & Mid$(X, 1, 1) & "'"
    If Mid$(X, 2, 1) <> "_" Then xWhere = xWhere & " and " & "SWIFTACOM = '" & Mid$(X, 2, 1) & "'"
    If Mid$(X, 3, 1) <> "_" Then xWhere = xWhere & " and " & "SWIFTASUP = '" & Mid$(X, 3, 1) & "'"
End If


X = Trim(txtSelect_SWIFTAREF)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "SWIFTAREF like '%" & X & "%'"
End If

X = Trim(cboSelect_SWIFTADE1)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "SWIFTADE1 = '" & X & "'"
End If

X = Trim(txtSelect_SWIFTADES)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "SWIFTADES like '%" & X & "%'"
End If

X = Trim(txtSelect_SWIFTAMES)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "SWIFTAMES like '" & X & "%'"
End If

X = Trim(cboSelect_SWIFTASER)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "SWIFTASER like '%" & X & "%'"
End If

currentAction = "arrZSWIFTA0_SQL"
arrZSWIFTA0_SQL xWhere

If chkSelect_ZSWIHIA0 = "1" Then
    fgZSWIFTA0.Height = fgZSWIFTA0_Height
Else
    fgZSWIFTA0.Height = fgZSWIFTA0_Height + fgZSWIHIA0.Height
End If


fgZSWIFTA0_Display

If chkSelect_ZSWIHIA0 = "1" Then
    If Trim(txtSelect_SWIFTAREF) = "" Then
        Call lstErr_AddItem(lstErr, cmdContext, "!!! Historique : préciser une référence"): DoEvents
    Else
        currentAction = "arrZSWIHIA0_SQL"
        arrZSWIHIA0_SQL xWhere
        fgZSWIHIA0_Display
    End If
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgZSWIFTA0_DisplayLine(lIndex As Long)
On Error Resume Next
fgZSWIFTA0.Rows = fgZSWIFTA0.Rows + 1
fgZSWIFTA0.Row = fgZSWIFTA0.Rows - 1
fgZSWIFTA0.Col = 0: fgZSWIFTA0.Text = xZSWIFTA0.SWIFTADES
fgZSWIFTA0.Col = 1: fgZSWIFTA0.Text = xZSWIFTA0.SWIFTAMES
fgZSWIFTA0.Col = 2: fgZSWIFTA0.Text = xZSWIFTA0.SWIFTAREF
fgZSWIFTA0.Col = 3: fgZSWIFTA0.Text = Format$(xZSWIFTA0.SWIFTAMON, "### ### ### ###.00")
fgZSWIFTA0.Col = 4: fgZSWIFTA0.Text = xZSWIFTA0.SWIFTADE1
fgZSWIFTA0.Col = 5: fgZSWIFTA0.Text = dateIBM10(xZSWIFTA0.SWIFTADVA, True)
fgZSWIFTA0.Col = 6: fgZSWIFTA0.Text = xZSWIFTA0.SWIFTAVAL
fgZSWIFTA0.Col = 7: fgZSWIFTA0.Text = xZSWIFTA0.SWIFTACOM
fgZSWIFTA0.Col = 8: fgZSWIFTA0.Text = xZSWIFTA0.SWIFTASUP
fgZSWIFTA0.Col = 9: fgZSWIFTA0.Text = Format$(xZSWIFTA0.SWIFTANUM, "### ### ### ###")
fgZSWIFTA0.Col = 10: fgZSWIFTA0.Text = dateIBM10(xZSWIFTA0.SWIFTADEN, True) & " " & timeImp8(xZSWIFTA0.SWIFTAHEN)

fgZSWIFTA0.Col = fgZSWIFTA0_arrIndex: fgZSWIFTA0.Text = lIndex
End Sub



Public Sub fgZSWIFTA0_Reset()
fgZSWIFTA0.Clear
fgZSWIFTA0_Sort1 = 0: fgZSWIFTA0_Sort2 = 0
fgZSWIFTA0_Sort1_Old = -1
fgZSWIFTA0_RowDisplay = 0: fgZSWIFTA0_RowClick = 0
fgZSWIFTA0_arrIndex = fgZSWIFTA0.Cols - 1
blnfgZSWIFTA0_DisplayLine = False
End Sub


Public Sub fgZSWIFTA0_Sort()
If fgZSWIFTA0.Rows > 1 Then
    fgZSWIFTA0.Row = 1
    fgZSWIFTA0.RowSel = fgZSWIFTA0.Rows - 1
    
    If fgZSWIFTA0_Sort1_Old = fgZSWIFTA0_Sort1 Then
        If fgZSWIFTA0_SortAD = 5 Then
            fgZSWIFTA0_SortAD = 6
        Else
            fgZSWIFTA0_SortAD = 5
        End If
    Else
        fgZSWIFTA0_SortAD = 5
    End If
    fgZSWIFTA0_Sort1_Old = fgZSWIFTA0_Sort1
    
    fgZSWIFTA0.Col = fgZSWIFTA0_Sort1
    fgZSWIFTA0.ColSel = fgZSWIFTA0_Sort2
    fgZSWIFTA0.Sort = fgZSWIFTA0_SortAD
End If
'cboDevise_Reset
End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xYSWIMON0.SAAUNIT
fgSelect.Col = 1: fgSelect.Text = xYSWIMON0.SWIMONXMT
fgSelect.Col = 2: fgSelect.Text = xYSWIMON0.SWIMONX20
fgSelect.Col = 3: fgSelect.Text = Format$(xYSWIMON0.SWIMONX32A, "### ### ### ###.00")
fgSelect.Col = 4: fgSelect.Text = xYSWIMON0.SWIMONX32D
fgSelect.Col = 5: fgSelect.Text = dateImp10(xYSWIMON0.SWIMONX32V)

fgSelect.Col = 6: fgSelect.Text = xYSWIMON0.SAAQUEUE
fgSelect.Col = 7: fgSelect.Text = xYSWIMON0.SWIMONID
fgSelect.Col = 8: fgSelect.Text = xYSWIMON0.SWIMONSTA
fgSelect.Col = 9: fgSelect.Text = dateImp10(xYSWIMON0.SWIMONSTAD) & " " & Format$(xYSWIMON0.SWIMONSTAH, "@@:@@:@@")
fgSelect.Col = 10: fgSelect.Text = xYSWIMON0.SWIMONFLUQ & " " & dateImp10(xYSWIMON0.SWIMONFLUD) & " " & Format$(xYSWIMON0.SWIMONFLUH, "@@:@@:@@") & " " & xYSWIMON0.SWIMONFLUS
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub


Public Sub fgZSWIHIA0_DisplayLine(lIndex As Integer)
On Error Resume Next
fgZSWIHIA0.Rows = fgZSWIHIA0.Rows + 1
fgZSWIHIA0.Row = fgZSWIHIA0.Rows - 1
fgZSWIHIA0.Col = 0: fgZSWIHIA0.Text = xZSWIHIA0.SWIHIADES
fgZSWIHIA0.Col = 1: fgZSWIHIA0.Text = xZSWIHIA0.SWIHIAMES
fgZSWIHIA0.Col = 2: fgZSWIHIA0.Text = xZSWIHIA0.SWIHIAREF
fgZSWIHIA0.Col = 3: fgZSWIHIA0.Text = Format$(xZSWIHIA0.SWIHIAMON, "### ### ### ###.00")
fgZSWIHIA0.Col = 4: fgZSWIHIA0.Text = xZSWIHIA0.SWIHIADE1
fgZSWIHIA0.Col = 5: fgZSWIHIA0.Text = dateIBM10(xZSWIHIA0.SWIHIADVA, True)
fgZSWIHIA0.Col = 6: fgZSWIHIA0.Text = xZSWIHIA0.SWIHIAVAL
fgZSWIHIA0.Col = 7: fgZSWIHIA0.Text = xZSWIHIA0.SWIHIACOM
fgZSWIHIA0.Col = 8: fgZSWIHIA0.Text = xZSWIHIA0.SWIHIASUP
fgZSWIHIA0.Col = 9: fgZSWIHIA0.Text = Format$(xZSWIHIA0.SWIHIANUM, "### ### ### ###")
fgZSWIHIA0.Col = 10: fgZSWIHIA0.Text = dateIBM10(xZSWIHIA0.SWIHIADEN, True) & " " & timeNImp8(xZSWIHIA0.SWIHIAHEN)

fgZSWIHIA0.Col = fgZSWIHIA0_arrIndex: fgZSWIHIA0.Text = lIndex

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

Public Sub fgZSWIHIA0_Reset()
fgZSWIHIA0.Clear
fgZSWIHIA0_Sort1 = 0: fgZSWIHIA0_Sort2 = 0
fgZSWIHIA0_Sort1_Old = -1
fgZSWIHIA0_RowDisplay = 0: fgZSWIHIA0_RowClick = 0
fgZSWIHIA0_arrIndex = fgZSWIHIA0.Cols - 1
blnfgZSWIHIA0_DisplayLine = False
fgZSWIHIA0_SortAD = 6
fgZSWIHIA0.LeftCol = 0

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
Public Sub fgZSWIHIA0_Sort()
If fgZSWIHIA0.Rows > 1 Then
    fgZSWIHIA0.Row = 1
    fgZSWIHIA0.RowSel = fgZSWIHIA0.Rows - 1
    
    If fgZSWIHIA0_Sort1_Old = fgZSWIHIA0_Sort1 Then
        If fgZSWIHIA0_SortAD = 5 Then
            fgZSWIHIA0_SortAD = 6
        Else
            fgZSWIHIA0_SortAD = 5
        End If
    Else
        fgZSWIHIA0_SortAD = 5
    End If
    fgZSWIHIA0_Sort1_Old = fgZSWIHIA0_Sort1
    
    fgZSWIHIA0.Col = fgZSWIHIA0_Sort1
    fgZSWIHIA0.ColSel = fgZSWIHIA0_Sort2
    fgZSWIHIA0.Sort = fgZSWIHIA0_SortAD
End If

End Sub

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        X = fgSelect.Text
    Else
        X = ""
    End If
    
    fgSelect.Col = 3
    X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

Public Sub fgZSWIHIA0_SortX(lK As Integer)
Dim I As Integer, X As String, K As Integer
For I = 1 To fgZSWIHIA0.Rows - 1
    fgZSWIHIA0.Row = I
    fgZSWIHIA0.Col = lK
    Select Case lK
        Case 3: X = Format$(Val(fgZSWIHIA0.Text), "000000000000000.00")
        Case 5: Call dateJMA_AMJ(Trim(fgZSWIHIA0.Text), X)
        Case 9: X = Format$(Val(fgZSWIHIA0.Text), "000000000000000")
        Case 10:
                fgZSWIHIA0.Col = fgZSWIHIA0_arrIndex:  K = CLng(fgZSWIHIA0.Text)
                X = arrZSWIHIA0(K).SWIHIADEN & Format$(arrZSWIHIA0(K).SWIHIAHEN, "000000")
    End Select
    fgZSWIHIA0.Col = fgZSWIHIA0_arrIndex - 1
    fgZSWIHIA0.Text = X
Next I


fgZSWIHIA0_Sort1 = fgZSWIHIA0_arrIndex - 1: fgZSWIHIA0_Sort2 = fgZSWIHIA0_arrIndex - 1
fgZSWIHIA0_Sort
End Sub

Public Sub fgZSWIFTA0_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgZSWIFTA0.Rows - 1
    fgZSWIFTA0.Row = I
    fgZSWIFTA0.Col = lK
    Select Case lK
        Case 3: X = Format$(Val(fgZSWIFTA0.Text), "000000000000000.00")
        Case 5: Call dateJMA_AMJ(Trim(fgZSWIFTA0.Text), X)
        Case 9: X = Format$(Val(fgZSWIFTA0.Text), "000000000000000")
    End Select
    fgZSWIFTA0.Col = fgZSWIFTA0_arrIndex - 1
    fgZSWIFTA0.Text = X
Next I


fgZSWIFTA0_Sort1 = fgZSWIFTA0_arrIndex - 1: fgZSWIFTA0_Sort2 = fgZSWIFTA0_arrIndex - 1
fgZSWIFTA0_Sort
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub mnuSelect_SAAAID_Status()   ' RECHERCHE DE STATUT COURANT DU MESSAGE

Dim V, W
Dim xSQL As String
Dim blnOFAC As Boolean, blnNAK As Boolean, blnCompleted As Boolean


' Toujours 1 seule instance dont les valeurs des zones utilisées changent
Set rsSIDE_DB = Nothing
xSQL = "select * from rInst " _
        & "where Aid = " & xYSWIMON0.SAAAID _
        & " and inst_s_umidl = " & xYSWIMON0.SAAUMIDL _
        & " and inst_s_umidh  = " & xYSWIMON0.SAAUMIDH _
        & " and inst_num = 0 "
        
blnOFAC = False
blnCompleted = False

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
If Not rsSIDE_DB.EOF Then
    V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
    If Not IsNull(V) Then
        ' Test si bloqué dans OFAC par rIntv - S210-
        blnOFAC = True
    Else
        xYSWIMON0.SAAQUEUE = xrInst.inst_rp_name
        If xrInst.inst_mpfn_name = "OFCA_Check" Then xYSWIMON0.SAAQOFAC = 1
        
        Select Case xrInst.inst_status
            ' Message en LIVE
            Case "LIVE":
                Select Case xYSWIMON0.SAAQUEUE
                    Case "_MP_mod_text": xYSWIMON0.SWIMONSTA = "S220"
                    Case "_MP_authorisation": xYSWIMON0.SWIMONSTA = "S230"
                End Select
            ' Message COMPLETED - Voir si SUPPRIME ou ACK ou NAK ou REJECTED
            Case "COMPLETED":
                If Trim(xrInst.inst_mpfn_name) = "mpm" And _
                   Trim(xrInst.inst_auth_oper_nickname) = "" And _
                   Trim(xrInst.x_last_emi_appe_date_time) = "00:00:00" And _
                   xrInst.x_last_emi_appe_seq_nbr = 0 Then
                   
                   xYSWIMON0.SWIMONSTA = "S904"
                Else
                
                    blnCompleted = True
                End If
        End Select
        
    End If
End If

' ==> Soit blnOFAC / Soit blnCompleted MAIS JAMAIS les deux en même temps

If blnOFAC Then    'Violation et blocage dans OFAC

    Set rsSIDE_DB = Nothing
    xSQL = "select * from rIntv where intv_inty_category = 'INTY_OTHER' and " _
                & "intv_inst_num = 0 and intv_mpfn_name = 'OFCS_Detect' and " _
                & "Aid = " & xYSWIMON0.SAAAID _
                & " and intv_s_umidl = " & xYSWIMON0.SAAUMIDL _
                & " and intv_s_umidh  = " & xYSWIMON0.SAAUMIDH
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
    Do While Not rsSIDE_DB.EOF
        V = srvrIntv_GetBuffer_ODBC(rsSIDE_DB, xrIntv)
    
         If Not IsNull(V) Then
            MsgBox "ERR02 -" & " " & xYSWIMON0.SWIMONID & " : " & "Pas d'instance, pas ds OFAC !! mnuSelect_SAAAID"
            Exit Sub
         Else   ' 1 seule enregistrement SI violation dans OFAC
            xYSWIMON0.SWIMONSTA = "S210"
            xYSWIMON0.SAAQOFAC = 1
        End If
        rsSIDE_DB.MoveNext
    Loop
    
End If  ' Fin test blnOFAC

If blnCompleted Then    ' ACK ou NAK ou REJECTED ??

    Set rsSIDE_DB = Nothing
    xSQL = "select * from rAppe where appe_iapp_name = 'SWIFT' and " _
                & "appe_inst_num = 0 and appe_network_delivery_status = 'DLV_ACKED' and " _
                & "Aid = " & xYSWIMON0.SAAAID _
                & " and appe_s_umidl = " & xYSWIMON0.SAAUMIDL _
                & " and appe_s_umidh  = " & xYSWIMON0.SAAUMIDH
                
    blnNAK = False
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
    Do While Not rsSIDE_DB.EOF
        V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)
    
         If Not IsNull(V) Then ' Positionner un flag pour lire NAK
            blnNAK = True
         Else
            xYSWIMON0.SWIMONSTA = "S901"
         End If
         rsSIDE_DB.MoveNext
    Loop
    
    If blnNAK Then
    
        Set rsSIDE_DB = Nothing
        xSQL = "select * from rAppe where appe_iapp_name = 'SWIFT' and " _
                    & "appe_inst_num = 0 and appe_network_delivery_status = 'DLV_NACKED' and " _
                    & "Aid = " & xYSWIMON0.SAAAID _
                    & " and appe_s_umidl = " & xYSWIMON0.SAAUMIDL _
                    & " and appe_s_umidh  = " & xYSWIMON0.SAAUMIDH
                    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        Do While Not rsSIDE_DB.EOF
            V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)
        
             If Not IsNull(V) Then ' Positionner un flag pour dire...
                MsgBox "ERR03 -" & " " & xYSWIMON0.SWIMONID & " : " & "COMPLETED ni ACK, ni NAK !! mnuSelect_SAAAID"
             Else
                xYSWIMON0.SWIMONSTA = "S220"
             End If
             rsSIDE_DB.MoveNext
        Loop
    End If   ' Fin test blnNAK
                        
End If  ' Fin test blnCompleted

' Mise à jour - YSWIMON0 -
xYSWIMON0.SWIMONSTAD = DSys
xYSWIMON0.SWIMONSTAH = time_Hms
W = sqlYSWIMON0_Update(xYSWIMON0, arrYSWIMON0(arrYSWIMON0_Index), cnsab)
If Not IsNull(W) Then MsgBox W, vbCritical, "MAJ YSWIMON0 au mnuSelect_SAAAID"


End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate
blnDebug = False

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SWI_MESSAGES_Aut)
Form_Init


Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
     Case "@SAA_SORTANT":    blnAuto = True
                            mnuZSWIALI0_Update_Click
                            Unload Me
    Case "@AUTO_SAA":   '$JPL 2012-04-12 blnAuto = True
                        '$JPL 2012-04-12    mnuAuto_Status_Click
                         '$JPL 2012-04-12    mnuAuto_Status_Complément_Click
                         '$JPL 2012-04-12    mnuAuto_Status_S200_Click
                            '19/03/2021 DR suppression de la fermeture de appExcelPublic
                            'If xlsManual Then
                            '    If Not appExcelPublic Is Nothing Then
                            '        appExcelPublic.Quit
                            '        Set appExcelPublic = Nothing
                            '    End If
                            'End If
                            Unload Me

   Case "@SAA_LISTES":   blnAuto = True
                            SSTab1.Tab = 1
                           Call DTPicker_Set(txtSelect_from_crea_date_time, YBIATAB0_DATE_CPT_JP0)
                           Call DTPicker_Set(txtSelect_to_crea_date_time, YBIATAB0_DATE_CPT_J)
   'Messages MP_CREATION
                            Call cbo_Scan("6 ", cboSelect_SAA)
                            fraSAA_Options.Enabled = True
                            cmdSAA_Ok_Click
                            If xlsManual Then
                                Call cmdPrint_Click_xlsManual
                            Else
                                cmdPrint_Click
                            End If
   'Messages automatiques MP_MODIFICATION
                             Call cbo_Scan("7 ", cboSelect_SAA)
                            fraSAA_Options.Enabled = True
                            cmdSAA_Ok_Click
                            If xlsManual Then
                                Call cmdPrint_Click_xlsManual
                                If Not appExcelPublic Is Nothing Then
                                    appExcelPublic.Quit
                                    Set appExcelPublic = Nothing
                                End If
                            Else
                                cmdPrint_Click
                            End If
                            Unload Me
   Case Else: blnAuto = False
End Select


End Sub

Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
SSTab1.Tab = 0
If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmSwift_Messages.param_init"
    Unload Me
Else
    lstErr.Clear
End If
If Not IsNull(paramSAA_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmSwift_Messages.paramSAA_Init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgrTextField_FormatString = fgrTextField.FormatString
fgSAA_FormatString = fgSAA.FormatString
fgZSWIFTA0_FormatString = fgZSWIFTA0.FormatString
fgZSWIFTA0_Height = fgZSWIFTA0.Height

fgZSWIHIA0_FormatString = fgZSWIHIA0.FormatString
fgStatus_FormatString = fgStatus.FormatString

mnuReprise_H.Enabled = False '''SWI_MESSAGES_Aut.Xspécial
mnuReprise_Restauration.Enabled = SWI_MESSAGES_Aut.Xspécial

mnuZSWIALI0_Update.Enabled = SWI_MESSAGES_Aut.Xspécial
fgSelect.Enabled = True
cmdReset

Me.Enabled = True
Me.MousePointer = 0
End Sub


Private Sub cboSelect_SAA_Click()

    Select Case Mid$(cboSelect_SAA, 1, 2)
        Case "1 ":  lblSelect_Utilisateur.Visible = False
                    txtSelect_Utilisateur.Visible = False
        Case "2 ":  lblSelect_Utilisateur.Visible = False
                    txtSelect_Utilisateur.Visible = False
        Case "3 ":  lblSelect_Utilisateur.Visible = False
                    txtSelect_Utilisateur.Visible = False
        Case "4 ":  lblSelect_Utilisateur.Visible = False
                    txtSelect_Utilisateur.Visible = False
        Case "5 ":  lblSelect_Utilisateur.Visible = False
                    txtSelect_Utilisateur.Visible = False
        Case "6 ":  lblSelect_Utilisateur.Visible = False
                    txtSelect_Utilisateur.Visible = False
                    
        Case "99":  txtSelect_Utilisateur = ""
                    lblSelect_Utilisateur.Visible = True
                    txtSelect_Utilisateur.Visible = True
    End Select

End Sub


Private Sub cmdImport_File_Click()
Dim X As String, xFile As String
Me.Enabled = False: Me.MousePointer = vbHourglass
xFile = Trim(txtImport_File_In)
X = Dir(xFile)
If X = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "? fichier import MsgFile non trouvé")
Else
    cmdImport_File_Exe xFile
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSaa_CB_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
    cmdSaa_CB_Exe
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdStatus_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraStatus_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Swift_cmdstatus_Ok ........"): DoEvents

fgStatus.Clear
If blnOk Or blnAuto Then
    cmdStatus_Ok.Caption = "Options"
    cmdStatus_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraStatus_Options.BackColor = &H8000000F
    Call usrColor_Container(fraStatus_Options, fraStatus_Options.BackColor)
    fraStatus_Options.Enabled = False
    cmdStatus_SQL
Else
    cmdStatus_Ok.Caption = constcmdRechercher
    cmdStatus_Ok.BackColor = &HC0FFC0
    fraStatus_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraStatus_Options, fraStatus_Options.BackColor)
    fraStatus_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Swift_cmdStatus_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdStatus_Update_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
Dim K As Long, I As Integer
Call lstErr_Clear(lstErr, cmdContext, "cmdStatus_Update : " & fgStatus.Rows - 1)

For I = 1 To fgStatus.Rows - 1
    fgStatus.Row = I
    fgStatus.Col = fgStatus_arrIndex:  K = CLng(fgStatus.Text)
    merInst_Status = arrrInst_Status(K)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "cmdStatus_Update : " & merInst_Status.inst_crea_date_time)
    mnuStatus_Actualiser_YSWIMON0
Next I

For I = 1 To arrrIntv_Flag_Nb
    merIntv_Flag = arrrIntv_Flag(arrrIntv_Flag_Nb)
    mnuStatus_Flag_YSWIMON0

Next I

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub mnuAuto_Status_Click()
Dim X As String
Dim K1 As Integer, K2 As Integer, lenText As Integer
Dim V, xSQL As String

On Error GoTo Error_Handler
V = Null

Call lstErr_Clear(lstErr, cmdContext, "> mnuAuto_Status_Click ........"): DoEvents

If IsNull(mnuStatus_Actualiser_YSWIMON0_Fct("Select -2", wAmjMin, wHmsMin)) Then
    
    Call DTPicker_Set(txtStatus_Amj, wAmjMin)
    txtStatus_Hms = wHmsMin
    '2004.10.13 anomalies : retrancher 30 sec ?
    txtStatus_Hms = Time_Sss_Hms(Time_Hms_Sss(Format$(wHmsMin, "000000")) - 30)   'wHmsMin
    fraStatus_Options.Enabled = True
    cmdStatus_Ok_Click
    If arrrInst_Status_Nb > 0 Then
        Me.Enabled = False: Me.MousePointer = vbHourglass

      cmdStatus_Update_Click
    
      'MAJ : tester si Erreur pendant MAJ ==> ? actualisation SWIMONID = -2
      '----------------------------------------------------------------------
    '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        blnTransaction_Set
      X = res_date_time & " "  ' il faut ajouter 1 espace pour timeHMS_Scan
      K1 = 0
      wAmjMax = DateJMA_Scan(X, K1)
      wHmsMax = CLng(TimeHMS_Scan(X, K1))
      Call mnuStatus_Actualiser_YSWIMON0_Fct("Update -2", wAmjMax, wHmsMax)

    If Not IsNull(V) Then
    xSQL = "Rollback"
Else
    xSQL = "Commit"
End If

Set rsSab_Update = cnsab.Execute(xSQL)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
End If
End If

Error_Handler:
Call lstErr_AddItem(lstErr, cmdContext, "< mnuAuto_Status_Click"): DoEvents
Me.Enabled = True: Me.MousePointer = 0



End Sub

Private Sub cmdZSWIFTA0_Update_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdZSWIFTA0_Update.Visible = False
cmdZSWIFTA0_Update_Transaction_Queue
fgZSWIHIA0_Reset
fgZSWIFTA0_Reset
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub fgSAA_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSAA.RowHeightMin Then
    Select Case fgSAA.Col
        Case 0: fgSAA_Sort1 = 0: fgSAA_Sort2 = 1: fgSAA_Sort
        Case 1:  fgSAA_Sort1 = 1: fgSAA_Sort2 = 1: fgSAA_Sort
        Case 2: fgSAA_Sort1 = 2: fgSAA_Sort2 = 2: fgSAA_Sort
        Case 3: fgSAA_Sort1 = 3: fgSAA_Sort2 = 3: fgSAA_Sort
        Case 4: fgSAA_Sort1 = 4: fgSAA_Sort2 = 4: fgSAA_Sort
        Case 5: fgSAA_Sort1 = 5: fgSAA_Sort2 = 5: fgSAA_SortX 5
        Case 6: fgSAA_Sort1 = 6: fgSAA_Sort2 = 6: fgSAA_Sort
        Case 7: fgSAA_Sort1 = 7: fgSAA_Sort2 = 7: fgSAA_Sort
        Case 8: fgSAA_Sort1 = 8: fgSAA_Sort2 = 8: fgSAA_Sort
        Case 9: fgSAA_Sort1 = 9: fgSAA_Sort2 = 9: fgSAA_Sort
        Case 10: fgSAA_Sort1 = 10: fgSAA_Sort2 = 10: fgSAA_Sort
        Case 11:  fgSAA_Sort1 = 11: fgSAA_Sort2 = 11: fgSAA_Sort
        Case 12: fgSAA_Sort1 = 12: fgSAA_Sort2 = 12: fgSAA_Sort
        Case 13: fgSAA_Sort1 = 13: fgSAA_Sort2 = 13: fgSAA_SortX 13
        Case 14: fgSAA_Sort1 = 14: fgSAA_Sort2 = 14: fgSAA_Sort
        Case 15: fgSAA_Sort1 = 15: fgSAA_Sort2 = 15: fgSAA_Sort
   End Select
Else
    If fgSAA.Rows > 1 Then
        Call fgSAA_Color(fgSAA_RowClick, MouseMoveUsr.BackColor, fgSAA_ColorClick)
        fgSAA.Col = fgSAA_arrIndex:  K = CLng(fgSAA.Text)
        xrMesg = arrrMesg(K)
        Call fgrTextField_Display(xrMesg.Aid, xrMesg.mesg_s_umidl, xrMesg.mesg_s_umidh)
        
        'xrIntv = arrrIntv(K)
        'srvrIntv_ElpDisplay xrIntv

   End If
End If

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
srvZSWIFTA0_Init meZSWIFTA0
xZSWIFTA0 = meZSWIFTA0
fraSelect_Options.Enabled = True
fgrTextField.Visible = False
'cmdSelect_Ok_Click
fraTab4.Visible = SWI_MESSAGES_Aut.Xspécial
cmdSaa_CB.Visible = SWI_MESSAGES_Aut.Xspécial
txtSAA_CB.Visible = SWI_MESSAGES_Aut.Xspécial

cmdZSWIFTA0_Update.Visible = False
mnuZSWIFTA0_Update_BackOffice.Enabled = SWI_MESSAGES_Aut.Swift
mnuZSWIFTA0_Update_BOTC_Jour.Enabled = SWI_MESSAGES_Aut.Swift
mnuZSWIFTA0_Update_BOTC_MT3.Enabled = SWI_MESSAGES_Aut.Swift
mnuZSWIFTA0_Update_Manuel.Enabled = SWI_MESSAGES_Aut.Swift
mnuZSWIHIA0_Reprise_YSWIMON0.Enabled = False   '''SWI_MESSAGES_Aut.Xspécial
mnuAuto_Status.Enabled = SWI_MESSAGES_Aut.Xspécial
mnuAuto_Status_S200.Enabled = SWI_MESSAGES_Aut.Xspécial
mnuAuto_Status_Complément.Enabled = SWI_MESSAGES_Aut.Xspécial

fraStatus.Enabled = SWI_MESSAGES_Aut.Xspécial
mnuSelect_S999.Enabled = SWI_MESSAGES_Aut.Saisir

cboSelect_SWIMONX32D.Enabled = False
chkSelect_SAAAID.Enabled = False
txtSelect_SWIMONFLUD.Enabled = True 'SWI_MESSAGES_Aut.Xspécial
txtSelect_SWIMONFLUD_Max.Enabled = False


'1 ère étape :  INTERDIRE Queue SI_to_SWIFT
mnuSAA_Queue_SWIFT.Enabled = SWI_MESSAGES_Aut.Xspécial ' False

blnControl = True



End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, ". BIA_Swift_Import cbo"): DoEvents

fgSelect.Visible = False

txtSelect_Utilisateur = ""
lblSelect_Utilisateur.Visible = False
txtSelect_Utilisateur.Visible = False

cboSelect_SAA.Clear
cboSelect_SAA.AddItem "1 - Intervention OFCS   "
cboSelect_SAA.AddItem "2 - En violation OFCS   "
cboSelect_SAA.AddItem "3 - Messages ACK        "
cboSelect_SAA.AddItem "4 - Messages NAK        "
cboSelect_SAA.AddItem "5 - Messages LIVE       "
cboSelect_SAA.AddItem "6 - Messages manuels -I-"
cboSelect_SAA.AddItem "7 - Messages automatiques modifiés -I-"
cboSelect_SAA.AddItem "8 - Statistiques "
cboSelect_SAA.AddItem "8x - Statistiques (export .xlsx)"
If SWI_MESSAGES_Aut.Valider Then cboSelect_SAA.AddItem "99- Interventions/utilisateur"
cboSelect_SAA.ListIndex = 4

cboSelect_IO.Clear
cboSelect_IO.AddItem "      "
cboSelect_IO.AddItem "Input "
cboSelect_IO.AddItem "Output"
cboSelect_IO.ListIndex = 0

cboSelect_Unit.Clear
cboSelect_Unit.AddItem "    "
cboSelect_Unit.AddItem "BOTC"
cboSelect_Unit.AddItem "COBK"
cboSelect_Unit.AddItem "DAFI"
cboSelect_Unit.AddItem "DCOM"
cboSelect_Unit.AddItem "None"
cboSelect_Unit.AddItem "ORPA"
cboSelect_Unit.AddItem "SCLE"
cboSelect_Unit.AddItem "SOBF"
cboSelect_Unit.AddItem "SOBI"
cboSelect_Unit.AddItem "STLX"
cboSelect_Unit.ListIndex = 0

Call rsYBIATAB0_cboK2("DEVISE", "ISO", cboSelect_SWIFTADE1)

cboSelect_SWIFTAXXX.AddItem "   "
cboSelect_SWIFTAXXX.AddItem "OON : Validé, Complet, non Supprimé" '''' TOUJOURS EN POSITION 1
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
cboSelect_SWIFTAXXX.AddItem "__N : non Supprimé"
cboSelect_SWIFTAXXX.AddItem "O__ : Validé      "
cboSelect_SWIFTAXXX.AddItem "N__ : non Validé  "
cboSelect_SWIFTAXXX.AddItem "N_N : non Validé,non Supprimé  "
cboSelect_SWIFTAXXX.AddItem "_O_ : Complet     "
cboSelect_SWIFTAXXX.AddItem "_N_ : non Complet "
cboSelect_SWIFTAXXX.AddItem "_NN : non Complet,non Supprimé "
cboSelect_SWIFTAXXX.AddItem "__O : Supprimé    "
cboSelect_SWIFTAXXX.ListIndex = 1


cboSelect_SWIFTASER.Clear
cboSelect_SWIFTASER.AddItem "      "
cboSelect_SWIFTASER.AddItem "00"
cboSelect_SWIFTASER.AddItem "TC"
cboSelect_SWIFTASER.ListIndex = 0

Call rsYBIATAB0_cboK2("DEVISE", "ISO", cboSelect_SWIMONX32D)


fgSelect.Visible = True
Call DTPicker_Set(txtSelect_SWIFTADVA, dateElp("Ouvré", 2, DSys))
Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAb_  Stock_Import"): DoEvents

fgSelect.Visible = True
Call DTPicker_Set(txtSelect_from_crea_date_time, DSys)
Call DTPicker_Set(txtSelect_to_crea_date_time, DSys)
Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAb_  Stock_Import"): DoEvents


Call DTPicker_Set(txtSelect_SWIMONFLUD, DSys)
Call DTPicker_Set(txtSelect_SWIMONFLUD_Max, DSys)

Call DTPicker_Set(txtStatus_Amj, DSys)
If IsNull(mnuStatus_Actualiser_YSWIMON0_Fct("Select -2", wAmjMin, wHmsMin)) Then
    Call DTPicker_Set(txtStatus_Amj, wAmjMin)
    txtStatus_Hms = wHmsMin
End If

Dim X As String
X = dateElp("Jour", -7, DSys)
ts_DSys_7 = "{ts '" & Mid$(X, 1, 4) & "-" & Mid$(X, 5, 2) & "-" & Mid$(X, 7, 2) & " 00:00:00.000'}"

Me.Enabled = True: Me.MousePointer = 0

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


Private Sub cmdSAA_Ok_Click()
Dim blnOk As Boolean

blnOk = fraSAA_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmddSAA_Ok ........"): DoEvents
fraSAA.Visible = False
cmdPrint.Visible = False

fgSAA.Clear
If blnOk Then
    cmdSAA_Ok.Caption = "Options"
    cmdSAA_Ok.BackColor = &HC0FFFF
    fraSAA_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSAA_Options, fraSAA_Options.BackColor)
    fraSAA_Options.Enabled = False
    Select Case Mid$(cboSelect_SAA, 1, 2)
    Case "1 ": cmdSAA_SQL_01
    Case "2 ": cmdSAA_SQL_02
    Case "3 ": cmdSAA_SQL_03
    Case "4 ": cmdSAA_SQL_04
    Case "5 ": cmdSAA_SQL_05
    Case "6 ": Call cbo_Scan("Input", cboSelect_IO)
               cmdSAA_SQL_06
               cmdPrint.Visible = True
    Case "7 ": Call cbo_Scan("Input", cboSelect_IO)
               cmdSAA_SQL_07
               cmdPrint.Visible = True
    Case "8 ": cmdSAA_SQL_08
               cmdPrint.Visible = True
    Case "8x":  SAA_Statistiques_Export
               cmdPrint.Visible = True

               cmdPrint.Visible = True
    Case "99": cmdSAA_SQL_99
             
    Case Else: MsgBox "Non programmé", vbCritical, Me.Name & " : " & cboSelect_SAA
                Exit Sub
    End Select

    fraSAA.Visible = True
Else
    cmdSAA_Ok.Caption = constcmdRechercher
    cmdSAA_Ok.BackColor = &HC0FFC0
    fraSAA_Options.BackColor = &HC0FFFF
    Call usrColor_Container(fraSAA_Options, fraSAA_Options.BackColor)
    fraSAA_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< ScmddSAA_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdZSWIFTA0_Ok_Click()
Dim blnOk As Boolean

blnOk = fraZSWIFTA0_Options.Enabled

cmdZSWIFTA0_Update.Visible = False
ReDim arrZSWIFTA0_SAA_Queue(0)     '!!!!!!!!!!!!!!! utilisé UNIQUEMENT SI UPDATE
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Swift_cmdZSWIFTA0_Ok ........"): DoEvents

fraZSWIFTA0.Visible = False
fgZSWIFTA0.Clear
fgZSWIHIA0.Clear
If blnOk Then
    cmdZSWIFTA0_Ok.Caption = "Options"
    cmdZSWIFTA0_Ok.BackColor = &HC0FFFF '&HFFFFFA   '
    fraZSWIFTA0_Options.BackColor = &H8000000F
    Call usrColor_Container(fraZSWIFTA0_Options, fraZSWIFTA0_Options.BackColor)
    fraZSWIFTA0_Options.Enabled = False
    cmdZSWIFTA0_SQL
    fraZSWIFTA0.Visible = True

Else
    cmdZSWIFTA0_Ok.Caption = constcmdRechercher
    cmdZSWIFTA0_Ok.BackColor = &HC0FFC0
    fraZSWIFTA0_Options.BackColor = &HC0FFFF
    Call usrColor_Container(fraZSWIFTA0_Options, fraZSWIFTA0_Options.BackColor)
    fraZSWIFTA0_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Swift_cmdZSWIFTA0_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long

On Error Resume Next
If y <= fgStatus.RowHeightMin Then
'
Else
    If fgStatus.Rows > 1 Then
        Call fgStatus_Color(fgStatus_RowClick, MouseMoveUsr.BackColor, fgStatus_ColorClick)
        fgStatus.Col = fgStatus_arrIndex:  arrrInst_Status_Index = CLng(fgStatus.Text)
        fgStatus.LeftCol = 0
        merInst_Status = arrrInst_Status(arrrInst_Status_Index)
        
       '''Me.PopupMenu mnuStatus, vbPopupMenuLeftButton

   End If
End If
fgStatus.LeftCol = 0

End Sub


Private Sub fgZSWI_D_Click()
lstW.Visible = Not lstW.Visible
End Sub

Private Sub fgZSWIHIA0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long, xWhere As String
On Error Resume Next
If y <= fgZSWIHIA0.RowHeightMin Then
    Select Case fgZSWIHIA0.Col
        Case 0: fgZSWIHIA0_Sort1 = 0: fgZSWIHIA0_Sort2 = 1: fgZSWIHIA0_Sort
        Case 1:  fgZSWIHIA0_Sort1 = 1: fgZSWIHIA0_Sort2 = 1: fgZSWIHIA0_Sort
        Case 2: fgZSWIHIA0_Sort1 = 2: fgZSWIHIA0_Sort2 = 2: fgZSWIHIA0_Sort
        Case 3: fgZSWIHIA0_Sort1 = 3: fgZSWIHIA0_Sort2 = 3: fgZSWIHIA0_SortX 3
        Case 4: fgZSWIHIA0_Sort1 = 4: fgZSWIHIA0_Sort2 = 4: fgZSWIHIA0_Sort
        Case 5: fgZSWIHIA0_Sort1 = 5: fgZSWIHIA0_Sort2 = 5: fgZSWIHIA0_SortX 5
        Case 6: fgZSWIHIA0_Sort1 = 6: fgZSWIHIA0_Sort2 = 6: fgZSWIHIA0_Sort
        Case 7: fgZSWIHIA0_Sort1 = 7: fgZSWIHIA0_Sort2 = 7: fgZSWIHIA0_Sort
        Case 8: fgZSWIHIA0_Sort1 = 8: fgZSWIHIA0_Sort2 = 8: fgZSWIHIA0_Sort
        Case 9: fgZSWIHIA0_Sort1 = 9: fgZSWIHIA0_Sort2 = 9: fgZSWIHIA0_SortX 9
        Case 10: fgZSWIHIA0_Sort1 = 10: fgZSWIHIA0_Sort2 = 10: fgZSWIHIA0_SortX 10
    End Select
Else
    If fgZSWIHIA0.Rows > 1 Then
        Call fgZSWIHIA0_Color(fgZSWIHIA0_RowClick, MouseMoveUsr.BackColor, fgZSWIHIA0_ColorClick)
        fgZSWIHIA0.Col = fgZSWIHIA0_arrIndex:  K = CLng(fgZSWIHIA0.Text)
        fgZSWIHIA0.LeftCol = 0
        xZSWIHIA0 = arrZSWIHIA0(K)
        If SWI_MESSAGES_Aut.Xspécial Then
            Me.PopupMenu mnuReprise, vbPopupMenuLeftButton
        Else
            fgZSWI_D.Clear
            fgZSWI_D.Visible = True
            xWhere = " where SWIHIBNUM = " & xZSWIHIA0.SWIHIANUM
            srvZSWIHIA0_fgDisplay xZSWIHIA0, fgZSWI_D
            arrZSWIHIB0_SQL xWhere
            arrZSWIHIB0_lstW
       End If

   End If
End If

End Sub


Private Sub mnuAuto_Status_Complément_Click()

' Mise à jour des dossiers à partir de YSWIMON0  (pb synchronisation sporadique entre rIntv et rInst)
'====================================================================================================
Dim V, xSQL As String
Dim I As Integer

On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
V = Null

Call lstErr_Clear(lstErr, cmdContext, "> mnuAuto_Status_Complément ........"): DoEvents

' rechercher les messages dont le statut est < 'S900'
'====================================================
chkSelect_SWIMONSTA = "1"
cmdSelect_SQL
If arrYSWIMON0_Nb > 0 Then
    ReDim arrrInst_Status(arrYSWIMON0_Nb)
    arrrInst_Status_Max = arrYSWIMON0_Nb: arrrInst_Status_Nb = 0
    

' Lecture de l'instance dans SIDE_DB
'====================================================
    For I = 1 To arrYSWIMON0_Nb
         
        xYSWIMON0 = arrYSWIMON0(I)
        If xYSWIMON0.SWIMONID > 0 Then
                Set rsSIDE_DB = Nothing
                
                xSQL = "select * from rInst where Aid = " & xYSWIMON0.SAAAID _
                        & " and inst_s_umidl = " & xYSWIMON0.SAAUMIDL _
                        & " and inst_s_umidh = " & xYSWIMON0.SAAUMIDH _
                        & " and inst_num = 0"
                        
                Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
                
                Do While Not rsSIDE_DB.EOF
                    V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
                
                     If Not IsNull(V) Then
                         MsgBox V, vbCritical, "mnuAuto_Status_Complément_Click : rInst"
                         ''Exit Sub ''20070530 JPL
                     Else
                         arrrInst_Status_Nb = arrrInst_Status_Nb + 1
                         If arrrInst_Status_Nb > arrrInst_Status_Max Then
                             arrrInst_Status_Max = arrrInst_Status_Max + 50
                             ReDim Preserve arrrInst_Status(arrrInst_Status_Max)
                         End If
                         
                         arrrInst_Status(arrrInst_Status_Nb) = xrInst
                    End If
                    rsSIDE_DB.MoveNext
                
                Loop
        End If
    Next I

    fgStatus_Display
' Mise à jour YSWIMON0.STATUS
'====================================================

    cmdStatus_Update_Click
       
End If

Error_Handler:
SSTab1.Tab = 0
fgSelect_Reset
Call lstErr_AddItem(lstErr, cmdContext, "< mnuAuto_Status_Complément: FIN"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuAuto_Status_S200_Click()
'====================================================================================================
' envoi EMail des dossiers envoyés par SAB non intégrés dans ALLIANCE après 5 minutes
'====================================================================================================
Dim V, xSQL As String
Dim I As Integer, SSS As Long
Dim Nb As Long
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
V = Null

Call lstErr_Clear(lstErr, cmdContext, "> mnuAuto_Status_S200 ........"): DoEvents

' rechercher les messages dont le statut est = 'S200' depuis 12 minutes
'====================================================
SSS = Time_Hms_Sss(time_Hms) - 720


xSQL = " where SWIMONSTA = 'S200' and SWIMONFLUH <= " & Time_Sss_Hms(SSS)
arrYSWIMON0_SQL xSQL

If arrYSWIMON0_Nb > 0 Then
'====================================================
    For I = 1 To arrYSWIMON0_Nb
         
        oldYSWIMON0_Status = arrYSWIMON0(I)
        If xYSWIMON0.SWIMONID > 0 Then
        '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                meYSWIMON0_Status = oldYSWIMON0_Status
                meYSWIMON0_Status.SWIMONSTA = "S201"
                blnTransaction_Set
                
                V = sqlYSWIMON0_Update(meYSWIMON0_Status, oldYSWIMON0_Status, cnsab)
                
                If Not IsNull(V) Then
                    xSQL = "Rollback"
                Else
                    xSQL = "Commit"
                End If
                
                Set rsSab_Update = cnsab.Execute(xSQL, Nb)
        '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        End If
    Next I

      
End If

Error_Handler:
SSTab1.Tab = 0
Call lstErr_AddItem(lstErr, cmdContext, "< mnuAuto_Status_S200: FIN"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuReprise_H_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdZSWIHIA0_Insert
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuReprise_Restauration_Click()
X = MsgBox("!!! Confirmez-vous la restauration de ce message pour l'émettre à nouveau dans SAB ?", vbQuestion + vbYesNo, "BIA_SWIFT : mnuReprise_Restauration_Click")
If X = vbYes Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    cmdZSWIHIA0_Restauration
    Me.Enabled = True: Me.MousePointer = 0
End If

End Sub

Private Sub mnuSAA_Queue_Autorisation_Click()
fgZSWIFTA0_Update_SAA_Queue paramSAA_Queue_Autorisation
End Sub

Private Sub fgZSWIFTA0_Update_SAA_Queue(lSAA_Queue As String)
Dim X As String
X = lSAA_Queue
If arrZSWIFTA0(arrZSWIFTA0_Index).SWIFTAETA <> 1 Then X = paramSAA_Queue_TRF_en_Cours

arrZSWIFTA0_SAA_Queue(arrZSWIFTA0_Index) = X
fgZSWIFTA0.Col = 10: fgZSWIFTA0.Text = X
Select Case X
    Case "":     Call fgZSWIFTA0_Color(0, fgZSWIFTA0.BackColor, fgZSWIFTA0.BackColor)
    Case paramSAA_Queue_TRF_en_Cours:     Call fgZSWIFTA0_Color(0, vbMagenta, fgZSWIFTA0.BackColor)
    Case Else:   Call fgZSWIFTA0_Color(0, cmdZSWIFTA0_Update.BackColor, fgZSWIFTA0.BackColor)
End Select
End Sub

Private Sub mnuSAA_Queue_Modification_Click()
fgZSWIFTA0_Update_SAA_Queue paramSAA_Queue_Modification
End Sub


Private Sub mnuSAA_Queue_SWIFT_Click()
fgZSWIFTA0_Update_SAA_Queue paramSAA_Queue_SWIFT

End Sub


Private Sub mnuSelect_SAAAID_Click()

' arrYSWIMON0(arrYSWIMON0_Index) = Old record prefix / xYSWIMON0 = New record prefix

If xYSWIMON0.SWIMONSTA Like "S9*" Then MsgBox "Statut final -" & xYSWIMON0.SWIMONSTA & "- pour le message. Plus de traitement ": Exit Sub

mnuSelect_SAAAID_Evolution

End Sub

Private Sub mnuSelect_S999_Click()
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
meYSWIMON0_Status = xYSWIMON0
meYSWIMON0_Status.SWIMONSTA = "S999"
meYSWIMON0_Status.SWIMONSTAD = DSys
meYSWIMON0_Status.SWIMONSTAH = time_Hms

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
blnTransaction_Set

V = sqlYSWIMON0_Update(meYSWIMON0_Status, xYSWIMON0, cnsab)

If Not IsNull(V) Then
    xSQL = "Rollback"
Else
    xSQL = "Commit"
    xYSWIMON0 = meYSWIMON0_Status
End If

Set rsSab_Update = cnsab.Execute(xSQL)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
fgSelect_DisplayLine arrYSWIMON0_Index
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuStatus_Actualiser_Click()
'mnuStatus_Actualiser_YSWIMON0
End Sub

Private Sub mnuZSWIALI0_Update_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass

arrZSWIALI0_SQL " where SWIALIETA > 0 "
If arrZSWIALI0_Nb > 0 Then
    'If paramEnvironnement = constProduction Then
        cmdZSWIALI0_Update_Transaction "SW", paramSAA_Queue_SWIFT
    'Else
    '    cmdZSWIALI0_Update_Transaction "SW", paramSAA_Queue_SWIFT
    'End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuZSWIFTA0_Update_BackOffice_Click()
fraZSWIFTA0_Options_Reset
cbo_Scan "00", cboSelect_SWIFTASER
cmdZSWIFTA0_Ok_Click
If arrZSWIFTA0_Nb > 0 Then fgZSWIFTA0_Update_Enabled "" 'paramSAA_Queue_SWIFT 'Autorisation

End Sub

Private Sub fraZSWIFTA0_Options_Reset()
SSTab1.Tab = 2
fraZSWIFTA0_Options.Enabled = True
cbo_Scan "OON", cboSelect_SWIFTAXXX
cbo_Scan "  ", cboSelect_SWIFTASER
cbo_Scan "   ", cboSelect_SWIFTADE1
txtSelect_SWIFTAREF = ""
txtSelect_SWIFTAMES = ""
chkSelect_ZSWIHIA0.value = "0"
Call DTPicker_Set(txtSelect_SWIFTADVA, dateElp("Ouvré", 2, DSys))

End Sub

Private Sub mnuZSWIFTA0_Update_BOTC_Jour_Click()
fraZSWIFTA0_Options_Reset
cbo_Scan "TC", cboSelect_SWIFTASER
Call DTPicker_Set(txtSelect_SWIFTADVA, DSys)
cmdZSWIFTA0_Ok_Click
If arrZSWIFTA0_Nb > 0 Then fgZSWIFTA0_Update_Enabled paramSAA_Queue_SWIFT 'Autorisation
End Sub

Private Sub mnuZSWIFTA0_Update_BOTC_MT3_Click()
fraZSWIFTA0_Options_Reset
cbo_Scan "TC", cboSelect_SWIFTASER
txtSelect_SWIFTAMES = "3"

cmdZSWIFTA0_Ok_Click
If arrZSWIFTA0_Nb > 0 Then fgZSWIFTA0_Update_Enabled paramSAA_Queue_SWIFT 'Autorisation
End Sub

Private Sub mnuZSWIFTA0_Update_Manuel_48H_Click()
fraZSWIFTA0_Options_Reset
cmdZSWIFTA0_Ok_Click
'''If arrZSWIFTA0_Nb > 0 Then cmdZSWIFTA0_Update.Visible = SWI_MESSAGES_Aut.Swift
If arrZSWIFTA0_Nb > 0 Then fgZSWIFTA0_Update_Enabled "" 'paramSAA_Queue_SWIFT 'Autorisation

End Sub

Private Sub mnuZSWIFTA0_Update_Manuel_Click()
fraZSWIFTA0_Options_Reset
Call DTPicker_Set(txtSelect_SWIFTADVA, "29991231")

cmdZSWIFTA0_Ok_Click
''If arrZSWIFTA0_Nb > 0 Then cmdZSWIFTA0_Update.Visible = SWI_MESSAGES_Aut.Swift
If arrZSWIFTA0_Nb > 0 Then fgZSWIFTA0_Update_Enabled "" 'paramSAA_Queue_SWIFT 'Autorisation

End Sub

Private Sub mnuZSWIHIA0_Display_Click()
Dim xWhere As String, wAmj7 As Long
Call DTPicker_Amj7(txtSelect_SWIFTADVA, wAmj7)
xWhere = " where SWIHIADEN = " & wAmj7 & " "

arrZSWIHIA0_SQL xWhere
fgZSWIHIA0_Display

End Sub

Private Sub mnuZSWIHIA0_Reprise_YSWIMON0_Click()
Dim xSQL As String
Dim wSWIMONID As Long, wAmj7 As Long

Me.Enabled = False: Me.MousePointer = vbHourglass

xSQL = "Select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" & " where SWIMONID= -1"
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    MsgBox "MAnque YSWIMON0.SWIMONID = -1  : ", vbCritical, "frmSwift_Messages.mnuZSWIHIA0_Reprise_YSWIMON0_Click"
    Exit Sub
End If
wSWIMONID = rsSab("SWISABNUM")
xSQL = "Select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" & " where SWIMONID = " & wSWIMONID
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    Call srvYSWIMON0_GetBuffer_ODBC(rsSab, xYSWIMON0)
    wAmj7 = xYSWIMON0.SWIMONFLUD - 19000000
    xSQL = " where SWIHIADEN > " & wAmj7 & " or ( SWIHIADEN = " & wAmj7 & " and  SWIHIAHEN > " & xYSWIMON0.SWIMONFLUH & " )"
    arrZSWIHIA0_SQL xSQL
    fgZSWIHIA0_Display
    cmdZSWIHIA0_Insert_Auto
Else
    MsgBox "MAnque YSWIMON0.SWIMONID = " & wSWIMONID, vbCritical, "frmSwift_Messages.mnuZSWIHIA0_Reprise_YSWIMON0_Click"
    Exit Sub
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub







Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 1:
            Select Case Mid$(cboSelect_SAA, 1, 2)
                Case "6 ": cmdPrint_List6_Ok
                Case "7 ": cmdPrint_List6_Ok
                Case "8 ": cmdPrint_List8_Ok
                
               ' Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
            End Select
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdYSWIMON0_SQL"
If chkSelect_SWIMONSTA = "1" Then
    xWhere = " where SWIMONSTA < 'S900'"
Else
    Call DTPicker_Control(txtSelect_SWIMONFLUD, wAmjMin)
    Call DTPicker_Control(txtSelect_SWIMONFLUD_Max, wAmjMax)
    xWhere = " where SWIMONFLUD >= " & wAmjMin & " and SWIMONFLUD <= " & wAmjMax
    
    If chkSelect_SAAAID = "1" Then xWhere = xWhere & " and SAAAID = 0"
End If
arrYSWIMON0_SQL xWhere


fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdStatus_SQL()
Dim V
Dim X As String
Dim xWhere As String
Dim wHms As Long
On Error GoTo Error_Handler

currentAction = "cmdStatus_SQL"
Call DTPicker_Control(txtStatus_Amj, wAmjMin)
wHms = time_N6(txtStatus_Hms)

xWhere = " where intv_inst_num = 0  and intv_date_time >= " & SQL_Date_Time(wAmjMin, wHms)

arrrInst_Status_SQL xWhere

fgStatus_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 3
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSAA_SQL_Modele()
Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer

ReDim arrrIntv(101)
arrrIntv_Max = 100: arrrIntv_Nb = 0

blnOk = False
Set rsSIDE_DB = Nothing

xSQL = "select * from rintv where intv_inty_category = 'INTY_ROUTING' and" _
       & "intv_inst_num = 0 and intv_mpfn_name = 'OFCS_Detect' and month(intv_date_time) = 6"

Select Case Mid$(cboSelect_SAA, 1, 2)
   Case "1 ": xSQL = "select * from rintv where intv_inty_category = 'INTY_ROUTING' and" _
       & "intv_inst_num = 0 and intv_mpfn_name = 'OFCS_Detect' and month(intv_date_time) = 6"
   Case Else: MsgBox "Non programmé", vbCritical, Me.Name & " : " & cboSelect_SAA
               Exit Sub
End Select



Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrIntv_GetBuffer_ODBC(rsSIDE_DB, xrIntv)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
        Exit Sub
     Else
         arrrIntv_Nb = arrrIntv_Nb + 1
         If arrrIntv_Nb = arrrIntv_Max Then   '>
            If arrrIntv_Max >= 300 Then
                MsgBox "300 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
   
                arrrIntv_Max = arrrIntv_Max + 50
                ReDim Preserve arrrIntv(arrrIntv_Max + 1)
            End If
         End If
         
         arrrIntv(arrrIntv_Nb) = xrIntv
    End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rIntv : " & arrrIntv_Nb): DoEvents

' lecture rMesg
'================
arrrMesg_Max = arrrIntv_Max: arrrMesg_Nb = arrrIntv_Nb
ReDim arrrMesg(arrrMesg_Max)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rMesg " _
        & "where Aid = " & xrIntv.Aid _
        & " and mesg_s_umidl = " & xrIntv.intv_s_umidl _
        & " and mesg_s_umidh  = " & xrIntv.intv_s_umidh
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rMesg"
              '  Exit Sub
            Else
                arrrMesg(I) = xrMesg
            End If
        End If
    
Next I
'=========

fgSAA_Display

End Sub

Private Sub cmdSAA_SQL_05()     ' ***** MESSAGES  LIVE
Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long

' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

ReDim arrrMesg(101)
arrrMesg_Max = 100: arrrMesg_Nb = 0

blnOk = False
Set rsSIDE_DB = Nothing

If Mid$(cboSelect_IO, 1, 1) = " " Then
    xSQL = "select * from rMesg where mesg_status = 'LIVE' and " _
               & "mesg_frmt_name = 'Swift' and " _
               & "mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
               & "mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "
Else
    xSQL = "select * from rMesg where mesg_status = 'LIVE' and " _
               & "mesg_frmt_name = 'Swift' and substring(mesg_sub_format, 1, 1) = '" & Mid$(cboSelect_IO, 1, 1) & "' and " _
               & "mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
               & "mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "
End If

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
        Exit Sub
     Else
         arrrMesg_Nb = arrrMesg_Nb + 1
         If arrrMesg_Nb = arrrMesg_Max Then   '>
            If arrrMesg_Max >= 10000 Then
                MsgBox "10000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
   
                arrrMesg_Max = arrrMesg_Max + 50
                ReDim Preserve arrrMesg(arrrMesg_Max + 1)
            End If
         End If
         
         arrrMesg(arrrMesg_Nb) = xrMesg
    End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rMesg : " & arrrMesg_Nb): DoEvents


' lecture rInst
'================
arrrInst_Max = arrrMesg_Max: arrrInst_Nb = arrrMesg_Nb
ReDim arrrInst(arrrInst_Max)

For I = 1 To arrrMesg_Nb
    xrMesg = arrrMesg(I)
    xSQL = "select * from rInst " _
        & "where Aid = " & xrMesg.Aid _
        & " and inst_s_umidl = " & xrMesg.mesg_s_umidl _
        & " and inst_s_umidh  = " & xrMesg.mesg_s_umidh _
        & " and inst_num  = 0"
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rInst"
            Else
                arrrInst(I) = xrInst
            End If
        End If
    
Next I

' lecture rAppe
'================
arrrAppe_Max = arrrMesg_Max: arrrAppe_Nb = arrrMesg_Nb
ReDim arrrAppe(arrrAppe_Max)
ReDim arrrAppe_E(arrrAppe_Max)
ReDim arrrAppe_R(arrrAppe_Max)

srvrAppe_Init arrrAppe_E(0)

For I = 1 To arrrMesg_Nb
    xrMesg = arrrMesg(I)
    xSQL = "select * from rAppe " _
        & "where Aid = " & xrMesg.Aid _
        & " and appe_s_umidl = " & xrMesg.mesg_s_umidl _
        & " and appe_s_umidh  = " & xrMesg.mesg_s_umidh _
        & " and appe_inst_num  = 0"
    
    arrrAppe_E(I) = arrrAppe_E(0)
    arrrAppe_R(I) = arrrAppe_E(0)
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
  
    For Boucle = 1 To 2
        If Not rsSIDE_DB.EOF Then
            V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rAppe"
            Else
                If Boucle = 1 Then
                    arrrAppe(I) = xrAppe
                End If
               ' Debug.Print I; Boucle; xrAppe.appe_type
                If xrAppe.appe_type = "APPE_EMISSION" Then
                    arrrAppe_E(I) = xrAppe
                Else
                    arrrAppe_R(I) = xrAppe
                End If
            End If
            rsSIDE_DB.MoveNext
        End If
    Next Boucle

Next I

' lecture rIntv : CHARGER SEULEMENT le nbr utilisé dans fgSAA_Display
'================
arrrIntv_Max = arrrMesg_Max: arrrIntv_Nb = arrrMesg_Nb
ReDim arrrIntv(arrrIntv_Max)

srvrIntv_Init arrrIntv(0)

For I = 1 To arrrMesg_Nb
    arrrIntv(I) = arrrIntv(0)
Next I


fgSAA_Display

End Sub

Private Sub cmdSAA_SQL_06()     ' ***** MESSAGES CREES MANUELLEMENT DANS ALLIANCE - Période par rapport à la date de création des messages

Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long

On Error GoTo Error_Handler
' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

ReDim arrrInst(501)
arrrInst_Max = 500: arrrInst_Nb = 0

blnOk = False
Set rsSIDE_DB = Nothing

' Par défaut, cette requête n'est valable que pour les Input SAA
' Mid$(Me.cboSelect_IO, 1, 6) = "Input "    ' Préparer l'affichage dans fgSAA_DisplayLine

'xSql = "select * from rInst where inst_num = 0 and inst_crea_rp_name = '_MP_creation' and " _
'           & "inst_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
'           & "inst_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "


'$JPL 2013-07-22
'xSql = "select * from rInst where inst_num = 0 and " _
'           & "inst_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
'           & "inst_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'}  and inst_crea_rp_name = '_MP_creation'"

xSQL = "select * from rMesg , rInst " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'}" _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'}" _
          & " and substring(mesg_uumid, 1, 1) = 'I'" _
          & " and  rInst.Aid =  rMesg.Aid" _
          & " and inst_s_umidl = mesg_s_umidl" _
          & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0" _
          & " and inst_crea_rp_name = '_MP_creation'"



Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
'If arrrInst_Nb > 59 Then
'    Debug.Print arrrInst_Nb
'End If
    
    V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
        Exit Sub
     Else
         arrrInst_Nb = arrrInst_Nb + 1
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "rInst : " & arrrInst_Nb): DoEvents

         If arrrInst_Nb = arrrInst_Max Then   '>
            If arrrInst_Max >= 10000 Then
                MsgBox "10000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
   
                arrrInst_Max = arrrInst_Max + 500
                ReDim Preserve arrrInst(arrrInst_Max + 1)
            End If
         End If
         
         arrrInst(arrrInst_Nb) = xrInst
    End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rInst : " & arrrInst_Nb): DoEvents


' lecture rMesg
'================
arrrMesg_Max = arrrInst_Max: arrrMesg_Nb = arrrInst_Nb
ReDim arrrMesg(arrrInst_Max)

For I = 1 To arrrInst_Nb
    xrInst = arrrInst(I)
    xSQL = "select * from rMesg " _
        & "where Aid = " & xrInst.Aid _
        & " and mesg_s_umidl = " & xrInst.inst_s_umidl _
        & " and mesg_s_umidh = " & xrInst.inst_s_umidh
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rMesg"
            Else
                arrrMesg(I) = xrMesg
            End If
        End If
    
Next I

' lecture rAppe
'================
arrrAppe_Max = arrrInst_Max: arrrAppe_Nb = arrrInst_Nb
ReDim arrrAppe(arrrAppe_Max)
ReDim arrrAppe_E(arrrAppe_Max)
ReDim arrrAppe_R(arrrAppe_Max)

srvrAppe_Init arrrAppe_E(0)

For I = 1 To arrrInst_Nb
    xrInst = arrrInst(I)
    xSQL = "select * from rAppe " _
        & "where Aid = " & xrInst.Aid _
        & " and appe_s_umidl = " & xrInst.inst_s_umidl _
        & " and appe_s_umidh  = " & xrInst.inst_s_umidh _
        & " and appe_inst_num  = 0"
    
    arrrAppe_E(I) = arrrAppe_E(0)
    arrrAppe_R(I) = arrrAppe_E(0)
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
  
    For Boucle = 1 To 2
        If Not rsSIDE_DB.EOF Then
            V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rAppe"
            Else
                If Boucle = 1 Then
                    arrrAppe(I) = xrAppe
                End If
               ' Debug.Print I; Boucle; xrAppe.appe_type
                If xrAppe.appe_type = "APPE_EMISSION" Then
                    arrrAppe_E(I) = xrAppe
                Else
                    arrrAppe_R(I) = xrAppe
                End If
            End If
            rsSIDE_DB.MoveNext
        End If
    Next Boucle

Next I

' lecture rIntv : CHARGER SEULEMENT le nbr utilisé dans fgSAA_Display
'================
arrrIntv_Max = arrrInst_Max: arrrIntv_Nb = arrrInst_Nb
ReDim arrrIntv(arrrIntv_Max)

srvrIntv_Init arrrIntv(0)

For I = 1 To arrrInst_Nb
    arrrIntv(I) = arrrIntv(0)
Next I


fgSAA_Display
Exit Sub

Error_Handler:
    Call lstErr_AddItem(lstErr, cmdContext, "rInst : " & arrrInst_Nb & Error): DoEvents
    
    '''Wait_SS 1
    Call lstErr_AddItem(lstErr, cmdContext, "temporisation"): DoEvents
   ''' If InStr(Error, "Erreur Automation") > 0 Then Resume 0
   ''' Resume 0
End Sub


Private Sub cmdSAA_SQL_07()     ' ***** MESSAGES AUTOMATIQUES MODIFIES DANS ALLIANCE - Période par rapport à la date de création des messages

Dim xSQL As String, xWhere As String, xAnd As String, W As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long

On Error GoTo Error_Handler
' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

ReDim arrrInst(5001)
arrrInst_Max = 5000: arrrInst_Nb = 0

blnOk = False
Set rsSIDE_DB = Nothing

' Par défaut, cette requête n'est valable que pour les Input SAA
' Mid$(Me.cboSelect_IO, 1, 6) = "Input "    - Préparer l'affichage dans fgSAA_DisplayLine
'$JPL 2013-07-22
'xSql = "select * from rInst where inst_num = 0 and inst_crea_rp_name = '_AI_from_APPLI' and " _
'           & "inst_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
'           & "inst_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "
xSQL = "select * from rMesg , rInst " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'}" _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'}" _
          & " and substring(mesg_uumid, 1, 1) = 'I'" _
          & " and  rInst.Aid =  rMesg.Aid" _
          & " and inst_s_umidl = mesg_s_umidl" _
          & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0" _
          & " and inst_crea_rp_name = '_AI_from_APPLI'"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)

    If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
         Exit Sub
    Else
         arrrInst_Nb = arrrInst_Nb + 1
         If arrrInst_Nb = arrrInst_Max Then   '>
            If arrrInst_Max >= 10000 Then
                MsgBox "20000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
   
                arrrInst_Max = arrrInst_Max + 500
                ReDim Preserve arrrInst(arrrInst_Max + 1)
            End If
         End If
         
         arrrInst(arrrInst_Nb) = xrInst
    End If
    DoEvents

    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rInst : " & arrrInst_Nb): DoEvents


' lecture rMesg
'================
arrrMesg_Max = arrrInst_Max: arrrMesg_Nb = arrrInst_Nb
ReDim arrrMesg(arrrInst_Max)

For I = 1 To arrrInst_Nb
    xrInst = arrrInst(I)
    xSQL = "select * from rMesg " _
        & "where Aid = " & xrInst.Aid _
        & " and mesg_s_umidl = " & xrInst.inst_s_umidl _
        & " and mesg_s_umidh = " & xrInst.inst_s_umidh

        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rMesg"
            Else
                arrrMesg(I) = xrMesg
            End If
        End If

Next I

' lecture rAppe
'================
arrrAppe_Max = arrrInst_Max: arrrAppe_Nb = arrrInst_Nb
ReDim arrrAppe(arrrAppe_Max)
ReDim arrrAppe_E(arrrAppe_Max)
ReDim arrrAppe_R(arrrAppe_Max)

srvrAppe_Init arrrAppe_E(0)

For I = 1 To arrrInst_Nb
    xrInst = arrrInst(I)
    xSQL = "select * from rAppe " _
        & "where Aid = " & xrInst.Aid _
        & " and appe_s_umidl = " & xrInst.inst_s_umidl _
        & " and appe_s_umidh  = " & xrInst.inst_s_umidh _
        & " and appe_inst_num  = 0"
    
    arrrAppe_E(I) = arrrAppe_E(0)
    arrrAppe_R(I) = arrrAppe_E(0)
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
  
    For Boucle = 1 To 2
        If Not rsSIDE_DB.EOF Then
            V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rAppe"
            Else
                If Boucle = 1 Then
                    arrrAppe(I) = xrAppe
                End If
               ' Debug.Print I; Boucle; xrAppe.appe_type
                If xrAppe.appe_type = "APPE_EMISSION" Then
                    arrrAppe_E(I) = xrAppe
                Else
                    arrrAppe_R(I) = xrAppe
                End If
            End If
            rsSIDE_DB.MoveNext
        End If
    Next Boucle

Next I

' lecture rIntv : CHARGER SEULEMENT le nbr utilisé dans fgSAA_Display
'================
arrrIntv_Max = arrrInst_Max: arrrIntv_Nb = arrrInst_Nb
ReDim arrrIntv(arrrIntv_Max)

srvrIntv_Init arrrIntv(0)

For I = 1 To arrrInst_Nb
    arrrIntv(I) = arrrIntv(0)
Next I


fgSAA_Display
Exit Sub

Error_Handler:
    Call lstErr_AddItem(lstErr, cmdContext, "rInst : " & arrrInst_Nb & Error): DoEvents
    
  '''  Wait_SS 1
    Call lstErr_AddItem(lstErr, cmdContext, "temporisation"): DoEvents
  '''  If InStr(Error, "Erreur Automation") > 0 Then Resume 0
  '''  Resume 0

End Sub


Private Sub cmdSAA_SQL_99()     ' ***** INTERVENTIONS PAR UTILISATEUR

Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long

' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

' lecture rIntv : Interventions par utilisateur
'===============================================

ReDim arrrIntv(101)
arrrIntv_Max = 100: arrrIntv_Nb = 0

blnOk = False
Set rsSIDE_DB = Nothing

xSQL = "select * from rIntv where intv_inst_num = 0 and " _
           & "intv_oper_nickname = '" & Trim(txtSelect_Utilisateur) & "' and " _
           & "intv_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
           & "intv_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
           & "order by intv_date_time "
           
Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrIntv_GetBuffer_ODBC(rsSIDE_DB, xrIntv)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL_99"
        Exit Sub
     Else
         arrrIntv_Nb = arrrIntv_Nb + 1
         If arrrIntv_Nb = arrrIntv_Max Then   '>
            If arrrIntv_Max >= 10000 Then
                MsgBox "10000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
                arrrIntv_Max = arrrIntv_Max + 50
                ReDim Preserve arrrIntv(arrrIntv_Max + 1)
            End If
         End If
         
         arrrIntv(arrrIntv_Nb) = xrIntv
    End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rIntv : " & arrrIntv_Nb): DoEvents

' Lecture rMesg
'================

arrrMesg_Max = arrrIntv_Max: arrrMesg_Nb = arrrIntv_Nb
ReDim arrrMesg(arrrMesg_Max)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rMesg " _
        & "where Aid = " & xrIntv.Aid _
        & " and mesg_s_umidl = " & xrIntv.intv_s_umidl _
        & " and mesg_s_umidh  = " & xrIntv.intv_s_umidh
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rMesg"
            Else
                arrrMesg(I) = xrMesg
            End If
        End If
Next I

' lecture rInst
'================
arrrInst_Max = arrrIntv_Max: arrrInst_Nb = arrrIntv_Nb
ReDim arrrInst(arrrInst_Max)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rInst " _
        & "where Aid = " & xrMesg.Aid _
        & " and inst_s_umidl = " & xrIntv.intv_s_umidl _
        & " and inst_s_umidh  = " & xrIntv.intv_s_umidh _
        & " and inst_num  = 0"
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rInst"
            Else
                arrrInst(I) = xrInst
            End If
        End If
    
Next I

' lecture rAppe
'================
arrrAppe_Max = arrrIntv_Max: arrrAppe_Nb = arrrIntv_Nb
ReDim arrrAppe(arrrAppe_Max)
ReDim arrrAppe_E(arrrAppe_Max)
ReDim arrrAppe_R(arrrAppe_Max)

srvrAppe_Init arrrAppe_E(0)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rAppe " _
        & "where Aid = " & xrIntv.Aid _
        & " and appe_s_umidl = " & xrIntv.intv_s_umidl _
        & " and appe_s_umidh  = " & xrIntv.intv_s_umidh _
        & " and appe_inst_num  = 0"
    
    arrrAppe_E(I) = arrrAppe_E(0)
    arrrAppe_R(I) = arrrAppe_E(0)
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
  
    For Boucle = 1 To 2
        If Not rsSIDE_DB.EOF Then
            V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rAppe"
            Else
                If Boucle = 1 Then
                    arrrAppe(I) = xrAppe
                End If
               ' Debug.Print I; Boucle; xrAppe.appe_type
                If xrAppe.appe_type = "APPE_EMISSION" Then
                    arrrAppe_E(I) = xrAppe
                Else
                    arrrAppe_R(I) = xrAppe
                End If
            End If
            rsSIDE_DB.MoveNext
        End If
    Next Boucle

Next I

' Affichage écran :
'==================

fgSAA_Display_99

End Sub


Private Sub cmdSAA_SQL_01()     ' *****  INTERVENTIONS OFCS

Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long

' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

ReDim arrrIntv(101)
arrrIntv_Max = 100: arrrIntv_Nb = 0

blnOk = False
Set rsSIDE_DB = Nothing

xSQL = "select * from rIntv where intv_inty_category = 'INTY_ROUTING' and " _
           & "intv_inst_num = 0 and intv_mpfn_name = 'OFCS_Detect' and " _
           & "intv_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
           & "intv_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "

'Select Case mId$(cboSelect_SAA, 1, 2)
'
'    Case "1 ": xSQL = "select * from rintv where intv_inty_category = 'INTY_ROUTING' and " _
'        & "intv_inst_num = 0 and intv_mpfn_name = 'OFCS_Detect' and " _
'        & "intv_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
'        & "intv_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "
'
'    Case Else: MsgBox "Non programmé", vbCritical, Me.Name & " : " & cboSelect_SAA
'                Exit Sub
'End Select

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrIntv_GetBuffer_ODBC(rsSIDE_DB, xrIntv)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
        Exit Sub
     Else
         arrrIntv_Nb = arrrIntv_Nb + 1
         If arrrIntv_Nb = arrrIntv_Max Then   '>
            If arrrIntv_Max >= 10000 Then
                MsgBox "10000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
   
                arrrIntv_Max = arrrIntv_Max + 50
                ReDim Preserve arrrIntv(arrrIntv_Max + 1)
            End If
         End If
         
         arrrIntv(arrrIntv_Nb) = xrIntv
    End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rIntv : " & arrrIntv_Nb): DoEvents

' lecture rMesg
'================
arrrMesg_Max = arrrIntv_Max: arrrMesg_Nb = arrrIntv_Nb
ReDim arrrMesg(arrrMesg_Max)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rMesg " _
        & "where Aid = " & xrIntv.Aid _
        & " and mesg_s_umidl = " & xrIntv.intv_s_umidl _
        & " and mesg_s_umidh  = " & xrIntv.intv_s_umidh
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rMesg"
              '  Exit Sub
            Else
                arrrMesg(I) = xrMesg
            End If
        End If
    
Next I

' lecture rInst
'================
arrrInst_Max = arrrIntv_Max: arrrInst_Nb = arrrIntv_Nb
ReDim arrrInst(arrrInst_Max)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rInst " _
        & "where Aid = " & xrIntv.Aid _
        & " and inst_s_umidl = " & xrIntv.intv_s_umidl _
        & " and inst_s_umidh  = " & xrIntv.intv_s_umidh _
        & " and inst_num  = " & xrIntv.intv_inst_num
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rInst"
            Else
                arrrInst(I) = xrInst
            End If
        End If
    
Next I

' lecture rAppe
'================
arrrAppe_Max = arrrIntv_Max: arrrAppe_Nb = arrrIntv_Nb
ReDim arrrAppe(arrrAppe_Max)
ReDim arrrAppe_E(arrrAppe_Max)
ReDim arrrAppe_R(arrrAppe_Max)

srvrAppe_Init arrrAppe_E(0)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rAppe " _
        & "where Aid = " & xrIntv.Aid _
        & " and appe_s_umidl = " & xrIntv.intv_s_umidl _
        & " and appe_s_umidh  = " & xrIntv.intv_s_umidh _
        & " and appe_inst_num  = " & xrIntv.intv_inst_num
    
    arrrAppe_E(I) = arrrAppe_E(0)
    arrrAppe_R(I) = arrrAppe_E(0)
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
  
    For Boucle = 1 To 2
        If Not rsSIDE_DB.EOF Then
            V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rAppe"
            Else
                If Boucle = 1 Then
                    arrrAppe(I) = xrAppe
                End If
               ' Debug.Print I; Boucle; xrAppe.appe_type
                If xrAppe.appe_type = "APPE_EMISSION" Then
                    arrrAppe_E(I) = xrAppe
                Else
                    arrrAppe_R(I) = xrAppe
                End If
            End If
            rsSIDE_DB.MoveNext
        End If
    Next Boucle

Next I

fgSAA_Display

End Sub



Private Sub cmdSAA_SQL_08()     ' *****  Statistiques

Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long
Dim K1 As Integer, K2 As Integer, K3 As Integer, xUnit As String

' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret
For K1 = 0 To 1000
    For K2 = 0 To 11
        For K3 = 0 To 2
            arrMt_Unit(K1, K2, K3) = 0
        Next K3
    Next K2
Next K1

blnOk = False
Set rsSIDE_DB = Nothing

xSQL = "select count(*)  as Tally , x_inst0_unit_name, mesg_sub_format, mesg_type from rMesg where " _
           & "Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
           & "Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
           & " group by x_inst0_unit_name, mesg_sub_format, mesg_type"
        

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
      xUnit = rsSIDE_DB("x_inst0_unit_name")
      K2 = cmdSAA_SQL_08_Unit(xUnit)
      K3 = IIf(rsSIDE_DB("mesg_sub_format") = "INPUT", 1, 2)
      If IsNull(rsSIDE_DB("mesg_type")) Then
         K1 = 1
       Else
         K1 = CInt(rsSIDE_DB("mesg_type"))
        End If
      If Not IsNull(rsSIDE_DB("Tally")) Then
         arrMt_Unit(K1, K2, K3) = arrMt_Unit(K1, K2, K3) + rsSIDE_DB("Tally")
         arrMt_Unit(K1, 0, 0) = 1
      End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rMesg : " & arrrMesg_Nb): DoEvents

' Affectation des swifts entrants 'None'

xSQL = "select mesg_sub_format, mesg_type, mesg_sender_x1 from rMesg where " _
           & "Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
           & "Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
           & " and x_inst0_unit_name = 'None' and mesg_sub_format= 'OUTPUT'"
        

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
      K3 = IIf(rsSIDE_DB("mesg_sub_format") = "INPUT", 1, 2)
      If IsNull(rsSIDE_DB("mesg_type")) Then
         K1 = 1
      Else
         K1 = CInt(rsSIDE_DB("mesg_type"))
      End If
        I = Int(K1 / 100)
        Select Case I
            Case 1, 2
                If rsSIDE_DB("mesg_sender_x1") = "SOGEFRPPTGV" Then
                    xUnit = "SOBF"
                Else
                    xUnit = "ORPA"
                End If
            Case 3: xUnit = "BOTC"
            Case 4: xUnit = "ORPA"
            Case 5: xUnit = "DAFI"
            Case 7: xUnit = "SOBI"
            Case 8: xUnit = "SOBF"
            Case 9:
                Select Case K1
                    Case 950: xUnit = "CSOP"
                    Case 900, 910: xUnit = "SOBF"
                    Case 960, 961, 962, 963: xUnit = "SCLE"
                    Case Else: xUnit = "None"
                End Select
            Case Else: xUnit = "None"
        End Select
      K2 = cmdSAA_SQL_08_Unit(xUnit)
      If K2 <> 10 Then
         arrMt_Unit(K1, 10, K3) = arrMt_Unit(K1, 10, K3) - 1
         arrMt_Unit(K1, K2, K3) = arrMt_Unit(K1, K2, K3) + 1
         arrMt_Unit(K1, 0, 0) = 1
      End If
    rsSIDE_DB.MoveNext

Loop
            
arrMt_Unit(1000, 0, 0) = 1
For K1 = 1 To 999
    For K2 = 1 To 10
            arrMt_Unit(K1, 11, 1) = arrMt_Unit(K1, 11, 1) + arrMt_Unit(K1, K2, 1)
            arrMt_Unit(K1, 11, 2) = arrMt_Unit(K1, 11, 2) + arrMt_Unit(K1, K2, 2)
            arrMt_Unit(1000, K2, 1) = arrMt_Unit(1000, K2, 1) + arrMt_Unit(K1, K2, 1)
            arrMt_Unit(1000, K2, 2) = arrMt_Unit(1000, K2, 2) + arrMt_Unit(K1, K2, 2)
    Next K2
Next K1
For K2 = 1 To 10
        arrMt_Unit(1000, 11, 1) = arrMt_Unit(1000, 11, 1) + arrMt_Unit(1000, K2, 1)
        arrMt_Unit(1000, 11, 2) = arrMt_Unit(1000, 11, 2) + arrMt_Unit(1000, K2, 2)
Next K2


fgSAA_Display_08
cmdPrint.Enabled = True
End Sub

Private Sub cmdSAA_SQL_02()     ' *****  EN VIOLATION OFCS
Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long

' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

ReDim arrrIntv(101)
arrrIntv_Max = 100: arrrIntv_Nb = 0

blnOk = False
Set rsSIDE_DB = Nothing

xSQL = "select * from rIntv where intv_inty_category = 'INTY_OTHER' and " _
           & "intv_inst_num = 0 and intv_mpfn_name = 'OFCS_Detect' and " _
           & "intv_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
           & "intv_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrIntv_GetBuffer_ODBC(rsSIDE_DB, xrIntv)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
        Exit Sub
     Else
         arrrIntv_Nb = arrrIntv_Nb + 1
         If arrrIntv_Nb = arrrIntv_Max Then   '>
            If arrrIntv_Max >= 10000 Then
                MsgBox "10000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
   
                arrrIntv_Max = arrrIntv_Max + 50
                ReDim Preserve arrrIntv(arrrIntv_Max + 1)
            End If
         End If
         
         arrrIntv(arrrIntv_Nb) = xrIntv
    End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rIntv : " & arrrIntv_Nb): DoEvents

' lecture rMesg
'================
arrrMesg_Max = arrrIntv_Max: arrrMesg_Nb = arrrIntv_Nb
ReDim arrrMesg(arrrMesg_Max)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rMesg " _
        & "where Aid = " & xrIntv.Aid _
        & " and mesg_s_umidl = " & xrIntv.intv_s_umidl _
        & " and mesg_s_umidh  = " & xrIntv.intv_s_umidh
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rMesg"
              '  Exit Sub
            Else
                arrrMesg(I) = xrMesg
            End If
        End If
    
Next I

' lecture rInst
'================
arrrInst_Max = arrrIntv_Max: arrrInst_Nb = arrrIntv_Nb
ReDim arrrInst(arrrInst_Max)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rInst " _
        & "where Aid = " & xrIntv.Aid _
        & " and inst_s_umidl = " & xrIntv.intv_s_umidl _
        & " and inst_s_umidh  = " & xrIntv.intv_s_umidh _
        & " and inst_num  = " & xrIntv.intv_inst_num
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rInst"
            Else
                arrrInst(I) = xrInst
            End If
        End If
    
Next I

' lecture rAppe
'================
arrrAppe_Max = arrrIntv_Max: arrrAppe_Nb = arrrIntv_Nb
ReDim arrrAppe(arrrAppe_Max)
ReDim arrrAppe_E(arrrAppe_Max)
ReDim arrrAppe_R(arrrAppe_Max)

srvrAppe_Init arrrAppe_E(0)

For I = 1 To arrrIntv_Nb
    xrIntv = arrrIntv(I)
    xSQL = "select * from rAppe " _
        & "where Aid = " & xrIntv.Aid _
        & " and appe_s_umidl = " & xrIntv.intv_s_umidl _
        & " and appe_s_umidh  = " & xrIntv.intv_s_umidh _
        & " and appe_inst_num  = " & xrIntv.intv_inst_num
    
    arrrAppe_E(I) = arrrAppe_E(0)
    arrrAppe_R(I) = arrrAppe_E(0)
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
  
    For Boucle = 1 To 2
        If Not rsSIDE_DB.EOF Then
            V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rAppe"
            Else
                If Boucle = 1 Then
                    arrrAppe(I) = xrAppe
                End If
               ' Debug.Print I; Boucle; xrAppe.appe_type
                If xrAppe.appe_type = "APPE_EMISSION" Then
                    arrrAppe_E(I) = xrAppe
                Else
                    arrrAppe_R(I) = xrAppe
                End If
            End If
            rsSIDE_DB.MoveNext
        End If
    Next Boucle

Next I

fgSAA_Display

End Sub

Private Sub cmdSAA_SQL_03()     ' *****  MESSAGES  ACK
Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long

' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

ReDim arrrAppe(101)
arrrAppe_Max = 100: arrrAppe_Nb = 0

ReDim arrrAppe_E(101)
ReDim arrrAppe_R(101)

srvrAppe_Init arrrAppe_E(0)
srvrAppe_Init arrrAppe_R(0)
   
blnOk = False
Set rsSIDE_DB = Nothing

xSQL = "select * from rAppe where appe_iapp_name = 'SWIFT' and " _
           & "appe_inst_num = 0 and appe_network_delivery_status = 'DLV_ACKED' and " _
           & "appe_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
           & "appe_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
        Exit Sub
     Else
         arrrAppe_Nb = arrrAppe_Nb + 1
         If arrrAppe_Nb = arrrAppe_Max Then   '>
            If arrrAppe_Max >= 10000 Then
                MsgBox "10000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
                arrrAppe_Max = arrrAppe_Max + 50
                ReDim Preserve arrrAppe(arrrAppe_Max + 1)
                ReDim Preserve arrrAppe_E(arrrAppe_Max + 1)
                ReDim Preserve arrrAppe_R(arrrAppe_Max + 1)
            End If
         End If
         
         arrrAppe(arrrAppe_Nb) = xrAppe
         
         ' Pour préparer la ligne d'affichage au fgSAA_displayLine
         arrrAppe_E(arrrAppe_Nb) = arrrAppe_E(0)
         arrrAppe_R(arrrAppe_Nb) = arrrAppe_R(0)
         If xrAppe.appe_type = "APPE_EMISSION" Then
            arrrAppe_E(arrrAppe_Nb) = xrAppe
         Else
            arrrAppe_R(arrrAppe_Nb) = xrAppe
         End If

    End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rAppe : " & arrrAppe_Nb): DoEvents

' lecture rMesg
'================
arrrMesg_Max = arrrAppe_Max: arrrMesg_Nb = arrrAppe_Nb
ReDim arrrMesg(arrrMesg_Max)

For I = 1 To arrrAppe_Nb
    xrAppe = arrrAppe(I)
    xSQL = "select * from rMesg " _
        & "where Aid = " & xrAppe.Aid _
        & " and mesg_s_umidl = " & xrAppe.appe_s_umidl _
        & " and mesg_s_umidh  = " & xrAppe.appe_s_umidh
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rMesg"
              '  Exit Sub
            Else
                arrrMesg(I) = xrMesg
            End If
        End If
    
Next I

' lecture rInst
'================
arrrInst_Max = arrrAppe_Max: arrrInst_Nb = arrrAppe_Nb
ReDim arrrInst(arrrInst_Max)

For I = 1 To arrrAppe_Nb
    xrAppe = arrrAppe(I)
    xSQL = "select * from rInst " _
        & "where Aid = " & xrAppe.Aid _
        & " and inst_s_umidl = " & xrAppe.appe_s_umidl _
        & " and inst_s_umidh  = " & xrAppe.appe_s_umidh _
        & " and inst_num  = " & xrAppe.appe_inst_num
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rInst"
            Else
                arrrInst(I) = xrInst
            End If
        End If
    
Next I

' lecture rIntv : CHARGER SEULEMENT le nbr utilisé dans fgSAA_Display
'================
arrrIntv_Max = arrrAppe_Max: arrrIntv_Nb = arrrAppe_Nb
ReDim arrrIntv(arrrIntv_Max)

srvrIntv_Init arrrIntv(0)

For I = 1 To arrrAppe_Nb
    arrrIntv(I) = arrrIntv(0)
Next I


fgSAA_Display

End Sub

Private Sub cmdSAA_SQL_04()     ' *****  MESSAGES  NACK
Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim I As Integer
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim Boucle As Long

' Transformer la période pour pouvoir comparer avec les dates de SIDE_DB

Call DTPicker_Amj8_tiret(txtSelect_from_crea_date_time, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect_to_crea_date_time, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

ReDim arrrAppe(101)
arrrAppe_Max = 100: arrrAppe_Nb = 0

ReDim arrrAppe_E(101)
ReDim arrrAppe_R(101)

srvrAppe_Init arrrAppe_E(0)
srvrAppe_Init arrrAppe_R(0)
   
blnOk = False
Set rsSIDE_DB = Nothing

xSQL = "select * from rAppe where appe_iapp_name = 'SWIFT' and " _
           & "appe_inst_num = 0 and appe_network_delivery_status = 'DLV_NACKED' and " _
           & "appe_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} and " _
           & "appe_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} "

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, xrAppe)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.cmdSelect_SQL"
        Exit Sub
     Else
         arrrAppe_Nb = arrrAppe_Nb + 1
         If arrrAppe_Nb = arrrAppe_Max Then   '>
            If arrrAppe_Max >= 10000 Then
                MsgBox "10000 lignes max.", vbCritical, Me.Name & " : " & currentAction
                Exit Do
            Else
                arrrAppe_Max = arrrAppe_Max + 50
                ReDim Preserve arrrAppe(arrrAppe_Max + 1)
                ReDim Preserve arrrAppe_E(arrrAppe_Max + 1)
                ReDim Preserve arrrAppe_R(arrrAppe_Max + 1)
            End If
         End If
         
         arrrAppe(arrrAppe_Nb) = xrAppe
         
         ' Pour préparer la ligne d'affichage au fgSAA_displayLine
         arrrAppe_E(arrrAppe_Nb) = arrrAppe_E(0)
         arrrAppe_R(arrrAppe_Nb) = arrrAppe_R(0)
         If xrAppe.appe_type = "APPE_EMISSION" Then
            arrrAppe_E(arrrAppe_Nb) = xrAppe
         Else
            arrrAppe_R(arrrAppe_Nb) = xrAppe
         End If

    End If
    rsSIDE_DB.MoveNext

Loop
Call lstErr_AddItem(lstErr, cmdContext, "rAppe : " & arrrAppe_Nb): DoEvents

' lecture rMesg
'================
arrrMesg_Max = arrrAppe_Max: arrrMesg_Nb = arrrAppe_Nb
ReDim arrrMesg(arrrMesg_Max)

For I = 1 To arrrAppe_Nb
    xrAppe = arrrAppe(I)
    xSQL = "select * from rMesg " _
        & "where Aid = " & xrAppe.Aid _
        & " and mesg_s_umidl = " & xrAppe.appe_s_umidl _
        & " and mesg_s_umidh  = " & xrAppe.appe_s_umidh
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rMesg"
              '  Exit Sub
            Else
                arrrMesg(I) = xrMesg
            End If
        End If
    
Next I

' lecture rInst
'================
arrrInst_Max = arrrAppe_Max: arrrInst_Nb = arrrAppe_Nb
ReDim arrrInst(arrrInst_Max)

For I = 1 To arrrAppe_Nb
    xrAppe = arrrAppe(I)
    xSQL = "select * from rInst " _
        & "where Aid = " & xrAppe.Aid _
        & " and inst_s_umidl = " & xrAppe.appe_s_umidl _
        & " and inst_s_umidh  = " & xrAppe.appe_s_umidh _
        & " and inst_num  = " & xrAppe.appe_inst_num
    
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            V = srvrInst_GetBuffer_ODBC(rsSIDE_DB, xrInst)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "rInst"
            Else
                arrrInst(I) = xrInst
            End If
        End If
    
Next I

' lecture rIntv : CHARGER SEULEMENT le nbr utilisé dans fgSAA_Display
'================
arrrIntv_Max = arrrAppe_Max: arrrIntv_Nb = arrrAppe_Nb
ReDim arrrIntv(arrrIntv_Max)

srvrIntv_Init arrrIntv(0)

For I = 1 To arrrAppe_Nb
    arrrIntv(I) = arrrIntv(0)
Next I


fgSAA_Display

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long


blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Swift_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fgZSWIFTA0.Clear
fgZSWIHIA0.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Swift_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYSWIMON0_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        xYSWIMON0 = arrYSWIMON0(arrYSWIMON0_Index)
        
       If xYSWIMON0.SAAAID = 0 Then Me.PopupMenu mnuSelect, vbPopupMenuLeftButton

   End If
End If
fgSelect.LeftCol = 0
End Sub

Private Sub fgZSWIFTA0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Dim xWhere As String

On Error Resume Next
If y <= fgZSWIFTA0.RowHeightMin Then
    Select Case fgZSWIFTA0.Col
        Case 0: fgZSWIFTA0_Sort1 = 0: fgZSWIFTA0_Sort2 = 1: fgZSWIFTA0_Sort
        Case 1:  fgZSWIFTA0_Sort1 = 1: fgZSWIFTA0_Sort2 = 1: fgZSWIFTA0_Sort
        Case 2: fgZSWIFTA0_Sort1 = 2: fgZSWIFTA0_Sort2 = 2: fgZSWIFTA0_Sort
        Case 3: fgZSWIFTA0_Sort1 = 3: fgZSWIFTA0_Sort2 = 3: fgZSWIFTA0_SortX 3
        Case 4: fgZSWIFTA0_Sort1 = 4: fgZSWIFTA0_Sort2 = 4: fgZSWIFTA0_Sort
        Case 5: fgZSWIFTA0_Sort1 = 5: fgZSWIFTA0_Sort2 = 5: fgZSWIFTA0_SortX 5
        Case 6: fgZSWIFTA0_Sort1 = 6: fgZSWIFTA0_Sort2 = 6: fgZSWIFTA0_Sort
        Case 7: fgZSWIFTA0_Sort1 = 7: fgZSWIFTA0_Sort2 = 7: fgZSWIFTA0_Sort
        Case 8: fgZSWIFTA0_Sort1 = 8: fgZSWIFTA0_Sort2 = 8: fgZSWIFTA0_Sort
        Case 9: fgZSWIFTA0_Sort1 = 9: fgZSWIFTA0_Sort2 = 9: fgZSWIFTA0_SortX 9
    End Select
Else
    If fgZSWIFTA0.Rows > 1 Then
        fgZSWIFTA0.Col = fgZSWIFTA0_arrIndex:  arrZSWIFTA0_Index = CLng(fgZSWIFTA0.Text)
        xZSWIFTA0 = arrZSWIFTA0(arrZSWIFTA0_Index)
        fgZSWIFTA0.LeftCol = 0

       If cmdZSWIFTA0_Update.Visible Then
'Sélection /Désélection
            If fgZSWIFTA0.CellBackColor = cmdZSWIFTA0_Update.BackColor Then
                fgZSWIFTA0_Update_SAA_Queue ""
            Else
                Me.PopupMenu mnuSAA_Queue, vbPopupMenuLeftButton
           End If
            
        Else
            Call fgZSWIFTA0_Color(fgZSWIFTA0_RowClick, MouseMoveUsr.BackColor, fgZSWIFTA0_ColorClick)

            fgZSWI_D.Clear
            fgZSWI_D.Visible = True
            xWhere = " where SWIFTBNUM = " & xZSWIFTA0.SWIFTANUM
            srvZSWIFTA0_fgDisplay xZSWIFTA0, fgZSWI_D
            arrZSWIFTB0_SQL xWhere, False
            arrZSWIFTB0_lstW
        End If
   End If
End If
fgZSWIFTA0.LeftCol = 0
fgZSWIFTA0.Col = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'If blnControl Then
    cnSIDE_DB.Close
    Set cnSIDE_DB = Nothing

'End If
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
If cmdZSWIFTA0_Update.Visible Then
    If mnuSAA_Queue.Visible Then
        Exit Sub
    Else
        X = MsgBox("Voulez-vous abandonner la mise à jour ?", vbYesNo, "BIA_SWIFT : Emission")
        If X = vbYes Then cmdZSWIFTA0_Update.Visible = False
        Exit Sub
    End If
End If
If lstW.Visible Then lstW.Visible = False: Exit Sub
If fgrTextField.Visible And SSTab1.Tab = 1 Then fgrTextField.Visible = False: Exit Sub
If fgZSWI_D.Visible And SSTab1.Tab = 2 Then fgZSWI_D.Visible = False: Exit Sub
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    fgSelect.Row = fgSelect.TopRow
    fgSelect.Col = fgSelect_arrIndex: ' wK1 = fgSelect.Text
    'cmdSelect txtSelect ''fgSelect.Text

'    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False


cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
Exit Sub

Error_Handler:

blnControl = False
MsgBox Error
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
'Call txt_GotFocus(txt)
'Call txt_LostFocus(txt)

End Sub


Public Sub fgStatus_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgStatus.Row

If lRow > 0 And lRow < fgStatus.Rows Then
    fgStatus.Row = lRow
    For I = 0 To fgStatus_arrIndex
        fgStatus.Col = I: fgStatus.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgStatus.Row = mRow
    If fgStatus.Row > 0 Then
        lRow = fgStatus.Row
        lColor_Old = fgStatus.CellBackColor
        For I = 0 To fgStatus_arrIndex
          fgStatus.Col = I: fgStatus.CellBackColor = lColor
        Next I
        fgStatus.Col = 0
    End If
End If

End Sub
Private Sub fgStatus_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 3
fgStatus.Visible = False
fgStatus_Reset
cmdPrint.Enabled = False

fgStatus.Rows = 1
fgStatus.FormatString = fgStatus_FormatString
currentAction = "fgStatus_Display"
    
For I = 1 To arrrInst_Status_Nb
         
    xrInst = arrrInst_Status(I)
    fgStatus_DisplayLine I
Next I

fgStatus.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrrInst_Status_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 3
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Public Sub fgStatus_DisplayLine(lIndex As Long)
On Error Resume Next
fgStatus.Rows = fgStatus.Rows + 1
fgStatus.Row = fgStatus.Rows - 1
fgStatus.Col = 0: fgStatus.Text = xrInst.Aid
fgStatus.Col = 1: fgStatus.Text = xrInst.inst_s_umidl
fgStatus.Col = 2: fgStatus.Text = xrInst.inst_s_umidh
fgStatus.Col = 3: fgStatus.Text = xrInst.x_last_emi_appe_date_time
fgStatus.Col = fgStatus_arrIndex: fgStatus.Text = lIndex

End Sub
Public Sub fgStatus_Reset()
fgStatus.Clear
fgStatus_Sort1 = 0: fgStatus_Sort2 = 0
fgStatus_Sort1_Old = -1
fgStatus_RowDisplay = 0: fgStatus_RowClick = 0
fgStatus_arrIndex = fgStatus.Cols - 1
blnfgStatus_DisplayLine = False
fgStatus_SortAD = 6
fgStatus.LeftCol = 0

End Sub


Public Sub fgStatus_Sort()
If fgStatus.Rows > 1 Then
    fgStatus.Row = 1
    fgStatus.RowSel = fgStatus.Rows - 1
    
    If fgStatus_Sort1_Old = fgStatus_Sort1 Then
        If fgStatus_SortAD = 5 Then
            fgStatus_SortAD = 6
        Else
            fgStatus_SortAD = 5
        End If
    Else
        fgStatus_SortAD = 5
    End If
    fgStatus_Sort1_Old = fgStatus_Sort1
    
    fgStatus.Col = fgStatus_Sort1
    fgStatus.ColSel = fgStatus_Sort2
    fgStatus.Sort = fgStatus_SortAD
End If

End Sub




Private Sub mnuSelect_Print_Détail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"D "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"L "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
    Case 1: fgSAA.LeftCol = 0
    Case 2: fgZSWIFTA0.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok()
End Sub


Public Sub cmdPrint_List6_Ok()
Dim iRow As Integer, K As Integer, I As Integer, X As String
Dim blnOk As Boolean, blnOpen As Boolean
Dim xCOMPTEINT As String
Dim m_unit_name As String
Dim mNrequest As String

mNrequest = Mid$(cboSelect_SAA, 1, 2)
fgSAA.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSAA.Rows - 1)

If fgSAA.Rows > 1 Then
    fgSAA_Sort1_Old = -1
    fgSAA_SortX 13
End If

blnOpen = False
m_unit_name = ""

For iRow = 1 To fgSAA.Rows - 1
    
    fgSAA.Row = iRow
    fgSAA.Col = fgSAA_arrIndex:  K = CLng(fgSAA.Text)
    xrMesg = arrrMesg(K)
    xrAppe_E = arrrAppe_E(K)
    xrAppe_R = arrrAppe_R(K)
    xrInst = arrrInst(K)
' Rupture Service : ligne totale de l'ancien code de service

    If m_unit_name <> xrMesg.x_inst0_unit_name Then
        If blnOpen Then
            ''prtSWI_Messages_List6_Rupture m_unit_name, mNrequest
            prtSWI_Messages_List6_Close m_unit_name, mNrequest
            If blnAuto Then Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "")
        End If
        m_unit_name = xrMesg.x_inst0_unit_name
        ' If blnAuto Then Call Printer_Set_Unit(m_unit_name)
        X = Table_Unit_SSI("", m_unit_name)
        If blnAuto Then Call frmElpPrt.prtIMP_PDF_NoPaper_Init(X, "BIA-SAA-CTL", "Archive")
        prtSWI_Messages_List6_Open Mid$(cboSelect_SAA, 1, 2), txtSelect_from_crea_date_time, txtSelect_to_crea_date_time
        blnOpen = True
   End If

    prtSWI_Messages_List6_Line Mid$(cboSelect_SAA, 1, 2), xrMesg, xrAppe_E, xrAppe_R, xrInst

Next iRow

If blnOpen Then
   '' prtSWI_Messages_List6_Rupture m_unit_name, mNrequest
    prtSWI_Messages_List6_Close m_unit_name, mNrequest
    If blnAuto Then Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "")

End If
fgSAA.Visible = True
Me.Show

End Sub

Public Sub cmdPrint_List8_Ok()
Dim iRow As Integer, K As Integer, I As Integer
Dim blnOk As Boolean, blnOpen As Boolean
Dim xCOMPTEINT As String
Dim m_unit_name As String
Dim mNrequest As String
Dim K1 As Integer, K2 As Integer, K3 As Integer, xUnit As String

mNrequest = Mid$(cboSelect_SAA, 1, 2)
fgSAA.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Statistiques")

prtSWI_Messages_List6_Open mNrequest, txtSelect_from_crea_date_time, txtSelect_to_crea_date_time

arrMt_Unit(1000, 0, 0) = 1
For K1 = 0 To 1000
    If arrMt_Unit(K1, 0, 0) <> 0 Then

        If XPrt.CurrentY + 300 > prtMaxY Then
            frmElpPrt.prtNewPage
            prtSWI_Messages_List8_Form
        End If

' Impression de la ligne courante
        XPrt.CurrentX = prtMinX + 500
        XPrt.ForeColor = vbBlack
        If K1 <> 1000 Then
            XPrt.FontBold = False
            XPrt.Print Format(K1, "000");
        Else
            XPrt.FontBold = True
            XPrt.Print "===";
        End If
        For K2 = 1 To 11
            If K2 = 11 Then XPrt.FontBold = True
            X = Format(arrMt_Unit(K1, K2, 1), "#####")
            XPrt.CurrentX = prtMinX + 1300 * K2 + 600 - XPrt.TextWidth(X)
            XPrt.ForeColor = vbRed
            XPrt.Print X;
            X = Format(arrMt_Unit(K1, K2, 2), "#####")
            XPrt.CurrentX = prtMinX + 1300 * K2 + 1230 - XPrt.TextWidth(X)
            XPrt.ForeColor = vbBlue
            XPrt.Print X;
        Next K2
         XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
        XPrt.CurrentY = XPrt.CurrentY + 50
   End If
    
Next K1

prtSWI_Messages_List6_Close "*", mNrequest



fgSAA.Visible = True
Me.Show

End Sub

Public Sub mnuSelect_SAAAID_Evolution()

Dim V, W
Dim Err_Code As String, Msg As String
Dim wZone As String, wAmj8_tiret As String, wTime As String, wTime8C_From As String, wTime8C_To As String
Dim xSQL As String, wBatchRef As String
Dim SSS As Long
Dim blnCcy As Boolean, blnAmt As Boolean, blnOFAC As Boolean, blnNAK As Boolean, blnStatus As Boolean


' arrYSWIMON0(arrYSWIMON0_Index) = Old record prefix / xYSWIMON0 = New record prefix

ReDim arrrMesg(101)
arrrMesg_Max = 100: arrrMesg_Nb = 0
Set rsSIDE_DB = Nothing

Select Case xYSWIMON0.SWIMONSTA

' >>>>>  Rapprochement ID SAB et ID SAA
    Case "S200":
    
        ' Recherche messages avec -S200- dans rMesg de SIDE
       
        ' Zones nécessaires pour la requête SQL sur rMesg
        
        wZone = xYSWIMON0.SWIMONFLUD
        wAmj8_tiret = Mid$(wZone, 1, 4) & "-" & Mid$(wZone, 5, 2) & "-" & Mid$(wZone, 7, 2)
        ' 180 secondes = Durée entre messages sortis de SAB et messages constitués ds Alliance
        wZone = xYSWIMON0.SWIMONFLUH
        wTime8C_From = Mid$(wZone, 1, 2) & ":" & Mid$(wZone, 3, 2) & ":" & Mid$(wZone, 5, 2) & ".000"
        SSS = Time_Hms_Sss(wZone) + 180
        wZone = Time_Sss_Hms(SSS)
        wTime8C_To = Mid$(wZone, 1, 2) & ":" & Mid$(wZone, 3, 2) & ":" & Mid$(wZone, 5, 2) & ".000"
        wBatchRef = xYSWIMON0.SWIMONFLUD & "_" & Mid$(xYSWIMON0.SWIMONFLUH, 1, 2) & "%_YSWIALL0_0%.pcc/" & xYSWIMON0.SWIMONFLUS & "/"
        
        xSQL = "select * from rMesg where mesg_sub_format = 'INPUT' and mesg_crea_rp_name like '_AI_from_APPLI%' and " _
                   & "mesg_frmt_name = 'Swift' and mesg_type = '" & xYSWIMON0.SWIMONXMT & "' and " _
                   & "mesg_trn_ref like '" & Trim(xYSWIMON0.SWIMONX20) & "%' And " _
                   & "mesg_batch_reference like  '" & wBatchRef & "%' And " _
                   & "mesg_crea_date_time >= {ts '" & wAmj8_tiret & " " & wTime8C_From & "'} and " _
                   & "mesg_crea_date_time < {ts '" & wAmj8_tiret & " " & wTime8C_To & "'} "
       
    Case Else:
        ' Recherche d'un message déjà répertorié
        
        xSQL = "select * from rMesg " _
                & "where Aid = " & xYSWIMON0.SAAAID _
                & " and mesg_s_umidl = " & xYSWIMON0.SAAUMIDL _
                & " and mesg_s_umidh  = " & xYSWIMON0.SAAUMIDH
                
End Select
       
' Lecture rMesg dans SIDE

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Do While Not rsSIDE_DB.EOF
    V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)

    If Not IsNull(V) Then
        MsgBox "ERR01 -" & " " & xYSWIMON0.SWIMONID & " : " & "lecture rMesg échouée - rapprochement ID non trouvé"
        Exit Sub
    Else
        arrrMesg_Nb = arrrMesg_Nb + 1
        If arrrMesg_Nb = arrrMesg_Max Then   '>
            If arrrMesg_Max >= 10000 Then
            MsgBox "10000 lignes max.", vbCritical, "mnuSelect_SAAAID"
            Exit Do
        Else
            arrrMesg_Max = arrrMesg_Max + 50
            ReDim Preserve arrrMesg(arrrMesg_Max + 1)
            End If
        End If
        
        arrrMesg(arrrMesg_Nb) = xrMesg
    End If
    rsSIDE_DB.MoveNext
Loop

' 2) Traitement de tous les messages lus à partir de rMesg de SIDE

For I = 1 To arrrMesg_Nb
    
    xrMesg = arrrMesg(I)
    ' Critères de rapprochement supplémentaire
    
    blnCcy = True
    If Trim(xrMesg.x_fin_ccy) = "" Or IsNull(xrMesg.x_fin_ccy) Then
    Else
        If Mid$(xrMesg.x_fin_ccy, 1, 3) <> xYSWIMON0.SWIMONX32D Then blnCcy = False
    End If
    
    blnAmt = True
    If Not IsNull(xrMesg.x_fin_amount) Then
        If Fix(xrMesg.x_fin_amount) <> Fix(xYSWIMON0.SWIMONX32A) Then blnAmt = False
    End If
   
    If Trim(xrMesg.mesg_rel_trn_ref) = Trim(xYSWIMON0.SWIMONX21) And _
       blnCcy And blnAmt Then
       
        If xYSWIMON0.SAAAID = 0 Then xYSWIMON0.SAAAID = xrMesg.Aid
        If xYSWIMON0.SAAUMIDL = 0 Then xYSWIMON0.SAAUMIDL = xrMesg.mesg_s_umidl
        If xYSWIMON0.SAAUMIDH = 0 Then xYSWIMON0.SAAUMIDH = xrMesg.mesg_s_umidh
       
        xYSWIMON0.SAAQMOD = xrMesg.mesg_is_text_modified
        xYSWIMON0.SAAUNIT = xrMesg.x_inst0_unit_name
        
        mnuSelect_SAAAID_Status
        
    Else
        ' >>>>> QUE FAIRE SI IMPOSSIBLE DE RAPPROCHER
        
    End If
Next I

End Sub


Private Sub txtSelect_SWIFTADES_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSelect_SWIFTAMES_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSelect_SWIFTAREF_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Public Sub fgZSWIFTA0_Update_Enabled(lSAA_Queue As String)
Dim I As Integer
ReDim arrZSWIFTA0_SAA_Queue(arrZSWIFTA0_Max)

cmdZSWIFTA0_Update.Visible = SWI_MESSAGES_Aut.Swift
For I = 1 To fgZSWIFTA0.Rows - 1
        fgZSWIFTA0.Row = I
        fgZSWIFTA0.Col = fgZSWIFTA0_arrIndex:  arrZSWIFTA0_Index = CLng(fgZSWIFTA0.Text)
        fgZSWIFTA0.LeftCol = 0
        fgZSWIFTA0_Update_SAA_Queue lSAA_Queue
        
Next I
End Sub

Public Function cmdZSWIFTA0_Update_Transaction_RJE()
'=========================================================================
Dim V
Dim blnChamp_Préfixe As Boolean, mChamp As String, lenX As Integer, K As Integer
Dim blnOk As Boolean, wX1 As String, wX2 As String
V = Null

Dim X As String, I As Integer
Dim wBlock1 As String, wBlock2 As String
On Error GoTo Error_Handler

If Mid$(arrZSWIFTB0(1).SWIFTBDET, 1, 2) <> "+1" Then V = "cmdZSWIFTA0_Update_Transaction_RJE : Erreur Block +1": Exit Function
If Mid$(arrZSWIFTB0(2).SWIFTBDET, 1, 2) <> "+2" Then V = "cmdZSWIFTA0_Update_Transaction_RJE : Erreur Block +2": Exit Function
If Mid$(arrZSWIFTB0(arrZSWIFTB0_Nb).SWIFTBDET, 1, 2) <> "-1" Then V = "cmdZSWIFTA0_Update_Transaction_RJE : Erreur Block -1": Exit Function

wBlock1 = "{1:F01" & paramBic8 & "AXXX0000000000}"
wBlock2 = "{2:" & Mid$(arrZSWIFTB0(2).SWIFTBDET, 3, 12) & "X" & Mid$(arrZSWIFTB0(2).SWIFTBDET, 15, 3) & Mid$(arrZSWIFTB0(2).SWIFTBDET, 19, 1) & "}{4:" ''' Ignorer le block 3

Call SAA_from_SAB_Block2(wBlock2)

If newYSWIMON0.SWIMONFLUS = 1 Then
    Print #2, wBlock1 & wBlock2
Else
    Print #2, "$" & wBlock1 & wBlock2
End If

For I = 3 To arrZSWIFTB0_Nb - 1
    If arrSWIFTCSIG(I) = ">" Then
        X = Trim(arrZSWIFTB0(I).SWIFTBDET)
        
'=========================================================================
'Ne pas générer de lignes intempestives par exemple
'           :50K:/
'           blablabla
'et préfixer, le cas échéant,la ligne suivante parle code du champ
'           :50K:blablabla


        blnOk = True
        If Mid$(X, 1, 1) = ":" Then
            
            K = InStr(2, X, ":")
            If K > 0 Then
                blnChamp_Préfixe = False
                mChamp = Mid$(X, 1, K)
                ' Supprimer les espaces entre la fin du code et le 1 er caractère significatif
                '-----------------------------------------------------------------------------
                lenX = Len(X)
                If lenX > K Then
                    wX1 = Trim(Mid$(X, K + 1, lenX - K))
                    X = mChamp & wX1
                End If
                wX1 = mChamp & "/"
                wX2 = mChamp & "/ /"
                If X = wX1 Or X = wX2 Then blnOk = False: blnChamp_Préfixe = True
            End If
        Else
            If blnChamp_Préfixe Then
                blnChamp_Préfixe = False
                X = mChamp & X
            End If
        End If
'=========================================================================
                
        
        If blnOk Then
            lenX = Len(X)
            Select Case mChamp
                Case ":20:": X = ":20:" & newYSWIMON0.SWIMONX20 '''pour test : Mid$(X, 5, 4) = "TEST"
                Case ":21:": newYSWIMON0.SWIMONX21 = Mid$(X, 5, Len(X) - 4)
                Case ":32A:": Call SAA_X32(X, newYSWIMON0.SWIMONX32V, newYSWIMON0.SWIMONX32D, newYSWIMON0.SWIMONX32A)
                
            End Select
            If lenX > 8 And Mid$(X, 5, 4) = ":/ /" Then X = Mid$(X, 1, 5) & Mid$(X, 8, lenX - 7)
            
            Print #2, Trim(X)
        End If
    End If
Next I
Print #2, "-}"


GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : cmdZSWIFTA0_Update_Transaction_Historique"
    
Exit_Function:
    On Error Resume Next
    cmdZSWIFTA0_Update_Transaction_RJE = V

        
'=========================================================================
            ''Case Chr$(&HE9): Mid$(lSwift, K, 1) = Chr$(&H7B)
            ''Case Chr$(&HE8): Mid$(lSwift, K, 1) = Chr$(&H7D)

'=========================================================================

End Function

Public Function cmdZSWIALI0_Update_Transaction_RJE(lSWIALIDON As String, lenSWIALIDON As Long)
'=========================================================================
Dim V
Dim blnChamp_Préfixe As Boolean, mChamp As String, lenX As Integer, K As Integer
Dim blnExit As Boolean, wX1 As String, wX2 As String
Dim K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer
Dim kSG As Integer, kX0 As Integer, kX1 As Integer, kX2 As Integer

V = Null

Dim X As String, I As Integer
Dim wBlock1 As String, wBlock2 As String, wBlock3 As String
Dim debutBlock3 As Long
Dim finBlock3 As Long

On Error GoTo Error_Handler

wBlock1 = "{1:F01" & paramBic8 & "AXXX0000000000}"

K2 = InStr(4, lSWIALIDON, "é2:"): Mid$(lSWIALIDON, K2, 1) = "{"
K3 = InStr(4, lSWIALIDON, "é3:")
K4 = InStr(4, lSWIALIDON, "é4:")
If K3 > 0 Then
    wBlock2 = Mid$(lSWIALIDON, K2, K3 - K2)
    wBlock3 = Mid$(lSWIALIDON, K3, K4 - K3)
    wBlock3 = Replace(wBlock3, "é", "{")
    wBlock3 = Replace(wBlock3, "è", "}")
    '04/12/2018 Prendre uniquement le code '{121:' dans le block 3
    debutBlock3 = InStr(1, wBlock3, "{121:")
    If debutBlock3 >= 1 Then
        finBlock3 = InStr(debutBlock3, wBlock3, "}")
        wBlock3 = Mid(wBlock3, debutBlock3, (finBlock3 - debutBlock3) + 1)
        wBlock3 = "{3:" & wBlock3 & "}"
    Else
        wBlock3 = ""
    End If
    
    '03/12/2018 DENIS. On garde le block3 dans tous les cas
'    'xxxx Modif du 04/11/2010 Denis
'    'wBlock3 = ""
'    If wBlock3 = "{3:{119:COV}}" Then
'        'on garde le block 3
'    Else
'        wBlock3 = ""
'    End If
'    'xxxx FIN modif 04/11/2010 Denis
    'FIN modif du 03/12/2018
Else
    wBlock2 = Mid$(lSWIALIDON, K2, K4 - K2)
    wBlock3 = ""
End If

Call SAA_from_SAB_Block2(wBlock2)                   'correction certains terminaux

' Uniquement si BDFEFRPPXXX et MT200
'=============================================================
kSG = InStr(wBlock2, "2:I200BDFEFRPP")
If kSG > 0 Then Mid$(wBlock2, kSG, 18) = "2:I200BDFEFRPPXNRO"

' Uniquement si USD : voir cmdZSWIALI0_Update_Transaction_RJE
'=============================================================
kSG = InStr(wBlock2, "SOGEFRPPXXXX")
If kSG > 0 Then
    kX1 = InStr(kSG, lSWIALIDON, ":32A:")
    If kX1 = 0 Then kX1 = InStr(kSG, lSWIALIDON, ":32B:")
    If kX1 > 0 Then
       kX2 = InStr(kX1 + 6, lSWIALIDON, ":")
       If kX2 > 0 Then
          kX0 = InStr(1, Mid$(lSWIALIDON, kX1, kX2 - kX1), "USD")
 '___________________________________________________________________________________________
'$JPL 2012-05-02 uniquement si MT 1** ou 2**
          'If kX0 > 0 Then Mid$(wBlock2, kSG, 12) = "SOGEFRPPXADB"
            If kX0 > 0 Then
                If InStr(wBlock2, "2:I1") Then
                    Mid$(wBlock2, kSG, 12) = "SOGEFRPPXADB"
                Else
                    If InStr(wBlock2, "2:I2") Then Mid$(wBlock2, kSG, 12) = "SOGEFRPPXADB"
                End If
            End If
'$JPL 2012-05-02 uniquement si MT 1** ou 2**
 '___________________________________________________________________________________________
        End If
    End If
End If
' Uniquement MT300 & MT320 SOGEFRPP:
'=============================================================
kSG = InStr(wBlock2, "2:I320SOGEFRPP")
If kSG > 0 Then Mid$(wBlock2, kSG, 18) = "2:I320SOGEFRPPXHCM"
kSG = InStr(wBlock2, "2:I300SOGEFRPP")
If kSG > 0 Then Mid$(wBlock2, kSG, 18) = "2:I300SOGEFRPPXHCM"     '$2007-07-18 JPL

'=============================================================
If newYSWIMON0.SWIMONFLUS = 1 Then
    Print #2, wBlock1 & wBlock2 & wBlock3 & "{4:"
Else
    Print #2, "$" & wBlock1 & wBlock2 & wBlock3 & "{4:"
End If

K1 = InStr(K4, lSWIALIDON, Asc13) + 2 '1 pour debug
blnExit = False
Do
    K2 = InStr(K1, lSWIALIDON, Asc13)
    If K2 > 0 Then
        X = Mid$(lSWIALIDON, K1, K2 - K1)
        K1 = K2 + 2 '1 pour debug
        'détecter un champ :57A: ou :57D: ou :71D: ou :72Z: ou :79Z: inclus dans un autre champ, notamment :32D: vide
        If InStr(X, ":57A:") > 0 Then
            K2 = InStr(X, ":57A:")
            X = Mid(X, K2)
        ElseIf InStr(X, ":57D:") > 0 Then
            K2 = InStr(X, ":57D:")
            X = Mid(X, K2)
        ElseIf InStr(X, ":71D:") > 0 Then
            K2 = InStr(X, ":71D:")
            X = Mid(X, K2)
        ElseIf InStr(X, ":72Z:") > 0 Then
            K2 = InStr(X, ":72Z:")
            X = Mid(X, K2)
        ElseIf InStr(X, ":79Z:") > 0 Then
            K2 = InStr(X, ":79Z:")
            X = Mid(X, K2)
        End If
        Select Case Mid$(X, 1, 4)
            Case ":20:": newYSWIMON0.SWIMONX20 = SAA_from_SAB_TRN(Mid$(X, 5, Len(X) - 4), wSWISABUnit) ' TRN modifié => SAA
                        X = ":20:" & newYSWIMON0.SWIMONX20 '''pour test : Mid$(X, 5, 4) = "TEST"
            Case ":21:": newYSWIMON0.SWIMONX21 = Mid$(X, 5, Len(X) - 4)
            Case ":32A": Call SAA_X32(X, newYSWIMON0.SWIMONX32V, newYSWIMON0.SWIMONX32D, newYSWIMON0.SWIMONX32A)
        End Select
        Print #2, X
    Else
        blnExit = True
    End If
Loop Until blnExit

Print #2, "-}"


GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : cmdZSWIALI0_Update_Transaction_Historique"
    
Exit_Function:
    On Error Resume Next
    cmdZSWIALI0_Update_Transaction_RJE = V

        
'=========================================================================
            ''Case Chr$(&HE9): Mid$(lSWIALIDON, K, 1) = Chr$(&H7B)
            ''Case Chr$(&HE8): Mid$(lSWIALIDON, K, 1) = Chr$(&H7D)

'=========================================================================

End Function

Public Sub SAA_from_SAB_Block2(wBlock2 As String)
'''pour test :Mid$(wBlock2, 8, 12) = paramBic8 & "XXXX"
Dim K As Integer

K = InStr(1, wBlock2, "è")
If K > 0 Then Mid$(wBlock2, K, 1) = "}"

Select Case Mid$(wBlock2, 1, 19)
    Case "{2:I754BEXADZALXXXX": Mid$(wBlock2, 1, 19) = "{2:I754BEXADZALXDOE"
    Case "{2:I202CRESCHZZXXXX": Mid$(wBlock2, 1, 19) = "{2:I202CRESCHZZX80A"
    Case "{2:I103CRESCHZZXXXX": Mid$(wBlock2, 1, 19) = "{2:I103CRESCHZZX80A"
    Case "{2:I103CRESCHZZXXXX": Mid$(wBlock2, 1, 19) = "{2:I103CRESCHZZX80A"
' Uniquement si USD : voir cmdZSWIALI0_Update_Transaction_RJE
'=============================================================
End Select


End Sub

Public Sub cmdZSWIHIA0_Insert()
Dim xSQL As String, Nb As Integer

xSQL = "Select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" & " where SWISABNUM = " & xZSWIHIA0.SWIHIANUM
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    MsgBox "Message déjà enregistré dans YSWIMON0 : " & rsSab("SWIMONID"), vbCritical, "frmSwift_Messages.mnuReprise_H_Click"
    Exit Sub
End If
Call ZSWICLA0_Sql(xZSWIHIA0.SWIHIANUM, wSWISABCOP, wSWISABDOS, wSWISABUnit)

srvYSWIMON0_Init newYSWIMON0

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
blnTransaction_Set
 
 V = sqlYSWIMON0_Init(newYSWIMON0, cnsab)
 If IsNull(V) Then
        newYSWIMON0.SWISABNUM = xZSWIHIA0.SWIHIANUM
        newYSWIMON0.SWISABCOP = wSWISABCOP
        newYSWIMON0.SWISABDOS = wSWISABDOS
        newYSWIMON0.SWIMONFLUX = "S"
        newYSWIMON0.SWIMONFLUD = xZSWIHIA0.SWIHIADEN + 19000000
        newYSWIMON0.SWIMONFLUH = xZSWIHIA0.SWIHIAHEN
        newYSWIMON0.SWIMONSTA = "S200"
        newYSWIMON0.SWIMONXMT = xZSWIHIA0.SWIHIAMES
        newYSWIMON0.SWIMONX20 = SAA_from_SAB_TRN(xZSWIHIA0.SWIHIAREF, wSWISABUnit) ' TRN modifié => SAA
        newYSWIMON0.SWIMONX32A = xZSWIHIA0.SWIHIAMON
        newYSWIMON0.SWIMONX32D = xZSWIHIA0.SWIHIADE1
        newYSWIMON0.SWIMONX32V = xZSWIHIA0.SWIHIADVA + 19000000
        
        
        '''xSql = "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
        '''Set rsSab_Update = cnsab.Execute(xSql, nb)
        
                V = sqlYSWIMON0_Insert(newYSWIMON0, cnsab)
        If Not IsNull(V) Then
            xSQL = "Rollback"
        Else
            xSQL = "Commit"
        End If
        
        Set rsSab_Update = cnsab.Execute(xSQL, Nb)
        '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
    End If
If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdZSWIHIA0_Insert_Auto()

Dim K As Long, I As Integer
Call lstErr_Clear(lstErr, cmdContext, "> cmdZSWIHIA0_Insert_Auto : " & fgZSWIHIA0.Rows - 1)
Call lstErr_AddItem(lstErr, cmdContext, "- cmdZSWIHIA0_Insert_Auto : " & fgZSWIHIA0.Rows - 1)

For I = 1 To fgZSWIHIA0.Rows - 1
    fgZSWIHIA0.Row = I
        fgZSWIHIA0.Col = fgZSWIHIA0_arrIndex:  K = CLng(fgZSWIHIA0.Text)
        xZSWIHIA0 = arrZSWIHIA0(K)
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "Insert : " & xZSWIHIA0.SWIHIAREF)
    cmdZSWIHIA0_Insert
Next I
Call lstErr_AddItem(lstErr, cmdContext, "< cmdZSWIHIA0_Insert_Auto : " & fgZSWIHIA0.Rows - 1)
End Sub

Public Function mnuStatus_Actualiser_YSWIMON0()
'me********_Status instance courante : rINst , rMesg , YSWIMON0
'--------------------------------------------------

Dim V
Dim X As String, xSQL As String
Dim Nb As Integer

On Error GoTo Error_Handler

'REchercher dans YSWIMON0 avec cles de l'instance
'--------------------------------------------------
V = Null

currentAction = "mnuStatus_Actualiser_YSWIMON0 : phase 1"
Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" _
     & " where SAAAID = " & merInst_Status.Aid _
     & " and SAAUMIDL = " & merInst_Status.inst_s_umidl _
     & " and SAAUMIDH = " & merInst_Status.inst_s_umidh

Set rsSab = cnsab.Execute(xSQL)
Nb = 0
Do While Not rsSab.EOF
    V = srvYSWIMON0_GetBuffer_ODBC(rsSab, oldYSWIMON0_Status)

    If Not IsNull(V) Then GoTo Error_MsgBox
    
    meYSWIMON0_Status = oldYSWIMON0_Status
    Nb = Nb + 1
    rsSab.MoveNext
Loop

'Tester : Doublon, pas de réponse (=> rapprocher) ou  réponse unique ?
'--------------------------------------------------------------------
If Nb > 1 Then V = " ! doublon : ": GoTo Error_MsgBox
If Nb = 0 Then
    V = mnuStatus_Actualiser_YSWIMON0_Rapprochement
    If Not IsNull(V) Then GoTo Exit_Function
End If

'réponse unique : Maj dans YSWIMON0 du statut du message
'--------------------------------------------------
currentAction = "mnuStatus_Actualiser_YSWIMON0 : phase 2"
V = mnuStatus_Actualiser_YSWIMON0_Statut
If Not IsNull(V) Then GoTo Exit_Function

'Maj SAB073 / YSWIMON0
'--------------------------------------------------

meYSWIMON0_Status.SWIMONSTAD = DSys
meYSWIMON0_Status.SWIMONSTAH = time_Hms

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
blnTransaction_Set

V = sqlYSWIMON0_Update(meYSWIMON0_Status, oldYSWIMON0_Status, cnsab)

If Not IsNull(V) Then
    xSQL = "Rollback"
Else
    xSQL = "Commit"
End If

Set rsSab_Update = cnsab.Execute(xSQL, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

GoTo Exit_Function

Error_Handler:
    V = Error
Error_MsgBox:
    ''MsgBox V, vbCritical, Me.Name & " : " & currentAction
    Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

Exit_Function:
   mnuStatus_Actualiser_YSWIMON0 = V
   
   
End Function

Public Function mnuStatus_Flag_YSWIMON0()
'me********_Status instance courante : rINst , rMesg , YSWIMON0
'--------------------------------------------------
Dim V
Dim X As String, xSQL As String
Dim Nb As Integer

On Error GoTo Error_Handler

'REchercher dans YSWIMON0 avec cles de l'instance
'--------------------------------------------------
V = Null

currentAction = "mnuStatus_Flag_YSWIMON0: phase 1"
Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" _
     & " where SAAAID = " & merIntv_Flag.Aid _
     & " and SAAUMIDL = " & merIntv_Flag.intv_s_umidl _
     & " and SAAUMIDH = " & merIntv_Flag.intv_s_umidh

Set rsSab = cnsab.Execute(xSQL)
Nb = 0
Do While Not rsSab.EOF
    V = srvYSWIMON0_GetBuffer_ODBC(rsSab, oldYSWIMON0_Status)

    If Not IsNull(V) Then GoTo Error_MsgBox
    
    meYSWIMON0_Status = oldYSWIMON0_Status
    Nb = Nb + 1
    rsSab.MoveNext
Loop

'Tester : Doublon, pas de réponse (=> rapprocher) ou  réponse unique ?
'--------------------------------------------------------------------
If Nb > 1 Then V = " ! doublon : ": GoTo Error_MsgBox
If Nb = 0 Then V = " ! non rapproché : ": GoTo Error_MsgBox

'réponse unique : Maj dans YSWIMON0 du statut du message
'--------------------------------------------------
currentAction = "mnuStatus_Actualiser_YSWIMON0 : phase 2"

Select Case Trim(arrrIntv_Flag(arrrIntv_Flag_Nb).intv_appl_serv_name)
   Case "Mesg Modification": meYSWIMON0_Status.SAAQMOD = 1

   Case "OFCA_Interface": meYSWIMON0_Status.SAAQOFAC = 1

End Select


'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
blnTransaction_Set

V = sqlYSWIMON0_Update(meYSWIMON0_Status, oldYSWIMON0_Status, cnsab)

If Not IsNull(V) Then
    xSQL = "Rollback"
Else
    xSQL = "Commit"
End If

Set rsSab_Update = cnsab.Execute(xSQL, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

GoTo Exit_Function

Error_Handler:
    V = Error
Error_MsgBox:
    ''MsgBox V, vbCritical, Me.Name & " : " & currentAction
    Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

Exit_Function:
   mnuStatus_Flag_YSWIMON0 = V
End Function


Public Function mnuStatus_Actualiser_YSWIMON0_Rapprochement()
'merInst_Status instance courante
'--------------------------------------------------
Dim V
Dim X As String, xSQL As String
Dim Nb As Integer, K1 As Integer, K2 As Integer, K3 As Integer
Dim xSWIMONFLUQ As String

On Error GoTo Error_Handler
V = Null


'Rechercher dans rMesg avec cles de l'instance
'--------------------------------------------------
currentAction = "mnuStatus_Actualiser_YSWIMON0_Rapprochement : phase 1"
Set rsSIDE_DB = Nothing
xSQL = "select * from rMesg" _
     & " where AID = " & merInst_Status.Aid _
     & " and mesg_s_UMIDL = " & merInst_Status.inst_s_umidl _
     & " and mesg_s_UMIDH = " & merInst_Status.inst_s_umidh

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)

Nb = 0
Do While Not rsSIDE_DB.EOF
    V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, merMesg_Status)
    If Not IsNull(V) Then GoTo Error_MsgBox
    Nb = Nb + 1
    rsSIDE_DB.MoveNext
Loop


'Tester : Doublon, pas de réponse (=> pb SIDE_DB) ou  réponse unique (OK)
'--------------------------------------------------------------------
If Nb > 1 Then V = " ! doublon : ": GoTo Error_MsgBox
If Nb = 0 Then
    V = "! manque rMesg !"
    GoTo Error_MsgBox
End If

'REchercher dans YSWIMON0 avec référence du flux
'--------------------------------------------------
currentAction = "mnuStatus_Actualiser_YSWIMON0_Rapprochement : phase 2"
If merMesg_Status.mesg_type = "950" Then V = "ignorer MT950": GoTo Exit_Function

Set rsSab = Nothing
X = merMesg_Status.mesg_batch_reference
K1 = InStr(1, X, "/"): K2 = 0: Nb = 0
If K1 > 0 Then K2 = InStr(K1 + 1, X, "/")
If K2 > 0 Then
    Nb = CInt(Mid$(X, K1 + 1, K2 - K1 - 1))
Else
    V = "! manque mesg_batch_reference : " & merMesg_Status.mesg_trn_ref
    GoTo Error_MsgBox
End If

K1 = InStr(1, X, "_")
K2 = InStr(K1 + 1, X, "_")
K3 = InStr(K2 + 1, X, ".")
xSWIMONFLUQ = Mid$(X, K2 + 1, K3 - K2 - 1)

If xSWIMONFLUQ = "MA" Or xSWIMONFLUQ = "MM" Or xSWIMONFLUQ = "SW" Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" _
         & " where SWIMONFLUQ = '" & xSWIMONFLUQ & "'" _
         & " and SWIMONFLUD = " & Mid$(X, 1, K1 - 1) _
         & " and SWIMONFLUH = " & Mid$(X, K1 + 1, K2 - K1 - 1) _
         & " and SWIMONFLUS = " & Nb
Else
    '2004.09.02 jpl rapprochement temporaire avec TYPE,TRN
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" _
         & " where SWIMONXMT = '" & merMesg_Status.mesg_type & "'" _
         & " and SWIMONX20 = '" & merMesg_Status.mesg_trn_ref & "'"
End If

Set rsSab = cnsab.Execute(xSQL)
Nb = 0
Do While Not rsSab.EOF
    V = srvYSWIMON0_GetBuffer_ODBC(rsSab, oldYSWIMON0_Status)

    If Not IsNull(V) Then GoTo Error_MsgBox
    
    meYSWIMON0_Status = oldYSWIMON0_Status
    Nb = Nb + 1
    rsSab.MoveNext
Loop

'Tester : Doublon, pas de réponse (=> rapprocher) ou  réponse unique ?
'--------------------------------------------------------------------
If Nb > 1 Then V = " ! doublon !" & merMesg_Status.mesg_trn_ref: GoTo Error_MsgBox
If Nb = 0 Then
    V = "! manque YSWIMON0 : " & merMesg_Status.mesg_trn_ref
    GoTo Error_MsgBox
End If

'réponse unique : Maj dans YSWIMON0 SAAAID .....
'--------------------------------------------------
meYSWIMON0_Status.SAAAID = merInst_Status.Aid
meYSWIMON0_Status.SAAUMIDL = merInst_Status.inst_s_umidl
meYSWIMON0_Status.SAAUMIDH = merInst_Status.inst_s_umidh
meYSWIMON0_Status.SAAQMOD = merMesg_Status.mesg_is_text_modified
meYSWIMON0_Status.SAAUNIT = merMesg_Status.x_inst0_unit_name


GoTo Exit_Function

Error_Handler:
    V = Error
Error_MsgBox:
    ''MsgBox V, vbCritical, Me.Name & " : " & currentAction
      Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

Exit_Function:
    mnuStatus_Actualiser_YSWIMON0_Rapprochement = V

End Function

Public Function mnuStatus_Actualiser_YSWIMON0_Fct(lFct As String, lAMJ As String, lHMS As Long)
Static sYSWIMON0_2 As typeYSWIMON0
Dim V, xSQL As String

V = Null
Select Case lFct
    Case "Select -2"
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" & " where  SWIMONID =  -2"
        Set rsSab = Nothing
        Set rsSab = cnsab.Execute(xSQL)
        V = srvYSWIMON0_GetBuffer_ODBC(rsSab, sYSWIMON0_2)
        
        If Not IsNull(V) Then GoTo Error_MsgBox
        lAMJ = sYSWIMON0_2.SWIMONSTAD
        lHMS = sYSWIMON0_2.SWIMONSTAH
        
     Case "Update -2"
        xYSWIMON0 = sYSWIMON0_2
        xYSWIMON0.SWIMONSTAD = lAMJ
        xYSWIMON0.SWIMONSTAH = lHMS
        V = sqlYSWIMON0_Update(xYSWIMON0, sYSWIMON0_2, cnsab)
        If Not IsNull(V) Then GoTo Exit_Function

        
    Case Else: V = "non programmé " & lFct
End Select
GoTo Exit_Function

Error_Handler:
    V = Error
Error_MsgBox:
    ''MsgBox V, vbCritical, Me.Name & " : mnuStatus_Actualiser_YSWIMON0_Fct"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

Exit_Function:
   mnuStatus_Actualiser_YSWIMON0_Fct = V

End Function

Public Function cmdZSWIFTA0_Update_Transaction(lSWIMONFLUQ As String, lQueue As String)
Dim wCommit_Rollback As String
Dim xSQL As String, nbSql As Long, xWhere As String, xSet As String
Dim V, I As Integer, Nb As Long
Dim wFile_Export As String, wFile_Log As String
Dim X As String, xFile As String


On Error GoTo Error_Handler

V = Null
wCommit_Rollback = "ROLLBACK"

srvYSWIMON0_Init newYSWIMON0
newYSWIMON0.SWIMONFLUD = DSys
newYSWIMON0.SWIMONFLUH = time_Hms


wFile_Log = newYSWIMON0.SWIMONFLUD & "_" & newYSWIMON0.SWIMONFLUH & lQueue & ".log"
X = paramSAA_DataF_from_SAB & wFile_Log
Call lstErr_Clear(lstErr, cmdContext, X)

Open X For Output As #1
currentAction = "1.1 réservation des messages (SWIFTAETA = -1)"
'=============================================================
Print #1, "1.0 : "; Time, currentAction
'----------------------------------------------
blnTransaction_Set
'=============================================================
Nb = 0
For arrZSWIFTA0_Index = 1 To arrZSWIFTA0_Nb
    If arrZSWIFTA0_SAA_Queue(arrZSWIFTA0_Index) = lQueue Then   ''' = paramSAA_Queue_Autorisation Then
            
        xZSWIFTA0 = arrZSWIFTA0(arrZSWIFTA0_Index)
        xSet = " set SWIFTAETA = -1"

        xWhere = " where SWIFTANUM = " & xZSWIFTA0.SWIFTANUM _
               & " and SWIFTANEN = '" & xZSWIFTA0.SWIFTANEN & "'" _
               & " and SWIFTAETA = " & xZSWIFTA0.SWIFTAETA
               
        xSQL = "update " & paramIBM_Library_SAB & ".ZSWIFTA0" & xSet & xWhere
        Call FEU_ROUGE
        Print #1, "1.1 : "; xSQL
        '----------------------------------------------
        Set rsSab = cnsab.Execute(xSQL, nbSql)
        Call FEU_VERT
        If nbSql = 0 Then
            V = "Erreur màj : " & xZSWIFTA0.SWIFTANUM
            GoTo Error_MsgBox
        End If
        Nb = Nb + 1
    End If
Next arrZSWIFTA0_Index

Print #1, "1.9 : "; Nb; " messages à traiter."
'----------------------------------------------
If Nb = 0 Then GoTo Exit_Function
'=============================================================
currentAction = "2 Incrémentation YSWIMON0.SWIMONNUM => SWIMONID "
'=============================================================
 
newYSWIMON0.SAAAID = 0
newYSWIMON0.SAAUMIDL = 0
newYSWIMON0.SAAUMIDH = 0
newYSWIMON0.SAAQUEUE = 0
newYSWIMON0.SAAQMOD = 0
newYSWIMON0.SAAQOFAC = 0
newYSWIMON0.SAAUNIT = 0
newYSWIMON0.SWIMONSTA = "S200"
newYSWIMON0.SWIMONFLUQ = lSWIMONFLUQ

newYSWIMON0.SWIMONFLUS = 0
newYSWIMON0.SWIMONFLUX = "S"

wFile_Export = newYSWIMON0.SWIMONFLUD & "_" & newYSWIMON0.SWIMONFLUH & "_" & newYSWIMON0.SWIMONFLUQ
xFile = paramSAA_DataF_from_SAB & wFile_Export & paramSAA_Data_from_SAB_ExtensionP_sab
Print #1, "2.0 : "; xFile
'----------------------------------------------
Call lstErr_AddItem(lstErr, cmdContext, wFile_Export)

Open xFile For Output As #2

For arrZSWIFTA0_Index = 1 To arrZSWIFTA0_Nb
    If arrZSWIFTA0_SAA_Queue(arrZSWIFTA0_Index) = lQueue Then
        xZSWIFTA0 = arrZSWIFTA0(arrZSWIFTA0_Index)
        Call lstErr_ChangeLastItem(lstErr, cmdContext, xZSWIFTA0.SWIFTANUM)
        xWhere = " where SWIFTBNUM = " & xZSWIFTA0.SWIFTANUM _
               & " and SWIFTBNEN = " & xZSWIFTA0.SWIFTANEN _
               & " and SWIFTBETA = " & xZSWIFTA0.SWIFTAETA
        Print #1, "2.1 : "; xWhere
        '----------------------------------------------

        arrZSWIFTB0_SQL xWhere, True
        Print #1, "2.2 : "; "cmdZSWIFTA0_Update_Transaction_YSWIMON0_Init"
        '----------------------------------------------
        
        V = cmdZSWIFTA0_Update_Transaction_YSWIMON0_Init
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "2.3 : "; "cmdZSWIFTA0_Update_Transaction_RJE"
        '----------------------------------------------
        V = cmdZSWIFTA0_Update_Transaction_RJE
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "2.4 : "; "sqlYSWIMON0_Insert newYSWIMON0 : "; newYSWIMON0.SWIMONID
        '----------------------------------------------
        V = sqlYSWIMON0_Insert(newYSWIMON0, cnsab)
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "2.5 : "; "cmdZSWIFTA0_Update_Transaction_Historique"
        '----------------------------------------------
        V = cmdZSWIFTA0_Update_Transaction_Historique
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "2.9 : "; Time; "-----------------------------------------------------------------"
        '----------------------------------------------

    End If
        
Next arrZSWIFTA0_Index
'=============================================================

wCommit_Rollback = "COMMIT"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    ''MsgBox V, vbCritical, Me.Name & " : cmdZSWIFTA0_Update_Transaction"
      Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

    Print #1, "ERREUR : "; Time, V
'----------------------------------------------

Exit_Function:
    On Error Resume Next
    cmdZSWIFTA0_Update_Transaction = V
    xSQL = wCommit_Rollback
    Call lstErr_AddItem(lstErr, cmdContext, xSQL)

    Print #1, "FIN : "; xSQL
    Set rsSab_Update = cnsab.Execute(xSQL, Nb)
    
'=============================================================
currentAction = "9 Archivage \\SWIFTPROD\....\Archive  & mise à disposition du fichier .sab en .rje "
'=====================================================================================================
  
    If Nb = 0 And wCommit_Rollback = "COMMIT" Then
        On Error GoTo END_Function
        Close #2
        xFile = paramSAA_DataF_from_SAB & wFile_Export & paramSAA_Data_from_SAB_ExtensionP_sab
        
        X = paramSAA_DataF_Archive & "\SAA_from_SAB_" & wFile_Export & ".sav"
        Print #1, "Archive : "; Time, X
        msFileSystem.CopyFile xFile, X
        
        X = paramSAA_DataF_from_SAB & wFile_Export & paramSAA_Data_from_SAB_ExtensionP_rje
        Print #1, "Rename : "; Time, X
        msFileSystem.MoveFile xFile, X
        
        xFile = paramSAA_DataF_from_SAB & wFile_Log
        X = paramSAA_DataF_Log & "\SAA_from_SAB_" & wFile_Log
        Print #1, "Archive Log : "; Time, X
        Print #1, "===== : "; Time, Nb
        Close #1
        
        msFileSystem.MoveFile xFile, X
    
    Else
        Print #1, "ERREUR : "; Time, Nb & "????????????????????????????????????????"
         Close
        X = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
            & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
            & htmlFontColor("BLUE") & "<BR><BR>" & "Plantage pendant le traitement des messages SWIFT : SAB => SAA" _
            & "<BR><BR>" & "Vérifier le contenu de la pièce jointe : " & wFile_Log

        Call Email_Alerte("ALERTE", "INFO", "BIA_SWIFT > frmSwift_Messages > cmdZSWIFTA0_Update_Transaction", X, True, paramSAA_DataF_from_SAB & wFile_Log)
   End If
    
    
END_Function:
   Close
'=============================================================

End Function

Public Function cmdZSWIALI0_Update_Transaction(lSWIMONFLUQ As String, lQueue As String)
Dim wCommit_Rollback As String
Dim xSQL As String, nbSql As Long, xWhere As String, xSet As String
Dim V, I As Integer, Nb As Long
Dim wFile_Export As String, wFile_Log As String
Dim X As String, xFile As String
Dim wSWIALIDON As String, lenSWIALIDON As Long
Dim blnExit As Boolean

On Error GoTo Error_Handler

V = Null
wCommit_Rollback = "ROLLBACK"

srvYSWIMON0_Init newYSWIMON0
newYSWIMON0.SWIMONFLUD = DSys
newYSWIMON0.SWIMONFLUH = time_Hms


wFile_Log = newYSWIMON0.SWIMONFLUD & "_" & newYSWIMON0.SWIMONFLUH & lQueue & ".log"
X = paramSAA_DataF_from_SAB & wFile_Log
Call lstErr_Clear(lstErr, cmdContext, X)

Open X For Output As #1
currentAction = "1.1 réservation des messages (SWIALIETA = -1)"
'=============================================================
Print #1, "1.0 : "; Time, currentAction
'----------------------------------------------
blnTransaction_Set
'=============================================================
Nb = 0
For arrZSWIALI0_Index = 1 To arrZSWIALI0_Nb
            
        xZSWIALI0 = arrZSWIALI0(arrZSWIALI0_Index)
        xSet = " set SWIALIETA = -1"

        xWhere = " where SWIALINUM = " & xZSWIALI0.SWIALINUM _
               & " and SWIALINEN = '" & xZSWIALI0.SWIALINEN & "'" _
               & " and SWIALINLI = " & xZSWIALI0.SWIALINLI _
               & " and SWIALIETA = " & xZSWIALI0.SWIALIETA
               
        xSQL = "update " & paramIBM_Library_SAB & ".ZSWIALI0" & xSet & xWhere
                Call FEU_ROUGE
            Print #1, "1.1 : "; xSQL
        '----------------------------------------------
        Set rsSab = cnsab.Execute(xSQL, nbSql)
                Call FEU_VERT

        If nbSql = 0 Then
            V = "Erreur màj : " & xZSWIALI0.SWIALINUM
            GoTo Error_MsgBox
        End If
        Nb = Nb + 1
Next arrZSWIALI0_Index

Print #1, "1.9 : "; Nb; " messages à traiter."
'----------------------------------------------
If Nb = 0 Then GoTo Exit_Function
'=============================================================
currentAction = "2 Incrémentation YSWIMON0.SWIMONNUM => SWIMONID "
'=============================================================
 
newYSWIMON0.SAAAID = 0
newYSWIMON0.SAAUMIDL = 0
newYSWIMON0.SAAUMIDH = 0
newYSWIMON0.SAAQUEUE = 0
newYSWIMON0.SAAQMOD = 0
newYSWIMON0.SAAQOFAC = 0
newYSWIMON0.SAAUNIT = 0
newYSWIMON0.SWIMONSTA = "S200"
newYSWIMON0.SWIMONFLUQ = lSWIMONFLUQ

newYSWIMON0.SWIMONFLUS = 0
newYSWIMON0.SWIMONFLUX = "S"

wFile_Export = newYSWIMON0.SWIMONFLUD & "_" & newYSWIMON0.SWIMONFLUH & "_" & newYSWIMON0.SWIMONFLUQ
xFile = paramSAA_DataF_from_SAB & wFile_Export & paramSAA_Data_from_SAB_ExtensionP_sab
Print #1, "2.0 : "; xFile
'----------------------------------------------
Call lstErr_AddItem(lstErr, cmdContext, wFile_Export)

Open xFile For Output As #2

For arrZSWIALI0_Index = 1 To arrZSWIALI0_Nb
    If arrZSWIALI0(arrZSWIALI0_Index).SWIALINLI = 1 Then
        xZSWIALI0 = arrZSWIALI0(arrZSWIALI0_Index)
        Call lstErr_ChangeLastItem(lstErr, cmdContext, xZSWIALI0.SWIALINUM)
        xWhere = " where SWIFTBNUM = " & xZSWIALI0.SWIALINUM _
               & " and SWIFTBNEN = " & xZSWIALI0.SWIALINEN _
               & " and SWIFTBETA = " & xZSWIALI0.SWIALIETA
        Print #1, "2.1 : SWIALINUM : "; xZSWIALI0.SWIALINUM, arrZSWIALI0_Index
        '----------------------------------------------
        wSWIALIDON = ""
        blnExit = False
        Do
            wSWIALIDON = wSWIALIDON & Trim(arrZSWIALI0(arrZSWIALI0_Index).SWIALIDON)
            lenSWIALIDON = Len(wSWIALIDON)
            If Mid$(wSWIALIDON, lenSWIALIDON, 1) = Asc03 Then
                blnExit = True
            Else
                arrZSWIALI0_Index = arrZSWIALI0_Index + 1
                Print #1, "2.1 : SWIALINUM + "; xZSWIALI0.SWIALINUM, arrZSWIALI0_Index
        '----------------------------------------------

                If arrZSWIALI0_Index > arrZSWIALI0_Nb Then
                    V = "Manque HEX '03' fin de message. Erreur Index > Nb : " & arrZSWIALI0_Index & " > " & arrZSWIALI0_Nb
                    GoTo Error_MsgBox
                Else
                    If xZSWIALI0.SWIALINUM <> arrZSWIALI0(arrZSWIALI0_Index).SWIALINUM Then
                        V = "Manque HEX '03' fin de message : " & xZSWIALI0.SWIALINUM
                        GoTo Error_MsgBox
                    End If
                End If
            End If
                            
        Loop Until blnExit
        Print #1, "2.2 : "; "cmdZSWIALI0_Update_Transaction_YSWIMON0_Init"
        '----------------------------------------------
        
        V = cmdZSWIALI0_Update_Transaction_YSWIMON0_Init
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "2.3 : "; "cmdZSWIALI0_Update_Transaction_RJE"
        '----------------------------------------------
        V = cmdZSWIALI0_Update_Transaction_RJE(wSWIALIDON, lenSWIALIDON)
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "2.4 : "; "sqlYSWIMON0_Insert newYSWIMON0 : "; newYSWIMON0.SWIMONID
        '----------------------------------------------
        V = sqlYSWIMON0_Insert(newYSWIMON0, cnsab)
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "2.5 : "; "cmdZSWIALI0_Update_Transaction_Historique"
        '----------------------------------------------
        V = cmdZSWIALI0_Update_Transaction_Historique
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "2.9 : "; Time; "-----------------------------------------------------------------"
        '----------------------------------------------

    End If
        
Next arrZSWIALI0_Index
'=============================================================

wCommit_Rollback = "COMMIT"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    ''MsgBox V, vbCritical, Me.Name & " : cmdZSWIALI0_Update_Transaction"
      Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

    Print #1, "ERREUR : "; Time, V
'----------------------------------------------

Exit_Function:
    On Error Resume Next
    cmdZSWIALI0_Update_Transaction = V
    xSQL = wCommit_Rollback
    Call lstErr_AddItem(lstErr, cmdContext, xSQL)

    Print #1, "FIN : "; xSQL
    Set rsSab_Update = cnsab.Execute(xSQL, Nb)
    
'=============================================================
currentAction = "9 Archivage \\SWIFTPROD\....\Archive  & mise à disposition du fichier .sab en .rje "
'=====================================================================================================
  
    If Nb = 0 And wCommit_Rollback = "COMMIT" Then
        On Error GoTo END_Function
        Close #2
        xFile = paramSAA_DataF_from_SAB & wFile_Export & paramSAA_Data_from_SAB_ExtensionP_sab
        
        X = paramSAA_DataF_Archive & "\SAA_from_SAB_" & wFile_Export & ".sav"
        Print #1, "Archive : "; Time, X
        msFileSystem.CopyFile xFile, X
        
        X = paramSAA_DataF_from_SAB & wFile_Export & paramSAA_Data_from_SAB_ExtensionP_rje
        Print #1, "Rename : "; Time, X
        msFileSystem.MoveFile xFile, X
        
        xFile = paramSAA_DataF_from_SAB & wFile_Log
        X = paramSAA_DataF_Log & "\SAA_from_SAB_" & wFile_Log
        Print #1, "Archive Log : "; Time, X
        Print #1, "===== : "; Time, Nb
        Close #1
        
        msFileSystem.MoveFile xFile, X
    
    Else
        Print #1, "ERREUR : "; Time, Nb & "????????????????????????????????????????"
        Close
        X = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
            & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
            & htmlFontColor("BLUE") & "<BR><BR>" & "Plantage pendant le traitement des messages SWIFT : SAB => SAA" _
            & "<BR><BR>" & "Vérifier le contenu de la pièce jointe : " & wFile_Log

        Call Email_Alerte("ALERTE", "INFO", "BIA_SWIFT > frmSwift_Messages > cmdZSWIALI0_Update_Transaction", X, True, paramSAA_DataF_from_SAB & wFile_Log)

    End If
    
    
END_Function:
   Close
'=============================================================

End Function


Public Function cmdZSWIHIA0_Restauration()
Dim wCommit_Rollback As String
Dim xSQL As String, nbSql As Long, xWhere As String, xSet As String
Dim V, I As Integer, Nb As Long
Dim wFile_Export As String, wFile_Log As String
Dim X As String, xFile As String

On Error GoTo Error_Handler

V = Null
wCommit_Rollback = "ROLLBACK"

srvYSWIMON0_Init newYSWIMON0
newYSWIMON0.SWIMONFLUD = DSys
newYSWIMON0.SWIMONFLUH = time_Hms

wFile_Log = newYSWIMON0.SWIMONFLUD & "_" & newYSWIMON0.SWIMONFLUH & "_BIA_Swift_Restauration.log"
X = paramSAA_DataF_from_SAB & wFile_Log
Call lstErr_Clear(lstErr, cmdContext, X)
Call FEU_ROUGE
Open X For Output As #1
currentAction = "1.1 Restauration"
'=============================================================
Print #1, "1.0 : "; Time, currentAction
'----------------------------------------------
blnTransaction_Set
'''xSql = "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
'''Set rsSab_Update = cnsab.Execute(xSql, nb)
'=============================================================

        xSQL = "delete from " & paramIBM_Library_SABSPE & ".ySWIMON0 where SWISABNUM = " & xZSWIHIA0.SWIHIANUM
         Print #1, "1.1 : "; xSQL
        '----------------------------------------------
       Set rsSab = cnsab.Execute(xSQL, nbSql)

        
        Print #1, "1.5 : "; "cmdZSWIHIA0_Restauration_Historique"
        '----------------------------------------------
        V = cmdZSWIHIA0_Restauration_Historique
        If Not IsNull(V) Then GoTo Error_MsgBox
        Print #1, "1.9 : "; Time; "-----------------------------------------------------------------"
        '----------------------------------------------

        
'=============================================================

wCommit_Rollback = "COMMIT"
GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " : cmdZSWIHIA0_Restauration"
    Print #1, "ERREUR : "; Time, V
'----------------------------------------------

Exit_Function:
    On Error Resume Next
    cmdZSWIHIA0_Restauration = V
    xSQL = wCommit_Rollback
    Call lstErr_AddItem(lstErr, cmdContext, xSQL)

    Print #1, "FIN : "; xSQL
    Set rsSab_Update = cnsab.Execute(xSQL, Nb)
    
'=====================================================================================================
  
    If Nb = 0 And wCommit_Rollback = "COMMIT" Then
        Print #1, "===== : "; Time, Nb
        
        xFile = paramSAA_DataF_from_SAB & wFile_Log
        X = paramSAA_DataF_Log & "\SAA_from_SAB_" & wFile_Log
        Print #1, "Archive Log : "; Time, X
        Print #1, "===== : "; Time, Nb
        Close #1
        
        msFileSystem.MoveFile xFile, X
    
    Else
        Print #1, "ERREUR : "; Time, Nb & "????????????????????????????????????????"
    End If
    
    
END_Function:
   Close
Call FEU_VERT
'=============================================================

End Function


Public Function cmdZSWIHIA0_Restauration_Historique()
Dim xSQL As String, nbSql As Long, xWhere As String, xSet As String
Dim V
On Error GoTo Error_Handler

V = Null

'ZSWIHIA0 => ZSWIFTA0  :  date et heure du tranfert =0
'=========================================================================================
xZSWIHIA0.SWIHIADEN = 0
xZSWIHIA0.SWIHIAHEN = 0

V = srvZSWIFTA0_Sql_Restauration(xZSWIHIA0, xSQL)
If Not IsNull(V) Then GoTo Error_MsgBox
Call FEU_ROUGE
Print #1, "1.5.1 : "; xSQL
'----------------------------------------------
Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Insert ZSWIFTA0 : " & xZSWIHIA0.SWIHIANUM
    GoTo Error_MsgBox
End If
xWhere = " where SWIHIANUM = " & xZSWIHIA0.SWIHIANUM _
       & " and SWIHIANEN = '" & xZSWIHIA0.SWIHIANEN & "'" _
       & " and SWIHIAETA = " & xZSWIHIA0.SWIHIAETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWIHIA0" & xWhere
Print #1, "1.5.2 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWIHIA0 : " & xZSWIHIA0.SWIHIANUM
    GoTo Error_MsgBox
End If

'ZSWIHIB0 => ZSWIFTB0
'====================================================================================================

 xWhere = " where SWIHIBNUM = " & xZSWIHIA0.SWIHIANUM _
        & " and SWIHIBNEN = " & xZSWIHIA0.SWIHIANEN _
        & " and SWIHIBETA = " & xZSWIHIA0.SWIHIAETA
Print #1, "1.5.3 : "; "ZSWIHIB0 => ZSWIFTB0 : " & xWhere
'----------------------------------------------
arrZSWIHIB0_SQL xWhere

For I = 1 To arrZSWIHIB0_Nb
    
    V = srvZSWIFTB0_Sql_Restauration(arrZSWIHIB0(I), xSQL)
    If Not IsNull(V) Then GoTo Error_MsgBox
    Set rsSab = cnsab.Execute(xSQL, nbSql)
    
    If nbSql = 0 Then
        V = "Erreur Insert ZSWIFTB0 : " & xZSWIHIB0.SWIHIBNUM & " " & xZSWIHIB0.SWIHIBNEN & " " & xZSWIHIB0.SWIHIBNLI
        GoTo Error_MsgBox
    End If
Next I

xWhere = " where SWIHIBNUM = " & xZSWIHIA0.SWIHIANUM _
       & " and SWIHIBNEN = " & xZSWIHIA0.SWIHIANEN _
       & " and SWIHIBETA = " & xZSWIHIA0.SWIHIAETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWIHIB0" & xWhere
Print #1, "1.5.4 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWIHIB0 : " & xZSWIHIB0.SWIHIBNUM
    GoTo Error_MsgBox
End If

'ZSWIHIC0 => ZSWIFTC0
'==============================================================================
 xWhere = " where SWIHICNUM = " & xZSWIHIA0.SWIHIANUM _
        & " and SWIHICNEN = " & xZSWIHIA0.SWIHIANEN _
        & " and SWIHICETA = " & xZSWIHIA0.SWIHIAETA
Print #1, "1.5.5 : "; "ZSWIHIC0 => ZSWIFTC0 : " & xWhere
'----------------------------------------------
arrZSWIHIC0_SQL xWhere

For I = 1 To arrZSWIHIC0_Nb
    
    V = srvZSWIFTC0_Sql_Restauration(arrZSWIHIC0(I), xSQL)
    If Not IsNull(V) Then GoTo Error_MsgBox
    Set rsSab = cnsab.Execute(xSQL, nbSql)
    
    If nbSql = 0 Then
        V = "Erreur Insert ZSWIFTC0 : " & xZSWIHIC0.SWIHICNUM & " " & xZSWIHIC0.SWIHICNEN & " " & xZSWIHIC0.SWIHICNLI
        GoTo Error_MsgBox
    End If
Next I

xWhere = " where SWIHICNUM = " & xZSWIHIA0.SWIHIANUM _
       & " and SWIHICNEN = " & xZSWIHIA0.SWIHIANEN _
       & " and SWIHICETA = " & xZSWIHIA0.SWIHIAETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWIHIC0" & xWhere
Print #1, "1.5.6 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWIHIC0 : " & xZSWIHIC0.SWIHICNUM
    GoTo Error_MsgBox
End If


'ZSWIHIT0 => ZSWITEM0
'====================

 xWhere = " where SWIHITNUM = " & xZSWIHIA0.SWIHIANUM _
        & " and SWIHITNEN = " & xZSWIHIA0.SWIHIANEN _
        & " and SWIHITETA = " & xZSWIHIA0.SWIHIAETA
Print #1, "1.5.7 : "; "ZSWIHIT0 => ZSWITEM0 : " & xWhere
'----------------------------------------------
arrZSWIHIT0_SQL xWhere


For I = 1 To arrZSWIHIT0_Nb
    
    V = srvZSWITEM0_Sql_Restauration(arrZSWIHIT0(I), xSQL)
    If Not IsNull(V) Then GoTo Error_MsgBox
    Set rsSab = cnsab.Execute(xSQL, nbSql)
    
    If nbSql = 0 Then
        V = "Erreur Insert ZSWITEM0 : " & xZSWIHIT0.SWIHITNUM & " " & xZSWIHIT0.SWIHITNEN & " " & xZSWIHIT0.SWIHITCON
        GoTo Error_MsgBox
    End If
Next I

xWhere = " where SWIHITNUM = " & xZSWIHIA0.SWIHIANUM _
       & " and SWIHITNEN = " & xZSWIHIA0.SWIHIANEN _
       & " and SWIHITETA = " & xZSWIHIA0.SWIHIAETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWIHIT0" & xWhere
Print #1, "1.5.8 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWIHIT0 : " & xZSWIHIT0.SWIHITNUM
    GoTo Error_MsgBox
End If

GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : cmdZSWIHIA0_Restauration_Historique"
    
Exit_Function:
    On Error Resume Next
    cmdZSWIHIA0_Restauration_Historique = V
    Call FEU_VERT
    
End Function



Public Function cmdZSWIFTA0_Update_Transaction_YSWIMON0_Init()
Dim xSQL As String
Dim wSWISABCOP As String, wSWISABDOS As Long
Dim V
V = Null

xSQL = "Select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" & " where SWISABNUM = " & xZSWIFTA0.SWIFTANUM
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    V = "Message déjà enregistré dans YSWIMON0.SWISABNUM : " & xZSWIFTA0.SWIFTANUM
    GoTo Exit_Function
End If
Call ZSWICLA0_Sql(xZSWIFTA0.SWIFTANUM, wSWISABCOP, wSWISABDOS, wSWISABUnit)

 V = sqlYSWIMON0_Init(newYSWIMON0, cnsab)
 If IsNull(V) Then
        newYSWIMON0.SWISABNUM = xZSWIFTA0.SWIFTANUM
        newYSWIMON0.SWISABCOP = wSWISABCOP
        newYSWIMON0.SWISABDOS = wSWISABDOS

        newYSWIMON0.SWIMONFLUS = newYSWIMON0.SWIMONFLUS + 1
        newYSWIMON0.SWIMONXMT = xZSWIFTA0.SWIFTAMES
        newYSWIMON0.SWIMONX20 = SAA_from_SAB_TRN(xZSWIFTA0.SWIFTAREF, wSWISABUnit) ' TRN modifié => SAA
        newYSWIMON0.SWIMONX21 = ""
        newYSWIMON0.SWIMONX32A = xZSWIFTA0.SWIFTAMON
        newYSWIMON0.SWIMONX32D = xZSWIFTA0.SWIFTADE1
        newYSWIMON0.SWIMONX32V = xZSWIFTA0.SWIFTADVA + 19000000
    End If
Exit_Function:
    cmdZSWIFTA0_Update_Transaction_YSWIMON0_Init = V
End Function
Public Function cmdZSWIALI0_Update_Transaction_YSWIMON0_Init()
Dim xSQL As String
Dim V
V = Null

xSQL = "Select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" & " where SWISABNUM = " & xZSWIALI0.SWIALINUM
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    V = "Message déjà enregistré dans YSWIMON0.SWISABNUM : " & xZSWIALI0.SWIALINUM
    GoTo Exit_Function
End If
'20041116 jpl         xSql = "Select SWICLAOPR,SWICLANUM from " & paramIBM_Library_SAB & ".ZSWICLA0" & " where SWICLAINT = " & xZSWIALI0.SWIALINUM
'20041116 jpl         Set rsSab = cnsab.Execute(xSql)
'20041116 jpl         If Not rsSab.EOF Then
'20041116 jpl             wSWISABCOP = rsSab("SWICLAOPR"): wSWISABDOS = rsSab("SWICLANUM")
'20041116 jpl             If wSWISABCOP = "CDE" Or wSWISABCOP = "CDI" Then wSWISABDOS = wSWISABDOS / 10000
'20041116 jpl             If wSWISABCOP = "RDE" Or wSWISABCOP = "RDI" Then wSWISABDOS = wSWISABDOS / 100
'20041116 jpl         Else
'20041116 jpl             wSWISABCOP = "": wSWISABDOS = 0
'20041116 jpl         End If

Call ZSWICLA0_Sql(xZSWIALI0.SWIALINUM, wSWISABCOP, wSWISABDOS, wSWISABUnit)

 V = sqlYSWIMON0_Init(newYSWIMON0, cnsab)
 If IsNull(V) Then
        newYSWIMON0.SWISABNUM = xZSWIALI0.SWIALINUM
        newYSWIMON0.SWISABCOP = wSWISABCOP
        newYSWIMON0.SWISABDOS = wSWISABDOS

        newYSWIMON0.SWIMONFLUS = newYSWIMON0.SWIMONFLUS + 1
        newYSWIMON0.SWIMONXMT = xZSWIALI0.SWIALIMES
        newYSWIMON0.SWIMONX20 = ""              'SAA_from_SAB_TRN(xZSWIALI0.SWIALIREF)  ' TRN modifié => SAA
        newYSWIMON0.SWIMONX21 = ""
        newYSWIMON0.SWIMONX32A = 0           'xZSWIALI0.SWIALIMON
        newYSWIMON0.SWIMONX32D = ""             'xZSWIALI0.SWIALIDE1
        newYSWIMON0.SWIMONX32V = 0            'xZSWIALI0.SWIALIDVA + 19000000
    End If
Exit_Function:
    cmdZSWIALI0_Update_Transaction_YSWIMON0_Init = V
End Function


Public Function cmdZSWIFTA0_Update_Transaction_Historique()
Dim xSQL As String, nbSql As Long, xWhere As String, xSet As String
Dim V
On Error GoTo Error_Handler

V = Null

'ZSWIFTA0 => ZSWIHIA0  :  date et heure du tranfert
'=====================================================
xZSWIFTA0.SWIFTADEN = newYSWIMON0.SWIMONFLUD - 19000000
xZSWIFTA0.SWIFTAHEN = newYSWIMON0.SWIMONFLUH

V = srvZSWIHIA0_Sql_Sauvegarde(xZSWIFTA0, xSQL)
If Not IsNull(V) Then GoTo Error_MsgBox
Call FEU_ROUGE
Print #1, "2.5.1 : "; xSQL
'----------------------------------------------
Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Insert ZSWIHIA0 : " & xZSWIFTA0.SWIFTANUM
    GoTo Error_MsgBox
End If
xWhere = " where SWIFTANUM = " & xZSWIFTA0.SWIFTANUM _
       & " and SWIFTANEN = '" & xZSWIFTA0.SWIFTANEN & "'" _
       & " and SWIFTAETA = -1                                   "
                            '!!!!!!!! " & xZSWIFTA0.SWIFTAETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWIFTA0" & xWhere
Print #1, "2.5.2 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWIFTA0 : " & xZSWIFTA0.SWIFTANUM
    GoTo Error_MsgBox
End If

'ZSWIFTB0 => ZSWIHIB0
'====================
Print #1, "2.5.3 : "; "ZSWIFTB0 => ZSWIHIB0"
'----------------------------------------------
For I = 1 To arrZSWIFTB0_Nb
    
    V = srvZSWIHIB0_Sql_Sauvegarde(arrZSWIFTB0(I), xSQL)
    If Not IsNull(V) Then GoTo Error_MsgBox
    Set rsSab = cnsab.Execute(xSQL, nbSql)
    
    If nbSql = 0 Then
        V = "Erreur Insert ZSWIHIB0 : " & xZSWIFTB0.SWIFTBNUM & " " & xZSWIFTB0.SWIFTBNEN & " " & xZSWIFTB0.SWIFTBNLI
        GoTo Error_MsgBox
    End If
Next I

xWhere = " where SWIFTBNUM = " & xZSWIFTA0.SWIFTANUM _
       & " and SWIFTBNEN = " & xZSWIFTA0.SWIFTANEN _
       & " and SWIFTBETA = " & xZSWIFTA0.SWIFTAETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWIFTB0" & xWhere
Print #1, "2.5.4 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWIFTB0 : " & xZSWIFTB0.SWIFTBNUM
    GoTo Error_MsgBox
End If

'ZSWIFTC0 => ZSWIHIC0
'====================
Print #1, "2.5.5 : "; "ZSWIFTC0 => ZSWIHIC0"
'----------------------------------------------
For I = 1 To arrZSWIFTC0_Nb
    
    V = srvZSWIHIC0_Sql_Sauvegarde(arrZSWIFTC0(I), xSQL)
    If Not IsNull(V) Then GoTo Error_MsgBox
    Set rsSab = cnsab.Execute(xSQL, nbSql)
    
    If nbSql = 0 Then
        V = "Erreur Insert ZSWIHIC0 : " & xZSWIFTC0.SWIFTCNUM & " " & xZSWIFTC0.SWIFTCNEN & " " & xZSWIFTC0.SWIFTCNLI
        GoTo Error_MsgBox
    End If
Next I

xWhere = " where SWIFTCNUM = " & xZSWIFTA0.SWIFTANUM _
       & " and SWIFTCNEN = " & xZSWIFTA0.SWIFTANEN _
       & " and SWIFTCETA = " & xZSWIFTA0.SWIFTAETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWIFTC0" & xWhere
Print #1, "2.5.6 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWIFTC0 : " & xZSWIFTC0.SWIFTCNUM
    GoTo Error_MsgBox
End If


'ZSWITEM0 => ZSWIHIT0
'====================
Print #1, "2.5.7 : "; "ZSWITEM0 => ZSWIHIT00"
'----------------------------------------------

 xWhere = " where SWITEMNUM = " & xZSWIFTA0.SWIFTANUM _
        & " and SWITEMNEN = " & xZSWIFTA0.SWIFTANEN _
        & " and SWITEMETA = " & xZSWIFTA0.SWIFTAETA
arrZSWITEM0_SQL xWhere


For I = 1 To arrZSWITEM0_Nb
    
    V = srvZSWIHIT0_Sql_Sauvegarde(arrZSWITEM0(I), xSQL)
    If Not IsNull(V) Then GoTo Error_MsgBox
    Set rsSab = cnsab.Execute(xSQL, nbSql)
    
    If nbSql = 0 Then
        V = "Erreur Insert ZSWIHIT0 : " & xZSWITEM0.SWITEMNUM & " " & xZSWITEM0.SWITEMNEN & " " & xZSWITEM0.SWITEMCON
        GoTo Error_MsgBox
    End If
Next I

xWhere = " where SWItemnUM = " & xZSWIFTA0.SWIFTANUM _
       & " and SWItemNEN = " & xZSWIFTA0.SWIFTANEN _
       & " and SWItemETA = " & xZSWIFTA0.SWIFTAETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWITEM0" & xWhere
Print #1, "2.5.8 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWITEM0 : " & xZSWITEM0.SWITEMNUM
    GoTo Error_MsgBox
End If

GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : cmdZSWIFTA0_Update_Transaction_Historique"
      Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
Exit_Function:
    On Error Resume Next
    cmdZSWIFTA0_Update_Transaction_Historique = V
    Call FEU_VERT
    
End Function
Public Function cmdZSWIALI0_Update_Transaction_Historique()
Dim xSQL As String, nbSql As Long, xWhere As String, xSet As String
Dim V
On Error GoTo Error_Handler

V = Null

xWhere = " where SWIALINUM = " & xZSWIALI0.SWIALINUM _
       & " and SWIALINEN = '" & xZSWIALI0.SWIALINEN & "'" _
       & " and SWIALIETA = -1                                   "
                            '!!!!!!!! " & xZSWIALI0.SWIALIETA
       
xSQL = "delete from " & paramIBM_Library_SAB & ".ZSWIALI0" & xWhere
Call FEU_ROUGE
Print #1, "2.5.2 : "; xSQL
'----------------------------------------------

Set rsSab = cnsab.Execute(xSQL, nbSql)

If nbSql = 0 Then
    V = "Erreur Delete ZSWIALI0 : " & xZSWIALI0.SWIALINUM
    GoTo Error_MsgBox
End If


GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : cmdZSWIALI0_Update_Transaction_Historique"
      Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
Exit_Function:
    On Error Resume Next
    cmdZSWIALI0_Update_Transaction_Historique = V
    Call FEU_VERT
    
End Function

Public Function mnuStatus_Actualiser_YSWIMON0_Statut()
'merInst_Status instance courante
'--------------------------------------------------
Dim V
Dim X As String, xSQL As String
Dim blnCompleted As Boolean

' Toujours 1 seule instance dont les valeurs des zones utilisées changent
  
On Error GoTo Error_Handler
V = Null

blnCompleted = False

meYSWIMON0_Status.SAAQUEUE = merInst_Status.inst_rp_name
If merInst_Status.inst_mpfn_name = "OFCA_Check" Then meYSWIMON0_Status.SAAQOFAC = 1
        
Select Case merInst_Status.inst_status

    ' Message en LIVE
    Case "LIVE":
        Select Case Trim(meYSWIMON0_Status.SAAQUEUE)
        
            Case "OFCS_Validate": meYSWIMON0_Status.SWIMONSTA = "S210"
            Case "_MP_mod_text": meYSWIMON0_Status.SWIMONSTA = "S220"
            Case "_MP_authorisation": meYSWIMON0_Status.SWIMONSTA = "S230"
            Case "_MP_mod_transmis": meYSWIMON0_Status.SWIMONSTA = "S240"
            
            Case "_MP_verification": meYSWIMON0_Status.SWIMONSTA = "S250"
            Case "_MP_mod_reception": meYSWIMON0_Status.SWIMONSTA = "S260"
            
            Case "_SI_to_SWIFT": meYSWIMON0_Status.SWIMONSTA = "S800"
            Case Else: meYSWIMON0_Status.SWIMONSTA = "S888"
            
        End Select
    
    ' Message COMPLETED - Voir si SUPPRIME ou ACK ou NAK ou REJECTED
    Case "COMPLETED":
        If Trim(merInst_Status.inst_mpfn_name) = "mpm" And _
           Trim(merInst_Status.inst_auth_oper_nickname) = "" And _
           Trim(merInst_Status.x_last_emi_appe_date_time) = "00:00:00" And _
           merInst_Status.x_last_emi_appe_seq_nbr = 0 Then
           
           meYSWIMON0_Status.SWIMONSTA = "S904"  ' Message supprimé
        Else
        
            blnCompleted = True
        End If
End Select

If blnCompleted Then    ' ACK ou NAK ou REJECTED ??

    'Rechercher dans rAppe avec cles de l'instance : 1 seul enreg ACK lu
    '----------------------------------------------
    currentAction = "mnuStatus_Actualiser_YSWIMON0_Statut"

    Set rsSIDE_DB = Nothing
    xSQL = "select * from rAppe where appe_iapp_name = 'SWIFT' and " _
                & "appe_inst_num = 0" _
                & " and Aid = " & merInst_Status.Aid _
                & " and appe_s_umidl = " & merInst_Status.inst_s_umidl _
                & " and appe_s_umidh  = " & merInst_Status.inst_s_umidh _
                & " order by appe_date_time, appe_seq_nbr "
                
                
 '                 & " and (appe_network_delivery_status = 'DLV_ACKED' or " _
 '               & "appe_network_delivery_status = 'DLV_NACKED' or " _
 '               & "appe_network_delivery_status = 'DLV_REJECTED_LOCALLY')" _

    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
    Do While Not rsSIDE_DB.EOF
        V = srvrAppe_GetBuffer_ODBC(rsSIDE_DB, merAppe_Status)
    
         If Not IsNull(V) Then ' Positionner un flag pour dire ...
            V = "Message ni supprimé, ni ACK, ni NAK, ni REJECTED... ??": GoTo Error_MsgBox
         Else
            Select Case merAppe_Status.appe_network_delivery_status
            
                Case "DLV_ACKED": meYSWIMON0_Status.SWIMONSTA = "S901"
                ' Case "DLV_NACKED": meYSWIMON0_Status.SWIMONSTA = "S220"
                ' Case "DLV_REJECTED_LOCALLY":
                '
                '    Select Case Trim(meYSWIMON0_Status.SAAQUEUE)
                '
                '        Case "_OFAC_Validate": meYSWIMON0_Status.SWIMONSTA = "S210"
                '        Case "_MP_mod_text": meYSWIMON0_Status.SWIMONSTA = "S220"
                '        Case "_MP_authorisation": meYSWIMON0_Status.SWIMONSTA = "S230"
                '
                '        Case Else: meYSWIMON0_Status.SWIMONSTA = "S903"
                '    End Select
                Case Else:
                    If Mid$(merAppe_Status.appe_network_delivery_status, 1, 4) = "DLV_" Then meYSWIMON0_Status.SWIMONSTA = "S905" '2006-12-22 JPL "S700"
            End Select
            
         End If
         rsSIDE_DB.MoveNext
    Loop
End If      ' Fin test sur blnCompleted
    
     
GoTo Exit_Function



Error_Handler:
    V = Error
Error_MsgBox:
    ''MsgBox V, vbCritical, Me.Name & " : " & currentAction
      Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

Exit_Function:
    mnuStatus_Actualiser_YSWIMON0_Statut = V

End Function





Public Sub blnTransaction_Set()
If Not blnTransaction Then
    blnTransaction = True
    Set rsSab_Update = cnsab.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub

Public Sub cmdZSWIFTA0_Update_Transaction_Queue()
Dim blnSAA_Queue_Autorisation As Boolean
Dim blnSAA_Queue_Modification As Boolean
Dim blnSAA_Queue_Swift As Boolean

blnSAA_Queue_Autorisation = False
blnSAA_Queue_Modification = False
blnSAA_Queue_Swift = False

For arrZSWIFTA0_Index = 1 To arrZSWIFTA0_Nb
    Select Case arrZSWIFTA0_SAA_Queue(arrZSWIFTA0_Index)
        Case paramSAA_Queue_Autorisation: blnSAA_Queue_Autorisation = True
        Case paramSAA_Queue_Modification: blnSAA_Queue_Modification = True
        Case paramSAA_Queue_SWIFT: blnSAA_Queue_Swift = True
    End Select
Next arrZSWIFTA0_Index

If blnSAA_Queue_Autorisation Then cmdZSWIFTA0_Update_Transaction "MA", paramSAA_Queue_Autorisation
If blnSAA_Queue_Modification Then cmdZSWIFTA0_Update_Transaction "MM", paramSAA_Queue_Modification
If blnSAA_Queue_Swift Then cmdZSWIFTA0_Update_Transaction "SW", paramSAA_Queue_SWIFT

End Sub

Private Sub txtSelect_Utilisateur_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Public Sub ZSWICLA0_Sql(lNum As Long, lSWISABCOP As String, lSWISABDOS As Long, lSWISABUnit As String)
Dim X As String, xSQL As String
Dim Vdec
' Rechercher le code opération et numéro dossier
lSWISABCOP = ""
lSWISABDOS = 0
lSWISABUnit = ""
xSQL = "Select SWICLAOPR,SWICLANUM from " & paramIBM_Library_SAB & ".ZSWICLA0" & " where SWICLAINT = " & lNum
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    lSWISABCOP = rsSab("SWICLAOPR")
    Vdec = CDec(rsSab("SWICLANUM"))
    If lSWISABCOP = "CDE" Or lSWISABCOP = "CDI" Then lSWISABDOS = CLng(Vdec / 10000)
    
    If lSWISABCOP = "RDE" Or lSWISABCOP = "RDI" Then
        lSWISABDOS = CLng(Vdec / 100)
        lSWISABUnit = Table_Ope_Unit_RDE(lSWISABCOP, lSWISABDOS, cnsab, rsSab)
    End If
End If

End Sub


Public Sub cmdImport_File_Exe(lFile As String)
Dim Seq As Long, xIn As String
Dim blnOk As Boolean
Open lFile For Input As #1
Open Trim(txtImport_File_Out) For Output As #2
Seq = 0
blnOk = False
Call lstErr_AddItem(lstErr, cmdContext, "Lecture : ")

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
        If Mid$(xIn, 1, 13) = "U-UMID      =" Then
            blnOk = True
            Seq = Seq + 1
            Call lstErr_ChangeLastItem(lstErr, cmdContext, Mid$(xIn, 15, 11) & Seq)
        End If
    If Mid$(xIn, 1, 17) = "Message History =" Then blnOk = False
    
    If blnOk Then Print #2, xIn
Loop
Close

End Sub
Public Sub cmdSaa_CB_Exe()
Dim Seq As Long, xIn As String
Dim blnOk As Boolean

Open Trim(txtSAA_CB) & "_Mesg.txt" For Output As #2
Open Trim(txtSAA_CB) & "_Block4.txt" For Output As #4

Seq = 0
blnOk = False
Call lstErr_AddItem(lstErr, cmdContext, "Lecture : ")

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
        If Mid$(xIn, 1, 13) = "U-UMID      =" Then
            blnOk = True
            Seq = Seq + 1
            Call lstErr_ChangeLastItem(lstErr, cmdContext, Mid$(xIn, 15, 11) & Seq)
        End If
    If Mid$(xIn, 1, 17) = "Message History =" Then blnOk = False
    
    Print #2, xIn
Loop

Close

End Sub


Public Function cmdSAA_SQL_08_Unit(lUnit As String) As Integer
Select Case lUnit
    Case "BOTC": cmdSAA_SQL_08_Unit = 1
    Case "CSOP": cmdSAA_SQL_08_Unit = 2
    Case "DAFI": cmdSAA_SQL_08_Unit = 3
    Case "DCOM": cmdSAA_SQL_08_Unit = 4
    Case "ORPA": cmdSAA_SQL_08_Unit = 5
    Case "SCLE": cmdSAA_SQL_08_Unit = 6
    Case "SOBF": cmdSAA_SQL_08_Unit = 7
    Case "SOBI": cmdSAA_SQL_08_Unit = 8
    Case "None": cmdSAA_SQL_08_Unit = 10
    Case Else: cmdSAA_SQL_08_Unit = 9
End Select

End Function

