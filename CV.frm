VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCV 
   Caption         =   "Contre-valeur + Euro"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9390
   Begin VB.Timer timerControl 
      Interval        =   200
      Left            =   5160
      Top             =   0
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   9000
      Picture         =   "CV.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   0
      Width           =   400
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5760
      TabIndex        =   56
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton cmdContext 
      Caption         =   "Calcul"
      Height          =   300
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   0
      Width           =   1200
   End
   Begin VB.PictureBox picCompta 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00E0E0E0&
      Height          =   2000
      Left            =   50
      ScaleHeight     =   1935
      ScaleWidth      =   9240
      TabIndex        =   52
      Top             =   4320
      Width           =   9300
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   6
      Top             =   320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Contre_Valeur (Cours  T.C.)"
      TabPicture(0)   =   "CV.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraDevise"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picCV"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Comptabilité"
      TabPicture(1)   =   "CV.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCRI"
      Tab(1).Control(1)=   "cmdOk"
      Tab(1).Control(2)=   "fraCrédit"
      Tab(1).Control(3)=   "fraDébit"
      Tab(1).Control(4)=   "libNUMPIE"
      Tab(1).ControlCount=   5
      Begin VB.Frame fraCRI 
         Caption         =   "CRI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -67080
         TabIndex        =   60
         Top             =   1080
         Width           =   1335
         Begin VB.OptionButton optCriSG2 
            Caption         =   "SG 550.02.9"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton optCriUbaf 
            Caption         =   "UBAF"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   2280
            Width           =   855
         End
         Begin VB.OptionButton optCriBdf 
            Caption         =   "BDF"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1680
            Width           =   1095
         End
         Begin VB.OptionButton optCriSG 
            Caption         =   "SG 550.01.1"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ok"
         Height          =   525
         Left            =   -67080
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   480
         Width           =   1260
      End
      Begin VB.Frame fraCrédit 
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
         Height          =   1815
         Left            =   -74880
         TabIndex        =   47
         Top             =   2040
         Width           =   7695
         Begin VB.CheckBox chkCréditAvis 
            Caption         =   "imprimer un avis"
            Height          =   255
            Left            =   2760
            TabIndex        =   65
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtCréditCompte 
            Height          =   285
            Left            =   1080
            TabIndex        =   41
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtCréditLibellé 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   42
            Top             =   900
            Width           =   6495
         End
         Begin MSMask.MaskEdBox txtCréditDval 
            Height          =   300
            Left            =   1080
            TabIndex        =   43
            Top             =   1300
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "## - ## - ####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblCréditCompte 
            Caption         =   "Compte"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   615
         End
         Begin VB.Label libCréditCompte 
            Caption         =   "-"
            Height          =   255
            Left            =   2880
            TabIndex        =   50
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label lblCréditLibellé 
            Caption         =   "Libellé"
            Height          =   255
            Left            =   1080
            TabIndex        =   49
            Top             =   600
            Width           =   6495
         End
         Begin VB.Label lblCréditDval 
            Caption         =   "Date Valeur"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1300
            Width           =   855
         End
      End
      Begin VB.Frame fraDébit 
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
         Height          =   1700
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   7695
         Begin VB.CheckBox chkDébitAvis 
            Caption         =   "imprimer un avis"
            Height          =   255
            Left            =   2760
            TabIndex        =   64
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtDébitLibellé 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   39
            Top             =   960
            Width           =   6495
         End
         Begin VB.TextBox txtDébitCompte 
            Height          =   285
            Left            =   1080
            TabIndex        =   38
            Top             =   240
            Width           =   1575
         End
         Begin MSMask.MaskEdBox txtDébitDval 
            Height          =   300
            Left            =   1080
            TabIndex        =   40
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "## - ## - ####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblDébitDval 
            Caption         =   "Date Valeur"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1305
            Width           =   855
         End
         Begin VB.Label lblDébitLibellé 
            BackColor       =   &H8000000A&
            Caption         =   "Libellé"
            Height          =   200
            Left            =   1080
            TabIndex        =   45
            Top             =   600
            Width           =   6255
         End
         Begin VB.Label libDébitCompte 
            Caption         =   "-"
            Height          =   255
            Left            =   2880
            TabIndex        =   44
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label lblDébitCompte 
            Caption         =   "Compte"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picCV 
         AutoRedraw      =   -1  'True
         FillColor       =   &H00E0E0E0&
         Height          =   675
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   9060
         TabIndex        =   35
         Top             =   3240
         Width           =   9120
      End
      Begin VB.Frame FraDevise 
         ForeColor       =   &H00C00000&
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9120
         Begin VB.ListBox lstDevise 
            Height          =   2400
            Left            =   2160
            TabIndex        =   67
            Top             =   360
            Width           =   3015
         End
         Begin VB.Frame fraDevise2 
            Height          =   1095
            Left            =   50
            TabIndex        =   22
            Top             =   1150
            Width           =   7320
            Begin VB.CheckBox chkDevise2 
               Caption         =   "montant manuel"
               Height          =   250
               Left            =   2880
               TabIndex        =   4
               Top             =   720
               Width           =   1695
            End
            Begin VB.Frame fraDevise2Sens 
               Height          =   500
               Left            =   2880
               TabIndex        =   26
               Top             =   120
               Width           =   1700
               Begin VB.OptionButton optDevise2Cr 
                  Caption         =   "Crédit"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   28
                  Top             =   200
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton optDevise2Db 
                  Caption         =   "Débit"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   27
                  Top             =   200
                  Width           =   735
               End
            End
            Begin VB.Frame fraDevise2BME 
               Height          =   500
               Left            =   4920
               TabIndex        =   23
               Top             =   120
               Width           =   2175
               Begin VB.OptionButton optDevise2EnCompte 
                  Caption         =   "Compte"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   25
                  Top             =   200
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton optDevise2Espèces 
                  Caption         =   "Espèces"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   24
                  Top             =   200
                  Width           =   975
               End
            End
            Begin VB.TextBox txtDevise2Montant 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   120
               TabIndex        =   3
               Top             =   240
               Width           =   1600
            End
            Begin VB.TextBox txtDevise2 
               Height          =   285
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   2
               Top             =   240
               Width           =   500
            End
            Begin VB.Label lblDevise2montant 
               Alignment       =   1  'Right Justify
               Caption         =   "xxxx"
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
               Left            =   120
               TabIndex        =   29
               Top             =   720
               Width           =   1605
            End
         End
         Begin VB.Frame FraDevise1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   50
            TabIndex        =   15
            Top             =   400
            Width           =   7320
            Begin VB.Frame fraDevise1BME 
               Height          =   500
               Left            =   4920
               TabIndex        =   19
               Top             =   120
               Width           =   2175
               Begin VB.OptionButton optDevise1Espèces 
                  Caption         =   "Espèces"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   21
                  Top             =   200
                  Width           =   975
               End
               Begin VB.OptionButton optDevise1EnCompte 
                  Caption         =   "Compte"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   20
                  Top             =   200
                  Value           =   -1  'True
                  Width           =   975
               End
            End
            Begin VB.Frame fraDevise1Sens 
               Height          =   500
               Left            =   2880
               TabIndex        =   16
               Top             =   120
               Width           =   1700
               Begin VB.OptionButton optDevise1Db 
                  Caption         =   "Débit "
                  Height          =   255
                  Left            =   120
                  TabIndex        =   18
                  Top             =   200
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton optDevise1Cr 
                  Caption         =   "Crédit "
                  Height          =   255
                  Left            =   840
                  TabIndex        =   17
                  Top             =   200
                  Width           =   735
               End
            End
            Begin VB.TextBox txtDevise1Montant 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   120
               TabIndex        =   0
               Top             =   240
               Width           =   1600
            End
            Begin VB.TextBox txtDevise1 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   1
               Top             =   240
               Width           =   500
            End
         End
         Begin VB.TextBox txtDevise3Montant 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   2400
            Width           =   1600
         End
         Begin VB.TextBox txtDevise3 
            Height          =   285
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   13
            Top             =   2400
            Width           =   500
         End
         Begin VB.Frame FraDeviseCours 
            Caption         =   "Cours T.C."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   7440
            TabIndex        =   8
            Top             =   120
            Width           =   1545
            Begin VB.OptionButton optVirCompte 
               Caption         =   "Vir Cpt / Cpt"
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   2280
               Width           =   1335
            End
            Begin VB.OptionButton optCriAller 
               Caption         =   "CRI Aller"
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   2000
               Width           =   1335
            End
            Begin VB.OptionButton optCriRetour 
               Caption         =   "CRI Retour"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   1680
               Width           =   1335
            End
            Begin VB.OptionButton optPivot 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Pivot"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optEnCompte 
               Caption         =   "En Compte"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   1320
               Value           =   -1  'True
               Width           =   1150
            End
            Begin VB.OptionButton optPrivilégié 
               Caption         =   "BME Privilégié"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton optNormal 
               Caption         =   "BME Normal"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   600
               Width           =   1215
            End
         End
         Begin MSMask.MaskEdBox txtAmj 
            Height          =   300
            Left            =   3840
            TabIndex        =   5
            Top             =   2400
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "## - ## - ####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblCours2_1 
            Caption         =   "x"
            Height          =   255
            Left            =   5640
            TabIndex        =   34
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lblCours1_2 
            Caption         =   "x"
            Height          =   255
            Left            =   5640
            TabIndex        =   33
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label lblMontant 
            Caption         =   "Montant"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   150
            Width           =   855
         End
         Begin VB.Label lblDevise 
            Caption         =   "Devise"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   31
            Top             =   150
            Width           =   735
         End
         Begin VB.Label lblOpéAMJ 
            Caption         =   "Date cours"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   30
            Top             =   2400
            Width           =   975
         End
      End
      Begin VB.Label libNUMPIE 
         Caption         =   "-"
         Height          =   255
         Left            =   -66840
         TabIndex        =   54
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Menu mnuDevise 
      Caption         =   "mnuDevise"
      Visible         =   0   'False
      Begin VB.Menu mnuDeviseDisplay 
         Caption         =   "Détail devise"
      End
      Begin VB.Menu mnuDevisePrint 
         Caption         =   "imprimer les cours"
      End
   End
End
Attribute VB_Name = "frmCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor

Dim blnMsgBox_Quit As Boolean
Dim CVAut As typeAuthorization
Dim Msg As String
Dim valAMJ As String, valAMJ1 As String, valAMJ2 As String

Dim CV As typeCV
Dim CV1 As typeCV, CV2 As typeCV, CV3 As typeCV
Dim wCV1 As typeCV, wCV2 As typeCV

Dim recCompte As typeCompte
Dim recCréditCompte As typeCompte
Dim recDébitCompte As typeCompte
Dim mDval1 As String * 8, mDval2 As String * 8, mDval3 As String * 8
Dim Sens1 As String, Sens2 As String, Conversion As String

Dim arrCV030(6) As typeCpj030W0
Dim arrCV030Nb As Integer
Dim mAMJP1 As String * 8, mAMJN2 As String * 8
Dim blnTimer As Boolean

Dim mDébitDval As String * 8, mCréditDval As String * 8, X8 As String * 8
Dim blnDébitCompte As Boolean, blnCréditCompte As Boolean
Dim mDébitCompte As String * 11, mCréditCompte As String * 11

Dim blnDébitLibellé As Boolean, blnCréditlibellé As Boolean
Dim curX As Currency
Dim mCriCompte As String
Dim recOpCpt As typeOpCpt
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
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
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
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

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
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Private Sub chkCréditAvis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCréditAvis

End Sub


Private Sub chkDébitAvis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDébitAvis

End Sub


Private Sub chkDevise2_Click()
timerControl.Enabled = True

End Sub

Private Sub chkDevise2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDevise2

End Sub


Private Sub cmdContext_Click()
timerControl.Enabled = True
End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub


Private Sub cmdOk_Click()
Dim I As Integer, wDEVISE As String * 4, wNUMLOT As String * 4, wNUMPIE As String * 7
Dim intNOLIGN As Integer
Dim mIMPHMS As String * 6

Me.Enabled = False

lstErr.Clear
Cv_Compta

If lstErr.ListCount = 0 Then
    wDEVISE = "0"
    wNUMLOT = "0000"
    wNUMPIE = "0000000"
    intNOLIGN = 1
    mIMPHMS = time_Hms
    For I = 1 To arrCV030Nb
        If arrCV030(I).MONDEV <> 0 Then
            If Val(wDEVISE) = Val(arrCV030(I).Devise) Then
                arrCV030(I).NUMPIE = wNUMPIE
                intNOLIGN = intNOLIGN + 1
            Else
                intNOLIGN = 1
            End If
            arrCV030(I).NUMLOT = wNUMLOT
            arrCV030(I).NOLIGN = Format$(intNOLIGN, "0000")
            arrCV030(I).CTLSTA = "1"
            arrCV030(I).IMPAMJ = DSys
            arrCV030(I).IMPHMS = mIMPHMS
            
            srvCpj030W0_Update arrCV030(I)
            
            wDEVISE = arrCV030(I).Devise
            wNUMLOT = arrCV030(I).NUMLOT
            wNUMPIE = arrCV030(I).NUMPIE
    End If
   Next I
End If

cmdPrintX constDemandeDeValidation

If chkCréditAvis = "1" Then
    If arrCV030(1).SENECR = "C" Then
        prtCv_Avis 1
    Else
        prtCv_Avis 3
    End If
End If

If chkDébitAvis = "1" Then
    If arrCV030(1).SENECR = "D" Then
        prtCv_Avis 1
    Else
        prtCv_Avis 3
    End If
End If
Cv_Compta_Clear
libNUMPIE = arrCV030(1).NUMPIE
SSTab1.Tab = 0

    Me.Enabled = True
    AppActivate Me.Caption


End Sub


Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOk

End Sub


Private Sub cmdPrint_Click()
timerControl.Enabled = True
cmdPrintX ""

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  SendKeys "{TAB}" 'timerControl.Enabled = True '
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select
End Sub



'---------------------------------------------------------
Public Sub CV_Cours_Display()
'---------------------------------------------------------
Dim X As String, XX As String, strAv As String

'picCV.FontBold = True

picCV.ForeColor = libUsr.ForeColor

picCV.CurrentY = picCV.CurrentY + 300
picCV.CurrentX = 50: picCV.Print Format$(CV.DeviseN, "000") & "   " & CV.DeviseLibellé;

picCV.ForeColor = warnUsrColor
picCV.CurrentX = 1800
If CV.EuroIn Then
    picCV.Print "Devise IN";
Else
    picCV.Print "Devise OUT";
End If

picCV.CurrentX = 3000
'If CV.CotationCertain Then
'    picCV.Print "Cotation au certain";
'Else
'    picCV.Print "Cotation à l'incertain";
'End If
strAv = ""
If CV.Normal <> " " Then
    Select Case CV.AchatVente
        Case "A": picCV.Print "(Achat) ";: strAv = "  (Vente)"
        Case "V": picCV.Print "(Vente) ";: strAv = "  (Achat)"
    End Select
End If

picCV.ForeColor = libUsr.ForeColor
picCV.CurrentX = 3600
If CV.CotationCertain Then
    picCV.Print CV3.DeviseIso & "  /  " & CV.DeviseIso;
Else
    picCV.Print CV.DeviseIso & "  /  " & CV3.DeviseIso;
End If
picCV.ForeColor = warnUsrColor
picCV.Print strAv;
picCV.ForeColor = libUsr.ForeColor

X = Format$(CV.Cours, "## ##0.00 000 00")
picCV.CurrentX = 6500 - picCV.TextWidth(X)
picCV.Print X;
       
picCV.ForeColor = warnUsrColor
picCV.CurrentX = 6800

Select Case CV.Normal
    Case "N": picCV.Print "Normal ";
    Case "P": picCV.Print "Privilégié ";
    Case "C": picCV.Print "en Compte ";
    Case Else: picCV.Print "Cours Pivot ";
End Select

picCV.CurrentX = 8000
picCV.ForeColor = libUsr.ForeColor
picCV.Print dateImp(CV.CoursAmj);
End Sub

'---------------------------------------------------------
Public Sub CV_Compta_Display()
'---------------------------------------------------------
Dim X As String, I As Integer
Dim mCurrentX1 As Integer, mForeColor1 As Long
Dim mCurrentX2 As Integer, mForeColor2 As Long
Dim curTotal As Currency

DoEvents: picCompta.Cls
picCompta.ForeColor = libUsr.ForeColor
picCompta.Line (0, 600)-(9300, 600)
picCompta.Line (0, 1200)-(9300, 1200)
picCompta.CurrentY = 50
curTotal = 0
For I = 1 To arrCV030Nb
    If arrCV030(I).MONDEV <> 0 Then

        If arrCV030(I).SENECR = "D" Then
            curTotal = curTotal - arrCV030(I).MONDEV
            mCurrentX1 = 8000: mForeColor1 = errUsr.ForeColor
            mCurrentX2 = 9200: mForeColor2 = libUsr.ForeColor
        Else
            curTotal = curTotal + arrCV030(I).MONDEV
            mCurrentX2 = 8000: mForeColor2 = errUsr.ForeColor
            mCurrentX1 = 9200: mForeColor1 = libUsr.ForeColor
        End If
    
        picCompta.FontBold = False
        
        picCompta.ForeColor = libUsr.ForeColor
        
        picCompta.CurrentX = 50: picCompta.Print Format$(arrCV030(I).Devise, "000") & "." & Compte_Imp(arrCV030(I).Compte);
        
        picCompta.FontBold = False
        If Val(arrCV030(I).Compte) = 0 Then
            If CVAut.Saisir Then Call lstErr_AddItem(lstErr, picCompta, "? compte à préciser")
        Else
            picCompta.CurrentX = 1600
            recCompte.Method = "SeekL1"
            recCompte.Société = arrCV030(I).COSOC
            recCompte.Agence = arrCV030(I).Agence
            recCompte.Devise = Format$(Val(arrCV030(I).Devise), "000")
            recCompte.Numéro = arrCV030(I).Compte
            If IsNull(srvCompteFind(recCompte)) Then
                 If recCompte.Situation <> " " And recCompte.Situation <> "E" And recCompte.Situation <> "F" Then
                    picCompta.ForeColor = errUsr.ForeColor
                    picCompta.Print Trim(recCompte.Intitulé) & "Annulé/Bloqué";
                    Call lstErr_AddItem(lstErr, picCompta, "? compte  annulé/bloqué")
                Else
                    If recCompte.TypeGA = "A" And recCompte.BiaTyp <> "001" _
                                              And recCompte.BiaTyp <> "002" And recCompte.BiaTyp <> "550" _
                                              And recCompte.BiaTyp <> "028" And recCompte.BiaTyp <> "010" _
                                              And recCompte.BiaTyp <> "013" And recCompte.BiaTyp <> "025" Then
                        picCompta.ForeColor = errUsr.ForeColor
                        Call lstErr_AddItem(lstErr, picCompta, "? type  = 001 ou 002 ou 550 ou 028 ou 010 ou 13 ou 25")
                    End If
                    If recCompte.TypeGA = "G" And recCompte.Numéro >= 90000000 Then
                        picCompta.ForeColor = errUsr.ForeColor
                        Call lstErr_AddItem(lstErr, picCompta, "? compte Hors-Bilan ")
                    End If
                    picCompta.Print Trim(recCompte.Intitulé);
               End If
            Else
                picCompta.ForeColor = errUsr.ForeColor
                picCompta.Print "???????";
                Call lstErr_AddItem(lstErr, picCompta, "? compte inconnu")
            End If
        End If
        
        If picCompta.CurrentX < 5500 Then picCompta.CurrentX = 5500
        picCompta.ForeColor = warnUsrColor
        picCompta.Print dateImp(arrCV030(I).AMJVAL);
        
        picCompta.FontBold = True
        picCompta.ForeColor = mForeColor1
        X = Format$(arrCV030(I).MONDEV, "### ### ### ### ##0.00")
        
        picCompta.CurrentX = mCurrentX1 - picCompta.TextWidth(X)
        picCompta.Print X;
        picCompta.CurrentY = picCompta.CurrentY + 300
    End If
Next I
If curTotal <> 0 Then Call lstErr_AddItem(lstErr, picCompta, "? pièce non équilibrée")
End Sub



Public Sub cmdContext_Quit()
If lstDevise.Visible Then
    lstDevise.Visible = False
Else
    Unload Me
End If
End Sub


Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)


mDébitDval = "00000000"
mCréditDval = "00000000"
blnTimer = False: timerControl.Enabled = False

fraCRI.Visible = False
optCriSG = True: mCriCompte = constCriCompteSG

'txtDevise2Montant.Enabled = False
txtDevise3Montant.Enabled = False
txtDevise3.Enabled = False
lblDevise2montant.ForeColor = errUsr.ForeColor
lblDevise2montant = ""
fraDevise1BME.Enabled = False
fraDevise2BME.Enabled = False
lblCours1_2 = "": lblCours1_2.ForeColor = errUsr.ForeColor
lblCours2_1 = "": lblCours2_1.ForeColor = errUsr.ForeColor
dbDeviseChange_Replication
Cv_Compta_Clear

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraCrédit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDébit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub FraDevise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise1BME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise1Sens_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise2BME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise2Sens_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub FraDeviseCours_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub lblCréditCompte_Click()
currentActiveControl_Name = "txtCréditCompte"
frmCompte_Click

End Sub

Private Sub lblCréditCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lblCréditCompte
End Sub


Private Sub lblDébitCompte_Click()
currentActiveControl_Name = "txtDébitCompte"
frmCompte_Click
End Sub

Private Sub lblDébitCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lblDébitCompte
End Sub


Private Sub lblDevise_Click()
lstDevise.Visible = True
Call LstDictio(889, lstDevise)

End Sub

Private Sub lblDevise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lblDevise

End Sub

Private Sub mnuDeviseDisplay_Click()
If lstDevise.ListIndex < 0 Then lstDevise.ListIndex = 1
    Set XListBox = Me.lstDevise
    frmDevise.Show vbModal
'End If
End Sub

Private Sub lstDevise_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
        Me.PopupMenu mnuDevise, vbPopupMenuRightButton
Else
    Select Case currentActiveControl_Name
        Case "txtDevise1": txtDevise1 = mId$(lstDevise.Text, 1, 3): If txtDevise1.Enabled Then txtDevise1.SetFocus
        Case "txtDevise2": txtDevise2 = mId$(lstDevise.Text, 1, 3): If txtDevise2.Enabled Then txtDevise2.SetFocus
    End Select
    timerControl.Enabled = True
End If
End Sub


Private Sub mnuDevisePrint_Click()
Dim Msg As String
Msg = Space$(50)
Msg = Format$(1, "000000") & Format$(999, "000000") & DSys & "TI"

prtDeviseX Msg

End Sub

Private Sub optCriAller_Click()
If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
chkDébitAvis = "1"
timerControl.Enabled = True

End Sub

Private Sub optCriAller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optCriAller
End Sub


Private Sub optCriBdf_Click()
mCriCompte = constCriCompteBDF
timerControl.Enabled = True
End Sub

Private Sub optCriRetour_Click()
If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
chkCréditAvis = "1"
timerControl.Enabled = True

End Sub

Private Sub optCriRetour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optCriRetour
End Sub


Private Sub optCriSG_Click()
mCriCompte = constCriCompteSG
timerControl.Enabled = True

End Sub

Private Sub optCriSG2_Click()
mCriCompte = constCriCompteSG2
timerControl.Enabled = True

End Sub


Private Sub optCriUbaf_Click()
mCriCompte = constCriCompteUBAF
timerControl.Enabled = True

End Sub


Private Sub optDevise1Cr_Click()
optDevise2Db = True
timerControl.Enabled = True
End Sub

Private Sub optDevise1Cr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1Cr

End Sub


Private Sub optDevise1Db_Click()
optDevise2Cr = True
timerControl.Enabled = True
End Sub


Private Sub optDevise1Db_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1Db

End Sub


Private Sub optDevise1EnCompte_Click()
timerControl.Enabled = True
End Sub

Private Sub optDevise1EnCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1EnCompte

End Sub


Private Sub optDevise1Espèces_Click()
timerControl.Enabled = True

End Sub


Private Sub optDevise1Espèces_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1Espèces

End Sub


Private Sub optDevise2Cr_Click()
optDevise1Db = True
'timerControl.Enabled = True

End Sub

Private Sub optDevise2Cr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2Cr

End Sub


Private Sub optDevise2Db_Click()
optDevise1Cr = True
'timerControl.Enabled = True
End Sub


Private Sub optDevise2Db_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2Db

End Sub


Private Sub optDevise2EnCompte_Click()
timerControl.Enabled = True

End Sub

Private Sub optDevise2EnCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2EnCompte

End Sub


Private Sub optDevise2Espèces_Click()
timerControl.Enabled = True

End Sub


Private Sub optDevise2Espèces_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2Espèces

End Sub


Private Sub optEnCompte_Click()
If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
timerControl.Enabled = True
End Sub


Private Sub optEnCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEnCompte

End Sub


Private Sub optNormal_Click()
If optDevise1EnCompte And optDevise2EnCompte Then optDevise1Espèces = True
If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
timerControl.Enabled = True
End Sub


Private Sub optNormal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optNormal

End Sub


Private Sub optPivot_Click()
If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
timerControl.Enabled = True
End Sub


Private Sub optPivot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPivot

End Sub


Private Sub optPrivilégié_Click()
If optDevise1EnCompte And optDevise2EnCompte Then optDevise1Espèces = True
If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
timerControl.Enabled = True
End Sub


Private Sub optPrivilégié_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPrivilégié

End Sub


Private Sub optVirCompte_Click()
If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
timerControl.Enabled = True

End Sub

Private Sub optVirCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optVirCompte
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
Else
    If fraDébit.Visible And txtDébitCompte.Enabled Then
        If txtDébitCompte.Enabled Then txtDébitCompte.SetFocus
    Else
        If fraCrédit.Visible And txtCréditCompte.Enabled Then txtCréditCompte.SetFocus
    End If
End If

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1

End Sub


Private Sub timerControl_Timer()
If Not blnTimer Then
    timerControl.Enabled = False
    cmdControl
End If

End Sub

Private Sub txtAmj_Change()
form_Clear
End Sub

Private Sub txtCréditCompte_GotFocus()
Call txt_GotFocus(txtCréditCompte)

End Sub


Private Sub txtCréditCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtCréditCompte_LostFocus()
txt_LostFocus txtCréditCompte
txtCréditCompte_Control
If blnCréditCompte Then
    If Not optVirCompte And Trim(txtDébitCompte) = "" Then txtDébitCompte = txtCréditCompte ': txtDébitCompte_Control
End If
timerControl.Enabled = True 'Cv_Compta

End Sub

Private Sub txtCréditDval_GotFocus()
Call txt_GotFocus(txtCréditDval)

End Sub


Private Sub txtCréditDval_LostFocus()
Dim X As String, X2 As String
Call txt_LostFocus(txtCréditDval)
lstErr.Clear

mCréditDval = "00000000"
X = dateCtl(txtCréditDval.Text)
If Not IsNumeric(X) Then
    txtCréditDval.ForeColor = errUsr.ForeColor
    Call lstErr_Clear(lstErr, txtCréditDval, "erreur date")
Else
    X2 = mId$(X, 1, 8)
    If X2 <> "00000000" Then
    
        X = vbYes
        If X2 <> X8 Then
            If X2 < DsysValueMin Or X2 > DsysValueMax Then
                X = MsgBox("Date valeur au Crédit hors limites ( +- 7 jours), cofirmez-vous ?", vbQuestion + vbYesNo, Me.Name & "contrôle date")
            End If
        End If
        
        If X = vbYes Then
            txtCréditDval.Text = dateImp(X2)
            mCréditDval = X2
        Else
            txtCréditDval.Text = constDateZ
        End If
    End If
End If
Cv_Compta

End Sub


Private Sub txtCréditLibellé_GotFocus()
Call txt_GotFocus(txtCréditLibellé)

End Sub


Private Sub txtCréditLibellé_LostFocus()
Call txt_LostFocus(txtCréditLibellé)

End Sub


Private Sub txtDébitCompte_GotFocus()
Call txt_GotFocus(txtDébitCompte)

End Sub


Private Sub txtDébitCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtDébitCompte_LostFocus()
txt_LostFocus txtDébitCompte
txtDébitCompte_Control
If blnDébitCompte Then
    If Not optVirCompte And Trim(txtCréditCompte) = "" Then txtCréditCompte = txtDébitCompte ': txtCréditCompte_Control
End If
'timerControl.Enabled = True 'Cv_Compta
timerControl.Enabled = True
End Sub

Private Sub txtDébitDval_GotFocus()
Call txt_GotFocus(txtDébitDval)

End Sub


Private Sub txtDébitDval_LostFocus()
Dim X As String, X2 As String
Call txt_LostFocus(txtDébitDval)

lstErr.Clear
X8 = mDébitDval
mDébitDval = "00000000"

X = dateCtl(txtDébitDval.Text)
If Not IsNumeric(X) Then
    txtDébitDval.ForeColor = errUsr.ForeColor
    Call lstErr_Clear(lstErr, txtDébitDval, "erreur date")
Else
    X2 = mId$(X, 1, 8)
    If X2 <> "00000000" Then
    
        X = vbYes
        If X2 <> X8 Then
            If X2 < DsysValueMin Or X2 > DsysValueMax Then
                X = MsgBox("Date valeur au débit hors limites ( +- 7 jours), cofirmez-vous ?", vbQuestion + vbYesNo, Me.Name & "contrôle date")
            End If
        End If
        
        If X = vbYes Then
            txtDébitDval.Text = dateImp(X2)
            mDébitDval = X2
        Else
            txtDébitDval.Text = constDateZ
        End If
    End If
End If

Cv_Compta

End Sub


Private Sub txtDébitLibellé_GotFocus()
Call txt_GotFocus(txtDébitLibellé)

End Sub


Private Sub txtDébitLibellé_LostFocus()
Call txt_LostFocus(txtDébitLibellé)
End Sub


Private Sub txtDevise1_Change()
form_Clear
End Sub

Private Sub txtDevise1_GotFocus()
Call txt_GotFocus(txtDevise1)
txtDevise1 = Trim(txtDevise1)
MouseMoveActiveControl_Set lblDevise
End Sub
Public Sub Msg_Rcv(txtMsg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(txtMsg, 13, 12)))
    Case Is = "FRMCOMPTE":
        Select Case currentActiveControl_Name
            Case "txtDébitCompte": txtDébitCompte = Trim(mId$(txtMsg, 38, 11)): txtDébitCompte_Control
            Case "txtCréditCompte":  txtCréditCompte = Trim(mId$(txtMsg, 38, 11)): txtCréditCompte_Control
        End Select
        timerControl.Enabled = True
    Case Else
        Call BiaPgmAut_Init(txtMsg, CVAut)
        CV_Compta_Init
        recCompteInit recCompte
End Select

End Sub
Private Sub frmCompte_Click()
Dim X As String
X = Space$(100)
Mid$(X, 1, 12) = "frmCompte   "
Mid$(X, 13, 12) = "frmCV      "
Mid$(X, 25, 10) = Space$(10)
Select Case currentActiveControl_Name
    Case "txtDébitCompte":
        If IsNumeric(txtDébitCompte) Then
            Mid$(X, 38, 11) = Format$(Val(txtDébitCompte), "###########")
        Else
            Mid$(X, 38, 11) = Format$(RTrim(txtDébitCompte), "@@@@@@@@@@@")
        End If
        If optDevise1Db Then
            Mid$(X, 35, 3) = CV1.DeviseN
        Else
            Mid$(X, 35, 3) = CV2.DeviseN
        End If
    Case "txtCréditCompte"
        If IsNumeric(txtCréditCompte) Then
            Mid$(X, 38, 11) = Format$(Val(txtCréditCompte), "###########")
        Else
            Mid$(X, 38, 11) = Format$(RTrim(txtCréditCompte), "@@@@@@@@@@@")
        End If
        If optDevise1Cr Then
            Mid$(X, 35, 3) = CV1.DeviseN
        Else
            Mid$(X, 35, 3) = CV2.DeviseN
        End If
End Select
Msg_Monitor X
End Sub




Private Sub txtDevise1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtDevise1_LostFocus()
Call txt_LostFocus(txtDevise1)
timerControl.Enabled = True

End Sub


Private Sub txtAmj_GotFocus()
Call txt_GotFocus(txtAmj)
End Sub


Private Sub txtAmj_LostFocus()
Dim X As String
Call txt_LostFocus(txtAmj)

lstErr.Clear

X = dateCtl(txtAmj.Text)
If Not IsNumeric(X) Then
    txtAmj.ForeColor = errUsr.ForeColor
    Call lstErr_Clear(lstErr, txtAmj, "erreur date")
Else
    valAMJ = mId$(X, 1, 8)
    If valAMJ <> "00000000" Then
        txtAmj.Text = dateImp(valAMJ)
    End If
    valAMJ1 = valAMJ: valAMJ2 = valAMJ
    timerControl.Enabled = True
End If

End Sub


Private Sub txtDevise1Montant_Change()
form_Clear
End Sub

Private Sub txtDevise1Montant_GotFocus()
Call txt_GotFocus(txtDevise1Montant)

End Sub


Private Sub txtDevise1Montant_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtDevise1Montant)

End Sub


Private Sub txtDevise1Montant_LostFocus()
Dim X As String
Call txt_LostFocus(txtDevise1Montant)
curX = num_CDec(txtDevise1Montant)
If curX = 0 Then
    txtDevise1Montant = ""
Else
    txtDevise1Montant = Format$(curX, "### ### ### ###.00")
   
End If
timerControl.Enabled = True

End Sub


Private Sub txtDevise2_Change()
form_Clear
End Sub

Private Sub txtDevise2_GotFocus()
Call txt_GotFocus(txtDevise2)
txtDevise2 = Trim(txtDevise2)
MouseMoveActiveControl_Set lblDevise
End Sub


Private Sub txtDevise2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtDevise2_LostFocus()
Call txt_LostFocus(txtDevise2)
timerControl.Enabled = True

End Sub


Private Sub txtDevise2Montant_Change()
form_Clear
End Sub

Private Sub txtDevise2Montant_GotFocus()
Call txt_GotFocus(txtDevise2Montant)

End Sub


Private Sub txtDevise2Montant_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtDevise2Montant)

End Sub


Private Sub txtDevise2Montant_LostFocus()
Call txt_LostFocus(txtDevise2Montant)
curX = num_CDec(txtDevise2Montant)
If curX = 0 Then
    txtDevise2Montant = ""
Else
    txtDevise2Montant = Format$(curX, "### ### ### ###.00")
End If
timerControl.Enabled = True

End Sub




Public Sub form_Clear()
picCompta.Cls
picCV.Cls
txtDevise3Montant = ""
'lblDevise2montant = ""
lblCours1_2 = ""
lblCours2_1 = ""
fraDébit.Visible = False
fraCrédit.Visible = False
cmdOk.Visible = False
libNUMPIE = ""
'''recCompteInit recDébitCompte: recCréditCompte = recDébitCompte
End Sub

Private Sub txtDevise3_Change()
form_Clear
End Sub


Public Sub txtDébitCompte_Control()

blnDébitCompte = False

recCompteInit recDébitCompte
If Val(txtDébitCompte) = 0 Then recDébitCompte.Intitulé = "compte à débiter": Exit Sub
recDébitCompte.Method = "SeekL1"
recDébitCompte.Société = SocId$
recDébitCompte.Agence = SocAgence$
If optDevise1Db Then
    recDébitCompte.Devise = CV1.DeviseN
Else
    recDébitCompte.Devise = CV2.DeviseN
End If

recDébitCompte.Numéro = Val(txtDébitCompte)
If IsNull(srvCompteFind(recDébitCompte)) Then
    blnDébitCompte = True
    libDébitCompte = recDébitCompte.Intitulé
    libDébitCompte.ForeColor = libUsr.ForeColor
    If recDébitCompte.Situation <> " " And recDébitCompte.Situation <> "E" Then
        Call lstErr_AddItem(lstErr, txtDébitCompte, "? compte à débiter annulé/bloqué")
    End If
    
    
Else
    libDébitCompte.ForeColor = errUsr.ForeColor
    libDébitCompte = "? compte à débiter inconnu"
    Call lstErr_AddItem(lstErr, txtDébitCompte, "? compte à débiter inconnu")
End If

End Sub
Public Sub txtCréditCompte_Control()

blnCréditCompte = False
recCompteInit recCréditCompte
If Val(txtCréditCompte) = 0 Then recCréditCompte.Intitulé = "compte à créditer": Exit Sub

recCréditCompte.Method = "SeekL1"
recCréditCompte.Société = SocId$
recCréditCompte.Agence = SocAgence$
If optDevise1Cr Then
    recCréditCompte.Devise = CV1.DeviseN
Else
    recCréditCompte.Devise = CV2.DeviseN
End If

recCréditCompte.Numéro = Val(txtCréditCompte)
If IsNull(srvCompteFind(recCréditCompte)) Then
    blnCréditCompte = True
    libCréditCompte = recCréditCompte.Intitulé
    libCréditCompte.ForeColor = libUsr.ForeColor
    If recCréditCompte.Situation <> " " Then
        Call lstErr_AddItem(lstErr, txtCréditCompte, "? compte à Créditer annulé/bloqué")
    End If
    
Else
    libCréditCompte.ForeColor = errUsr.ForeColor
    libCréditCompte = "? compte à Créditer inconnu"
    Call lstErr_AddItem(lstErr, txtCréditCompte, "? compte à Créditer inconnu")
End If

End Sub


Public Sub Cv_Compta()

CV_Compta_Gen
lstErr.Clear

CV_Compta_Display
'If Not optVirCompte Then
'    If CV1.DeviseN = CV2.DeviseN Then Call lstErr_AddItem(lstErr, picCompta, "! Devise 1 = Devise 2")
'End If

cmdOk.Visible = IIf(lstErr.ListCount = 0, CVAut.Saisir, False)

End Sub

Public Sub CV_Compta_Init()
recCpj030W0_Init arrCV030(0)
arrCV030(0).Method = "AddNew"
arrCV030(0).COSOC = SocId$
arrCV030(0).Agence = SocAgence$
arrCV030(0).AGEMET = SocAgence$
arrCV030(0).BIACOP = "S013"
arrCV030(0).SERVIC = "001"
arrCV030(0).AMJSAI = DSys
arrCV030(0).AMJVAL = DSys
arrCV030(0).AMJOPE = DSys
arrCV030(0).NOMOP = usrId
arrCV030(0).JJCPLT = "0"

mAMJP1 = dateElp("Ouvré", -1, DSys)
mAMJN2 = dateElp("Ouvré", 2, DSys)

End Sub
Public Sub CV_Compta_Gen()
Dim wDébitLibellé As String, wCréditLibellé As String
Dim strConversion As String, strMontant1 As String, strMontant2 As String, strMontant3 As String
Dim wAMJP1 As String * 8, wAMJN2 As String * 8


arrCV030(1) = arrCV030(0)
arrCV030(2) = arrCV030(0)
arrCV030(3) = arrCV030(0)
arrCV030(4) = arrCV030(0)
arrCV030(5) = arrCV030(0)
arrCV030(6) = arrCV030(0)

strMontant1 = CV1.DeviseIso & " " & Trim(Format$(CV1.Montant, "##### ### ##0.00"))
strMontant2 = CV2.DeviseIso & " " & Trim(Format$(CV2.Montant, "##### ### ##0.00"))
strMontant3 = CV3.DeviseIso & " " & Trim(Format$(CV3.Montant, "##### ### ##0.00"))

Select Case Conversion
    Case "C":   arrCV030Nb = 4
                strConversion = "CONVERSION "
                arrCV030(1).LIBELE = strConversion & strMontant1 & " / " & strMontant2
                arrCV030(2).LIBELE = CV2.DeviseN & " " & strConversion & strMontant2
                arrCV030(3).LIBELE = strConversion & strMontant2 & " / " & strMontant1
                arrCV030(4).LIBELE = CV1.DeviseN & " " & strConversion & strMontant1
    
    Case "B":   arrCV030Nb = 6
                strConversion = "ARBITRAGE "
                arrCV030(1).LIBELE = strConversion & strMontant1 & " / " & strMontant2
                arrCV030(2).LIBELE = CV3.DeviseN & " " & strConversion & strMontant3
                arrCV030(3).LIBELE = strConversion & strMontant2 & " / " & strMontant1
                arrCV030(4).LIBELE = CV3.DeviseN & " " & strConversion & strMontant3
                arrCV030(5).LIBELE = CV1.DeviseN & " " & strConversion & strMontant1
                arrCV030(6).LIBELE = CV2.DeviseN & " " & strConversion & strMontant2
   
    Case "A":   arrCV030Nb = 4
                strConversion = "ARBITRAGE "
                arrCV030(1).LIBELE = strConversion & strMontant1 & " / " & strMontant2
                arrCV030(2).LIBELE = CV2.DeviseN & " " & strConversion & strMontant2
                arrCV030(3).LIBELE = strConversion & strMontant2 & " / " & strMontant1
                arrCV030(4).LIBELE = CV1.DeviseN & " " & strConversion & strMontant1
    
   
End Select

wAMJP1 = mAMJP1
wAMJN2 = mAMJN2
If optCriRetour Then
    wAMJP1 = DSys: wAMJN2 = DSys
    arrCV030(1).LIBELE = "CRI RECU F/ "
End If

If optCriAller Then
    wAMJP1 = DSys: wAMJN2 = DSys
    arrCV030(1).LIBELE = "CRI EMIS O/ "
End If

If optVirCompte Or optPivot Or optEnCompte Then
    wAMJP1 = DSys: wAMJN2 = DSys
End If

wDébitLibellé = Trim(txtDébitLibellé)
wCréditLibellé = Trim(txtCréditLibellé)

arrCV030(1).MONDEV = CV1.Montant
arrCV030(2).MONDEV = CV1.Montant
arrCV030(3).MONDEV = CV2.Montant
arrCV030(4).MONDEV = CV2.Montant
arrCV030(5).MONDEV = CV3.Montant
arrCV030(6).MONDEV = CV3.Montant
If optVirCompte Then arrCV030(2).MONDEV = 0: arrCV030(4).MONDEV = 0


arrCV030(1).Devise = CV1.DeviseN
arrCV030(2).Devise = CV1.DeviseN
arrCV030(3).Devise = CV2.DeviseN
arrCV030(4).Devise = CV2.DeviseN
arrCV030(5).Devise = CV3.DeviseN
arrCV030(6).Devise = CV3.DeviseN

arrCV030(1).NOLIGN = 1
arrCV030(2).NOLIGN = 2
arrCV030(3).NOLIGN = 3
arrCV030(4).NOLIGN = 4
arrCV030(5).NOLIGN = 5
arrCV030(6).NOLIGN = 6

If optDevise1Db Then
    arrCV030(1).SENECR = "D"
    arrCV030(4).SENECR = "D"
    arrCV030(5).SENECR = "D"
    arrCV030(2).SENECR = "C"
    arrCV030(3).SENECR = "C"
    arrCV030(6).SENECR = "C"
    arrCV030(1).AMJVAL = wAMJP1
    If mDébitDval <> "00000000" Then arrCV030(1).AMJVAL = mDébitDval
    arrCV030(3).AMJVAL = wAMJN2
    If mCréditDval <> "00000000" Then arrCV030(3).AMJVAL = mCréditDval
    
    arrCV030(1).Compte = Format$(Val(txtDébitCompte), "00000000000")
    arrCV030(3).Compte = Format$(Val(txtCréditCompte), "00000000000")
    
    lblDébitLibellé = arrCV030(1).LIBELE
    If wDébitLibellé <> "" Then arrCV030(1).LIBELE = wDébitLibellé
    
    lblCréditLibellé = arrCV030(3).LIBELE
    If wCréditLibellé <> "" Then arrCV030(3).LIBELE = wCréditLibellé
Else
    arrCV030(1).SENECR = "C"
    arrCV030(4).SENECR = "C"
    arrCV030(5).SENECR = "C"
    arrCV030(2).SENECR = "D"
    arrCV030(3).SENECR = "D"
    arrCV030(6).SENECR = "D"
    arrCV030(1).AMJVAL = wAMJN2
    If mCréditDval <> "00000000" Then arrCV030(1).AMJVAL = mCréditDval
    arrCV030(3).AMJVAL = wAMJP1
    If mDébitDval <> "00000000" Then arrCV030(3).AMJVAL = mDébitDval
    
    arrCV030(3).Compte = Format$(Val(txtCréditCompte), "00000000000")
    
    arrCV030(1).Compte = Format$(Val(txtCréditCompte), "00000000000")
    arrCV030(3).Compte = Format$(Val(txtDébitCompte), "00000000000")
    
    lblCréditLibellé = arrCV030(1).LIBELE
    If wCréditLibellé <> "" Then arrCV030(1).LIBELE = wCréditLibellé
    
    lblDébitLibellé = arrCV030(3).LIBELE
    If wDébitLibellé <> "" Then arrCV030(3).LIBELE = wDébitLibellé
End If

If CV1.Normal = "N" Or CV1.Normal = "P" Then
    If optDevise1Db Then
        fraDébit.Visible = False
    Else
        fraCrédit.Visible = False
    End If
    
    arrCV030(1).AMJVAL = DSys
    If CV1.EuroIn Then
        arrCV030(1).Compte = "00010100006"
    Else
        arrCV030(1).Compte = "00010110001"

    End If
End If

If Conversion = "C" Or (CV1.EuroIn And Conversion = "B") Then
        arrCV030(2).Compte = "00038212005"
Else
    If CV1.Normal = "N" Or CV.Normal = "P" Then
        arrCV030(2).Compte = "00038393005"
    Else
        arrCV030(2).Compte = "00038210003"
   End If
End If


If CV2.Normal = "N" Or CV2.Normal = "P" Then
    If optDevise2Db Then
        fraDébit.Visible = False
    Else
        fraCrédit.Visible = False
    End If
    
    arrCV030(3).AMJVAL = DSys
    If CV2.EuroIn Then
        arrCV030(3).Compte = "00010100006"
    Else
        arrCV030(3).Compte = "00010110001"
    End If
End If

If Conversion = "C" Or (CV2.EuroIn And Conversion = "B") Then
        arrCV030(4).Compte = "00038212005"
Else
    If CV2.Normal = "N" Or CV.Normal = "P" Then
        arrCV030(4).Compte = "00038393005"
    Else
        arrCV030(4).Compte = "00038210003"
   End If
End If

arrCV030(5).Compte = arrCV030(2).Compte
arrCV030(6).Compte = arrCV030(4).Compte

End Sub


Public Sub Cv_Compta_Clear()
cmdOk.Visible = False
txtDevise1Montant = "": txtDevise2Montant = "":: txtDevise3Montant = ""
chkDébitAvis = 0: chkCréditAvis = 0
mDébitCompte = "": mCréditCompte = ""
libDébitCompte = ""
lblDébitLibellé = ""
txtDébitLibellé = ""
txtDébitDval = constDateZ
txtDébitCompte = ""
mDébitDval = "00000000"

libCréditCompte = ""
lblCréditLibellé = ""
txtCréditLibellé = ""
txtCréditDval = constDateZ
txtCréditCompte = ""
mCréditDval = "00000000"
recCompteInit recDébitCompte: recCréditCompte = recDébitCompte
optEnCompte = True
blnDébitLibellé = False: blnCréditlibellé = False
lstDevise.Visible = False

valAMJ = DSys
txtAmj.Text = dateImp(valAMJ)
 valAMJ1 = valAMJ: valAMJ2 = valAMJ
End Sub

Public Sub cmdPrintX(Text As String)
Dim I As Integer, Msg As String

prtCompta.CV1 = CV1
prtCompta.CV2 = CV2
prtCompta.CV3 = CV3

For I = 1 To 6
    prtCompta.arrCV030(I) = arrCV030(I)
Next I

Msg = Format$(1, "000000") & Format$(arrCV030Nb, "000000")

prtCompta_Monitor Msg, Text, "Contre-Valeur"

End Sub

Public Sub cmdControl()
Dim X As String
Dim valX As String
Dim valDevise2Montant As Currency, curX As Currency
Dim dblX As Double
Dim C As Control, blnFocus As Boolean

lstDevise.Visible = False

If blnTimer Then Exit Sub 'timerControl.Enabled = True: Exit Sub
frmCV.Enabled = False   ' jpl 10-05-99

blnTimer = True
blnFocus = False
For Each C In Me.Controls
    If TypeOf C Is TextBox Then
        If C.BackColor = focusUsr.BackColor Then
            blnFocus = True
            Exit For
        End If
    End If
Next C


lstErr.Clear
lstErr.Height = 200

txtDevise2.Enabled = True
txtDevise2Montant.Enabled = True
fraDevise1Sens.Enabled = True
fraDevise2Sens.Enabled = True
txtDébitCompte.Enabled = True
txtCréditCompte.Enabled = True
If optVirCompte Then
    txtDevise2.Enabled = False
    txtDevise2Montant.Enabled = False
    txtDevise2 = txtDevise1
    txtDevise2Montant = txtDevise1Montant
End If

If optCriRetour Then
    txtDevise2.Enabled = False
    txtDevise2 = "EUR"
    fraDevise1Sens.Enabled = False
    fraDevise2Sens.Enabled = False
    optDevise2Db = True
    txtDébitCompte.Enabled = False
    txtDébitCompte = mCriCompte
End If

If optCriAller Then
    txtDevise2.Enabled = False
    txtDevise2 = "EUR"
    fraDevise1Sens.Enabled = False
    fraDevise2Sens.Enabled = False
    optDevise1Db = True
    txtCréditCompte.Enabled = False
    txtCréditCompte = mCriCompte
End If
If optCriRetour Or optCriAller Then
   fraCRI.Visible = True
Else
    fraCRI.Visible = False
End If

If optNormal Or optPrivilégié Then
    fraDevise1BME.Enabled = True
    fraDevise2BME.Enabled = True
Else
    fraDevise1BME.Enabled = False
    optDevise1EnCompte = True
    fraDevise2BME.Enabled = False
    optDevise2EnCompte = True
End If


chkDevise2.Visible = optEnCompte

lstErr.Clear
CV_Init CV1
CV1.OpéAmj = valAMJ

CV2 = CV1
CV3 = CV_Euro

X = Trim(txtDevise1)
If X = "" Then Call lstErr_Clear(lstErr, cmdContext, "? devise1"): GoTo ExitSub
If X = "EUR" Or X = "978" Then optDevise1EnCompte = True
If IsNumeric(X) Then
    CV1.DeviseN = X
Else
    CV1.DeviseIso = X
End If
X = num_Control(txtDevise1Montant, valX, 13, 2)
CV1.Montant = valX
If CV1.Montant = 0 Then Call lstErr_Clear(lstErr, cmdContext, "? montant"): GoTo ExitSub

X = Trim(txtDevise2)
If X = "" Then Call lstErr_Clear(lstErr, cmdContext, "? devise2"): GoTo ExitSub
If X = "EUR" Or X = "978" Then optDevise2EnCompte = True

If IsNumeric(X) Then
    CV2.DeviseN = X
Else
    CV2.DeviseIso = X
End If

If chkDevise2 = 1 Then
    X = num_Control(txtDevise2Montant, valX, 13, 2)
    valDevise2Montant = valX
    If valDevise2Montant = 0 Then Call lstErr_Clear(lstErr, cmdContext, "? montant devise2"): GoTo ExitSub
End If
If optDevise1Db Then
    CV1.AchatVente = "V"
    CV2.AchatVente = "A"
Else
    If optDevise1Cr Then
        CV1.AchatVente = "A"
        CV2.AchatVente = "V"
    End If
End If

If optPivot Or optVirCompte Then
    CV1.Normal = " "
    CV2.Normal = " "
    CV3.Normal = " "
Else
    If optNormal Then
        CV1.Normal = "N"
        CV2.Normal = "N"
        CV3.Normal = "N"
    Else
        If optPrivilégié Then
            CV1.Normal = "P"
            CV2.Normal = "P"
            CV3.Normal = "P"
        Else
            CV1.Normal = "C"
            CV2.Normal = "C"
            CV3.Normal = "C"
        End If
    End If
End If

If optNormal Or optPrivilégié Then
    If optDevise1EnCompte Then CV1.Normal = "C"
    If optDevise2EnCompte Then CV2.Normal = "C"
End If

If optDevise1Cr Then
    txtDevise1Montant.ForeColor = vbBlue
    optDevise1Cr.ForeColor = vbBlue
    optDevise1Db.ForeColor = lblUsr.ForeColor
Else
    txtDevise1Montant.ForeColor = vbRed
    optDevise1Cr.ForeColor = lblUsr.ForeColor
    optDevise1Db.ForeColor = vbRed
End If


If optDevise2Cr Then
    txtDevise2Montant.ForeColor = vbBlue
    optDevise2Cr.ForeColor = vbBlue
    optDevise2Db.ForeColor = lblUsr.ForeColor
Else
    txtDevise2Montant.ForeColor = vbRed
    optDevise2Cr.ForeColor = lblUsr.ForeColor
    optDevise2Db.ForeColor = vbRed
End If

If optDevise1EnCompte Then
    optDevise1EnCompte.ForeColor = vbBlue
    optDevise1Espèces.ForeColor = lblUsr.ForeColor
Else
    optDevise1EnCompte.ForeColor = lblUsr.ForeColor
    optDevise1Espèces.ForeColor = vbBlue
End If


If optDevise2EnCompte Then
    optDevise2EnCompte.ForeColor = vbBlue
    optDevise2Espèces.ForeColor = lblUsr.ForeColor
Else
    optDevise2EnCompte.ForeColor = lblUsr.ForeColor
    optDevise2Espèces.ForeColor = vbBlue
End If

Call CV_Transitoire(CV1, CV2, CV3, Conversion)

If CV1.EuroIn Or CV1.DeviseIso = "EUR" Then
    If CV2.EuroIn Or CV2.DeviseIso = "EUR" Then
        chkDevise2.Visible = False
        chkDevise2.Value = 0
    End If
End If

txtDevise1 = CV1.DeviseIso
txtDevise1Montant = Format$(CV1.Montant, "### ### ### ##0.00")

txtDevise2 = CV2.DeviseIso
If chkDevise2 = 1 Then
    If Not CV1.EuroIn And CV2.EuroIn Then
        If CV1.DeviseIso <> "EUR" Then
            wCV1 = CV2: wCV1.Montant = valDevise2Montant
            wCV2 = CV1
            CV_Calc wCV1, wCV2, CV3
        End If
    End If
    
    curX = valDevise2Montant - CV2.Montant
    X = "Différence : " & Format$(curX, "### ### ### ##0.00")
    Call lstErr_Clear(lstErr, cmdContext, X)
    lblDevise2montant = Format$(CV2.Montant, "### ### ### ##0.00")
    CV2.Montant = valDevise2Montant
Else
    lblDevise2montant = ""
End If
txtDevise2Montant = Format$(CV2.Montant, "### ### ### ##0.00")

txtDevise3 = CV3.DeviseIso
txtDevise3Montant = Format$(CV3.Montant, "### ### ### ##0.00")

'picCV.Cls
picCV.CurrentY = -250

CV = CV1: CV_Cours_Display
CV = CV2: CV_Cours_Display

fraDébit.Visible = CVAut.Saisir
fraCrédit.Visible = CVAut.Saisir
''''cmdOk.Visible = CVAut.Saisir

If mDébitCompte = recDébitCompte.Numéro Then
    blnCréditlibellé = True
Else
    mDébitCompte = recDébitCompte.Numéro
    blnCréditlibellé = False
End If
If mCréditCompte = recCréditCompte.Numéro Then
    blnDébitLibellé = True
Else
    mCréditCompte = recCréditCompte.Numéro
    blnDébitLibellé = False
End If

If CV1.Montant <> 0 Then
    If optCriRetour Then
        txtDébitCompte_Control
        If Not blnCréditlibellé Then txtCréditLibellé = "VIRT RECU "
        If Not blnDébitLibellé Then txtDébitLibellé = mId$("CRI RECU F/ " & recCréditCompte.Intitulé, 1, 50)
   End If
    If optCriAller Then
        txtCréditCompte_Control
        If Not blnDébitLibellé Then txtDébitLibellé = "VIRT EMIS "
        If Not blnCréditlibellé Then txtCréditLibellé = mId$("CRI EMIS O/ " & recDébitCompte.Intitulé, 1, 50)
    End If
    If optVirCompte Then
        CV2.Montant = CV1.Montant
        txtDébitCompte_Control
        If Not blnDébitLibellé Then txtDébitLibellé = mId$("VIRT F/ " & recCréditCompte.Intitulé, 1, 50)
        txtCréditCompte_Control
        If Not blnCréditlibellé Then txtCréditLibellé = mId$("VIRT O/ " & recDébitCompte.Intitulé, 1, 50)
   End If
End If
Cv_Compta

If CV1.Cours <> 0 Then dblX = CV2.Cours / CV1.Cours:    lblCours1_2 = Format$(dblX, "## ##0.00 000 00") & "  / "
If CV2.Cours <> 0 Then dblX = CV1.Cours / CV2.Cours:    lblCours2_1 = Format$(dblX, "## ##0.00 000 00")

If Not CV1.EuroIn And valAMJ1 <> CV1.CoursAmj Then
    X = MsgBox("! cours " & txtDevise1 & " au " & dateImp(CV1.CoursAmj) & "; confirmez-vous ?", vbQuestion + vbYesNo, "Contre-Valeur : contrôle date du cours ")
    If X = vbYes Then
        valAMJ1 = CV1.CoursAmj
    Else
        Call lstErr_AddItem(lstErr, cmdContext, "? date du cours " & txtDevise1)
    End If
End If


If Not CV2.EuroIn And valAMJ2 <> CV2.CoursAmj Then
    X = MsgBox("! cours " & txtDevise2 & " au " & dateImp(CV2.CoursAmj) & "; confirmez-vous ?", vbQuestion + vbYesNo, "Contre-Valeur : contrôle date du cours ")
    If X = vbYes Then
        valAMJ2 = CV2.CoursAmj
    Else
        Call lstErr_AddItem(lstErr, cmdContext, "? date du cours " & txtDevise2)
    End If
End If

ExitSub:
'$JPL 99-05-05 timerControl.Enabled = False
frmCV.Enabled = True ' jpl 10-05-99
If blnFocus Then txt_GotFocus C: If C.Enabled Then C.SetFocus
blnTimer = False

End Sub

Public Sub prtCv_Avis(I As Integer)
recOpCpt_Init recOpCpt
recOpCpt.Référence = Format(Val(arrCV030(I).NUMPIE), "000000")
recOpCpt.CodeOpération = arrCV030(I).BIACOP
recOpCpt.Société = arrCV030(I).COSOC
recOpCpt.Agence = arrCV030(I).Agence
recOpCpt.Devise = Format(Val(arrCV030(I).Devise), "000")
recOpCpt.Compte = arrCV030(I).Compte
recOpCpt.Brut = arrCV030(I).MONDEV
recOpCpt.Sens = arrCV030(I).SENECR
recOpCpt.AmjOpération = arrCV030(I).AMJOPE
recOpCpt.AmjValeur = arrCV030(I).AMJVAL
recOpCpt.Libellé = arrCV030(I).LIBELE
recOpCpt.optAvis = "1"
prtAvisX recOpCpt, " "
End Sub
