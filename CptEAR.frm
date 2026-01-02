VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmCptEAR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "CptEAR : écritures à régulariser"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9420
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5400
      TabIndex        =   8
      Top             =   0
      Width           =   3500
   End
   Begin TabDlg.SSTab sstab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "CptEAR.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOption"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Liste des écritures"
      TabPicture(1)   =   "CptEAR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picInfo2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraUpdate"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "picInfo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fgSelect"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Détail"
      TabPicture(2)   =   "CptEAR.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.PictureBox picInfo2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1400
         Left            =   -74880
         ScaleHeight     =   1335
         ScaleWidth      =   3225
         TabIndex        =   20
         Top             =   4920
         Width           =   3285
      End
      Begin VB.Frame fraUpdate 
         Caption         =   "Régularisation de l'écriture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   -71400
         TabIndex        =   17
         Top             =   3360
         Width           =   5655
         Begin VB.Frame fraUpdateS 
            Height          =   1455
            Left            =   120
            TabIndex        =   26
            Top             =   1540
            Width           =   4455
            Begin VB.TextBox txtUpdateS 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   2900
               TabIndex        =   30
               Top             =   200
               Width           =   1335
            End
            Begin VB.TextBox txtUpdateD 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   2900
               TabIndex        =   29
               Top             =   500
               Width           =   1335
            End
            Begin VB.TextBox txtUpdateE 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   2900
               TabIndex        =   28
               Top             =   800
               Width           =   1335
            End
            Begin VB.TextBox txtUpdateF 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   2900
               TabIndex        =   27
               Top             =   1100
               Width           =   1335
            End
            Begin VB.Label lblUpdateS 
               Caption         =   "solde provisoire"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   200
               Width           =   1815
            End
            Begin VB.Label lblUpdateD 
               Caption         =   "découvert autorisé"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   500
               Width           =   2175
            End
            Begin VB.Label lblUpdateE 
               Caption         =   "régularisation"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   800
               Width           =   2175
            End
            Begin VB.Label lblUpdateF 
               Caption         =   "solde après mouvement"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   1100
               Width           =   2175
            End
         End
         Begin VB.Frame fraUpdateDta 
            Height          =   1350
            Left            =   120
            TabIndex        =   21
            Top             =   200
            Width           =   5415
            Begin VB.OptionButton optUpdateSIT 
               Caption         =   "rejet SIT"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   600
               Width           =   1095
            End
            Begin VB.OptionButton optUpdateCRI 
               Caption         =   "rejet CRI"
               Height          =   255
               Left            =   1320
               TabIndex        =   35
               Top             =   600
               Width           =   1095
            End
            Begin VB.OptionButton optUpdateCompte 
               Caption         =   "imputer sur le compte :"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton optUpdateAnnulation 
               Caption         =   "ignorer définitivement"
               Height          =   300
               Left            =   2880
               TabIndex        =   23
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox txtUpdateCompte 
               Height          =   285
               Left            =   2880
               TabIndex        =   22
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label libUpdateCompte 
               Caption         =   "-"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   960
               Width           =   5055
               WordWrap        =   -1  'True
            End
         End
         Begin VB.CommandButton cmdUpdateNOK 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Invalider"
            Height          =   495
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2400
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdateOk 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Valider"
            Height          =   495
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1680
            Width           =   855
         End
      End
      Begin VB.PictureBox picInfo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   4320
         Left            =   -74880
         ScaleHeight     =   4260
         ScaleWidth      =   3225
         TabIndex        =   16
         Top             =   480
         Width           =   3285
      End
      Begin VB.Frame fraOption 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   8895
         Begin VB.OptionButton optSelectAll 
            Caption         =   "toutes les écritures "
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2520
            Width           =   2055
         End
         Begin VB.OptionButton optSelectAnnulation 
            Caption         =   "écritures annulées"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   2040
            Width           =   1815
         End
         Begin VB.OptionButton optSelectàValider 
            Caption         =   "écritures de régularisation à valider"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            Width           =   3255
         End
         Begin VB.OptionButton optSelectEnCours 
            Caption         =   "écritures à régulariser"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.CheckBox chkSelectService 
            Caption         =   "Service"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher"
            Height          =   975
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3720
            Width           =   2415
         End
         Begin VB.TextBox txtSelectService 
            Height          =   285
            Left            =   2880
            TabIndex        =   6
            Top             =   600
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker txtSelectAmjMax 
            Height          =   300
            Left            =   4560
            TabIndex        =   7
            Top             =   2280
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
            Format          =   65601539
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtSelectAmjMin 
            Height          =   300
            Left            =   2880
            TabIndex        =   11
            Top             =   2280
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
            Format          =   65601539
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   2850
         Left            =   -71400
         TabIndex        =   4
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   5027
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   14737632
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"CptEAR.frx":0054
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "CptEAR.frx":00E9
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
      Height          =   500
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOption 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnufgSelect 
      Caption         =   "Opération"
      Visible         =   0   'False
      Begin VB.Menu mnufgSelect_Display 
         Caption         =   "Afficher cette écriture"
      End
      Begin VB.Menu mnufgSelect_Print 
         Caption         =   "Imprimer cette écriture"
      End
   End
   Begin VB.Menu mnucmdPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnucmdPrint_fgSelect_Row 
         Caption         =   "Imprimer cette écriture"
      End
      Begin VB.Menu mnucmdPrint_fgSelect_All 
         Caption         =   "Imprimer toutes les écritures"
      End
      Begin VB.Menu mnucmdPrint_fgSelect 
         Caption         =   "Imprimer la liste"
      End
   End
End
Attribute VB_Name = "frmCptEAR"
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
Dim CptEARAut As typeAuthorization

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim recCptEAR As typeCptEAR, xCptEAR As typeCptEAR, mCptEAR As typeCptEAR

Dim meCptEAR() As typeCptEAR
Dim meCptEAR_Nb As Integer, meCptEAR_Index As Integer, meCptEAR_NbMax As Integer

Dim blncmdUpdateOK_Visible As Boolean, blnErr As Boolean, blncmdUpdateNOK_Visible As Boolean
Dim blnfgSelect_DisplayLine As Boolean, blnfgEchéance_DisplayLine As Boolean


Dim blnSetfocus As Boolean
Dim blnSelectService As Boolean, wSelectService As String * 3, blnSelectService_Enabled As Boolean
Dim blnSelectAmj As Boolean, wSelectAmjMin As String * 8, wSelectAmjMax As String * 8

Dim meCompteOri As typeCompte, meCompteDes As typeCompte
Dim moptAnnulation As Boolean
Dim meCV1 As typeCV, meCV2 As typeCV, meCV3 As typeCV

Dim meCptMvt As typeCptMvt
Dim meBiaLog As typeBiaLog

Dim blnCompteSituation_Saisie As Boolean, blnCompteSituation_Validation As Boolean, blnCompteSituation_Forçage As Boolean
Dim AmjOpéMin As String * 8
'Dim curX As Currency

Dim meCompte() As typeCompte
Dim meCompte_Nb As Integer, meCompte_Index As Integer, meCompte_NbMax As Integer

Dim paramCPTEAR_CRI As String * 11, paramCPTEAR_SIT As String * 11

Public Sub prtCptEAR_Avis()
Dim X1 As String * 1
Dim yMax As Integer

Bialog_Load
If meCompteOri.Numéro <> mCptEAR.Compte Then
    meCompteOri.Devise = mCptEAR.Devise
    meCompteOri.Numéro = mCptEAR.Compte
    V = mdbCptP0_Find(meCompteOri)
End If

recCptMvtInit meCptMvt
meCptMvt.Pièce = Format(Val(mCptEAR.NUMPIE), "000000")
meCptMvt.CodeOpération = mCptEAR.BIACOP
meCptMvt.Société = mCptEAR.COSOC
meCptMvt.Agence = mCptEAR.Agence
meCptMvt.Devise = Format(Val(mCptEAR.Devise), "000")
meCptMvt.Compte = mCptEAR.Compte
meCptMvt.MT = mCptEAR.MONDEV
meCptMvt.AmjOpération = mCptEAR.AMJOPE
meCptMvt.AmjValeur = mCptEAR.AMJVAL
meCptMvt.Libellé = mCptEAR.LIBELE

Call prtAvis_CptEar(meCptMvt, "")


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX - 10, XPrt.CurrentY + prtlineHeight * 2, " ", "230")

XPrt.FontSize = 12
XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + 100
X = "ECRITURE A REGULARISER :   " & mCptEAR.EARIdRef
''XPrt.CurrentX = prtMinX + 100: XPrt.Print X;
XPrt.FontUnderline = True
frmElpPrt.prtCentré (prtMaxX - prtMinX) / 2, X
XPrt.FontUnderline = False

''XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.3
XPrt.CurrentX = prtMinX + 100: XPrt.Print "Service :   " & mCptEAR.SERVIC;
X = "Compta du  " & dateImp(mCptEAR.EARCptAmj)
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2.5
XPrt.CurrentX = prtMinX: XPrt.Print "Motif du rejet";
XPrt.CurrentX = prtMinX + 1600: XPrt.Print ":    ";: XPrt.Print DicLib(524, mCptEAR.LogCodErr);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinX + 1600: XPrt.Print "     ";: XPrt.Print meBiaLog.Log_Texte1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinX + 1600: XPrt.Print "     ";: XPrt.Print meBiaLog.Log_Texte2;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinX: XPrt.Print "Compte d'origine";
XPrt.CurrentX = prtMinX + 1600: XPrt.Print ":    ";: XPrt.Print Compte_Imp(mCptEAR.EARCptOri) & "    " & meCompteOri.Intitulé;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinX: XPrt.Print "Compte EAR";
XPrt.CurrentX = prtMinX + 1600: XPrt.Print ":    ";: XPrt.Print Compte_Imp(mCptEAR.EARCptEAR);

XPrt.FontBold = True: XPrt.FontSize = 12
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
yMax = XPrt.CurrentY + prtlineHeight * 1.5
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX - 10, yMax, " ", "230")
recStatut.K2 = mCptEAR.EARStatus: tableElpTable_Read recStatut
XPrt.CurrentY = XPrt.CurrentY + 100
X = "REGULARISATION :    " & Trim(recStatut.Name)
frmElpPrt.prtCentré (prtMaxX - prtMinX) / 2, X

XPrt.FontBold = False: XPrt.FontSize = 8

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinX: XPrt.Print "Compte destination";
XPrt.CurrentX = prtMinX + 1600: XPrt.Print ":    ";

If mCptEAR.EARNumPie > 0 Then

     XPrt.Print Compte_Imp(mCptEAR.EARCptDes) & "    " & meCompteDes.Intitulé;
     
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinX: XPrt.Print "     Lot";
    XPrt.CurrentX = prtMinX + 1600: XPrt.Print ":    ";: XPrt.Print Format$(mCptEAR.EARNumLot, "#####");
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinX: XPrt.Print "     Pièce";
    XPrt.CurrentX = prtMinX + 1600: XPrt.Print ":    ";: XPrt.Print Format$(mCptEAR.EARNumPie, "########");
End If

XPrt.CurrentY = yMax + prtlineHeight
Call frmElpPrt.prtTrame(prtMaxX - 4000, XPrt.CurrentY, prtMaxX - 10, XPrt.CurrentY + prtlineHeight * 5, " ", "230")
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMaxX - 2000, Trim("Demande de validation")

XPrt.FontUnderline = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMaxX - 3900
XPrt.Print mCptEAR.EARMajUsr;
XPrt.Print dateImp(mCptEAR.EARMajAmj) & "    " & timeImp(mCptEAR.EARMajHms);


XPrt.CurrentY = yMax + prtlineHeight * 7
Call frmElpPrt.prtTrame(prtMaxX - 4000, XPrt.CurrentY, prtMaxX - 10, XPrt.CurrentY + prtlineHeight * 5, " ", "230")

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMaxX - 2000, Trim("VALIDATION")

XPrt.FontUnderline = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMaxX - 3900
XPrt.Print mCptEAR.EARValUsr;
XPrt.Print dateImp(mCptEAR.EARValAmj) & "    " & timeImp(mCptEAR.EARValHms);

XPrt.CurrentY = yMax + prtlineHeight * 13
Call frmElpPrt.prtTrame(prtMaxX - 4000, XPrt.CurrentY, prtMaxX - 10, XPrt.CurrentY + prtlineHeight * 5, " ", "230")

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMaxX - 2000, Trim("ARBITRAGE")
XPrt.FontUnderline = False

recCompteInit meCompte(1)
meCompte(1).Method = "SnapL5"
meCompte(1).Société = SocId$
meCompte(1).Agence = SocAgence$
If mCptEAR.EARNumPie > 0 Then
    meCompte(1).Numéro = mCptEAR.EARCptDes
Else
    meCompte(1).Numéro = mCptEAR.EARCptOri
End If
meCompte(1).Devise = "000"
meCompte(0) = meCompte(1)
meCompte(0).Devise = "999"

Compte_Sel
meCV2.DeviseIso = "EUR"

prtCurrentY = prtMaxY - prtlineHeight * 10
XPrt.CurrentY = prtCurrentY
yMax = prtMaxY
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, "B", "230")
XPrt.Line (prtMinX, prtCurrentY)-(prtMinX, yMax)
XPrt.Line (prtMinX + 2000, prtCurrentY)-(prtMinX + 2000, yMax)
XPrt.Line (prtMinX + 6000, prtCurrentY)-(prtMinX + 6000, yMax)
XPrt.Line (prtMinX + 8000, prtCurrentY)-(prtMinX + 8000, yMax)
XPrt.Line (prtMaxX, prtCurrentY)-(prtMaxX, yMax)

XPrt.CurrentY = prtCurrentY + 20

XPrt.CurrentX = prtMinX + 100: XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 3500: XPrt.Print "Solde en devise";
XPrt.CurrentX = prtMinX + 6500: XPrt.Print "CV solde en EUR";
XPrt.CurrentX = prtMinX + 8600: XPrt.Print "Autorisation en devise";


For I = meCompte_Nb To 1 Step -1
    meCV1.DeviseIso = ""
    meCV1.DeviseN = meCompte(I).Devise
    meCV1.Montant = meCompte(I).SoldeInstantané
    Call CV_Transitoire(meCV1, meCV2, meCV3, X1)
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 100
    XPrt.Print Compte_Imp(meCompte(I).Numéro) & "    " & meCV1.DeviseIso;
    X = Format$(meCV1.Montant, "### ### ### ##0.00")
    If meCV1.Montant < 0 Then
        XPrt.CurrentX = prtMinX + 3900 - XPrt.TextWidth(X)
    Else
        XPrt.CurrentX = prtMinX + 5900 - XPrt.TextWidth(X)
    End If
    XPrt.Print X;
    
    X = Format$(meCV2.Montant, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 7900 - XPrt.TextWidth(X)
    XPrt.Print X;
    
    If Val(meCompte(I).DécouvertAmj) > DSys Then
        XPrt.CurrentX = prtMinX + 8100
        XPrt.Print dateImp(meCompte(I).DécouvertAmj);
        X = Format$(meCompte(I).DécouvertMontant, "### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If

Next I

frmElpPrt.prtStdBottom

prtAvis_Close

End Sub

Private Sub cmdUpdateNOK_Click()

Me.Enabled = False
lstErr.Clear
fraUpdate.Enabled = False
cmdUpdateOk.Visible = False
cmdUpdateNOK.Visible = False
currentAction = constInvalider

xCptEAR = mCptEAR
xCptEAR.Method = constUpdate
xCptEAR.EARStatus = ""
xCptEAR.EARMajAmj = DSys
xCptEAR.EARMajHms = time_Hms
xCptEAR.EARMajUsr = ""

cmdUpdate_Db

Exit_Sub:

currentAction = ""

Me.Enabled = True
AppActivate Me.Caption


End Sub

Private Sub cmdUpdateOk_Click()

Me.Enabled = False

lstErr.Clear
xCptEAR = mCptEAR: cmdUpdate_Control

If lstErr.ListCount <> 0 Then GoTo Exit_Sub
Me.Enabled = False
fraUpdate.Enabled = False
cmdUpdateOk.Visible = False
cmdUpdateNOK.Visible = False

xCptEAR.Method = constUpdate
Select Case currentAction
    Case constSaisie
        xCptEAR.EARMajAmj = DSys
        xCptEAR.EARMajHms = time_Hms
        xCptEAR.EARMajUsr = usrId
    Case constValider
        If Not blnCompteSituation_Validation Then
            Call lstErr_AddItem(lstErr, cmdContext, "? Situation du compte")
        Else

            If Not CptEARAut.Xspécial And Trim(xCptEAR.EARMajUsr) = Trim(usrId) Then
                Call MsgBox("Vous ne pouvez pas valider vos propres opérations.", vbCritical, "TC : Validation ")
                Call lstErr_AddItem(lstErr, cmdContext, "? validation interdite")
            Else
                xCptEAR.EARValAmj = DSys
                xCptEAR.EARValHms = time_Hms
                xCptEAR.EARValUsr = usrId
            End If
        End If
    Case Else
        Call lstErr_AddItem(lstErr, cmdContext, "? cmdOk : " & cmdUpdateOk.Caption)
End Select

cmdUpdate_Db

Exit_Sub:

currentAction = ""
Me.Enabled = True
AppActivate Me.Caption



End Sub

Private Sub cmdUpdateQuit()
fgSelect.Row = 0
Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
picInfo.Cls
picInfo2.Cls
txtUpdateCompte = ""
fraUpdate.Enabled = False
libRéférenceInterne = ""
End Sub

Private Sub mnucmdPrint_fgSelect_All_Click()
Me.Enabled = False

For meCptEAR_Index = 1 To meCptEAR_Nb
    If meCptEAR(meCptEAR_Index).Method <> constIgnore And meCptEAR(meCptEAR_Index).Method <> constDelete Then
        mCptEAR = meCptEAR(meCptEAR_Index)
        prtCptEAR_Avis
    End If

Next meCptEAR_Index

fgSelect_Display
Me.Enabled = True

End Sub

Private Sub mnucmdPrint_fgSelect_Row_Click()
Me.Enabled = False

prtCptEAR_Avis

Me.Enabled = True

End Sub

Private Sub mnufgSelect_Display_Click()
srvCptEAR_ElpDisplay mCptEAR

End Sub

Private Sub mnufgSelect_Print_Click()
prtCptEAR_Avis
End Sub

Private Sub optSelectAll_Click()
On Error GoTo Exit_Sub
blnSelectAmj = True
txtSelectAmjMin.Visible = True: txtSelectAmjMax.Visible = True
If blnSetfocus Then txtSelectAmjMin.SetFocus
Exit_Sub:

End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear
If fraUpdate.Enabled Then
    cmdUpdateQuit
Else
    If sstab1.Tab <> 0 Then
        sstab1.Tab = 0
    Else
        If currentAction = "" Then
            If blnMsgBox_Quit Then
                X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
            Else
               X = vbYes
            End If
            If X = vbYes Then Unload Me
        Else
            cmdReset
        End If
    End If
End If


End Sub
Public Sub cmdControl()

If Not Me.Enabled Then Exit Sub
Me.Enabled = False

'cmdOk.Visible = False
'cmdSave.Visible = False
blnControl = False
'blnSetfocus = False

lstErr.Clear
lstErr.Height = 200

blnSelectService = IIf(chkSelectService = "1", True, False)
wSelectService = Format$(Val(Trim(txtSelectService)), "000")
If blnSelectService Then
    If wSelectService = "000" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le service")
End If

Call DTPicker_Control(txtSelectAmjMin, wSelectAmjMin)
Call DTPicker_Control(txtSelectAmjMax, wSelectAmjMax)
If blnSelectAmj Then
    If wSelectAmjMin = "00000000" Then
        Call lstErr_AddItem(lstErr, cmdContext, "? préciser le amj min")
    Else
        If wSelectAmjMax = "00000000" Then wSelectAmjMax = wSelectAmjMin
    End If
    If wSelectAmjMin > wSelectAmjMax Then Call lstErr_AddItem(lstErr, cmdContext, "? amj min > amj max")
End If


If lstErr.ListCount > 0 Then
    lstErr.Visible = True
End If

ExitSub:

Me.Enabled = True
    
blnControl = True

End Sub

Private Sub chkSelectService_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSelectService

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdSelect

End Sub

Private Sub fraOption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub optSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSelectAll

End Sub

Private Sub optSelectAnnulation_Click()
On Error GoTo Exit_Sub
blnSelectAmj = True

    txtSelectAmjMin.Visible = True: txtSelectAmjMax.Visible = True
    If blnSetfocus Then txtSelectAmjMin.SetFocus
Exit_Sub:

End Sub

Private Sub optSelectAnnulation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSelectAnnulation
End Sub


Private Sub optSelectàValider_Click()
blnSelectAmj = False
txtSelectAmjMin.Visible = False: txtSelectAmjMax.Visible = False

End Sub


Private Sub optSelectàValider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSelectàValider
End Sub


Private Sub optSelectEnCours_Click()
txtSelectAmjMin.Visible = False: txtSelectAmjMax.Visible = False
blnSelectAmj = False
End Sub


Private Sub optSelectEnCours_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSelectEnCours
End Sub


Private Sub optUpdateAnnulation_Click()
optUpdate_Color
End Sub

Private Sub optUpdateCompte_Click()
optUpdate_Color
End Sub


Private Sub optUpdateCRI_Click()
optUpdate_Color

End Sub

Private Sub optUpdateSIT_Click()
optUpdate_Color

End Sub


Private Sub txtSelectAmjMax_GotFocus()
DTPicker_GotFocus txtSelectAmjMax


End Sub


Private Sub txtSelectAmjMax_LostFocus()
DTPicker_LostFocus txtSelectAmjMax

End Sub


Private Sub txtSelectAmjMin_GotFocus()
DTPicker_GotFocus txtSelectAmjMin

End Sub


Private Sub txtSelectAmjMin_LostFocus()
DTPicker_LostFocus txtSelectAmjMin

End Sub


Private Sub txtUpdateCompte_GotFocus()

txt_GotFocus txtUpdateCompte

End Sub


Private Sub txtUpdateCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtUpdateCompte_LostFocus()
txt_LostFocus txtUpdateCompte
txtUpdateCompte_Control
End Sub

Private Sub txtSelectService_GotFocus()
txt_GotFocus txtSelectService
End Sub

Private Sub txtSelectService_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub

Private Sub txtSelectService_LostFocus()
txt_LostFocus txtSelectService
End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
picInfo.BackColor = MouseMoveUsr.BackColor
'picInfo2.BackColor = MouseMoveUsr.BackColor

wSelectAmjMin = AmjCptVeille
wSelectAmjMax = AmjCptVeille
Call DTPicker_Set(txtSelectAmjMin, wSelectAmjMin)
Call DTPicker_Set(txtSelectAmjMax, wSelectAmjMax)

AmjOpéMin = dateElp("FinDeMoisP", 0, DSys)
If mId$(DSys, 7, 2) < 4 Then AmjOpéMin = dateElp("FinDeMoisP", 0, AmjOpéMin)

meCV1 = CV_Euro
meCV1.CoursCompta = "C"
meCV1.OpéAmj = DSys
meCV1.Normal = "P"
meCV1.AchatVente = " "
meCV2 = meCV1: meCV3 = meCV1

recCompteInit meCompteOri
meCompteOri.Société = SocId$
meCompteOri.Agence = SocAgence$
meCompteDes = meCompteOri

recBiaLog_Init meBiaLog

currentAction = ""
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
fgSelect_Reset
blncmdUpdateOK_Visible = False: blncmdUpdateNOK_Visible = False

txtSelectAmjMin.Visible = False: txtSelectAmjMax.Visible = False
optSelectEnCours.Value = "1"

wSelectService = "000"
If usrService_DisplayAll Then
    chkSelectService.Value = "0": txtSelectService = "": txtSelectService.Visible = False
Else
    txtSelectService = usrService: txtSelectService.Visible = True
    chkSelectService.Value = "1"
    chkSelectService.Enabled = False: txtSelectService.Enabled = False
End If
txtUpdateS.Locked = False
txtUpdateD.Locked = False
txtUpdateE.Locked = False
txtUpdateF.Locked = False


blnControl = True
End Sub


Private Sub chkSelectService_Click()
On Error GoTo Exit_Sub
If chkSelectService = "1" Then
    txtSelectService.Visible = True: If blnSetfocus Then txtSelectService.SetFocus
Else
    txtSelectService.Visible = False
End If
Exit_Sub:

End Sub

Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 Then
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

Private Sub fgSelect_Display()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

cmdUpdateQuit
sstab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For meCptEAR_Index = 1 To meCptEAR_Nb
    If meCptEAR(meCptEAR_Index).Method <> constIgnore And meCptEAR(meCptEAR_Index).Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next meCptEAR_Index

fgSelect_SortAD = 5
fgSelect_Sort1_Old = 1: fgSelect_Sort1 = 1
If fgSelect.Rows > 1 Then fgSelect_SortX 1

End Sub
Public Sub fgSelect_DisplayLine()

fgSelect.Col = 0:
If meCptEAR(meCptEAR_Index).EARStatus <> "   " Then
    recStatut.K2 = meCptEAR(meCptEAR_Index).EARStatus: recStatut.Name = recStatut.K2: tableElpTable_Read recStatut: fgSelect.Text = recStatut.Name
Else
    fgSelect.Text = ""
End If
fgSelect.Col = 1: fgSelect.Text = meCptEAR(meCptEAR_Index).EARIdRef

fgSelect.Col = 2
Call CV_AttributS(meCptEAR(meCptEAR_Index).Devise, meCV1)
If meCptEAR(meCptEAR_Index).MONDEV < 0 Then
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.CellForeColor = vbBlue
End If
fgSelect.Text = Format$(meCptEAR(meCptEAR_Index).MONDEV, "### ### ###.00") & "  " & meCV1.DeviseIso
fgSelect.Col = 3: fgSelect.Text = meCptEAR(meCptEAR_Index).LIBELE

fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = meCptEAR_Index

End Sub
Public Sub fgSelect_Load()
Dim X As String, mMethod As String

recCptEAR_Init xCptEAR

xCptEAR.Method = "SnapL0"

xCptEAR.COSOC = SocId$
xCptEAR.Agence = SocAgence$
xCptEAR.SERVIC = wSelectService

meCptEAR(0) = xCptEAR
meCptEAR(0).EARIdRef = 9999999
If Not blnSelectService Then
    xCptEAR.SERVIC = "000"
    meCptEAR(0).SERVIC = "999"
End If

If blnSelectAmj Then
    xCptEAR.Method = "SnapL1"
    xCptEAR.EARCptAmj = wSelectAmjMin
    meCptEAR(0).EARCptAmj = wSelectAmjMax
End If

If optSelectàValider Then xCptEAR.EARStatus = "à": meCptEAR(0).EARStatus = "à99"
If optSelectAnnulation Then xCptEAR.EARStatus = "A": meCptEAR(0).EARStatus = "A99"

Call srvCptEAR_Load(xCptEAR, meCptEAR(0))

meCptEAR_Nb = srvCptEAR.arrCptEAR_Nb
meCptEAR_NbMax = meCptEAR_Nb + 1: ReDim meCptEAR(meCptEAR_NbMax)
For I = 1 To meCptEAR_Nb
    meCptEAR(I) = srvCptEAR.arrCptEAR(I)
Next I

fgSelect_Display
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
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    meCptEAR_Index = Val(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
    Select Case lK
        Case 1: fgSelect.Text = Format$(meCptEAR(meCptEAR_Index).EARIdRef, "00000000")
        Case 2: fgSelect.Text = Format$(meCptEAR(meCptEAR_Index).MONDEV, "000000000000000.00")
        Case fgSelect_arrIndex: fgSelect.Text = Format$(meCptEAR_Index, "0000000000")
    End Select
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

sstab1.Tab = 0
ReDim meCptEAR(10)

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
ReDim meCompte(10)

paramCPTEAR_CRI = "00038151005"
paramCPTEAR_SIT = "00038152001"

cmdReset

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
Me.PopupMenu mnucmdPrint, vbPopupMenuLeftButton


End Sub

Private Sub cmdSelect_Click()
cmdControl
If lstErr.ListCount = 0 Then fgSelect_Load
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

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String

If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1: fgSelect_SortX 1
        Case 2:  fgSelect_SortX 2
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex
        meCptEAR_Index = Val(fgSelect.Text)
        mCptEAR = meCptEAR(meCptEAR_Index)
        xCptEAR = mCptEAR
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        
        fraUpdate.Enabled = False
        cmdUpdateOk.Visible = False
        cmdUpdateNOK.Visible = False
        optUpdateCompte = True
        moptAnnulation = False
        currentAction = ""
        meCompteDes.Numéro = "???"
        
        Select Case Trim(mCptEAR.EARStatus)
            Case "": currentAction = constSaisie
                    fraUpdate.Enabled = CptEARAut.Saisir: fraUpdateDta.Enabled = CptEARAut.Saisir
                    cmdUpdateOk.Visible = CptEARAut.Saisir: cmdUpdateOk.Caption = constàValider
            Case "àV": currentAction = constValider
                    fraUpdate.Enabled = CptEARAut.Valider: fraUpdateDta.Enabled = False
                    cmdUpdateNOK.Visible = CptEARAut.Valider: cmdUpdateOk.Visible = CptEARAut.Valider: cmdUpdateOk.Caption = constValider
            Case "àA": currentAction = constValider
                    fraUpdate.Enabled = CptEARAut.Valider: fraUpdateDta.Enabled = False
                    cmdUpdateNOK.Visible = CptEARAut.Valider: cmdUpdateOk.Visible = CptEARAut.Valider: cmdUpdateOk.Caption = constValider: moptAnnulation = True
            Case "àC": currentAction = constValider
                    fraUpdate.Enabled = CptEARAut.Valider: fraUpdateDta.Enabled = False
                    cmdUpdateNOK.Visible = CptEARAut.Valider: cmdUpdateOk.Visible = CptEARAut.Valider: cmdUpdateOk.Caption = constValider
            Case "A":  moptAnnulation = True
           
        End Select
     picInfo_Display
   
        If Button = vbRightButton Then
            mnufgSelect_Display = CptEARAut.Consulter
            mnufgSelect_Print = CptEARAut.Consulter
            Me.PopupMenu mnufgSelect, vbPopupMenuLeftButton
        End If
    End If
End If

End Sub
Public Sub picInfo_Display()
Dim X As String, xErr As String
Dim wLineHeight As Integer, wCol1 As Integer

wLineHeight = (picInfo.Height - 20) / 15
wCol1 = picInfo.TextWidth("Référence---")

xErr = Trim(DicLib(524, mCptEAR.LogCodErr))
libRéférenceInterne = "EAR : " & mCptEAR.EARIdRef & " : " & xErr

libUpdateCompte = ""
If mCptEAR.EARCptDes = "00000000000" Then
    txtUpdateCompte = mCptEAR.EARCptOri
Else
    txtUpdateCompte = mCptEAR.EARCptDes
End If

txtUpdateCompte_Control

optUpdateAnnulation = moptAnnulation

If mCptEAR.SERVIC = "001" Then optUpdateCRI.Enabled = True: optUpdateSIT.Enabled = True
Select Case mCptEAR.EARCptDes
    Case paramCPTEAR_CRI: optUpdateCRI.Value = True
    Case paramCPTEAR_SIT: optUpdateSIT.Value = True
End Select

optUpdate_Color

picInfo.Cls: picInfo2.Cls


picInfo.CurrentY = 0: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Référence";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print mCptEAR.EARIdRef;
picInfo.ForeColor = warnUsrColor
recStatut.K2 = mCptEAR.EARStatus: tableElpTable_Read recStatut: picInfo.Print " " & recStatut.Name;

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Rejet";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor: picInfo.FontBold = True: picInfo.Print Compte_Imp(mCptEAR.EARCptEAR);

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50
picInfo.ForeColor = warnUsrColor: picInfo.Print xErr;


picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Compte";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor: picInfo.FontBold = True: picInfo.Print Compte_Imp(mCptEAR.Compte);

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = libUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50
If meCompteOri.Numéro <> mCptEAR.Compte Then
    meCompteOri.Devise = mCptEAR.Devise
    meCompteOri.Numéro = mCptEAR.Compte
    V = mdbCptP0_Find(meCompteOri)
    If Not IsNull(V) Then meCompteOri.Numéro = mCptEAR.Compte: meCompteOri.Intitulé = "?compte inconnu"

End If
picInfo.Print Trim(meCompteOri.Intitulé);

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50: picInfo.Print "Montant";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
If mCptEAR.MONDEV < 0 Then
    picInfo.ForeColor = errUsr.ForeColor: txtUpdateE.ForeColor = errUsr.ForeColor
Else
    picInfo.ForeColor = libUsr.ForeColor: txtUpdateE.ForeColor = libUsr.ForeColor
End If
Call CV_AttributS(mCptEAR.Devise, meCV1)
X = Trim(Format$(mCptEAR.MONDEV, "#### ### ### ### ##0.00"))
picInfo.Print X & " " & meCV1.DeviseIso;
txtUpdateE = X

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = libUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50
picInfo.Print Trim(mCptEAR.LIBELE);

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Service";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print Trim(DicLib(4, mCptEAR.SERVIC));

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Opération";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print Trim(DicLib(27, mCptEAR.BIACOP));

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Date opé";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print dateImp(mCptEAR.AMJOPE);

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Date valeur";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print dateImp(mCptEAR.AMJVAL);

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Lot.Pièce.";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print Trim(Format$(mCptEAR.NUMLOT, "###0")) & "_" & Trim(Format$(mCptEAR.NUMPIE, "######0")) & "_" & Trim(Format$(mCptEAR.NOLIGN, "###0"));

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Opérateur";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print mCptEAR.NOMOP;

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Programme";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print mCptEAR.NOPROG;

picInfo.CurrentY = picInfo.CurrentY + wLineHeight: picInfo.ForeColor = lblUsr.ForeColor: picInfo.FontBold = False
picInfo.CurrentX = 50:  picInfo.Print "Dossier";
picInfo.CurrentX = wCol1: picInfo.Print ": ";
picInfo.ForeColor = libUsr.ForeColor:  picInfo.Print mCptEAR.REFCON;


wLineHeight = (picInfo2.Height - 20) / 5

picInfo2.CurrentY = 0: picInfo2.ForeColor = lblUsr.ForeColor: picInfo2.FontBold = False

picInfo2.CurrentX = 50:  picInfo2.Print "Saisie";
picInfo2.CurrentX = wCol1: picInfo2.Print ": ";
picInfo2.ForeColor = libUsr.ForeColor:  picInfo2.Print mCptEAR.EARMajUsr;

picInfo2.CurrentY = picInfo2.CurrentY + wLineHeight: picInfo2.ForeColor = lblUsr.ForeColor: picInfo2.FontBold = False
picInfo2.CurrentX = 50:  picInfo2.Print "";
picInfo2.CurrentX = wCol1: picInfo2.Print ": ";
picInfo2.ForeColor = libUsr.ForeColor:  picInfo2.Print dateImp10(mCptEAR.EARMajAmj) & "   " & timeImp(mCptEAR.EARMajHms);

picInfo2.CurrentY = picInfo2.CurrentY + wLineHeight: picInfo2.ForeColor = lblUsr.ForeColor: picInfo2.FontBold = False
picInfo2.CurrentX = 50:  picInfo2.Print "Validation";
picInfo2.CurrentX = wCol1: picInfo2.Print ": ";
picInfo2.ForeColor = libUsr.ForeColor:  picInfo2.Print mCptEAR.EARValUsr;

picInfo2.CurrentY = picInfo2.CurrentY + wLineHeight: picInfo2.ForeColor = lblUsr.ForeColor: picInfo2.FontBold = False
picInfo2.CurrentX = 50:  picInfo2.Print "";
picInfo2.CurrentX = wCol1: picInfo2.Print ": ";
picInfo2.ForeColor = libUsr.ForeColor:  picInfo2.Print dateImp10(mCptEAR.EARValAmj) & "   " & timeImp(mCptEAR.EARValHms);

picInfo2.CurrentY = picInfo2.CurrentY + wLineHeight: picInfo2.ForeColor = lblUsr.ForeColor: picInfo2.FontBold = False
picInfo2.CurrentX = 50:  picInfo2.Print "Lot.Pièce.";
picInfo2.CurrentX = wCol1: picInfo2.Print ": ";
picInfo2.ForeColor = libUsr.ForeColor:  picInfo2.Print Trim(Format$(mCptEAR.EARNumLot, "###0")) & "_" & Trim(Format$(mCptEAR.EARNumPie, "######0")) & "_" & Trim(Format$(mCptEAR.EARNoLign, "###0"));




''Bialog_Load
End Sub


Private Sub txtXXX_GotFocus()

'KeyAscii = convUCase(KeyAscii)

'txt_GotFocus txtXXX

'txt_LostFocus txtXXX
'If blnControl Then cmdControl

'DTPicker_GotFocus txtXXX

'DTPicker_LostFocus txtXXX
'If blnControl Then cmdControl


' Change : txtAmjfin_control
'MouseMoveActiveControl_Set txtXXX

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
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), CptEARAut)    ' "EAR"

blnSetfocus = True
Form_Init
If UCase$(Trim(mId$(Msg, 13, 12))) = "BIA_EXPLOIT" Then
    optSelectAll.Value = True
    cmdSelect_Click
    mnucmdPrint_fgSelect_All_Click
    Unload Me
End If

End Sub


Public Sub cmdContext_Return()
    SendKeys "{TAB}"
    

End Sub



Public Sub fgSelect_Reset()
fgSelect_Sort1 = 1: fgSelect_Sort2 = 1
fgSelect_Sort1_Old = 0
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 6
blnfgSelect_DisplayLine = False

End Sub

Public Sub cmdUpdate_Control()
lstErr.Clear
lstErr.Height = 200
 Select Case Trim(mCptEAR.EARStatus)   ' mCPTEAR !!!!!!
    Case "":
        If optUpdateAnnulation Then
            xCptEAR.EARStatus = "àA"
        Else
            txtUpdateCompte_Control
            xCptEAR.EARStatus = "àC"
 '           xCptEAR.REFCON = "EAR" & Format$(mCptEAR.EARIdRef, "0000000")
 '           xCptEAR.LIBELE = "EARX " & Trim(Format$(mCptEAR.EARIdRef, "#######")) & dateImp(mCptEAR.EARCptAmj)
        End If
    Case "àC":
        '$JPL 2002.11.04 date d'opé = date de validation cf note COMPTA
        '    If xCptEAR.AMJOPE <= AmjOpéMin Then
        '        If Not usrSituationCompte_Forçage Then
        '            X = "? Date Opération : " & xCptEAR.AMJOPE & " <= " & AmjOpéMin
        '            Call MsgBox(X & Chr$(13) & "Validation uniquement par le service comptable", vbCritical, "frmCptEAR.cmdUpdate_Control")
        '            Call lstErr_AddItem(lstErr, cmdContext, X)
        '        End If
        '    End If
            meCompteDes.Numéro = -1
            txtUpdateCompte_Control
            xCptEAR.EARStatus = IIf(blnCompteSituation_Forçage, "CF", "C")
            xCptEAR.REFCON = "EAR" & Format$(mCptEAR.EARIdRef, "0000000")
            xCptEAR.LIBELE = "EARX : " & Trim(Format$(mCptEAR.EARIdRef, "0000000")) & " du " & dateImp10(mCptEAR.EARCptAmj) _
                                 & " Réf : " & Trim(Format$(mCptEAR.NUMLOT, "###0")) & "_" & Trim(Format$(mCptEAR.NUMPIE, "######0")) & "_" & Trim(Format$(mCptEAR.NOLIGN, "###0"))
            

     Case "àA": xCptEAR.EARStatus = "A": blnCompteSituation_Validation = True
     Case Else:             Call lstErr_AddItem(lstErr, cmdContext, "? Status (frmCptEAR.cmdUpdate_Control)")

End Select
End Sub

Public Sub optUpdate_Color()
optUpdateCompte.ForeColor = fraUpdate.ForeColor
optUpdateCRI.ForeColor = fraUpdate.ForeColor
optUpdateSIT.ForeColor = fraUpdate.ForeColor
optUpdateAnnulation.ForeColor = fraUpdate.ForeColor
txtUpdateCompte.Enabled = False

If optUpdateAnnulation Then
    optUpdateAnnulation.ForeColor = warnUsrColor
Else
    If optUpdateSIT Then
        optUpdateSIT.ForeColor = warnUsrColor
        txtUpdateCompte = paramCPTEAR_SIT
    Else
        If optUpdateCRI Then
            optUpdateCRI.ForeColor = warnUsrColor
            txtUpdateCompte = paramCPTEAR_CRI
        Else
            optUpdateCompte.ForeColor = warnUsrColor
            txtUpdateCompte.Enabled = True
        End If
    End If
End If

End Sub

Public Sub cmdUpdate_Db()
If lstErr.ListCount = 0 Then
    V = srvCptEAR_Update(xCptEAR)
    If IsNull(V) Then
        meCptEAR(meCptEAR_Index) = xCptEAR
        mCptEAR = xCptEAR
        fgSelect_DisplayLine
        picInfo_Display
        If currentAction = constValider Then prtCptEAR_Avis
    End If
End If

End Sub

Public Sub txtUpdateCompte_Control()

xCptEAR.EARCptDes = Format$(Val(txtUpdateCompte), "00000000000")
If xCptEAR.EARCptDes <> meCompteDes.Numéro Then
    meCompteDes.Société = SocId$
    meCompteDes.Agence = SocAgence$
    meCompteDes.Devise = xCptEAR.Devise
    meCompteDes.Numéro = xCptEAR.EARCptDes
    
    V = srvCompte_InitFind(meCompteDes)
    
    If Not IsNull(V) Then
        blnCompteSituation_Saisie = False
        recCompteInit meCompteDes
        V = "? compte inconnu : " & xCptEAR.Devise & " _" & xCptEAR.EARCptDes
    Else
        Select Case currentAction
            Case constValider: V = CompteSituation_Validation(meCompteDes, blnCompteSituation_Validation, blnCompteSituation_Forçage)
            Case constSaisie: V = CompteSituation_Saisie(meCompteDes, blnCompteSituation_Saisie)
        End Select
    End If
    If Not IsNull(V) Then meCompteDes.Numéro = "$$$": Call lstErr_AddItem(lstErr, lstErr, "? " & V):

    libUpdateCompte = meCompteDes.Intitulé
    picInfo_DisplaySolde

End If
End Sub

Public Sub Bialog_Load()

meBiaLog.Method = "SeekP0"
meBiaLog.Log_Cosoc = mCptEAR.COSOC
meBiaLog.Log_Agence = mCptEAR.Agence

meBiaLog.Log_CptAmj = mCptEAR.EARCptAmj
meBiaLog.Log_Cpteur = mCptEAR.LogCpteur

If Not IsNull(srvBiaLog_Monitor(meBiaLog)) Then recBiaLog_Init meBiaLog

End Sub

Public Sub picInfo_DisplaySolde()
If meCompteDes.SoldeInstantané < 0 Then
    txtUpdateS.ForeColor = errUsr.ForeColor
Else
    txtUpdateS.ForeColor = libUsr.ForeColor
End If
txtUpdateS = Format$(meCompteDes.SoldeInstantané, "#### ### ### ### ##0.00")
curX = meCompteDes.SoldeInstantané + mCptEAR.MONDEV

lblUpdateD = "Découvert => " & dateImp10(meCompteDes.DécouvertAmj)
If Val(meCompteDes.DécouvertAmj) > DSys Then
    curX = curX + meCompteDes.DécouvertMontant
    txtUpdateD = Format$(meCompteDes.DécouvertMontant, "#### ### ### ### ##0.00")
Else
    txtUpdateD = ""
End If

If curX < 0 Then
    txtUpdateF.ForeColor = errUsr.ForeColor
Else
    txtUpdateF.ForeColor = libUsr.ForeColor
End If
txtUpdateF = Format$(curX, "#### ### ### ### ##0.00")


End Sub

Public Sub Compte_Sel()
If Not IsNull(selCompte_Load(meCompte(1), meCompte(0), "Init")) Then Exit Sub

meCompte_Nb = selCompte_Nb
meCompte_NbMax = selCompte_Nb + 1
ReDim meCompte(meCompte_NbMax)

For I = 1 To selCompte_Nb
    meCompte(I) = selCompte(I)
Next I

Call selCompte_Load(meCompte(1), meCompte(0), "End")

End Sub
